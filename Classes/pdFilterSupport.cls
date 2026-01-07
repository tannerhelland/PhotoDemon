VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdFilterSupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Filter Support Class
'Copyright 2013-2026 by Tanner Helland
'Created: 15/January/13
'Last updated: 02/November/22
'Last update: new helper for handling edges on normalized values (range [0, 1]); the Perspective transform now
'             uses this for fast edge-handling when custom foreshortening is active
'
'Per its name, this class provides support routines for certain types of image filters, namely: filters
' that move pixels.  Automated edge-handling (with a variety of approaches) and interpolation are key features.
'
'The main utilizer of this class is the Effects -> Distort menu.  See any Distort tool for example usage.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private m_FinalX As Double, m_FinalY As Double
Private m_imgWidth As Double, m_imgHeight As Double
Private m_EdgeMethod As PD_EdgeOperator
Private m_Interpolate As Boolean
Private m_DstX As Double, m_DstY As Double

'PD now allows functions to "unsafely" pass a reference to a source image array.
' The benefit of doing this in advance is that it makes subsequent array references very fast.
' The downside is that you *MUST* free the reference manually or crashes will occur.
'
'(This is technically bypassable by setting different flags in the passed SafeArray header,
' but I like to deliberately enforce strict array freeing behavior in PhotoDemon as it removes
' the risk for surprise crashes down the road.)
'
'Anyway - not all filters use these faster, less-safe wrappers yet, but I hope to migrate
' the whole program to them in the coming days.
Private m_PixelSA2D As SafeArray2D, m_Pixels() As RGBQuad

'New, unsafe mechanism for persistently aliasing a target DIB.
' YOU MUST MANUALLY UNALIAS BEFORE RELEASE THIS CLASS.
'Alias/unalias the target pixel array, specifically.
Friend Sub AliasTargetDIB(ByRef srcDIB As pdDIB)
    With m_PixelSA2D
        .cbElements = 4
        .cDims = 2
        .cLocks = 1
        .Bounds(0).lBound = 0
        .Bounds(0).cElements = srcDIB.GetDIBHeight()
        .Bounds(1).lBound = 0
        .Bounds(1).cElements = srcDIB.GetDIBWidth()
        .pvData = srcDIB.GetDIBPointer()
    End With
    PutMem4 VarPtrArray(m_Pixels()), VarPtr(m_PixelSA2D)
End Sub

Friend Sub UnaliasTargetDIB()
    PutMem4 VarPtrArray(m_Pixels()), 0&
End Sub

'If a function simply wants RGB/A values returned for a given source position, then can use this function to capture
' those values.  This is helpful for tools that calculate multiple source positions (like a Blur tool), and need to
' do something with the colors from those various positions.
'
'Said another way, the caller need only calculate srcX and srcY, then this class will do the rest!
Friend Function GetColorsFromSource(ByVal srcX As Double, ByVal srcY As Double, ByVal origX As Long, ByVal origY As Long) As RGBQuad
    
    'Cache the original x/y values as necessary
    If (m_EdgeMethod = pdeo_Original) Then
        m_DstX = origX
        m_DstY = origY
    End If
    
    'fixDistort will only be set to TRUE when the current pixel needs to be erased
    ' (e.g. the edge mode is set to pdeo_Erase and this pixel lies outside the image).
    '
    'When this occurs, note that the function return has already been intialized to
    ' [0, 0, 0, 0] so we don't need to do anything.  If the function returns FALSE,
    ' however, srcX and srcY will contain potentially modified source coordinates
    ' (with the current edge-handling settings taken into account).  We then need to
    ' interpolate a value for that position, or simply return an integer-aligned
    ' copy of the source at that position.
    If (Not FixDistortEdges(srcX, srcY)) Then
        
        'Interpolate a new pixel value
        If m_Interpolate Then
            GetColorsFromSource = BilinearInterpolate(srcX, srcY)
            
        'Clamp to the nearest integer coordinate value, and note that we *cannot* round here
        ' (because srcX and srcY may be e.g. 99.99 on a 99-px image, and rounding could cause access errors)
        Else
            GetColorsFromSource = m_Pixels(Int(srcX), Int(srcY))
        End If
        
    End If

End Function

'This function is a fast variant of the standard GetColorsFromSource() function, with this IMPORTANT CAVEAT:
' it does *NOT* support edge-handling.  If the source pixel lies off-image, this function will return a
' transparent pixel.  This is helpful for e.g. the Droste effect, which needs to iterate several times before
' caring about edge-handling, so on the early iterations, we can skip edge-handling entirely for a nice perf boost.
Friend Function GetColorsFromSource_FastErase(ByVal srcX As Double, ByVal srcY As Double) As RGBQuad
    If (srcX >= 0#) Then
        If (srcX <= m_FinalX) Then
            If (srcY >= 0#) Then
                If (srcY <= m_FinalY) Then
                    If m_Interpolate Then
                        GetColorsFromSource_FastErase = BilinearInterpolate(srcX, srcY)
                    Else
                        GetColorsFromSource_FastErase = m_Pixels(Int(srcX + 0.5), Int(srcY + 0.5))
                    End If
                End If
            End If
        End If
    End If
End Function

'When this support class is first initialized by a function, a number of parameters are specified.  It is much faster
' to set these once and store them instead of passing them every time a pixel needs to be handled.
Friend Sub SetDistortParameters(ByVal edgeMethod As PD_EdgeOperator, ByVal toInterpolate As Boolean, ByVal finalX As Long, ByVal finalY As Long)
    m_EdgeMethod = edgeMethod
    m_Interpolate = toInterpolate
    m_FinalX = finalX
    m_FinalY = finalY
    m_imgWidth = m_FinalX + 0.99999999  'Cheat to avoid the need for subsequent bounds-checking
    m_imgHeight = m_FinalY + 0.99999999
End Sub

'To use FixDistortEdges externally, call this safe wrapper.  Returns TRUE if pixel needs to be erased; FALSE otherwise.
' By design, this function always returns coordinates on the range [0, 1] where [1 = img_size_in_x/y_direction].
Friend Function HandleEdgesOnly_Normalized(ByRef srcX As Double, ByRef srcY As Double, ByRef useOrigCoords As Boolean) As Boolean
    
    useOrigCoords = False
    
    Select Case m_EdgeMethod
    
        Case pdeo_Clamp
        
            If (srcX < 0#) Then
                srcX = 0#
            Else
                If (srcX > 1#) Then srcX = 1#
            End If
            
            If (srcY < 0#) Then
                srcY = 0#
            Else
                If (srcY > 1#) Then srcY = 1#
            End If
            
        Case pdeo_Reflect
        
            srcX = PDMath.Modulo(srcX, 2#)
            srcY = PDMath.Modulo(srcY, 2#)
            If (srcX > 1#) Then srcX = 2# - srcX
            If (srcY > 1#) Then srcY = 2# - srcY
            
        Case pdeo_Wrap
            
            srcX = srcX - Int(srcX)
            srcY = srcY - Int(srcY)
            
        Case pdeo_Erase
        
            If (srcX < 0#) Then
                HandleEdgesOnly_Normalized = True
                Exit Function
            End If
            
            If (srcY < 0#) Then
                HandleEdgesOnly_Normalized = True
                Exit Function
            End If
            
            If (srcX > 1#) Then
                HandleEdgesOnly_Normalized = True
                Exit Function
            End If
            
            If (srcY > 1#) Then
                HandleEdgesOnly_Normalized = True
                Exit Function
            End If
            
        Case pdeo_Original
            
            If (srcX < 0#) Or (srcY < 0#) Then
                useOrigCoords = True
                Exit Function
            End If
            
            If (srcX > 1#) Or (srcY > 1#) Then
                useOrigCoords = True
                Exit Function
            End If
            
    End Select
    
End Function

'IMPORTANT NOTE: this function does not handle pixel erasing in ERASE mode.  This function also assumes
' the source x/y position actually exists in-bounds (because this is how the normal color handler works).
' To use this shortcut function, as PhotoDemon does in the PerspectiveImage function, you must:
' 1) Handle pixel erasing manually, as relevant
' 2) Ensure srcX and srcY are IN-BOUNDS
Friend Function HandleInterpolationOnly(ByVal srcX As Double, ByVal srcY As Double) As RGBQuad

    'Interpolate a new pixel value
    If m_Interpolate Then
        HandleInterpolationOnly = BilinearInterpolate(srcX, srcY)
        
    'Clamp to the nearest integer coordinate value, and note that we *cannot* round here
    ' (because srcX and srcY may be e.g. 99.99 on a 99-px image, and rounding could cause access errors)
    Else
        HandleInterpolationOnly = m_Pixels(Int(srcX), Int(srcY))
    End If
    
End Function

'If a pixel lies outside image boundaries, move it in-bounds using one of several methods
' If the edge handle method is "Erase", this function will return a boolean indicating whether the supplied pixel
' must be erased.  If FALSE is returned, the pixel can be handled normally.
Private Function FixDistortEdges(ByRef srcX As Double, ByRef srcY As Double) As Boolean

    Select Case m_EdgeMethod
    
        Case pdeo_Clamp
        
            If (srcX >= 0#) Then
                If (srcX > m_FinalX) Then srcX = m_FinalX
            Else
                srcX = 0#
            End If
            
            If (srcY >= 0#) Then
                If (srcY > m_FinalY) Then srcY = m_FinalY
            Else
                srcY = 0#
            End If
            
        Case pdeo_Reflect
        
            srcX = PDMath.Modulo(srcX, m_FinalX * 2)
            srcY = PDMath.Modulo(srcY, m_FinalY * 2)
            If (srcX > m_FinalX) Then srcX = m_FinalX * 2 - srcX
            If (srcY > m_FinalY) Then srcY = m_FinalY * 2 - srcY
            
        Case pdeo_Wrap
            
            If (srcX >= 0#) Then
                If (srcX >= m_imgWidth) Then srcX = PDMath.Modulo(srcX, m_imgWidth)
            Else
                srcX = PDMath.Modulo(srcX, m_imgWidth)
            End If
            
            If (srcY >= 0#) Then
                If (srcY >= m_imgHeight) Then srcY = PDMath.Modulo(srcY, m_imgHeight)
            Else
                srcY = PDMath.Modulo(srcY, m_imgHeight)
            End If
            
        Case pdeo_Erase
        
            If (srcX < 0#) Then
                FixDistortEdges = True
                Exit Function
            End If
            
            If (srcY < 0#) Then
                FixDistortEdges = True
                Exit Function
            End If
            
            If (srcX > m_FinalX) Then
                FixDistortEdges = True
                Exit Function
            End If
            
            If (srcY > m_FinalY) Then
                FixDistortEdges = True
                Exit Function
            End If
            
        Case pdeo_Original
        
            If (srcX < 0#) Or (srcY < 0#) Then
                srcX = m_DstX
                srcY = m_DstY
                Exit Function
            End If
            
            If (srcX > m_FinalX) Or (srcY > m_FinalY) Then
                srcX = m_DstX
                srcY = m_DstY
                Exit Function
            End If
            
    End Select
    
End Function

'This function takes an x and y value - as floating-point - and uses their position to calculate an interpolated value
' for an imaginary pixel in that location.  Results are returned as an RGBQuad.
Private Function BilinearInterpolate(ByVal x1 As Double, ByVal y1 As Double) As RGBQuad
    
    'We first need to retrieve the four surrounding pixel values
    Dim topLeft As RGBQuad, topRight As RGBQuad, bottomLeft As RGBQuad, bottomRight As RGBQuad
    
    'Pixel coordinates are always guaranteed in-bounds on the left/top boundary
    Dim intLeft As Long, intTop As Long
    intLeft = Int(x1)
    intTop = Int(y1)
    topLeft = m_Pixels(intLeft, intTop)
    
    'The other three pixels can potentially lie out-of-bounds, so we need to bounds-check 'em
    Dim fixX As Double, fixY As Double
    
    'Pixels at the far edges of the image require special treatment during interpolation
    ' (treatment that varies depending on the current edge-handling model).
    
    'Top-right next
    If (x1 < m_FinalX) Then
    
        topRight = m_Pixels(intLeft + 1, intTop)
        
        'We know the x-coord is in-bounds; perform a quick check on y to see if we can short-circuit
        ' the rest of the function.  (If x is in-bounds, the odds of y also being in-bounds is very high.)
        If (y1 < m_FinalY) Then
            intTop = intTop + 1
            bottomLeft = m_Pixels(intLeft, intTop)
            bottomRight = m_Pixels(intLeft + 1, intTop)
            GoTo ShortCircuit
        End If
        
    Else
        fixX = x1 + 1#
        fixY = y1
        If FixDistortEdges(fixX, fixY) Then PutMem4 VarPtr(topRight), 0& Else topRight = m_Pixels(Int(fixX), Int(fixY))
    End If
    
    'Bottom-left
    If (y1 < m_FinalY) Then
        bottomLeft = m_Pixels(intLeft, intTop + 1)
    Else
        fixX = x1
        fixY = y1 + 1#
        If FixDistortEdges(fixX, fixY) Then PutMem4 VarPtr(bottomLeft), 0& Else bottomLeft = m_Pixels(Int(fixX), Int(fixY))
    End If
    
    If (x1 < m_FinalX) And (y1 < m_FinalY) Then
        bottomRight = m_Pixels(intLeft + 1, intTop + 1)
    Else
        fixX = x1 + 1#
        fixY = y1 + 1#
        If FixDistortEdges(fixX, fixY) Then PutMem4 VarPtr(bottomRight), 0& Else bottomRight = m_Pixels(Int(fixX), Int(fixY))
    End If
    
ShortCircuit:
    
    'Calculate blend ratios
    Dim yBlend As Double, xBlend As Double, xBlendInv As Double, yBlendInv As Double
    yBlend = y1 - Int(y1)
    yBlendInv = 1# - yBlend
    xBlend = x1 - Int(x1)
    xBlendInv = 1# - xBlend
    
    'Blend in the x-direction, then y-direction, for each color component
    Dim topRow As Double, bottomRow As Double
    
    topRow = topRight.Blue * xBlend + topLeft.Blue * xBlendInv
    bottomRow = bottomRight.Blue * xBlend + bottomLeft.Blue * xBlendInv
    BilinearInterpolate.Blue = bottomRow * yBlend + topRow * yBlendInv
    
    topRow = topRight.Green * xBlend + topLeft.Green * xBlendInv
    bottomRow = bottomRight.Green * xBlend + bottomLeft.Green * xBlendInv
    BilinearInterpolate.Green = bottomRow * yBlend + topRow * yBlendInv
    
    topRow = topRight.Red * xBlend + topLeft.Red * xBlendInv
    bottomRow = bottomRight.Red * xBlend + bottomLeft.Red * xBlendInv
    BilinearInterpolate.Red = bottomRow * yBlend + topRow * yBlendInv
    
    topRow = topRight.Alpha * xBlend + topLeft.Alpha * xBlendInv
    bottomRow = bottomRight.Alpha * xBlend + bottomLeft.Alpha * xBlendInv
    BilinearInterpolate.Alpha = bottomRow * yBlend + topRow * yBlendInv
    
End Function
