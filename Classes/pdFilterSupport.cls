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
'Copyright 2013-2020 by Tanner Helland
'Created: 15/January/13
'Last updated: 21/May/14
'Last update: add new getColorsFromSource function, which can be used to RGB values from anywhere in the image,
'              but with all the usual support for edge wrapping, interpolation, and more.  The updated Fragment
'              tool makes use of these capabilities.
'
'Per its name, this class provides support routines for certain types of image filters, namely: filters
' that move pixels.  Automated edge-handling (with a variety of approaches) and interpolation are key features.
'
'The main utilizer of this class is the Effects -> Distort menu.  See any Distort tool for example usage.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private m_FinalX As Long, m_FinalY As Long
Private m_imgWidth As Long, m_imgHeight As Long
Private m_EdgeMethod As Long
Private m_Interpolate As Boolean
Private m_DstX As Long, m_DstY As Long

'PD now allows functions to "unsafely" pass a reference to a source image array.
' The benefit of doing this in advance is that it makes subsequent array references very fast.
' The downside is that you *MUST* free the reference manually or crashes will occur.
' (This is technically bypassable by setting flags in the SafeArray header, but I like to enforce
' strict array freeing behavior in PhotoDemon as it removes the risk for surprise crashes down
' the road.)
'
'Anyway - not all filters use the faster, less-safe wrappers yet, but I hope to migrate
' the whole program to them in the future.
Private Declare Sub SafeArrayLock Lib "oleaut32" (ByVal ptrToSA As Long)
Private Declare Sub SafeArrayUnlock Lib "oleaut32" (ByVal ptrToSA As Long)
Private m_PixelSA2D As SafeArray2D, m_Pixels() As RGBQuad
Private m_ByteArrayPtr As Long, m_ArrayWidth As Long, m_ArrayHeight As Long

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
    CopyMemory ByVal VarPtrArray(m_Pixels()), VarPtr(m_PixelSA2D), 4&
End Sub

Friend Sub UnaliasTargetDIB()
    CopyMemory ByVal VarPtrArray(m_Pixels()), 0&, 4&
End Sub

'If a function simply wants RGB/A values returned for a given source position, then can use this function to capture
' those values.  This is helpful for tools that calculate multiple source positions (like a Blur tool), and need to
' do something with the colors from those various positions.
'
'Said another way, the caller need only calculate srcX and srcY, then this class will do the rest!
Friend Sub GetColorsFromSource(ByRef r As Long, ByRef g As Long, ByRef b As Long, ByRef a As Long, ByRef srcX As Double, ByRef srcY As Double, ByRef srcData() As Byte, Optional ByVal origX As Long, Optional ByVal origY As Long)
    
    'Cache the original x/y values as necessary
    If (m_EdgeMethod = EDGE_ORIGINAL) Then
        m_DstX = origX
        m_DstY = origY
    End If
    
    'First, fix any coordinates that lie outside the image
    Dim fixDistort As Boolean
    fixDistort = FixDistortEdges(srcX, srcY)
    
    'Now, interpolate as necessary
    Dim quickValX As Long, quickValY As Long
    
    'fixDistort will only be set to TRUE when the current pixel needs to be erased
    If fixDistort Then
        r = 0
        g = 0
        b = 0
        a = 0
    Else
    
        'Interpolate a new pixel value
        If m_Interpolate Then
            
            b = PInterpolate(srcX, srcY, srcData, 0)
            g = PInterpolate(srcX, srcY, srcData, 1)
            r = PInterpolate(srcX, srcY, srcData, 2)
            a = PInterpolate(srcX, srcY, srcData, 3)
        
        'Round to the nearest coordinate value
        Else
    
            quickValX = Int(srcX) * 4
            quickValY = Int(srcY)
            
            b = srcData(quickValX, quickValY)
            g = srcData(quickValX + 1, quickValY)
            r = srcData(quickValX + 2, quickValY)
            a = srcData(quickValX + 3, quickValY)
            
        End If
        
    End If

End Sub

'If a function simply wants RGB/A values returned for a given source position, then can use this function to capture
' those values.  This is helpful for tools that calculate multiple source positions (like a Blur tool), and need to
' do something with the colors from those various positions.
'
'Said another way, the caller need only calculate srcX and srcY, then this class will do the rest!
Friend Function GetColorsFromSource_Fast(ByVal srcX As Double, ByVal srcY As Double, ByVal origX As Long, ByVal origY As Long) As RGBQuad
    
    'Cache the original x/y values as necessary
    If (m_EdgeMethod = EDGE_ORIGINAL) Then
        m_DstX = origX
        m_DstY = origY
    End If
    
    'fixDistort will only be set to TRUE when the current pixel needs to be erased
    ' (e.g. the edge mode is set to EDGE_ERASE and this pixel lies outside the image).
    ' When this occurs, note that the function return has already been intialized to
    ' [0, 0, 0, 0] so we don't need to do anything.  If the function returns FALSE,
    ' however, srcX and srcY will contain the newly calculated coordinates, which
    ' we will then pass to the interpolator as necessary.
    If (Not FixDistortEdges(srcX, srcY)) Then
        
        'Interpolate a new pixel value
        If m_Interpolate Then
            GetColorsFromSource_Fast = PInterpolate_Fast(srcX, srcY)
            
        'Clamp to the nearest integer coordinate value
        Else
            GetColorsFromSource_Fast = m_Pixels(Int(srcX), Int(srcY))
        End If
        
    End If

End Function

'After a transform has generated a set of reverse-mapped pixel coordinates, they can be passed to
' this function.  Based on already established parameters (such as interpolation, edge handling, etc),
' new pixel values are calculated and assigned automagically.
Friend Sub SetPixels(ByRef x As Long, ByRef y As Long, ByRef srcX As Double, ByRef srcY As Double, ByRef srcData() As Byte, ByRef dstData() As Byte)

    If (m_EdgeMethod = EDGE_ORIGINAL) Then
        m_DstX = x
        m_DstY = y
    End If

    Dim fixDistort As Boolean

    'First, fix any pixels that lie outside the image
    fixDistort = FixDistortEdges(srcX, srcY)
    
    'Now, interpolate as necessary
    Dim i As Long, quickVal As Long, quickVal2 As Long
    
    quickVal = x * 4
    
    'fixDistort will only be set to TRUE when the current pixel needs to be erased
    If fixDistort Then
        For i = 0 To 3
            dstData(quickVal + i, y) = 0
        Next i
    Else
        If m_Interpolate Then
            For i = 0 To 3
                dstData(quickVal + i, y) = PInterpolate(srcX, srcY, srcData, i)
            Next i
        Else
    
            quickVal2 = Int(srcX) * 4
        
            For i = 0 To 3
                dstData(quickVal + i, y) = srcData(quickVal2 + i, Int(srcY))
            Next i
        End If
    End If

End Sub

'When this support class is first initialized by a function, a number of parameters are specified.  It is much faster
' to set these once and store them instead of passing them every time a pixel needs to be handled.
Friend Sub SetDistortParameters(ByVal imgColorDepth As Long, ByVal edgeMethod As EDGE_OPERATOR, ByVal toInterpolate As Boolean, ByVal finalX As Long, ByVal finalY As Long)
    m_EdgeMethod = edgeMethod
    m_Interpolate = toInterpolate
    m_FinalX = finalX
    m_FinalY = finalY
    m_imgWidth = finalX + 1
    m_imgHeight = finalY + 1
End Sub

'If a pixel lies outside image boundaries, move it in-bounds using one of several methods
' If the edge handle method is "Erase", this function will return a boolean indicating whether the supplied pixel
' must be erased.  If FALSE is returned, the pixel can be handled normally.
Friend Function FixDistortEdges(ByRef srcX As Double, ByRef srcY As Double) As Boolean

    Select Case m_EdgeMethod
    
        Case EDGE_CLAMP
        
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
            
        Case EDGE_REFLECT
        
            srcX = PDMath.Modulo(srcX, m_FinalX * 2)
            srcY = PDMath.Modulo(srcY, m_FinalY * 2)
            If (srcX > m_FinalX) Then srcX = (m_FinalX - (srcX - m_FinalX))
            If (srcY > m_FinalY) Then srcY = (m_FinalY - (srcY - m_FinalY))
            
        Case EDGE_WRAP
            
            If (srcX >= 0#) Then
                If (srcX > m_FinalX) Then srcX = PDMath.Modulo(srcX, m_imgWidth)
            Else
                srcX = PDMath.Modulo(srcX, m_imgWidth)
            End If
            
            If (srcY >= 0#) Then
                If (srcY > m_FinalY) Then srcY = PDMath.Modulo(srcY, m_imgHeight)
            Else
                srcY = PDMath.Modulo(srcY, m_imgHeight)
            End If
            
        Case EDGE_ERASE
        
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
            
        Case EDGE_ORIGINAL
        
            If (srcX < 0#) Or (srcY < 0#) Then
                srcX = m_DstX
                srcY = m_DstY
            End If
            
            If (srcX > m_FinalX) Or (srcY > m_FinalY) Then
                srcX = m_DstX
                srcY = m_DstY
            End If
            
    End Select
    
End Function

'This function takes an x and y value - as floating-point - and uses their position to calculate an interpolated value
' for an imaginary pixel in that location.  Offset (r/g/b/alpha) and image color depth are also required.
Private Function PInterpolate(ByVal x1 As Double, ByVal y1 As Double, ByRef iData() As Byte, ByVal iOffset As Long) As Byte
    
    'Retrieve the four surrounding pixel values
    Dim topLeft As Double, topRight As Double, bottomLeft As Double, bottomRight As Double
    topLeft = iData(Int(x1) * 4 + iOffset, Int(y1))
    Dim fixX As Double, fixY As Double
    
    'Pixels at the far edges of the image require special treatment during interpolation
    If x1 < m_FinalX Then
        topRight = iData(Int(x1 + 1) * 4 + iOffset, Int(y1))
    Else
        fixX = x1 + 1
        fixY = y1
        If FixDistortEdges(fixX, fixY) Then
            topRight = 0
        Else
            topRight = iData(Int(fixX) * 4 + iOffset, Int(y1))
        End If
    End If
    If y1 < m_FinalY Then
        bottomLeft = iData(Int(x1) * 4 + iOffset, Int(y1 + 1))
    Else
        fixX = x1
        fixY = y1 + 1
        If FixDistortEdges(fixX, fixY) Then
            bottomLeft = 0
        Else
            bottomLeft = iData(Int(x1) * 4 + iOffset, Int(fixY))
        End If
    End If
    If x1 < m_FinalX Then
        If y1 < m_FinalY Then
            bottomRight = iData(Int(x1 + 1) * 4 + iOffset, Int(y1 + 1))
        Else
            fixX = x1 + 1
            fixY = y1 + 1
            If FixDistortEdges(fixX, fixY) Then
                bottomRight = 0
            Else
                bottomRight = iData(Int(x1 + 1) * 4 + iOffset, Int(fixY))
            End If
        End If
    Else
        fixX = x1 + 1
        fixY = y1 + 1
        If FixDistortEdges(fixX, fixY) Then
            bottomRight = 0
        Else
            If y1 < m_FinalY Then
                bottomRight = iData(Int(fixX) * 4 + iOffset, Int(y1 + 1))
            Else
                bottomRight = iData(Int(fixX) * 4 + iOffset, Int(fixY))
            End If
        End If
    End If
    
    'Calculate blend ratios
    Dim yBlend As Double
    Dim xBlend As Double, xBlendInv As Double
    yBlend = y1 - Int(y1)
    xBlend = x1 - Int(x1)
    xBlendInv = 1# - xBlend
    
    'Blend in the x-direction
    Dim topRowColor As Double, bottomRowColor As Double
    topRowColor = topRight * xBlend + topLeft * xBlendInv
    bottomRowColor = bottomRight * xBlend + bottomLeft * xBlendInv
    
    'Blend in the y-direction
    PInterpolate = bottomRowColor * yBlend + topRowColor * (1# - yBlend)

End Function

'This function takes an x and y value - as floating-point - and uses their position to calculate an interpolated value
' for an imaginary pixel in that location.  Results are returned as an RGBQuad.
Private Function PInterpolate_Fast(ByVal x1 As Double, ByVal y1 As Double) As RGBQuad
    
    'We first need to retrieve the four surrounding pixel values
    Dim topLeft As RGBQuad, topRight As RGBQuad, bottomLeft As RGBQuad, bottomRight As RGBQuad
    
    'Pixel coordinates are always guaranteed in-bounds on the left/top boundary
    topLeft = m_Pixels(Int(x1), Int(y1))
    
    'The other three pixels can potentially lie out-of-bounds, so we need to bounds-check 'em
    Dim fixX As Double, fixY As Double
    
    'Pixels at the far edges of the image require special treatment during interpolation
    ' (treatment that varies depending on the current edge-handling model).
    
    'Top-right next
    If (x1 < m_FinalX) Then
    
        topRight = m_Pixels(Int(x1 + 1#), Int(y1))
        
        'We know the x-coord is in-bounds; perform a quick check on y to see if we can short-circuit
        ' the rest of the function.  (If x is in-bounds, the odds of y also being in-bounds is very high.)
        If (y1 < m_FinalY) Then
            bottomLeft = m_Pixels(Int(x1), Int(y1 + 1#))
            bottomRight = m_Pixels(Int(x1 + 1#), Int(y1 + 1#))
            GoTo ShortCircuit
        End If
        
    Else
        fixX = x1 + 1#
        fixY = y1
        If FixDistortEdges(fixX, fixY) Then PutMem4 VarPtr(topRight), 0& Else topRight = m_Pixels(Int(fixX), Int(fixY))
    End If
    
    If (y1 < m_FinalY) Then
        bottomLeft = m_Pixels(Int(x1), Int(y1 + 1#))
    Else
        fixX = x1
        fixY = y1 + 1#
        If FixDistortEdges(fixX, fixY) Then PutMem4 VarPtr(bottomLeft), 0& Else bottomLeft = m_Pixels(Int(fixX), Int(fixY))
    End If
    
    If (x1 < m_FinalX) And (y1 < m_FinalY) Then
        bottomRight = m_Pixels(Int(x1 + 1#), Int(y1 + 1#))
    Else
        fixX = x1 + 1#
        fixY = y1 + 1#
        If FixDistortEdges(fixX, fixY) Then
            PutMem4 VarPtr(bottomRight), 0&
        Else
            bottomRight = m_Pixels(Int(fixX), Int(fixY))
        End If
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
    PInterpolate_Fast.Blue = bottomRow * yBlend + topRow * yBlendInv
    
    topRow = topRight.Green * xBlend + topLeft.Green * xBlendInv
    bottomRow = bottomRight.Green * xBlend + bottomLeft.Green * xBlendInv
    PInterpolate_Fast.Green = bottomRow * yBlend + topRow * yBlendInv
    
    topRow = topRight.Red * xBlend + topLeft.Red * xBlendInv
    bottomRow = bottomRight.Red * xBlend + bottomLeft.Red * xBlendInv
    PInterpolate_Fast.Red = bottomRow * yBlend + topRow * yBlendInv
    
    topRow = topRight.Alpha * xBlend + topLeft.Alpha * xBlendInv
    bottomRow = bottomRight.Alpha * xBlend + bottomLeft.Alpha * xBlendInv
    PInterpolate_Fast.Alpha = bottomRow * yBlend + topRow * yBlendInv
    
End Function
