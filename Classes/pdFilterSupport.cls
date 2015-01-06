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
'Copyright 2013-2015 by Tanner Helland
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
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private m_FinalX As Long, m_FinalY As Long
Private m_EdgeMethod As Long, m_ColorDepth As Long, m_ColorDepthLoop As Long
Private m_Interpolate As Boolean
Private m_DstX As Long, m_DstY As Long

'A function may want to force parts of an image to a certain color.  This simplified function can be used to do that.
Public Sub forcePixels(ByRef x As Long, ByRef y As Long, ByRef newR As Long, ByRef newG As Long, ByRef newB As Long, ByRef newA As Long, ByRef dstData() As Byte)

    Dim QuickVal As Long
    QuickVal = x * m_ColorDepth
    dstData(QuickVal, y) = newB
    dstData(QuickVal + 1, y) = newG
    dstData(QuickVal + 2, y) = newR
    If m_ColorDepth = 4 Then dstData(QuickVal + 3, y) = newA

End Sub

'If a function simply wants RGB/A values returned for a given source position, then can use this function to capture
' those values.  This is helpful for tools that calculate multiple source positions (like a Blur tool), and need to
' do something with the colors from those various positions.
'
'Said another way, the caller need only calculate srcX and srcY, then this class will do the rest!
Public Sub getColorsFromSource(ByRef r As Long, ByRef g As Long, ByRef b As Long, ByRef a As Long, ByRef srcX As Double, ByRef srcY As Double, ByRef srcData() As Byte, Optional ByRef origX As Long, Optional ByRef origY As Long)
    
    'Cache the original x/y values as necessary
    If m_EdgeMethod = EDGE_ORIGINAL Then
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
        If m_ColorDepth = 4 Then a = 0
    Else
    
        'Interpolate a new pixel value
        If m_Interpolate Then
            
            r = pInterpolate(srcX, srcY, srcData, 2)
            g = pInterpolate(srcX, srcY, srcData, 1)
            b = pInterpolate(srcX, srcY, srcData, 0)
            If m_ColorDepth = 4 Then a = pInterpolate(srcX, srcY, srcData, 3)
        
        'Round to the nearest coordinate value
        Else
    
            quickValX = Int(srcX) * m_ColorDepth
            quickValY = Int(srcY)
            
            r = srcData(quickValX + 2, quickValY)
            g = srcData(quickValX + 1, quickValY)
            b = srcData(quickValX, quickValY)
            If m_ColorDepth = 4 Then a = srcData(quickValX + 3, quickValY)
            
        End If
        
    End If

End Sub

'After a transform has generate a set of reverse-mapped pixel coordinates, they are passed to this function.
' Based on already established parameters (such as interpolation, edge handling, etc), new pixel values are set.
Public Sub setPixels(ByRef x As Long, ByRef y As Long, ByRef srcX As Double, ByRef srcY As Double, ByRef srcData() As Byte, ByRef dstData() As Byte)

    If m_EdgeMethod = EDGE_ORIGINAL Then
        m_DstX = x
        m_DstY = y
    End If

    Dim fixDistort As Boolean

    'First, fix any pixels that lie outside the image
    fixDistort = FixDistortEdges(srcX, srcY)
    
    'Now, interpolate as necessary
    Dim i As Long, QuickVal As Long, QuickVal2 As Long
    
    QuickVal = x * m_ColorDepth
    
    'fixDistort will only be set to TRUE when the current pixel needs to be erased
    If fixDistort Then
        For i = 0 To m_ColorDepthLoop
            dstData(QuickVal + i, y) = 0
        Next i
    Else
        If m_Interpolate Then
            For i = 0 To m_ColorDepthLoop
                dstData(QuickVal + i, y) = pInterpolate(srcX, srcY, srcData, i)
            Next i
        Else
    
            QuickVal2 = Int(srcX) * m_ColorDepth
        
            For i = 0 To m_ColorDepthLoop
                dstData(QuickVal + i, y) = srcData(QuickVal2 + i, Int(srcY))
            Next i
        End If
    End If

End Sub

'When this support class is first initialized by a function, a number of parameters are specified.  It is much faster
' to set these once and store them instead of passing them every time a pixel needs to be handled.
Public Sub setDistortParameters(ByVal imgColorDepth As Long, ByVal edgeMethod As EDGE_OPERATOR, ByVal toInterpolate As Boolean, ByVal finalX As Long, ByVal finalY As Long)
    
    m_ColorDepth = imgColorDepth
    m_ColorDepthLoop = m_ColorDepth - 1
    m_EdgeMethod = edgeMethod
    m_Interpolate = toInterpolate
    m_FinalX = finalX
    m_FinalY = finalY
    
End Sub

'If a pixel lies outside image boundaries, move it in-bounds using one of several methods
' If the edge handle method is "Erase", this function will return a boolean indicating whether the supplied pixel
' must be erased.  If FALSE is returned, the pixel can be handled normally.
Public Function FixDistortEdges(ByRef srcX As Double, ByRef srcY As Double) As Boolean

    Select Case m_EdgeMethod
    
        'Clamp
        Case EDGE_CLAMP
        
            If srcX < 0 Then srcX = 0
            If srcY < 0 Then srcY = 0
            If srcX > m_FinalX Then srcX = m_FinalX
            If srcY > m_FinalY Then srcY = m_FinalY
        
        'Reflect
        Case EDGE_REFLECT
        
            If srcX < 0 Then srcX = Abs(srcX)
            If srcY < 0 Then srcY = Abs(srcY)
            If srcX > m_FinalX Then srcX = m_FinalX - (srcX - m_FinalX)
            If srcY > m_FinalY Then srcY = m_FinalY - (srcY - m_FinalY)
        
            'If the modified pixel STILL lies outside the image, use modulo to move it in-bounds
            If srcX < 0 Then srcX = Modulo(srcX, (m_FinalX + 1))
            If srcY < 0 Then srcY = Modulo(srcY, (m_FinalY + 1))
            If srcX > m_FinalX Then srcX = Modulo(srcX, (m_FinalX + 1))
            If srcY > m_FinalY Then srcY = Modulo(srcY, (m_FinalY + 1))
        
        'Wrap
        Case EDGE_WRAP
        
            If srcX < 0 Then srcX = Modulo(srcX, (m_FinalX + 1))
            If srcY < 0 Then srcY = Modulo(srcY, (m_FinalY + 1))
            If srcX > m_FinalX Then srcX = Modulo(srcX, (m_FinalX + 1))
            If srcY > m_FinalY Then srcY = Modulo(srcY, (m_FinalY + 1))
        
        'Erase
        Case EDGE_ERASE
        
            If srcX < 0 Then
                FixDistortEdges = True
                Exit Function
            End If
            
            If srcY < 0 Then
                FixDistortEdges = True
                Exit Function
            End If
            
            If srcX > m_FinalX Then
                FixDistortEdges = True
                Exit Function
            End If
            
            If srcY > m_FinalY Then
                FixDistortEdges = True
                Exit Function
            End If
            
        Case EDGE_ORIGINAL
            If srcX < 0 Then
                srcX = m_DstX
                srcY = m_DstY
            End If
            
            If srcY < 0 Then
                srcX = m_DstX
                srcY = m_DstY
            End If
            
            If srcX > m_FinalX Then
                srcX = m_DstX
                srcY = m_DstY
            End If
            
            If srcY > m_FinalY Then
                srcX = m_DstX
                srcY = m_DstY
            End If
            
    End Select
    
    FixDistortEdges = False

End Function

'This function takes an x and y value - as floating-point - and uses their position to calculate an interpolated value
' for an imaginary pixel in that location.  Offset (r/g/b/alpha) and image color depth are also required.
Private Function pInterpolate(ByVal x1 As Double, ByVal y1 As Double, ByRef iData() As Byte, ByRef iOffset As Long) As Byte
        
    'Retrieve the four surrounding pixel values
    Dim topLeft As Double, topRight As Double, bottomLeft As Double, bottomRight As Double
    topLeft = iData(Int(x1) * m_ColorDepth + iOffset, Int(y1))
    Dim fixX As Double, fixY As Double
    
    'Pixels at the far edges of the image require special treatment during interpolation
    If x1 < m_FinalX Then
        topRight = iData(Int(x1 + 1) * m_ColorDepth + iOffset, Int(y1))
    Else
        fixX = x1 + 1
        fixY = y1
        If FixDistortEdges(fixX, fixY) Then
            topRight = 0
        Else
            topRight = iData(Int(fixX) * m_ColorDepth + iOffset, Int(y1))
        End If
    End If
    If y1 < m_FinalY Then
        bottomLeft = iData(Int(x1) * m_ColorDepth + iOffset, Int(y1 + 1))
    Else
        fixX = x1
        fixY = y1 + 1
        If FixDistortEdges(fixX, fixY) Then
            bottomLeft = 0
        Else
            bottomLeft = iData(Int(x1) * m_ColorDepth + iOffset, Int(fixY))
        End If
    End If
    If x1 < m_FinalX Then
        If y1 < m_FinalY Then
            bottomRight = iData(Int(x1 + 1) * m_ColorDepth + iOffset, Int(y1 + 1))
        Else
            fixX = x1 + 1
            fixY = y1 + 1
            If FixDistortEdges(fixX, fixY) Then
                bottomRight = 0
            Else
                bottomRight = iData(Int(x1 + 1) * m_ColorDepth + iOffset, Int(fixY))
            End If
        End If
    Else
        fixX = x1 + 1
        fixY = y1 + 1
        If FixDistortEdges(fixX, fixY) Then
            bottomRight = 0
        Else
            If y1 < m_FinalY Then
                bottomRight = iData(Int(fixX) * m_ColorDepth + iOffset, Int(y1 + 1))
            Else
                bottomRight = iData(Int(fixX) * m_ColorDepth + iOffset, Int(fixY))
            End If
        End If
    End If
    
    'Calculate blend ratios
    Dim yBlend As Double
    Dim xBlend As Double, xBlendInv As Double
    yBlend = y1 - Int(y1)
    xBlend = x1 - Int(x1)
    xBlendInv = 1 - xBlend
    
    'Blend in the x-direction
    Dim topRowColor As Double, bottomRowColor As Double
    topRowColor = topRight * xBlend + topLeft * xBlendInv
    bottomRowColor = bottomRight * xBlend + bottomLeft * xBlendInv
    
    'Blend in the y-direction
    pInterpolate = bottomRowColor * yBlend + topRowColor * (1 - yBlend)

End Function
