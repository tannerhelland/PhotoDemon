Attribute VB_Name = "Histogram_Analysis"
'***************************************************************************
'Histogram Analysis tools
'Copyright 2000-2015 by Tanner Helland
'Created: 13/October/00
'Last updated: 07/September/15
'Last update: start collecting all of PD's various histogram-specific tools into a single location.
'
'Histograms pop up a lot in an image editor like PD, so rather than re-implement the same functions dozens of times,
' I've tried to condense key histogram functionality into this single module.
'
'Note that some UI code pops up here as well, as various PD tools provide a histogram overlay.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Fill the supplied arrays with histogram data for the current image
' In order, the arrays that need to be supplied are:
' 1) array for histogram data (dimensioned [0,3][0,255] - the first ordinal specifies channel, in RGB[L] order)
' 2) array for logarithmic histogram data (dimensioned same as hData)
' 3) Array for max channel values (dimensioned [0,3])
' 4) Array for max log channel values
' 5) Array of where the maximum channel values occur (histogram index)
Public Sub fillHistogramArrays(ByRef hData() As Double, ByRef hDataLog() As Double, ByRef channelMax() As Double, ByRef channelMaxLog() As Double, ByRef channelMaxPosition() As Byte)
    
    'Redimension the various arrays
    ReDim hData(0 To 3, 0 To 255) As Double
    ReDim hDataLog(0 To 3, 0 To 255) As Double
    ReDim channelMax(0 To 3) As Double
    ReDim channelMaxLog(0 To 3) As Double
    ReDim channelMaxPosition(0 To 3) As Byte
    
    'Create a local array and point it at the pixel data we want to scan
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    
    prepImageData tmpSA, , , , True
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'These variables will hold temporary histogram values
    Dim r As Long, g As Long, b As Long, l As Long
    
    'If the histogram has already been used, we need to clear out all the
    'maximum values and histogram values
    Dim hMax As Double, hMaxLog As Double
    hMax = 0:    hMaxLog = 0
    
    For x = 0 To 3
        channelMax(x) = 0
        channelMaxLog(x) = 0
        For y = 0 To 255
            hData(x, y) = 0
        Next y
    Next x
    
    'Build a look-up table for luminance conversion; 765 = 255 * 3
    Dim lumLookup(0 To 765) As Byte
    
    For x = 0 To 765
        lumLookup(x) = x \ 3
    Next x
    
    'Run a quick loop through the image, gathering what we need to calculate our histogram
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        'We have to gather the red, green, and blue in order to calculate luminance
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Rather than generate authentic luminance (which requires a costly HSL conversion routine), we use a simpler average value.
        l = lumLookup(r + g + b)
        
        'Increment each value in the array, depending on its present value; this will let us see how many pixels of
        ' each color value (and luminance value) there are in the image
        
        'Red
        hData(0, r) = hData(0, r) + 1
        'Green
        hData(1, g) = hData(1, g) + 1
        'Blue
        hData(2, b) = hData(2, b) + 1
        'Luminance
        hData(3, l) = hData(3, l) + 1
        
    Next y
    Next x
    
    'With our dataset successfully collected, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Run a quick loop through the completed array to find maximum values
    For x = 0 To 3
        For y = 0 To 255
            If hData(x, y) > channelMax(x) Then
                channelMax(x) = hData(x, y)
                channelMaxPosition(x) = y
            End If
        Next y
    Next x
    
    'Now calculate the logarithmic version of the histogram
    For x = 0 To 3
        If channelMax(x) <> 0 Then channelMaxLog(x) = Log(channelMax(x)) Else channelMaxLog(x) = 0
    Next x
    
    For x = 0 To 3
        For y = 0 To 255
            If hData(x, y) <> 0 Then
                hDataLog(x, y) = Log(hData(x, y))
            Else
                hDataLog(x, y) = 0
            End If
        Next y
    Next x
    
End Sub

'Given a set of histogram arrays generated by fillHistogramArrays(), above, produce a set of matching pdDIB objects.
' For consistency reasons, the caller doesn't get much control over these images; just width/height, and this function
' controls the rest (including color + transparency decisions).  The finished images are 32-bpp and suitable for layering,
' as well.
'
'Note: this function only takes one set of input histogram data, so if you want images for both log and non-log variants,
' you'll need to call this function *twice*, once for each set.
Public Sub generateHistogramImages(ByRef histogramData() As Double, ByRef channelMax() As Double, ByRef dstDIBs() As pdDIB, ByVal imgWidth As Long, ByVal imgHeight As Long)
    
    'The incoming histogramData() and channelMax() arrays are already filled, and must not be modified.
    
    'The DIBs, however, are fully under our control
    ReDim dstDIBs(0 To 3) As pdDIB
    
    Dim tmpPath As pdGraphicsPath, histogramShape() As POINTFLOAT
    Dim tmpPen As Long, tmpBrush As Long, hColor As Long
    Dim i As Long, j As Long
    Dim yMax As Double
    
    'Build a look-up table of x-positions for the histogram data; these are equally distributed across the width of
    ' the target image (with a little room left for padding).
    Dim hLookupX() As Double
    ReDim hLookupX(0 To 255) As Double
    For j = 0 To 255
        hLookupX(j) = (CSng(j + 1) / 257) * CSng(imgWidth)
    Next j
    
    For i = 0 To 3
        
        'Initialize this channel's DIB
        Set dstDIBs(i) = New pdDIB
        dstDIBs(i).createBlank imgWidth, imgHeight, 32, vbBlack
        
        yMax = 0.9 * imgHeight
        
        'The color of the histogram changes for each channel
        Select Case i
        
            'Red
            Case 0
                hColor = RGB(255, 60, 80)
            
            'Green
            Case 1
                hColor = RGB(60, 210, 80)
            
            'Blue
            Case 2
                hColor = RGB(60, 100, 255)
            
            'Luminance
            Case 3
                hColor = RGB(66, 74, 74)
        
        
        End Select
                
        'New strategy!  Use the awesome pdGraphicsPath class to construct a matching polygon for each histogram.
        ' Then, stroke and fill the polygon in one fell swoop (much faster).
        ReDim histogramShape(0 To 260) As POINTFLOAT
        
        For j = 0 To 255
            histogramShape(j).x = hLookupX(j)
            histogramShape(j).y = imgHeight - (histogramData(i, j) / channelMax(i)) * yMax
        Next j
        
        'Complete each shape by tracing the outline of the DIB
        histogramShape(256).x = imgWidth + 1
        histogramShape(256).y = imgHeight
        histogramShape(257).x = imgWidth
        histogramShape(257).y = imgHeight + 1
        
        histogramShape(258).x = 0
        histogramShape(258).y = imgHeight
        histogramShape(259).x = -1
        histogramShape(259).y = imgHeight + 1
        
        'Populate shape objects using those point lists
        Set tmpPath = New pdGraphicsPath
        tmpPath.addPolygon 260, VarPtr(histogramShape(0)), True, True
        
        'Prep pens and brushes in the current color
        tmpPen = GDI_Plus.getGDIPlusPenHandle(hColor, 255, 1, LineCapRound, LineJoinRound)
        tmpBrush = GDI_Plus.getGDIPlusSolidBrushHandle(hColor, 64)
        
        'Render the paths to their target DIBs
        tmpPath.fillPathToDIB_BareBrush tmpBrush, dstDIBs(i)
        tmpPath.strokePathToDIB_BarePen tmpPen, dstDIBs(i)
        
        'Free our pen and brush resources
        GDI_Plus.releaseGDIPlusPen tmpPen
        GDI_Plus.releaseGDIPlusBrush tmpBrush
        
        'Mark alpha premultiplication status
        dstDIBs(i).setInitialAlphaPremultiplicationState True
        
    Next i
    
End Sub
