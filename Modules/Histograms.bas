Attribute VB_Name = "Histograms"
'***************************************************************************
'Histogram Analysis tools
'Copyright 2000-2019 by Tanner Helland
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
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
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
'
'TODO: add an option for ignoring transparent pixels; this would improve output on images with variable opacity
Public Sub FillHistogramArrays(ByRef hData() As Long, ByRef hDataLog() As Double, ByRef channelMax() As Long, ByRef channelMaxLog() As Double, ByRef channelMaxPosition() As Byte, Optional ByVal allowDownsample As Boolean = False)
    
    'Redimension the various arrays
    ReDim hData(0 To 3, 0 To 255) As Long
    ReDim hDataLog(0 To 3, 0 To 255) As Double
    ReDim channelMax(0 To 3) As Long
    ReDim channelMaxLog(0 To 3) As Double
    ReDim channelMaxPosition(0 To 3) As Byte
    
    'If the image is large, it's faster to grab a downsampled image and simply use that;
    ' however, some dialogs (like Display Histogram) require full image data.
    Dim srcDIB As pdDIB
    Set srcDIB = New pdDIB
    
    If allowDownsample Then
        
        'Downsample to two megapixel as necessary; that's more than enough data for a good histogram
        Const MAX_PX_SIZE As Long = 2000000
        With PDImages.GetActiveImage.GetActiveDIB
            If (.GetDIBWidth * .GetDIBHeight) > MAX_PX_SIZE Then
                DIBs.ResizeDIBByPixelCount PDImages.GetActiveImage.GetActiveDIB, srcDIB, MAX_PX_SIZE, GP_IM_Bilinear
            Else
                srcDIB.CreateFromExistingDIB PDImages.GetActiveImage.GetActiveDIB
            End If
        End With
        
    Else
        srcDIB.CreateFromExistingDIB PDImages.GetActiveImage.GetActiveDIB
    End If
    
    'Create a local array and point it at the pixel data we want to scan
    Dim imageData() As Byte, tmpSA1D As SafeArray1D
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim imgDepth As Long
    imgDepth = srcDIB.GetDIBColorDepth \ 8
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = (srcDIB.GetDIBWidth - 1) * imgDepth
    finalY = srcDIB.GetDIBHeight - 1
    
    'These variables will hold temporary histogram values
    Dim r As Long, g As Long, b As Long, l As Long
    
    'If the histogram has already been used, we need to clear out all the
    'maximum values and histogram values
    Dim hMax As Long, hMaxLog As Double
    hMax = 0
    hMaxLog = 0#
    
    'Build a look-up table for luminance conversion; 765 = 255 * 3
    Dim lumLookup(0 To 765) As Byte
    
    For x = 0 To 765
        lumLookup(x) = x \ 3
    Next x
    
    Dim dibPtr As Long, dibStride As Long
    srcDIB.WrapArrayAroundScanline imageData, tmpSA1D, 0
    dibPtr = srcDIB.GetDIBPointer
    dibStride = srcDIB.GetDIBStride
    
    'Run a quick loop through the image, gathering what we need to calculate our histogram
    For y = initY To finalY
        tmpSA1D.pvData = dibPtr + dibStride * y
    For x = initX To finalX Step imgDepth
        
        'Gather RGB and calculate luminance
        b = imageData(x)
        g = imageData(x + 1)
        r = imageData(x + 2)
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
        
    Next x
    Next y
    
    'With our dataset successfully collected, point ImageData() away from the DIB and deallocate it
    srcDIB.UnwrapArrayFromDIB imageData
    
    'Run a quick loop through the completed array to find maximum values
    For x = 0 To 3
        For y = 0 To 255
            If (hData(x, y) > channelMax(x)) Then
                channelMax(x) = hData(x, y)
                channelMaxPosition(x) = y
            End If
        Next y
    Next x
    
    'Now calculate the logarithmic version of the histogram
    For x = 0 To 3
        If (channelMax(x) <> 0) Then channelMaxLog(x) = Log(channelMax(x)) Else channelMaxLog(x) = 0#
    Next x
    
    For x = 0 To 3
        For y = 0 To 255
            If (hData(x, y) <> 0) Then
                hDataLog(x, y) = Log(hData(x, y))
            Else
                hDataLog(x, y) = 0#
            End If
        Next y
    Next x
    
End Sub

'Given a set of histogram arrays generated by FillHistogramArrays(), above, produce a set of matching pdDIB objects.
' For consistency reasons, the caller doesn't get much control over these images; just width/height, and this function
' controls the rest (including color + transparency decisions).  The finished images are 32-bpp and suitable for layering,
' as well.
'
'Note: this function only takes one set of input histogram data, so if you want images for both log and non-log variants,
' you'll need to call this function *twice*, once for each set.
Public Sub GenerateHistogramImages(ByRef histogramData() As Long, ByRef channelMax() As Long, ByRef dstDIBs() As pdDIB, ByVal imgWidth As Long, ByVal imgHeight As Long)
    
    'The incoming histogramData() and channelMax() arrays are already filled, and must not be modified.
    
    'The DIBs, however, are fully under our control
    ReDim dstDIBs(0 To 3) As pdDIB
    
    Dim tmpPath As pd2DPath, histogramShape() As PointFloat
    Dim hColor As Long
    Dim i As Long, j As Long
    Dim yMax As Double
    
    'Build a look-up table of x-positions for the histogram data; these are equally distributed across the width of
    ' the target image (with a little room left for padding).
    Dim hLookupX() As Double
    ReDim hLookupX(0 To 255) As Double
    For j = 0 To 255
        hLookupX(j) = (CSng(j + 1) / 257#) * CSng(imgWidth)
    Next j
    
    Dim cSurface As pd2DSurface, cPen As pd2DPen, cBrush As pd2DBrush
    
    For i = 0 To 3
        
        'Initialize this channel's DIB
        Set dstDIBs(i) = New pdDIB
        dstDIBs(i).CreateBlank imgWidth, imgHeight, 32, vbBlack
        
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
                
        'New strategy!  Use the awesome pd2DPath class to construct a matching polygon for each histogram.
        ' Then, stroke and fill the polygon in one fell swoop (much faster).
        ReDim histogramShape(0 To 260) As PointFloat
        
        For j = 0 To 255
            histogramShape(j).x = hLookupX(j)
            If (channelMax(i) > 0) Then
                histogramShape(j).y = imgHeight - (histogramData(i, j) / channelMax(i)) * yMax
            Else
                histogramShape(j).y = imgHeight - yMax
            End If
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
        Set tmpPath = New pd2DPath
        tmpPath.AddPolygon 260, VarPtr(histogramShape(0)), True, True
        
        'Prep pens and brushes in the current color
        Drawing2D.QuickCreateSolidPen cPen, 1#, hColor, 100#, P2_LJ_Round, P2_LC_Round
        Drawing2D.QuickCreateSolidBrush cBrush, hColor, 25#
        
        'Render the paths to their target DIBs
        Drawing2D.QuickCreateSurfaceFromDC cSurface, dstDIBs(i).GetDIBDC, True
        PD2D.FillPath cSurface, cBrush, tmpPath
        PD2D.DrawPath cSurface, cPen, tmpPath
        
        'Mark alpha premultiplication status
        dstDIBs(i).SetInitialAlphaPremultiplicationState True
        
    Next i
    
    'All pd2D paint objects are self-freeing
    
End Sub

'Stretch the histogram to reach from 0 to 255 (white balance correction is a far better method, FYI)
Public Sub StretchHistogram()
   
    Message "Analyzing image histogram for maximum and minimum values..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte
    Dim tmpSA As SafeArray2D
    
    EffectPrep.PrepImageData tmpSA
    CopyMemory ByVal VarPtrArray(imageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim quickVal As Long, qvDepth As Long
    qvDepth = curDIBValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    
    'Max and min values
    Dim rMax As Long, gMax As Long, bMax As Long
    Dim RMin As Long, gMin As Long, bMin As Long
    RMin = 255
    gMin = 255
    bMin = 255
        
    'Loop through each pixel in the image, checking max/min values as we go
    For x = initX To finalX
        quickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        r = imageData(quickVal + 2, y)
        g = imageData(quickVal + 1, y)
        b = imageData(quickVal, y)
        
        If r < RMin Then RMin = r
        If r > rMax Then rMax = r
        If g < gMin Then gMin = g
        If g > gMax Then gMax = g
        If b < bMin Then bMin = b
        If b > bMax Then bMax = b
        
    Next y
    Next x
    
    Message "Stretching histogram..."
    Dim rDif As Long, gDif As Long, bDif As Long
    
    rDif = rMax - RMin
    gDif = gMax - gMin
    bDif = bMax - bMin
    
    'Lookup tables make the stretching go faster
    Dim rLookup(0 To 255) As Byte, gLookup(0 To 255) As Byte, bLookup(0 To 255) As Byte
    
    For x = 0 To 255
        If rDif <> 0 Then
            r = 255 * ((x - RMin) / rDif)
            If r < 0 Then r = 0
            If r > 255 Then r = 255
            rLookup(x) = r
        Else
            rLookup(x) = x
        End If
        If gDif <> 0 Then
            g = 255 * ((x - gMin) / gDif)
            If g < 0 Then g = 0
            If g > 255 Then g = 255
            gLookup(x) = g
        Else
            gLookup(x) = x
        End If
        If bDif <> 0 Then
            b = 255 * ((x - bMin) / bDif)
            If b < 0 Then b = 0
            If b > 255 Then b = 255
            bLookup(x) = b
        Else
            bLookup(x) = x
        End If
    Next x
    
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        quickVal = x * qvDepth
    For y = initY To finalY
    
        'Get the source pixel color values
        r = imageData(quickVal + 2, y)
        g = imageData(quickVal + 1, y)
        b = imageData(quickVal, y)
                
        imageData(quickVal + 2, y) = rLookup(r)
        imageData(quickVal + 1, y) = gLookup(g)
        imageData(quickVal, y) = bLookup(b)
        
    Next y
        If (x And progBarCheck) = 0 Then SetProgBarVal x
    Next x
    
    'Safely deallocate imageData()
    CopyMemory ByVal VarPtrArray(imageData), 0&, 4
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData
        
End Sub

