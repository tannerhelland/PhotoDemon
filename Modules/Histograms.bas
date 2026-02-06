Attribute VB_Name = "Histograms"
'***************************************************************************
'Histogram Analysis tools
'Copyright 2000-2026 by Tanner Helland
'Created: 13/October/00
'Last updated: 07/September/15
'Last update: start collecting all of PD's various histogram-specific tools into a single location.
'
'Histograms pop up a lot in an image editor like PD, so rather than re-implement the same functions dozens of times,
' I've tried to condense key histogram functionality into this single module.
'
'Note that some UI code pops up here as well, as various PD tools provide a histogram overlay.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
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
        lumLookup(x) = Int(CDbl(x) / 3# + 0.5)
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
Public Sub GenerateHistogramImages(ByRef histogramData() As Long, ByRef channelMax() As Long, ByRef dstDIBs() As pdDIB, ByVal imgWidth As Long, ByVal imgHeight As Long, Optional ByVal paintBorder As Boolean = False)
    
    'The incoming histogramData() and channelMax() arrays are already filled, and must not be modified.
    
    'The DIBs, however, are fully under our control
    ReDim dstDIBs(0 To 3) As pdDIB
    
    Dim tmpImage As pdDIB
    Dim cCompositor As pdCompositor
    Set cCompositor = New pdCompositor
    
    'Width/height padding for the histogram image itself
    Const HIST_WIDTH_PADDING As Single = 2!
    Const HIST_HEIGHT_PADDING As Single = 8!
    
    'tHeight is used to determine the height of the maximum value in the histogram.  We want it to be slightly
    ' shorter than the height of the picture box; this way the tallest histogram value fills the entire box
    Dim dstWidth As Single, dstHeight As Single
    dstWidth = imgWidth - HIST_WIDTH_PADDING * 2
    dstHeight = imgHeight - HIST_HEIGHT_PADDING * 2
    
    'pd2D will be used for rendering, so we simply need to construct a polyline for it to draw.
    ' If the user wants us to *fill* the histogram, we will need to add corner points to the
    ' finished line to construct a filled shape - two extra points exist so that the left and right
    ' histogram points extend to the edge of the image (so 255 + 2), plus another 2 points for the
    ' bottom two corners (255 + 2 + 2.)
    Dim listOfPoints() As PointFloat
    ReDim listOfPoints(0 To 259) As PointFloat
    
    Dim i As Long, j As Long
    Dim curChannelMax As Long, targetColor As Long
    
    'Build a look-up table of x-positions for the histogram data; these are equally distributed across the width of
    ' the target image (with a little room left for padding).
    Dim hLookupX() As Double
    ReDim hLookupX(0 To 255) As Double
    For j = 0 To 255
        hLookupX(j) = (CSng(j + 1) / 257#) * CSng(imgWidth)
    Next j
    
    Dim cSurface As pd2DSurface, cPen As pd2DPen, cBrush As pd2DBrush
    
    'Find the max of all channels
    curChannelMax = PDMath.Max3Int(channelMax(0), channelMax(1), channelMax(2))
    If (curChannelMax = 0) Then curChannelMax = 1
    
    For i = 0 To 3
        
        'Initialize this channel's DIB
        Set dstDIBs(i) = New pdDIB
        dstDIBs(i).CreateBlank imgWidth, imgHeight, 32, 0, 0
        
        'Individual color channels are handled differently from the "merged" RGB channel
        If (i < 3) Then
            
            'The color of the histogram changes for each channel
            Select Case i
                Case 0
                    targetColor = g_Themer.GetGenericUIColor(UI_ChannelRed)
                Case 1
                    targetColor = g_Themer.GetGenericUIColor(UI_ChannelGreen)
                Case 2
                    targetColor = g_Themer.GetGenericUIColor(UI_ChannelBlue)
            End Select
            
            'Iterate through the histogram and construct a matching on-screen point for each value
            For j = 0 To 255
                listOfPoints(j + 1).x = HIST_WIDTH_PADDING + (CSng(j) * dstWidth) / 255!
                listOfPoints(j + 1).y = HIST_HEIGHT_PADDING + (dstHeight - (histogramData(i, j) * dstHeight) / curChannelMax)
            Next j
            
            'Manually populate the first and last points
            listOfPoints(0).x = 0!
            listOfPoints(0).y = listOfPoints(1).y
            listOfPoints(257).x = imgWidth
            listOfPoints(257).y = listOfPoints(256).y
            
            'Apply gentle smoothing to the line to improve its visual appearance
            Dim numOfPoints As Long
            numOfPoints = 257
            PDMath.SmoothLineY listOfPoints, numOfPoints, 0.5
            PDMath.SimplifyLine listOfPoints, numOfPoints, 0.25
            
            'Re-fill the first and last points to ensure the histogram is filled correctly.
            listOfPoints(0).x = 0!
            listOfPoints(0).y = listOfPoints(1).y
            listOfPoints(numOfPoints).x = imgWidth
            listOfPoints(numOfPoints).y = listOfPoints(numOfPoints - 1).y
            
            'Also fill in the end points of the polyline, so we can treat it as a polygon
            listOfPoints(numOfPoints + 1).x = imgWidth + 1
            listOfPoints(numOfPoints + 1).y = imgHeight * 2
            listOfPoints(numOfPoints + 2).x = -1!
            listOfPoints(numOfPoints + 2).y = imgHeight
            
            numOfPoints = numOfPoints + 3
            
            'Assemble a drawing surface
            Drawing2D.QuickCreateSurfaceFromDIB cSurface, dstDIBs(i), True
            cSurface.SetSurfacePixelOffset P2_PO_Half
            
            'Construct a matching fill brush, then fill the histogram region
            Drawing2D.QuickCreateSolidBrush cBrush, targetColor, 15!
            PD2D.FillPolygonF_FromPtF cSurface, cBrush, numOfPoints, VarPtr(listOfPoints(0)), True, 0.25
            
            'Next, stroke the outline, then free all rendering objects
            Drawing2D.QuickCreateSolidPen cPen, 1!, targetColor, 100!, P2_LJ_Round, P2_LC_Round
            PD2D.DrawLinesF_FromPtF cSurface, cPen, numOfPoints, VarPtr(listOfPoints(0)), True, 0.25
            cSurface.ReleaseSurface
            
            'Mark the DIB's alpha state
            dstDIBs(i).SetInitialAlphaPremultiplicationState True
        
        'For the "merged" RGB image, we want to merge all three channel DIBs together,
        ' but we need to rebuild them against the max of *all* channels
        Else
            
            'Prepare a "temporary" DIB to receive the merged image
            Set tmpImage = New pdDIB
            tmpImage.CreateBlank imgWidth, imgHeight, 32, 0, 0
            tmpImage.SetInitialAlphaPremultiplicationState True
            
            'Merge all previous images onto it
            Dim targetBM As PD_BlendMode
            targetBM = BM_Overlay
            
            cCompositor.QuickMergeTwoDibsOfEqualSize tmpImage, dstDIBs(0), targetBM
            cCompositor.QuickMergeTwoDibsOfEqualSize tmpImage, dstDIBs(1), targetBM
            cCompositor.QuickMergeTwoDibsOfEqualSize tmpImage, dstDIBs(2), targetBM
            
            'Fill the target DIB with the current background UI color, then blend the
            ' final result onto it
            dstDIBs(i).FillWithColor g_Themer.GetGenericUIColor(UI_Background), 100!
            tmpImage.AlphaBlendToDC dstDIBs(i).GetDIBDC
            
        End If
        
    Next i
    
    'Finalize the individual channel DIBs by merging them onto the current theme background color
    For i = 0 To 2
        tmpImage.FillWithColor g_Themer.GetGenericUIColor(UI_Background), 100!
        dstDIBs(i).AlphaBlendToDC tmpImage.GetDIBDC
        dstDIBs(i).CreateFromExistingDIB tmpImage
    Next i
    
    'If the caller wants us to paint a border, do that last
    If paintBorder Then
        
        cPen.SetPenWidth 1!
        cPen.SetPenColor g_Themer.GetGenericUIColor(UI_GrayNeutral)
        cPen.SetPenLineJoin P2_LJ_Miter
        
        For i = 0 To 3
            cSurface.WrapSurfaceAroundPDDIB dstDIBs(i)
            cSurface.SetSurfaceAntialiasing P2_AA_None
            cSurface.SetSurfacePixelOffset P2_PO_Normal
            PD2D.DrawRectangleI cSurface, cPen, 0, 0, imgWidth - 1, imgHeight - 1
        Next i
    
    End If
    
End Sub

'Stretch the histogram to reach from 0 to 255 (white balance correction is a far better method, FYI)
Public Sub StretchHistogram()
   
    Message "Analyzing image..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte, tmpSA As SafeArray2D, tmpSA1D As SafeArray1D
    EffectPrep.PrepImageData tmpSA, False
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    Dim xStride As Long
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    ProgressBars.SetProgBarMax finalY
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Color variables
    Dim r As Long, g As Long, b As Long
    
    'Max and min values
    Dim rMax As Long, gMax As Long, bMax As Long
    Dim rMin As Long, gMin As Long, bMin As Long
    rMin = 255
    gMin = 255
    bMin = 255
        
    'Loop through each pixel in the image, checking max/min values as we go
    For y = initY To finalY
        workingDIB.WrapArrayAroundScanline imageData, tmpSA1D, y
    For x = initX To finalX
        
        'Get the source pixel color values
        xStride = x * 4
        b = imageData(xStride)
        g = imageData(xStride + 1)
        r = imageData(xStride + 2)
        
        'Find max/min values in the image
        If (r < rMin) Then rMin = r
        If (r > rMax) Then rMax = r
        If (g < gMin) Then gMin = g
        If (g > gMax) Then gMax = g
        If (b < bMin) Then bMin = b
        If (b > bMax) Then bMax = b
        
    Next x
    Next y
    
    Message "Stretching histogram..."
    Dim rDif As Long, gDif As Long, bDif As Long
    
    rDif = rMax - rMin
    gDif = gMax - gMin
    bDif = bMax - bMin
    
    'Lookup tables make the stretching go faster
    Dim rLookup(0 To 255) As Byte, gLookup(0 To 255) As Byte, bLookup(0 To 255) As Byte
    
    For x = 0 To 255
        
        If (rDif <> 0) Then
            r = 255 * ((x - rMin) / rDif)
            If (r < 0) Then r = 0
            If (r > 255) Then r = 255
            rLookup(x) = r
        Else
            rLookup(x) = x
        End If
        
        If (gDif <> 0) Then
            g = 255 * ((x - gMin) / gDif)
            If (g < 0) Then g = 0
            If (g > 255) Then g = 255
            gLookup(x) = g
        Else
            gLookup(x) = x
        End If
        
        If (bDif <> 0) Then
            b = 255 * ((x - bMin) / bDif)
            If (b < 0) Then b = 0
            If (b > 255) Then b = 255
            bLookup(x) = b
        Else
            bLookup(x) = x
        End If
        
    Next x
    
    'Loop through each pixel in the image, converting values as we go
    For y = initY To finalY
        workingDIB.WrapArrayAroundScanline imageData, tmpSA1D, y
    For x = initX To finalX
        
        'Get the source pixel color values
        xStride = x * 4
        b = imageData(xStride)
        g = imageData(xStride + 1)
        r = imageData(xStride + 2)
        
        imageData(xStride) = bLookup(b)
        imageData(xStride + 1) = gLookup(g)
        imageData(xStride + 2) = rLookup(r)
        
    Next x
        If (y And progBarCheck) = 0 Then ProgressBars.SetProgBarVal y
    Next y
    
    'Safely deallocate imageData()
    workingDIB.UnwrapArrayFromDIB imageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    EffectPrep.FinalizeImageData
        
End Sub
