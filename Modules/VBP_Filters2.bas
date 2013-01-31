Attribute VB_Name = "Filters_Miscellaneous"
'***************************************************************************
'Filter Module
'Copyright ©2000-2013 by Tanner Helland
'Created: 13/October/00
'Last updated: 05/September/12
'Last update: rewrote all code against the new pdLayer class.
'
'The general image filter module; contains unorganized routines at present.
'
'***************************************************************************

Option Explicit

'Given two layers, fill one with a gaussian-blur version of the other.
Public Sub CreateGaussianBlurLayer(ByVal gRadius As Long, ByRef srcLayer As pdLayer, ByRef dstLayer As pdLayer, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0)
            
    'Create a local array and point it at the pixel data of the destination image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepSafeArray dstSA, dstLayer
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Do the same for the source image
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepSafeArray srcSA, srcLayer
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    'Create one more local array.  This will contain the intermediate copy of the gaussian blur, as it must be done in two passes.
    Dim gaussLayer As pdLayer
    Set gaussLayer = New pdLayer
    gaussLayer.createFromExistingLayer srcLayer
    
    Dim GaussImageData() As Byte
    Dim gaussSA As SAFEARRAY2D
    prepSafeArray gaussSA, gaussLayer
    CopyMemory ByVal VarPtrArray(GaussImageData()), VarPtr(gaussSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim X As Long, Y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcLayer.getLayerWidth - 1
    finalY = srcLayer.getLayerHeight - 1
    
    'Make sure we were passed a valid radius
    If gRadius < 1 Then gRadius = 1
    If finalX > finalY Then
        If gRadius > finalX Then gRadius = finalX
    Else
        If gRadius > finalY Then gRadius = finalY
    End If
        
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, QuickValInner As Long, qvDepth As Long
    qvDepth = srcLayer.getLayerColorDepth \ 8
    
    Dim chkAlpha As Boolean
    If qvDepth = 4 Then chkAlpha = True Else chkAlpha = False
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If modifyProgBarMax = -1 Then
        SetProgBarMax finalY + finalY
    Else
        SetProgBarMax modifyProgBarMax
    End If
    progBarCheck = findBestProgBarValue()
    
    'Create a one-dimensional Gaussian kernel using the requested radius
    Dim gKernel() As Single
    ReDim gKernel(-gRadius To gRadius) As Single
    
    Dim numPixels As Long
    numPixels = (gRadius * 2) + 1
    
    'Calculate a standard deviation (sigma) using the GIMP formula:
    Dim stdDev As Double, stdDev2 As Double, stdDev3 As Double
    If gRadius > 1 Then
        stdDev = Sqr(-(gRadius * gRadius) / (2 * Log(1# / 255#)))
    'Note that this is my addition - for a radius of 1 the GIMP formula results in too small of a sigma value
    Else
        stdDev = 0.5
    End If
    stdDev2 = stdDev * stdDev
    stdDev3 = stdDev * 3
    
    'Populate the kernel using that sigma
    Dim i As Long
    Dim curVal As Double, sumVal As Double
    sumVal = 0
    
    For i = -gRadius To gRadius
        curVal = (1 / (Sqr(PI_DOUBLE) * stdDev)) * (EULER ^ (-1 * ((i * i) / (2 * stdDev2))))
        
        'Ignore values less than 3 sigma
        If curVal < stdDev3 Then
            sumVal = sumVal + curVal
            gKernel(i) = curVal
        Else
            gKernel(i) = 0
        End If
    Next i
    
    'Find new bounds, which may exist if parts of the kernel lie outside the 3-sigma relevance limit
    Dim gLB As Long, gUB As Long
    
    gLB = -gRadius
    gUB = gRadius
    If gRadius > 1 Then
        For i = gLB To 0
            If gKernel(i) = 0 Then gLB = i + 1
        Next i
   
        For i = gUB To 0 Step -1
            If gKernel(i) = 0 Then gUB = i - 1
        Next i
   
    End If
    
    'Finally, normalize the kernel so that all values sum to 1
    For i = gLB To gUB
        gKernel(i) = gKernel(i) / sumVal
    Next i
        
    'We now have a normalized 1-dimensional gaussian kernel available for convolution.
    
    'Color variables - in this case, sums for each color component
    Dim rSum As Double, gSum As Double, bSum As Double, aSum As Double
    
    'To increase speed, we now build a look-up table of gaussian values.  This can be used in place of floating-point multiplication.
    Dim glLookup() As Single
    ReDim glLookup(0 To 255, gLB To gUB) As Single
    For X = gLB To gUB
        For Y = 0 To 255
            glLookup(Y, X) = Y * gKernel(X)
        Next Y
    Next X
        
    'Next, prepare 1D arrays that will be used to point at source and destination pixel data.  VB accesses 1D arrays more quickly
    ' than 2D arrays, and this technique shaves precious time off the final calculation.
    Dim scanlineSize As Long
    scanlineSize = srcLayer.getLayerArrayWidth
    Dim origDIBPointer As Long
    origDIBPointer = srcLayer.getLayerDIBits
    Dim dstDIBPointer As Long
    dstDIBPointer = gaussLayer.getLayerDIBits
    
    Dim tmpImageData() As Byte
    Dim tmpSA As SAFEARRAY1D
    With tmpSA
        .cbElements = 1
        .cDims = 1
        .lBound = 0
        .cElements = scanlineSize
        .pvData = origDIBPointer
    End With
        
    Dim tmpDstImageData() As Byte
    Dim tmpDstSA As SAFEARRAY1D
    With tmpDstSA
        .cbElements = 1
        .cDims = 1
        .lBound = 0
        .cElements = scanlineSize
        .pvData = dstDIBPointer
    End With
    
    'We now convolve the image twice - once in the horizontal direction, then again in the vertical direction.  This is
    ' referred to as "separable" convolution, and it's much faster than than traditional convolution, especially for
    ' large radii (the exact speed gain for a P x Q kernel is PQ/(P + Q) - so for a radius of 4 (which is an actual kernel
    ' of 9x9) the processing time is 4.5x faster).
    
    'First, perform a horizontal convolution.
        
    Dim chkX As Long
    
    'Loop through each pixel in the image, converting values as we go
    For Y = 0 To finalY
        
        'Accessing multidimensional arrays in VB is slow.  We cheat this by pointing a one-dimensional array
        ' at the current source and destination lines, then using that to access pixel data.
        tmpSA.pvData = origDIBPointer + scanlineSize * Y
        CopyMemory ByVal VarPtrArray(tmpImageData()), VarPtr(tmpSA), 4
        
        tmpDstSA.pvData = dstDIBPointer + scanlineSize * Y
        CopyMemory ByVal VarPtrArray(tmpDstImageData()), VarPtr(tmpDstSA), 4
                
    For X = initX To finalX
        
        QuickVal = X * qvDepth
    
        rSum = 0
        gSum = 0
        bSum = 0
                
        'Apply the convolution to the intermediate gaussian array
        For i = gLB To gUB
                        
            chkX = X + i
            
            'We need to give special treatment to pixels that lie off the image
            If chkX < initX Then
                chkX = initX
            Else
                If chkX > finalX Then chkX = finalX
            End If
            
            QuickValInner = chkX * qvDepth
                
            rSum = rSum + glLookup(tmpImageData(QuickValInner + 2), i)
            gSum = gSum + glLookup(tmpImageData(QuickValInner + 1), i)
            bSum = bSum + glLookup(tmpImageData(QuickValInner), i)
       
        Next i
                
        'We now have sums for each of red, green, blue (and potentially alpha).  Apply those values to the source array.
        tmpDstImageData(QuickVal + 2) = rSum
        tmpDstImageData(QuickVal + 1) = gSum
        tmpDstImageData(QuickVal) = bSum
        
        'If alpha must be checked, do it now
        If chkAlpha Then
            
            aSum = 0
            
            For i = gLB To gUB
            
                'curFactor = gKernel(i)
                chkX = X + i
                If chkX < initX Then chkX = initX
                If chkX > finalX Then chkX = finalX
                aSum = aSum + glLookup(tmpImageData(chkX * qvDepth + 3), i)
                
            Next i
            
            tmpDstImageData(QuickVal + 3) = aSum
            
        End If
        
    Next X
        If Not suppressMessages Then
            If (Y And progBarCheck) = 0 Then SetProgBarVal Y + modifyProgBarOffset
        End If
    Next Y
    
    CopyMemory ByVal VarPtrArray(tmpImageData()), 0&, 4
    CopyMemory ByVal VarPtrArray(tmpDstImageData()), 0&, 4
    
    dstDIBPointer = dstLayer.getLayerDIBits
    tmpDstSA.pvData = dstDIBPointer
    
    'The source array now contains a horizontally convolved image.  We now need to convolve it vertically.
    Dim chkY As Long
    
    For Y = initY To finalY
    
        'Accessing multidimensional arrays in VB is slow.  We cheat this by pointing a one-dimensional array
        ' at the current destination line, then using that to access pixel data.
        tmpDstSA.pvData = dstDIBPointer + scanlineSize * Y
        CopyMemory ByVal VarPtrArray(tmpDstImageData()), VarPtr(tmpDstSA), 4
    
    For X = initX To finalX
    
        QuickVal = X * qvDepth
    
        rSum = 0
        gSum = 0
        bSum = 0
        aSum = 0
    
        'Apply the convolution to the destination array, using the gaussian array as the source.
        For i = gLB To gUB
        
            chkY = Y + i
            
            'We need to give special treatment to pixels that lie off the image
            If chkY < initY Then
                chkY = initY
            Else
                If chkY > finalY Then chkY = finalY
            End If
                                    
            rSum = rSum + glLookup(GaussImageData(QuickVal + 2, chkY), i)
            gSum = gSum + glLookup(GaussImageData(QuickVal + 1, chkY), i)
            bSum = bSum + glLookup(GaussImageData(QuickVal, chkY), i)
                    
        Next i
        
        'We now have sums for each of red, green, blue (and potentially alpha).  Apply those values to the source array.
        tmpDstImageData(QuickVal + 2) = rSum
        tmpDstImageData(QuickVal + 1) = gSum
        tmpDstImageData(QuickVal) = bSum
        
        'If alpha must be checked, do it now
        If chkAlpha Then
        
            'Apply the convolution to the destination array, using the gaussian array as the source.
            For i = gLB To gUB
                'curFactor = gKernel(i)
                chkY = Y + i
                If chkY < initY Then chkY = initY
                If chkY > finalY Then chkY = finalY
                aSum = aSum + glLookup(GaussImageData(QuickVal + 3, chkY), i)
            Next i
        
            tmpDstImageData(QuickVal + 3) = aSum
        
        End If
                
    Next X
        If Not suppressMessages Then
            If (Y And progBarCheck) = 0 Then SetProgBarVal (Y + finalY) + modifyProgBarOffset
        End If
    Next Y
        
    'With our work complete, point all ImageData() arrays away from their DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(tmpDstImageData()), 0&, 4
    
    CopyMemory ByVal VarPtrArray(GaussImageData), 0&, 4
    Erase GaussImageData
    
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    'We can also erase our intermediate gaussian layer
    gaussLayer.eraseLayer
    Set gaussLayer = Nothing
        
End Sub

'Convert the image's color depth to a new value.  (Currently, only 24bpp and 32bpp is allowed.)
Public Sub ConvertImageColorDepth(ByVal newColorDepth As Long)

    Message "Converting image mode..."

    If newColorDepth = 24 Then
    
        'Ask the current layer to convert itself to 24bpp mode
        pdImages(CurrentImage).mainLayer.convertTo24bpp
    
        'Change the menu entries to match
        tInit tImgMode32bpp, False
        
    Else
    
        'Ask the current layer to convert itself to 32bpp mode
        pdImages(CurrentImage).mainLayer.convertTo32bpp
    
        'Change the menu entries to match
        tInit tImgMode32bpp, True
    
    End If
    
    Message "Finished."
    
    'Redraw the main window
    ScrollViewport FormMain.ActiveForm

End Sub

'Load the last Undo file and alpha-blend it with the current image
Public Sub MenuFadeLastEffect()

    Message "Fading last effect..."
    
    'Create a temporary layer and use it to load the last Undo file's pixel data
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer
    tmpLayer.createFromFile GetLastUndoFile()
    
    'Create a local array and point it at the pixel data of that undo file
    Dim uImageData() As Byte
    Dim uSA As SAFEARRAY2D
    prepSafeArray uSA, tmpLayer
    CopyMemory ByVal VarPtrArray(uImageData()), VarPtr(uSA), 4
        
    'Create another array, but point it at the pixel data of the current image
    Dim cImageData() As Byte
    Dim cSA As SAFEARRAY2D
    prepImageData cSA
    CopyMemory ByVal VarPtrArray(cImageData()), VarPtr(cSA), 4
    
    'Because the undo file and current image may be different sizes (if the last action was a resize, for example), we need to
    ' find the minimum width and height to make sure there are no out-of-bound errors.
    Dim minWidth As Long, minHeight As Long
    If tmpLayer.getLayerWidth < pdImages(CurrentImage).Width Then minWidth = tmpLayer.getLayerWidth Else minWidth = pdImages(CurrentImage).Width
    If tmpLayer.getLayerHeight < pdImages(CurrentImage).Height Then minHeight = tmpLayer.getLayerHeight Else minHeight = pdImages(CurrentImage).Height
        
    'Set the progress bar maximum value to that minimum width value
    SetProgBarMax minWidth
    
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, QuickValUndo As Long, qvDepth As Long, qvDepthUndo As Long
    qvDepth = pdImages(CurrentImage).mainLayer.getLayerColorDepth \ 8
    qvDepthUndo = tmpLayer.getLayerColorDepth \ 8
        
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Local loop variables can be more efficiently cached by VB's compiler
    Dim X As Long, Y As Long
    
    'Finally, prepare a look-up table for the alpha-blend
    Dim aLookUp(0 To 255, 0 To 255) As Byte
    Dim tmpCalc As Long
    
    For X = 0 To 255
    For Y = 0 To 255
        tmpCalc = (X + Y) \ 2
        aLookUp(X, Y) = CByte(tmpCalc)
    Next Y
    Next X
    
    'Loop through both images, alpha-blending pixels as we go
    For X = 0 To minWidth - 1
        QuickVal = X * qvDepth
        QuickValUndo = X * qvDepthUndo
    For Y = 0 To minHeight - 1
        cImageData(QuickVal, Y) = aLookUp(cImageData(QuickVal, Y), uImageData(QuickValUndo, Y))
        cImageData(QuickVal + 1, Y) = aLookUp(cImageData(QuickVal + 1, Y), uImageData(QuickValUndo + 1, Y))
        cImageData(QuickVal + 2, Y) = aLookUp(cImageData(QuickVal + 2, Y), uImageData(QuickValUndo + 2, Y))
    Next Y
        If (X And progBarCheck) = 0 Then SetProgBarVal X
    Next X
        
    'With our work complete, point both ImageData() arrays away from their respective DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(cImageData), 0&, 4
    Erase cImageData
    
    CopyMemory ByVal VarPtrArray(uImageData), 0&, 4
    Erase uImageData
    
    'Erase our temporary layer as well
    tmpLayer.eraseLayer
    Set tmpLayer = Nothing
    
    'Finally, pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData
    
End Sub

'Render an image using faux thermography; basically, map luminance values as if they were heat, and use a modified hue spectrum for representation.
' (I have manually tweaked the values at certain ranges to better approximate actual thermography.)
Public Sub MenuHeatMap()

    Message "Performing thermographic analysis..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepImageData tmpSA
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim X As Long, Y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long
    Dim grayVal As Long
    Dim hVal As Double, sVal As Double, lVal As Double
    Dim h As Double, s As Double, l As Double
    
    'Because gray values are constant, we can use a look-up table to calculate them
    Dim gLookup(0 To 765) As Byte
    For X = 0 To 765
        gLookup(X) = CByte(X \ 3)
    Next X
        
    'Apply the filter
    For X = initX To finalX
        QuickVal = X * qvDepth
    For Y = initY To finalY
        
        r = ImageData(QuickVal + 2, Y)
        g = ImageData(QuickVal + 1, Y)
        b = ImageData(QuickVal, Y)
        
        grayVal = gLookup(r + g + b)
        
        'Based on the luminance of this pixel, apply a predetermined hue gradient (stretching between -1 and 5)
        hVal = (CSng(grayVal) / 255) * 360
        
        'If the hue is "below" blue, gradually darken the corresponding luminance value
        If hVal < 120 Then
            lVal = (0.35 * (hVal / 120)) + 0.15
        Else
            lVal = 0.5
        End If
        
        'Invert the hue
        hVal = 360 - hVal
                
        'Place hue in the range of -1 to 5, per the requirements of our HSL conversion algorithm
        hVal = (hVal - 60) / 60
        
        'Use nearly full saturation (for dramatic effect)
        sVal = 0.8
        
        'Use RGB to calculate hue, saturation, and luminance
        tRGBToHSL r, g, b, h, s, l
        
        'Now convert those HSL values back to RGB, but substitute in our artificial hue value (calculated above)
        tHSLToRGB hVal, sVal, lVal, r, g, b
        
        ImageData(QuickVal + 2, Y) = r
        ImageData(QuickVal + 1, Y) = g
        ImageData(QuickVal, Y) = b
        
    Next Y
        If (X And progBarCheck) = 0 Then SetProgBarVal X
    Next X
        
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData
    
End Sub

'A very neat comic-book filter that actually blends together a number of other filters into one!
Public Sub MenuComicBook()
    
    Dim gRadius As Long
    gRadius = 20
    
    Dim gThreshold As Long
    gThreshold = 8
    
    Message "Animating image (stage 1 of 3)..."
                
    'More color variables - in this case, sums for each color component
    Dim r As Long, g As Long, b As Long
    Dim r2 As Long, g2 As Long, b2 As Long
    Dim tDelta As Long
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstSA As SAFEARRAY2D
    prepImageData dstSA
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent blurred pixel values from spreading across the image as we go.)
    Dim srcLayer As pdLayer
    Set srcLayer = New pdLayer
    srcLayer.createFromExistingLayer workingLayer
    
    Dim gaussLayer As pdLayer
    Set gaussLayer = New pdLayer
    gaussLayer.createFromExistingLayer workingLayer
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim X As Long, Y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
    
    CreateGaussianBlurLayer gRadius, srcLayer, gaussLayer, False, finalY + finalY + finalX + finalX
        
    'Now that we have a gaussian layer created in gaussLayer, we can point arrays toward it and the source layer
    Dim dstImageData() As Byte
    prepImageData dstSA
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepSafeArray srcSA, srcLayer
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    Dim GaussImageData() As Byte
    Dim gaussSA As SAFEARRAY2D
    prepSafeArray gaussSA, gaussLayer
    CopyMemory ByVal VarPtrArray(GaussImageData()), VarPtr(gaussSA), 4
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
        
    Message "Animating image (stage 2 of 3)..."
        
    Dim blendVal As Double
    
    'The final step of the smart blur function is to find edges, and replace them with the blurred data as necessary
    For X = initX To finalX
        QuickVal = X * qvDepth
    For Y = initY To finalY
        
        'Retrieve the original image's pixels
        r = srcImageData(QuickVal + 2, Y)
        g = srcImageData(QuickVal + 1, Y)
        b = srcImageData(QuickVal, Y)
        
        tDelta = (213 * r + 715 * g + 72 * b) \ 1000
        
        'Now, retrieve the gaussian pixels
        r2 = GaussImageData(QuickVal + 2, Y)
        g2 = GaussImageData(QuickVal + 1, Y)
        b2 = GaussImageData(QuickVal, Y)
        
        'Calculate a delta between the two
        tDelta = tDelta - ((213 * r2 + 715 * g2 + 72 * b2) \ 1000)
        If tDelta < 0 Then tDelta = -tDelta
                
        'If the delta is below the specified threshold, replace it with the blurred data.
        If tDelta > gThreshold Then
            If tDelta <> 0 Then blendVal = 1 - (gThreshold / tDelta) Else blendVal = 0
            dstImageData(QuickVal + 2, Y) = BlendColors(srcImageData(QuickVal + 2, Y), GaussImageData(QuickVal + 2, Y), blendVal)
            dstImageData(QuickVal + 1, Y) = BlendColors(srcImageData(QuickVal + 1, Y), GaussImageData(QuickVal + 1, Y), blendVal)
            dstImageData(QuickVal, Y) = BlendColors(srcImageData(QuickVal, Y), GaussImageData(QuickVal, Y), blendVal)
            If qvDepth = 4 Then dstImageData(QuickVal + 3, Y) = BlendColors(srcImageData(QuickVal + 3, Y), GaussImageData(QuickVal + 3, Y), blendVal)
        End If
        
    Next Y
        If (X And progBarCheck) = 0 Then SetProgBarVal X + (finalY * 2)
    Next X
        
    'With our work complete, release all arrays
    CopyMemory ByVal VarPtrArray(GaussImageData), 0&, 4
    Erase GaussImageData
    
    gaussLayer.eraseLayer
    Set gaussLayer = Nothing
    
    'The last thing we need to do is sketch in the edges of the image.
    
    Message "Animating image (stage 3 of 3)..."
    
    'We can't do this at the borders of the image, so shrink the functional area by one in each dimension.
    initX = initX + 1
    initY = initY + 1
    finalX = finalX - 1
    finalY = finalY - 1
    
    Dim QuickValRight As Long, QuickValLeft As Long, tmpColor As Long, tMin As Long
    Dim z As Long
        
    'Loop through each pixel in the image, converting values as we go
    For X = initX To finalX
        QuickVal = X * qvDepth
        QuickValRight = (X + 1) * qvDepth
        QuickValLeft = (X - 1) * qvDepth
    For Y = initY To finalY
        For z = 0 To 2
    
            tMin = 255
            tmpColor = srcImageData(QuickValRight + z, Y)
            If tmpColor < tMin Then tMin = tmpColor
            tmpColor = srcImageData(QuickValRight + z, Y - 1)
            If tmpColor < tMin Then tMin = tmpColor
            tmpColor = srcImageData(QuickValRight + z, Y + 1)
            If tmpColor < tMin Then tMin = tmpColor
            tmpColor = srcImageData(QuickValLeft + z, Y)
            If tmpColor < tMin Then tMin = tmpColor
            tmpColor = srcImageData(QuickValLeft + z, Y - 1)
            If tmpColor < tMin Then tMin = tmpColor
            tmpColor = srcImageData(QuickValLeft + z, Y + 1)
            If tmpColor < tMin Then tMin = tmpColor
            tmpColor = srcImageData(QuickVal + z, Y)
            If tmpColor < tMin Then tMin = tmpColor
            tmpColor = srcImageData(QuickVal + z, Y - 1)
            If tmpColor < tMin Then tMin = tmpColor
            tmpColor = srcImageData(QuickVal + z, Y + 1)
            If tmpColor < tMin Then tMin = tmpColor
            
            If tMin > 255 Then tMin = 255
            If tMin < 0 Then tMin = 0
            
            Select Case z
            
                Case 0
                    b = 255 - (srcImageData(QuickVal, Y) - tMin)
            
                Case 1
                    g = 255 - (srcImageData(QuickVal + 1, Y) - tMin)
                    
                Case 2
                    r = 255 - (srcImageData(QuickVal + 2, Y) - tMin)
            
            End Select
                    
        Next z
        
        r2 = dstImageData(QuickVal + 2, Y)
        g2 = dstImageData(QuickVal + 1, Y)
        b2 = dstImageData(QuickVal, Y)
        
        r = ((CSng(r) / 255) * (CSng(r2) / 255)) * 255
        g = ((CSng(g) / 255) * (CSng(g2) / 255)) * 255
        b = ((CSng(b) / 255) * (CSng(b2) / 255)) * 255
        
        dstImageData(QuickVal + 2, Y) = r
        dstImageData(QuickVal + 1, Y) = g
        dstImageData(QuickVal, Y) = b
        
    Next Y
        If (X And progBarCheck) = 0 Then SetProgBarVal X + finalX + (finalY * 2)
    Next X
    
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData

End Sub

'Wacky filter discovered by trial-and-error.  I named it "synthesize".
Public Sub MenuSynthesize()

    Message "Synthesizing new image..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepImageData tmpSA
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim X As Long, Y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long
    Dim grayVal As Long
    
    'Because gray values are constant, we can use a look-up table to calculate them
    Dim gLookup(0 To 765) As Byte
    For X = 0 To 765
        gLookup(X) = CByte(X \ 3)
    Next X
        
    'Apply the filter
    For X = initX To finalX
        QuickVal = X * qvDepth
    For Y = initY To finalY
        
        r = ImageData(QuickVal + 2, Y)
        g = ImageData(QuickVal + 1, Y)
        b = ImageData(QuickVal, Y)
        
        grayVal = gLookup(r + g + b)
        
        r = g + b - grayVal
        g = r + b - grayVal
        b = r + g - grayVal
        
        If r > 255 Then r = 255
        If r < 0 Then r = 0
        If g > 255 Then g = 255
        If g < 0 Then g = 0
        If b > 255 Then b = 255
        If b < 0 Then b = 0
        
        ImageData(QuickVal + 2, Y) = r
        ImageData(QuickVal + 1, Y) = g
        ImageData(QuickVal, Y) = b
        
    Next Y
        If (X And progBarCheck) = 0 Then SetProgBarVal X
    Next X
        
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData

End Sub

'Another random filter discovered by trial-and-error.  "Alien" effect.
Public Sub MenuAlien()

    Message "Abducting image and probing it for weaknesses..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepImageData tmpSA
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim X As Long, Y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long
    Dim newR As Long, newG As Long, newB As Long
        
    'Apply the filter
    For X = initX To finalX
        QuickVal = X * qvDepth
    For Y = initY To finalY
        
        r = ImageData(QuickVal + 2, Y)
        g = ImageData(QuickVal + 1, Y)
        b = ImageData(QuickVal, Y)
        
        newR = b + g - r
        newG = r + b - g
        newB = r + g - b
        
        If newR > 255 Then newR = 255
        If newR < 0 Then newR = 0
        If newG > 255 Then newG = 255
        If newG < 0 Then newG = 0
        If newB > 255 Then newB = 255
        If newB < 0 Then newB = 0
        
        ImageData(QuickVal + 2, Y) = newR
        ImageData(QuickVal + 1, Y) = newG
        ImageData(QuickVal, Y) = newB
        
    Next Y
        If (X And progBarCheck) = 0 Then SetProgBarVal X
    Next X
        
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData
  
End Sub

'Very improved version of "sepia".  This is more involved than a typical "change to brown" effect - the white balance and
' shading is also adjusted to give the image a more "antique" look.
Public Sub MenuAntique()
    
    Message "Accelerating to 88mph in order to antique-ify this image..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepImageData tmpSA
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim X As Long, Y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'We're going to need grayscale values as part of the effect; grayscale is easily optimized via a look-up table
    Dim gLookup(0 To 765) As Byte
    For X = 0 To 765
        gLookup(X) = CByte(X \ 3)
    Next X
    
    'We're going to use gamma conversion as part of the effect; gamma is easily optimized via a look-up table
    Dim gammaLookUp(0 To 255) As Byte
    Dim tmpVal As Double
    For X = 0 To 255
        tmpVal = X / 255
        tmpVal = tmpVal ^ (1 / 1.6)
        tmpVal = tmpVal * 255
        If tmpVal > 255 Then tmpVal = 255
        If tmpVal < 0 Then tmpVal = 0
        gammaLookUp(X) = CByte(tmpVal)
    Next X
    
    'Finally, we also need to adjust brightness.  A look-up table is once again invaluable
    Dim bLookup(0 To 255) As Byte
    For X = 0 To 255
        tmpVal = X * 1.75
        If tmpVal > 255 Then tmpVal = 255
        bLookup(X) = CByte(tmpVal)
    Next X
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long
    Dim newR As Long, newG As Long, newB As Long
    Dim gray As Long
        
    'Apply the filter
    For X = initX To finalX
        QuickVal = X * qvDepth
    For Y = initY To finalY
    
        r = ImageData(QuickVal + 2, Y)
        g = ImageData(QuickVal + 1, Y)
        b = ImageData(QuickVal, Y)
        
        gray = gLookup(r + g + b)
        
        r = (r + gray) \ 2
        g = (g + gray) \ 2
        b = (b + gray) \ 2
        
        r = (g * b) \ 256
        g = (b * r) \ 256
        b = (r * g) \ 256
        
        newR = bLookup(r)
        newG = bLookup(g)
        newB = bLookup(b)
        
        newR = gammaLookUp(newR)
        newG = gammaLookUp(newG)
        newB = gammaLookUp(newB)
        
        ImageData(QuickVal + 2, Y) = newR
        ImageData(QuickVal + 1, Y) = newG
        ImageData(QuickVal, Y) = newB
        
    Next Y
        If (X And progBarCheck) = 0 Then SetProgBarVal X
    Next X
        
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData
    
End Sub

'Dull but standard "sepia" transformation.  Values are derived from the w3c standard at:
' https://dvcs.w3.org/hg/FXTF/raw-file/tip/filters/index.html#sepiaEquivalent
Public Sub MenuSepia()
    
    Message "Engaging hipsters to perform sepia conversion..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepImageData tmpSA
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim X As Long, Y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long
    Dim newR As Double, newG As Double, newB As Double
        
    'Apply the filter
    For X = initX To finalX
        QuickVal = X * qvDepth
    For Y = initY To finalY
    
        r = ImageData(QuickVal + 2, Y)
        g = ImageData(QuickVal + 1, Y)
        b = ImageData(QuickVal, Y)
                
        newR = CSng(r) * 0.393 + CSng(g) * 0.769 + CSng(b) * 0.189
        newG = CSng(r) * 0.349 + CSng(g) * 0.686 + CSng(b) * 0.168
        newB = CSng(r) * 0.272 + CSng(g) * 0.534 + CSng(b) * 0.131
        
        r = newR
        g = newG
        b = newB
        
        If r > 255 Then r = 255
        If g > 255 Then g = 255
        If b > 255 Then b = 255
        
        ImageData(QuickVal + 2, Y) = r
        ImageData(QuickVal + 1, Y) = g
        ImageData(QuickVal, Y) = b
        
    Next Y
        If (X And progBarCheck) = 0 Then SetProgBarVal X
    Next X
        
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData
    
End Sub

'Makes the picture appear like it has been shaken
Public Sub MenuVibrate()
    g_FilterSize = 5
    ReDim g_FM(-2 To 2, -2 To 2) As Long
    g_FM(-2, -2) = 1
    g_FM(-1, -1) = -1
    g_FM(0, 0) = 1
    g_FM(1, 1) = -1
    g_FM(2, 2) = 1
    g_FM(-1, 1) = 1
    g_FM(-2, 2) = -1
    g_FM(1, -1) = 1
    g_FM(2, -2) = -1
    g_FilterWeight = 1
    g_FilterBias = 0
    DoFilter "Vibrate"
End Sub

'Another filter found by trial-and-error.  "Dream" effect.
Public Sub MenuDream()

    Message "Putting image to sleep, then measuring its REM cycles..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepImageData tmpSA
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim X As Long, Y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long
    Dim newR As Long, newG As Long, newB As Long
    Dim grayVal As Long
    
    'Because gray values are constant, we can use a look-up table to calculate them
    Dim gLookup(0 To 765) As Byte
    For X = 0 To 765
        gLookup(X) = CByte(X \ 3)
    Next X
        
    'Apply the filter
    For X = initX To finalX
        QuickVal = X * qvDepth
    For Y = initY To finalY
        
        newR = ImageData(QuickVal + 2, Y)
        newG = ImageData(QuickVal + 1, Y)
        newB = ImageData(QuickVal, Y)
        
        grayVal = gLookup(newR + newG + newB)
        
        r = Abs(newR - grayVal) + Abs(newR - newG) + Abs(newR - newB) + (newR \ 2)
        g = Abs(newG - grayVal) + Abs(newG - newB) + Abs(newG - newR) + (newG \ 2)
        b = Abs(newB - grayVal) + Abs(newB - newR) + Abs(newB - newG) + (newB \ 2)
        
        If r > 255 Then r = 255
        If r < 0 Then r = 0
        If g > 255 Then g = 255
        If g < 0 Then g = 0
        If b > 255 Then b = 255
        If b < 0 Then b = 0
        
        ImageData(QuickVal + 2, Y) = r
        ImageData(QuickVal + 1, Y) = g
        ImageData(QuickVal, Y) = b
        
    Next Y
        If (X And progBarCheck) = 0 Then SetProgBarVal X
    Next X
        
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData

End Sub

'A bright-green filter I've aptly named "radioactive".
Public Sub MenuRadioactive()

    Message "Injecting image with non-ionizing radiation..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepImageData tmpSA
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim X As Long, Y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long
    Dim newR As Long, newG As Long, newB As Long
        
    'Apply the filter
    For X = initX To finalX
        QuickVal = X * qvDepth
    For Y = initY To finalY
        
        r = ImageData(QuickVal + 2, Y)
        g = ImageData(QuickVal + 1, Y)
        b = ImageData(QuickVal, Y)
        
        If r = 0 Then r = 1
        If g = 0 Then g = 1
        If b = 0 Then b = 1
        
        newR = (g * b) \ r
        newG = (r * b) \ g
        newB = (r * g) \ b
        
        If newR > 255 Then newR = 255
        If newG > 255 Then newG = 255
        If newB > 255 Then newB = 255
        
        newG = 255 - newG
        
        ImageData(QuickVal + 2, Y) = newR
        ImageData(QuickVal + 1, Y) = newG
        ImageData(QuickVal, Y) = newB
        
    Next Y
        If (X And progBarCheck) = 0 Then SetProgBarVal X
    Next X
        
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData

End Sub

'Stretch out the contrast and convert the image to dramatic black and white.  "Comic book" filter.
Public Sub MenuFilmNoir()

    Message "Embuing image with the essence of F. Miller..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepImageData tmpSA
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim X As Long, Y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long
    Dim grayVal As Long
    
    'Because gray values are constant, we can use a look-up table to calculate them
    Dim gLookup(0 To 765) As Byte
    For X = 0 To 765
        gLookup(X) = CByte(X \ 3)
    Next X
        
    'Apply the filter
    For X = initX To finalX
        QuickVal = X * qvDepth
    For Y = initY To finalY
        
        r = ImageData(QuickVal + 2, Y)
        g = ImageData(QuickVal + 1, Y)
        b = ImageData(QuickVal, Y)
        
        r = Abs(r * (g - b + g + r)) / 255
        g = Abs(r * (b - g + b + r)) / 255
        b = Abs(g * (b - g + b + r)) / 255
        
        If r > 255 Then r = 255
        If g > 255 Then g = 255
        If b > 255 Then b = 255
        
        grayVal = gLookup(r + g + b)
        
        ImageData(QuickVal + 2, Y) = grayVal
        ImageData(QuickVal + 1, Y) = grayVal
        ImageData(QuickVal, Y) = grayVal
        
    Next Y
        If (X And progBarCheck) = 0 Then SetProgBarVal X
    Next X
        
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData

End Sub

'Subroutine for counting the number of unique colors in an image
Public Sub MenuCountColors()
    
    Message "Counting the number of unique colors in this image..."
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepImageData tmpSA
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim X As Long, Y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'This array will track whether or not a given color has been detected in the image
    Dim UniqueColors() As Boolean
    ReDim UniqueColors(0 To 16777216) As Boolean
    
    'Total number of unique colors counted so far
    Dim totalCount As Long
    totalCount = 0
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long
    Dim chkValue As Long
        
    'Apply the filter
    For X = initX To finalX
        QuickVal = X * qvDepth
    For Y = initY To finalY
        
        r = ImageData(QuickVal + 2, Y)
        g = ImageData(QuickVal + 1, Y)
        b = ImageData(QuickVal, Y)
        
        chkValue = RGB(r, g, b)
        If UniqueColors(chkValue) = False Then
            totalCount = totalCount + 1
            UniqueColors(chkValue) = True
        End If
        
    Next Y
        If (X And progBarCheck) = 0 Then SetProgBarVal X
    Next X
        
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Also, erase the counting array
    Erase UniqueColors
    
    'Reset the progress bar
    SetProgBarVal 0
    
    'Show the user our finaly tally
    Message "Total number of unique colors: " & totalCount
    MsgBox "This image contains " & totalCount & " unique colors.", vbOKOnly + vbApplicationModal + vbInformation, "Count Image Colors"
    
End Sub

'You can use this section of code to test out your own filters.  I've left some sample code below.
Public Sub MenuTest()
    
    MsgBox "This menu item only appears in the Visual Basic IDE." & vbCrLf & vbCrLf & "You can use the MenuTest() sub in the Filters_Miscellaneous module to test out your own filters.  I typically do this first, then once the filter is working properly, I give it a subroutine of its own.", vbInformation + vbOKOnly + vbApplicationModal, PROGRAMNAME & " Pro Tip"
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepImageData tmpSA
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim X As Long, Y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = curLayerValues.BytesPerPixel
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
    
    'Because gray values are constant, we can use a look-up table to calculate them
    Dim gLookup(0 To 765) As Byte
    For X = 0 To 765
        gLookup(X) = CByte(X \ 3)
    Next X
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long, grayVal As Long
    Dim newR As Long, newG As Long, newB As Long
    Dim hVal As Double, sVal As Double, lVal As Double
    Dim h As Double, s As Double, l As Double
        
    'Apply the filter
    For X = initX To finalX
        QuickVal = X * qvDepth
    For Y = initY To finalY
        
        r = ImageData(QuickVal + 2, Y)
        g = ImageData(QuickVal + 1, Y)
        b = ImageData(QuickVal, Y)
        
        grayVal = gLookup(r + g + b)
        
        'Put interesting color transformations here.  As an example, here's one possible sepia formula.
        newR = grayVal + 40
        newG = grayVal + 20
        newB = grayVal - 30
                                
        If newR < 0 Then newR = 0
        If newG < 0 Then newG = 0
        If newB < 0 Then newB = 0
        
        If newR > 255 Then newR = 255
        If newG > 255 Then newG = 255
        If newB > 255 Then newB = 255
                
        ImageData(QuickVal + 2, Y) = newR
        ImageData(QuickVal + 1, Y) = newG
        ImageData(QuickVal, Y) = newB
                
    Next Y
        If (X And progBarCheck) = 0 Then SetProgBarVal X
    Next X
        
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Pass control to finalizeImageData, which will handle the rest of the rendering
    finalizeImageData

    
End Sub

'This function will return the luminance value of an RGB triplet.  Note that the value will be in the [0,255] range instead
' of the usual [0,1.0] one.
Public Function getLuminance(ByVal r As Long, ByVal g As Long, ByVal b As Long) As Long
    Dim Max As Long, Min As Long
    Max = MaximumInt(r, g, b)
    Min = MinimumInt(r, g, b)
    getLuminance = (Max + Min) \ 2
End Function

'HSL <-> RGB conversion routines
Public Sub tRGBToHSL(r As Long, g As Long, b As Long, h As Double, s As Double, l As Double)
    
    Dim Max As Double, Min As Double
    Dim delta As Double
    Dim rR As Double, rG As Double, rB As Double
    
    rR = r / 255
    rG = g / 255
    rB = b / 255

    'Note: HSL are calculated in the following ranges:
    ' Hue: [-1,5]
    ' Saturation: [0,1] (Note that if saturation = 0, hue is technically undefined)
    ' Lightness: [0,1]

    Max = Maximum(rR, rG, rB)
    Min = Minimum(rR, rG, rB)
        
    'Calculate luminance
    l = (Max + Min) / 2
        
    'If the maximum and minimum are identical, this image is gray, meaning it has no saturation and an undefined hue.
    If Max = Min Then
        s = 0
        h = 0
    Else
        
        'Calculate saturation
        If l <= 0.5 Then
            s = (Max - Min) / (Max + Min)
        Else
            s = (Max - Min) / (2 - Max - Min)
        End If
        
        'Calculate hue
        delta = Max - Min

        If rR = Max Then
            h = (rG - rB) / delta    '{Resulting color is between yellow and magenta}
        ElseIf rG = Max Then
            h = 2 + (rB - rR) / delta '{Resulting color is between cyan and yellow}
        ElseIf rB = Max Then
            h = 4 + (rR - rG) / delta '{Resulting color is between magenta and cyan}
        End If
        
        'If you prefer hue in the [0,360] range instead of [-1, 5] you can use this code
        'h = h * 60
        'If h < 0 Then h = h + 360

    End If

    'Tanner's final note: if byte values are preferred to floating-point, this code will return hue on [0,240],
    ' saturation on [0,255], and luminance on [0,255]
    'H = Int(H * 40 + 40)
    'S = Int(S * 255)
    'L = Int(L * 255)
    
End Sub

'Convert HSL values to RGB values
Public Sub tHSLToRGB(h As Double, s As Double, l As Double, r As Long, g As Long, b As Long)

    Dim rR As Double, rG As Double, rB As Double
    Dim Min As Double, Max As Double

    'Unsaturated pixels do not technically have hue - they only have luminance
    If s = 0 Then
        rR = l: rG = l: rB = l
    Else
        If l <= 0.5 Then
            Min = l * (1 - s)
        Else
            Min = l - s * (1 - l)
        End If
      
        Max = 2 * l - Min
      
        If (h < 1) Then
            
            rR = Max
            
            If (h < 0) Then
                rG = Min
                rB = rG - h * (Max - Min)
            Else
                rB = Min
                rG = h * (Max - Min) + rB
            End If
        
        ElseIf (h < 3) Then
            
            rG = Max
         
            If (h < 2) Then
                rB = Min
                rR = rB - (h - 2) * (Max - Min)
            Else
                rR = Min
                rB = (h - 2) * (Max - Min) + rR
            End If
        
        Else
        
            rB = Max
            
            If (h < 4) Then
                rR = Min
                rG = rR - (h - 4) * (Max - Min)
            Else
                rG = Min
                rR = (h - 4) * (Max - Min) + rG
            End If
         
        End If
            
   End If
   
   r = rR * 255
   g = rG * 255
   b = rB * 255
   
   'Failsafe added 29 August '12
   'This should never return RGB values > 255, but it doesn't hurt to make sure.
   If r > 255 Then r = 255
   If g > 255 Then g = 255
   If b > 255 Then b = 255
   
End Sub

'Return the maximum of three variables
Public Function Maximum(rR As Double, rG As Double, rB As Double) As Double
   If (rR > rG) Then
      If (rR > rB) Then
         Maximum = rR
      Else
         Maximum = rB
      End If
   Else
      If (rB > rG) Then
         Maximum = rB
      Else
         Maximum = rG
      End If
   End If
End Function

'Return the minimum of three variables
Public Function Minimum(rR As Double, rG As Double, rB As Double) As Double
   If (rR < rG) Then
      If (rR < rB) Then
         Minimum = rR
      Else
         Minimum = rB
      End If
   Else
      If (rB < rG) Then
         Minimum = rB
      Else
         Minimum = rG
      End If
   End If
End Function

'Return the maximum of three variables
Public Function MaximumInt(rR As Long, rG As Long, rB As Long) As Long
   If (rR > rG) Then
      If (rR > rB) Then
         MaximumInt = rR
      Else
         MaximumInt = rB
      End If
   Else
      If (rB > rG) Then
         MaximumInt = rB
      Else
         MaximumInt = rG
      End If
   End If
End Function

'Return the minimum of three variables
Public Function MinimumInt(rR As Long, rG As Long, rB As Long) As Long
   If (rR < rG) Then
      If (rR < rB) Then
         MinimumInt = rR
      Else
         MinimumInt = rB
      End If
   Else
      If (rB < rG) Then
         MinimumInt = rB
      Else
         MinimumInt = rG
      End If
   End If
End Function
