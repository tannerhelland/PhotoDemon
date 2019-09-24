Attribute VB_Name = "Filters_Layers"
'***************************************************************************
'DIB Filters Module
'Copyright 2013-2019 by Tanner Helland
'Created: 15/February/13
'Last updated: 09/August/17
'Last update: new functions for blurring just a sub-region of an image
'
'Some filters in PhotoDemon are capable of operating "on-demand" on any supplied DIBs.  In a perfect world, *all*
' filters would work this way - but alas I did not design the program very well up front.  Going forward I will be
' moving more filters to an "on-demand" model.
'
'The benefit of filters like this is that any function can call them.  This means that a tool like "gaussian blur"
' need only be written once, and then any other function can use it at will.  This is useful for stacking multiple
' filters to create more complex effects.  It also cuts down on code maintenance because I need only perfect a
' formula once, then reference it externally.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Enum GDI_StretchBltMode
    sbm_ColorOnColor = 3
    sbm_Halftone = 4
End Enum

#If False Then
    Private Const sbm_ColorOnColor = 3, sbm_Halftone = 4
#End If

Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDestDC As Long, ByVal nStretchMode As GDI_StretchBltMode) As Long

'Pad a DIB with blank space.  This will (obviously) resize the DIB as necessary.
Public Function PadDIB(ByRef srcDIB As pdDIB, ByVal paddingSize As Long) As Boolean

    'Make a copy of the current DIB
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    tmpDIB.CreateFromExistingDIB srcDIB
    
    'Resize the source DIB to accommodate the new padding
    srcDIB.CreateBlank srcDIB.GetDIBWidth + paddingSize * 2, srcDIB.GetDIBHeight + paddingSize * 2, srcDIB.GetDIBColorDepth, 0, 0
    srcDIB.SetInitialAlphaPremultiplicationState tmpDIB.GetAlphaPremultiplication
    
    'Copy the old DIB into the center of the new DIB
    GDI.BitBltWrapper srcDIB.GetDIBDC, paddingSize, paddingSize, tmpDIB.GetDIBWidth, tmpDIB.GetDIBHeight, tmpDIB.GetDIBDC, 0, 0, vbSrcCopy
    
    'Erase the temporary DIB
    Set tmpDIB = Nothing
    
    PadDIB = True

End Function

'Pad a DIB with blank space, using a RECT so that each side can be independently resized.  Note that the rect specifies how many pixels
' on each side the image should be expanded.  It does not specify the rect of the new image (because that wouldn't tell us where to
' place the image on the new rect).
' Note that this function will (obviously) resize the DIB as part of padding it.
Public Function PadDIBRect(ByRef srcDIB As pdDIB, ByRef paddingRect As RECT) As Boolean

    'Make a copy of the current DIB
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    tmpDIB.CreateFromExistingDIB srcDIB
    
    'Resize the source DIB to accommodate the new padding
    srcDIB.CreateBlank srcDIB.GetDIBWidth + paddingRect.Left + paddingRect.Right, srcDIB.GetDIBHeight + paddingRect.Top + paddingRect.Bottom, srcDIB.GetDIBColorDepth, 0, 0
    
    'Copy the old DIB into the center of the new DIB
    GDI.BitBltWrapper srcDIB.GetDIBDC, paddingRect.Left, paddingRect.Top, tmpDIB.GetDIBWidth, tmpDIB.GetDIBHeight, tmpDIB.GetDIBDC, 0, 0, vbSrcCopy
    
    'Erase the temporary DIB
    Set tmpDIB = Nothing
    
    PadDIBRect = True

End Function

'If the application needs to quickly blur a DIB and it doesn't care how, use this function.  It will lean on GDI+ if
' available (unless otherwise requested), or fall back to a high-speed internal box blur.
Public Function QuickBlurDIB(ByRef srcDIB As pdDIB, ByVal blurRadius As Long, Optional ByVal useGDIPlusIfAvailable As Boolean = True) As Boolean

    If (blurRadius > 0) Then
    
        'If GDI+ 1.1 exists, use it for a faster blur operation.  If only v1.0 is found, fall back to one of our internal blur functions.
        '
        'ADDENDUM JAN '15: GDI+ exhibits broken behavior on Windows 8+, if the radius is less than 20px.  (Only a horizontal blur
        ' is applied, for reasons unknown.)  This problem has persisted throug multiple Windows 10 builds as well, so I think it's
        ' unlikely to be fixed.
        '
        'Either way, we provide necessary internal fallbacks to compensate, and external functions can always request that we
        ' avoid GDI+ if they don't want to deal with the headache.
        Dim gdiPlusIsAcceptable As Boolean
        
        'Attempt to see if GDI+ v1.1 (or later) is available.
        If GDI_Plus.IsGDIPlusV11Available And useGDIPlusIfAvailable Then
        
            'Next, make sure one of two things are true:
            ' 1) We are on Windows 7, OR...
            ' 2) We are on Windows 8+ and the blur radius is > 20.  Below this radius, Windows 8 doesn't blur correctly,
            '    and we've gone long enough without a patch (years!) that I don't expect MS to fix it.
            If OS.IsWin8OrLater And (blurRadius <= 20) Then
                gdiPlusIsAcceptable = False
            Else
                gdiPlusIsAcceptable = True
            End If
        
        'On XP or Vista, don't bother with GDI+
        Else
            gdiPlusIsAcceptable = False
        End If
        
        'If we think GDI+ will work, try it now.  (Note that GDI+ blurs are prone to failure, so we *definitely*
        ' need to provide a fallback mechanism.
        If gdiPlusIsAcceptable Then gdiPlusIsAcceptable = GDIPlusBlurDIB(srcDIB, blurRadius * 2, 0, 0, srcDIB.GetDIBWidth, srcDIB.GetDIBHeight)
        
        'If GDI+ is unacceptable (or if it failed), use our internal quick blur functionality.
        If (Not gdiPlusIsAcceptable) Then
        
            'Create a copy of the current DIB; we need this to hold an intermediate blur copy
            Dim tmpDIB As pdDIB
            Set tmpDIB = New pdDIB
            tmpDIB.CreateFromExistingDIB srcDIB
            If (CreateHorizontalBlurDIB(blurRadius, blurRadius, srcDIB, tmpDIB, True) <> 0) Then
                QuickBlurDIB = (CreateVerticalBlurDIB(blurRadius, blurRadius, tmpDIB, srcDIB, True) <> 0)
            End If
        
        End If
        
    End If
    
    QuickBlurDIB = True
    
End Function

'Want to blur just some sub-portion of a DIB?  Use this function instead.
Public Function QuickBlurDIBRegion(ByRef srcDIB As pdDIB, ByVal blurRadius As Long, ByRef blurBounds As RectF) As Boolean
    
    'Create a copy of the current DIB; we need this to hold an intermediate blur copy
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    tmpDIB.CreateFromExistingDIB srcDIB
    If (Filters_Layers.HorizontalBlur_SubRegion(blurRadius, blurRadius, srcDIB, tmpDIB, blurBounds) <> 0) Then
        QuickBlurDIBRegion = (Filters_Layers.VerticalBlur_SubRegion(blurRadius, blurRadius, tmpDIB, srcDIB, blurBounds) <> 0)
    End If
    
End Function

'Given a 32bpp DIB, return a "shadow" version.  (It's pretty simple, really - black out the DIB but retain alpha values.)
Public Function CreateShadowDIB(ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB) As Boolean

    'If the source DIB is not 32bpp, exit.
    If (srcDIB.GetDIBColorDepth <> 32) Then
        CreateShadowDIB = False
        Exit Function
    End If

    'Start by copying the source DIB into the destination
    dstDIB.CreateFromExistingDIB srcDIB
    
    'Create a local array and point it at the pixel data of the destination image
    Dim dstImageData() As Byte
    Dim dstSA As SafeArray2D
    PrepSafeArray dstSA, dstDIB
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, finalX As Long, finalY As Long, quickX As Long
    finalX = dstDIB.GetDIBWidth - 1
    finalY = dstDIB.GetDIBHeight - 1
    
    'Loop through all pixels in the destination image and set them to black.  Easy as pie!
    For x = 0 To finalX
        quickX = x * 4
    For y = 0 To finalY
    
        dstImageData(quickX + 2, y) = 0
        dstImageData(quickX + 1, y) = 0
        dstImageData(quickX, y) = 0
        
    Next y
    Next x
    
    'Release our array reference and exit
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    CreateShadowDIB = True

End Function

'Given two DIBs, fill one with a median-filtered version of the other.
' Per PhotoDemon convention, this function will return a non-zero value if successful, and 0 if canceled.
Public Function CreateMedianDIB(ByVal mRadius As Long, ByVal mPercent As Double, ByVal kernelShape As PD_PIXEL_REGION_SHAPE, ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SafeArray2D
    PrepSafeArray dstSA, dstDIB
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.GetDIBWidth - 1
    finalY = srcDIB.GetDIBHeight - 1
    
    'Just to be safe, make sure the radius isn't larger than the image itself
    If (finalY - initY) < (finalX - initX) Then
        If (mRadius > (finalY - initY)) Then mRadius = finalY - initY
    Else
        If (mRadius > (finalX - initX)) Then mRadius = finalX - initX
    End If
    
    If (mRadius < 1) Then mRadius = 1
        
    mPercent = mPercent * 0.01
    If (mPercent < 0.01) Then mPercent = 0.01
        
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim qvDepth As Long
    qvDepth = srcDIB.GetDIBColorDepth \ 8
    
    'The x-dimension of the image has a stride of (width * 4) for 32-bit images; precalculate this, to spare us some
    ' processing time in the inner loop.
    initX = initX * qvDepth
    finalX = finalX * qvDepth
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not suppressMessages) Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalX
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'The number of pixels in the current median box are tracked dynamically.
    Dim numOfPixels As Long
    numOfPixels = 0
            
    'We use an optimized histogram technique for calculating means, which means a lot of intermediate values are required
    Dim rValues() As Long, gValues() As Long, bValues() As Long, aValues() As Long
    ReDim rValues(0 To 255) As Long
    ReDim gValues(0 To 255) As Long
    ReDim bValues(0 To 255) As Long
    ReDim aValues(0 To 255) As Long
    
    Dim cutoffTotal As Long
    Dim r As Long, g As Long, b As Long
    Dim startY As Long, stopY As Long, yStep As Long, i As Long
    
    Dim directionDown As Boolean
    directionDown = True
    
    'Prep the pixel iterator
    Dim cPixelIterator As pdPixelIterator
    Set cPixelIterator = New pdPixelIterator
    
    If cPixelIterator.InitializeIterator(srcDIB, mRadius, mRadius, kernelShape) Then
    
        numOfPixels = cPixelIterator.LockTargetHistograms_RGBA(rValues, gValues, bValues, aValues, False)
        
        'Loop through each pixel in the image, applying the filter as we go
        For x = initX To finalX Step qvDepth
            
            'Based on the direction we're traveling, reverse the interior loop boundaries as necessary.
            If directionDown Then
                startY = initY
                stopY = finalY
                yStep = 1
            Else
                startY = finalY
                stopY = initY
                yStep = -1
            End If
            
            'Process the next column.  This step is pretty much identical to the row steps above (but in a vertical direction, obviously)
            For y = startY To stopY Step yStep
            
                'With a local histogram successfully built for the area surrounding this pixel, we now need to find the
                ' actual median value.
                
                'Loop through each color component histogram, until we've passed the desired percentile of pixels
                r = 0
                g = 0
                b = 0
                cutoffTotal = (mPercent * numOfPixels)
                If cutoffTotal = 0 Then cutoffTotal = 1
        
                i = -1
                Do
                    i = i + 1
                    r = r + rValues(i)
                Loop Until (r >= cutoffTotal)
                r = i

                i = -1
                Do
                    i = i + 1
                    g = g + gValues(i)
                Loop Until (g >= cutoffTotal)
                g = i

                i = -1
                Do
                    i = i + 1
                    b = b + bValues(i)
                Loop Until (b >= cutoffTotal)
                b = i
                
                'Finally, apply the results to the image.
                dstImageData(x, y) = b
                dstImageData(x + 1, y) = g
                dstImageData(x + 2, y) = r
                
                'Move the iterator in the correct direction
                If directionDown Then
                    If (y < finalY) Then numOfPixels = cPixelIterator.MoveYDown
                Else
                    If (y > initY) Then numOfPixels = cPixelIterator.MoveYUp
                End If
        
            Next y
            
            'Reverse y-directionality on each pass
            directionDown = Not directionDown
            If (x < finalX) Then numOfPixels = cPixelIterator.MoveXRight
            
            'Update the progress bar every (progBarCheck) lines
            If Not suppressMessages Then
                If (x And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal x + modifyProgBarOffset
                End If
            End If
            
        Next x
        
        'Release the pixel iterator
        cPixelIterator.ReleaseTargetHistograms_RGBA rValues, gValues, bValues, aValues
        
        'Release our local array that points to the target DIB
        CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
            
        If g_cancelCurrentAction Then CreateMedianDIB = 0 Else CreateMedianDIB = 1
    
    Else
        CreateMedianDIB = 0
    End If
    
End Function

'White balance a given DIB.
' Per PhotoDemon convention, this function will return a non-zero value if successful, and 0 if canceled.
Public Function WhiteBalanceDIB(ByVal percentIgnore As Double, ByRef srcDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long

    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte
    Dim tmpSA As SafeArray2D
    PrepSafeArray tmpSA, srcDIB
    CopyMemory ByVal VarPtrArray(imageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.GetDIBWidth - 1
    finalY = srcDIB.GetDIBHeight - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim quickVal As Long, qvDepth As Long
    qvDepth = srcDIB.GetDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not suppressMessages) Then
        If (modifyProgBarMax = -1) Then
            SetProgBarMax finalX
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'Color values
    Dim r As Long, g As Long, b As Long
    
    'Maximum and minimum values, which will be detected by our initial histogram run
    Dim rMax As Byte, gMax As Byte, bMax As Byte
    Dim rMin As Byte, gMin As Byte, bMin As Byte
    rMax = 0: gMax = 0: bMax = 0
    rMin = 255: gMin = 255: bMin = 255
    
    'Shrink the percentIgnore value down to 1% of the value we are passed (you'll see why in a moment)
    percentIgnore = percentIgnore * 0.01
    
    'Prepare histogram arrays
    Dim rCount(0 To 255) As Long, gCount(0 To 255) As Long, bCount(0 To 255) As Long
    For x = 0 To 255
        rCount(x) = 0
        gCount(x) = 0
        bCount(x) = 0
    Next x
    
    'Build the image histogram
    Dim startX As Long, stopX As Long
    startX = initX * qvDepth
    stopX = finalX * qvDepth
    
    For y = initY To finalY
    For x = startX To stopX Step qvDepth
        b = imageData(x, y)
        g = imageData(x + 1, y)
        r = imageData(x + 2, y)
        rCount(r) = rCount(r) + 1
        gCount(g) = gCount(g) + 1
        bCount(b) = bCount(b) + 1
    Next x
    Next y
    
     'With the histogram complete, we can now figure out how to stretch the RGB channels. We do this by calculating a min/max
    ' ratio where the top and bottom 0.05% (or user-specified value) of pixels are ignored.
    
    Dim foundYet As Boolean
    foundYet = False
    
    Dim numOfPixels As Long
    numOfPixels = (finalX + 1) * (finalY + 1)
    
    Dim wbThreshold As Long
    wbThreshold = numOfPixels * percentIgnore
    
    r = 0: g = 0: b = 0
    
    Dim rTally As Long, gTally As Long, bTally As Long
    rTally = 0: gTally = 0: bTally = 0
    
    'Find minimum values of red, green, and blue
    Do
        If (rCount(r) + rTally < wbThreshold) Then
            r = r + 1
            rTally = rTally + rCount(r)
        Else
            rMin = r
            foundYet = True
        End If
    Loop While foundYet = False
        
    foundYet = False
        
    Do
        If (gCount(g) + gTally < wbThreshold) Then
            g = g + 1
            gTally = gTally + gCount(g)
        Else
            gMin = g
            foundYet = True
        End If
    Loop While foundYet = False
    
    foundYet = False
    
    Do
        If (bCount(b) + bTally < wbThreshold) Then
            b = b + 1
            bTally = bTally + bCount(b)
        Else
            bMin = b
            foundYet = True
        End If
    Loop While foundYet = False
    
    'Now, find maximum values of red, green, and blue
    foundYet = False
    
    r = 255: g = 255: b = 255
    rTally = 0: gTally = 0: bTally = 0
    
    Do
        If (rCount(r) + rTally < wbThreshold) Then
            r = r - 1
            rTally = rTally + rCount(r)
        Else
            rMax = r
            foundYet = True
        End If
    Loop While foundYet = False
        
    foundYet = False
        
    Do
        If (gCount(g) + gTally < wbThreshold) Then
            g = g - 1
            gTally = gTally + gCount(g)
        Else
            gMax = g
            foundYet = True
        End If
    Loop While foundYet = False
    
    foundYet = False
    
    Do
        If (bCount(b) + bTally < wbThreshold) Then
            b = b - 1
            bTally = bTally + bCount(b)
        Else
            bMax = b
            foundYet = True
        End If
    Loop While foundYet = False
    
    'Finally, calculate the difference between max and min for each color
    Dim rDif As Long, gDif As Long, bDif As Long
    rDif = CLng(rMax) - CLng(rMin)
    gDif = CLng(gMax) - CLng(gMin)
    bDif = CLng(bMax) - CLng(bMin)
    
    'We can now build a final set of look-up tables that contain the results of every possible color transformation
    Dim rFinal(0 To 255) As Byte, gFinal(0 To 255) As Byte, bFinal(0 To 255) As Byte
    
    For x = 0 To 255
        If (rDif <> 0) Then r = 255# * ((x - rMin) / rDif) Else r = x
        If (gDif <> 0) Then g = 255# * ((x - gMin) / gDif) Else g = x
        If (bDif <> 0) Then b = 255# * ((x - bMin) / bDif) Else b = x
        If (r > 255) Then r = 255
        If (r < 0) Then r = 0
        If (g > 255) Then g = 255
        If (g < 0) Then g = 0
        If (b > 255) Then b = 255
        If (b < 0) Then b = 0
        rFinal(x) = r
        gFinal(x) = g
        bFinal(x) = b
    Next x
    
    'Now we can loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        quickVal = x * qvDepth
    For y = initY To finalY
            
        'Adjust white balance in a single pass (thanks to the magic of look-up tables)
        imageData(quickVal, y) = bFinal(imageData(quickVal, y))
        imageData(quickVal + 1, y) = gFinal(imageData(quickVal + 1, y))
        imageData(quickVal + 2, y) = rFinal(imageData(quickVal + 2, y))
        
    Next y
        If Not suppressMessages Then
            If (x And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal x + modifyProgBarOffset
            End If
        End If
    Next x
    
    'Safely deallocate imageData()
    CopyMemory ByVal VarPtrArray(imageData), 0&, 4
    
    If g_cancelCurrentAction Then WhiteBalanceDIB = 0 Else WhiteBalanceDIB = 1
    
End Function

'Contrast-correct a given DIB.  This function is similar to white-balance, except that it operates *only on luminance*, meaning individual
' color channel ratios are not changed - just luminance.  It's helpful for auto-spreading luminance across the full spectrum range, without
' changing color balance at all.
' Per PhotoDemon convention, this function will return a non-zero value if successful, and 0 if canceled.
Public Function ContrastCorrectDIB(ByVal percentIgnore As Double, ByRef srcDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long

    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte
    Dim tmpSA As SafeArray2D
    PrepSafeArray tmpSA, srcDIB
    CopyMemory ByVal VarPtrArray(imageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.GetDIBWidth - 1
    finalY = srcDIB.GetDIBHeight - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim quickVal As Long, qvDepth As Long
    qvDepth = srcDIB.GetDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalX
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'Color values
    Dim r As Long, g As Long, b As Long, grayVal As Long
    
    'Maximum and minimum values, which will be detected by our initial histogram run
    Dim lMax As Byte, lMin As Byte
    lMax = 0
    lMin = 255
    
    'Shrink the percentIgnore value down to 1% of the value we are passed (you'll see why in a moment)
    percentIgnore = percentIgnore / 100
    
    'Prepare a histogram array
    Dim lCount(0 To 255) As Long
    For x = 0 To 255
        lCount(x) = 0
    Next x
    
    'Build the image histogram
    For x = initX To finalX
        quickVal = x * qvDepth
    For y = initY To finalY
    
        b = imageData(quickVal, y)
        g = imageData(quickVal + 1, y)
        r = imageData(quickVal + 2, y)
        
        'Calculate a grayscale value using the original ITU-R recommended formula (BT.709, specifically)
        grayVal = (213 * r + 715 * g + 72 * b) * 0.001
        If grayVal > 255 Then grayVal = 255
        
        'Increment the histogram at this position
        lCount(grayVal) = lCount(grayVal) + 1
        
    Next y
    Next x
    
     'With the histogram complete, we can now figure out how to stretch the RGB channels. We do this by calculating a min/max
    ' ratio where the top and bottom 0.05% (or user-specified value) of pixels are ignored.
    Dim foundYet As Boolean
    foundYet = False
    
    Dim numOfPixels As Long
    numOfPixels = (finalX + 1) * (finalY + 1)
    
    Dim wbThreshold As Long
    wbThreshold = numOfPixels * percentIgnore
    
    grayVal = 0
    
    Dim lTally As Long
    lTally = 0
    
    'Find minimum and maximum luminance values in the current image
    Do
        If lCount(grayVal) + lTally < wbThreshold Then
            grayVal = grayVal + 1
            lTally = lTally + lCount(grayVal)
        Else
            lMin = grayVal
            foundYet = True
        End If
    Loop While foundYet = False
        
    foundYet = False
    
    grayVal = 255
    lTally = 0
    
    Do
        If lCount(grayVal) + lTally < wbThreshold Then
            grayVal = grayVal - 1
            lTally = lTally + lCount(grayVal)
        Else
            lMax = grayVal
            foundYet = True
        End If
    Loop While foundYet = False
    
    'Calculate the difference between max and min
    Dim lDif As Long
    lDif = CLng(lMax) - CLng(lMin)
    
    'Build a final set of look-up tables that contain the results of the requisite luminance transformation
    Dim lFinal(0 To 255) As Byte
    
    For x = 0 To 255
        If lDif <> 0 Then grayVal = 255 * ((x - lMin) / lDif) Else grayVal = x
        
        If grayVal > 255 Then grayVal = 255
        If grayVal < 0 Then grayVal = 0
        
        lFinal(x) = grayVal
        
    Next x
    
    'Now we can loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        quickVal = x * qvDepth
    For y = initY To finalY
            
        'Adjust white balance in a single pass (thanks to the magic of look-up tables)
        imageData(quickVal, y) = lFinal(imageData(quickVal, y))
        imageData(quickVal + 1, y) = lFinal(imageData(quickVal + 1, y))
        imageData(quickVal + 2, y) = lFinal(imageData(quickVal + 2, y))
        
    Next y
        If Not suppressMessages Then
            If (x And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal x + modifyProgBarOffset
            End If
        End If
    Next x
    
    'Safely deallocate imageData()
    CopyMemory ByVal VarPtrArray(imageData), 0&, 4
    
    If g_cancelCurrentAction Then ContrastCorrectDIB = 0 Else ContrastCorrectDIB = 1
    
End Function

'Given two DIBs, fill one with an artistically contoured (edge detect) version of the other.
' Per PhotoDemon convention, this function will return a non-zero value if successful, and 0 if canceled.
Public Function CreateContourDIB(ByVal blackBackground As Boolean, ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long
 
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SafeArray2D, dstSA1D As SafeArray1D
    PrepSafeArray dstSA, dstDIB
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent already embossed pixels from screwing up our results for later pixels.)
    Dim srcImageData() As Byte
    Dim srcSA As SafeArray2D, srcSA1D As SafeArray1D
    PrepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 1
    initY = 1
    finalX = srcDIB.GetDIBWidth - 2
    finalY = srcDIB.GetDIBHeight - 2
    
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim xOffset As Long, xOffsetRight As Long, xOffsetLeft As Long, qvDepth As Long
    qvDepth = srcDIB.GetDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not suppressMessages) Then
        If (modifyProgBarMax = -1) Then
            SetProgBarMax finalX
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'Color variables
    Dim rMin As Long, gMin As Long, bMin As Long
    Dim r As Long, g As Long, b As Long
    
    'Prep a one-dimensional safearray for the source image
    Dim srcBits As Long, srcStride As Long
    srcBits = srcDIB.GetDIBPointer
    srcStride = srcDIB.GetDIBStride
    
    With srcSA1D
        .cbElements = 1
        .cDims = 1
        .cLocks = 1
        .lBound = 0
        .cElements = srcDIB.GetDIBStride
    End With
    
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA1D), 4&
    
    '...and another one for the destination image
    Dim dstBits As Long, dstStride As Long
    dstBits = dstDIB.GetDIBPointer
    dstStride = dstDIB.GetDIBStride
    
    With dstSA1D
        .cbElements = 1
        .cDims = 1
        .cLocks = 1
        .lBound = 0
        .cElements = dstDIB.GetDIBStride
    End With
    
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA1D), 4&
    
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        xOffset = x * qvDepth
        xOffsetRight = (x + 1) * qvDepth
        xOffsetLeft = (x - 1) * qvDepth
    For y = initY To finalY
        
        'Find the smallest RGB values in the local vicinity of this pixel
        rMin = 255
        gMin = 255
        bMin = 255
        
        'Previous line
        srcSA1D.pvData = srcBits + (y - 1) * srcStride
        b = srcImageData(xOffsetLeft)
        g = srcImageData(xOffsetLeft + 1)
        r = srcImageData(xOffsetLeft + 2)
        If (b < bMin) Then bMin = b
        If (g < gMin) Then gMin = g
        If (r < rMin) Then rMin = r
        
        b = srcImageData(xOffset)
        g = srcImageData(xOffset + 1)
        r = srcImageData(xOffset + 2)
        If (b < bMin) Then bMin = b
        If (g < gMin) Then gMin = g
        If (r < rMin) Then rMin = r
        
        b = srcImageData(xOffsetRight)
        g = srcImageData(xOffsetRight + 1)
        r = srcImageData(xOffsetRight + 2)
        If (b < bMin) Then bMin = b
        If (g < gMin) Then gMin = g
        If (r < rMin) Then rMin = r
        
        'Current line
        srcSA1D.pvData = srcBits + y * srcStride
        b = srcImageData(xOffsetLeft)
        g = srcImageData(xOffsetLeft + 1)
        r = srcImageData(xOffsetLeft + 2)
        If (b < bMin) Then bMin = b
        If (g < gMin) Then gMin = g
        If (r < rMin) Then rMin = r
        
        b = srcImageData(xOffset)
        g = srcImageData(xOffset + 1)
        r = srcImageData(xOffset + 2)
        If (b < bMin) Then bMin = b
        If (g < gMin) Then gMin = g
        If (r < rMin) Then rMin = r
        
        b = srcImageData(xOffsetRight)
        g = srcImageData(xOffsetRight + 1)
        r = srcImageData(xOffsetRight + 2)
        If (b < bMin) Then bMin = b
        If (g < gMin) Then gMin = g
        If (r < rMin) Then rMin = r
        
        'Next line
        srcSA1D.pvData = srcBits + (y + 1) * srcStride
        b = srcImageData(xOffsetLeft)
        g = srcImageData(xOffsetLeft + 1)
        r = srcImageData(xOffsetLeft + 2)
        If (b < bMin) Then bMin = b
        If (g < gMin) Then gMin = g
        If (r < rMin) Then rMin = r
        
        b = srcImageData(xOffset)
        g = srcImageData(xOffset + 1)
        r = srcImageData(xOffset + 2)
        If (b < bMin) Then bMin = b
        If (g < gMin) Then gMin = g
        If (r < rMin) Then rMin = r
        
        b = srcImageData(xOffsetRight)
        g = srcImageData(xOffsetRight + 1)
        r = srcImageData(xOffsetRight + 2)
        If (b < bMin) Then bMin = b
        If (g < gMin) Then gMin = g
        If (r < rMin) Then rMin = r
        
        'Subtract the minimum value from the current pixel value
        srcSA1D.pvData = srcBits + y * srcStride
        dstSA1D.pvData = dstBits + y * dstStride
        
        If blackBackground Then
            dstImageData(xOffset) = srcImageData(xOffset) - bMin
            dstImageData(xOffset + 1) = srcImageData(xOffset + 1) - gMin
            dstImageData(xOffset + 2) = srcImageData(xOffset + 2) - rMin
        Else
            dstImageData(xOffset) = 255 - (srcImageData(xOffset) - bMin)
            dstImageData(xOffset + 1) = 255 - (srcImageData(xOffset + 1) - gMin)
            dstImageData(xOffset + 2) = 255 - (srcImageData(xOffset + 2) - rMin)
        End If
        
    Next y
        If (Not suppressMessages) Then
            If (x And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal x + modifyProgBarOffset
            End If
        End If
    Next x
    
    'The edges of the image will always be missed, so manually apply their values now
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Top row
    For x = initX To finalX
        xOffset = x * qvDepth
        dstImageData(xOffset, 0) = dstImageData(xOffset, 1)
        dstImageData(xOffset + 1, 0) = dstImageData(xOffset + 1, 1)
        dstImageData(xOffset + 2, 0) = dstImageData(xOffset + 2, 1)
    Next x
    
    'Bottom row
    y = finalY + 1
    For x = initX To finalX
        xOffset = x * qvDepth
        dstImageData(xOffset, y) = dstImageData(xOffset, finalY)
        dstImageData(xOffset + 1, y) = dstImageData(xOffset + 1, finalY)
        dstImageData(xOffset + 2, y) = dstImageData(xOffset + 2, finalY)
    Next x
    
    'Left row
    For y = initY To finalY
        dstImageData(0, y) = dstImageData(qvDepth, y)
        dstImageData(1, y) = dstImageData(qvDepth + 1, y)
        dstImageData(2, y) = dstImageData(qvDepth + 2, y)
    Next y
    
    'Right row
    xOffset = finalX * qvDepth
    xOffsetRight = (finalX + 1) * qvDepth
    For y = initY To finalY
        dstImageData(xOffsetRight, y) = dstImageData(xOffset, y)
        dstImageData(xOffsetRight + 1, y) = dstImageData(xOffset + 1, y)
        dstImageData(xOffsetRight + 2, y) = dstImageData(xOffset + 2, y)
    Next y
    
    'Safely deallocate all image arrays
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    
    If g_cancelCurrentAction Then CreateContourDIB = 0 Else CreateContourDIB = 1
    
End Function

'Make shadows, midtone, and/or highlight adjustments to a given DIB.
' Per PhotoDemon convention, this function will return a non-zero value if successful, and 0 if canceled.
Public Function AdjustDIBShadowHighlight(ByVal shadowAmount As Double, ByVal midtoneAmount As Double, ByVal highlightAmount As Double, ByVal shadowWidth As Long, ByVal shadowRadius As Double, ByVal highlightWidth As Long, ByVal highlightRadius As Double, ByRef srcDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long
    
    'As of March 2015, this function has been entirely rewritten, using a system similar to PhotoShop's
    ' (I think... but it's impossible to know for sure, since I don't have a copy for testing!
    '  Theoretically it should be very close.)
    '
    'This overhaul greatly improved the usefulness of this tool, but because it completed changed the
    ' input parameters, ranges, and UI of the associated dialog, it is incompatible with past versions
    ' of the tool.  As such, the processor call that wraps this function has been changed to prevent
    ' conflicts with old macros.
        
    'Start by converting input parameters to desired ranges.
    shadowAmount = shadowAmount / 100#
    highlightAmount = -1 * (highlightAmount * 0.01)
    midtoneAmount = -1 * (midtoneAmount * 0.01)
    
    'Also, make absolute-value copies of the amount input.  (This is faster than constantly re-calculating
    ' absolute values inside the per-pixel adjustment loops.)
    Dim absShadowAmount As Double, absHighlightAmount As Double, absMidtoneAmount As Double
    absShadowAmount = Abs(shadowAmount)
    absHighlightAmount = Abs(highlightAmount)
    absMidtoneAmount = Abs(midtoneAmount)
    
    'From here, processing becomes more intensive.  Prep the progress bar as necessary.
    If (Not suppressMessages) Then
        If (modifyProgBarMax < 0) Then SetProgBarMax 6 Else SetProgBarMax modifyProgBarMax
        SetProgBarVal modifyProgBarOffset
    End If
    
    'Next we will create shadow, midtone, and highlight lookup tables.  These will simplify the process of
    ' identifying luminance regions in the base image.
    
    'These lookup tables will be Single-type, and they will contain a value on the range [0, 1] for each
    ' 8-bit channel value [0, 255].  0 signifies a lookup entry outside that range, while 1 indicates a
    ' value fully within the target range.  Some feathering is used to make the transition between ranges
    ' appear more natural.  (The feathering used is a place where it would be really nice to have PhotoShop
    ' for comparisons, as I'm curious how they blend between shadow/midtone/highlight ranges...)
    Dim sLookup() As Single, mLookup() As Single, hLookup() As Single
    ReDim sLookup(0 To 255) As Single
    ReDim mLookup(0 To 255) As Single
    ReDim hLookup(0 To 255) As Single
    
    'Before generating the tables, generate shadow and highlight cut-offs, using the values supplied by the user.
    Dim sCutoff As Long, hCutoff As Long
    sCutoff = shadowWidth
    If (sCutoff = 0) Then sCutoff = 1
    
    hCutoff = 255 - highlightWidth
    If (hCutoff = 255) Then hCutoff = 254
    
    'Next, automatically determine midtone cut-offs, using the supplied shadow/highlight values as our guide
    Dim mCutoffLow As Long, mCutoffHigh As Long, mRange As Long, mMidpoint As Long
    mCutoffLow = sCutoff
    mCutoffHigh = hCutoff
    
    'If artificially low shadow/highlight ranges are used, shrink midtones accordingly
    If (mCutoffLow < 64) Then mCutoffLow = 64
    If (mCutoffHigh > 192) Then mCutoffHigh = 192
    mRange = mCutoffHigh - mCutoffLow
    mMidpoint = (mRange \ 2)
    
    Dim tmpCalc As Double
    
    'Now we can generate lookup tables
    Dim i As Long
    For i = 0 To 255
    
        'Shadows use a power curve maximized at 0, and descending toward the cutoff point
        If (i < sCutoff) Then
            tmpCalc = i / sCutoff
            tmpCalc = tmpCalc * tmpCalc
            sLookup(i) = 1 - tmpCalc
        End If
        
        'Highlights use a power curve maximized at 255, and descending toward the cutoff point
        If (i > hCutoff) Then
            tmpCalc = (255 - i) / (255 - hCutoff)
            tmpCalc = tmpCalc * tmpCalc
            hLookup(i) = 1 - tmpCalc
        End If
        
        'Midtones use a bell curve stretching between mCutoffLow and mCutoffHigh
        If (i > mCutoffLow) And (i < mCutoffHigh) Then
            tmpCalc = (i - mCutoffLow)
            tmpCalc = 1 - (tmpCalc / mMidpoint)
            tmpCalc = tmpCalc * tmpCalc
            mLookup(i) = 1 - tmpCalc
        End If
    
    Next i
    
    'With shadow, midtone, and highlight ranges now established, we can start applying the user's changes.
    
    If (Not suppressMessages) Then SetProgBarVal modifyProgBarOffset + 1
    
    'First, if the shadow and highlight regions have different radius values, we need to make a backup copy
    ' of the current DIB.
    Dim backupDIB As pdDIB
    
    If (shadowRadius <> highlightRadius) Or (shadowAmount = 0#) Then
        Set backupDIB = New pdDIB
        backupDIB.CreateFromExistingDIB srcDIB
    End If
    
    'Next, we need to make a duplicate copy of the source image.  To improve output, this copy will be blurred,
    ' and we will use it to identify shadow/highlight regions.  (The blur naturally creates smoother transitions
    ' between light and dark parts of the image.)
    Dim blurDIB As pdDIB
    Set blurDIB = New pdDIB
    blurDIB.CreateFromExistingDIB srcDIB
    
    'Shadows are handled first.  If the user requested a radius > 0, blur the reference image now.
    If (shadowAmount <> 0#) And (shadowRadius > 0) Then QuickBlurDIB blurDIB, shadowRadius, False
        
    'Unfortunately, the next step of the operation requires manual pixel-by-pixel blending.  Prep all required
    ' loop objects now.
    If (Not suppressMessages) Then SetProgBarVal modifyProgBarOffset + 2
    
    'Create local arrays and point them at the source DIB and blurred DIB
    Dim srcImageData() As Byte, blurImageData() As Byte
    Dim srcSA As SafeArray2D, blurSA As SafeArray2D
    
    PrepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    PrepSafeArray blurSA, blurDIB
    CopyMemory ByVal VarPtrArray(blurImageData()), VarPtr(blurSA), 4
        
    'Prep local loop variables
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.GetDIBWidth - 1
    finalY = srcDIB.GetDIBHeight - 1
            
    'Prep stride ofsets.  (This is required because the image array may be 24 or 32 bits per pixel, and we want
    ' to handle both cases.)
    Dim quickX As Long, qvDepth As Long
    qvDepth = srcDIB.GetDIBColorDepth \ 8
    
    'Prep color retrieval variables (Long-type, because intermediate calculates may exceed byte range)
    Dim rSrc As Double, gSrc As Double, bSrc As Double
    Dim rDst As Long, gDst As Long, bDst As Long
    Dim rBlur As Double, gBlur As Double, bBlur As Double
    Dim srcBlur As Long, grayBlur As Long
    Dim pxShadowCorrection As Double, pxHighlightCorrection As Double, pxMidtoneCorrection As Double
    Const ONE_DIV_255 As Double = 1# / 255#
    
    'Start processing shadow pixels
    If (shadowAmount <> 0) Then
    
        For x = initX To finalX
            quickX = x * qvDepth
        For y = initY To finalY
            
            'Calculate luminance for this pixel in the *blurred* image.  (We use the blurred copy for luminance
            ' detection, to improve transitions between light and dark regions in the image.)
            bBlur = blurImageData(quickX, y)
            gBlur = blurImageData(quickX + 1, y)
            rBlur = blurImageData(quickX + 2, y)
            
            grayBlur = (213 * rBlur + 715 * gBlur + 72 * bBlur) \ 1000
            If (grayBlur > 255) Then grayBlur = 255
            
            'If the luminance of this pixel falls within the shadow range, continue processing; otherwise, ignore it and
            ' move on to the next pixel.
            If (sLookup(grayBlur) > 0!) Then
                
                'Invert the blur pixel values, and convert to the range [0, 1]
                If (shadowAmount > 0) Then
                    rBlur = 1# - (rBlur * ONE_DIV_255)
                    gBlur = 1# - (gBlur * ONE_DIV_255)
                    bBlur = 1# - (bBlur * ONE_DIV_255)
                Else
                    rBlur = (rBlur * ONE_DIV_255)
                    gBlur = (gBlur * ONE_DIV_255)
                    bBlur = (bBlur * ONE_DIV_255)
                End If
                
                'Retrieve source pixel values and convert to the range [0, 1]
                bSrc = srcImageData(quickX, y)
                gSrc = srcImageData(quickX + 1, y)
                rSrc = srcImageData(quickX + 2, y)
                
                rSrc = rSrc * ONE_DIV_255
                gSrc = gSrc * ONE_DIV_255
                bSrc = bSrc * ONE_DIV_255
                
                'Calculate a maximum strength adjustment value.
                ' (This code is actually just the Overlay compositor formula.)
                If (rSrc < 0.5) Then rBlur = 2# * rSrc * rBlur Else rBlur = 1# - 2# * (1# - rSrc) * (1# - rBlur)
                If (gSrc < 0.5) Then gBlur = 2# * gSrc * gBlur Else gBlur = 1# - 2# * (1# - gSrc) * (1# - gBlur)
                If (bSrc < 0.5) Then bBlur = 2# * bSrc * bBlur Else bBlur = 1# - 2# * (1# - bSrc) * (1# - bBlur)
                
                'Calculate a final shadow correction amount, which is a combination of...
                ' 1) The user-supplied shadow correction amount
                ' 2) The shadow lookup table for this value
                pxShadowCorrection = absShadowAmount * sLookup(grayBlur)
                
                'Modify the maximum strength adjustment value by the user-supplied shadow correction amount
                bDst = 255 * ((pxShadowCorrection * bBlur) + ((1# - pxShadowCorrection) * bSrc))
                gDst = 255 * ((pxShadowCorrection * gBlur) + ((1# - pxShadowCorrection) * gSrc))
                rDst = 255 * ((pxShadowCorrection * rBlur) + ((1# - pxShadowCorrection) * rSrc))
                
                'Save the modified values into the source image
                srcImageData(quickX, y) = bDst
                srcImageData(quickX + 1, y) = gDst
                srcImageData(quickX + 2, y) = rDst
                
            End If
            
        Next y
            If Not suppressMessages Then
                If (x And 63) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                End If
            End If
        Next x
        
    End If
    
    'With our shadow work complete, point all local arrays away from their respective DIBs
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    CopyMemory ByVal VarPtrArray(blurImageData), 0&, 4
    
    If (Not suppressMessages) Then SetProgBarVal modifyProgBarOffset + 3
    
    'Next, it's time to operate on highlights.  The steps involved are pretty much identical to shadows, but we obviously
    ' use the highlight lookup table to determine valid correction candidates.
    If (highlightAmount <> 0) And (Not g_cancelCurrentAction) Then
    
        'Before starting per-pixel processing, see if a highlight radius was specified.  If it was, and the radius differs
        ' from the shadow radius, calculate a new blur DIB now.
        If (highlightRadius <> shadowRadius) Or (shadowAmount = 0#) Then
            
            blurDIB.CreateFromExistingDIB backupDIB
            If (highlightRadius <> 0) Then QuickBlurDIB blurDIB, highlightRadius, False
            
            'Note that we can now free our backup DIB, as it's no longer needed
            Set backupDIB = Nothing
            
        End If
        
        If (Not suppressMessages) Then SetProgBarVal modifyProgBarOffset + 4
        
        'Once again, point arrays at both the source and blur DIBs
        PrepSafeArray srcSA, srcDIB
        CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
            
        PrepSafeArray blurSA, blurDIB
        CopyMemory ByVal VarPtrArray(blurImageData()), VarPtr(blurSA), 4
        
        'Start per-pixel highlight processing!
        For x = initX To finalX
            quickX = x * qvDepth
        For y = initY To finalY
            
            'Calculate luminance for this pixel in the *blurred* image.  (We use the blurred copy for luminance detection, to improve
            ' transitions between light and dark regions in the image.)
            bBlur = blurImageData(quickX, y)
            gBlur = blurImageData(quickX + 1, y)
            rBlur = blurImageData(quickX + 2, y)
            
            grayBlur = (213 * rBlur + 715 * gBlur + 72 * bBlur) \ 1000
            If (grayBlur > 255) Then grayBlur = 255
            
            'If the luminance of this pixel falls within the highlight range, continue processing; otherwise, ignore it and
            ' move on to the next pixel.
            If (hLookup(grayBlur) > 0!) Then
                
                'Invert the blur pixel values, and convert to the range [0, 1]
                If (highlightAmount > 0#) Then
                    rBlur = 1# - (rBlur * ONE_DIV_255)
                    gBlur = 1# - (gBlur * ONE_DIV_255)
                    bBlur = 1# - (bBlur * ONE_DIV_255)
                Else
                    rBlur = (rBlur * ONE_DIV_255)
                    gBlur = (gBlur * ONE_DIV_255)
                    bBlur = (bBlur * ONE_DIV_255)
                End If
                
                'Retrieve source pixel values and convert to the range [0, 1]
                bSrc = srcImageData(quickX, y)
                gSrc = srcImageData(quickX + 1, y)
                rSrc = srcImageData(quickX + 2, y)
                
                rSrc = rSrc * ONE_DIV_255
                gSrc = gSrc * ONE_DIV_255
                bSrc = bSrc * ONE_DIV_255
                
                'Calculate a maximum strength adjustment value.
                ' (This code is actually just the Overlay compositor formula.)
                If (rSrc < 0.5) Then rBlur = 2# * rSrc * rBlur Else rBlur = 1# - 2# * (1# - rSrc) * (1# - rBlur)
                If (gSrc < 0.5) Then gBlur = 2# * gSrc * gBlur Else gBlur = 1# - 2# * (1# - gSrc) * (1# - gBlur)
                If (bSrc < 0.5) Then bBlur = 2# * bSrc * bBlur Else bBlur = 1# - 2# * (1# - bSrc) * (1# - bBlur)
                
                'Calculate a final highlight correction amount, which is a combination of...
                ' 1) The user-supplied highlight correction amount
                ' 2) The highlight lookup table for this value
                pxHighlightCorrection = absHighlightAmount * hLookup(grayBlur)
                
                'Modify the maximum strength adjustment value by the user-supplied highlight correction amount
                bDst = 255 * ((pxHighlightCorrection * bBlur) + ((1# - pxHighlightCorrection) * bSrc))
                gDst = 255 * ((pxHighlightCorrection * gBlur) + ((1# - pxHighlightCorrection) * gSrc))
                rDst = 255 * ((pxHighlightCorrection * rBlur) + ((1# - pxHighlightCorrection) * rSrc))
                
                'Save the modified values into the source image
                srcImageData(quickX, y) = bDst
                srcImageData(quickX + 1, y) = gDst
                srcImageData(quickX + 2, y) = rDst
                
            End If
            
        Next y
            If Not suppressMessages Then
                If (x And 63) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                End If
            End If
        Next x
        
        'With our highlight work complete, point all local arrays away from their respective DIBs
        CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
        CopyMemory ByVal VarPtrArray(blurImageData), 0&, 4
    
    End If
    
    If (Not suppressMessages) Then SetProgBarVal modifyProgBarOffset + 5
    
    'We are now done with the blur DIB, so let's free it regardless of what comes next
    Set blurDIB = Nothing
    
    'Last up is midtone correction.  The steps involved are pretty much identical to shadow and highlight correction, but we obviously
    ' use the midtone lookup table to determine valid correction candidates.  (Also, we do not use a blurred copy of the DIB.)
    If (midtoneAmount <> 0) And (Not g_cancelCurrentAction) Then
        
        'Once again, point an array at the source DIB
        PrepSafeArray srcSA, srcDIB
        CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
        'Start per-pixel midtone processing!
        For x = initX To finalX
            quickX = x * qvDepth
        For y = initY To finalY
            
            'Calculate luminance for this pixel in the *source* image.
            bSrc = srcImageData(quickX, y)
            gSrc = srcImageData(quickX + 1, y)
            rSrc = srcImageData(quickX + 2, y)
            
            srcBlur = (213 * rSrc + 715 * gSrc + 72 * bSrc) \ 1000
            If (srcBlur > 255) Then srcBlur = 255
            
            'If the luminance of this pixel falls within the highlight range, continue processing; otherwise, ignore it and
            ' move on to the next pixel.
            If (mLookup(srcBlur) > 0!) Then
                
                'Convert the source pixel values to the range [0, 1]
                bSrc = bSrc * ONE_DIV_255
                gSrc = gSrc * ONE_DIV_255
                rSrc = rSrc * ONE_DIV_255
                
                'To cut down on the need for additional local variables, we're going to simply re-use the blur variable names here.
                If (midtoneAmount > 0) Then
                    rBlur = 1# - rSrc
                    gBlur = 1# - gSrc
                    bBlur = 1# - bSrc
                Else
                    rBlur = rSrc
                    gBlur = gSrc
                    bBlur = bSrc
                End If
                
                'Calculate a maximum strength adjustment value.
                ' (This code is actually just the Overlay compositor formula.)
                If (rSrc < 0.5) Then rBlur = 2# * rSrc * rBlur Else rBlur = 1# - 2# * (1# - rSrc) * (1# - rBlur)
                If (gSrc < 0.5) Then gBlur = 2# * gSrc * gBlur Else gBlur = 1# - 2# * (1# - gSrc) * (1# - gBlur)
                If (bSrc < 0.5) Then bBlur = 2# * bSrc * bBlur Else bBlur = 1# - 2# * (1# - bSrc) * (1# - bBlur)
                
                'Calculate a final midtone correction amount, which is a combination of...
                ' 1) The user-supplied midtone correction amount
                ' 2) The midtone lookup table for this value
                pxMidtoneCorrection = absMidtoneAmount * mLookup(srcBlur)
                
                'Modify the maximum strength adjustment value by the user-supplied midtone correction amount
                bDst = 255 * ((pxMidtoneCorrection * bBlur) + ((1# - pxMidtoneCorrection) * bSrc))
                gDst = 255 * ((pxMidtoneCorrection * gBlur) + ((1# - pxMidtoneCorrection) * gSrc))
                rDst = 255 * ((pxMidtoneCorrection * rBlur) + ((1# - pxMidtoneCorrection) * rSrc))
                
                'Save the modified values into the source image
                srcImageData(quickX, y) = bDst
                srcImageData(quickX + 1, y) = gDst
                srcImageData(quickX + 2, y) = rDst
                
            End If
            
        Next y
            If Not suppressMessages Then
                If (x And 63) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                End If
            End If
        Next x
        
        'With our highlight work complete, point all local arrays away from their respective DIBs
        CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
        
    End If
    
    If (Not suppressMessages) Then SetProgBarVal modifyProgBarOffset + 6
    
    If g_cancelCurrentAction Then AdjustDIBShadowHighlight = 0 Else AdjustDIBShadowHighlight = 1
    
End Function

'Given two DIBs, fill one with an approximated gaussian-blur version of the other.
' Per the Central Limit Theorem, a Gaussian function can be approximated within 3% by three iterations of a matching box function.
' Gaussian blur and a 3x box blur are thus "roughly" identical, but there are some trade-offs - PD's Gaussian Blur uses a modified
' standard deviation function, which results in a higher-quality blur.  It also supports floating-point radii.  Both these options
' are lost if this approximate function is used.  That said, the performance trade-off (20x faster in most cases) is well worth it
' for all but the most stringent blur needs.
'
' Per PhotoDemon convention, this function will return a non-zero value if successful, and 0 if canceled.
Public Function CreateApproximateGaussianBlurDIB(ByVal equivalentGaussianRadius As Double, ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, Optional ByVal numIterations As Long = 3, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (modifyProgBarMax = -1) Then modifyProgBarMax = srcDIB.GetDIBWidth * numIterations + srcDIB.GetDIBHeight * numIterations
    If (Not suppressMessages) Then SetProgBarMax modifyProgBarMax
    
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'Modify the Gaussian radius, and convert it to an integer.  (Box blurs don't work on floating-point radii.)
    Dim comparableRadius As Long
    
    'If the number of iterations = 3, we can approximate a correct radius using a piece-wise quadratic convolution kernel.
    ' This should result in a kernel that's ~97% identical to a Gaussian kernel.  For a more in-depth explanation of
    ' converting between standard deviation and a box blur estimation, please see this W3 spec:
    ' http://www.w3.org/TR/SVG11/filters.html#feGaussianBlurElement
    If (numIterations = 3) Then
        Dim stdDev As Double
        stdDev = Sqr(-(equivalentGaussianRadius * equivalentGaussianRadius) / (2# * Log(1# / 255#)))
        comparableRadius = Int(stdDev * 2.37997232 * 0.5 + 0.5)
    Else
    
        'For larger iterations, it's not worth the trouble to perform a fine estimation, as repeat iterations
        ' eliminate small discrepancies.  Use a quick-and-dirty computation to calculate radius:
        comparableRadius = Int(equivalentGaussianRadius / numIterations + 0.5)
    
    End If
    
    'Box blurs require a radius of at least 1, so force it to that
    If (comparableRadius < 1) Then comparableRadius = 1
    
    'We now want to populate an array of radii, each of which will be used at one of the iterations.  We do this
    ' so that we can more accurately imitate the original gaussian blur with a series of box blurs.
    Dim radiiTable() As Long
    ReDim radiiTable(0 To numIterations - 1) As Long
    
    Dim netRadii As Long, thisRadii As Long
    
    Dim i As Long
    For i = 0 To numIterations - 1
        
        'For all (n-1) iterations, use the comparable radius calculated above
        If (i < numIterations - 1) Then
            thisRadii = comparableRadius
        
        'For the final iteration, use the remainder (e.g. the difference between how many radii we've calculated,
        ' and how many radii remain to be calculated)
        Else
            
            thisRadii = (equivalentGaussianRadius - netRadii)
            
            'If the remainder is <= 0, simply remove the final iteration of the blur to compensate
            If (thisRadii < 1) Then
                thisRadii = 1
                numIterations = numIterations - 1
                If (numIterations < 1) Then numIterations = 1
            End If
            
        End If
        
        radiiTable(i) = thisRadii
        
        netRadii = netRadii + thisRadii
        
    Next i
    
    'Create an extra intermediate DIB.  This is needed to cache the results of the horizontal blur, before we apply the vertical pass to it.
    Dim gaussDIB As pdDIB
    Set gaussDIB = New pdDIB
    gaussDIB.CreateFromExistingDIB srcDIB
    dstDIB.CreateFromExistingDIB gaussDIB
    
    'Iterate a box blur, switching between the gauss and destination DIBs as we go
    For i = 0 To numIterations - 1
        
        If (radiiTable(i) > 0) Then
            If (CreateHorizontalBlurDIB(radiiTable(i), radiiTable(i), dstDIB, gaussDIB, suppressMessages, modifyProgBarMax, modifyProgBarOffset + (gaussDIB.GetDIBWidth * i) + (gaussDIB.GetDIBHeight * i)) > 0) Then
                If (CreateVerticalBlurDIB(radiiTable(i), radiiTable(i), gaussDIB, dstDIB, suppressMessages, modifyProgBarMax, modifyProgBarOffset + (gaussDIB.GetDIBWidth * (i + 1)) + (gaussDIB.GetDIBHeight * i)) = 0) Then Exit For
            Else
                Exit For
            End If
        End If
    
    Next i
    
    'Erase the temporary DIB and exit
    gaussDIB.EraseDIB
    Set gaussDIB = Nothing
    
    If g_cancelCurrentAction Then CreateApproximateGaussianBlurDIB = 0 Else CreateApproximateGaussianBlurDIB = 1

End Function

'Given two DIBs, fill one with a polar-coordinate conversion of the other.
' Per PhotoDemon convention, this function will return a non-zero value if successful, and 0 if canceled.
Public Function CreatePolarCoordDIB(ByVal conversionMethod As Long, ByVal polarRadius As Double, ByVal edgeHandling As Long, ByVal useBilinear As Boolean, ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long

    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SafeArray2D
    PrepSafeArray dstSA, dstDIB
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent medianred pixel values from spreading across the image as we go.)
    Dim srcImageData() As Byte
    Dim srcSA As SafeArray2D
    PrepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.GetDIBWidth - 1
    finalY = srcDIB.GetDIBHeight - 1
        
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim quickVal As Long, qvDepth As Long
    qvDepth = srcDIB.GetDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not suppressMessages) Then
        If (modifyProgBarMax = -1) Then SetProgBarMax finalX Else SetProgBarMax modifyProgBarMax
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'Create a filter support class, which will aid with edge handling and interpolation
    Dim fSupport As pdFilterSupport
    Set fSupport = New pdFilterSupport
    fSupport.SetDistortParameters qvDepth, edgeHandling, useBilinear, finalX, finalY
    
    'Polar conversion requires a number of specialized variables
    
    'Calculate the center of the image
    Dim midX As Double, midY As Double
    midX = CDbl(finalX - initX) * 0.5
    midX = midX + initX
    midY = CDbl(finalY - initY) * 0.5
    midY = midY + initY
    
    'Rotation values
    Dim theta As Double, sRadius As Double, sRadius2 As Double, sDistance As Double
    Dim r As Double, t As Double
    
    'X and Y values, remapped around a center point of (0, 0)
    Dim nX As Double, nY As Double
    
    'Source X and Y values, which may or may not be used as part of a bilinear interpolation function
    Dim srcX As Double, srcY As Double
        
    'Max radius is calculated as the distance from the center of the image to a corner
    Dim tWidth As Long, tHeight As Long
    tWidth = finalX - initX
    tHeight = finalY - initY
    sRadius = Sqr(tWidth * tWidth + tHeight * tHeight) / 2
              
    sRadius = sRadius * (polarRadius / 100#)
    If (sRadius < 1E-20) Then sRadius = 1E-20
    sRadius2 = sRadius * sRadius
    
    polarRadius = 1# / (polarRadius / 100#)
    
    'Check for extremely small images and exit, to avoid OOB problems
    If (tWidth <= 1) Or (tHeight <= 1) Then
        CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
        CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
        CreatePolarCoordDIB = 1
        Exit Function
    End If
    
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        quickVal = x * qvDepth
    For y = initY To finalY
    
        'Each polar conversion requires a unique set of code
        Select Case conversionMethod
        
            'Rectangular to polar
            Case 0
                            
                'Remap the coordinates around a center point of (0, 0)
                nX = x - midX
                nY = y - midY
                
                'Calculate distance automatically
                sDistance = (nX * nX) + (nY * nY)
                
                If (sDistance <= sRadius2) Then
                
                    'X is handled differently based on its relation to the center of the image
                    If x >= midX Then
                        nX = x - midX
                        If y > midY Then
                            theta = PI - Atn(nX / nY)
                            r = Sqr(sDistance)
                        ElseIf y < midY Then
                            theta = Atn(nX / (midY - y))
                            r = Sqr(nX * nX + (midY - y) * (midY - y))
                        Else
                            theta = PI_HALF
                            r = nX
                        End If
                    Else
                        nX = midX - x
                        If y > midY Then
                            theta = PI + Atn(nX / nY)
                            r = Sqr(sDistance)
                        ElseIf y < midY Then
                            theta = PI_DOUBLE - Atn(nX / (midY - y))
                            r = Sqr(nX * nX + (midY - y) * (midY - y))
                        Else
                            theta = PI * 1.5
                            r = nX
                        End If
                    End If
                                        
                    srcX = finalX - (finalX / PI_DOUBLE * theta)
                    srcY = finalY * (r / sRadius)
                    
                Else
                
                    srcX = x
                    srcY = y
                    
                End If
                
            'Polar to rectangular
            Case 1
            
                'Remap the coordinates around a center point of (0, 0)
                nX = x - midX
                nY = y - midY
                
                'Calculate distance automatically
                sDistance = (nX * nX) + (nY * nY)
            
                If sDistance <= sRadius2 Then
                
                    theta = (x / finalX) * PI_DOUBLE
                    
                    If theta >= (PI * 1.5) Then
                        t = PI_DOUBLE - theta
                    ElseIf theta >= PI Then
                        t = theta - PI
                    ElseIf theta > PI_HALF Then
                        t = PI - theta
                    Else
                        t = theta
                    End If
                    
                    r = sRadius * (y / finalY)
                    
                    nX = -r * Sin(t)
                    nY = r * Cos(t)
                    
                    If theta >= 1.5 * PI Then
                        srcX = midX - nX
                        srcY = midY - nY
                    ElseIf theta >= PI Then
                        srcX = midX - nX
                        srcY = midY + nY
                    ElseIf theta >= PI_HALF Then
                        srcX = midX + nX
                        srcY = midY + nY
                    Else
                        srcX = midX + nX
                        srcY = midY - nY
                    End If
                    
                Else
                
                    srcX = x
                    srcY = y
                
                End If
                            
            'Polar inversion
            Case 2
            
                'Remap the coordinates around a center point of (0, 0)
                nX = x - midX
                nY = y - midY
                
                'Calculate distance automatically
                sDistance = (nX * nX) + (nY * nY)
                
                If (sDistance <> 0#) Then
                    srcX = midX + midX * midX * (nX / sDistance) * polarRadius
                    srcY = midY + midY * midY * (nY / sDistance) * polarRadius
                    srcX = PDMath.Modulo(srcX, finalX)
                    srcY = PDMath.Modulo(srcY, finalY)
                Else
                    srcX = x
                    srcY = y
                End If
            
        End Select
        
        'The lovely .setPixels routine will handle edge detection and interpolation for us as necessary
        fSupport.SetPixels x, y, srcX, srcY, srcImageData, dstImageData
                
    Next y
        If Not suppressMessages Then
            If (x And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal x + modifyProgBarOffset
            End If
        End If
    Next x
    
    'Safely deallocate all image arrays
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    
    If g_cancelCurrentAction Then CreatePolarCoordDIB = 0 Else CreatePolarCoordDIB = 1

End Function

'Given two DIBs, fill one with a polar-coordinate conversion of the other.
' Per PhotoDemon convention, this function will return a non-zero value if successful, and 0 if canceled.
' NOTE: unlike the traditional polar conversion function above, this one swaps x and y values.  There is no canonical definition for
'       how to polar convert an image, so we allow the user to choose whichever method they prefer.
Public Function CreateXSwappedPolarCoordDIB(ByVal conversionMethod As Long, ByVal polarRadius As Double, ByVal edgeHandling As Long, ByVal useBilinear As Boolean, ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long

    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SafeArray2D
    PrepSafeArray dstSA, dstDIB
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent medianred pixel values from spreading across the image as we go.)
    Dim srcImageData() As Byte
    Dim srcSA As SafeArray2D
    PrepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.GetDIBWidth - 1
    finalY = srcDIB.GetDIBHeight - 1
        
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim quickVal As Long, qvDepth As Long
    qvDepth = srcDIB.GetDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalX
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'Create a filter support class, which will aid with edge handling and interpolation
    Dim fSupport As pdFilterSupport
    Set fSupport = New pdFilterSupport
    fSupport.SetDistortParameters qvDepth, edgeHandling, useBilinear, finalX, finalY
    
    'Polar conversion requires a number of specialized variables
    
    'Calculate the center of the image
    Dim midX As Double, midY As Double
    midX = CDbl(finalX - initX) / 2
    midX = midX + initX
    midY = CDbl(finalY - initY) / 2
    midY = midY + initY
    
    'Rotation values
    Dim theta As Double, sRadius As Double, sRadius2 As Double, sDistance As Double
    Dim r As Double, t As Double
    
    'X and Y values, remapped around a center point of (0, 0)
    Dim nX As Double, nY As Double
    
    'Source X and Y values, which may or may not be used as part of a bilinear interpolation function
    Dim srcX As Double, srcY As Double
        
    'Max radius is calculated as the distance from the center of the image to a corner
    Dim tWidth As Long, tHeight As Long
    tWidth = finalX - initX
    tHeight = finalY - initY
    sRadius = Sqr(tWidth * tWidth + tHeight * tHeight) / 2
              
    sRadius = sRadius * (polarRadius / 100#)
    If (sRadius < 1E-20) Then sRadius = 1E-20
    sRadius2 = sRadius * sRadius
        
    polarRadius = 1# / (polarRadius / 100#)
    
    'Check for extremely small images and exit, to avoid OOB problems
    If (tWidth <= 1) Or (tHeight <= 1) Then
        CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
        CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
        CreateXSwappedPolarCoordDIB = 1
        Exit Function
    End If
    
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        quickVal = x * qvDepth
    For y = initY To finalY
    
        'Each polar conversion requires a unique set of code
        Select Case conversionMethod
        
            'Rectangular to polar
            Case 0
                            
                'Remap the coordinates around a center point of (0, 0)
                nX = x - midX
                nY = y - midY
                
                'Calculate distance automatically
                sDistance = (nX * nX) + (nY * nY)
                
                If sDistance <= sRadius2 Then
                
                    'X is handled differently based on its relation to the center of the image
                    If y >= midY Then
                        nY = y - midY
                        If x > midX Then
                            theta = PI - Atn(nY / nX)
                            r = Sqr(sDistance)
                        ElseIf x < midX Then
                            theta = Atn(nY / (midX - x))
                            r = Sqr(nY * nY + (midX - x) * (midX - x))
                        Else
                            theta = PI_HALF
                            r = nY
                        End If
                    Else
                        nY = midY - y
                        If x > midX Then
                            theta = PI + Atn(nY / nX)
                            r = Sqr(sDistance)
                        ElseIf x < midX Then
                            theta = PI_DOUBLE - Atn(nY / (midX - x))
                            r = Sqr(nY * nY + (midX - x) * (midX - x))
                        Else
                            theta = PI * 1.5
                            r = nY
                        End If
                    End If
                                        
                    srcY = finalY - (finalY / PI_DOUBLE * theta)
                    srcX = finalX * (r / sRadius)
                    
                Else
                
                    srcX = x
                    srcY = y
                    
                End If
                
            'Polar to rectangular
            Case 1
            
                'Remap the coordinates around a center point of (0, 0)
                nX = x - midX
                nY = y - midY
                
                'Calculate distance automatically
                sDistance = (nX * nX) + (nY * nY)
            
                If sDistance <= sRadius2 Then
                
                    theta = (y / finalY) * PI_DOUBLE
                    
                    If theta >= (PI * 1.5) Then
                        t = PI_DOUBLE - theta
                    ElseIf theta >= PI Then
                        t = theta - PI
                    ElseIf theta > PI_HALF Then
                        t = PI - theta
                    Else
                        t = theta
                    End If
                    
                    r = sRadius * (x / finalX)
                    
                    nY = -r * Sin(t)
                    nX = r * Cos(t)
                    
                    If theta >= 1.5 * PI Then
                        srcY = midY - nY
                        srcX = midX - nX
                    ElseIf theta >= PI Then
                        srcY = midY - nY
                        srcX = midX + nX
                    ElseIf theta >= PI_HALF Then
                        srcY = midY + nY
                        srcX = midX + nX
                    Else
                        srcY = midY + nY
                        srcX = midX - nX
                    End If
                    
                Else
                
                    srcX = x
                    srcY = y
                
                End If
                            
            'Polar inversion
            Case 2
            
                'Remap the coordinates around a center point of (0, 0)
                nX = x - midX
                nY = y - midY
                
                'Calculate distance automatically
                sDistance = (nX * nX) + (nY * nY)
                
                If sDistance <> 0 Then
                    srcX = midX + midX * midX * (nX / sDistance) * polarRadius
                    srcY = midY + midY * midY * (nY / sDistance) * polarRadius
                    srcX = Modulo(srcX, finalX)
                    srcY = Modulo(srcY, finalY)
                Else
                    srcX = x
                    srcY = y
                End If
            
        End Select
        
        'The lovely .setPixels routine will handle edge detection and interpolation for us as necessary
        fSupport.SetPixels x, y, srcX, srcY, srcImageData, dstImageData
                
    Next y
        If Not suppressMessages Then
            If (x And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal x + modifyProgBarOffset
            End If
        End If
    Next x
    
    'Safely deallocate all image arrays
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    
    If g_cancelCurrentAction Then CreateXSwappedPolarCoordDIB = 0 Else CreateXSwappedPolarCoordDIB = 1

End Function

'Given two DIBs, fill one with a horizontally blurred version of the other.  A highly-optimized modified accumulation algorithm
' is used to improve performance.
'Input: left and right distance to blur (I call these radii, because the final box size is (leftoffset + rightoffset + 1)
'
'IMPORTANT NOTE!  As of v7.0, this function was modified to *require* 32-bpp source images.  Passing it 24-bpp images will fail.
Public Function CreateHorizontalBlurDIB(ByVal lRadius As Long, ByVal rRadius As Long, ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long
    
    'As of v7.0, only 32-bpp RGBA images are supported.  (This matches internal design changes to PD.)
    If (srcDIB.GetDIBColorDepth <> 32) Then
        PDDebug.LogAction "WARNING!  CreateHorizontalBlurDIB requires 32-bpp inputs.  Function abandoned."
        Exit Function
    End If
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SafeArray2D
    PrepSafeArray dstSA, dstDIB
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array.  This will contain a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent blurred pixel values from spreading across the image as we go.)
    Dim srcImageData() As Byte
    Dim srcSA As SafeArray2D
    PrepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.GetDIBWidth - 1
    finalY = srcDIB.GetDIBHeight - 1
        
    Dim xStride As Long
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not suppressMessages) Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalX
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    Dim xRadius As Long
    xRadius = finalX - initX
    
    'Limit the left and right offsets to the width of the image
    If (lRadius > xRadius) Then lRadius = xRadius
    If (rRadius > xRadius) Then rRadius = xRadius
        
    'The number of pixels in the current horizontal line are tracked dynamically.  (This lets us weight edges differently,
    ' yielding a much nicer blur along boundary pixels.)
    Dim numOfPixels As Long
    numOfPixels = 0
    
    'Left and right bounds of the current accumulator
    Dim lbX As Long, ubX As Long
    
    'This horizontal blur algorithm is based on the principle of "not redoing work that's already been done."  To that end,
    ' we will store the accumulated blur total for each horizontal line, and only update it when we move one column to the right.
    Dim rTotals() As Long, gTotals() As Long, bTotals() As Long, aTotals() As Long
    ReDim rTotals(initY To finalY) As Long
    ReDim gTotals(initY To finalY) As Long
    ReDim bTotals(initY To finalY) As Long
    ReDim aTotals(initY To finalY) As Long
    
    Dim avgSample As Double
    
    'Populate the initial arrays.  We can ignore the left offset at this point, as we are starting at column 0 (and there are no
    ' pixels left of that!)
    For x = initX To initX + rRadius - 1
        xStride = x * 4
    For y = initY To finalY
        bTotals(y) = bTotals(y) + srcImageData(xStride, y)
        gTotals(y) = gTotals(y) + srcImageData(xStride + 1, y)
        rTotals(y) = rTotals(y) + srcImageData(xStride + 2, y)
        aTotals(y) = aTotals(y) + srcImageData(xStride + 3, y)
    Next y
        numOfPixels = numOfPixels + 1
    Next x
                
    'Loop through each column in the image, tallying blur values as we go
    For x = initX To finalX
        
        'Remove trailing values from the blur collection if they lie outside the processing radius
        lbX = x - lRadius
        If (lbX > 0) Then
        
            xStride = (lbX - 1) * 4
        
            For y = initY To finalY
                bTotals(y) = bTotals(y) - srcImageData(xStride, y)
                gTotals(y) = gTotals(y) - srcImageData(xStride + 1, y)
                rTotals(y) = rTotals(y) - srcImageData(xStride + 2, y)
                aTotals(y) = aTotals(y) - srcImageData(xStride + 3, y)
            Next y
            
            numOfPixels = numOfPixels - 1
        
        End If
        
        'Add leading values to the blur box if they lie inside the processing radius
        ubX = x + rRadius
        If (ubX <= finalX) Then
        
            xStride = ubX * 4
            
            For y = initY To finalY
                bTotals(y) = bTotals(y) + srcImageData(xStride, y)
                gTotals(y) = gTotals(y) + srcImageData(xStride + 1, y)
                rTotals(y) = rTotals(y) + srcImageData(xStride + 2, y)
                aTotals(y) = aTotals(y) + srcImageData(xStride + 3, y)
            Next y
            
            numOfPixels = numOfPixels + 1
            
        End If
            
        'Process the current column.  This simply involves calculating blur values, and applying them to the destination image
        xStride = x * 4
        avgSample = 1# / numOfPixels
        
        For y = initY To finalY
            dstImageData(xStride, y) = Int(bTotals(y) * avgSample + 0.5)
            dstImageData(xStride + 1, y) = Int(gTotals(y) * avgSample + 0.5)
            dstImageData(xStride + 2, y) = Int(rTotals(y) * avgSample + 0.5)
            dstImageData(xStride + 3, y) = Int(aTotals(y) * avgSample + 0.5)
        Next y
        
        'Halt for external events, like ESC-to-cancel and progress bar updates
        If (Not suppressMessages) Then
            If (x And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal x + modifyProgBarOffset
            End If
        End If
        
    Next x
        
    'Safely deallocate all image arrays
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    
    If g_cancelCurrentAction Then CreateHorizontalBlurDIB = 0 Else CreateHorizontalBlurDIB = 1
    
End Function

'Given two DIBs, fill one with a vertically blurred version of the other.  A highly-optimized modified accumulation algorithm
' is used to improve performance.
'Input: up and down distance to blur (I call these radii, because the final box size is (upoffset + downoffset + 1)
'
'IMPORTANT NOTE!  As of v7.0, this function was modified to *require* 32-bpp source images.  Passing it 24-bpp images will fail.
Public Function CreateVerticalBlurDIB(ByVal uRadius As Long, ByVal dRadius As Long, ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long
    
    'As of v7.0, only 32-bpp RGBA images are supported.  (This matches internal design changes to PD.)
    If (srcDIB.GetDIBColorDepth <> 32) Then
        PDDebug.LogAction "WARNING!  CreateVerticalBlurDIB requires 32-bpp inputs.  Function abandoned."
        Exit Function
    End If
    
    'Wrap 1D arrays around the source and destination images
    Dim dstImageData() As Byte, dstSA1D As SafeArray1D
    dstDIB.WrapArrayAroundScanline dstImageData, dstSA1D, 0
    
    Dim dstDibPointer As Long, dstDibStride As Long
    dstDibPointer = dstSA1D.pvData
    dstDibStride = dstSA1D.cElements
    
    Dim srcImageData() As Byte, srcSA1D As SafeArray1D
    srcDIB.WrapArrayAroundScanline srcImageData, srcSA1D, 0
    
    Dim srcDibPointer As Long, srcDibStride As Long
    srcDibPointer = srcSA1D.pvData
    srcDibStride = srcSA1D.cElements
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.GetDIBWidth - 1
    finalY = srcDIB.GetDIBHeight - 1
        
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim xStride As Long, quickY As Long
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalY
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    Dim yRadius As Long
    yRadius = finalY - initY
    
    'Limit the up and down offsets to the height of the image
    If (uRadius > yRadius) Then uRadius = yRadius
    If (dRadius > yRadius) Then dRadius = yRadius
        
    'The number of pixels in the current vertical line are tracked dynamically.
    Dim numOfPixels As Long
    numOfPixels = 0
            
    'Blurring takes a lot of variables
    Dim lbY As Long, ubY As Long
        
    'This vertical blur algorithm is based on the principle of "not redoing work that's already been done."  To that end,
    ' we will store the accumulated blur total for each vertical line, and only update it when we move one row down.
    Dim rTotals() As Long, gTotals() As Long, bTotals() As Long, aTotals() As Long
    ReDim rTotals(initX To finalX) As Long
    ReDim gTotals(initX To finalX) As Long
    ReDim bTotals(initX To finalX) As Long
    ReDim aTotals(initX To finalX) As Long
    
    'Populate the initial arrays.  We can ignore the up offset at this point, as we are starting at row 0 (and there are no
    ' pixels above that!)
    
    For y = initY To initY + dRadius - 1
        srcSA1D.pvData = srcDibPointer + srcDibStride * y
    For x = initX To finalX
        xStride = x * 4
        bTotals(x) = bTotals(x) + srcImageData(xStride)
        gTotals(x) = gTotals(x) + srcImageData(xStride + 1)
        rTotals(x) = rTotals(x) + srcImageData(xStride + 2)
        aTotals(x) = aTotals(x) + srcImageData(xStride + 3)
    Next x
        numOfPixels = numOfPixels + 1
    Next y
    
    Dim avgSample As Double
    
    'Loop through each row in the image, tallying blur values as we go
    For y = initY To finalY
                
        'Remove trailing values from the blur collection if they lie outside the processing radius
        lbY = y - uRadius
        If (lbY > 0) Then
        
            quickY = lbY - 1
            srcSA1D.pvData = srcDibPointer + srcDibStride * quickY
            
            For x = initX To finalX
                xStride = x * 4
                bTotals(x) = bTotals(x) - srcImageData(xStride)
                gTotals(x) = gTotals(x) - srcImageData(xStride + 1)
                rTotals(x) = rTotals(x) - srcImageData(xStride + 2)
                aTotals(x) = aTotals(x) - srcImageData(xStride + 3)
            Next x
            
            numOfPixels = numOfPixels - 1
        
        End If
        
        'Add leading values to the blur box if they lie inside the processing radius
        ubY = y + dRadius
        If (ubY <= finalY) Then
        
            quickY = ubY
            srcSA1D.pvData = srcDibPointer + srcDibStride * quickY
            
            For x = initX To finalX
                xStride = x * 4
                bTotals(x) = bTotals(x) + srcImageData(xStride)
                gTotals(x) = gTotals(x) + srcImageData(xStride + 1)
                rTotals(x) = rTotals(x) + srcImageData(xStride + 2)
                aTotals(x) = aTotals(x) + srcImageData(xStride + 3)
            Next x
            
            numOfPixels = numOfPixels + 1
            
        End If
        
        avgSample = 1# / numOfPixels
        
        dstSA1D.pvData = dstDibPointer + dstDibStride * y
        
        'Process the current row.  This simply involves calculating blur values, and applying them to the destination image.
        For x = initX To finalX
            
            xStride = x * 4
            
            'With the blur box successfully calculated, we can finally apply the results to the image.
            dstImageData(xStride) = Int(bTotals(x) * avgSample + 0.5)
            dstImageData(xStride + 1) = Int(gTotals(x) * avgSample + 0.5)
            dstImageData(xStride + 2) = Int(rTotals(x) * avgSample + 0.5)
            dstImageData(xStride + 3) = Int(aTotals(x) * avgSample + 0.5)
    
        Next x
        
        'Halt for external events, like ESC-to-cancel and progress bar updates
        If (Not suppressMessages) Then
            If (y And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal y + modifyProgBarOffset
            End If
        End If
        
    Next y
        
    'Safely deallocate all image arrays
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    
    If g_cancelCurrentAction Then CreateVerticalBlurDIB = 0 Else CreateVerticalBlurDIB = 1
    
End Function

'Given two DIBs, horizontally blur some sub-region of the source, and place the results inside the destination.
' A highly-optimized modified accumulation algorithm is used to improve performance.
'Input: left and right distance to blur (I call these radii, because the final box size is (leftoffset + rightoffset + 1)
'
'NOTE: source and destination DIBs must be 32-bpp
Public Function HorizontalBlur_SubRegion(ByVal lRadius As Long, ByVal rRadius As Long, ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, ByRef blurBounds As RectF) As Long
    
    'As of v7.0, only 32-bpp RGBA images are supported.  (This matches internal design changes to PD.)
    If (srcDIB.GetDIBColorDepth <> 32) Then
        PDDebug.LogAction "WARNING!  HorizontalBlur_SubRegion requires 32-bpp inputs.  Function abandoned."
        Exit Function
    End If
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SafeArray2D
    PrepSafeArray dstSA, dstDIB
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array.  This will contain a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent blurred pixel values from spreading across the image as we go.)
    Dim srcImageData() As Byte
    Dim srcSA As SafeArray2D
    PrepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = PDMath.Max2Int(Int(blurBounds.Left), 0)
    initY = PDMath.Max2Int(Int(blurBounds.Top), 0)
    finalX = PDMath.Min2Int(srcDIB.GetDIBWidth - 1, Int(blurBounds.Left + blurBounds.Width + 0.999999))
    finalY = PDMath.Min2Int(srcDIB.GetDIBHeight - 1, Int(blurBounds.Top + blurBounds.Height + 0.999999))
    
    Dim xLimit As Long, xStride As Long
    xLimit = srcDIB.GetDIBWidth - 1
    
    Dim xRadius As Long
    xRadius = finalX - initX
    
    'Limit the left and right offsets to the width of the sub-region
    If (lRadius > xRadius) Then lRadius = xRadius
    If (rRadius > xRadius) Then rRadius = xRadius
        
    'The number of pixels in the current horizontal line are tracked dynamically.  (This lets us weight edges differently,
    ' yielding a much nicer blur along boundary pixels.)
    Dim numOfPixels As Long
    numOfPixels = 0
    
    'Left and right bounds of the current accumulator
    Dim lbX As Long, ubX As Long
    
    'This horizontal blur algorithm is based on the principle of "not redoing work that's already been done."  To that end,
    ' we will store the accumulated blur total for each horizontal line, and only update it when we move one column to the right.
    Dim rTotals() As Long, gTotals() As Long, bTotals() As Long, aTotals() As Long
    ReDim rTotals(initY To finalY) As Long
    ReDim gTotals(initY To finalY) As Long
    ReDim bTotals(initY To finalY) As Long
    ReDim aTotals(initY To finalY) As Long
    
    Dim avgSample As Double
    
    'Populate the initial arrays.
    Dim startX As Long
    startX = initX - lRadius
    If (startX < 0) Then startX = 0
    
    For x = startX To initX + rRadius - 1
        xStride = x * 4
    For y = initY To finalY
        bTotals(y) = bTotals(y) + srcImageData(xStride, y)
        gTotals(y) = gTotals(y) + srcImageData(xStride + 1, y)
        rTotals(y) = rTotals(y) + srcImageData(xStride + 2, y)
        aTotals(y) = aTotals(y) + srcImageData(xStride + 3, y)
    Next y
        numOfPixels = numOfPixels + 1
    Next x
                
    'Loop through each column in the image, tallying blur values as we go
    For x = initX To finalX
        
        'Remove trailing values from the blur collection if they lie outside the processing radius
        lbX = x - lRadius
        If (lbX > startX) Then
        
            xStride = (lbX - 1) * 4
        
            For y = initY To finalY
                bTotals(y) = bTotals(y) - srcImageData(xStride, y)
                gTotals(y) = gTotals(y) - srcImageData(xStride + 1, y)
                rTotals(y) = rTotals(y) - srcImageData(xStride + 2, y)
                aTotals(y) = aTotals(y) - srcImageData(xStride + 3, y)
            Next y
            
            numOfPixels = numOfPixels - 1
        
        End If
        
        'Add leading values to the blur box if they lie inside the processing radius
        ubX = x + rRadius
        If (ubX <= xLimit) Then
        
            xStride = ubX * 4
            
            For y = initY To finalY
                bTotals(y) = bTotals(y) + srcImageData(xStride, y)
                gTotals(y) = gTotals(y) + srcImageData(xStride + 1, y)
                rTotals(y) = rTotals(y) + srcImageData(xStride + 2, y)
                aTotals(y) = aTotals(y) + srcImageData(xStride + 3, y)
            Next y
            
            numOfPixels = numOfPixels + 1
            
        End If
            
        'Process the current column.  This simply involves calculating blur values, and applying them to the destination image
        xStride = x * 4
        avgSample = 1# / CDbl(numOfPixels)
        For y = initY To finalY
            dstImageData(xStride, y) = Int(CDbl(bTotals(y)) * avgSample)
            dstImageData(xStride + 1, y) = Int(CDbl(gTotals(y)) * avgSample)
            dstImageData(xStride + 2, y) = Int(CDbl(rTotals(y)) * avgSample)
            dstImageData(xStride + 3, y) = Int(CDbl(aTotals(y)) * avgSample)
        Next y
        
    Next x
        
    'Safely deallocate all image arrays
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    
    If g_cancelCurrentAction Then HorizontalBlur_SubRegion = 0 Else HorizontalBlur_SubRegion = 1
    
End Function

'Given two DIBs, vertically blur some sub-region of the source, and place the results inside the destination.
' A highly-optimized modified accumulation algorithm is used to improve performance.
'Input: up and down distance to blur (I call these radii, because the final box size is (upoffset + downoffset + 1)
'
'NOTE: source and destination DIBs must be 32-bpp
Public Function VerticalBlur_SubRegion(ByVal uRadius As Long, ByVal dRadius As Long, ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, ByRef blurBounds As RectF) As Long
    
    'As of v7.0, only 32-bpp RGBA images are supported.  (This matches internal design changes to PD.)
    If (srcDIB.GetDIBColorDepth <> 32) Then
        PDDebug.LogAction "WARNING!  VerticalBlur_SubRegion requires 32-bpp inputs.  Function abandoned."
        Exit Function
    End If
    
    'Wrap 1D arrays around the source and destination images
    Dim dstImageData() As Byte, dstSA1D As SafeArray1D
    dstDIB.WrapArrayAroundScanline dstImageData, dstSA1D, 0
    
    Dim dstDibPointer As Long, dstDibStride As Long
    dstDibPointer = dstSA1D.pvData
    dstDibStride = dstSA1D.cElements
    
    Dim srcImageData() As Byte, srcSA1D As SafeArray1D
    srcDIB.WrapArrayAroundScanline srcImageData, srcSA1D, 0
    
    Dim srcDibPointer As Long, srcDibStride As Long
    srcDibPointer = srcSA1D.pvData
    srcDibStride = srcSA1D.cElements
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = PDMath.Max2Int(Int(blurBounds.Left), 0)
    initY = PDMath.Max2Int(Int(blurBounds.Top), 0)
    finalX = PDMath.Min2Int(srcDIB.GetDIBWidth - 1, Int(blurBounds.Left + blurBounds.Width + 0.999999))
    finalY = PDMath.Min2Int(srcDIB.GetDIBHeight - 1, Int(blurBounds.Top + blurBounds.Height + 0.999999))
    
    Dim yLimit As Long
    yLimit = srcDIB.GetDIBHeight - 1
        
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim xStride As Long, quickY As Long
    
    Dim yRadius As Long
    yRadius = finalY - initY
    
    'Limit the up and down offsets to the height of the image
    If (uRadius > yRadius) Then uRadius = yRadius
    If (dRadius > yRadius) Then dRadius = yRadius
        
    'The number of pixels in the current vertical line are tracked dynamically.
    Dim numOfPixels As Long
    numOfPixels = 0
            
    'Blurring takes a lot of variables
    Dim lbY As Long, ubY As Long
        
    'This vertical blur algorithm is based on the principle of "not redoing work that's already been done."  To that end,
    ' we will store the accumulated blur total for each vertical line, and only update it when we move one row down.
    Dim rTotals() As Long, gTotals() As Long, bTotals() As Long, aTotals() As Long
    ReDim rTotals(initX To finalX) As Long
    ReDim gTotals(initX To finalX) As Long
    ReDim bTotals(initX To finalX) As Long
    ReDim aTotals(initX To finalX) As Long
    
    'Populate the initial arrays.
    
    Dim startY As Long
    startY = initY - uRadius
    If (startY < 0) Then startY = 0
    
    For y = startY To initY + dRadius - 1
        srcSA1D.pvData = srcDibPointer + srcDibStride * y
    For x = initX To finalX
        xStride = x * 4
        bTotals(x) = bTotals(x) + srcImageData(xStride)
        gTotals(x) = gTotals(x) + srcImageData(xStride + 1)
        rTotals(x) = rTotals(x) + srcImageData(xStride + 2)
        aTotals(x) = aTotals(x) + srcImageData(xStride + 3)
    Next x
        numOfPixels = numOfPixels + 1
    Next y
    
    Dim avgSample As Double
    
    'Loop through each row in the image, tallying blur values as we go
    For y = initY To finalY
                
        'Remove trailing values from the blur collection if they lie outside the processing radius
        lbY = y - uRadius
        If (lbY > startY) Then
        
            quickY = lbY - 1
            srcSA1D.pvData = srcDibPointer + srcDibStride * quickY
            
            For x = initX To finalX
                xStride = x * 4
                bTotals(x) = bTotals(x) - srcImageData(xStride)
                gTotals(x) = gTotals(x) - srcImageData(xStride + 1)
                rTotals(x) = rTotals(x) - srcImageData(xStride + 2)
                aTotals(x) = aTotals(x) - srcImageData(xStride + 3)
            Next x
            
            numOfPixels = numOfPixels - 1
        
        End If
        
        'Add leading values to the blur box if they lie inside the processing radius
        ubY = y + dRadius
        If (ubY <= yLimit) Then
        
            quickY = ubY
            srcSA1D.pvData = srcDibPointer + srcDibStride * quickY
            
            For x = initX To finalX
                xStride = x * 4
                bTotals(x) = bTotals(x) + srcImageData(xStride)
                gTotals(x) = gTotals(x) + srcImageData(xStride + 1)
                rTotals(x) = rTotals(x) + srcImageData(xStride + 2)
                aTotals(x) = aTotals(x) + srcImageData(xStride + 3)
            Next x
            
            numOfPixels = numOfPixels + 1
            
        End If
        
        avgSample = 1# / numOfPixels
        
        dstSA1D.pvData = dstDibPointer + dstDibStride * y
        
        'Process the current row.  This simply involves calculating blur values, and applying them to the destination image.
        For x = initX To finalX
            
            xStride = x * 4
            
            'With the blur box successfully calculated, we can finally apply the results to the image.
            dstImageData(xStride) = Int(bTotals(x) * avgSample)
            dstImageData(xStride + 1) = Int(gTotals(x) * avgSample)
            dstImageData(xStride + 2) = Int(rTotals(x) * avgSample)
            dstImageData(xStride + 3) = Int(aTotals(x) * avgSample)
    
        Next x
        
    Next y
        
    'Safely deallocate all image arrays
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    
    If g_cancelCurrentAction Then VerticalBlur_SubRegion = 0 Else VerticalBlur_SubRegion = 1
    
End Function

'Given two DIBs, fill one with a rotated version of the other.
' Per PhotoDemon convention, this function will return a non-zero value if successful, and 0 if canceled.
Public Function CreateRotatedDIB(ByVal rotateAngle As Double, ByVal edgeHandling As Long, ByVal useBilinear As Boolean, ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, Optional ByVal centerX As Double = 0.5, Optional ByVal centerY As Double = 0.5, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long

    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SafeArray2D
    PrepSafeArray dstSA, dstDIB
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent rotated pixel values from spreading across the image as we go.)
    Dim srcImageData() As Byte
    Dim srcSA As SafeArray2D
    PrepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.GetDIBWidth - 1
    finalY = srcDIB.GetDIBHeight - 1
        
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim quickVal As Long, qvDepth As Long
    qvDepth = srcDIB.GetDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalX
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'Create a filter support class, which will aid with edge handling and interpolation
    Dim fSupport As pdFilterSupport
    Set fSupport = New pdFilterSupport
    fSupport.SetDistortParameters qvDepth, edgeHandling, useBilinear, finalX, finalY
    
    'Calculate the center of the image
    Dim midX As Double, midY As Double
    midX = CDbl(finalX - initX) * centerX
    midX = midX + initX
    midY = CDbl(finalY - initY) * centerY
    midY = midY + initY
    
    'Convert the rotation angle to radians
    rotateAngle = rotateAngle * (PI / 180#)
    
    'Find the cos and sin of this angle and store the values
    Dim cosTheta As Double, sinTheta As Double
    cosTheta = Cos(rotateAngle)
    sinTheta = Sin(rotateAngle)
    
    'Using those values, build 4 lookup tables, one each for x/y times sin/cos
    Dim xSin() As Double, xCos() As Double
    ReDim xSin(initX To finalX) As Double
    ReDim xCos(initX To finalX) As Double
    
    For x = initX To finalX
        xSin(x) = (x - midX) * sinTheta + midY
        xCos(x) = (x - midX) * cosTheta + midX
    Next
    
    Dim ySin() As Double, yCos() As Double
    ReDim ySin(initY To finalY) As Double
    ReDim yCos(initY To finalY) As Double
    For y = initY To finalY
        ySin(y) = (y - midY) * sinTheta
        yCos(y) = (y - midY) * cosTheta
    Next y
        
    'Source X and Y values, which may or may not be used as part of a bilinear interpolation function
    Dim srcX As Double, srcY As Double
    
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        quickVal = x * qvDepth
    For y = initY To finalY
        
        srcX = xCos(x) - ySin(y)
        srcY = yCos(y) + xSin(x)
        
        'The lovely .setPixels routine will handle edge detection and interpolation for us as necessary
        fSupport.SetPixels x, y, srcX, srcY, srcImageData, dstImageData
        
    Next y
        If Not suppressMessages Then
            If (x And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal x + modifyProgBarOffset
            End If
        End If
    Next x
    
    'Safely deallocate all image arrays
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    
    If g_cancelCurrentAction Then CreateRotatedDIB = 0 Else CreateRotatedDIB = 1

End Function

'Given two DIBs, fill one with an enlarged and edge-extended version of the other.  (This is often useful when something
' needs to be done to an image and edge output is tough to handle.  By extending image borders and clamping the extended
' area to the nearest valid pixels, the function can be run without specialized edge handling.)
'
'Please note that the extension value is for a SINGLE side.  The function will automatically double the horizontal and
' vertical measurements, so that matching image sides receive identical extensions.
'
'Per PhotoDemon convention, this function will return a non-zero value if successful, and 0 if canceled.
Public Function PadDIBClampedPixels(ByVal hExtend As Long, ByVal vExtend As Long, ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB) As Long

    'Start by resizing the destination DIB
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    dstDIB.CreateBlank srcDIB.GetDIBWidth + hExtend * 2, srcDIB.GetDIBHeight + vExtend * 2, srcDIB.GetDIBColorDepth
    
    'Copy the valid part of the source image into the center of the destination image
    GDI.BitBltWrapper dstDIB.GetDIBDC, hExtend, vExtend, srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, srcDIB.GetDIBDC, 0, 0, vbSrcCopy
    
    'We now need to fill the blank areas (borders) of the destination canvas with clamped values from the source image.  We do this
    ' by extending the nearest valid pixels across the empty area.
    
    'Start with the four edges, and use COLORONCOLOR as we don't want to waste time with interpolation
    SetStretchBltMode dstDIB.GetDIBDC, sbm_ColorOnColor
    
    'Top, bottom
    GDI.StretchBltWrapper dstDIB.GetDIBDC, hExtend, 0, srcDIB.GetDIBWidth, vExtend, srcDIB.GetDIBDC, 0, 0, srcDIB.GetDIBWidth, 1, vbSrcCopy
    GDI.StretchBltWrapper dstDIB.GetDIBDC, hExtend, vExtend + srcDIB.GetDIBHeight, srcDIB.GetDIBWidth, vExtend, srcDIB.GetDIBDC, 0, srcDIB.GetDIBHeight - 1, srcDIB.GetDIBWidth, 1, vbSrcCopy
    
    'Left, right
    GDI.StretchBltWrapper dstDIB.GetDIBDC, 0, vExtend, hExtend, srcDIB.GetDIBHeight, srcDIB.GetDIBDC, 0, 0, 1, srcDIB.GetDIBHeight, vbSrcCopy
    GDI.StretchBltWrapper dstDIB.GetDIBDC, srcDIB.GetDIBWidth + hExtend, vExtend, hExtend, srcDIB.GetDIBHeight, srcDIB.GetDIBDC, srcDIB.GetDIBWidth - 1, 0, 1, srcDIB.GetDIBHeight, vbSrcCopy
    
    'Next, the four corners
    
    'Top-left, top-right
    GDI.StretchBltWrapper dstDIB.GetDIBDC, 0, 0, hExtend, vExtend, srcDIB.GetDIBDC, 0, 0, 1, 1, vbSrcCopy
    GDI.StretchBltWrapper dstDIB.GetDIBDC, srcDIB.GetDIBWidth + hExtend, 0, hExtend, vExtend, srcDIB.GetDIBDC, srcDIB.GetDIBWidth - 1, 0, 1, 1, vbSrcCopy
    
    'Bottom-left, bottom-right
    GDI.StretchBltWrapper dstDIB.GetDIBDC, 0, srcDIB.GetDIBHeight + vExtend, hExtend, vExtend, srcDIB.GetDIBDC, 0, srcDIB.GetDIBHeight - 1, 1, 1, vbSrcCopy
    GDI.StretchBltWrapper dstDIB.GetDIBDC, srcDIB.GetDIBWidth + hExtend, srcDIB.GetDIBHeight + vExtend, hExtend, vExtend, srcDIB.GetDIBDC, srcDIB.GetDIBWidth - 1, srcDIB.GetDIBHeight - 1, 1, 1, vbSrcCopy
    
    'The destination DIB now contains a fully clamped, extended copy of the original image
    PadDIBClampedPixels = 1
    
End Function

'Variation on PadDIBClampedPixels, but this one takes a precise new width/height value.  This is useful with
' FFT functions that require a power-of-two size.  IMPORTANTLY, make sure the newwidth/height values are
' *larger* than the source ones, or the function may not behave as expected.
'
'Note that as a convenience, this function also returns (via parameter) the destination x/y of the image.
' This makes it simple to retrieve the image after processing.
'
'Per PhotoDemon convention, this function will return a non-zero value if successful, and 0 if canceled.
Public Function PadDIBClampedPixelsEx(ByVal newWidth As Long, ByVal newHeight As Long, ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, ByRef dstX As Long, ByRef dstY As Long) As Long

    'Start by resizing the destination DIB
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    dstDIB.CreateBlank newWidth, newHeight, srcDIB.GetDIBColorDepth
    
    'Copy the valid part of the source image into the center of the destination image
    dstX = (newWidth - srcDIB.GetDIBWidth) \ 2
    dstY = (newHeight - srcDIB.GetDIBHeight) \ 2
    GDI.BitBltWrapper dstDIB.GetDIBDC, dstX, dstY, srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, srcDIB.GetDIBDC, 0, 0, vbSrcCopy
    
    'We now need to fill the blank areas (borders) of the destination canvas with clamped values from the source image.  We do this
    ' by extending the nearest valid pixels across the empty area.
    
    'Start with the four edges, and use COLORONCOLOR as we don't want to waste time with interpolation
    SetStretchBltMode dstDIB.GetDIBDC, sbm_ColorOnColor
    
    'Top, bottom
    GDI.StretchBltWrapper dstDIB.GetDIBDC, dstX, 0, srcDIB.GetDIBWidth, dstY, srcDIB.GetDIBDC, 0, 0, srcDIB.GetDIBWidth, 1, vbSrcCopy
    GDI.StretchBltWrapper dstDIB.GetDIBDC, dstX, dstY + srcDIB.GetDIBHeight, srcDIB.GetDIBWidth, dstY, srcDIB.GetDIBDC, 0, srcDIB.GetDIBHeight - 1, srcDIB.GetDIBWidth, 1, vbSrcCopy
    
    'Left, right
    GDI.StretchBltWrapper dstDIB.GetDIBDC, 0, dstY, dstX, srcDIB.GetDIBHeight, srcDIB.GetDIBDC, 0, 0, 1, srcDIB.GetDIBHeight, vbSrcCopy
    GDI.StretchBltWrapper dstDIB.GetDIBDC, srcDIB.GetDIBWidth + dstX, dstY, dstX, srcDIB.GetDIBHeight, srcDIB.GetDIBDC, srcDIB.GetDIBWidth - 1, 0, 1, srcDIB.GetDIBHeight, vbSrcCopy
    
    'Next, the four corners
    
    'Top-left, top-right
    GDI.StretchBltWrapper dstDIB.GetDIBDC, 0, 0, dstX, dstY, srcDIB.GetDIBDC, 0, 0, 1, 1, vbSrcCopy
    GDI.StretchBltWrapper dstDIB.GetDIBDC, srcDIB.GetDIBWidth + dstX, 0, dstX, dstY, srcDIB.GetDIBDC, srcDIB.GetDIBWidth - 1, 0, 1, 1, vbSrcCopy
    
    'Bottom-left, bottom-right
    GDI.StretchBltWrapper dstDIB.GetDIBDC, 0, srcDIB.GetDIBHeight + dstY, dstX, dstY, srcDIB.GetDIBDC, 0, srcDIB.GetDIBHeight - 1, 1, 1, vbSrcCopy
    GDI.StretchBltWrapper dstDIB.GetDIBDC, srcDIB.GetDIBWidth + dstX, srcDIB.GetDIBHeight + dstY, dstX, dstY, srcDIB.GetDIBDC, srcDIB.GetDIBWidth - 1, srcDIB.GetDIBHeight - 1, 1, 1, vbSrcCopy
    
    'The destination DIB now contains a fully clamped, extended copy of the original image
    PadDIBClampedPixelsEx = 1
    
End Function

'Quickly grayscale a given DIB.
' Per PhotoDemon convention, this function will return a non-zero value if successful, and 0 if canceled.
Public Function GrayscaleDIB(ByRef srcDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long

    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte
    Dim tmpSA As SafeArray2D
    PrepSafeArray tmpSA, srcDIB
    CopyMemory ByVal VarPtrArray(imageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.GetDIBWidth - 1
    finalY = srcDIB.GetDIBHeight - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim quickVal As Long, qvDepth As Long
    qvDepth = srcDIB.GetDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalX
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'Color values
    Dim r As Long, g As Long, b As Long, grayVal As Long
    
    'Now we can loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        quickVal = x * qvDepth
    For y = initY To finalY
            
        'Get the source pixel color values
        b = imageData(quickVal, y)
        g = imageData(quickVal + 1, y)
        r = imageData(quickVal + 2, y)
        
        'Calculate a grayscale value using the original ITU-R recommended formula (BT.709, specifically)
        grayVal = (213 * r + 715 * g + 72 * b) \ 1000
        If grayVal > 255 Then grayVal = 255
        
        'Assign that gray value to each color channel
        imageData(quickVal, y) = grayVal
        imageData(quickVal + 1, y) = grayVal
        imageData(quickVal + 2, y) = grayVal
        
    Next y
        If Not suppressMessages Then
            If (x And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal x + modifyProgBarOffset
            End If
        End If
    Next x
    
    'Safely deallocate imageData()
    CopyMemory ByVal VarPtrArray(imageData), 0&, 4
    
    If g_cancelCurrentAction Then GrayscaleDIB = 0 Else GrayscaleDIB = 1
    
End Function

'Quickly modify RGB values by some constant factor.
' Per PhotoDemon convention, this function will return a non-zero value if successful, and 0 if canceled.
Public Function ScaleDIBRGBValues(ByRef srcDIB As pdDIB, Optional ByVal scaleAmount As Long = 0, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long

    'Unpremultiply the source DIB, as necessary
    If srcDIB.GetDIBColorDepth = 32 Then srcDIB.SetAlphaPremultiplication False
    
    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte
    Dim tmpSA As SafeArray2D
    PrepSafeArray tmpSA, srcDIB
    CopyMemory ByVal VarPtrArray(imageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.GetDIBWidth - 1
    finalY = srcDIB.GetDIBHeight - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim quickVal As Long, qvDepth As Long
    qvDepth = srcDIB.GetDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalX
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = ProgressBars.FindBestProgBarValue()
    End If
    
    'Color values
    Dim r As Long, g As Long, b As Long
    
    'Look-up tables are the easiest way to handle this type of conversion
    Dim scaleLookup() As Byte
    ReDim scaleLookup(0 To 255) As Byte
    
    For x = 0 To 255
        r = x + scaleAmount
        If r < 0 Then r = 0
        If r > 255 Then r = 255
        scaleLookup(x) = r
    Next x
    
    'Now we can loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        quickVal = x * qvDepth
    For y = initY To finalY
            
        'Get the source pixel color values
        r = imageData(quickVal + 2, y)
        g = imageData(quickVal + 1, y)
        b = imageData(quickVal, y)
        
        'Assign the look-up table values
        imageData(quickVal + 2, y) = scaleLookup(r)
        imageData(quickVal + 1, y) = scaleLookup(g)
        imageData(quickVal, y) = scaleLookup(b)
                
    Next y
        If Not suppressMessages Then
            If (x And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal x + modifyProgBarOffset
            End If
        End If
    Next x
    
    'Safely deallocate imageData()
    CopyMemory ByVal VarPtrArray(imageData), 0&, 4
    
    'Premultiply the source DIB, as necessary
    If (srcDIB.GetDIBColorDepth = 32) Then srcDIB.SetAlphaPremultiplication True
    
    If g_cancelCurrentAction Then ScaleDIBRGBValues = 0 Else ScaleDIBRGBValues = 1
    
End Function

'Given a DIB, scan it and find the max/min luminance values.  This function makes no changes to the DIB itself.
Public Sub GetDIBMaxMinLuminance(ByRef srcDIB As pdDIB, ByRef dibLumMin As Long, ByRef dibLumMax As Long)

    'Create a local array and point it at the pixel data we want to operate on
    Dim imageData() As Byte
    Dim tmpSA As SafeArray2D
    PrepSafeArray tmpSA, srcDIB
    CopyMemory ByVal VarPtrArray(imageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.GetDIBWidth - 1
    finalY = srcDIB.GetDIBHeight - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim quickVal As Long, qvDepth As Long
    qvDepth = srcDIB.GetDIBColorDepth \ 8
    
    'Color values
    Dim r As Long, g As Long, b As Long, grayVal As Long
    
    'Max and min values
    Dim lMax As Long, lMin As Long
    lMin = 255
    lMax = 0
    
    'Calculate max/min values for each channel
    For x = initX To finalX
        quickVal = x * qvDepth
    For y = initY To finalY
            
        'Get the source pixel color values
        r = imageData(quickVal + 2, y)
        g = imageData(quickVal + 1, y)
        b = imageData(quickVal, y)
        
        'Calculate a grayscale value using the original ITU-R recommended formula (BT.709, specifically)
        grayVal = (213 * r + 715 * g + 72 * b) \ 1000
        
        'Check max/min
        If grayVal > lMax Then
            lMax = grayVal
        ElseIf grayVal < lMin Then
            lMin = grayVal
        End If
        
    Next y
    Next x
    
    'Safely deallocate imageData()
    CopyMemory ByVal VarPtrArray(imageData), 0&, 4
    
    'Return the max/min values we calculated
    dibLumMin = lMin
    dibLumMax = lMax
    
End Sub

'Quickly modify a DIB's gamma values.  A single value is used to correct all channels.  Only 32-bpp DIBs are supported.
' NOTE: progress updated are not provided, by design; the goal here is to be as fast as possible!
Public Function FastGammaDIB(ByRef srcDIB As pdDIB, ByVal newGamma As Double) As Long
    
    'Ensure gamma is valid; bad crashes will occur otherwise
    If (newGamma <= 0#) Or (newGamma >= 10000#) Then
        PDDebug.LogAction "Invalid gamma requested in Filters_Layers.FastGammaDIB().  Gamma correction was *not* applied."
        FastGammaDIB = 0
        Exit Function
    End If
    
    'Unpremultiply the source DIB, as necessary
    If srcDIB.GetAlphaPremultiplication Then srcDIB.SetAlphaPremultiplication False

    'Create a local array and point it at the pixel data we want to operate on
    Dim imgPtr As Long, imgStride As Long
    imgPtr = srcDIB.GetDIBPointer()
    imgStride = srcDIB.GetDIBStride
    
    Dim imageData() As Byte, tmpSA As SafeArray1D
    srcDIB.WrapArrayAroundScanline imageData, tmpSA, 0&
    
    Dim x As Long, y As Long, finalX As Long, finalY As Long
    finalX = (srcDIB.GetDIBWidth - 1) * 4
    finalY = srcDIB.GetDIBHeight - 1
    
    'Look-up tables are the easiest way to handle this type of conversion
    Dim pixelLookup() As Byte
    ReDim pixelLookup(0 To 255) As Byte
    
    Dim tmpVal As Double
    Const ONE_DIV_255 As Double = 1# / 255#
    
    newGamma = 1# / newGamma
    
    For x = 0 To 255
        tmpVal = (x * ONE_DIV_255) ^ newGamma
        tmpVal = tmpVal * 255#
        If (tmpVal > 255#) Then tmpVal = 255#
        If (tmpVal < 0#) Then tmpVal = 0#
        pixelLookup(x) = tmpVal
    Next x
    
    'Now we can loop through each pixel in the image, converting values as we go
    For y = 0 To finalY
        tmpSA.pvData = imgPtr + imgStride * y
    For x = 0 To finalX Step 4
        imageData(x) = pixelLookup(imageData(x))
        imageData(x + 1) = pixelLookup(imageData(x + 1))
        imageData(x + 2) = pixelLookup(imageData(x + 2))
    Next x
    Next y
    
    srcDIB.UnwrapArrayFromDIB imageData
    
    'Premultiply the source DIB, as necessary
    If (Not srcDIB.GetAlphaPremultiplication) Then srcDIB.SetAlphaPremultiplication True
    If g_cancelCurrentAction Then FastGammaDIB = 0 Else FastGammaDIB = 1
    
End Function

'Apply bilateral smoothing (separable implementation, so faster but lower quality) to an arbitrary DIB.
' PROGRESS BAR: one call of this function requires (2 * width) progress bar range
' INPUT RANGES:
' 1) kernelRadius: Any integer 1+
' 2) spatialFactor: [0, 100]
' 3) spatialPower: [0.01, 10] - defaults to 2, generally shouldn't be set to any other value unless you understand the technical implications
' 4) colorFactor: [0, 100]
' 5) colorPower: [0.01, 10]
Public Function CreateBilateralDIB(ByRef srcDIB As pdDIB, ByVal kernelRadius As Long, ByVal spatialFactor As Double, ByVal spatialPower As Double, ByVal colorFactor As Double, ByVal colorPower As Double, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long
    
    Const colorsCount As Long = 256
    Dim spatialFunc() As Double, colorFunc() As Double

    'As a convenience to the user, we display spatial and color factors with a [0, 100].  The color factor can
    ' actually be bumped a bit, to [0, 255], so apply that now.
    colorFactor = colorFactor * 2.55
    
    'Spatial factor is left on a [0, 100] scale as a convenience to the user, but any value larger than about 10
    ' tends to produce meaningless results.  As such, shrink the input by a factor of 10.
    spatialFactor = spatialFactor / 10
    If (spatialFactor < 1#) Then spatialFactor = 1#
    
    'Spatial power is currently hidden from the user.  As such, default it to value 2.
    spatialPower = 2#
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SafeArray2D
    PrepSafeArray dstSA, srcDIB
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.GetDIBWidth - 1
    finalY = srcDIB.GetDIBHeight - 1
    
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim quickValDst As Long, quickValSrc As Long, QuickYSrc As Long, qvDepth As Long
    qvDepth = srcDIB.GetDIBColorDepth \ 8
    
    'If messages are not being suppressed, and the user did not specify a custom progress bar maximum, calculate a
    ' maximum value relevant to this function.
    If (Not suppressMessages) Then
        If (modifyProgBarMax = -1) Then SetProgBarMax finalX * 2
    End If
    
    'The kernel must be at least 1 in either direction; otherwise, we'll get range errors
    If (kernelRadius < 1) Then kernelRadius = 1
    
    'Create a second local array. This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent already-processed pixels from affecting the results of later pixels.)
    Dim srcImageData() As Byte
    Dim srcSA As SafeArray2D
    
    'To simplify the edge-handling required by this function, we're actually going to resize the source DIB with
    ' clamped pixel edges.  This removes the need for any edge handling whatsoever.
    Dim srcDIBPadded As pdDIB
    Set srcDIBPadded = New pdDIB
    PadDIBClampedPixels kernelRadius, kernelRadius, srcDIB, srcDIBPadded
    
    PrepSafeArray srcSA, srcDIBPadded
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
    
    'As part of our separable implementation, we'll also be producing an intermediate copy of the filter in either direction
    Dim midDIB As pdDIB
    Set midDIB = New pdDIB
    midDIB.CreateFromExistingDIB srcDIBPadded
    
    Dim midImageData() As Byte
    Dim midSA As SafeArray2D
    PrepSafeArray midSA, midDIB
    CopyMemory ByVal VarPtrArray(midImageData()), VarPtr(midSA), 4
        
    'To keep processing quick, only update the progress bar when absolutely necessary. This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If (Not suppressMessages) Then progBarCheck = ProgressBars.FindBestProgBarValue()
        
    'Color variables
    Dim srcR As Long, srcG As Long, srcB As Long
    Dim newR As Long, newG As Long, newB As Long
    Dim srcR0 As Long, srcG0 As Long, srcB0 As Long
    
    Dim sCoefR As Double, sCoefG As Double, sCoefB As Double
    Dim sMembR As Double, sMembG As Double, sMembB As Double
    Dim coefR As Double, coefG As Double, coefB As Double
    Dim xOffset As Long, yOffset As Long, xMax As Long, yMax As Long, xMin As Long, yMin As Long
    Dim spacialFuncCache As Double
    Dim srcPixelX As Long
    Dim i As Long, k As Long
    
    'For performance improvements, color and spatial functions are precalculated prior to starting filter.
    
    'Prepare the spatial function
    ReDim spatialFunc(-kernelRadius To kernelRadius) As Double
    
    For i = -kernelRadius To kernelRadius
        spatialFunc(i) = Exp(-0.5 * (Abs(i) / spatialFactor) ^ spatialPower)
    Next i
    
    'Prepare the color function
    ReDim colorFunc(0 To colorsCount - 1, 0 To colorsCount - 1)
    
    For i = 0 To colorsCount - 1
        For k = 0 To colorsCount - 1
            colorFunc(i, k) = Exp(-0.5 * ((Abs(i - k) / colorFactor) ^ colorPower))
        Next k
    Next i
    
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        quickValSrc = (x + kernelRadius) * qvDepth
    For y = initY To finalY
    
        sCoefR = 0
        sCoefG = 0
        sCoefB = 0
        sMembR = 0
        sMembG = 0
        sMembB = 0
        
        QuickYSrc = y + kernelRadius
        
        srcB0 = srcImageData(quickValSrc, QuickYSrc)
        srcG0 = srcImageData(quickValSrc + 1, QuickYSrc)
        srcR0 = srcImageData(quickValSrc + 2, QuickYSrc)
        
        'Cache y-loop boundaries so that they do not have to be re-calculated on the interior loop.  (X boundaries
        ' don't matter, but since we're doing it, for y, mirror it to x.)
        xMax = x + kernelRadius
        xMin = x - kernelRadius
        
        For xOffset = xMin To xMax
                
            'Cache the source pixel's x and y locations
            srcPixelX = (xOffset + kernelRadius) * qvDepth
            
            srcB = srcImageData(srcPixelX, QuickYSrc)
            srcG = srcImageData(srcPixelX + 1, QuickYSrc)
            srcR = srcImageData(srcPixelX + 2, QuickYSrc)
            
            spacialFuncCache = spatialFunc(xOffset - x)
            
            'As a general rule, when convolving data against a kernel, any kernel value below 3-sigma can effectively
            ' be ignored (as its contribution to the convolution total is not statistically meaningful). Rather than
            ' calculating an actual sigma against a standard deviation for this kernel, we can approximate a threshold
            ' because we know that our source data - RGB colors - will only ever be on a [0, 255] range.  As such,
            ' let's assume that any spatial value below 1 / 255 (roughly 0.0039) is unlikely to have a meaningful
            ' impact on the final image; by simply ignoring values below that limit, we can save ourselves additional
            ' processing time when the incoming spatial parameters are low (as is common for the cartoon-like effect).
            If (spacialFuncCache > 0.0039) Then
                
                coefR = spacialFuncCache * colorFunc(srcR, srcR0)
                coefG = spacialFuncCache * colorFunc(srcG, srcG0)
                coefB = spacialFuncCache * colorFunc(srcB, srcB0)
                
                'We could perform an additional 3-sigma check here to account for meaningless colorFunc values, but
                ' because we'd have to perform the check for each of R, G, and B, we risk inadvertently increasing
                ' processing time when the color modifiers are consistently high.  As such, I think it's best to
                ' limit our check to just the spatial modifier at present.
                
                sCoefR = sCoefR + coefR
                sCoefG = sCoefG + coefG
                sCoefB = sCoefB + coefB
                
                sMembR = sMembR + coefR * srcR
                sMembG = sMembG + coefG * srcG
                sMembB = sMembB + coefB * srcB
                
            End If
            
        Next xOffset
        
        If (sCoefR <> 0) Then newR = sMembR / sCoefR
        If (sCoefG <> 0) Then newG = sMembG / sCoefG
        If (sCoefB <> 0) Then newB = sMembB / sCoefB
                        
        'Assign the new values to each color channel
        midImageData(quickValSrc, QuickYSrc) = newB
        midImageData(quickValSrc + 1, QuickYSrc) = newG
        midImageData(quickValSrc + 2, QuickYSrc) = newR
        
    Next y
        If Not suppressMessages Then
            If (x And progBarCheck) = 0 Then
                If Interface.UserPressedESC() Then Exit For
                SetProgBarVal modifyProgBarOffset + x
            End If
        End If
    Next x
    
    'Our first pass is now complete, and the results have been cached inside midImageData.  To prevent edge distortion,
    ' we are now going to trim the mid DIB, then re-pad it with its processed edge values.
    
    If (Not g_cancelCurrentAction) Then
    
        'Release our array
        CopyMemory ByVal VarPtrArray(midImageData), 0&, 4
        Erase midImageData
        
        'Copy the contents of midDIB to the working DIB
        GDI.BitBltWrapper srcDIB.GetDIBDC, 0, 0, srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, midDIB.GetDIBDC, kernelRadius, kernelRadius, vbSrcCopy
        
        'Re-pad working DIB
        PadDIBClampedPixels kernelRadius, kernelRadius, srcDIB, midDIB
        
        'Reclaim a pointer to the DIB data
        PrepSafeArray midSA, midDIB
        CopyMemory ByVal VarPtrArray(midImageData()), VarPtr(midSA), 4
        
        'We will now apply a second bilateral pass, in the Y direction, using midImageData as the base.
        
        'Loop through each pixel in the image, converting values as we go
        For x = initX To finalX
            quickValDst = x * qvDepth
            quickValSrc = (x + kernelRadius) * qvDepth
        For y = initY To finalY
        
            sCoefR = 0
            sCoefG = 0
            sCoefB = 0
            sMembR = 0
            sMembG = 0
            sMembB = 0
            
            QuickYSrc = y + kernelRadius
            
            'IMPORTANT!  One of the tricks we use in this function is that on this second pass, we use the unmodified
            ' (well, null-padded but otherwise unmodified) copy of the image as the base of our kernel.  We then
            ' convolve those original RGB values against the already-convolved RGB values from the first pass, which
            ' gives us a better approximation of the naive convolution's "true" result.
            srcB0 = srcImageData(quickValSrc, QuickYSrc)
            srcG0 = srcImageData(quickValSrc + 1, QuickYSrc)
            srcR0 = srcImageData(quickValSrc + 2, QuickYSrc)
            
            'Cache y-loop boundaries so that they do not have to be re-calculated on the interior loop.  (X boundaries
            ' don't matter, but since we're doing it, for y, mirror it to x.)
            yMin = QuickYSrc - kernelRadius
            yMax = QuickYSrc + kernelRadius
            
                For yOffset = yMin To yMax
                    
                    'Cache the kernel pixel's x and y locations
                    srcB = midImageData(quickValSrc, yOffset)
                    srcG = midImageData(quickValSrc + 1, yOffset)
                    srcR = midImageData(quickValSrc + 2, yOffset)
                    
                    spacialFuncCache = spatialFunc(yOffset - QuickYSrc)
                    
                    'As a general rule, when convolving data against a kernel, any kernel value below 3-sigma can effectively
                    ' be ignored (as its contribution to the convolution total is not statistically meaningful). Rather than
                    ' calculating an actual sigma against a standard deviation for this kernel, we can approximate a threshold
                    ' because we know that our source data - RGB colors - will only ever be on a [0, 255] range.  As such,
                    ' let's assume that any spatial value below 1 / 255 (roughly 0.0039) is unlikely to have a meaningful
                    ' impact on the final image; by simply ignoring values below that limit, we can save ourselves additional
                    ' processing time when the incoming spatial parameters are low (as is common for the cartoon-like effect).
                    If (spacialFuncCache > 0.0039) Then
                        
                        coefR = spacialFuncCache * colorFunc(srcR, srcR0)
                        coefG = spacialFuncCache * colorFunc(srcG, srcG0)
                        coefB = spacialFuncCache * colorFunc(srcB, srcB0)
                        
                        'We could perform an additional 3-sigma check here to account for meaningless colorFunc values, but
                        ' because we'd have to perform the check for each of R, G, and B, we risk inadvertently increasing
                        ' processing time when the color modifiers are consistently high.  As such, I think it's best to
                        ' limit our check to just the spatial modifier at present.
                        
                        sCoefR = sCoefR + coefR
                        sCoefG = sCoefG + coefG
                        sCoefB = sCoefB + coefB
                        
                        sMembR = sMembR + coefR * srcR
                        sMembG = sMembG + coefG * srcG
                        sMembB = sMembB + coefB * srcB
                        
                    End If
                            
                Next yOffset
            
            If (sCoefR <> 0) Then newR = sMembR / sCoefR
            If (sCoefG <> 0) Then newG = sMembG / sCoefG
            If (sCoefB <> 0) Then newB = sMembB / sCoefB
            
            'Assign the new values to each color channel
            dstImageData(quickValDst, y) = newB
            dstImageData(quickValDst + 1, y) = newG
            dstImageData(quickValDst + 2, y) = newR
            
        Next y
            If (Not suppressMessages) Then
                If (x And progBarCheck) = 0 Then
                    If Interface.UserPressedESC() Then Exit For
                    SetProgBarVal modifyProgBarOffset + finalX + x
                End If
            End If
        Next x
        
    End If
    
    'With our work complete, point all ImageData() arrays away from their DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(midImageData), 0&, 4
    Erase midImageData
    Set midDIB = Nothing
    
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    Set srcDIBPadded = Nothing
    
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    If g_cancelCurrentAction Then CreateBilateralDIB = 0 Else CreateBilateralDIB = 1

End Function
