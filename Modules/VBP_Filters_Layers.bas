Attribute VB_Name = "Filters_Layers"
'***************************************************************************
'DIB Filters Module
'Copyright 2013-2015 by Tanner Helland
'Created: 15/February/13
'Last updated: 17/September/13
'Last update: removed the old dedicated box blur routine.  A horizontal/vertical two-pass is waaaaay faster!
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
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Constants required for creating a gamma curve from .1 to 10
Private Const MAXGAMMA As Double = 1.8460498941512
Private Const MIDGAMMA As Double = 0.68377223398334
Private Const ROOT10 As Double = 3.16227766

'Pad a DIB with blank space.  This will (obviously) resize the DIB as necessary.
Public Function padDIB(ByRef srcDIB As pdDIB, ByVal paddingSize As Long) As Boolean

    'Make a copy of the current DIB
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    tmpDIB.createFromExistingDIB srcDIB
    
    'Resize the source DIB to accommodate the new padding
    srcDIB.createBlank srcDIB.getDIBWidth + paddingSize * 2, srcDIB.getDIBHeight + paddingSize * 2, srcDIB.getDIBColorDepth, 0, 0
    srcDIB.setInitialAlphaPremultiplicationState tmpDIB.getAlphaPremultiplication
    
    'Copy the old DIB into the center of the new DIB
    BitBlt srcDIB.getDIBDC, paddingSize, paddingSize, tmpDIB.getDIBWidth, tmpDIB.getDIBHeight, tmpDIB.getDIBDC, 0, 0, vbSrcCopy
    
    'Erase the temporary DIB
    Set tmpDIB = Nothing
    
    padDIB = True

End Function

'Pad a DIB with blank space, using a RECT so that each side can be independently resized.  Note that the rect specifies how many pixels
' on each side the image should be expanded.  It does not specify the rect of the new image (because that wouldn't tell us where to
' place the image on the new rect).
' Note that this function will (obviously) resize the DIB as part of padding it.
Public Function padDIBRect(ByRef srcDIB As pdDIB, ByRef paddingRect As RECT) As Boolean

    'Make a copy of the current DIB
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    tmpDIB.createFromExistingDIB srcDIB
    
    'Resize the source DIB to accommodate the new padding
    srcDIB.createBlank srcDIB.getDIBWidth + paddingRect.Left + paddingRect.Right, srcDIB.getDIBHeight + paddingRect.Top + paddingRect.Bottom, srcDIB.getDIBColorDepth, 0, 0
    
    'Copy the old DIB into the center of the new DIB
    BitBlt srcDIB.getDIBDC, paddingRect.Left, paddingRect.Top, tmpDIB.getDIBWidth, tmpDIB.getDIBHeight, tmpDIB.getDIBDC, 0, 0, vbSrcCopy
    
    'Erase the temporary DIB
    Set tmpDIB = Nothing
    
    padDIBRect = True

End Function

'If the application needs to quickly blur a DIB and it doesn't care how, use this function.  It will lean on GDI+ if
' available (unless otherwise requested), or fall back to a high-speed internal box blur.
Public Function quickBlurDIB(ByRef srcDIB As pdDIB, ByVal blurRadius As Long, Optional ByVal useGDIPlusIfAvailable As Boolean = True) As Boolean

    If blurRadius > 0 Then
    
        'If GDI+ 1.1 exists, use it for a faster blur operation.  If only v1.0 is found, fall back to one of our internal blur functions.
        '
        'ADDENDUM JAN '15: it has come to my attention that GDI+ exhibits broken behavior on Windows 8, if the radius is less than 20px.
        '                   (Only a horizontal blur is applied, for reasons unknown.)  I have added an extra check for these circumstances,
        '                   and will revisit once Windows 10 builds have stabilized.
        Dim gdiPlusIsAcceptable As Boolean
        
        'Attempt to see if GDI+ v1.1 (or later) is available.
        If g_GDIPlusFXAvailable And useGDIPlusIfAvailable Then
        
            'Next, make sure one of two things are true:
            ' 1) We are on Windows 7, OR
            ' 2) We are on Windows 8+ and the blur radius is > 20.  Below this radius, Windows 8 doesn't blur correctly, and we've gone long
            '    enough without a patch (years!) that I don't expect MS to fix it.
            If g_IsWin8OrLater And (blurRadius <= 20) Then
                gdiPlusIsAcceptable = False
            Else
                gdiPlusIsAcceptable = True
            End If
        
        'On XP or Vista, don't bother with GDI+
        Else
            gdiPlusIsAcceptable = False
        End If
        
        Dim tmpDIB As pdDIB
        
        If gdiPlusIsAcceptable Then
        
            'GDI+ blurs are prone to failure, so as a failsafe, provide a fallback to internal PD mechanisms.
            If Not GDIPlusBlurDIB(srcDIB, blurRadius * 2, 0, 0, srcDIB.getDIBWidth, srcDIB.getDIBHeight) Then
                
                Set tmpDIB = New pdDIB
                tmpDIB.createFromExistingDIB srcDIB
                CreateApproximateGaussianBlurDIB blurRadius, tmpDIB, srcDIB, 1, True
                
            End If
            
        Else
            
            Set tmpDIB = New pdDIB
            tmpDIB.createFromExistingDIB srcDIB
            CreateApproximateGaussianBlurDIB blurRadius, tmpDIB, srcDIB, 1, True
            
        End If
    
    End If
    
    quickBlurDIB = True
    
End Function

'Given a 32bpp DIB, return a "shadow" version.  (It's pretty simple, really - black out the DIB but retain alpha values.)
Public Function createShadowDIB(ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB) As Boolean

    'If the source DIB is not 32bpp, exit.
    If srcDIB.getDIBColorDepth <> 32 Then
        createShadowDIB = False
        Exit Function
    End If

    'Start by copying the source DIB into the destination
    dstDIB.createFromExistingDIB srcDIB
    
    'Create a local array and point it at the pixel data of the destination image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepSafeArray dstSA, dstDIB
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, finalX As Long, finalY As Long, QuickX As Long
    finalX = dstDIB.getDIBWidth - 1
    finalY = dstDIB.getDIBHeight - 1
    
    'Loop through all pixels in the destination image and set them to black.  Easy as pie!
    For x = 0 To finalX
        QuickX = x * 4
    For y = 0 To finalY
    
        dstImageData(QuickX + 2, y) = 0
        dstImageData(QuickX + 1, y) = 0
        dstImageData(QuickX, y) = 0
        
    Next y
    Next x
    
    'Release our array reference and exit
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    createShadowDIB = True

End Function

'Given two DIBs, fill one with a median-filtered version of the other.
' Per PhotoDemon convention, this function will return a non-zero value if successful, and 0 if canceled.
Public Function CreateMedianDIB(ByVal mRadius As Long, ByVal mPercent As Double, ByVal kernelShape As PD_PIXEL_REGION_SHAPE, ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long

    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepSafeArray dstSA, dstDIB
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.getDIBWidth - 1
    finalY = srcDIB.getDIBHeight - 1
    
    'Just to be safe, make sure the radius isn't larger than the image itself
    If (finalY - initY) < (finalX - initX) Then
        If mRadius > (finalY - initY) Then mRadius = finalY - initY
    Else
        If mRadius > (finalX - initX) Then mRadius = finalX - initX
    End If
    
    If mRadius < 1 Then mRadius = 1
        
    mPercent = mPercent / 100
    If mPercent < 0.01 Then mPercent = 0.01
        
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, QuickValInner As Long, QuickY As Long, qvDepth As Long
    qvDepth = srcDIB.getDIBColorDepth \ 8
    
    'The x-dimension of the image has a stride of (width * 4) for 32-bit images; precalculate this, to spare us some
    ' processing time in the inner loop.
    initX = initX * qvDepth
    finalX = finalX * qvDepth
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalX
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = findBestProgBarValue()
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
    
        numOfPixels = cPixelIterator.LockTargetHistograms(rValues, gValues, bValues, aValues, False)
        
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
                    If rValues(i) > 0 Then r = r + rValues(i)
                Loop Until (r >= cutoffTotal)
                r = i
                
                i = -1
                Do
                    i = i + 1
                    If gValues(i) > 0 Then g = g + gValues(i)
                Loop Until (g >= cutoffTotal)
                g = i
                
                i = -1
                Do
                    i = i + 1
                    If bValues(i) > 0 Then b = b + bValues(i)
                Loop Until (b >= cutoffTotal)
                b = i
                
                'Finally, apply the results to the image.
                dstImageData(x, y) = b
                dstImageData(x + 1, y) = g
                dstImageData(x + 2, y) = r
                
                'Move the iterator in the correct direction
                If directionDown Then
                    If y < finalY Then numOfPixels = cPixelIterator.MoveYDown
                Else
                    If y > initY Then numOfPixels = cPixelIterator.MoveYUp
                End If
        
            Next y
            
            'Reverse y-directionality on each pass
            directionDown = Not directionDown
            If x < finalX Then numOfPixels = cPixelIterator.MoveXRight
            
            'Update the progress bar every (progBarCheck) lines
            If Not suppressMessages Then
                If (x And progBarCheck) = 0 Then
                    If userPressedESC() Then Exit For
                    SetProgBarVal x + modifyProgBarOffset
                End If
            End If
            
        Next x
        
        'Release the pixel iterator
        cPixelIterator.ReleaseTargetHistograms rValues, gValues, bValues, aValues
        
        'Release our local array that points to the target DIB
        CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
            
        If cancelCurrentAction Then CreateMedianDIB = 0 Else CreateMedianDIB = 1
    
    Else
        CreateMedianDIB = 0
    End If
    
End Function

'White balance a given DIB.
' Per PhotoDemon convention, this function will return a non-zero value if successful, and 0 if canceled.
Public Function WhiteBalanceDIB(ByVal percentIgnore As Double, ByRef srcDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long

    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepSafeArray tmpSA, srcDIB
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.getDIBWidth - 1
    finalY = srcDIB.getDIBHeight - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = srcDIB.getDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalX
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = findBestProgBarValue()
    End If
    
    'Color values
    Dim r As Long, g As Long, b As Long
    
    'Maximum and minimum values, which will be detected by our initial histogram run
    Dim RMax As Byte, gMax As Byte, bMax As Byte
    Dim RMin As Byte, gMin As Byte, bMin As Byte
    RMax = 0: gMax = 0: bMax = 0
    RMin = 255: gMin = 255: bMin = 255
    
    'Shrink the percentIgnore value down to 1% of the value we are passed (you'll see why in a moment)
    percentIgnore = percentIgnore / 100
    
    'Prepare histogram arrays
    Dim rCount(0 To 255) As Long, gCount(0 To 255) As Long, bCount(0 To 255) As Long
    For x = 0 To 255
        rCount(x) = 0
        gCount(x) = 0
        bCount(x) = 0
    Next x
    
    'Build the image histogram
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        rCount(r) = rCount(r) + 1
        gCount(g) = gCount(g) + 1
        bCount(b) = bCount(b) + 1
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
    
    r = 0: g = 0: b = 0
    
    Dim rTally As Long, gTally As Long, bTally As Long
    rTally = 0: gTally = 0: bTally = 0
    
    'Find minimum values of red, green, and blue
    Do
        If rCount(r) + rTally < wbThreshold Then
            r = r + 1
            rTally = rTally + rCount(r)
        Else
            RMin = r
            foundYet = True
        End If
    Loop While foundYet = False
        
    foundYet = False
        
    Do
        If gCount(g) + gTally < wbThreshold Then
            g = g + 1
            gTally = gTally + gCount(g)
        Else
            gMin = g
            foundYet = True
        End If
    Loop While foundYet = False
    
    foundYet = False
    
    Do
        If bCount(b) + bTally < wbThreshold Then
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
        If rCount(r) + rTally < wbThreshold Then
            r = r - 1
            rTally = rTally + rCount(r)
        Else
            RMax = r
            foundYet = True
        End If
    Loop While foundYet = False
        
    foundYet = False
        
    Do
        If gCount(g) + gTally < wbThreshold Then
            g = g - 1
            gTally = gTally + gCount(g)
        Else
            gMax = g
            foundYet = True
        End If
    Loop While foundYet = False
    
    foundYet = False
    
    Do
        If bCount(b) + bTally < wbThreshold Then
            b = b - 1
            bTally = bTally + bCount(b)
        Else
            bMax = b
            foundYet = True
        End If
    Loop While foundYet = False
    
    'Finally, calculate the difference between max and min for each color
    Dim rDif As Long, gDif As Long, bDif As Long
    rDif = CLng(RMax) - CLng(RMin)
    gDif = CLng(gMax) - CLng(gMin)
    bDif = CLng(bMax) - CLng(bMin)
    
    'We can now build a final set of look-up tables that contain the results of every possible color transformation
    Dim rFinal(0 To 255) As Byte, gFinal(0 To 255) As Byte, bFinal(0 To 255) As Byte
    
    For x = 0 To 255
        If rDif <> 0 Then r = 255 * ((x - RMin) / rDif) Else r = x
        If gDif <> 0 Then g = 255 * ((x - gMin) / gDif) Else g = x
        If bDif <> 0 Then b = 255 * ((x - bMin) / bDif) Else b = x
        If r > 255 Then r = 255
        If r < 0 Then r = 0
        If g > 255 Then g = 255
        If g < 0 Then g = 0
        If b > 255 Then b = 255
        If b < 0 Then b = 0
        rFinal(x) = r
        gFinal(x) = g
        bFinal(x) = b
    Next x
    
    'Now we can loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
            
        'Adjust white balance in a single pass (thanks to the magic of look-up tables)
        ImageData(QuickVal + 2, y) = rFinal(ImageData(QuickVal + 2, y))
        ImageData(QuickVal + 1, y) = gFinal(ImageData(QuickVal + 1, y))
        ImageData(QuickVal, y) = bFinal(ImageData(QuickVal, y))
        
    Next y
        If Not suppressMessages Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x + modifyProgBarOffset
            End If
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    If cancelCurrentAction Then WhiteBalanceDIB = 0 Else WhiteBalanceDIB = 1
    
End Function

'Contrast-correct a given DIB.  This function is similar to white-balance, except that it operates *only on luminance*, meaning individual
' color channel ratios are not changed - just luminance.  It's helpful for auto-spreading luminance across the full spectrum range, without
' changing color balance at all.
' Per PhotoDemon convention, this function will return a non-zero value if successful, and 0 if canceled.
Public Function ContrastCorrectDIB(ByVal percentIgnore As Double, ByRef srcDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long

    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepSafeArray tmpSA, srcDIB
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.getDIBWidth - 1
    finalY = srcDIB.getDIBHeight - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = srcDIB.getDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalX
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = findBestProgBarValue()
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
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Calculate a grayscale value using the original ITU-R recommended formula (BT.709, specifically)
        grayVal = (213 * r + 715 * g + 72 * b) \ 1000
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
        QuickVal = x * qvDepth
    For y = initY To finalY
            
        'Adjust white balance in a single pass (thanks to the magic of look-up tables)
        ImageData(QuickVal + 2, y) = lFinal(ImageData(QuickVal + 2, y))
        ImageData(QuickVal + 1, y) = lFinal(ImageData(QuickVal + 1, y))
        ImageData(QuickVal, y) = lFinal(ImageData(QuickVal, y))
        
    Next y
        If Not suppressMessages Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x + modifyProgBarOffset
            End If
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    If cancelCurrentAction Then ContrastCorrectDIB = 0 Else ContrastCorrectDIB = 1
    
End Function

'Given two DIBs, fill one with an artistically contoured (edge detect) version of the other.
' Per PhotoDemon convention, this function will return a non-zero value if successful, and 0 if canceled.
Public Function CreateContourDIB(ByVal blackBackground As Boolean, ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long
 
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepSafeArray dstSA, dstDIB
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent already embossed pixels from screwing up our results for later pixels.)
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, z As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 1
    initY = 1
    finalX = srcDIB.getDIBWidth - 2
    finalY = srcDIB.getDIBHeight - 2
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, QuickValRight As Long, QuickValLeft As Long, qvDepth As Long
    qvDepth = srcDIB.getDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalX
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = findBestProgBarValue()
    End If
    
    'Color variables
    Dim tmpColor As Long, tMin As Long
        
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
        QuickValRight = (x + 1) * qvDepth
        QuickValLeft = (x - 1) * qvDepth
    For y = initY To finalY
        For z = 0 To 2
    
            tMin = 255
            tmpColor = srcImageData(QuickValRight + z, y)
            If tmpColor < tMin Then tMin = tmpColor
            tmpColor = srcImageData(QuickValRight + z, y - 1)
            If tmpColor < tMin Then tMin = tmpColor
            tmpColor = srcImageData(QuickValRight + z, y + 1)
            If tmpColor < tMin Then tMin = tmpColor
            tmpColor = srcImageData(QuickValLeft + z, y)
            If tmpColor < tMin Then tMin = tmpColor
            tmpColor = srcImageData(QuickValLeft + z, y - 1)
            If tmpColor < tMin Then tMin = tmpColor
            tmpColor = srcImageData(QuickValLeft + z, y + 1)
            If tmpColor < tMin Then tMin = tmpColor
            tmpColor = srcImageData(QuickVal + z, y)
            If tmpColor < tMin Then tMin = tmpColor
            tmpColor = srcImageData(QuickVal + z, y - 1)
            If tmpColor < tMin Then tMin = tmpColor
            tmpColor = srcImageData(QuickVal + z, y + 1)
            If tmpColor < tMin Then tMin = tmpColor
            
            If tMin > 255 Then tMin = 255
            If tMin < 0 Then tMin = 0
            
            If blackBackground Then
                dstImageData(QuickVal + z, y) = srcImageData(QuickVal + z, y) - tMin
            Else
                dstImageData(QuickVal + z, y) = 255 - (srcImageData(QuickVal + z, y) - tMin)
            End If
            
            'The edges of the image will always be missed, so manually check for and correct that
            If x = initX Then dstImageData(QuickValLeft + z, y) = dstImageData(QuickVal + z, y)
            If x = finalX Then dstImageData(QuickValRight + z, y) = dstImageData(QuickVal + z, y)
            If y = initY Then dstImageData(QuickVal + z, y - 1) = dstImageData(QuickVal + z, y)
            If y = finalY Then dstImageData(QuickVal + z, y + 1) = dstImageData(QuickVal + z, y)
        
        Next z
    Next y
        If Not suppressMessages Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x + modifyProgBarOffset
            End If
        End If
    Next x
    
    'With our work complete, point both ImageData() arrays away from their DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    If cancelCurrentAction Then CreateContourDIB = 0 Else CreateContourDIB = 1
    
End Function

'Make shadows, midtone, and/or highlight adjustments to a given DIB.
' Per PhotoDemon convention, this function will return a non-zero value if successful, and 0 if canceled.
Public Function AdjustDIBShadowHighlight(ByVal shadowAmount As Double, ByVal midtoneAmount As Double, ByVal highlightAmount As Double, ByVal shadowWidth As Long, ByVal shadowRadius As Double, ByVal highlightWidth As Long, ByVal highlightRadius As Double, ByRef srcDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long
    
    'As of March 2015, this function has been entirely rewritten, using a system similar to PhotoShop's (I think...
    ' but it's impossible to know for sure, since I don't have a copy for testing!  Theoretically it should be very close.)
    '
    'This overhaul greatly improved the usefulness of this tool, but because it completed changed the input parameters, ranges, and UI
    ' of the associated form, it is incompatible with past versions of the tool.  As such, the processor call that wraps this function
    ' has been changed to prevent conflicts with old macros.
        
    'Start by converting input parameters to desired ranges.
    shadowAmount = shadowAmount / 100
    highlightAmount = -1 * (highlightAmount / 100)
    midtoneAmount = -1 * (midtoneAmount / 100)
    
    'Also, make absolute-value copies of the amount input.  (This is faster than constantly re-calculating absolute values
    ' inside the per-pixel adjustment loops.)
    Dim absShadowAmount As Double, absHighlightAmount As Double, absMidtoneAmount As Double
    absShadowAmount = Abs(shadowAmount)
    absHighlightAmount = Abs(highlightAmount)
    absMidtoneAmount = Abs(midtoneAmount)
    
    'From here, processing becomes more intensive.  Prep the progress bar as necessary.
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax 6
        Else
            SetProgBarMax modifyProgBarMax
        End If
    End If
    
    If Not suppressMessages Then SetProgBarVal 0
    
    'Next we will create shadow, midtone, and highlight lookup tables.  These will simplify the process of identifying luminance regions
    ' in the base image.
    
    'These lookup tables will be Single-type, and they will contain a value on the range [0, 1] for each 8-bit channel value [0, 255].
    ' 0 signifies a lookup entry outside that range, while 1 indicates a value fully within the target range.  Some feathering is
    ' used to make the transition between ranges appear more natural.  (The feathering used is a place where it would be really
    ' nice to have PhotoShop for comparisons, as I'm curious how they blend between shadow/midtone/highlight ranges...)
    Dim sLookup() As Single, mLookup() As Single, hLookup() As Single
    ReDim sLookup(0 To 255) As Single
    ReDim mLookup(0 To 255) As Single
    ReDim hLookup(0 To 255) As Single
    
    'Before generating the tables, generate shadow and highlight cut-offs, using the values supplied by the user.
    Dim sCutoff As Long, hCutoff As Long
    sCutoff = shadowWidth
    If sCutoff = 0 Then sCutoff = 1
    
    hCutoff = 255 - highlightWidth
    If hCutoff = 255 Then hCutoff = 254
    
    'Next, automatically determine midtone cut-offs, using the supplied shadow/highlight values as our guide
    Dim mCutoffLow As Long, mCutoffHigh As Long, mRange As Long, mMidpoint As Long
    mCutoffLow = sCutoff
    mCutoffHigh = hCutoff
    
    'If artificially low shadow/highlight ranges are used, shrink midtones accordingly
    If mCutoffLow < 64 Then mCutoffLow = 64
    If mCutoffHigh > 192 Then mCutoffHigh = 192
    mRange = mCutoffHigh - mCutoffLow
    mMidpoint = (mRange \ 2)
    
    Dim tmpCalc As Double
    
    'Now we can generate lookup tables
    Dim i As Long
    For i = 0 To 255
    
        'Shadows use a power curve maximized at 0, and descending toward the cutoff point
        If i < sCutoff Then
            tmpCalc = i / sCutoff
            tmpCalc = tmpCalc * tmpCalc
            sLookup(i) = 1 - tmpCalc
        End If
        
        'Highlights use a power curve maximized at 255, and descending toward the cutoff point
        If i > hCutoff Then
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
    
    If Not suppressMessages Then SetProgBarVal 1
    
    'First, if the shadow and highlight regions have different radius values, we need to make a backup copy of the current DIB.
    Dim backupDIB As pdDIB
    
    If shadowRadius <> highlightRadius Then
        Set backupDIB = New pdDIB
        backupDIB.createFromExistingDIB srcDIB
    End If
    
    'Next, we need to make a duplicate copy of the source image.  To improve output, this copy will be blurred, and we will use it to
    ' identify shadow/highlight regions.  (The blur naturally creates smoother transitions between light and dark parts of the image.)
    Dim blurDIB As pdDIB
    Set blurDIB = New pdDIB
    blurDIB.createFromExistingDIB srcDIB
    
    'Shadows are handled first.  If the user requested a radius > 0, blur the reference image now.
    If (shadowAmount <> 0) And (shadowRadius > 0) Then quickBlurDIB blurDIB, shadowRadius, False
        
    'Unfortunately, the next step of the operation requires manual pixel-by-pixel blending.  Prep all required loop objects now.
    
    If Not suppressMessages Then SetProgBarVal 2
    
    'Create local arrays and point them at the source DIB and blurred DIB
    Dim srcImageData() As Byte, blurImageData() As Byte
    Dim srcSA As SAFEARRAY2D, blurSA As SAFEARRAY2D
    
    prepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    prepSafeArray blurSA, blurDIB
    CopyMemory ByVal VarPtrArray(blurImageData()), VarPtr(blurSA), 4
        
    'Prep local loop variables
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.getDIBWidth - 1
    finalY = srcDIB.getDIBHeight - 1
            
    'Prep stride ofsets.  (This is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickX As Long, qvDepth As Long
    qvDepth = srcDIB.getDIBColorDepth \ 8
    
    'Prep color retrieval variables (Long-type, because intermediate calculates may exceed byte range)
    Dim rSrc As Double, gSrc As Double, bSrc As Double
    Dim rDst As Long, gDst As Long, bDst As Long
    Dim rBlur As Double, gBlur As Double, bBlur As Double
    Dim srcBlur As Long, grayBlur As Long
    Dim pxShadowCorrection As Double, pxHighlightCorrection As Double, pxMidtoneCorrection As Double
        
    'Start processing shadow pixels
    If shadowAmount <> 0 Then
    
        For x = initX To finalX
            QuickX = x * qvDepth
        For y = initY To finalY
            
            'Calculate luminance for this pixel in the *blurred* image.  (We use the blurred copy for luminance detection, to improve
            ' transitions between light and dark regions in the image.)
            bBlur = blurImageData(QuickX, y)
            gBlur = blurImageData(QuickX + 1, y)
            rBlur = blurImageData(QuickX + 2, y)
            
            grayBlur = (213 * rBlur + 715 * gBlur + 72 * bBlur) \ 1000
            If grayBlur > 255 Then grayBlur = 255
            
            'If the luminance of this pixel falls within the shadow range, continue processing; otherwise, ignore it and
            ' move on to the next pixel.
            If sLookup(grayBlur) > 0 Then
                
                'Invert the blur pixel values, and convert to the range [0, 1]
                If shadowAmount > 0 Then
                    rBlur = 1 - (rBlur / 255)
                    gBlur = 1 - (gBlur / 255)
                    bBlur = 1 - (bBlur / 255)
                Else
                    rBlur = (rBlur / 255)
                    gBlur = (gBlur / 255)
                    bBlur = (bBlur / 255)
                End If
                
                'Retrieve source pixel values and convert to the range [0, 1]
                bSrc = srcImageData(QuickX, y)
                gSrc = srcImageData(QuickX + 1, y)
                rSrc = srcImageData(QuickX + 2, y)
                
                rSrc = rSrc / 255
                gSrc = gSrc / 255
                bSrc = bSrc / 255
                
                'Calculate a maximum strength adjustment value.
                ' (This code is actually just the Overlay compositor formula.)
                If rSrc < 0.5 Then rBlur = 2 * rSrc * rBlur Else rBlur = 1 - 2 * (1 - rSrc) * (1 - rBlur)
                If gSrc < 0.5 Then gBlur = 2 * gSrc * gBlur Else gBlur = 1 - 2 * (1 - gSrc) * (1 - gBlur)
                If bSrc < 0.5 Then bBlur = 2 * bSrc * bBlur Else bBlur = 1 - 2 * (1 - bSrc) * (1 - bBlur)
                
                'Calculate a final shadow correction amount, which is a combination of...
                ' 1) The user-supplied shadow correction amount
                ' 2) The shadow lookup table for this value
                pxShadowCorrection = absShadowAmount * sLookup(grayBlur)
                
                'Modify the maximum strength adjustment value by the user-supplied shadow correction amount
                bDst = 255 * ((pxShadowCorrection * bBlur) + ((1 - pxShadowCorrection) * bSrc))
                gDst = 255 * ((pxShadowCorrection * gBlur) + ((1 - pxShadowCorrection) * gSrc))
                rDst = 255 * ((pxShadowCorrection * rBlur) + ((1 - pxShadowCorrection) * rSrc))
                
                'Save the modified values into the source image
                srcImageData(QuickX, y) = bDst
                srcImageData(QuickX + 1, y) = gDst
                srcImageData(QuickX + 2, y) = rDst
                
            End If
            
        Next y
            If Not suppressMessages Then
                If (x And 63) = 0 Then
                    If userPressedESC() Then Exit For
                End If
            End If
        Next x
        
    End If
    
    'With our shadow work complete, point all local arrays away from their respective DIBs
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    CopyMemory ByVal VarPtrArray(blurImageData), 0&, 4
    
    If Not suppressMessages Then SetProgBarVal 3
    
    'Next, it's time to operate on highlights.  The steps involved are pretty much identical to shadows, but we obviously
    ' use the highlight lookup table to determine valid correction candidates.
    If (highlightAmount <> 0) And (Not cancelCurrentAction) Then
    
        'Before starting per-pixel processing, see if a highlight radius was specified.  If it was, and the radius differs
        ' from the shadow radius, calculate a new blur DIB now.
        If (highlightRadius <> shadowRadius) Then
            
            blurDIB.createFromExistingDIB backupDIB
            If (highlightRadius <> 0) Then quickBlurDIB blurDIB, highlightRadius, False
            
            'Note that we can now free our backup DIB, as it's no longer needed
            Set backupDIB = Nothing
            
        End If
        
        If Not suppressMessages Then SetProgBarVal 4
        
        'Once again, point arrays at both the source and blur DIBs
        prepSafeArray srcSA, srcDIB
        CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
            
        prepSafeArray blurSA, blurDIB
        CopyMemory ByVal VarPtrArray(blurImageData()), VarPtr(blurSA), 4
        
        'Start per-pixel highlight processing!
        For x = initX To finalX
            QuickX = x * qvDepth
        For y = initY To finalY
            
            'Calculate luminance for this pixel in the *blurred* image.  (We use the blurred copy for luminance detection, to improve
            ' transitions between light and dark regions in the image.)
            bBlur = blurImageData(QuickX, y)
            gBlur = blurImageData(QuickX + 1, y)
            rBlur = blurImageData(QuickX + 2, y)
            
            grayBlur = (213 * rBlur + 715 * gBlur + 72 * bBlur) \ 1000
            If grayBlur > 255 Then grayBlur = 255
            
            'If the luminance of this pixel falls within the highlight range, continue processing; otherwise, ignore it and
            ' move on to the next pixel.
            If hLookup(grayBlur) > 0 Then
                
                'Invert the blur pixel values, and convert to the range [0, 1]
                If highlightAmount > 0 Then
                    rBlur = 1 - (rBlur / 255)
                    gBlur = 1 - (gBlur / 255)
                    bBlur = 1 - (bBlur / 255)
                Else
                    rBlur = (rBlur / 255)
                    gBlur = (gBlur / 255)
                    bBlur = (bBlur / 255)
                End If
                
                'Retrieve source pixel values and convert to the range [0, 1]
                bSrc = srcImageData(QuickX, y)
                gSrc = srcImageData(QuickX + 1, y)
                rSrc = srcImageData(QuickX + 2, y)
                
                rSrc = rSrc / 255
                gSrc = gSrc / 255
                bSrc = bSrc / 255
                
                'Calculate a maximum strength adjustment value.
                ' (This code is actually just the Overlay compositor formula.)
                If rSrc < 0.5 Then rBlur = 2 * rSrc * rBlur Else rBlur = 1 - 2 * (1 - rSrc) * (1 - rBlur)
                If gSrc < 0.5 Then gBlur = 2 * gSrc * gBlur Else gBlur = 1 - 2 * (1 - gSrc) * (1 - gBlur)
                If bSrc < 0.5 Then bBlur = 2 * bSrc * bBlur Else bBlur = 1 - 2 * (1 - bSrc) * (1 - bBlur)
                
                'Calculate a final highlight correction amount, which is a combination of...
                ' 1) The user-supplied highlight correction amount
                ' 2) The highlight lookup table for this value
                pxHighlightCorrection = absHighlightAmount * hLookup(grayBlur)
                
                'Modify the maximum strength adjustment value by the user-supplied highlight correction amount
                bDst = 255 * ((pxHighlightCorrection * bBlur) + ((1 - pxHighlightCorrection) * bSrc))
                gDst = 255 * ((pxHighlightCorrection * gBlur) + ((1 - pxHighlightCorrection) * gSrc))
                rDst = 255 * ((pxHighlightCorrection * rBlur) + ((1 - pxHighlightCorrection) * rSrc))
                
                'Save the modified values into the source image
                srcImageData(QuickX, y) = bDst
                srcImageData(QuickX + 1, y) = gDst
                srcImageData(QuickX + 2, y) = rDst
                
            End If
            
        Next y
            If Not suppressMessages Then
                If (x And 63) = 0 Then
                    If userPressedESC() Then Exit For
                End If
            End If
        Next x
        
        'With our highlight work complete, point all local arrays away from their respective DIBs
        CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
        CopyMemory ByVal VarPtrArray(blurImageData), 0&, 4
    
    End If
    
    If Not suppressMessages Then SetProgBarVal 5
    
    'We are now done with the blur DIB, so let's free it regardless of what comes next
    Set blurDIB = Nothing
    
    'Last up is midtone correction.  The steps involved are pretty much identical to shadow and highlight correction, but we obviously
    ' use the midtone lookup table to determine valid correction candidates.  (Also, we do not use a blurred copy of the DIB.)
    If (midtoneAmount <> 0) And (Not cancelCurrentAction) Then
        
        'Once again, point an array at the source DIB
        prepSafeArray srcSA, srcDIB
        CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
        'Start per-pixel midtone processing!
        For x = initX To finalX
            QuickX = x * qvDepth
        For y = initY To finalY
            
            'Calculate luminance for this pixel in the *source* image.
            bSrc = srcImageData(QuickX, y)
            gSrc = srcImageData(QuickX + 1, y)
            rSrc = srcImageData(QuickX + 2, y)
            
            srcBlur = (213 * rSrc + 715 * gSrc + 72 * bSrc) \ 1000
            If srcBlur > 255 Then srcBlur = 255
            
            'If the luminance of this pixel falls within the highlight range, continue processing; otherwise, ignore it and
            ' move on to the next pixel.
            If mLookup(srcBlur) > 0 Then
                
                'Convert the source pixel values to the range [0, 1]
                bSrc = bSrc / 255
                gSrc = gSrc / 255
                rSrc = rSrc / 255
                
                'To cut down on the need for additional local variables, we're going to simply re-use the blur variable names here.
                If midtoneAmount > 0 Then
                    rBlur = 1 - rSrc
                    gBlur = 1 - gSrc
                    bBlur = 1 - bSrc
                Else
                    rBlur = rSrc
                    gBlur = gSrc
                    bBlur = bSrc
                End If
                
                'Calculate a maximum strength adjustment value.
                ' (This code is actually just the Overlay compositor formula.)
                If rSrc < 0.5 Then rBlur = 2 * rSrc * rBlur Else rBlur = 1 - 2 * (1 - rSrc) * (1 - rBlur)
                If gSrc < 0.5 Then gBlur = 2 * gSrc * gBlur Else gBlur = 1 - 2 * (1 - gSrc) * (1 - gBlur)
                If bSrc < 0.5 Then bBlur = 2 * bSrc * bBlur Else bBlur = 1 - 2 * (1 - bSrc) * (1 - bBlur)
                
                'Calculate a final midtone correction amount, which is a combination of...
                ' 1) The user-supplied midtone correction amount
                ' 2) The midtone lookup table for this value
                pxMidtoneCorrection = absMidtoneAmount * mLookup(srcBlur)
                
                'Modify the maximum strength adjustment value by the user-supplied midtone correction amount
                bDst = 255 * ((pxMidtoneCorrection * bBlur) + ((1 - pxMidtoneCorrection) * bSrc))
                gDst = 255 * ((pxMidtoneCorrection * gBlur) + ((1 - pxMidtoneCorrection) * gSrc))
                rDst = 255 * ((pxMidtoneCorrection * rBlur) + ((1 - pxMidtoneCorrection) * rSrc))
                
                'Save the modified values into the source image
                srcImageData(QuickX, y) = bDst
                srcImageData(QuickX + 1, y) = gDst
                srcImageData(QuickX + 2, y) = rDst
                
            End If
            
        Next y
            If Not suppressMessages Then
                If (x And 63) = 0 Then
                    If userPressedESC() Then Exit For
                End If
            End If
        Next x
        
        'With our highlight work complete, point all local arrays away from their respective DIBs
        CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
        
    End If
    
    If Not suppressMessages Then SetProgBarVal 6
    
    If cancelCurrentAction Then AdjustDIBShadowHighlight = 0 Else AdjustDIBShadowHighlight = 1
    
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

    'Create an extra temp DIB.  This will contain the intermediate copy of our horizontal/vertical blurs.
    Dim gaussDIB As pdDIB
    Set gaussDIB = New pdDIB
    gaussDIB.createFromExistingDIB srcDIB
    dstDIB.createFromExistingDIB gaussDIB
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If modifyProgBarMax = -1 Then modifyProgBarMax = gaussDIB.getDIBWidth * numIterations + gaussDIB.getDIBHeight * numIterations
    If Not suppressMessages Then SetProgBarMax modifyProgBarMax
    
    progBarCheck = findBestProgBarValue()
    
    'Modify the Gaussian radius, and convert it to an integer.  (Box blurs don't work on floating-point radii.)
    Dim comparableRadius As Long
    
    'If the number of iterations = 3, we can approximate a correct radius using a piece-wise quadratic convolution kernel.
    ' This should result in a kernel that's ~97% identical to a Gaussian kernel.  For a more in-depth explanation of
    ' converting between standard deviation and a box blur estimation, please see this W3 spec:
    ' http://www.w3.org/TR/SVG11/filters.html#feGaussianBlurElement
    If numIterations = 3 Then
        Dim stdDev As Double
        stdDev = Sqr(-(equivalentGaussianRadius * equivalentGaussianRadius) / (2 * Log(1# / 255#)))
        comparableRadius = Int(stdDev * 2.37997232) / 2 - 1
    Else
    
        'For larger iterations, it's not worth the trouble to perform a fine estimation, as the repeat iterations will
        ' eliminate any smaller discrepancies.  Use a quick-and-dirty computation to calculate radius:
        comparableRadius = Int(equivalentGaussianRadius / numIterations + 0.5)
    
    End If
    
    'Box blurs require a radius of at least 1, so force it to that
    If comparableRadius < 1 Then comparableRadius = 1
    
    'Iterate a box blur, switching between the gauss and destination DIBs as we go
    Dim i As Long
    For i = 1 To numIterations
    
        If CreateHorizontalBlurDIB(comparableRadius, comparableRadius, dstDIB, gaussDIB, suppressMessages, modifyProgBarMax, modifyProgBarOffset + (gaussDIB.getDIBWidth * (i - 1)) + (gaussDIB.getDIBHeight * (i - 1))) > 0 Then
            If CreateVerticalBlurDIB(comparableRadius, comparableRadius, gaussDIB, dstDIB, suppressMessages, modifyProgBarMax, modifyProgBarOffset + (gaussDIB.getDIBWidth * i) + (gaussDIB.getDIBHeight * (i - 1))) = 0 Then
                Exit For
            End If
        Else
            Exit For
        End If
    
    Next i
    
    'Erase the temporary DIB and exit
    gaussDIB.eraseDIB
    Set gaussDIB = Nothing
    
    If cancelCurrentAction Then CreateApproximateGaussianBlurDIB = 0 Else CreateApproximateGaussianBlurDIB = 1

End Function

'Given two DIBs, fill one with a gaussian-blur version of the other.
' This is an extremely optimized, integer-based version of a standard gaussian blur routine.  It uses some standard optimizations
' (e.g. separable kernels) as well as a number of VB-specific optimizations.  As such, it may not be appropriate for direct translation to
' other languages.
'
' Per PhotoDemon convention, this function will return a non-zero value if successful, and 0 if canceled.
Public Function CreateGaussianBlurDIB(ByVal userRadius As Double, ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long
            
    'Create a local array and point it at the pixel data of the destination image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepSafeArray dstSA, dstDIB
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Do the same for the source image
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    'Create one more local array.  This will contain the intermediate copy of the gaussian blur, as it must be done in two passes.
    Dim gaussDIB As pdDIB
    Set gaussDIB = New pdDIB
    gaussDIB.createFromExistingDIB srcDIB
    
    Dim GaussImageData() As Byte
    Dim gaussSA As SAFEARRAY2D
    prepSafeArray gaussSA, gaussDIB
    CopyMemory ByVal VarPtrArray(GaussImageData()), VarPtr(gaussSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.getDIBWidth - 1
    finalY = srcDIB.getDIBHeight - 1
    
    'Make sure we were passed a valid radius
    If userRadius < 0.1 Then userRadius = 0.1
    If finalX > finalY Then
        If userRadius > finalX Then userRadius = finalX
    Else
        If userRadius > finalY Then userRadius = finalY
    End If
    
    'Because the radius can now be a floating-point value, make the actual radius one larger as necessary
    Dim gRadius As Long
    Dim gRadiusModifier As Double
    
    If userRadius - Int(userRadius) > 0.0001 Then
        gRadiusModifier = userRadius - Int(userRadius)
        gRadius = Int(userRadius + 1)
    Else
        gRadiusModifier = 0
        gRadius = Int(userRadius)
    End If
        
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, QuickValInner As Long, qvDepth As Long
    qvDepth = srcDIB.getDIBColorDepth \ 8
    
    Dim chkAlpha As Boolean
    If qvDepth = 4 Then chkAlpha = True Else chkAlpha = False
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalY + finalY
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = findBestProgBarValue()
    End If
    
    'Create a one-dimensional Gaussian kernel using the requested radius
    Dim gKernel() As Double
    ReDim gKernel(-gRadius To gRadius) As Double
    
    Dim gKernelInt() As Long
    ReDim gKernelInt(-gRadius To gRadius) As Long
        
    'Calculate a standard deviation (sigma) using the GIMP formula:
    Dim stdDev As Double, stdDev2 As Double, stdDev3 As Double
    If gRadius > 1 Then
        stdDev = Sqr(-(gRadius * gRadius) / (2 * Log(1# / 255#)))
    Else
        'Note that this is my addition - for a radius of 1 the GIMP formula results in too small of a sigma value
        stdDev = gRadius    '0.5
    End If
    
    stdDev2 = stdDev * stdDev
    stdDev3 = stdDev * 3
    
    'Populate the kernel using that sigma
    Dim i As Long
    Dim curVal As Double
    
    For i = -gRadius To gRadius
        curVal = (1 / (Sqr(PI_DOUBLE) * stdDev)) * (EULER ^ (-1 * ((i * i) / (2 * stdDev2))))
        
        'Ignore values less than 3 sigma
        If curVal < stdDev3 Then
            gKernel(i) = curVal
        Else
            gKernel(i) = 0
        End If
    Next i
    
    'Because floating-point radii are now allowed, adjust the far ends of the gaussian look-up manually
    If gRadiusModifier > 0 Then
        gKernel(-gRadius) = gKernel(-gRadius) * gRadiusModifier
        gKernel(gRadius) = gKernel(gRadius) * gRadiusModifier
    End If
    
    'Find new bounds, which may exist if parts of the kernel lie outside the 3-sigma relevance limit
    Dim gLB As Long, gUB As Long
    
    gLB = -gRadius
    gUB = gRadius
    
    For i = gLB To 0
        If gKernel(i) = 0 Then gLB = i + 1
    Next i
        
    For i = gUB To 0 Step -1
        If gKernel(i) = 0 Then gUB = i - 1
    Next i
        
    'For the integer version of this function, we need to find the smallest value in the gaussian table.
    Dim gMin As Double
    gMin = 1
    For i = gLB To gUB
        If gKernel(i) < gMin Then gMin = gKernel(i)
    Next i
    
    'Fill the integer version of the gaussian table with normalized values
    For i = gLB To gUB
        gKernelInt(i) = gKernel(i) / gMin
    Next i
    
    'Finally, sum all the values in the table to find our divisor
    Dim gaussSum As Long
    gaussSum = 0
    For i = gLB To gUB
        gaussSum = gaussSum + gKernelInt(i)
    Next i
        
    'We now have a normalized 1-dimensional integer-based gaussian kernel available for convolution.
    
    'Finally, generate a specialized sum look-up table for the low end of the gaussian kernel.  We will use this to "pre-compute"
    ' the values for pixels that lie off the image (by mirroring the pixel on the edge in their place).
    Dim gLookupLow() As Long
    ReDim gLookupLow(gLB To gUB) As Long
    
    Dim runningSum As Long
    runningSum = 0
    
    For i = gLB To 0
        runningSum = runningSum + gKernelInt(i)
        gLookupLow(i) = runningSum
    Next i
    
    '...and repeat for the high end
    runningSum = 0
    
    For i = gUB To 1 Step -1
        runningSum = runningSum + gKernelInt(i)
        gLookupLow(i) = runningSum
    Next i
    
    'Color variables - in this case, sums for each color component
    Dim rSum As Long, gSum As Long, bSum As Long, aSum As Long
        
    'Next, prepare 1D arrays that will be used to point at source and destination pixel data.  VB accesses 1D arrays more quickly
    ' than 2D arrays, and this technique shaves precious time off the final calculation.
    Dim scanlineSize As Long
    scanlineSize = srcDIB.getDIBArrayWidth
    Dim origDIBPointer As Long
    origDIBPointer = srcDIB.getActualDIBBits
    Dim dstDIBPointer As Long
    dstDIBPointer = gaussDIB.getActualDIBBits
    
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
        
    Dim chkX As Long, finalChkX As Long
    finalChkX = finalX * qvDepth
    
    'Loop through each pixel in the image, converting values as we go
    For y = 0 To finalY
        
        'Accessing multidimensional arrays in VB is slow.  We cheat this by pointing a one-dimensional array
        ' at the current source and destination lines, then using that to access pixel data.
        tmpSA.pvData = origDIBPointer + scanlineSize * y
        CopyMemory ByVal VarPtrArray(tmpImageData()), VarPtr(tmpSA), 4
        
        tmpDstSA.pvData = dstDIBPointer + scanlineSize * y
        CopyMemory ByVal VarPtrArray(tmpDstImageData()), VarPtr(tmpDstSA), 4
        
    For x = initX To finalX
        
        QuickVal = x * qvDepth
    
        rSum = 0
        gSum = 0
        bSum = 0
                
        'Apply the convolution to the intermediate gaussian array
        For i = gLB To gUB
                        
            chkX = x + i
            
            'We need to give special treatment to pixels that lie off the image
            If chkX >= initX Then
                If chkX < finalX Then
                    QuickValInner = chkX * qvDepth
                    rSum = rSum + tmpImageData(QuickValInner + 2) * gKernelInt(i)
                    gSum = gSum + tmpImageData(QuickValInner + 1) * gKernelInt(i)
                    bSum = bSum + tmpImageData(QuickValInner) * gKernelInt(i)
                Else
                    chkX = i
                    rSum = rSum + tmpImageData(finalChkX + 2) * gLookupLow(chkX)
                    gSum = gSum + tmpImageData(finalChkX + 1) * gLookupLow(chkX)
                    bSum = bSum + tmpImageData(finalChkX) * gLookupLow(chkX)
                    Exit For
                End If
            Else
                chkX = gLB + Abs(chkX)
                rSum = tmpImageData(2) * gLookupLow(chkX)
                gSum = tmpImageData(1) * gLookupLow(chkX)
                bSum = tmpImageData(0) * gLookupLow(chkX)
                i = chkX
            End If
                   
        Next i
        
        'We now have sums for each of red, green, blue (and potentially alpha).  Apply those values to the source array.
        tmpDstImageData(QuickVal + 2) = rSum \ gaussSum
        tmpDstImageData(QuickVal + 1) = gSum \ gaussSum
        tmpDstImageData(QuickVal) = bSum \ gaussSum
        
        'If alpha must be checked, do it now
        If chkAlpha Then
            
            aSum = 0
            
            For i = gLB To gUB
            
                'curFactor = gKernel(i)
                chkX = x + i
                If chkX < initX Then chkX = initX
                If chkX > finalX Then chkX = finalX
                aSum = aSum + tmpImageData(chkX * qvDepth + 3) * gKernelInt(i)
                
            Next i
            
            tmpDstImageData(QuickVal + 3) = aSum \ gaussSum
            
        End If
        
    Next x
        If Not suppressMessages Then
            If (y And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal y + modifyProgBarOffset
            End If
        End If
    Next y
    
    CopyMemory ByVal VarPtrArray(tmpImageData()), 0&, 4
    CopyMemory ByVal VarPtrArray(tmpDstImageData()), 0&, 4
    
    'Because this function occurs in multiple passes, it requires specialized cancel behavior.  All array references must be dropped
    ' or the program will experience a hard-freeze.
    If cancelCurrentAction Then
        CopyMemory ByVal VarPtrArray(dstImageData()), 0&, 4
        CopyMemory ByVal VarPtrArray(srcImageData()), 0&, 4
        CopyMemory ByVal VarPtrArray(GaussImageData()), 0&, 4
        CreateGaussianBlurDIB = 0
        Exit Function
    End If
    
    dstDIBPointer = dstDIB.getActualDIBBits
    tmpDstSA.pvData = dstDIBPointer
    
    'The source array now contains a horizontally convolved image.  We now need to convolve it vertically.
    Dim chkY As Long
    
    For y = initY To finalY
    
        'Accessing multidimensional arrays in VB is slow.  We cheat this by pointing a one-dimensional array
        ' at the current destination line, then using that to access pixel data.
        tmpDstSA.pvData = dstDIBPointer + scanlineSize * y
        CopyMemory ByVal VarPtrArray(tmpDstImageData()), VarPtr(tmpDstSA), 4
    
    For x = initX To finalX
    
        QuickVal = x * qvDepth
    
        rSum = 0
        gSum = 0
        bSum = 0
        
        'Apply the convolution to the destination array, using the gaussian array as the source.
        For i = gLB To gUB
        
            chkY = y + i
            
            'We need to give special treatment to pixels that lie off the image
            If chkY >= initY Then
                If chkY > finalY Then chkY = finalY
                rSum = rSum + GaussImageData(QuickVal + 2, chkY) * gKernelInt(i)
                gSum = gSum + GaussImageData(QuickVal + 1, chkY) * gKernelInt(i)
                bSum = bSum + GaussImageData(QuickVal, chkY) * gKernelInt(i)
            Else
                chkY = gLB + Abs(chkY)
                rSum = GaussImageData(QuickVal + 2, 0) * gLookupLow(chkY)
                gSum = GaussImageData(QuickVal + 1, 0) * gLookupLow(chkY)
                bSum = GaussImageData(QuickVal, 0) * gLookupLow(chkY)
                i = chkY
            End If
                                
        Next i
        
        'We now have sums for each of red, green, blue (and potentially alpha).  Apply those values to the source array.
        tmpDstImageData(QuickVal + 2) = rSum \ gaussSum
        tmpDstImageData(QuickVal + 1) = gSum \ gaussSum
        tmpDstImageData(QuickVal) = bSum \ gaussSum
        
        'If alpha must be checked, do it now
        If chkAlpha Then
        
            aSum = 0
        
            'Apply the convolution to the destination array, using the gaussian array as the source.
            For i = gLB To gUB
                'curFactor = gKernel(i)
                chkY = y + i
                If chkY < initY Then chkY = initY
                If chkY > finalY Then chkY = finalY
                aSum = aSum + GaussImageData(QuickVal + 3, chkY) * gKernelInt(i)
            Next i
        
            tmpDstImageData(QuickVal + 3) = aSum \ gaussSum
        
        End If
                
    Next x
        If Not suppressMessages Then
            If (y And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal (y + finalY) + modifyProgBarOffset
            End If
        End If
    Next y
        
    'With our work complete, point all ImageData() arrays away from their DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(tmpDstImageData()), 0&, 4
    
    CopyMemory ByVal VarPtrArray(GaussImageData), 0&, 4
    Erase GaussImageData
    
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    'We can also erase our intermediate gaussian DIB
    gaussDIB.eraseDIB
    Set gaussDIB = Nothing
    
    If cancelCurrentAction Then CreateGaussianBlurDIB = 0 Else CreateGaussianBlurDIB = 1
    
End Function

'Given two DIBs, fill one with a polar-coordinate conversion of the other.
' Per PhotoDemon convention, this function will return a non-zero value if successful, and 0 if canceled.
Public Function CreatePolarCoordDIB(ByVal conversionMethod As Long, ByVal polarRadius As Double, ByVal edgeHandling As Long, ByVal useBilinear As Boolean, ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long

    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepSafeArray dstSA, dstDIB
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent medianred pixel values from spreading across the image as we go.)
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.getDIBWidth - 1
    finalY = srcDIB.getDIBHeight - 1
        
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = srcDIB.getDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalX
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = findBestProgBarValue()
    End If
    
    'Create a filter support class, which will aid with edge handling and interpolation
    Dim fSupport As pdFilterSupport
    Set fSupport = New pdFilterSupport
    fSupport.setDistortParameters qvDepth, edgeHandling, useBilinear, finalX, finalY
    
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
              
    sRadius = sRadius * (polarRadius / 100)
    sRadius2 = sRadius * sRadius
        
    polarRadius = 1 / (polarRadius / 100)
        
    Dim iAspect As Double
    iAspect = tHeight / tWidth
              
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
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
        fSupport.setPixels x, y, srcX, srcY, srcImageData, dstImageData
                
    Next y
        If Not suppressMessages Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x + modifyProgBarOffset
            End If
        End If
    Next x
    
    'With our work complete, point both ImageData() arrays away from their DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    If cancelCurrentAction Then CreatePolarCoordDIB = 0 Else CreatePolarCoordDIB = 1

End Function

'Given two DIBs, fill one with a polar-coordinate conversion of the other.
' Per PhotoDemon convention, this function will return a non-zero value if successful, and 0 if canceled.
' NOTE: unlike the traditional polar conversion function above, this one swaps x and y values.  There is no canonical definition for
'       how to polar convert an image, so we allow the user to choose whichever method they prefer.
Public Function CreateXSwappedPolarCoordDIB(ByVal conversionMethod As Long, ByVal polarRadius As Double, ByVal edgeHandling As Long, ByVal useBilinear As Boolean, ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long

    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepSafeArray dstSA, dstDIB
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent medianred pixel values from spreading across the image as we go.)
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.getDIBWidth - 1
    finalY = srcDIB.getDIBHeight - 1
        
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = srcDIB.getDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalX
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = findBestProgBarValue()
    End If
    
    'Create a filter support class, which will aid with edge handling and interpolation
    Dim fSupport As pdFilterSupport
    Set fSupport = New pdFilterSupport
    fSupport.setDistortParameters qvDepth, edgeHandling, useBilinear, finalX, finalY
    
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
              
    sRadius = sRadius * (polarRadius / 100)
    sRadius2 = sRadius * sRadius
        
    polarRadius = 1 / (polarRadius / 100)
        
    Dim iAspect As Double
    iAspect = tHeight / tWidth
              
    'Loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
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
        fSupport.setPixels x, y, srcX, srcY, srcImageData, dstImageData
                
    Next y
        If Not suppressMessages Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x + modifyProgBarOffset
            End If
        End If
    Next x
    
    'With our work complete, point both ImageData() arrays away from their DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    If cancelCurrentAction Then CreateXSwappedPolarCoordDIB = 0 Else CreateXSwappedPolarCoordDIB = 1

End Function

'Given two DIBs, fill one with a horizontally blurred version of the other.  A highly-optimized modified accumulation algorithm
' is used to improve performance.
'Input: left and right distance to blur (I call these radii, because the final box size is (leftoffset + rightoffset + 1)
Public Function CreateHorizontalBlurDIB(ByVal lRadius As Long, ByVal rRadius As Long, ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long

    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepSafeArray dstSA, dstDIB
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array.  This will contain a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent blurred pixel values from spreading across the image as we go.)
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.getDIBWidth - 1
    finalY = srcDIB.getDIBHeight - 1
        
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, QuickValInner As Long, qvDepth As Long
    qvDepth = srcDIB.getDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalX
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = findBestProgBarValue()
    End If
    
    Dim xRadius As Long
    xRadius = finalX - initX
    
    'Limit the left and right offsets to the width of the image
    If lRadius > xRadius Then lRadius = xRadius
    If rRadius > xRadius Then rRadius = xRadius
        
    'The number of pixels in the current horizontal line are tracked dynamically.
    Dim numOfPixels As Long
    numOfPixels = 0
            
    'Blurring takes a lot of variables
    Dim lbX As Long, ubX As Long
    Dim obuX As Boolean
    
    'This horizontal blur algorithm is based on the principle of "not redoing work that's already been done."  To that end,
    ' we will store the accumulated blur total for each horizontal line, and only update it when we move one column to the right.
    Dim rTotals() As Long, gTotals() As Long, bTotals() As Long, aTotals() As Long
    ReDim rTotals(initY To finalY) As Long
    ReDim gTotals(initY To finalY) As Long
    ReDim bTotals(initY To finalY) As Long
    ReDim aTotals(initY To finalY) As Long
    
    'Populate the initial arrays.  We can ignore the left offset at this point, as we are starting at column 0 (and there are no
    ' pixels left of that!)
    For x = initX To initX + rRadius - 1
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        rTotals(y) = rTotals(y) + srcImageData(QuickVal + 2, y)
        gTotals(y) = gTotals(y) + srcImageData(QuickVal + 1, y)
        bTotals(y) = bTotals(y) + srcImageData(QuickVal, y)
        If qvDepth = 4 Then aTotals(y) = aTotals(y) + srcImageData(QuickVal + 3, y)
        
    Next y
        'Increase the pixel tally
        numOfPixels = numOfPixels + 1
    Next x
                
    'Loop through each column in the image, tallying blur values as we go
    For x = initX To finalX
            
        QuickVal = x * qvDepth
        
        'Determine the loop bounds of the current blur box in the X direction
        lbX = x - lRadius
        If lbX < 0 Then lbX = 0
        ubX = x + rRadius
        
        If ubX > finalX Then
            obuX = True
            ubX = finalX
        Else
            obuX = False
        End If
                
        'Remove trailing values from the blur collection if they lie outside the processing radius
        If lbX > 0 Then
        
            QuickValInner = (lbX - 1) * qvDepth
        
            For y = initY To finalY
                rTotals(y) = rTotals(y) - srcImageData(QuickValInner + 2, y)
                gTotals(y) = gTotals(y) - srcImageData(QuickValInner + 1, y)
                bTotals(y) = bTotals(y) - srcImageData(QuickValInner, y)
                If qvDepth = 4 Then aTotals(y) = aTotals(y) - srcImageData(QuickValInner + 3, y)
            Next y
            
            numOfPixels = numOfPixels - 1
        
        End If
        
        'Add leading values to the blur box if they lie inside the processing radius
        If Not obuX Then
        
            QuickValInner = ubX * qvDepth
            
            For y = initY To finalY
                rTotals(y) = rTotals(y) + srcImageData(QuickValInner + 2, y)
                gTotals(y) = gTotals(y) + srcImageData(QuickValInner + 1, y)
                bTotals(y) = bTotals(y) + srcImageData(QuickValInner, y)
                If qvDepth = 4 Then aTotals(y) = aTotals(y) + srcImageData(QuickValInner + 3, y)
            Next y
            
            numOfPixels = numOfPixels + 1
            
        End If
            
        'Process the current column.  This simply involves calculating blur values, and applying them to the destination image
        For y = initY To finalY
                
            'With the blur box successfully calculated, we can finally apply the results to the image.
            dstImageData(QuickVal + 2, y) = rTotals(y) \ numOfPixels
            dstImageData(QuickVal + 1, y) = gTotals(y) \ numOfPixels
            dstImageData(QuickVal, y) = bTotals(y) \ numOfPixels
            If qvDepth = 4 Then dstImageData(QuickVal + 3, y) = aTotals(y) \ numOfPixels
    
        Next y
        
        'Halt for external events, like ESC-to-cancel and progress bar updates
        If Not suppressMessages Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x + modifyProgBarOffset
            End If
        End If
        
    Next x
        
    'With our work complete, point both ImageData() arrays away from their DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    If cancelCurrentAction Then CreateHorizontalBlurDIB = 0 Else CreateHorizontalBlurDIB = 1
    
End Function

'Given two DIBs, fill one with a vertically blurred version of the other.  A highly-optimized modified accumulation algorithm
' is used to improve performance.
'Input: up and down distance to blur (I call these radii, because the final box size is (upoffset + downoffset + 1)
Public Function CreateVerticalBlurDIB(ByVal uRadius As Long, ByVal dRadius As Long, ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long

    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepSafeArray dstSA, dstDIB
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array.  This will contain a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent blurred pixel values from spreading across the image as we go.)
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.getDIBWidth - 1
    finalY = srcDIB.getDIBHeight - 1
        
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, QuickY As Long, qvDepth As Long
    qvDepth = srcDIB.getDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalY
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = findBestProgBarValue()
    End If
    
    Dim yRadius As Long
    yRadius = finalY - initY
    
    'Limit the up and down offsets to the height of the image
    If uRadius > yRadius Then uRadius = yRadius
    If dRadius > yRadius Then dRadius = yRadius
        
    'The number of pixels in the current vertical line are tracked dynamically.
    Dim numOfPixels As Long
    numOfPixels = 0
            
    'Blurring takes a lot of variables
    Dim lbY As Long, ubY As Long
    Dim obuY As Boolean
        
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
    For x = initX To finalX
        QuickVal = x * qvDepth
        rTotals(x) = rTotals(x) + srcImageData(QuickVal + 2, y)
        gTotals(x) = gTotals(x) + srcImageData(QuickVal + 1, y)
        bTotals(x) = bTotals(x) + srcImageData(QuickVal, y)
        If qvDepth = 4 Then aTotals(x) = aTotals(x) + srcImageData(QuickVal + 3, y)
    Next x
        'Increase the pixel tally
        numOfPixels = numOfPixels + 1
    Next y
                
    'Loop through each row in the image, tallying blur values as we go
    For y = initY To finalY
        
        'Determine the loop bounds of the current blur box in the Y direction
        lbY = y - uRadius
        If lbY < 0 Then lbY = 0
        ubY = y + dRadius
        
        If ubY > finalY Then
            obuY = True
            ubY = finalY
        Else
            obuY = False
        End If
                
        'Remove trailing values from the blur collection if they lie outside the processing radius
        If lbY > 0 Then
        
            QuickY = lbY - 1
        
            For x = initX To finalX
                QuickVal = x * qvDepth
                rTotals(x) = rTotals(x) - srcImageData(QuickVal + 2, QuickY)
                gTotals(x) = gTotals(x) - srcImageData(QuickVal + 1, QuickY)
                bTotals(x) = bTotals(x) - srcImageData(QuickVal, QuickY)
                If qvDepth = 4 Then aTotals(x) = aTotals(x) - srcImageData(QuickVal + 3, QuickY)
            Next x
            
            numOfPixels = numOfPixels - 1
        
        End If
        
        'Add leading values to the blur box if they lie inside the processing radius
        If Not obuY Then
        
            QuickY = ubY
            
            For x = initX To finalX
                QuickVal = x * qvDepth
                rTotals(x) = rTotals(x) + srcImageData(QuickVal + 2, QuickY)
                gTotals(x) = gTotals(x) + srcImageData(QuickVal + 1, QuickY)
                bTotals(x) = bTotals(x) + srcImageData(QuickVal, QuickY)
                If qvDepth = 4 Then aTotals(x) = aTotals(x) + srcImageData(QuickVal + 3, QuickY)
            Next x
            
            numOfPixels = numOfPixels + 1
            
        End If
            
        'Process the current row.  This simply involves calculating blur values, and applying them to the destination image.
        For x = initX To finalX
            
            QuickVal = x * qvDepth
            
            'With the blur box successfully calculated, we can finally apply the results to the image.
            dstImageData(QuickVal + 2, y) = rTotals(x) \ numOfPixels
            dstImageData(QuickVal + 1, y) = gTotals(x) \ numOfPixels
            dstImageData(QuickVal, y) = bTotals(x) \ numOfPixels
            If qvDepth = 4 Then dstImageData(QuickVal + 3, y) = aTotals(x) \ numOfPixels
    
        Next x
        
        'Halt for external events, like ESC-to-cancel and progress bar updates
        If Not suppressMessages Then
            If (y And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal y + modifyProgBarOffset
            End If
        End If
        
    Next y
        
    'With our work complete, point both ImageData() arrays away from their DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    If cancelCurrentAction Then CreateVerticalBlurDIB = 0 Else CreateVerticalBlurDIB = 1
    
End Function

'Given two DIBs, fill one with a rotated version of the other.
' Per PhotoDemon convention, this function will return a non-zero value if successful, and 0 if canceled.
Public Function CreateRotatedDIB(ByVal rotateAngle As Double, ByVal edgeHandling As Long, ByVal useBilinear As Boolean, ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, Optional ByVal centerX As Double = 0.5, Optional ByVal centerY As Double = 0.5, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long

    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepSafeArray dstSA, dstDIB
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent rotated pixel values from spreading across the image as we go.)
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepSafeArray srcSA, srcDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.getDIBWidth - 1
    finalY = srcDIB.getDIBHeight - 1
        
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = srcDIB.getDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalX
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = findBestProgBarValue()
    End If
    
    'Create a filter support class, which will aid with edge handling and interpolation
    Dim fSupport As pdFilterSupport
    Set fSupport = New pdFilterSupport
    fSupport.setDistortParameters qvDepth, edgeHandling, useBilinear, finalX, finalY
    
    'Calculate the center of the image
    Dim midX As Double, midY As Double
    midX = CDbl(finalX - initX) * centerX
    midX = midX + initX
    midY = CDbl(finalY - initY) * centerY
    midY = midY + initY
    
    'Convert the rotation angle to radians
    rotateAngle = rotateAngle * (PI / 180)
    
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
        QuickVal = x * qvDepth
    For y = initY To finalY
                            
        srcX = xCos(x) - ySin(y)
        srcY = yCos(y) + xSin(x)
        
        'The lovely .setPixels routine will handle edge detection and interpolation for us as necessary
        fSupport.setPixels x, y, srcX, srcY, srcImageData, dstImageData
                
    Next y
        If Not suppressMessages Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x + modifyProgBarOffset
            End If
        End If
    Next x
    
    'With our work complete, point both ImageData() arrays away from their DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    If cancelCurrentAction Then CreateRotatedDIB = 0 Else CreateRotatedDIB = 1

End Function

'Given two DIBs, fill one with an enlarged and edge-extended version of the other.  (This is often useful when something
' needs to be done to an image and edge output is tough to handle.  By extending image borders and clamping the extended
' area to the nearest valid pixels, the function can be run without specialized edge handling.)
'
'Please note that the extension value is for a SINGLE side.  The function will automatically double the horizontal and
' vertical measurements, so that matching image sides receive identical extensions.
'
'Per PhotoDemon convention, this function will return a non-zero value if successful, and 0 if canceled.
Public Function padDIBClampedPixels(ByVal hExtend As Long, ByVal vExtend As Long, ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB) As Long

    'Start by resizing the destination DIB
    dstDIB.createBlank srcDIB.getDIBWidth + hExtend * 2, srcDIB.getDIBHeight + vExtend * 2, srcDIB.getDIBColorDepth
    
    'Copy the valid part of the source image into the center of the destination image
    BitBlt dstDIB.getDIBDC, hExtend, vExtend, srcDIB.getDIBWidth, srcDIB.getDIBHeight, srcDIB.getDIBDC, 0, 0, vbSrcCopy
    
    'We now need to fill the blank areas (borders) of the destination canvas with clamped values from the source image.  We do this
    ' by extending the nearest valid pixels across the empty area.
    
    'Start with the four edges, and use COLORONCOLOR as we don't want to waste time with interpolation
    SetStretchBltMode dstDIB.getDIBDC, STRETCHBLT_COLORONCOLOR
    
    'Top, bottom
    StretchBlt dstDIB.getDIBDC, hExtend, 0, srcDIB.getDIBWidth, vExtend, srcDIB.getDIBDC, 0, 0, srcDIB.getDIBWidth, 1, vbSrcCopy
    StretchBlt dstDIB.getDIBDC, hExtend, vExtend + srcDIB.getDIBHeight, srcDIB.getDIBWidth, vExtend, srcDIB.getDIBDC, 0, srcDIB.getDIBHeight - 1, srcDIB.getDIBWidth, 1, vbSrcCopy
    
    'Left, right
    StretchBlt dstDIB.getDIBDC, 0, vExtend, hExtend, srcDIB.getDIBHeight, srcDIB.getDIBDC, 0, 0, 1, srcDIB.getDIBHeight, vbSrcCopy
    StretchBlt dstDIB.getDIBDC, srcDIB.getDIBWidth + hExtend, vExtend, hExtend, srcDIB.getDIBHeight, srcDIB.getDIBDC, srcDIB.getDIBWidth - 1, 0, 1, srcDIB.getDIBHeight, vbSrcCopy
    
    'Next, the four corners
    
    'Top-left, top-right
    StretchBlt dstDIB.getDIBDC, 0, 0, hExtend, vExtend, srcDIB.getDIBDC, 0, 0, 1, 1, vbSrcCopy
    StretchBlt dstDIB.getDIBDC, srcDIB.getDIBWidth + hExtend, 0, hExtend, vExtend, srcDIB.getDIBDC, srcDIB.getDIBWidth - 1, 0, 1, 1, vbSrcCopy
    
    'Bottom-left, bottom-right
    StretchBlt dstDIB.getDIBDC, 0, srcDIB.getDIBHeight + vExtend, hExtend, vExtend, srcDIB.getDIBDC, 0, srcDIB.getDIBHeight - 1, 1, 1, vbSrcCopy
    StretchBlt dstDIB.getDIBDC, srcDIB.getDIBWidth + hExtend, srcDIB.getDIBHeight + vExtend, hExtend, vExtend, srcDIB.getDIBDC, srcDIB.getDIBWidth - 1, srcDIB.getDIBHeight - 1, 1, 1, vbSrcCopy
    
    'The destination DIB now contains a fully clamped, extended copy of the original image
    padDIBClampedPixels = 1
    
End Function

'Quickly grayscale a given DIB.
' Per PhotoDemon convention, this function will return a non-zero value if successful, and 0 if canceled.
Public Function GrayscaleDIB(ByRef srcDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long

    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepSafeArray tmpSA, srcDIB
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.getDIBWidth - 1
    finalY = srcDIB.getDIBHeight - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = srcDIB.getDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalX
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = findBestProgBarValue()
    End If
    
    'Color values
    Dim r As Long, g As Long, b As Long, grayVal As Long
    
    'Now we can loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
            
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Calculate a grayscale value using the original ITU-R recommended formula (BT.709, specifically)
        grayVal = (213 * r + 715 * g + 72 * b) \ 1000
        If grayVal > 255 Then grayVal = 255
        
        'Assign that gray value to each color channel
        ImageData(QuickVal, y) = grayVal
        ImageData(QuickVal + 1, y) = grayVal
        ImageData(QuickVal + 2, y) = grayVal
        
    Next y
        If Not suppressMessages Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x + modifyProgBarOffset
            End If
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    If cancelCurrentAction Then GrayscaleDIB = 0 Else GrayscaleDIB = 1
    
End Function

'Quickly modify RGB values by some constant factor.
' Per PhotoDemon convention, this function will return a non-zero value if successful, and 0 if canceled.
Public Function ScaleDIBRGBValues(ByRef srcDIB As pdDIB, Optional ByVal scaleAmount As Long = 0, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long

    'Unpremultiply the source DIB, as necessary
    If srcDIB.getDIBColorDepth = 32 Then srcDIB.SetAlphaPremultiplication False

    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepSafeArray tmpSA, srcDIB
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.getDIBWidth - 1
    finalY = srcDIB.getDIBHeight - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = srcDIB.getDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalX
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = findBestProgBarValue()
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
        QuickVal = x * qvDepth
    For y = initY To finalY
            
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Assign the look-up table values
        ImageData(QuickVal + 2, y) = scaleLookup(r)
        ImageData(QuickVal + 1, y) = scaleLookup(g)
        ImageData(QuickVal, y) = scaleLookup(b)
                
    Next y
        If Not suppressMessages Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x + modifyProgBarOffset
            End If
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Premultiply the source DIB, as necessary
    If srcDIB.getDIBColorDepth = 32 Then srcDIB.SetAlphaPremultiplication True
    
    If cancelCurrentAction Then ScaleDIBRGBValues = 0 Else ScaleDIBRGBValues = 1
    
End Function

'Given a DIB, scan it and find the max/min luminance values.  This function makes no changes to the DIB itself.
Public Sub getDIBMaxMinLuminance(ByRef srcDIB As pdDIB, ByRef dibLumMin As Long, ByRef dibLumMax As Long)

    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepSafeArray tmpSA, srcDIB
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.getDIBWidth - 1
    finalY = srcDIB.getDIBHeight - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = srcDIB.getDIBColorDepth \ 8
    
    'Color values
    Dim r As Long, g As Long, b As Long, grayVal As Long
    
    'Max and min values
    Dim lMax As Long, lMin As Long
    lMin = 255
    lMax = 0
    
    'Calculate max/min values for each channel
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
            
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
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
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    
    'Return the max/min values we calculated
    dibLumMin = lMin
    dibLumMax = lMax
    
End Sub

'Quickly modify a DIB's gamma values.  A single value is used to correct all channels.
' TODO!  Look at wrapping GDI+ gamma correction, if available.  That may be faster than correcting gamma manually.
' Per PhotoDemon convention, this function will return a non-zero value if successful, and 0 if canceled.
Public Function GammaCorrectDIB(ByRef srcDIB As pdDIB, ByVal newGamma As Double, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long
    
    'Make sure the supplied gamma is valid
    If newGamma <= 0 Then
        
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "Invalid gamma requested in GammaCorrectDIB.  Gamma correction was not applied."
        #End If
        
        GammaCorrectDIB = 0
        Exit Function
        
    End If
    
    'Unpremultiply the source DIB, as necessary
    If srcDIB.getDIBColorDepth = 32 Then srcDIB.SetAlphaPremultiplication False

    'Create a local array and point it at the pixel data we want to operate on
    Dim ImageData() As Byte
    Dim tmpSA As SAFEARRAY2D
    prepSafeArray tmpSA, srcDIB
    CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(tmpSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.getDIBWidth - 1
    finalY = srcDIB.getDIBHeight - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = srcDIB.getDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then
            SetProgBarMax finalX
        Else
            SetProgBarMax modifyProgBarMax
        End If
        progBarCheck = findBestProgBarValue()
    End If
    
    'Color values
    Dim r As Long, g As Long, b As Long
    
    'Look-up tables are the easiest way to handle this type of conversion
    Dim pixelLookup() As Byte
    ReDim pixelLookup(0 To 255) As Byte
    
    Dim tmpVal As Double
    
    For x = 0 To 255
    
        tmpVal = x / 255
        tmpVal = tmpVal ^ (1 / newGamma)
        tmpVal = tmpVal * 255
        
        If tmpVal > 255 Then tmpVal = 255
        If tmpVal < 0 Then tmpVal = 0
        
        pixelLookup(x) = tmpVal
        
    Next x
    
    'Now we can loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
            
        'Get the source pixel color values
        r = ImageData(QuickVal + 2, y)
        g = ImageData(QuickVal + 1, y)
        b = ImageData(QuickVal, y)
        
        'Assign the look-up table values
        ImageData(QuickVal + 2, y) = pixelLookup(r)
        ImageData(QuickVal + 1, y) = pixelLookup(g)
        ImageData(QuickVal, y) = pixelLookup(b)
                
    Next y
        If Not suppressMessages Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal x + modifyProgBarOffset
            End If
        End If
    Next x
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(ImageData), 0&, 4
    Erase ImageData
    
    'Premultiply the source DIB, as necessary
    If srcDIB.getDIBColorDepth = 32 Then srcDIB.SetAlphaPremultiplication True
    
    If cancelCurrentAction Then GammaCorrectDIB = 0 Else GammaCorrectDIB = 1
    
End Function

'Apply bilateral smoothing (separable implementation, so faster but lower quality) to an arbitrary DIB.
' PROGRESS BAR: one call of this function requires (2 * width) progress bar range
' INPUT RANGES:
' 1) kernelRadius: Any integer 1+
' 2) spatialFactor: [0, 100]
' 3) spatialPower: [0.01, 10] - defaults to 2, generally shouldn't be set to any other value unless you understand the technical implications
' 4) colorFactor: [0, 100]
' 5) colorPower: [0.01, 10]
Public Function createBilateralDIB(ByRef srcDIB As pdDIB, ByVal kernelRadius As Long, ByVal spatialFactor As Double, ByVal spatialPower As Double, ByVal colorFactor As Double, ByVal colorPower As Double, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long

    Const maxKernelSize As Long = 256
    Const colorsCount As Long = 256

    Dim spatialFunc() As Double, colorFunc() As Double

    'As a convenience to the user, we display spatial and color factors with a [0, 100].  The color factor can
    ' actually be bumped a bit, to [0, 255], so apply that now.
    colorFactor = colorFactor * 2.55
    
    'Spatial factor is left on a [0, 100] scale as a convenience to the user, but any value larger than about 10
    ' tends to produce meaningless results.  As such, shrink the input by a factor of 10.
    spatialFactor = spatialFactor / 10
    If spatialFactor < 1# Then spatialFactor = 1#
    
    'Spatial power is currently hidden from the user.  As such, default it to value 2.
    spatialPower = 2#
    
    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepSafeArray dstSA, srcDIB
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = 0
    initY = 0
    finalX = srcDIB.getDIBWidth - 1
    finalY = srcDIB.getDIBHeight - 1
    
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickValDst As Long, QuickValSrc As Long, QuickYSrc As Long, qvDepth As Long
    qvDepth = srcDIB.getDIBColorDepth \ 8
    
    'If messages are not being suppressed, and the user did not specify a custom progress bar maximum, calculate a
    ' maximum value relevant to this function.
    If Not suppressMessages Then
        If modifyProgBarMax = -1 Then SetProgBarMax finalX * 2
    End If
    
    'The kernel must be at least 1 in either direction; otherwise, we'll get range errors
    If kernelRadius < 1 Then kernelRadius = 1
    
    'Create a second local array. This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent already-processed pixels from affecting the results of later pixels.)
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    
    'To simplify the edge-handling required by this function, we're actually going to resize the source DIB with
    ' clamped pixel edges.  This removes the need for any edge handling whatsoever.
    Dim srcDIBPadded As pdDIB
    Set srcDIBPadded = New pdDIB
    padDIBClampedPixels kernelRadius, kernelRadius, srcDIB, srcDIBPadded
    
    prepSafeArray srcSA, srcDIBPadded
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
    
    'As part of our separable implementation, we'll also be producing an intermediate copy of the filter in either direction
    Dim midDIB As pdDIB
    Set midDIB = New pdDIB
    midDIB.createFromExistingDIB srcDIBPadded
    
    Dim midImageData() As Byte
    Dim midSA As SAFEARRAY2D
    prepSafeArray midSA, midDIB
    CopyMemory ByVal VarPtrArray(midImageData()), VarPtr(midSA), 4
        
    'To keep processing quick, only update the progress bar when absolutely necessary. This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    If Not suppressMessages Then progBarCheck = findBestProgBarValue()
        
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
        QuickValSrc = (x + kernelRadius) * qvDepth
    For y = initY To finalY
    
        sCoefR = 0
        sCoefG = 0
        sCoefB = 0
        sMembR = 0
        sMembG = 0
        sMembB = 0
        
        QuickYSrc = y + kernelRadius
        
        srcR0 = srcImageData(QuickValSrc + 2, QuickYSrc)
        srcG0 = srcImageData(QuickValSrc + 1, QuickYSrc)
        srcB0 = srcImageData(QuickValSrc, QuickYSrc)
        
        'Cache y-loop boundaries so that they do not have to be re-calculated on the interior loop.  (X boundaries
        ' don't matter, but since we're doing it, for y, mirror it to x.)
        xMax = x + kernelRadius
        xMin = x - kernelRadius
        
        For xOffset = xMin To xMax
                
            'Cache the source pixel's x and y locations
            srcPixelX = (xOffset + kernelRadius) * qvDepth
            
            srcR = srcImageData(srcPixelX + 2, QuickYSrc)
            srcG = srcImageData(srcPixelX + 1, QuickYSrc)
            srcB = srcImageData(srcPixelX, QuickYSrc)
            
            spacialFuncCache = spatialFunc(xOffset - x)
            
            'As a general rule, when convolving data against a kernel, any kernel value below 3-sigma can effectively
            ' be ignored (as its contribution to the convolution total is not statistically meaningful). Rather than
            ' calculating an actual sigma against a standard deviation for this kernel, we can approximate a threshold
            ' because we know that our source data - RGB colors - will only ever be on a [0, 255] range.  As such,
            ' let's assume that any spatial value below 1 / 255 (roughly 0.0039) is unlikely to have a meaningful
            ' impact on the final image; by simply ignoring values below that limit, we can save ourselves additional
            ' processing time when the incoming spatial parameters are low (as is common for the cartoon-like effect).
            If spacialFuncCache > 0.0039 Then
                
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
        
        If sCoefR <> 0 Then newR = sMembR / sCoefR
        If sCoefG <> 0 Then newG = sMembG / sCoefG
        If sCoefB <> 0 Then newB = sMembB / sCoefB
                        
        'Assign the new values to each color channel
        midImageData(QuickValSrc + 2, QuickYSrc) = newR
        midImageData(QuickValSrc + 1, QuickYSrc) = newG
        midImageData(QuickValSrc, QuickYSrc) = newB
        
    Next y
        If Not suppressMessages Then
            If (x And progBarCheck) = 0 Then
                If userPressedESC() Then Exit For
                SetProgBarVal modifyProgBarOffset + x
            End If
        End If
    Next x
    
    'Our first pass is now complete, and the results have been cached inside midImageData.  To prevent edge distortion,
    ' we are now going to trim the mid DIB, then re-pad it with its processed edge values.
    
    If Not cancelCurrentAction Then
    
        'Release our array
        CopyMemory ByVal VarPtrArray(midImageData), 0&, 4
        Erase midImageData
        
        'Copy the contents of midDIB to the working DIB
        BitBlt srcDIB.getDIBDC, 0, 0, srcDIB.getDIBWidth, srcDIB.getDIBHeight, midDIB.getDIBDC, kernelRadius, kernelRadius, vbSrcCopy
        
        'Re-pad working DIB
        padDIBClampedPixels kernelRadius, kernelRadius, srcDIB, midDIB
        
        'Reclaim a pointer to the DIB data
        prepSafeArray midSA, midDIB
        CopyMemory ByVal VarPtrArray(midImageData()), VarPtr(midSA), 4
        
        'We will now apply a second bilateral pass, in the Y direction, using midImageData as the base.
        
        'Loop through each pixel in the image, converting values as we go
        For x = initX To finalX
            QuickValDst = x * qvDepth
            QuickValSrc = (x + kernelRadius) * qvDepth
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
            srcR0 = srcImageData(QuickValSrc + 2, QuickYSrc)
            srcG0 = srcImageData(QuickValSrc + 1, QuickYSrc)
            srcB0 = srcImageData(QuickValSrc, QuickYSrc)
            
            'Cache y-loop boundaries so that they do not have to be re-calculated on the interior loop.  (X boundaries
            ' don't matter, but since we're doing it, for y, mirror it to x.)
            yMin = QuickYSrc - kernelRadius
            yMax = QuickYSrc + kernelRadius
            
                For yOffset = yMin To yMax
                    
                    'Cache the kernel pixel's x and y locations
                    srcR = midImageData(QuickValSrc + 2, yOffset)
                    srcG = midImageData(QuickValSrc + 1, yOffset)
                    srcB = midImageData(QuickValSrc, yOffset)
                    
                    spacialFuncCache = spatialFunc(yOffset - QuickYSrc)
                    
                    'As a general rule, when convolving data against a kernel, any kernel value below 3-sigma can effectively
                    ' be ignored (as its contribution to the convolution total is not statistically meaningful). Rather than
                    ' calculating an actual sigma against a standard deviation for this kernel, we can approximate a threshold
                    ' because we know that our source data - RGB colors - will only ever be on a [0, 255] range.  As such,
                    ' let's assume that any spatial value below 1 / 255 (roughly 0.0039) is unlikely to have a meaningful
                    ' impact on the final image; by simply ignoring values below that limit, we can save ourselves additional
                    ' processing time when the incoming spatial parameters are low (as is common for the cartoon-like effect).
                    If spacialFuncCache > 0.0039 Then
                        
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
            
            If sCoefR <> 0 Then newR = sMembR / sCoefR
            If sCoefG <> 0 Then newG = sMembG / sCoefG
            If sCoefB <> 0 Then newB = sMembB / sCoefB
            
            'Assign the new values to each color channel
            dstImageData(QuickValDst + 2, y) = newR
            dstImageData(QuickValDst + 1, y) = newG
            dstImageData(QuickValDst, y) = newB
            
        Next y
            If Not suppressMessages Then
                If (x And progBarCheck) = 0 Then
                    If userPressedESC() Then Exit For
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
    
    If cancelCurrentAction Then createBilateralDIB = 0 Else createBilateralDIB = 1

End Function
