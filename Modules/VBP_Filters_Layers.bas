Attribute VB_Name = "Filters_Layers"
'***************************************************************************
'DIB Filters Module
'Copyright ©2013-2014 by Tanner Helland
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
        If g_GDIPlusFXAvailable And useGDIPlusIfAvailable Then
            GDIPlusBlurDIB srcDIB, blurRadius * 2, 0, 0, srcDIB.getDIBWidth, srcDIB.getDIBHeight
        Else
            Dim tmpDIB As pdDIB
            Set tmpDIB = New pdDIB
            tmpDIB.createFromExistingDIB srcDIB
            CreateApproximateGaussianBlurDIB blurRadius, tmpDIB, srcDIB, 1, True
            Set tmpDIB = Nothing
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
Public Function CreateMedianDIB(ByVal mRadius As Long, ByVal mPercent As Double, ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long

    'Create a local array and point it at the pixel data of the current image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    prepSafeArray dstSA, dstDIB
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array.  This will contain the a copy of the current image, and we will use it as our source reference
    ' (This is necessary to prevent median-calculated pixel values from spreading across the image as we go.)
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
    
    'Just to be safe, make sure the radius isn't larger than the image itself
    If (finalY - initY) < (finalX - initX) Then
        If mRadius > (finalY - initY) Then mRadius = finalY - initY
    Else
        If mRadius > (finalX - initX) Then mRadius = finalX - initX
    End If
        
    mPercent = mPercent / 100
        
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, QuickValInner As Long, QuickY As Long, qvDepth As Long
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
    
    'The number of pixels in the current median box are tracked dynamically.
    Dim NumOfPixels As Long
    NumOfPixels = 0
            
    'Median filtering takes a lot of variables
    Dim rValues(0 To 255) As Long, gValues(0 To 255) As Long, bValues(0 To 255) As Long
    Dim lbX As Long, lbY As Long, ubX As Long, ubY As Long
    Dim obuX As Boolean, obuY As Boolean, oblY As Boolean
    Dim i As Long, j As Long
    Dim cutoffTotal As Long
    Dim r As Long, g As Long, b As Long
    Dim midR As Long, midG As Long, midB As Long
    'Dim rBins As Long, gBins As Long, bBins As Long
    
    Dim atBottom As Boolean
    atBottom = True
    
    Dim startY As Long, stopY As Long, yStep As Long
    
    NumOfPixels = 0
    
    'Generate an initial array of median data for the first pixel
    For x = initX To initX + mRadius - 1
        QuickVal = x * qvDepth
    For y = initY To initY + mRadius '- 1
    
        r = srcImageData(QuickVal + 2, y)
        g = srcImageData(QuickVal + 1, y)
        b = srcImageData(QuickVal, y)
        rValues(r) = rValues(r) + 1
        gValues(g) = gValues(g) + 1
        bValues(b) = bValues(b) + 1
        
        'Increase the pixel tally
        NumOfPixels = NumOfPixels + 1
        
    Next y
    Next x
                
    'Loop through each pixel in the image, tallying median values as we go
    For x = initX To finalX
            
        QuickVal = x * qvDepth
        
        'Determine the bounds of the current median box in the X direction
        lbX = x - mRadius
        If lbX < 0 Then lbX = 0
        
        ubX = x + mRadius
        If ubX > finalX Then
            obuX = True
            ubX = finalX
        Else
            obuX = False
        End If
                
        'As part of my accumulation algorithm, I swap the inner loop's direction with each iteration.
        ' Set y-related loop variables depending on the direction of the next cycle.
        If atBottom Then
            lbY = 0
            ubY = mRadius
        Else
            lbY = finalY - mRadius
            ubY = finalY
        End If
        
        'Remove trailing values from the median box if they lie outside the processing radius
        If lbX > 0 Then
        
            QuickValInner = (lbX - 1) * qvDepth
        
            For j = lbY To ubY
                r = srcImageData(QuickValInner + 2, j)
                g = srcImageData(QuickValInner + 1, j)
                b = srcImageData(QuickValInner, j)
                rValues(r) = rValues(r) - 1
                gValues(g) = gValues(g) - 1
                bValues(b) = bValues(b) - 1
                NumOfPixels = NumOfPixels - 1
            Next j
        
        End If
        
        'Add leading values to the median box if they lie inside the processing radius
        If Not obuX Then
        
            QuickValInner = ubX * qvDepth
            
            For j = lbY To ubY
                r = srcImageData(QuickValInner + 2, j)
                g = srcImageData(QuickValInner + 1, j)
                b = srcImageData(QuickValInner, j)
                rValues(r) = rValues(r) + 1
                gValues(g) = gValues(g) + 1
                bValues(b) = bValues(b) + 1
                NumOfPixels = NumOfPixels + 1
            Next j
            
        End If
        
        'Depending on the direction we are moving, remove a line of pixels from the median box
        ' (because the interior loop will add it back in).
        If atBottom Then
        
            For i = lbX To ubX
                QuickValInner = i * qvDepth
                r = srcImageData(QuickValInner + 2, mRadius)
                g = srcImageData(QuickValInner + 1, mRadius)
                b = srcImageData(QuickValInner, mRadius)
                rValues(r) = rValues(r) - 1
                gValues(g) = gValues(g) - 1
                bValues(b) = bValues(b) - 1
                NumOfPixels = NumOfPixels - 1
            Next i
       
        Else
       
            QuickY = finalY - mRadius
       
            For i = lbX To ubX
                QuickValInner = i * qvDepth
                r = srcImageData(QuickValInner + 2, QuickY)
                g = srcImageData(QuickValInner + 1, QuickY)
                b = srcImageData(QuickValInner, QuickY)
                rValues(r) = rValues(r) - 1
                gValues(g) = gValues(g) - 1
                bValues(b) = bValues(b) - 1
                NumOfPixels = NumOfPixels - 1
            Next i
       
        End If
        
        'Based on the direction we're traveling, reverse the interior loop boundaries as necessary.
        If atBottom Then
            startY = 0
            stopY = finalY
            yStep = 1
        Else
            startY = finalY
            stopY = 0
            yStep = -1
        End If
            
    'Process the next column.  This step is pretty much identical to the row steps above (but in a vertical direction, obviously)
    For y = startY To stopY Step yStep
            
        'If we are at the bottom and moving up, we will REMOVE rows from the bottom and ADD them at the top.
        'If we are at the top and moving down, we will REMOVE rows from the top and ADD them at the bottom.
        'As such, there are two copies of this function, one per possible direction.
        If atBottom Then
        
            'Calculate bounds
            lbY = y - mRadius
            If lbY < 0 Then lbY = 0
            
            ubY = y + mRadius
            If ubY > finalY Then
                obuY = True
                ubY = finalY
            Else
                obuY = False
            End If
                                
            'Remove trailing values from the box
            If lbY > 0 Then
            
                QuickY = lbY - 1
            
                For i = lbX To ubX
                    QuickValInner = i * qvDepth
                    r = srcImageData(QuickValInner + 2, QuickY)
                    g = srcImageData(QuickValInner + 1, QuickY)
                    b = srcImageData(QuickValInner, QuickY)
                    rValues(r) = rValues(r) - 1
                    gValues(g) = gValues(g) - 1
                    bValues(b) = bValues(b) - 1
                    NumOfPixels = NumOfPixels - 1
                Next i
                        
            End If
                    
            'Add leading values
            If Not obuY Then
            
                For i = lbX To ubX
                    QuickValInner = i * qvDepth
                    r = srcImageData(QuickValInner + 2, ubY)
                    g = srcImageData(QuickValInner + 1, ubY)
                    b = srcImageData(QuickValInner, ubY)
                    rValues(r) = rValues(r) + 1
                    gValues(g) = gValues(g) + 1
                    bValues(b) = bValues(b) + 1
                    NumOfPixels = NumOfPixels + 1
                Next i
            
            End If
            
        'The exact same code as above, but in the opposite direction
        Else
        
            lbY = y - mRadius
            If lbY < 0 Then
                oblY = True
                lbY = 0
            Else
                oblY = False
            End If
            
            ubY = y + mRadius
            If ubY > finalY Then ubY = finalY
                                
            If ubY < finalY Then
            
                QuickY = ubY + 1
            
                For i = lbX To ubX
                    QuickValInner = i * qvDepth
                    r = srcImageData(QuickValInner + 2, QuickY)
                    g = srcImageData(QuickValInner + 1, QuickY)
                    b = srcImageData(QuickValInner, QuickY)
                    rValues(r) = rValues(r) - 1
                    gValues(g) = gValues(g) - 1
                    bValues(b) = bValues(b) - 1
                    NumOfPixels = NumOfPixels - 1
                Next i
                        
            End If
                    
            If Not oblY Then
            
                For i = lbX To ubX
                    QuickValInner = i * qvDepth
                    r = srcImageData(QuickValInner + 2, lbY)
                    g = srcImageData(QuickValInner + 1, lbY)
                    b = srcImageData(QuickValInner, lbY)
                    rValues(r) = rValues(r) + 1
                    gValues(g) = gValues(g) + 1
                    bValues(b) = bValues(b) + 1
                    NumOfPixels = NumOfPixels + 1
                Next i
            
            End If
        
        End If
                
        'With the median box successfully calculated, we can now find the actual median for this pixel.
                
        'Loop through each color component histogram, until we've passed the desired percentile of pixels
        midR = 0
        midG = 0
        midB = 0
        cutoffTotal = (mPercent * NumOfPixels)
        If cutoffTotal = 0 Then cutoffTotal = 1
        
        i = -1
        Do
            i = i + 1
            If rValues(i) > 0 Then midR = midR + rValues(i)
        Loop Until (midR >= cutoffTotal)
        midR = i
        
        i = -1
        Do
            i = i + 1
            If gValues(i) > 0 Then midG = midG + gValues(i)
        Loop Until (midG >= cutoffTotal)
        midG = i
        
        i = -1
        Do
            i = i + 1
            If bValues(i) > 0 Then midB = midB + bValues(i)
        Loop Until (midB >= cutoffTotal)
        midB = i
                
        'Finally, apply the results to the image.
        dstImageData(QuickVal + 2, y) = midR
        dstImageData(QuickVal + 1, y) = midG
        dstImageData(QuickVal, y) = midB
        
    Next y
        atBottom = Not atBottom
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
    
    If cancelCurrentAction Then CreateMedianDIB = 0 Else CreateMedianDIB = 1

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
    Dim rMax As Byte, gMax As Byte, bMax As Byte
    Dim rMin As Byte, gMin As Byte, bMin As Byte
    rMax = 0: gMax = 0: bMax = 0
    rMin = 255: gMin = 255: bMin = 255
    
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
    
    Dim NumOfPixels As Long
    NumOfPixels = (finalX + 1) * (finalY + 1)
    
    Dim wbThreshold As Long
    wbThreshold = NumOfPixels * percentIgnore
    
    r = 0: g = 0: b = 0
    
    Dim rTally As Long, gTally As Long, bTally As Long
    rTally = 0: gTally = 0: bTally = 0
    
    'Find minimum values of red, green, and blue
    Do
        If rCount(r) + rTally < wbThreshold Then
            r = r + 1
            rTally = rTally + rCount(r)
        Else
            rMin = r
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
            rMax = r
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
    rDif = CLng(rMax) - CLng(rMin)
    gDif = CLng(gMax) - CLng(gMin)
    bDif = CLng(bMax) - CLng(bMin)
    
    'We can now build a final set of look-up tables that contain the results of every possible color transformation
    Dim rFinal(0 To 255) As Byte, gFinal(0 To 255) As Byte, bFinal(0 To 255) As Byte
    
    For x = 0 To 255
        If rDif <> 0 Then r = 255 * ((x - rMin) / rDif) Else r = x
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
Public Function AdjustDIBShadowHighlight(ByVal shadowClipping As Double, ByVal highlightClipping As Double, ByVal targetMidtone As Long, ByRef srcDIB As pdDIB, Optional ByVal suppressMessages As Boolean = False, Optional ByVal modifyProgBarMax As Long = -1, Optional ByVal modifyProgBarOffset As Long = 0) As Long

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
    Dim rMax As Byte, gMax As Byte, bMax As Byte
    Dim rMin As Byte, gMin As Byte, bMin As Byte
    rMax = 0: gMax = 0: bMax = 0
    rMin = 255: gMin = 255: bMin = 255
    
    'Shrink the percentIgnore value down to 1% of the value we are passed (you'll see why in a moment)
    shadowClipping = shadowClipping / 100
    highlightClipping = highlightClipping / 100
    
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
    
    Dim NumOfPixels As Long
    NumOfPixels = (finalX + 1) * (finalY + 1)
    
    Dim shadowThreshold As Long
    shadowThreshold = NumOfPixels * shadowClipping
    
    Dim highlightThreshold As Long
    highlightThreshold = NumOfPixels * highlightClipping
    
    r = 0: g = 0: b = 0
    
    Dim rTally As Long, gTally As Long, bTally As Long
    rTally = 0: gTally = 0: bTally = 0
    
    'Find minimum values of red, green, and blue
    Do
        If rCount(r) + rTally < shadowThreshold Then
            r = r + 1
            rTally = rTally + rCount(r)
        Else
            rMin = r
            foundYet = True
        End If
    Loop While foundYet = False
        
    foundYet = False
        
    Do
        If gCount(g) + gTally < shadowThreshold Then
            g = g + 1
            gTally = gTally + gCount(g)
        Else
            gMin = g
            foundYet = True
        End If
    Loop While foundYet = False
    
    foundYet = False
    
    Do
        If bCount(b) + bTally < shadowThreshold Then
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
        If rCount(r) + rTally < highlightThreshold Then
            r = r - 1
            rTally = rTally + rCount(r)
        Else
            rMax = r
            foundYet = True
        End If
    Loop While foundYet = False
        
    foundYet = False
        
    Do
        If gCount(g) + gTally < highlightThreshold Then
            g = g - 1
            gTally = gTally + gCount(g)
        Else
            gMax = g
            foundYet = True
        End If
    Loop While foundYet = False
    
    foundYet = False
    
    Do
        If bCount(b) + bTally < highlightThreshold Then
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
        If rDif <> 0 Then r = 255 * ((x - rMin) / rDif) Else r = x
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
    
    'Now it is time to handle the target midtone calculation.  Start by extracting the red, green, and blue components
    Dim targetRed As Long, targetGreen As Long, targetBlue As Long
    targetRed = 255 - ExtractR(targetMidtone)
    targetGreen = 255 - ExtractG(targetMidtone)
    targetBlue = 255 - ExtractB(targetMidtone)
    
    'We now re-use some logic from the Levels tool to remap midtones according to the target color we've been given.
    
    'Look-up tables for the midtone (gamma) leveled values
    Dim lValues(0 To 255) As Double
    
    'WARNING: This next chunk of code is a lot of messy math.  Don't worry too much
    ' if you can't make sense of it ;)
    
    'Fill the gamma table with appropriate gamma values (from 10 to .1, ranged quadratically)
    ' NOTE: This table is constant, and could theoretically be loaded from file instead of generated
    ' every time we run this function.
    Dim gStep As Double
    gStep = (MAXGAMMA + MIDGAMMA) / 127
    For x = 0 To 127
        lValues(x) = (CDbl(x) / 127) * MIDGAMMA
    Next x
    For x = 128 To 255
        lValues(x) = MIDGAMMA + (CDbl(x - 127) * gStep)
    Next x
    For x = 0 To 255
        lValues(x) = 1 / ((lValues(x) + 1 / ROOT10) ^ 2)
    Next x
    
    'Calculate a look-up table of gamma-corrected values based on the midtones scrollbar
    Dim rValues(0 To 255) As Byte, gValues(0 To 255) As Byte, bValues(0 To 255) As Byte
    Dim tmpRed As Double, tmpGreen As Double, tmpBlue As Double
    For x = 0 To 255
        tmpRed = CDbl(x) / 255
        tmpGreen = CDbl(x) / 255
        tmpBlue = CDbl(x) / 255
        tmpRed = tmpRed ^ (1 / lValues(targetRed))
        tmpGreen = tmpGreen ^ (1 / lValues(targetGreen))
        tmpBlue = tmpBlue ^ (1 / lValues(targetBlue))
        tmpRed = tmpRed * 255
        tmpGreen = tmpGreen * 255
        tmpBlue = tmpBlue * 255
        If tmpRed > 255 Then tmpRed = 255
        If tmpRed < 0 Then tmpRed = 0
        If tmpGreen > 255 Then tmpGreen = 255
        If tmpGreen < 0 Then tmpGreen = 0
        If tmpBlue > 255 Then tmpBlue = 255
        If tmpBlue < 0 Then tmpBlue = 0
        rValues(x) = tmpRed
        gValues(x) = tmpGreen
        bValues(x) = tmpBlue
    Next x
    
    'Now we can loop through each pixel in the image, converting values as we go
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
            
        'Adjust white balance in a single pass (thanks to the magic of look-up tables)
        ImageData(QuickVal + 2, y) = rValues(rFinal(ImageData(QuickVal + 2, y)))
        ImageData(QuickVal + 1, y) = gValues(gFinal(ImageData(QuickVal + 1, y)))
        ImageData(QuickVal, y) = bValues(bFinal(ImageData(QuickVal, y)))
        
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
        stdDev = 0.5
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
    Dim QuickVal As Long, QuickValInner As Long, QuickY As Long, qvDepth As Long
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
    Dim QuickVal As Long, QuickValInner As Long, QuickY As Long, qvDepth As Long
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
    Dim QuickVal As Long, QuickValInner As Long, QuickY As Long, qvDepth As Long
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
    Dim NumOfPixels As Long
    NumOfPixels = 0
            
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
        NumOfPixels = NumOfPixels + 1
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
            
            NumOfPixels = NumOfPixels - 1
        
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
            
            NumOfPixels = NumOfPixels + 1
            
        End If
            
        'Process the current column.  This simply involves calculating blur values, and applying them to the destination image
        For y = initY To finalY
                
            'With the blur box successfully calculated, we can finally apply the results to the image.
            dstImageData(QuickVal + 2, y) = rTotals(y) \ NumOfPixels
            dstImageData(QuickVal + 1, y) = gTotals(y) \ NumOfPixels
            dstImageData(QuickVal, y) = bTotals(y) \ NumOfPixels
            If qvDepth = 4 Then dstImageData(QuickVal + 3, y) = aTotals(y) \ NumOfPixels
    
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
    Dim NumOfPixels As Long
    NumOfPixels = 0
            
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
        NumOfPixels = NumOfPixels + 1
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
            
            NumOfPixels = NumOfPixels - 1
        
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
            
            NumOfPixels = NumOfPixels + 1
            
        End If
            
        'Process the current row.  This simply involves calculating blur values, and applying them to the destination image.
        For x = initX To finalX
            
            QuickVal = x * qvDepth
            
            'With the blur box successfully calculated, we can finally apply the results to the image.
            dstImageData(QuickVal + 2, y) = rTotals(x) \ NumOfPixels
            dstImageData(QuickVal + 1, y) = gTotals(x) \ NumOfPixels
            dstImageData(QuickVal, y) = bTotals(x) \ NumOfPixels
            If qvDepth = 4 Then dstImageData(QuickVal + 3, y) = aTotals(x) \ NumOfPixels
    
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
    Dim QuickVal As Long, QuickValInner As Long, QuickY As Long, qvDepth As Long
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
'Per PhotoDemon convention, this function will return a non-zero value if successful, and 0 if canceled.
Public Function CreateExtendedDIB(ByVal hExtend As Long, ByVal vExtend As Long, ByRef srcDIB As pdDIB, ByRef dstDIB As pdDIB) As Long

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
    CreateExtendedDIB = 1
    
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
