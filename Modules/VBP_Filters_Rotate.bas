Attribute VB_Name = "Filters_Transform"
'***************************************************************************
'Image Transformations Interface (including flip/mirror/rotation/crop/etc)
'Copyright ©2003-2014 by Tanner Helland
'Created: 25/January/03
'Last updated: 17/May/13
'Last update: CropToSelection now handles non-rectangular selections correctly.  (Unselected areas are made transparent.)
'
'Runs all image transformations, including rotate, flip, mirror and crop at present.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Automatically crop the image.  An optional threshold can be supplied; pixels must be this close before they will be cropped.
' (The threshold is required for JPEG images; pixels may not be identical due to lossy compression.)
Public Sub AutocropImage(Optional ByVal cThreshold As Long = 15)

    'TODO: rework this to operate on layers.  In theory, we can simply crop the pdImage width height, without
    '      actually modifying any individual layers!  The best way to do it may be to retrieve a composited
    '      copy of the image, autocrop it, then use its dimensions to change the original image's height/width.
    '      (NOTE: for left/top, all layer offsets will need to be adjusted to match.)

    'If the image contains an active selection, disable it before transforming the canvas
    If pdImages(g_CurrentImage).selectionActive Then
        pdImages(g_CurrentImage).selectionActive = False
        pdImages(g_CurrentImage).mainSelection.lockRelease
    End If

    'The image will be cropped in four steps.  Each edge will be cropped separately, starting with the top.
    
    Message "Analyzing top edge of image..."
    
    'Make a copy of the current image
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    'tmpDIB.createFromExistingDIB pdImages(g_CurrentImage).mainDIB
    
    'Point an array at the DIB data
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepSafeArray srcSA, tmpDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    finalX = pdImages(g_CurrentImage).Width - 1
    finalY = pdImages(g_CurrentImage).Height - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    'qvDepth = pdImages(g_CurrentImage).mainDIB.getDIBColorDepth \ 8
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    SetProgBarMax 4
    
    'Build a grayscale lookup table.  We will only be comparing luminance - not colors - when determining where to crop.
    Dim gLookup(0 To 765) As Long
    For x = 0 To 765
        gLookup(x) = CByte(x \ 3)
    Next x
    
    'The new edges of the image will mark these values for us
    Dim newTop As Long, newBottom As Long, newLeft As Long, newRight As Long
    
    'First, scan the top of the image.
    
    'All edges follow the same formula, so I'm only commenting this first section.
    
    '1-1) Start by determining the color of the top-left pixel.  This will be our baseline.
    Dim initColor As Long, curColor As Long
    initColor = gLookup(CLng(srcImageData(0, 0)) + CLng(srcImageData(1, 0)) + CLng(srcImageData(2, 0)))
    
    Dim colorFails As Boolean
    colorFails = False
    
    'Scan the image, starting at the top-left and moving right
    For y = 0 To finalY
    For x = 0 To finalX
        QuickVal = x * qvDepth
        curColor = gLookup(CLng(srcImageData(QuickVal, y)) + CLng(srcImageData(QuickVal + 1, y)) + CLng(srcImageData(QuickVal + 2, y)))
        
        'If pixel color DOES NOT match the baseline, keep scanning.  Otherwise, note that we have found a mismatched color
        ' and exit the loop.
        If Abs(curColor - initColor) > cThreshold Then colorFails = True
        
        If colorFails Then Exit For
        
    Next x
        If colorFails Then Exit For
    Next y
    
    'We have now reached one of two conditions:
    '1) The entire image is one solid color
    '2) The loop progressed part-way through the image and terminated
    
    'Check for case (1) and warn the user if it occurred
    If Not colorFails Then
        CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
        Erase srcImageData
        
        SetProgBarVal 0
        releaseProgressBar
        
        Message "Image is all one color.  Autocrop unnecessary."
        Exit Sub
    
    'Next, check for case (2)
    Else
        newTop = y
    End If
    
    initY = newTop
    
    'Repeat the above steps, but tracking the left edge instead.  Note also that we will only be scanning from wherever
    ' the top crop failed - this saves processing time.
    colorFails = False
    
    Message "Analyzing left edge of image..."
    initColor = gLookup(CLng(srcImageData(0, initY)) + CLng(srcImageData(1, initY)) + CLng(srcImageData(2, initY)))
    SetProgBarVal 1
    
    For x = 0 To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        curColor = gLookup(CLng(srcImageData(QuickVal, y)) + CLng(srcImageData(QuickVal + 1, y)) + CLng(srcImageData(QuickVal + 2, y)))
        
        'If pixel color DOES NOT match the baseline, keep scanning.  Otherwise, note that we have found a mismatched color
        ' and exit the loop.
        If Abs(curColor - initColor) > cThreshold Then colorFails = True
        
        If colorFails Then Exit For
        
    Next y
        If colorFails Then Exit For
    Next x
    
    newLeft = x
    
    'Repeat the above steps, but tracking the right edge instead.  Note also that we will only be scanning from wherever
    ' the top crop failed - this saves processing time.
    colorFails = False
    
    Message "Analyzing right edge of image..."
    QuickVal = finalX * qvDepth
    initColor = gLookup(CLng(srcImageData(QuickVal, initY)) + CLng(srcImageData(QuickVal + 1, 0)) + CLng(srcImageData(QuickVal + 2, 0)))
    SetProgBarVal 2
    
    For x = finalX To 0 Step -1
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        curColor = gLookup(CLng(srcImageData(QuickVal, y)) + CLng(srcImageData(QuickVal + 1, y)) + CLng(srcImageData(QuickVal + 2, y)))
        
        'If pixel color DOES NOT match the baseline, keep scanning.  Otherwise, note that we have found a mismatched color
        ' and exit the loop.
        If Abs(curColor - initColor) > cThreshold Then colorFails = True
        
        If colorFails Then Exit For
        
    Next y
        If colorFails Then Exit For
    Next x
    
    newRight = x
    
    'Finally, repeat the steps above for the bottom of the image.  Note also that we will only be scanning from wherever
    ' the left and right crops failed - this saves processing time.
    colorFails = False
    initX = newLeft
    finalX = newRight
    QuickVal = initX * qvDepth
    initColor = gLookup(CLng(srcImageData(QuickVal, finalY)) + CLng(srcImageData(QuickVal + 1, finalY)) + CLng(srcImageData(QuickVal + 2, finalY)))
    
    Message "Analyzing bottom edge of image..."
    SetProgBarVal 3
    
    For y = finalY To initY Step -1
    For x = initX To finalX
        QuickVal = x * qvDepth
        curColor = gLookup(CLng(srcImageData(QuickVal, y)) + CLng(srcImageData(QuickVal + 1, y)) + CLng(srcImageData(QuickVal + 2, y)))
        
        'If pixel color DOES NOT match the baseline, keep scanning.  Otherwise, note that we have found a mismatched color
        ' and exit the loop.
        If Abs(curColor - initColor) > cThreshold Then colorFails = True
        
        If colorFails Then Exit For
        
    Next x
        If colorFails Then Exit For
    Next y
    
    newBottom = y
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    
    'We now know where to crop the image.  Apply the crop.
    
    If (newTop = 0) And (newBottom = pdImages(g_CurrentImage).Height - 1) And (newLeft = 0) And (newRight = pdImages(g_CurrentImage).Width - 1) Then
        SetProgBarVal 0
        releaseProgressBar
        Message "Image is already cropped intelligently.  Autocrop abandoned.  (No changes were made to the image.)"
    Else
    
        Message "Cropping image to new dimensions..."
        SetProgBarVal 4
        
        'Resize the current image's main DIB
        'pdImages(g_CurrentImage).mainDIB.createBlank newRight - newLeft, newBottom - newTop, tmpDIB.getDIBColorDepth
        
        'Copy the autocropped area to the new main DIB
        'BitBlt pdImages(g_CurrentImage).mainDIB.getDIBDC, 0, 0, pdImages(g_CurrentImage).mainDIB.getDIBWidth, pdImages(g_CurrentImage).mainDIB.getDIBHeight, tmpDIB.getDIBDC, newLeft, newTop, vbSrcCopy
    
        'Erase the temporary DIB
        tmpDIB.eraseDIB
        Set tmpDIB = Nothing
    
        'Update the current image size
        pdImages(g_CurrentImage).updateSize
        DisplaySize pdImages(g_CurrentImage)
        
        Message "Finished. "
        SetProgBarVal 0
        releaseProgressBar
        
        'Redraw the image
        PrepareViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0), "Autocrop image"
    
    End If

End Sub

'Crop the image to the current selection
Public Sub MenuCropToSelection()

    'TODO: make this work with layers.

    'First, make sure there is an active selection
    If Not pdImages(g_CurrentImage).selectionActive Then
        Message "No active selection found.  Crop abandoned."
        Exit Sub
    End If
    
    Message "Cropping image to selected area..."
    
    'Create a new DIB the size of the active selection
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    pdImages(g_CurrentImage).retrieveProcessedSelection tmpDIB, True
    
    'NOTE: historically, the entire rectangular bounding region of the selection was included in the crop.  (This is GIMP's behavior.)
    ' I now fully crop the image, which means that for non-square selections, all unselected pixels are set to transparent.  For non-square
    ' selections, this will always result in a 32bpp image.
    '
    'The old code will be left here few a few releases, in case I decide to provide a preference for alternate behavior, per user request.
    ' (Note: comment added for the 6.0 release; consider removing by 6.4 if no complaints received.)
    '
    'Copy the selection area to the temporary DIB
    'tmpDIB.createBlank pdImages(g_CurrentImage).mainSelection.boundWidth, pdImages(g_CurrentImage).mainSelection.boundHeight, pdImages(g_CurrentImage).mainDIB.getDIBColorDepth
    'BitBlt tmpDIB.getDIBDC, 0, 0, pdImages(g_CurrentImage).mainSelection.boundWidth, pdImages(g_CurrentImage).mainSelection.boundHeight, pdImages(g_CurrentImage).mainDIB.getDIBDC, pdImages(g_CurrentImage).mainSelection.boundLeft, pdImages(g_CurrentImage).mainSelection.boundTop, vbSrcCopy
    
    'Transfer the newly cropped image back into the main DIB object
    'pdImages(g_CurrentImage).mainDIB.createFromExistingDIB tmpDIB
    
    'Erase the temporary DIB
    tmpDIB.eraseDIB
    Set tmpDIB = Nothing
    
    'Update the current image size
    pdImages(g_CurrentImage).updateSize
    DisplaySize pdImages(g_CurrentImage)
    
    Message "Finished. "
    
    'Deactivate the current selection, as it's no longer needed
    'Clear selections after "Crop to Selection"
    If g_UserPreferences.GetPref_Boolean("Tools", "Clear Selection After Crop", True) Then
        pdImages(g_CurrentImage).selectionActive = False
        pdImages(g_CurrentImage).mainSelection.lockRelease
        Message "Crop complete.  (Note: the selected area was automatically unselected.)"
    Else
        pdImages(g_CurrentImage).mainSelection.lockRelease
        pdImages(g_CurrentImage).mainSelection.selLeft = 0
        pdImages(g_CurrentImage).mainSelection.selTop = 0
        pdImages(g_CurrentImage).mainSelection.selWidth = pdImages(g_CurrentImage).Width
        pdImages(g_CurrentImage).mainSelection.selHeight = pdImages(g_CurrentImage).Height
        pdImages(g_CurrentImage).mainSelection.lockIn
        Dim i As Long
        For i = 0 To toolbar_Tools.cmbSelRender.Count - 1
            toolbar_Tools.cmbSelRender(i).ListIndex = sHighlightRed
        Next i
        Message "Crop complete.  Selection drawing mode changed to make selection visible."
    End If
    
    'Redraw the image
    PrepareViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0), "Crop to selection"

End Sub

'Flip an image vertically
Public Sub MenuFlip()

    'TODO: make this function work with layers.

    'If the image contains an active selection, disable it before transforming the canvas
    If pdImages(g_CurrentImage).selectionActive Then
        pdImages(g_CurrentImage).selectionActive = False
        pdImages(g_CurrentImage).mainSelection.lockRelease
    End If
    
    Message "Flipping image..."
    'StretchBlt pdImages(g_CurrentImage).mainDIB.getDIBDC, 0, 0, pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height, pdImages(g_CurrentImage).mainDIB.getDIBDC, 0, pdImages(g_CurrentImage).Height - 1, pdImages(g_CurrentImage).Width, -pdImages(g_CurrentImage).Height, vbSrcCopy
    Message "Finished. "
    
    ScrollViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

'Flip an image horizontally
Public Sub MenuMirror()
    
    'TODO: make this function work with layers.
    
    'If the image contains an active selection, disable it before transforming the canvas
    If pdImages(g_CurrentImage).selectionActive Then
        pdImages(g_CurrentImage).selectionActive = False
        pdImages(g_CurrentImage).mainSelection.lockRelease
    End If

    Message "Mirroring image..."
    'StretchBlt pdImages(g_CurrentImage).mainDIB.getDIBDC, 0, 0, pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height, pdImages(g_CurrentImage).mainDIB.getDIBDC, pdImages(g_CurrentImage).Width - 1, 0, -pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height, vbSrcCopy
    Message "Finished. "
    
    ScrollViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

'Rotate an image 90° clockwise
Public Sub MenuRotate90Clockwise()

    'TODO: make this function work with layers.

    'If a selection is active, remove it.  (This is not the most elegant solution - the elegant solution would be rotating
    ' the selection to match the new image, but we can fix that at a later date.)
    If pdImages(g_CurrentImage).selectionActive Then
        pdImages(g_CurrentImage).selectionActive = False
        pdImages(g_CurrentImage).mainSelection.lockRelease
    End If

    Message "Rotating image clockwise..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepImageData srcSA
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
    
    'Create a second local array.  This will contain the pixel data of the new, rotated image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    
    Dim dstDIB As pdDIB
    Set dstDIB = New pdDIB
    'dstDIB.createBlank pdImages(g_CurrentImage).mainDIB.getDIBHeight, pdImages(g_CurrentImage).mainDIB.getDIBWidth, pdImages(g_CurrentImage).mainDIB.getDIBColorDepth
    
    prepSafeArray dstSA, dstDIB
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, i As Long
    Dim initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long, QuickValY
    qvDepth = curDIBValues.BytesPerPixel
    
    Dim iWidth As Long, iHeight As Long
    iWidth = finalX * qvDepth
    iHeight = finalY * qvDepth
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
        
    'Rotate the source image into the destination image, using the arrays provided
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
        QuickValY = y * qvDepth
        
        For i = 0 To qvDepth - 1
            dstImageData(iHeight - QuickValY + i, finalX - x) = srcImageData(iWidth - QuickVal + i, y)
        Next i
        
    Next y
        If (x And progBarCheck) = 0 Then SetProgBarVal x
    Next x
    
    'With our work complete, point both ImageData() arrays away from their respective DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    'If the original image was 32bpp, we need to re-apply premultiplication (because prepImageData above removed it)
    If dstDIB.getDIBColorDepth = 32 Then dstDIB.fixPremultipliedAlpha True
    
    'dstImageData now contains the rotated image.  We need to transfer that back into the current image.
    'pdImages(g_CurrentImage).mainDIB.createFromExistingDIB dstDIB
    
    'With that transfer complete, we can erase our temporary DIB
    dstDIB.eraseDIB
    Set dstDIB = Nothing
    
    'Update the current image size
    pdImages(g_CurrentImage).updateSize
    DisplaySize pdImages(g_CurrentImage)
    
    Message "Finished. "
    
    PrepareViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0), "image rotated"
    
    'Reset the progress bar to zero
    SetProgBarVal 0
    releaseProgressBar
    
End Sub

'Rotate an image 180°
Public Sub MenuRotate180()

    'TODO: make this function work with layers.

    'If the image contains an active selection, disable it before transforming the canvas
    If pdImages(g_CurrentImage).selectionActive Then
        pdImages(g_CurrentImage).selectionActive = False
        pdImages(g_CurrentImage).mainSelection.lockRelease
    End If

    'Fun fact: rotating 180 degrees can be accomplished by flipping and then mirroring it.
    Message "Rotating image..."
        
    'StretchBlt pdImages(g_CurrentImage).mainDIB.getDIBDC, 0, 0, pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height, pdImages(g_CurrentImage).mainDIB.getDIBDC, pdImages(g_CurrentImage).Width - 1, pdImages(g_CurrentImage).Height - 1, -pdImages(g_CurrentImage).Width, -pdImages(g_CurrentImage).Height, vbSrcCopy
        
    Message "Finished. "
    
    ScrollViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

'Rotate an image 90° counter-clockwise
Public Sub MenuRotate270Clockwise()

    'TODO: make this function work with layers.

    'If a selection is active, remove it.  (This is not the most elegant solution - the elegant solution would be rotating
    ' the selection to match the new image, but we can fix that at a later date.)
    If pdImages(g_CurrentImage).selectionActive Then
        pdImages(g_CurrentImage).selectionActive = False
        pdImages(g_CurrentImage).mainSelection.lockRelease
    End If

    Message "Rotating image counter-clockwise..."
    
    'Create a local array and point it at the pixel data of the current image
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepImageData srcSA
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
    
    'Create a second local array.  This will contain the pixel data of the new, rotated image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    
    Dim dstDIB As pdDIB
    Set dstDIB = New pdDIB
    'dstDIB.createBlank pdImages(g_CurrentImage).mainDIB.getDIBHeight, pdImages(g_CurrentImage).mainDIB.getDIBWidth, pdImages(g_CurrentImage).mainDIB.getDIBColorDepth
    
    prepSafeArray dstSA, dstDIB
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, i As Long
    Dim initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curDIBValues.Left
    initY = curDIBValues.Top
    finalX = curDIBValues.Right
    finalY = curDIBValues.Bottom
    
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long, QuickValY
    qvDepth = curDIBValues.BytesPerPixel
    
    Dim iWidth As Long
    iWidth = finalX * qvDepth
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long
    progBarCheck = findBestProgBarValue()
        
    'Rotate the source image into the destination image, using the arrays provided
    For x = initX To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
        QuickValY = y * qvDepth
        
        For i = 0 To qvDepth - 1
            dstImageData(QuickValY + i, x) = srcImageData(iWidth - QuickVal + i, y)
        Next i
        
    Next y
        If (x And progBarCheck) = 0 Then SetProgBarVal x
    Next x
    
    'With our work complete, point both ImageData() arrays away from their respective DIBs and deallocate them
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    Erase dstImageData
    
    'If the original image was 32bpp, we need to re-apply premultiplication (because prepImageData above removed it)
    If dstDIB.getDIBColorDepth = 32 Then dstDIB.fixPremultipliedAlpha True
    
    'dstImageData now contains the rotated image.  We need to transfer that back into the current image.
    'pdImages(g_CurrentImage).mainDIB.createFromExistingDIB dstDIB
    
    'With that transfer complete, we can erase our temporary DIB
    dstDIB.eraseDIB
    Set dstDIB = Nothing
    
    'Update the current image size
    pdImages(g_CurrentImage).updateSize
    DisplaySize pdImages(g_CurrentImage)
    
    Message "Finished. "
    
    'Redraw the image
    PrepareViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0), "image rotated"
    
    'Reset the progress bar to zero
    SetProgBarVal 0
    releaseProgressBar
    
End Sub
