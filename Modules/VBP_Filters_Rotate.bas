Attribute VB_Name = "Filters_Transform"
'***************************************************************************
'Image Transformations Interface (including flip/mirror/rotation/crop/etc)
'Copyright ©2000-2013 by Tanner Helland
'Created: 25/January/03
'Last updated: 05/October/12
'Last update: Added cropping to selection.
'
'Runs all image transformations, including rotate, flip, mirror and crop at present.
'
'***************************************************************************

Option Explicit

'Automatically crop the image.  An optional threshold can be supplied; pixels must be this close before they will be cropped.
' (The threshold is required for JPEG images; pixels may not be identical due to lossy compression.)
Public Sub AutocropImage(Optional ByVal cThreshold As Long = 15)

    'The image will be cropped in four steps.  Each edge will be cropped separately, starting with the top.
    
    Message "Analyzing top edge of image..."
    
    'Make a copy of the current image
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer
    tmpLayer.createFromExistingLayer pdImages(CurrentImage).mainLayer
    
    'Point an array at the DIB data
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    prepSafeArray srcSA, tmpLayer
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    finalX = pdImages(CurrentImage).Width - 1
    finalY = pdImages(CurrentImage).Height - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long
    qvDepth = pdImages(CurrentImage).mainLayer.getLayerColorDepth \ 8

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
    
    If (newTop = 0) And (newBottom = pdImages(CurrentImage).Height - 1) And (newLeft = 0) And (newRight = pdImages(CurrentImage).Width - 1) Then
        SetProgBarVal 0
        Message "Image is already cropped intelligently.  Autocrop abandoned.  (No changes were made to the image.)"
    Else
    
        Message "Cropping image to new dimensions..."
        SetProgBarVal 4
        
        'Resize the current image's main layer
        pdImages(CurrentImage).mainLayer.createBlank newRight - newLeft, newBottom - newTop, tmpLayer.getLayerColorDepth
        
        'Copy the autocropped area to the new main layer
        BitBlt pdImages(CurrentImage).mainLayer.getLayerDC, 0, 0, pdImages(CurrentImage).mainLayer.getLayerWidth, pdImages(CurrentImage).mainLayer.getLayerHeight, tmpLayer.getLayerDC, newLeft, newTop, vbSrcCopy
    
        'Erase the temporary layer
        tmpLayer.eraseLayer
        Set tmpLayer = Nothing
    
        'Update the current image size
        pdImages(CurrentImage).updateSize
        DisplaySize pdImages(CurrentImage).Width, pdImages(CurrentImage).Height
        
        Message "Finished. "
        SetProgBarVal 0
        
        'Redraw the image
        PrepareViewport FormMain.ActiveForm, "Autocrop image"
    
    End If

End Sub

'Crop the image to the current selection
Public Sub MenuCropToSelection()

    'First, make sure there is an active selection
    If pdImages(CurrentImage).selectionActive = False Then
        Message "No active selection found.  Crop abandoned."
        Exit Sub
    End If
    
    Message "Cropping image to selected area..."
    
    'Create a new layer the size of the active selection
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer
    tmpLayer.createBlank pdImages(CurrentImage).mainSelection.selWidth, pdImages(CurrentImage).mainSelection.selHeight, pdImages(CurrentImage).mainLayer.getLayerColorDepth
    
    'Copy the selection area to the temporary layer
    BitBlt tmpLayer.getLayerDC, 0, 0, pdImages(CurrentImage).mainSelection.selWidth, pdImages(CurrentImage).mainSelection.selHeight, pdImages(CurrentImage).mainLayer.getLayerDC, pdImages(CurrentImage).mainSelection.selLeft, pdImages(CurrentImage).mainSelection.selTop, vbSrcCopy
    
    'Transfer the newly cropped image back into the main layer object
    pdImages(CurrentImage).mainLayer.createFromExistingLayer tmpLayer
    
    'Erase the temporary layer
    tmpLayer.eraseLayer
    Set tmpLayer = Nothing
    
    'Update the current image size
    pdImages(CurrentImage).updateSize
    DisplaySize pdImages(CurrentImage).Width, pdImages(CurrentImage).Height
    
    Message "Finished. "
    
    'Deactivate the current selection, as it's no longer needed
    'Clear selections after "Crop to Selection"
    If g_UserPreferences.GetPreference_Boolean("Tool Preferences", "ClearSelectionAfterCrop", True) Then
        pdImages(CurrentImage).selectionActive = False
        tInit tSelection, False
        Message "Crop complete.  (Note: the selected area was automatically unselected.)"
    Else
        pdImages(CurrentImage).mainSelection.lockRelease
        pdImages(CurrentImage).mainSelection.selLeft = 0
        pdImages(CurrentImage).mainSelection.selTop = 0
        pdImages(CurrentImage).mainSelection.selWidth = pdImages(CurrentImage).Width
        pdImages(CurrentImage).mainSelection.selHeight = pdImages(CurrentImage).Height
        pdImages(CurrentImage).mainSelection.refreshTextBoxes
        pdImages(CurrentImage).mainSelection.lockIn pdImages(CurrentImage).containingForm
        g_selectionRenderPreference = sHighlightRed
        FormMain.cmbSelRender.ListIndex = 2
        Message "Crop complete.  Selection drawing mode changed to make selection visible."
    End If
    
    'Redraw the image
    PrepareViewport FormMain.ActiveForm, "Crop to selection"

End Sub

'Flip an image vertically
Public Sub MenuFlip()

    Message "Flipping image..."
    
    StretchBlt pdImages(CurrentImage).mainLayer.getLayerDC, 0, 0, pdImages(CurrentImage).Width, pdImages(CurrentImage).Height, pdImages(CurrentImage).mainLayer.getLayerDC, 0, pdImages(CurrentImage).Height - 1, pdImages(CurrentImage).Width, -pdImages(CurrentImage).Height, vbSrcCopy
        
    Message "Finished. "
    
    ScrollViewport FormMain.ActiveForm
    
End Sub

'Flip an image horizontally
Public Sub MenuMirror()

    Message "Mirroring image..."
    
    StretchBlt pdImages(CurrentImage).mainLayer.getLayerDC, 0, 0, pdImages(CurrentImage).Width, pdImages(CurrentImage).Height, pdImages(CurrentImage).mainLayer.getLayerDC, pdImages(CurrentImage).Width - 1, 0, -pdImages(CurrentImage).Width, pdImages(CurrentImage).Height, vbSrcCopy
    
    Message "Finished. "
    
    ScrollViewport FormMain.ActiveForm
    
End Sub

'Rotate an image 90° clockwise
Public Sub MenuRotate90Clockwise()

    'If a selection is active, remove it.  (This is not the most elegant solution - the elegant solution would be rotating
    ' the selection to match the new image, but we can fix that at a later date.)
    If pdImages(CurrentImage).selectionActive Then
        pdImages(CurrentImage).selectionActive = False
        tInit tSelection, False
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
    
    Dim dstLayer As pdLayer
    Set dstLayer = New pdLayer
    dstLayer.createBlank pdImages(CurrentImage).mainLayer.getLayerHeight, pdImages(CurrentImage).mainLayer.getLayerWidth, pdImages(CurrentImage).mainLayer.getLayerColorDepth
    
    prepSafeArray dstSA, dstLayer
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, i As Long
    Dim initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
    
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long, QuickValY
    qvDepth = curLayerValues.BytesPerPixel
    
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
    
    'dstImageData now contains the rotated image.  We need to transfer that back into the current image.
    pdImages(CurrentImage).mainLayer.createFromExistingLayer dstLayer
    
    'With that transfer complete, we can erase our temporary layer
    dstLayer.eraseLayer
    Set dstLayer = Nothing
    
    'Update the current image size
    pdImages(CurrentImage).updateSize
    DisplaySize pdImages(CurrentImage).Width, pdImages(CurrentImage).Height
    
    Message "Finished. "
    
    'Redraw the image
    FitWindowToImage
    
    'Reset the progress bar to zero
    SetProgBarVal 0
    
End Sub

'Rotate an image 180°
Public Sub MenuRotate180()

    'Fun fact: rotating 180 degrees can be accomplished by flipping and then mirroring it.
    Message "Rotating image..."
        
    StretchBlt pdImages(CurrentImage).mainLayer.getLayerDC, 0, 0, pdImages(CurrentImage).Width, pdImages(CurrentImage).Height, pdImages(CurrentImage).mainLayer.getLayerDC, pdImages(CurrentImage).Width - 1, pdImages(CurrentImage).Height - 1, -pdImages(CurrentImage).Width, -pdImages(CurrentImage).Height, vbSrcCopy
        
    Message "Finished. "
    
    ScrollViewport FormMain.ActiveForm
    
End Sub

'Rotate an image 90° counter-clockwise
Public Sub MenuRotate270Clockwise()

    'If a selection is active, remove it.  (This is not the most elegant solution - the elegant solution would be rotating
    ' the selection to match the new image, but we can fix that at a later date.)
    If pdImages(CurrentImage).selectionActive Then
        pdImages(CurrentImage).selectionActive = False
        tInit tSelection, False
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
    
    Dim dstLayer As pdLayer
    Set dstLayer = New pdLayer
    dstLayer.createBlank pdImages(CurrentImage).mainLayer.getLayerHeight, pdImages(CurrentImage).mainLayer.getLayerWidth, pdImages(CurrentImage).mainLayer.getLayerColorDepth
    
    prepSafeArray dstSA, dstLayer
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
        
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, i As Long
    Dim initX As Long, initY As Long, finalX As Long, finalY As Long
    initX = curLayerValues.Left
    initY = curLayerValues.Top
    finalX = curLayerValues.Right
    finalY = curLayerValues.Bottom
    
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim QuickVal As Long, qvDepth As Long, QuickValY
    qvDepth = curLayerValues.BytesPerPixel
    
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
    
    'dstImageData now contains the rotated image.  We need to transfer that back into the current image.
    pdImages(CurrentImage).mainLayer.createFromExistingLayer dstLayer
    
    'With that transfer complete, we can erase our temporary layer
    dstLayer.eraseLayer
    Set dstLayer = Nothing
    
    'Update the current image size
    pdImages(CurrentImage).updateSize
    DisplaySize pdImages(CurrentImage).Width, pdImages(CurrentImage).Height
    
    Message "Finished. "
    
    'Redraw the image
    FitWindowToImage
    
    'Reset the progress bar to zero
    SetProgBarVal 0
    
End Sub
