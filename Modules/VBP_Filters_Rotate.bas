Attribute VB_Name = "Filters_Transform"
'***************************************************************************
'Image Transformations Interface (including flip/mirror/rotation/crop/etc)
'Copyright ©2003-2014 by Tanner Helland
'Created: 25/January/03
'Last updated: 08/May/14
'Last update: convert rotate 90/270 functions to use GDI+; this gives a small but welcome speed improvement
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
    Dim gLookUp(0 To 765) As Long
    For x = 0 To 765
        gLookUp(x) = CByte(x \ 3)
    Next x
    
    'The new edges of the image will mark these values for us
    Dim newTop As Long, newBottom As Long, newLeft As Long, newRight As Long
    
    'First, scan the top of the image.
    
    'All edges follow the same formula, so I'm only commenting this first section.
    
    '1-1) Start by determining the color of the top-left pixel.  This will be our baseline.
    Dim initColor As Long, curColor As Long
    initColor = gLookUp(CLng(srcImageData(0, 0)) + CLng(srcImageData(1, 0)) + CLng(srcImageData(2, 0)))
    
    Dim colorFails As Boolean
    colorFails = False
    
    'Scan the image, starting at the top-left and moving right
    For y = 0 To finalY
    For x = 0 To finalX
        QuickVal = x * qvDepth
        curColor = gLookUp(CLng(srcImageData(QuickVal, y)) + CLng(srcImageData(QuickVal + 1, y)) + CLng(srcImageData(QuickVal + 2, y)))
        
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
    initColor = gLookUp(CLng(srcImageData(0, initY)) + CLng(srcImageData(1, initY)) + CLng(srcImageData(2, initY)))
    SetProgBarVal 1
    
    For x = 0 To finalX
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        curColor = gLookUp(CLng(srcImageData(QuickVal, y)) + CLng(srcImageData(QuickVal + 1, y)) + CLng(srcImageData(QuickVal + 2, y)))
        
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
    initColor = gLookUp(CLng(srcImageData(QuickVal, initY)) + CLng(srcImageData(QuickVal + 1, 0)) + CLng(srcImageData(QuickVal + 2, 0)))
    SetProgBarVal 2
    
    For x = finalX To 0 Step -1
        QuickVal = x * qvDepth
    For y = initY To finalY
    
        curColor = gLookUp(CLng(srcImageData(QuickVal, y)) + CLng(srcImageData(QuickVal + 1, y)) + CLng(srcImageData(QuickVal + 2, y)))
        
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
    initColor = gLookUp(CLng(srcImageData(QuickVal, finalY)) + CLng(srcImageData(QuickVal + 1, finalY)) + CLng(srcImageData(QuickVal + 2, finalY)))
    
    Message "Analyzing bottom edge of image..."
    SetProgBarVal 3
    
    For y = finalY To initY Step -1
    For x = initX To finalX
        QuickVal = x * qvDepth
        curColor = gLookUp(CLng(srcImageData(QuickVal, y)) + CLng(srcImageData(QuickVal + 1, y)) + CLng(srcImageData(QuickVal + 2, y)))
        
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

'Crop the image to the current selection.
Public Sub MenuCropToSelection()
    
    'First, make sure there is an active selection
    If Not pdImages(g_CurrentImage).selectionActive Then
        Message "No active selection found.  Crop abandoned."
        Exit Sub
    End If
    
    Message "Cropping image to selected area..."
    
    'NOTE: historically, the entire rectangular bounding region of the selection was included in the crop.  (This is GIMP's behavior.)
    ' I now fully crop the image, which means that for non-square selections, all unselected pixels are set to transparent.  For non-square
    ' selections, this will always result in an image with some transparent regions.
    
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    'Arrays will be pointed at three sets of pixels: the current layer, the selection mask, and a destination layer.
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    
    'Point our selection array at the selection mask in advance; this only needs to be done once, as the same mask is used for all layers.
    Dim selData() As Byte
    Dim selSA As SAFEARRAY2D
    prepSafeArray selSA, pdImages(g_CurrentImage).mainSelection.selMask
    CopyMemory ByVal VarPtrArray(selData()), VarPtr(selSA), 4
    
    'Lots of helper variables for a function like this
    Dim leftOffset As Long, topOffset As Long
    leftOffset = pdImages(g_CurrentImage).mainSelection.boundLeft
    topOffset = pdImages(g_CurrentImage).mainSelection.boundTop
    
    Dim r As Long, g As Long, b As Long
    Dim thisAlpha As Long, origAlpha As Long, blendAlpha As Double
    Dim srcQuickX As Long, srcQuickY As Long, dstQuickX As Long, selQuickX As Long
    
    Dim x As Long, y As Long
    Dim imgWidth As Long, imgHeight As Long
    imgWidth = pdImages(g_CurrentImage).Width
    imgHeight = pdImages(g_CurrentImage).Height
    
    Dim selectionWidth As Long, selectionHeight As Long
    selectionWidth = pdImages(g_CurrentImage).mainSelection.boundWidth
    selectionHeight = pdImages(g_CurrentImage).mainSelection.boundHeight
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    Dim progBarCheck As Long, progBarOffsetX As Long
    SetProgBarMax pdImages(g_CurrentImage).getNumOfLayers * imgWidth
    progBarCheck = findBestProgBarValue()
    
    'Iterate through each layer, rotating them in turn
    Dim tmpLayerRef As pdLayer
    
    Dim i As Long
    For i = 0 To pdImages(g_CurrentImage).getNumOfLayers - 1
    
        'Update the progress bar counter for this layer
        progBarOffsetX = i * imgWidth
    
        'Retrieve a pointer to the layer of interest
        Set tmpLayerRef = pdImages(g_CurrentImage).getLayerByIndex(i)
        
        'Null-pad the layer
        tmpLayerRef.convertToNullPaddedLayer pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height
        
        'Create a temporary layer at the relevant size of the selection, and retrieve a pointer to its pixel data
        tmpDIB.createBlank selectionWidth, selectionHeight, 32, 0
        prepSafeArray dstSA, tmpDIB
        CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
        
        'Point another array at the original image layer
        prepSafeArray srcSA, tmpLayerRef.layerDIB
        CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
        
        'Iterate through all relevant pixels in this layer (e.g. only those that actually lie within the interesting region
        ' of the selection), copying them to the destination as necessary.
        For x = 0 To selectionWidth - 1
            dstQuickX = x * 4
            srcQuickX = (leftOffset + x) * 4
            selQuickX = (leftOffset + x) * 3
        For y = 0 To selectionHeight - 1
        
            srcQuickY = topOffset + y
            thisAlpha = selData(selQuickX, srcQuickY)
            
            If thisAlpha > 0 Then
            
                'Check the image's alpha value.  If it's zero, we have no reason to process it further
                origAlpha = srcImageData(srcQuickX + 3, srcQuickY)
                
                If origAlpha > 0 Then
                    
                    'Source pixel data will be premultiplied, which saves us a bunch of processing time.  (That is why
                    ' we premultiply alpha, after all!)
                    r = srcImageData(srcQuickX + 2, srcQuickY)
                    g = srcImageData(srcQuickX + 1, srcQuickY)
                    b = srcImageData(srcQuickX, srcQuickY)
                    
                    'Calculate a new multiplier, based on the strength of the selection at this location
                    blendAlpha = thisAlpha / 255
                    
                    'Apply the multiplier to the existing pixel data (which is already premultiplied, saving us a bunch of time now)
                    dstImageData(dstQuickX + 2, y) = r * blendAlpha
                    dstImageData(dstQuickX + 1, y) = g * blendAlpha
                    dstImageData(dstQuickX, y) = b * blendAlpha
                    
                    'Finish our work by calculating a new alpha channel value for this pixel, which is a blend of
                    ' the original alpha value, and the selection mask value at this location.
                    dstImageData(dstQuickX + 3, y) = origAlpha * blendAlpha
                    
                End If
                
            End If
            
        Next y
            If ((progBarOffsetX + x) And progBarCheck) = 0 Then SetProgBarVal (progBarOffsetX + x)
        Next x
        
        'With our work complete, point both ImageData() arrays away from their respective DIBs and deallocate them
        CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
        Erase srcImageData
        CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
        Erase dstImageData
        
        'Replace the current layer DIB with our destination one
        tmpLayerRef.layerDIB.createFromExistingDIB tmpDIB
                
        'Release our temporary DIB
        tmpDIB.eraseDIB
        
        'Remove any null-padding from the layer
        tmpLayerRef.cropNullPaddedLayer
        
        'Notify the target layer that its DIB data has been changed; the layer will use this to regenerate various internal caches
        tmpLayerRef.notifyLayerModified
        
    Next i
    
    'Clear the selection mask array reference
    CopyMemory ByVal VarPtrArray(selData), 0&, 4
    Erase selData
    
    'Update the current image size
    pdImages(g_CurrentImage).updateSize False, selectionWidth, selectionHeight
    DisplaySize pdImages(g_CurrentImage)
    
    'Deactivate the current selection, as it's no longer needed
    ' NOTE: this preference is still created in new preference files, but it is no longer exposed to the user.  I'm leaving the preference check
    '        here (just commented out) in case this is ever revisited, but for now, the active selection is always cleared after a crop has been
    '        applied.  The work involved in properly translating the crop (including as-of-yet unwritten code for non-transformable crops) is
    '        extensive, and its dubious benefit has left me unwilling to work on it just yet.
    'If g_UserPreferences.GetPref_Boolean("Tools", "Clear Selection After Crop", True) Then
        pdImages(g_CurrentImage).selectionActive = False
        pdImages(g_CurrentImage).mainSelection.lockRelease
        Message "Crop complete.  (Note: the selected area was automatically unselected.)"
    'End If
    
    PrepareViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0), "Crop to selection"
    
    'Reset the progress bar to zero
    SetProgBarVal 0
    releaseProgressBar
    
End Sub

'Flip an image vertically.  If no layer is specified (e.g. if targetLayerIndex = -1), all layers will be flipped.
Public Sub MenuFlip(Optional ByVal targetLayerIndex As Long = -1)

    Dim flipAllLayers As Boolean
    If targetLayerIndex = -1 Then flipAllLayers = True Else flipAllLayers = False

    'If the image contains an active selection, disable it before transforming the canvas
    If flipAllLayers And pdImages(g_CurrentImage).selectionActive Then
        pdImages(g_CurrentImage).selectionActive = False
        pdImages(g_CurrentImage).mainSelection.lockRelease
    End If
    
    Message "Flipping image..."
    
    'Iterate through each layer, flipping them in turn
    Dim tmpLayerRef As pdLayer
    
    Dim i As Long, lStart As Long, lEnd As Long
    
    'If the user wants us to process all layers, we will iterate through the full layer stack, applying the transformation to each in turn.
    ' Otherwise, we will only transform the specified layer.  To cut down on code duplication, we simply modify the endpoints of the loop.
    If flipAllLayers Then
        lStart = 0
        lEnd = pdImages(g_CurrentImage).getNumOfLayers - 1
    Else
        lStart = targetLayerIndex
        lEnd = targetLayerIndex
    End If
    
    'Loop through all relevant layers, transforming each as we go
    For i = lStart To lEnd
    
        'Retrieve a pointer to the layer of interest
        Set tmpLayerRef = pdImages(g_CurrentImage).getLayerByIndex(i)
        
        'Null-pad the layer
        If flipAllLayers Then tmpLayerRef.convertToNullPaddedLayer pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height
        
        'Flip it
        StretchBlt tmpLayerRef.layerDIB.getDIBDC, 0, 0, tmpLayerRef.layerDIB.getDIBWidth, tmpLayerRef.layerDIB.getDIBHeight, tmpLayerRef.layerDIB.getDIBDC, 0, tmpLayerRef.layerDIB.getDIBHeight - 1, tmpLayerRef.layerDIB.getDIBWidth, -tmpLayerRef.layerDIB.getDIBHeight, vbSrcCopy
        
        'Remove any null-padding
        If flipAllLayers Then tmpLayerRef.cropNullPaddedLayer
        
        'Notify the target layer that its DIB data has been changed; the layer will use this to regenerate various internal caches
        tmpLayerRef.notifyLayerModified
    
    Next i
    
    Message "Finished. "
    
    'Redraw the viewport
    ScrollViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

'Flip an image horizontally (mirror)
Public Sub MenuMirror(Optional ByVal targetLayerIndex As Long = -1)

    Dim flipAllLayers As Boolean
    If targetLayerIndex = -1 Then flipAllLayers = True Else flipAllLayers = False
    
    'If the image contains an active selection, disable it before transforming the canvas
    If flipAllLayers And pdImages(g_CurrentImage).selectionActive Then
        pdImages(g_CurrentImage).selectionActive = False
        pdImages(g_CurrentImage).mainSelection.lockRelease
    End If

    Message "Mirroring image..."
    
    'Iterate through each layer, mirroring them in turn
    Dim tmpLayerRef As pdLayer
    
    Dim i As Long, lStart As Long, lEnd As Long
    
    'If the user wants us to process all layers, we will iterate through the full layer stack, applying the transformation to each in turn.
    ' Otherwise, we will only transform the specified layer.  To cut down on code duplication, we simply modify the endpoints of the loop.
    If flipAllLayers Then
        lStart = 0
        lEnd = pdImages(g_CurrentImage).getNumOfLayers - 1
    Else
        lStart = targetLayerIndex
        lEnd = targetLayerIndex
    End If
    
    'Loop through all relevant layers, transforming each as we go
    For i = lStart To lEnd
    
        'Retrieve a pointer to the layer of interest
        Set tmpLayerRef = pdImages(g_CurrentImage).getLayerByIndex(i)
        
        'Null-pad the layer
        If flipAllLayers Then tmpLayerRef.convertToNullPaddedLayer pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height
        
        'Mirror it
        StretchBlt tmpLayerRef.layerDIB.getDIBDC, 0, 0, tmpLayerRef.layerDIB.getDIBWidth, tmpLayerRef.layerDIB.getDIBHeight, tmpLayerRef.layerDIB.getDIBDC, tmpLayerRef.layerDIB.getDIBWidth - 1, 0, -tmpLayerRef.layerDIB.getDIBWidth, tmpLayerRef.layerDIB.getDIBHeight, vbSrcCopy
        
        'Remove any null-padding
        If flipAllLayers Then tmpLayerRef.cropNullPaddedLayer
        
        'Notify the target layer that its DIB data has been changed; the layer will use this to regenerate various internal caches
        tmpLayerRef.notifyLayerModified
    
    Next i
    
    Message "Finished."
    
    'Redraw the viewport
    ScrollViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

'Rotate an image 90° clockwise
' TODO: test PlgBlt as an alternative implementation (PD currently uses GDI+, which is not the fastest kid on the block)
Public Sub MenuRotate90Clockwise(Optional ByVal targetLayerIndex As Long = -1)

    Dim flipAllLayers As Boolean
    If targetLayerIndex = -1 Then flipAllLayers = True Else flipAllLayers = False
    
    'If the image contains an active selection, disable it before transforming the canvas
    If flipAllLayers And pdImages(g_CurrentImage).selectionActive Then
        pdImages(g_CurrentImage).selectionActive = False
        pdImages(g_CurrentImage).mainSelection.lockRelease
    End If

    Message "Rotating image clockwise..."
    
    'A temporary DIB will hold the contents of the layer as it is being rotated
    Dim copyDIB As pdDIB
    Set copyDIB = New pdDIB
    
    'Lots of helper variables for a function like this
    Dim imgWidth As Long, imgHeight As Long
    imgWidth = pdImages(g_CurrentImage).Width
    imgHeight = pdImages(g_CurrentImage).Height
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    If flipAllLayers Then
        SetProgBarMax pdImages(g_CurrentImage).getNumOfLayers - 1
    Else
        SetProgBarMax targetLayerIndex
    End If
    
    'Iterate through each layer, rotating them in turn
    Dim tmpLayerRef As pdLayer
    
    Dim i As Long, lStart As Long, lEnd As Long
    
    'If the user wants us to process all layers, we will iterate through the full layer stack, applying the transformation to each in turn.
    ' Otherwise, we will only transform the specified layer.  To cut down on code duplication, we simply modify the endpoints of the loop.
    If flipAllLayers Then
        lStart = 0
        lEnd = pdImages(g_CurrentImage).getNumOfLayers - 1
    Else
        lStart = targetLayerIndex
        lEnd = targetLayerIndex
    End If
    
    'Loop through all relevant layers, transforming each as we go
    For i = lStart To lEnd
        
        'Retrieve a pointer to the layer of interest
        Set tmpLayerRef = pdImages(g_CurrentImage).getLayerByIndex(i)
        
        'Null-pad the layer
        If flipAllLayers Then tmpLayerRef.convertToNullPaddedLayer pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height
        
        'Make a copy of the layer, which we will use as our source during the transform
        copyDIB.createFromExistingDIB tmpLayerRef.layerDIB
        
        'Create a blank destination DIB to receive the transformed pixels
        tmpLayerRef.layerDIB.createBlank tmpLayerRef.layerDIB.getDIBHeight, tmpLayerRef.layerDIB.getDIBWidth, 32
        
        'Use GDI+ to apply the rotation
        GDIPlusRotateDIB tmpLayerRef.layerDIB, (tmpLayerRef.layerDIB.getDIBWidth - copyDIB.getDIBWidth) / 2, (tmpLayerRef.layerDIB.getDIBHeight - copyDIB.getDIBHeight) / 2, copyDIB.getDIBWidth, copyDIB.getDIBHeight, copyDIB, 0, 0, copyDIB.getDIBWidth, copyDIB.getDIBHeight, 90, InterpolationModeNearestNeighbor
        
        'Remove any null-padding
        If flipAllLayers Then tmpLayerRef.cropNullPaddedLayer
        
        'Notify the target layer that its DIB data has been changed; the layer will use this to regenerate various internal caches
        tmpLayerRef.notifyLayerModified
    
        'Update the progress bar (really only relevant if rotating the entire image)
        SetProgBarVal i
    
    Next i
    
    'Update the current image size, if necessary
    If flipAllLayers Then
        pdImages(g_CurrentImage).updateSize False, imgHeight, imgWidth
        DisplaySize pdImages(g_CurrentImage)
    End If
    
    Message "Finished. "
    
    PrepareViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0), "image rotated"
    
    'Reset the progress bar to zero
    SetProgBarVal 0
    releaseProgressBar
    
End Sub

'Rotate an image 180°
Public Sub MenuRotate180(Optional ByVal targetLayerIndex As Long = -1)

    Dim flipAllLayers As Boolean
    If targetLayerIndex = -1 Then flipAllLayers = True Else flipAllLayers = False
    
    'If the image contains an active selection, disable it before transforming the canvas
    If flipAllLayers And pdImages(g_CurrentImage).selectionActive Then
        pdImages(g_CurrentImage).selectionActive = False
        pdImages(g_CurrentImage).mainSelection.lockRelease
    End If

    'Fun fact: rotating 180 degrees can be accomplished by flipping and then mirroring it.
    Message "Rotating image..."
    
    'Iterate through each layer, rotating them in turn
    Dim tmpLayerRef As pdLayer
    
    Dim i As Long, lStart As Long, lEnd As Long
    
    'If the user wants us to process all layers, we will iterate through the full layer stack, applying the transformation to each in turn.
    ' Otherwise, we will only transform the specified layer.  To cut down on code duplication, we simply modify the endpoints of the loop.
    If flipAllLayers Then
        lStart = 0
        lEnd = pdImages(g_CurrentImage).getNumOfLayers - 1
    Else
        lStart = targetLayerIndex
        lEnd = targetLayerIndex
    End If
    
    'Loop through all relevant layers, transforming each as we go
    For i = lStart To lEnd
    
        'Retrieve a pointer to the layer of interest
        Set tmpLayerRef = pdImages(g_CurrentImage).getLayerByIndex(i)
        
        'Null-pad the layer
        If flipAllLayers Then tmpLayerRef.convertToNullPaddedLayer pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height
        
        'Rotate it by inverting both directions of a StretchBlt call
        StretchBlt tmpLayerRef.layerDIB.getDIBDC, 0, 0, tmpLayerRef.layerDIB.getDIBWidth, tmpLayerRef.layerDIB.getDIBHeight, tmpLayerRef.layerDIB.getDIBDC, tmpLayerRef.layerDIB.getDIBWidth - 1, tmpLayerRef.layerDIB.getDIBHeight - 1, -tmpLayerRef.layerDIB.getDIBWidth, -tmpLayerRef.layerDIB.getDIBHeight, vbSrcCopy
        
        'Remove any null-padding
        If flipAllLayers Then tmpLayerRef.cropNullPaddedLayer
        
        'Notify the target layer that its DIB data has been changed; the layer will use this to regenerate various internal caches
        tmpLayerRef.notifyLayerModified
    
    Next i
            
    Message "Finished. "
    
    ScrollViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

'Rotate an image 90° counter-clockwise
' TODO: test PlgBlt as an alternative implementation (PD currently uses GDI+, which is not the fastest kid on the block)
Public Sub MenuRotate270Clockwise(Optional ByVal targetLayerIndex As Long = -1)

    Dim flipAllLayers As Boolean
    If targetLayerIndex = -1 Then flipAllLayers = True Else flipAllLayers = False
    
    'If the image contains an active selection, disable it before transforming the canvas
    If flipAllLayers And pdImages(g_CurrentImage).selectionActive Then
        pdImages(g_CurrentImage).selectionActive = False
        pdImages(g_CurrentImage).mainSelection.lockRelease
    End If

    Message "Rotating image counter-clockwise..."
    
    'A temporary DIB will hold the contents of the layer as it is being rotated
    Dim copyDIB As pdDIB
    Set copyDIB = New pdDIB
    
    'Lots of helper variables for a function like this
    Dim imgWidth As Long, imgHeight As Long
    imgWidth = pdImages(g_CurrentImage).Width
    imgHeight = pdImages(g_CurrentImage).Height
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    If flipAllLayers Then
        SetProgBarMax pdImages(g_CurrentImage).getNumOfLayers - 1
    Else
        SetProgBarMax targetLayerIndex
    End If
    
    'Iterate through each layer, rotating them in turn
    Dim tmpLayerRef As pdLayer
    
    Dim i As Long, lStart As Long, lEnd As Long
    
    'If the user wants us to process all layers, we will iterate through the full layer stack, applying the transformation to each in turn.
    ' Otherwise, we will only transform the specified layer.  To cut down on code duplication, we simply modify the endpoints of the loop.
    If flipAllLayers Then
        lStart = 0
        lEnd = pdImages(g_CurrentImage).getNumOfLayers - 1
    Else
        lStart = targetLayerIndex
        lEnd = targetLayerIndex
    End If
    
    'Loop through all relevant layers, transforming each as we go
    For i = lStart To lEnd
        
        'Retrieve a pointer to the layer of interest
        Set tmpLayerRef = pdImages(g_CurrentImage).getLayerByIndex(i)
        
        'Null-pad the layer
        If flipAllLayers Then tmpLayerRef.convertToNullPaddedLayer pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height
        
        'Make a copy of the layer, which we will use as our source during the transform
        copyDIB.createFromExistingDIB tmpLayerRef.layerDIB
        
        'Create a blank destination DIB to receive the transformed pixels
        tmpLayerRef.layerDIB.createBlank tmpLayerRef.layerDIB.getDIBHeight, tmpLayerRef.layerDIB.getDIBWidth, 32
        
        'Use GDI+ to apply the rotation
        GDIPlusRotateDIB tmpLayerRef.layerDIB, (tmpLayerRef.layerDIB.getDIBWidth - copyDIB.getDIBWidth) / 2, (tmpLayerRef.layerDIB.getDIBHeight - copyDIB.getDIBHeight) / 2, copyDIB.getDIBWidth, copyDIB.getDIBHeight, copyDIB, 0, 0, copyDIB.getDIBWidth, copyDIB.getDIBHeight, -90, InterpolationModeNearestNeighbor
        
        'Remove any null-padding
        If flipAllLayers Then tmpLayerRef.cropNullPaddedLayer
        
        'Notify the target layer that its DIB data has been changed; the layer will use this to regenerate various internal caches
        tmpLayerRef.notifyLayerModified
        
        'Update the progress bar (really only relevant if rotating the entire image)
        SetProgBarVal i
    
    Next i
    
    'Update the current image size, if necessary
    If flipAllLayers Then
        pdImages(g_CurrentImage).updateSize False, imgHeight, imgWidth
        DisplaySize pdImages(g_CurrentImage)
    End If
    
    Message "Finished. "
    
    PrepareViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0), "image rotated"
    
    'Reset the progress bar to zero
    SetProgBarVal 0
    releaseProgressBar
    
End Sub

'This function takes an x and y value - as floating-point - and uses their position to calculate an interpolated value
' for an imaginary pixel in that location.  Offset (r/g/b/alpha) and image color depth are also required.
Public Function getInterpolatedVal(ByVal x1 As Double, ByVal y1 As Double, ByRef iData() As Byte, ByRef iOffset As Long, ByRef iDepth As Long) As Byte
        
    'Retrieve the four surrounding pixel values
    Dim topLeft As Double, topRight As Double, bottomLeft As Double, bottomRight As Double
    topLeft = iData(Int(x1) * iDepth + iOffset, Int(y1))
    topRight = iData(Int(x1 + 1) * iDepth + iOffset, Int(y1))
    bottomLeft = iData(Int(x1) * iDepth + iOffset, Int(y1 + 1))
    bottomRight = iData(Int(x1 + 1) * iDepth + iOffset, Int(y1 + 1))
    
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
    getInterpolatedVal = bottomRowColor * yBlend + topRowColor * (1 - yBlend)

End Function

'This function takes an x and y value - as floating-point - and uses their position to calculate an interpolated value
' for an imaginary pixel in that location.  Offset (r/g/b/alpha) and image color depth are also required.
Public Function getInterpolatedValWrap(ByVal x1 As Double, ByVal y1 As Double, ByVal xMax As Long, yMax As Long, ByRef iData() As Byte, ByRef iOffset As Long, ByRef iDepth As Long) As Byte
        
    'Retrieve the four surrounding pixel values
    Dim topLeft As Double, topRight As Double, bottomLeft As Double, bottomRight As Double
    topLeft = iData(Int(x1) * iDepth + iOffset, Int(y1))
    If Int(x1) = xMax Then
        topRight = iData(0 + iOffset, Int(y1))
    Else
        topRight = iData(Int(x1 + 1) * iDepth + iOffset, Int(y1))
    End If
    If Int(y1) = yMax Then
        bottomLeft = iData(Int(x1) * iDepth + iOffset, 0)
    Else
        bottomLeft = iData(Int(x1) * iDepth + iOffset, Int(y1 + 1))
    End If
    
    If Int(x1) = xMax Then
        If Int(y1) = yMax Then
            bottomRight = iData(0 + iOffset, 0)
        Else
            bottomRight = iData(0 + iOffset, Int(y1 + 1))
        End If
    Else
        If Int(y1) = yMax Then
            bottomRight = iData(Int(x1 + 1) * iDepth + iOffset, 0)
        Else
            bottomRight = iData(Int(x1 + 1) * iDepth + iOffset, Int(y1 + 1))
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
    getInterpolatedValWrap = bottomRowColor * yBlend + topRowColor * (1 - yBlend)

End Function

'Fit the image canvas around the current layer
Public Sub MenuFitCanvasToLayer(ByVal dstLayerIndex As Long)
    
    Message "Fitting image canvas around layer..."
    
    'If the image contains an active selection, disable it before transforming the canvas
    If pdImages(g_CurrentImage).selectionActive Then
        pdImages(g_CurrentImage).selectionActive = False
        pdImages(g_CurrentImage).mainSelection.lockRelease
    End If
    
    'Start by calculating a new offset, based on the current layer's offsets
    Dim dstX As Long, dstY As Long
    dstX = pdImages(g_CurrentImage).getLayerByIndex(dstLayerIndex).getLayerOffsetX
    dstY = pdImages(g_CurrentImage).getLayerByIndex(dstLayerIndex).getLayerOffsetY
    
    'Now that we have new top-left corner coordinates (and new width/height values), resizing the canvas
    ' is actually very easy.  In PhotoDemon, there is no such thing as "image data"; an image is just an
    ' imaginary bounding box around the layers collection.  Because of this, we don't actually need to
    ' resize any pixel data - we just need to modify all layer offsets to account for the new top-left corner!
    Dim i As Long
    For i = 0 To pdImages(g_CurrentImage).getNumOfLayers - 1
    
        With pdImages(g_CurrentImage).getLayerByIndex(i)
            .setLayerOffsetX .getLayerOffsetX - dstX
            .setLayerOffsetY .getLayerOffsetY - dstY
        End With
    
    Next i
    
    'Finally, update the parent image's size and DPI values
    pdImages(g_CurrentImage).updateSize False, pdImages(g_CurrentImage).getLayerByIndex(dstLayerIndex).layerDIB.getDIBWidth, pdImages(g_CurrentImage).getLayerByIndex(dstLayerIndex).layerDIB.getDIBHeight
    DisplaySize pdImages(g_CurrentImage)
    
    'In other functions, we would refresh the layer box here; however, because we haven't actually changed the
    ' appearance of any of the layers, we can leave it as-is!
    
    'Fit the new image on-screen and redraw its viewport
    PrepareViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0), "Fit canvas to layer"
    
    Message "Finished."
    
End Sub

'Fit the canvas around all layers present in the image
Public Sub MenuFitCanvasToAllLayers()
    
    Message "Fitting image canvas around layer..."
    
    'If the image contains an active selection, disable it before transforming the canvas
    If pdImages(g_CurrentImage).selectionActive Then
        pdImages(g_CurrentImage).selectionActive = False
        pdImages(g_CurrentImage).mainSelection.lockRelease
    End If
    
    'Start by finding two things:
    ' 1) The lowest x/y offsets in the current layer stack
    ' 2) The highest width/height in the current layer stack (while accounting for offsets as well!)
    Dim dstLeft As Long, dstTop As Long, dstRight As Long, dstBottom As Long
    dstLeft = &HFFFFFF
    dstTop = &HFFFFFF
    dstRight = -1 * &HFFFFFF
    dstBottom = -1 * &HFFFFFF
    
    Dim i As Long
    
    For i = 0 To pdImages(g_CurrentImage).getNumOfLayers - 1
    
        With pdImages(g_CurrentImage).getLayerByIndex(i)
        
            'Check for new minimum offsets
            If .getLayerOffsetX < dstLeft Then dstLeft = .getLayerOffsetX
            If .getLayerOffsetY < dstTop Then dstTop = .getLayerOffsetY
            
            'Check for new maximum right/top
            If .getLayerOffsetX + .layerDIB.getDIBWidth > dstRight Then dstRight = .getLayerOffsetX + .layerDIB.getDIBWidth
            If .getLayerOffsetY + .layerDIB.getDIBHeight > dstBottom Then dstBottom = .getLayerOffsetY + .layerDIB.getDIBHeight
        
        End With
    
    Next i
    
    'Now that we have new top-left corner coordinates (and new width/height values), resizing the canvas
    ' is actually very easy.  In PhotoDemon, there is no such thing as "image data"; an image is just an
    ' imaginary bounding box around the layers collection.  Because of this, we don't actually need to
    ' resize any pixel data - we just need to modify all layer offsets to account for the new top-left corner!
    For i = 0 To pdImages(g_CurrentImage).getNumOfLayers - 1
    
        With pdImages(g_CurrentImage).getLayerByIndex(i)
            .setLayerOffsetX .getLayerOffsetX - dstLeft
            .setLayerOffsetY .getLayerOffsetY - dstTop
        End With
    
    Next i
    
    'Finally, update the parent image's size
    pdImages(g_CurrentImage).updateSize False, (dstRight - dstLeft), (dstBottom - dstTop)
    DisplaySize pdImages(g_CurrentImage)
    
    'In other functions, we would refresh the layer box here; however, because we haven't actually changed the
    ' appearance of any of the layers, we can leave it as-is!
    
    'Fit the new image on-screen and redraw its viewport
    PrepareViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0), "Fit canvas to all layers"
    
    Message "Finished."
    
End Sub

'Automatically trim empty borders from an image.  Empty borders are defined as borders comprised only of 100% transparent pixels.
Public Sub TrimImage()

    'If the image contains an active selection, disable it before transforming the canvas
    If pdImages(g_CurrentImage).selectionActive Then
        pdImages(g_CurrentImage).selectionActive = False
        pdImages(g_CurrentImage).mainSelection.lockRelease
    End If

    'The image will be trimmed in four steps.  Each edge will be trimmed separately, starting with the top.
    
    Message "Analyzing top edge of image..."
    
    'Retrieve a copy of the composited image
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    pdImages(g_CurrentImage).getCompositedImage tmpDIB
    
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
    Dim QuickVal As Long
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    SetProgBarMax 4
    
    'The new edges of the image will mark these values for us
    Dim newTop As Long, newBottom As Long, newLeft As Long, newRight As Long
    
    'When a non-transparent pixel is found, this check value will be set to TRUE
    Dim colorFails As Boolean
    colorFails = False
    
    'Scan the image, starting at the top-left and moving right
    For y = 0 To finalY
    For x = 0 To finalX
        
        'If this pixel is transparent, keep scanning.  Otherwise, note that we have found a non-transparent pixel
        ' and exit the loop.
        If srcImageData(x * 4 + 3, y) > 0 Then colorFails = True
        If colorFails Then Exit For
        
    Next x
        If colorFails Then Exit For
    Next y
    
    'We have now reached one of two conditions:
    '1) The entire image is transparent
    '2) The loop progressed part-way through the image and terminated
    
    'Check for case (1) and warn the user if it occurred
    If Not colorFails Then
    
        CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
        Erase srcImageData
        
        SetProgBarVal 0
        releaseProgressBar
        
        Message "Image is fully transparent.  Trim abandoned."
        Exit Sub
    
    'Next, check for case (2)
    Else
        newTop = y
    End If
    
    initY = newTop
    
    'Repeat the above steps, but tracking the left edge instead.  Note also that we will only be scanning from wherever
    ' the top trim failed - this saves processing time.
    colorFails = False
    
    Message "Analyzing left edge of image..."
    SetProgBarVal 1
    
    For x = 0 To finalX
        QuickVal = x * 4
    For y = initY To finalY
    
        If srcImageData(QuickVal + 3, y) > 0 Then colorFails = True
        If colorFails Then Exit For
        
    Next y
        If colorFails Then Exit For
    Next x
    
    newLeft = x
    
    'Repeat the above steps, but tracking the right edge instead.  Note also that we will only be scanning from wherever
    ' the top trim failed - this saves processing time.
    colorFails = False
    
    Message "Analyzing right edge of image..."
    SetProgBarVal 2
    
    For x = finalX To 0 Step -1
        QuickVal = x * 4
    For y = initY To finalY
    
        If srcImageData(QuickVal + 3, y) > 0 Then colorFails = True
        If colorFails Then Exit For
        
    Next y
        If colorFails Then Exit For
    Next x
    
    newRight = x
    
    'Finally, repeat the steps above for the bottom of the image.  Note also that we will only be scanning from wherever
    ' the left and right trims failed - this saves processing time.
    colorFails = False
    initX = newLeft
    finalX = newRight
    
    Message "Analyzing bottom edge of image..."
    SetProgBarVal 3
    
    For y = finalY To initY Step -1
    For x = initX To finalX
        
        If srcImageData(x * 4 + 3, y) > 0 Then colorFails = True
        If colorFails Then Exit For
        
    Next x
        If colorFails Then Exit For
    Next y
    
    newBottom = y
    
    'With our work complete, point ImageData() away from the DIB and deallocate it
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    
    'Erase the temporary DIB
    Set tmpDIB = Nothing
    
    'We now know where to crop the image.  Apply the crop.
    If (newTop = 0) And (newBottom = pdImages(g_CurrentImage).Height - 1) And (newLeft = 0) And (newRight = pdImages(g_CurrentImage).Width - 1) Then
        SetProgBarVal 0
        releaseProgressBar
        Message "Image is already trimmed.  (No changes were made to the image.)"
    Else
    
        Message "Trimming image to new dimensions..."
        SetProgBarVal 4
        
        'Now that we have new top-left corner coordinates (and new width/height values), resizing the canvas
        ' is actually very easy.  In PhotoDemon, there is no such thing as "image data"; an image is just an
        ' imaginary bounding box around the layers collection.  Because of this, we don't actually need to
        ' resize any pixel data - we just need to modify all layer offsets to account for the new top-left corner!
        Dim i As Long
        For i = 0 To pdImages(g_CurrentImage).getNumOfLayers - 1
        
            With pdImages(g_CurrentImage).getLayerByIndex(i)
                .setLayerOffsetX .getLayerOffsetX - newLeft
                .setLayerOffsetY .getLayerOffsetY - newTop
            End With
        
        Next i
    
        'Finally, update the parent image's size
        pdImages(g_CurrentImage).updateSize False, (newRight - newLeft), (newBottom - newTop)
        DisplaySize pdImages(g_CurrentImage)
    
        'In other functions, we would refresh the layer box here; however, because we haven't actually changed the
        ' appearance of any of the layers, we can leave it as-is!
        
        Message "Finished. "
        SetProgBarVal 0
        releaseProgressBar
        
        'Redraw the image
        PrepareViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0), "Trim empty borders"
    
    End If

End Sub

