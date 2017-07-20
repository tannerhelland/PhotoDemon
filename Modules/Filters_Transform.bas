Attribute VB_Name = "Filters_Transform"
'***************************************************************************
'Image Transformations Interface (including flip/mirror/rotation/crop/etc)
'Copyright 2003-2017 by Tanner Helland
'Created: 25/January/03
'Last updated: 13/June/17
'Last update: routine code-cleanup, minor optimizations
'
'Functions for generic 2D transformations, including rotate, flip, mirror and crop.
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
    If pdImages(g_CurrentImage).IsSelectionActive Then
        pdImages(g_CurrentImage).SetSelectionActive False
        pdImages(g_CurrentImage).MainSelection.LockRelease
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
    PrepSafeArray srcSA, tmpDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    finalX = pdImages(g_CurrentImage).Width - 1
    finalY = pdImages(g_CurrentImage).Height - 1
            
    'These values will help us access locations in the array more quickly.
    ' (qvDepth is required because the image array may be 24 or 32 bits per pixel, and we want to handle both cases.)
    Dim quickVal As Long, qvDepth As Long
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
        quickVal = x * qvDepth
        curColor = gLookUp(CLng(srcImageData(quickVal, y)) + CLng(srcImageData(quickVal + 1, y)) + CLng(srcImageData(quickVal + 2, y)))
        
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
        ReleaseProgressBar
        
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
        quickVal = x * qvDepth
    For y = initY To finalY
    
        curColor = gLookUp(CLng(srcImageData(quickVal, y)) + CLng(srcImageData(quickVal + 1, y)) + CLng(srcImageData(quickVal + 2, y)))
        
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
    quickVal = finalX * qvDepth
    initColor = gLookUp(CLng(srcImageData(quickVal, initY)) + CLng(srcImageData(quickVal + 1, 0)) + CLng(srcImageData(quickVal + 2, 0)))
    SetProgBarVal 2
    
    For x = finalX To 0 Step -1
        quickVal = x * qvDepth
    For y = initY To finalY
    
        curColor = gLookUp(CLng(srcImageData(quickVal, y)) + CLng(srcImageData(quickVal + 1, y)) + CLng(srcImageData(quickVal + 2, y)))
        
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
    quickVal = initX * qvDepth
    initColor = gLookUp(CLng(srcImageData(quickVal, finalY)) + CLng(srcImageData(quickVal + 1, finalY)) + CLng(srcImageData(quickVal + 2, finalY)))
    
    Message "Analyzing bottom edge of image..."
    SetProgBarVal 3
    
    For y = finalY To initY Step -1
    For x = initX To finalX
        quickVal = x * qvDepth
        curColor = gLookUp(CLng(srcImageData(quickVal, y)) + CLng(srcImageData(quickVal + 1, y)) + CLng(srcImageData(quickVal + 2, y)))
        
        'If pixel color DOES NOT match the baseline, keep scanning.  Otherwise, note that we have found a mismatched color
        ' and exit the loop.
        If Abs(curColor - initColor) > cThreshold Then colorFails = True
        
        If colorFails Then Exit For
        
    Next x
        If colorFails Then Exit For
    Next y
    
    newBottom = y
    
    'Safely deallocate imageData()
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    
    'We now know where to crop the image.  Apply the crop.
    
    If (newTop = 0) And (newBottom = pdImages(g_CurrentImage).Height - 1) And (newLeft = 0) And (newRight = pdImages(g_CurrentImage).Width - 1) Then
        SetProgBarVal 0
        ReleaseProgressBar
        Message "Image is already cropped intelligently.  Autocrop abandoned.  (No changes were made to the image.)"
    Else
    
        Message "Cropping image to new dimensions..."
        SetProgBarVal 4
        
        'Resize the current image's main DIB
        'pdImages(g_CurrentImage).mainDIB.createBlank newRight - newLeft, newBottom - newTop, tmpDIB.getDIBColorDepth
        
        'Copy the autocropped area to the new main DIB
        'BitBlt pdImages(g_CurrentImage).mainDIB.getDIBDC, 0, 0, pdImages(g_CurrentImage).mainDIB.getDIBWidth, pdImages(g_CurrentImage).mainDIB.getDIBHeight, tmpDIB.getDIBDC, newLeft, newTop, vbSrcCopy
    
        'Erase the temporary DIB
        tmpDIB.EraseDIB
        Set tmpDIB = Nothing
    
        'Update the current image size
        pdImages(g_CurrentImage).UpdateSize
        Interface.DisplaySize pdImages(g_CurrentImage)
        
        Message "Finished. "
        SetProgBarVal 0
        ReleaseProgressBar
        
        'Redraw the image
        ViewportEngine.Stage1_InitializeBuffer pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
    End If

End Sub

'Determine if a non-destructive crop is possible.  Pure rectangular selections allow this, because we can simply modify canvas
' boundaries and layer offsets to arrive at the crop shape.
Public Sub SeeIfCropCanBeAppliedNonDestructively()
    
    'First, make sure there is an active selection
    If (Not pdImages(g_CurrentImage).IsSelectionActive) Then
        Message "No active selection found.  Crop abandoned."
        
    Else
        
        'Query the active selection object; if it's a pure rectangular region, we can apply a non-destructive crop (which is not
        ' only much faster, but it doesn't require rasterizing vector layers!)
        With pdImages(g_CurrentImage).MainSelection
            
            'Start by seeing if we're even working with a rectangle.  If we are, we can check a few extra criteria as well; if we aren't,
            ' only a destructive crop is possible.
            Dim selectionIsPureRectangle As Boolean
            selectionIsPureRectangle = (.GetSelectionShape = ss_Rectangle)
            
            If selectionIsPureRectangle Then
                selectionIsPureRectangle = selectionIsPureRectangle And (.GetSelectionProperty_Long(sp_RoundedCornerRadius) = 0)
                selectionIsPureRectangle = selectionIsPureRectangle And (.GetSelectionProperty_Long(sp_Area) = sa_Interior)
                selectionIsPureRectangle = selectionIsPureRectangle And (.GetSelectionProperty_Long(sp_Smoothing) = ss_None)
            End If
            
            'If that huge list of above criteria are met, we can apply a non-destructive crop operation.
            Dim cParams As pdParamXML
            Set cParams = New pdParamXML
            cParams.AddParam "nondestructive", selectionIsPureRectangle
            Processor.Process "Crop", False, cParams.GetParamString(), UNDO_EVERYTHING
            
        End With
    
    End If
    
End Sub

'XML-based wrapper for CropToSelection, below
Public Sub CropToSelection_XML(ByVal processParameters As String)
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString processParameters
    Filters_Transform.CropToSelection , cParams.GetBool("nondestructive", False)
End Sub

'Crop the image to the current selection.  To crop only a single layer, specify a target layer index.
' (Optionally, full-image crops can be applied non-destructively, by simply modifying layer offsets and
'  image dimensions.  Single layers cannot be modified non-destructively, unfortunately.)
Public Sub CropToSelection(Optional ByVal targetLayerIndex As Long = -1, Optional ByVal applyNonDestructively As Boolean = False)
    
    'First, make sure there is an active selection
    If (Not pdImages(g_CurrentImage).IsSelectionActive) Then
        Message "No active selection found.  Crop abandoned."
        Exit Sub
    End If
    
    Message "Cropping image to selected area..."
    
    Dim progBarCheck As Long, progBarOffsetX As Long
    Dim tmpLayerRef As pdLayer
    Dim i As Long
    
    Dim selectionWidth As Long, selectionHeight As Long, selBounds As RECTF
    selBounds = pdImages(g_CurrentImage).MainSelection.GetBoundaryRect
    selectionWidth = selBounds.Width
    selectionHeight = selBounds.Height
    
    'Crop can be applied in two ways.
    ' - If the current selection is a pure rectangle with no feathering or rounded corners, and it's a full-image crop,
    '   we can crop the image non-destructively.  (This simply modifies layer offsets and canvas size, and it doesn't
    '   require rasterization of vector layers.)
    ' - If the current selection is any other shape, or if only a single layer is being cropped, we have to rasterize
    '   vector layers and apply per-pixel crops against the current mask.
    
    'This function doesn't actually determine whether a crop can be handled non-destructively; that is up to the
    ' SeeIfCropCanBeAppliedNonDestructively() function, above.
    If applyNonDestructively And (targetLayerIndex = -1) Then
    
        SetProgBarMax pdImages(g_CurrentImage).GetNumOfLayers
        
        'Non-destructive crops are very easy to handle.  In PhotoDemon, there is no such thing as "image data"; an image is just an
        ' imaginary bounding box around the layer collection.  Because of this, we don't actually need to resize any pixel data -
        ' we just need to modify all layer offsets to account for a new top-left corner!
        For i = 0 To pdImages(g_CurrentImage).GetNumOfLayers - 1
        
            SetProgBarVal i
            
            With pdImages(g_CurrentImage).GetLayerByIndex(i)
                .SetLayerOffsetX .GetLayerOffsetX - selBounds.Left
                .SetLayerOffsetY .GetLayerOffsetY - selBounds.Top
            End With
        
        Next i
        
        'That's all there is to it!
    
    'A complex shape is in use, or only a single layer is being cropped.  Crop using per-pixel raster mask analysis.
    Else
    
        'NOTE: historically, the entire rectangular bounding region of the selection was included in the crop.  (This is GIMP's behavior.)
        ' I now fully crop the image, which means that for non-square selections, all unselected pixels are set to transparent.  For non-square
        ' selections, this will always result in an image with some transparent regions.
        
        'Images will be processed into a temporary DIB
        Dim tmpDIB As pdDIB
        Set tmpDIB = New pdDIB
        
        'Arrays will be pointed at three sets of pixels: the current layer, the selection mask, and a destination layer.
        Dim srcImageData() As Byte, srcSA As SAFEARRAY2D
        Dim dstImageData() As Byte, dstSA As SAFEARRAY2D
        
        'Point our selection array at the selection mask in advance; this only needs to be done once, as the same mask is used for all layers.
        Dim selData() As Byte
        Dim selSA As SAFEARRAY2D
        pdImages(g_CurrentImage).MainSelection.GetMaskDIB.WrapArrayAroundDIB selData, selSA
        
        'Lots of helper variables for a function like this
        Dim leftOffset As Long, topOffset As Long
        leftOffset = selBounds.Left
        topOffset = selBounds.Top
        
        Dim r As Long, g As Long, b As Long
        Dim thisAlpha As Long, origAlpha As Long, blendAlpha As Double
        Dim srcQuickX As Long, srcQuickY As Long, dstQuickX As Long, selQuickX As Long
        Const ONE_DIVIDED_BY_255 As Double = 1# / 255#
        
        Dim x As Long, y As Long
        Dim imgWidth As Long, imgHeight As Long
        imgWidth = pdImages(g_CurrentImage).Width
        imgHeight = pdImages(g_CurrentImage).Height
        
        'Figure out loop boundaries.  If the entire image is being cropped, we'll need to process each layer in turn.
        Dim numLayersToCrop As Long, startLayerIndex As Long, endLayerIndex As Long
        If (targetLayerIndex = -1) Then
            numLayersToCrop = pdImages(g_CurrentImage).GetNumOfLayers
            startLayerIndex = 0
            endLayerIndex = pdImages(g_CurrentImage).GetNumOfLayers - 1
        Else
            numLayersToCrop = 1
            startLayerIndex = targetLayerIndex
            endLayerIndex = targetLayerIndex
        End If
        
        'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
        ' based on the size of the area to be processed.
        SetProgBarMax numLayersToCrop * imgWidth
        progBarCheck = FindBestProgBarValue()
        
        'Iterate through each layer, cropping them in turn
        For i = startLayerIndex To endLayerIndex
        
            'Update the progress bar counter for this layer
            progBarOffsetX = i * imgWidth
        
            'Retrieve a pointer to the layer of interest
            Set tmpLayerRef = pdImages(g_CurrentImage).GetLayerByIndex(i)
            
            'Null-pad the layer
            tmpLayerRef.ConvertToNullPaddedLayer pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height
            
            'Create a temporary layer at the relevant size of the selection, and retrieve a pointer to its pixel data
            If (tmpDIB.GetDIBWidth <> selectionWidth) Or (tmpDIB.GetDIBHeight <> selectionHeight) Then
                tmpDIB.CreateBlank selectionWidth, selectionHeight, 32, 0, 0
            Else
                tmpDIB.ResetDIB 0
            End If
            
            tmpDIB.WrapArrayAroundDIB dstImageData, dstSA
            tmpLayerRef.layerDIB.WrapArrayAroundDIB srcImageData, srcSA
            
            Dim selMaskDepth As Long
            selMaskDepth = (pdImages(g_CurrentImage).MainSelection.GetMaskDIB.GetDIBColorDepth \ 8)
            
            'Iterate through all relevant pixels in this layer (e.g. only those that actually lie within the interesting region
            ' of the selection), copying them to the destination as necessary.
            For x = 0 To selectionWidth - 1
                dstQuickX = x * 4
                srcQuickX = (leftOffset + x) * 4
                selQuickX = (leftOffset + x) * selMaskDepth
            For y = 0 To selectionHeight - 1
            
                srcQuickY = topOffset + y
                thisAlpha = selData(selQuickX, srcQuickY)
                
                If (thisAlpha > 0) Then
                
                    'Check the image's alpha value.  If it's zero, we have no reason to process it further
                    origAlpha = srcImageData(srcQuickX + 3, srcQuickY)
                    
                    If (origAlpha > 0) Then
                        
                        'Source pixel data will be premultiplied, which saves us a bunch of processing time.  (That is why
                        ' we premultiply alpha, after all!)
                        b = srcImageData(srcQuickX, srcQuickY)
                        g = srcImageData(srcQuickX + 1, srcQuickY)
                        r = srcImageData(srcQuickX + 2, srcQuickY)
                        
                        'Calculate a new multiplier, based on the strength of the selection at this location
                        blendAlpha = thisAlpha * ONE_DIVIDED_BY_255
                        
                        'Apply the multiplier to the existing pixel data (which is already premultiplied, saving us a bunch of time now)
                        dstImageData(dstQuickX, y) = b * blendAlpha
                        dstImageData(dstQuickX + 1, y) = g * blendAlpha
                        dstImageData(dstQuickX + 2, y) = r * blendAlpha
                        
                        'Finish our work by calculating a new alpha channel value for this pixel, which is a blend of
                        ' the original alpha value, and the selection mask value at this location.
                        dstImageData(dstQuickX + 3, y) = origAlpha * blendAlpha
                        
                    End If
                    
                End If
                
            Next y
                If ((progBarOffsetX + x) And progBarCheck) = 0 Then SetProgBarVal (progBarOffsetX + x)
            Next x
            
            tmpDIB.UnwrapArrayFromDIB dstImageData
            tmpLayerRef.layerDIB.UnwrapArrayFromDIB srcImageData
            
            'Replace the current layer DIB with our destination one.
            tmpDIB.SetInitialAlphaPremultiplicationState True
            tmpLayerRef.layerDIB.CreateFromExistingDIB tmpDIB
            
            'Update the layer offsets, if any.  Note that the exact approach to this varies by crop type; for single-layer crops,
            ' we need to manually update the offsets (as they aren't guaranteed to be at (0, 0), like they are for a
            ' full-image crop).
            If (targetLayerIndex = -1) Then
            
                'Remove any null-padding from the layer
                tmpLayerRef.CropNullPaddedLayer
                
            Else
                
                'Manually update layer offsets to point at the selection's top-left point
                tmpLayerRef.SetLayerOffsetX selBounds.Left
                tmpLayerRef.SetLayerOffsetY selBounds.Top
                
            End If
            
            'Notify the parent of the change
            pdImages(g_CurrentImage).NotifyImageChanged UNDO_LAYER, i
            
        Next i
        
        'Clear the selection mask array reference
        pdImages(g_CurrentImage).MainSelection.GetMaskDIB.UnwrapArrayFromDIB selData
        
    End If
        
    'From here, we do some generic clean-up that's identical for both destructive and non-destructive modes.
    ' (But generally speaking, only relevant when all layers are being cropped.)
    If (targetLayerIndex = -1) Then
    
        'The selection is now going to be out of sync with the image.  Forcibly clear it.
        pdImages(g_CurrentImage).MainSelection.LockRelease
        pdImages(g_CurrentImage).SetSelectionActive False
        pdImages(g_CurrentImage).MainSelection.EraseCustomTrackers
        SyncTextToCurrentSelection g_CurrentImage
        
    End If
    
    'Update the viewport.  For full-image crops, we need to refresh the entire viewport pipeline (as the image size
    ' may have changed).
    If (targetLayerIndex = -1) Then
        pdImages(g_CurrentImage).UpdateSize False, selectionWidth, selectionHeight
        Interface.DisplaySize pdImages(g_CurrentImage)
        ViewportEngine.Stage1_InitializeBuffer pdImages(g_CurrentImage), FormMain.mainCanvas(0)
        CanvasManager.CenterOnScreen
    
    'For individual layers, we can use some existing viewport pipeline data
    Else
        ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
    
    'Reset the progress bar to zero, then exit
    SetProgBarVal 0
    ReleaseProgressBar
    
End Sub

'Flip an image vertically.  If no layer is specified (e.g. if targetLayerIndex = -1), all layers will be flipped.
Public Sub MenuFlip(Optional ByVal targetLayerIndex As Long = -1)

    Dim flipAllLayers As Boolean
    flipAllLayers = (targetLayerIndex = -1)

    'If the image contains an active selection, disable it before transforming the canvas
    If flipAllLayers And pdImages(g_CurrentImage).IsSelectionActive Then
        pdImages(g_CurrentImage).SetSelectionActive False
        pdImages(g_CurrentImage).MainSelection.LockRelease
    End If
    
    Message "Flipping image..."
    
    'Iterate through each layer, flipping them in turn
    Dim tmpLayerRef As pdLayer
    
    Dim i As Long, lStart As Long, lEnd As Long
    
    'If the user wants us to process all layers, we will iterate through the full layer stack, applying the transformation to each in turn.
    ' Otherwise, we will only transform the specified layer.  To cut down on code duplication, we simply modify the endpoints of the loop.
    If flipAllLayers Then
        lStart = 0
        lEnd = pdImages(g_CurrentImage).GetNumOfLayers - 1
    Else
        lStart = targetLayerIndex
        lEnd = targetLayerIndex
    End If
    
    'Loop through all relevant layers, transforming each as we go
    For i = lStart To lEnd
    
        'Retrieve a pointer to the layer of interest
        Set tmpLayerRef = pdImages(g_CurrentImage).GetLayerByIndex(i)
        
        'Null-pad the layer
        If flipAllLayers Then tmpLayerRef.ConvertToNullPaddedLayer pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height
        
        'Flip it
        StretchBlt tmpLayerRef.layerDIB.GetDIBDC, 0, 0, tmpLayerRef.GetLayerWidth(False), tmpLayerRef.GetLayerHeight(False), tmpLayerRef.layerDIB.GetDIBDC, 0, tmpLayerRef.GetLayerHeight(False) - 1, tmpLayerRef.GetLayerWidth(False), -tmpLayerRef.GetLayerHeight(False), vbSrcCopy
        
        'Remove any null-padding
        If flipAllLayers Then tmpLayerRef.CropNullPaddedLayer
        
        'Notify the parent image of the change
        pdImages(g_CurrentImage).NotifyImageChanged UNDO_LAYER, i
        
    Next i
    
    'Notify the parent image that the entire image now needs to be recomposited
    pdImages(g_CurrentImage).NotifyImageChanged UNDO_IMAGE
    
    Message "Finished. "
    
    'Redraw the viewport
    ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

'Flip an image horizontally (mirror)
Public Sub MenuMirror(Optional ByVal targetLayerIndex As Long = -1)

    Dim flipAllLayers As Boolean
    flipAllLayers = (targetLayerIndex = -1)
    
    'If the image contains an active selection, disable it before transforming the canvas
    If flipAllLayers And pdImages(g_CurrentImage).IsSelectionActive Then
        pdImages(g_CurrentImage).SetSelectionActive False
        pdImages(g_CurrentImage).MainSelection.LockRelease
    End If

    Message "Mirroring image..."
    
    'Iterate through each layer, mirroring them in turn
    Dim tmpLayerRef As pdLayer
    
    Dim i As Long, lStart As Long, lEnd As Long
    
    'If the user wants us to process all layers, we will iterate through the full layer stack, applying the transformation to each in turn.
    ' Otherwise, we will only transform the specified layer.  To cut down on code duplication, we simply modify the endpoints of the loop.
    If flipAllLayers Then
        lStart = 0
        lEnd = pdImages(g_CurrentImage).GetNumOfLayers - 1
    Else
        lStart = targetLayerIndex
        lEnd = targetLayerIndex
    End If
    
    'Loop through all relevant layers, transforming each as we go
    For i = lStart To lEnd
    
        'Retrieve a pointer to the layer of interest
        Set tmpLayerRef = pdImages(g_CurrentImage).GetLayerByIndex(i)
        
        'Null-pad the layer
        If flipAllLayers Then tmpLayerRef.ConvertToNullPaddedLayer pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height
        
        'Mirror it
        StretchBlt tmpLayerRef.layerDIB.GetDIBDC, 0, 0, tmpLayerRef.GetLayerWidth(False), tmpLayerRef.GetLayerHeight(False), tmpLayerRef.layerDIB.GetDIBDC, tmpLayerRef.GetLayerWidth(False) - 1, 0, -tmpLayerRef.GetLayerWidth(False), tmpLayerRef.GetLayerHeight(False), vbSrcCopy
        
        'Remove any null-padding
        If flipAllLayers Then tmpLayerRef.CropNullPaddedLayer
        
        'Notify the parent image of the change
        pdImages(g_CurrentImage).NotifyImageChanged UNDO_LAYER, i
        
    Next i
    
    'Notify the parent image that the entire image now needs to be recomposited
    pdImages(g_CurrentImage).NotifyImageChanged UNDO_IMAGE
    
    Message "Finished."
    
    'Redraw the viewport
    ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

'Rotate an image 90 degrees clockwise
Public Sub MenuRotate90Clockwise(Optional ByVal targetLayerIndex As Long = -1)

    Dim flipAllLayers As Boolean
    flipAllLayers = (targetLayerIndex = -1)
    
    'If the image contains an active selection, disable it before transforming the canvas
    If flipAllLayers And pdImages(g_CurrentImage).IsSelectionActive Then
        pdImages(g_CurrentImage).SetSelectionActive False
        pdImages(g_CurrentImage).MainSelection.LockRelease
    End If

    Message "Rotating image clockwise..."
    
    Dim copyDIB As pdDIB
    Set copyDIB = New pdDIB
    
    Dim imgWidth As Long, imgHeight As Long
    imgWidth = pdImages(g_CurrentImage).Width
    imgHeight = pdImages(g_CurrentImage).Height
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    If flipAllLayers Then
        SetProgBarMax pdImages(g_CurrentImage).GetNumOfLayers - 1
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
        lEnd = pdImages(g_CurrentImage).GetNumOfLayers - 1
    Else
        lStart = targetLayerIndex
        lEnd = targetLayerIndex
    End If
    
    'Loop through all relevant layers, transforming each as we go
    For i = lStart To lEnd
        
        'Retrieve a pointer to the layer of interest
        Set tmpLayerRef = pdImages(g_CurrentImage).GetLayerByIndex(i)
        
        'Null-pad the layer
        If flipAllLayers Then tmpLayerRef.ConvertToNullPaddedLayer pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height
        
        'Make a copy of the layer, which we will use as our source during the transform
        copyDIB.CreateFromExistingDIB tmpLayerRef.layerDIB
        
        'Create a blank destination DIB to receive the transformed pixels
        tmpLayerRef.layerDIB.CreateBlank imgHeight, imgWidth, 32
        
        'Use GDI+ to apply the rotation
        GDI_Plus.GDIPlusRotateFlipDIB copyDIB, tmpLayerRef.layerDIB, GP_RF_90FlipNone
        
        'Mark the correct alpha state and remove any null-padding
        tmpLayerRef.layerDIB.SetInitialAlphaPremultiplicationState True
        If flipAllLayers Then tmpLayerRef.CropNullPaddedLayer
        
        'Notify the parent of the change
        pdImages(g_CurrentImage).NotifyImageChanged UNDO_LAYER, i
    
        'Update the progress bar (really only relevant if rotating the entire image)
        SetProgBarVal i
    
    Next i
    
    'Update the current image size, if necessary
    If flipAllLayers Then
        pdImages(g_CurrentImage).UpdateSize False, imgHeight, imgWidth
        Interface.DisplaySize pdImages(g_CurrentImage)
    End If
    
    Message "Finished. "
    
    ViewportEngine.Stage1_InitializeBuffer pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
    'Reset the progress bar to zero
    SetProgBarVal 0
    ReleaseProgressBar
    
End Sub

'Rotate an image 180 degrees
Public Sub MenuRotate180(Optional ByVal targetLayerIndex As Long = -1)

    Dim flipAllLayers As Boolean
    flipAllLayers = (targetLayerIndex = -1)
    
    'If the image contains an active selection, disable it before transforming the canvas
    If flipAllLayers And pdImages(g_CurrentImage).IsSelectionActive Then
        pdImages(g_CurrentImage).SetSelectionActive False
        pdImages(g_CurrentImage).MainSelection.LockRelease
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
        lEnd = pdImages(g_CurrentImage).GetNumOfLayers - 1
    Else
        lStart = targetLayerIndex
        lEnd = targetLayerIndex
    End If
    
    'Loop through all relevant layers, transforming each as we go
    For i = lStart To lEnd
    
        'Retrieve a pointer to the layer of interest
        Set tmpLayerRef = pdImages(g_CurrentImage).GetLayerByIndex(i)
        
        'Null-pad the layer
        If flipAllLayers Then tmpLayerRef.ConvertToNullPaddedLayer pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height
        
        'Rotate it by inverting both directions of a StretchBlt call
        StretchBlt tmpLayerRef.layerDIB.GetDIBDC, 0, 0, tmpLayerRef.GetLayerWidth(False), tmpLayerRef.GetLayerHeight(False), tmpLayerRef.layerDIB.GetDIBDC, tmpLayerRef.GetLayerWidth(False) - 1, tmpLayerRef.GetLayerHeight(False) - 1, -tmpLayerRef.GetLayerWidth(False), -tmpLayerRef.GetLayerHeight(False), vbSrcCopy
        
        'Remove any null-padding
        If flipAllLayers Then tmpLayerRef.CropNullPaddedLayer
        
        'Notify the parent image of the change
        pdImages(g_CurrentImage).NotifyImageChanged UNDO_LAYER, i
        
    Next i
    
    'Notify the parent image that the entire image now needs to be recomposited
    pdImages(g_CurrentImage).NotifyImageChanged UNDO_IMAGE
            
    Message "Finished. "
    
    ViewportEngine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

'Rotate an image 90 degrees counter-clockwise
Public Sub MenuRotate270Clockwise(Optional ByVal targetLayerIndex As Long = -1)

    Dim flipAllLayers As Boolean
    flipAllLayers = (targetLayerIndex = -1)
    
    'If the image contains an active selection, disable it before transforming the canvas
    If flipAllLayers And pdImages(g_CurrentImage).IsSelectionActive Then
        pdImages(g_CurrentImage).SetSelectionActive False
        pdImages(g_CurrentImage).MainSelection.LockRelease
    End If

    Message "Rotating image counter-clockwise..."
    
    Dim copyDIB As pdDIB
    Set copyDIB = New pdDIB
    
    Dim imgWidth As Long, imgHeight As Long
    imgWidth = pdImages(g_CurrentImage).Width
    imgHeight = pdImages(g_CurrentImage).Height
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    If flipAllLayers Then
        SetProgBarMax pdImages(g_CurrentImage).GetNumOfLayers - 1
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
        lEnd = pdImages(g_CurrentImage).GetNumOfLayers - 1
    Else
        lStart = targetLayerIndex
        lEnd = targetLayerIndex
    End If
    
    'Loop through all relevant layers, transforming each as we go
    For i = lStart To lEnd
        
        'Retrieve a pointer to the layer of interest
        Set tmpLayerRef = pdImages(g_CurrentImage).GetLayerByIndex(i)
        
        'Null-pad the layer
        If flipAllLayers Then tmpLayerRef.ConvertToNullPaddedLayer pdImages(g_CurrentImage).Width, pdImages(g_CurrentImage).Height
        
        'Make a copy of the layer, which we will use as our source during the transform
        copyDIB.CreateFromExistingDIB tmpLayerRef.layerDIB
        
        'Create a blank destination DIB to receive the transformed pixels
        tmpLayerRef.layerDIB.CreateBlank imgHeight, imgWidth, 32
        
        'Use GDI+ to apply the rotation
        GDI_Plus.GDIPlusRotateFlipDIB copyDIB, tmpLayerRef.layerDIB, GP_RF_270FlipNone
        
        'Mark the correct alpha state and remove any null-padding
        tmpLayerRef.layerDIB.SetInitialAlphaPremultiplicationState True
        If flipAllLayers Then tmpLayerRef.CropNullPaddedLayer
        
        'Notify the parent of the change
        pdImages(g_CurrentImage).NotifyImageChanged UNDO_LAYER, i
    
        'Update the progress bar (really only relevant if rotating the entire image)
        SetProgBarVal i
    
    Next i
    
    'Update the current image size, if necessary
    If flipAllLayers Then
        pdImages(g_CurrentImage).UpdateSize False, imgHeight, imgWidth
        Interface.DisplaySize pdImages(g_CurrentImage)
    End If
    
    Message "Finished. "
    
    ViewportEngine.Stage1_InitializeBuffer pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
    'Reset the progress bar to zero
    SetProgBarVal 0
    ReleaseProgressBar
    
End Sub

'This function takes an x and y value - as floating-point - and uses their position to calculate an interpolated value
' for an imaginary pixel in that location.  Offset (r/g/b/alpha) and image color depth are also required.
Public Function GetInterpolatedVal(ByVal x1 As Double, ByVal y1 As Double, ByRef iData() As Byte, ByRef iOffset As Long, ByRef iDepth As Long) As Byte
        
    'Retrieve the four surrounding pixel values
    Dim topLeft As Single, topRight As Single, bottomLeft As Single, bottomRight As Single
    topLeft = iData(Int(x1) * iDepth + iOffset, Int(y1))
    topRight = iData(Int(x1 + 1) * iDepth + iOffset, Int(y1))
    bottomLeft = iData(Int(x1) * iDepth + iOffset, Int(y1 + 1))
    bottomRight = iData(Int(x1 + 1) * iDepth + iOffset, Int(y1 + 1))
    
    'Calculate blend ratios
    Dim yBlend As Single
    Dim xBlend As Single, xBlendInv As Single
    yBlend = y1 - Int(y1)
    xBlend = x1 - Int(x1)
    xBlendInv = 1 - xBlend
    
    'Blend in the x-direction
    Dim topRowColor As Single, bottomRowColor As Single
    topRowColor = topRight * xBlend + topLeft * xBlendInv
    bottomRowColor = bottomRight * xBlend + bottomLeft * xBlendInv
    
    'Blend in the y-direction
    GetInterpolatedVal = bottomRowColor * yBlend + topRowColor * (1 - yBlend)

End Function

'This function takes an x and y value - as floating-point - and uses their position to calculate an interpolated value
' for an imaginary pixel in that location.  Offset (r/g/b/alpha) and image color depth are also required.
Public Function GetInterpolatedValWrap(ByVal x1 As Double, ByVal y1 As Double, ByVal xMax As Long, yMax As Long, ByRef iData() As Byte, ByRef iOffset As Long, ByRef iDepth As Long) As Byte
        
    'Retrieve the four surrounding pixel values
    Dim topLeft As Single, topRight As Single, bottomLeft As Single, bottomRight As Single
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
    Dim yBlend As Single
    Dim xBlend As Single, xBlendInv As Single
    yBlend = y1 - Int(y1)
    xBlend = x1 - Int(x1)
    xBlendInv = 1 - xBlend
    
    'Blend in the x-direction
    Dim topRowColor As Single, bottomRowColor As Single
    topRowColor = topRight * xBlend + topLeft * xBlendInv
    bottomRowColor = bottomRight * xBlend + bottomLeft * xBlendInv
    
    'Blend in the y-direction
    GetInterpolatedValWrap = bottomRowColor * yBlend + topRowColor * (1 - yBlend)

End Function

'XML-param wrapper for MenuFitCanvasToLayer, below
Public Sub FitCanvasToLayer_XML(ByVal processParameters As String)
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString processParameters
    MenuFitCanvasToLayer cParams.GetLong("targetlayer", pdImages(g_CurrentImage).GetActiveLayerIndex)
End Sub

'Fit the image canvas around the current layer
Public Sub MenuFitCanvasToLayer(ByVal dstLayerIndex As Long)
    
    Message "Fitting image canvas around layer..."
    
    'If the image contains an active selection, disable it before transforming the canvas
    If pdImages(g_CurrentImage).IsSelectionActive Then
        pdImages(g_CurrentImage).SetSelectionActive False
        pdImages(g_CurrentImage).MainSelection.LockRelease
    End If
    
    'Start by calculating a new offset, based on the current layer's offsets.
    Dim curLayerBounds As RECTF
    pdImages(g_CurrentImage).GetLayerByIndex(dstLayerIndex).GetLayerBoundaryRect curLayerBounds
    
    Dim dstX As Long, dstY As Long
    dstX = curLayerBounds.Left
    dstY = curLayerBounds.Top
    
    'Now that we have new top-left corner coordinates (and new width/height values), resizing the canvas
    ' is actually very easy.  In PhotoDemon, there is no such thing as "image data"; an image is just an
    ' imaginary bounding box around the layers collection.  Because of this, we don't actually need to
    ' resize any pixel data - we just need to modify all layer offsets to account for the new top-left corner!
    Dim i As Long
    For i = 0 To pdImages(g_CurrentImage).GetNumOfLayers - 1
    
        With pdImages(g_CurrentImage).GetLayerByIndex(i)
            .SetLayerOffsetX .GetLayerOffsetX - dstX
            .SetLayerOffsetY .GetLayerOffsetY - dstY
        End With
    
    Next i
    
    'Finally, update the parent image's size and DPI values
    pdImages(g_CurrentImage).UpdateSize False, curLayerBounds.Width, curLayerBounds.Height
    Interface.DisplaySize pdImages(g_CurrentImage)
    
    'In other functions, we would refresh the layer box here; however, because we haven't actually changed the
    ' appearance of any of the layers, we can leave it as-is!
    
    'Fit the new image on-screen and redraw its viewport
    ViewportEngine.Stage1_InitializeBuffer pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
    Message "Finished."
    
End Sub

'Fit the canvas around all layers present in the image
Public Sub MenuFitCanvasToAllLayers()
    
    Message "Fitting image canvas around layer..."
    
    'If the image contains an active selection, disable it before transforming the canvas
    If pdImages(g_CurrentImage).IsSelectionActive Then
        pdImages(g_CurrentImage).SetSelectionActive False
        pdImages(g_CurrentImage).MainSelection.LockRelease
    End If
    
    'Start by finding two things:
    ' 1) The lowest x/y offsets in the current layer stack
    ' 2) The highest width/height in the current layer stack (while accounting for offsets as well!)
    Dim dstLeft As Long, dstTop As Long, dstRight As Long, dstBottom As Long
    dstLeft = &HFFFFFF
    dstTop = &HFFFFFF
    dstRight = -1 * &HFFFFFF
    dstBottom = -1 * &HFFFFFF
    
    Dim curLayerBounds As RECTF
    Dim i As Long
    
    For i = 0 To pdImages(g_CurrentImage).GetNumOfLayers - 1
        
        'Get a new boundary rect, with all affine transforms accounted for
        pdImages(g_CurrentImage).GetLayerByIndex(i).GetLayerBoundaryRect curLayerBounds
        
        With curLayerBounds
        
            'Check for new minimum offsets
            If (.Left < dstLeft) Then dstLeft = .Left
            If (.Top < dstTop) Then dstTop = .Top
            
            'Check for new maximum right/top
            If (.Left + .Width > dstRight) Then dstRight = .Left + .Width
            If (.Top + .Height > dstBottom) Then dstBottom = .Top + .Height
        
        End With
    
    Next i
    
    'Now that we have new top-left corner coordinates (and new width/height values), resizing the canvas
    ' is actually very easy.  In PhotoDemon, there is no such thing as "image data"; an image is just an
    ' imaginary bounding box around the layers collection.  Because of this, we don't actually need to
    ' resize any pixel data - we just need to modify all layer offsets to account for the new top-left corner!
    For i = 0 To pdImages(g_CurrentImage).GetNumOfLayers - 1
    
        With pdImages(g_CurrentImage).GetLayerByIndex(i)
            .SetLayerOffsetX .GetLayerOffsetX - dstLeft
            .SetLayerOffsetY .GetLayerOffsetY - dstTop
        End With
    
    Next i
    
    'Finally, update the parent image's size
    pdImages(g_CurrentImage).UpdateSize False, (dstRight - dstLeft), (dstBottom - dstTop)
    Interface.DisplaySize pdImages(g_CurrentImage)
    
    'In other functions, we would refresh the layer box here; however, because we haven't actually changed the
    ' appearance of any of the layers, we can leave it as-is!
    
    'Fit the new image on-screen and redraw its viewport
    ViewportEngine.Stage1_InitializeBuffer pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
    Message "Finished."
    
End Sub

'Automatically trim empty borders from an image.  Empty borders are defined as borders comprised only of 100% transparent pixels.
Public Sub TrimImage()

    'If the image contains an active selection, disable it before transforming the canvas
    If pdImages(g_CurrentImage).IsSelectionActive Then
        pdImages(g_CurrentImage).SetSelectionActive False
        pdImages(g_CurrentImage).MainSelection.LockRelease
    End If

    'The image will be trimmed in four steps.  Each edge will be trimmed separately, starting with the top.
    
    Message "Analyzing top edge of image..."
    
    'Retrieve a copy of the composited image
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    pdImages(g_CurrentImage).GetCompositedImage tmpDIB
    
    'Point an array at the DIB data
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    PrepSafeArray srcSA, tmpDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    finalX = pdImages(g_CurrentImage).Width - 1
    finalY = pdImages(g_CurrentImage).Height - 1
            
    'These values will help us access locations in the array more quickly.
    Dim quickVal As Long
    
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
        If (srcImageData(x * 4 + 3, y) <> 0) Then
            colorFails = True
            Exit For
        End If
        
    Next x
        If colorFails Then Exit For
    Next y
    
    'We have now reached one of two conditions:
    '1) The entire image is transparent
    '2) The loop progressed part-way through the image and terminated
    
    'Check for case (1) and warn the user if it occurred
    If (Not colorFails) Then
    
        CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
        SetProgBarVal 0
        ReleaseProgressBar
        
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
        quickVal = x * 4
    For y = initY To finalY
    
        If (srcImageData(quickVal + 3, y) <> 0) Then
            colorFails = True
            Exit For
        End If
        
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
        quickVal = x * 4
    For y = initY To finalY
    
        If (srcImageData(quickVal + 3, y) <> 0) Then
            colorFails = True
            Exit For
        End If
        
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
        
        If (srcImageData(x * 4 + 3, y) <> 0) Then
            colorFails = True
            Exit For
        End If
        
    Next x
        If colorFails Then Exit For
    Next y
    
    newBottom = y
    
    'Safely deallocate imageData()
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    Erase srcImageData
    
    'Erase the temporary DIB
    Set tmpDIB = Nothing
    
    'We now know where to crop the image.  Apply the crop.
    If (newTop = 0) And (newBottom = pdImages(g_CurrentImage).Height - 1) And (newLeft = 0) And (newRight = pdImages(g_CurrentImage).Width - 1) Then
        SetProgBarVal 0
        ReleaseProgressBar
        Message "Image is already trimmed.  (No changes were made to the image.)"
    Else
    
        Message "Trimming image to new dimensions..."
        SetProgBarVal 4
        
        'Now that we have new top-left corner coordinates (and new width/height values), resizing the canvas
        ' is actually very easy.  In PhotoDemon, there is no such thing as "image data"; an image is just an
        ' imaginary bounding box around the layers collection.  Because of this, we don't actually need to
        ' resize any pixel data - we just need to modify all layer offsets to account for the new top-left corner!
        Dim i As Long
        For i = 0 To pdImages(g_CurrentImage).GetNumOfLayers - 1
        
            With pdImages(g_CurrentImage).GetLayerByIndex(i)
                .SetLayerOffsetX .GetLayerOffsetX - newLeft
                .SetLayerOffsetY .GetLayerOffsetY - newTop
            End With
        
        Next i
    
        'Finally, update the parent image's size
        pdImages(g_CurrentImage).UpdateSize False, (newRight - newLeft), (newBottom - newTop)
        Interface.DisplaySize pdImages(g_CurrentImage)
    
        'In other functions, we would refresh the layer box here; however, because we haven't actually changed the
        ' appearance of any of the layers, we can leave it as-is!
        
        Message "Finished. "
        SetProgBarVal 0
        ReleaseProgressBar
        
        'Redraw the image
        ViewportEngine.Stage1_InitializeBuffer pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
    End If

End Sub
