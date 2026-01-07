Attribute VB_Name = "Filters_Transform"
'***************************************************************************
'Image Transformations Interface (including flip/mirror/rotation/crop/etc)
'Copyright 2003-2026 by Tanner Helland
'Created: 25/January/03
'Last updated: 10/September/21
'Last update: rewrite Crop function for greatly reduced memory usage and improved performance, plus more
'              predictable behavior on layers with excessive transparent padding
'
'Functions for generic 2D transformations, including rotate, flip, mirror and crop.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'NOTE: Autocrop is currently disabled pending further testing.

'Automatically crop the image.  An optional threshold can be supplied; pixels must be this close before they will be cropped.
' (The threshold is required for JPEG images; pixels may not be identical due to lossy compression.)
Public Sub AutocropImage(Optional ByVal cThreshold As Long = 15)

    'TODO: rework this to operate on layers.  In theory, we can simply crop the pdImage width height, without
    '      actually modifying any individual layers!  The best way to do it may be to retrieve a composited
    '      copy of the image, autocrop it, then use its dimensions to change the original image's height/width.
    '      (NOTE: for left/top, all layer offsets will need to be adjusted to match.)

    'If the image contains an active selection, disable it before transforming the canvas
    'TODO: this is now handled in the central processor
    'If PDImages.GetActiveImage.IsSelectionActive Then
    '    PDImages.GetActiveImage.SetSelectionActive False
    '    PDImages.GetActiveImage.MainSelection.LockRelease
    'End If

    'The image will be cropped in four steps.  Each edge will be cropped separately, starting with the top.
    
    Message "Analyzing image..."
    PDDebug.LogAction "Analyzing top edge of image..."
    
    'Make a copy of the current image
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    tmpDIB.CreateFromExistingDIB PDImages.GetActiveImage.GetActiveDIB
    
    'Point an array at the DIB data
    Dim srcImageData() As Byte, srcSA As SafeArray2D
    tmpDIB.WrapArrayAroundDIB srcImageData, srcSA
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    finalX = PDImages.GetActiveImage.Width - 1
    finalY = PDImages.GetActiveImage.Height - 1
    
    Dim xStride As Long
    
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
        xStride = x * 4
        curColor = gLookup(CLng(srcImageData(xStride, y)) + CLng(srcImageData(xStride + 1, y)) + CLng(srcImageData(xStride + 2, y)))
        
        'If pixel color DOES NOT match the baseline, keep scanning.  Otherwise, note that we have found a mismatched color
        ' and exit the loop.
        colorFails = (Abs(curColor - initColor) > cThreshold)
        If colorFails Then Exit For
        
    Next x
        If colorFails Then Exit For
    Next y
    
    'We have now reached one of two conditions:
    '1) The entire image is one solid color
    '2) The loop progressed part-way through the image and terminated
    
    'Check for case (1) and warn the user if it occurred
    If (Not colorFails) Then
    
        tmpDIB.UnwrapArrayFromDIB srcImageData
        ProgressBars.SetProgBarVal 0
        ProgressBars.ReleaseProgressBar
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
    
    PDDebug.LogAction "Analyzing left edge of image..."
    initColor = gLookup(CLng(srcImageData(0, initY)) + CLng(srcImageData(1, initY)) + CLng(srcImageData(2, initY)))
    SetProgBarVal 1
    
    For x = 0 To finalX
        xStride = x * 4
    For y = initY To finalY
    
        curColor = gLookup(CLng(srcImageData(xStride, y)) + CLng(srcImageData(xStride + 1, y)) + CLng(srcImageData(xStride + 2, y)))
        
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
    
    PDDebug.LogAction "Analyzing right edge of image..."
    xStride = finalX * 4
    initColor = gLookup(CLng(srcImageData(xStride, initY)) + CLng(srcImageData(xStride + 1, 0)) + CLng(srcImageData(xStride + 2, 0)))
    SetProgBarVal 2
    
    For x = finalX To 0 Step -1
        xStride = x * 4
    For y = initY To finalY
    
        curColor = gLookup(CLng(srcImageData(xStride, y)) + CLng(srcImageData(xStride + 1, y)) + CLng(srcImageData(xStride + 2, y)))
        
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
    xStride = initX * 4
    initColor = gLookup(CLng(srcImageData(xStride, finalY)) + CLng(srcImageData(xStride + 1, finalY)) + CLng(srcImageData(xStride + 2, finalY)))
    
    PDDebug.LogAction "Analyzing bottom edge of image..."
    SetProgBarVal 3
    
    For y = finalY To initY Step -1
    For x = initX To finalX
        xStride = x * 4
        curColor = gLookup(CLng(srcImageData(xStride, y)) + CLng(srcImageData(xStride + 1, y)) + CLng(srcImageData(xStride + 2, y)))
        
        'If pixel color DOES NOT match the baseline, keep scanning.  Otherwise, note that we have found a mismatched color
        ' and exit the loop.
        If Abs(curColor - initColor) > cThreshold Then colorFails = True
        
        If colorFails Then Exit For
        
    Next x
        If colorFails Then Exit For
    Next y
    
    newBottom = y
    
    'Safely deallocate imageData()
    tmpDIB.UnwrapArrayFromDIB srcImageData
    
    'We now know where to crop the image.  Apply the crop.
    
    If (newTop = 0) And (newBottom = PDImages.GetActiveImage.Height - 1) And (newLeft = 0) And (newRight = PDImages.GetActiveImage.Width - 1) Then
        SetProgBarVal 0
        ReleaseProgressBar
        Message "Image is already cropped intelligently.  Autocrop abandoned.  (No changes were made to the image.)"
    Else
    
        Message "Cropping image..."
        SetProgBarVal 4
        
        'Resize the current image's main DIB
        'PDImages.GetActiveImage.mainDIB.createBlank newRight - newLeft, newBottom - newTop, tmpDIB.getDIBColorDepth
        
        'Copy the autocropped area to the new main DIB
        'GDI.BitBltWrapper PDImages.GetActiveImage.mainDIB.getDIBDC, 0, 0, PDImages.GetActiveImage.mainDIB.getDIBWidth, PDImages.GetActiveImage.mainDIB.getDIBHeight, tmpDIB.getDIBDC, newLeft, newTop, vbSrcCopy
    
        'Erase the temporary DIB
        tmpDIB.EraseDIB
        Set tmpDIB = Nothing
    
        'Update the current image size
        PDImages.GetActiveImage.UpdateSize
        Interface.DisplaySize PDImages.GetActiveImage()
        Tools.NotifyImageSizeChanged
        
        Message "Finished. "
        SetProgBarVal 0
        ReleaseProgressBar
        
        'Redraw the image
        Viewport.Stage1_InitializeBuffer PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
    End If

End Sub

'Determine if a non-destructive crop is possible.  Pure rectangular selections allow this, because we can simply modify canvas
' boundaries and layer offsets to arrive at the crop shape.
Public Sub SeeIfCropCanBeAppliedNonDestructively()
    
    If (Not PDImages.IsImageActive()) Then Exit Sub
    
    'First, make sure there is an active selection
    If (Not PDImages.GetActiveImage.IsSelectionActive) Then
        Message "No active selection found.  Crop abandoned."
    Else
        
        'Query the active selection object; if it's a pure rectangular region, we can apply a non-destructive crop
        ' to vector layers (which will allow them to remain editable).
        With PDImages.GetActiveImage.MainSelection
            
            'Start by seeing if we're even working with a rectangle.  If we are, we can check a few extra criteria
            ' as well; if we aren't, only a destructive crop will work.
            Dim selectionIsPureRectangle As Boolean
            selectionIsPureRectangle = (.GetSelectionShape = ss_Rectangle)
            
            If selectionIsPureRectangle Then
                selectionIsPureRectangle = selectionIsPureRectangle And (.GetSelectionProperty_Float(sp_RoundedCornerRadius) = 0!)
                selectionIsPureRectangle = selectionIsPureRectangle And (.GetSelectionProperty_Long(sp_Area) = sa_Interior)
                selectionIsPureRectangle = selectionIsPureRectangle And ((.GetSelectionProperty_Long(sp_Smoothing) = es_None) Or (.GetSelectionProperty_Long(sp_Smoothing) = es_Antialiased))
            End If
            
            'If the selection is a pure rectangle, we can add a further improvement by retaining vector layers.
            ' (This only works if vector layers have no non-destructive transforms applied.)
            If selectionIsPureRectangle Then
                
                Dim i As Long
                For i = 0 To PDImages.GetActiveImage.GetNumOfLayers() - 1
                    
                    'If a vector layer has active non-destructive transforms, we will need to make those transforms
                    ' permanent as part of the crop - thus, we *will* need to rasterize vector layers.
                    With PDImages.GetActiveImage.GetLayerByIndex(i)
                        If (Not .IsLayerRaster) And .AffineTransformsActive(True) Then
                            selectionIsPureRectangle = False
                            Exit For
                        End If
                    End With
                    
                Next i
                
            End If
            
            'If that huge list of above criteria are met, we can apply a non-destructive crop operation.
            Dim cParams As pdSerialize
            Set cParams = New pdSerialize
            cParams.AddParam "nondestructive", selectionIsPureRectangle
            Processor.Process "Crop", False, cParams.GetParamString(), UNDO_Everything
            
        End With
    
    End If
    
End Sub

'XML-based wrapper for CropToSelection, below.  (This exists primarily for macro support.)
Public Sub CropToSelection_XML(ByRef processParameters As String)
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString processParameters
    Filters_Transform.CropToSelection -1, cParams.GetBool("nondestructive", False)
End Sub

'Crop the image to the current selection.
' - To crop only a single layer, specify a target layer index.  (Layer > Crop menu will do this.)
' - Optionally, full-image crops on multi-layer images can sometimes be applied non-destructively (for example,
'    rectangular crops without feathering can be performed by simply modifying layer offsets and image dimensions.)
' - Non-destructive cropping is *ONLY* used on vector layers (including text).  Raster layers are always cropped
'    destructively.  (This was different in old versions of PD, but it proved confusing to users that cropping an
'    image resulted in layers with boundaries outside the image.  However, I'm willing to risk this confusion for
'    vector layers because it allows those layers to still be editable.)
Public Sub CropToSelection(Optional ByVal targetLayerIndex As Long = -1, Optional ByVal applyNonDestructively As Boolean = False)
    
    'Errors are never expected; this is an extreme failsafe, only
    On Error GoTo CropProblem
    
    'First, make sure there is an active selection
    If (Not PDImages.GetActiveImage.IsSelectionActive()) Then
        Message "No active selection found.  Crop abandoned."
        Exit Sub
    End If
    
    Message "Cropping image..."
    
    'Progress bar updates are provided by this tool
    Dim progBarCheck As Long, progBarOffsetX As Long
    
    'Layers are cropped one-at-a-time
    Dim i As Long, tmpLayerRef As pdLayer
    
    'Start by querying the current selection for its boundaries (in image coordinates).  Note that these boundaries
    ' aren't relevant to the selection mask - that's always image-sized - but we can use the relevant selection rect
    ' to minimize how many per-pixel checks we perform.
    Dim selBounds As RectF
    selBounds = PDImages.GetActiveImage.MainSelection.GetCompositeBoundaryRect
    
    'Before proceeding further, calculate a "final" new image rect (the intersection of the current image rect
    ' and the selection rect).
    Dim curImageRect As RectF
    curImageRect.Left = 0
    curImageRect.Top = 0
    curImageRect.Width = PDImages.GetActiveImage.Width
    curImageRect.Height = PDImages.GetActiveImage.Height
    
    Dim finalImageRect As RectF
    GDI_Plus.IntersectRectF finalImageRect, curImageRect, selBounds
    
    'In the past, PD would sometimes attempt to crop the image non-destructively.  (Rectangular selections allowed
    ' this because we could simply change layer offsets.)  In 9.0 this was revisited because it was unintuitive
    ' for users accustomed to other editors, which tend to *always* crop destructively.  Note that vector layers
    ' are a specific exception to this rule - they *will* be non-destructively cropped when possible, to allow
    ' the layers to remain in vector mode (instead of being rasterized).
    
    'As part of a broader "crop overhaul" in 9.0, crop logic was also revisited to reduce memory usage, improve
    ' performance, and solve some long-standing bugs when cropping layers with existing transparent borders.
    
    'Layers with non-destructive transforms active will need to be processed into a temporary DIB before cropping.
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    'Point a normal VB array at the selection mask bits.  (This is an alias, *not* a copy - so we need to release
    ' the alias before this function exits.)
    Dim selData() As Byte, selSA As SafeArray2D
    PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB.WrapArrayAroundDIB selData, selSA
    
    'Lots of helper variables follow.  They are declared this way because VB6 is pesky.
    Dim thisAlpha As Long, blendAlpha As Double
    Dim selQuickX As Long, selQuickY As Long
    Const ONE_DIVIDED_BY_255 As Double = 1# / 255#
    
    Dim x As Long, y As Long
    
    'Progress is tracked in the y-direction only
    Dim imgHeight As Long
    imgHeight = PDImages.GetActiveImage.Height
    
    'On a full image crop, we need to iterate all layers.  For a single layer crop, we do not.
    ' Determine indices for an outer loop that traverses all crop targets.
    Dim numLayersToCrop As Long, startLayerIndex As Long, endLayerIndex As Long
    If (targetLayerIndex = -1) Then
        numLayersToCrop = PDImages.GetActiveImage.GetNumOfLayers
        startLayerIndex = 0
        endLayerIndex = PDImages.GetActiveImage.GetNumOfLayers - 1
    Else
        numLayersToCrop = 1
        startLayerIndex = targetLayerIndex
        endLayerIndex = targetLayerIndex
    End If
    
    'To keep processing quick, we will only update the progress bar when absolutely necessary.
    ' This function calculates that value based on the size of the area to be processed.
    ProgressBars.SetProgBarMax numLayersToCrop * imgHeight
    progBarCheck = ProgressBars.FindBestProgBarValue()
    
    'New layer rects will be assigned based on the union of the crop rect and each layer's
    ' original rect.  (Note that complex crop shapes - e.g. circles - will do additional
    ' per-pixel work inside the new boundary rect to erase any unselected pixels, and to
    ' feather partially selected pixels.)
    Dim origLayerRect As RectF, newLayerRect As RectF
    
    'We will attempt to crop vector layers non-destructively (if we can) by simply changing
    ' layer offsets.  Note that this only works on rectangular selections; non-rectangular ones
    ' will always require rasterization.  (In the future, layer masks could be used to avoid this.)
    
    'Iterate through each layer, cropping them in turn
    For i = startLayerIndex To endLayerIndex
    
        'Update the progress bar counter for this layer
        progBarOffsetX = i * imgHeight
        
        'Point a local reference at the layer of interest
        Set tmpLayerRef = PDImages.GetActiveImage.GetLayerByIndex(i)
        
        'Cache a copy of the layer's current boundary rect.  (We'll refer to this later
        ' to determine ideal layer offsets inside the newly cropped image.)
        tmpLayerRef.GetLayerBoundaryRect origLayerRect
        
        'If this is a vector layer, and the current selection is rectangular, we can do a "fake" crop
        ' by simply changing layer offsets within the image.  This lets us avoid rasterization.
        If ((Not tmpLayerRef.IsLayerRaster()) And applyNonDestructively) Then
        
            With tmpLayerRef
                .SetLayerOffsetX .GetLayerOffsetX - selBounds.Left
                .SetLayerOffsetY .GetLayerOffsetY - selBounds.Top
            End With
            
            'Notify the parent of the change
            PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer_VectorSafe, i
        
        'This is a raster layer and/or a non-rectangular selection.  We have to do a pixel-by-pixel scan.
        Else
            
            'Make sure this layer overlaps at least partially with the selection.  If it doesn't,
            ' we can skip per-pixel processing entirely.
            If GDI_Plus.IntersectRectF(newLayerRect, origLayerRect, selBounds) Then
            
                'This layer intersects the selection region.
                
                'If the target layer has non-destructive transforms, convert it to a null-padded layer,
                ' trim it, then recalculate the intersection rect between the layer and the selection.
                If tmpLayerRef.AffineTransformsActive(True) Then
                    tmpLayerRef.ConvertToNullPaddedLayer PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height, True
                    tmpLayerRef.CropNullPaddedLayer
                    tmpLayerRef.GetLayerBoundaryRect origLayerRect
                    GDI_Plus.IntersectRectF newLayerRect, origLayerRect, selBounds
                End If
                
                'Create a new DIB at the size of the intersection between the layer and the selection mask.
                ' (This will become the backing bits for the new layer copy.)
                Set tmpDIB = New pdDIB
                tmpDIB.CreateBlank newLayerRect.Width, newLayerRect.Height, 32, 0, 0
                
                'To remove the need for a copy of the original layer bits, we are now going to copy the relevant
                ' portion of the source layer into the temporary surface we just created.  As a nice perf bonus,
                ' this will greatly reduce cache misses while applying any per-pixel selection mask processing.
                GDI.BitBltWrapper tmpDIB.GetDIBDC, 0, 0, newLayerRect.Width, newLayerRect.Height, tmpLayerRef.GetLayerDIB.GetDIBDC, newLayerRect.Left - origLayerRect.Left, newLayerRect.Top - origLayerRect.Top, vbSrcCopy
                
                'We no longer need the source layer's pixel data.  Free it.
                tmpLayerRef.GetLayerDIB.EraseDIB
                
                'Alias a VB6 array around the temporary surface
                Dim dstImageData() As RGBQuad, dstSA As SafeArray1D, tmpQuad As RGBQuad
                
                'PD's selection masks used to be 24-bpp surfaces.  Now they are 32-bpp surfaces.  As a failsafe
                ' against future changes, we don't assume a given bit-depth here.
                Dim selMaskDepth As Long
                selMaskDepth = (PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB.GetDIBColorDepth \ 8)
                
                'Offsets between the new target layer and the selection mask are fixed from this point forward.
                Dim leftOffset As Long, topOffset As Long
                leftOffset = Int(newLayerRect.Left)
                topOffset = Int(newLayerRect.Top)
                
                'We are now going to iterate through the destination surface pixel-by-pixel.  We already have a
                ' guarantee that loop boundaries are safe within the intersect rect calculated above (because that
                ' rect is the intersection of the selection mask and the new layer mask - so it lies safely
                ' within bounds of *both* surfaces).  This lets us skip pesky boundary checks in the inner loop.
                For y = 0 To newLayerRect.Height - 1
                    selQuickY = topOffset + y
                    tmpDIB.WrapRGBQuadArrayAroundScanline dstImageData, dstSA, y
                For x = 0 To newLayerRect.Width - 1
                    
                    'Calculate pixel offsets into both the destination surface (new target layer) and
                    ' the selection mask (which is never modified - it's effectively read-only).
                    selQuickX = (leftOffset + x) * selMaskDepth
                    
                    'Probe the selection mask at this point.  If it is non-white, we need to feather and/or blank
                    ' the target pixel in this location.
                    thisAlpha = selData(selQuickX, selQuickY)
                    If (thisAlpha < 255) Then
                    
                        'Check the underlying layer's alpha value.  If it's zero, we can ignore it.
                        tmpQuad = dstImageData(x)
                        If (tmpQuad.Alpha > 0) Then
                            
                            'Original pixel data will be premultiplied, which saves us a bunch of processing time.
                            ' (That's why we premultiply alpha, after all!)
                            
                            'Calculate a new multiplier, based on the strength of the selection at this location
                            blendAlpha = thisAlpha * ONE_DIVIDED_BY_255
                            
                            'Apply the multiplier to the existing pixel data (which is already premultiplied, saving us a bunch of time now)
                            tmpQuad.Blue = Int(tmpQuad.Blue * blendAlpha + 0.5)
                            tmpQuad.Green = Int(tmpQuad.Green * blendAlpha + 0.5)
                            tmpQuad.Red = Int(tmpQuad.Red * blendAlpha + 0.5)
                            
                            'Finish our work by calculating a new alpha channel value for this pixel, which is a blend of
                            ' the original alpha value, and the selection mask value at this location.
                            tmpQuad.Alpha = Int(CLng(tmpQuad.Alpha) * blendAlpha + 0.5)
                            dstImageData(x) = tmpQuad
                            
                        End If
                        
                    End If
                    
                Next x
                    If ((progBarOffsetX + y) And progBarCheck) = 0 Then ProgressBars.SetProgBarVal (progBarOffsetX + y)
                Next y
                
                'Free our unsafe array reference
                tmpDIB.UnwrapRGBQuadArrayFromDIB dstImageData
                
                'Mark target alpha as premultiplied
                tmpDIB.SetInitialAlphaPremultiplicationState True
                
                'Update the target layer's backing surface with the newly composited result
                tmpLayerRef.SetLayerDIB tmpDIB
                
                'Update the layer's offsets to match.
                If (targetLayerIndex = -1) Then
                    tmpLayerRef.SetLayerOffsetX newLayerRect.Left - selBounds.Left
                    tmpLayerRef.SetLayerOffsetY newLayerRect.Top - selBounds.Top
                Else
                    tmpLayerRef.SetLayerOffsetX newLayerRect.Left
                    tmpLayerRef.SetLayerOffsetY newLayerRect.Top
                End If
                
            'This layer does *not* intersect the newly cropped image.  I'm not entirely
            ' sure what the best option is here - ideally we'd probably just delete the
            ' damn layer (since it now exists entirely off-image), but because that
            ' could have problematic knock-on effects, let's instead just replace it
            ' with a fully transparent DIB at the current selection size.
            Else
                
                'Start by resetting all non-destructive layer transforms.
                ' (This is a nop if the layer hasn't been transformed non-destructively.)
                tmpLayerRef.MakeCanvasTransformsPermanent
                
                'Next, create a blank layer at the size of the current selection
                tmpLayerRef.GetLayerDIB.CreateBlank selBounds.Width, selBounds.Height, 32, 0, 0
                tmpLayerRef.GetLayerDIB.SetInitialAlphaPremultiplicationState True
                
                'Reset layer offsets to match the new size.  If we are resizing *all* layers,
                ' set the offset to the top-left of the new image, but if we are only cropping
                ' a single layer, instead set its top-left position to the current selection's.
                If (targetLayerIndex = -1) Then
                    tmpLayerRef.SetLayerOffsetX 0
                    tmpLayerRef.SetLayerOffsetY 0
                Else
                    tmpLayerRef.SetLayerOffsetX selBounds.Left
                    tmpLayerRef.SetLayerOffsetY selBounds.Top
                End If
                
            End If
            
            'Notify the parent of the change
            PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer, i
        
        '/end LayerIsVector and NonDestructiveCropPossible
        End If
        
        Set tmpLayerRef = Nothing
        
    Next i

    'Clear the selection mask array alias
    PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB.UnwrapArrayFromDIB selData
    
    'From here, we do some generic clean-up that's identical for both destructive
    ' and non-destructive modes. (But generally speaking, it's only relevant when
    ' *all* layers are being cropped.)
    
    'For a full-image crop, the selection is potentially out of sync with the new image size.
    ' Forcibly clear it.
    If (targetLayerIndex = -1) Then Selections.RemoveCurrentSelection False
    
    'Update the viewport.  For full-image crops, we need to refresh the entire viewport pipeline
    ' (as the image size may have changed).
    If (targetLayerIndex = -1) Then
        
        'Notify the image of its new size
        PDImages.GetActiveImage.UpdateSize False, finalImageRect.Width, finalImageRect.Height
        Interface.DisplaySize PDImages.GetActiveImage()
        Tools.NotifyImageSizeChanged
        
        'Reset the viewport to center the newly cropped image on-screen
        CanvasManager.CenterOnScreen True
        Viewport.Stage1_InitializeBuffer PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
    'For individual layers, we can use some existing viewport pipeline data
    Else
        Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    End If
    
    'Reset the progress bar to zero, then exit
    ProgressBars.SetProgBarVal 0
    ProgressBars.ReleaseProgressBar
    
    Message "Finished. "
    
    Exit Sub
    
CropProblem:
    PDDebug.LogAction "WARNING!  Filters_Transform.CropToSelection error #" & Err.Number & ": " & Err.Description
    
End Sub

'Flip an image vertically.  If no layer is specified (e.g. if targetLayerIndex = -1), all layers will be flipped.
Public Sub MenuFlip(Optional ByVal targetLayerIndex As Long = -1)

    Dim flipAllLayers As Boolean
    flipAllLayers = (targetLayerIndex = -1)
    
    Message "Flipping image..."
    
    'Iterate through each layer, flipping them in turn
    Dim tmpLayerRef As pdLayer
    
    Dim i As Long, lStart As Long, lEnd As Long
    
    'If the user wants us to process all layers, we will iterate through the full layer stack, applying the transformation to each in turn.
    ' Otherwise, we will only transform the specified layer.  To cut down on code duplication, we simply modify the endpoints of the loop.
    If flipAllLayers Then
        lStart = 0
        lEnd = PDImages.GetActiveImage.GetNumOfLayers - 1
    Else
        lStart = targetLayerIndex
        lEnd = targetLayerIndex
    End If
    
    'Loop through all relevant layers, transforming each as we go
    For i = lStart To lEnd
    
        'Retrieve a pointer to the layer of interest
        Set tmpLayerRef = PDImages.GetActiveImage.GetLayerByIndex(i)
        
        'Null-pad the layer
        If flipAllLayers Then tmpLayerRef.ConvertToNullPaddedLayer PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height
        
        'Flip it
        GDI.StretchBltWrapper tmpLayerRef.GetLayerDIB.GetDIBDC, 0, 0, tmpLayerRef.GetLayerWidth(False), tmpLayerRef.GetLayerHeight(False), tmpLayerRef.GetLayerDIB.GetDIBDC, 0, tmpLayerRef.GetLayerHeight(False) - 1, tmpLayerRef.GetLayerWidth(False), -tmpLayerRef.GetLayerHeight(False), vbSrcCopy
        
        'Remove any null-padding
        If flipAllLayers Then tmpLayerRef.CropNullPaddedLayer
        
        'Notify the parent image of the change
        PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer, i
        
    Next i
    
    'Notify the parent image that the entire image now needs to be recomposited
    PDImages.GetActiveImage.NotifyImageChanged UNDO_Image
    
    Message "Finished. "
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

'Flip an image horizontally (mirror)
Public Sub MenuMirror(Optional ByVal targetLayerIndex As Long = -1)

    Dim flipAllLayers As Boolean
    flipAllLayers = (targetLayerIndex = -1)
    
    Message "Mirroring image..."
    
    'Iterate through each layer, mirroring them in turn
    Dim tmpLayerRef As pdLayer
    
    Dim i As Long, lStart As Long, lEnd As Long
    
    'If the user wants us to process all layers, we will iterate through the full layer stack, applying the transformation to each in turn.
    ' Otherwise, we will only transform the specified layer.  To cut down on code duplication, we simply modify the endpoints of the loop.
    If flipAllLayers Then
        lStart = 0
        lEnd = PDImages.GetActiveImage.GetNumOfLayers - 1
    Else
        lStart = targetLayerIndex
        lEnd = targetLayerIndex
    End If
    
    'Loop through all relevant layers, transforming each as we go
    For i = lStart To lEnd
    
        'Retrieve a pointer to the layer of interest
        Set tmpLayerRef = PDImages.GetActiveImage.GetLayerByIndex(i)
        
        'Null-pad the layer
        If flipAllLayers Then tmpLayerRef.ConvertToNullPaddedLayer PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height
        
        'Mirror it
        GDI.StretchBltWrapper tmpLayerRef.GetLayerDIB.GetDIBDC, 0, 0, tmpLayerRef.GetLayerWidth(False), tmpLayerRef.GetLayerHeight(False), tmpLayerRef.GetLayerDIB.GetDIBDC, tmpLayerRef.GetLayerWidth(False) - 1, 0, -tmpLayerRef.GetLayerWidth(False), tmpLayerRef.GetLayerHeight(False), vbSrcCopy
        
        'Remove any null-padding
        If flipAllLayers Then tmpLayerRef.CropNullPaddedLayer
        
        'Notify the parent image of the change
        PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer, i
        
    Next i
    
    'Notify the parent image that the entire image now needs to be recomposited
    PDImages.GetActiveImage.NotifyImageChanged UNDO_Image
    
    Message "Finished. "
    
    'Redraw the viewport
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

'Rotate an image 90 degrees clockwise
Public Sub MenuRotate90Clockwise(Optional ByVal targetLayerIndex As Long = -1)

    Dim flipAllLayers As Boolean
    flipAllLayers = (targetLayerIndex = -1)
    
    Message "Rotating image..."
    
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    Dim imgWidth As Long, imgHeight As Long
    imgWidth = PDImages.GetActiveImage.Width
    imgHeight = PDImages.GetActiveImage.Height
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    If flipAllLayers Then
        SetProgBarMax PDImages.GetActiveImage.GetNumOfLayers - 1
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
        lEnd = PDImages.GetActiveImage.GetNumOfLayers - 1
    Else
        lStart = targetLayerIndex
        lEnd = targetLayerIndex
    End If
    
    'Loop through all relevant layers, transforming each as we go
    For i = lStart To lEnd
        
        'Retrieve a pointer to the layer of interest
        Set tmpLayerRef = PDImages.GetActiveImage.GetLayerByIndex(i)
        
        'Null-pad the layer
        If flipAllLayers Then tmpLayerRef.ConvertToNullPaddedLayer PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height
        
        'Make a copy of the layer, which we will use as our source during the transform
        tmpDIB.CreateFromExistingDIB tmpLayerRef.GetLayerDIB
        
        'Create a blank destination DIB to receive the transformed pixels
        tmpLayerRef.GetLayerDIB.CreateBlank imgHeight, imgWidth, 32
        
        'Use GDI+ to apply the rotation
        GDI_Plus.GDIPlusRotateFlipDIB tmpDIB, tmpLayerRef.GetLayerDIB, GP_RF_90FlipNone
        
        'Mark the correct alpha state and remove any null-padding
        tmpLayerRef.GetLayerDIB.SetInitialAlphaPremultiplicationState True
        If flipAllLayers Then tmpLayerRef.CropNullPaddedLayer
        
        'Notify the parent of the change
        PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer, i
    
        'Update the progress bar (really only relevant if rotating the entire image)
        SetProgBarVal i
    
    Next i
    
    'Update the current image size, if necessary
    If flipAllLayers Then
        PDImages.GetActiveImage.UpdateSize False, imgHeight, imgWidth
        Interface.DisplaySize PDImages.GetActiveImage()
        Tools.NotifyImageSizeChanged
    End If
    
    Message "Finished. "
    
    Viewport.Stage1_InitializeBuffer PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
    'Reset the progress bar to zero
    SetProgBarVal 0
    ReleaseProgressBar
    
End Sub

'Rotate an image 180 degrees
Public Sub MenuRotate180(Optional ByVal targetLayerIndex As Long = -1)

    Dim flipAllLayers As Boolean
    flipAllLayers = (targetLayerIndex = -1)
    
    'Fun fact: rotating 180 degrees can be accomplished by flipping and then mirroring it.
    Message "Rotating image..."
    
    'Iterate through each layer, rotating them in turn
    Dim tmpLayerRef As pdLayer
    
    Dim i As Long, lStart As Long, lEnd As Long
    
    'If the user wants us to process all layers, we will iterate through the full layer stack, applying the transformation to each in turn.
    ' Otherwise, we will only transform the specified layer.  To cut down on code duplication, we simply modify the endpoints of the loop.
    If flipAllLayers Then
        lStart = 0
        lEnd = PDImages.GetActiveImage.GetNumOfLayers - 1
    Else
        lStart = targetLayerIndex
        lEnd = targetLayerIndex
    End If
    
    'Loop through all relevant layers, transforming each as we go
    For i = lStart To lEnd
    
        'Retrieve a pointer to the layer of interest
        Set tmpLayerRef = PDImages.GetActiveImage.GetLayerByIndex(i)
        
        'Null-pad the layer
        If flipAllLayers Then tmpLayerRef.ConvertToNullPaddedLayer PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height
        
        'Rotate it by inverting both directions of a StretchBlt call
        GDI.StretchBltWrapper tmpLayerRef.GetLayerDIB.GetDIBDC, 0, 0, tmpLayerRef.GetLayerWidth(False), tmpLayerRef.GetLayerHeight(False), tmpLayerRef.GetLayerDIB.GetDIBDC, tmpLayerRef.GetLayerWidth(False) - 1, tmpLayerRef.GetLayerHeight(False) - 1, -tmpLayerRef.GetLayerWidth(False), -tmpLayerRef.GetLayerHeight(False), vbSrcCopy
        
        'Remove any null-padding
        If flipAllLayers Then tmpLayerRef.CropNullPaddedLayer
        
        'Notify the parent image of the change
        PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer, i
        
    Next i
    
    'Notify the parent image that the entire image now needs to be recomposited
    PDImages.GetActiveImage.NotifyImageChanged UNDO_Image
            
    Message "Finished. "
    
    Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

'Rotate an image 90 degrees counter-clockwise
Public Sub MenuRotate270Clockwise(Optional ByVal targetLayerIndex As Long = -1)

    Dim flipAllLayers As Boolean
    flipAllLayers = (targetLayerIndex = -1)
    
    Message "Rotating image..."
    
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    Dim imgWidth As Long, imgHeight As Long
    imgWidth = PDImages.GetActiveImage.Width
    imgHeight = PDImages.GetActiveImage.Height
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    If flipAllLayers Then
        SetProgBarMax PDImages.GetActiveImage.GetNumOfLayers - 1
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
        lEnd = PDImages.GetActiveImage.GetNumOfLayers - 1
    Else
        lStart = targetLayerIndex
        lEnd = targetLayerIndex
    End If
    
    'Loop through all relevant layers, transforming each as we go
    For i = lStart To lEnd
        
        'Retrieve a pointer to the layer of interest
        Set tmpLayerRef = PDImages.GetActiveImage.GetLayerByIndex(i)
        
        'Null-pad the layer
        If flipAllLayers Then tmpLayerRef.ConvertToNullPaddedLayer PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height
        
        'Make a copy of the layer, which we will use as our source during the transform
        tmpDIB.CreateFromExistingDIB tmpLayerRef.GetLayerDIB
        
        'Create a blank destination DIB to receive the transformed pixels
        tmpLayerRef.GetLayerDIB.CreateBlank imgHeight, imgWidth, 32
        
        'Use GDI+ to apply the rotation
        GDI_Plus.GDIPlusRotateFlipDIB tmpDIB, tmpLayerRef.GetLayerDIB, GP_RF_270FlipNone
        
        'Mark the correct alpha state and remove any null-padding
        tmpLayerRef.GetLayerDIB.SetInitialAlphaPremultiplicationState True
        If flipAllLayers Then tmpLayerRef.CropNullPaddedLayer
        
        'Notify the parent of the change
        PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer, i
    
        'Update the progress bar (really only relevant if rotating the entire image)
        SetProgBarVal i
    
    Next i
    
    'Update the current image size, if necessary
    If flipAllLayers Then
        PDImages.GetActiveImage.UpdateSize False, imgHeight, imgWidth
        Interface.DisplaySize PDImages.GetActiveImage()
        Tools.NotifyImageSizeChanged
    End If
    
    Message "Finished. "
    
    Viewport.Stage1_InitializeBuffer PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
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
Public Sub FitCanvasToLayer_XML(ByRef processParameters As String)
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString processParameters
    MenuFitCanvasToLayer cParams.GetLong("targetlayer", PDImages.GetActiveImage.GetActiveLayerIndex)
End Sub

'Fit the image canvas around the current layer
Public Sub MenuFitCanvasToLayer(ByVal dstLayerIndex As Long)
    
    Message "Fitting image canvas around layer..."
    
    'Start by calculating a new offset, based on the current layer's offsets.
    Dim curLayerBounds As RectF
    PDImages.GetActiveImage.GetLayerByIndex(dstLayerIndex).GetLayerBoundaryRect curLayerBounds
    
    Dim dstX As Long, dstY As Long
    dstX = curLayerBounds.Left
    dstY = curLayerBounds.Top
    
    'Now that we have new top-left corner coordinates (and new width/height values), resizing the canvas
    ' is actually very easy.  In PhotoDemon, there is no such thing as "image data"; an image is just an
    ' imaginary bounding box around the layers collection.  Because of this, we don't actually need to
    ' resize any pixel data - we just need to modify all layer offsets to account for the new top-left corner!
    Dim i As Long
    For i = 0 To PDImages.GetActiveImage.GetNumOfLayers - 1
    
        With PDImages.GetActiveImage.GetLayerByIndex(i)
            .SetLayerOffsetX .GetLayerOffsetX - dstX
            .SetLayerOffsetY .GetLayerOffsetY - dstY
        End With
    
    Next i
    
    'Finally, update the parent image's size and DPI values
    PDImages.GetActiveImage.UpdateSize False, curLayerBounds.Width, curLayerBounds.Height
    Interface.DisplaySize PDImages.GetActiveImage()
    Tools.NotifyImageSizeChanged
    
    'In other functions, we would refresh the layer box here; however, because we haven't actually changed the
    ' appearance of any of the layers, we can leave it as-is!
    
    'Fit the new image on-screen and redraw its viewport
    Viewport.Stage1_InitializeBuffer PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
    Message "Finished. "
    
End Sub

'Fit the canvas around all layers present in the image
Public Sub MenuFitCanvasToAllLayers()
    
    Message "Fitting image canvas around layer..."
    
    'Start by finding two things:
    ' 1) The lowest x/y offsets in the current layer stack
    ' 2) The highest width/height in the current layer stack (while accounting for offsets as well!)
    Dim dstLeft As Long, dstTop As Long, dstRight As Long, dstBottom As Long
    dstLeft = &HFFFFFF
    dstTop = &HFFFFFF
    dstRight = -1 * &HFFFFFF
    dstBottom = -1 * &HFFFFFF
    
    Dim curLayerBounds As RectF
    Dim i As Long
    
    For i = 0 To PDImages.GetActiveImage.GetNumOfLayers - 1
        
        'Get a new boundary rect, with all affine transforms accounted for
        PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerBoundaryRect curLayerBounds
        
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
    For i = 0 To PDImages.GetActiveImage.GetNumOfLayers - 1
    
        With PDImages.GetActiveImage.GetLayerByIndex(i)
            .SetLayerOffsetX .GetLayerOffsetX - dstLeft
            .SetLayerOffsetY .GetLayerOffsetY - dstTop
        End With
    
    Next i
    
    'Finally, update the parent image's size
    PDImages.GetActiveImage.UpdateSize False, (dstRight - dstLeft), (dstBottom - dstTop)
    Interface.DisplaySize PDImages.GetActiveImage()
    Tools.NotifyImageSizeChanged
    
    'In other functions, we would refresh the layer box here; however, because we haven't actually changed the
    ' appearance of any of the layers, we can leave it as-is!
    
    'Fit the new image on-screen and redraw its viewport
    Viewport.Stage1_InitializeBuffer PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
    Message "Finished. "
    
End Sub

'Automatically trim empty borders from an image.  Empty borders are defined as borders comprised only of 100% transparent pixels.
Public Sub TrimImage()
    
    'The image will be trimmed in four steps.  Each edge will be trimmed separately, starting with the top.
    Message "Analyzing image..."
    PDDebug.LogAction "Analyzing top edge of image..."
    
    'Retrieve a copy of the composited image
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    PDImages.GetActiveImage.GetCompositedImage tmpDIB
    
    'Point an array at the DIB data
    Dim srcImageData() As Byte, srcSA As SafeArray2D
    tmpDIB.WrapArrayAroundDIB srcImageData, srcSA
    
    Dim x As Long, y As Long, initX As Long, initY As Long, finalX As Long, finalY As Long
    finalX = PDImages.GetActiveImage.Width - 1
    finalY = PDImages.GetActiveImage.Height - 1
            
    'These values will help us access locations in the array more quickly.
    Dim xStride As Long
    
    'To keep processing quick, only update the progress bar when absolutely necessary.  This function calculates that value
    ' based on the size of the area to be processed.
    ProgressBars.SetProgBarMax 4
    
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
        
        tmpDIB.UnwrapArrayFromDIB srcImageData
        ProgressBars.SetProgBarVal 0
        ProgressBars.ReleaseProgressBar
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
    
    PDDebug.LogAction "Analyzing left edge of image..."
    SetProgBarVal 1
    
    For x = 0 To finalX
        xStride = x * 4
    For y = initY To finalY
    
        If (srcImageData(xStride + 3, y) <> 0) Then
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
    
    PDDebug.LogAction "Analyzing right edge of image..."
    SetProgBarVal 2
    
    For x = finalX To 0 Step -1
        xStride = x * 4
    For y = initY To finalY
    
        If (srcImageData(xStride + 3, y) <> 0) Then
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
    
    PDDebug.LogAction "Analyzing bottom edge of image..."
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
    tmpDIB.UnwrapArrayFromDIB srcImageData
    
    'Erase the temporary DIB
    Set tmpDIB = Nothing
    
    'We now know where to crop the image.  Apply the crop.
    If (newTop = 0) And (newBottom = PDImages.GetActiveImage.Height - 1) And (newLeft = 0) And (newRight = PDImages.GetActiveImage.Width - 1) Then
        SetProgBarVal 0
        ReleaseProgressBar
        Message "Image is already trimmed.  (No changes were made to the image.)"
    Else
    
        Message "Trimming image..."
        SetProgBarVal 4
        
        'Now that we have new top-left corner coordinates (and new width/height values), resizing the canvas
        ' is actually very easy.  In PhotoDemon, there is no such thing as "image data"; an image is just an
        ' imaginary bounding box around the layers collection.  Because of this, we don't actually need to
        ' resize any pixel data - we just need to modify all layer offsets to account for the new top-left corner!
        Dim i As Long
        For i = 0 To PDImages.GetActiveImage.GetNumOfLayers - 1
        
            With PDImages.GetActiveImage.GetLayerByIndex(i)
                .SetLayerOffsetX .GetLayerOffsetX - newLeft
                .SetLayerOffsetY .GetLayerOffsetY - newTop
            End With
        
        Next i
    
        'Finally, update the parent image's size
        PDImages.GetActiveImage.UpdateSize False, (newRight - newLeft), (newBottom - newTop)
        Interface.DisplaySize PDImages.GetActiveImage()
        Tools.NotifyImageSizeChanged
    
        'In other functions, we would refresh the layer box here; however, because we haven't actually changed the
        ' appearance of any of the layers, we can leave it as-is!
        
        Message "Finished. "
        SetProgBarVal 0
        ReleaseProgressBar
        
        'Redraw the image
        Viewport.Stage1_InitializeBuffer PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
    End If

End Sub
