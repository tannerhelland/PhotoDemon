Attribute VB_Name = "Tools_Crop"
'***************************************************************************
'Crop tool interface
'Copyright 2024-2026 by Tanner Helland
'Created: 12/November/24
'Last updated: 25/February/25
'Last update: implement "target image" vs "target layer" setting
'
'The crop tool performs a roughly identical task as "rectangular selection tool + Image > Crop menu".
'
'Because a number of crop-related options are user-exposed via this tool, this function has its own
' crop function (so it can handle the complex matrix of options and how they all interact).  Cropping
' via selection uses a separate function.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'TRUE if _MouseDown event was received, and _MouseUp has not yet arrived
Private m_LMBDown As Boolean

'Populated in _MouseDown
Private Const INVALID_X_COORD As Single = SINGLE_MAX, INVALID_Y_COORD As Single = SINGLE_MAX
Private m_InitImgX As Double, m_InitImgY As Double

'Populate in _MouseMove
Private m_LastImgX As Double, m_LastImgY As Double

'Current crop rect, always *in image coordinates*
Private m_CropRectF As RectF

'Some crop attributes can be independently locked (e.g. aspect ratio).
Private m_IsWidthLocked As Boolean, m_IsHeightLocked As Boolean, m_IsAspectLocked As Boolean
Private m_LockedWidth As Long, m_LockedHeight As Long
Private m_LockedAspectNumerator As Double, m_LockedAspectDenominator As Double, m_LockedAspectRatio As Double

'Index of the current mouse_over point, if any.
Private m_idxHover As PD_PointOfInterest

'On _MouseDown events, this will be set to a specific POI if the user is interacting with that POI.
' (If the user is creating a new Crop from scratch, this will be poi_Undefined.)
Private m_idxMouseDown As PD_PointOfInterest

'On _MouseMove events, the user may drag the current crop point past another crop point (for example,
' drag the lower-right point above the upper-right point).  This would cause the current POI to change.
' This constant stores the numerical index of the currently interacting point, *as defined by its
' actual position in the crop rect right now*.  (Same goes for edge-dragging.)
Private m_idxMouseDownActual As Long, m_poiEdgeMouseDownActual As PD_PointOfInterest

'If the user is modifying an existing crop boundary, this list of coordinates will be filled in _MouseDown
Private m_CornerCoords() As PointFloat, m_numCornerCoords As Long

'To correctly detect clicks, we cache the current crop rect (if any) at _MouseDown
Private m_CropRectAtMouseDown As RectF

'Because we allow _DoubleClick events to trigger a final crop event, we need to set some flags
' prior so that the subsequent _MouseUp doesn't cause UI flicker.
Private m_CropInProgress As Boolean

'Allow the crop to extend beyond image boundaries (for enlarging the image).  To accompany this,
' the caller also needs to supply allowable max/min values (and if those change, we need to be
' re-notified so we can update any existing crop region as necessary).
Private m_AllowEnlarge As Boolean, m_MaxCropWidth As Long, m_MaxCropHeight As Long

'Highlighting the retained area (by "shielding", as Photoshop calls it, the cut areas) is user-modifiable.
Private m_HighlightCrop As Boolean, m_HighlightColor As Long, m_HighlightOpacity As Single

'PhotoDemon can retain pixels outside the crop area (e.g. non-destructive cropping).
' This setting is user-modifiable.
Private m_DeleteCroppedPixels As Boolean

'Users can crop either the full image (default) or only the active layer.
Private m_CropAllLayers As Boolean

'Users can ask for a guide overlay in various ratios
Private Enum PD_CropGuide
    cg_None = 0
    cg_Centers = 1
    cg_RuleOfThirds = 2
    cg_RuleOfFifths = 3
    cg_GoldenRatio = 4
    cg_Diagonals = 5
End Enum

#If False Then
    Private Const cg_None = 0, cg_Centers = 1, cg_RuleOfThirds = 2, cg_RuleOfFifths = 3, cg_GoldenRatio = 4, cg_Diagonals = 5
#End If

Private m_CropGuides As PD_CropGuide

'Apply a crop operation from a param string constructed by the GetCropParamString() function.
' (This function is a thin wrapper to Crop_ApplyCrop(); it just extracts relevant params from the
'  param string and forwards the results.)
Public Sub Crop_ApplyFromString(ByRef paramString As String)
    
    Const FUNC_NAME As String = "Crop_ApplyFromString"
     
    'Extract string parameters and convert them to actual types
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString paramString
    
    Dim targetCropRectF As RectF
    Dim applyNonDestructively As Boolean, targetLayerOnly As Boolean, idxTargetLayer As Long
    
    'Only rectangular crops are currently supported
    If Strings.StringsNotEqual("rect", cParams.GetString("type", "rect", True), True) Then
        InternalError FUNC_NAME, "no crop rect"
        Exit Sub
    Else
        
        With cParams
            targetCropRectF.Left = .GetDouble("left", 0#, True)
            targetCropRectF.Top = .GetDouble("top", 0#, True)
            targetCropRectF.Width = .GetDouble("width", 0#, True)
            targetCropRectF.Height = .GetDouble("height", 0#, True)
            applyNonDestructively = Not .GetBool("delete-pixels", m_DeleteCroppedPixels, True)
            targetLayerOnly = .GetBool("target-layer", False, True)
        End With
        
        idxTargetLayer = -1
        If targetLayerOnly Then idxTargetLayer = PDImages.GetActiveImage.GetActiveLayerIndex
        
        Crop_ApplyCropRect targetCropRectF, idxTargetLayer, applyNonDestructively
    
    End If
    
End Sub

'Crop the image against coordinates supplied by the on-canvas crop tool.
' - To crop only a single layer, specify a target layer index.
' - Full-image crops on multi-layer images can be applied non-destructively (by simply modifying layer offsets
'    and parent image dimensions.)
' - Destructive cropping requires rasterization of vector layers, *if* they overlap crop boundaries.
Private Sub Crop_ApplyCropRect(ByRef cropRectF As RectF, Optional ByVal targetLayerIndex As Long = -1, Optional ByVal applyNonDestructively As Boolean = True)
    
    Const FUNC_NAME As String = "Crop_ApplyCrop"
    
    'Errors are never expected; this is an extreme failsafe, only
    On Error GoTo CropProblem
    
    'A few more failsafe checks for valid crop areas.
    m_CropRectF = cropRectF
    If (Not IsValidCropActive()) Then
        InternalError FUNC_NAME, "bad crop rect"
        Exit Sub
    End If
    
    'Before we apply a crop, selections must be removed
    If Selections.SelectionsAllowed(False) Then Selections.RemoveCurrentSelection False
    
    'Just an FYI before we begin: an important distinction between this tool and selection-based cropping
    ' is that this tool can explicitly *grow* image dimensions if/when the crop boundary lies outside the image.
    ' As such, the overlap between any given layer and the crop rect is highly variable, and it's fully
    ' possible for layers and cropped regions to not overlap at all during a crop.
    
    Message "Cropping image..."
    
    'Layers are cropped one-at-a-time
    Dim i As Long, tmpLayerRef As pdLayer
    
    'Layers with active transforms need to be rasterized into a temporary DIB before cropping.
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    'This function can crop the entire image, or individual layer(s)
    Dim croppingWholeImage As Boolean
    croppingWholeImage = (targetLayerIndex < 0)
    
    'On a full image crop, we need to iterate all layers.  For a single layer crop, we do not.
    ' Determine indices for an outer loop that traverses all crop targets.
    Dim numLayersToCrop As Long, startLayerIndex As Long, endLayerIndex As Long
    If croppingWholeImage Then
        numLayersToCrop = PDImages.GetActiveImage.GetNumOfLayers
        startLayerIndex = 0
        endLayerIndex = PDImages.GetActiveImage.GetNumOfLayers - 1
    Else
        numLayersToCrop = 1
        startLayerIndex = targetLayerIndex
        endLayerIndex = targetLayerIndex
    End If
    
    'To keep processing quick, we only update the progress bar when absolutely necessary.
    ' This function calculates that value based on the size of the area to be processed.
    ProgressBars.SetProgBarMax numLayersToCrop
    
    'New layer rects are assigned based on the union of the crop rect and each layer's original boundary rect.
    Dim origLayerRect As RectF, newLayerRect As RectF
    Dim numLayersProcessed As Long
    
    'Iterate through each layer, cropping them in turn
    For i = startLayerIndex To endLayerIndex
        
        ProgressBars.SetProgBarVal numLayersProcessed + 1
        
        'Point a local reference at the layer of interest
        Set tmpLayerRef = PDImages.GetActiveImage.GetLayerByIndex(i)
        
        'Cache a copy of the layer's current boundary rect.  (We'll refer to this later
        ' to determine ideal layer offsets inside the newly cropped image.)
        tmpLayerRef.GetLayerBoundaryRect origLayerRect
        
        'If this is a vector layer, and the current selection is rectangular, we can do a "fake" crop
        ' by simply changing layer offsets within the image.  This lets us avoid rasterization.
        If applyNonDestructively Then
        
            With tmpLayerRef
                .SetLayerOffsetX .GetLayerOffsetX - cropRectF.Left
                .SetLayerOffsetY .GetLayerOffsetY - cropRectF.Top
            End With
            
            'Notify the parent of the change
            PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer_VectorSafe, i
        
        'This is a raster layer and/or a non-rectangular selection.  We have to do a pixel-by-pixel scan.
        Else
            
            'Make sure this layer overlaps at least partially with the crop area.
            ' (If it lies fully within the crop, we can simply move it, and if it lies fully outside the crop,
            '  we can just delete it.)
            If GDI_Plus.IntersectRectF(newLayerRect, origLayerRect, cropRectF) Then
            
                'This layer intersects the selection region.
                
                'If the target layer has non-destructive transforms, rasterize it, trim it,
                ' then recalculate the intersection rect between the layer and the selection.
                If tmpLayerRef.AffineTransformsActive(True) Then
                    tmpLayerRef.ConvertToNullPaddedLayer PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height, True
                    tmpLayerRef.CropNullPaddedLayer
                    tmpLayerRef.GetLayerBoundaryRect origLayerRect
                    GDI_Plus.IntersectRectF newLayerRect, origLayerRect, cropRectF
                End If
                
                'TODO: look for a selection that's fully contained and skip the next step.
                
                'Create a new DIB at the size of the intersection between the layer and the crop rect.
                ' (This will become the backing bits for the new layer copy.)
                Set tmpDIB = New pdDIB
                tmpDIB.CreateBlank newLayerRect.Width, newLayerRect.Height, 32, 0, 0
                
                'To remove the need for a copy of the original layer bits, we are now going to copy the relevant
                ' portion of the source layer into the temporary surface we just created.  As a nice perf bonus,
                ' this will greatly reduce cache misses while applying any per-pixel modifications.
                GDI.BitBltWrapper tmpDIB.GetDIBDC, 0, 0, newLayerRect.Width, newLayerRect.Height, tmpLayerRef.GetLayerDIB.GetDIBDC, newLayerRect.Left - origLayerRect.Left, newLayerRect.Top - origLayerRect.Top, vbSrcCopy
                
                'We no longer need the source layer's pixel data.  Free it.
                tmpLayerRef.GetLayerDIB.EraseDIB
                
                'Mark target alpha as premultiplied
                tmpDIB.SetInitialAlphaPremultiplicationState True
                
                'Update the target layer's backing surface with the newly composited result
                tmpLayerRef.SetLayerDIB tmpDIB
                
                'Update the layer's offsets to match.
                If croppingWholeImage Then
                    tmpLayerRef.SetLayerOffsetX newLayerRect.Left - cropRectF.Left
                    tmpLayerRef.SetLayerOffsetY newLayerRect.Top - cropRectF.Top
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
                tmpLayerRef.GetLayerDIB.CreateBlank cropRectF.Width, cropRectF.Height, 32, 0, 0
                tmpLayerRef.GetLayerDIB.SetInitialAlphaPremultiplicationState True
                
                'Reset layer offsets to match the new size.  If we are resizing *all* layers,
                ' set the offset to the top-left of the new image, but if we are only cropping
                ' a single layer, instead set its top-left position to the current selection's.
                If croppingWholeImage Then
                    tmpLayerRef.SetLayerOffsetX 0
                    tmpLayerRef.SetLayerOffsetY 0
                Else
                    tmpLayerRef.SetLayerOffsetX cropRectF.Left
                    tmpLayerRef.SetLayerOffsetY cropRectF.Top
                End If
                
            End If
            
            'Notify the parent of the change
            PDImages.GetActiveImage.NotifyImageChanged UNDO_Layer, i
        
        '/end LayerIsVector and NonDestructiveCropPossible
        End If
        
        Set tmpLayerRef = Nothing
        numLayersProcessed = numLayersProcessed + 1
        
    Next i
    
    'From here, we do some generic clean-up that's identical for both destructive
    ' and non-destructive modes. (But generally speaking, it's only relevant when
    ' *all* layers are being cropped.)
    
    'Start clean-up by removing the active crop rect and updating some crop tool UI bits
    ResetCropRectF
    toolpanel_Crop.cmdCommit(0).Enabled = Tools_Crop.IsValidCropActive()
    toolpanel_Crop.cmdCommit(1).Enabled = Tools_Crop.IsValidCropActive()
    
    'If cropping the entire image, notify the parent image object of the new size.
    ' (Layers don't need this notification; they always auto-sync against their backing surface.)
    If croppingWholeImage Then
        PDImages.GetActiveImage.UpdateSize False, cropRectF.Width, cropRectF.Height
        Interface.DisplaySize PDImages.GetActiveImage()
        Tools.NotifyImageSizeChanged
    End If
    
    'Update the viewport.  For full-image crops, we need to refresh the entire viewport pipeline
    ' (as the image size may have changed).
    If croppingWholeImage Then
    
        'Reset the viewport to center the newly cropped image on-screen
        CanvasManager.CenterOnScreen True
        Viewport.Stage1_InitializeBuffer PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
    'When cropping individual layers, we can reuse some existing viewport pipeline data
    Else
        Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    End If
    
    'Reset the progress bar to zero, then exit
    ProgressBars.SetProgBarVal 0
    ProgressBars.ReleaseProgressBar
    
    Message "Finished. "
    
    Exit Sub
    
CropProblem:
    InternalError FUNC_NAME, "WARNING! Error #" & Err.Number & ": " & Err.Description
    
End Sub

'Commit (apply) the current crop rectangle.  *All* layers will be affected by this.
' This function takes the current crop settings, builds a string parameter list from them, then calls
' PD's central processor with the resulting string.  (This allows the crop to be recorded.)
Public Sub CommitCurrentCrop()
    
    'TODO: undo type needs to be flagged as vector_safe here pending a UI toggle for erasing cropped areas
    
    If Tools_Crop.IsValidCropActive() Then
        Process "Crop tool", False, GetCropParamString(), UNDO_Image, ND_CROP
    End If
    
End Sub

'All operations in PD derive from parameter strings.  This allows us to save operations to file in a human-readable format.
Private Function GetCropParamString() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    'Start with generic crop settings
    ' (TODO)
    
    'Finally, add the corner coordinates of the crop region
    With cParams
        .AddParam "type", "rect"
        .AddParam "left", m_CropRectF.Left
        .AddParam "top", m_CropRectF.Top
        .AddParam "width", m_CropRectF.Width
        .AddParam "height", m_CropRectF.Height
        .AddParam "delete-pixels", m_DeleteCroppedPixels
        .AddParam "target-layer", (Not m_CropAllLayers)
    End With
    
    GetCropParamString = cParams.GetParamString()
    
End Function

'Turn off the current crop rectangle.  No image modifications will be made.
Public Sub RemoveCurrentCrop()
    ResetCropRectF
    RelayCropChangesToUI
    Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage, FormMain.MainCanvas(0)
End Sub

Public Sub DrawCanvasUI(ByRef dstCanvas As pdCanvas, ByRef srcImage As pdImage)
    
    'Update the status bar with the curent crop rectangle (if any)
    dstCanvas.SetSelectionState ValidateCropRectF
    
    'Skip if the current rect is invalid
    If (Not ValidateCropRectF()) Then Exit Sub
    
    'Clone a pair of UI pens from the main rendering module.  (Note that we clone them because we need
    ' to modify some rendering properties, and we don't want to fuck up the central cache.)
    Dim basePenInactive As pd2DPen, topPenInactive As pd2DPen
    Dim basePenActive As pd2DPen, topPenActive As pd2DPen
    Drawing.CloneCachedUIPens basePenInactive, topPenInactive, False
    Drawing.CloneCachedUIPens basePenActive, topPenActive, True
    
    'Specify rounded line edges for our pens; this looks better for this particular tool
    basePenInactive.SetPenLineCap P2_LC_Round
    topPenInactive.SetPenLineCap P2_LC_Round
    basePenActive.SetPenLineCap P2_LC_Round
    topPenActive.SetPenLineCap P2_LC_Round
    
    basePenInactive.SetPenLineJoin P2_LJ_Round
    topPenInactive.SetPenLineJoin P2_LJ_Round
    basePenActive.SetPenLineJoin P2_LJ_Round
    topPenActive.SetPenLineJoin P2_LJ_Round
    
    'Create other required pd2D drawing tools (a surface)
    Dim cSurface As pd2DSurface
    Drawing2D.QuickCreateSurfaceFromDC cSurface, dstCanvas.hDC, True
    cSurface.SetSurfacePixelOffset P2_PO_Normal
    
    'Convert the crop rectangle to canvas coordinates
    Dim cropRectCanvas As RectF
    Drawing.ConvertImageCoordsToCanvasCoords_RectF dstCanvas, srcImage, m_CropRectF, cropRectCanvas, False
    
    Dim i As Long
    
    'If the user wants the retained area highlighted (or more accurately, the cut regions shadowed),
    ' we'll render that first, then render the outline and interactive bits atop it.
    If m_HighlightCrop Then
        
        cSurface.SetSurfacePixelOffset P2_PO_Half
        
        'Get the current canvas image rectangle
        Dim imgCanvasRectF As RectF
        PDImages.GetActiveImage.ImgViewport.GetIntersectRectCanvas imgCanvasRectF
        
        'Because GDI+ precision is iffy, expand the image canvas rect by one pixel in each direction
        imgCanvasRectF.Left = imgCanvasRectF.Left - 1!
        imgCanvasRectF.Top = imgCanvasRectF.Top - 1!
        imgCanvasRectF.Width = imgCanvasRectF.Width + 2!
        imgCanvasRectF.Height = imgCanvasRectF.Height + 2!
        
        'Subtract the crop rect from the image rect, using regions.
        Dim cShadowRegion As pd2DRegion
        Set cShadowRegion = New pd2DRegion
        cShadowRegion.AddRectangle_FromRectF imgCanvasRectF, P2_CM_Replace
        
        '(Again, note that we modify the rect slightly to ensure pretty rendering.)
        Dim tmpCropCanvas As RectF
        tmpCropCanvas.Left = cropRectCanvas.Left - 1!
        tmpCropCanvas.Top = cropRectCanvas.Top - 1!
        tmpCropCanvas.Width = cropRectCanvas.Width + 2!
        tmpCropCanvas.Height = cropRectCanvas.Height + 2!
        cShadowRegion.AddRectangle_FromRectF tmpCropCanvas, P2_CM_Exclude
        
        'Retrieve the final, composited region as a collection of rectangles
        Dim numRects As Long, regionAsRectFs() As RectF
        If cShadowRegion.GetRegionAsRectFs(numRects, regionAsRectFs) Then
            
            'Shadow each region using a brush filled with the user's highlight settings
            Dim tmpCropBrush As pd2DBrush
            Drawing2D.QuickCreateSolidBrush tmpCropBrush, m_HighlightColor, m_HighlightOpacity
            
            If (numRects > 0) Then
                If (UBound(regionAsRectFs) >= numRects - 1) Then
                    For i = 0 To numRects - 1
                        PD2D.FillRectangleF_FromRectF cSurface, tmpCropBrush, regionAsRectFs(i)
                    Next i
                End If
            End If
            
            Set tmpCropBrush = Nothing
            
        '/failed to convert the region to rectangles; the crop region is probably non-existent,
        ' so no Else required
        End If
        
        cSurface.SetSurfacePixelOffset P2_PO_Normal
        
    End If
    
    'Now it's time to render the crop outline, crop guides, and any on-canvas interactive UI bits
    Dim basePenToUse As pd2DPen, topPenToUse As pd2DPen
    
    'Start with guides (so that the crop rect itself can "overlay" the guides, if any)
    Dim ptGuides() As PointFloat, numPtsRender As Long, renderPointsNow As Boolean
    
    Set basePenToUse = basePenInactive
    Set topPenToUse = topPenInactive
    
    Select Case m_CropGuides
        
        Case cg_None
            renderPointsNow = False
            
        Case cg_Centers
            renderPointsNow = True
            numPtsRender = 2
            ReDim ptGuides(0 To numPtsRender * 2 - 1) As PointFloat
            ptGuides(0).x = cropRectCanvas.Left + cropRectCanvas.Width \ 2
            ptGuides(0).y = cropRectCanvas.Top
            ptGuides(1).x = ptGuides(0).x
            ptGuides(1).y = cropRectCanvas.Top + cropRectCanvas.Height
            ptGuides(2).x = cropRectCanvas.Left
            ptGuides(2).y = cropRectCanvas.Top + cropRectCanvas.Height \ 2
            ptGuides(3).x = cropRectCanvas.Left + cropRectCanvas.Width
            ptGuides(3).y = ptGuides(2).y
            
        Case cg_RuleOfThirds
            renderPointsNow = True
            numPtsRender = 4
            ReDim ptGuides(0 To numPtsRender * 2 - 1) As PointFloat
            For i = 0 To numPtsRender - 1 Step 2
                ptGuides(i).x = cropRectCanvas.Left + (cropRectCanvas.Width * (i \ 2 + 1) \ 3)
                ptGuides(i).y = cropRectCanvas.Top
                ptGuides(i + 1).x = ptGuides(i).x
                ptGuides(i + 1).y = cropRectCanvas.Top + cropRectCanvas.Height
            Next i
            For i = 0 To numPtsRender - 1 Step 2
                ptGuides(numPtsRender + i).x = cropRectCanvas.Left
                ptGuides(numPtsRender + i).y = cropRectCanvas.Top + (cropRectCanvas.Height * (i \ 2 + 1) \ 3)
                ptGuides(numPtsRender + i + 1).x = cropRectCanvas.Left + cropRectCanvas.Width
                ptGuides(numPtsRender + i + 1).y = ptGuides(numPtsRender + i).y
            Next i
            
        Case cg_RuleOfFifths
            renderPointsNow = True
            numPtsRender = 8
            ReDim ptGuides(0 To numPtsRender * 2 - 1) As PointFloat
            For i = 0 To numPtsRender - 1 Step 2
                ptGuides(i).x = cropRectCanvas.Left + (cropRectCanvas.Width * (i \ 2 + 1) \ 5)
                ptGuides(i).y = cropRectCanvas.Top
                ptGuides(i + 1).x = ptGuides(i).x
                ptGuides(i + 1).y = cropRectCanvas.Top + cropRectCanvas.Height
            Next i
            For i = 0 To numPtsRender - 1 Step 2
                ptGuides(numPtsRender + i).x = cropRectCanvas.Left
                ptGuides(numPtsRender + i).y = cropRectCanvas.Top + (cropRectCanvas.Height * (i \ 2 + 1) \ 5)
                ptGuides(numPtsRender + i + 1).x = cropRectCanvas.Left + cropRectCanvas.Width
                ptGuides(numPtsRender + i + 1).y = ptGuides(numPtsRender + i).y
            Next i
            
        Case cg_GoldenRatio
            renderPointsNow = True
            Const GOLDEN_RATIO As Double = 0.61803398874989
            numPtsRender = 4
            ReDim ptGuides(0 To numPtsRender * 2 - 1) As PointFloat
            ptGuides(0).x = cropRectCanvas.Left + (cropRectCanvas.Width * GOLDEN_RATIO)
            ptGuides(0).y = cropRectCanvas.Top
            ptGuides(1).x = ptGuides(0).x
            ptGuides(1).y = cropRectCanvas.Top + cropRectCanvas.Height
            
            ptGuides(2).x = cropRectCanvas.Left + cropRectCanvas.Width - (cropRectCanvas.Width * GOLDEN_RATIO)
            ptGuides(2).y = cropRectCanvas.Top
            ptGuides(3).x = ptGuides(2).x
            ptGuides(3).y = cropRectCanvas.Top + cropRectCanvas.Height
            
            ptGuides(4).x = cropRectCanvas.Left
            ptGuides(4).y = cropRectCanvas.Top + (cropRectCanvas.Height * GOLDEN_RATIO)
            ptGuides(5).x = cropRectCanvas.Left + cropRectCanvas.Width
            ptGuides(5).y = ptGuides(4).y
            
            ptGuides(6).x = cropRectCanvas.Left
            ptGuides(6).y = cropRectCanvas.Top + cropRectCanvas.Height - (cropRectCanvas.Height * GOLDEN_RATIO)
            ptGuides(7).x = cropRectCanvas.Left + cropRectCanvas.Width
            ptGuides(7).y = ptGuides(6).y
            
        Case cg_Diagonals
            renderPointsNow = True
            numPtsRender = 4
            ReDim ptGuides(0 To numPtsRender * 2 - 1) As PointFloat
            
            Dim minDimension As Single
            minDimension = PDMath.Min2Float_Single(cropRectCanvas.Width, cropRectCanvas.Height)
            
            ptGuides(0).x = cropRectCanvas.Left
            ptGuides(0).y = cropRectCanvas.Top
            ptGuides(1).x = ptGuides(0).x + minDimension
            ptGuides(1).y = ptGuides(0).y + minDimension
            
            ptGuides(2).x = cropRectCanvas.Left + cropRectCanvas.Width
            ptGuides(2).y = cropRectCanvas.Top
            ptGuides(3).x = ptGuides(2).x - minDimension
            ptGuides(3).y = cropRectCanvas.Top + minDimension
            
            ptGuides(4).x = cropRectCanvas.Left
            ptGuides(4).y = cropRectCanvas.Top + cropRectCanvas.Height
            ptGuides(5).x = cropRectCanvas.Left + minDimension
            ptGuides(5).y = cropRectCanvas.Top + cropRectCanvas.Height - minDimension
            
            ptGuides(6).x = cropRectCanvas.Left + cropRectCanvas.Width
            ptGuides(6).y = cropRectCanvas.Top + cropRectCanvas.Height
            ptGuides(7).x = cropRectCanvas.Left + cropRectCanvas.Width - minDimension
            ptGuides(7).y = ptGuides(6).y - minDimension
        
    End Select
    
    If renderPointsNow Then
        For i = 0 To numPtsRender * 2 - 1 Step 2
            PD2D.DrawLineF_FromPtF cSurface, basePenInactive, ptGuides(i), ptGuides(i + 1)
        Next i
        For i = 0 To numPtsRender * 2 - 1 Step 2
            PD2D.DrawLineF_FromPtF cSurface, topPenInactive, ptGuides(i), ptGuides(i + 1)
        Next i
    End If
    
    Set basePenToUse = Nothing
    Set topPenToUse = Nothing
    
    'Determine if the user is interacting with a corner node or edge node (or by correlation, neither)
    Dim userIsCornerDragging As Boolean, userIsEdgeDragging As Boolean
    If (m_idxMouseDown >= 0) And (m_idxMouseDown <= 3) Then
        userIsCornerDragging = True
    ElseIf (m_idxMouseDown >= poi_EdgeW) And (m_idxMouseDown <= poi_EdgeN) Then
        userIsEdgeDragging = True
    End If
    
    Dim userIsCornerHover As Boolean, userIsEdgeHover As Boolean
    If (m_idxHover >= 0) And (m_idxHover <= 3) Then
        userIsCornerHover = True
    ElseIf (m_idxHover >= poi_EdgeW) And (m_idxHover <= poi_EdgeN) Then
        userIsEdgeHover = True
    End If
    
    Dim drawActiveOutline As Boolean
    
    'Edge boundaries need to be rendered individually, because we highlight the currently interacted edge
    ' (if any).
    If (userIsEdgeDragging Or userIsEdgeHover) Then
        
        'Use the current hovered or clicked edge index, whichever is relevant to mouse state
        Dim targetPOI As PD_PointOfInterest
        targetPOI = m_idxHover
        If m_LMBDown Then targetPOI = m_idxMouseDown
        
        'Map that to a 0-based index
        Dim idxTarget As Long
        Select Case targetPOI
            Case poi_EdgeN
                idxTarget = 0
            Case poi_EdgeE
                idxTarget = 1
            Case poi_EdgeS
                idxTarget = 2
            Case poi_EdgeW
                idxTarget = 3
            Case Else
                idxTarget = -1
        End Select
        
        'Map the crop rectangle to an arbitrary list of points, and include the first point twice
        ' (once at the start, once at the end to simplify rendering).
        Dim listOfRectPts() As PointFloat
        ReDim listOfRectPts(0 To 4) As PointFloat
        With cropRectCanvas
            listOfRectPts(0).x = .Left
            listOfRectPts(0).y = .Top
            listOfRectPts(1).x = .Left + .Width
            listOfRectPts(1).y = .Top
            listOfRectPts(2).x = .Left + .Width
            listOfRectPts(2).y = .Top + .Height
            listOfRectPts(3).x = .Left
            listOfRectPts(3).y = .Top + .Height
        End With
        
        listOfRectPts(4) = listOfRectPts(0)
        
        'Render each edge in turn, highlighting as relevant
        For i = 0 To 3
            
            If (idxTarget = i) Then
                Set basePenToUse = basePenActive
                Set topPenToUse = topPenActive
            Else
                Set basePenToUse = basePenInactive
                Set topPenToUse = topPenInactive
            End If
            
            PD2D.DrawLineF_FromPtF cSurface, basePenToUse, listOfRectPts(i), listOfRectPts(i + 1)
            PD2D.DrawLineF_FromPtF cSurface, topPenToUse, listOfRectPts(i), listOfRectPts(i + 1)
            
        Next i
        
        Set basePenToUse = Nothing: Set topPenToUse = Nothing
        
    Else
        
        Dim cropOutline As pd2DPath
        Set cropOutline = New pd2DPath
        cropOutline.AddRectangle_RectF cropRectCanvas
        
        'If mouse is down, use the clicked corner instead of the hovered one (as the hovered one may change if
        ' corner nodes are dragged extremely close together).
        drawActiveOutline = (m_idxHover = poi_Interior)
        If m_LMBDown Then drawActiveOutline = (m_idxMouseDown = poi_Interior)
        
        'Stroke the outline path
        If drawActiveOutline Then
            PD2D.DrawPath cSurface, basePenActive, cropOutline
            PD2D.DrawPath cSurface, topPenActive, cropOutline
        Else
            PD2D.DrawPath cSurface, basePenInactive, cropOutline
            PD2D.DrawPath cSurface, topPenInactive, cropOutline
        End If
        
    End If
    
    'Next, we want to render drag outlines over the four corner vertices.
    Dim ptCorners() As PointFloat, numPoints As Long
    numPoints = GetCropCorners(ptCorners)
    
    'Next, convert each corner from image coordinate space to the active viewport coordinate space
    Drawing.ConvertListOfImageCoordsToCanvasCoords dstCanvas, srcImage, ptCorners, False
    
    Dim cornerSize As Single, halfCornerSize As Single
    cornerSize = SQUARE_CORNER_SIZE
    halfCornerSize = cornerSize * 0.5!
    
    Dim idxActive As PD_PointOfInterest
    idxActive = poi_Undefined
    If m_LMBDown Then
        If (m_idxMouseDown <> poi_Interior) Then idxActive = m_idxMouseDownActual
    Else
        idxActive = m_idxHover
    End If
    
    'Convert from PD constants to numerical indices
    Select Case idxActive
        Case poi_CornerNW
            idxActive = 0
        Case poi_CornerNE
            idxActive = 1
        Case poi_CornerSW
            idxActive = 2
        Case poi_CornerSW
            idxActive = 3
    End Select
    
    For i = 0 To numPoints - 1
        If (i = idxActive) And Not (userIsEdgeDragging Or userIsEdgeHover) Then
            PD2D.DrawRectangleF cSurface, basePenActive, ptCorners(i).x - halfCornerSize, ptCorners(i).y - halfCornerSize, cornerSize, cornerSize
            PD2D.DrawRectangleF cSurface, topPenActive, ptCorners(i).x - halfCornerSize, ptCorners(i).y - halfCornerSize, cornerSize, cornerSize
        Else
            PD2D.DrawRectangleF cSurface, basePenInactive, ptCorners(i).x - halfCornerSize, ptCorners(i).y - halfCornerSize, cornerSize, cornerSize
            PD2D.DrawRectangleF cSurface, topPenInactive, ptCorners(i).x - halfCornerSize, ptCorners(i).y - halfCornerSize, cornerSize, cornerSize
        End If
    Next i
    
    'Free rendering objects
    Set cSurface = Nothing
    Set basePenInactive = Nothing: Set topPenInactive = Nothing
    Set basePenActive = Nothing: Set topPenActive = Nothing
    
End Sub

Public Sub NotifyDoubleClick(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    If (Not PDImages.IsImageActive()) Then Exit Sub
    If (Not IsValidCropActive()) Then Exit Sub
    
    If ((Button And pdLeftButton) = pdLeftButton) Then
    
        'Translate the canvas coordinates to image coordinates
        Dim imgX As Double, imgY As Double
        Drawing.ConvertCanvasCoordsToImageCoords FormMain.MainCanvas(0), PDImages.GetActiveImage(), x, y, imgX, imgY, False
        
        'Set a module-level flag; other _MouseEvent handles will ignore mouse behavior until this flag is un-set
        m_CropInProgress = True
        
        'Check mouse position.  If the mouse is *outside* the crop rect that existed at _MouseDown,
        ' simply clear the current rect.
        If PDMath.IsPointInRectF(imgX, imgY, m_CropRectAtMouseDown) Then
            
            'Set a module-level flag so that the subsequent _MouseUp doesn't fire
            m_CropInProgress = True
            CommitCurrentCrop
            ResetCropRectF 'Failsafe because sometimes laptop touchpads generate superfluous mouse events
            
            'm_CropInProgress can only be reset by a subsequent _MouseDown or _MouseUp event;
            ' do not clear it here.
            
        End If
        
    End If
    
End Sub

Public Sub NotifyKeyDown(ByVal Shift As ShiftConstants, ByVal vkCode As Long, ByRef markEventHandled As Boolean)
    
    If (vkCode = VK_ESCAPE) Then
        Tools_Crop.RemoveCurrentCrop
        markEventHandled = True
    ElseIf (vkCode = VK_RETURN) Then
        If Tools_Crop.IsValidCropActive() Then
            Tools_Crop.CommitCurrentCrop
            markEventHandled = True
        End If
    End If
    
End Sub

Public Sub NotifyMouseDown(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByRef srcCanvas As pdCanvas, ByRef srcImage As pdImage, ByVal canvasX As Single, ByVal canvasY As Single)
    
    'As a failsafe only, reset the "in the midst of a double-click" flag.
    ' (Under normal circumstances, this will be cleared by the _MouseUp that fires *after* a double-click flag,
    '  but failsafes are useful because laptop touchpads and bluetooth mice do not always behave.)
    m_CropInProgress = False
    
    'Cache initial x/y positions
    m_LMBDown = ((Button And pdLeftButton) = pdLeftButton)
    
    If m_LMBDown Then
        
        'Cache the current crop rect, if any.
        ' (on _Click events, we'll compare click coords against *this* rect.)
        m_CropRectAtMouseDown = m_CropRectF
        
        'Translate the canvas coordinates to image coordinates
        Dim imgX As Double, imgY As Double
        Drawing.ConvertCanvasCoordsToImageCoords srcCanvas, srcImage, canvasX, canvasY, imgX, imgY, False
        
        'Cache mouse coords so we can calculate translation amounts in future _MouseMove events
        m_InitImgX = imgX
        m_InitImgY = imgY
        m_LastImgX = imgX
        m_LastImgY = imgY
        
        'See if the user is creating a new crop, or interacting with an existing point
        If IsValidCropActive Then
            m_idxMouseDown = UpdateMousePOI(imgX, imgY)
        Else
            m_idxMouseDown = poi_Undefined
        End If
        
        'If the user is initiating a new crop, start it now
        If (m_idxMouseDown = poi_Undefined) Then
            
            'If the user wants cropping locked to image boundaries, apply that now
            ResetCropRectF imgX, imgY
            
            'Notate the bottom-right point as the current interactive target.
            m_idxMouseDown = 3
            m_numCornerCoords = GetCropCorners(m_CornerCoords)
            
        'The user is interacting with an existing crop boundary.
        ' How we modify the crop rect depends on the point being interacted with.
        Else
            
            'Cache a full list of current boundary coordinates
            m_numCornerCoords = GetCropCorners(m_CornerCoords)
            
        End If
            
    Else
        m_InitImgX = INVALID_X_COORD
        m_InitImgY = INVALID_Y_COORD
    End If
    
End Sub

Public Sub NotifyMouseMove(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByRef srcCanvas As pdCanvas, ByRef srcImage As pdImage, ByVal canvasX As Single, ByVal canvasY As Single)
    
    'Convert the current x/y position from canvas to image coordinates
    Dim imgX As Double, imgY As Double
    Drawing.ConvertCanvasCoordsToImageCoords srcCanvas, srcImage, canvasX, canvasY, imgX, imgY, False
    
    'Cache current x/y positions.
    ' (m_CropInProgress is TRUE in the midst of a double-click event, while the active crop is still being applied
    '  It is not uncommon for touchpads or mice to fire one or two superfluous WM_MOUSEMOVE messages between clicks,
    '  particularly at high sensitivity, and we want to ignore these until the _DoubleClick event completes.)
    If m_LMBDown And (Not m_CropInProgress) Then
        
        'If the user is interacting with an existing crop, modify the crop rect accordingly.
        If (m_idxMouseDown <> poi_Undefined) Then
            
            'Calculate an offset from the original click to this new location
            Dim xOffset As Long, yOffset As Long
            xOffset = imgX - m_InitImgX
            yOffset = imgY - m_InitImgY
            
            Dim i As Long, tmpCornerCoords() As PointFloat
            If (m_numCornerCoords > 0) Then ReDim tmpCornerCoords(0 To UBound(m_CornerCoords)) As PointFloat
            
            'On a "move" event, the entire crop area is moving.  Calculate an offset and move all crop points that far.
            If (m_idxMouseDown = poi_Interior) Then
                
                'Failsafe only; _MouseDown will always set m_numCornerCoords to 4
                If (m_numCornerCoords > 0) Then
                    For i = 0 To m_numCornerCoords - 1
                        tmpCornerCoords(i).x = m_CornerCoords(i).x + xOffset
                        tmpCornerCoords(i).y = m_CornerCoords(i).y + yOffset
                    Next i
                End If
                
                'If the user is not allowed to enlarge, force points in-bounds.
                ' (Note that a subsequent step - UpdateCropRectF() - will apply failsafe fixes to this as necessary.)
                If (Not m_AllowEnlarge) Then
                    
                    Dim fixOffset As Single
                    If (tmpCornerCoords(0).x < 0) Then
                        fixOffset = Abs(tmpCornerCoords(0).x)
                        For i = 0 To m_numCornerCoords - 1
                            tmpCornerCoords(i).x = tmpCornerCoords(i).x + fixOffset
                        Next i
                    End If
                    If (tmpCornerCoords(0).y < 0) Then
                        fixOffset = Abs(tmpCornerCoords(0).y)
                        For i = 0 To m_numCornerCoords - 1
                            tmpCornerCoords(i).y = tmpCornerCoords(i).y + fixOffset
                        Next i
                    End If
                    If (tmpCornerCoords(1).x > m_MaxCropWidth) Then
                        fixOffset = m_MaxCropWidth - tmpCornerCoords(1).x
                        For i = 0 To m_numCornerCoords - 1
                            tmpCornerCoords(i).x = tmpCornerCoords(i).x + fixOffset
                        Next i
                    End If
                    If (tmpCornerCoords(2).y > m_MaxCropHeight) Then
                        fixOffset = m_MaxCropHeight - tmpCornerCoords(2).y
                        For i = 0 To m_numCornerCoords - 1
                            tmpCornerCoords(i).y = tmpCornerCoords(i).y + fixOffset
                        Next i
                    End If
                End If
                
            'The user is click-dragging a specific point or edge.
            Else
                
                'Failsafe only; _MouseDown will always set m_numCornerCoords to 4
                If (m_numCornerCoords > 0) Then
                    
                    For i = 0 To m_numCornerCoords - 1
                        tmpCornerCoords(i).x = m_CornerCoords(i).x
                        tmpCornerCoords(i).y = m_CornerCoords(i).y
                    Next i
                    
                    'Check edges first; we need to manually match corner IDs for them
                    Select Case m_idxMouseDown
                    
                        Case poi_EdgeN
                            tmpCornerCoords(0).y = tmpCornerCoords(0).y + yOffset
                            tmpCornerCoords(1).y = tmpCornerCoords(1).y + yOffset
                        Case poi_EdgeE
                            tmpCornerCoords(1).x = tmpCornerCoords(1).x + xOffset
                            tmpCornerCoords(3).x = tmpCornerCoords(3).x + xOffset
                        Case poi_EdgeS
                            tmpCornerCoords(2).y = tmpCornerCoords(2).y + yOffset
                            tmpCornerCoords(3).y = tmpCornerCoords(3).y + yOffset
                        Case poi_EdgeW
                            tmpCornerCoords(0).x = tmpCornerCoords(0).x + xOffset
                            tmpCornerCoords(2).x = tmpCornerCoords(2).x + xOffset
                        
                        'Any remaining cases refer to specific corner node indices
                        Case Else
                            
                            If (m_idxMouseDown >= 0) And (m_idxMouseDown <= 3) Then
                                
                                'Add the cursor offsets to the point being interacted with
                                tmpCornerCoords(m_idxMouseDown).x = tmpCornerCoords(m_idxMouseDown).x + xOffset
                                tmpCornerCoords(m_idxMouseDown).y = tmpCornerCoords(m_idxMouseDown).y + yOffset
                                
                                'Because the point-list-to-rect function operates on max/min values,
                                ' we need to adjust adjoining corners too.
                                If (m_idxMouseDown = 0) Then
                                    tmpCornerCoords(1).y = tmpCornerCoords(1).y + yOffset
                                    tmpCornerCoords(2).x = tmpCornerCoords(2).x + xOffset
                                ElseIf (m_idxMouseDown = 1) Then
                                    tmpCornerCoords(0).y = tmpCornerCoords(0).y + yOffset
                                    tmpCornerCoords(3).x = tmpCornerCoords(3).x + xOffset
                                ElseIf (m_idxMouseDown = 2) Then
                                    tmpCornerCoords(0).x = tmpCornerCoords(0).x + xOffset
                                    tmpCornerCoords(3).y = tmpCornerCoords(3).y + yOffset
                                Else
                                    tmpCornerCoords(1).x = tmpCornerCoords(1).x + xOffset
                                    tmpCornerCoords(2).y = tmpCornerCoords(2).y + yOffset
                                End If
                                
                            End If
                            
                    End Select
                    
                End If
                
            End If
            
            'Update the crop rect to reflect any changes made to individual coordinates.
            ' (This will also handle locked aspect ratio, and "force crop in-bounds" settings.)
            UpdateCropRectF_FromPtFList tmpCornerCoords, m_idxMouseDown, imgX, imgY
            
        End If
        
    '/LMB is *not* down
    End If
    
    'Update the currently hovered crop corner, if any
    m_idxHover = UpdateMousePOI(imgX, imgY)
    
    'Finally, request a viewport redraw (to reflect any changes)
    Viewport.Stage4_FlipBufferAndDrawUI srcImage, FormMain.MainCanvas(0)
    
End Sub

Public Sub NotifyMouseUp(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByRef srcCanvas As pdCanvas, ByRef srcImage As pdImage, ByVal canvasX As Single, ByVal canvasY As Single, ByVal numOfMouseMovements As Long, ByVal clickEventAlsoFiring As Boolean)
    
    'Update cached button status now, before handling the actual event.  (This ensures that
    ' a viewport redraw, if any, will wipe all UI elements we may have rendered over the canvas.)
    If ((Button And pdLeftButton) = pdLeftButton) Then
        m_LMBDown = False
        m_idxMouseDown = poi_Undefined
    End If
    
    'Ignore the _MouseUp event that follows a double-click (by design)
    If m_CropInProgress Then
        m_CropInProgress = False
        RemoveCurrentCrop
        Exit Sub
    End If
    
    'Convert the current x/y position from canvas to image coordinates
    Dim imgX As Double, imgY As Double
    Drawing.ConvertCanvasCoordsToImageCoords srcCanvas, srcImage, canvasX, canvasY, imgX, imgY, False
        
    'In Photoshop, double-click commits the crop.  In GIMP, it's single-click.
    '
    'PD uses double-click-inside-the-rect to commit.  Single-click outside to reset.
    ' (I tried using single-click-inside to commit, but it was way too easy to accidentally apply the crop!)
    If clickEventAlsoFiring Then
        
        'Check mouse position.  If the mouse is *outside* the crop rect that existed at _MouseDown,
        ' simply clear the current rect.
        If (Not PDMath.IsPointInRectF(imgX, imgY, m_CropRectAtMouseDown)) Then
            RemoveCurrentCrop
        End If
        
    'If this is a click-drag event, we need to simply resize the crop area to match.
    ' (Commit happens in a separate step.)
    Else
        
        'If the user has modified the current crop in an unuseable way, we need to update the UI accordingly
        ' (by e.g. disabling the "commit crop" button)
        Dim cropAintGood As Boolean
        cropAintGood = False
        
        If (m_InitImgX = INVALID_X_COORD) Or (m_InitImgY = INVALID_Y_COORD) Then cropAintGood = True
        If (Not ValidateCropRectF()) Then cropAintGood = True
        If cropAintGood Then
            RemoveCurrentCrop
        
        'Finally, request a viewport redraw (to reflect any changes)
        Else
            
            'If the UI looks OK, update the viewport to match
            Viewport.Stage4_FlipBufferAndDrawUI srcImage, FormMain.MainCanvas(0)
            
        End If
        
    End If
    
End Sub

'Set new max crop bounds.  Should be called any time the active image is changed.
'
'Returns: TRUE if an active crop rectangle was modified by the newly set sizes
Public Function NotifyCropMaxSizes(ByVal newMaxWidth As Long, ByVal newMaxHeight As Long) As Boolean
    
    'Standardize un-set max width/height against 0
    If (newMaxWidth <= 0) Then m_MaxCropWidth = 0 Else m_MaxCropWidth = newMaxWidth
    If (newMaxHeight <= 0) Then m_MaxCropHeight = 0 Else m_MaxCropHeight = newMaxHeight
    
    'Whenever max values change, we need to re-assess the current crop region (to make sure it fits in-bounds)
    If (Not m_AllowEnlarge) Then NotifyCropMaxSizes = ForceCropRectInBounds()
    
End Function

Public Function GetCropAllLayers() As Boolean
    GetCropAllLayers = m_CropAllLayers
End Function

Public Sub SetCropAllLayers(ByVal newValue As Boolean)
    m_CropAllLayers = newValue
End Sub

Public Function GetCropAllowEnlarge() As Boolean
    GetCropAllowEnlarge = m_AllowEnlarge
End Function

Public Sub SetCropAllowEnlarge(ByVal newValue As Boolean)
    
    If (m_AllowEnlarge <> newValue) Then
        
        m_AllowEnlarge = newValue
        
        'After allowing enlarge behavior, we may need to force the current crop rect in-bounds.
        If m_AllowEnlarge Then
            m_MaxCropWidth = 0
            m_MaxCropHeight = 0
        Else
            ForceCropRectInBounds
        End If
        
        'Changing this toggle can cause position or size to change; relay any changes back to the UI now
        RelayCropChangesToUI
        
        'Update the viewport as necessary
        If PDImages.IsImageActive Then Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage(), FormMain.MainCanvas(0)
        
    End If
    
End Sub

'Only returns usable data when IsValidCropActive() is TRUE
Public Function GetCropRectF() As RectF
    GetCropRectF = m_CropRectF
End Function

Public Function GetCropDeletePixels() As Boolean
    GetCropDeletePixels = m_DeleteCroppedPixels
End Function

Public Sub SetCropDeletePixels(ByVal newValue As Boolean)
    m_DeleteCroppedPixels = newValue
End Sub

Public Function GetCropGuides() As Long
    GetCropGuides = m_CropGuides
End Function

Public Sub SetCropGuide(ByVal newValue As Long)
    If (m_CropGuides <> newValue) Then
        m_CropGuides = newValue
        If PDImages.IsImageActive And IsValidCropActive() Then Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    End If
End Sub

Public Function GetCropHighlight() As Boolean
    GetCropHighlight = m_HighlightCrop
End Function

Public Sub SetCropHighlight(ByVal newValue As Boolean)
    If (m_HighlightCrop <> newValue) Then
        m_HighlightCrop = newValue
        If PDImages.IsImageActive And IsValidCropActive() Then Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    End If
End Sub

Public Function GetCropHighlightColor() As Long
    GetCropHighlightColor = m_HighlightColor
End Function

Public Sub SetCropHighlightColor(ByVal newColor As Long)
    If (m_HighlightColor <> newColor) Then
        m_HighlightColor = newColor
        If PDImages.IsImageActive And IsValidCropActive() Then Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    End If
End Sub

Public Function GetCropHighlightOpacity() As Single
    GetCropHighlightOpacity = m_HighlightOpacity
End Function

Public Sub SetCropHighlightOpacity(ByVal newOpacity As Single)
    If (m_HighlightOpacity <> newOpacity) Then
        m_HighlightOpacity = newOpacity
        If PDImages.IsImageActive And IsValidCropActive() Then Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    End If
End Sub

'Called whenever the crop tool is selected
Public Sub InitializeCropTool()
    ResetCropRectF
    RelayCropChangesToUI
End Sub

'Check before using anything from GetCropRectF, above
Public Function IsValidCropActive() As Boolean
    IsValidCropActive = ValidateCropRectF()
End Function

'Some properties can be independently locked (e.g. width or height or aspect ratio).
' When a property is locked, it cannot be changed by additional UI inputs.
Public Sub LockProperty(ByVal selProperty As PD_SelectionLockable, ByVal lockedValue As Variant, Optional ByVal lockedValue2 As Variant)

    If (selProperty = pdsl_Width) Then
        m_IsWidthLocked = True
        m_LockedWidth = lockedValue
        m_IsHeightLocked = False
        m_IsAspectLocked = False
    ElseIf (selProperty = pdsl_Height) Then
        m_IsHeightLocked = True
        m_LockedHeight = lockedValue
        m_IsWidthLocked = False
        m_IsAspectLocked = False
    ElseIf (selProperty = pdsl_AspectRatio) Then
        m_IsAspectLocked = True
        m_LockedAspectNumerator = lockedValue
        m_LockedAspectDenominator = lockedValue2
        If (m_LockedAspectNumerator > 0#) And (m_LockedAspectDenominator > 0#) Then
            m_LockedAspectRatio = m_LockedAspectNumerator / m_LockedAspectDenominator
        Else
            m_LockedAspectRatio = 1#
        End If
        m_IsWidthLocked = False
        m_IsHeightLocked = False
    End If
    
End Sub

Public Sub UnlockProperty(ByVal selProperty As PD_SelectionLockable)
    If (selProperty = pdsl_Width) Then
        m_IsWidthLocked = False
    ElseIf (selProperty = pdsl_Height) Then
        m_IsHeightLocked = False
    ElseIf (selProperty = pdsl_AspectRatio) Then
        m_IsAspectLocked = False
    End If
End Sub

Public Sub ReadyForCursor(ByRef srcCanvasView As pdCanvasView)
    
    Dim idxRelevant As PD_PointOfInterest
    If m_LMBDown Then idxRelevant = m_idxMouseDown Else idxRelevant = m_idxHover
    
    Select Case idxRelevant
    
        'Mouse not over the crop
        Case poi_Undefined
            srcCanvasView.RequestCursor_System IDC_ARROW
            
        'Mouse is within the crop region, but not over a corner node
        Case poi_Interior
            If m_LMBDown Then
                If (m_idxMouseDown = poi_Interior) Then
                    srcCanvasView.RequestCursor_System IDC_SIZEALL
                Else
                    srcCanvasView.RequestCursor_System IDC_ARROW
                End If
            Else
                srcCanvasView.RequestCursor_System IDC_SIZEALL
            End If
            
        'Mouse is over one of the four crop corners
        Case 0
            srcCanvasView.RequestCursor_System IDC_SIZENWSE
        Case 1
            srcCanvasView.RequestCursor_System IDC_SIZENESW
        Case 2
            srcCanvasView.RequestCursor_System IDC_SIZENESW
        Case 3
            srcCanvasView.RequestCursor_System IDC_SIZENWSE
            
        'Mouse is over a crop edge (horizontal or vertical)
        Case poi_EdgeN
            srcCanvasView.RequestCursor_System IDC_SIZENS
        Case poi_EdgeE
            srcCanvasView.RequestCursor_System IDC_SIZEWE
        Case poi_EdgeS
            srcCanvasView.RequestCursor_System IDC_SIZENS
        Case poi_EdgeW
            srcCanvasView.RequestCursor_System IDC_SIZEWE
            
        'Should never trigger; failsafe only
        Case Else
            srcCanvasView.RequestCursor_System IDC_ARROW
        
    End Select
    
End Sub

'On _MouseMove events, update the current POI (if any)
Private Function UpdateMousePOI(ByVal imgX As Long, ByVal imgY As Long) As PD_PointOfInterest
    
    'MouseAccuracy in PD is a global value, but because we are working in image coordinates, we must compensate for the
    ' current zoom value.  (Otherwise, when zoomed out the user would be forced to work with tighter accuracy.)
    ' (TODO: come up with a better solution for this.  Accuracy should *really* be handled in the canvas coordinate space,
    '        so perhaps the caller should specify an image x/y and a radius...?)
    Dim mouseAccuracy As Double
    mouseAccuracy = Drawing.ConvertCanvasSizeToImageSize(Interface.GetStandardInteractionDistance(), PDImages.GetActiveImage)
    
    'Find the smallest distance for this mouse position
    Dim minDistance As Single
    minDistance = mouseAccuracy
    
    'Retrieve current crop corners (in image space)
    Dim cropCorners() As PointFloat, numPoints As Long
    numPoints = GetCropCorners(cropCorners)
    
    'Find the nearest point (if any) to the mouse pointer
    UpdateMousePOI = poi_Undefined
    UpdateMousePOI = PDMath.FindClosestPointInFloatArray(imgX, imgY, minDistance, cropCorners)
    
    'If the mouse is not near a corner, perform an additional check for the crop rect's interior
    If (UpdateMousePOI = poi_Undefined) And (numPoints = 4) And IsValidCropActive() Then
        
        'Next, look check the cursor against horizontal annd vertical edges of the crop rect
        Dim testPath As pd2DPath
        Set testPath = New pd2DPath
        
        Dim testPen As pd2DPen
        Set testPen = New pd2DPen
        testPen.SetPenWidth mouseAccuracy * 2
        testPen.CreatePen
        testPen.SetPenLineCap P2_LC_Flat
        
        Dim i As Long, idxOtherPoint As Long
        For i = 0 To numPoints - 1
            testPath.ResetPath
            Select Case i
                Case 0
                    idxOtherPoint = 1
                Case 1
                    idxOtherPoint = 3
                Case 2
                    idxOtherPoint = 0
                Case 3
                    idxOtherPoint = 2
            End Select
            
            'Add the line and convert it to radius [mouseAccuracy]
            testPath.AddLine cropCorners(i).x, cropCorners(i).y, cropCorners(idxOtherPoint).x, cropCorners(idxOtherPoint).y
            testPath.ConvertPath_PenTrace testPen
            
            If testPath.IsPointInsidePathF(imgX, imgY) Then
                Select Case i
                    Case 0
                        UpdateMousePOI = poi_EdgeN
                    Case 1
                        UpdateMousePOI = poi_EdgeE
                    Case 2
                        UpdateMousePOI = poi_EdgeW
                    Case 3
                        UpdateMousePOI = poi_EdgeS
                End Select
            End If
            
        Next i
        
        If (UpdateMousePOI = poi_Undefined) Then
            
            Dim tmpPath As pd2DPath
            Set tmpPath = New pd2DPath
            tmpPath.AddRectangle_RectF m_CropRectF
            
            If tmpPath.IsPointInsidePathF(imgX, imgY) Then
                UpdateMousePOI = poi_Interior
            Else
                UpdateMousePOI = poi_Undefined
            End If
            
        End If
        
    End If
    
End Function

'Get the current crop rect coordinates as a list of points (in image coordinates).
' Returns: number of points in the array, and a guaranteed ReDim to [0, [n] - 1].
Private Function GetCropCorners(ByRef dstCropPts() As PointFloat) As Long
    
    GetCropCorners = 4
    
    ReDim dstCropPts(0 To GetCropCorners - 1) As PointFloat
    dstCropPts(0).x = m_CropRectF.Left
    dstCropPts(0).y = m_CropRectF.Top
    dstCropPts(1).x = m_CropRectF.Left + m_CropRectF.Width
    dstCropPts(1).y = dstCropPts(0).y
    dstCropPts(2).x = dstCropPts(0).x
    dstCropPts(2).y = m_CropRectF.Top + m_CropRectF.Height
    dstCropPts(3).x = dstCropPts(1).x
    dstCropPts(3).y = dstCropPts(2).y
    
End Function

'You can only pass a PointFloat array sized [0, 3] to this function.  It will use that array to produce an updated
' RectF of the boundary coords of the passed list.
Private Sub UpdateCropRectF_FromPtFList(ByRef srcPoints() As PointFloat, ByVal srcPOI As PD_PointOfInterest, ByVal imgX As Double, ByVal imgY As Double)
    
    'Find the min/max points in the source point list
    Dim xMin As Single, yMin As Single, xMax As Single, yMax As Single
    xMin = srcPoints(0).x
    yMin = srcPoints(0).y
    xMax = srcPoints(0).x
    yMax = srcPoints(0).y
    
    Dim i As Long
    For i = 1 To UBound(srcPoints)
        If (srcPoints(i).x < xMin) Then xMin = srcPoints(i).x
        If (srcPoints(i).x > xMax) Then xMax = srcPoints(i).x
        If (srcPoints(i).y < yMin) Then yMin = srcPoints(i).y
        If (srcPoints(i).y > yMax) Then yMax = srcPoints(i).y
    Next i
    
    'Re-order the points in standard order (with index 0 = top-left, index 3 = bottom-right)
    srcPoints(0).x = xMin
    srcPoints(0).y = yMin
    srcPoints(1).x = xMax
    srcPoints(1).y = yMin
    srcPoints(2).x = xMin
    srcPoints(2).y = yMax
    srcPoints(3).x = xMax
    srcPoints(3).y = yMax
    
    'On a scale of [0, 3] figure out which point of the crop rect the user is *actually* interacting with.
    ' (This is important because they may e.g. drag the lower-right corner above the upper-right corner,
    ' causing the currently interactive point to switch position in the rect - and how we handle each
    ' corner is different when maintaining aspect ratio (because we anchor against the *opposite* corner.)
    Dim curDistance As Double, minDistance As Double
    minDistance = DOUBLE_MAX
    For i = 0 To 3
        curDistance = PDMath.DistanceTwoPointsShortcut(imgX, imgY, srcPoints(i).x, srcPoints(i).y)
        If (curDistance < minDistance) Then
            minDistance = curDistance
            m_idxMouseDownActual = i
        End If
    Next i
    
    'm_idxMouseDownActual now points to the srcPoints index that the user is currently interacting with.
    
    'In the event that the user is dragging a crop edge (e.g. a line segment), we also want to know
    ' the nearest line-segment of the crop boundary rect.
    m_poiEdgeMouseDownActual = poi_Undefined
    minDistance = DOUBLE_MAX
    For i = 0 To 3
        Select Case i
            Case 0
                curDistance = PDMath.DistancePerpendicular(imgX, imgY, srcPoints(0).x, srcPoints(0).y, srcPoints(1).x, srcPoints(1).y)
            Case 1
                curDistance = PDMath.DistancePerpendicular(imgX, imgY, srcPoints(1).x, srcPoints(1).y, srcPoints(3).x, srcPoints(3).y)
            Case 2
                curDistance = PDMath.DistancePerpendicular(imgX, imgY, srcPoints(0).x, srcPoints(0).y, srcPoints(2).x, srcPoints(2).y)
            Case 3
                curDistance = PDMath.DistancePerpendicular(imgX, imgY, srcPoints(2).x, srcPoints(2).y, srcPoints(3).x, srcPoints(3).y)
        End Select
        If (curDistance < minDistance) Then
            minDistance = curDistance
            m_poiEdgeMouseDownActual = i
        End If
    Next i
    
    'Translate the final "nearest edge" index into a POI constant
    If (m_poiEdgeMouseDownActual = 0) Then
        m_poiEdgeMouseDownActual = poi_EdgeN
    ElseIf (m_poiEdgeMouseDownActual = 1) Then
        m_poiEdgeMouseDownActual = poi_EdgeE
    ElseIf (m_poiEdgeMouseDownActual = 2) Then
        m_poiEdgeMouseDownActual = poi_EdgeW
    ElseIf (m_poiEdgeMouseDownActual = 3) Then
        m_poiEdgeMouseDownActual = poi_EdgeS
    End If
    
    'm_poiEdgeMouseDownActual now points to the crop edge that the user is currently interacting with.
    
    'To simplify further handling, check the "interaction POI at _MouseDown" value to determine if the user
    ' is edge-dragging or corner-dragging.
    Dim userIsCornerDragging As Boolean, userIsEdgeDragging As Boolean
    If (m_idxMouseDown >= 0) And (m_idxMouseDown <= 3) Then
        userIsCornerDragging = True
    ElseIf (m_idxMouseDown >= poi_EdgeW) And (m_idxMouseDown <= poi_EdgeN) Then
        userIsEdgeDragging = True
    End If
    
    'Next, let's consider the different types of modifications the user may be performing
    ' to the crop region.
    
    'Moving the active crop rect is the simplest modification.  (The only special case we have to
    ' consider is keeping the crop rect in-bounds, if canvas enlarging is disallowed.)
    
    'Because it's so simple, separate crop rect movement into its own case.
    Dim actionIsMoving As Boolean
    actionIsMoving = (m_idxMouseDown = poi_Interior)
    
    'We can now use the calculated max/min values to calculate a new crop boundary rect -
    ' but first, we have to apply any locked dimensions and/or aspect ratios.
    Dim tmpOverlap As RectF
    
    'When aspect ratio is locked and the user is click-dragging a corner point, this operation is actually
    ' somewhat involved.  We need to reshape the current crop rectangle to match the requested aspect ratio,
    ' while also keeping it in-bounds, anchored against the opposite corner/edge of wherever the user is dragging.
    If m_IsAspectLocked And (Not actionIsMoving) And (m_LockedAspectRatio > 0!) Then
        
        'Calculate the rect's current width/height
        Dim newWidth As Single, newHeight As Single
        newWidth = xMax - xMin
        If (newWidth < 1!) Then newWidth = 1!
        newHeight = yMax - yMin
        If (newHeight < 1!) Then newHeight = 1!
        
        'Split handling by edge-dragging vs corner-dragging.  When edge-dragging, we *always* want to leave
        ' the dimension being dragged as-is (and adjust the *opposite* dimension for aspect-ratio).
        If userIsEdgeDragging Then
            
            If (m_poiEdgeMouseDownActual = poi_EdgeE) Or (m_poiEdgeMouseDownActual = poi_EdgeW) Then
                newHeight = newWidth * (1# / m_LockedAspectRatio)
            Else
                newWidth = newHeight * m_LockedAspectRatio
            End If
            
        Else
            
            'Width > height (so adjust height as necessary)
            If (m_LockedAspectRatio >= 1!) Then
                newHeight = newWidth * (1# / m_LockedAspectRatio)
                
            'Height > width (so adjust width as necessary)
            Else
                newWidth = newHeight * m_LockedAspectRatio
            End If
            
        End If
        
        'newWidth and newHeight now represent the current crop rectangle, corrected for aspect ratio.
        
        'If the caller is allowed to enlarge the crop area, simply use the new width and height values as-is
        ' (because we don't care if they exceed image boundaries).
        With m_CropRectF
            .Width = newWidth
            .Height = newHeight
            
            'Anchor the resize against the opposite point of the active interaction node
            If userIsCornerDragging Then
                
                Select Case m_idxMouseDownActual
                    Case 0
                        .Left = xMax - newWidth
                        .Top = yMax - newHeight
                    Case 1
                        .Left = xMin
                        .Top = yMax - newHeight
                    Case 2
                        .Left = xMax - newWidth
                        .Top = yMin
                    Case 3
                        .Left = xMin
                        .Top = yMin
                End Select
                
            ElseIf userIsEdgeDragging Then
                
                Select Case m_poiEdgeMouseDownActual
                    Case poi_EdgeN
                        .Left = xMin
                        .Top = yMax - newHeight
                    Case poi_EdgeW
                        .Left = xMax - newWidth
                        .Top = yMin
                    Case Else
                        .Left = xMin
                        .Top = yMin
                End Select
                
            End If
            
        End With
        
        'If we are *not* allowed to enlarge the image, we now need to make sure our calculated rect
        ' actually fits within the image!
        If (Not m_AllowEnlarge) Then
            
            'We have calculated a new width/height for the image, but unfortunately, we don't know
            ' if our new sizes still fit within image boundaries.
            
            'Before doing anything else, try using the sizes as-is.  (If they work, great!)
            PDMath.GetIntClampedRectF m_CropRectF
            GDI_Plus.IntersectRectF tmpOverlap, m_CropRectF, PDImages.GetActiveImage.GetBoundaryRectF
            PDMath.GetIntClampedRectF tmpOverlap
            
            If (Not VBHacks.MemCmp(VarPtr(m_CropRectF), VarPtr(tmpOverlap), 16)) Then
                
                'Argh, our newly calculated rectangle doesn't work.  We need to calculate a new
                ' aspect-ratio-preserved rectangle that actually fits within image boundaries.
                
                'Calculate what happens when we shrink either width or height (because based on anchor positioning,
                ' there's no guarantee which direction we should shrink to keep the rect in-bounds).
                Dim testRectF As RectF, testRectF2 As RectF
                testRectF = tmpOverlap
                
                Dim areaOne As Single
                testRectF.Height = testRectF.Width * (1# / m_LockedAspectRatio)
                areaOne = testRectF.Width * testRectF.Height
                
                testRectF2 = tmpOverlap
                testRectF2.Width = testRectF2.Height * m_LockedAspectRatio
                
                'Take the smaller of the two areas and save it to testRectF
                If ((testRectF2.Width * testRectF2.Height) < areaOne) Then testRectF = testRectF2
                
                '...then use that as the basis for m_CropRectF
                m_CropRectF = testRectF
                
                'Finally, anchor the modified rectangle against the *opposite* point/edge the user is interacting with
                If userIsCornerDragging Then
                    
                    Select Case m_idxMouseDownActual
                        Case 0
                            m_CropRectF.Left = (tmpOverlap.Left + tmpOverlap.Width) - testRectF.Width
                            m_CropRectF.Top = (tmpOverlap.Top + tmpOverlap.Height) - testRectF.Height
                        Case 1
                            m_CropRectF.Top = (tmpOverlap.Top + tmpOverlap.Height) - testRectF.Height
                        Case 2
                            m_CropRectF.Left = (tmpOverlap.Left + tmpOverlap.Width) - testRectF.Width
                        Case 3
                    End Select
                    
                ElseIf userIsEdgeDragging Then
                    
                    Select Case m_poiEdgeMouseDownActual
                        Case poi_EdgeN
                            m_CropRectF.Top = (tmpOverlap.Top + tmpOverlap.Height) - testRectF.Height
                        Case poi_EdgeS
                        Case poi_EdgeE
                        Case poi_EdgeW
                            m_CropRectF.Left = (tmpOverlap.Left + tmpOverlap.Width) - testRectF.Width
                    End Select
                    
                End If
                
            End If
            
        End If
        
    'When aspect ratio is not locked, this step is easy: just use max/min values as calculated above
    Else
        
        With m_CropRectF
            If actionIsMoving Or (Not m_IsWidthLocked) Then .Left = xMin
            If (Not m_IsWidthLocked) Then .Width = xMax - xMin
            If actionIsMoving Or (Not m_IsHeightLocked) Then .Top = yMin
            If (Not m_IsHeightLocked) Then .Height = yMax - yMin
        End With
        
    End If
    
    'If the user wants the crop clamped to image boundaries, calculate a final, failsafe overlap now.
    ' (Nothing should change as a result of the above calculations, but better safe than sorry.)
    If (Not m_AllowEnlarge) Then
        GDI_Plus.IntersectRectF tmpOverlap, m_CropRectF, PDImages.GetActiveImage.GetBoundaryRectF
        m_CropRectF = tmpOverlap
    End If
    
    'Lock position and height to their nearest integer equivalent
    PDMath.GetIntClampedRectF m_CropRectF
    
    'If the crop rect is valid, relay its values to the toolpanel
    If ValidateCropRectF() Then RelayCropChangesToUI
    
End Sub

'If the crop is not allowed to enlarge the image, any crop changes must be forced in-bounds.
' This is best handled in individual size modification steps (because it makes behavior more intuitive
' for the user), but when toggling the "allow enlarge" setting, we need to forcibly bring the current
' crop rect (if any) into bounds - this function will do that.
'
'Returns: TRUE if the crop rect was changed as a result of forcing in-bounds
Private Function ForceCropRectInBounds() As Boolean
    
    ForceCropRectInBounds = False
    
    'If the user is not currently enforcing boundaries, do nothing
    If m_AllowEnlarge Then Exit Function
    
    'The user is enforcing boundaries.  They *must* have set a max width/height in a previous step.
    If (m_MaxCropWidth = 0) Or (m_MaxCropHeight = 0) Then Exit Function
    
    Dim widthOrHeightChanged As Boolean
    widthOrHeightChanged = False
    
    'There's a nebulous interplay when the user locks width/height but those dimensions are too big.
    ' When this happens, ignore the locked width/height values as necessary to enforce the max width/height.
    If m_IsWidthLocked Then
        If (m_CropRectF.Width > m_MaxCropWidth) Then
            m_CropRectF.Width = m_MaxCropWidth
            widthOrHeightChanged = True
        End If
        If (m_CropRectF.Height > m_MaxCropHeight) Then
            m_CropRectF.Height = m_MaxCropHeight
            widthOrHeightChanged = True
        End If
    End If
    
    'Now we get the fun of trying to make sure the crop rect "fits" within the boundaries set by the user.
    
    'Start with left/top position, which are the easiest to handle.
    If (m_CropRectF.Left < 0) Then m_CropRectF.Left = 0
    If (m_CropRectF.Top < 0) Then m_CropRectF.Top = 0
    
    'Next, reconcile width
    If (m_CropRectF.Left + m_CropRectF.Width > m_MaxCropWidth) Then
        
        'First, attempt to bring the crop "in-bounds" by shifting it left.
        m_CropRectF.Left = m_MaxCropWidth - m_CropRectF.Width
        
        'If that isn't enough to "fix" the crop, simply limit it to max size
        If (m_CropRectF.Left < 0) Then
            m_CropRectF.Left = 0
            m_CropRectF.Width = m_MaxCropWidth
            widthOrHeightChanged = True
        End If
        
    End If
    
    'Repeat the above steps, but for height
    If (m_CropRectF.Top + m_CropRectF.Height > m_MaxCropHeight) Then
        
        'First, attempt to bring the crop "in-bounds" by shifting it up.
        m_CropRectF.Top = m_MaxCropHeight - m_CropRectF.Height
        
        'If that isn't enough to "fix" the crop, simply limit it to max size
        If (m_CropRectF.Top < 0) Then
            m_CropRectF.Top = 0
            m_CropRectF.Height = m_MaxCropHeight
            widthOrHeightChanged = True
        End If
        
    End If
    
    'This function can only *shrink* width/height, not enlarge it - but that leaves us with the problem of aspect ratio.
    ' If the above changes messed up a locked aspect ratio, try to preserve it.
    If m_IsAspectLocked And (m_LockedAspectRatio <> 0#) And widthOrHeightChanged Then
        
        'We were forced to change width/height to bring the crop region in-bounds.
        ' Try to fix the problem by adjusting the problematic dimensions.
        If (m_LockedAspectRatio >= 1#) Then
            m_CropRectF.Height = m_CropRectF.Width * (1# / m_LockedAspectRatio)
            If (m_CropRectF.Height > m_MaxCropHeight) Then
                m_CropRectF.Top = m_MaxCropHeight - m_CropRectF.Height
                If (m_CropRectF.Top < 0) Then
                    m_CropRectF.Top = 0
                    m_CropRectF.Height = m_MaxCropHeight
                    widthOrHeightChanged = True
                End If
            End If
        Else
            m_CropRectF.Width = m_CropRectF.Height * m_LockedAspectRatio
            If (m_CropRectF.Width > m_MaxCropWidth) Then
                m_CropRectF.Left = m_MaxCropWidth - m_CropRectF.Width
                If (m_CropRectF.Left < 0) Then
                    m_CropRectF.Left = 0
                    m_CropRectF.Width = m_MaxCropWidth
                    widthOrHeightChanged = True
                End If
            End If
        End If
        
    End If
    
    ForceCropRectInBounds = widthOrHeightChanged
    
End Function

Private Sub ResetCropRectF(Optional ByVal initX As Single = 0!, Optional ByVal initY As Single = 0!)
    With m_CropRectF
        .Left = initX
        .Top = initY
        .Width = 0!
        .Height = 0!
    End With
End Sub

'Returns TRUE if the current crop rect is valid; if FALSE, do *not* attempt to crop
Private Function ValidateCropRectF() As Boolean
    
    ValidateCropRectF = True
    
    'Validate left/right/width/height
    With m_CropRectF
        If (.Left = INVALID_X_COORD) Or (.Top = INVALID_Y_COORD) Then ValidateCropRectF = False
        If (.Width <= 0!) Or (.Height <= 0!) Then ValidateCropRectF = False
    End With
    
    'Next, ensure the crop rect actually overlaps the image *somewhere*.
    If PDImages.IsImageActive() Then
        Dim dummyIntersectRectF As RectF
        If (Not GDI_Plus.IntersectRectF(dummyIntersectRectF, m_CropRectF, PDImages.GetActiveImage.GetBoundaryRectF)) Then ValidateCropRectF = False
    End If
    
End Function

Private Sub RelayCropChangesToUI()
    
    'Lock updates to prevent circular references
    Tools.SetToolBusyState True
    
    toolpanel_Crop.tudCrop(0).Value = m_CropRectF.Left
    toolpanel_Crop.tudCrop(1).Value = m_CropRectF.Top
    
    toolpanel_Crop.tudCrop(2).Value = m_CropRectF.Width
    toolpanel_Crop.tudCrop(3).Value = m_CropRectF.Height
    
    'Aspect ratio requires a bit of math
    Dim fracNumerator As Long, fracDenominator As Long
    GetAspectRatioAsFraction fracNumerator, fracDenominator
    If m_IsAspectLocked Then
        toolpanel_Crop.tudCrop(4).Value = m_LockedAspectNumerator
        toolpanel_Crop.tudCrop(5).Value = m_LockedAspectDenominator
    Else
        toolpanel_Crop.tudCrop(4).Value = fracNumerator
        toolpanel_Crop.tudCrop(5).Value = fracDenominator
    End If
    
    toolpanel_Crop.cmdCommit(0).Enabled = Tools_Crop.IsValidCropActive()
    toolpanel_Crop.cmdCommit(1).Enabled = Tools_Crop.IsValidCropActive()
    
    'Unlock updates
    Tools.SetToolBusyState False
    
End Sub

'The crop toolpanel relays user input via this function.  Left/top/width/height as passed as integers; aspect ratio as a float.
Public Sub RelayCropChangesFromUI(ByVal changedProperty As PD_Dimension, Optional ByVal newPropI As Long = 0, Optional ByVal newPropF As Single = 0!, Optional ByVal newPropI2 As Long = 0&)
    
    'For calculating aspect ratio as a fraction
    Dim fracNumerator As Long, fracDenominator As Long
    
    Select Case changedProperty
        
        Case pdd_Left
            m_CropRectF.Left = newPropI
            
        Case pdd_Top
            m_CropRectF.Top = newPropI
        
        'Changing width/height requires us to update aspect ratio to match
        Case pdd_Width, pdd_Height
            
            'If the current aspect ratio is locked, we need to update the dimension passed to its exact value,
            ' then modify the opposite dimension to match the locked aspect ratio.
            If m_IsAspectLocked And (m_LockedAspectRatio > 0#) Then
                
                'User is changing width
                If (changedProperty = pdd_Width) Then
                    If toolpanel_Crop.tudCrop(2).IsValid(False) Then
                        m_CropRectF.Width = newPropI
                        m_CropRectF.Height = m_CropRectF.Width * (1# / m_LockedAspectRatio)
                        toolpanel_Crop.tudCrop(3).Value = m_CropRectF.Height
                    End If
                    
                'User is changing height
                Else
                    If toolpanel_Crop.tudCrop(3).IsValid(False) Then
                        m_CropRectF.Height = newPropI
                        m_CropRectF.Width = m_CropRectF.Height * m_LockedAspectRatio
                        toolpanel_Crop.tudCrop(2).Value = m_CropRectF.Width
                    End If
                End If
            
            'If aspect ratio is *not* locked, freely modify either value, and cache any locked values
            ' so they can be used elsewhere.
            Else
                
                If (changedProperty = pdd_Width) Then
                    m_CropRectF.Width = newPropI
                    If m_IsWidthLocked Then m_LockedWidth = m_CropRectF.Width
                Else
                    m_CropRectF.Height = newPropI
                    If m_IsHeightLocked Then m_LockedHeight = m_CropRectF.Height
                End If
                
                GetAspectRatioAsFraction fracNumerator, fracDenominator
                toolpanel_Crop.tudCrop(4).Value = fracNumerator
                toolpanel_Crop.tudCrop(5).Value = fracDenominator
                
            End If
        
        'When changing aspect ratio, the UI passes both the new aspect ratio, and the current size of
        ' the *opposite* dimension.
        Case pdd_AspectRatioW
            m_CropRectF.Width = newPropI
            toolpanel_Crop.tudCrop(2).Value = newPropI
            If m_IsAspectLocked Then
                m_LockedAspectNumerator = newPropF
                m_LockedAspectDenominator = newPropI2
                If (m_LockedAspectNumerator > 0) And (m_LockedAspectDenominator > 0) Then
                    m_LockedAspectRatio = m_LockedAspectNumerator / m_LockedAspectDenominator
                Else
                    m_LockedAspectRatio = 1!
                End If
            End If
            
        Case pdd_AspectRatioH
            m_CropRectF.Height = newPropI
            toolpanel_Crop.tudCrop(3).Value = newPropI
            If m_IsAspectLocked Then
                m_LockedAspectNumerator = newPropF
                m_LockedAspectDenominator = newPropI2
                If (m_LockedAspectNumerator > 0) And (m_LockedAspectDenominator > 0) Then
                    m_LockedAspectRatio = m_LockedAspectNumerator / m_LockedAspectDenominator
                Else
                    m_LockedAspectRatio = 1!
                End If
            End If
        
        'Swap aspect ratio (e.g. 4:3 becomes 3:4).  This toggle ignores any/all locked values.
        Case pdd_SwapAspectRatio
            
            Tools.SetToolBusyState True
            
            newPropF = m_CropRectF.Width
            m_CropRectF.Width = m_CropRectF.Height
            m_CropRectF.Height = newPropF
            
            'If the crop is currently constrained to image dimensions, shrink it as necessary
            If (Not m_AllowEnlarge) Then
                
                'Width, then height
                Dim allowedSize As Double, reduceFactor As Double
                If (m_CropRectF.Left + m_CropRectF.Width > m_MaxCropWidth) Then
                    allowedSize = m_MaxCropWidth - m_CropRectF.Left
                    reduceFactor = allowedSize / m_CropRectF.Width
                    m_CropRectF.Width = allowedSize
                    m_CropRectF.Height = m_CropRectF.Height * reduceFactor
                    If (m_CropRectF.Height < 1!) Then m_CropRectF.Height = 1!
                End If
                
                If (m_CropRectF.Top + m_CropRectF.Height > m_MaxCropHeight) Then
                    allowedSize = m_MaxCropHeight - m_CropRectF.Top
                    reduceFactor = allowedSize / m_CropRectF.Height
                    m_CropRectF.Height = allowedSize
                    m_CropRectF.Width = m_CropRectF.Width * reduceFactor
                    If (m_CropRectF.Width < 1!) Then m_CropRectF.Width = 1!
                End If
                
            End If
            
            'Now we have to relay changes to a bunch of places: width/height/aspect ratio
            toolpanel_Crop.tudCrop(2).Value = m_CropRectF.Width
            toolpanel_Crop.tudCrop(3).Value = m_CropRectF.Height
            
            Dim tmpDouble As Double
            If m_IsAspectLocked And (m_CropRectF.Height >= 1!) Then
                
                'Swap the stored numerator and denominator values
                tmpDouble = m_LockedAspectDenominator
                m_LockedAspectDenominator = m_LockedAspectNumerator
                m_LockedAspectNumerator = tmpDouble
                
                'Calculate a new locked aspect ratio, then relay the changes back to the UI
                m_LockedAspectRatio = m_CropRectF.Width / m_CropRectF.Height
            
                toolpanel_Crop.tudCrop(4).Value = m_LockedAspectNumerator
                toolpanel_Crop.tudCrop(5).Value = m_LockedAspectDenominator
            
            'Simply swap the text box values as-is
            Else
                tmpDouble = toolpanel_Crop.tudCrop(4).Value
                toolpanel_Crop.tudCrop(4).Value = toolpanel_Crop.tudCrop(5).Value
                toolpanel_Crop.tudCrop(5).Value = Int(tmpDouble)
            End If
            
            Tools.SetToolBusyState False
            
        'When neither width nor height is locked, the UI will pass aspect ratio as separate width/height values;
        ' we need to ensure those stay the same while also keeping the active crop rectangle (if any) in-bounds.
        Case pdd_AspectBoth
            
            'If aspect ratio is locked, store the updated value now
            If (m_IsAspectLocked And (newPropF > 0!)) Then
                m_LockedAspectRatio = CDbl(newPropI) / newPropF
                m_LockedAspectNumerator = newPropI
                m_LockedAspectDenominator = newPropF
            End If
            
            If Tools_Crop.IsValidCropActive() Then
                
                'A crop is active.  While preserving aspect ratio, calculate new width/height for the current crop
                ' (while also keeping it in-bounds, as necessary).
                Tools.SetToolBusyState True
                
                Dim tmpRatio As Double
                If (newPropF > 0!) Then
                    tmpRatio = CDbl(newPropI) / newPropF
                    m_CropRectF.Width = m_CropRectF.Height * tmpRatio
                End If
                
                'We may now need to keep the crop in-bounds
                If (Not m_AllowEnlarge) Then
                    
                    If (m_CropRectF.Left + m_CropRectF.Width > m_MaxCropWidth) Then
                        allowedSize = m_MaxCropWidth - m_CropRectF.Left
                        reduceFactor = allowedSize / m_CropRectF.Width
                        m_CropRectF.Width = allowedSize
                        m_CropRectF.Height = m_CropRectF.Height * reduceFactor
                        If (m_CropRectF.Height < 1!) Then m_CropRectF.Height = 1!
                    End If
                    
                    If (m_CropRectF.Top + m_CropRectF.Height > m_MaxCropHeight) Then
                        allowedSize = m_MaxCropHeight - m_CropRectF.Top
                        reduceFactor = allowedSize / m_CropRectF.Height
                        m_CropRectF.Height = allowedSize
                        m_CropRectF.Width = m_CropRectF.Width * reduceFactor
                        If (m_CropRectF.Width < 1!) Then m_CropRectF.Width = 1!
                    End If
                    
                End If
                
                'Now we have to relay changes to a bunch of places: width/height/aspect ratio
                toolpanel_Crop.tudCrop(2).Value = m_CropRectF.Width
                toolpanel_Crop.tudCrop(3).Value = m_CropRectF.Height
                If m_IsAspectLocked And (m_CropRectF.Height >= 1!) Then m_LockedAspectRatio = m_CropRectF.Width / m_CropRectF.Height
                
                'Don't relay aspect ratio; we want to leave the current UI values untouched
                'fracNumerator = toolpanel_Crop.tudCrop(4).Value
                'toolpanel_Crop.tudCrop(4).Value = toolpanel_Crop.tudCrop(5).Value
                'toolpanel_Crop.tudCrop(5).Value = fracNumerator
                    
                Tools.SetToolBusyState False
                
            'No crop is active; do nothing
            'Else
            End If
            
    End Select
    
    'After a property change, we must validate the crop rectangle to make sure no funny business occurred
    If ValidateCropRectF() Then
        
        'If the crop region looks OK *and* the user doesn't want to allow enlarging, perform a final
        ' forcible check to bring the region "back" in-bounds as necessary.
        If (Not m_AllowEnlarge) Then
            ForceCropRectInBounds
            RelayCropChangesToUI
        End If
        
    End If
    
    'We don't want to relay *all* crop settings back to the UI (because this function is meant to
    ' relay values *from* the UI), but we may need to adjust the apply/commit buttons if the user's
    ' changes changed crop validity state.
    toolpanel_Crop.cmdCommit(0).Enabled = Tools_Crop.IsValidCropActive()
    toolpanel_Crop.cmdCommit(1).Enabled = Tools_Crop.IsValidCropActive()
    
    '...then request a viewport redraw to reflect any changes
    Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

'Get the current aspect ratio of m_CropRectF, split into integer numerator and denominator values
Private Sub GetAspectRatioAsFraction(ByRef dstNumerator As Long, ByRef dstDenominator As Long)

    'Changes to width/height (with unlocked aspect ratio) will obviously change the
    ' current aspect ratio.  Re-calculate aspect ratio and update those spinners accordingly.
    If (m_CropRectF.Width >= 1!) And (m_CropRectF.Height >= 1!) Then
        
        PDMath.ConvertToFraction m_CropRectF.Width / m_CropRectF.Height, dstNumerator, dstDenominator, 0.005
        
        'Aspect ratios are typically given in terms of base 10 if possible, so change values like 8:5 to 16:10
        If (dstDenominator = 5) Or (dstNumerator = 5) Then
            dstNumerator = dstNumerator * 2
            dstDenominator = dstDenominator * 2
        End If
    
    Else
        dstNumerator = 1
        dstDenominator = 1
    End If
        
End Sub

'Pass crop-tool-specific errors here
Private Sub InternalError(ByRef funcName As String, ByRef srcErrMsg As String)
    PDDebug.LogAction "WARNING!  Tools_Crop module error in " & funcName & ": " & srcErrMsg
End Sub
