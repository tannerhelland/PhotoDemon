Attribute VB_Name = "Tools_Crop"
'***************************************************************************
'Crop tool interface
'Copyright 2024-2025 by Tanner Helland
'Created: 12/November/24
'Last updated: 06/February/25
'Last update: add aspect ratio swap, for converting between landscape and portrait modes
'
'The crop tool performs identical operations to the rectangular selection tool + Image > Crop menu.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'TRUE if _MouseDown event was received, and _MouseUp has not yet arrived
Private m_LMBDown As Boolean

'Populated in _MouseDown
Private Const INVALID_X_COORD As Double = DOUBLE_MAX, INVALID_Y_COORD As Double = DOUBLE_MAX
Private m_InitImgX As Double, m_InitImgY As Double

'Populate in _MouseMove
Private m_LastImgX As Double, m_LastImgY As Double

'Current crop rect, always *in image coordinates*
Private m_CropRectF As RectF

'Some crop attributes can be independently locked (e.g. aspect ratio).
Private m_IsWidthLocked As Boolean, m_IsHeightLocked As Boolean, m_IsAspectLocked As Boolean
Private m_LockedWidth As Long, m_LockedHeight As Long, m_LockedAspectRatio As Double

'Index of the current mouse_over point, if any.
Private m_idxHover As PD_PointOfInterest

'On _MouseDown events, this will be set to a specific POI if the user is interacting with that POI.
' (If the user is creating a new Crop from scratch, this will be poi_Undefined.)
Private m_idxMouseDown As PD_PointOfInterest

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
        End With
        
        Crop_ApplyCropRect targetCropRectF
    
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
    
    'Now it's time to render the crop outline and any on-canvas interactive UI bits
    
    'We now want to add all lines to a path, which we'll render all at once
    Dim cropOutline As pd2DPath
    Set cropOutline = New pd2DPath
    cropOutline.AddRectangle_RectF cropRectCanvas
    
    Dim drawActiveOutline As Boolean
    drawActiveOutline = (m_idxHover = poi_Interior)
    If m_LMBDown Then drawActiveOutline = drawActiveOutline And (m_idxMouseDown = poi_Interior)
    
    'Stroke the outline path
    If drawActiveOutline Then
        PD2D.DrawPath cSurface, basePenActive, cropOutline
        PD2D.DrawPath cSurface, topPenActive, cropOutline
    Else
        PD2D.DrawPath cSurface, basePenInactive, cropOutline
        PD2D.DrawPath cSurface, topPenInactive, cropOutline
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
    If m_LMBDown Then idxActive = m_idxMouseDown Else idxActive = m_idxHover
    
    For i = 0 To numPoints - 1
        If (i = idxActive) Then
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
        
        'TODO: add module flag here, and ignore other mouse events when flag is set (clear flag at commit end)
        m_CropInProgress = True
        
        'Check mouse position.  If the mouse is *outside* the crop rect that existed at _MouseDown,
        ' simply clear the current rect.
        If PDMath.IsPointInRectF(imgX, imgY, m_CropRectAtMouseDown) Then
            
            'Set a module-level flag so that the subsequent _MouseUp doesn't fire
            m_CropInProgress = True
            CommitCurrentCrop
            ResetCropRectF 'Failsafe because sometimes laptop touchpads generate superfluous mouse events
            
            'm_CropInProgress can only be reset by a subsequent _MouseDown; do not do anything with it here.
            
        End If
        
    End If
    
End Sub

Public Sub NotifyMouseDown(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByRef srcCanvas As pdCanvas, ByRef srcImage As pdImage, ByVal canvasX As Single, ByVal canvasY As Single)
    
    'Reset the "in the midst of a double-click" flag
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
        m_idxMouseDown = UpdateMousePOI(imgX, imgY)
        
        'If the user is initiating a new crop, start it now
        If (m_idxMouseDown = poi_Undefined) Then
            
            'If the user wants cropping locked to image boundaries, apply that now
            ResetCropRectF imgX, imgY
            
        'The user is interacting with an existing crop boundary.
        ' How we modify the crop rect depends on the point being interacted with.
        Else
            
            'Cache a full list of boundary coordinates
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
    
    'Cache current x/y positions
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
                
            'The user is click-dragging a specific point.
            Else
                
                'Failsafe only; _MouseDown will always set m_numCornerCoords to 4
                If (m_numCornerCoords > 0) Then
                    
                    For i = 0 To m_numCornerCoords - 1
                        tmpCornerCoords(i).x = m_CornerCoords(i).x
                        tmpCornerCoords(i).y = m_CornerCoords(i).y
                        If (i = m_idxMouseDown) Then
                            tmpCornerCoords(i).x = tmpCornerCoords(i).x + xOffset
                            tmpCornerCoords(i).y = tmpCornerCoords(i).y + yOffset
                        End If
                    Next i
                    
                    'Because the point-list-to-rect function operates on max/min values, we need to adjust adjoining
                    ' corners too.
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
                
            End If
            
            'Update the crop rect to reflect any changes made to individual coordinates
            UpdateCropRectF_FromPtFList tmpCornerCoords, (m_idxMouseDown = poi_Interior), m_idxMouseDown
            
        '...otherwise, simply create a new crop to match the mouse movement
        Else
            
            'Move coordinates around to ensure positive width/height
            UpdateCropRectF imgX, imgY
            m_LastImgX = imgX
            m_LastImgY = imgY
            
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
    If ((Button And pdLeftButton) = pdLeftButton) Then m_LMBDown = False
    
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
    'I'm still debating which way to go in PD, but I'm currently leaning toward GIMP as it's also the
    ' traditional PD convention (click-outside-to-cancel, click-inside-to-apply).
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
            
        'If the user is interacting with an existing crop, modify the crop rect accordingly.
        If (m_idxMouseDown <> poi_Undefined) Then
        
        '...otherwise, simply create a new crop to match the mouse movement
        Else
                
            'Update the crop rectangle against these (final) coordinates
            UpdateCropRectF imgX, imgY
            m_LastImgX = imgX
            m_LastImgY = imgY
            
        End If
        
        'Bail if any of the final coordinates produce an unuseable crop rect
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

Public Function GetCropHighlight() As Boolean
    GetCropHighlight = m_HighlightCrop
End Function

Public Sub SetCropHighlight(ByVal newValue As Boolean)
    m_HighlightCrop = newValue
    If PDImages.IsImageActive And IsValidCropActive() Then Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage(), FormMain.MainCanvas(0)
End Sub

Public Function GetCropHighlightColor() As Long
    GetCropHighlightColor = m_HighlightColor
End Function

Public Sub SetCropHighlightColor(ByVal newColor As Long)
    m_HighlightColor = newColor
    If PDImages.IsImageActive And IsValidCropActive() Then Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage(), FormMain.MainCanvas(0)
End Sub

Public Function GetCropHighlightOpacity() As Single
    GetCropHighlightOpacity = m_HighlightOpacity
End Function

Public Sub SetCropHighlightOpacity(ByVal newOpacity As Single)
    m_HighlightOpacity = newOpacity
    If PDImages.IsImageActive And IsValidCropActive() Then Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage(), FormMain.MainCanvas(0)
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
Public Sub LockProperty(ByVal selProperty As PD_SelectionLockable, ByVal lockedValue As Variant)

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
        m_LockedAspectRatio = lockedValue
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

Private Sub UpdateCropRectF(ByVal newX As Single, ByVal newY As Single)
    
    Dim newHeight As Single
    If m_IsHeightLocked Then
        newHeight = m_LockedHeight
    Else
        newHeight = Abs(m_InitImgY - newY)
        If (newHeight < 1!) Then newHeight = 1!
    End If
    
    Dim newWidth As Single
    If m_IsWidthLocked Then
        newWidth = m_LockedWidth
    Else
        If m_IsAspectLocked Then
            newWidth = newHeight * m_LockedAspectRatio
        Else
            newWidth = Abs(m_InitImgX - newX)
        End If
        If (newWidth < 1!) Then newWidth = 1!
    End If
    
    With m_CropRectF
        
        .Width = newWidth
        
        If (newX < m_InitImgX) Then
            If (m_IsWidthLocked Or m_IsAspectLocked) Then
                .Left = m_InitImgX - newWidth
            Else
                .Left = newX
            End If
        Else
            .Left = m_InitImgX
        End If
        
        .Height = newHeight
        
        If (newY < m_InitImgY) Then
            If (m_IsHeightLocked Or m_IsAspectLocked) Then
                .Top = m_InitImgY - newHeight
            Else
                .Top = newY
            End If
        Else
            .Top = m_InitImgY
        End If
        
    End With
    
    'If the user wants the crop clamped to image boundaries, calculate overlap now.
    If (Not m_AllowEnlarge) Then
        Dim tmpOverlap As RectF
        GDI_Plus.IntersectRectF tmpOverlap, m_CropRectF, PDImages.GetActiveImage.GetBoundaryRectF
        m_CropRectF = tmpOverlap
    End If
    
    'Lock position and height to their nearest integer equivalent
    PDMath.GetIntClampedRectF m_CropRectF
    
    'If the crop rect is valid, relay its values to the toolpanel
    If ValidateCropRectF() Then RelayCropChangesToUI
    
End Sub

'You can only pass a PointFloat array sized [0, 3] to this function.  It will use that array to produce an updated
' RectF of the boundary coords of the passed list.
Private Sub UpdateCropRectF_FromPtFList(ByRef srcPoints() As PointFloat, Optional ByVal okToMove As Boolean = False, Optional ByVal srcPOI As PD_PointOfInterest = poi_Undefined)
    
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
    
    'We can now use the calculated max/min values to calculate a boundary rect (but note that we must
    ' also consider any locked dimensions and/or aspect ratio).
    If m_IsAspectLocked And (Not okToMove) Then
        
        'When aspect ratio is locked, we need to manually calculate new width/height values
        Dim newWidth As Single, newHeight As Single
        newHeight = yMax - yMin
        If (newHeight < 1!) Then newHeight = 1!
        newWidth = Int(newHeight * m_LockedAspectRatio + 0.5!)
        If (newWidth < 1!) Then newWidth = 1!
        m_CropRectF.Width = newWidth
        m_CropRectF.Height = newHeight
        
        'We now need to position the selection to either the left or right, depending on the current POI
        If (srcPOI = 0) Then
            If (srcPoints(0).x < srcPoints(1).x) Then
                m_CropRectF.Left = srcPoints(1).x - newWidth
            Else
                m_CropRectF.Left = xMin
            End If
            If (srcPoints(0).y < srcPoints(2).y) Then
                m_CropRectF.Top = srcPoints(2).y - newHeight
            Else
                m_CropRectF.Top = yMin
            End If
        ElseIf (srcPOI = 1) Then
            If (srcPoints(1).x < srcPoints(0).x) Then
                m_CropRectF.Left = srcPoints(0).x - newWidth
            Else
                m_CropRectF.Left = xMin
            End If
            If (srcPoints(0).y < srcPoints(3).y) Then
                m_CropRectF.Top = srcPoints(3).y - newHeight
            Else
                m_CropRectF.Top = yMin
            End If
        ElseIf (srcPOI = 2) Then
            If (srcPoints(2).x < srcPoints(3).x) Then
                m_CropRectF.Left = srcPoints(3).x - newWidth
            Else
                m_CropRectF.Left = xMin
            End If
            If (srcPoints(2).y < srcPoints(0).y) Then
                m_CropRectF.Top = srcPoints(0).y - newHeight
            Else
                m_CropRectF.Top = yMin
            End If
        ElseIf (srcPOI = 3) Then
            If (srcPoints(3).x < srcPoints(2).x) Then
                m_CropRectF.Left = srcPoints(2).x - newWidth
            Else
                m_CropRectF.Left = xMin
            End If
            If (srcPoints(3).y < srcPoints(1).y) Then
                m_CropRectF.Top = srcPoints(1).y - newHeight
            Else
                m_CropRectF.Top = yMin
            End If
        End If
    
    'When aspect ratio is not locked, this step is easy: just use max/min values as calculated above
    Else
        
        With m_CropRectF
            If okToMove Or (Not m_IsWidthLocked) Then .Left = xMin
            If (Not m_IsWidthLocked) Then .Width = xMax - xMin
            If okToMove Or (Not m_IsHeightLocked) Then .Top = yMin
            If (Not m_IsHeightLocked) Then .Height = yMax - yMin
        End With
        
    End If
    
    'If the user wants the crop clamped to image boundaries, calculate a final, failsafe overlap now.
    If (Not m_AllowEnlarge) Then
        Dim tmpOverlap As RectF
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
    toolpanel_Crop.tudCrop(4).Value = fracNumerator
    toolpanel_Crop.tudCrop(5).Value = fracDenominator
    
    toolpanel_Crop.cmdCommit(0).Enabled = Tools_Crop.IsValidCropActive()
    toolpanel_Crop.cmdCommit(1).Enabled = Tools_Crop.IsValidCropActive()
    
    'Unlock updates
    Tools.SetToolBusyState False
    
End Sub

'The crop toolpanel relays user input via this function.  Left/top/width/height as passed as integers; aspect ratio as a float.
Public Sub RelayCropChangesFromUI(ByVal changedProperty As PD_Dimension, Optional ByVal newPropI As Long = 0, Optional ByVal newPropF As Single = 0!)
    
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
            If m_IsAspectLocked Then m_LockedAspectRatio = newPropF
            
        Case pdd_AspectRatioH
            m_CropRectF.Height = newPropI
            toolpanel_Crop.tudCrop(3).Value = newPropI
            If m_IsAspectLocked And (newPropF <> 0!) Then m_LockedAspectRatio = 1! / newPropF
        
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
            If m_IsAspectLocked And (m_CropRectF.Height >= 1!) Then m_LockedAspectRatio = m_CropRectF.Width / m_CropRectF.Height
            
            fracNumerator = toolpanel_Crop.tudCrop(4).Value
            toolpanel_Crop.tudCrop(4).Value = toolpanel_Crop.tudCrop(5).Value
            toolpanel_Crop.tudCrop(5).Value = fracNumerator
                
            Tools.SetToolBusyState False
            
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
