Attribute VB_Name = "Tools_Crop"
'***************************************************************************
'Crop tool interface
'Copyright 2024-2024 by Tanner Helland
'Created: 12/November/24
'Last updated: 12/November/24
'Last update: initial build
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
    
    'We now want to add all lines to a path, which we'll render all at once
    Dim cropOutline As pd2DPath
    Set cropOutline = New pd2DPath
    cropOutline.AddRectangle_RectF cropRectCanvas
    
    'Stroke the outline path
    PD2D.DrawPath cSurface, basePenInactive, cropOutline
    PD2D.DrawPath cSurface, topPenInactive, cropOutline
    
    'Next, we want to render drag outlines over the four corner vertices.
    Dim ptCorners() As PointFloat, numPoints As Long
    numPoints = GetCropCorners(ptCorners)
    
    'Next, convert each corner from image coordinate space to the active viewport coordinate space
    Drawing.ConvertListOfImageCoordsToCanvasCoords dstCanvas, srcImage, ptCorners, False
    
    Dim cornerSize As Single, halfCornerSize As Single
    cornerSize = SQUARE_CORNER_SIZE
    halfCornerSize = cornerSize * 0.5!
    
    Dim i As Long
    For i = 0 To numPoints - 1
        If (i = m_idxHover) Then
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

Public Sub NotifyMouseDown(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByRef srcCanvas As pdCanvas, ByRef srcImage As pdImage, ByVal canvasX As Single, ByVal canvasY As Single)
    
    'Cache initial x/y positions
    m_LMBDown = ((Button And pdLeftButton) = pdLeftButton)
    
    If m_LMBDown Then
        
        'Translate the canvas coordinates to image coordinates
        Dim imgX As Double, imgY As Double
        Drawing.ConvertCanvasCoordsToImageCoords srcCanvas, srcImage, canvasX, canvasY, imgX, imgY, False
        ResetCropRectF imgX, imgY
        
        m_InitImgX = imgX
        m_InitImgY = imgY
        m_LastImgX = imgX
        m_LastImgY = imgY
        
    Else
        m_InitImgX = INVALID_X_COORD
        m_InitImgY = INVALID_Y_COORD
    End If
    
End Sub

Public Sub NotifyMouseMove(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByRef srcCanvas As pdCanvas, ByRef srcImage As pdImage, ByVal canvasX As Single, ByVal canvasY As Single)
    
    'Convert the current x/y position from canvas to image coordinates
    Dim imgX As Double, imgY As Double
    Drawing.ConvertCanvasCoordsToImageCoords srcCanvas, srcImage, canvasX, canvasY, imgX, imgY, False
    
    'Regardless of mousedown state, see if the mouse is over any of the corner nodes of the crop region.
    'if (
    
    'Cache current x/y positions
    If m_LMBDown Then
        
        'Move coordinates around to ensure positive width/height
        UpdateCropRectF imgX, imgY
        m_LastImgX = imgX
        m_LastImgY = imgY
        
        'TODO: apply aspect ratio, if any
        
    End If
    
    'Update the currently hovered crop corner, if any
    UpdateMousePOI imgX, imgY
    
    'Finally, request a viewport redraw (to reflect any changes)
    Viewport.Stage4_FlipBufferAndDrawUI srcImage, FormMain.MainCanvas(0)
    
End Sub

Public Sub NotifyMouseUp(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByRef srcCanvas As pdCanvas, ByRef srcImage As pdImage, ByVal canvasX As Single, ByVal canvasY As Single, ByVal numOfMouseMovements As Long, ByVal clickEventAlsoFiring As Boolean)
    
    'Update cached button status now, before handling the actual event.  (This ensures that
    ' a viewport redraw, if any, will wipe all UI elements we may have rendered over the canvas.)
    If ((Button And pdLeftButton) = pdLeftButton) Then m_LMBDown = False
    
    'Convert the current x/y position from canvas to image coordinates
    Dim imgX As Double, imgY As Double
    Drawing.ConvertCanvasCoordsToImageCoords srcCanvas, srcImage, canvasX, canvasY, imgX, imgY, False
        
    'Double-click commits the crop, so we may need to flag something here?  idk yet...
    If clickEventAlsoFiring Then
        
        'Dim zoomIn As Boolean
        'If ((Button And pdLeftButton) <> 0) Then
        '    zoomIn = True
        'ElseIf ((Button And pdRightButton) <> 0) Then
        '    zoomIn = False
        'Else
        '    m_InitImgX = INVALID_X_COORD
        '    m_InitImgY = INVALID_Y_COORD
        '    Exit Sub
        'End If
        '
        'Tools_Zoom.RelayCanvasZoom srcCanvas, srcImage, canvasX, canvasY, zoomIn
    
    'If this is a click-drag event, we need to simply resize the crop area to match.
    ' (Commit happens in a separate step.)
    Else
        
        'Update the crop rectangle against these (final) coordinates
        UpdateCropRectF imgX, imgY
        m_LastImgX = imgX
        m_LastImgY = imgY
        
        'Bail if coordinates are bad
        If (m_InitImgX = INVALID_X_COORD) Or (m_InitImgY = INVALID_Y_COORD) Then Exit Sub
        If (Not ValidateCropRectF()) Then Exit Sub
        
        'We now need to retrieve the current viewport rect in screen space (actual pixels)
        'Dim viewportWidth As Double, viewportHeight As Double
        'viewportWidth = FormMain.MainCanvas(0).GetCanvasWidth
        'viewportHeight = FormMain.MainCanvas(0).GetCanvasHeight
        
        'Calculate a width and height ratio in advance, and note that we know width/height
        ' are non-zero (thanks to a check above).
        'Dim horizontalRatio As Double, verticalRatio As Double
        'horizontalRatio = viewportWidth / rectImageCoords.Width
        'verticalRatio = viewportHeight / rectImageCoords.Height
        
        'The smaller of the two ratios is our limiting factor
        'Dim targetRatio As Double
        'targetRatio = PDMath.Min2Float_Single(horizontalRatio, verticalRatio)
        
        'TODO: everything past this point
        'With all calculations complete, we just need to assign the new values!
        
        'Suspend automatic viewport rendering, then assign new zoom
        srcCanvas.SetRedrawSuspension True
        
        'Reinstate canvas redraws, then reset the viewport buffer (while passing the new scrollbar
        ' values that we want to use - we pass them here and let the viewport assign them, because
        ' it will also determine new max/min values for the scroll bars as part of the zoom calculation).
        srcCanvas.SetRedrawSuspension False
        Viewport.Stage1_InitializeBuffer srcImage, srcCanvas
        
        'Notify any other UI elements of the change (e.g. the top-right navigator window)
        Viewport.NotifyEveryoneOfViewportChanges
        
    End If
    
    'Reset x/y trackers
    'm_InitImgX = INVALID_X_COORD
    'm_InitImgY = INVALID_Y_COORD
        
End Sub

'Only returns usable data when IsValidCropActive() is TRUE
Public Function GetCropRectF() As RectF
    GetCropRectF = m_CropRectF
End Function

'Check before using anything from GetCropRectF, above
Public Function IsValidCropActive() As Boolean
    IsValidCropActive = ValidateCropRectF
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

    Select Case m_idxHover
    
        'Mouse not over the crop
        Case poi_Undefined
            srcCanvasView.RequestCursor_System IDC_ARROW
            
        'Mouse is within the crop region, but not over a corner node
        Case poi_Interior
            srcCanvasView.RequestCursor_System IDC_SIZEALL
        
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
Private Sub UpdateMousePOI(ByVal imgX As Long, ByVal imgY As Long)
    
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
    m_idxHover = poi_Undefined
    m_idxHover = PDMath.FindClosestPointInFloatArray(imgX, imgY, minDistance, cropCorners)
    
    'If the mouse is not near a corner, perform an additional check for the crop rect's interior
    If (m_idxHover = poi_Undefined) Then
        
        Dim tmpPath As pd2DPath
        Set tmpPath = New pd2DPath
        tmpPath.AddRectangle_RectF m_CropRectF
        
        If tmpPath.IsPointInsidePathF(imgX, imgY) Then
            m_idxHover = poi_Interior
        Else
            m_idxHover = poi_Undefined
        End If
        
    End If
    
End Sub

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
        
    With m_CropRectF
    
        If (newX < m_InitImgX) Then
            If m_IsWidthLocked Then
                .Left = (m_InitImgX - m_LockedWidth)
                .Width = m_LockedWidth
            Else
                .Left = newX
                .Width = m_InitImgX - newX
            End If
        Else
            .Left = m_InitImgX
            If m_IsWidthLocked Then
                .Width = m_LockedWidth
            Else
                .Width = newX - m_InitImgX
            End If
        End If
        
        If (newY < m_InitImgY) Then
            If m_IsHeightLocked Then
                .Top = (m_InitImgY - m_LockedHeight)
                .Height = m_LockedHeight
            Else
                .Top = newY
                .Height = m_InitImgY - newY
            End If
        Else
            .Top = m_InitImgY
            If m_IsHeightLocked Then
                .Height = m_LockedHeight
            Else
                .Height = newY - m_InitImgY
            End If
        End If
        
    End With
    
    'Lock position and height to their nearest integer equivalent
    PDMath.GetIntClampedRectF m_CropRectF
    
    'If the crop rect is valid, relay its values to the toolpanel
    If ValidateCropRectF() Then RelayCropChangesToUI
    
End Sub

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
    With m_CropRectF
        If (.Left = INVALID_X_COORD) Or (.Top = INVALID_Y_COORD) Then ValidateCropRectF = False
        If (.Width <= 0) Or (.Height <= 0) Then ValidateCropRectF = False
    End With
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
    PDMath.ConvertToFraction m_CropRectF.Width / m_CropRectF.Height, fracNumerator, fracDenominator, 0.005
    
    'Aspect ratios are typically given in terms of base 10 if possible, so change values like 8:5 to 16:10
    If (fracDenominator = 5) Then
        fracNumerator = fracNumerator * 2
        fracDenominator = fracDenominator * 2
    End If
    
    toolpanel_Crop.tudCrop(4).Value = fracNumerator
    toolpanel_Crop.tudCrop(5).Value = fracDenominator
    
    'Unlock updates
    Tools.SetToolBusyState False
    
End Sub

'The crop toolpanel relays user input via this function.  Left/top/width/height as passed as integers; aspect ratio as a float.
Public Sub RelayCropChangesFromUI(ByVal changedProperty As PD_Dimension, Optional ByVal newPropI As Long = 0, Optional ByVal newPropF As Single = 0!)
    
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
                
                'Changes to width/height (with unlocked aspect ratio) will obviously change the
                ' current aspect ratio.  Re-calculate aspect ratio and update those spinners accordingly.
                Dim fracNumerator As Long, fracDenominator As Long
                PDMath.ConvertToFraction m_CropRectF.Width / m_CropRectF.Height, fracNumerator, fracDenominator, 0.005
                
                'Aspect ratios are typically given in terms of base 10 if possible, so change values like 8:5 to 16:10
                If (fracDenominator = 5) Then
                    fracNumerator = fracNumerator * 2
                    fracDenominator = fracDenominator * 2
                End If
                
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
            
    End Select
    
    'After a property change, we must validate the crop rectangle to make sure no funny business occurred
    ValidateCropRectF
    
    '...then request a viewport redraw to reflect any changes
    Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub
