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
    
    'Stroke the path
    PD2D.DrawPath cSurface, basePenInactive, cropOutline
    PD2D.DrawPath cSurface, topPenInactive, cropOutline
    
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
    
    'Cache current x/y positions
    If m_LMBDown Then
        
        'Convert the current x/y position from canvas to image coordinates
        Dim imgX As Double, imgY As Double
        Drawing.ConvertCanvasCoordsToImageCoords srcCanvas, srcImage, canvasX, canvasY, imgX, imgY, False
        
        'Move coordinates around to ensure positive width/height
        UpdateCropRectF imgX, imgY
        m_LastImgX = imgX
        m_LastImgY = imgY
        
        'TODO: apply aspect ratio, if any
        
        'Request a viewport redraw too
        Viewport.Stage4_FlipBufferAndDrawUI srcImage, FormMain.MainCanvas(0)
    
    End If
    
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

Private Sub UpdateCropRectF(ByVal newX As Single, ByVal newY As Single)
    
    With m_CropRectF
        If (newX < m_InitImgX) Then
            .Left = newX
            .Width = m_InitImgX - newX
        Else
            .Left = m_InitImgX
            .Width = newX - m_InitImgX
        End If
        If (newY < m_InitImgY) Then
            .Top = newY
            .Height = m_InitImgY - newY
        Else
            .Top = m_InitImgY
            .Height = newY - m_InitImgY
        End If
    End With
    
    'Lock position and height to the nearest edge of the image
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

Public Sub RelayCropChangesFromUI()

End Sub
