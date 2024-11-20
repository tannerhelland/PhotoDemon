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
Private m_InitCanvasX As Double, m_InitCanvasY As Double

'Populate in _MouseMove
Private m_LastCanvasX As Double, m_LastCanvasY As Double

Public Sub DrawCanvasUI(ByRef dstCanvas As pdCanvas, ByRef srcImage As pdImage)
    
    'Because coords are stored in canvas coordinate space, rendering the UI is easy.
    
    
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
    
    'We now want to add all lines to a path, which we'll render all at once
    Dim cropOutline As pd2DPath
    Set cropOutline = New pd2DPath
    cropOutline.AddRectangle_Absolute m_InitCanvasX, m_InitCanvasY, m_LastCanvasX, m_LastCanvasY
    
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
        m_InitCanvasX = canvasX
        m_InitCanvasY = canvasY
        m_LastCanvasX = canvasX
        m_LastCanvasY = canvasY
    Else
        m_InitCanvasX = INVALID_X_COORD
        m_InitCanvasY = INVALID_Y_COORD
    End If
    
End Sub

Public Sub NotifyMouseMove(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByRef srcCanvas As pdCanvas, ByRef srcImage As pdImage, ByVal canvasX As Single, ByVal canvasY As Single)
    
    'Cache current x/y positions
    If m_LMBDown Then
        
        m_LastCanvasX = canvasX
        m_LastCanvasY = canvasY
        
        'TODO: apply aspect ratio, if any
        
        'Request a viewport redraw too
        Viewport.Stage4_FlipBufferAndDrawUI srcImage, FormMain.MainCanvas(0)
    
    End If
    
End Sub

Public Sub NotifyMouseUp(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByRef srcCanvas As pdCanvas, ByRef srcImage As pdImage, ByVal canvasX As Single, ByVal canvasY As Single, ByVal numOfMouseMovements As Long, ByVal clickEventAlsoFiring As Boolean)
    
    'Update cached button status now, before handling the actual event.  (This ensures that
    ' a viewport redraw, if any, will wipe all UI elements we may have rendered over the canvas.)
    If ((Button And pdLeftButton) = pdLeftButton) Then m_LMBDown = False
    
    'Double-click commits the crop, so we may need to flag something here?  idk yet...
    If clickEventAlsoFiring Then
        
        'Dim zoomIn As Boolean
        'If ((Button And pdLeftButton) <> 0) Then
        '    zoomIn = True
        'ElseIf ((Button And pdRightButton) <> 0) Then
        '    zoomIn = False
        'Else
        '    m_InitCanvasX = INVALID_X_COORD
        '    m_InitCanvasY = INVALID_Y_COORD
        '    Exit Sub
        'End If
        '
        'Tools_Zoom.RelayCanvasZoom srcCanvas, srcImage, canvasX, canvasY, zoomIn
    
    'If this is a click-drag event, we need to simply resize the crop area to match.
    ' (Commit happens in a separate step.)
    Else
        
        'Bail if initial coordinates are bad
        If (m_InitCanvasX = INVALID_X_COORD) Or (m_InitCanvasY = INVALID_Y_COORD) Then Exit Sub
        
        'TODO: update coords?
        m_LastCanvasX = canvasX
        m_LastCanvasY = canvasY
        
        'Start by solving for the size of the crop region, in image coordinates.
        'Dim rectImageCoords As RectF
        'FillZoomRect_ImageCoords srcCanvas, srcImage, rectImageCoords
        '
        ''Failsafe check for DBZ errors
        'If (rectImageCoords.Width <= 0!) Or (rectImageCoords.Height <= 0!) Then Exit Sub
        
        'We now need to retrieve the current viewport rect in screen space (actual pixels)
        Dim viewportWidth As Double, viewportHeight As Double
        viewportWidth = FormMain.MainCanvas(0).GetCanvasWidth
        viewportHeight = FormMain.MainCanvas(0).GetCanvasHeight
        
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
    m_InitCanvasX = INVALID_X_COORD
    m_InitCanvasY = INVALID_Y_COORD
        
End Sub
