Attribute VB_Name = "Tools_Zoom"
'***************************************************************************
'Zoom on-canvas tool interface
'Copyright 2021-2026 by Tanner Helland
'Created: 14/December/21
'Last updated: 14/December/21
'Last update: start migrating zoom bits from elsewhere into this dedicated module
'
'PD's Zoom tool is very straightforward.  It basically relays simple zoom commands from the canvas
' to PD's viewport engine.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'TRUE if a _MouseDown event was received
Private m_LMBDown As Boolean

'Populated in _MouseDown
Private Const INVALID_X_COORD As Double = DOUBLE_MAX, INVALID_Y_COORD As Double = DOUBLE_MAX
Private m_InitCanvasX As Double, m_InitCanvasY As Double

'Populate in _MouseMove
Private m_LastCanvasX As Double, m_LastCanvasY As Double

Public Sub DrawCanvasUI(ByRef dstCanvas As pdCanvas, ByRef srcImage As pdImage)
    
    'We only care about rendering if the left mouse-button is down, and a valid set of
    ' mouse coordinates has been stored.
    Dim okToRender As Boolean
    okToRender = m_LMBDown
    okToRender = okToRender And (m_InitCanvasX <> INVALID_X_COORD) And (m_InitCanvasY <> INVALID_Y_COORD)
    okToRender = okToRender And (m_InitCanvasX <> m_LastCanvasX) And (m_InitCanvasY <> m_LastCanvasY)
    If (Not okToRender) Then Exit Sub
    
    'Still here?  Guess we'd better render a UI!
    
    'Clone a pair of UI pens from the main rendering module.  (Note that we clone them because we need
    ' to modify some rendering properties, and we don't want to fuck up the central cache.)
    Dim basePenInactive As pd2DPen, topPenInactive As pd2DPen
    Drawing.CloneCachedUIPens basePenInactive, topPenInactive, False
    
    'Specify squared-off line joins for our pens; this looks better for this particular tool
    basePenInactive.SetPenLineJoin P2_LJ_Bevel
    topPenInactive.SetPenLineJoin P2_LJ_Bevel
    
    'Create other required pd2D drawing tools (a surface)
    Dim cSurface As pd2DSurface
    Drawing2D.QuickCreateSurfaceFromDC cSurface, dstCanvas.hDC, True
    cSurface.SetSurfacePixelOffset P2_PO_Normal
    
    'Because stored coordinates are in canvas space, we can use them as-is
    Dim renderRect As RectF
    With renderRect
        .Left = PDMath.Min2Float_Single(m_InitCanvasX, m_LastCanvasX)
        .Top = PDMath.Min2Float_Single(m_InitCanvasY, m_LastCanvasY)
        .Width = PDMath.Max2Float_Single(m_InitCanvasX, m_LastCanvasX) - .Left
        .Height = PDMath.Max2Float_Single(m_InitCanvasY, m_LastCanvasY) - .Top
    End With
    
    PD2D.DrawRectangleF_FromRectF cSurface, basePenInactive, renderRect
    PD2D.DrawRectangleF_FromRectF cSurface, topPenInactive, renderRect
    
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
        
        'Request a viewport redraw too
        Viewport.Stage4_FlipBufferAndDrawUI srcImage, FormMain.MainCanvas(0)
    
    End If
    
End Sub

Public Sub NotifyMouseUp(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByRef srcCanvas As pdCanvas, ByRef srcImage As pdImage, ByVal canvasX As Single, ByVal canvasY As Single, ByVal numOfMouseMovements As Long, ByVal clickEventAlsoFiring As Boolean)
    
    'Update cached button status now, before handling the actual event.  (This ensures that
    ' a viewport redraw, if any, will wipe all UI elements we may have rendered over the canvas.)
    If ((Button And pdLeftButton) = pdLeftButton) Then m_LMBDown = False
    
    'Left-click zooms in, right-click zooms out (per convention with other software)
    If clickEventAlsoFiring Then
        
        Dim zoomIn As Boolean
        If ((Button And pdLeftButton) <> 0) Then
            zoomIn = True
        ElseIf ((Button And pdRightButton) <> 0) Then
            zoomIn = False
        Else
            m_InitCanvasX = INVALID_X_COORD
            m_InitCanvasY = INVALID_Y_COORD
            Exit Sub
        End If
        
        Tools_Zoom.RelayCanvasZoom srcCanvas, srcImage, canvasX, canvasY, zoomIn
    
    'If this is a click-drag event, we need to solve a more difficult equation
    Else
        
        'Bail if initial coordinates are bad
        If (m_InitCanvasX = INVALID_X_COORD) Or (m_InitCanvasY = INVALID_Y_COORD) Then Exit Sub
        
        'Using the zoom tool, the user can click-drag a region to select it for zooming.
        ' Our job is to find the "best" zoom value for that rectangle, so that the entire rectangle
        ' fits nicely inside the viewport area.
        m_LastCanvasX = canvasX
        m_LastCanvasY = canvasY
        
        'Start by solving for the size of the selected region, in image coordinates.
        Dim rectImageCoords As RectF
        FillZoomRect_ImageCoords srcCanvas, srcImage, rectImageCoords
        
        'Failsafe check for DBZ errors
        If (rectImageCoords.Width <= 0!) Or (rectImageCoords.Height <= 0!) Then Exit Sub
        
        'We now need to retrieve the current viewport rect in screen space (actual pixels)
        Dim viewportWidth As Double, viewportHeight As Double
        viewportWidth = FormMain.MainCanvas(0).GetCanvasWidth
        viewportHeight = FormMain.MainCanvas(0).GetCanvasHeight
        
        'Calculate a width and height ratio in advance, and note that we know width/height
        ' are non-zero (thanks to a check above).
        Dim horizontalRatio As Double, verticalRatio As Double
        horizontalRatio = viewportWidth / rectImageCoords.Width
        verticalRatio = viewportHeight / rectImageCoords.Height
        
        'The smaller of the two ratios is our limiting factor
        Dim targetRatio As Double
        targetRatio = PDMath.Min2Float_Single(horizontalRatio, verticalRatio)
                        
        'We now need to find the closest zoom factor to this one (from the pre-set list of zoom factors
        ' that PD exposes directly to the user).
        Dim nearestZoomIndex As Long, nearestZoomRatio As Double
        nearestZoomIndex = Zoom.GetNearestZoomOutIndex_FromRatio(targetRatio)
        nearestZoomRatio = Zoom.GetZoomRatioFromIndex(nearestZoomIndex)
        
        'With the appropriate zoom level established, all that's left to do is center the selected region
        ' inside the viewport, at the selected zoom.  To do this, figure out how many pixels (in image
        ' coordinates) we'll have to work with at the new viewport size.
        Dim newViewportWidth As Long, newViewportHeight As Long
        newViewportWidth = viewportWidth / nearestZoomRatio
        newViewportHeight = viewportHeight / nearestZoomRatio
        
        'Use the original [top, left] position we calculated, then add necessary padding based
        ' on the *actual* size of the viewport vs the "idealized" size calculated before.
        Dim newOffsetX As Long, newOffsetY As Long
        newOffsetX = rectImageCoords.Left - (newViewportWidth - rectImageCoords.Width) \ 2
        newOffsetY = rectImageCoords.Top - (newViewportHeight - rectImageCoords.Height) \ 2
        
        'With all calculations complete, we just need to assign the new values!
        
        'Suspend automatic viewport rendering, then assign new zoom
        srcCanvas.SetRedrawSuspension True
        srcCanvas.SetZoomDropDownIndex nearestZoomIndex
        PDImages.GetActiveImage.SetZoomIndex nearestZoomIndex
        
        'Reinstate canvas redraws, then reset the viewport buffer (while passing the new scrollbar
        ' values that we want to use - we pass them here and let the viewport assign them, because
        ' it will also determine new max/min values for the scroll bars as part of the zoom calculation).
        srcCanvas.SetRedrawSuspension False
        Viewport.Stage1_InitializeBuffer srcImage, srcCanvas, VSR_ResetToCustom, newOffsetX, newOffsetY
        
        'Notify any other UI elements of the change (e.g. the top-right navigator window)
        Viewport.NotifyEveryoneOfViewportChanges
        
    End If
    
    'Reset x/y trackers
    m_InitCanvasX = INVALID_X_COORD
    m_InitCanvasY = INVALID_Y_COORD
        
End Sub

'When a canvas event initiates zoom (mousewheel, zoom tool, etc), send the relevant canvas info here
' and this function will perform the actual zoom change.  It will return TRUE if the caller needs to
Public Sub RelayCanvasZoom(ByRef srcCanvas As pdCanvas, ByRef srcImage As pdImage, ByVal canvasX As Double, ByVal canvasY As Double, ByVal zoomIn As Boolean)

    If (Not srcCanvas.IsCanvasInteractionAllowed()) Then Exit Sub
    
    'Before doing anything else, cache the current mouse coordinates (in both Canvas and Image coordinate spaces)
    Dim imgX As Double, imgY As Double
    Drawing.ConvertCanvasCoordsToImageCoords srcCanvas, srcImage, canvasX, canvasY, imgX, imgY, True
    
    'Suspend automatic viewport redraws until we are done with our calculations.
    ' (Same goes for the canvas, which needs to stop handling scroll bar synchronization until we're done.)
    Viewport.DisableRendering
    srcCanvas.SetRedrawSuspension True
    
    'Calculate a new zoom value
    If srcCanvas.IsZoomEnabled() Then
        If zoomIn Then
            If (srcCanvas.GetZoomDropDownIndex > 0) Then srcCanvas.SetZoomDropDownIndex Zoom.GetNearestZoomInIndex(srcCanvas.GetZoomDropDownIndex)
        Else
            If (srcCanvas.GetZoomDropDownIndex <> Zoom.GetZoomCount) Then srcCanvas.SetZoomDropDownIndex Zoom.GetNearestZoomOutIndex(srcCanvas.GetZoomDropDownIndex)
        End If
    End If
    
    'Relay the new zoom value to the target pdImage object (pdImage objects store their current zoom value,
    ' so we can preserve it when switching between images)
    srcImage.SetZoomIndex srcCanvas.GetZoomDropDownIndex()
    
    'Re-enable automatic viewport redraws
    Viewport.EnableRendering
    srcCanvas.SetRedrawSuspension False
    
    'Request a manual redraw from Viewport.Stage1_InitializeBuffer, while supplying our x/y coordinates so that
    ' it can preserve mouse position relative to the underlying image.
    Viewport.Stage1_InitializeBuffer srcImage, srcCanvas, VSR_PreservePointPosition, canvasX, canvasY, imgX, imgY
    
    'Notify external UI elements of the change
    Viewport.NotifyEveryoneOfViewportChanges

End Sub

'Private functions follow

'Unlike other tools, the zoom tool stores coordinates in canvas space.  When solving the equations
' for click-drag zoom behavior, we need to translate those coordinates into image space.
Private Sub FillZoomRect_ImageCoords(ByRef srcCanvas As pdCanvas, ByRef srcImage As pdImage, ByRef rectImageCoords As RectF)

    Dim rectCanvasCoords As RectF
    With rectCanvasCoords
        .Left = PDMath.Min2Float_Single(m_InitCanvasX, m_LastCanvasX)
        .Top = PDMath.Min2Float_Single(m_InitCanvasY, m_LastCanvasY)
        .Width = PDMath.Max2Float_Single(m_InitCanvasX, m_LastCanvasX) - .Left
        .Height = PDMath.Max2Float_Single(m_InitCanvasY, m_LastCanvasY) - .Top
    End With
    
    Dim newX As Double, newY As Double
    With rectCanvasCoords
        Drawing.ConvertCanvasCoordsToImageCoords srcCanvas, srcImage, .Left, .Top, newX, newY, False
    End With
    
    rectImageCoords.Left = newX
    rectImageCoords.Top = newY
    
    With rectCanvasCoords
        Drawing.ConvertCanvasCoordsToImageCoords srcCanvas, srcImage, .Left + .Width, .Top + .Height, newX, newY, False
    End With
    
    rectImageCoords.Width = newX - rectImageCoords.Left
    rectImageCoords.Height = newY - rectImageCoords.Top
    
End Sub
