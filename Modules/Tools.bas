Attribute VB_Name = "Tools"
'***************************************************************************
'Helper functions for various PhotoDemon tools
'Copyright 2014-2026 by Tanner Helland
'Created: 06/February/14
'Last updated: 16/January/25
'Last update: add snap support when rotating a layer
'
'To keep the pdCanvas user control codebase lean, many of its MouseMove events redirect here, to specialized
' functions that take mouse actions on the canvas and translate them into tool actions.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Various constants related to custom tool behavior
Public Const PD_TEXT_TOOL_CREATED_NEW_LAYER As Long = &H1

'The drag-to-pan tool uses these values to store the original image offset
Private m_InitHScroll As Long, m_InitVScroll As Long

'Upon initiating a layer interaction, the move/size tool caches two sets of original layer values: the layer's transformed coordinates,
' which include the results of any affine transforms (e.g. rotation), and the layer's "pure" coordinates, e.g. without affine transforms.
' These coordinates are crucial for establishing the difference between the original layer offsets/dimensions, and any new ones created
' by canvas interactions.
'
'As a convenience, we also cache the layer's aspect ratio.  This is important for operations that support use of the SHIFT key.
'
'Finally, the initial mouse x/y values are also supplied, in case they are needed later on.  (We call these m_InitImageX/Y as a
' reminder that they exist in the *image* coordinate space, not the *canvas* coordinate space.)  We also make a copy of these values
' in the current layer's coordinate space (e.g. with affine transforms considered)
Private m_InitLayerCoords_Transformed(0 To 3) As PointFloat
Private m_InitLayerCoords_Pure(0 To 3) As PointFloat
Private m_LayerAspectRatio As Double
Private m_InitImageX As Double, m_InitImageY As Double, m_InitLayerX As Single, m_InitLayerY As Single
Private m_InitLayerRotateCenterX As Single, m_InitLayerRotateCenterY As Single

'If a point of interest is being modified by a tool action, its ID will be stored here.  Make sure to clear this value
' (to -1, which means "no point of interest") when you are finished with it (typically after MouseUp).
Private m_CurPOI As PD_PointOfInterest

'Some tools require complicated synchronization between UI elements (e.g. text up/downs that display coordinates), and PD internals
' (e.g. layer positions).  To prevent infinite update loops (UI change > internal change > UI change > ...), tools can mark their
' current state as "busy".  Subsequent UI refreshes will be rejected until the "busy" state is reset.
Private m_ToolIsBusy As Boolean

'Some tools may perform different actions under different circumstances.  At MouseDown, they can set this value to anything
' they want.  At MouseUp, this value can be retrieved to know what kind of action occurred.  (For example, the text tool uses
' this to know if the previous MouseDown actually created the current text layer, or if it is just editing an existing layer.)
'
'IMPORTANT: after retrieval, this value is forcibly reset to zero.  Do not check it more than once without internally caching it.
Private m_CustomToolMarker As Long

'Some tool-related user preferences are cached locally, to improve performance (vs pulling them from the
' central resource manager).
Private m_HighResMouseInputAllowed As Boolean

'While using paint tools, the user can use the ALT key to temporarily swap to the color picker tool.
' When the ALT key is released, PhotoDemon will automatically active their original tool.  Because
' this requires cross-tool communication, this module needs to store the relevant tracker.
Private m_PaintToolAltState As Boolean

'As of 8.4, middle-clicks automatically engage the pan/move tool; we track this state independently,
' because we need to restore previous tool state when the button is released.
Private m_MiddleMouseState As Boolean

'As of 9.0, the Move/Size tool behaves differently when used on a selection.  (The selected pixels
' will be auto-copied/cut into their own layer, and that new layer will be moved instead of the
' original one.)
Private m_MoveSelectedPixels As Boolean

'Get/Set the "alternate" state for a paint tool (typically triggered by pressing ALT)
Public Function GetToolAltState() As Boolean
    GetToolAltState = m_PaintToolAltState
End Function

Public Sub SetToolAltState(ByVal newState As Boolean)
    m_PaintToolAltState = newState
End Sub

'Get/Set tool "busy" state; when a tool is busy, many operations are suspended for performance reasons
Public Function GetToolBusyState() As Boolean
    GetToolBusyState = m_ToolIsBusy
End Function

Public Sub SetToolBusyState(ByVal newState As Boolean)
    m_ToolIsBusy = newState
End Sub

'Middle-mouse button state
Public Function GetToolMMBState() As Boolean
    GetToolMMBState = m_MiddleMouseState
End Function

Public Sub SetToolMMBState(ByVal newState As Boolean)
    m_MiddleMouseState = newState
End Sub

Public Function GetCustomToolState() As Long
    GetCustomToolState = m_CustomToolMarker
    m_CustomToolMarker = 0
End Function

Public Sub SetCustomToolState(ByVal newState As Long)
    m_CustomToolMarker = newState
End Sub

Public Function IsSelectionToolActive() As Boolean
    IsSelectionToolActive = (g_CurrentTool = SELECT_CIRC) Or (g_CurrentTool = SELECT_LASSO) Or (g_CurrentTool = SELECT_POLYGON) Or (g_CurrentTool = SELECT_RECT) Or (g_CurrentTool = SELECT_WAND)
End Function

'When a tool is finished processing, it can call this function to release all tool tracking variables
Public Sub TerminateGenericToolTracking()
    
    'Reset the current POI, if any
    m_CurPOI = poi_Undefined
    
    'Reset any selection tracking
    m_MoveSelectedPixels = False
    
End Sub

'The move tool uses this function to set various initial parameters for layer interactions.
Public Sub SetInitialLayerToolValues(ByRef srcImage As pdImage, ByRef srcLayer As pdLayer, ByVal mouseX_ImageSpace As Double, ByVal mouseY_ImageSpace As Double, Optional ByVal relevantPOI As PD_PointOfInterest = poi_Undefined, Optional ByVal useSelectedPixels As Boolean = False, Optional ByVal Shift As ShiftConstants)
    
    'Note whether we a selection is active
    m_MoveSelectedPixels = useSelectedPixels
    
    'If a selection is active, we need to cut (or copy) the currently selected pixels into their own new layer.
    If m_MoveSelectedPixels Then
        
        'Create the new layer.  Note that Copy vs Cut is determined by first noting the user's default setting,
        ' then toggling that setting if ALT was pressed.
        Dim eraseOriginalPixels As Boolean
        eraseOriginalPixels = Tools_Move.GetMoveSelectedPixels_DefaultCut()
        If ((Shift And vbAltMask) = vbAltMask) Then eraseOriginalPixels = (Not eraseOriginalPixels)
        Layers.AddLayerViaSelection True, Tools_Move.GetMoveSelectedPixels_SampleMerged, eraseOriginalPixels
        
        'Silently point the layer reference at the newly created layer (we don't care about the original layer ref)
        Set srcLayer = srcImage.GetActiveLayer
        
        'Remove the active selection
        Selections.RemoveCurrentSelection False
        
    End If
    
    'Cache the initial mouse values.  Note that, per the parameter names, these must have already been converted to the image's
    ' coordinate space (NOT the canvas's!)
    m_InitImageX = mouseX_ImageSpace
    m_InitImageY = mouseY_ImageSpace
    
    'Also, make a copy of those coordinates in the current layer space
    Drawing.ConvertImageCoordsToLayerCoords srcImage, srcLayer, m_InitImageX, m_InitImageY, m_InitLayerX, m_InitLayerY
    
    'Make a copy of the current layer coordinates, with any affine transforms applied (rotation, etc)
    srcLayer.GetLayerCornerCoordinates m_InitLayerCoords_Transformed
    
    'Make a copy of the current layer coordinates, *without* affine transforms applied.  This is basically the rect of
    ' the layer as it would appear if no affine modifiers were active (e.g. without rotation, etc)
    Dim i As Long
    For i = 0 To 3
        Drawing.ConvertImageCoordsToLayerCoords srcImage, srcLayer, m_InitLayerCoords_Transformed(i).x, m_InitLayerCoords_Transformed(i).y, m_InitLayerCoords_Pure(i).x, m_InitLayerCoords_Pure(i).y
    Next i
    
    'Make a copy of the layer's rotational center point, in absolute image coordinates
    m_InitLayerRotateCenterX = m_InitLayerCoords_Pure(0).x + (srcLayer.GetLayerRotateCenterX * (m_InitLayerCoords_Pure(1).x - m_InitLayerCoords_Pure(0).x))
    m_InitLayerRotateCenterY = m_InitLayerCoords_Pure(0).y + (srcLayer.GetLayerRotateCenterY * (m_InitLayerCoords_Pure(2).y - m_InitLayerCoords_Pure(0).y))
    
    'Cache the layer's aspect ratio.  Note that this *does not include any current non-destructive transforms*!
    ' (We will use this to handle the SHIFT key, which typically means "preserve original image aspect ratio".)
    If (srcLayer.GetLayerHeight(False) <> 0#) Then
        m_LayerAspectRatio = srcLayer.GetLayerWidth(False) / srcLayer.GetLayerHeight(False)
    Else
        m_LayerAspectRatio = 1#
    End If
    
    'If a relevant POI was supplied, store it as well.  Note that not all tools make use of this.
    m_CurPOI = relevantPOI
        
End Sub

'The drag-to-pan tool uses this function to set the initial scroll bar values for a pan operation
Public Sub SetInitialCanvasScrollValues(ByRef srcCanvas As pdCanvas)
    m_InitHScroll = srcCanvas.GetScrollValue(pdo_Horizontal)
    m_InitVScroll = srcCanvas.GetScrollValue(pdo_Vertical)
End Sub

'The drag-to-pan tool uses this function to actually scroll the viewport area
Public Sub PanImageCanvas(ByVal initX As Long, ByVal initY As Long, ByVal curX As Long, ByVal curY As Long, ByRef srcImage As pdImage, ByRef srcCanvas As pdCanvas)

    'Prevent the canvas from redrawing itself until our pan operation is complete.  (This prevents juddery movement.)
    srcCanvas.SetRedrawSuspension True
    
    'Sub-pixel panning is now allowed (because we're awesome like that)
    Dim zoomRatio As Double
    zoomRatio = Zoom.GetZoomRatioFromIndex(srcImage.GetZoomIndex())
    
    'Calculate new scroll values
    Dim hOffset As Long, vOffset As Long
    If (zoomRatio <> 0#) Then
        hOffset = (initX - curX) / zoomRatio
        vOffset = (initY - curY) / zoomRatio
    End If
        
    'Factor in the initial scroll bar values
    hOffset = m_InitHScroll + hOffset
    vOffset = m_InitVScroll + vOffset
        
    'If these values lie within the bounds of their respective scroll bar(s), apply 'em
    If (hOffset < srcCanvas.GetScrollMin(pdo_Horizontal)) Then
        srcCanvas.SetScrollValue pdo_Horizontal, srcCanvas.GetScrollMin(pdo_Horizontal)
    ElseIf (hOffset > srcCanvas.GetScrollMax(pdo_Horizontal)) Then
        srcCanvas.SetScrollValue pdo_Horizontal, srcCanvas.GetScrollMax(pdo_Horizontal)
    Else
        srcCanvas.SetScrollValue pdo_Horizontal, hOffset
    End If
    
    If (vOffset < srcCanvas.GetScrollMin(pdo_Vertical)) Then
        srcCanvas.SetScrollValue pdo_Vertical, srcCanvas.GetScrollMin(pdo_Vertical)
    ElseIf (vOffset > srcCanvas.GetScrollMax(pdo_Vertical)) Then
        srcCanvas.SetScrollValue pdo_Vertical, srcCanvas.GetScrollMax(pdo_Vertical)
    Else
        srcCanvas.SetScrollValue pdo_Vertical, vOffset
    End If
    
    'Reinstate canvas redraws
    srcCanvas.SetRedrawSuspension False
    
    'Request the scroll-specific viewport pipeline stage
    Viewport.Stage2_CompositeAllLayers srcImage, FormMain.MainCanvas(0)
    
    'As of v8.0, rulers also need to be notified of this change.  (Normally they are notified
    ' of all canvas mouse events, but this tool is a little strange because we move the canvas
    ' *after* rulers have received mouse move notifications - so their coordinates are out of
    ' date by the time this function finishes).
    srcCanvas.RequestRulerUpdate
    
End Sub

'This function can be used to move and/or non-destructively resize an image layer.
'
'If this action occurs during a Mouse_Up event, the finalizeTransform parameter should be set to TRUE. This instructs the function
' to forward the transformation request to PD's central processor, so it can generate Undo/Redo data, be recorded as part of macros, etc.
Public Sub TransformCurrentLayer(ByVal curImageX As Double, ByVal curImageY As Double, ByRef srcImage As pdImage, ByRef srcLayer As pdLayer, ByRef srcCanvas As pdCanvas, Optional ByVal isShiftDown As Boolean = False, Optional ByVal finalizeTransform As Boolean = False)
    
    'Prevent the canvas from redrawing itself until our movement calculations are complete.
    ' (This prevents juddery movement.)
    srcCanvas.SetRedrawSuspension True
    
    'Also, mark the tool engine as busy to prevent re-entrance issues
    Tools.SetToolBusyState True
    
    'For operations only involving a single point of transformation (e.g. resizing a layer by corner-dragging),
    ' we can apply snapping *now*, to the mouse coordinate itself.
    '
    'For operations that transform multiple points (like moving an entire layer), we need to snap points *besides*
    ' the mouse pointer (e.g. the layer edges, which are not located at the mouse position), so we'll need to wait
    ' to snap until the transform has been applied to the underlying layer.
    '
    '(Note: snapping angle is not handled this way; it's handled later in the function.)
    Dim srcPtF As PointFloat, snappedPtF As PointFloat
    If Snap.GetSnap_Any() Then
        
        Select Case m_CurPOI
            Case poi_CornerNW, poi_CornerNE, poi_CornerSW, poi_CornerSE
                srcPtF.x = curImageX
                srcPtF.y = curImageY
                Snap.SnapPointByMoving srcPtF, snappedPtF
                curImageX = snappedPtF.x
                curImageY = snappedPtF.y
                
        End Select
        
    End If
    
    'Convert the current x/y pair to the layer coordinate space.  This takes into account any active affine transforms
    ' on the image (e.g. rotation), which may place the point in a totally different position relative to the underlying layer.
    Dim curLayerX As Single, curLayerY As Single
    Drawing.ConvertImageCoordsToLayerCoords srcImage, srcLayer, curImageX, curImageY, curLayerX, curLayerY
    
    'As a convenience for later calculations, calculate offsets between the initial transform coordinates (set at MouseDown)
    ' and the current ones.  Repeat this for both the image and layer coordinate spaces, as we need different ones for different
    ' transform types.
    Dim hOffsetLayer As Double, vOffsetLayer As Double, hOffsetImage As Double, vOffsetImage As Double
    hOffsetLayer = curLayerX - m_InitLayerX
    vOffsetLayer = curLayerY - m_InitLayerY
    
    hOffsetImage = curImageX - m_InitImageX
    vOffsetImage = curImageY - m_InitImageY
        
    'To prevent the user from flipping or mirroring the image, we must do some bound checking on their changes,
    ' and disallow anything that results in invalid coordinates or sizes.
    Dim newLeft As Single, newTop As Single, newRight As Single, newBottom As Single
    
    'The way we assign new offsets and/or sizes to the layer depends on the POI (point of interest) the user is interacting with.
    ' Layers currently support nine points of interest: each of their 4 corners, 4 rotational points (lying on the center of
    ' each edge), and anywhere in the layer interior (for moving the layer).
    
    'If this layer has an active rotation transform (e.g. srcLayer.GetLayerAngle <> 0), we may need to modify the layer's
    ' rotational center to compensate for positional and width/height changes.  This is only necessary for move/resize events,
    ' *not* rotation events (which is confusing, I know, but rotation events use a fixed rotation point).
    Dim rotateCleanupRequired As Boolean
    rotateCleanupRequired = False
    
    'Aspect ratio is locked on SHIFT keypress, or with the fixed toggle on the move/size toolpanel
    Dim lockAspectRatio As Boolean
    lockAspectRatio = isShiftDown
    If (g_CurrentTool = NAV_MOVE) Then lockAspectRatio = lockAspectRatio Or toolpanel_MoveSize.chkAspectRatio.Value
    
    'Check the POI we were given, and update the layer accordingly.
    With srcLayer
    
        Select Case m_CurPOI
            
            '-1: the mouse is not over the layer.  Do nothing.
            Case poi_Undefined
                Tools.SetToolBusyState False
                Snap.NotifyNoSnapping
                srcCanvas.SetRedrawSuspension False
                Exit Sub
                
            '0: the mouse is dragging the top-left corner of the layer.
            ' (The comments here are uniform for all POIs, so for brevity's sake, I'll keep the others short.)
            Case poi_CornerNW
                
                'The opposite corner coordinate (bottom-left) stays in exactly the same place
                newRight = m_InitLayerCoords_Pure(1).x
                newBottom = m_InitLayerCoords_Pure(3).y
                
                'Set the new left/top position to match the mouse coordinates, while also accounting for the shift key
                ' (which locks the current aspect ratio).
                If ((newRight - curLayerX) > 1!) Then newLeft = curLayerX Else newLeft = newRight - 1!
                If lockAspectRatio Then newTop = newBottom - (newRight - newLeft) / m_LayerAspectRatio Else newTop = curLayerY
                If ((newBottom - newTop) < 1!) Then newTop = newBottom - 1!
                
                'Locked aspect ratios and snapping have complex interactions.  While we originally snapped
                ' mouse coordinates at the start of this function, if the user has locked aspect ratio,
                ' correcting for aspect ratio likely pulled us away from our snap target.  We now need to
                ' manually calculate snap against the aspect-ratio corrected coordinates, and then *post*-snap,
                ' ensure aspect ratio is still OK.
                If lockAspectRatio And Snap.GetSnap_Any() Then
                    
                    'Snap the aspect-ratio-corrected points into place
                    srcPtF.x = newLeft
                    srcPtF.y = newTop
                    Snap.SnapPointByMoving srcPtF, snappedPtF
                    
                    'We have to pick an arbitrary direction to prioritize when snapping *and* preserving
                    ' aspect ratio, because we can't do both.  I've arbitrarily selected the x-position,
                    ' and if x is not snapped, we'll snap y (if we can).
                    If Snap.IsSnapped_X() Then
                        newLeft = snappedPtF.x
                        newTop = newBottom - (newRight - newLeft) / m_LayerAspectRatio
                        Snap.NotifyNoSnapping_Y
                    ElseIf Snap.IsSnapped_Y() Then
                        newTop = snappedPtF.y
                        newLeft = newRight - (newBottom - newTop) * m_LayerAspectRatio
                        Snap.NotifyNoSnapping_X
                    End If
                    
                    'Still important to validate layer rect before assigning
                    If (newLeft > newRight - 1!) Then newLeft = newRight - 1!
                    If (newTop > newBottom - 1!) Then newTop = newBottom - 1!
                    
                End If
                
                'Immediately relay the new coordinates to the source layer
                srcLayer.SetOffsetsAndModifiersTogether newLeft, newTop, newRight, newBottom
                
                'A helper block at the end of this function cleans up any rotation-related parameters to match the new coordinate
                rotateCleanupRequired = True
                
            '1: top-right corner
            Case poi_CornerNE
            
                newLeft = m_InitLayerCoords_Pure(0).x
                newBottom = m_InitLayerCoords_Pure(2).y
                
                If ((curLayerX - newLeft) > 1!) Then newRight = curLayerX Else newRight = newLeft + 1!
                If lockAspectRatio Then newTop = newBottom - (newRight - newLeft) / m_LayerAspectRatio Else newTop = curLayerY
                If ((newBottom - newTop) < 1!) Then newTop = newBottom - 1!
                
                If lockAspectRatio And Snap.GetSnap_Any() Then
                    
                    srcPtF.x = newRight
                    srcPtF.y = newTop
                    Snap.SnapPointByMoving srcPtF, snappedPtF
                    
                    If Snap.IsSnapped_X() Then
                        newRight = snappedPtF.x
                        newTop = newBottom - (newRight - newLeft) / m_LayerAspectRatio
                        Snap.NotifyNoSnapping_Y
                    ElseIf Snap.IsSnapped_Y() Then
                        newTop = snappedPtF.y
                        newRight = newLeft + (newBottom - newTop) * m_LayerAspectRatio
                        Snap.NotifyNoSnapping_X
                    End If
                    
                    If (newRight < newLeft + 1!) Then newRight = newLeft + 1!
                    If (newTop > newBottom - 1!) Then newTop = newBottom - 1!
                    
                End If
                
                srcLayer.SetOffsetsAndModifiersTogether newLeft, newTop, newRight, newBottom
                rotateCleanupRequired = True
                
            '2: bottom-left
            Case poi_CornerSW
                
                newRight = m_InitLayerCoords_Pure(1).x
                newTop = m_InitLayerCoords_Pure(0).y
                
                If ((newRight - curLayerX) > 1!) Then newLeft = curLayerX Else newLeft = newRight - 1!
                If lockAspectRatio Then newBottom = newTop + (newRight - newLeft) / m_LayerAspectRatio Else newBottom = curLayerY
                If ((newBottom - newTop) < 1!) Then newBottom = newTop + 1!
                
                If lockAspectRatio And Snap.GetSnap_Any() Then
                    
                    srcPtF.x = newLeft
                    srcPtF.y = newBottom
                    Snap.SnapPointByMoving srcPtF, snappedPtF
                    
                    If Snap.IsSnapped_X() Then
                        newLeft = snappedPtF.x
                        newBottom = newTop + (newRight - newLeft) / m_LayerAspectRatio
                        Snap.NotifyNoSnapping_Y
                    ElseIf Snap.IsSnapped_Y() Then
                        newBottom = snappedPtF.y
                        newLeft = newRight - (newBottom - newTop) * m_LayerAspectRatio
                        Snap.NotifyNoSnapping_X
                    End If
                    
                    If (newLeft > newRight - 1!) Then newLeft = newRight - 1!
                    If (newBottom < newTop + 1!) Then newBottom = newTop + 1!
                    
                End If
                
                srcLayer.SetOffsetsAndModifiersTogether newLeft, newTop, newRight, newBottom
                rotateCleanupRequired = True
                
            '3: bottom-right
            Case poi_CornerSE
                
                newLeft = m_InitLayerCoords_Pure(0).x
                newTop = m_InitLayerCoords_Pure(0).y
                
                'Finish calculating things like required minimum layer size and aspect ratio preservation
                If ((curLayerX - newLeft) > 1!) Then newRight = curLayerX Else newRight = newLeft + 1!
                If lockAspectRatio Then newBottom = newTop + (newRight - newLeft) / m_LayerAspectRatio Else newBottom = curLayerY
                If ((newBottom - newTop) < 1!) Then newBottom = newTop + 1!
                
                If lockAspectRatio And Snap.GetSnap_Any() Then
                    
                    srcPtF.x = newRight
                    srcPtF.y = newBottom
                    Snap.SnapPointByMoving srcPtF, snappedPtF
                    
                    If Snap.IsSnapped_X() Then
                        newRight = snappedPtF.x
                        newBottom = newTop + (newRight - newLeft) / m_LayerAspectRatio
                        Snap.NotifyNoSnapping_Y
                    ElseIf Snap.IsSnapped_Y() Then
                        newBottom = snappedPtF.y
                        newRight = newLeft + (newBottom - newTop) * m_LayerAspectRatio
                        Snap.NotifyNoSnapping_X
                    End If
                    
                    If (newRight < newLeft + 1!) Then newRight = newLeft + 1!
                    If (newBottom < newTop + 1!) Then newBottom = newTop + 1!
                    
                End If
                
                srcLayer.SetOffsetsAndModifiersTogether newLeft, newTop, newRight, newBottom
                rotateCleanupRequired = True
                
            '4-7: rotation nodes
            Case poi_EdgeN, poi_EdgeE, poi_EdgeS, poi_EdgeW
                
                'Disable smart guide rendering
                Snap.NotifyNoSnapping
                
                'Layer rotation is different because it involves finding the angle between two lines; specifically, the angle between
                ' a flat origin line and the current node-to-origin line of the rotation node.
                Dim ptIntersect As PointFloat, pt1 As PointFloat, pt2 As PointFloat
                
                'The intersect point is the center of the image.  This point is the same for all rotation nodes.
                ptIntersect.x = m_InitLayerCoords_Pure(0).x + (m_InitLayerCoords_Pure(3).x - m_InitLayerCoords_Pure(0).x) / 2
                ptIntersect.y = m_InitLayerCoords_Pure(0).y + (m_InitLayerCoords_Pure(3).y - m_InitLayerCoords_Pure(0).y) / 2
                
                'The first non-intersecting point varies by rotation node (as they lie in 90-degree increments).  Note that the
                ' 100 offset is totally arbitrary; we just need a line of some non-zero length for the angle calculation to work.
                If (m_CurPOI = poi_EdgeE) Then
                    pt1.x = ptIntersect.x + 100#
                    pt1.y = ptIntersect.y
                ElseIf (m_CurPOI = poi_EdgeS) Then
                    pt1.x = ptIntersect.x
                    pt1.y = ptIntersect.y + 100#
                ElseIf (m_CurPOI = poi_EdgeW) Then
                    pt1.x = ptIntersect.x - 100#
                    pt1.y = ptIntersect.y
                ElseIf (m_CurPOI = poi_EdgeN) Then
                    pt1.x = ptIntersect.x
                    pt1.y = ptIntersect.y - 100#
                End If
                                                
                'The second non-intersecting point is the current mouse position.
                pt2.x = curImageX
                pt2.y = curImageY
                
                'If shearing is active on the current layer, we need to account for its effect on the current mouse location.
                ' (Note that we could apply this matrix transformation regardless of current shear values, as values of zero
                ' will simply return an identity matrix, but why do extra math if it's not required?)
                If (srcLayer.GetLayerShearX <> 0#) Or (srcLayer.GetLayerShearY <> 0#) Then
                
                    'Apply the current layer's shear effect to the mouse position.  This gives us its unadulterated equivalent,
                    ' e.g. its location in the same coordinate space as the two points we've already calculated.
                    Dim tmpMatrix As pd2DTransform
                    Set tmpMatrix = New pd2DTransform
                    
                    tmpMatrix.ApplyShear srcLayer.GetLayerShearX, srcLayer.GetLayerShearY, ptIntersect.x, ptIntersect.y
                    tmpMatrix.InvertTransform
                    
                    tmpMatrix.ApplyTransformToPointF pt2
                
                End If
                
                'Find the angle between the two lines we've calculated
                Dim newAngle As Double
                newAngle = PDMath.AngleBetweenTwoIntersectingLines(ptIntersect, pt1, pt2, True)
                
                'Because the angle function finds the absolute inner angle, it will never be greater than 180 degrees.
                ' This also means that +90 and -90 (from a UI standpoint) return the same 90 result.  A simple workaround
                ' is to force the sign to match the difference between the relevant coordinate of the intersecting lines.
                ' (The relevant coordinate varies based on the orientation of the default, non-rotated line defined by
                ' ptIntersect and pt1.)
                If (m_CurPOI = poi_EdgeE) Then
                    If (pt2.y < pt1.y) Then newAngle = -newAngle
                ElseIf (m_CurPOI = poi_EdgeS) Then
                    If (pt2.x > pt1.x) Then newAngle = -newAngle
                ElseIf (m_CurPOI = poi_EdgeW) Then
                    If (pt2.y > pt1.y) Then newAngle = -newAngle
                ElseIf (m_CurPOI = poi_EdgeN) Then
                    If (pt2.x < pt1.x) Then newAngle = -newAngle
                End If
                
                'If the SHIFT key is down, snap the angle
                If isShiftDown Then
                    newAngle = Snap.SnapAngle_Arbitrary(newAngle, 15!)
                    
                'If the SHIFT key is *not* down, rely on the View > Snap To... menu for snap behavior
                Else
                    newAngle = Snap.SnapAngle(newAngle)
                End If
                
                'Apply the angle to the layer, and our work here is done!
                .SetLayerAngle newAngle
                
            '5: interior of the layer (e.g. move the layer instead of resize it)
            Case poi_Interior
                
                'Pass the new coordinates to the layer engine, then retrieve the new layer rect
                ' the transform produces
                .SetLayerOffsetX m_InitLayerCoords_Pure(0).x + hOffsetImage
                .SetLayerOffsetY m_InitLayerCoords_Pure(0).y + vOffsetImage
                
                'Apply snapping (contingent on user settings).
                If Snap.GetSnap_Any() Then
                    
                    Dim listOfCorners() As PointFloat
                    ReDim listOfCorners(0 To 3) As PointFloat
                    .GetLayerCornerCoordinates listOfCorners
                    
                    Dim snapOffsetX As Long, snapOffsetY As Long
                    Snap.SnapPointListByMoving listOfCorners, 4, snapOffsetX, snapOffsetY
                    
                    'Hand the layer corners off to the snap calculator, then take whatever it returns and
                    ' forward the original left/top position + snapped offsets to the source layer
                    .SetLayerOffsetX .GetLayerOffsetX + snapOffsetX
                    .SetLayerOffsetY .GetLayerOffsetY + snapOffsetY
                    
                End If
                
        End Select
        
        'If this layer is moved and/or resized while rotation is active, we need to adjust the layer's rotational center
        ' point to match the new position and/or size.
        If rotateCleanupRequired Then
        
            'The goal here is to modify the layer's point of rotation so that it remains fixed in space, even though
            ' the layer's offsets and/or dimensions are changing in real-time.  (If we allow the point of rotation to
            ' auto-calculate to the center of the image, as it does by default, the layer will "jump around" during
            ' mouse interactions because the center of the image is constantly changing due to the corresponding
            ' position/dimension changes.)
            
            'First, while the user is still moving the mouse, we want to set a new, temporary layer rotation point that is
            ' identical to the original rotation point in absolute image coordinates.  (Rotation points are stored as
            ' *ratios* inside pdLayer, e.g. [0.5, 0.5] for the center of the image, which helpfully makes them independent
            ' of current layer width/height - helpful everywhere but here, alas.  Inside this function, we must manually
            ' convert those ratios to an absolute physical coordinate, and we want to make sure that physical coordinate
            ' remains fixed during the duration of the interaction.)
            Dim adjustedWidth As Single, adjustedHeight As Single
            adjustedWidth = (newRight - newLeft)
            adjustedHeight = (newBottom - newTop)
            
            If (adjustedWidth <> 0#) Then srcLayer.SetLayerRotateCenterX (m_InitLayerRotateCenterX - newLeft) / adjustedWidth
            If (adjustedHeight <> 0#) Then srcLayer.SetLayerRotateCenterY (m_InitLayerRotateCenterY - newTop) / adjustedHeight
            
            'If the mouse has just been released, we want to reset the layer's rotational point to its default value
            ' (the center of the image).  This ensures that future move/size events behave as expected.
            '
            '(Note that resetting the rotational center point requires us to redefine the layer's offsets to match
            ' any new positions and/or dimensions the user has set via on-canvas mouse input.)
            If finalizeTransform Then
            
                'Note the layer's "proper" center of rotation, in absolute image coordinates
                Dim tmpPoints() As PointFloat
                ReDim tmpPoints(0 To 3) As PointFloat
                srcLayer.GetLayerCornerCoordinates tmpPoints
                
                Dim curCenter As PointFloat
                PDMath.FindCenterOfFloatPoints curCenter, tmpPoints
                
                'Reset the layer's center of rotation
                srcLayer.SetLayerRotateCenterX 0.5
                srcLayer.SetLayerRotateCenterY 0.5
                
                'Resetting the rotational point will cause the layer to "jump" to a new position.  Retrieve the
                ' layer's new center of rotation, in absolute coordinates.
                srcLayer.GetLayerCornerCoordinates tmpPoints
                Dim newCenter As PointFloat
                PDMath.FindCenterOfFloatPoints newCenter, tmpPoints
                
                'Apply new (x, y) layer offsets to ensure that the layer's on-screen position hasn't changed
                srcLayer.SetLayerOffsetX srcLayer.GetLayerOffsetX + (curCenter.x - newCenter.x)
                srcLayer.SetLayerOffsetY srcLayer.GetLayerOffsetY + (curCenter.y - newCenter.y)
                
            End If
                
        End If
        
    End With
    
    'Manually synchronize the new values against their on-screen UI elements
    Tools.SyncToolOptionsUIToCurrentLayer
    
    'Free the tool engine
    Tools.SetToolBusyState False
    
    'Reinstate canvas redraws
    srcCanvas.SetRedrawSuspension False
    
    'If this is the final step of a transform (e.g. if the user has just released the mouse), forward this
    ' request to PD's central processor, so an Undo/Redo entry can be generated.
    If finalizeTransform Then
        
        Dim cParams As pdSerialize
        Set cParams = New pdSerialize
        
        'As a convenience to the user, layer resize and move operations are listed separately.
        Select Case m_CurPOI
        
            'Move/resize transformations
            Case poi_CornerNW, poi_CornerNE, poi_CornerSW, poi_CornerSE
                
                With cParams
                    .AddParam "layer-offsetx", srcImage.GetActiveLayer.GetLayerOffsetX
                    .AddParam "layer-offsety", srcImage.GetActiveLayer.GetLayerOffsetY
                    
                    'Image layers need an x/y modifier pair, while vector layers need an absolute size; we store both
                    ' and let the loader sort it out later.
                    .AddParam "layer-modifierx", srcImage.GetActiveLayer.GetLayerCanvasXModifier
                    .AddParam "layer-modifiery", srcImage.GetActiveLayer.GetLayerCanvasYModifier
                    .AddParam "layer-sizex", srcImage.GetActiveLayer.GetLayerWidth
                    .AddParam "layer-sizey", srcImage.GetActiveLayer.GetLayerHeight
                    
                End With
                
                With srcImage.GetActiveLayer
                    Process "Resize layer (on-canvas)", False, cParams.GetParamString(), UNDO_LayerHeader
                End With
                
            'Rotation
            Case poi_EdgeE, poi_EdgeS, poi_EdgeW, poi_EdgeN
                
                cParams.AddParam "layer-angle", srcImage.GetActiveLayer.GetLayerAngle
                
                With srcImage.GetActiveLayer
                    Process "Rotate layer (on-canvas)", False, cParams.GetParamString(), UNDO_LayerHeader
                End With
            
            'Move-only transformations
            Case poi_Interior
            
                With cParams
                    .AddParam "layer-offsetx", srcImage.GetActiveLayer.GetLayerOffsetX
                    .AddParam "layer-offsety", srcImage.GetActiveLayer.GetLayerOffsetY
                End With
                
                'If the user is moving pixels from an active selection, a new layer was automatically created
                ' on _MouseDown.  We need to create a different type of undo entry (a full image stack) for
                ' that special case.  (On a normal move event, we're literally just setting different position
                ' flags inside a layer header - so we don't need to backup any pixel data.)
                If m_MoveSelectedPixels Then
                    Process "Move selected pixels", False, cParams.GetParamString(), UNDO_Image_VectorSafe
                Else
                    Process "Move layer", False, cParams.GetParamString(), UNDO_LayerHeader
                End If
                
        End Select
    
    'If the transformation is still active (e.g. the user has the mouse pressed down), just redraw the viewport, but don't
    ' process Undo/Redo or any macro stuff.
    Else
    
        'Manually request a canvas redraw
        Dim tmpViewportParams As PD_ViewportParams
        tmpViewportParams = Viewport.GetDefaultParamObject()
        tmpViewportParams.curPOI = m_CurPOI
        Viewport.Stage2_CompositeAllLayers srcImage, srcCanvas, VarPtr(tmpViewportParams)
    
    End If
    
End Sub

'Are on-canvas tools currently allowed?
' This central function evaluates all relevant program states related to on-canvas tool operations
' (e.g. "no open images", "main form locked against input").
Public Function CanvasToolsAllowed(Optional ByVal alsoCheckBusyState As Boolean = True, Optional ByVal checkMainWindowEnabled As Boolean = True) As Boolean

    CanvasToolsAllowed = False
    
    'Make sure an image is loaded and active
    If PDImages.IsImageActive() Then
    
        'Make sure the main form has not been disabled by a modal dialog
        If FormMain.Enabled Or (Not checkMainWindowEnabled) Then
            
            'Finally, make sure another process hasn't locked the active canvas.  Note that the caller can disable this behavior
            ' if they don't need it.
            If alsoCheckBusyState Then
                CanvasToolsAllowed = (Not Processor.IsProgramBusy) And (Not Tools.GetToolBusyState)
            Else
                CanvasToolsAllowed = True
            End If
            
        End If
    
    End If
    
End Function

'Some tools (paintbrushes, most notably), have to initialize themselves against the current image
' (prepping a scratch layer, for example).  To reduce stuttering on first tool use, we initialize
'  this behavior whenever...
' 1) Such a tool is selected, or...
' 2) The tool is already selected and the user switches images, or...
' 3) The active image's size changes
Public Sub InitializeToolsDependentOnImage(Optional ByVal activeImageChanged As Boolean = False)
    
    If PDImages.IsImageActive() Then
        
        'The measurement tool has two settings: it can either share measurements across images
        ' (great for unifying measurements), or it can allow each image to have its own measurement.
        ' What we do when changing images depends on this setting.
        If (g_CurrentTool = ND_MEASURE) Then toolpanel_Measure.NotifyActiveImageChanged
        
        'An active crop region (if any) can be moved between images, but max/min values on the toolpanel's
        ' spin controls need to change to match.
        If (g_CurrentTool = ND_CROP) Then toolpanel_Crop.NotifyActiveImageChanged
        
        'Paint tools are handled as a special case
        Dim toolIsPaint As Boolean
        toolIsPaint = (g_CurrentTool = PAINT_PENCIL) Or (g_CurrentTool = PAINT_SOFTBRUSH)
        If (Not toolIsPaint) Then toolIsPaint = (g_CurrentTool = PAINT_ERASER) Or (g_CurrentTool = PAINT_CLONE)
        If (Not toolIsPaint) Then toolIsPaint = (g_CurrentTool = PAINT_FILL) Or (g_CurrentTool = PAINT_GRADIENT)
        
        If toolIsPaint Then
            
            'A couple things require us to reset the scratch layer...
            ' 1) If it hasn't been initialized at all
            ' 2) If it doesn't match the current image's size
            Dim scratchLayerResetRequired As Boolean: scratchLayerResetRequired = False
            scratchLayerResetRequired = (PDImages.GetActiveImage.ScratchLayer Is Nothing)
            If (Not scratchLayerResetRequired) Then
                scratchLayerResetRequired = (PDImages.GetActiveImage.ScratchLayer.GetLayerWidth <> PDImages.GetActiveImage.Width)
                If (Not scratchLayerResetRequired) Then scratchLayerResetRequired = (PDImages.GetActiveImage.ScratchLayer.GetLayerHeight <> PDImages.GetActiveImage.Height)
            End If
            
            If scratchLayerResetRequired Then PDImages.GetActiveImage.ResetScratchLayer True
            
            'If the active image has changed, or the image state has changed enough to warrant
            ' creating a new scratch layer, we also need to reset some other paint tool parameters
            ' (such as last stroke position tracking)
            If activeImageChanged Or scratchLayerResetRequired Then
                If (g_CurrentTool = PAINT_PENCIL) Then
                    Tools_Pencil.NotifyActiveImageChanged
                ElseIf (g_CurrentTool = PAINT_SOFTBRUSH) Or (g_CurrentTool = PAINT_ERASER) Then
                    Tools_Paint.NotifyActiveImageChanged
                ElseIf (g_CurrentTool = PAINT_CLONE) Then
                    Tools_Clone.NotifyActiveImageChanged
                End If
            End If
            
        Else
            
            'The scratch layer is not required for non-paint tools, and releasing it frees a lot of memory
            Set PDImages.GetActiveImage.ScratchLayer = Nothing
            
            'As a failsafe, restore default mouse settings, which may have been modified by various paintbrush tools
            FormMain.MainCanvas(0).SetMouseInput_AutoDrop True
            FormMain.MainCanvas(0).SetMouseInput_HighRes False
            
        End If
        
    End If
    
End Sub

'When image (or layer) size changes, call this function.
'
'Some tools (like Measure or the Clone Brush) rely on saved image coordinates.  When image size changes,
' those coordinates may no longer be valid.
Public Sub NotifyImageSizeChanged()
    If (g_CurrentTool = ND_MEASURE) Then
        Tools_Measure.ResetPoints True
    ElseIf (g_CurrentTool = ND_CROP) Then
        If PDImages.IsImageActive() Then
            toolpanel_Crop.NotifyActiveImageChanged
        Else
            Tools_Crop.RemoveCurrentCrop
        End If
    ElseIf (g_CurrentTool = PAINT_CLONE) Then
        Tools_Clone.NotifyImageSizeChanged
    End If
End Sub

'When the active layer changes, call this function.  It synchronizes various layer-specific tool panels against the
' currently active layer.  (Note that you also need to call this whenever a new tool panel is selected, as the newly
' loaded panel will reflect default values otherwise.)
Public Sub SyncToolOptionsUIToCurrentLayer()
    
    'Failsafe checks
    If (Not PDImages.IsImageActive()) Then Exit Sub
    If (PDImages.GetActiveImage.GetActiveLayer Is Nothing) Then Exit Sub
    
    'Before doing anything else, make sure canvas tool operations are allowed.  (They are disallowed if no images
    ' are loaded, for example.)
    If (Not CanvasToolsAllowed(False, False)) Then
        
        'Some panels may still wish to redraw their contents, even if no images are loaded.  (Text panels use this
        ' opportunity to hide the "convert typography to text or vice-versa" panels that are visible by default.)
        If (g_CurrentTool = TEXT_BASIC) Then
            toolpanel_TextBasic.UpdateAgainstCurrentLayer
        ElseIf (g_CurrentTool = TEXT_ADVANCED) Then
            toolpanel_TextAdvanced.UpdateAgainstCurrentLayer
        End If
        
        'Exit now, as subsequent checks in this function require one or more active images
        Exit Sub
        
    End If
    
    'Next, figure out if the current tool is a type that requires syncing.  (Some tools, like paintbrushes, don't need
    ' to be synched against layer changes.  Others, like the move/size tool, obviously do.)
    Dim layerToolActive As Boolean
    
    Select Case g_CurrentTool
        
        Case NAV_MOVE
            layerToolActive = True
            
        Case COLOR_PICKER
            layerToolActive = True
        
        'Text layers only require a sync if the current layer is a text layer.
        Case TEXT_BASIC, TEXT_ADVANCED
            If PDImages.GetActiveImage.GetActiveLayer.IsLayerText Then
                layerToolActive = True
            Else
            
                'Hide the "convert to different type of text" panel prompts
                If (g_CurrentTool = TEXT_BASIC) Then
                    toolpanel_TextBasic.UpdateAgainstCurrentLayer
                ElseIf (g_CurrentTool = TEXT_ADVANCED) Then
                    toolpanel_TextAdvanced.UpdateAgainstCurrentLayer
                End If
            
            End If
        
        Case Else
            layerToolActive = False
        
    End Select
    
    'To improve performance, only continue with a UI sync if a layer-specific tool is active, and the tool options
    ' panel is visible.  (The user can choose to disable this panel... though why they would, I don't know.)
    If (Not toolbar_Options.Visible) And (Not layerToolActive) Then Exit Sub
    
    If layerToolActive Then
        
        'Mark the tool engine as busy; this prevents things like changing control values from triggering automatic
        ' viewport redraws.
        Tools.SetToolBusyState True
        
        'Start iterating various layer properties, and reflecting them across their corresponding UI elements.
        ' (Obviously, this step is separated by tool type.)
        Select Case g_CurrentTool
        
            Case NAV_MOVE
                
                'The interface module actually has nice functions that handle this
                Interface.SetUIGroupState PDUI_LayerTools, True
                Interface.SyncUI_CurrentLayerSettings
                
                'Reset tool busy state (because it will be reset by the Interface module call, above)
                Tools.SetToolBusyState True
                
            Case TEXT_BASIC
                
                With toolpanel_TextBasic
                    .txtTextTool.Text = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_Text)
                    .cboTextFontFace.ListIndex = .cboTextFontFace.ListIndexByString(PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_FontFace), vbTextCompare)
                    .sldTextFontSize.Value = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_FontSize)
                    .csTextFontColor.Color = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_FontColor)
                    .cboTextRenderingHint.ListIndex = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_TextAntialiasing)
                    .sltTextClarity.Value = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_TextContrast)
                    .btnFontStyles(0).Value = CBool(PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_FontBold))
                    .btnFontStyles(1).Value = CBool(PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_FontItalic))
                    .btnFontStyles(2).Value = CBool(PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_FontUnderline))
                    .btnFontStyles(3).Value = CBool(PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_FontStrikeout))
                    .btsHAlignment.ListIndex = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_HorizontalAlignment)
                    .btsVAlignment.ListIndex = PDImages.GetActiveImage.GetActiveLayer.GetTextLayerProperty(ptp_VerticalAlignment)
                End With
                
                'Display the "convert to basic text layer" panel as necessary
                toolpanel_TextBasic.UpdateAgainstCurrentLayer
                
            Case TEXT_ADVANCED
                
                'Either sync all UI objects against the current layer's text settings, or display a
                ' "convert this basic text layer to an advanced text layer" prompt.
                toolpanel_TextAdvanced.SyncSettingsToCurrentLayer
                toolpanel_TextAdvanced.UpdateAgainstCurrentLayer
        
        End Select
        
        'Free the tool engine
        Tools.SetToolBusyState False
    
    End If
    
End Sub

'this function is the reverse of syncToolOptionsUIToCurrentLayer(), above.  If you want to copy all current UI settings into
' the currently active layer, call this function.
Public Sub SyncCurrentLayerToToolOptionsUI()
    
    'Before doing anything else, make sure canvas tool operations are allowed
    If (Not CanvasToolsAllowed(False)) Then Exit Sub
    
    'To improve performance, we'll only sync the UI if a layer-specific tool is active, and the tool options panel is currently
    ' set to VISIBLE.
    If (Not toolbar_Options.Visible) Then Exit Sub
    
    Dim layerToolActive As Boolean
    
    Select Case g_CurrentTool
        
        Case NAV_MOVE
            layerToolActive = True
        
        Case TEXT_BASIC, TEXT_ADVANCED
            If PDImages.GetActiveImage.GetActiveLayer.IsLayerText Then layerToolActive = True
        
        Case Else
            layerToolActive = False
        
    End Select
    
    If layerToolActive Then
        
        'Mark the tool engine as busy; this prevents each change from triggering viewport redraws
        Tools.SetToolBusyState True
        
        'Start iterating various layer properties, and reflecting them across their corresponding UI elements.
        ' (Obviously, this step is separated by tool type.)
        Select Case g_CurrentTool
        
            Case NAV_MOVE
            
                'The Layer Move tool has four text up/downs: two for layer position (x, y) and two for layer size (w, y)
                PDImages.GetActiveImage.GetActiveLayer.SetLayerOffsetX toolpanel_MoveSize.tudLayerMove(0).Value
                PDImages.GetActiveImage.GetActiveLayer.SetLayerOffsetY toolpanel_MoveSize.tudLayerMove(1).Value
                
                'Setting layer width and height isn't activated at present, on purpose
                'PDImages.GetActiveImage.getActiveLayer.setLayerWidth toolpanel_MoveSize.tudLayerMove(2).Value
                'PDImages.GetActiveImage.getActiveLayer.setLayerHeight toolpanel_MoveSize.tudLayerMove(3).Value
                
                'The layer resize quality combo box also needs to be synched
                PDImages.GetActiveImage.GetActiveLayer.SetLayerResizeQuality toolpanel_MoveSize.cboLayerResizeQuality.ListIndex
                
                'Layer angle and shear are newly available as of 7.0
                PDImages.GetActiveImage.GetActiveLayer.SetLayerAngle toolpanel_MoveSize.sltLayerAngle.Value
                PDImages.GetActiveImage.GetActiveLayer.SetLayerShearX toolpanel_MoveSize.sltLayerShearX.Value
                PDImages.GetActiveImage.GetActiveLayer.SetLayerShearY toolpanel_MoveSize.sltLayerShearY.Value
            
            Case TEXT_BASIC
                
                With PDImages.GetActiveImage.GetActiveLayer
                    .SetTextLayerProperty ptp_Text, toolpanel_TextBasic.txtTextTool.Text
                    .SetTextLayerProperty ptp_FontFace, toolpanel_TextBasic.cboTextFontFace.List(toolpanel_TextBasic.cboTextFontFace.ListIndex)
                    .SetTextLayerProperty ptp_FontSize, toolpanel_TextBasic.sldTextFontSize.Value
                    .SetTextLayerProperty ptp_FontColor, toolpanel_TextBasic.csTextFontColor.Color
                    .SetTextLayerProperty ptp_TextAntialiasing, toolpanel_TextBasic.cboTextRenderingHint.ListIndex
                    .SetTextLayerProperty ptp_TextContrast, toolpanel_TextBasic.sltTextClarity.Value
                    .SetTextLayerProperty ptp_FontBold, toolpanel_TextBasic.btnFontStyles(0).Value
                    .SetTextLayerProperty ptp_FontItalic, toolpanel_TextBasic.btnFontStyles(1).Value
                    .SetTextLayerProperty ptp_FontUnderline, toolpanel_TextBasic.btnFontStyles(2).Value
                    .SetTextLayerProperty ptp_FontStrikeout, toolpanel_TextBasic.btnFontStyles(3).Value
                    .SetTextLayerProperty ptp_HorizontalAlignment, toolpanel_TextBasic.btsHAlignment.ListIndex
                    .SetTextLayerProperty ptp_VerticalAlignment, toolpanel_TextBasic.btsVAlignment.ListIndex
                End With
                
                'This is a little weird, but we also make sure to synchronize the current text rendering engine when
                ' the UI is synched, because this property changes according to the type of text layer.
                ' (Basic text layers are rendered using the built-in Windows text renderer.)
                PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_RenderingEngine, te_WAPI
            
            Case TEXT_ADVANCED
                
                With PDImages.GetActiveImage.GetActiveLayer
                    .SetTextLayerProperty ptp_Text, toolpanel_TextAdvanced.txtTextTool.Text
                    .SetTextLayerProperty ptp_FontFace, toolpanel_TextAdvanced.cboTextFontFace.List(toolpanel_TextAdvanced.cboTextFontFace.ListIndex)
                    .SetTextLayerProperty ptp_FontSize, toolpanel_TextAdvanced.sldTextFontSize.Value
                    .SetTextLayerProperty ptp_StretchToFit, toolpanel_TextAdvanced.btsStretch.ListIndex
                    .SetTextLayerProperty ptp_TextAntialiasing, toolpanel_TextAdvanced.cboTextRenderingHint.ListIndex
                    .SetTextLayerProperty ptp_TextHinting, (toolpanel_TextAdvanced.btsHinting.ListIndex = 1)
                    .SetTextLayerProperty ptp_FontBold, toolpanel_TextAdvanced.btnFontStyles(0).Value
                    .SetTextLayerProperty ptp_FontItalic, toolpanel_TextAdvanced.btnFontStyles(1).Value
                    .SetTextLayerProperty ptp_FontUnderline, toolpanel_TextAdvanced.btnFontStyles(2).Value
                    .SetTextLayerProperty ptp_FontStrikeout, toolpanel_TextAdvanced.btnFontStyles(3).Value
                    .SetTextLayerProperty ptp_HorizontalAlignment, toolpanel_TextAdvanced.btsHAlignment.ListIndex
                    .SetTextLayerProperty ptp_VerticalAlignment, toolpanel_TextAdvanced.btsVAlignment.ListIndex
                    .SetTextLayerProperty ptp_WordWrap, toolpanel_TextAdvanced.cboWordWrap.ListIndex
                    .SetTextLayerProperty ptp_FillActive, toolpanel_TextAdvanced.chkFillText.Value
                    .SetTextLayerProperty ptp_FillBrush, toolpanel_TextAdvanced.bsText.Brush
                    .SetTextLayerProperty ptp_OutlineActive, toolpanel_TextAdvanced.chkOutlineText.Value
                    .SetTextLayerProperty ptp_OutlinePen, toolpanel_TextAdvanced.psText.Pen
                    .SetTextLayerProperty ptp_BackgroundActive, toolpanel_TextAdvanced.chkBackground.Value
                    .SetTextLayerProperty ptp_BackgroundBrush, toolpanel_TextAdvanced.bsTextBackground.Brush
                    .SetTextLayerProperty ptp_BackBorderActive, toolpanel_TextAdvanced.chkBackgroundBorder.Value
                    .SetTextLayerProperty ptp_BackBorderPen, toolpanel_TextAdvanced.psTextBackground.Pen
                    .SetTextLayerProperty ptp_LineSpacing, toolpanel_TextAdvanced.sldLineSpacing.Value
                    .SetTextLayerProperty ptp_MarginLeft, toolpanel_TextAdvanced.tudMargin(0).Value
                    .SetTextLayerProperty ptp_MarginRight, toolpanel_TextAdvanced.tudMargin(1).Value
                    .SetTextLayerProperty ptp_MarginTop, toolpanel_TextAdvanced.tudMargin(2).Value
                    .SetTextLayerProperty ptp_MarginBottom, toolpanel_TextAdvanced.tudMargin(3).Value
                    .SetTextLayerProperty ptp_CharInflation, toolpanel_TextAdvanced.sltCharInflation.Value
                    .SetTextLayerProperty ptp_CharJitterX, toolpanel_TextAdvanced.tudJitter(0).Value
                    .SetTextLayerProperty ptp_CharJitterY, toolpanel_TextAdvanced.tudJitter(1).Value
                    .SetTextLayerProperty ptp_CharMirror, toolpanel_TextAdvanced.cboCharMirror.ListIndex
                    .SetTextLayerProperty ptp_CharOrientation, toolpanel_TextAdvanced.sltCharOrientation.Value
                    .SetTextLayerProperty ptp_CharRemap, toolpanel_TextAdvanced.cboCharCase.ListIndex
                    .SetTextLayerProperty ptp_CharSpacing, toolpanel_TextAdvanced.sltCharSpacing.Value
                    .SetTextLayerProperty ptp_AlignLastLine, toolpanel_TextAdvanced.btsHAlignJustify.ListIndex
                    .SetTextLayerProperty ptp_OutlineAboveFill, toolpanel_TextAdvanced.chkFillFirst.Value
                End With
                
                'Advanced text layers are rendered using a PhotoDemon-specific renderer.
                PDImages.GetActiveImage.GetActiveLayer.SetTextLayerProperty ptp_RenderingEngine, te_PhotoDemon
        
        End Select
        
        'Free the tool engine
        Tools.SetToolBusyState False
    
    End If
    
End Sub

'Some tool-related user preferences are cached locally, to improve performance (vs pulling them from the
' central resource manager).  You may interact with these settings via the following safe wrapper functions.
Public Function GetToolSetting_HighResMouse() As Boolean
    GetToolSetting_HighResMouse = m_HighResMouseInputAllowed
End Function

Public Sub SetToolSetting_HighResMouse(ByVal newSetting As Boolean)
    m_HighResMouseInputAllowed = newSetting
End Sub

'Debug helper only; useful for logging tool-specific data in a human-readable way
Public Function GetNameOfTool(ByVal toolIndex As PDTools) As String

    Select Case toolIndex
        Case NAV_DRAG
            GetNameOfTool = "Hand"
        Case NAV_ZOOM
            GetNameOfTool = "Zoom"
        Case NAV_MOVE
            GetNameOfTool = "Move"
        Case COLOR_PICKER
            GetNameOfTool = "Color picker"
        Case ND_MEASURE
            GetNameOfTool = "Measure"
        Case ND_CROP
            GetNameOfTool = "Crop"
        Case SELECT_RECT
            GetNameOfTool = "Rectangle selection"
        Case SELECT_CIRC
            GetNameOfTool = "Circle selection"
        Case SELECT_POLYGON
            GetNameOfTool = "Polygon selection"
        Case SELECT_LASSO
            GetNameOfTool = "Lasso selection"
        Case SELECT_WAND
            GetNameOfTool = "Magic wand selection"
        Case TEXT_BASIC
            GetNameOfTool = "Basic text"
        Case TEXT_ADVANCED
            GetNameOfTool = "Advanced text"
        Case PAINT_PENCIL
            GetNameOfTool = "Pencil"
        Case PAINT_SOFTBRUSH
            GetNameOfTool = "Paintbrush"
        Case PAINT_ERASER
            GetNameOfTool = "Eraser"
        Case PAINT_CLONE
            GetNameOfTool = "Clone brush"
        Case PAINT_FILL
            GetNameOfTool = "Paint bucket"
        Case PAINT_GRADIENT
            GetNameOfTool = "Gradient"
    End Select

End Function

'Some generic tool-related actions are implemented here.  These are (typically) activated via hotkey,
' and they are designed to work on *any* relevant tool.  (For example, increasing "brush size" works
' on any brush-like tool.)

'Hardness changes at 25% increments, like PS.
Public Sub QuickToolAction_HardnessDown()
    
    'Before adjusting anything, ensure a relevant tool is active
    Dim curValue As Double
    
    Select Case g_CurrentTool
        Case PAINT_SOFTBRUSH
            curValue = toolpanel_Paintbrush.sltBrushSetting(2).Value
        Case PAINT_ERASER
            curValue = toolpanel_Eraser.sltBrushSetting(2).Value
        Case PAINT_CLONE
            curValue = toolpanel_Clone.sltBrushSetting(2).Value
        Case Else
            Exit Sub
    End Select
    
    'Lock to the nearest multiple of 25
    curValue = Int((curValue + 24.99) / 25) * 25 - 25
    
    'Ensure valid minimum
    If (curValue < 1) Then curValue = 1
    
    'Assign the new value
    Select Case g_CurrentTool
        Case PAINT_SOFTBRUSH
            toolpanel_Paintbrush.sltBrushSetting(2).Value = curValue
        Case PAINT_ERASER
            toolpanel_Eraser.sltBrushSetting(2).Value = curValue
        Case PAINT_CLONE
            toolpanel_Clone.sltBrushSetting(2).Value = curValue
    End Select
    
End Sub

Public Sub QuickToolAction_HardnessUp()

    'Before adjusting anything, ensure a relevant tool is active
    Dim curValue As Double
    
    Select Case g_CurrentTool
        Case PAINT_SOFTBRUSH
            curValue = toolpanel_Paintbrush.sltBrushSetting(2).Value
        Case PAINT_ERASER
            curValue = toolpanel_Eraser.sltBrushSetting(2).Value
        Case PAINT_CLONE
            curValue = toolpanel_Clone.sltBrushSetting(2).Value
        Case Else
            Exit Sub
    End Select
    
    'Lock to the nearest multiple of 25
    curValue = (Int(curValue / 25) + 1) * 25
    
    'Ensure valid maximum
    If (curValue > 100) Then curValue = 100
    
    'Assign the new value
    Select Case g_CurrentTool
        Case PAINT_SOFTBRUSH
            toolpanel_Paintbrush.sltBrushSetting(2).Value = curValue
        Case PAINT_ERASER
            toolpanel_Eraser.sltBrushSetting(2).Value = curValue
        Case PAINT_CLONE
            toolpanel_Clone.sltBrushSetting(2).Value = curValue
    End Select
    
End Sub

'Size changes are more complex.  How much we increase or decrease size varies based on the current size;
' increments generally increase proportional to brush size.
Public Sub QuickToolAction_SizeDown()
    
    'Before adjusting anything, ensure a relevant tool is active
    Dim curValue As Double
    
    Select Case g_CurrentTool
        Case PAINT_PENCIL
            curValue = toolpanel_Pencil.sltBrushSetting(0).Value
        Case PAINT_SOFTBRUSH
            curValue = toolpanel_Paintbrush.sltBrushSetting(0).Value
        Case PAINT_ERASER
            curValue = toolpanel_Eraser.sltBrushSetting(0).Value
        Case PAINT_CLONE
            curValue = toolpanel_Clone.sltBrushSetting(0).Value
        Case Else
            Exit Sub
    End Select
    
    'Size changes vary by range.  This pattern is a direct copy of Photoshop CS2's strategy, for better or worse.
    If (curValue > 300) Then
        curValue = Int((curValue + 99.99) / 100) * 100 - 100
    ElseIf (curValue > 200) Then
        curValue = Int((curValue + 49.99) / 50) * 50 - 50
    ElseIf (curValue > 100) Then
        curValue = Int((curValue + 24.99) / 25) * 25 - 25
    ElseIf (curValue > 10) Then
        curValue = Int((curValue + 9.99) / 10) * 10 - 10
    Else
        curValue = Int(curValue + 0.99) - 1
    End If
    
    'Ensure valid minimum
    If (curValue < 1#) Then curValue = 1#
    
    'Assign the new value
    Select Case g_CurrentTool
        Case PAINT_PENCIL
            toolpanel_Pencil.sltBrushSetting(0).Value = curValue
        Case PAINT_SOFTBRUSH
            toolpanel_Paintbrush.sltBrushSetting(0).Value = curValue
        Case PAINT_ERASER
            toolpanel_Eraser.sltBrushSetting(0).Value = curValue
        Case PAINT_CLONE
            toolpanel_Clone.sltBrushSetting(0).Value = curValue
    End Select
    
End Sub

Public Sub QuickToolAction_SizeUp()

    'Before adjusting anything, ensure a relevant tool is active
    Dim curValue As Double
    
    Select Case g_CurrentTool
        Case PAINT_PENCIL
            curValue = toolpanel_Pencil.sltBrushSetting(0).Value
        Case PAINT_SOFTBRUSH
            curValue = toolpanel_Paintbrush.sltBrushSetting(0).Value
        Case PAINT_ERASER
            curValue = toolpanel_Eraser.sltBrushSetting(0).Value
        Case PAINT_CLONE
            curValue = toolpanel_Clone.sltBrushSetting(0).Value
        Case Else
            Exit Sub
    End Select
    
    'Size changes vary by range.  This pattern is a direct copy of Photoshop CS2's strategy, for better or worse.
    If (curValue >= 300) Then
        curValue = Int(curValue / 100) * 100 + 100
    ElseIf (curValue >= 200) Then
        curValue = Int(curValue / 50) * 50 + 50
    ElseIf (curValue >= 100) Then
        curValue = Int(curValue / 25) * 25 + 25
    ElseIf (curValue >= 10) Then
        curValue = Int(curValue / 10) * 10 + 10
    Else
        curValue = Int(curValue) + 1
    End If
    
    'Ensure valid maximum (may vary by tool)
    Select Case g_CurrentTool
        Case PAINT_PENCIL
            If (curValue > toolpanel_Pencil.sltBrushSetting(0).Max) Then curValue = toolpanel_Pencil.sltBrushSetting(0).Max
        Case PAINT_SOFTBRUSH
            If (curValue > toolpanel_Paintbrush.sltBrushSetting(0).Max) Then curValue = toolpanel_Paintbrush.sltBrushSetting(0).Max
        Case PAINT_ERASER
            If (curValue > toolpanel_Eraser.sltBrushSetting(0).Max) Then curValue = toolpanel_Eraser.sltBrushSetting(0).Max
        Case PAINT_CLONE
            If (curValue > toolpanel_Clone.sltBrushSetting(0).Max) Then curValue = toolpanel_Clone.sltBrushSetting(0).Max
        Case Else
            Exit Sub
    End Select
    
    'Assign the new value
    Select Case g_CurrentTool
        Case PAINT_PENCIL
            toolpanel_Pencil.sltBrushSetting(0).Value = curValue
        Case PAINT_SOFTBRUSH
            toolpanel_Paintbrush.sltBrushSetting(0).Value = curValue
        Case PAINT_ERASER
            toolpanel_Eraser.sltBrushSetting(0).Value = curValue
        Case PAINT_CLONE
            toolpanel_Clone.sltBrushSetting(0).Value = curValue
    End Select
    
End Sub
