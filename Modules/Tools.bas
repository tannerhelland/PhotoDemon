Attribute VB_Name = "Tools"
'***************************************************************************
'Helper functions for various PhotoDemon tools
'Copyright 2014-2018 by Tanner Helland
'Created: 06/February/14
'Last updated: 25/June/14
'Last update: add new makeQuickFixesPermanent() function
'
'To keep the pdCanvas user control codebase lean, many of its MouseMove events redirect here, to specialized
' functions that take mouse actions on the canvas and translate them into tool actions.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
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

Public Function GetToolBusyState() As Boolean
    GetToolBusyState = m_ToolIsBusy
End Function

Public Sub SetToolBusyState(ByVal newState As Boolean)
    m_ToolIsBusy = newState
End Sub

Public Function GetCustomToolState() As Long
    GetCustomToolState = m_CustomToolMarker
    m_CustomToolMarker = 0
End Function

Public Sub SetCustomToolState(ByVal newState As Long)
    m_CustomToolMarker = newState
End Sub

Public Function IsSelectionToolActive() As Boolean
    IsSelectionToolActive = (g_CurrentTool = SELECT_CIRC) Or (g_CurrentTool = SELECT_LASSO) Or (g_CurrentTool = SELECT_LINE) Or (g_CurrentTool = SELECT_POLYGON) Or (g_CurrentTool = SELECT_RECT) Or (g_CurrentTool = SELECT_WAND)
End Function

'When a tool is finished processing, it can call this function to release all tool tracking variables
Public Sub TerminateGenericToolTracking()
    
    'Reset the current POI, if any
    m_CurPOI = poi_Undefined
    
End Sub

'The move tool uses this function to set various initial parameters for layer interactions.
Public Sub SetInitialLayerToolValues(ByRef srcImage As pdImage, ByRef srcLayer As pdLayer, ByVal mouseX_ImageSpace As Double, ByVal mouseY_ImageSpace As Double, Optional ByVal relevantPOI As PD_PointOfInterest = poi_Undefined)
    
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
    m_InitHScroll = srcCanvas.GetScrollValue(PD_HORIZONTAL)
    m_InitVScroll = srcCanvas.GetScrollValue(PD_VERTICAL)
End Sub

'The drag-to-pan tool uses this function to actually scroll the viewport area
Public Sub PanImageCanvas(ByVal initX As Long, ByVal initY As Long, ByVal curX As Long, ByVal curY As Long, ByRef srcImage As pdImage, ByRef srcCanvas As pdCanvas)

    'Prevent the canvas from redrawing itself until our pan operation is complete.  (This prevents juddery movement.)
    srcCanvas.SetRedrawSuspension True
    
    'Sub-pixel panning is now allowed (because we're awesome like that)
    Dim zoomRatio As Double
    zoomRatio = g_Zoom.GetZoomValue(srcImage.GetZoom)
    
    'Calculate new scroll values
    Dim hOffset As Long, vOffset As Long
    hOffset = (initX - curX) / zoomRatio
    vOffset = (initY - curY) / zoomRatio
        
    'Factor in the initial scroll bar values
    hOffset = m_InitHScroll + hOffset
    vOffset = m_InitVScroll + vOffset
        
    'If these values lie within the bounds of their respective scroll bar(s), apply 'em
    If (hOffset < srcCanvas.GetScrollMin(PD_HORIZONTAL)) Then
        srcCanvas.SetScrollValue PD_HORIZONTAL, srcCanvas.GetScrollMin(PD_HORIZONTAL)
    ElseIf (hOffset > srcCanvas.GetScrollMax(PD_HORIZONTAL)) Then
        srcCanvas.SetScrollValue PD_HORIZONTAL, srcCanvas.GetScrollMax(PD_HORIZONTAL)
    Else
        srcCanvas.SetScrollValue PD_HORIZONTAL, hOffset
    End If
    
    If (vOffset < srcCanvas.GetScrollMin(PD_VERTICAL)) Then
        srcCanvas.SetScrollValue PD_VERTICAL, srcCanvas.GetScrollMin(PD_VERTICAL)
    ElseIf (vOffset > srcCanvas.GetScrollMax(PD_VERTICAL)) Then
        srcCanvas.SetScrollValue PD_VERTICAL, srcCanvas.GetScrollMax(PD_VERTICAL)
    Else
        srcCanvas.SetScrollValue PD_VERTICAL, vOffset
    End If
    
    'Reinstate canvas redraws
    srcCanvas.SetRedrawSuspension False
    
    'Request the scroll-specific viewport pipeline stage
    ViewportEngine.Stage2_CompositeAllLayers srcImage, FormMain.MainCanvas(0)
    
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
    Dim newLeft As Double, newTop As Double, newRight As Double, newBottom As Double
    
    'The way we assign new offsets and/or sizes to the layer depends on the POI (point of interest) the user is interacting with.
    ' Layers currently support nine points of interest: each of their 4 corners, 4 rotational points (lying on the center of
    ' each edge), and anywhere in the layer interior (for moving the layer).
    
    'If this layer has an active rotation transform (e.g. srcLayer.GetLayerAngle <> 0), we may need to modify the layer's
    ' rotational center to compensate for positional and width/height changes.  This is only necessary for move/resize events,
    ' *not* rotation events (which is confusing, I know, but rotation events use a fixed rotation point).
    Dim rotateCleanupRequired As Boolean
    rotateCleanupRequired = False
    
    'Check the POI we were given, and update the layer accordingly.
    With srcLayer
    
        Select Case m_CurPOI
            
            '-1: the mouse is not over the layer.  Do nothing.
            Case poi_Undefined
                Tools.SetToolBusyState False
                srcCanvas.SetRedrawSuspension False
                Exit Sub
                
            '0: the mouse is dragging the top-left corner of the layer.  The comments here are uniform for all POIs, so for brevity's sake,
            ' I'll keep the others short.
            Case poi_CornerNW
                
                'The opposite corner coordinate (bottom-left) stays in exactly the same place
                newRight = m_InitLayerCoords_Pure(1).x
                newBottom = m_InitLayerCoords_Pure(3).y
                
                'Set the new left/top position to match the mouse coordinates, while also accounting for the shift key
                ' (which locks the current aspect ratio).
                If ((newRight - curLayerX) > 1#) Then newLeft = curLayerX Else newLeft = newRight - 1#
                If isShiftDown Then newTop = newBottom - (newRight - newLeft) / m_LayerAspectRatio Else newTop = curLayerY
                If ((newBottom - newTop) < 1#) Then newTop = newBottom - 1#
                
                'Immediately relay the new coordinates to the source layer
                srcLayer.SetOffsetsAndModifiersTogether newLeft, newTop, newRight, newBottom
                
                'A helper block at the end of this function cleans up any rotation-related parameters to match the new coordinate
                rotateCleanupRequired = True
                
            '1: top-right corner
            Case poi_CornerNE
            
                newLeft = m_InitLayerCoords_Pure(0).x
                newBottom = m_InitLayerCoords_Pure(2).y
                
                If ((curLayerX - newLeft) > 1#) Then newRight = curLayerX Else newRight = newLeft + 1#
                If isShiftDown Then newTop = newBottom - (newRight - newLeft) / m_LayerAspectRatio Else newTop = curLayerY
                If ((newBottom - newTop) < 1#) Then newTop = newBottom - 1#
                
                srcLayer.SetOffsetsAndModifiersTogether newLeft, newTop, newRight, newBottom
                rotateCleanupRequired = True
                
            '2: bottom-left
            Case poi_CornerSW
                
                newRight = m_InitLayerCoords_Pure(1).x
                newTop = m_InitLayerCoords_Pure(0).y
                
                If ((newRight - curLayerX) > 1#) Then newLeft = curLayerX Else newLeft = newRight - 1#
                If isShiftDown Then newBottom = (newRight - newLeft) / m_LayerAspectRatio Else newBottom = curLayerY
                If ((newBottom - newTop) < 1#) Then newBottom = newTop + 1#
                
                srcLayer.SetOffsetsAndModifiersTogether newLeft, newTop, newRight, newBottom
                rotateCleanupRequired = True
                
            '3: bottom-right
            Case poi_CornerSE
                
                newLeft = m_InitLayerCoords_Pure(0).x
                newTop = m_InitLayerCoords_Pure(0).y
                
                If ((curLayerX - newLeft) > 1#) Then newRight = curLayerX Else newRight = newLeft + 1#
                If isShiftDown Then newBottom = (newRight - newLeft) / m_LayerAspectRatio Else newBottom = curLayerY
                If ((newBottom - newTop) < 1#) Then newBottom = newTop + 1#
                
                srcLayer.SetOffsetsAndModifiersTogether newLeft, newTop, newRight, newBottom
                rotateCleanupRequired = True
                
            '4-7: rotation nodes
            Case poi_EdgeN, poi_EdgeE, poi_EdgeS, poi_EdgeW
            
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
                
                'Because the angle function finds the absolute inner angle, it will never be greater than 180 degrees.  This also means
                ' that +90 and -90 (from a UI standpoint) return the same 90 result.  A simple workaround is to force the sign to
                ' match the difference between the relevant coordinate of the intersecting lines.  (The relevant coordinate varies
                ' based on the orientation of the default, non-rotated line defined by ptIntersect and pt1.)
                If (m_CurPOI = poi_EdgeE) Then
                    If (pt2.y < pt1.y) Then newAngle = -newAngle
                ElseIf (m_CurPOI = poi_EdgeS) Then
                    If (pt2.x > pt1.x) Then newAngle = -newAngle
                ElseIf (m_CurPOI = poi_EdgeW) Then
                    If (pt2.y > pt1.y) Then newAngle = -newAngle
                ElseIf (m_CurPOI = poi_EdgeN) Then
                    If (pt2.x < pt1.x) Then newAngle = -newAngle
                End If
                
                'Apply the angle to the layer, and our work here is done!
                .SetLayerAngle newAngle
                
            '5: interior of the layer (e.g. move the layer instead of resize it)
            Case poi_Interior
                .SetLayerOffsetX m_InitLayerCoords_Pure(0).x + hOffsetImage
                .SetLayerOffsetY m_InitLayerCoords_Pure(0).y + vOffsetImage
            
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
        
        Dim cParams As pdParamXML
        Set cParams = New pdParamXML
        
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
                
                With srcImage.GetActiveLayer
                    Process "Move layer", False, cParams.GetParamString(), UNDO_LayerHeader
                End With
                
            'The caller can specify other dummy values if they don't want us to redraw the screen
        
        End Select
    
    'If the transformation is still active (e.g. the user has the mouse pressed down), just redraw the viewport, but don't
    ' process Undo/Redo or any macro stuff.
    Else
    
        'Manually request a canvas redraw
        ViewportEngine.Stage2_CompositeAllLayers srcImage, srcCanvas, m_CurPOI
    
    End If
    
End Sub

'Are on-canvas tools currently allowed?  This master function will evaluate all relevant program states for allowing on-canvas
' tool operations (e.g. "no open images", "main form locked").
Public Function CanvasToolsAllowed(Optional ByVal alsoCheckBusyState As Boolean = True, Optional ByVal checkMainWindowEnabled As Boolean = True) As Boolean

    'Start with a few failsafe checks
    
    'Make sure an image is loaded and active
    If (g_OpenImageCount > 0) Then
    
        'Make sure the main form has not been disabled by a modal dialog
        If FormMain.Enabled Or (Not checkMainWindowEnabled) Then
            
            'Finally, make sure another process hasn't locked the active canvas.  Note that the caller can disable this behavior
            ' if they don't need it.
            If alsoCheckBusyState Then
                CanvasToolsAllowed = (Not Processor.IsProgramBusy) And (Not Tools.GetToolBusyState)
            Else
                CanvasToolsAllowed = True
            End If
            
        Else
            CanvasToolsAllowed = False
        End If
    Else
        CanvasToolsAllowed = False
    End If
    
End Function

'Some tools (paintbrushes, most notably), have to initialize themselves against the current image (prepping a scratch layer,
' for example).  To reduce stuttering on first tool use, we initialize this behavior whenever...
' 1) Such a tool is selected, or...
' 2) The tool is already selected and the user switches images
Public Sub InitializeToolsDependentOnImage()
    
    If (g_OpenImageCount > 0) Then
        If (g_CurrentTool = PAINT_BASICBRUSH) Or (g_CurrentTool = PAINT_SOFTBRUSH) Or (g_CurrentTool = PAINT_ERASER) Or (g_CurrentTool = PAINT_FILL) Then
            
            'A couple things require us to reset the scratch layer...
            ' 1) If it hasn't been initialized at all
            ' 2) If it doesn't match the current image's size
            Dim scratchLayerResetRequired As Boolean: scratchLayerResetRequired = False
            scratchLayerResetRequired = (pdImages(g_CurrentImage).ScratchLayer Is Nothing)
            If (Not scratchLayerResetRequired) Then
                scratchLayerResetRequired = (pdImages(g_CurrentImage).ScratchLayer.GetLayerWidth <> pdImages(g_CurrentImage).Width)
                If (Not scratchLayerResetRequired) Then scratchLayerResetRequired = (pdImages(g_CurrentImage).ScratchLayer.GetLayerHeight <> pdImages(g_CurrentImage).Height)
            End If
            
            If scratchLayerResetRequired Then pdImages(g_CurrentImage).ResetScratchLayer True
            
        Else
            
            'The scratch layer is not required for non-paint tools, and releasing it frees a lot of memory
            Set pdImages(g_CurrentImage).ScratchLayer = Nothing
            
            'As a failsafe, restore default mouse settings, which may have been modified by various paintbrush tools
            FormMain.MainCanvas(0).SetMouseInput_AutoDrop True
            FormMain.MainCanvas(0).SetMouseInput_HighRes False
            
        End If
    End If
    
End Sub

'When the active layer changes, call this function.  It synchronizes various layer-specific tool panels against the
' currently active layer.  (Note that you also need to call this whenever a new tool panel is selected, as the newly
' loaded panel will reflect default values otherwise.)
Public Sub SyncToolOptionsUIToCurrentLayer()
    
    'Failsafe checks
    If (pdImages(g_CurrentImage) Is Nothing) Then Exit Sub
    If (Not pdImages(g_CurrentImage).IsActive) Then Exit Sub
    If (pdImages(g_CurrentImage).GetActiveLayer Is Nothing) Then Exit Sub
    
    'Before doing anything else, make sure canvas tool operations are allowed.  (They are disallowed if no images
    ' are loaded, for example.)
    If (Not CanvasToolsAllowed(False, False)) Then
        
        'Some panels may still wish to redraw their contents, even if no images are loaded.  (Text panels use this
        ' opportunity to hide the "convert typography to text or vice-versa" panels that are visible by default.)
        If (g_CurrentTool = VECTOR_TEXT) Then
            toolpanel_Text.UpdateAgainstCurrentLayer
        ElseIf (g_CurrentTool = VECTOR_FANCYTEXT) Then
            toolpanel_FancyText.UpdateAgainstCurrentLayer
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
        Case VECTOR_TEXT, VECTOR_FANCYTEXT
            If pdImages(g_CurrentImage).GetActiveLayer.IsLayerText Then
                layerToolActive = True
            Else
            
                'Hide the "convert to different type of text" panel prompts
                If (g_CurrentTool = VECTOR_TEXT) Then
                    toolpanel_Text.UpdateAgainstCurrentLayer
                ElseIf (g_CurrentTool = VECTOR_FANCYTEXT) Then
                    toolpanel_FancyText.UpdateAgainstCurrentLayer
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
                
                'The interface module actually has a nice function that already handles this
                Interface.SetUIGroupState PDUI_LayerTools, True
                
                'Reset tool busy state (because it will be reset by the Interface module call, above)
                Tools.SetToolBusyState True
                
            Case VECTOR_TEXT
                
                With toolpanel_Text
                    .txtTextTool.Text = pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_Text)
                    .cboTextFontFace.ListIndex = .cboTextFontFace.ListIndexByString(pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_FontFace), vbTextCompare)
                    .tudTextFontSize.Value = pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_FontSize)
                    .csTextFontColor.Color = pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_FontColor)
                    .cboTextRenderingHint.ListIndex = pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_TextAntialiasing)
                    .sltTextClarity.Value = pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_TextContrast)
                    .btnFontStyles(0).Value = CBool(pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_FontBold))
                    .btnFontStyles(1).Value = CBool(pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_FontItalic))
                    .btnFontStyles(2).Value = CBool(pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_FontUnderline))
                    .btnFontStyles(3).Value = CBool(pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_FontStrikeout))
                    .btsHAlignment.ListIndex = pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_HorizontalAlignment)
                    .btsVAlignment.ListIndex = pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_VerticalAlignment)
                End With
                
                'Display the "convert to basic text layer" panel as necessary
                toolpanel_Text.UpdateAgainstCurrentLayer
                
            Case VECTOR_FANCYTEXT
                
                With toolpanel_FancyText
                    .txtTextTool.Text = pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_Text)
                    .cboTextFontFace.ListIndex = .cboTextFontFace.ListIndexByString(pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_FontFace), vbTextCompare)
                    .tudTextFontSize.Value = pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_FontSize)
                    .cboTextRenderingHint.ListIndex = pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_TextAntialiasing)
                    .chkHinting.Value = IIf(pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_TextHinting), vbChecked, vbUnchecked)
                    .btnFontStyles(0).Value = CBool(pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_FontBold))
                    .btnFontStyles(1).Value = CBool(pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_FontItalic))
                    .btnFontStyles(2).Value = CBool(pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_FontUnderline))
                    .btnFontStyles(3).Value = CBool(pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_FontStrikeout))
                    .btsHAlignment.ListIndex = pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_HorizontalAlignment)
                    .btsVAlignment.ListIndex = pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_VerticalAlignment)
                    .cboWordWrap.ListIndex = pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_WordWrap)
                    .chkFillText.Value = IIf(pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_FillActive), vbChecked, vbUnchecked)
                    .bsText.Brush = pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_FillBrush)
                    .chkOutlineText.Value = IIf(pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_OutlineActive), vbChecked, vbUnchecked)
                    .psText.Pen = pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_OutlinePen)
                    .chkBackground.Value = IIf(pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_BackgroundActive), vbChecked, vbUnchecked)
                    .bsTextBackground.Brush = pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_BackgroundBrush)
                    .chkBackgroundBorder.Value = IIf(pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_BackBorderActive), vbChecked, vbUnchecked)
                    .psTextBackground.Pen = pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_BackBorderPen)
                    .tudMargin(0).Value = pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_MarginLeft)
                    .tudMargin(1).Value = pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_MarginRight)
                    .tudMargin(2).Value = pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_MarginTop)
                    .tudMargin(3).Value = pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_MarginBottom)
                    .tudLineSpacing.Value = pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_LineSpacing)
                    .sltCharInflation.Value = pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_CharInflation)
                    .tudJitter(0).Value = pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_CharJitterX)
                    .tudJitter(1).Value = pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_CharJitterY)
                    .cboCharMirror.ListIndex = pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_CharMirror)
                    .sltCharOrientation.Value = pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_CharOrientation)
                    .cboCharCase.ListIndex = pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_CharRemap)
                    .sltCharSpacing.Value = pdImages(g_CurrentImage).GetActiveLayer.GetTextLayerProperty(ptp_CharSpacing)
                End With
                
                'Display the "convert to typography layer" panel as necessary
                toolpanel_FancyText.UpdateAgainstCurrentLayer
        
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
        
        Case VECTOR_TEXT, VECTOR_FANCYTEXT
            If pdImages(g_CurrentImage).GetActiveLayer.IsLayerText Then layerToolActive = True
        
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
                pdImages(g_CurrentImage).GetActiveLayer.SetLayerOffsetX toolpanel_MoveSize.tudLayerMove(0).Value
                pdImages(g_CurrentImage).GetActiveLayer.SetLayerOffsetY toolpanel_MoveSize.tudLayerMove(1).Value
                
                'Setting layer width and height isn't activated at present, on purpose
                'pdImages(g_CurrentImage).getActiveLayer.setLayerWidth toolpanel_MoveSize.tudLayerMove(2).Value
                'pdImages(g_CurrentImage).getActiveLayer.setLayerHeight toolpanel_MoveSize.tudLayerMove(3).Value
                
                'The layer resize quality combo box also needs to be synched
                pdImages(g_CurrentImage).GetActiveLayer.SetLayerResizeQuality toolpanel_MoveSize.cboLayerResizeQuality.ListIndex
                
                'Layer angle and shear are newly available as of 7.0
                pdImages(g_CurrentImage).GetActiveLayer.SetLayerAngle toolpanel_MoveSize.sltLayerAngle.Value
                pdImages(g_CurrentImage).GetActiveLayer.SetLayerShearX toolpanel_MoveSize.sltLayerShearX.Value
                pdImages(g_CurrentImage).GetActiveLayer.SetLayerShearY toolpanel_MoveSize.sltLayerShearY.Value
            
            Case VECTOR_TEXT
                
                With pdImages(g_CurrentImage).GetActiveLayer
                    .SetTextLayerProperty ptp_Text, toolpanel_Text.txtTextTool.Text
                    .SetTextLayerProperty ptp_FontFace, toolpanel_Text.cboTextFontFace.List(toolpanel_Text.cboTextFontFace.ListIndex)
                    .SetTextLayerProperty ptp_FontSize, toolpanel_Text.tudTextFontSize.Value
                    .SetTextLayerProperty ptp_FontColor, toolpanel_Text.csTextFontColor.Color
                    .SetTextLayerProperty ptp_TextAntialiasing, toolpanel_Text.cboTextRenderingHint.ListIndex
                    .SetTextLayerProperty ptp_TextContrast, toolpanel_Text.sltTextClarity.Value
                    .SetTextLayerProperty ptp_FontBold, toolpanel_Text.btnFontStyles(0).Value
                    .SetTextLayerProperty ptp_FontItalic, toolpanel_Text.btnFontStyles(1).Value
                    .SetTextLayerProperty ptp_FontUnderline, toolpanel_Text.btnFontStyles(2).Value
                    .SetTextLayerProperty ptp_FontStrikeout, toolpanel_Text.btnFontStyles(3).Value
                    .SetTextLayerProperty ptp_HorizontalAlignment, toolpanel_Text.btsHAlignment.ListIndex
                    .SetTextLayerProperty ptp_VerticalAlignment, toolpanel_Text.btsVAlignment.ListIndex
                End With
                
                'This is a little weird, but we also make sure to synchronize the current text rendering engine when the UI is synched.
                ' This is because this property changes according to the active text tool.
                pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_RenderingEngine, tre_WAPI
            
            Case VECTOR_FANCYTEXT
                
                With pdImages(g_CurrentImage).GetActiveLayer
                    .SetTextLayerProperty ptp_Text, toolpanel_FancyText.txtTextTool.Text
                    .SetTextLayerProperty ptp_FontFace, toolpanel_FancyText.cboTextFontFace.List(toolpanel_FancyText.cboTextFontFace.ListIndex)
                    .SetTextLayerProperty ptp_FontSize, toolpanel_FancyText.tudTextFontSize.Value
                    .SetTextLayerProperty ptp_TextAntialiasing, toolpanel_FancyText.cboTextRenderingHint.ListIndex
                    .SetTextLayerProperty ptp_TextHinting, CBool(toolpanel_FancyText.chkHinting.Value)
                    .SetTextLayerProperty ptp_FontBold, toolpanel_FancyText.btnFontStyles(0).Value
                    .SetTextLayerProperty ptp_FontItalic, toolpanel_FancyText.btnFontStyles(1).Value
                    .SetTextLayerProperty ptp_FontUnderline, toolpanel_FancyText.btnFontStyles(2).Value
                    .SetTextLayerProperty ptp_FontStrikeout, toolpanel_FancyText.btnFontStyles(3).Value
                    .SetTextLayerProperty ptp_HorizontalAlignment, toolpanel_FancyText.btsHAlignment.ListIndex
                    .SetTextLayerProperty ptp_VerticalAlignment, toolpanel_FancyText.btsVAlignment.ListIndex
                    .SetTextLayerProperty ptp_WordWrap, toolpanel_FancyText.cboWordWrap.ListIndex
                    .SetTextLayerProperty ptp_FillActive, CBool(toolpanel_FancyText.chkFillText.Value)
                    .SetTextLayerProperty ptp_FillBrush, toolpanel_FancyText.bsText.Brush
                    .SetTextLayerProperty ptp_OutlineActive, CBool(toolpanel_FancyText.chkOutlineText.Value)
                    .SetTextLayerProperty ptp_OutlinePen, toolpanel_FancyText.psText.Pen
                    .SetTextLayerProperty ptp_BackgroundActive, CBool(toolpanel_FancyText.chkBackground.Value)
                    .SetTextLayerProperty ptp_BackgroundBrush, toolpanel_FancyText.bsTextBackground.Brush
                    .SetTextLayerProperty ptp_BackBorderActive, CBool(toolpanel_FancyText.chkBackgroundBorder.Value)
                    .SetTextLayerProperty ptp_BackBorderPen, toolpanel_FancyText.psTextBackground.Pen
                    .SetTextLayerProperty ptp_LineSpacing, toolpanel_FancyText.tudLineSpacing.Value
                    .SetTextLayerProperty ptp_MarginLeft, toolpanel_FancyText.tudMargin(0).Value
                    .SetTextLayerProperty ptp_MarginRight, toolpanel_FancyText.tudMargin(1).Value
                    .SetTextLayerProperty ptp_MarginTop, toolpanel_FancyText.tudMargin(2).Value
                    .SetTextLayerProperty ptp_MarginBottom, toolpanel_FancyText.tudMargin(3).Value
                    .SetTextLayerProperty ptp_CharInflation, toolpanel_FancyText.sltCharInflation.Value
                    .SetTextLayerProperty ptp_CharJitterX, toolpanel_FancyText.tudJitter(0).Value
                    .SetTextLayerProperty ptp_CharJitterY, toolpanel_FancyText.tudJitter(1).Value
                    .SetTextLayerProperty ptp_CharMirror, toolpanel_FancyText.cboCharMirror.ListIndex
                    .SetTextLayerProperty ptp_CharOrientation, toolpanel_FancyText.sltCharOrientation.Value
                    .SetTextLayerProperty ptp_CharRemap, toolpanel_FancyText.cboCharCase.ListIndex
                    .SetTextLayerProperty ptp_CharSpacing, toolpanel_FancyText.sltCharSpacing.Value
                End With
                
                'This is a little weird, but we also make sure to synchronize the current text rendering engine when the UI is synched.
                ' This is because this property changes according to the active text tool.
                pdImages(g_CurrentImage).GetActiveLayer.SetTextLayerProperty ptp_RenderingEngine, tre_PHOTODEMON
        
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
