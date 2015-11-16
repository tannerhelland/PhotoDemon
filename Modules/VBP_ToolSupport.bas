Attribute VB_Name = "Tool_Support"
'***************************************************************************
'Helper functions for various PhotoDemon tools
'Copyright 2014-2015 by Tanner Helland
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
Private m_InitLayerCoords_Transformed(0 To 3) As POINTFLOAT
Private m_InitLayerCoords_Pure(0 To 3) As POINTFLOAT
Private m_LayerAspectRatio As Double
Private m_InitImageX As Double, m_InitImageY As Double, m_InitLayerX As Single, m_InitLayerY As Single

'If a point of interest is being modified by a tool action, its ID will be stored here.  Make sure to clear this value
' (to -1, which means "no point of interest") when you are finished with it (typically after MouseUp).
Private m_curPOI As Long

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

Public Function getToolBusyState() As Boolean
    getToolBusyState = m_ToolIsBusy
End Function

Public Sub setToolBusyState(ByVal newState As Boolean)
    m_ToolIsBusy = newState
End Sub

Public Function getCustomToolState() As Long
    getCustomToolState = m_CustomToolMarker
    m_CustomToolMarker = 0
End Function

Public Sub setCustomToolState(ByVal newState As Long)
    m_CustomToolMarker = newState
End Sub

'When a tool is finished processing, it can call this function to release all tool tracking variables
Public Sub terminateGenericToolTracking()
    
    'Reset the current POI, if any
    m_curPOI = -1
    
End Sub

'The move tool uses this function to set various initial parameters for layer interactions.
Public Sub setInitialLayerToolValues(ByRef srcImage As pdImage, ByRef srcLayer As pdLayer, ByVal mouseX_ImageSpace As Double, ByVal mouseY_ImageSpace As Double, Optional ByVal relevantPOI As Long = -1)
    
    'Cache the initial mouse values.  Note that, per the parameter names, these must have already been converted to the image's
    ' coordinate space (NOT the canvas's!)
    m_InitImageX = mouseX_ImageSpace
    m_InitImageY = mouseY_ImageSpace
    
    'Also, make a copy of those coordinates in the current layer space
    Drawing.convertImageCoordsToLayerCoords srcImage, srcLayer, m_InitImageX, m_InitImageY, m_InitLayerX, m_InitLayerY
    
    'Make a copy of the current layer coordinates, with any affine transforms applied (rotation, etc)
    srcLayer.getLayerCornerCoordinates m_InitLayerCoords_Transformed
    
    'Finally, make a copy of the current layer coordinates, *without* affine transforms applied.  This is basically the rect of
    ' the layer as it would appear if no affine modifiers were active (e.g. without rotation, etc)
    Dim i As Long
    For i = 0 To 3
        Drawing.convertImageCoordsToLayerCoords srcImage, srcLayer, m_InitLayerCoords_Transformed(i).x, m_InitLayerCoords_Transformed(i).y, m_InitLayerCoords_Pure(i).x, m_InitLayerCoords_Pure(i).y
    Next i
    
    'Cache the layer's aspect ratio.  Note that this *does include any current non-destructive transforms*!
    If srcLayer.getLayerHeight(False) <> 0 Then
        m_LayerAspectRatio = srcLayer.getLayerWidth(False) / srcLayer.getLayerHeight(False)
    Else
        m_LayerAspectRatio = 1
    End If
    
    'If a relevant POI was supplied, store it as well.  Note that not all tools make use of this.
    m_curPOI = relevantPOI
        
End Sub

'The drag-to-pan tool uses this function to set the initial scroll bar values for a pan operation
Public Sub setInitialCanvasScrollValues(ByRef srcCanvas As pdCanvas)

    m_InitHScroll = srcCanvas.getScrollValue(PD_HORIZONTAL)
    m_InitVScroll = srcCanvas.getScrollValue(PD_VERTICAL)

End Sub

'The drag-to-pan tool uses this function to actually scroll the viewport area
Public Sub panImageCanvas(ByVal initX As Long, ByVal initY As Long, ByVal curX As Long, ByVal curY As Long, ByRef srcImage As pdImage, ByRef srcCanvas As pdCanvas)

    'Prevent the canvas from redrawing itself until our pan operation is complete.  (This prevents juddery movement.)
    srcCanvas.setRedrawSuspension True
    
    'Sub-pixel panning is now allowed (because we're awesome like that)
    Dim zoomRatio As Double
    zoomRatio = g_Zoom.getZoomValue(srcImage.currentZoomValue)
    
    'Calculate new scroll values
    Dim hOffset As Long, vOffset As Long
    hOffset = (initX - curX) / zoomRatio
    vOffset = (initY - curY) / zoomRatio
        
    'Factor in the initial scroll bar values
    hOffset = m_InitHScroll + hOffset
    vOffset = m_InitVScroll + vOffset
        
    'If these values lie within the bounds of their respective scroll bar(s), apply 'em
    If (hOffset < srcCanvas.getScrollMin(PD_HORIZONTAL)) Then
        srcCanvas.setScrollValue PD_HORIZONTAL, srcCanvas.getScrollMin(PD_HORIZONTAL)
    ElseIf (hOffset > srcCanvas.getScrollMax(PD_HORIZONTAL)) Then
        srcCanvas.setScrollValue PD_HORIZONTAL, srcCanvas.getScrollMax(PD_HORIZONTAL)
    Else
        srcCanvas.setScrollValue PD_HORIZONTAL, hOffset
    End If
    
    If (vOffset < srcCanvas.getScrollMin(PD_VERTICAL)) Then
        srcCanvas.setScrollValue PD_VERTICAL, srcCanvas.getScrollMin(PD_VERTICAL)
    ElseIf (vOffset > srcCanvas.getScrollMax(PD_VERTICAL)) Then
        srcCanvas.setScrollValue PD_VERTICAL, srcCanvas.getScrollMax(PD_VERTICAL)
    Else
        srcCanvas.setScrollValue PD_VERTICAL, vOffset
    End If
    
    'Reinstate canvas redraws
    srcCanvas.setRedrawSuspension False
    
    'Request the scroll-specific viewport pipeline stage
    Viewport_Engine.Stage3_ExtractRelevantRegion srcImage, FormMain.mainCanvas(0)
    
End Sub

'This function can be used to move and/or non-destructively resize an image layer.
'
'If this action occurs during a Mouse_Up event, the finalizeTransform parameter should be set to TRUE. This instructs the function
' to forward the transformation request to PD's central processor, so it can generate Undo/Redo data, be recorded as part of macros, etc.
Public Sub transformCurrentLayer(ByVal curImageX As Double, ByVal curImageY As Double, ByRef srcImage As pdImage, ByRef srcLayer As pdLayer, ByRef srcCanvas As pdCanvas, Optional ByVal isShiftDown As Boolean = False, Optional ByVal finalizeTransform As Boolean = False)
    
    'Prevent the canvas from redrawing itself until our movement calculations are complete.
    ' (This prevents juddery movement.)
    srcCanvas.setRedrawSuspension True
    
    'Also, mark the tool engine as busy to prevent re-entrance issues
    Tool_Support.setToolBusyState True
    
    'Convert the current x/y pair to the layer coordinate space.  This takes into account any active affine transforms
    ' on the image (e.g. rotation), which may place the point in a totally different position relative to the underlying layer.
    Dim curLayerX As Single, curLayerY As Single
    Drawing.convertImageCoordsToLayerCoords srcImage, srcLayer, curImageX, curImageY, curLayerX, curLayerY
            
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
    Dim newX As Double, newY As Double, newWidth As Double, newHeight As Double
    
    'The way we assign new offsets and/or sizes to the layer depends on the POI (point of interest) the user is interacting with.
    ' Layers currently support five points of interest: each of their 4 corners, and anywhere in the layer interior
    ' (for moving the layer).
    
    'Because the various POI evaluators share similar code (they all just set a new boundary rect), this value will be set to TRUE
    ' if a POI was successfully evaluated.  This triggers a set of uniform code checks, including safe boundaries and SHIFT key handling.
    Dim poiCleanupRequired As Boolean
    poiCleanupRequired = False
    
    'Check the POI we were given, and update the layer accordingly.
    With srcLayer
    
        Select Case m_curPOI
            
            '-1: the mouse is not over the layer.  Do nothing.
            Case -1
                Tool_Support.setToolBusyState False
                srcCanvas.setRedrawSuspension False
                Exit Sub
                
            '0: the mouse is dragging the top-left corner of the layer.  The comments here are uniform for all POIs, so for brevity's sake,
            ' I'll keep the others short.
            Case 0
                
                'The new (x, y) offset for this layer is simply the current mouse coordinates, transformed to the layer's coordinate space
                newLeft = curLayerX
                newTop = curLayerY
                
                'As of PD 7.0, corner interactions cause the layer to naturally resize around its current center point.  As such, we need
                ' to calculate new width/height values now.
                newRight = m_InitLayerCoords_Pure(3).x - hOffsetLayer
                newBottom = m_InitLayerCoords_Pure(3).y - vOffsetLayer
                
                poiCleanupRequired = True
                
            '1: top-right corner
            Case 1
            
                'Calculate a new boundary rect
                newRight = curLayerX
                newTop = curLayerY
                newLeft = m_InitLayerCoords_Pure(0).x - hOffsetLayer
                newBottom = m_InitLayerCoords_Pure(3).y - vOffsetLayer
                
                poiCleanupRequired = True
                
            '3: bottom-left
            Case 2
                
                'Calculate a new boundary rect
                newLeft = curLayerX
                newBottom = curLayerY
                newRight = m_InitLayerCoords_Pure(3).x - hOffsetLayer
                newTop = m_InitLayerCoords_Pure(0).y - vOffsetLayer
                
                poiCleanupRequired = True
                
            '2: bottom-right
            Case 3
                
                'Calculate a new boundary rect
                newRight = curLayerX
                newBottom = curLayerY
                newLeft = m_InitLayerCoords_Pure(0).x - hOffsetLayer
                newTop = m_InitLayerCoords_Pure(0).y - vOffsetLayer
                
                poiCleanupRequired = True
            
            '4-7: rotation nodes
            Case 4 To 7
            
                'Layer rotation is different because it involves finding the angle between two lines; specifically, the angle between
                ' a flat origin line and the current node-to-origin line of the rotation node.
                Dim ptIntersect As POINTFLOAT, pt1 As POINTFLOAT, pt2 As POINTFLOAT
                Dim ptIntersect_T As POINTFLOAT, pt1_T As POINTFLOAT, pt2_T As POINTFLOAT
                
                'The intersect point is the center of the image.  This point is the same for all rotation nodes.
                ptIntersect.x = m_InitLayerCoords_Pure(0).x + (m_InitLayerCoords_Pure(3).x - m_InitLayerCoords_Pure(0).x) / 2
                ptIntersect.y = m_InitLayerCoords_Pure(0).y + (m_InitLayerCoords_Pure(3).y - m_InitLayerCoords_Pure(0).y) / 2
                
                'The first non-intersecting point varies by rotation node (as they lie in 90-degree increments).  Note that the
                ' 100 offset is totally arbitrary; we just need a line of some non-zero length for the angle calculation to work.
                If m_curPOI = 4 Then
                    pt1.x = ptIntersect.x + 100
                    pt1.y = ptIntersect.y
                ElseIf m_curPOI = 5 Then
                    pt1.x = ptIntersect.x
                    pt1.y = ptIntersect.y + 100
                ElseIf m_curPOI = 6 Then
                    pt1.x = ptIntersect.x - 100
                    pt1.y = ptIntersect.y
                Else
                    pt1.x = ptIntersect.x
                    pt1.y = ptIntersect.y - 100
                End If
                                                
                'The second non-intersecting point is the current mouse position.
                pt2.x = curImageX
                pt2.y = curImageY
                
                'If shearing is active on the current layer, we need to account for its effect on the current mouse location.
                ' (Note that we could apply this matrix transformation regardless of current shear values, as values of zero
                ' will simply return an identity matrix, but why do extra math if it's not required?)
                If (srcLayer.getLayerShearX <> 0) Or (srcLayer.getLayerShearY <> 0) Then
                
                    'Apply the current layer's shear effect to the mouse position.  This gives us its unadulterated equivalent,
                    ' e.g. its location in the same coordinate space as the two points we've already calculated.
                    Dim tmpMatrix As pdGraphicsMatrix
                    Set tmpMatrix = New pdGraphicsMatrix
                    
                    tmpMatrix.ShearMatrix srcLayer.getLayerShearX, srcLayer.getLayerShearY, ptIntersect.x, ptIntersect.y
                    tmpMatrix.InvertMatrix
                    
                    tmpMatrix.applyMatrixToPointF pt2
                
                End If
                
                'Find the angle between the two lines we've calculated
                Dim newAngle As Double
                newAngle = Math_Functions.angleBetweenTwoIntersectingLines(ptIntersect, pt1, pt2, True)
                
                'Because the angle function finds the absolute inner angle, it will never be greater than 180 degrees.  This also means
                ' that +90 and -90 (from a UI standpoint) return the same 90 result.  A simple workaround is to force the sign to
                ' match the difference between the relevant coordinate of the intersecting lines.  (The relevant coordinate varies
                ' based on the orientation of the default, non-rotated line defined by ptIntersect and pt1.)
                If m_curPOI = 4 Then
                    If pt2.y < pt1.y Then newAngle = -newAngle
                ElseIf m_curPOI = 5 Then
                    If pt2.x > pt1.x Then newAngle = -newAngle
                ElseIf m_curPOI = 6 Then
                    If pt2.y > pt1.y Then newAngle = -newAngle
                Else
                    If pt2.x < pt1.x Then newAngle = -newAngle
                End If
                
                'Apply the angle to the layer, and our work here is done!
                .setLayerAngle newAngle
                            
            '5: interior of the layer (e.g. move the layer instead of resize it)
            Case 8
                .setLayerOffsetX m_InitLayerCoords_Pure(0).x + hOffsetImage
                .setLayerOffsetY m_InitLayerCoords_Pure(0).y + vOffsetImage
            
        End Select
        
        'If a POI was successfully evaluated, we need to perform some generic clean-up on the calculated boundary rect.
        ' (Note that moving the layer doesn't trigger these checks, as movement alone can't result in invalid bounds.)
        If poiCleanupRequired Then
        
            'If the SHIFT key is down, lock the image's aspect ratio
            If isShiftDown Then
            
                newHeight = (newRight - newLeft) / m_LayerAspectRatio
                
                'Shift the top coordinate offset to compensate for the newly calculated height
                newY = newTop + (newBottom - newTop) / 2
                newBottom = newY + (newHeight / 2)
                newTop = newBottom - newHeight
                
            End If
            
            'Make sure the new (x, y) values don't result in negative width/height modifiers
            If (newRight > newLeft) And (newBottom > newTop) Then .setOffsetsAndModifiersTogether newLeft, newTop, newRight, newBottom
        
        End If
        
    End With
    
    'Manually synchronize the new values against their on-screen UI elements
    Tool_Support.syncToolOptionsUIToCurrentLayer
    
    'Free the tool engine
    Tool_Support.setToolBusyState False
    
    'Reinstate canvas redraws
    srcCanvas.setRedrawSuspension False
    
    'If this is the final step of a transform (e.g. if the user has just released the mouse), forward this
    ' request to PD's central processor, so an Undo/Redo entry can be generated.
    If finalizeTransform Then
        
        'As a convenience to the user, layer resize and move operations are listed separately.
        Select Case m_curPOI
        
            'Move/resize transformations.
            Case 0 To 3
            
                With srcImage.getActiveLayer
                    Process "Resize layer (on-canvas)", False, buildParams(.getLayerOffsetX, .getLayerOffsetY, .getLayerCanvasXModifier, .getLayerCanvasYModifier), UNDO_LAYERHEADER
                End With
                
            'Rotation
            Case 4 To 7
                With srcImage.getActiveLayer
                    Process "Rotate layer (on-canvas)", False, buildParams(.getLayerAngle), UNDO_LAYERHEADER
                End With
            
            'Move-only transformations
            Case 8
                
                With srcImage.getActiveLayer
                    Process "Move layer", False, buildParams(.getLayerOffsetX, .getLayerOffsetY), UNDO_LAYERHEADER
                End With
                
            'The caller can specify other dummy values if they don't want us to redraw the screen
        
        End Select
    
    'If the transformation is still active (e.g. the user has the mouse pressed down), just redraw the viewport, but don't
    ' process Undo/Redo or any macro stuff.
    Else
    
        'Manually request a canvas redraw
        Viewport_Engine.Stage2_CompositeAllLayers srcImage, srcCanvas, False, m_curPOI
    
    End If
    
End Sub

'Assuming the user has made one or more edits via the Quick-Fix function, permanently apply those changes to the image now.
Public Sub makeQuickFixesPermanent()

    'Prepare a PD Compositor object, which will handle the actual compositing step
    Dim tmpCompositor As pdCompositor
    Set tmpCompositor = New pdCompositor
    
    'Apply the quick-fix adjustments
    tmpCompositor.applyNDFXToDIB pdImages(g_CurrentImage).getActiveLayer, pdImages(g_CurrentImage).getActiveDIB
    
    'Reset the quick-fix settings stored inside the pdLayer object
    Dim i As Long
    For i = 0 To toolpanel_NDFX.sltQuickFix.Count - 1
        pdImages(g_CurrentImage).getActiveLayer.setLayerNonDestructiveFXState i, 0
    Next i
    
End Sub

'Are on-canvas tools currently allowed?  This master function will evaluate all relevant program states for allowing on-canvas
' tool operations (e.g. "no open images", "main form locked").
Public Function canvasToolsAllowed(Optional ByVal alsoCheckBusyState As Boolean = True) As Boolean

    'Start with a few failsafe checks
    
    'Make sure an image is loaded and active
    If g_OpenImageCount > 0 Then
    
        'Make sure the main form has not been disabled by a modal dialog
        If FormMain.Enabled Then
            
            'Finally, make sure another process hasn't locked the active canvas.  Note that the caller can disable this behavior
            ' if they don't need it.
            If alsoCheckBusyState Then
                
                If (Not Processor.IsProgramBusy) And (Not getToolBusyState) Then
                    canvasToolsAllowed = True
                Else
                    canvasToolsAllowed = False
                End If
            
            Else
                canvasToolsAllowed = True
            End If
            
        Else
            canvasToolsAllowed = False
        End If
    Else
        canvasToolsAllowed = False
    End If
    
End Function

'When the active layer changes, call this function.  It synchronizes various layer-specific tool panels against the
' currently active layer.
Public Sub syncToolOptionsUIToCurrentLayer()
    
    'Before doing anything else, make sure canvas tool operations are allowed
    If Not canvasToolsAllowed(False) Then
        
        'Some panels may redraw their contents if no images are loaded
        If g_CurrentTool = VECTOR_TEXT Then
            toolpanel_Text.updateAgainstCurrentLayer
        ElseIf g_CurrentTool = VECTOR_FANCYTEXT Then
            toolpanel_FancyText.updateAgainstCurrentLayer
        End If
        
        'Exit now, as subsequent checks in this function require one or more active images
        Exit Sub
        
    End If
    
    Dim layerToolActive As Boolean
    
    Select Case g_CurrentTool
        
        Case NAV_MOVE
            layerToolActive = True
        
        Case VECTOR_TEXT, VECTOR_FANCYTEXT
            If pdImages(g_CurrentImage).getActiveLayer.isLayerText Then
                layerToolActive = True
            Else
            
                'Hide the "convert to different type of text panel" prompts
                If g_CurrentTool = VECTOR_TEXT Then
                    toolpanel_Text.updateAgainstCurrentLayer
                ElseIf g_CurrentTool = VECTOR_FANCYTEXT Then
                    toolpanel_FancyText.updateAgainstCurrentLayer
                End If
            
            End If
        
        Case Else
            layerToolActive = False
        
    End Select
    
    'To improve performance, we'll only sync the UI if a layer-specific tool is active, and the tool options panel is
    ' currently visible.
    If (Not toolbar_Options.Visible) And (Not layerToolActive) Then Exit Sub
    
    If layerToolActive Then
        
        'Mark the tool engine as busy; this prevents each change from triggering viewport redraws
        Tool_Support.setToolBusyState True
        
        'Start iterating various layer properties, and reflecting them across their corresponding UI elements.
        ' (Obviously, this step is separated by tool type.)
        Select Case g_CurrentTool
        
            Case NAV_MOVE
            
                'The Layer Move tool has four text up/downs: two for layer position (x, y) and two for layer size (w, y)
                toolpanel_MoveSize.tudLayerMove(0).Value = pdImages(g_CurrentImage).getActiveLayer.getLayerOffsetX
                toolpanel_MoveSize.tudLayerMove(1).Value = pdImages(g_CurrentImage).getActiveLayer.getLayerOffsetY
                toolpanel_MoveSize.tudLayerMove(2).Value = pdImages(g_CurrentImage).getActiveLayer.getLayerWidth
                toolpanel_MoveSize.tudLayerMove(3).Value = pdImages(g_CurrentImage).getActiveLayer.getLayerHeight
                
                'The layer resize quality combo box also needs to be synched
                toolpanel_MoveSize.cboLayerResizeQuality.ListIndex = pdImages(g_CurrentImage).getActiveLayer.getLayerResizeQuality
                
                'Layer angle and shear are newly available as of 7.0
                toolpanel_MoveSize.sltLayerAngle.Value = pdImages(g_CurrentImage).getActiveLayer.getLayerAngle
                toolpanel_MoveSize.sltLayerShearX.Value = pdImages(g_CurrentImage).getActiveLayer.getLayerShearX
                toolpanel_MoveSize.sltLayerShearY.Value = pdImages(g_CurrentImage).getActiveLayer.getLayerShearY
            
            Case VECTOR_TEXT
                
                With toolpanel_Text
                    .txtTextTool.Text = pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_Text)
                    .cboTextFontFace.setListIndexByString pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_FontFace)
                    .tudTextFontSize.Value = pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_FontSize)
                    .csTextFontColor.Color = pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_FontColor)
                    .cboTextRenderingHint.ListIndex = pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_TextAntialiasing)
                    .sltTextClarity.Value = pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_TextContrast)
                    .btnFontStyles(0).Value = CBool(pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_FontBold))
                    .btnFontStyles(1).Value = CBool(pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_FontItalic))
                    .btnFontStyles(2).Value = CBool(pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_FontUnderline))
                    .btnFontStyles(3).Value = CBool(pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_FontStrikeout))
                    .btsHAlignment.ListIndex = pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_HorizontalAlignment)
                    .btsVAlignment.ListIndex = pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_VerticalAlignment)
                End With
                
                'Display the "convert to basic text layer" panel as necessary
                toolpanel_Text.updateAgainstCurrentLayer
                
            Case VECTOR_FANCYTEXT
                
                With toolpanel_FancyText
                    .txtTextTool.Text = pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_Text)
                    .cboTextFontFace.setListIndexByString pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_FontFace)
                    .tudTextFontSize.Value = pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_FontSize)
                    .cboTextRenderingHint.ListIndex = pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_TextAntialiasing)
                    .chkHinting.Value = IIf(pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_TextHinting), vbChecked, vbUnchecked)
                    .btnFontStyles(0).Value = CBool(pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_FontBold))
                    .btnFontStyles(1).Value = CBool(pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_FontItalic))
                    .btnFontStyles(2).Value = CBool(pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_FontUnderline))
                    .btnFontStyles(3).Value = CBool(pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_FontStrikeout))
                    .btsHAlignment.ListIndex = pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_HorizontalAlignment)
                    .btsVAlignment.ListIndex = pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_VerticalAlignment)
                    .cboWordWrap.ListIndex = pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_WordWrap)
                    .chkFillText.Value = IIf(pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_FillActive), vbChecked, vbUnchecked)
                    .bsText.Brush = pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_FillBrush)
                    .chkOutlineText.Value = IIf(pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_OutlineActive), vbChecked, vbUnchecked)
                    .psText.Pen = pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_OutlinePen)
                    .chkBackground.Value = IIf(pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_BackgroundActive), vbChecked, vbUnchecked)
                    .bsTextBackground.Brush = pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_BackgroundBrush)
                    .chkBackgroundBorder.Value = IIf(pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_BackBorderActive), vbChecked, vbUnchecked)
                    .psTextBackground.Pen = pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_BackBorderPen)
                    .tudMargin(0).Value = pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_MarginLeft)
                    .tudMargin(1).Value = pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_MarginRight)
                    .tudMargin(2).Value = pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_MarginTop)
                    .tudMargin(3).Value = pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_MarginBottom)
                    .sltCharInflation.Value = pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_CharInflation)
                    .tudJitter(0).Value = pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_CharJitterX)
                    .tudJitter(1).Value = pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_CharJitterY)
                    .cboCharMirror.ListIndex = pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_CharMirror)
                    .sltCharOrientation.Value = pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_CharOrientation)
                    .cboCharCase.ListIndex = pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_CharRemap)
                    .sltCharSpacing.Value = pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_CharSpacing)
                End With
                
                'Display the "convert to typography layer" panel as necessary
                toolpanel_FancyText.updateAgainstCurrentLayer
        
        End Select
        
        'Free the tool engine
        Tool_Support.setToolBusyState False
    
    End If
    
End Sub

'this function is the reverse of syncToolOptionsUIToCurrentLayer(), above.  If you want to copy all current UI settings into
' the currently active layer, call this function.
Public Sub syncCurrentLayerToToolOptionsUI()
    
    'Before doing anything else, make sure canvas tool operations are allowed
    If Not canvasToolsAllowed(False) Then Exit Sub
    
    'To improve performance, we'll only sync the UI if a layer-specific tool is active, and the tool options panel is currently
    ' set to VISIBLE.
    If Not toolbar_Options.Visible Then Exit Sub
    
    Dim layerToolActive As Boolean
    
    Select Case g_CurrentTool
        
        Case NAV_MOVE
            layerToolActive = True
        
        Case VECTOR_TEXT, VECTOR_FANCYTEXT
            If pdImages(g_CurrentImage).getActiveLayer.isLayerText Then layerToolActive = True
        
        Case Else
            layerToolActive = False
        
    End Select
    
    If layerToolActive Then
        
        'Mark the tool engine as busy; this prevents each change from triggering viewport redraws
        Tool_Support.setToolBusyState True
        
        'Start iterating various layer properties, and reflecting them across their corresponding UI elements.
        ' (Obviously, this step is separated by tool type.)
        Select Case g_CurrentTool
        
            Case NAV_MOVE
            
                'The Layer Move tool has four text up/downs: two for layer position (x, y) and two for layer size (w, y)
                pdImages(g_CurrentImage).getActiveLayer.setLayerOffsetX toolpanel_MoveSize.tudLayerMove(0).Value
                pdImages(g_CurrentImage).getActiveLayer.setLayerOffsetY toolpanel_MoveSize.tudLayerMove(1).Value
                
                'Setting layer width and height isn't activated at present, on purpose
                'pdImages(g_CurrentImage).getActiveLayer.setLayerWidth toolpanel_MoveSize.tudLayerMove(2).Value
                'pdImages(g_CurrentImage).getActiveLayer.setLayerHeight toolpanel_MoveSize.tudLayerMove(3).Value
                
                'The layer resize quality combo box also needs to be synched
                pdImages(g_CurrentImage).getActiveLayer.setLayerResizeQuality toolpanel_MoveSize.cboLayerResizeQuality.ListIndex
                
                'Layer angle and shear are newly available as of 7.0
                pdImages(g_CurrentImage).getActiveLayer.setLayerAngle toolpanel_MoveSize.sltLayerAngle.Value
                pdImages(g_CurrentImage).getActiveLayer.setLayerShearX toolpanel_MoveSize.sltLayerShearX.Value
                pdImages(g_CurrentImage).getActiveLayer.setLayerShearY toolpanel_MoveSize.sltLayerShearY.Value
            
            Case VECTOR_TEXT
                
                With pdImages(g_CurrentImage).getActiveLayer
                    .setTextLayerProperty ptp_Text, toolpanel_Text.txtTextTool.Text
                    .setTextLayerProperty ptp_FontFace, toolpanel_Text.cboTextFontFace.List(toolpanel_Text.cboTextFontFace.ListIndex)
                    .setTextLayerProperty ptp_FontSize, toolpanel_Text.tudTextFontSize.Value
                    .setTextLayerProperty ptp_FontColor, toolpanel_Text.csTextFontColor.Color
                    .setTextLayerProperty ptp_TextAntialiasing, toolpanel_Text.cboTextRenderingHint.ListIndex
                    .setTextLayerProperty ptp_TextContrast, toolpanel_Text.sltTextClarity.Value
                    .setTextLayerProperty ptp_FontBold, toolpanel_Text.btnFontStyles(0).Value
                    .setTextLayerProperty ptp_FontItalic, toolpanel_Text.btnFontStyles(1).Value
                    .setTextLayerProperty ptp_FontUnderline, toolpanel_Text.btnFontStyles(2).Value
                    .setTextLayerProperty ptp_FontStrikeout, toolpanel_Text.btnFontStyles(3).Value
                    .setTextLayerProperty ptp_HorizontalAlignment, toolpanel_Text.btsHAlignment.ListIndex
                    .setTextLayerProperty ptp_VerticalAlignment, toolpanel_Text.btsVAlignment.ListIndex
                End With
                
                'This is a little weird, but we also make sure to synchronize the current text rendering engine when the UI is synched.
                ' This is because this property changes according to the active text tool.
                pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_RenderingEngine, tre_WAPI
            
            Case VECTOR_FANCYTEXT
                
                With pdImages(g_CurrentImage).getActiveLayer
                    .setTextLayerProperty ptp_Text, toolpanel_FancyText.txtTextTool.Text
                    .setTextLayerProperty ptp_FontFace, toolpanel_FancyText.cboTextFontFace.List(toolpanel_FancyText.cboTextFontFace.ListIndex)
                    .setTextLayerProperty ptp_FontSize, toolpanel_FancyText.tudTextFontSize.Value
                    .setTextLayerProperty ptp_TextAntialiasing, toolpanel_FancyText.cboTextRenderingHint.ListIndex
                    .setTextLayerProperty ptp_TextHinting, CBool(toolpanel_FancyText.chkHinting.Value)
                    .setTextLayerProperty ptp_FontBold, toolpanel_FancyText.btnFontStyles(0).Value
                    .setTextLayerProperty ptp_FontItalic, toolpanel_FancyText.btnFontStyles(1).Value
                    .setTextLayerProperty ptp_FontUnderline, toolpanel_FancyText.btnFontStyles(2).Value
                    .setTextLayerProperty ptp_FontStrikeout, toolpanel_FancyText.btnFontStyles(3).Value
                    .setTextLayerProperty ptp_HorizontalAlignment, toolpanel_FancyText.btsHAlignment.ListIndex
                    .setTextLayerProperty ptp_VerticalAlignment, toolpanel_FancyText.btsVAlignment.ListIndex
                    .setTextLayerProperty ptp_WordWrap, toolpanel_FancyText.cboWordWrap.ListIndex
                    .setTextLayerProperty ptp_FillActive, CBool(toolpanel_FancyText.chkFillText.Value)
                    .setTextLayerProperty ptp_FillBrush, toolpanel_FancyText.bsText.Brush
                    .setTextLayerProperty ptp_OutlineActive, CBool(toolpanel_FancyText.chkOutlineText.Value)
                    .setTextLayerProperty ptp_OutlinePen, toolpanel_FancyText.psText.Pen
                    .setTextLayerProperty ptp_BackgroundActive, CBool(toolpanel_FancyText.chkBackground.Value)
                    .setTextLayerProperty ptp_BackgroundBrush, toolpanel_FancyText.bsTextBackground.Brush
                    .setTextLayerProperty ptp_BackBorderActive, CBool(toolpanel_FancyText.chkBackgroundBorder.Value)
                    .setTextLayerProperty ptp_BackBorderPen, toolpanel_FancyText.psTextBackground.Pen
                    .setTextLayerProperty ptp_LineSpacing, toolpanel_FancyText.tudLineSpacing.Value
                    .setTextLayerProperty ptp_MarginLeft, toolpanel_FancyText.tudMargin(0).Value
                    .setTextLayerProperty ptp_MarginRight, toolpanel_FancyText.tudMargin(1).Value
                    .setTextLayerProperty ptp_MarginTop, toolpanel_FancyText.tudMargin(2).Value
                    .setTextLayerProperty ptp_MarginBottom, toolpanel_FancyText.tudMargin(3).Value
                    .setTextLayerProperty ptp_CharInflation, toolpanel_FancyText.sltCharInflation.Value
                    .setTextLayerProperty ptp_CharJitterX, toolpanel_FancyText.tudJitter(0).Value
                    .setTextLayerProperty ptp_CharJitterY, toolpanel_FancyText.tudJitter(1).Value
                    .setTextLayerProperty ptp_CharMirror, toolpanel_FancyText.cboCharMirror.ListIndex
                    .setTextLayerProperty ptp_CharOrientation, toolpanel_FancyText.sltCharOrientation.Value
                    .setTextLayerProperty ptp_CharRemap, toolpanel_FancyText.cboCharCase.ListIndex
                    .setTextLayerProperty ptp_CharSpacing, toolpanel_FancyText.sltCharSpacing.Value
                End With
                
                'This is a little weird, but we also make sure to synchronize the current text rendering engine when the UI is synched.
                ' This is because this property changes according to the active text tool.
                pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_RenderingEngine, tre_PHOTODEMON
        
        End Select
        
        'Free the tool engine
        Tool_Support.setToolBusyState False
    
    End If
    
End Sub

