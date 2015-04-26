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

'The move tool uses these values to store the original layer offset
Private m_InitX As Double, m_InitY As Double

'The move tool uses these values to store the original layer canvas x/y modifiers
Private m_InitCanvasXMod As Double, m_InitCanvasYMod As Double

'If a point of interest is being modified by a tool action, its ID will be stored here.  Make sure to clear this value
' (to -1, which means "no point of interest") when you are finished with it (typically after MouseUp).
Private curPOI As Long

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

Public Sub setToolBusyState(ByVal NewState As Boolean)
    m_ToolIsBusy = NewState
End Sub

Public Function getCustomToolState() As Long
    getCustomToolState = m_CustomToolMarker
    m_CustomToolMarker = 0
End Function

Public Sub setCustomToolState(ByVal NewState As Long)
    m_CustomToolMarker = NewState
End Sub

'When a tool is finished processing, it can call this function to release all tool tracking variables
Public Sub terminateGenericToolTracking()
    
    'Reset the current POI, if any
    curPOI = -1
    
End Sub

'The move tool uses this function to set the initial layer offsets for a move operation
Public Sub setInitialLayerOffsets(ByRef srcLayer As pdLayer, Optional ByVal relevantPOI As Long = -1)
    
    'Store the layer's initial offset values (before any MouseMove events have occurred)
    m_InitX = srcLayer.getLayerOffsetX
    m_InitY = srcLayer.getLayerOffsetY
    
    'Store the layer's initial canvas x/y offset values
    m_InitCanvasXMod = srcLayer.getLayerCanvasXModifier
    m_InitCanvasYMod = srcLayer.getLayerCanvasYModifier
    
    'If a relevant POI was supplied, store it as well
    curPOI = relevantPOI
    
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

    'If the canvas in question has a horizontal scrollbar, process it
    If srcCanvas.getScrollVisibility(PD_HORIZONTAL) Then
    
        'Calculate a new scroll value
        Dim hOffset As Long
        hOffset = (initX - curX)
        
        'When zoomed-in, sub-pixel scrolling is not allowed.  Compensate for that now
        If g_Zoom.getZoomValue(srcImage.currentZoomValue) > 1 Then
            hOffset = hOffset / g_Zoom.getZoomOffsetFactor(srcImage.currentZoomValue)
        End If
        
        'Factor in the initial scroll bar value
        hOffset = m_InitHScroll + hOffset
        
        'If that value lies within the bounds of the scroll bar, apply it
        If (hOffset < srcCanvas.getScrollMin(PD_HORIZONTAL)) Then
            srcCanvas.setScrollValue PD_HORIZONTAL, srcCanvas.getScrollMin(PD_HORIZONTAL)
        ElseIf (hOffset > srcCanvas.getScrollMax(PD_HORIZONTAL)) Then
            srcCanvas.setScrollValue PD_HORIZONTAL, srcCanvas.getScrollMax(PD_HORIZONTAL)
        Else
            srcCanvas.setScrollValue PD_HORIZONTAL, hOffset
        End If
    
    End If
    
    'If the canvas in question has a vertical scrollbar, process it
    If srcCanvas.getScrollVisibility(PD_VERTICAL) Then
    
        'Calculate a new scroll value
        Dim vOffset As Long
        vOffset = (initY - curY)
        
        'When zoomed-in, sub-pixel scrolling is not allowed.  Compensate for that now
        If g_Zoom.getZoomValue(srcImage.currentZoomValue) > 1 Then
            vOffset = vOffset / g_Zoom.getZoomOffsetFactor(srcImage.currentZoomValue)
        End If
        
        'Factor in the initial scroll bar value
        vOffset = m_InitVScroll + vOffset
        
        'If that value lies within the bounds of the scroll bar, apply it
        If (vOffset < srcCanvas.getScrollMin(PD_VERTICAL)) Then
            srcCanvas.setScrollValue PD_VERTICAL, srcCanvas.getScrollMin(PD_VERTICAL)
        ElseIf (vOffset > srcCanvas.getScrollMax(PD_VERTICAL)) Then
            srcCanvas.setScrollValue PD_VERTICAL, srcCanvas.getScrollMax(PD_VERTICAL)
        Else
            srcCanvas.setScrollValue PD_VERTICAL, vOffset
        End If
    
    End If
    
    'Reinstate canvas redraws
    srcCanvas.setRedrawSuspension False
    
    'Request the scroll-specific viewport pipeline stage
    Viewport_Engine.Stage3_ExtractRelevantRegion srcImage, FormMain.mainCanvas(0)
    
End Sub

'The nav tool uses this function to move and/or resize the current layer.
' If this action occurs during a Mouse_Up event, the finalizeTransform parameter should be set to TRUE.
' This will instruct the function to forward the request to PD's central processor, so it can generate
' Undo/Redo data, be recorded as part of macros, etc.
Public Sub transformCurrentLayer(ByVal initX As Long, ByVal initY As Long, ByVal curX As Long, ByVal curY As Long, ByRef srcImage As pdImage, ByRef srcCanvas As pdCanvas, Optional ByVal isShiftDown As Boolean = False, Optional ByVal finalizeTransform As Boolean = False)
    
    'Prevent the canvas from redrawing itself until our movement calculations are complete.
    ' (This prevents juddery movement.)
    srcCanvas.setRedrawSuspension True
    
    'Also, mark the tool engine as busy
    Tool_Support.setToolBusyState True
    
    'Start by converting the mouse coordinates we were passed from screen units to image units
    Dim initImgX As Double, initImgY As Double, curImgX As Double, curImgY As Double
    convertCanvasCoordsToImageCoords srcCanvas, srcImage, initX, initY, initImgX, initImgY
    convertCanvasCoordsToImageCoords srcCanvas, srcImage, curX, curY, curImgX, curImgY
    
    'Calculate offsets between the initial mouse coordinates and the current ones
    Dim hOffset As Long, vOffset As Long
    hOffset = curImgX - initImgX
    vOffset = curImgY - initImgY
    
    'To help us more easily process the transformation's effect on the layer, store the layer's original position
    ' and size inside a RECT.  Note that we make two copies: one with canvas modifications (such as dynamic x/y
    ' changes caused by on-canvas resizing), and one without.
    Dim origLayerRect As RECT, origLayerRectModified As RECT
    Layer_Handler.fillRectForLayer srcImage.getActiveLayer, origLayerRect
    Layer_Handler.fillRectForLayer srcImage.getActiveLayer, origLayerRectModified, True
    
    'Calculate original width/height values, which will simplify our x/y modifier calculations later on
    Dim origWidth As Double, origHeight As Double
    origWidth = origLayerRect.Right - origLayerRect.Left
    origHeight = origLayerRect.Bottom - origLayerRect.Top
    If origWidth < 1 Then origWidth = 1
    If origHeight < 1 Then origHeight = 1
    
    'To prevent the user from flipping or mirroring the image, we must do some bound checking on their changes,
    ' and disallow anything that results in an invalid image coordinate.
    Dim newX As Double, newY As Double
    
    'The way we assign new offsets to the layer depends on the POI (point of interest) the user has used to move the image.
    ' Layers currently support five points of interest: each of their 4 corners, and anywhere in the layer interior
    ' (for moving the layer).
    
    'Check the POI we were given, and update the layer accordingly.
    With srcImage.getActiveLayer
    
        Select Case curPOI
            
            '-1: the mouse is not over the layer.  Do nothing.
            Case -1
                srcCanvas.setRedrawSuspension False
                Exit Sub
                
            '0: the mouse is dragging the top-left corner of the layer
            Case 0
                newX = m_InitX + hOffset
                newY = m_InitY + vOffset
                If newX > origLayerRectModified.Right - 1 Then newX = origLayerRectModified.Right - 1
                If newY > origLayerRectModified.Bottom - 1 Then newY = origLayerRectModified.Bottom - 1
                .setLayerOffsetX newX
                .setLayerOffsetY newY
                .setLayerCanvasXModifier (origLayerRectModified.Right - .getLayerOffsetX) / origWidth
                .setLayerCanvasYModifier (origLayerRectModified.Bottom - .getLayerOffsetY) / origHeight
                
                'If the user is pressing the SHIFT key, lock the image's aspect ratio
                If isShiftDown Then
                    .setLayerCanvasXModifier .getLayerCanvasYModifier
                    .setLayerOffsetX origLayerRectModified.Right - (.getLayerCanvasXModifier * origWidth)
                End If
            
            '1: top-right corner
            Case 1
                newY = m_InitY + vOffset
                If newY > origLayerRectModified.Bottom - 1 Then newY = origLayerRectModified.Bottom - 1
                .setLayerOffsetY newY
                .setLayerCanvasXModifier (curImgX - origLayerRect.Left) / origWidth
                .setLayerCanvasYModifier (origLayerRectModified.Bottom - .getLayerOffsetY) / origHeight
                
                'If the user is pressing the SHIFT key, lock the image's aspect ratio
                If isShiftDown Then .setLayerCanvasXModifier .getLayerCanvasYModifier
            
            '2: bottom-right
            Case 2
                .setLayerCanvasXModifier (curImgX - origLayerRect.Left) / origWidth
                .setLayerCanvasYModifier (curImgY - origLayerRect.Top) / origHeight
                
                'If the user is pressing the SHIFT key, lock the image's aspect ratio
                If isShiftDown Then .setLayerCanvasYModifier .getLayerCanvasXModifier
            
            '3: bottom-left
            Case 3
                newX = m_InitX + hOffset
                If newX > origLayerRectModified.Right - 1 Then newX = origLayerRectModified.Right - 1
                .setLayerOffsetX newX
                .setLayerCanvasXModifier (origLayerRectModified.Right - .getLayerOffsetX) / origWidth
                .setLayerCanvasYModifier (curImgY - origLayerRect.Top) / origHeight
                
                'If the user is pressing the SHIFT key, lock the image's aspect ratio
                If isShiftDown Then .setLayerCanvasYModifier .getLayerCanvasXModifier
            
            '4: interior of the layer (e.g. move the layer instead of resize it)
            Case 4
                .setLayerOffsetX m_InitX + hOffset
                .setLayerOffsetY m_InitY + vOffset
            
        End Select
        
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
        Select Case curPOI
        
            'Resize transformations.  (Note that resize transformations include some layer movement as well.)
            Case 0 To 3
            
                With srcImage.getActiveLayer
                    Process "Resize layer (on-canvas)", False, buildParams(.getLayerOffsetX, .getLayerOffsetY, .getLayerCanvasXModifier, .getLayerCanvasYModifier), UNDO_LAYERHEADER
                End With
            
            'Move-only transformations
            Case 4
                
                With srcImage.getActiveLayer
                    Process "Move layer", False, buildParams(.getLayerOffsetX, .getLayerOffsetY), UNDO_LAYERHEADER
                End With
                
            'The caller can specify other dummy values if they don't want us to redraw the screen
        
        End Select
    
    'If the transformation is still active (e.g. the user has the mouse pressed down), just redraw the viewport, but don't
    ' process Undo/Redo or any macro stuff.
    Else
    
        'Manually request a canvas redraw
        Viewport_Engine.Stage2_CompositeAllLayers srcImage, srcCanvas
    
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
    For i = 0 To toolbar_Options.sltQuickFix.Count - 1
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
                
                If (Not Processor.Processing) And (Not getToolBusyState) Then
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
    If Not canvasToolsAllowed(False) Then Exit Sub
    
    'To improve performance, we'll only sync the UI if a layer-specific tool is active, and the tool options panel is currently
    ' set to VISIBLE.
    If Not toolbar_Options.Visible Then Exit Sub
    
    Dim layerToolActive As Boolean
    
    Select Case g_CurrentTool
        
        Case NAV_MOVE
            layerToolActive = True
        
        Case VECTOR_TEXT
            If pdImages(g_CurrentImage).getActiveLayer.getLayerType = PDL_TEXT Then layerToolActive = True
        
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
                toolbar_Options.tudLayerMove(0).Value = pdImages(g_CurrentImage).getActiveLayer.getLayerOffsetX
                toolbar_Options.tudLayerMove(1).Value = pdImages(g_CurrentImage).getActiveLayer.getLayerOffsetY
                toolbar_Options.tudLayerMove(2).Value = pdImages(g_CurrentImage).getActiveLayer.getLayerWidth
                toolbar_Options.tudLayerMove(3).Value = pdImages(g_CurrentImage).getActiveLayer.getLayerHeight
                
                'The layer resize quality combo box also needs to be synched
                toolbar_Options.cboLayerResizeQuality.ListIndex = pdImages(g_CurrentImage).getActiveLayer.getLayerResizeQuality
            
            Case VECTOR_TEXT
                
                With toolbar_Options
                    .txtTextTool.Text = pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_Text)
                    .cboTextFontFace.setListIndexByString pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_FontFace)
                    .tudTextFontSize.Value = pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_FontSize)
                    .csTextFontColor.Color = pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_FontColor)
                    .cboTextRenderingHint.ListIndex = pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_TextRenderingHint)
                    .tudTextClarity.Value = pdImages(g_CurrentImage).getActiveLayer.getTextLayerProperty(ptp_TextContrast)
                End With
        
        End Select
        
        'Free the tool engine
        Tool_Support.setToolBusyState False
    
    End If
    
End Sub
