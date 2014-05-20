Attribute VB_Name = "Tool_Support"
'***************************************************************************
'Helper functions for various PhotoDemon tools
'Copyright ©2013-2014 by Tanner Helland
'Created: 06/February/14
'Last updated: 05/May/14
'Last update: add bounds checking to transformCurrentLayer to prevent the user from resizing an image layer into oblivion
'
'To keep the pdCanvas user control codebase lean, much of its MouseMove events redirect here, to specialized
' functions that take mouse actions on the canvas and translate them into tool actions.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'The drag-to-pan tool uses these values to store the original image offset
Private m_InitHScroll As Long, m_InitVScroll As Long

'The move tool uses these values to store the original layer offset
Private m_InitX As Double, m_InitY As Double

'The move tool uses these values to store the original layer canvas x/y modifiers
Private m_InitCanvasXMod As Double, m_InitCanvasYMod As Double

'If a point of interest is being modified by a tool action, its ID will be stored here.  Make sure to clear this value
' (to -1, which means "no point of interest") when you are finished with it (typically after MouseUp).
Private curPOI As Long

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
        If srcImage.currentZoomValue < g_Zoom.getZoom100Index Then
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
        If srcImage.currentZoomValue < g_Zoom.getZoom100Index Then
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
    
    'Manually request a canvas redraw
    ScrollViewport srcImage, srcCanvas

End Sub

'The nav tool uses this function to move and/or resize the current layer.
' If this action occurs during a Mouse_Up event, the finalizeTransform parameter should be set to TRUE.
' This will instruct the function to forward the request to PD's central processor, so it can generate
' Undo/Redo data, be recorded as part of macros, etc.
Public Sub transformCurrentLayer(ByVal initX As Long, ByVal initY As Long, ByVal curX As Long, ByVal curY As Long, ByRef srcImage As pdImage, ByRef srcCanvas As pdCanvas, Optional ByVal isShiftDown As Boolean = False, Optional ByVal finalizeTransform As Boolean = False)
    
    'Prevent the canvas from redrawing itself until our movement calculations are complete.
    ' (This prevents juddery movement.)
    srcCanvas.setRedrawSuspension True
    
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
                If isShiftDown Then .setLayerCanvasYModifier .getLayerCanvasXModifier
            
            '1: top-right corner
            Case 1
                newY = m_InitY + vOffset
                If newY > origLayerRectModified.Bottom - 1 Then newY = origLayerRectModified.Bottom - 1
                .setLayerOffsetY newY
                .setLayerCanvasXModifier (curImgX - origLayerRect.Left) / origWidth
                .setLayerCanvasYModifier (origLayerRectModified.Bottom - .getLayerOffsetY) / origHeight
                
                'If the user is pressing the SHIFT key, lock the image's aspect ratio
                If isShiftDown Then .setLayerCanvasYModifier .getLayerCanvasXModifier
            
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
        
        End Select
    
    Else
    
        'Manually request a canvas redraw
        ScrollViewport srcImage, srcCanvas
    
    End If
    
End Sub

