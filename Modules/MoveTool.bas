Attribute VB_Name = "Tools_Move"
'***************************************************************************
'PhotoDemon Move/Size Tool Manager
'Copyright 2014-2026 by Tanner Helland
'Created: 24/May/14
'Last updated: 06/May/24
'Last update: move rendering layer boundaries to universal "View > Show" menu; start work on new
'             "show distances" feature for the move tool, specifically.
'
'This module interfaces between the layer move/size UI and actual layer backend.  Look in the relevant
' tool panel form for more details on how the UI relays relevant tool data here.
'
'As of 2024, this module also handles move-related duties like snapping to various features.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'The move/size tool exposes a number of UI-only options (like drawing distances between objects).
' To improve viewport performance, we cache those settings locally, and the viewport queries us instead
' of directly querying the associated UI elements.
Private m_DrawDistances As Boolean, m_DrawCornerNodes As Boolean, m_DrawRotateNodes As Boolean

'Set to TRUE when LMB mouse is actively down; FALSE when LMB is not
Private m_LMBDown As Boolean

'Same goes for various selection-related move settings (for moving selected pixels).  These are simple
' flags whose value is relayed from the Move/Size options panel.
Private m_SelectionDefaultCut As Boolean, m_SelectionSampleMerged As Boolean

Public Sub DrawCanvasUI(ByRef dstCanvas As pdCanvas, ByRef srcImage As pdImage, Optional ByVal curPOI As PD_PointOfInterest = poi_Undefined, Optional ByVal srcIsTextLayer As Boolean = False)
    
    'Layer boundary rendering now exists as a standalone option in the View > Show menu
    ' (rather than being tied to the move tool, specifically).
    'If (Tools_Move.GetDrawLayerBorders() Or srcIsTextLayer) Then Drawing.DrawLayerBoundaries dstCanvas, srcImage, srcImage.GetActiveLayer
    
    'Tool-specific UI elements
    If (Tools_Move.GetDrawLayerCornerNodes() Or srcIsTextLayer) Then Drawing.DrawLayerCornerNodes dstCanvas, srcImage, srcImage.GetActiveLayer, curPOI
    If (Tools_Move.GetDrawLayerRotateNodes() Or srcIsTextLayer) Then Drawing.DrawLayerRotateNode dstCanvas, srcImage, srcImage.GetActiveLayer, curPOI
    If (m_LMBDown And Drawing.Get_ShowSmartGuides()) Then Drawing.DrawSmartGuides dstCanvas, srcImage
    If ((g_CurrentTool = NAV_MOVE) And Tools_Move.GetDrawDistances) Then Drawing.DrawLayerDistances dstCanvas, srcImage, srcImage.GetActiveLayer, curPOI
    
End Sub

Public Sub NotifyKeyDown(ByVal Shift As ShiftConstants, ByVal vkCode As Long, ByRef markEventHandled As Boolean)
    
    Dim hOffset As Long, vOffset As Long
    Dim canvasUpdateRequired As Boolean
        
    'Handle arrow keys first
    If (vkCode = VK_UP) Or (vkCode = VK_DOWN) Or (vkCode = VK_LEFT) Or (vkCode = VK_RIGHT) Then

        'Calculate offset modifiers for the current layer
        If (vkCode = VK_UP) Then vOffset = vOffset - 1
        If (vkCode = VK_DOWN) Then vOffset = vOffset + 1
        If (vkCode = VK_LEFT) Then hOffset = hOffset - 1
        If (vkCode = VK_RIGHT) Then hOffset = hOffset + 1
        
        canvasUpdateRequired = True
        
        'Apply the offsets
        With PDImages.GetActiveImage.GetActiveLayer
            .SetLayerOffsetX .GetLayerOffsetX + hOffset
            .SetLayerOffsetY .GetLayerOffsetY + vOffset
        End With
        
        'Redraw the viewport if necessary
        If canvasUpdateRequired Then
            markEventHandled = True
            Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
        End If
        
    'Handle non-arrow keys next
    Else
        
        'Delete key: delete the active layer (if allowed)
        If (vkCode = VK_DELETE) And (PDImages.GetActiveImage.GetNumOfLayers > 1) Then
            markEventHandled = True
            Process "Delete layer", False, TextSupport.BuildParamList("layerindex", PDImages.GetActiveImage.GetActiveLayerIndex), UNDO_Image_VectorSafe
        End If
        
        'Insert: raise Add New Layer dialog
        If (vkCode = VK_INSERT) Then
            markEventHandled = True
            Process "Add new layer", True
        End If
                
        'Tab and Shift+Tab: move through layer stack
        If (vkCode = VK_TAB) Then
            
            markEventHandled = True
            
            'Retrieve the active layer index
            Dim curLayerIndex As Long
            curLayerIndex = PDImages.GetActiveImage.GetActiveLayerIndex
            
            'Advance the layer index according to the Shift key
            If ((Shift And vbShiftMask) <> 0) Then curLayerIndex = curLayerIndex + 1 Else curLayerIndex = curLayerIndex - 1
            If (curLayerIndex < 0) Then curLayerIndex = PDImages.GetActiveImage.GetNumOfLayers - 1
            If (curLayerIndex > PDImages.GetActiveImage.GetNumOfLayers - 1) Then curLayerIndex = 0
            
            'Activate the new layer, then redraw the viewport and interface to match
            PDImages.GetActiveImage.SetActiveLayerByIndex curLayerIndex
            Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
            SyncInterfaceToCurrentImage
            
        End If
                
        'Space bar: toggle active layer visibility
        If (vkCode = VK_SPACE) Then
            markEventHandled = True
            PDImages.GetActiveImage.GetActiveLayer.SetLayerVisibility (Not PDImages.GetActiveImage.GetActiveLayer.GetLayerVisibility)
            Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
            Interface.SyncInterfaceToCurrentImage
        End If
        
        'Control can be used to "jump" the layer to the current mouse position
        If (vkCode = VK_CONTROL) Then Message "Ctrl+click to move the active layer here"
        
    End If
    
End Sub

'We also cover hand tool key behavior in this module, despite it being a separate tool.
' (There is not currently a separate hand-tool module.)
Public Sub NotifyKeyDown_HandTool(ByVal Shift As ShiftConstants, ByVal vkCode As Long, ByRef markEventHandled As Boolean)
        
    'Handle arrow keys as standard scroll events.
    If (vkCode = VK_UP) Or (vkCode = VK_DOWN) Or (vkCode = VK_LEFT) Or (vkCode = VK_RIGHT) Then
    
        'Set focus to the canvas (if it isn't already)
        FormMain.MainCanvas(0).SetFocusToCanvasView
        
        If (vkCode = VK_UP) Then
            FormMain.MainCanvas(0).CanvasView_MouseWheelVertical 0&, 0&, FormMain.MainCanvas(0).GetLastMouseX, FormMain.MainCanvas(0).GetLastMouseY, 1#
        ElseIf (vkCode = VK_DOWN) Then
            FormMain.MainCanvas(0).CanvasView_MouseWheelVertical 0&, 0&, FormMain.MainCanvas(0).GetLastMouseX, FormMain.MainCanvas(0).GetLastMouseY, -1#
        ElseIf (vkCode = VK_LEFT) Then
            FormMain.MainCanvas(0).CanvasView_MouseWheelHorizontal 0&, 0&, FormMain.MainCanvas(0).GetLastMouseX, FormMain.MainCanvas(0).GetLastMouseY, -1#
        ElseIf (vkCode = VK_RIGHT) Then
            FormMain.MainCanvas(0).CanvasView_MouseWheelHorizontal 0&, 0&, FormMain.MainCanvas(0).GetLastMouseX, FormMain.MainCanvas(0).GetLastMouseY, 1#
        End If
        
        markEventHandled = True
        
    End If
    
End Sub

Public Sub NotifyMouseDown(ByRef srcCanvas As pdCanvas, ByVal Shift As ShiftConstants, ByVal imgX As Single, ByVal imgY As Single)
    
    'Failsafe check only
    If (Not PDImages.IsImageActive) Then Exit Sub
    
    m_LMBDown = True
    
    'See if a selection is active.  If it is, we need to see if the user has clicked within the selected region.
    ' (If they have, we will allow them to move just the *selected* pixels.)
    Dim useSelectedPixels As Boolean: useSelectedPixels = False
    If PDImages.GetActiveImage.IsSelectionActive Then
        useSelectedPixels = PDImages.GetActiveImage.MainSelection.IsPointSelected(imgX, imgY)
    End If
    
    'Some move settings allow for additional parameters to be passed (such as the selection check
    ' we just performed above)
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.AddParam "use-selected-pixels", useSelectedPixels
    
    'See if the control key is down; if it is, we want to move the active layer to the current position.
    ' (Note that this behavior does *NOT* work with selections, by design, because there's no way for
    ' the user to clarify the desired move source - unlike a normal click-drag operation, which explicitly
    ' identifies the source on the initial click.)
    If ((Shift And vbCtrlMask) = vbCtrlMask) Then
        
        With cParams
            .AddParam "layer-offsetx", imgX
            .AddParam "layer-offsety", imgY
        End With
        
        Process "Move layer", False, cParams.GetParamString(), UNDO_LayerHeader
        
    Else
        
        'Prior to moving or transforming a layer, we need to check the state of the "auto-activate layer beneath mouse"
        ' option; if it is set, check (and possibly modify) the active layer based on the mouse position.
        If toolpanel_MoveSize.chkAutoActivateLayer.Value Then
            
            Dim layerUnderMouse As Long
            layerUnderMouse = Layers.GetLayerUnderMouse(imgX, imgY, True)
            
            'The "GetLayerUnderMouse" function returns a layer index >= 0 *if* the mouse is over a layer.
            If (layerUnderMouse >= 0) Then
            
                'If the layer under the mouse is not already active, activate it now
                If (layerUnderMouse <> PDImages.GetActiveImage.GetActiveLayerIndex) Then
                    Layers.SetActiveLayerByIndex layerUnderMouse, False
                    Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
                End If
            
            End If
        
        End If
        
        'Initiate the layer transformation engine.
        ' (Note that nothing will happen until the user actually moves the mouse.)
        '
        'If a selection is active, the only valid transform is movement.  Otherwise, the transform may
        ' be moving or resizing or rotating or some combination of these.
        Dim curPOI As PD_PointOfInterest
        curPOI = PDImages.GetActiveImage.GetActiveLayer.CheckForPointOfInterest(imgX, imgY, Tools_Move.GetDrawLayerRotateNodes())
        
        'Give preferential treatment to corner and edge nodes; if neither of these are selected,
        ' we then allow the selected area to "take over".
        If useSelectedPixels Then
            If (curPOI = poi_Interior) Or (curPOI = poi_Undefined) Then
                curPOI = poi_Interior
            Else
                useSelectedPixels = False
            End If
        End If
        
        Tools.SetInitialLayerToolValues PDImages.GetActiveImage(), PDImages.GetActiveImage.GetActiveLayer, imgX, imgY, curPOI, useSelectedPixels, Shift
        
    End If
    
End Sub

Public Function NotifyMouseMove(ByVal lmbDown As Boolean, ByVal Shift As ShiftConstants, ByVal imgX As Single, ByVal imgY As Single) As Long
    
    'Left mouse button down
    If lmbDown Then
        Message "Shift key: preserve layer aspect ratio", "DONOTLOG"
        Tools.TransformCurrentLayer imgX, imgY, PDImages.GetActiveImage(), PDImages.GetActiveImage.GetActiveLayer, FormMain.MainCanvas(0), (Shift And vbShiftMask)
    
    'Left mouse button *not* down
    Else
        
        'If a selection is active, the selection region will take precedence over any layer(s)
        ' beneath the current selection.  Determine this in advance.
        Dim mouseOverSelection As Boolean: mouseOverSelection = False
        If PDImages.GetActiveImage.IsSelectionActive Then
            mouseOverSelection = PDImages.GetActiveImage.MainSelection.IsPointSelected(imgX, imgY)
        End If
        
        'To spare the debug logger from receiving too many events, we will forcibly prevent logging
        ' basic status bar messages while in debug mode.
        Dim strAppendForDebug As String
        If UserPrefs.GenerateDebugLogs Then strAppendForDebug = "DONOTLOG" Else strAppendForDebug = vbNullString
        
        'If the Ctrl key is down, the user can ctrl+click to "jump" the active layer to
        ' the current mouse position.  We do not want to display target layer information
        ' in this case, as the "auto-select layer under mouse" behavior will be disabled.
        If ((Shift And vbCtrlMask) = 0) Then
            
            'If the "auto-activate layer beneath mouse" option is active, report the current layer name in the message bar;
            ' this is helpful for letting the user know which layer will be affected by an action in the current position.
            If toolpanel_MoveSize.chkAutoActivateLayer.Value Then
                
                Dim layerUnderMouse As Long
                layerUnderMouse = Layers.GetLayerUnderMouse(imgX, imgY, True)
                If (layerUnderMouse >= 0) Then
                
                    NotifyMouseMove = layerUnderMouse
                    
                    If mouseOverSelection Then
                        Message "Target: selected pixels", strAppendForDebug
                    Else
                        Message "Target layer: %1", PDImages.GetActiveImage.GetLayerByIndex(layerUnderMouse).GetLayerName, strAppendForDebug
                    End If
                    
                'The mouse is *not* over a layer.  Default to the active layer, which allows the user
                ' to interact with the layer even if it lies off-canvas.
                Else
                
                    NotifyMouseMove = PDImages.GetActiveImage.GetActiveLayerIndex
                    
                    If mouseOverSelection Then
                        Message "Target: selected pixels", strAppendForDebug
                    Else
                        Message "Target layer: %1", g_Language.TranslateMessage("(none)"), strAppendForDebug
                    End If
                    
                End If
            
            'Auto-activation is disabled.  In this state we only need to differentiate between the active layer
            ' and pixels under the current selection (if one exists).
            Else
                
                NotifyMouseMove = PDImages.GetActiveImage.GetActiveLayerIndex
                
                If mouseOverSelection Then
                    Message "Target: selected pixels", strAppendForDebug
                Else
                    Message "Target layer: %1", PDImages.GetActiveImage.GetActiveLayer.GetLayerName(), strAppendForDebug
                End If
                
            End If
        
        '/end Ctrl key down
        End If
            
    End If

End Function

Public Sub NotifyMouseUp(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal imgX As Single, ByVal imgY As Single, ByVal numOfMouseMovements As Long)
    
    m_LMBDown = False
    
    'Pass a final transform request to the layer handler.  This will initiate Undo/Redo creation, among other things.
    ' (Note Aug 2024: this line previously only triggered after this check:
    ' If (numOfMouseMovements > 0) Then
    ' This check breaks "move selected pixels" behavior as reported here: https://github.com/tannerhelland/PhotoDemon/issues/584
    ' I have now removed the check, but am leaving this comment pending additional testing to see if any
    ' unexpected interactions occur.
    Tools.TransformCurrentLayer imgX, imgY, PDImages.GetActiveImage(), PDImages.GetActiveImage.GetActiveLayer, FormMain.MainCanvas(0), (Shift And vbShiftMask), True
    
    'Reset the generic tool mouse tracking function
    Tools.TerminateGenericToolTracking
                
End Sub

'Relay functions for layer move/size node and border rendering
Public Function GetDrawDistances() As Boolean
    GetDrawDistances = m_DrawDistances
End Function

Public Function GetDrawLayerCornerNodes() As Boolean
    GetDrawLayerCornerNodes = m_DrawCornerNodes
End Function

Public Function GetDrawLayerRotateNodes() As Boolean
    GetDrawLayerRotateNodes = m_DrawRotateNodes
End Function

Public Sub SetDrawDistances(ByVal newState As Boolean)
    m_DrawDistances = newState
End Sub

Public Sub SetDrawLayerCornerNodes(ByVal newState As Boolean)
    m_DrawCornerNodes = newState
End Sub

Public Sub SetDrawLayerRotateNodes(ByVal newState As Boolean)
    m_DrawRotateNodes = newState
End Sub

'Relay functions for move selected pixels behavior
Public Function GetMoveSelectedPixels_DefaultCut() As Boolean
    GetMoveSelectedPixels_DefaultCut = m_SelectionDefaultCut
End Function

Public Function GetMoveSelectedPixels_SampleMerged() As Boolean
    GetMoveSelectedPixels_SampleMerged = m_SelectionSampleMerged
End Function

Public Sub SetMoveSelectedPixels_DefaultCut(ByVal newState As Boolean)
    m_SelectionDefaultCut = newState
End Sub

Public Sub SetMoveSelectedPixels_SampleMerged(ByVal newState As Boolean)
    m_SelectionSampleMerged = newState
End Sub
