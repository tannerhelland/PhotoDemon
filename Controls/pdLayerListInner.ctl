VERSION 5.00
Begin VB.UserControl pdLayerListInner 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "pdLayerListInner.ctx":0000
   Begin PhotoDemon.pdTextBox txtLayerName 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
   End
End
Attribute VB_Name = "pdLayerListInner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Layer Listbox (inner portion only)
'Copyright 2014-2026 by Tanner Helland
'Created: 25/March/14
'Last updated: 18/April/22
'Last update: fix UI glitch when starting to move layer that is *not* the active layer
'
'In a surprise to precisely no one, PhotoDemon has many unique UI/UX needs - needs that the
' intrinsic VB controls can't handle.  These range from the obnoxious (lack of an "autosize"
' property for anything but labels) to the critical (no Unicode support).
'
'As such, I've created many of my own UCs for the program.  All are owner-drawn, with the
' goal of maintaining visual fidelity across the program, while also enabling key features
' like Unicode support.
'
'A few notes on this layer listbox control, specifically:
'
' 1) This control bares no relation to pdListBox, for better or worse.
' 2) High DPI settings are handled automatically.
' 3) A hand cursor is automatically applied, and clicks are returned via the Click event.
' 4) Color decisions are automatically handled by PD's internal theming engine.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This control does not contain a scrollbar; instead, it bubbles scroll-related events upward, so its parent can
' show/hide/update an embedded scroll bar accordingly.
Public Event ScrollMaxChanged(ByVal newMax As Long)
Public Event ScrollValueChanged(ByVal newValue As Long)

'Because VB focus events are wonky, especially when we use CreateWindow within a UC, this control raises its own
' specialized focus events.  If you need to track focus, use these instead of the default VB functions.
Public Event GotFocusAPI()
Public Event LostFocusAPI()

'The rectangle where the list is actually rendered
Private m_ListRect As RectF

'User control support class.  Historically, many classes (and associated subclassers) were required by each user control,
' but I've since wrapped these into a single central support class.
Private WithEvents ucSupport As pdUCSupport
Attribute ucSupport.VB_VarHelpID = -1

'Local list of themable colors.  This list includes all potential colors used by the control, regardless of state change
' or internal control settings.  The list is updated by calling the UpdateColorList function.
' (Note also that this list does not include variants, e.g. "BorderColor" vs "BorderColor_Hovered".  Variant values are
'  automatically calculated by the color management class, and they are retrieved by passing boolean modifiers to that
'  class, rather than treating every imaginable variant as a separate constant.)
Private Enum PDLAYERBOX_COLOR_LIST
    [_First] = 0
    PDLB_Background = 0
    PDLB_Border = 1
    PDLB_SelectedItemFill = 2
    PDLB_SelectedItemBorder = 3
    PDLB_SelectedItemText = 4
    PDLB_UnselectedItemFill = 5
    PDLB_UnselectedItemBorder = 6
    PDLB_UnselectedItemText = 7
    [_Last] = 7
    [_Count] = 8
End Enum

'Color retrieval and storage is handled by a dedicated class; this allows us to optimize theme interactions,
' without worrying about the details locally.
Private m_Colors As pdThemeColors

'This control needs to store thumbnails for all layers in the current image.  This storage needs to be easy
' to manipulate, not just because layers can change frequently - but because layer *order* can also change frequently.
' For this reason, layers are tracked by their canonical ID, *not* their index (position in the layer stack).
Private Type LayerThumbDisplay
    spriteIDInt As Long
    spriteIDStr As String
    CanonicalLayerID As Long
End Type

Private m_LayerThumbnails() As LayerThumbDisplay
Private m_NumOfThumbnails As Long
Private m_ThumbCollection As pdSpriteSheet

'Thumbnail image width/height values, declared as variables so we can dynamically adjust them at run-time.
Private m_ThumbWidth As Long, m_ThumbHeight As Long

'Height of each layer content block.  Note that this is effectively a "magic number", in pixels, representing the
' height of each layer block in the layer selection UI.  This number will be dynamically resized per the current
' screen DPI by the "RedrawLayerList" and "RenderLayerBlock" functions.
Private Const LAYER_BLOCK_HEIGHT As Long = 40&

'I don't want thumbnails to fill the full height of their blocks, so a border is automatically applied to each
' side of the thumbnail.  (Like all other interface elements, it is dynamically modified for DPI as necessary.)
Private Const THUMBNAIL_PADDING As Long = 4&

'Finally, some horizontal padding is applied to each element of a layer block (visibility icon, lock icon, thumb,
' layer name, etc)
Private Const HORIZONTAL_ITEM_PADDING As Long = 4&

'The currently hovered layer entry.  (Note that the currently *selected* layer is retrieved from the active
' pdImage object, rather than stored locally.)
Private m_CurLayerHover As Long

'The renderer highlights clickable elements, so it needs access to the current mouse coordinates
Private m_MouseX As Single, m_MouseY As Single

'Extra interface images are loaded as resources at run-time
Private img_EyeOpen As pdDIB, img_EyeClosed As pdDIB

'Some UI elements are dynamically rendered onto the layer box.  To simplify hit detection, their RECTs are stored
' at render-time, which allows the mouse actions to easily check hits regardless of layer box position.
Private m_LayerHoverRect As RECT, m_VisibilityRect As RECT, m_NameRect As RECT, m_NameEditRect As RECT

'While in OLE drag/drop mode (e.g. dragging files from Explorer), ignore any mouse actions on the main layer box
Private m_InOLEDragDropMode As Boolean

'While in our own custom layer box drag/drop mode (e.g. rearranging layers), this will be set to TRUE.
' Also, the layer-to-be-moved is tracked, as is the initial layer index (which is required for processing the final
' action, e.g. the one that triggers Undo/Redo creation).
Private m_LayerRearrangingMode As Boolean, m_LayerMovingID As Long, m_LayerIndexToRearrange As Long, m_InitialLayerIndex As Long
Private m_ActiveLayerIDatMouseDown As Long

'When the user is in "edit layer name" mode, this will be set to TRUE
Private m_LayerNameEditMode As Boolean

'When the mouse is over the layer list, this will be set to TRUE
Private m_MouseOverLayerBox As Boolean

'Current scroll bar value and maximum.  Note that this control does not possess a scroll bar; instead, it bubbles
' changes upward, and our parent control maintains the actual scroll bar object.
Private m_ScrollValue As Long, m_ScrollMax As Long

'Popup menu when a layer is right-clicked
Private WithEvents m_PopupMenu As pdPopupMenu
Attribute m_PopupMenu.VB_VarHelpID = -1
Private m_RightClickIndex As Long

'In 2020, I added a "flash" action; the canvas uses this to draw attention to the active layer when
' the user attempts to interact with it despite it being hidden.  (In the future, this could also be
' useful if layer locking is added.)
Private WithEvents m_FlashTimer As pdTimer
Attribute m_FlashTimer.VB_VarHelpID = -1
Private m_FlashCount As Long, m_FlashTimeElapsed As Long, m_FlashLength As Long

'Normally, we want to ensure the currently selected layer is forcibly made visible on-screen.
' In rare circumstances, however, (like toggling visibility via this box), we want to suspend
' this behavior.  Set this toggle to prevent auto-positioning, but you *must* remember to
' reset it to FALSE when you're done!
Private m_SuspendAutoShowActiveLayer As Boolean

Public Function GetControlType() As PD_ControlType
    GetControlType = pdct_LayerListInner
End Function

Public Function GetControlName() As String
    GetControlName = UserControl.Extender.Name
End Function

'Workaround for VB6 quirks; see VBHacks.InControlArray()
Public Function IsChildInControlArray(ByRef ctlChild As Object) As Boolean
    IsChildInControlArray = Not UserControl.Controls(ctlChild.Name) Is ctlChild
End Function

'The Enabled property is a bit unique; see http://msdn.microsoft.com/en-us/library/aa261357%28v=vs.60%29.aspx
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    UserControl.Enabled = newValue
    If ucSupport.AmIVisible() Then RedrawBackBuffer
    PropertyChanged "Enabled"
End Property

'hWnds aren't exposed by default
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'Container hWnd must be exposed for external tooltip handling
Public Property Get ContainerHwnd() As Long
    ContainerHwnd = UserControl.ContainerHwnd
End Property

'To support high-DPI settings properly, we expose specialized move+size functions
Public Function GetLeft() As Long
    GetLeft = ucSupport.GetControlLeft
End Function

Public Sub SetLeft(ByVal newLeft As Long)
    ucSupport.RequestNewPosition newLeft, , True
End Sub

Public Function GetTop() As Long
    GetTop = ucSupport.GetControlTop
End Function

Public Sub SetTop(ByVal newTop As Long)
    ucSupport.RequestNewPosition , newTop, True
End Sub

Public Function GetWidth() As Long
    GetWidth = ucSupport.GetControlWidth
End Function

Public Sub SetWidth(ByVal newWidth As Long)
    ucSupport.RequestNewSize newWidth, , True
End Sub

Public Function GetHeight() As Long
    GetHeight = ucSupport.GetControlHeight
End Function

Public Sub SetHeight(ByVal newHeight As Long)
    ucSupport.RequestNewSize , newHeight, True
End Sub

Public Sub SetPositionAndSize(ByVal newLeft As Long, ByVal newTop As Long, ByVal newWidth As Long, ByVal newHeight As Long)
    ucSupport.RequestFullMove newLeft, newTop, newWidth, newHeight, True
End Sub

Public Function ScrollMax() As Long
    ScrollMax = m_ScrollMax
End Function

Public Property Get ScrollValue() As Long
    ScrollValue = m_ScrollValue
End Property

Public Property Let ScrollValue(ByRef newValue As Long)
    If (m_ScrollValue <> newValue) Then
        m_ScrollValue = newValue
        RedrawBackBuffer
    End If
End Property

Private Sub m_FlashTimer_Timer()

    m_FlashCount = m_FlashCount + 1
    m_FlashTimeElapsed = m_FlashTimeElapsed + m_FlashTimer.Interval
    
    'Only flash for three seconds (by default; the caller can configure this manually)
    If (m_FlashTimeElapsed >= m_FlashLength) Then
        m_FlashTimer.StopTimer
        m_FlashCount = 0
        m_FlashTimeElapsed = 0
        RedrawBackBuffer True
    Else
        RedrawBackBuffer True
    End If
    
End Sub

Private Sub m_PopupMenu_MenuClicked(ByRef clickedMenuID As String, ByVal idxMenuTop As Long, ByVal idxMenuSub As Long)
    
    'Make sure a valid layer was clicked
    If (m_RightClickIndex < 0) Then Exit Sub
    
    'As a failsafe, before modifying the image, force any non-destructive changes (like visibility toggles)
    ' to be included in the current image state, because right-click toggles can do things like modify
    ' the active layer (which can cause non-destructive changes to be lost).
    PDImages.GetActiveImage.UndoManager.ForceLastUndoDataToIncludeEverything
    
    'All popup menu items simply reference core PD actions, but some require us to temporarily suspend action validation.
    ' (This is because some action validations can only be performed against the currently active layer, and via popup,
    ' the user can select other layers.)
    Actions.LaunchAction_ByName clickedMenuID, pdas_Menu, (clickedMenuID = "layer_mergeup") Or (clickedMenuID = "layer_mergedown"), m_RightClickIndex
    
End Sub

'If the layer name textbox is visible and the Enter key is pressed, commit the changed layer name and hide the text box
Private Sub txtLayerName_KeyPress(ByVal Shift As ShiftConstants, ByVal vKey As Long, preventFurtherHandling As Boolean)
    
    If (vKey = VK_RETURN) Then
        
        preventFurtherHandling = True
        
        'Set the active layer name, then hide the text box
        PDImages.GetActiveImage.GetActiveLayer.SetLayerName txtLayerName.Text
        
        'Fun fact: as of v8.0, PhotoDemon can "hide" animation frame data right inside
        ' a layer's name (e.g. Frame 1 (100 ms)).  This provides a convenient way for the
        ' user to tweak frame time settings in e.g. an animated GIF without popping into a
        ' separate screen.
        
        'But this brings new complexities.  In particular, if the user changes a layer's
        ' name, they may actually be changing the frame time for that layer.  We need to
        ' look for any changes potentially related to frame time, and update them accordingly.
        If PDImages.GetActiveImage.IsAnimated Then
        
            'Look for frame time changes
            Dim newFrameTime As Long, oldFrameTime As Long
            newFrameTime = Animation.GetFrameTimeFromLayerName(PDImages.GetActiveImage.GetActiveLayer.GetLayerName())
            oldFrameTime = PDImages.GetActiveImage.GetActiveLayer.GetLayerFrameTimeInMS()
            
            'If we parsed a viable frame time from the layer name, store it immediately
            If (newFrameTime >= 0) And (newFrameTime <> oldFrameTime) Then
                PDImages.GetActiveImage.GetActiveLayer.SetLayerFrameTimeInMS newFrameTime
            End If
            
            'If frame time was changed, we need to notify the animation window of the change
            layerpanel_Navigator.NotifyFrameTimeChange PDImages.GetActiveImage.GetActiveLayerIndex, newFrameTime
        
        End If
        
        'If the user changed the name, set an Undo/Redo point now.
        ' (Note that this also catches frame time changes, as the user can't change
        ' layer frame time without *also* changing layer name.)
        If Tools.CanvasToolsAllowed Then Processor.FlagFinalNDFXState_Generic pgp_Name, PDImages.GetActiveImage.GetActiveLayer.GetLayerName
        
        'Re-enable hotkeys now that editing is finished
        m_LayerNameEditMode = False
        
        'Redraw the layer box with the new name
        RedrawBackBuffer
        
        'Hide the text box
        txtLayerName.Visible = False
        txtLayerName.Text = vbNullString
        
        'Transfer focus back to the layer box itself (instead of letting Windows forward it somewhere random)
        If (Not g_WindowManager Is Nothing) Then g_WindowManager.SetFocusAPI Me.hWnd
        
    ElseIf (vKey = VK_ESCAPE) Then
        preventFurtherHandling = True
        
        If txtLayerName.Visible Then txtLayerName.Visible = False
        txtLayerName.Text = vbNullString
        If (Not g_WindowManager Is Nothing) Then g_WindowManager.SetFocusAPI Me.hWnd
        m_LayerNameEditMode = False
        
    Else
        preventFurtherHandling = False
    End If

End Sub

'If the text box loses focus mid-edit, hide it and discard any changes
Private Sub txtLayerName_LostFocusAPI()
    If txtLayerName.Visible Then txtLayerName.Visible = False
    m_LayerNameEditMode = False
End Sub

Private Sub ucSupport_CustomMessage(ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bHandled As Boolean, lReturn As Long)
    
    Select Case wMsg
        
        'Different screen profile selected; this requires a redraw of layer thumbnails
        Case WM_PD_COLOR_MANAGEMENT_CHANGE
            Me.RequestRedraw True
            
        'User attempted to interact with a hidden/locked layer; flash the active layer's status
        ' with wParam specifying flash length in ms (typically 500) and lParam specifying total
        ' flash duration in ms (typically 3000)
        Case WM_PD_FLASH_ACTIVE_LAYER
            Set m_FlashTimer = New pdTimer
            m_FlashTimer.Interval = IIf(wParam <= 0, 500, wParam)
            m_FlashLength = IIf(lParam <= 0, 3000, lParam)
            m_FlashTimeElapsed = 0
            m_FlashTimer.StartTimer
            
    End Select

End Sub

'Double-clicks on the layer box raise "layer title edit mode", if the mouse is within a layer's title area
Private Sub ucSupport_DoubleClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)

    'Ignore user interaction while in drag/drop mode
    If m_InOLEDragDropMode Then Exit Sub
    
    If (PDMath.IsPointInRect(x, y, m_NameRect) And (Button = pdLeftButton)) Then
        
        'Just in case, forcibly deactivate "layer rearrange" mode
        m_LayerRearrangingMode = False
        
        'Move the text layer box into position
        txtLayerName.SetPositionAndSize m_NameEditRect.Left, m_NameEditRect.Top, m_NameEditRect.Right - m_NameEditRect.Left, m_NameEditRect.Bottom - m_NameEditRect.Top
        txtLayerName.ZOrder 0
        txtLayerName.Visible = True
        
        'Disable hotkeys until editing is finished
        m_LayerNameEditMode = True
        
        'Fill the text box with the current layer name, and select it
        txtLayerName.Text = PDImages.GetActiveImage.GetLayerByIndex(GetLayerAtPosition(x, y)).GetLayerName
        
        'Set an Undo/Redo marker for the existing layer name
        Processor.FlagInitialNDFXState_Generic pgp_Name, PDImages.GetActiveImage.GetLayerByIndex(GetLayerAtPosition(x, y)).GetLayerName, PDImages.GetActiveImage.GetLayerByIndex(GetLayerAtPosition(x, y)).GetLayerID
        
        txtLayerName.SetFocusToEditBox
        txtLayerName.SelectAll
    
    'Hide the text box if it isn't already
    Else
        txtLayerName.Visible = False
    End If

End Sub

'When the control receives focus, if the focus isn't received via mouse click, display a focus rect around the active button
Private Sub ucSupport_GotFocusAPI()
    RedrawBackBuffer
    RaiseEvent GotFocusAPI
End Sub

'When the control loses focus, erase any focus rects it may have active
Private Sub ucSupport_LostFocusAPI()
    
    'Check for any non-destructive changes that may have been set via this window (e.g. visibility)
    If PDImages.IsImageActive() Then
        Processor.FlagFinalNDFXState_Generic pgp_Visibility, PDImages.GetActiveImage.GetActiveLayer.GetLayerVisibility
    End If
    
    RedrawBackBuffer
    RaiseEvent LostFocusAPI
    
End Sub

Private Sub ucSupport_ClickCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    
    'Ignore user interaction while in drag/drop mode
    If m_InOLEDragDropMode Then Exit Sub
    
    m_RightClickIndex = -1
    
    Dim clickedLayer As Long
    clickedLayer = GetLayerAtPosition(x, y)
    
    'Basic validation checks
    If (clickedLayer >= 0) Then
    If PDImages.IsImageActive() Then
        
        'Left button
        If ((Button And pdLeftButton) <> 0) Then
            
            'Check the clicked position against a series of rects, each rect representing a
            ' unique area of the layer box (with different allowed interactions).
            
            'Has the user clicked a visibility rectangle?
            If PDMath.IsPointInRect(x, y, m_VisibilityRect) Then
                
                'When toggling visibility, we want to suspend this control's default behavior
                ' of forcing the currently active layer on-screen.
                m_SuspendAutoShowActiveLayer = True
                
                'Look for a modifier key being down.  Alt/Shift/Ctrl can be used to ONLY
                ' show the current layer.
                
                'No modifiers; toggle visibility of only the clicked layer.
                ' (Note that we explicitly request the "ImageHeader" Undo type - this is because
                ' the undo engine only supports "LayerHeader" changes for the *currently active layer*,
                ' which may or may not be the one getting its visibility toggled.)
                If (Shift = 0) Then
                    Process "Toggle layer visibility", False, BuildParamList("layerindex", clickedLayer), UNDO_ImageHeader
                    
                'One or more modifiers
                Else
                    
                    'See if any other layers are visible *besides* the clicked layer
                    Dim otherLayersVisible As Boolean
                    otherLayersVisible = False
                    
                    Dim i As Long
                    For i = 0 To PDImages.GetActiveImage.GetNumOfLayers - 1
                        If (i <> clickedLayer) Then
                            otherLayersVisible = PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerVisibility()
                            If otherLayersVisible Then Exit For
                        End If
                    Next i
                    
                    'If one or more other layer are visible, turn all of them off and make only the
                    ' current layer visible.
                    If otherLayersVisible Then
                        Process "Show only this layer", False, BuildParamList("layerindex", clickedLayer), UNDO_ImageHeader
                        
                    'If all other layers in the image are invisible, make *all* layers visible.
                    Else
                        Process "Show all layers", False, vbNullString, UNDO_ImageHeader
                    End If
                    
                End If
                
                'Before exiting, reset the "show active layer on-screen automatically" flag
                m_SuspendAutoShowActiveLayer = False
            
            'The user has not clicked any item of interest.  Assume that they want to make the clicked layer
            ' the active layer.
            Else
            
                'See if the clicked layer differs from the active layer at _MouseDown; if it does, redraw the viewport.
                ' (This ensures correct UI behavior on tools like move/size that render decorations on the current layer.)
                If (m_ActiveLayerIDatMouseDown <> PDImages.GetActiveImage.GetLayerByIndex(clickedLayer).GetLayerID) Then
                    Processor.FlagFinalNDFXState_Generic pgp_Visibility, PDImages.GetActiveImage.GetActiveLayer.GetLayerVisibility
                    Layers.SetActiveLayerByIndex clickedLayer, False
                    Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
                End If
                
            End If
            
            'Redraw the layer box to represent any changes from this interaction.
            ' NOTE: this is not currently necessary, as all interactions automatically force a redraw on their own.
            'RedrawLayerBox
            
        ElseIf ((Button And pdRightButton) <> 0) Then
        
            'Note the clicked layer index; we'll need this if the user chooses to do something to this layer
            m_RightClickIndex = clickedLayer
            
            'Display the popup menu; it will raise additional events, as necessary
            ShowLayerPopupMenu x, y
            
        End If
                        
    'End failsafe checks
    End If
    End If
    
End Sub

'The user has right-clicked a specific layer.  The layer index is saved in the module-level m_RightClickIndex value.
Private Sub ShowLayerPopupMenu(ByVal srcX As Long, ByVal srcY As Long)
    
    m_PopupMenu.Reset
    
    'Construct the menu.  (The current layout of this menu largely mimics Photoshop.)
    With m_PopupMenu
        .AddMenuItem g_Language.TranslateMessage("Show this layer"), "layer_show", 0, menuIsChecked:=PDImages.GetActiveImage.GetLayerByIndex(m_RightClickIndex, True).GetLayerVisibility()
        .AddMenuItem g_Language.TranslateMessage("Show only this layer"), "layer_showonly", 1
        .AddMenuItem g_Language.TranslateMessage("Hide only this layer"), "layer_hideonly", 2
        .AddMenuItem "-", "-", 3
        .AddMenuItem g_Language.TranslateMessage("Duplicate layer"), "layer_duplicate", 4
        .AddMenuItem g_Language.TranslateMessage("Delete layer"), "layer_deletecurrent", 5, menuIsEnabled:=(PDImages.GetActiveImage.GetNumOfLayers() > 1)
        .AddMenuItem "-", "-", 6
        .AddMenuItem g_Language.TranslateMessage("Rasterize layer"), "layer_rasterizecurrent", 7, menuIsEnabled:=PDImages.GetActiveImage.GetLayerByIndex(m_RightClickIndex, True).IsLayerVector()
        .AddMenuItem "-", "-", 8
        
        Dim allowMergeUp As Boolean, allowMergeDown As Boolean
        If (PDImages.GetActiveImage.GetNumOfLayers() > 1) Then
            allowMergeUp = (Layers.IsLayerAllowedToMergeAdjacent(m_RightClickIndex, False) <> -1)
            allowMergeDown = (Layers.IsLayerAllowedToMergeAdjacent(m_RightClickIndex, True) <> -1)
        Else
            allowMergeUp = False
            allowMergeDown = False
        End If
        .AddMenuItem g_Language.TranslateMessage("Merge up"), "layer_mergeup", 9, menuIsEnabled:=allowMergeUp
        .AddMenuItem g_Language.TranslateMessage("Merge down"), "layer_mergedown", 10, menuIsEnabled:=allowMergeDown
        
        .AddMenuItem g_Language.TranslateMessage("Merge visible layers"), "image_mergevisible", 11, menuIsEnabled:=(PDImages.GetActiveImage.GetNumOfVisibleLayers() > 1)
        .AddMenuItem g_Language.TranslateMessage("Flatten image..."), "image_flatten", 12, menuIsEnabled:=(PDImages.GetActiveImage.GetNumOfLayers() > 1)
    End With
    
    m_PopupMenu.ShowMenu Me.hWnd, srcX, srcY
    
End Sub

'Key events are TODO
Private Sub ucSupport_KeyDownCustom(ByVal Shift As ShiftConstants, ByVal vkCode As Long, markEventHandled As Boolean)
        
    'Ignore user interaction while in drag/drop mode
    If m_InOLEDragDropMode Then Exit Sub
    
    'Ignore keypresses if the user is currently editing a layer name
    If m_LayerNameEditMode Then
        markEventHandled = False
        Exit Sub
    End If
    
    'Ignore key presses unless an image has been loaded
    If PDImages.IsImageActive() Then
    
        'Up key activates the next layer upward
        If (vkCode = VK_UP) And (PDImages.GetActiveImage.GetActiveLayerIndex < PDImages.GetActiveImage.GetNumOfLayers - 1) Then
            Layers.SetActiveLayerByIndex PDImages.GetActiveImage.GetActiveLayerIndex + 1, True
        End If
        
        'Down key activates the next layer downward
        If (vkCode = VK_DOWN) And PDImages.GetActiveImage.GetActiveLayerIndex > 0 Then
            Layers.SetActiveLayerByIndex PDImages.GetActiveImage.GetActiveLayerIndex - 1, True
        End If
        
        'Right key increases active layer opacity
        If (vkCode = VK_RIGHT) And (PDImages.GetActiveImage.GetActiveLayer.GetLayerVisibility) Then
            'TODO!  Bubble up opacity changes
            'sltLayerOpacity.Value = PDImages.GetActiveImage.GetActiveLayer.GetLayerOpacity + 10
            'Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.mainCanvas(0)
        End If
        
        'Left key decreases active layer opacity
        If (vkCode = VK_LEFT) And (PDImages.GetActiveImage.GetActiveLayer.GetLayerVisibility) Then
            'TODO!  Bubble up opacity changes
            'sltLayerOpacity.Value = PDImages.GetActiveImage.GetActiveLayer.GetLayerOpacity - 10
            'Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.mainCanvas(0)
        End If
        
        'Delete key: delete the active layer (if allowed)
        If (vkCode = VK_DELETE) And PDImages.GetActiveImage.GetNumOfLayers > 1 Then
            Process "Delete layer", False, BuildParamList("layerindex", PDImages.GetActiveImage.GetActiveLayerIndex), UNDO_Image_VectorSafe
        End If
        
        'Insert: raise Add New Layer dialog
        If (vkCode = VK_INSERT) Then
            Process "Add new layer", True
            If (Not g_WindowManager Is Nothing) Then g_WindowManager.SetFocusAPI Me.hWnd
        End If
        
        'Tab and Shift+Tab: move through layer stack
        If (vkCode = VK_TAB) Then
            
            'Retrieve the active layer index
            Dim curLayerIndex As Long
            curLayerIndex = PDImages.GetActiveImage.GetActiveLayerIndex
            
            'Advance the layer index according to the Shift modifier
            If (Shift And vbShiftMask) <> 0 Then
                curLayerIndex = curLayerIndex + 1
            Else
                curLayerIndex = curLayerIndex - 1
            End If
            
            'I'm currently working on letting the user tab through the layer list, then tab *out of the control* upon reaching
            ' the last layer.  But this requires some changes to the pdCanvas control (it's complicated), so this doesn't work just yet.
            If (curLayerIndex >= 0) And (curLayerIndex < PDImages.GetActiveImage.GetNumOfLayers) Then
                
                'Activate the new layer
                PDImages.GetActiveImage.SetActiveLayerByIndex curLayerIndex
                
                'Redraw the viewport and interface to match
                Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
                SyncInterfaceToCurrentImage
                
                'All that interface stuff may have messed up focus; retain it on the layer box
                If (Not g_WindowManager Is Nothing) Then g_WindowManager.SetFocusAPI Me.hWnd
            
            Else
                markEventHandled = False
            End If
            
        End If
        
        'Space bar: toggle active layer visibility
        If (vkCode = VK_SPACE) Then
            PDImages.GetActiveImage.GetActiveLayer.SetLayerVisibility (Not PDImages.GetActiveImage.GetActiveLayer.GetLayerVisibility)
            Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
            SyncInterfaceToCurrentImage
        End If
        
    End If

End Sub

'MouseDown is used for drag/drop layer reordering
Private Sub ucSupport_MouseDownCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)

    'Ignore user interaction while in drag/drop mode
    If m_InOLEDragDropMode Then Exit Sub
    
    'Retrieve the layer under this position
    Dim clickedLayer As Long
    clickedLayer = GetLayerAtPosition(x, y)
    
    'Don't proceed unless the user has the mouse over a valid layer
    If (clickedLayer >= 0) And PDImages.IsImageActive() Then
        
        'If the image is a multilayer image, and they're using the left mouse button,
        ' initiate drag/drop layer reordering - but *only* if the mouse is not in the
        ' "visibility toggle" segment of the control.
        If (PDImages.GetActiveImage.GetNumOfLayers > 1) And (Button = pdLeftButton) And (Not PDMath.IsPointInRect(x, y, m_VisibilityRect)) Then
            
            'Enter layer rearranging mode
            m_LayerRearrangingMode = True
            
            'Note the ID of the active layer *before* the mouse button was pressed - we may activate a new layer,
            ' and at _MouseUp, we'll need to account for that.
            m_ActiveLayerIDatMouseDown = PDImages.GetActiveImage.GetActiveLayerID
            
            'If the selected layer is not the active one, make it so
            If (clickedLayer <> PDImages.GetActiveImage.GetActiveLayerIndex) Then Layers.SetActiveLayerByIndex clickedLayer, False
            
            'Note the layer being rearranged
            m_LayerMovingID = PDImages.GetActiveImage.GetLayerByIndex(clickedLayer, True).GetLayerID
            m_LayerIndexToRearrange = clickedLayer
            m_InitialLayerIndex = m_LayerIndexToRearrange
        
        End If
        
    End If

End Sub

Private Sub ucSupport_MouseEnter(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_MouseOverLayerBox = True
    ucSupport.RequestCursor IDC_HAND
    RedrawBackBuffer
End Sub

'When the mouse leaves the UC, we must repaint the list (as an item is no longer hovered)
Private Sub ucSupport_MouseLeave(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long)
    m_MouseOverLayerBox = False
    UpdateHoveredLayer -1
    m_MouseX = -1
    m_MouseY = -1
    ucSupport.RequestCursor IDC_DEFAULT
    RedrawBackBuffer
End Sub

'When the mouse enters the button, we must initiate a repaint (to reflect its hovered state)
Private Sub ucSupport_MouseMoveCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal timeStamp As Long)
    
    'Ignore user interaction while in drag/drop mode
    If m_InOLEDragDropMode Then Exit Sub
    
    'Only display the hand cursor if the cursor is over a layer
    If (GetLayerAtPosition(x, y) <> -1) Then
        ucSupport.RequestCursor IDC_HAND
    Else
        ucSupport.RequestCursor IDC_ARROW
    End If
    
    'Don't process further MouseMove events if no images are loaded
    If (Not PDImages.IsImageActive()) Then Exit Sub
    
    'Store the mouse coords at module-level; the renderer may use these to highlight clickable elements
    m_MouseX = x
    m_MouseY = y
    
    'Process any important interactions first.  If a live interaction is taking place (such as drag/drop layer reordering),
    ' other MouseMove events will be suspended until the drag/drop is completed.
    
    'Check for drag/drop reordering
    If m_LayerRearrangingMode Then
    
        'The user is in the middle of a drag/drop reorder.  Give them a live update!
        
        'Retrieve the layer under this position
        Dim layerIndexUnderMouse As Long
        layerIndexUnderMouse = GetLayerAtPosition(x, y, True)
                
        'Ask the parent pdImage to move the layer for us
        If PDImages.GetActiveImage.MoveLayerToArbitraryIndex(m_LayerIndexToRearrange, layerIndexUnderMouse) Then
        
            'Note that the layer currently being moved has changed
            m_LayerIndexToRearrange = layerIndexUnderMouse
            
            'Keep the current layer as the active one
            SetActiveLayerByIndex layerIndexUnderMouse, False
            
            'Redraw the layer box, and note that thumbnails need to be re-cached
            Me.RequestRedraw True
            
            'Redraw the viewport
            Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
        
        End If
        
    End If
    
    'If a layer other than the active one is being hovered, highlight that box
    If (Not UpdateHoveredLayer(GetLayerAtPosition(x, y))) Then RedrawBackBuffer
    
End Sub

Private Sub ucSupport_MouseUpCustom(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal clickEventAlsoFiring As Boolean, ByVal timeStamp As Long)
    
    'Ignore user interaction while in drag/drop mode
    If m_InOLEDragDropMode Then Exit Sub
    
    'Retrieve the layer under this position
    Dim layerIndexUnderMouse As Long
    layerIndexUnderMouse = GetLayerAtPosition(x, y, True)
    
    'Don't proceed further unless an image has been loaded, and the user is not just clicking the layer box
    If PDImages.IsImageActive() And (Not clickEventAlsoFiring) Then
        
        'If we're in drag/drop mode, and the left mouse button is pressed, terminate drag/drop layer reordering
        If m_LayerRearrangingMode And (Button = pdLeftButton) Then
        
            'Exit layer rearranging mode
            m_LayerRearrangingMode = False
            
            'Ask the parent pdImage to move the layer for us; the MouseMove event has probably taken care of this already.
            ' In that case, this function will return FALSE and we don't have to do anything extra.
            If PDImages.GetActiveImage.MoveLayerToArbitraryIndex(m_LayerIndexToRearrange, layerIndexUnderMouse) Then
    
                'Keep the current layer as the active one
                SetActiveLayerByIndex layerIndexUnderMouse, False
                
                'Redraw the layer box, and note that thumbnails need to be re-cached
                Me.RequestRedraw True
                
                'Redraw the viewport
                Viewport.Stage2_CompositeAllLayers PDImages.GetActiveImage(), FormMain.MainCanvas(0)
                
            End If
            
            'If the new position differs from the layer's original position, call a dummy Processor call, which will create
            ' an Undo/Redo entry at this point.
            If (m_InitialLayerIndex <> layerIndexUnderMouse) Then Process "Rearrange layers", False, vbNullString, UNDO_ImageHeader
        
        End If
        
    End If
    
    'If we haven't already, exit layer rearranging mode
    m_LayerRearrangingMode = False
    
    'TODO: optimize this call; we may not need it if we redrew the buffer previously in this function
    RedrawBackBuffer
    
End Sub

Private Sub ucSupport_MouseWheelVertical(ByVal Button As PDMouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Long, ByVal y As Long, ByVal scrollAmount As Double)
    
    If (m_ScrollMax > 0) And (scrollAmount <> 0) Then
    
        Dim newScrollValue As Long: newScrollValue = m_ScrollValue
        
        If (scrollAmount > 0) Then
            newScrollValue = newScrollValue - Interface.FixDPI(LAYER_BLOCK_HEIGHT) \ 2
        Else
            newScrollValue = newScrollValue + Interface.FixDPI(LAYER_BLOCK_HEIGHT) \ 2
        End If
        
        If (newScrollValue < 0) Then newScrollValue = 0
        If (newScrollValue > m_ScrollMax) Then newScrollValue = m_ScrollMax
        m_ScrollValue = newScrollValue
        
        'Because a new layer will now be under the mouse, calculate that now (*prior* to rendering the updated box)
        UpdateHoveredLayer GetLayerAtPosition(x, y)
        RedrawBackBuffer
        
        'Last of all, notify our parent so they can synchronize the scroll bar to our new internal scroll value
        RaiseEvent ScrollValueChanged(newScrollValue)
        
    End If
        
End Sub

Private Sub ucSupport_RepaintRequired(ByVal updateLayoutToo As Boolean)
    If updateLayoutToo Then
        If (Not UpdateControlLayout) Then RedrawBackBuffer
    Else
        RedrawBackBuffer
    End If
End Sub

Private Sub UserControl_Initialize()
    
    'Initialize a user control support class
    Set ucSupport = New pdUCSupport
    ucSupport.RegisterControl UserControl.hWnd, True
    ucSupport.RequestExtraFunctionality True, True
    ucSupport.SpecifyRequiredKeys VK_UP, VK_DOWN, VK_RIGHT, VK_LEFT, VK_DELETE, VK_INSERT, VK_SPACE, VK_TAB
    ucSupport.SubclassCustomMessage WM_PD_COLOR_MANAGEMENT_CHANGE, True
    ucSupport.SubclassCustomMessage WM_PD_FLASH_ACTIVE_LAYER, True
    
    'Prep the color manager and load default colors
    Set m_Colors = New pdThemeColors
    Dim colorCount As PDLAYERBOX_COLOR_LIST: colorCount = [_Count]
    m_Colors.InitializeColorList "PDLayerBoxInner", colorCount
    If (Not PDMain.IsProgramRunning()) Then UpdateColorList
    
    'Reset all internal storage objects (used to track layer thumbnails, among other things)
    m_NumOfThumbnails = 0
    ReDim m_LayerThumbnails(0 To m_NumOfThumbnails) As LayerThumbDisplay
    Set m_ThumbCollection = New pdSpriteSheet
    
    m_MouseOverLayerBox = False
    m_LayerRearrangingMode = False
    m_CurLayerHover = -1
    m_MouseX = -1
    m_MouseY = -1
    m_ScrollValue = 0
    m_ScrollMax = 0
    Set m_PopupMenu = New pdPopupMenu
    
End Sub

Private Sub UserControl_InitProperties()
    Enabled = True
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'Make sure the form is available (e.g. a modal form hasn't stolen focus)
    If (Not g_AllowDragAndDrop) Then Exit Sub
    
    'Use the external function (in the clipboard handler, as the code is roughly identical to clipboard pasting)
    ' to load the OLE source.  This allows us to support Unicode filenames.
    m_InOLEDragDropMode = True
    g_Clipboard.LoadImageFromDragDrop Data, Effect, True
    m_InOLEDragDropMode = False
    
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

    'PD supports a lot of potential drop sources.  These values are defined and addressed by the main
    ' clipboard handler, as Drag/Drop and clipboard actions share a ton of code.
    If g_Clipboard.IsObjectDragDroppable(Data) Then
        Effect = vbDropEffectCopy And Effect
    Else
        Effect = vbDropEffectNone
    End If
    
End Sub

'At run-time, painting is handled by PD's pdWindowPainter class.  In the IDE, however, we must rely on VB's internal paint event.
Private Sub UserControl_Paint()
    If (Not PDMain.IsProgramRunning()) Then ucSupport.RequestIDERepaint UserControl.hDC
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Enabled = .ReadProperty("Enabled", True)
    End With
End Sub

Private Sub UserControl_Resize()
    If (Not PDMain.IsProgramRunning()) Then ucSupport.RequestRepaint True
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Enabled", Me.Enabled, True
    End With
End Sub

'Call to retrieve current list item height, in pixels.  (Our parent control needs this for
' setting a good .SmallChange value for the accompanying scrollbar.)
Public Function GetListItemHeight() As Long
    GetListItemHeight = Interface.FixDPI(LAYER_BLOCK_HEIGHT)
End Function

'External functions can request a redraw of the layer box by calling this function.  (This is necessary
' whenever layers are added, deleted, re-ordered, etc.)  If the action requires us to rebuild our thumbnail
' cache (because we switched images, maybe) make sure to clarify that via the matching parameter.
Public Sub RequestRedraw(Optional ByVal refreshThumbnailCache As Boolean = True, Optional ByVal layerID As Long = -1)
    
    'Redraws happen in multiple steps.
    
    'First, before doing anything else, update layer thumbnails when requested.
    ' (This step affects all subsequent steps, because it will re-calculate things like the
    ' size of the underlying list, if the number of layers in the image has changed.)
    If refreshThumbnailCache Then CacheLayerThumbnails layerID
    
    'Next, we need to make sure the currently selected layer is visible on-screen.
    If (Not m_SuspendAutoShowActiveLayer) Then EnsureActiveLayerIsVisible
    
    'Finally, we need to redraw the back buffer to reflect any changes from previous steps.
    RedrawBackBuffer
    
End Sub

Private Sub EnsureActiveLayerIsVisible()

    'This step is only necessary if a scrollbar is required to navigate the layer list,
    ' which in turn only matters if this is a multi-layer image.
    If PDImages.IsImageActive() Then
    If (PDImages.GetActiveImage.GetNumOfLayers > 1) Then
    If Me.IsScrollbarRequiredForHeight(Me.GetHeight()) Then
        
        'We potentially need to adjust the current scroll bar position.
        
        'Start by figuring out where the currently active layer is positioned.
        Dim activeLayerIndex As Long
        activeLayerIndex = (PDImages.GetActiveImage.GetNumOfLayers - PDImages.GetActiveImage.GetActiveLayerIndex) - 1   '0-based
        
        Dim BLOCKHEIGHT As Long
        BLOCKHEIGHT = Interface.FixDPI(LAYER_BLOCK_HEIGHT)
        
        Dim yCurrent As Long
        yCurrent = activeLayerIndex * BLOCKHEIGHT - m_ScrollValue
        
        'If the current layer lies *above* 0, reposition it to 0.
        If (yCurrent < 0) Then
            m_ScrollValue = activeLayerIndex * BLOCKHEIGHT
            RaiseEvent ScrollValueChanged(m_ScrollValue)
            Exit Sub
        End If
        
        'If it lies below the bottom boundary of the list (in pixels), scroll down accordingly
        Dim bottomOfControl As Long
        bottomOfControl = (Me.GetHeight - 3) - BLOCKHEIGHT
        If (yCurrent > bottomOfControl) Then
            m_ScrollValue = activeLayerIndex * BLOCKHEIGHT - bottomOfControl
            RaiseEvent ScrollValueChanged(m_ScrollValue)
            Exit Sub
        End If
        
    End If
    End If
    End If

End Sub

'Re-cache all thumbnails for all layers in the current image.  This is required when the user switches to a new image,
' or when an image is first loaded.
Private Sub CacheLayerThumbnails(Optional ByVal layerID As Long = -1)

    'Do not attempt to cache thumbnails if there are no open images
    If PDImages.IsImageActive() Then
    
        'Make sure the active image has at least one layer.  (This should always be true, but better safe than sorry.)
        If (PDImages.GetActiveImage.GetNumOfLayers > 0) Then
            
            'We now have two options.
            ' - If a valid layerID (>= 0) was specified, we can update just that layer.
            ' - If a layerID was NOT specified, we need to update all layers.
            Dim layerUpdateSuccessful As Boolean: layerUpdateSuccessful = False
            Dim tmpDIB As pdDIB
            Dim i As Long
            
            'A layer ID was provided.  Search our thumbnail collection for this layer ID.
            ' If we have an entry for it, update the thumbnail to match.
            If (layerID >= 0) And (m_NumOfThumbnails > 0) Then
                For i = 0 To m_NumOfThumbnails - 1
                    If (m_LayerThumbnails(i).CanonicalLayerID = layerID) Then
                        PDImages.GetActiveImage.GetLayerByIndex(i).RequestThumbnail tmpDIB, m_ThumbHeight
                        m_LayerThumbnails(i).spriteIDStr = Trim$(Str$(layerID))
                        m_LayerThumbnails(i).spriteIDInt = m_ThumbCollection.AddImage(tmpDIB, m_LayerThumbnails(i).spriteIDStr)
                        layerUpdateSuccessful = True
                        Exit For
                    End If
                Next i
            End If
            
            'If we failed to find the requested layer in our collection (or if our collection
            ' is currently empty, or if *all* thumbnails need to be modified), then we need to
            ' tackle the uglier task of rebuilding our entire thumbnail collection from scratch.
            If (Not layerUpdateSuccessful) Then
                
                'Retrieve the number of layers in the current image and prepare the thumbnail cache
                m_NumOfThumbnails = PDImages.GetActiveImage.GetNumOfLayers
                If (UBound(m_LayerThumbnails) <> (m_NumOfThumbnails - 1)) Then ReDim m_LayerThumbnails(0 To m_NumOfThumbnails - 1) As LayerThumbDisplay
                m_ThumbCollection.ResetCache
                
                If (m_NumOfThumbnails > 0) Then
                
                    For i = 0 To m_NumOfThumbnails - 1
                        
                        'Note that alongside the thumbnail, we also note each layer's canonical ID; this lets us
                        ' reuse thumbnails if layer order changes.
                        If (Not PDImages.GetActiveImage.GetLayerByIndex(i) Is Nothing) Then
                            m_LayerThumbnails(i).CanonicalLayerID = PDImages.GetActiveImage.GetLayerByIndex(i).GetLayerID
                            m_LayerThumbnails(i).spriteIDStr = Str$(m_LayerThumbnails(i).CanonicalLayerID)
                            PDImages.GetActiveImage.GetLayerByIndex(i).RequestThumbnail tmpDIB, m_ThumbHeight
                            m_LayerThumbnails(i).spriteIDInt = m_ThumbCollection.AddImage(tmpDIB, m_LayerThumbnails(i).spriteIDStr)
                        End If
                        
                    Next i
                    
                    'Whenever we perform a full refresh of the entire layer thumbnail collection
                    ' (after e.g. loading a new image), immediately attempt to minimize spritesheet
                    ' memory usage.  This will suspend all DIBs in the current thumb collection, and in
                    ' images with very large layer counts (e.g. animated images), this can save a lot
                    ' of memory because thumbnails outside the current view won't be refreshed for awhile.
                    ' (And when they are, they'll be silently paged and left in their uncompressed state.)
                    m_ThumbCollection.MinimizeMemory
                    
                End If
                
            End If
            
        Else
            m_NumOfThumbnails = 0
            If (UBound(m_LayerThumbnails) <> 0) Then ReDim m_LayerThumbnails(0) As LayerThumbDisplay
            m_ThumbCollection.ResetCache
        End If
        
    Else
        m_NumOfThumbnails = 0
        If (UBound(m_LayerThumbnails) <> 0) Then ReDim m_LayerThumbnails(0) As LayerThumbDisplay
        m_ThumbCollection.ResetCache
    End If
    
    'See if the list's scrollability has changed
    UpdateLayerScrollbarVisibility
    
End Sub

'Update the currently hovered layer.  Note that this sets a module-level flag, rather than returning a specific value.
Private Function UpdateHoveredLayer(ByVal newLayerUnderMouse As Long) As Boolean
    
    UpdateHoveredLayer = False
    
    If (Not PDMain.IsProgramRunning()) Then Exit Function
    
    'If a layer other than the active one is being hovered, highlight that box
    If (m_CurLayerHover <> newLayerUnderMouse) Then
        m_CurLayerHover = newLayerUnderMouse
        RedrawBackBuffer
        UpdateHoveredLayer = True
    End If

End Function

'Unlike other control's UpdateControlLayout() function(s), this control can't easily predict whether or not
' it will need to redraw the control (because it depends on some unpredictable factors like "does our new size
' require a scroll bar").  If this control *does* redraw the underlying window, it will return TRUE.  Use this
' to know whether you need to manually call RedrawBackBuffer after updating the control's layout.
Private Function UpdateControlLayout() As Boolean
    
    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    'Determine the position of the list rect.  A slight border allows us to apply chunky borders on focus.
    With m_ListRect
        .Left = 1
        .Top = 1
        .Width = (bWidth - 2) - .Left
        .Height = (bHeight - 2) - .Top
    End With
    
    'While here, setup a specific thumbnail width/height, which is calculated relative to the available block
    ' width/height of each layer entry.
    m_ThumbHeight = Interface.FixDPI(LAYER_BLOCK_HEIGHT) - Interface.FixDPI(THUMBNAIL_PADDING) * 2
    m_ThumbWidth = m_ThumbHeight
    
    'See if a scroll bar needs to be displayed.  Note that this will return TRUE if a redraw was requested -
    ' in that case, we can skip requesting our own redraw.
    UpdateControlLayout = (Not UpdateLayerScrollbarVisibility)
    If UpdateControlLayout Then RedrawBackBuffer
            
End Function

'Use this function to completely redraw the back buffer from scratch.
' Note that this is computationally expensive compared to just flipping the existing buffer to the screen,
' so only redraw the backbuffer if the control state has somehow changed.
Private Sub RedrawBackBuffer(Optional ByVal raiseImmediateDrawEvent As Boolean = False)
    
    Dim enabledState As Boolean
    enabledState = Me.Enabled
    
    'Retrieve DPI-aware control dimensions from the support class
    Dim bWidth As Long, bHeight As Long
    bWidth = ucSupport.GetBackBufferWidth
    bHeight = ucSupport.GetBackBufferHeight
    
    'Request the back buffer DC, and ask the support module to erase any existing rendering for us.
    Dim bufferDC As Long
    bufferDC = ucSupport.GetBackBufferDC(True, m_Colors.RetrieveColor(PDLB_Background, enabledState))
    If (bufferDC = 0) Then Exit Sub
    
    'This bunch of checks are basically failsafes to ensure we have valid pdLayer objects to pull from
    If PDMain.IsProgramRunning() Then
        
        'Some special circumstances require us to perform rendering even when we'd normally
        ' skip a step.  (For example, if the user attempts to paint on an invisible layer,
        ' the canvas will post a message that tells us to flash the visibility icon as a
        ' warning message is displayed.)
        Dim needToRenderAnimation As Boolean
        needToRenderAnimation = (m_FlashCount > 0)
        
        'If the list either 1) has keyboard focus, or 2) is actively being hovered by the mouse, we render
        ' it differently, using PD's standard hover behavior (accent colors and chunky border)
        Dim listHasFocus As Boolean
        listHasFocus = ucSupport.DoIHaveFocus Or ucSupport.IsMouseInside
        
        'Wrap the rendering area in a pd2D surface; this greatly simplifies paint ops
        Dim cSurface As pd2DSurface, cBrush As pd2DBrush, cPen As pd2DPen
        Drawing2D.QuickCreateSurfaceFromDC cSurface, bufferDC
        
        'Determine an offset based on the current scroll bar value
        Dim scrollOffset As Long
        scrollOffset = m_ScrollValue
        
        Dim layerIndex As Long, offsetX As Long, offsetY As Long, paintColor As Long
        Dim layerHoverIndex As Long, layerSelectedIndex As Long, layerIsHovered As Boolean, layerIsSelected As Boolean
        layerHoverIndex = -1: layerSelectedIndex = -1
        
        'Determine if we're in "zero layer" mode.  "Zero layer" mode lets us skip a lot of rendering details.
        Dim zeroLayers As Boolean
        If PDImages.IsImageActive() Then
            zeroLayers = (PDImages.GetActiveImage.GetNumOfLayers <= 0)
        Else
            zeroLayers = True
        End If
        
        'If we are not in "zero layers" mode, proceed with drawing the various list items
        If (Not zeroLayers) Then
            
            'Ensure we have the same number of thumbnails as the active image.
            ' (This won't trigger unless image state has somehow changed without us being notified.)
            If (m_NumOfThumbnails <> PDImages.GetActiveImage.GetNumOfLayers) Then
                PDDebug.LogAction "A UI change was initiated before proper notification - double check recent changes!"
                CacheLayerThumbnails
            End If
            
            'Cache colors in advance, so we can simply reuse them in the inner loop
            Dim itemColorSelectedBorder As Long, itemColorSelectedFill As Long
            Dim itemColorSelectedBorderHover As Long, itemColorSelectedFillHover As Long
            Dim itemColorUnselectedBorder As Long, itemColorUnselectedFill As Long
            Dim itemColorUnselectedBorderHover As Long, itemColorUnselectedFillHover As Long
            Dim fontColorSelected As Long, fontColorSelectedHover As Long
            Dim fontColorUnselected As Long, fontColorUnselectedHover As Long
            
            itemColorUnselectedBorder = m_Colors.RetrieveColor(PDLB_UnselectedItemBorder, enabledState, False, False)
            itemColorUnselectedBorderHover = m_Colors.RetrieveColor(PDLB_UnselectedItemBorder, enabledState, False, True)
            itemColorUnselectedFill = m_Colors.RetrieveColor(PDLB_UnselectedItemFill, enabledState, False, False)
            itemColorUnselectedFillHover = m_Colors.RetrieveColor(PDLB_UnselectedItemFill, enabledState, False, True)
            itemColorSelectedBorder = m_Colors.RetrieveColor(PDLB_SelectedItemBorder, enabledState, False, False)
            itemColorSelectedBorderHover = m_Colors.RetrieveColor(PDLB_SelectedItemBorder, enabledState, False, True)
            itemColorSelectedFill = m_Colors.RetrieveColor(PDLB_SelectedItemFill, enabledState, False, False)
            itemColorSelectedFillHover = m_Colors.RetrieveColor(PDLB_SelectedItemFill, enabledState, False, True)
            
            fontColorSelected = m_Colors.RetrieveColor(PDLB_SelectedItemText, enabledState, False, False)
            fontColorSelectedHover = m_Colors.RetrieveColor(PDLB_SelectedItemText, enabledState, False, True)
            fontColorUnselected = m_Colors.RetrieveColor(PDLB_UnselectedItemText, enabledState, False, False)
            fontColorUnselectedHover = m_Colors.RetrieveColor(PDLB_UnselectedItemText, enabledState, False, True)
            
            'Retrieve a font object that we can use for rendering layer names
            Dim layerFont As pdFont
            Set layerFont = Fonts.GetMatchingUIFont(10!)
            layerFont.AttachToDC bufferDC
            layerFont.SetTextAlignment vbLeftJustify
            
            Dim blockHeightDPIAware As Long
            blockHeightDPIAware = Interface.FixDPI(LAYER_BLOCK_HEIGHT)
            
            'Loop through the current layer list, drawing layers as we go
            Dim i As Long, tmpLayerRef As pdLayer
            For i = 0 To PDImages.GetActiveImage.GetNumOfLayers() - 1
            
                'Because layers are displayed in reverse order (layer 0 is displayed at the bottom of the list, not the top),
                ' we need to convert our For loop index into a matching layer index
                layerIndex = (PDImages.GetActiveImage.GetNumOfLayers() - 1) - i
                offsetX = m_ListRect.Left
                offsetY = m_ListRect.Top + Interface.FixDPI(i * LAYER_BLOCK_HEIGHT) - scrollOffset
                
                'Start by figuring out if this layer is even visible in the current box; if it isn't,
                ' skip drawing entirely
                If (((offsetY + blockHeightDPIAware) >= 0) And (offsetY < m_ListRect.Top + m_ListRect.Height)) Or needToRenderAnimation Then
                    
                    'For performance reasons, retrieve a local reference to the corresponding pdLayer object.
                    ' We need to pull a *lot* of information from this object.
                    Set tmpLayerRef = PDImages.GetActiveImage.GetLayerByIndex(layerIndex)
                    
                    If (Not tmpLayerRef Is Nothing) Then
                        
                        layerIsHovered = (layerIndex = m_CurLayerHover)
                        If (layerIsHovered) Then layerHoverIndex = layerIndex
                        layerIsSelected = (tmpLayerRef.GetLayerID = PDImages.GetActiveImage.GetActiveLayerID)
                        If (layerIsSelected) Then layerSelectedIndex = layerIndex
                        
                        'If the user is currently dragging this layer into a new position, we want to indent
                        ' it horizontally by an arbitrary pixel amount; this makes it more obvious that its
                        ' current state is "abnormal".
                        If (m_LayerRearrangingMode And (m_LayerMovingID = tmpLayerRef.GetLayerID)) Then offsetX = offsetX + (HORIZONTAL_ITEM_PADDING * 6)
                        
                        'To simplify drawing, convert the current block area into a rect; we'll use this for subsequent
                        ' layout decisions.
                        Dim blockRect As RectF
                        With blockRect
                            .Left = offsetX
                            .Top = offsetY
                            .Width = m_ListRect.Width - offsetX
                            .Height = blockHeightDPIAware
                        End With
                        
                        If layerIsHovered Then
                            With m_LayerHoverRect
                                .Left = blockRect.Left
                                .Top = blockRect.Top
                                .Right = .Left + blockRect.Width
                                .Bottom = .Top + blockRect.Height
                            End With
                        End If
                        
                        'Fill this block with the appropriate color.  (The actively selected layer is highlighted.)
                        If layerIsSelected Then paintColor = itemColorSelectedFill Else paintColor = itemColorUnselectedFill
                        Drawing2D.QuickCreateSolidBrush cBrush, paintColor
                        PD2D.FillRectangleF_FromRectF cSurface, cBrush, blockRect
                        
                        'Next, start painting block elements in LTR order.
                        Dim objOffsetX As Long, objOffsetY As Long
                        
                        'First, the layer visibility toggle
                        If (Not img_EyeOpen Is Nothing) Then
                            
                            'Start by calculating the "clickable" rect where the visibility icon lives.
                            ' (It is the left-most item in the layer box.)  This rect is module-level,
                            ' and this UC also uses it for hit-detection, so it needs to be pixel-accurate.
                            ' Note that we calculate this rect under two conditions:
                            ' 1) the layer is actively hovered by the mouse (so we need to highlight it), or...
                            ' 2) the user has attempted to interact with an invisible layer (so we need to flash it)
                            If layerIsHovered Or (needToRenderAnimation And layerIsSelected) Then
                                
                                With m_VisibilityRect
                                    .Left = blockRect.Left
                                    .Right = .Left + HORIZONTAL_ITEM_PADDING * 2 + img_EyeOpen.GetDIBWidth
                                    .Top = blockRect.Top
                                    .Bottom = .Top + blockHeightDPIAware
                                End With
                                
                                'If the icon area needs to be "flickered", do so now
                                If (needToRenderAnimation And layerIsSelected) Then
                                    If ((m_FlashCount And 1&) = 1&) Then
                                        Drawing2D.QuickCreateSolidBrush cBrush, itemColorUnselectedFill
                                        PD2D.FillRectangleI_AbsoluteCoords cSurface, cBrush, m_VisibilityRect.Left, m_VisibilityRect.Top, m_VisibilityRect.Right, m_VisibilityRect.Bottom
                                    End If
                                End If
                                
                            End If
                            
                            'Paint the appropriate visibility icon, centered in the current area
                            objOffsetX = blockRect.Left + HORIZONTAL_ITEM_PADDING
                            objOffsetY = blockRect.Top + (blockHeightDPIAware - img_EyeOpen.GetDIBHeight) \ 2
                            
                            If tmpLayerRef.GetLayerVisibility Then
                                img_EyeOpen.AlphaBlendToDC bufferDC, 255, objOffsetX, objOffsetY
                                img_EyeOpen.FreeFromDC
                            Else
                                img_EyeClosed.AlphaBlendToDC bufferDC, 255, objOffsetX, objOffsetY
                                img_EyeClosed.FreeFromDC
                            End If
                            
                            'Move the running offsets right
                            offsetX = offsetX + HORIZONTAL_ITEM_PADDING * 2 + img_EyeOpen.GetDIBWidth
                            
                        End If
                        
                        'Next comes the layer thumbnail.  If the layer is not currently visible, render it at partial opacity.
                        If m_ThumbCollection.DoesImageExist(m_LayerThumbnails(layerIndex).spriteIDStr) Then
                            
                            objOffsetX = offsetX + HORIZONTAL_ITEM_PADDING
                            objOffsetY = offsetY + (blockHeightDPIAware - m_ThumbHeight) \ 2
                            
                            If tmpLayerRef.GetLayerVisibility Then
                                m_ThumbCollection.PaintCachedImage bufferDC, objOffsetX, objOffsetY, m_LayerThumbnails(layerIndex).spriteIDInt
                            Else
                                m_ThumbCollection.PaintCachedImage bufferDC, objOffsetX, objOffsetY, m_LayerThumbnails(layerIndex).spriteIDInt, 127
                            End If
                            
                        End If
                        
                        'Move the running offsets right
                        offsetX = offsetX + m_ThumbWidth + HORIZONTAL_ITEM_PADDING * 2
                        
                        'Next comes the layer name
                        Dim drawString As String
                        drawString = tmpLayerRef.GetLayerName
                        
                        'Retrieve a matching font object from the UI font cache, and prep it with the proper display settings
                        If layerIsSelected Then
                            If layerIsHovered Then paintColor = fontColorSelectedHover Else paintColor = fontColorSelected
                        Else
                            If layerIsHovered Then paintColor = fontColorUnselectedHover Else paintColor = fontColorUnselected
                        End If
                        
                        layerFont.SetFontColor paintColor
                        
                        'Calculate where the string will actually lie.  This is important, as the text region is clickable
                        ' (the user can double-click to edit the layer's name).
                        Dim xTextOffset As Long, yTextOffset As Long, xTextWidth As Long, yTextHeight As Long
                        xTextOffset = offsetX + HORIZONTAL_ITEM_PADDING
                        xTextWidth = m_ListRect.Width - (xTextOffset + HORIZONTAL_ITEM_PADDING)
                        If (LenB(drawString) <> 0) Then yTextHeight = layerFont.GetHeightOfString(drawString) Else yTextHeight = Fonts.GetDefaultStringHeight(layerFont.GetFontSize)
                        yTextOffset = offsetY + (Interface.FixDPI(LAYER_BLOCK_HEIGHT) - yTextHeight) \ 2
                        layerFont.FastRenderTextWithClipping xTextOffset, yTextOffset, xTextWidth, yTextHeight, drawString
                        
                        'Store the resulting text area in the text rect; if the user clicks this, they can modify the layer name
                        With m_NameRect
                            If layerIsHovered Then
                                .Left = offsetX
                                .Top = offsetY
                                .Right = m_ListRect.Left + m_ListRect.Width - 2
                                .Bottom = offsetY + blockHeightDPIAware
                            End If
                        End With
                        
                        'The edit rect is where the edit text box is actually displayed.  We want this to closely mimic the
                        ' actual position of the text, rather than encompassing the entire clickable area.
                        With m_NameEditRect
                            If layerIsHovered Then
                                .Left = xTextOffset - 2
                                .Top = yTextOffset - 2
                                .Right = xTextOffset + xTextWidth + 2
                                .Bottom = yTextOffset + yTextHeight + 2
                            End If
                        End With
                        
                    'Layer is non-empty
                    End If
                    
                    Set tmpLayerRef = Nothing
                
                'Layer is not visible
                End If
                
            Next i
            
            layerFont.ReleaseFromDC
            
            'After painting all layers, if a layer is currently hovered by the mouse, highlight any clickable regions
            If (layerHoverIndex >= 0) Or needToRenderAnimation Then
                
                'First, draw a thin border around the hovered layer
                If (layerHoverIndex = layerSelectedIndex) Or needToRenderAnimation Then paintColor = itemColorSelectedBorderHover Else paintColor = itemColorUnselectedBorderHover
                Drawing2D.QuickCreateSolidPen cPen, 1, paintColor
                If (layerHoverIndex >= 0) Then PD2D.DrawRectangleF_AbsoluteCoords cSurface, cPen, m_LayerHoverRect.Left, m_LayerHoverRect.Top, m_LayerHoverRect.Right, m_LayerHoverRect.Bottom
                
                'Next, if the mouse is specifically within the "toggle layer visibility" rect, paint *that* region
                ' with a chunky border.
                If PDMath.IsPointInRect(m_MouseX, m_MouseY, m_VisibilityRect) Or needToRenderAnimation Then
                    Drawing2D.QuickCreateSolidPen cPen, 3!, paintColor
                    PD2D.DrawRectangleF_AbsoluteCoords cSurface, cPen, m_VisibilityRect.Left, m_VisibilityRect.Top, m_VisibilityRect.Right, m_VisibilityRect.Bottom
                End If
                
            End If
        
        'End zero-layer mode check
        End If
        
        'Last of all, render the listbox border.  Note that we actually draw *two* borders.
        ' The actual border, which is slightly inset from the list box boundaries, then a
        ' second border - pure background color, erasing any item rendering that may have
        ' fallen outside the clipping area.
        Dim borderWidth As Single, borderColor As Long
        If (listHasFocus And Not zeroLayers) Then borderWidth = 3! Else borderWidth = 1!
        borderColor = m_Colors.RetrieveColor(PDLB_Border, enabledState, listHasFocus And (Not zeroLayers))
        
        Drawing2D.QuickCreateSolidPen cPen, borderWidth, borderColor
        PD2D.DrawRectangleF_FromRectF cSurface, cPen, m_ListRect
        
        If (Not listHasFocus) Then
            Drawing2D.QuickCreateSolidPen cPen, 1!, m_Colors.RetrieveColor(PDLB_Background, enabledState)
            PD2D.DrawRectangleI cSurface, cPen, 0, 0, bWidth - 1, bHeight - 1
        End If
        
        Set cSurface = Nothing: Set cBrush = Nothing: Set cPen = Nothing
        
    End If
    
    'Paint the final result to the screen, as relevant
    ucSupport.RequestRepaint raiseImmediateDrawEvent
    If (Not PDMain.IsProgramRunning()) Then UserControl.Refresh
    
End Sub

'Given mouse coordinates over the control, return the layer at that location.  The optional parameter
' "reportNearestLayer" will return the index of the top layer if the mouse is in the invalid area
' *above* the top-most layer, and the bottom layer if in the invalid area *beneath* the bottom-most layer.
Private Function GetLayerAtPosition(ByVal x As Long, ByVal y As Long, Optional ByVal reportNearestLayer As Boolean = False) As Long
    
    If (Not PDImages.IsImageActive()) Then
        GetLayerAtPosition = -1
    Else
        
        Dim tmpLayerCheck As Long
        tmpLayerCheck = (y + m_ScrollValue) \ Interface.FixDPI(LAYER_BLOCK_HEIGHT)
    
        'It's a bit counterintuitive, but we draw the layer box in reverse order: layer 0 (the image's first layer)
        ' is at the BOTTOM of our box, while layer(max) is at the TOP.  Because of this, all layer positioning checks
        ' must be reversed.
        tmpLayerCheck = (PDImages.GetActiveImage.GetNumOfLayers - 1) - tmpLayerCheck
        
        'Is the mouse over an actual layer, or just dead space in the box?
        If (tmpLayerCheck >= 0) And (tmpLayerCheck < PDImages.GetActiveImage.GetNumOfLayers) Then
            GetLayerAtPosition = tmpLayerCheck
        Else
        
            'If the user wants us to report the *nearest* valid layer
            If reportNearestLayer Then
                If (tmpLayerCheck < 0) Then
                    GetLayerAtPosition = 0
                Else
                    GetLayerAtPosition = PDImages.GetActiveImage.GetNumOfLayers - 1
                End If
            Else
                GetLayerAtPosition = -1
            End If
            
        End If
    
    End If
    
End Function

'When an action occurs that potentially affects the visibility of the vertical scroll bar (such as resizing the form
' vertically, or adding a new layer to the image), call this function.  Any changes will be bubbled upward to our parent.
'
'Returns: TRUE if the backbuffer was redrawn due to a visibility change; FALSE otherwise.  Use this to determine if you
' need to provide your own redraw.
Private Function UpdateLayerScrollbarVisibility() As Boolean
    
    UpdateLayerScrollbarVisibility = False
    
    Dim maxBoxSize As Long
    maxBoxSize = Interface.FixDPIFloat(LAYER_BLOCK_HEIGHT) * m_NumOfThumbnails - 1
    
    If (maxBoxSize < m_ListRect.Height) Then
        m_ScrollValue = 0
        If (m_ScrollMax <> 0) Then
            m_ScrollMax = 0
            RaiseEvent ScrollMaxChanged(0)
            RedrawBackBuffer
            UpdateLayerScrollbarVisibility = True
        End If
    Else
        If (m_ScrollMax <> (maxBoxSize - m_ListRect.Height)) Then
            m_ScrollMax = (maxBoxSize - m_ListRect.Height) + 1
            RaiseEvent ScrollMaxChanged(m_ScrollMax)
            RedrawBackBuffer
            UpdateLayerScrollbarVisibility = True
        End If
    End If
    
End Function

'Want to know if a scrollbar is required at some arbitrary height?  Use this function to test.
Public Function IsScrollbarRequiredForHeight(ByVal testHeight As Long) As Boolean
    
    Dim maxBoxSize As Long
    maxBoxSize = Interface.FixDPIFloat(LAYER_BLOCK_HEIGHT) * m_NumOfThumbnails - 1
    
    'Note that this control requires three vertical pixels for padding; that's the reason for the
    ' magic "3" number, below.
    IsScrollbarRequiredForHeight = (maxBoxSize >= (testHeight - 3))
    
End Function

'Before this control does any painting, we need to retrieve relevant colors from PD's primary theming class.  Note that this
' step must also be called if/when PD's visual theme settings change.
Private Sub UpdateColorList()
    With m_Colors
        .LoadThemeColor PDLB_Background, "Background", IDE_WHITE
        .LoadThemeColor PDLB_Border, "Border", IDE_GRAY
        .LoadThemeColor PDLB_SelectedItemFill, "SelectedItemFill", IDE_BLUE
        .LoadThemeColor PDLB_SelectedItemBorder, "SelectedItemBorder", IDE_BLUE
        .LoadThemeColor PDLB_SelectedItemText, "SelectedItemText", IDE_WHITE
        .LoadThemeColor PDLB_UnselectedItemFill, "UnselectedItemFill", IDE_WHITE
        .LoadThemeColor PDLB_UnselectedItemBorder, "UnselectedItemBorder", IDE_WHITE
        .LoadThemeColor PDLB_UnselectedItemText, "UnselectedItemText", IDE_BLACK
    End With
End Sub

'External functions can call this to request a redraw.  This is helpful for live-updating theme settings, as in the Preferences dialog.
Public Sub UpdateAgainstCurrentTheme(Optional ByVal hostFormhWnd As Long = 0)
    
    'Load all hover UI image resources
    If ucSupport.ThemeUpdateRequired Then
        
        If PDMain.IsProgramRunning() Then
            Dim iconSize As Long
            iconSize = Interface.FixDPI(16)
            LoadResourceToDIB "generic_visible", img_EyeOpen, iconSize, iconSize, usePDResamplerInstead:=rf_Box
            LoadResourceToDIB "generic_invisible", img_EyeClosed, iconSize, iconSize, usePDResamplerInstead:=rf_Box
        End If
        
        UpdateColorList
        If PDMain.IsProgramRunning() Then NavKey.NotifyControlLoad Me, hostFormhWnd
        If PDMain.IsProgramRunning() Then ucSupport.UpdateAgainstThemeAndLanguage
        txtLayerName.UpdateAgainstCurrentTheme
        
    End If
    
End Sub

'By design, PD prefers to not use design-time tooltips.  Apply tooltips at run-time, using this function.
' (IMPORTANT NOTE: translations are handled automatically.  Always pass the original English text!)
Public Sub AssignTooltip(ByRef newTooltip As String, Optional ByRef newTooltipTitle As String = vbNullString, Optional ByVal raiseTipsImmediately As Boolean = False)
    ucSupport.AssignTooltip UserControl.ContainerHwnd, newTooltip, newTooltipTitle, raiseTipsImmediately
End Sub
