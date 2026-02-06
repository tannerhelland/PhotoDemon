Attribute VB_Name = "Interface"
'***************************************************************************
'Miscellaneous functions related to UI
'Copyright 2001-2026 by Tanner Helland
'Created: 6/12/01
'Last updated: 17/August/17
'Last update: overhaul PDMsgBox to use an internal renderer
'
'Miscellaneous routines related to rendering and handling PhotoDemon's interface.  As the program's complexity has
' increased, so has the need for specialized handling of certain UI elements.
'
'Many of the functions in this module rely on subclassing, either directly or through things like PD's window manager.
' As such, some functions may operate differently (or not at all) while in the IDE.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As PointAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As winRect) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

'System keyboard repeat settings are mimicked by internal PD controls
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoW" (ByVal uiAction As Long, ByVal uiParam As Long, ByRef pvParam As Long, ByVal fWinIni As Long) As Long
Private Const SPI_GETKEYBOARDDELAY As Long = &H16
Private Const SPI_GETKEYBOARDSPEED As Long = &HA

'Types and API calls for processing ESC keypresses mid-loop
Private Type winMsg
    hWnd As Long
    sysMsg As Long
    wParam As Long
    lParam As Long
    msgTime As Long
    ptX As Long
    ptY As Long
End Type

Private Declare Function GetInputState Lib "user32" () As Long
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (ByRef lpMsg As winMsg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Const WM_KEYFIRST As Long = &H100
Private Const WM_KEYLAST As Long = &H108
Private Const PM_REMOVE As Long = &H1

'We use a lot of PeekMessage() calls to look for user canceling of long-running actions.  To keep the function small,
' we simply reuse a module-level variable on each call.
Private m_tmpMsg As winMsg

Public g_cancelCurrentAction As Boolean

'Some UI elements are always enabled or disabled as a group.  For example, the PDUI_Save group will simultaneously en/disable the
' File -> Save menu, toolbar Save button(s), Ctrl+S hotkey, etc.  This enum is used throughout the interface manager to optimize the
' way PD en/disables large groups of related controls
Public Enum PD_UI_Group
    PDUI_Save = 0
    PDUI_SaveAs = 1
    PDUI_Close = 2
    PDUI_Undo = 3
    PDUI_Redo = 4
    PDUI_EditCopyCut = 5
    PDUI_Paste = 6
    PDUI_View = 7
    PDUI_ImageMenu = 8
    PDUI_Metadata = 9
    PDUI_GPSMetadata = 10
    PDUI_Macros = 11
    PDUI_Selections = 12
    PDUI_SelectionTransforms = 13
    PDUI_LayerTools = 14
    PDUI_ICCProfile = 15
    PDUI_FileOnDisk = 16
End Enum

#If False Then
    Private Const PDUI_Save = 0, PDUI_SaveAs = 1, PDUI_Close = 2, PDUI_Undo = 3, PDUI_Redo = 4, PDUI_Copy = 5, PDUI_Paste = 6, PDUI_View = 7
    Private Const PDUI_ImageMenu = 8, PDUI_Metadata = 9, PDUI_GPSMetadata = 10, PDUI_Macros = 11, PDUI_Selections = 12
    Private Const PDUI_SelectionTransforms = 13, PDUI_LayerTools = 14, PDUI_ICCProfile = 15, PDUI_FileOnDisk = 16
#End If

'PhotoDemon is designed against pixels at an expected screen resolution of 96 DPI.
' Other DPI settings mess up spacing calculations. To remedy this, we dynamically modify
' pixel measurements at run-time, using the current screen resolution as our guide.
Private m_DPIRatio As Double

'System DPI is used frequently for UI positioning calculations.  Because it's costly to constantly retrieve it via APIs,
' this module prefers to cache it only when the value changes.  Call the CacheSystemDPI() sub to update the value when
' appropriate, and the corresponding GetSystemDPI() function to retrieve the cached value.
Private m_CurrentSystemDPI As Single

'When a modal dialog is displayed, a reference to it is saved in this variable.
' If subsequent modal dialogs are displayed (for example, if a tool dialog displays a
' color selection dialog), the previous modal dialog is given ownership over the new dialog.
Private currentDialogReference As Form
Private isSecondaryDialog As Boolean

'When the central "ShowPDDialog" function is called, the dialog it raises must possess one of
' PD's command bar instances. The command bar will set a global "OK/Cancel" value that subsequent
' functions can retrieve, if they're curious.  (For example, a "cancel" result usually means that
' you can skip subsequent UI syncs, as the image's status has not changed.)
Private m_LastShowDialogResult As VbMsgBoxResult

'When a message is displayed to the user in the message portion of the status bar,
' we automatically cache the message's contents. If a subsequent request is raised
' with the exact same text, we can skip the whole message display process.
' (Note that this is the *unparsed, English-language* version of the message!
'  It is a much faster candidate for pattern-matching.)
Private m_PrevMessage As String

'Same as m_PrevMessage, but with all translations and/or custom parsing applied
Private m_LastFullMessage As String

'Syncing the entire program's UI to current image settings is a time-consuming process.  To try and shortcut it whenever possible,
' we track the last sync operation we performed.  If we receive a duplicate sync request, we can safely ignore it.
Private m_LastUISync_HadNoImages As PD_BOOL, m_LastUISync_HadNoLayers As PD_BOOL, m_LastUISync_HadMultipleLayers As PD_BOOL

'Popup dialogs present problems on non-Aero window managers, as VB's iconless approach results in the program "disappearing"
' from places like the Alt+Tab menu.  As of v7.0, we now track nested popup windows and manually handle their icon updates.
Private Const NESTED_POPUP_LIMIT As Long = 16&
Private m_PopupHWnds() As Long, m_NumOfPopupHWnds As Long
Private m_PopupIconsSmall() As Long, m_PopupIconsLarge() As Long

'A unique string that tracks the current theme and language combination.  If either one changes, we need to redraw
' large swaths of the interface to match.
Private m_CurrentInterfaceID As String

'Various program functions related to the main window disable themselves while modal dialogs (including common dialogs)
' are active.  We track dialog state internally, although things like modal dialogs require external notifications.
Private m_ModalDialogActive As Boolean, m_SystemDialogActive As Boolean

'When a dialog enters a modal resize loop, it should call us to set this flag; child controls may
' query the flag to know whether to perform intensive processing (or wait until the resize ends)
Private m_DialogActivelyResizing As Boolean

'Because the Interface handler is a module and not a class, like I prefer, we need to use a dedicated initialization function.
Public Sub InitializeInterfaceBackend()

    m_LastUISync_HadNoImages = PD_BOOL_UNKNOWN
    m_LastUISync_HadNoLayers = PD_BOOL_UNKNOWN
    m_LastUISync_HadMultipleLayers = PD_BOOL_UNKNOWN
    
    'vbIgnore is used internally as the "no result" value for a dialog box, as PD never provides an actual "ignore" option
    ' in its dialogs.
    m_LastShowDialogResult = vbIgnore
    
End Sub

'Get/set system DPI *as a ratio*, e.g. 96 DPI ("100%") should be cached as 1.0.  This gives us an easy modifier for calculating
' new window layouts and sizes.
Public Sub CacheSystemDPIRatio(ByVal newDPI As Single)
    m_CurrentSystemDPI = newDPI
    If (m_CurrentSystemDPI = 0!) Then m_CurrentSystemDPI = 1!
End Sub

Public Function GetSystemDPIRatio() As Single
    If (m_CurrentSystemDPI = 0!) Then m_CurrentSystemDPI = 1!
    GetSystemDPIRatio = m_CurrentSystemDPI
End Function

'Returns the last session's DPI, as a *ratio* (e.g. 96 DPI is returned as 1.0, 150% DPI is returned as 1.5)
Public Function GetLastSessionDPI_Ratio() As Single
    GetLastSessionDPI_Ratio = UserPrefs.GetPref_Float("Toolbox", "LastSessionDPI", Interface.GetSystemDPIRatio())
    If (GetLastSessionDPI_Ratio < 1!) Then GetLastSessionDPI_Ratio = 1!
    If (GetLastSessionDPI_Ratio > 4!) Then GetLastSessionDPI_Ratio = 4!
End Function

'PD's canvas uses interactive elements (like clickable "nodes") in many different contexts.  To ensure uniform
' UX behavior across tools, a standardized "close enough to interact" distance is used.  This value varies by
' run-time display DPI.
Public Function GetStandardInteractionDistance(Optional ByRef srcImageForCanvasData As pdImage = Nothing) As Single
    Const INTERACTION_DISTANCE_96_PPI As Single = 7!
    GetStandardInteractionDistance = Interface.FixDPIFloat(INTERACTION_DISTANCE_96_PPI)
End Function

'Generate a unique interface ID string that describes the current visual theme + language combination.
' UI elements can query this via GetCurrentInterfaceID() before rendering to see if they need to
' revisit layout and/or color decisions.
Public Sub GenerateInterfaceID()
    
    m_CurrentInterfaceID = vbNullString
    
    Dim i18nActive As Boolean
    i18nActive = (Not g_Language Is Nothing)
    If i18nActive Then i18nActive = g_Language.TranslationActive()
    
    If (Not g_Themer Is Nothing) Then
        m_CurrentInterfaceID = g_Themer.GetCurrentThemeID
        If i18nActive Then m_CurrentInterfaceID = m_CurrentInterfaceID & "|" & CStr(g_Language.GetCurrentLanguageIndex)
    Else
        If i18nActive Then m_CurrentInterfaceID = CStr(g_Language.GetCurrentLanguageIndex)
    End If
    
End Sub

Public Function GetCurrentInterfaceID() As String
    GetCurrentInterfaceID = m_CurrentInterfaceID
End Function

Public Function GetDialogResizeFlag() As Boolean
    GetDialogResizeFlag = m_DialogActivelyResizing
End Function

Public Sub SetDialogResizeFlag(ByVal newFlag As Boolean)
    m_DialogActivelyResizing = newFlag
End Sub

Public Sub SetFormCaptionW(ByRef dstForm As Form, ByVal srcCaption As String)
    If (LenB(srcCaption) > 0) Then srcCaption = " " & srcCaption
    If (Not g_WindowManager Is Nothing) Then
        g_WindowManager.SetWindowCaptionW dstForm.hWnd, srcCaption
    Else
        dstForm.Caption = srcCaption
    End If
End Sub

'Previously, various PD functions had to manually enable/disable button and menu state based on their actions.  This is no longer necessary.
' Simply call this function whenever an action has done something that will potentially affect the interface, and this function will iterate
' through all potential image/interface interactions, dis/enabling buttons and menus as necessary.
'
'TODO: look at having an optional "layerID" parameter, so we can skip certain steps if only a single layer is affected by a change.
Public Sub SyncInterfaceToCurrentImage()
        
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    'Interface dis/enabling falls into two rough categories: stuff that changes based on the current image (e.g. Undo), and stuff that changes
    ' based on the *total* number of available images (e.g. visibility of the Effects menu).
    
    'Start by breaking our interface decisions into two broad categories: "no images are loaded" and "one or more images are loaded".
    
    'If no images are loaded, we can disable a whole swath of controls
    If (Not PDImages.IsImageActive()) Then
    
        'Because this set of UI changes is immutable, there is no reason to repeat it if it was the last synchronization we performed.
        If (m_LastUISync_HadNoImages <> PD_BOOL_TRUE) Then
            SetUIMode_NoImages
            m_LastUISync_HadNoImages = PD_BOOL_TRUE
        End If
        
    'If one or more images are loaded, our job is trickier.  Some controls (such as Copy to Clipboard) are enabled no matter what,
    ' while others (Undo and Redo) are only enabled if the current image requires it.
    Else
        
        'Start with controls that are *always* enabled if at least one image is active.  These controls only need to be addressed when
        ' we move between the "no images" and "at least one image" state.
        If Not (m_LastUISync_HadNoImages = PD_BOOL_FALSE) Then
            SetUIMode_AtLeastOneImage
            m_LastUISync_HadNoImages = PD_BOOL_FALSE
        End If
    
        'Next, we have controls whose appearance varies according to the current image's state.  These include things like the
        ' window caption (which changes if the filename changes), Undo/Redo state (changes based on tool actions), image size
        ' and zoom indicators, etc.  These settings are more difficult to cache, because they can legitimately change for the
        ' same image object, so detecting meaningful vs repeat changes is trickier.
        SyncUI_CurrentImageSettings
        
        'Next, we are going to deal with layer-specific settings.
        
        'Start with settings that are ALWAYS visible if there is at least one layer in the image.
        ' (NOTE: PD doesn't currently support 0-layer images, so this is primarily a failsafe measure.)
        
        'Update all layer menus; some will be disabled depending on just how many layers are available, how many layers
        ' are visible, and other criteria.
        If (PDImages.GetActiveImage.GetNumOfLayers > 0) And (Not PDImages.GetActiveImage.GetActiveLayer Is Nothing) Then
            
            'Activate any generic layer UI elements (e.g. elements whose enablement is consistent for any number of layers)
            If Not (m_LastUISync_HadNoLayers = PD_BOOL_FALSE) Then
                SetUIMode_AtLeastOneLayer
                m_LastUISync_HadNoLayers = PD_BOOL_FALSE
            End If
            
            'Next, activate UI parameters whose behavior changes depending on the current layer's settings
            SyncUI_CurrentLayerSettings
            
            'Next, we must deal with controls whose enablement depends on how many layers are in the image.  Some options
            ' (like "Flatten" or "Delete layer") are only relevant if this is a multi-layer image.
            
            'If only one layer is present, a number of layer menu items (Delete, Flatten, Merge, Order) will be disabled.
            If (PDImages.GetActiveImage.GetNumOfLayers = 1) Then
            
                If Not (m_LastUISync_HadMultipleLayers = PD_BOOL_FALSE) Then
                    SetUIMode_OnlyOneLayer
                    m_LastUISync_HadMultipleLayers = PD_BOOL_FALSE
                End If
                
            'This image contains multiple layers.  Enable additional menu items (if they aren't already).
            Else
                
                If Not (m_LastUISync_HadMultipleLayers = PD_BOOL_TRUE) Then
                    SetUIMode_MultipleLayers
                    m_LastUISync_HadMultipleLayers = PD_BOOL_TRUE
                End If
                
                'Next, activate UI parameters whose behavior changes depending on the settings of multiple layers in the image
                ' (e.g. "delete hidden layers" requires at least one hidden layer in the image)
                SyncUI_MultipleLayerSettings
                
            End If
            
        'This Else branch should never be triggered, because PD doesn't allow zero-layer images, by design.
        Else
            If Not (m_LastUISync_HadNoLayers = PD_BOOL_TRUE) Then
                SetUIMode_NoLayers
                m_LastUISync_HadNoLayers = PD_BOOL_TRUE
            End If
            m_LastUISync_HadMultipleLayers = PD_BOOL_FALSE
        End If
            
        
        'TODO: move selection settings into the tool handler; they're too low-level for this function
        'If a selection is active on this image, update the text boxes to match
        If PDImages.GetActiveImage.IsSelectionActive And (Not PDImages.GetActiveImage.MainSelection Is Nothing) Then
            SetUIGroupState PDUI_Selections, True
            SetUIGroupState PDUI_SelectionTransforms, PDImages.GetActiveImage.MainSelection.IsTransformable()
            SyncTextToCurrentSelection PDImages.GetActiveImageID()
        Else
            SetUIGroupState PDUI_Selections, False
            SetUIGroupState PDUI_SelectionTransforms, False
        End If
        
        'Finally, synchronize various tool settings.  I've optimized this so that only the settings relative to the current tool
        ' are updated; others will be modified if/when the active tool is changed.
        Tools.SyncToolOptionsUIToCurrentLayer
        
    End If
        
    'Perform a special check if 2 or more images are loaded; if that is the case, enable a few additional controls, like
    ' the "Next/Previous" Window menu items.
    Menus.SetMenuEnabled "window_next", (PDImages.GetNumOpenImages() > 1)
    Menus.SetMenuEnabled "window_previous", (PDImages.GetNumOpenImages() > 1)
    
    'Similarly, the "assemble images into this image as layers" menu requires multiple images to be loaded
    Menus.SetMenuEnabled "layer_splitimagestolayers", (PDImages.GetNumOpenImages() > 1)
    
    'Redraw the layer box
    toolbar_Layers.NotifyLayerChange
    
    PDDebug.LogAction "Interface.SyncInterfaceToCurrentImage finished in " & VBHacks.GetTimeDiffNowAsString(startTime)
    
End Sub

'Synchronize all settings whose behavior and/or appearance depends on:
' 1) At least 2+ valid layers in the current image
' 2) Different behavior and/or appearances for different layer settings
' If a UI element appears the same for ANY amount of multiple layers (e.g. "Delete Layer"), use the SetUIMode_MultipleLayers() function.
Private Sub SyncUI_MultipleLayerSettings()
    
    'Delete hidden layers is only available if one or more layers are hidden, but not ALL layers are hidden.
    Menus.SetMenuEnabled "layer_deletehidden", (PDImages.GetActiveImage.GetNumOfHiddenLayers > 0) And (PDImages.GetActiveImage.GetNumOfHiddenLayers < PDImages.GetActiveImage.GetNumOfLayers)
    
    'Merge up/down are not available for layers at the top and bottom of the image
    Menus.SetMenuEnabled "layer_mergeup", (Layers.IsLayerAllowedToMergeAdjacent(PDImages.GetActiveImage.GetActiveLayerIndex, False) <> -1)
    Menus.SetMenuEnabled "layer_mergedown", (Layers.IsLayerAllowedToMergeAdjacent(PDImages.GetActiveImage.GetActiveLayerIndex, True) <> -1)
    
    'Within the order menu, certain items are disabled based on layer position.  Note that "move up" and
    ' "move to top" are both disabled for top layers (similarly for bottom layers and "move down/bottom"),
    ' so we can mirror the same enabled state for both options.
    
    'Activate top/next layer up
    Dim mnuEnabled As Boolean
    mnuEnabled = (PDImages.GetActiveImage.GetActiveLayerIndex < PDImages.GetActiveImage.GetNumOfLayers - 1)
    Menus.SetMenuEnabled "layer_gotop", mnuEnabled
    Menus.SetMenuEnabled "layer_goup", mnuEnabled
    
    'Activate bottom/next layer down
    mnuEnabled = (PDImages.GetActiveImage.GetActiveLayerIndex > 0)
    Menus.SetMenuEnabled "layer_godown", mnuEnabled
    Menus.SetMenuEnabled "layer_gobottom", mnuEnabled
    
    'Move to top/move up
    mnuEnabled = (PDImages.GetActiveImage.GetActiveLayerIndex < PDImages.GetActiveImage.GetNumOfLayers - 1)
    Menus.SetMenuEnabled "layer_movetop", mnuEnabled
    Menus.SetMenuEnabled "layer_moveup", mnuEnabled
    
    'Move to bottom/move down
    mnuEnabled = (PDImages.GetActiveImage.GetActiveLayerIndex > 0)
    Menus.SetMenuEnabled "layer_movedown", mnuEnabled
    Menus.SetMenuEnabled "layer_movebottom", mnuEnabled
    
    'Reverse layer order is always available for multi-layer images
    Menus.SetMenuEnabled "layer_reverse", True
    
    'Merge visible is only available if *two* or more layers are visible
    Menus.SetMenuEnabled "image_mergevisible", (PDImages.GetActiveImage.GetNumOfVisibleLayers > 1)
    
    'Flatten is only available if one or more layers are actually *visible*
    Menus.SetMenuEnabled "image_flatten", (PDImages.GetActiveImage.GetNumOfVisibleLayers > 0)
    
End Sub

'Synchronize all settings whose behavior and/or appearance depends on:
' 1) At least one valid layer in the current image
' 2) Different behavior and/or appearances for different layers
' If a UI element appears the same for ANY layer (e.g. toggling visibility), use the SetUIMode_AtLeastOneLayer() function.
Public Sub SyncUI_CurrentLayerSettings()
    
    'First, determine if the current layer is using any form of non-destructive resizing
    Dim nonDestructiveResizeActive As Boolean
    nonDestructiveResizeActive = (PDImages.GetActiveImage.GetActiveLayer.GetLayerCanvasXModifier <> 1#) Or (PDImages.GetActiveImage.GetActiveLayer.GetLayerCanvasYModifier <> 1#)
    
    'If non-destructive resizing is active, the "reset layer size" menu (and corresponding Move Tool button) must be enabled.
    If (Menus.IsMenuEnabled("layer_resetsize") <> nonDestructiveResizeActive) Then Menus.SetMenuEnabled "layer_resetsize", nonDestructiveResizeActive
    
    If (g_CurrentTool = NAV_MOVE) Then
        toolpanel_MoveSize.cmdLayerAffinePermanent.Enabled = PDImages.GetActiveImage.GetActiveLayer.AffineTransformsActive(True)
    End If
    
    'Layer visibility
    If (Menus.IsMenuChecked("layer_show") <> PDImages.GetActiveImage.GetActiveLayer.GetLayerVisibility()) Then Menus.SetMenuChecked "layer_show", PDImages.GetActiveImage.GetActiveLayer.GetLayerVisibility()
    
    'Layer rasterization depends on the current layer type
    If (Menus.IsMenuEnabled("layer_rasterizecurrent") <> PDImages.GetActiveImage.GetActiveLayer.IsLayerVector) Then Menus.SetMenuEnabled "layer_rasterizecurrent", PDImages.GetActiveImage.GetActiveLayer.IsLayerVector
    If (Menus.IsMenuEnabled("layer_rasterizeall") <> (PDImages.GetActiveImage.GetNumOfVectorLayers > 0)) Then Menus.SetMenuEnabled "layer_rasterizeall", (PDImages.GetActiveImage.GetNumOfVectorLayers > 0)
    
End Sub

'Synchronize all settings whose behavior and/or appearance depends on:
' 1) At least one valid, loaded image
' 2) Different behavior and/or appearances for different images
' If a UI element appears the same for ANY loaded image (e.g. activating the main canvas), use the SetUIMode_AtLeastOneImage() function.
Private Sub SyncUI_CurrentImageSettings()
            
    'Reset all Undo/Redo and related menus.  (Note that this also controls the SAVE BUTTON, as the image's save state is modified
    ' by PD's Undo/Redo engine.)
    Interface.SyncUndoRedoInterfaceElements True
    
    'Because Undo/Redo changes may modify menu captions, menu icons need to be reset (as they are tied to menu captions)
    IconsAndCursors.ResetMenuIcons
    
    'Determine whether metadata is present, and dis/enable metadata menu items accordingly
    If (Not PDImages.GetActiveImage.ImgMetadata Is Nothing) Then
        SetUIGroupState PDUI_Metadata, PDImages.GetActiveImage.ImgMetadata.HasMetadata
        SetUIGroupState PDUI_GPSMetadata, PDImages.GetActiveImage.ImgMetadata.HasGPSMetadata()
    Else
        SetUIGroupState PDUI_Metadata, False
        SetUIGroupState PDUI_GPSMetadata, False
    End If
    
    'If the image has an embedded ICC profile, expose the `File > Export > ICC profile` menu
    SetUIGroupState PDUI_ICCProfile, (LenB(PDImages.GetActiveImage.GetColorProfile_Original()) <> 0)
    
    'If the image exists on-disk, expose the `Image > Show location in Explorer` menu
    SetUIGroupState PDUI_FileOnDisk, (LenB(PDImages.GetActiveImage.ImgStorage.GetEntry_String("CurrentLocationOnDisk", vbNullString)) <> 0)
    
    'Display the image's path in the title bar.
    If (Not g_WindowManager Is Nothing) Then
        g_WindowManager.SetWindowCaptionW FormMain.hWnd, Interface.GetWindowCaption(PDImages.GetActiveImage())
    Else
        FormMain.Caption = Interface.GetWindowCaption(PDImages.GetActiveImage())
    End If
    
    'Display the image's size in the status bar
    If (PDImages.GetActiveImage.Width <> 0) Then Interface.DisplaySize PDImages.GetActiveImage()
    
    'Update the form's icon to match the current image; if a custom icon is not available, use the stock PD one
    If (PDImages.GetActiveImage.GetImageIcon(False) = 0) Or (PDImages.GetActiveImage.GetImageIcon(True) = 0) Then IconsAndCursors.CreateCustomFormIcons PDImages.GetActiveImage()
    IconsAndCursors.ChangeAppIcons PDImages.GetActiveImage.GetImageIcon(False), PDImages.GetActiveImage.GetImageIcon(True)
    
    'Restore the zoom value for this particular image (again, only if the form has been initialized)
    If (PDImages.GetActiveImage.Width <> 0) Then
        Viewport.DisableRendering
        FormMain.MainCanvas(0).SetZoomDropDownIndex PDImages.GetActiveImage.GetZoomIndex()
        Viewport.EnableRendering
    End If
    
End Sub

'If an image has multiple layers, call this function to enable any UI elements that
' operate on multiple layers.

'Note that some multi-layer settings require certain additional criteria to be met,
' e.g. "Merge Visible Layers" requires at least two visible layers, so it must still
' be handled specially.  This function is only for functions that are ALWAYS available
' if multiple layers are present in an image.
Private Sub SetUIMode_MultipleLayers()
    Menus.SetMenuEnabled "layer_delete", True
    Menus.SetMenuEnabled "layer_order", True
    Menus.SetMenuEnabled "layer_splitlayertoimage", True
    Menus.SetMenuEnabled "layer_splitalllayerstoimages", True
End Sub

'If an image has only one layer (e.g. a loaded JPEG), call this function to disable any UI elements
' that require multiple layers.
Private Sub SetUIMode_OnlyOneLayer()
    Menus.SetMenuEnabled "image_flatten", False
    Menus.SetMenuEnabled "image_mergevisible", False
    Menus.SetMenuEnabled "layer_delete", False
    Menus.SetMenuEnabled "layer_mergeup", False
    Menus.SetMenuEnabled "layer_mergedown", False
    Menus.SetMenuEnabled "layer_order", False
    Menus.SetMenuEnabled "layer_splitlayertoimage", False
    Menus.SetMenuEnabled "layer_splitalllayerstoimages", False
End Sub

'If an image has at least one valid layer (as they always do in PD), call this function to enable relevant layer menus and controls.
Private Sub SetUIMode_AtLeastOneLayer()
    Menus.SetMenuEnabled "layer_orientation", True
    Menus.SetMenuEnabled "layer_resize", True
    Menus.SetMenuEnabled "layer_transparency", True
    Menus.SetMenuEnabled "layer_rasterize", True
End Sub

'If PD ever reaches a "no layers in the current image" state, this function should be called.
' (Such a state is currently unsupported, so this exists only as a failsafe.)
Private Sub SetUIMode_NoLayers()
    
    'Image menu
    Menus.SetMenuEnabled "image_flatten", False
    Menus.SetMenuEnabled "image_mergevisible", False
    
    'Layer menu
    Menus.SetMenuEnabled "layer_delete", False
    Menus.SetMenuEnabled "layer_mergeup", False
    Menus.SetMenuEnabled "layer_mergedown", False
    Menus.SetMenuEnabled "layer_order", False
    Menus.SetMenuEnabled "layer_visibility", False
    Menus.SetMenuEnabled "layer_orientation", False
    Menus.SetMenuEnabled "layer_resize", False
    Menus.SetMenuEnabled "layer_cropselection", False
    Menus.SetMenuEnabled "layer_transparency", False
    Menus.SetMenuEnabled "layer_rasterize", False
    
End Sub

'Whenever PD returns to a "no images loaded" state, this function should be called.  (There are a number of specialized UI decisions
' required by this this state, and it's important to keep those options in one place.)
Private Sub SetUIMode_NoImages()
    
    'Start by forcibly disabling every conceivable UI group that requires an underlying image
    SetUIGroupState PDUI_Save, False
    SetUIGroupState PDUI_SaveAs, False
    SetUIGroupState PDUI_Close, False
    SetUIGroupState PDUI_EditCopyCut, False
    SetUIGroupState PDUI_View, False
    SetUIGroupState PDUI_ImageMenu, False
    SetUIGroupState PDUI_Selections, False
    SetUIGroupState PDUI_Macros, False
    SetUIGroupState PDUI_LayerTools, False
    SetUIGroupState PDUI_Undo, False
    SetUIGroupState PDUI_Redo, False
    SetUIGroupState PDUI_ICCProfile, False
    SetUIGroupState PDUI_FileOnDisk, False
    
    'Disable various layer-related toolbox options as well
    If (g_CurrentTool = NAV_MOVE) Then
        toolpanel_MoveSize.cmdLayerAffinePermanent.Enabled = False
    End If
    
    'Multiple edit menu items must also be disabled
    Menus.SetMenuEnabled "edit_history", False
    Menus.SetMenuEnabled "edit_repeat", False
    Menus.SetMenuEnabled "edit_fade", False
    Menus.RequestCaptionChange_ByName "edit_repeat", g_Language.TranslateMessage("Repeat"), True
    Menus.RequestCaptionChange_ByName "edit_fade", g_Language.TranslateMessage("Fade..."), True
    
    'Reset the main window's caption to its default PD name and version
    If (Not g_WindowManager Is Nothing) Then
        g_WindowManager.SetWindowCaptionW FormMain.hWnd, Interface.GetWindowCaption(Nothing)
    Else
        FormMain.Caption = Updates.GetPhotoDemonNameAndVersion()
    End If
        
    'Ask the canvas to reset itself.  Note that this also covers the status bar area and the image tabstrip, if they were
    ' previously visible.
    FormMain.MainCanvas(0).ClearCanvas
    
    'Restore the default taskbar and titlebar icons
    IconsAndCursors.ResetAppIcons
        
    'With all menus reset to their default values, we can now redraw all associated menu icons.
    ' (IMPORTANT: this function must be called whenever menu captions change, because icons are associated by caption.)
    IconsAndCursors.ResetMenuIcons
    
    'Ensure the Windows menu does not list any open images.
    Menus.UpdateSpecialMenu_WindowsOpen
    
    'If no images are currently open, but images were previously opened during this session, release any memory associated
    ' with those images.  This helps minimize PD's memory usage at idle.
    If (PDImages.GetNumSessionImages >= 1) Then PDImages.ReleaseAllPDImageResources
    
    'Some tools rely on image size (e.g. Crop, which can be constrained to image size, or Clone Stamp,
    ' which may have sampled from a now-unloaded image).  Notify them of changes so they can potentially
    ' free resources.
    Tools.NotifyImageSizeChanged
    
    'Forcibly blank out the current message if no images are loaded
    Message vbNullString
    
End Sub

'Whenever PD enters an "at least one valid image loaded" state, this function should be called.  Note that this function does not
' set any image-specific information; instead, it simply reverses a number of UI options that are disabled when no images exist.
Private Sub SetUIMode_AtLeastOneImage()
    
    SetUIGroupState PDUI_SaveAs, True
    SetUIGroupState PDUI_Close, True
    SetUIGroupState PDUI_EditCopyCut, True
    SetUIGroupState PDUI_View, True
    SetUIGroupState PDUI_ImageMenu, True
    SetUIGroupState PDUI_Macros, True
    SetUIGroupState PDUI_LayerTools, True
    
    'Make sure scroll bars are enabled and positioned correctly on the canvas
    FormMain.MainCanvas(0).AlignCanvasView
    
End Sub

'Some non-destructive actions need to synchronize *only* Undo/Redo buttons and menus (and their
' related counterparts, e.g. "Fade").  To make these actions snappier, I have pulled all Undo/Redo
' UI sync code into this separate sub, which can be called on-demand as necessary.
'
'If the caller will be calling ResetMenuIcons() after using this function, make sure to pass the
' optional suspendAssociatedRedraws as TRUE to prevent unnecessary menu redraws.
'
'Finally, if no images are loaded, this function does absolutely nothing.  Refer to SetUIMode_NoImages(),
' above, for additional details.
Public Sub SyncUndoRedoInterfaceElements(Optional ByVal suspendAssociatedRedraws As Boolean = False)

    If PDImages.IsImageActive() Then
    
        'Save is a bit funny, because if the image HAS been saved to file, we DISABLE the save button.
        SetUIGroupState PDUI_Save, Not PDImages.GetActiveImage.GetSaveState(pdSE_AnySave)
        
        'Undo, Redo, Repeat and Fade are all closely related
        If Not (PDImages.GetActiveImage.UndoManager Is Nothing) Then
        
            SetUIGroupState PDUI_Undo, PDImages.GetActiveImage.UndoManager.GetUndoState
            SetUIGroupState PDUI_Redo, PDImages.GetActiveImage.UndoManager.GetRedoState
            
            'Undo history is enabled if either Undo or Redo is active
            Menus.SetMenuEnabled "edit_history", (PDImages.GetActiveImage.UndoManager.GetUndoState Or PDImages.GetActiveImage.UndoManager.GetRedoState)
            
            '"Edit > Repeat..." and "Edit > Fade..." are also handled by the current image's undo manager (as it
            ' maintains the list of changes applied to the image, and links to copies of previous image state DIBs).
            Dim tmpDIB As pdDIB, tmpLayerIndex As Long, tmpActionName As String
            
            'See if the "Find last relevant layer action" function in the Undo manager returns TRUE or FALSE.  If it returns TRUE,
            ' enable both Repeat and Fade, and rename each menu caption so the user knows what is being repeated/faded.
            If PDImages.GetActiveImage.UndoManager.FillDIBWithLastUndoCopy(tmpDIB, tmpLayerIndex, tmpActionName, True) Then
                Menus.RequestCaptionChange_ByName "edit_fade", g_Language.TranslateMessage("Fade: %1...", g_Language.TranslateMessage(tmpActionName)), True
                Menus.SetMenuEnabled "edit_fade", True
            Else
                Menus.RequestCaptionChange_ByName "edit_fade", g_Language.TranslateMessage("Fade..."), True
                Menus.SetMenuEnabled "edit_fade", False
            End If
            
            'Repeat the above steps, but use the "Repeat" detection algorithm (which uses slightly different criteria;
            ' e.g. "Rotate Whole Image" cannot be faded, but it can be repeated)
            If PDImages.GetActiveImage.UndoManager.DoesStackContainRepeatableCommand(tmpActionName) Then
                Menus.RequestCaptionChange_ByName "edit_repeat", g_Language.TranslateMessage("Repeat: %1", g_Language.TranslateMessage(tmpActionName)), True
                Menus.SetMenuEnabled "edit_repeat", True
            Else
                Menus.RequestCaptionChange_ByName "edit_repeat", g_Language.TranslateMessage("Repeat"), True
                Menus.SetMenuEnabled "edit_repeat", False
            End If
            
            'Because these changes may modify menu captions, menu icons need to be reset (as they are tied to menu captions)
            If (Not suspendAssociatedRedraws) Then IconsAndCursors.ResetMenuIcons
        
        End If
    
    End If

End Sub

'SetUIGroupState enables or disables a swath of controls related to a simple keyword
' (e.g. "Undo", which affects multiple menu items and toolbar buttons)
Public Sub SetUIGroupState(ByVal metaItem As PD_UI_Group, ByVal newState As Boolean)
    
    Dim i As Long
    
    Select Case metaItem
            
        'Save (left-hand panel button(s) AND menu item)
        Case PDUI_Save
            If (Menus.IsMenuEnabled("file_save") <> newState) Then
                toolbar_Toolbox.cmdFile(FILE_SAVE).Enabled = newState
                Menus.SetMenuEnabled "file_save", newState
                Menus.SetMenuEnabled "file_revert", newState
            End If
            
        'Save As (menu item only).  Note that Save Copy is also tied to Save As functionality,
        ' because they use the same rules for enablement (e.g. disabled if no images are loaded,
        ' always enabled otherwise)
        Case PDUI_SaveAs
            If (Menus.IsMenuEnabled("file_saveas") <> newState) Then
                toolbar_Toolbox.cmdFile(FILE_SAVEAS_LAYERS).Enabled = newState
                toolbar_Toolbox.cmdFile(FILE_SAVEAS_FLAT).Enabled = newState
                Menus.SetMenuEnabled "file_saveas", newState
                Menus.SetMenuEnabled "file_savecopy", newState
                Menus.SetMenuEnabled "file_export", newState
            End If
            
        'Close (and Close All)
        Case PDUI_Close
            If (Menus.IsMenuEnabled("file_close") <> newState) Then
                toolbar_Toolbox.cmdFile(FILE_CLOSE).Enabled = newState
                Menus.SetMenuEnabled "file_close", newState
                Menus.SetMenuEnabled "file_closeall", newState
            End If
        
        'Undo (left-hand panel button AND menu item).  Undo toggles also control the "Fade last action" button,
        ' because that operates directly on previously saved Undo data.
        Case PDUI_Undo
        
            toolbar_Toolbox.cmdFile(FILE_UNDO).Enabled = newState
            Menus.SetMenuEnabled "edit_undo", newState
            
            'If Undo is being enabled, change the text to match the relevant action that created this Undo file
            If newState Then
                toolbar_Toolbox.cmdFile(FILE_UNDO).AssignTooltip PDImages.GetActiveImage.UndoManager.GetUndoProcessID, "Undo"
                Menus.RequestCaptionChange_ByName "edit_undo", g_Language.TranslateMessage("Undo:") & " " & g_Language.TranslateMessage(PDImages.GetActiveImage.UndoManager.GetUndoProcessID), True
            Else
                toolbar_Toolbox.cmdFile(FILE_UNDO).AssignTooltip "Undo last action"
                Menus.RequestCaptionChange_ByName "edit_undo", g_Language.TranslateMessage("Undo"), True
            End If
            
            'NOTE: when changing menu text, icons must be reapplied.  Make sure to call the ResetMenuIcons() function after changing
            ' Undo/Redo enablement.
            
        'Redo (left-hand panel button AND menu item)
        Case PDUI_Redo
            toolbar_Toolbox.cmdFile(FILE_REDO).Enabled = newState
            Menus.SetMenuEnabled "edit_redo", newState
            
            'If Redo is being enabled, change the menu text to match the relevant action that created this Undo file
            If newState Then
                toolbar_Toolbox.cmdFile(FILE_REDO).AssignTooltip PDImages.GetActiveImage.UndoManager.GetRedoProcessID, "Redo"
                Menus.RequestCaptionChange_ByName "edit_redo", g_Language.TranslateMessage("Redo:") & " " & g_Language.TranslateMessage(PDImages.GetActiveImage.UndoManager.GetRedoProcessID), True
            Else
                toolbar_Toolbox.cmdFile(FILE_REDO).AssignTooltip "Redo previous action"
                Menus.RequestCaptionChange_ByName "edit_redo", g_Language.TranslateMessage("Redo"), True
            End If
            
            'NOTE: when changing menu text, icons must be reapplied.  Make sure to call the ResetMenuIcons() function after changing
            ' Undo/Redo enablement.
            
        'Copy (menu item only)
        Case PDUI_EditCopyCut
            Menus.SetMenuEnabled "edit_copylayer", newState
            Menus.SetMenuEnabled "edit_copymerged", newState
            Menus.SetMenuEnabled "edit_cutlayer", newState
            Menus.SetMenuEnabled "edit_cutmerged", newState
            Menus.SetMenuEnabled "edit_pasteaslayer", newState
            Menus.SetMenuEnabled "edit_pastetocursor", newState
            Menus.SetMenuEnabled "edit_specialcopy", newState
            Menus.SetMenuEnabled "edit_specialcut", newState
            
        'View (top-menu level)
        Case PDUI_View
            Menus.SetMenuEnabled "view_top", newState
            Menus.SetMenuChecked "show_layeredges", Drawing.Get_ShowLayerEdges()
            Menus.SetMenuChecked "show_smartguides", Drawing.Get_ShowSmartGuides()
            Menus.SetMenuChecked "snap_global", Snap.GetSnap_Global()
            Menus.SetMenuChecked "snap_canvasedge", Snap.GetSnap_CanvasEdge()
            Menus.SetMenuChecked "snap_centerline", Snap.GetSnap_Centerline()
            Menus.SetMenuChecked "snap_layer", Snap.GetSnap_Layer()
            Menus.SetMenuChecked "snap_angle_90", Snap.GetSnap_Angle90()
            Menus.SetMenuChecked "snap_angle_45", Snap.GetSnap_Angle45()
            Menus.SetMenuChecked "snap_angle_30", Snap.GetSnap_Angle30()
            
        'ImageOps is all Image-related menu items; it enables/disables the Image, Layer, Select, Color, and Print menus.
        ' (This flag is very useful for items that require at least one open image to operate.)
        Case PDUI_ImageMenu
            If (Menus.IsMenuEnabled("image_top") <> newState) Then
                Menus.SetMenuEnabled "image_top", newState
                Menus.SetMenuEnabled "layer_top", newState
                Menus.SetMenuEnabled "select_top", newState
                Menus.SetMenuEnabled "adj_top", newState
                Menus.SetMenuEnabled "effects_top", newState
                Menus.SetMenuEnabled "file_print", newState
                
                'The edit menu also contains items that require an open image to operate
                Menus.SetMenuEnabled "edit_clear", newState
                Menus.SetMenuEnabled "edit_fill", newState
                Menus.SetMenuEnabled "edit_stroke", newState
    
            End If
            
        'Macro (within the Tools menu)
        Case PDUI_Macros
            Menus.SetMenuEnabled "tools_macrocreatetop", newState
            Menus.SetMenuEnabled "tools_playmacro", newState
            Menus.SetMenuEnabled "tools_recentmacros", newState
            
        'Selections in general
        Case PDUI_Selections
            
            'If selections are not active, clear all selection value spin controls.
            ' (These used to be called text up/downs, per Windows convention, hence the tud- prefix.)
            If Tools.IsSelectionToolActive Then

                If (Not newState) Then
                    For i = 0 To toolpanel_Selections.tudSel.Count - 1
                        If (toolpanel_Selections.tudSel(i).Min > 0) Then
                            toolpanel_Selections.tudSel(i).Value = toolpanel_Selections.tudSel(i).Min
                        Else
                            toolpanel_Selections.tudSel(i).Value = 0
                        End If
                    Next i
                End If

                'Set selection text boxes to enable only when a selection is active.  Other selection
                'vcontrols can remain active even without a selection present; this allows the user to
                ' set certain parameters in advance, so when they actually draw a selection, it already
                ' has the attributes they want - but spin controls are an exception to this.
                For i = 0 To toolpanel_Selections.tudSel.Count - 1
                    toolpanel_Selections.tudSel(i).Enabled = newState
                Next i

            End If
            
            'En/disable all selection menu items that rely on an existing selection to operate
            If (Menus.IsMenuEnabled("select_none") <> newState) Then
                
                'Select none, invert selection
                Menus.SetMenuEnabled "select_none", newState
                Menus.SetMenuEnabled "select_invert", newState
                
                'Grow/shrink/border/feather/sharpen selection
                Menus.SetMenuEnabled "select_grow", newState
                Menus.SetMenuEnabled "select_shrink", newState
                Menus.SetMenuEnabled "select_border", newState
                Menus.SetMenuEnabled "select_feather", newState
                Menus.SetMenuEnabled "select_sharpen", newState
                
                'Modify selected pixels in various ways
                Menus.SetMenuEnabled "select_erasearea", newState
                Menus.SetMenuEnabled "select_fill", newState
                Menus.SetMenuEnabled "select_heal", newState
                Menus.SetMenuEnabled "select_stroke", newState
                
                'Save selection
                Menus.SetMenuEnabled "select_save", newState
                
                'Export selection top-level menu
                Menus.SetMenuEnabled "select_export", newState
                
            End If
                                    
            'Selection enabling/disabling also affects the two Crop to Selection commands (one in the Image menu, one in the Layer menu)
            Menus.SetMenuEnabled "image_crop", newState
            Menus.SetMenuEnabled "layer_cropselection", newState
            
            'The content-aware fill option in the edit menu also requires an active selection.
            Menus.SetMenuEnabled "edit_contentawarefill", newState
            
        'Transformable selection controls specifically
        Case PDUI_SelectionTransforms
            
            If Tools.IsSelectionToolActive Then
            
                'Under certain circumstances, it is desirable to disable only the selection location boxes
                For i = 0 To toolpanel_Selections.tudSel.Count - 1
                    
                    If (Not newState) Then
                        If (toolpanel_Selections.tudSel(i).Min > 0) Then
                            toolpanel_Selections.tudSel(i).Value = toolpanel_Selections.tudSel(i).Min
                        Else
                            toolpanel_Selections.tudSel(i).Value = 0
                        End If
                    End If
                    
                    toolpanel_Selections.tudSel(i).Enabled = newState
                    
                Next i
                
                If newState Then SelectionUI.SyncTextToCurrentSelection PDImages.GetActiveImageID
                
            End If
                
        'If the ExifTool plugin is not available, metadata will ALWAYS be disabled.  (We do not currently have a
        ' separate fallback for reading/browsing/writing metadata.)
        Case PDUI_Metadata
        
            If PluginManager.IsPluginCurrentlyEnabled(CCP_ExifTool) Then
                Menus.SetMenuEnabled "image_editmetadata", newState
                Menus.SetMenuEnabled "image_removemetadata", newState
            Else
                Menus.SetMenuEnabled "image_editmetadata", False
                Menus.SetMenuEnabled "image_removemetadata", False
            End If
        
        'GPS metadata is its own sub-category, and its activation is contigent upon an image having embedded GPS data
        Case PDUI_GPSMetadata
        
            If PluginManager.IsPluginCurrentlyEnabled(CCP_ExifTool) Then
                Menus.SetMenuEnabled "image_maplocation", newState
            Else
                Menus.SetMenuEnabled "image_maplocation", False
            End If
        
        'Various layer-related tools (move, etc) are exposed on the tool options dialog.  For consistency, we disable those UI elements
        ' when no images are loaded.
        Case PDUI_LayerTools
            
            'Because we're dealing with text up/downs, we need to set hard limits relative to the current image's size.
            ' I'm currently using the "rule of three" - max/min values are the current dimensions of the image, x3.
            Dim minLayerUIValue_Width As Long, maxLayerUIValue_Width As Long
            Dim minLayerUIValue_Height As Long, maxLayerUIValue_Height As Long
            
            If newState Then
                maxLayerUIValue_Width = PDImages.GetActiveImage.Width * 3
                maxLayerUIValue_Height = PDImages.GetActiveImage.Height * 3
            Else
                maxLayerUIValue_Width = 0
                maxLayerUIValue_Height = 0
            End If
            
            'Make sure width/height values are non-zero
            If (maxLayerUIValue_Width = 0) Then maxLayerUIValue_Width = 1
            If (maxLayerUIValue_Height = 0) Then maxLayerUIValue_Height = 1
            
            'Minimum values are simply the negative of the max values
            minLayerUIValue_Width = -1 * maxLayerUIValue_Width
            minLayerUIValue_Height = -1 * maxLayerUIValue_Height
            
            'Mark the tool engine as busy; this prevents control changes from triggering viewport redraws
            Tools.SetToolBusyState True
            
            'Enable/disable all UI elements as necessary
            If (g_CurrentTool = NAV_MOVE) Then
                
                'First, enable all move/size panels
                For i = 0 To toolpanel_MoveSize.tudLayerMove.Count - 1
                    If (toolpanel_MoveSize.tudLayerMove(i).Enabled <> newState) Then toolpanel_MoveSize.tudLayerMove(i).Enabled = newState
                Next i
                
                'Where relevant, also update control bounds and values
                If newState Then
                
                    For i = 0 To toolpanel_MoveSize.tudLayerMove.Count - 1
                        
                        'Even-numbered indices correspond to width; odd-numbered to height
                        If (i Mod 2 = 0) Then
                            toolpanel_MoveSize.tudLayerMove(i).Min = minLayerUIValue_Width
                            toolpanel_MoveSize.tudLayerMove(i).Max = maxLayerUIValue_Width
                        Else
                            toolpanel_MoveSize.tudLayerMove(i).Min = minLayerUIValue_Height
                            toolpanel_MoveSize.tudLayerMove(i).Max = maxLayerUIValue_Height
                        End If
                        
                    Next i
                    
                    'The Layer Move tool has four text up/downs: two for layer position (x, y) and two for layer size (w, y)
                    toolpanel_MoveSize.tudLayerMove(0).Value = PDImages.GetActiveImage.GetActiveLayer.GetLayerOffsetX
                    toolpanel_MoveSize.tudLayerMove(1).Value = PDImages.GetActiveImage.GetActiveLayer.GetLayerOffsetY
                    toolpanel_MoveSize.tudLayerMove(2).Value = PDImages.GetActiveImage.GetActiveLayer.GetLayerWidth
                    toolpanel_MoveSize.tudLayerMove(3).Value = PDImages.GetActiveImage.GetActiveLayer.GetLayerHeight
                    toolpanel_MoveSize.tudLayerMove(2).DefaultValue = PDImages.GetActiveImage.GetActiveLayer.GetLayerWidth(False)
                    toolpanel_MoveSize.tudLayerMove(3).DefaultValue = PDImages.GetActiveImage.GetActiveLayer.GetLayerHeight(False)
                    
                    'The layer resize quality combo box also needs to be synched
                    toolpanel_MoveSize.cboLayerResizeQuality.ListIndex = PDImages.GetActiveImage.GetActiveLayer.GetLayerResizeQuality
                    
                    'Layer angle and shear are newly available as of 7.0
                    toolpanel_MoveSize.sltLayerAngle.Value = PDImages.GetActiveImage.GetActiveLayer.GetLayerAngle
                    toolpanel_MoveSize.sltLayerShearX.Value = PDImages.GetActiveImage.GetActiveLayer.GetLayerShearX
                    toolpanel_MoveSize.sltLayerShearY.Value = PDImages.GetActiveImage.GetActiveLayer.GetLayerShearY
                    
                End If
                
            End If
            
            'Free the tool engine
            Tools.SetToolBusyState False
            
        'Images with embedded color profiles support extra features
        Case PDUI_ICCProfile
            Menus.SetMenuEnabled "file_export_colorprofile", newState
            
        Case PDUI_FileOnDisk
            Menus.SetMenuEnabled "image_showinexplorer", newState
            
    End Select
    
End Sub

Public Function IsModalDialogActive() As Boolean
    IsModalDialogActive = m_ModalDialogActive Or m_SystemDialogActive
End Function

Public Sub NotifySystemDialogState(ByVal dialogIsActive As Boolean)
    m_SystemDialogActive = dialogIsActive
End Sub

'For best results, any modal form should be shown via this function.  This function will automatically center the form over the main window,
' while also properly assigning ownership so that the dialog is truly on top of any active windows.  It also handles deactivation of
' other windows (to prevent click-through), and dynamic top-most behavior to ensure that the program doesn't steal focus if the user switches
' to another program while a modal dialog is active.
Public Sub ShowPDDialog(ByRef dialogModality As FormShowConstants, ByRef dialogForm As Form, Optional ByVal doNotUnload As Boolean = False)

    On Error GoTo ShowPDDialogError
    
    m_ModalDialogActive = True
    
    'Make sure PD's main form is visible
    If (FormMain.WindowState = vbMinimized) Then FormMain.WindowState = vbNormal
    
    'Turn off any async pipe connections or other listeners
    FormMain.ChangeSessionListenerState False
    
    'Reset our "last dialog result" tracker.  (We use "ignore" as the "default" value, as it's a value PD never utilizes internally.)
    m_LastShowDialogResult = vbIgnore
    
    'Start by loading the form and hiding it
    If (dialogForm Is Nothing) Then Load dialogForm
    dialogForm.Visible = False
    
    'Store a reference to this dialog; if subsequent dialogs are loaded, this dialog will be given ownership over them
    If (currentDialogReference Is Nothing) Then
        
        'This is a regular modal dialog, and the main form should be its owner
        isSecondaryDialog = False
        Set currentDialogReference = dialogForm
                
    Else
    
        'We already have a reference to a modal dialog - that means a modal dialog is raising *another* modal dialog.  Give the previous
        ' modal dialog ownership over this new dialog!
        isSecondaryDialog = True
        
    End If
    
    'Retrieve and cache the hWnd; we need access to this even if the form is unloaded, so we can properly deregister it
    ' with the window manager.
    Dim dialogHWnd As Long
    dialogHWnd = dialogForm.hWnd
    
    'If the window has a previous position stored, use that.
    Dim prevPositionStored As Boolean
    If (Not g_WindowManager Is Nothing) Then prevPositionStored = g_WindowManager.IsPreviousPositionStored(dialogForm)
    
    'If a previous position is *not* stored, center it against the main dialog.
    If (Not prevPositionStored) Then
        
        'Get the rect of the main form, which we will use to calculate a center position
        Dim ownerRect As winRect
        GetWindowRect FormMain.hWnd, ownerRect
        
        'Determine the center of that rect
        Dim centerX As Long, centerY As Long
        centerX = ownerRect.x1 + (ownerRect.x2 - ownerRect.x1) \ 2
        centerY = ownerRect.y1 + (ownerRect.y2 - ownerRect.y1) \ 2
        
        'Get the rect of the child dialog
        Dim dialogRect As winRect
        GetWindowRect dialogHWnd, dialogRect
        
        'Determine an upper-left point for the dialog based on its size
        Dim newLeft As Long, newTop As Long
        newLeft = centerX - ((dialogRect.x2 - dialogRect.x1) \ 2)
        newTop = centerY - ((dialogRect.y2 - dialogRect.y1) \ 2)
        
        'If this position results in the dialog sitting off-screen, move it so that its bottom-right corner is always on-screen.
        ' (All PD dialogs have bottom-right OK/Cancel buttons, so that's the most important part of the dialog to show.)
        If newLeft + (dialogRect.x2 - dialogRect.x1) > g_Displays.GetDesktopRight Then newLeft = g_Displays.GetDesktopRight - (dialogRect.x2 - dialogRect.x1)
        If newTop + (dialogRect.y2 - dialogRect.y1) > g_Displays.GetDesktopBottom Then newTop = g_Displays.GetDesktopBottom - (dialogRect.y2 - dialogRect.y1)
        
        'Move the dialog into place, but do not repaint it (that will be handled in a moment by the .Show event)
        MoveWindow dialogHWnd, newLeft, newTop, dialogRect.x2 - dialogRect.x1, dialogRect.y2 - dialogRect.y1, 0
        
    End If
    
    'Mirror the current run-time window icons to the dialog; this allows the icons to appear in places like Alt+Tab
    ' on older OSes, even though a toolbox window has focus.
    Interface.FixPopupWindow dialogHWnd, True
    
    'Use VB to actually display the dialog.  Note that code execution will pause here until the form is closed.
    ' (As usual, disclaimers apply to message-loop functions like DoEvents.)
    dialogForm.Show dialogModality, FormMain
    
    'Now that the dialog has finished, we must replace the windows icons with its original ones -
    ' otherwise, VB will mistakenly unload our custom icons with the window!
    Interface.FixPopupWindow dialogHWnd, False
    
    'Release our reference to this dialog
    If isSecondaryDialog Then
        isSecondaryDialog = False
    Else
        Set currentDialogReference = Nothing
    End If
    
    'If the form has not been unloaded, unload it now
    If (Not (dialogForm Is Nothing)) And (Not doNotUnload) Then
        Unload dialogForm
        Set dialogForm = Nothing
    End If
    
    'Reinstate any async listeners
    FormMain.ChangeSessionListenerState True
    
    m_ModalDialogActive = False
    
    Exit Sub
    
'For reasons I can't yet ascertain, this function will sometimes fail, claiming that a modal window is already active.  If that happens,
' we can just exit.
ShowPDDialogError:

    m_ModalDialogActive = False

End Sub

'If you need to raise a popup from a custom owner (like the Tools > Options panel, where PD messes with form ownership),
' use this function instead.  It doesn't provide the same auto-positioning behavior as PD's core "ShowDialog", but it
' allows you to specify a custom parent.
Public Sub ShowCustomPopup(ByRef dialogModality As FormShowConstants, ByRef dialogForm As Form, ByRef parentForm As Form, Optional ByVal doNotUnload As Boolean = False)

    On Error GoTo ShowModalDialogError
    
    Dim initModalDialogState As Boolean
    initModalDialogState = m_ModalDialogActive
    m_ModalDialogActive = True
    
    'Make sure PD's main form is visible
    If (FormMain.WindowState = vbMinimized) Then FormMain.WindowState = vbNormal
    
    'TODO: do any callers actually require this?
    'Reset our "last dialog result" tracker.  (We use "ignore" as the "default" value, as it's a value PD never utilizes internally.)
    'm_LastShowDialogResult = vbIgnore
    
    'Start by loading the form and hiding it
    If (dialogForm Is Nothing) Then Load dialogForm
    dialogForm.Visible = False
    
    'Retrieve and cache the hWnd; we need access to this even if the form is unloaded, so we can properly deregister it
    ' with the window manager.
    Dim dialogHWnd As Long
    dialogHWnd = dialogForm.hWnd
    
    'If the window has a previous position stored, use that.
    'Dim prevPositionStored As Boolean
    'If (Not g_WindowManager Is Nothing) Then prevPositionStored = g_WindowManager.IsPreviousPositionStored(dialogForm)
    
    'If a previous position is *not* stored, center it against the parent dialog.
    'If (Not prevPositionStored) Then
        
        'Get the rect of the main form, which we will use to calculate a center position
        Dim ownerRect As winRect
        GetWindowRect parentForm.hWnd, ownerRect
        
        'Determine the center of that rect
        Dim centerX As Long, centerY As Long
        centerX = ownerRect.x1 + (ownerRect.x2 - ownerRect.x1) \ 2
        centerY = ownerRect.y1 + (ownerRect.y2 - ownerRect.y1) \ 2
        
        'Get the rect of the child dialog
        Dim dialogRect As winRect
        GetWindowRect dialogHWnd, dialogRect
        
        'Determine an upper-left point for the dialog based on its size
        Dim newLeft As Long, newTop As Long
        newLeft = centerX - ((dialogRect.x2 - dialogRect.x1) \ 2)
        newTop = centerY - ((dialogRect.y2 - dialogRect.y1) \ 2)
        
        'If this position results in the dialog sitting off-screen, move it so that its bottom-right corner is always on-screen.
        ' (All PD dialogs have bottom-right OK/Cancel buttons, so that's the most important part of the dialog to show.)
        If newLeft + (dialogRect.x2 - dialogRect.x1) > g_Displays.GetDesktopRight Then newLeft = g_Displays.GetDesktopRight - (dialogRect.x2 - dialogRect.x1)
        If newTop + (dialogRect.y2 - dialogRect.y1) > g_Displays.GetDesktopBottom Then newTop = g_Displays.GetDesktopBottom - (dialogRect.y2 - dialogRect.y1)
        
        'Move the dialog into place, but do not repaint it (that will be handled in a moment by the .Show event)
        MoveWindow dialogHWnd, newLeft, newTop, dialogRect.x2 - dialogRect.x1, dialogRect.y2 - dialogRect.y1, 0
        
    'End If
        
    'Mirror the current run-time window icons to the dialog; this allows the icons to appear in places like Alt+Tab
    ' on older OSes, even though a toolbox window has focus.
    Interface.FixPopupWindow dialogHWnd, True
    
    'Use VB to actually display the dialog.  Note that code execution will pause here until the form is closed.
    ' (As usual, disclaimers apply to message-loop functions like DoEvents.)
    dialogForm.Show dialogModality, parentForm
    
    'Now that the dialog has finished, we must replace the windows icons with its original ones -
    ' otherwise, VB will mistakenly unload our custom icons with the window!
    Interface.FixPopupWindow dialogHWnd, False
    
    'If the form has not been unloaded, unload it now
    If (Not (dialogForm Is Nothing)) And (Not doNotUnload) Then
        Unload dialogForm
        Set dialogForm = Nothing
    End If
    
    m_ModalDialogActive = initModalDialogState
    
    Exit Sub
    
'For reasons I can't yet ascertain, this function will sometimes fail, claiming that a modal window is already active.  If that happens,
' we can just exit.
ShowModalDialogError:
    m_ModalDialogActive = False

End Sub

'Any commandbar-based dialog will automatically notify us of its "OK" or "Cancel" result; subsequent functions can check this return
' via GetLastShowDialogResult(), below.
Public Sub NotifyShowDialogResult(ByVal msgResult As VbMsgBoxResult, Optional ByVal nonStandardDialogSource As Boolean = False)
    
    'Only store the result if the dialog was initiated via ShowPDDialog, above
    If m_ModalDialogActive Or nonStandardDialogSource Then m_LastShowDialogResult = msgResult
    
End Sub

'This function will tell you if the last commandbar-based dialog was closed via OK or CANCEL.
'IMPORTANT NOTE: calling this function will RESET THE LAST-GENERATED RESULT, by design.
Public Function GetLastShowDialogResult() As VbMsgBoxResult
    GetLastShowDialogResult = m_LastShowDialogResult
    m_LastShowDialogResult = vbIgnore
End Function

'When raising a modal dialog, we want to set the window ownership to the top-most (relevant) window in the program, which may
' or may not be the main program window.  This function can called to determine the proper owner of an arbitrary modal dialog box.
'
'If the caller knows in advance that a modal dialog is owned by another modal dialog (for example, a tool dialog displaying a color
' selection dialog), it can explicitly mark the assumeSecondaryDialog function as TRUE.
Public Function GetModalOwner(Optional ByVal assumeSecondaryDialog As Boolean = False) As Form

    'If a modal dialog is already active, it gets ownership over subsequent dialogs
    If isSecondaryDialog Or assumeSecondaryDialog Then
        Set GetModalOwner = currentDialogReference
        
    'No modal dialog is active, making this the only one.  Give the main form ownership.
    Else
        Set GetModalOwner = FormMain
    End If
    
End Function

'When raising nested dialogs (e.g. a modal effect dialog raises a "new preset" window), we need to dynamically assign and release
' icons to each popup.  This ensures that PD stays visible in places like Alt+Tab, even on pre-Win10 systems.
Public Sub FixPopupWindow(ByVal targetHWnd As Long, Optional ByVal windowIsLoading As Boolean = False)

    If windowIsLoading Then
        
        'We could dynamically resize our tracking collection to precisely match the number of open windows, but this would only
        ' save us a few bytes.  Since we know we'll never exceed NESTED_POPUP_LIMIT, we just default to the max size off the bat.
        If (Not VBHacks.IsArrayInitialized(m_PopupHWnds)) Then
            m_NumOfPopupHWnds = 0
            ReDim m_PopupHWnds(0 To NESTED_POPUP_LIMIT - 1) As Long
            ReDim m_PopupIconsSmall(0 To NESTED_POPUP_LIMIT - 1) As Long
            ReDim m_PopupIconsLarge(0 To NESTED_POPUP_LIMIT - 1) As Long
        End If
        
        'We don't actually need to store the target window's hWnd (at present) but we grab it "just in case" future enhancements
        ' require it.
        m_PopupHWnds(m_NumOfPopupHWnds) = targetHWnd
        
        'We can't guarantee that Aero is active on Win 7 and earlier, so we must jump through some extra hoops to make
        ' sure the popup window appears inside any raised Alt+Tab dialogs
        If (Not OS.IsWin8OrLater) Then g_WindowManager.ForceWindowAppearInAltTab targetHWnd, True
        
        'While here, cache the window's current icons.  (VB may insert its own default icons for some window types.)
        ' When the dialog is closed, we will restore these icons to avoid leaking any of PD's custom icons.
        IconsAndCursors.MirrorCurrentIconsToWindow targetHWnd, True, m_PopupIconsSmall(m_NumOfPopupHWnds), m_PopupIconsLarge(m_NumOfPopupHWnds)
        m_NumOfPopupHWnds = m_NumOfPopupHWnds + 1
        
    Else
    
        m_NumOfPopupHWnds = m_NumOfPopupHWnds - 1
        If (m_NumOfPopupHWnds >= 0) Then
            IconsAndCursors.ChangeWindowIcon m_PopupHWnds(m_NumOfPopupHWnds), m_PopupIconsSmall(m_NumOfPopupHWnds), m_PopupIconsLarge(m_NumOfPopupHWnds)
        Else
            m_NumOfPopupHWnds = 0
            PDDebug.LogAction "WARNING!  Interface.FixPopupWindow() has somehow unloaded more windows than it's loaded."
        End If
    
    End If

End Sub

'Return the system keyboard delay, in seconds.  This isn't an exact science because the delay is actually
' hardware dependent (e.g. the system returns a value from 0 to 3), but we use a "good enough" approximation.
Public Function GetKeyboardDelay() As Double
    Dim keyDelayIndex As Long
    SystemParametersInfo SPI_GETKEYBOARDDELAY, 0, keyDelayIndex, 0
    GetKeyboardDelay = (keyDelayIndex + 1) * 0.25
End Function

'Return the system keyboard repeat rate, in seconds.  This isn't an exact science because the delay is actually
' hardware dependent (e.g. the system returns a value from 0 to 31), but we use a "good enough" approximation.
Public Function GetKeyboardRepeatRate() As Double
    
    Dim keyRepeatIndex As Long
    SystemParametersInfo SPI_GETKEYBOARDSPEED, 0, keyRepeatIndex, 0
    
    'Per MSDN (http://msdn.microsoft.com/en-us/library/windows/desktop/ms724947%28v=vs.85%29.aspx#Input)
    ' "Retrieves the keyboard repeat-speed setting, which is a value in the range from 0 (approximately 2.5 repetitions per second)
    ' through 31 (approximately 30 repetitions per second). The actual repeat rates are hardware-dependent and may vary from a linear
    ' scale by as much as 20%. The pvParam parameter must point to a DWORD variable that receives the setting."
    '
    'The formula below mimics this behavior pretty closely (35 repetitions per second at the high-end, but identical at the low end)
    GetKeyboardRepeatRate = (400 - (keyRepeatIndex * 12)) / 1000
    
End Function

Public Sub ToggleImageTabstripAlignment(ByVal newAlignment As AlignConstants, Optional ByVal suppressPrefUpdate As Boolean = False)
    
    'Reset the menu checkmarks
    Dim curMenuIndex As Long
    If (newAlignment = vbAlignLeft) Then
        curMenuIndex = 4
    ElseIf (newAlignment = vbAlignTop) Then
        curMenuIndex = 5
    ElseIf (newAlignment = vbAlignRight) Then
        curMenuIndex = 6
    ElseIf (newAlignment = vbAlignBottom) Then
        curMenuIndex = 7
    End If
    
    Dim i As Long
    For i = 4 To 7
        FormMain.MnuWindowTabstrip(i).Checked = (i = curMenuIndex)
    Next i
    
    'Write the preference out to file, then notify the canvas of the change
    If (Not suppressPrefUpdate) Then UserPrefs.SetPref_Long "Core", "Image Tabstrip Alignment", CLng(newAlignment)
    FormMain.MainCanvas(0).NotifyImageStripAlignment newAlignment
    
End Sub

'The image tabstrip can set to appear under a variety of circumstances.  Use this sub to change the current setting; it will
' automatically handle syncing with the preferences file.
Public Sub ToggleImageTabstripVisibility(ByVal newSetting As Long, Optional ByVal suppressPrefUpdate As Boolean = False)

    'Start by synchronizing menu checkmarks to the selected option
    Dim i As Long
    For i = 0 To 2
        If (newSetting = i) Then
            FormMain.MnuWindowTabstrip(i).Checked = True
        Else
            FormMain.MnuWindowTabstrip(i).Checked = False
        End If
    Next i

    'Write the matching preference out to file, then notify the primary canvas of the change
    If (Not suppressPrefUpdate) Then UserPrefs.SetPref_Long "Core", "Image Tabstrip Visibility", newSetting
    FormMain.MainCanvas(0).NotifyImageStripVisibilityMode newSetting

End Sub

Public Function FixDPI(ByVal pxMeasurement As Long) As Long

    'The first time this function is called, m_DPIRatio will be 0.  Calculate it.
    If (m_DPIRatio = 0#) Then
    
        'There are 1440 twips in one inch.  (Twips are resolution-independent.)  Use that knowledge to calculate DPI.
        m_DPIRatio = 1440# / TwipsPerPixelXFix()
        
        'FYI: if the screen resolution is 96 dpi, this function will return the original pixel measurement, due to
        ' this calculation.
        m_DPIRatio = m_DPIRatio * (1# / 96#)
    
    End If
    
    FixDPI = Int(m_DPIRatio * CDbl(pxMeasurement) + 0.5)
    
End Function

Public Function FixDPIFloat(ByVal pxMeasurement As Long) As Double

    'The first time this function is called, m_DPIRatio will be 0.  Calculate it.
    If (m_DPIRatio = 0#) Then
    
        'There are 1440 twips in one inch.  (Twips are resolution-independent.)  Use that knowledge to calculate DPI.
        m_DPIRatio = 1440# / TwipsPerPixelXFix()
        
        'FYI: if the screen resolution is 96 dpi, this function will return the original pixel measurement, due to
        ' this calculation.
        m_DPIRatio = m_DPIRatio * (1# / 96#)
    
    End If
    
    FixDPIFloat = m_DPIRatio * CDbl(pxMeasurement)
    
End Function

'Fun fact: there are 15 twips per pixel at 96 DPI.  Not fun fact: at 200% DPI (e.g. 192 DPI), VB's internal
' TwipsPerPixelXFix will return 7, when actually we need the value 7.5.  This causes problems when resizing
' certain controls (like SmartCheckBox) because the size will actually come up short due to rounding errors!
' So whenever TwipsPerPixelXFix/Y is required, use these functions instead.
Public Function TwipsPerPixelXFix() As Double
    If (m_CurrentSystemDPI = 0) Then
        If (Screen.TwipsPerPixelX = 7) Then TwipsPerPixelXFix = 7.5 Else TwipsPerPixelXFix = Screen.TwipsPerPixelX
    Else
        TwipsPerPixelXFix = 15# / m_CurrentSystemDPI
    End If
End Function

Public Function TwipsPerPixelYFix() As Double
    If (m_CurrentSystemDPI = 0) Then
        If (Screen.TwipsPerPixelY = 7) Then TwipsPerPixelYFix = 7.5 Else TwipsPerPixelYFix = Screen.TwipsPerPixelY
    Else
        TwipsPerPixelYFix = 15# / m_CurrentSystemDPI
    End If
End Function

'ScaleX and ScaleY functions do not work when converting from pixels to twips, thanks to the 15 / 2 <> 7
' bug described above.  Instead of using ScaleX/Y functions, use these wrapper.
Public Function PXToTwipsX(ByVal srcPixelWidth As Long) As Long
    PXToTwipsX = srcPixelWidth * TwipsPerPixelXFix
End Function

Public Function PXToTwipsY(ByVal srcPixelHeight As Long) As Long
    PXToTwipsY = srcPixelHeight * TwipsPerPixelYFix
End Function

Public Sub DisplayWaitScreen(ByVal waitTitle As String, ByRef ownerForm As Form, Optional ByVal descriptionText As String = vbNullString, Optional ByVal raiseModally As Boolean = False)
    
    FormWait.Visible = False
    FormWait.lblWaitTitle.Caption = waitTitle
    FormWait.lblWaitTitle.Visible = True
    
    FormWait.lblWaitDescription.Caption = descriptionText
    FormWait.lblWaitDescription.Visible = (LenB(descriptionText) <> 0)
    
    Screen.MousePointer = vbHourglass
    
    If raiseModally Then
        FormWait.Show vbModal, ownerForm
    Else
        FormWait.Show vbModeless, ownerForm
    End If
    
End Sub

Public Sub HideWaitScreen()
    g_UnloadWaitWindow = False
    Screen.MousePointer = vbDefault
    Unload FormWait
End Sub

'Because VB6 apps look terrible on modern version of Windows, I do a bit of beautification to every form upon at load-time.
' This routine is nice because every form calls it at least once, so I can make centralized changes without having to rewrite
' code in every individual form.  This is also where run-time translation occurs.
Public Sub ApplyThemeAndTranslations(ByRef dstForm As Form, Optional ByVal handleFormPainting As Boolean = True, Optional ByVal handleAutoResize As Boolean = False, Optional ByVal hWndCustomAnchor As Long = 0)
    
    'Some forms call this function during the load step, meaning they will be triggered during compilation; avoid this
    If (Not PDMain.IsProgramRunning()) Then Exit Sub
    If (dstForm Is Nothing) Then Exit Sub
    
    'This function can be a rather egregious time hog, so I profile it from time-to-time to check for regressions
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    'FORM STEP 1: apply any form-level changes (like backcolor), as child controls may pull this automatically
    dstForm.BackColor = g_Themer.GetGenericUIColor(UI_Background)
    If handleFormPainting Then g_Themer.AddWindowPainter dstForm.hWnd
    dstForm.MouseIcon = Nothing
    dstForm.MousePointer = 0
    
    'TODO: solve icon issues here?
    If (dstForm.Name <> "FormMain") Then Set dstForm.Icon = Nothing
    
    'While we're here, notify the tab manager of the newly loaded form, and make a note of the form's hWnd so we
    ' can relay it to various child controls.
    Dim hostFormhWnd As Long
    hostFormhWnd = dstForm.hWnd
    NavKey.NotifyFormLoading dstForm, handleAutoResize, hWndCustomAnchor
    
    Dim ctlThemedOK As Boolean
    On Error GoTo ControlIsNotPD
    
    'FORM STEP 2: Enumerate through every control on the form and apply theming on a per-control basis.
    Dim eControl As Control
    For Each eControl In dstForm.Controls
        
        'We now want to ignore all built-in VB6 controls.  PhotoDemon doesn't use many of these
        ' (menus are the exception) but a few picture boxes may still linger in the project...
        If (TypeOf eControl Is Menu) Then
            
            'Don't do anything; we just need the Else at the end to not trigger on this control
            
        'PD still uses generic picture boxes in a few places.  Picture boxes get confused by all the
        ' weird run-time UI APIs we call, so to ensure that their cursors work properly, we use an API
        ' to reset their cursors as well.
        ElseIf (TypeOf eControl Is PictureBox) Then
        
            IconsAndCursors.SetArrowCursor eControl
            
            'While we're here, forcibly remove TabStops from each picture box.  They should never receive focus,
            ' but I often forget to change this at design-time.
            eControl.TabStop = False
            
            'Picture boxes should be replaced with pdPictureBox!
            PDDebug.LogAction "Found a legacy picturebox control on " & dstForm.Name & ": consider removing it!"
            
        Else
            
            ctlThemedOK = False
            
            'All of PhotoDemon's custom UI controls implement an UpdateAgainstCurrentTheme function.
            ' This function updates two things:
            ' 1) The control's visual appearance (to reflect any changes to visual themes)
            ' 2) Updating any translatable text against the current translation
            
            'Request a theme and language update.  (As of 7.0, this function also handles registration
            ' with the central tab key manager.)
            eControl.UpdateAgainstCurrentTheme hostFormhWnd
            ctlThemedOK = True
            
ControlIsNotPD:
            If (Not ctlThemedOK) Then
                PDDebug.LogAction "WARNING!  Non-PD control found on " & dstForm.Caption & ": " & eControl.Name
                On Error GoTo 0
            End If
            
        End If
        
    Next
    
    'Next, we need to translate any VB objects on the form.  At present, this only includes the Form caption;
    ' everything else is handled internally.
    If g_Language.TranslationActive And dstForm.Enabled Then g_Language.ApplyTranslations dstForm
    
    'Finally, restore the previous window position (if any)
    If (Not g_WindowManager Is Nothing) Then g_WindowManager.RestoreWindowLocation dstForm
    
    'Report timing results here:
    PDDebug.LogAction "Interface.ApplyThemeAndTranslations updated " & dstForm.Name & " in " & VBHacks.GetTimeDiffNowAsString(startTime)
    
End Sub

'When a themed form is unloaded, it may be desirable to release certain changes made to it - or in our case, unsubclass it.
' This function should be called when any themed form is unloaded.
Public Sub ReleaseFormTheming(ByRef srcForm As Form)
    
    'This function may be triggered during compilation; avoid this
    If PDMain.IsProgramRunning() Then
        If (Not g_Themer Is Nothing) Then g_Themer.RemoveWindowPainter srcForm.hWnd
        NavKey.NotifyFormUnloading srcForm
    End If
    
End Sub

'Given a pdImage object, generate an appropriate caption for the main PhotoDemon window.
Public Function GetWindowCaption(ByRef srcImage As pdImage, Optional ByVal appendPDInfo As Boolean = True, Optional ByVal doNotAppendImageFormat As Boolean = False) As String

    Dim captionBase As String
    Dim appendFileFormat As Boolean: appendFileFormat = False
    
    If (Not srcImage Is Nothing) Then
    
        'Start by seeing if this image has some kind of filename.  This field should always be populated by the load function,
        ' but better safe than sorry!
        If (Not srcImage.ImgStorage Is Nothing) Then
            
            If (LenB(srcImage.ImgStorage.GetEntry_String("OriginalFileName", vbNullString)) <> 0) Then
            
                'This image has a filename!  Next, check the user's preference for long or short window captions
                
                'The user prefers short captions.  Use just the filename and extension (no folders) as the base.
                If (UserPrefs.GetPref_Long("Interface", "Window Caption Length", 0) = 0) Then
                    captionBase = srcImage.ImgStorage.GetEntry_String("OriginalFileName", vbNullString)
                    appendFileFormat = True
                Else
                
                    'The user prefers long captions.  Make sure this image has such a location; if they do not, fallback
                    ' and use just the filename.
                    If (LenB(srcImage.ImgStorage.GetEntry_String("CurrentLocationOnDisk", vbNullString)) <> 0) Then
                        captionBase = srcImage.ImgStorage.GetEntry_String("CurrentLocationOnDisk", vbNullString)
                    Else
                        captionBase = srcImage.ImgStorage.GetEntry_String("OriginalFileName", vbNullString)
                        appendFileFormat = True
                    End If
                    
                End If
            
            'This image does not have a filename.  Assign it a default title.
            Else
                captionBase = g_Language.TranslateMessage("[untitled image]")
            End If
        
            'File format can be useful when working with multiple copies of the same image; PD tries to append it, as relevant
            If appendFileFormat And (Not doNotAppendImageFormat) And (srcImage.GetOriginalFileFormat() <> PDIF_UNKNOWN) Then
                captionBase = captionBase & " [" & UCase$(ImageFormats.GetExtensionFromPDIF(srcImage.GetCurrentFileFormat())) & "]"
            End If
        
        Else
            captionBase = g_Language.TranslateMessage("[untitled image]")
        End If
        
        'Check image state.  If the image is unsaved, prepend an asterisk.
        If (Not srcImage.GetSaveState(pdSE_AnySave)) Then captionBase = "* " & captionBase
        
    'If no image exists, return an empty caption; this is handled later in the function
    Else
    
    End If
    
    'Append the current PhotoDemon version number and exit
    If appendPDInfo Then
        
        If (LenB(captionBase) <> 0) Then
            GetWindowCaption = captionBase & "  -  " & Updates.GetPhotoDemonNameAndVersion()
        Else
            GetWindowCaption = Updates.GetPhotoDemonNameAndVersion()
        End If
        
        'When devs send me screenshots, it's helpful to see if they're running in the IDE or not, as this can explain some issues
        If (Not OS.IsProgramCompiled) Then GetWindowCaption = GetWindowCaption & " [IDE]"
        
    Else
        GetWindowCaption = captionBase
    End If

End Function

'Display the specified size in the main form's status bar
Public Sub DisplaySize(ByRef srcImage As pdImage)
    If (Not srcImage Is Nothing) Then FormMain.MainCanvas(0).DisplayImageSize srcImage
End Sub

'This wrapper is used in place of the standard MsgBox function.  At present it's just a wrapper around MsgBox, but
' in the future I may replace the dialog function with something custom.
Public Function PDMsgBox(ByVal pMessage As String, ByVal pButtons As VbMsgBoxStyle, ByVal pTitle As String, ParamArray ExtraText() As Variant) As VbMsgBoxResult
    
    'Before passing the message (and any optional parameters) over to the message box dialog, we first need to
    ' plug-in any dynamic elements (e.g. "%n" entries in the message with the param array contents) and apply
    ' any active language translations.
    Dim newMessage As String, newTitle As String
    newMessage = pMessage
    newTitle = pTitle
    
    If (Not g_Language Is Nothing) Then
        If (g_Language.ReadyToTranslate And g_Language.TranslationActive) Then
            newMessage = g_Language.TranslateMessage(pMessage)
            newTitle = g_Language.TranslateMessage(pTitle)
        End If
    End If
    
    'With the message freshly translated, we can plug-in any dynamic text entries
    If (UBound(ExtraText) >= LBound(ExtraText)) Then
        Dim i As Long
        For i = LBound(ExtraText) To UBound(ExtraText)
            newMessage = Replace$(newMessage, "%" & i + 1, CStr(ExtraText(i)))
        Next i
    End If
    
    'Suspend any system-wide cursors, as necessary
    Dim cursorBackup As MousePointerConstants
    cursorBackup = Screen.MousePointer
    Screen.MousePointer = vbDefault
    
    Load dialog_MsgBox
    If dialog_MsgBox.ShowDialog(newMessage, pButtons, newTitle) Then
        PDMsgBox = dialog_MsgBox.DialogResult
    
    'If the dialog failed to load for whatever reason, fall back to a default system message box
    Else
        PDMsgBox = MsgBox(newMessage, pButtons, newTitle)
    End If
    
    'Restore cursor before exiting
    Screen.MousePointer = cursorBackup
    
    Unload dialog_MsgBox
    Set dialog_MsgBox = Nothing

End Function

'This popular function is used to display a message in the main form's status bar.
' INPUTS:
' 1) the message to be displayed (mString)
' *2) any values that must be calculated at run-time, which are labeled in the message string by "%n", e.g. "Download time remaining: %1", timeRemaining
Public Sub Message(ByVal mString As String, ParamArray ExtraText() As Variant)

    Dim i As Long

    'Before doing anything else, check for a duplicate message request.  They are automatically ignored.
    Dim tmpDupeCheckString As String
    tmpDupeCheckString = mString
    
    If (UBound(ExtraText) >= LBound(ExtraText)) Then
        
        For i = LBound(ExtraText) To UBound(ExtraText)
            If Strings.StringsNotEqual(CStr(ExtraText(i)), "DONOTLOG", True) Then
                tmpDupeCheckString = Replace$(tmpDupeCheckString, "%" & CStr(i + 1), CStr(ExtraText(i)))
            End If
        Next i
        
    End If
    
    'If the message request is for a novel string (e.g. one that differs from the previous message request), display it.
    ' Otherwise, exit now.
    If Strings.StringsNotEqual(m_PrevMessage, tmpDupeCheckString, False) Then
        
        'In debug mode, mirror the message output to PD's central Debugger.  Note that this behavior can be overridden by
        ' supplying the string "DONOTLOG" as the final entry in the ParamArray.
        If UserPrefs.GenerateDebugLogs Then
        
            If (UBound(ExtraText) < LBound(ExtraText)) Then
                PDDebug.LogAction tmpDupeCheckString, PDM_User_Message
            Else
            
                'Check the last param passed.  If it's the string "DONOTLOG", do not log this entry.  (PD sometimes uses this
                ' to avoid logging useless data, like layer hover events or download updates.)
                If Strings.StringsNotEqual(CStr(ExtraText(UBound(ExtraText))), "DONOTLOG", False) Then
                    PDDebug.LogAction tmpDupeCheckString, PDM_User_Message
                End If
            
            End If
        
        End If
        
        'Cache the contents of the untranslated message, so we can check for duplicates on the next message request
        m_PrevMessage = tmpDupeCheckString
                
        Dim newString As String
        newString = mString
    
        'All messages are translatable, but we don't want to translate them if the translation object isn't ready yet.
        ' This only happens for a few messages when the program is first loaded, and at some point, I will eventually getting
        ' around to removing them entirely.
        If (Not g_Language Is Nothing) Then
            If g_Language.ReadyToTranslate Then
                If g_Language.TranslationActive Then newString = g_Language.TranslateMessage(mString)
            End If
        End If
        
        'Once the message is translated, we can add back in any optional text supplied in the ParamArray
        If (UBound(ExtraText) >= LBound(ExtraText)) Then
            For i = LBound(ExtraText) To UBound(ExtraText)
                newString = Replace$(newString, "%" & i + 1, CStr(ExtraText(i)))
            Next i
        End If
        
        'While macros are active, append a "Recording" message to help orient the user
        If (Macros.GetMacroStatus = MacroSTART) Then newString = newString & " {-" & g_Language.TranslateMessage("Recording") & "-}"
        
        'Post the message to the screen
        If (Macros.GetMacroStatus <> MacroBATCH) Then FormMain.MainCanvas(0).DisplayCanvasMessage newString
        
        'Update the global "previous message" string, so external functions can access it.
        m_LastFullMessage = newString
        
    End If
    
End Sub

'Retrieve the last full message posted to the Message() function, above.  Note that *translations and optional parameter parsing*
' have already been applied to the text.
Public Function GetLastFullMessage() As String
    GetLastFullMessage = m_LastFullMessage
End Function

'When the mouse is moved outside the primary image, clear the image coordinates display
Public Sub ClearImageCoordinatesDisplay()
    FormMain.MainCanvas(0).DisplayCanvasCoordinates 0, 0, True
End Sub

'Populate the passed combo box with options related to distort filter edge-handle options.  Also, select the specified method by default.
Public Sub PopDistortEdgeBox(ByRef cboEdges As pdDropDown, Optional ByVal defaultEdgeMethod As PD_EdgeOperator)

    cboEdges.Clear
    cboEdges.AddItem "clamp"
    cboEdges.AddItem "reflect"
    cboEdges.AddItem "wrap"
    cboEdges.AddItem "erase"
    cboEdges.AddItem "ignore"
    cboEdges.ListIndex = defaultEdgeMethod
    
End Sub

'Populate the passed button strip with options related to convolution kernel shape.  The caller can also specify which method they
' want set as the default.
Public Sub PopKernelShapeButtonStrip(ByRef srcBTS As pdButtonStrip, Optional ByVal defaultShape As PD_PixelRegionShape = PDPRS_Rectangle)
    srcBTS.AddItem "square", 0
    srcBTS.AddItem "circle", 1
    srcBTS.AddItem "diamond", 2
    srcBTS.ListIndex = defaultShape
End Sub

'Use whenever you want the user to not be allowed to interact with the primary PD window.  Make sure that call EnableUserInput(),
' below, when you are done processing!
Public Sub DisableUserInput()

    'Set the "input disabled" flag, which individual functions can use to modify their own behavior
    g_DisableUserInput = True
    
    'We also forcibly disable drag/drop whenever the interface is locked.
    g_AllowDragAndDrop = False
    
    'Suspend any active UI animations
    If PDImages.IsImageActive() Then PDImages.GetActiveImage.NotifyAnimationsAllowed False
    
    'Forcibly disable the main form
    FormMain.Enabled = False

End Sub

'Sister function to DisableUserInput(), above
Public Sub EnableUserInput()
    
    'Restore our internal input-allowance flag(s)
    g_DisableUserInput = False
    g_AllowDragAndDrop = True
    
    'Restore any active UI animations
    If PDImages.IsImageActive() Then PDImages.GetActiveImage.NotifyAnimationsAllowed True
    
    'Because PD is fully synchronous (yay, VB6), we now need to perform some message queue shenanigans.
    
    'We want to re-establish the correctness of the mouse position relative to the primary canvas, if it
    ' still lies over the canvas.  (The likelihood of movement, especially between paint ops, is high,
    ' and if we don't do this, the mouse cursor won't reflect any position updates that have occurred since
    ' the last action was started.)
    Dim tmpPoint As PointAPI, mouseMustBeFaked As Boolean
    If (GetCursorPos(tmpPoint) <> 0) Then mouseMustBeFaked = FormMain.MainCanvas(0).IsScreenCoordInsideCanvasView(tmpPoint.x, tmpPoint.y)
    
    '*WHILE THE MAIN FORM IS STILL DISABLED*, flush the keyboard/mouse queue.  (This prevents any stray
    ' keypresses or mouse events, applied while a background task was running, from suddenly firing.)
    VBHacks.DoEvents_SingleHwnd OS.ThunderMainHWnd
    
    'The line below has been a source of some frustration.  The original goal was to halt any input
    ' actions that may have triggered while PD was busy, to try and resolve some difficult-to-track-down
    ' bugs that occur when users click buttons while PD is already processing something.  This creates
    ' new problems of its own, however, like eating WM_MOUSEUP events for controls that already received
    ' WM_MOUSEDOWN (so they think they still have mouse down behavior).  I am leaving this line here
    ' for now (as of Nov 2020, TODO: revisit before v9.0 ships to ensure this is what we want) pending
    ' further testing.
    'VBHacks.PurgeInputMessages OS.ThunderMainHWnd
    
    'Re-enable the main form
    FormMain.Enabled = True
    
    'If the mouse lies over the canvas, we now want to post a "fake" mouse movement message to that window,
    ' to ensure any custom cursors are painted correctly.
    If mouseMustBeFaked Then
        If (Not g_WindowManager Is Nothing) Then g_WindowManager.GetScreenToClient FormMain.MainCanvas(0).GetCanvasViewHWnd, tmpPoint
        FormMain.MainCanvas(0).ManuallyNotifyCanvasMouse tmpPoint.x, tmpPoint.y
        Dim tmpViewportParams As PD_ViewportParams
        tmpViewportParams = Viewport.GetDefaultParamObject()
        tmpViewportParams.curPOI = poi_ReuseLast
        Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage(), FormMain.MainCanvas(0), VarPtr(tmpViewportParams)
    End If
    
End Sub

'Given a combo box, populate it with all currently supported blend modes
Public Sub PopulateBlendModeDropDown(ByRef dstCombo As pdDropDown, Optional ByVal blendIndex As PD_BlendMode = BM_Normal, Optional ByVal allowReplaceMode As Boolean = False)
    
    dstCombo.SetAutomaticRedraws False
    
    dstCombo.Clear
    
    dstCombo.AddItem "Normal", 0, True
    dstCombo.AddItem "Darken"
    dstCombo.AddItem "Multiply"
    dstCombo.AddItem "Color burn"
    dstCombo.AddItem "Linear burn", , True
    dstCombo.AddItem "Lighten"
    dstCombo.AddItem "Screen"
    dstCombo.AddItem "Color dodge"
    dstCombo.AddItem "Linear dodge", , True
    dstCombo.AddItem "Overlay"
    dstCombo.AddItem "Soft light"
    dstCombo.AddItem "Hard light"
    dstCombo.AddItem "Vivid light"
    dstCombo.AddItem "Linear light"
    dstCombo.AddItem "Pin light"
    dstCombo.AddItem "Hard mix", , True
    dstCombo.AddItem "Difference"
    dstCombo.AddItem "Exclusion"
    dstCombo.AddItem "Subtract"
    dstCombo.AddItem "Divide", , True
    dstCombo.AddItem "Hue"
    dstCombo.AddItem "Saturation"
    dstCombo.AddItem "Color"
    dstCombo.AddItem "Luminosity", , True
    dstCombo.AddItem "Grain extract"
    dstCombo.AddItem "Grain merge", , True
    dstCombo.AddItem "Erase", , True
    dstCombo.AddItem "Behind", , allowReplaceMode
    
    'Overwrite mode is only allowed in certain tools
    If allowReplaceMode Then dstCombo.AddItem "Replace"
    
    dstCombo.ListIndex = blendIndex
    
    dstCombo.SetAutomaticRedraws True, True
    
End Sub

'Given a combo box, populate it with all currently supported alpha modes
Public Sub PopulateAlphaModeDropDown(ByRef dstCombo As pdDropDown, Optional ByVal alphaIndex As PD_AlphaMode = AM_Normal)
    
    dstCombo.SetAutomaticRedraws False
    
    dstCombo.Clear
    
    dstCombo.AddItem "Normal", 0, True
    dstCombo.AddItem "Inherit"
    dstCombo.AddItem "Locked"
    
    dstCombo.ListIndex = alphaIndex
    
    dstCombo.SetAutomaticRedraws True, True
    
End Sub

Public Sub PopulateRenderingIntentDropDown(ByRef dstCombo As pdDropDown, Optional ByVal intentIndex As LCMS_RENDERING_INTENT = INTENT_PERCEPTUAL)
    
    dstCombo.SetAutomaticRedraws False
    
    dstCombo.Clear

    dstCombo.AddItem "Perceptual", 0
    dstCombo.AddItem "Relative colorimetric"
    dstCombo.AddItem "Saturation"
    dstCombo.AddItem "Absolute colorimetric"
    
    dstCombo.ListIndex = intentIndex
    
    dstCombo.SetAutomaticRedraws True, True

End Sub

Public Sub PopulateFloodFillTypes(ByRef dstCombo As pdDropDown, Optional ByVal startingIndex As PD_FloodCompare = pdfc_Color)

    dstCombo.SetAutomaticRedraws False
    
    dstCombo.Clear

    dstCombo.AddItem "color", 0
    dstCombo.AddItem "color and opacity", 1
    dstCombo.AddItem "luminance", 2, True
    dstCombo.AddItem "red only", 3
    dstCombo.AddItem "green only", 4
    dstCombo.AddItem "blue only", 5
    dstCombo.AddItem "alpha only", 6
    
    dstCombo.ListIndex = startingIndex
    
    dstCombo.SetAutomaticRedraws True, True

End Sub

'In an attempt to better serve high-DPI users, some of PD's stock UI icons are now generated at runtime.
' Note that the requested size is in PIXELS, so it is up to the caller to determine the proper size IN PIXELS of
' any requested UI elements.  This value will be automatically scaled to the current DPI, so make sure the passed
' pixel value is relevant to 100% DPI only (96 DPI).
Public Function GetRuntimeUIDIB(ByVal dibType As PD_RuntimeIcon, Optional ByVal dibSize As Long = 16, Optional ByVal dibPadding As Long = 0) As pdDIB

    'Adjust the dib size and padding to account for DPI
    dibSize = Interface.FixDPI(dibSize)
    dibPadding = Interface.FixDPI(dibPadding)

    'Create the target DIB
    Set GetRuntimeUIDIB = New pdDIB
    GetRuntimeUIDIB.CreateBlank dibSize, dibSize, 32, 0, 0
    GetRuntimeUIDIB.SetInitialAlphaPremultiplicationState True
    
    'pd2D handles rendering duties, as you'd expect.
    ' (Only the surface is created up-front, however, since not all rendering items require
    ' the other objects.)
    Dim cSurface As pd2DSurface
    Set cSurface = New pd2DSurface
    cSurface.WrapSurfaceAroundPDDIB GetRuntimeUIDIB
    cSurface.SetSurfaceAntialiasing P2_AA_HighQuality
    
    Dim cBrush As pd2DBrush, cPen As pd2DPen, cPath As pd2DPath
    Set cBrush = New pd2DBrush: Set cPen = New pd2DPen: Set cPath = New pd2DPath
    
    Dim cPoints() As PointFloat, cRadius As Single
    Dim xCenter As Single, yCenter As Single
    Dim paintColor As Long
    
    'Dynamically create the requested icon
    Select Case dibType
    
        'Red, green, and blue channel icons are all created similarly.
        Case pdri_ChannelRed, pdri_ChannelGreen, pdri_ChannelBlue
            
            If (dibType = pdri_ChannelRed) Then
                paintColor = g_Themer.GetGenericUIColor(UI_ChannelRed)
            ElseIf (dibType = pdri_ChannelGreen) Then
                paintColor = g_Themer.GetGenericUIColor(UI_ChannelGreen)
            ElseIf (dibType = pdri_ChannelBlue) Then
                paintColor = g_Themer.GetGenericUIColor(UI_ChannelBlue)
            End If
            
            'Draw a colored circle just within the bounds of the DIB
            cBrush.SetBrushColor paintColor
            PD2D.FillCircleF cSurface, cBrush, dibSize / 2, dibSize / 2, (dibSize / 2) - dibPadding
            
        'The RGB DIB is a triad of the individual RGB circles
        Case pdri_ChannelRGB
        
            'Draw the red, green, and blue circles, with slight overlap toward the middle
            Dim circleSize As Long
            circleSize = (dibSize - dibPadding) * 0.55
            
            cBrush.SetBrushOpacity 80!
            cBrush.SetBrushColor g_Themer.GetGenericUIColor(UI_ChannelBlue)
            PD2D.FillEllipseF cSurface, cBrush, dibSize - circleSize - dibPadding, dibSize - circleSize - dibPadding, circleSize, circleSize
            
            cBrush.SetBrushColor g_Themer.GetGenericUIColor(UI_ChannelGreen)
            PD2D.FillEllipseF cSurface, cBrush, dibPadding, dibSize - circleSize - dibPadding, circleSize, circleSize
            
            cBrush.SetBrushColor g_Themer.GetGenericUIColor(UI_ChannelRed)
            PD2D.FillEllipseF cSurface, cBrush, dibSize / 2 - circleSize / 2, dibPadding, circleSize, circleSize
            
        'Arrows are all drawn using the same code
        Case pdri_ArrowUp, pdri_ArrowUpR, pdri_ArrowRight, pdri_ArrowDownR, pdri_ArrowDown, pdri_ArrowDownL, pdri_ArrowLeft, pdri_ArrowUpL
        
            'Calculate button points.  (Note that these are calculated uniformly for all arrow directions.)
            Dim buttonPts() As PointFloat
            ReDim buttonPts(0 To 2) As PointFloat
            
            buttonPts(0).x = dibSize / 6
            buttonPts(0).y = (dibSize / 3) * 2
        
            buttonPts(2).x = dibSize - buttonPts(0).x
            buttonPts(2).y = buttonPts(0).y
        
            buttonPts(1).x = buttonPts(0).x + (buttonPts(2).x - buttonPts(0).x) / 2
            buttonPts(1).y = dibSize / 3
            
            'Add those points to a generic path object
            Dim tmpPath As pd2DPath
            Set tmpPath = New pd2DPath
            tmpPath.AddPolygon 3, VarPtr(buttonPts(0)), True
            
            'Rotate the path, as necessary
            If (dibType = pdri_ArrowUpR) Then
                tmpPath.RotatePathAroundItsCenter 45#
            ElseIf (dibType = pdri_ArrowRight) Then
                tmpPath.RotatePathAroundItsCenter 90#
            ElseIf (dibType = pdri_ArrowDownR) Then
                tmpPath.RotatePathAroundItsCenter 135#
            ElseIf (dibType = pdri_ArrowDown) Then
                tmpPath.RotatePathAroundItsCenter 180#
            ElseIf (dibType = pdri_ArrowDownL) Then
                tmpPath.RotatePathAroundItsCenter 225#
            ElseIf (dibType = pdri_ArrowLeft) Then
                tmpPath.RotatePathAroundItsCenter 270#
            ElseIf (dibType = pdri_ArrowUpL) Then
                tmpPath.RotatePathAroundItsCenter 315#
            End If
            
            'Render the path
            cBrush.SetBrushColor g_Themer.GetGenericUIColor(UI_GrayDark)
            PD2D.FillPath cSurface, cBrush, tmpPath
            
        'Play/pause buttons are used by animation windows
        Case pdri_Play
            
            xCenter = dibSize * 0.5
            yCenter = dibSize * 0.5
            
            cRadius = (dibSize * 0.35)
            
            ReDim cPoints(0 To 2) As PointFloat
            cPoints(0).x = xCenter + cRadius
            cPoints(0).y = yCenter
            
            PDMath.RotatePointAroundPoint cPoints(0).x, cPoints(0).y, xCenter, yCenter, (2# * PI) / 3#, cPoints(1).x, cPoints(1).y
            PDMath.RotatePointAroundPoint cPoints(0).x, cPoints(0).y, xCenter, yCenter, -1 * (2# * PI) / 3#, cPoints(2).x, cPoints(2).y
            
            Set cPath = New pd2DPath
            cPath.AddLines 3, VarPtr(cPoints(0))
            cPath.CloseCurrentFigure
            
            'Re-center the path (as the triangle will be biased rightward due to the angles used)
            cPath.TranslatePath -1! * (cRadius - (xCenter - cPoints(1).x)) * 0.5, 0!
            
            cSurface.SetSurfacePixelOffset P2_PO_Half
            cBrush.SetBrushColor g_Themer.GetGenericUIColor(UI_Accent)
            PD2D.FillPath cSurface, cBrush, cPath
                
        Case pdri_Pause
            
            ReDim cPoints(0 To 3) As PointFloat
            cPoints(0).x = (dibSize * 0.33)
            cPoints(0).y = (dibSize * 0.2)
            cPoints(1).x = cPoints(0).x
            cPoints(1).y = dibSize - cPoints(0).y
            
            cPoints(2).x = dibSize - cPoints(0).x
            cPoints(2).y = cPoints(0).y
            cPoints(3).x = cPoints(2).x
            cPoints(3).y = cPoints(1).y
            
            Set cPath = New pd2DPath
            cPath.ResetPath
            cPath.AddLines 2, VarPtr(cPoints(0))
            cPath.CloseCurrentFigure
            cPath.AddLines 2, VarPtr(cPoints(2))
            cPath.CloseCurrentFigure
    
            cPen.SetPenWidth dibSize * 0.15
            cPen.SetPenColor g_Themer.GetGenericUIColor(UI_Accent)
            cSurface.SetSurfacePixelOffset P2_PO_Half
            PD2D.DrawPath cSurface, cPen, cPath
        
    End Select
    
    Set cBrush = Nothing: Set cPen = Nothing: Set cSurface = Nothing: Set cPath = Nothing
    
    'If the user requested any padding, apply it now
    If (dibPadding > 0) Then PadDIB GetRuntimeUIDIB, dibPadding
    
End Function

'Program shutting down?  Call this function to release any interface-related resources stored by this module
Public Sub ReleaseResources()
    Set currentDialogReference = Nothing
End Sub

'This function will quickly and efficiently check the last unprocessed keypress submitted by the user.  If an ESC keypress was found,
' this function will return TRUE.  It is then up to the calling function to determine how to proceed.
Public Function UserPressedESC(Optional ByVal displayConfirmationPrompt As Boolean = True) As Boolean
    
    g_cancelCurrentAction = False
    
    'GetInputState returns a non-0 value if key or mouse events are pending.  By Microsoft's own admission,
    ' it is much faster than PeekMessage, so to keep things quick we check it before manually inspecting
    ' individual messages (see http://support.microsoft.com/kb/35605 for more details)
    If (GetInputState() <> 0) Then
    
        'Use the WM_KEYFIRST/LAST constants to explicitly request only keypress messages.  If the user has pressed multiple
        ' keys besides just ESC, this function may not operate as intended.  (Per the MSDN documentation: "...the first queued
        ' message that matches the specified filter is retrieved.")  We could technically parse all keypress messages and look
        ' for just ESC, but this would slow the function without providing any clear benefit.
        If (PeekMessage(m_tmpMsg, 0, WM_KEYFIRST, WM_KEYLAST, PM_REMOVE) <> 0) Then
        
            'Look for an ESC keypress
            If (m_tmpMsg.wParam = vbKeyEscape) Then
                
                'If the calling function requested a confirmation prompt, display it now; otherwise exit immediately.
                If displayConfirmationPrompt Then
                    Dim msgReturn As VbMsgBoxResult
                    msgReturn = PDMsgBox("Are you sure you want to cancel %1?", vbInformation Or vbYesNo, "Cancel action", Processor.GetLastProcessorID)
                    g_cancelCurrentAction = (msgReturn = vbYes)
                Else
                    g_cancelCurrentAction = True
                End If
                
            End If
            
        End If
        
    End If
    
    UserPressedESC = g_cancelCurrentAction
    
End Function

'Calculate and display the current mouse position.
' INPUTS: x and y coordinates of the mouse cursor, current form, and optionally two Double-type variables to receive the relative
' coordinates (e.g. location on the image) of the current mouse position.
Public Sub DisplayImageCoordinates(ByVal x1 As Double, ByVal y1 As Double, ByRef srcImage As pdImage, ByRef srcCanvas As pdCanvas, Optional ByRef copyX As Double, Optional ByRef copyY As Double)
    
    'This function simply wraps the relevant Drawing module function
    If Drawing.ConvertCanvasCoordsToImageCoords(srcCanvas, srcImage, x1, y1, copyX, copyY, False) Then
        
        'If an image is open, relay the new coordinates to the relevant canvas; it will handle the actual drawing internally
        If PDImages.IsImageActive() Then srcCanvas.DisplayCanvasCoordinates copyX, copyY
        
    End If
    
End Sub

'When a function does something that modifies the current image's appearance, it needs to notify this function.  This function will take
' care of the messy business of notifying various UI elements (like the image tabstrip) of the change.
Public Sub NotifyImageChanged(Optional ByVal affectedImageIndex As Long = -1)
    
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    'If an image is *not* specified, assume this is in reference to the currently active image
    If (affectedImageIndex < 0) Then affectedImageIndex = PDImages.GetActiveImageID()
    
    If PDImages.IsImageActive(affectedImageIndex) And (Macros.GetMacroStatus <> MacroBATCH) Then
        
        'Generate new taskbar and titlebar icons for the affected image
        IconsAndCursors.CreateCustomFormIcons PDImages.GetImageByID(affectedImageIndex)
        
        'Notify the image tabstrip of any changes
        FormMain.MainCanvas(0).NotifyTabstripUpdatedImage affectedImageIndex
        
    End If
    
    'This function has historically been a target for hand-optimization; uncomment the code below to
    ' report timing results.
    'pdDebug.LogAction "Time spent in NotifyImageChanged: " & Format$(VBHacks.GetTimerDifferenceNow(startTime) * 1000, "0.00") & " ms"
    
End Sub

'When a function results in an entirely new image being added to the central PD collection, it needs to notify this function.
' This function will update all relevant UI elements to match.
Public Sub NotifyImageAdded(Optional ByVal newImageIndex As Long = -1)

    'If an image is *not* specified, assume this is in reference to the currently active image
    If (newImageIndex < 0) Then newImageIndex = PDImages.GetActiveImageID()
    
    'Generate an initial set of taskbar and titlebar icons
    IconsAndCursors.CreateCustomFormIcons PDImages.GetImageByID(newImageIndex)
    
    'Notify the image tabstrip of the addition.  (It has to make quite a few internal changes to accommodate new images.)
    FormMain.MainCanvas(0).NotifyTabstripAddNewThumb newImageIndex
    
End Sub

'When a function results in an image being removed from the central PD collection, it needs to notify this function.
' This function will update all relevant UI elements to match.  The optional "redrawImmediately" parameter is useful if multiple
' images are about to be removed back-to-back; in this case, the function will not force immediate refreshes.  (However, make sure
' that when the *last* image is unloaded, redrawImmediately is set to TRUE so that appropriate redraws can take place!)
Public Sub NotifyImageRemoved(Optional ByVal oldImageIndex As Long = -1, Optional ByVal redrawImmediately As Boolean = True)

    'If an image is *not* specified, assume this is in reference to the currently active image
    If (oldImageIndex < 0) Then oldImageIndex = PDImages.GetActiveImageID()
    
    'The image tabstrip has to recalculate internal metrics whenever an image is unloaded
    FormMain.MainCanvas(0).NotifyTabstripRemoveThumb oldImageIndex, redrawImmediately
    
    'Any active UI animations also need to be suspended, as they may be tied to the removed image
    layerpanel_Navigator.NotifyStopAnimations
    
End Sub

'When a new image has been activated, call this function to apply all relevant UI changes.
Public Sub NotifyNewActiveImage(Optional ByVal newImageIndex As Long = -1)
    
    'If an image is *not* specified, assume this is in reference to the currently active image
    If (newImageIndex < 0) Then newImageIndex = PDImages.GetActiveImageID()
    
    'The toolbar must redraw itself to match the newly activated image
    FormMain.MainCanvas(0).NotifyTabstripNewActiveImage newImageIndex
    
    'A newly activated image requires a whole swath of UI changes.  Ask SyncInterfaceToCurrentImage to handle this for us.
    Interface.SyncInterfaceToCurrentImage
    
    'Ensure the list of open windows (on the main form > Window menu) is up-to-date
    Menus.UpdateSpecialMenu_WindowsOpen
    
    'Notify the new image that animations are allowed
    PDImages.GetActiveImage.NotifyAnimationsAllowed True
    
End Sub

'This function should only be used if the entire tabstrip needs to be redrawn due to some massive display-related change
' (such as changing the display color management policy).
Public Sub RequestTabstripRedraw(Optional ByVal regenerateThumbsToo As Boolean = False)
    FormMain.MainCanvas(0).NotifyTabstripTotalRedrawRequired regenerateThumbsToo
End Sub

'If a preview control won't be activated for a given dialog, call this function to display a persistent
' "no preview available" message.  (Note: for this to work, you must not attempt to supply updated preview images
' to the underlying control.  If you do, those images will obviously overwrite this warning!)
Public Sub ShowDisabledPreviewImage(ByRef dstPreview As pdFxPreviewCtl)
    
    If PDMain.IsProgramRunning() Then
    
        Dim tmpDIB As pdDIB
        Set tmpDIB = New pdDIB
        tmpDIB.CreateBlank dstPreview.GetPreviewWidth, dstPreview.GetPreviewHeight
        
        Dim notifyFont As pdFont
        Set notifyFont = New pdFont
        notifyFont.SetFontFace Fonts.GetUIFontName()
        notifyFont.SetFontSize 14
        notifyFont.SetFontColor 0
        notifyFont.SetFontBold True
        notifyFont.SetTextAlignment vbCenter
        notifyFont.CreateFontObject
        notifyFont.AttachToDC tmpDIB.GetDIBDC
    
        notifyFont.FastRenderText tmpDIB.GetDIBWidth \ 2, tmpDIB.GetDIBHeight \ 2, g_Language.TranslateMessage("preview not available")
        dstPreview.SetOriginalImage tmpDIB
        dstPreview.SetFXImage tmpDIB
        
        notifyFont.ReleaseFromDC
        
    End If
    
End Sub

'When the user's changes the UI theme, call this function to force a redraw of all visible elements.  The optional "DoEvents" parameter
' does what you expect; it yields for periodic refreshes, so the user can "see" the transformation as it occurs.
Public Sub RedrawEntireUI(Optional ByVal useDoEvents As Boolean = False)
    
    If FormMain.Visible Then
        
        FormMain.UpdateAgainstCurrentTheme
        
        'Resync the interface to redraw any remaining text and/or buttons
        Interface.SyncInterfaceToCurrentImage
        If useDoEvents Then DoEvents
        
        'Redraw any/all toolbars as well
        toolbar_Toolbox.UpdateAgainstCurrentTheme
        toolbar_Toolbox.ResetToolButtonStates
        toolbar_Options.UpdateAgainstCurrentTheme
        toolbar_Layers.UpdateAgainstCurrentTheme
        If useDoEvents Then DoEvents
        
    End If

End Sub

'Open a file manager window (defaults to Windows Explorer) and auto-select the backing file
' for the currently loaded image
Public Sub ShowActiveImageFileInExplorer()
    If PDImages.IsImageActive() Then
        If (LenB(PDImages.GetActiveImage.ImgStorage.GetEntry_String("CurrentLocationOnDisk", vbNullString)) <> 0) Then
            Files.FileSelectInExplorer PDImages.GetActiveImage.ImgStorage.GetEntry_String("CurrentLocationOnDisk", vbNullString)
        End If
    End If
End Sub
