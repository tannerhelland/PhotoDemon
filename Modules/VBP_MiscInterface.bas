Attribute VB_Name = "Interface"
'***************************************************************************
'Miscellaneous Functions Related to the User Interface
'Copyright 2001-2016 by Tanner Helland
'Created: 6/12/01
'Last updated: 06/March/16
'Last update: start implementing fine-grained UI sync caching, so sync various controls only when necessary
'
'Miscellaneous routines related to rendering and handling PhotoDemon's interface.  As the program's complexity has
' increased, so has the need for specialized handling of certain UI elements.
'
'Many of the functions in this module rely on subclassing, either directly or through things like PD's window manager.
' As such, some functions may operate differently (or not at all) while in the IDE.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************


Option Explicit

Private Declare Function GetWindowRect Lib "user32" (ByVal hndWindow As Long, ByRef lpRect As winRect) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hndWindow As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

'These constants are used to toggle visibility of display elements.
Public Const VISIBILITY_TOGGLE As Long = 0
Public Const VISIBILITY_FORCEDISPLAY As Long = 1
Public Const VISIBILITY_FORCEHIDE As Long = 2

'These values are used to remember the user's current font smoothing setting.  We try to be polite and restore
' the original setting when the application terminates.
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoW" (ByVal uiAction As Long, ByVal uiParam As Long, ByRef pvParam As Long, ByVal fWinIni As Long) As Long

Private Const SPI_GETFONTSMOOTHING As Long = &H4A
Private Const SPI_SETFONTSMOOTHING As Long = &H4B
Private Const SPI_GETFONTSMOOTHINGTYPE As Long = &H200A
Private Const SPI_SETFONTSMOOTHINGTYPE As Long = &H200B
Private Const SmoothingClearType As Long = &H2
Private Const SmoothingStandardType As Long = &H1
Private Const SmoothingNone As Long = &H0
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

Public cancelCurrentAction As Boolean

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
    PDUI_NonDestructiveFX = 15
End Enum

#If False Then
    Private Const PDUI_Save = 0, PDUI_SaveAs = 1, PDUI_Close = 2, PDUI_Undo = 3, PDUI_Redo = 4, PDUI_Copy = 5, PDUI_Paste = 6, PDUI_View = 7
    Private Const PDUI_ImageMenu = 8, PDUI_Metadata = 9, PDUI_GPSMetadata = 10, PDUI_Macros = 11, PDUI_Selections = 12
    Private Const PDUI_SelectionTransforms = 13, PDUI_LayerTools = 14, PDUI_NonDestructiveFX = 15
#End If

'If PhotoDemon enabled font smoothing where there was none previously, it will restore the original setting upon exit.  This variable
' can contain the following values:
' 0: did not have to change smoothing, as ClearType is already enabled
' 1: had to change smoothing type from Standard to ClearType
' 2: had to turn on smoothing, as it was originally turned off
Private m_ClearTypeForciblySet As Long

'PhotoDemon is designed against pixels at an expected screen resolution of 96 DPI.  Other DPI settings mess up our calculations.
' To remedy this, we dynamically modify all pixels measurements at run-time, using the current screen resolution as our guide.
Private m_DPIRatio As Double

'When a modal dialog is displayed, a reference to it is saved in this variable.  If subsequent modal dialogs are displayed (for example,
' if a tool dialog displays a color selection dialog), the previous modal dialog is given ownership over the new dialog.
Private currentDialogReference As Form
Private isSecondaryDialog As Boolean

'When the master "ShowPDDialog" function is called, it's assumed that the dialog it raises is using one of PD's command bar instances.
' The command bar will set a global "OK/Cancel" value that subsequent functions can retrieve, if they're curious.  (For example,
' a "cancel" result usually means that you can skip subsequent UI syncs, as the image's status has not changed.)
Private m_LastShowDialogResult As VbMsgBoxResult

'When a message is displayed to the user in the message portion of the status bar, we automatically cache the message's contents.
' If a subsequent request is raised with the exact same text, we can skip the whole message display process.
Private m_PrevMessage As String

'System DPI is used frequently for UI positioning calculations.  Because it's costly to constantly retrieve it via APIs, this module
' prefers to cache it only when the value changes.  Call the CacheSystemDPI() sub to update the value when appropriate, and the
' corresponding GetSystemDPI() function to retrieve the cached value.
Private m_CurrentSystemDPI As Single

'Syncing the entire program's UI to current image settings is a time-consuming process.  To try and shortcut it whenever possible,
' we track the last sync operation we performed.  If we receive a duplicate sync request, we can safely ignore it.
Private m_LastUISync_HadNoImages As PD_BOOL, m_LastUISync_HadNoLayers As PD_BOOL, m_LastUISync_HadMultipleLayers As PD_BOOL
Private m_LastUILimitingSize_Small As Single, m_LastUILimitingSize_Large As Single

'Popup dialogs present problems on non-Aero window managers, as VB's iconless approach results in the program "disappearing"
' from places like the Alt+Tab menu.  As of v7.0, we now track nested popup windows and manually handle their icon updates.
Private Const NESTED_POPUP_LIMIT As Long = 16&
Private m_PopupHWnds() As Long, m_NumOfPopupHWnds As Long
Private m_PopupIconsSmall() As Long, m_PopupIconsLarge() As Long

'Because the Interface handler is a module and not a class, like I prefer, we need to use a dedicated initialization function.
Public Sub InitializeInterfaceBackend()

    m_LastUISync_HadNoImages = PD_BOOL_UNKNOWN
    m_LastUISync_HadNoLayers = PD_BOOL_UNKNOWN
    m_LastUISync_HadMultipleLayers = PD_BOOL_UNKNOWN
    
    'vbIgnore is used internally as the "no result" value for a dialog box, as PD never provides an actual "ignore" option
    ' in its dialogs.
    m_LastShowDialogResult = vbIgnore
    
End Sub

Public Sub CacheSystemDPI(ByVal newDPI As Single)
    m_CurrentSystemDPI = newDPI
End Sub

Public Function GetSystemDPI() As Single
    GetSystemDPI = m_CurrentSystemDPI
End Function

'Previously, various PD functions had to manually enable/disable button and menu state based on their actions.  This is no longer necessary.
' Simply call this function whenever an action has done something that will potentially affect the interface, and this function will iterate
' through all potential image/interface interactions, dis/enabling buttons and menus as necessary.
'
'TODO: look at having an optional "layerID" parameter, so we can skip certain steps if only a single layer is affected by a change.
Public Sub SyncInterfaceToCurrentImage()
    
    Dim i As Long
    
    'Interface dis/enabling falls into two rough categories: stuff that changes based on the current image (e.g. Undo), and stuff that changes
    ' based on the *total* number of available images (e.g. visibility of the Effects menu).
    
    'Start by breaking our interface decisions into two broad categories: "no images are loaded" and "one or more images are loaded".
    
    'If no images are loaded, we can disable a whole swath of controls
    If (g_OpenImageCount = 0) Then
    
        'Because this set of UI changes is immutable, there is no reason to repeat it if it was the last synchronization we performed.
        If Not (m_LastUISync_HadNoImages = PD_BOOL_TRUE) Then
            SetUIMode_NoImages
            m_LastUISync_HadNoImages = PD_BOOL_TRUE
        End If
        
    'If one or more images are loaded, our job is trickier.  Some controls (such as Copy to Clipboard) are enabled no matter what,
    ' while others (Undo and Redo) are only enabled if the current image requires it.
    Else
        
        If Not (pdImages(g_CurrentImage) Is Nothing) Then
        
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
            If (pdImages(g_CurrentImage).GetNumOfLayers > 0) And Not (pdImages(g_CurrentImage).GetActiveLayer Is Nothing) Then
                
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
                If (pdImages(g_CurrentImage).GetNumOfLayers = 1) Then
                
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
                    
        End If
        
        'TODO: move selection settings into the tool handler; they're too low-level for this function
        'If a selection is active on this image, update the text boxes to match
        If pdImages(g_CurrentImage).selectionActive And (Not pdImages(g_CurrentImage).mainSelection Is Nothing) Then
            SetUIGroupState PDUI_Selections, True
            SetUIGroupState PDUI_SelectionTransforms, pdImages(g_CurrentImage).mainSelection.isTransformable
            syncTextToCurrentSelection g_CurrentImage
        Else
            SetUIGroupState PDUI_Selections, False
            SetUIGroupState PDUI_SelectionTransforms, False
        End If
            
        'Finally, synchronize various tool settings.  I've optimized this so that only the settings relative to the current tool
        ' are updated; others will be modified if/when the active tool is changed.
        Tool_Support.SyncToolOptionsUIToCurrentLayer
        
    End If
        
    'Perform a special check if 2 or more images are loaded; if that is the case, enable a few additional controls, like
    ' the "Next/Previous" Window menu items.
    If (g_OpenImageCount >= 2) Then
        FormMain.MnuWindow(5).Enabled = True
        FormMain.MnuWindow(6).Enabled = True
    Else
        FormMain.MnuWindow(5).Enabled = False
        FormMain.MnuWindow(6).Enabled = False
    End If
        
    'Redraw the layer box
    toolbar_Layers.NotifyLayerChange
        
End Sub

'Synchronize all settings whose behavior and/or appearance depends on:
' 1) At least 2+ valid layers in the current image
' 2) Different behavior and/or appearances for different layer settings
' If a UI element appears the same for ANY amount of multiple layers (e.g. "Delete Layer"), use the SetUIMode_MultipleLayers() function.
Private Sub SyncUI_MultipleLayerSettings()
    
    'Delete hidden layers is only available if one or more layers are hidden, but not ALL layers are hidden.
    If (pdImages(g_CurrentImage).GetNumOfHiddenLayers > 0) And (pdImages(g_CurrentImage).GetNumOfHiddenLayers < pdImages(g_CurrentImage).GetNumOfLayers) Then
        FormMain.MnuLayerDelete(1).Enabled = True
    Else
        FormMain.MnuLayerDelete(1).Enabled = False
    End If

    'Merge up/down are not available for layers at the top and bottom of the image
    If IsLayerAllowedToMergeAdjacent(pdImages(g_CurrentImage).GetActiveLayerIndex, False) <> -1 Then
        FormMain.MnuLayer(3).Enabled = True
    Else
        FormMain.MnuLayer(3).Enabled = False
    End If
    
    If IsLayerAllowedToMergeAdjacent(pdImages(g_CurrentImage).GetActiveLayerIndex, True) <> -1 Then
        FormMain.MnuLayer(4).Enabled = True
    Else
        FormMain.MnuLayer(4).Enabled = False
    End If
    
    'Within the order menu, certain items are disabled based on layer position.  Note that "move up" and
    ' "move to top" are both disabled for top images (similarly for bottom images and "move down/bottom"),
    ' so we can mirror the same enabled state for both options.
    If pdImages(g_CurrentImage).GetActiveLayerIndex < pdImages(g_CurrentImage).GetNumOfLayers - 1 Then
        If Not FormMain.MnuLayerOrder(0).Enabled Then
            FormMain.MnuLayerOrder(0).Enabled = True
            FormMain.MnuLayerOrder(3).Enabled = True    '"raise to top" mirrors "raise layer"
        End If
    Else
        If FormMain.MnuLayerOrder(0).Enabled Then
            FormMain.MnuLayerOrder(0).Enabled = False
            FormMain.MnuLayerOrder(3).Enabled = False
        End If
    End If
    
    If pdImages(g_CurrentImage).GetActiveLayerIndex > 0 Then
        If Not FormMain.MnuLayerOrder(1).Enabled Then
            FormMain.MnuLayerOrder(1).Enabled = True
            FormMain.MnuLayerOrder(4).Enabled = True    '"lower to bottom" mirrors "lower layer"
        End If
    Else
        If FormMain.MnuLayerOrder(1).Enabled Then
            FormMain.MnuLayerOrder(1).Enabled = False
            FormMain.MnuLayerOrder(4).Enabled = False
        End If
    End If
    
    
    'Flatten is only available if one or more layers are actually *visible*
    If pdImages(g_CurrentImage).GetNumOfVisibleLayers > 0 Then
        If Not FormMain.MnuLayer(15).Enabled Then FormMain.MnuLayer(15).Enabled = True
    Else
        FormMain.MnuLayer(15).Enabled = False
    End If
    
    'Merge visible is only available if *two* or more layers are visible
    If pdImages(g_CurrentImage).GetNumOfVisibleLayers > 1 Then
        If Not FormMain.MnuLayer(16).Enabled Then FormMain.MnuLayer(16).Enabled = True
    Else
        FormMain.MnuLayer(16).Enabled = False
    End If
    
End Sub

'Synchronize all settings whose behavior and/or appearance depends on:
' 1) At least one valid layer in the current image
' 2) Different behavior and/or appearances for different layers
' If a UI element appears the same for ANY layer (e.g. toggling visibility), use the SetUIMode_AtLeastOneLayer() function.
Private Sub SyncUI_CurrentLayerSettings()
    
    'First, determine if the current layer is using any form of non-destructive resizing
    Dim nonDestructiveResizeActive As Boolean
    nonDestructiveResizeActive = False
    If (pdImages(g_CurrentImage).GetActiveLayer.GetLayerCanvasXModifier <> 1) Then
        nonDestructiveResizeActive = True
    ElseIf (pdImages(g_CurrentImage).GetActiveLayer.GetLayerCanvasYModifier <> 1) Then
        nonDestructiveResizeActive = True
    End If
    
    'If non-destructive resizing is active, the "reset layer size" menu (and corresponding Move Tool button) must be enabled.
    If FormMain.MnuLayerSize(0).Enabled <> nonDestructiveResizeActive Then
        FormMain.MnuLayerSize(0).Enabled = nonDestructiveResizeActive
        toolpanel_MoveSize.cmdLayerMove(0).Enabled = nonDestructiveResizeActive
    End If
    
    toolpanel_MoveSize.cmdLayerMove(1).Enabled = pdImages(g_CurrentImage).GetActiveLayer.AffineTransformsActive(True)
    
    'Similar logic is used for other non-destructive affine transforms
    toolpanel_MoveSize.cmdLayerAffinePermanent.Enabled = pdImages(g_CurrentImage).GetActiveLayer.AffineTransformsActive(True)
    
    'If non-destructive FX are active on the current layer, update the non-destructive tool enablement to match
    SetUIGroupState PDUI_NonDestructiveFX, True
    
    'Layer rasterization depends on the current layer type
    FormMain.MnuLayerRasterize(0).Enabled = pdImages(g_CurrentImage).GetActiveLayer.IsLayerVector
    FormMain.MnuLayerRasterize(1).Enabled = CBool(pdImages(g_CurrentImage).GetNumOfVectorLayers > 0)
    
End Sub

'Synchronize all settings whose behavior and/or appearance depends on:
' 1) At least one valid, loaded image
' 2) Different behavior and/or appearances for different images
' If a UI element appears the same for ANY loaded image (e.g. activating the main canvas), use the SetUIMode_AtLeastOneImage() function.
Private Sub SyncUI_CurrentImageSettings()
            
    'Reset all Undo/Redo and related menus.  (Note that this also controls the SAVE BUTTON, as the image's save state is modified
    ' by PD's Undo/Redo engine.)
    SyncUndoRedoInterfaceElements True
    
    'Because Undo/Redo changes may modify menu captions, menu icons need to be reset (as they are tied to menu captions)
    ResetMenuIcons
            
    'Determine whether metadata is present, and dis/enable metadata menu items accordingly
    If Not pdImages(g_CurrentImage).imgMetadata Is Nothing Then
        SetUIGroupState PDUI_Metadata, pdImages(g_CurrentImage).imgMetadata.HasMetadata
        SetUIGroupState PDUI_GPSMetadata, pdImages(g_CurrentImage).imgMetadata.HasGPSMetadata()
    Else
        SetUIGroupState PDUI_Metadata, False
        SetUIGroupState PDUI_GPSMetadata, False
    End If
    
    'Display the image's path in the title bar.
    If Not (g_WindowManager Is Nothing) Then
        g_WindowManager.SetWindowCaptionW FormMain.hWnd, Interface.GetWindowCaption(pdImages(g_CurrentImage))
    Else
        FormMain.Caption = Interface.GetWindowCaption(pdImages(g_CurrentImage))
    End If
            
    'Display the image's size in the status bar
    If (pdImages(g_CurrentImage).Width <> 0) Then DisplaySize pdImages(g_CurrentImage)
            
    'Update the form's icon to match the current image; if a custom icon is not available, use the stock PD one
    If (pdImages(g_CurrentImage).curFormIcon32 = 0) Or (pdImages(g_CurrentImage).curFormIcon16 = 0) Then CreateCustomFormIcons pdImages(g_CurrentImage)
    ChangeAppIcons pdImages(g_CurrentImage).curFormIcon16, pdImages(g_CurrentImage).curFormIcon32
    
    'Restore the zoom value for this particular image (again, only if the form has been initialized)
    If pdImages(g_CurrentImage).Width <> 0 Then
        g_AllowViewportRendering = False
        FormMain.mainCanvas(0).GetZoomDropDownReference().ListIndex = pdImages(g_CurrentImage).currentZoomValue
        g_AllowViewportRendering = True
    End If
    
End Sub

'If an image has multiple layers, call this function to enable any UI elements that operate on multiple layers.
' Note that some multi-layer settings require certain additional criteria to be met, e.g. "Merge Visible Layers" requires at least
' two visible layers, so it must still be handled specially.  This function is only for functions that are ALWAYS available if
' multiple layers are present in an image.
Private Sub SetUIMode_MultipleLayers()
    If (Not FormMain.MnuLayer(1).Enabled) Then FormMain.MnuLayer(1).Enabled = True    'Delete layer
    If (Not FormMain.MnuLayer(5).Enabled) Then FormMain.MnuLayer(5).Enabled = True    'Order submenu
End Sub

'If an image has only one layer (e.g. a loaded JPEG), call this function to disable any UI elements that operate on multiple layers.
Private Sub SetUIMode_OnlyOneLayer()
    FormMain.MnuLayer(1).Enabled = False    'Delete layer
    FormMain.MnuLayer(3).Enabled = False    'Merge up/down
    FormMain.MnuLayer(4).Enabled = False
    FormMain.MnuLayer(5).Enabled = False    'Layer order
    FormMain.MnuLayer(15).Enabled = False   'Flatten
    FormMain.MnuLayer(16).Enabled = False   'Merge visible
End Sub

'If an image has at least one valid layer (as they always do in PD), call this function to enable relevant layer menus and controls.
Private Sub SetUIMode_AtLeastOneLayer()
    
    If (Not FormMain.MnuLayer(7).Enabled) Then
        FormMain.MnuLayer(7).Enabled = True
        FormMain.MnuLayer(8).Enabled = True
        FormMain.MnuLayer(11).Enabled = True
        FormMain.MnuLayerTransparency(3).Enabled = True     'Because all PD layers are 32-bpp, we always enable "remove transparency"
    End If
            
End Sub

'If PD ever reaches a "no layers in the current image" state, this function should be called.  (Such a state is currently unsupported, so this
' exists only as a failsafe measure.)
Private Sub SetUIMode_NoLayers()
    
    FormMain.MnuLayer(1).Enabled = False
    FormMain.MnuLayer(3).Enabled = False
    FormMain.MnuLayer(4).Enabled = False
    FormMain.MnuLayer(5).Enabled = False
    FormMain.MnuLayer(7).Enabled = False
    FormMain.MnuLayer(8).Enabled = False
    FormMain.MnuLayer(9).Enabled = False
    FormMain.MnuLayer(11).Enabled = False
    FormMain.MnuLayer(13).Enabled = False
    FormMain.MnuLayer(15).Enabled = False
    FormMain.MnuLayer(16).Enabled = False
    SetUIGroupState PDUI_NonDestructiveFX, False
    
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
    SetUIGroupState PDUI_NonDestructiveFX, False
    SetUIGroupState PDUI_Undo, False
    SetUIGroupState PDUI_Redo, False
    
    'Disable various layer-related toolbox options as well
    toolpanel_MoveSize.cmdLayerMove(0).Enabled = False
    toolpanel_MoveSize.cmdLayerMove(1).Enabled = False
    toolpanel_MoveSize.cmdLayerAffinePermanent.Enabled = False
        
    'Multiple edit menu items must also be disabled
    FormMain.MnuEdit(2).Enabled = False     'Undo history
    FormMain.MnuEdit(4).Caption = g_Language.TranslateMessage("Repeat")
    FormMain.MnuEdit(4).Enabled = False     '"Repeat..." and "Fade..."
    FormMain.MnuEdit(5).Caption = g_Language.TranslateMessage("Fade...")
    FormMain.MnuEdit(5).Enabled = False
    
    'The fade option in the primary toolbar must also go
    toolbar_Toolbox.cmdFile(FILE_FADE).AssignTooltip g_Language.TranslateMessage("Fade last action")
    toolbar_Toolbox.cmdFile(FILE_FADE).Enabled = False
    
    'Reset the main window's caption to its default PD name and version
    If Not (g_WindowManager Is Nothing) Then
        g_WindowManager.SetWindowCaptionW FormMain.hWnd, Interface.GetWindowCaption(Nothing)
    Else
        FormMain.Caption = Update_Support.GetPhotoDemonNameAndVersion()
    End If
    
    'Ask the canvas to reset itself.  Note that this also covers the status bar area and the image tabstrip, if they were
    ' previously visible.
    FormMain.mainCanvas(0).ClearCanvas
    
    'Restore the default taskbar and titlebar icons and clear the custom icon cache
    Icons_and_Cursors.ResetAppIcons
    Icons_and_Cursors.DestroyAllIcons
    
    'With all menus reset to their default values, we can now redraw all associated menu icons.
    ' (IMPORTANT: this function must be called whenever menu captions change, because icons are associated by caption.)
    ResetMenuIcons
        
    'If no images are currently open, but images were previously opened during this session, release any memory associated
    ' with those images.  This helps minimize PD's memory usage at idle.
    If g_NumOfImagesLoaded >= 1 Then
    
        'Loop through all pdImage objects and make sure they've been deactivated
        Dim i As Long
        For i = 0 To UBound(pdImages)
            If (Not pdImages(i) Is Nothing) Then
                pdImages(i).DeactivateImage
                Set pdImages(i) = Nothing
            End If
        Next i
        
        'Reset all window tracking variables
        g_NumOfImagesLoaded = 0
        g_CurrentImage = 0
        g_OpenImageCount = 0
        
    End If
    
    'Forcibly blank out the current message if no images are loaded
    Message ""
    
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
    FormMain.mainCanvas(0).AlignCanvasView
    
End Sub

'Some non-destructive actions need to synchronize *only* Undo/Redo buttons and menus (and their related counterparts, e.g. "Fade").
' To make these actions snappier, I have pulled all Undo/Redo UI sync code out of syncInterfaceToImage, and into this separate sub,
' which can be called on-demand as necessary.
'
'If the caller will be calling ResetMenuIcons() after using this function, make sure to pass the optional suspendAssociatedRedraws as TRUE
' to prevent unnecessary redraws.
'
'Finally, if no images are loaded, this function does absolutely nothing.  Refer to SetUIMode_NoImages(), above, for details.
Public Sub SyncUndoRedoInterfaceElements(Optional ByVal suspendAssociatedRedraws As Boolean = False)

    If (g_OpenImageCount <> 0) Then
    
        'Save is a bit funny, because if the image HAS been saved to file, we DISABLE the save button.
        SetUIGroupState PDUI_Save, Not pdImages(g_CurrentImage).GetSaveState(pdSE_AnySave)
        
        'Undo, Redo, Repeat and Fade are all closely related
        If Not (pdImages(g_CurrentImage).undoManager Is Nothing) Then
        
            SetUIGroupState PDUI_Undo, pdImages(g_CurrentImage).undoManager.GetUndoState
            SetUIGroupState PDUI_Redo, pdImages(g_CurrentImage).undoManager.GetRedoState
            
            'Undo history is enabled if either Undo or Redo is active
            If pdImages(g_CurrentImage).undoManager.GetUndoState Or pdImages(g_CurrentImage).undoManager.GetRedoState Then
                FormMain.MnuEdit(2).Enabled = True
            Else
                FormMain.MnuEdit(2).Enabled = False
            End If
            
            '"Edit > Repeat..." and "Edit > Fade..." are also handled by the current image's undo manager (as it
            ' maintains the list of changes applied to the image, and links to copies of previous image state DIBs).
            Dim tmpDIB As pdDIB, tmpLayerIndex As Long, tmpActionName As String
            
            'See if the "Find last relevant layer action" function in the Undo manager returns TRUE or FALSE.  If it returns TRUE,
            ' enable both Repeat and Fade, and rename each menu caption so the user knows what is being repeated/faded.
            If pdImages(g_CurrentImage).undoManager.FillDIBWithLastUndoCopy(tmpDIB, tmpLayerIndex, tmpActionName, True) Then
                FormMain.MnuEdit(4).Caption = g_Language.TranslateMessage("Repeat: %1", g_Language.TranslateMessage(tmpActionName))
                FormMain.MnuEdit(5).Caption = g_Language.TranslateMessage("Fade: %1...", g_Language.TranslateMessage(tmpActionName))
                toolbar_Toolbox.cmdFile(FILE_FADE).AssignTooltip pdImages(g_CurrentImage).undoManager.GetUndoProcessID, "Fade last action"
                
                toolbar_Toolbox.cmdFile(FILE_FADE).Enabled = True
                FormMain.MnuEdit(4).Enabled = True
                FormMain.MnuEdit(5).Enabled = True
            Else
                FormMain.MnuEdit(4).Caption = g_Language.TranslateMessage("Repeat")
                FormMain.MnuEdit(5).Caption = g_Language.TranslateMessage("Fade...")
                toolbar_Toolbox.cmdFile(FILE_FADE).AssignTooltip "Fade last action"
                
                toolbar_Toolbox.cmdFile(FILE_FADE).Enabled = False
                FormMain.MnuEdit(4).Enabled = False
                FormMain.MnuEdit(5).Enabled = False
            End If
            
            'Because these changes may modify menu captions, menu icons need to be reset (as they are tied to menu captions)
            If (Not suspendAssociatedRedraws) Then ResetMenuIcons
        
        End If
    
    End If

End Sub

'SetUIGroupState enables or disables a swath of controls related to a simple keyword (e.g. "Undo", which affects multiple menu items
' and toolbar buttons)
Public Sub SetUIGroupState(ByVal metaItem As PD_UI_Group, ByVal newState As Boolean)
    
    Dim i As Long
    
    Select Case metaItem
            
        'Save (left-hand panel button(s) AND menu item)
        Case PDUI_Save
            If FormMain.MnuFile(8).Enabled <> newState Then
                toolbar_Toolbox.cmdFile(FILE_SAVE).Enabled = newState
                FormMain.MnuFile(8).Enabled = newState      'Save
                FormMain.MnuFile(11).Enabled = newState     'Revert
            End If
            
        'Save As (menu item only).  Note that Save Copy is also tied to Save As functionality, because they use the same rules
        ' for enablement (e.g. disabled if no images are loaded, always enabled otherwise)
        Case PDUI_SaveAs
            If FormMain.MnuFile(10).Enabled <> newState Then
                toolbar_Toolbox.cmdFile(FILE_SAVEAS_LAYERS).Enabled = newState
                toolbar_Toolbox.cmdFile(FILE_SAVEAS_FLAT).Enabled = newState
                FormMain.MnuFile(9).Enabled = newState      'Save as
                FormMain.MnuFile(10).Enabled = newState     'Save copy
            End If
            
        'Close (and Close All)
        Case PDUI_Close
            If FormMain.MnuFile(5).Enabled <> newState Then
                FormMain.MnuFile(5).Enabled = newState
                FormMain.MnuFile(6).Enabled = newState
                toolbar_Toolbox.cmdFile(FILE_CLOSE).Enabled = newState
            End If
        
        'Undo (left-hand panel button AND menu item).  Undo toggles also control the "Fade last action" button,
        ' because that operates directly on previously saved Undo data.
        Case PDUI_Undo
        
            If FormMain.MnuEdit(0).Enabled <> newState Then
                toolbar_Toolbox.cmdFile(FILE_UNDO).Enabled = newState
                FormMain.MnuEdit(0).Enabled = newState
            End If
            
            'If Undo is being enabled, change the text to match the relevant action that created this Undo file
            If newState Then
                toolbar_Toolbox.cmdFile(FILE_UNDO).AssignTooltip pdImages(g_CurrentImage).undoManager.GetUndoProcessID, "Undo"
                FormMain.MnuEdit(0).Caption = g_Language.TranslateMessage("Undo:") & " " & g_Language.TranslateMessage(pdImages(g_CurrentImage).undoManager.GetUndoProcessID) & vbTab & g_Language.TranslateMessage("Ctrl") & "+Z"
            Else
                toolbar_Toolbox.cmdFile(FILE_UNDO).AssignTooltip "Undo last action"
                FormMain.MnuEdit(0).Caption = g_Language.TranslateMessage("Undo") & vbTab & g_Language.TranslateMessage("Ctrl") & "+Z"
            End If
            
            'NOTE: when changing menu text, icons must be reapplied.  Make sure to call the ResetMenuIcons() function after changing
            ' Undo/Redo enablement.
            
        'Redo (left-hand panel button AND menu item)
        Case PDUI_Redo
            If FormMain.MnuEdit(1).Enabled <> newState Then
                toolbar_Toolbox.cmdFile(FILE_REDO).Enabled = newState
                FormMain.MnuEdit(1).Enabled = newState
            End If
            
            'If Redo is being enabled, change the menu text to match the relevant action that created this Undo file
            If newState Then
                toolbar_Toolbox.cmdFile(FILE_REDO).AssignTooltip pdImages(g_CurrentImage).undoManager.GetRedoProcessID, "Redo"
                FormMain.MnuEdit(1).Caption = g_Language.TranslateMessage("Redo:") & " " & g_Language.TranslateMessage(pdImages(g_CurrentImage).undoManager.GetRedoProcessID) & vbTab & g_Language.TranslateMessage("Ctrl") & "+Y"
            Else
                toolbar_Toolbox.cmdFile(FILE_REDO).AssignTooltip "Redo previous action"
                FormMain.MnuEdit(1).Caption = g_Language.TranslateMessage("Redo") & vbTab & g_Language.TranslateMessage("Ctrl") & "+Y"
            End If
            
            'NOTE: when changing menu text, icons must be reapplied.  Make sure to call the ResetMenuIcons() function after changing
            ' Undo/Redo enablement.
            
        'Copy (menu item only)
        Case PDUI_EditCopyCut
            If FormMain.MnuEdit(7).Enabled <> newState Then FormMain.MnuEdit(7).Enabled = newState
            If FormMain.MnuEdit(8).Enabled <> newState Then FormMain.MnuEdit(8).Enabled = newState
            If FormMain.MnuEdit(9).Enabled <> newState Then FormMain.MnuEdit(9).Enabled = newState
            If FormMain.MnuEdit(10).Enabled <> newState Then FormMain.MnuEdit(10).Enabled = newState
            If FormMain.MnuEdit(12).Enabled <> newState Then FormMain.MnuEdit(12).Enabled = newState
            
        'View (top-menu level)
        Case PDUI_View
            If FormMain.MnuView.Enabled <> newState Then FormMain.MnuView.Enabled = newState
        
        'ImageOps is all Image-related menu items; it enables/disables the Image, Layer, Select, Color, and Print menus
        Case PDUI_ImageMenu
            If FormMain.MnuImageTop.Enabled <> newState Then
                FormMain.MnuSelectTop.Enabled = newState
                FormMain.MnuImageTop.Enabled = newState
                FormMain.MnuLayerTop.Enabled = newState
                FormMain.MnuAdjustmentsTop.Enabled = newState
                FormMain.MnuEffectsTop.Enabled = newState
                FormMain.MnuFile(15).Enabled = newState     'File -> Print
            End If
            
        'Macro (within the Tools menu)
        Case PDUI_Macros
            If FormMain.mnuTool(3).Enabled <> newState Then
                FormMain.mnuTool(3).Enabled = newState
                FormMain.mnuTool(4).Enabled = newState
                FormMain.mnuTool(5).Enabled = newState
            End If
        
        'Selections in general
        Case PDUI_Selections
            
            'If selections are not active, clear all the selection value textboxes
            If Not newState Then
                For i = 0 To toolpanel_Selections.tudSel.Count - 1
                    toolpanel_Selections.tudSel(i).Value = 0
                Next i
            End If
            
            'Set selection text boxes to enable only when a selection is active.  Other selection controls can remain active
            ' even without a selection present; this allows the user to set certain parameters in advance, so when they actually
            ' draw a selection, it already has the attributes they want.
            For i = 0 To toolpanel_Selections.tudSel.Count - 1
                toolpanel_Selections.tudSel(i).Enabled = newState
            Next i
            
            'En/disable all selection menu items that rely on an existing selection to operate
            If FormMain.MnuSelect(2).Enabled <> newState Then
                
                'Select none, invert selection
                FormMain.MnuSelect(1).Enabled = newState
                FormMain.MnuSelect(2).Enabled = newState
                
                'Grow/shrink/border/feather/sharpen selection
                For i = 4 To 8
                    FormMain.MnuSelect(i).Enabled = newState
                Next i
                
                'Erase selected area
                FormMain.MnuSelect(10).Enabled = newState
                
                'Save selection
                FormMain.MnuSelect(13).Enabled = newState
                
                'Export selection top-level menu
                FormMain.MnuSelect(14).Enabled = newState
                
            End If
                                    
            'Selection enabling/disabling also affects the two Crop to Selection commands (one in the Image menu, one in the Layer menu)
            If FormMain.MnuImage(9).Enabled <> newState Then FormMain.MnuImage(9).Enabled = newState
            If FormMain.MnuLayer(9).Enabled <> newState Then FormMain.MnuLayer(9).Enabled = newState
            
        'Transformable selection controls specifically
        Case PDUI_SelectionTransforms
        
            'Under certain circumstances, it is desirable to disable only the selection location boxes
            For i = 0 To toolpanel_Selections.tudSel.Count - 1
                If (Not newState) Then toolpanel_Selections.tudSel(i).Value = 0
                toolpanel_Selections.tudSel(i).Enabled = newState
            Next i
                
        'If the ExifTool plugin is not available, metadata will ALWAYS be disabled.  (We do not currently have a separate fallback for
        ' reading/browsing/writing metadata.)
        Case PDUI_Metadata
        
            If g_ExifToolEnabled Then
                If FormMain.MnuMetadata(0).Enabled <> newState Then FormMain.MnuMetadata(0).Enabled = newState
            Else
                If FormMain.MnuMetadata(0).Enabled Then FormMain.MnuMetadata(0).Enabled = False
            End If
        
        'GPS metadata is its own sub-category, and its activation is contigent upon an image having embedded GPS data
        Case PDUI_GPSMetadata
        
            If g_ExifToolEnabled Then
                If FormMain.MnuMetadata(3).Enabled <> newState Then FormMain.MnuMetadata(3).Enabled = newState
            Else
                If FormMain.MnuMetadata(3).Enabled Then FormMain.MnuMetadata(3).Enabled = False
            End If
        
        'Various layer-related tools (move, etc) are exposed on the tool options dialog.  For consistency, we disable those UI elements
        ' when no images are loaded.
        Case PDUI_LayerTools
            
            'Because we're dealing with text up/downs, we need to set hard limits relative to the current image's size.
            ' I'm currently using the "rule of three" - max/min values are the current dimensions of the image, x3.
            Dim minLayerUIValue_Width As Long, maxLayerUIValue_Width As Long
            Dim minLayerUIValue_Height As Long, maxLayerUIValue_Height As Long
            
            If newState Then
                maxLayerUIValue_Width = pdImages(g_CurrentImage).Width * 3
                maxLayerUIValue_Height = pdImages(g_CurrentImage).Height * 3
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
            Tool_Support.SetToolBusyState True
            
            'Enable/disable all UI elements as necessary
            For i = 0 To toolpanel_MoveSize.tudLayerMove.Count - 1
                If (toolpanel_MoveSize.tudLayerMove(i).Enabled <> newState) Then toolpanel_MoveSize.tudLayerMove(i).Enabled = newState
            Next i
            
            'Where relevant, also update control bounds
            If newState Then
            
                For i = 0 To toolpanel_MoveSize.tudLayerMove.Count - 1
                    
                    'Even-numbered indices correspond to width; odd-numbered to height
                    If (i Mod 2 = 0) Then
                        
                        If (toolpanel_MoveSize.tudLayerMove(i).Min <> minLayerUIValue_Width) Then
                            toolpanel_MoveSize.tudLayerMove(i).Min = minLayerUIValue_Width
                            toolpanel_MoveSize.tudLayerMove(i).Max = maxLayerUIValue_Width
                        End If
                        
                    Else
                    
                        If (toolpanel_MoveSize.tudLayerMove(i).Min <> minLayerUIValue_Height) Then
                            toolpanel_MoveSize.tudLayerMove(i).Min = minLayerUIValue_Height
                            toolpanel_MoveSize.tudLayerMove(i).Max = maxLayerUIValue_Height
                        End If
                    
                    End If
                Next i
                
            End If
            
            'Free the tool engine
            Tool_Support.SetToolBusyState False
        
        'Non-destructive FX are effects that the user can apply to a layer, without permanently modifying the layer
        Case PDUI_NonDestructiveFX
        
            If newState Then
                
                'Start by enabling all non-destructive FX controls
                For i = 0 To toolpanel_NDFX.sltQuickFix.Count - 1
                    If Not toolpanel_NDFX.sltQuickFix(i).Enabled Then toolpanel_NDFX.sltQuickFix(i).Enabled = True
                Next i
                
                'Quick fix buttons are only relevant if the current image has some non-destructive events applied
                If Not (pdImages(g_CurrentImage).GetActiveLayer Is Nothing) Then
                    
                    With toolpanel_NDFX
                    
                        .setNDFXControlState False
                        
                        If pdImages(g_CurrentImage).GetActiveLayer.GetLayerNonDestructiveFXState() Then
                            For i = 0 To .cmdQuickFix.Count - 1
                                .cmdQuickFix(i).Enabled = True
                            Next i
                        Else
                            For i = 0 To .cmdQuickFix.Count - 1
                                .cmdQuickFix(i).Enabled = False
                            Next i
                        End If
                        
                        'The index of sltQuickFix controls aligns exactly with PD's constants for non-destructive effects.  This is by design.
                        For i = 0 To .sltQuickFix.Count - 1
                            .sltQuickFix(i).Value = pdImages(g_CurrentImage).GetActiveLayer.GetLayerNonDestructiveFXValue(i)
                        Next i
                        
                        .setNDFXControlState True
                        
                    End With
                    
                Else
                    For i = 0 To toolpanel_NDFX.sltQuickFix.Count - 1
                        toolpanel_NDFX.sltQuickFix(i).Value = 0
                    Next i
                End If
                
            Else
                
                'Disable all non-destructive FX controls
                For i = 0 To toolpanel_NDFX.sltQuickFix.Count - 1
                    If toolpanel_NDFX.sltQuickFix(i).Enabled Then toolpanel_NDFX.sltQuickFix(i).Enabled = False
                Next i
                
                For i = 0 To toolpanel_NDFX.cmdQuickFix.Count - 1
                    If toolpanel_NDFX.cmdQuickFix(i).Enabled Then toolpanel_NDFX.cmdQuickFix(i).Enabled = False
                Next i
                
            End If
            
    End Select
    
End Sub


'For best results, any modal form should be shown via this function.  This function will automatically center the form over the main window,
' while also properly assigning ownership so that the dialog is truly on top of any active windows.  It also handles deactivation of
' other windows (to prevent click-through), and dynamic top-most behavior to ensure that the program doesn't steal focus if the user switches
' to another program while a modal dialog is active.
Public Sub ShowPDDialog(ByRef dialogModality As FormShowConstants, ByRef dialogForm As Form, Optional ByVal doNotUnload As Boolean = False)

    On Error GoTo showPDDialogError
    
    g_ModalDialogActive = True
    
    'Reset our "last dialog result" tracker.  (We use "ignore" as the "default" value, as it's a value PD never utilizes internally.)
    m_LastShowDialogResult = vbIgnore
    
    'Start by loading the form and hiding it
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
    Dim dialogHwnd As Long
    dialogHwnd = dialogForm.hWnd
    
    'Get the rect of the main form, which we will use to calculate a center position
    Dim ownerRect As winRect
    GetWindowRect FormMain.hWnd, ownerRect
    
    'Determine the center of that rect
    Dim centerX As Long, centerY As Long
    centerX = ownerRect.x1 + (ownerRect.x2 - ownerRect.x1) \ 2
    centerY = ownerRect.y1 + (ownerRect.y2 - ownerRect.y1) \ 2
    
    'Get the rect of the child dialog
    Dim dialogRect As winRect
    GetWindowRect dialogHwnd, dialogRect
    
    'Determine an upper-left point for the dialog based on its size
    Dim newLeft As Long, newTop As Long
    newLeft = centerX - ((dialogRect.x2 - dialogRect.x1) \ 2)
    newTop = centerY - ((dialogRect.y2 - dialogRect.y1) \ 2)
    
    'If this position results in the dialog sitting off-screen, move it so that its bottom-right corner is always on-screen.
    ' (All PD dialogs have bottom-right OK/Cancel buttons, so that's the most important part of the dialog to show.)
    If newLeft + (dialogRect.x2 - dialogRect.x1) > g_Displays.GetDesktopRight Then newLeft = g_Displays.GetDesktopRight - (dialogRect.x2 - dialogRect.x1)
    If newTop + (dialogRect.y2 - dialogRect.y1) > g_Displays.GetDesktopBottom Then newTop = g_Displays.GetDesktopBottom - (dialogRect.y2 - dialogRect.y1)
    
    'Move the dialog into place, but do not repaint it (that will be handled in a moment by the .Show event)
    MoveWindow dialogHwnd, newLeft, newTop, dialogRect.x2 - dialogRect.x1, dialogRect.y2 - dialogRect.y1, 0
    
    'Mirror the current run-time window icons to the dialog; this allows the icons to appear in places like Alt+Tab
    ' on older OSes, even though a toolbox window has focus.
    Interface.FixPopupWindow dialogHwnd, True
    
    'Use VB to actually display the dialog.  Note that the sub will pause here until the form is closed.
    dialogForm.Show dialogModality, FormMain
    
    'Now that the dialog has finished, we must replace the windows icons with its original ones - otherwise, VB will mistakenly
    ' unload our custom icons with the window!
    Interface.FixPopupWindow dialogHwnd, False
    
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
    
    g_ModalDialogActive = False
    
    Exit Sub
    
'For reasons I can't yet ascertain, this function will sometimes fail, claiming that a modal window is already active.  If that happens,
' we can just exit.
showPDDialogError:

    g_ModalDialogActive = False

End Sub

'Any commandbar-based dialog will automatically notify us of its "OK" or "Cancel" result; subsequent functions can check this return
' via GetLastShowDialogResult(), below.
Public Sub NotifyShowDialogResult(ByVal msgResult As VbMsgBoxResult, Optional ByVal nonStandardDialogSource As Boolean = False)
    
    'Only store the result if the dialog was initiated via ShowPDDialog, above
    If g_ModalDialogActive Or nonStandardDialogSource Then m_LastShowDialogResult = msgResult
    
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
Public Sub FixPopupWindow(ByVal targetHwnd As Long, Optional ByVal windowIsLoading As Boolean = False)

    If windowIsLoading Then
        
        'We could dynamically resize our tracking collection to precisely match the number of open windows, but this would only
        ' save us a few bytes.  Since we know we'll never exceed NESTED_POPUP_LIMIT, we just default to the max size off the bat.
        If (Not VB_Hacks.IsArrayInitialized(m_PopupHWnds)) Then
            m_NumOfPopupHWnds = 0
            ReDim m_PopupHWnds(0 To NESTED_POPUP_LIMIT - 1) As Long
            ReDim m_PopupIconsSmall(0 To NESTED_POPUP_LIMIT - 1) As Long
            ReDim m_PopupIconsLarge(0 To NESTED_POPUP_LIMIT - 1) As Long
        End If
        
        'We don't actually need to store the target window's hWnd (at present) but we grab it "just in case" future enhancements
        ' require it.
        m_PopupHWnds(m_NumOfPopupHWnds) = targetHwnd
        
        'We can't guarantee that Aero is active on Win 7 and earlier, so we must jump through some extra hoops to make
        ' sure the popup window appears inside any raised Alt+Tab dialogs
        If (Not g_IsWin8OrLater) Then g_WindowManager.ForceWindowAppearInAltTab targetHwnd, True
        
        'While here, cache the window's current icons.  (VB may insert its own default icons for some window types.)
        ' When the dialog is closed, we will restore these icons to avoid leaking any of PD's custom icons.
        Icons_and_Cursors.MirrorCurrentIconsToWindow targetHwnd, True, m_PopupIconsSmall(m_NumOfPopupHWnds), m_PopupIconsLarge(m_NumOfPopupHWnds)
        m_NumOfPopupHWnds = m_NumOfPopupHWnds + 1
        
    Else
    
        m_NumOfPopupHWnds = m_NumOfPopupHWnds - 1
        If (m_NumOfPopupHWnds >= 0) Then
            Icons_and_Cursors.ChangeWindowIcon m_PopupHWnds(m_NumOfPopupHWnds), m_PopupIconsSmall(m_NumOfPopupHWnds), m_PopupIconsLarge(m_NumOfPopupHWnds)
        Else
            m_NumOfPopupHWnds = 0
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "WARNING!  Interface.FixPopupWindow() has somehow unloaded more windows than it's loaded."
            #End If
        End If
    
    End If

End Sub

'Return the system keyboard delay, in seconds.  This isn't an exact science because the delay is actually hardware dependent
' (e.g. the system returns a value from 0 to 3), but we can use a "good enough" approximation.
Public Function GetKeyboardDelay() As Double
    Dim keyDelayIndex As Long
    SystemParametersInfo SPI_GETKEYBOARDDELAY, 0, keyDelayIndex, 0
    GetKeyboardDelay = (keyDelayIndex + 1) * 0.25
End Function

'Return the system keyboard repeat rate, in seconds.  This isn't an exact science because the delay is actually hardware dependent
' (e.g. the system returns a value from 0 to 31), but we can use a "good enough" approximation.
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
        FormMain.MnuWindowTabstrip(i).Checked = CBool(i = curMenuIndex)
    Next i
    
    'Write the preference out to file, then notify the canvas of the change
    If (Not suppressPrefUpdate) Then g_UserPreferences.SetPref_Long "Core", "Image Tabstrip Alignment", CLng(newAlignment)
    FormMain.mainCanvas(0).NotifyImageStripAlignment newAlignment
    
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
    If (Not suppressPrefUpdate) Then g_UserPreferences.SetPref_Long "Core", "Image Tabstrip Visibility", newSetting
    FormMain.mainCanvas(0).NotifyImageStripVisibilityMode newSetting

End Sub

Public Function FixDPI(ByVal pxMeasurement As Long) As Long

    'The first time this function is called, m_DPIRatio will be 0.  Calculate it.
    If m_DPIRatio = 0# Then
    
        'There are 1440 twips in one inch.  (Twips are resolution-independent.)  Use that knowledge to calculate DPI.
        m_DPIRatio = 1440 / TwipsPerPixelXFix
        
        'FYI: if the screen resolution is 96 dpi, this function will return the original pixel measurement, due to
        ' this calculation.
        m_DPIRatio = m_DPIRatio / 96
    
    End If
    
    FixDPI = CLng(m_DPIRatio * CDbl(pxMeasurement))
    
End Function

Public Function FixDPIFloat(ByVal pxMeasurement As Long) As Double

    'The first time this function is called, m_DPIRatio will be 0.  Calculate it.
    If m_DPIRatio = 0# Then
    
        'There are 1440 twips in one inch.  (Twips are resolution-independent.)  Use that knowledge to calculate DPI.
        m_DPIRatio = 1440 / TwipsPerPixelXFix
        
        'FYI: if the screen resolution is 96 dpi, this function will return the original pixel measurement, due to
        ' this calculation.
        m_DPIRatio = m_DPIRatio / 96
    
    End If
    
    FixDPIFloat = m_DPIRatio * CDbl(pxMeasurement)
    
End Function

'Fun fact: there are 15 twips per pixel at 96 DPI.  Not fun fact: at 200% DPI (e.g. 192 DPI), VB's internal
' TwipsPerPixelXFix will return 7, when actually we need the value 7.5.  This causes problems when resizing
' certain controls (like SmartCheckBox) because the size will actually come up short due to rounding errors!
' So whenever TwipsPerPixelXFix/Y is required, use these functions instead.
Public Function TwipsPerPixelXFix() As Double
    
    If m_CurrentSystemDPI = 0 Then
    
        If Screen.TwipsPerPixelX = 7 Then
            TwipsPerPixelXFix = 7.5
        Else
            TwipsPerPixelXFix = Screen.TwipsPerPixelX
        End If
        
    Else
        TwipsPerPixelXFix = 15# / m_CurrentSystemDPI
    End If

End Function

Public Function TwipsPerPixelYFix() As Double
    
    If m_CurrentSystemDPI = 0 Then
    
        If Screen.TwipsPerPixelY = 7 Then
            TwipsPerPixelYFix = 7.5
        Else
            TwipsPerPixelYFix = Screen.TwipsPerPixelY
        End If
        
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
    
    If Len(descriptionText) > 0 Then
        FormWait.lblWaitDescription.Caption = descriptionText
        FormWait.lblWaitDescription.Visible = True
    Else
        FormWait.lblWaitDescription.Visible = False
    End If
    
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
Public Sub ApplyThemeAndTranslations(ByRef dstForm As Form, Optional ByVal useDoEvents As Boolean = False)
    
    'Some forms call this function during the load step, meaning they will be triggered during compilation; avoid this
    If Not g_IsProgramRunning Then Exit Sub
    
    'FORM STEP 1: apply any form-level changes (like backcolor), as child controls may pull this automatically
    dstForm.BackColor = Colors.GetRGBLongFromHex(g_Themer.LookUpColor("Default", "Background"))
    dstForm.MouseIcon = LoadPicture("")
    dstForm.MousePointer = 0
    
    Dim isPDControl As Boolean, isControlEnabled As Boolean
    
    'FORM STEP 2: Enumerate through every control on the form and apply theming on a per-control basis.
    Dim eControl As Control
    
    For Each eControl In dstForm.Controls
        
        '*******************************************
        ' NOTE: some of these steps are based on old code, and will shortly be dying.  There are only a few remaining places in PD
        '  where traditional VB controls are used, and I am actively replacing them as time allows.
        
        'STEP 1: give all clickable controls a hand icon instead of the default pointer.
        ' (Note: this code sets all command buttons, scroll bars, option buttons, check boxes, list boxes, combo boxes, and file/directory/drive boxes to use the system hand cursor)
        If (TypeOf eControl Is PictureBox) Then
            SetArrowCursor eControl
        Else
            If ((TypeOf eControl Is HScrollBar) Or (TypeOf eControl Is VScrollBar) Or (TypeOf eControl Is ListBox) Or (TypeOf eControl Is ComboBox) Or (TypeOf eControl Is FileListBox) Or (TypeOf eControl Is DirListBox) Or (TypeOf eControl Is DriveListBox)) Then
                SetHandCursor eControl
            End If
        End If
        
        'STEP 2: if the current system is Vista or later, and the user has requested modern typefaces via Edit -> Preferences,
        ' redraw all control fonts using Segoe UI.
        If ((TypeOf eControl Is TextBox) Or (TypeOf eControl Is ListBox) Or (TypeOf eControl Is ComboBox) Or (TypeOf eControl Is FileListBox) Or (TypeOf eControl Is DirListBox) Or (TypeOf eControl Is DriveListBox) Or (TypeOf eControl Is Label)) And (Not TypeOf eControl Is PictureBox) Then
            eControl.fontName = g_InterfaceFont
        End If
        
        'STEP 3: make common control drop-down boxes display their full drop-down contents, without a scroll bar.
        '         (Note: this behavior requires a manifest, so it's entirely useless inside the IDE.)
        '         (Also, once all combo boxes are replaced with PD's dedicated replacement, this line can be removed.)
        If (TypeOf eControl Is ComboBox) Then SendMessage eControl.hWnd, CB_SETMINVISIBLE, CLng(eControl.ListCount), ByVal 0&
        
        ' TODO 6.8: remove these steps once and for all
        '*******************************************
        
        'All of PhotoDemon's custom UI controls implement an UpdateAgainstCurrentTheme function.  This function updates two things:
        ' 1) The control's visual appearance (to reflect any changes to visual themes)
        ' 2) Updating any translatable text against the current translation
        
        isPDControl = False
        
        'These controls are fully compatible with PD's theming and translation engines:
        If (TypeOf eControl Is pdButtonStrip) Or (TypeOf eControl Is pdButtonStripVertical) Then
            isPDControl = True
        ElseIf (TypeOf eControl Is pdLabel) Or (TypeOf eControl Is pdHyperlink) Then
            isPDControl = True
        ElseIf (TypeOf eControl Is pdColorSelector) Or (TypeOf eControl Is pdBrushSelector) Then
            isPDControl = True
        ElseIf (TypeOf eControl Is pdGradientSelector) Or (TypeOf eControl Is pdPenSelector) Then
            isPDControl = True
        ElseIf (TypeOf eControl Is pdButton) Or (TypeOf eControl Is pdButtonToolbox) Then
            isPDControl = True
        ElseIf (TypeOf eControl Is pdScrollBar) Then
            isPDControl = True
        ElseIf (TypeOf eControl Is pdTextBox) Or (TypeOf eControl Is pdSpinner) Then
            isPDControl = True
        ElseIf (TypeOf eControl Is pdSlider) Or (TypeOf eControl Is pdSliderStandalone) Then
            isPDControl = True
        ElseIf (TypeOf eControl Is pdTitle) Or (TypeOf eControl Is pdMetadataExport) Then
            isPDControl = True
        ElseIf (TypeOf eControl Is pdFxPreviewCtl) Or (TypeOf eControl Is pdPreview) Then
            isPDControl = True
        ElseIf (TypeOf eControl Is pdCheckBox) Or (TypeOf eControl Is pdRadioButton) Then
            isPDControl = True
        ElseIf (TypeOf eControl Is pdColorVariants) Or (TypeOf eControl Is pdColorWheel) Then
            isPDControl = True
        ElseIf (TypeOf eControl Is pdNavigator) Or (TypeOf eControl Is pdNavigatorInner) Then
            isPDControl = True
        ElseIf (TypeOf eControl Is pdCommandBar) Or (TypeOf eControl Is pdCommandBarMini) Then
            isPDControl = True
        ElseIf (TypeOf eControl Is pdResize) Then
            isPDControl = True
        ElseIf (TypeOf eControl Is pdCanvas) Or (TypeOf eControl Is pdCanvasView) Then
            isPDControl = True
        ElseIf (TypeOf eControl Is pdListBox) Or (TypeOf eControl Is pdListBoxView) Then
            isPDControl = True
        ElseIf (TypeOf eControl Is pdListBoxOD) Or (TypeOf eControl Is pdListBoxViewOD) Then
            isPDControl = True
        ElseIf (TypeOf eControl Is pdDropDown) Then
            isPDControl = True
        End If
        
        'Combo boxes are hopelessly broken in their current incarnation.  They will shortly be rewritten, so please ignore
        ' their problematic behavior at present.
        If (TypeOf eControl Is pdComboBox_Font) Or (TypeOf eControl Is pdComboBox_Hatch) Then isPDControl = True
        
        'Disabled controls will ignore any function calls, so we must manually enable disabled controls prior to theming them
        If isPDControl Then
            isControlEnabled = eControl.Enabled
            If Not isControlEnabled Then eControl.Enabled = True
            eControl.UpdateAgainstCurrentTheme
            If Not isControlEnabled Then eControl.Enabled = False
        End If
        
        'While we're here, forcibly remove TabStops from each picture box.  They should never receive focus, but I often forget
        ' to change this at design-time.
        If (TypeOf eControl Is PictureBox) Then eControl.TabStop = False
        
        'Optionally, DoEvents can be called after each change.  This slows the process, but it allows external progress
        ' bars to be automatically refreshed.  We use this when the user actively changes the visual theme and/or language,
        ' as it allows the user to "see" the changes appear on the main PD window.
        ' (TODO 6.8: investigate where else this is used, if anywhere, and consider removal.)
        If useDoEvents Then DoEvents
        
    Next
    
    'FORM STEP 3: translate the form (and all controls on it)
    ' Note that this step is not as relevant as it used to be, because all PD controls apply their own translations if/when necessary
    ' during the above eControl.UpdateAgainstCurrentTheme step.  This translation step only handles the form caption (which must be
    ' set specially), and some other oddities like menus, which have not been replaced yet.
    ' TODO 6.8: once all controls are migrated, consider killing this step entirely, and moving the specialized translation bits here.
    If g_Language.TranslationActive And dstForm.Enabled Then
        g_Language.ApplyTranslations dstForm, useDoEvents
    End If
    
    'FORM STEP 4: force a refresh to ensure our changes are immediately visible
    If dstForm.Name <> "FormMain" Then
        dstForm.Refresh
    Else
        'The main from is a bit different - if it has been translated or changed, it needs menu icons reassigned, because they are
        ' inadvertently dropped when the menu captions change.
        If FormMain.Visible Then ApplyAllMenuIcons
    End If
    
End Sub

'Used to enable font smoothing if currently disabled.
Public Sub HandleClearType(ByVal startingProgram As Boolean)
    
    'At start-up, activate ClearType.  At shutdown, restore the original setting (as necessary).
    If startingProgram Then
    
        m_ClearTypeForciblySet = 0
    
        'Get current font smoothing setting
        Dim pv As Long
        SystemParametersInfo SPI_GETFONTSMOOTHING, 0, pv, 0
        
        'If font smoothing is disabled, mark it
        If pv = 0 Then m_ClearTypeForciblySet = 2
        
        'If font smoothing is enabled but set to Standard instead of ClearType, mark it
        If pv <> 0 Then
            SystemParametersInfo SPI_GETFONTSMOOTHINGTYPE, 0, pv, 0
            If pv = SmoothingStandardType Then m_ClearTypeForciblySet = 1
        End If
        
        Select Case m_ClearTypeForciblySet
        
            'ClearType is enabled, no changes necessary
            Case 0
            
            'Standard smoothing is enabled; switch it to ClearType for the duration of the program
            Case 1
                SystemParametersInfo SPI_SETFONTSMOOTHINGTYPE, 0, ByVal SmoothingClearType, 0
                
            'No smoothing is enabled; turn it on and activate ClearType for the duration of the program
            Case 2
                SystemParametersInfo SPI_SETFONTSMOOTHING, 1, pv, 0
                SystemParametersInfo SPI_SETFONTSMOOTHINGTYPE, 0, ByVal SmoothingClearType, 0
            
        End Select
    
    Else
        
        Select Case m_ClearTypeForciblySet
        
            'ClearType was enabled, no action necessary
            Case 0
            
            'Standard smoothing was enabled; restore it now
            Case 1
                SystemParametersInfo SPI_SETFONTSMOOTHINGTYPE, 0, ByVal SmoothingStandardType, 0
                
            'No smoothing was enabled; restore that setting now
            Case 2
                SystemParametersInfo SPI_SETFONTSMOOTHING, 0, pv, 0
                SystemParametersInfo SPI_SETFONTSMOOTHINGTYPE, 0, ByVal SmoothingNone, 0
        
        End Select
    
    End If
    
End Sub

'When a themed form is unloaded, it may be desirable to release certain changes made to it - or in our case, unsubclass it.
' This function should be called when any themed form is unloaded.
Public Sub ReleaseFormTheming(ByRef tForm As Object)
    Set tForm = Nothing
End Sub

'Given a pdImage object, generate an appropriate caption for the main PhotoDemon window.
Private Function GetWindowCaption(ByRef srcImage As pdImage) As String

    Dim captionBase As String
    Dim appendFileFormat As Boolean: appendFileFormat = False
    
    If (Not (srcImage Is Nothing)) Then
    
        'Start by seeing if this image has some kind of filename.  This field should always be populated by the load function,
        ' but better safe than sorry!
        If Len(srcImage.imgStorage.GetEntry_String("OriginalFileName", vbNullString)) <> 0 Then
        
            'This image has a filename!  Next, check the user's preference for long or short window captions
            
            'The user prefers short captions.  Use just the filename and extension (no folders ) as the base.
            If g_UserPreferences.GetPref_Long("Interface", "Window Caption Length", 0) = 0 Then
                captionBase = srcImage.imgStorage.GetEntry_String("OriginalFileName", vbNullString)
                appendFileFormat = True
            Else
            
                'The user prefers long captions.  Make sure this image has such a location; if they do not, fallback
                ' and use just the filename.
                If Len(srcImage.imgStorage.GetEntry_String("CurrentLocationOnDisk", vbNullString)) <> 0 Then
                    captionBase = srcImage.imgStorage.GetEntry_String("CurrentLocationOnDisk", vbNullString)
                Else
                    captionBase = srcImage.imgStorage.GetEntry_String("OriginalFileName", vbNullString)
                    appendFileFormat = True
                End If
                
            End If
        
        'This image does not have a filename.  Assign it a default title.
        Else
            captionBase = g_Language.TranslateMessage("[untitled image]")
        End If
        
        'File format can be useful when working with multiple copies of the same image; PD tries to append it, as relevant
        If appendFileFormat And (Len(srcImage.imgStorage.GetEntry_String("OriginalFileExtension", vbNullString)) <> 0) Then
            captionBase = captionBase & " [" & UCase(srcImage.imgStorage.GetEntry_String("OriginalFileExtension", vbNullString)) & "]"
        End If
        
    Else
    
    End If
    
    'Append the current PhotoDemon version number and exit
    If (Len(captionBase) <> 0) Then
        GetWindowCaption = captionBase & "  -  " & Update_Support.GetPhotoDemonNameAndVersion()
    Else
        GetWindowCaption = Update_Support.GetPhotoDemonNameAndVersion()
    End If
    
    'When devs send me screenshots, it's helpful to see if they're running in the IDE or not, as this can explain some issues
    If (Not g_IsProgramCompiled) Then GetWindowCaption = GetWindowCaption & " [IDE]"

End Function

'Display the specified size in the main form's status bar
Public Sub DisplaySize(ByRef srcImage As pdImage)
    
    If Not (srcImage Is Nothing) Then
        
        FormMain.mainCanvas(0).DisplayImageSize srcImage
        
        'Size is only displayed when it is changed, so if any controls have a maximum value linked to the size of the image,
        ' now is an excellent time to update them.
        Dim newLimitingSize_Small As Single, newLimitingSize_Large As Single
        
        If (srcImage.Width < srcImage.Height) Then
            newLimitingSize_Small = srcImage.Width
            newLimitingSize_Large = srcImage.Height
        Else
            newLimitingSize_Small = srcImage.Height
            newLimitingSize_Large = srcImage.Width
        End If
        
        If (m_LastUILimitingSize_Small <> newLimitingSize_Small) Or (m_LastUILimitingSize_Large <> newLimitingSize_Large) Then
            
            m_LastUILimitingSize_Small = newLimitingSize_Small
            m_LastUILimitingSize_Large = newLimitingSize_Large
            
            'Certain selection tools are size-limited by the current image; update those now!
            toolpanel_Selections.sltCornerRounding.Max = m_LastUILimitingSize_Small
            toolpanel_Selections.sltSelectionLineWidth.Max = m_LastUILimitingSize_Large
        
            Dim i As Long
            For i = 0 To toolpanel_Selections.sltSelectionBorder.Count - 1
                toolpanel_Selections.sltSelectionBorder(i).Max = m_LastUILimitingSize_Small
            Next i
            
        End If
    
    End If
    
End Sub

'This wrapper is used in place of the standard MsgBox function.  At present it's just a wrapper around MsgBox, but
' in the future I may replace the dialog function with something custom.
Public Function PDMsgBox(ByVal pMessage As String, ByVal pButtons As VbMsgBoxStyle, ByVal pTitle As String, ParamArray ExtraText() As Variant) As VbMsgBoxResult

    Dim newMessage As String, newTitle As String
    newMessage = pMessage
    newTitle = pTitle

    'All messages are translatable, but we don't want to translate them if the translation object isn't ready yet
    If (Not (g_Language Is Nothing)) Then
        If g_Language.ReadyToTranslate Then
            If g_Language.TranslationActive Then
                newMessage = g_Language.TranslateMessage(pMessage)
                newTitle = g_Language.TranslateMessage(pTitle)
            End If
        End If
    End If
    
    'Once the message is translated, we can add back in any optional parameters
    If UBound(ExtraText) >= LBound(ExtraText) Then
    
        Dim i As Long
        For i = LBound(ExtraText) To UBound(ExtraText)
            newMessage = Replace$(newMessage, "%" & i + 1, CStr(ExtraText(i)))
        Next i
    
    End If

    PDMsgBox = MsgBox(newMessage, pButtons, newTitle)

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
    
    If UBound(ExtraText) >= LBound(ExtraText) Then
        
        For i = LBound(ExtraText) To UBound(ExtraText)
            If StrComp(UCase$(ExtraText(i)), "DONOTLOG", vbBinaryCompare) <> 0 Then
                tmpDupeCheckString = Replace$(tmpDupeCheckString, "%" & CStr(i + 1), CStr(ExtraText(i)))
            End If
        Next i
        
    End If
    
    'If the message request is for a novel string (e.g. one that differs from the previous message request), display it.
    ' Otherwise, exit now.
    If StrComp(m_PrevMessage, tmpDupeCheckString, vbBinaryCompare) <> 0 Then
        
        'In debug mode, mirror the message output to PD's central Debugger.  Note that this behavior can be overridden by
        ' supplying the string "DONOTLOG" as the final entry in the ParamArray.
        #If DEBUGMODE = 1 Then
            If UBound(ExtraText) < LBound(ExtraText) Then
                pdDebug.LogAction tmpDupeCheckString, PDM_USER_MESSAGE
            Else
            
                'Check the last param passed.  If it's the string "DONOTLOG", do not log this entry.  (PD sometimes uses this
                ' to avoid logging useless data, like layer hover events or download updates.)
                If StrComp(UCase$(CStr(ExtraText(UBound(ExtraText)))), "DONOTLOG", vbBinaryCompare) <> 0 Then
                    pdDebug.LogAction tmpDupeCheckString, PDM_USER_MESSAGE
                End If
            
            End If
        #End If
        
        'Cache the contents of the untranslated message, so we can check for duplicates on the next message request
        m_PrevMessage = tmpDupeCheckString
                
        Dim newString As String
        newString = mString
    
        'All messages are translatable, but we don't want to translate them if the translation object isn't ready yet.
        ' This only happens for a few messages when the program is first loaded, and at some point, I will eventually getting
        ' around to removing them entirely.
        If (Not (g_Language Is Nothing)) Then
            If g_Language.ReadyToTranslate Then
                If g_Language.TranslationActive Then newString = g_Language.TranslateMessage(mString)
            End If
        End If
        
        'Once the message is translated, we can add back in any optional text supplied in the ParamArray
        If UBound(ExtraText) >= LBound(ExtraText) Then
        
            For i = LBound(ExtraText) To UBound(ExtraText)
                newString = Replace$(newString, "%" & i + 1, CStr(ExtraText(i)))
            Next i
        
        End If
        
        'While macros are active, append a "Recording" message to help orient the user
        If MacroStatus = MacroSTART Then newString = newString & " {-" & g_Language.TranslateMessage("Recording") & "-}"
        
        'Post the message to the screen
        If (MacroStatus <> MacroBATCH) Then FormMain.mainCanvas(0).DisplayCanvasMessage newString
        
        'Update the global "previous message" string, so external functions can access it.
        g_LastPostedMessage = newString
        
    End If
    
End Sub

'When the mouse is moved outside the primary image, clear the image coordinates display
Public Sub ClearImageCoordinatesDisplay()
    FormMain.mainCanvas(0).DisplayCanvasCoordinates 0, 0, True
End Sub

'Populate the passed combo box with options related to distort filter edge-handle options.  Also, select the specified method by default.
Public Sub PopDistortEdgeBox(ByRef cboEdges As pdDropDown, Optional ByVal defaultEdgeMethod As EDGE_OPERATOR)

    cboEdges.Clear
    cboEdges.AddItem " clamp them to the nearest available pixel"
    cboEdges.AddItem " reflect them across the nearest edge"
    cboEdges.AddItem " wrap them around the image"
    cboEdges.AddItem " erase them"
    cboEdges.AddItem " ignore them"
    cboEdges.ListIndex = defaultEdgeMethod
    
End Sub

'Populate the passed button strip with options related to convolution kernel shape.  The caller can also specify which method they
' want set as the default.
Public Sub PopKernelShapeButtonStrip(ByRef srcBTS As pdButtonStrip, Optional ByVal defaultShape As PD_PIXEL_REGION_SHAPE = PDPRS_Rectangle)
    
    srcBTS.AddItem "Square", 0
    srcBTS.AddItem "Circle", 1
    srcBTS.ListIndex = defaultShape
    
End Sub

'Use whenever you want the user to not be allowed to interact with the primary PD window.  Make sure that call "enableUserInput", below,
' when you are done processing!
Public Sub DisableUserInput()

    'Set the "input disabled" flag, which individual functions can use to modify their own behavior
    g_DisableUserInput = True
    
    'We also forcibly disable drag/drop whenever the interface is locked.
    g_AllowDragAndDrop = False
    
    'Forcibly disable the main form
    FormMain.Enabled = False

End Sub

'Sister function to "disableUserInput", above
Public Sub EnableUserInput()
    
    'Start a countdown timer on the main form.  When it terminates, user input will be restored.  A timer is required because
    ' we need a certain amount of "dead time" to elapse between double-clicks on a top-level dialog (like a common dialog)
    ' which may be incorrectly passed through to the main form.  (I know, this seems like a ridiculous solution, but I tried
    ' a thousand others before settling on this.  It's the least of many evils.)
    FormMain.tmrCountdown.Enabled = True
    
    'Drag/drop allowance doesn't suffer the issue described above, so we can enable it immediately
    g_AllowDragAndDrop = True
    
    'Re-enable the main form
    FormMain.Enabled = True

End Sub

'Given a combo box, populate it with all currently supported blend modes
Public Sub PopulateBlendModeComboBox(ByRef dstCombo As pdDropDown, Optional ByVal blendIndex As LAYER_BLENDMODE = BL_NORMAL)
    
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
    dstCombo.AddItem "Grain merge"
    
    dstCombo.ListIndex = blendIndex
    
End Sub

'In an attempt to better serve high-DPI users, some of PD's stock UI icons are now generated at runtime.
' Note that the requested size is in PIXELS, so it is up to the caller to determine the proper size IN PIXELS of
' any requested UI elements.  This value will be automatically scaled to the current DPI, so make sure the passed
' pixel value is relevant to 100% DPI only (96 DPI).
Public Function GetRuntimeUIDIB(ByVal dibType As PD_RUNTIME_UI_DIB, Optional ByVal dibSize As Long = 16, Optional ByVal dibPadding As Long = 0, Optional ByVal BackColor As Long = 0) As pdDIB

    'Adjust the dib size and padding to account for DPI
    dibSize = FixDPI(dibSize)
    dibPadding = FixDPI(dibPadding)

    'Create the target DIB
    Set GetRuntimeUIDIB = New pdDIB
    GetRuntimeUIDIB.CreateBlank dibSize, dibSize, 32, BackColor, 0
    GetRuntimeUIDIB.SetInitialAlphaPremultiplicationState True
    
    Dim paintColor As Long
    
    'Dynamically create the requested icon
    Select Case dibType
    
        'Red, green, and blue channel icons are all created similarly.
        Case PDRUID_CHANNEL_RED, PDRUID_CHANNEL_GREEN, PDRUID_CHANNEL_BLUE
            
            If dibType = PDRUID_CHANNEL_RED Then
                paintColor = g_Themer.GetThemeColor(PDTC_CHANNEL_RED)
            ElseIf dibType = PDRUID_CHANNEL_GREEN Then
                paintColor = g_Themer.GetThemeColor(PDTC_CHANNEL_GREEN)
            ElseIf dibType = PDRUID_CHANNEL_BLUE Then
                paintColor = g_Themer.GetThemeColor(PDTC_CHANNEL_BLUE)
            End If
            
            'Draw a colored circle just within the bounds of the DIB
            GDI_Plus.GDIPlusFillEllipseToDC GetRuntimeUIDIB.GetDIBDC, dibPadding, dibPadding, dibSize - dibPadding * 2, dibSize - dibPadding * 2, paintColor, True
        
        'The RGB DIB is a triad of the individual RGB circles
        Case PDRUID_CHANNEL_RGB
        
            'Draw the red, green, and blue circles, with slight overlap toward the middle
            Dim circleSize As Long
            circleSize = (dibSize - dibPadding) * 0.55
            
            GDI_Plus.GDIPlusFillEllipseToDC GetRuntimeUIDIB.GetDIBDC, dibSize - circleSize - dibPadding, dibSize - circleSize - dibPadding, circleSize, circleSize, g_Themer.GetThemeColor(PDTC_CHANNEL_BLUE), True, 210
            GDI_Plus.GDIPlusFillEllipseToDC GetRuntimeUIDIB.GetDIBDC, dibPadding, dibSize - circleSize - dibPadding, circleSize, circleSize, g_Themer.GetThemeColor(PDTC_CHANNEL_GREEN), True, 210
            GDI_Plus.GDIPlusFillEllipseToDC GetRuntimeUIDIB.GetDIBDC, dibSize \ 2 - circleSize \ 2, dibPadding, circleSize, circleSize, g_Themer.GetThemeColor(PDTC_CHANNEL_RED), True, 210
    
    End Select
    
    'If the user requested any padding, apply it now
    If dibPadding > 0 Then padDIB GetRuntimeUIDIB, dibPadding
    
End Function

'New test functions to (hopefully) help address high-DPI issues where VB's internal scale properties report false values
Public Function APIWidth(ByVal srcHwnd As Long) As Long
    Dim tmpRect As winRect
    GetWindowRect srcHwnd, tmpRect
    APIWidth = tmpRect.x2 - tmpRect.x1
End Function

Public Function APIHeight(ByVal srcHwnd As Long) As Long
    Dim tmpRect As winRect
    GetWindowRect srcHwnd, tmpRect
    APIHeight = tmpRect.y2 - tmpRect.y1
End Function

'Program shutting down?  Call this function to release any interface-related resources stored by this module
Public Sub ReleaseResources()
    Set currentDialogReference = Nothing
End Sub

'Wait for (n) milliseconds, while still providing some interactivity via DoEvents.  Thank you to vbforums user "anhn" for the
' original version of this function, available here: http://www.vbforums.com/showthread.php?546633-VB6-Sleep-Function.
' Please note that his original code has been modified for use in PhotoDemon.
Public Sub PauseProgram(ByRef secsDelay As Double)
   
   Dim TimeOut   As Double
   Dim PrevTimer As Double
   
   PrevTimer = Timer
   TimeOut = PrevTimer + secsDelay
   Do While PrevTimer < TimeOut
      Sleep 2 '-- Timer is only updated every 1/128 sec
      DoEvents
      If Timer < PrevTimer Then TimeOut = TimeOut - 86400 '-- pass midnight
      PrevTimer = Timer
   Loop
   
End Sub

'This function will quickly and efficiently check the last unprocessed keypress submitted by the user.  If an ESC keypress was found,
' this function will return TRUE.  It is then up to the calling function to determine how to proceed.
Public Function UserPressedESC(Optional ByVal displayConfirmationPrompt As Boolean = True) As Boolean

    Dim tmpMsg As winMsg
    
    'GetInputState returns a non-0 value if key or mouse events are pending.  By Microsoft's own admission, it is much faster
    ' than PeekMessage, so to keep things fast we check it before manually inspecting individual messages
    ' (see http://support.microsoft.com/kb/35605 for more details)
    If GetInputState() Then
    
        'Use the WM_KEYFIRST/LAST constants to explicitly request only keypress messages.  If the user has pressed multiple
        ' keys besides just ESC, this function may not operate as intended.  (Per the MSDN documentation: "...the first queued
        ' message that matches the specified filter is retrieved.")  We could technically parse all keypress messages and look
        ' for just ESC, but this would slow the function without providing any clear benefit.
        PeekMessage tmpMsg, 0, WM_KEYFIRST, WM_KEYLAST, PM_REMOVE
        
        'ESC keypress found!
        If tmpMsg.wParam = vbKeyEscape Then
            
            'If the calling function requested a confirmation prompt, display it now; otherwise exit immediately.
            If displayConfirmationPrompt Then
                Dim msgReturn As VbMsgBoxResult
                msgReturn = PDMsgBox("Are you sure you want to cancel %1?", vbInformation + vbYesNo + vbApplicationModal, "Cancel image processing", LastProcess.Id)
                If msgReturn = vbYes Then cancelCurrentAction = True Else cancelCurrentAction = False
            Else
                cancelCurrentAction = True
            End If
            
        Else
            cancelCurrentAction = False
        End If
        
    Else
        cancelCurrentAction = False
    End If
    
    UserPressedESC = cancelCurrentAction
    
End Function

'Calculate and display the current mouse position.
' INPUTS: x and y coordinates of the mouse cursor, current form, and optionally two Double-type variables to receive the relative
' coordinates (e.g. location on the image) of the current mouse position.
Public Sub DisplayImageCoordinates(ByVal x1 As Double, ByVal y1 As Double, ByRef srcImage As pdImage, ByRef srcCanvas As pdCanvas, Optional ByRef copyX As Double, Optional ByRef copyY As Double)
    
    'This function simply wraps the relevant Drawing module function
    If Drawing.ConvertCanvasCoordsToImageCoords(srcCanvas, srcImage, x1, y1, copyX, copyY) Then
        
        'If an image is open, relay the new coordinates to the relevant canvas; it will handle the actual drawing internally
        If g_OpenImageCount > 0 Then srcCanvas.DisplayCanvasCoordinates copyX, copyY
        
    End If
    
End Sub

'When a function does something that modifies the current image's appearance, it needs to notify this function.  This function will take
' care of the messy business of notifying various UI elements (like the image tabstrip) of the change.
'
'If the change only affects a single image or layer, pass their indices; we can use them to shortcut a number of UI syncing steps.
Public Sub NotifyImageChanged(Optional ByVal affectedImageIndex As Long = -1, Optional ByVal affectedLayerID As Long = -1)
    
    'If an image is *not* specified, assume this is in reference to the currently active image
    If (affectedImageIndex < 0) Then affectedImageIndex = g_CurrentImage
    
    'Generate new taskbar and titlebar icons for the affected image
    CreateCustomFormIcons pdImages(affectedImageIndex)
    
    'Notify the image tabstrip of any changes
    FormMain.mainCanvas(0).NotifyTabstripUpdatedImage affectedImageIndex
    
End Sub

'When a function results in an entirely new image being added to the central PD collection, it needs to notify this function.
' This function will update all relevant UI elements to match.
Public Sub NotifyImageAdded(Optional ByVal newImageIndex As Long = -1)

    'If an image is *not* specified, assume this is in reference to the currently active image
    If (newImageIndex < 0) Then newImageIndex = g_CurrentImage
    
    'Generate an initial set of taskbar and titlebar icons
    Icons_and_Cursors.CreateCustomFormIcons pdImages(newImageIndex)
    
    'Notify the image tabstrip of the addition.  (It has to make quite a few internal changes to accommodate new images.)
    FormMain.mainCanvas(0).NotifyTabstripAddNewThumb newImageIndex
    
End Sub

'When a function results in an image being removed from the central PD collection, it needs to notify this function.
' This function will update all relevant UI elements to match.  The optional "redrawImmediately" parameter is useful if multiple
' images are about to be removed back-to-back; in this case, the function will not force immediate refreshes.  (However, make sure
' that when the *last* image is unloaded, redrawImmediately is set to TRUE so that appropriate redraws can take place!)
Public Sub NotifyImageRemoved(Optional ByVal oldImageIndex As Long = -1, Optional ByVal redrawImmediately As Boolean = True)

    'If an image is *not* specified, assume this is in reference to the currently active image
    If (oldImageIndex < 0) Then oldImageIndex = g_CurrentImage
    
    'The image tabstrip has to recalculate internal metrics whenever an image is unloaded
    FormMain.mainCanvas(0).NotifyTabstripRemoveThumb oldImageIndex, redrawImmediately
    
End Sub

'When a new image has been activated, call this function to apply all relevant UI changes.
Public Sub NotifyNewActiveImage(Optional ByVal newImageIndex As Long = -1)
    
    'If an image is *not* specified, assume this is in reference to the currently active image
    If (newImageIndex < 0) Then newImageIndex = g_CurrentImage
    
    'The toolbar must redraw itself to match the newly activated image
    FormMain.mainCanvas(0).NotifyTabstripNewActiveImage newImageIndex
    
    'A newly activated image requires a whole swath of UI changes.  Ask SyncInterfaceToCurrentImage to handle this for us.
    SyncInterfaceToCurrentImage
    
End Sub

'I'm not very happy about needing this function.  If an action does something that requires a tabstrip redraw, it should be handled
' by the dedicated NotifyXYZ functions, above.  The tabstrip should not require special handling.  That said, this is a temporary stopgap
' until we fix some widespread UI synchronization issues throughout the project.
Public Sub RequestTabstripRedraw()
    FormMain.mainCanvas(0).NotifyTabstripTotalRedrawRequired
End Sub

'If a preview control won't be activated for a given dialog, call this function to display a persistent
' "no preview available" message.  (Note: for this to work, you must not attempt to supply updated preview images
' to the underlying control.  If you do, those images will obviously overwrite this warning!)
Public Sub ShowDisabledPreviewImage(ByRef dstPreview As pdFxPreviewCtl)
    
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    tmpDIB.CreateBlank dstPreview.GetPreviewWidth, dstPreview.GetPreviewHeight

    Dim notifyFont As pdFont
    Set notifyFont = New pdFont
    notifyFont.SetFontFace g_InterfaceFont
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
    
End Sub
