Attribute VB_Name = "Hotkeys"
'***************************************************************************
'PhotoDemon Custom Hotkey handler
'Copyright 2015-2021 by Tanner Helland and contributors
'Created: 06/November/15 (formally split off from a heavily modified vbaIHookControl by Steve McMahon)
'Last updated: 04/October/21
'Last update: create this new module to host actual hotkey management; the old pdAccelerator control
'             on FormMain is still used to for actual key-hooking and raising hotkey events, but it
'             is no longer responsible for *any* element of actual hotkey storage and management.
'
'In 2021, PhotoDemon *finally* (dramatic breath) provides a way for users to specify custom hotkeys.
' This module is responsible for managing those custom-hotkey assignments, and it also manages default
' hotkey behavior (which is what 99.9% of users will presumably be using).
'
'Actual keypress detection is handled by a specialized user control on FormMain.  Look there for
' details on hooking.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Each hotkey stores several additional (and sometimes optional) parameters.  This spares us from writing specialized
' handling code for each individual keypress.
Private Type PD_Hotkey
    hkKeyCode As Long
    hkShiftState As ShiftConstants
    hkIsProcessorString As Boolean
    hkRequiresOpenImage As Boolean
    hkShowProcDialog As Boolean
    hkProcUndo As PD_UndoType
    hkKeyName As String
    hkMenuNameIfAny As String
End Type

'The list of hotkeys is stored in a basic array.  This makes it easy to set/retrieve values using built-in VB functions,
' and because the list of keys is short, performance isn't in issue.
Private m_Hotkeys() As PD_Hotkey
Private m_NumOfHotkeys As Long
Private Const INITIAL_HOTKEY_LIST_SIZE As Long = 16&

'To improve performance when language translations are active, we cache certain common translations
' (such as "Ctrl+" for hotkey text) to minimize how many times we have to hit the language engine.
' (Similarly, whenever the active language changes, make sure this text gets updated!)
Private Enum PD_CommonMenuText
    cmt_Ctrl = 0
    cmt_Alt = 1
    cmt_Shift = 2
    cmt_NumEntries = 3
End Enum

#If False Then
    Private Const cmt_Ctrl = 0, cmt_Alt = 1, cmt_Shift = 2, cmt_NumEntries = 3
#End If

Private m_CommonMenuText() As String

'Add a new accelerator key combination to the collection.
' - "isProcessorString": if TRUE, hotKeyName is assumed to a be a string meant for PD's central processor.
'    It will be directly passed to the processor there when that hotkey is used.
' - "correspondingMenu": a reference to the menu associated with this hotkey.  The reference is used to
'    dynamically draw matching shortcut text onto the menu.  It is not otherwise used.
' - "requiresOpenImage": specifies that this action *must be disallowed* unless one (or more) image(s)
'    are loaded and active.
' - "showProcForm": controls the "showDialog" parameter of processor string directives.
' - "procUndo": controls the "createUndo" parameter of processor string directives.  Remember that UNDO_NOTHING
'    means "do not create Undo data."
Public Function AddHotkey(ByVal vKeyCode As KeyCodeConstants, Optional ByVal Shift As ShiftConstants = 0&, Optional ByVal HotKeyName As String = vbNullString, Optional ByRef correspondingMenu As String = vbNullString, Optional ByVal IsProcessorString As Boolean = False, Optional ByVal requiresOpenImage As Boolean = True, Optional ByVal showProcDialog As Boolean = True, Optional ByVal procUndo As PD_UndoType = UNDO_Nothing) As Long
    
    'If this hotkey already exists in the collection, we will overwrite it with the new hotkey target.
    ' (This works well for overwriting PD's default hotkeys with new ones specified by the user.)
    Const PRINT_WARNING_ON_HOTKEY_DUPE As Boolean = True
    
    Dim idxHotkey As Long
    idxHotkey = Hotkeys.GetHotkeyIndex(vKeyCode, Shift)
    If (idxHotkey >= 0) Then
        'TODO: notify old menu here, so it can remove hotkey info??
        PDDebug.LogAction "WARNING: duplicate hotkey: " & Hotkeys.HotKeyName(idxHotkey)
    End If
    
    'If this is a novel entry, enlarge the list accordingly
    If (idxHotkey < 0) Then
        
        If (m_NumOfHotkeys = 0) Then
            ReDim m_Hotkeys(0 To INITIAL_HOTKEY_LIST_SIZE - 1) As PD_Hotkey
        Else
            If (m_NumOfHotkeys > UBound(m_Hotkeys)) Then ReDim Preserve m_Hotkeys(0 To m_NumOfHotkeys * 2 - 1) As PD_Hotkey
        End If
        
        'Tag the current position and increment the total hotkey count accordingly
        idxHotkey = m_NumOfHotkeys
        m_NumOfHotkeys = m_NumOfHotkeys + 1
        
    Else
    
    End If
    
    'Add the new entry (or overwrite the previous one, doesn't matter)
    With m_Hotkeys(idxHotkey)
        .hkKeyCode = vKeyCode
        .hkShiftState = Shift
        .hkKeyName = HotKeyName
        .hkMenuNameIfAny = correspondingMenu
        .hkIsProcessorString = IsProcessorString
        .hkRequiresOpenImage = requiresOpenImage
        .hkShowProcDialog = showProcDialog
        .hkProcUndo = procUndo
    End With
    
    'Return the matching index
    AddHotkey = idxHotkey
    
    'TODO: consider notifying corresponding menu here?
    
End Function

'Outside functions can retrieve certain accelerator properties.  Note that - by design - these properties should
' only be retrieved from inside an Accelerator event.
Public Function GetNumOfHotkeys() As Long
    GetNumOfHotkeys = m_NumOfHotkeys
End Function

Public Function IsProcessorString(ByVal idxHotkey As Long) As Boolean
    If (idxHotkey >= 0) And (idxHotkey < m_NumOfHotkeys) Then
        IsProcessorString = m_Hotkeys(idxHotkey).hkIsProcessorString
    End If
End Function

Public Function IsImageRequired(ByVal idxHotkey As Long) As Boolean
    If (idxHotkey >= 0) And (idxHotkey < m_NumOfHotkeys) Then
        IsImageRequired = m_Hotkeys(idxHotkey).hkRequiresOpenImage
    End If
End Function

Public Function IsDialogDisplayed(ByVal idxHotkey As Long) As Boolean
    If (idxHotkey >= 0) And (idxHotkey < m_NumOfHotkeys) Then
        IsDialogDisplayed = m_Hotkeys(idxHotkey).hkShowProcDialog
    End If
End Function

Public Function HasMenu(ByVal idxHotkey As Long) As Boolean
    If (idxHotkey >= 0) And (idxHotkey < m_NumOfHotkeys) Then
        HasMenu = (LenB(m_Hotkeys(idxHotkey).hkMenuNameIfAny) <> 0)
    End If
End Function

Public Function HotKeyName(ByVal idxHotkey As Long) As String
    If (idxHotkey >= 0) And (idxHotkey < m_NumOfHotkeys) Then
        HotKeyName = m_Hotkeys(idxHotkey).hkKeyName
    End If
End Function

Public Function GetMenuName(ByVal idxHotkey As Long) As String
    If (idxHotkey >= 0) And (idxHotkey < m_NumOfHotkeys) Then
        GetMenuName = m_Hotkeys(idxHotkey).hkMenuNameIfAny
    End If
End Function

Public Function GetKeyCode(ByVal idxHotkey As Long) As KeyCodeConstants
    If (idxHotkey >= 0) And (idxHotkey < m_NumOfHotkeys) Then
        GetKeyCode = m_Hotkeys(idxHotkey).hkKeyCode
    End If
End Function

Public Function GetShift(ByVal idxHotkey As Long) As ShiftConstants
    If (idxHotkey >= 0) And (idxHotkey < m_NumOfHotkeys) Then
        GetShift = m_Hotkeys(idxHotkey).hkShiftState
    End If
End Function

Public Function ProcUndoValue(ByVal idxHotkey As Long) As PD_UndoType
    ProcUndoValue = m_Hotkeys(idxHotkey).hkProcUndo
End Function

'If an accelerator exists in our current collection, this will return a value >= 0 corresponding to
' its position in the master array.
Public Function GetHotkeyIndex(ByVal vKeyCode As KeyCodeConstants, ByVal Shift As ShiftConstants) As Long
    
    GetHotkeyIndex = -1
    
    If (m_NumOfHotkeys > 0) Then
        
        Dim i As Long
        For i = 0 To m_NumOfHotkeys - 1
            If (m_Hotkeys(i).hkKeyCode = vKeyCode) And (m_Hotkeys(i).hkShiftState = Shift) Then
                GetHotkeyIndex = i
                Exit For
            End If
        Next i
        
    End If

End Function

'Until I have a better place to stick this, hotkeys are handled here, by the menu module.
' This is primarily done because there is (somewhat) tight integration between hotkeys and
' menu captions, and both need to be handled together while accounting for the usual
' nightmares (like language translations).
Public Sub InitializeDefaultHotkeys()
    
    'Special hotkeys
    Hotkeys.AddHotkey vbKeyF, vbCtrlMask, "tool_search", , False, False, False
    
    'Tool hotkeys (e.g. keys not associated with menus)
    Hotkeys.AddHotkey vbKeyH, , "tool_activate_hand", , , , False
    Hotkeys.AddHotkey vbKeyM, , "tool_activate_move", , , , False
    Hotkeys.AddHotkey vbKeyI, , "tool_activate_colorpicker", , , , False
    
    'Note that some hotkeys do double-duty in tool selection; you can press some of these shortcuts multiple times
    ' to toggle between similar tools (e.g. rectangular and elliptical selections).  Details can be found in
    ' FormMain.pdHotkey event handlers.
    Hotkeys.AddHotkey vbKeyS, , "tool_activate_selectrect", , , , False
    Hotkeys.AddHotkey vbKeyL, , "tool_activate_selectlasso", , , , False
    Hotkeys.AddHotkey vbKeyW, , "tool_activate_selectwand", , , , False
    Hotkeys.AddHotkey vbKeyT, , "tool_activate_text", , , , False
    Hotkeys.AddHotkey vbKeyP, , "tool_activate_pencil", , , , False
    Hotkeys.AddHotkey vbKeyB, , "tool_activate_brush", , , , False
    Hotkeys.AddHotkey vbKeyE, , "tool_activate_eraser", , , , False
    Hotkeys.AddHotkey vbKeyC, , "tool_activate_clone", , , , False
    Hotkeys.AddHotkey vbKeyF, , "tool_activate_fill", , , , False
    Hotkeys.AddHotkey vbKeyG, , "tool_activate_gradient", , , , False
    
    'File menu
    Hotkeys.AddHotkey vbKeyN, vbCtrlMask, "New image", "file_new", True, False, True, UNDO_Nothing
    Hotkeys.AddHotkey vbKeyO, vbCtrlMask, "Open", "file_open", True, False, True, UNDO_Nothing
    Hotkeys.AddHotkey vbKeyW, vbCtrlMask, "Close", "file_close", True, True, True, UNDO_Nothing
    Hotkeys.AddHotkey vbKeyW, vbCtrlMask Or vbAltMask, "Close all", "file_closeall", True, True, True, UNDO_Nothing
    Hotkeys.AddHotkey vbKeyS, vbCtrlMask, "Save", "file_save", True, True, True, UNDO_Nothing
    Hotkeys.AddHotkey vbKeyS, vbCtrlMask Or vbAltMask Or vbShiftMask, "Save copy", "file_savecopy", True, False, True, UNDO_Nothing
    Hotkeys.AddHotkey vbKeyS, vbCtrlMask Or vbShiftMask, "Save as", "file_saveas", True, True, True, UNDO_Nothing
    Hotkeys.AddHotkey vbKeyF12, 0, "Revert", "file_revert", True, True, False, UNDO_Nothing
    Hotkeys.AddHotkey vbKeyB, vbCtrlMask, "Batch wizard", "file_batch_process", True, False, True, UNDO_Nothing
    Hotkeys.AddHotkey vbKeyP, vbCtrlMask, "Print", "file_print", True, True, True, UNDO_Nothing
    Hotkeys.AddHotkey vbKeyQ, vbCtrlMask, "Exit program", "file_quit", True, False, True, UNDO_Nothing
    
        'File -> Import submenu
        Hotkeys.AddHotkey vbKeyI, vbCtrlMask Or vbShiftMask Or vbAltMask, "Scan image", "file_import_scanner", True, False, True, UNDO_Nothing
        Hotkeys.AddHotkey vbKeyD, vbCtrlMask Or vbShiftMask, "Internet import", "file_import_web", True, False, True, UNDO_Nothing
        Hotkeys.AddHotkey vbKeyI, vbCtrlMask Or vbAltMask, "Screen capture", "file_import_screenshot", True, False, True, UNDO_Nothing
    
        'Most-recently used files.  Note that we cannot automatically associate these with a menu, as these menus may not
        ' exist at run-time.  (They are dynamically created as necessary.)
        Hotkeys.AddHotkey vbKey1, vbCtrlMask Or vbShiftMask, COMMAND_FILE_OPEN_RECENT & "0", requiresOpenImage:=False
        Hotkeys.AddHotkey vbKey2, vbCtrlMask Or vbShiftMask, COMMAND_FILE_OPEN_RECENT & "1", requiresOpenImage:=False
        Hotkeys.AddHotkey vbKey3, vbCtrlMask Or vbShiftMask, COMMAND_FILE_OPEN_RECENT & "2", requiresOpenImage:=False
        Hotkeys.AddHotkey vbKey4, vbCtrlMask Or vbShiftMask, COMMAND_FILE_OPEN_RECENT & "3", requiresOpenImage:=False
        Hotkeys.AddHotkey vbKey5, vbCtrlMask Or vbShiftMask, COMMAND_FILE_OPEN_RECENT & "4", requiresOpenImage:=False
        Hotkeys.AddHotkey vbKey6, vbCtrlMask Or vbShiftMask, COMMAND_FILE_OPEN_RECENT & "5", requiresOpenImage:=False
        Hotkeys.AddHotkey vbKey7, vbCtrlMask Or vbShiftMask, COMMAND_FILE_OPEN_RECENT & "6", requiresOpenImage:=False
        Hotkeys.AddHotkey vbKey8, vbCtrlMask Or vbShiftMask, COMMAND_FILE_OPEN_RECENT & "7", requiresOpenImage:=False
        Hotkeys.AddHotkey vbKey9, vbCtrlMask Or vbShiftMask, COMMAND_FILE_OPEN_RECENT & "8", requiresOpenImage:=False
        Hotkeys.AddHotkey vbKey0, vbCtrlMask Or vbShiftMask, COMMAND_FILE_OPEN_RECENT & "9", requiresOpenImage:=False
        
    'Edit menu
    Hotkeys.AddHotkey vbKeyZ, vbCtrlMask, "Undo", "edit_undo", True, True, False, UNDO_Nothing
    Hotkeys.AddHotkey vbKeyY, vbCtrlMask, "Redo", "edit_redo", True, True, False, UNDO_Nothing
    
    Hotkeys.AddHotkey vbKeyF, vbCtrlMask Or vbShiftMask, "Repeat last action", "edit_repeat", True, True, False, UNDO_Image
    
    Hotkeys.AddHotkey vbKeyX, vbCtrlMask, "Cut", "edit_cutlayer", True, True, False, UNDO_Image
    'This "cut from layer" hotkey combination is used as "crop to selection" in other software; as such,
    ' I am suspending this instance for now.
    'Hotkeys.AddHotkey vbKeyX, vbCtrlMask Or vbShiftMask, "Cut from layer", "edit_cutlayer", True, True, False, UNDO_Layer
    Hotkeys.AddHotkey vbKeyC, vbCtrlMask, "Copy", "edit_copylayer", True, True, False, UNDO_Nothing
    Hotkeys.AddHotkey vbKeyC, vbCtrlMask Or vbShiftMask, "Copy merged", "edit_copymerged", True, True, False, UNDO_Nothing
    Hotkeys.AddHotkey vbKeyV, vbCtrlMask, "Paste", "edit_pasteaslayer", True, False, False, UNDO_Image_VectorSafe
    Hotkeys.AddHotkey vbKeyV, vbCtrlMask Or vbAltMask, "Paste to cursor", "edit_pastetocursor", True, True, False, UNDO_Image_VectorSafe
    Hotkeys.AddHotkey vbKeyV, vbCtrlMask Or vbShiftMask, "Paste to new image", "edit_pasteasimage", True, False, False, UNDO_Nothing
    
    'Image menu
    Hotkeys.AddHotkey vbKeyA, vbCtrlMask Or vbShiftMask, "Duplicate image", "image_duplicate", True, True, False, UNDO_Nothing
    Hotkeys.AddHotkey vbKeyR, vbCtrlMask, "Resize image", "image_resize", True, True, True, UNDO_Nothing
    Hotkeys.AddHotkey vbKeyR, vbCtrlMask Or vbAltMask, "Canvas size", "image_canvassize", True, True, True, UNDO_ImageHeader
    Hotkeys.AddHotkey vbKeyX, vbCtrlMask Or vbShiftMask, "Crop", "image_crop", True, True, False, UNDO_Image
    Hotkeys.AddHotkey vbKeyX, vbCtrlMask Or vbAltMask, "Trim empty image borders", "image_trim", True, True, False, UNDO_ImageHeader
    
        'Image -> Rotate submenu
        Hotkeys.AddHotkey vbKeyR, vbCtrlMask Or vbShiftMask Or vbAltMask, "Arbitrary image rotation", "image_rotatearbitrary", True, True, True, UNDO_Nothing
    
    'Layer Menu
    Hotkeys.AddHotkey vbKeyN, vbCtrlMask Or vbShiftMask, "Add new layer", "layer_addbasic", True, True, True, UNDO_Nothing
    Hotkeys.AddHotkey vbKeyJ, vbCtrlMask, "Layer via copy", "layer_addviacopy", True, True, False, UNDO_Image_VectorSafe
    Hotkeys.AddHotkey vbKeyJ, vbCtrlMask Or vbShiftMask, "Layer via cut", "layer_addviacut", True, True, False, UNDO_Image
    Hotkeys.AddHotkey vbKeyPageUp, vbCtrlMask Or vbAltMask, "Go to top layer", "layer_gotop", True, True, False, UNDO_Nothing
    Hotkeys.AddHotkey vbKeyPageDown, vbCtrlMask Or vbAltMask, "Go to bottom layer", "layer_gobottom", True, True, False, UNDO_Nothing
    Hotkeys.AddHotkey vbKeyPageUp, vbAltMask, "Go to layer above", "layer_goup", True, True, False, UNDO_Nothing
    Hotkeys.AddHotkey vbKeyPageDown, vbAltMask, "Go to layer below", "layer_godown", True, True, False, UNDO_Nothing
    Hotkeys.AddHotkey vbKeyE, vbCtrlMask, "layer_mergedown", "layer_mergedown", False, True, False, UNDO_Image
    Hotkeys.AddHotkey vbKeyE, vbCtrlMask Or vbShiftMask, "Merge visible layers", "image_mergevisible", True, True, False, UNDO_Image
    Hotkeys.AddHotkey vbKeyF, vbCtrlMask Or vbShiftMask, "Flatten image", "image_flatten", True, True, True, UNDO_Nothing
    
    'Select Menu
    Hotkeys.AddHotkey vbKeyA, vbCtrlMask, "Select all", "select_all", True, True, False, UNDO_Selection
    Hotkeys.AddHotkey vbKeyD, vbCtrlMask, "Remove selection", "select_none", False, True, False, UNDO_Selection
    Hotkeys.AddHotkey vbKeyI, vbCtrlMask Or vbShiftMask, "Invert selection", "select_invert", True, True, False, UNDO_Selection
    'KeyCode VK_OEM_4 = {[  (next to the letter P), VK_OEM_6 = }]
    Hotkeys.AddHotkey VK_OEM_6, vbCtrlMask Or vbAltMask, "Grow selection", "select_grow", True, True, True, UNDO_Nothing
    Hotkeys.AddHotkey VK_OEM_4, vbCtrlMask Or vbAltMask, "Shrink selection", "select_shrink", True, True, True, UNDO_Nothing
    Hotkeys.AddHotkey vbKeyD, vbCtrlMask Or vbAltMask, "Feather selection", "select_feather", True, True, True, UNDO_Nothing
    
    'Adjustments Menu
    
    'Adjustments top shortcut menu
    Hotkeys.AddHotkey vbKeyL, vbCtrlMask Or vbShiftMask, "Auto correct", "adj_autocorrect", True, True, False, UNDO_Layer
    Hotkeys.AddHotkey vbKeyL, vbCtrlMask Or vbShiftMask Or vbAltMask, "Auto enhance", "adj_autoenhance", True, True, False, UNDO_Layer
    Hotkeys.AddHotkey vbKeyU, vbCtrlMask Or vbShiftMask, "Black and white", "adj_blackandwhite", True, True, True, UNDO_Nothing
    Hotkeys.AddHotkey vbKeyB, vbCtrlMask Or vbShiftMask, "Brightness and contrast", "adj_bandc", True, True, True, UNDO_Nothing
    Hotkeys.AddHotkey vbKeyC, vbCtrlMask Or vbAltMask, "Color balance", "adj_colorbalance", True, True, True, UNDO_Nothing
    Hotkeys.AddHotkey vbKeyM, vbCtrlMask, "Curves", "adj_curves", True, True, True, UNDO_Nothing
    Hotkeys.AddHotkey vbKeyL, vbCtrlMask, "Levels", "adj_levels", True, True, True, UNDO_Nothing
    Hotkeys.AddHotkey vbKeyH, vbCtrlMask Or vbShiftMask, "Shadows and highlights", "adj_sandh", True, True, True, UNDO_Nothing
    Hotkeys.AddHotkey vbKeyAdd, vbCtrlMask Or vbAltMask, "Vibrance", "adj_vibrance", True, True, True, UNDO_Nothing
    Hotkeys.AddHotkey VK_OEM_PLUS, vbCtrlMask Or vbAltMask, "Vibrance", , True, True, True, UNDO_Nothing
    
    'Ctrl+W has been remapped to File > Close
    'Hotkeys.AddHotkey vbKeyW, vbCtrlMask, "White balance", "adj_whitebalance", True, True, True, UNDO_Nothing
    
        'Color adjustments
        Hotkeys.AddHotkey vbKeyH, vbCtrlMask, "Hue and saturation", "adj_hsl", True, True, True, UNDO_Nothing
        Hotkeys.AddHotkey vbKeyT, vbCtrlMask, "Temperature", "adj_temperature", True, True, True, UNDO_Nothing
        
        'Lighting adjustments
        Hotkeys.AddHotkey vbKeyG, vbCtrlMask, "Gamma", "adj_gamma", True, True, True, UNDO_Nothing
        
        'Adjustments -> Invert submenu
        Hotkeys.AddHotkey vbKeyI, vbCtrlMask, "Invert RGB", "adj_invertRGB", True, True, False, UNDO_Layer
        
        'Adjustments -> Monochrome submenu
        Hotkeys.AddHotkey vbKeyB, vbCtrlMask Or vbAltMask Or vbShiftMask, "Color to monochrome", "adj_colortomonochrome", True, True, True, UNDO_Nothing
        
        'Adjustments -> Photography submenu
        Hotkeys.AddHotkey vbKeyE, vbCtrlMask Or vbAltMask, "Exposure", "adj_exposure", True, True, True, UNDO_Nothing
        Hotkeys.AddHotkey vbKeyP, vbCtrlMask Or vbAltMask, "Photo filter", "adj_photofilters", True, True, True, UNDO_Nothing
        
    
    'Effects Menu
    'Hotkeys.AddHotkey vbKeyZ, vbCtrlMask Or vbAltMask Or vbShiftMask, "Add RGB noise", FormMain.MnuNoise(1), True, True, True, False
    'Hotkeys.AddHotkey vbKeyG, vbCtrlMask Or vbAltMask Or vbShiftMask, "Gaussian blur", FormMain.MnuBlurFilter(1), True, True, True, False
    'Hotkeys.AddHotkey vbKeyY, vbCtrlMask Or vbAltMask Or vbShiftMask, "Correct lens distortion", FormMain.MnuDistortEffects(1), True, True, True, False
    'Hotkeys.AddHotkey vbKeyU, vbCtrlMask Or vbAltMask Or vbShiftMask, "Unsharp mask", FormMain.MnuSharpen(1), True, True, True, False
    
    'Tools menu
    'KeyCode 190 = >.  (two keys to the right of the M letter key)
    Hotkeys.AddHotkey 190, vbCtrlMask Or vbAltMask, "Play macro", "tools_playmacro", True, True, True, UNDO_Nothing
    
    'Previously, Alt+Enter was used for preferences; I dislike this, however, as holding down the Alt-key
    ' is useful for keyboard navigation of menus (via mnemonics), and if you use Enter to select a menu
    ' item, this accelerator overrides your menu click.  Photoshop uses Ctrl+K - maybe we should
    ' investigate that as an option?  TODO!
    'Hotkeys.AddHotkey vbKeyReturn, vbAltMask, "Preferences", "tools_options", False, False, True, UNDO_Nothing
    Hotkeys.AddHotkey vbKeyM, vbCtrlMask Or vbAltMask, "Plugin manager", "tools_3rdpartylibs", False, False, True, UNDO_Nothing
    
    'View menu
    Hotkeys.AddHotkey vbKey0, vbCtrlMask, "FitOnScreen", "view_fit", False, True, False, UNDO_Nothing
    Hotkeys.AddHotkey vbKeyAdd, vbCtrlMask, "Zoom_In", "view_zoomin", False, True, False, UNDO_Nothing
    Hotkeys.AddHotkey VK_OEM_PLUS, vbCtrlMask, "Zoom_In", , False, True, False, UNDO_Nothing
    Hotkeys.AddHotkey vbKeySubtract, vbCtrlMask, "Zoom_Out", "view_zoomout", False, True, False, UNDO_Nothing
    Hotkeys.AddHotkey VK_OEM_MINUS, vbCtrlMask, "Zoom_Out", , False, True, False, UNDO_Nothing
    Hotkeys.AddHotkey vbKey5, vbCtrlMask, "Zoom_161", "zoom_16_1", False, True, False, UNDO_Nothing
    Hotkeys.AddHotkey vbKey4, vbCtrlMask, "Zoom_81", "zoom_8_1", False, True, False, UNDO_Nothing
    Hotkeys.AddHotkey vbKey3, vbCtrlMask, "Zoom_41", "zoom_4_1", False, True, False, UNDO_Nothing
    Hotkeys.AddHotkey vbKey2, vbCtrlMask, "Zoom_21", "zoom_2_1", False, True, False, UNDO_Nothing
    Hotkeys.AddHotkey vbKey1, vbCtrlMask, "Actual_Size", "zoom_actual", False, True, False, UNDO_Nothing
    Hotkeys.AddHotkey vbKey2, vbShiftMask, "Zoom_12", "zoom_1_2", False, True, False, UNDO_Nothing
    Hotkeys.AddHotkey vbKey3, vbShiftMask, "Zoom_14", "zoom_1_4", False, True, False, UNDO_Nothing
    Hotkeys.AddHotkey vbKey4, vbShiftMask, "Zoom_18", "zoom_1_8", False, True, False, UNDO_Nothing
    Hotkeys.AddHotkey vbKey5, vbShiftMask, "Zoom_116", "zoom_1_16", False, True, False, UNDO_Nothing
    
    'Window menu
    Hotkeys.AddHotkey vbKeyPageDown, 0, "Next_Image", "window_next", False, True, False, UNDO_Nothing
    Hotkeys.AddHotkey vbKeyPageUp, 0, "Prev_Image", "window_previous", False, True, False, UNDO_Nothing
    
    'All hotkeys have now been added to the collection.
    
    'Activate hotkey detection on the main form.  (FormMain has a specialized user control that actually detects
    ' hotkey presses, then forwards the key combinations to us for translation and execution.)
    FormMain.HotkeyManager.Enabled = True
    FormMain.HotkeyManager.ActivateHook
    
    'Before exiting, relay all hotkey assignments to the menu manager; it will generate matching hotkey text
    ' and display it alongside appropriate menu items.
    CacheCommonTranslations
    
    Dim i As Long
    For i = 0 To m_NumOfHotkeys - 1
        If (LenB(m_Hotkeys(i).hkMenuNameIfAny) > 0) Then Menus.NotifyMenuHotkey m_Hotkeys(i).hkMenuNameIfAny, i
    Next i
    
End Sub

'If a menu has a hotkey associated with it, you can use this function to update the language-specific text representation of the hotkey.
' (This text is appended to the menu caption automatically.)
Public Function GetHotkeyText(ByVal hkID As Long) As String
    
    'Validate ID (which is really just an index into the menu array)
    If (hkID >= 0) And (hkID < m_NumOfHotkeys) Then
        
        With m_Hotkeys(hkID)
            
            If (.hkShiftState And vbCtrlMask) Then GetHotkeyText = m_CommonMenuText(cmt_Ctrl) & "+"
            If (.hkShiftState And vbAltMask) Then GetHotkeyText = GetHotkeyText & m_CommonMenuText(cmt_Alt) & "+"
            If (.hkShiftState And vbShiftMask) Then GetHotkeyText = GetHotkeyText & m_CommonMenuText(cmt_Shift) & "+"
            
            'Processing the string itself takes a bit of extra work, as some keyboard keys don't automatically map to a
            ' string equivalent.  (Also, translations need to be considered.)
            Dim sChar As String
            
            Select Case .hkKeyCode
            
                Case vbKeyAdd
                    sChar = "+"
                
                Case vbKeySubtract
                    sChar = "-"
                
                Case vbKeyReturn
                    sChar = g_Language.TranslateMessage("Enter")
                
                Case vbKeyPageUp
                    sChar = g_Language.TranslateMessage("Page Up")
                
                Case vbKeyPageDown
                    sChar = g_Language.TranslateMessage("Page Down")
                    
                Case vbKeyF1 To vbKeyF16
                    sChar = "F" & (.hkKeyCode - 111)
                
                'In the future I would like to enumerate virtual key bindings properly, using the data at this link:
                ' http://msdn.microsoft.com/en-us/library/windows/desktop/dd375731%28v=vs.85%29.aspx
                ' At the moment, however, they're implemented as magic numbers.
                Case 188
                    sChar = ","
                    
                Case 190
                    sChar = "."
                    
                Case 219
                    sChar = "["
                    
                Case 221
                    sChar = "]"
                
                'This is a stupid hack; APIs need to be used instead, although their results may be "unpredictable".
                ' See https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-mapvirtualkeyw
                Case Else
                    sChar = UCase$(ChrW$(.hkKeyCode))
                
            End Select
        
        End With
        
        GetHotkeyText = GetHotkeyText & sChar
        
    '/invalid hotkey ID
    Else
        GetHotkeyText = vbNullString
    End If
    
End Function

'Some hotkey-related text is accessed very frequently (e.g. "Ctrl"), so when a translation is active,
' we cache common translations locally instead of regenerating them over and over.
Private Sub CacheCommonTranslations()
    ReDim m_CommonMenuText(0 To cmt_NumEntries - 1) As String
    If (Not g_Language Is Nothing) Then
        m_CommonMenuText(cmt_Ctrl) = g_Language.TranslateMessage("Ctrl")
        m_CommonMenuText(cmt_Alt) = g_Language.TranslateMessage("Alt")
        m_CommonMenuText(cmt_Shift) = g_Language.TranslateMessage("Shift")
    Else
        PDDebug.LogAction "WARNING!  g_Language isn't available, so hotkey captions won't be correct."
    End If
End Sub

Public Sub UpdateHotkeyLocalization()
    CacheCommonTranslations
End Sub
