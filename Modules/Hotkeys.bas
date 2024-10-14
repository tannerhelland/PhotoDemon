Attribute VB_Name = "Hotkeys"
'***************************************************************************
'PhotoDemon Custom Hotkey handler
'Copyright 2015-2024 by Tanner Helland and contributors
'Created: 06/November/15 (formally split off from a heavily modified vbaIHookControl by Steve McMahon)
'Last updated: 04/October/21
'Last update: create this new module to host actual hotkey management; the old pdAccelerator control
'             on FormMain is still used to for actual key-hooking and raising hotkey events, but it
'             is no longer responsible for *any* element of actual hotkey storage and management.
'
'In 2024 (hopefully), PhotoDemon *finally* provides a way for users to specify custom hotkeys.
' This module is responsible for managing custom hotkey assignments, and it also manages default
' hotkey behavior (which is what 99.9% of users presumably use).
'
'Actual keypress detection is handled by a specialized user control on FormMain.  Look there for
' details on hooking.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Each hotkey must store a key code, shift state (can be 0), and action ID string.  The action ID string
' will be blindly forwarded to the Actions.LaunchAction_ByName() function, so make sure the spelling
' (and case! always lowercase!) match the action as it is declared there.
Public Type PD_Hotkey
    hkKeyCode As Long               'virtual key-code
    hkShiftState As ShiftConstants  'shift-key states
    hkAction As String              'action triggered by this hotkey
End Type

'The list of hotkeys is stored in a basic array.  This makes it easy to set/retrieve values using
' built-in VB functions, and because the list of keys is short, performance isn't in issue.
Private m_Hotkeys() As PD_Hotkey
Private m_NumOfHotkeys As Long
Private Const INITIAL_HOTKEY_LIST_SIZE As Long = 16&

'A list of PhotoDemon's *default* hotkeys.  The user can always default to these (or we can)
' if something goes catastrophically wrong during custom hotkey initialization.
Private m_DefaultHotkeys() As PD_Hotkey
Private m_NumOfDefaultHotkeys As Long

'To improve performance when language translations are active, we cache certain common translations
' (such as "Ctrl+" for hotkey text) to minimize how many times we have to hit the language engine.
' (Similarly, whenever the active language changes, make sure this text gets updated!)
Public Enum PD_CommonMenuText
    cmt_Ctrl = 0
    cmt_Alt = 1
    cmt_Shift = 2
    cmt_NumEntries = 3
End Enum

#If False Then
    Private Const cmt_Ctrl = 0, cmt_Alt = 1, cmt_Shift = 2, cmt_NumEntries = 3
#End If

Private Const HOTKEY_FILENAME As String = "HotkeysUser.xml"
Private m_CommonMenuText() As String

Private Declare Function GetKeyNameTextW Lib "user32" (ByVal lParam As Long, ByVal lpString As Long, ByVal cchSize As Long) As Long
Private Declare Function MapVirtualKeyW Lib "user32" (ByVal uCode As Long, ByVal uMapType As Long) As Long
Private Declare Function ToUnicode Lib "user32" (ByVal wVirtKey As Long, ByVal wScanCode As Long, ByVal ptrToKeyStateBuffer As Long, ByVal ptrToOutBufferW As Long, ByVal sizeOfOutBufferInWChars As Long, ByVal wFlags As Long) As Long

'Add a new hotkey to the collection.  While the hotKeyAction parameter is marked as OPTIONAL, that's purely to
' allow the preceding constant (shift modifiers, which are often null) to be optional.
' (TODO: reorder parameters in the future to make it abundantly clear that the hotkey action is 100% MANDATORY lol.)
'
'The final optional parameter should be TRUE if
'
'RETURNS: the ID (index) of the added hotkey
Public Function AddHotkey(ByVal vKeyCode As KeyCodeConstants, Optional ByVal Shift As ShiftConstants = 0&, Optional ByVal hotKeyAction As String = vbNullString, Optional ByVal hotKeyIsPDDefault As Boolean = False) As Long
    
    'If this hotkey already exists in the collection, we will overwrite it with the new hotkey target.
    ' (This works well for overwriting PD's default hotkeys with new ones specified by the user.)
    Const PRINT_WARNING_ON_HOTKEY_DUPE As Boolean = True
    
    Dim idxHotkey As Long
    idxHotkey = Hotkeys.GetHotkeyIndex(vKeyCode, Shift, hotKeyIsPDDefault)
    
    If (idxHotkey >= 0) Then
        'TODO: notify old menu here, so it can remove hotkey info??
        If PRINT_WARNING_ON_HOTKEY_DUPE Then PDDebug.LogAction "WARNING: duplicate hotkey: " & Hotkeys.GetHotKeyAction(idxHotkey)
    
    'If this is a novel entry, enlarge the list accordingly
    Else
        
        If hotKeyIsPDDefault Then
        
            If (m_NumOfDefaultHotkeys = 0) Then
                ReDim m_DefaultHotkeys(0 To INITIAL_HOTKEY_LIST_SIZE - 1) As PD_Hotkey
            Else
                If (m_NumOfDefaultHotkeys > UBound(m_DefaultHotkeys)) Then ReDim Preserve m_DefaultHotkeys(0 To m_NumOfDefaultHotkeys * 2 - 1) As PD_Hotkey
            End If
            
            'Tag the current position and increment the total hotkey count accordingly
            idxHotkey = m_NumOfDefaultHotkeys
            m_NumOfDefaultHotkeys = m_NumOfDefaultHotkeys + 1
            
        Else
            
            If (m_NumOfHotkeys = 0) Then
                ReDim m_Hotkeys(0 To INITIAL_HOTKEY_LIST_SIZE - 1) As PD_Hotkey
            Else
                If (m_NumOfHotkeys > UBound(m_Hotkeys)) Then ReDim Preserve m_Hotkeys(0 To m_NumOfHotkeys * 2 - 1) As PD_Hotkey
            End If
            
            'Tag the current position and increment the total hotkey count accordingly
            idxHotkey = m_NumOfHotkeys
            m_NumOfHotkeys = m_NumOfHotkeys + 1
            
        End If
            
    End If
    
    If hotKeyIsPDDefault Then
        
        'Add the new entry (or overwrite the previous one, doesn't matter)
        With m_DefaultHotkeys(idxHotkey)
            .hkKeyCode = vKeyCode
            .hkShiftState = Shift
            .hkAction = hotKeyAction
        End With
        
    Else
    
        'Add the new entry (or overwrite the previous one, doesn't matter)
        With m_Hotkeys(idxHotkey)
            .hkKeyCode = vKeyCode
            .hkShiftState = Shift
            .hkAction = hotKeyAction
        End With
        
    End If
    
    'Return the matching index
    AddHotkey = idxHotkey
    
    'TODO: consider notifying corresponding menu here, instead of in a batch operation at the end?
    
End Function

'Outside functions can retrieve certain accelerator properties.  Note that - by design - these properties should
' only be retrieved from inside an Accelerator event.
Public Function GetNumOfHotkeys() As Long
    GetNumOfHotkeys = m_NumOfHotkeys
End Function

Public Function GetHotKeyAction(ByVal idxHotkey As Long) As String
    If (idxHotkey >= 0) And (idxHotkey < m_NumOfHotkeys) Then
        GetHotKeyAction = m_Hotkeys(idxHotkey).hkAction
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

'If an accelerator exists in our current collection, this will return a value >= 0
' corresponding to its position in the primary tracking array.
Public Function GetHotkeyIndex(ByVal vKeyCode As KeyCodeConstants, ByVal Shift As ShiftConstants, Optional ByVal useDefaultTable As Boolean = False) As Long
    
    GetHotkeyIndex = -1
    Dim i As Long
    
    If useDefaultTable Then
    
        If (m_NumOfDefaultHotkeys > 0) Then
            For i = 0 To m_NumOfDefaultHotkeys - 1
                If (m_DefaultHotkeys(i).hkKeyCode = vKeyCode) And (m_DefaultHotkeys(i).hkShiftState = Shift) Then
                    GetHotkeyIndex = i
                    Exit For
                End If
            Next i
        End If
        
    Else
        
        If (m_NumOfHotkeys > 0) Then
            For i = 0 To m_NumOfHotkeys - 1
                If (m_Hotkeys(i).hkKeyCode = vKeyCode) And (m_Hotkeys(i).hkShiftState = Shift) Then
                    GetHotkeyIndex = i
                    Exit For
                End If
            Next i
        End If
        
    End If

End Function

'Return the current name of the file where custom hotkeys - if they exist - reside.
Public Function GetNameOfHotkeyFile() As String
    Const PERSISTENT_HOTKEY_FILENAME As String = "Current"
    GetNameOfHotkeyFile = UserPrefs.GetHotkeyPath() & PERSISTENT_HOTKEY_FILENAME & ".xml"
End Function

'Initialize a default set of program hotkeys.  For the most part, these attempt to mimic hotkey
' "conventions" in popular photo editors (with a strong emphasis on Photoshop).  Order does not
' matter when adding hotkeys.  Duplicate hotkeys are also okay; the second instance will simply
' overwrite the previous instance, if any.  Action strings need to ALWAYS be full lowercase,
' and identical to their corresponding action in the Menus module.  (This is how hotkeys get
' matched up to a corresponding menu, which is important since that's the primary mechanism for
' discovery!)
Public Sub InitializeDefaultHotkeys()
    
    'Special hotkeys
    Hotkeys.AddHotkey vbKeyF, vbCtrlMask, "tool_search", True
    
    'Tool hotkeys (e.g. keys not associated with menus)
    Hotkeys.AddHotkey vbKeyH, , "tool_hand", True
    Hotkeys.AddHotkey vbKeyZ, , "tool_zoom", True
    Hotkeys.AddHotkey vbKeyM, , "tool_move", True
    Hotkeys.AddHotkey vbKeyI, , "tool_colorselect", True
    
    'Note that some hotkeys do double-duty in tool selection; you can press some of these shortcuts multiple times
    ' to toggle between similar tools (e.g. rectangular and elliptical selections).  Details can be found in
    ' FormMain.pdHotkey event handlers.
    Hotkeys.AddHotkey vbKeyS, , "tool_select_rect", True
    Hotkeys.AddHotkey vbKeyL, , "tool_select_lasso", True
    Hotkeys.AddHotkey vbKeyW, , "tool_select_wand", True
    Hotkeys.AddHotkey vbKeyT, , "tool_text_basic", True
    Hotkeys.AddHotkey vbKeyP, , "tool_pencil", True
    Hotkeys.AddHotkey vbKeyB, , "tool_paintbrush", True
    Hotkeys.AddHotkey vbKeyE, , "tool_erase", True
    Hotkeys.AddHotkey vbKeyC, , "tool_clone", True
    Hotkeys.AddHotkey vbKeyF, , "tool_paintbucket", True
    Hotkeys.AddHotkey vbKeyG, , "tool_gradient", True
    
    'Tool modifiers; UI setting changes only!
    Hotkeys.AddHotkey VK_OEM_4, , "tool_active_sizedown", True
    Hotkeys.AddHotkey VK_OEM_6, , "tool_active_sizeup", True
    Hotkeys.AddHotkey VK_OEM_4, vbShiftMask, "tool_active_hardnessdown", True
    Hotkeys.AddHotkey VK_OEM_6, vbShiftMask, "tool_active_hardnessup", True
    Hotkeys.AddHotkey VK_CAPITAL, , "tool_active_togglecursor", True
    
    'File menu
    Hotkeys.AddHotkey vbKeyN, vbCtrlMask, "file_new", True
    Hotkeys.AddHotkey vbKeyO, vbCtrlMask, "file_open", True
    
        'Most-recently used files.  Note that we cannot automatically associate these with a menu,
        ' as these menus may not exist at run-time.  (They are created dynamically.)
        Hotkeys.AddHotkey vbKey1, vbCtrlMask Or vbShiftMask, COMMAND_FILE_OPEN_RECENT & "0", True
        Hotkeys.AddHotkey vbKey2, vbCtrlMask Or vbShiftMask, COMMAND_FILE_OPEN_RECENT & "1", True
        Hotkeys.AddHotkey vbKey3, vbCtrlMask Or vbShiftMask, COMMAND_FILE_OPEN_RECENT & "2", True
        Hotkeys.AddHotkey vbKey4, vbCtrlMask Or vbShiftMask, COMMAND_FILE_OPEN_RECENT & "3", True
        Hotkeys.AddHotkey vbKey5, vbCtrlMask Or vbShiftMask, COMMAND_FILE_OPEN_RECENT & "4", True
        Hotkeys.AddHotkey vbKey6, vbCtrlMask Or vbShiftMask, COMMAND_FILE_OPEN_RECENT & "5", True
        Hotkeys.AddHotkey vbKey7, vbCtrlMask Or vbShiftMask, COMMAND_FILE_OPEN_RECENT & "6", True
        Hotkeys.AddHotkey vbKey8, vbCtrlMask Or vbShiftMask, COMMAND_FILE_OPEN_RECENT & "7", True
        Hotkeys.AddHotkey vbKey9, vbCtrlMask Or vbShiftMask, COMMAND_FILE_OPEN_RECENT & "8", True
        Hotkeys.AddHotkey vbKey0, vbCtrlMask Or vbShiftMask, COMMAND_FILE_OPEN_RECENT & "9", True
        
        'File -> Import submenu
        Hotkeys.AddHotkey vbKeyI, vbCtrlMask Or vbShiftMask Or vbAltMask, "file_import_scanner", True
        Hotkeys.AddHotkey vbKeyD, vbCtrlMask Or vbShiftMask, "file_import_web", True
        Hotkeys.AddHotkey vbKeyI, vbCtrlMask Or vbAltMask, "file_import_screenshot", True
    
    Hotkeys.AddHotkey vbKeyW, vbCtrlMask, "file_close", True
    Hotkeys.AddHotkey vbKeyW, vbCtrlMask Or vbAltMask, "file_closeall", True
    Hotkeys.AddHotkey vbKeyS, vbCtrlMask, "file_save", True
    Hotkeys.AddHotkey vbKeyS, vbCtrlMask Or vbAltMask Or vbShiftMask, "file_savecopy", True
    Hotkeys.AddHotkey vbKeyS, vbCtrlMask Or vbShiftMask, "file_saveas", True
    Hotkeys.AddHotkey vbKeyF12, 0, "file_revert", True
    
        'File -> Export submenu
        Hotkeys.AddHotkey vbKeyW, vbCtrlMask Or vbShiftMask Or vbAltMask, "file_export_image", True
        
    Hotkeys.AddHotkey vbKeyB, vbCtrlMask, "file_batch_process", True
    Hotkeys.AddHotkey vbKeyP, vbCtrlMask, "file_print", True
    Hotkeys.AddHotkey vbKeyQ, vbCtrlMask, "file_quit", True
    
    'Edit menu
    Hotkeys.AddHotkey vbKeyZ, vbCtrlMask, "edit_undo", True
    Hotkeys.AddHotkey vbKeyY, vbCtrlMask, "edit_redo", True
    
    Hotkeys.AddHotkey vbKeyY, vbCtrlMask Or vbShiftMask, "edit_repeat", True
    
    Hotkeys.AddHotkey vbKeyX, vbCtrlMask, "edit_cutlayer", True
    Hotkeys.AddHotkey vbKeyC, vbCtrlMask, "edit_copylayer", True
    Hotkeys.AddHotkey vbKeyC, vbCtrlMask Or vbShiftMask, "edit_copymerged", True
    Hotkeys.AddHotkey vbKeyV, vbCtrlMask, "edit_pasteaslayer", True
    Hotkeys.AddHotkey vbKeyV, vbCtrlMask Or vbAltMask, "edit_pastetocursor", True
    Hotkeys.AddHotkey vbKeyV, vbCtrlMask Or vbShiftMask, "edit_pasteasimage", True
    
    Hotkeys.AddHotkey vbKeyF5, vbCtrlMask Or vbShiftMask, "edit_contentawarefill", True
    
    'Image menu
    Hotkeys.AddHotkey vbKeyA, vbCtrlMask Or vbShiftMask, "image_duplicate", True
    Hotkeys.AddHotkey vbKeyR, vbCtrlMask, "image_resize", True
    Hotkeys.AddHotkey vbKeyR, vbCtrlMask Or vbAltMask, "image_canvassize", True
    Hotkeys.AddHotkey vbKeyX, vbCtrlMask Or vbShiftMask, "image_crop", True
    Hotkeys.AddHotkey vbKeyX, vbCtrlMask Or vbAltMask, "image_trim", True
    
        'Image -> Rotate submenu
        Hotkeys.AddHotkey vbKeyR, vbCtrlMask Or vbShiftMask Or vbAltMask, "image_rotatearbitrary", True
    
    Hotkeys.AddHotkey vbKeyE, vbCtrlMask Or vbShiftMask, "image_mergevisible", True
    Hotkeys.AddHotkey vbKeyF, vbCtrlMask Or vbShiftMask, "image_flatten", True
    
    'Layer Menu
    Hotkeys.AddHotkey vbKeyN, vbCtrlMask Or vbShiftMask, "layer_addbasic", True
    Hotkeys.AddHotkey vbKeyJ, vbCtrlMask, "layer_addviacopy", True
    Hotkeys.AddHotkey vbKeyJ, vbCtrlMask Or vbShiftMask, "layer_addviacut", True
    Hotkeys.AddHotkey vbKeyPageUp, vbCtrlMask Or vbAltMask, "layer_gotop", True
    Hotkeys.AddHotkey vbKeyPageUp, vbAltMask, "layer_goup", True
    Hotkeys.AddHotkey vbKeyPageDown, vbAltMask, "layer_godown", True
    Hotkeys.AddHotkey vbKeyPageDown, vbCtrlMask Or vbAltMask, "layer_gobottom", True
    Hotkeys.AddHotkey vbKeyE, vbCtrlMask, "layer_mergedown", True
    
    'Select Menu
    Hotkeys.AddHotkey vbKeyA, vbCtrlMask, "select_all", True
    Hotkeys.AddHotkey vbKeyD, vbCtrlMask, "select_none", True
    Hotkeys.AddHotkey vbKeyI, vbCtrlMask Or vbShiftMask, "select_invert", True
    Hotkeys.AddHotkey VK_OEM_6, vbCtrlMask Or vbAltMask, "select_grow", True     'VK_OEM_6 = }]
    Hotkeys.AddHotkey VK_OEM_4, vbCtrlMask Or vbAltMask, "select_shrink", True   'VK_OEM_4 = {[  (next to the letter P)
    Hotkeys.AddHotkey vbKeyD, vbCtrlMask Or vbAltMask, "select_feather", True
    
    'Adjustments Menu
    
    'Adjustments top shortcut menu
    Hotkeys.AddHotkey vbKeyL, vbCtrlMask Or vbShiftMask, "adj_autocorrect", True
    Hotkeys.AddHotkey vbKeyL, vbCtrlMask Or vbShiftMask Or vbAltMask, "adj_autoenhance", True
    Hotkeys.AddHotkey vbKeyU, vbCtrlMask Or vbShiftMask, "adj_blackandwhite", True
    Hotkeys.AddHotkey vbKeyB, vbCtrlMask Or vbShiftMask, "adj_bandc", True
    Hotkeys.AddHotkey vbKeyC, vbCtrlMask Or vbAltMask, "adj_colorbalance", True
    Hotkeys.AddHotkey vbKeyM, vbCtrlMask, "adj_curves", True
    Hotkeys.AddHotkey vbKeyL, vbCtrlMask, "adj_levels", True
    Hotkeys.AddHotkey vbKeyH, vbCtrlMask Or vbShiftMask, "adj_sandh", True
    Hotkeys.AddHotkey vbKeyAdd, vbCtrlMask Or vbAltMask, "adj_vibrance", True
    
        'Color adjustments
        Hotkeys.AddHotkey vbKeyH, vbCtrlMask, "adj_hsl", True
        Hotkeys.AddHotkey vbKeyP, vbCtrlMask Or vbAltMask, "adj_photofilters", True
        Hotkeys.AddHotkey vbKeyT, vbCtrlMask, "adj_temperature", True
        
        'Lighting adjustments
        Hotkeys.AddHotkey vbKeyE, vbCtrlMask Or vbAltMask, "adj_exposure", True
        Hotkeys.AddHotkey vbKeyG, vbCtrlMask, "adj_gamma", True
        
        'Adjustments -> Invert submenu
        Hotkeys.AddHotkey vbKeyI, vbCtrlMask, "adj_invertrgb", True
        
        'Adjustments -> Monochrome submenu
        Hotkeys.AddHotkey vbKeyB, vbCtrlMask Or vbAltMask Or vbShiftMask, "adj_colortomonochrome", True
        
    'Tools menu
    Hotkeys.AddHotkey 190, vbCtrlMask Or vbAltMask, "tools_playmacro", True  'KeyCode 190 = >.  (two keys to the right of the M letter key)
    Hotkeys.AddHotkey vbKeyK, vbCtrlMask, "tools_options", True
    Hotkeys.AddHotkey vbKeyM, vbCtrlMask Or vbAltMask, "tools_3rdpartylibs", True
    
    'View menu
    Hotkeys.AddHotkey vbKey0, vbCtrlMask, "view_fit", True
    Hotkeys.AddHotkey vbKeyAdd, vbCtrlMask, "view_zoomin", True
    Hotkeys.AddHotkey vbKeySubtract, vbCtrlMask, "view_zoomout", True
    Hotkeys.AddHotkey vbKey5, vbCtrlMask, "zoom_16_1", True
    Hotkeys.AddHotkey vbKey4, vbCtrlMask, "zoom_8_1", True
    Hotkeys.AddHotkey vbKey3, vbCtrlMask, "zoom_4_1", True
    Hotkeys.AddHotkey vbKey2, vbCtrlMask, "zoom_2_1", True
    Hotkeys.AddHotkey vbKey1, vbCtrlMask, "zoom_actual", True
    Hotkeys.AddHotkey vbKey2, vbShiftMask, "zoom_1_2", True
    Hotkeys.AddHotkey vbKey3, vbShiftMask, "zoom_1_4", True
    Hotkeys.AddHotkey vbKey4, vbShiftMask, "zoom_1_8", True
    Hotkeys.AddHotkey vbKey5, vbShiftMask, "zoom_1_16", True
    Hotkeys.AddHotkey VK_OEM_1, vbCtrlMask Or vbShiftMask, "snap_global", True
    
    'Window menu
    Hotkeys.AddHotkey vbKeyPageDown, , "window_next", True
    Hotkeys.AddHotkey vbKeyPageUp, , "window_previous", True
    
End Sub

'Replace PD's current hotkey list with its default hotkey list.  This overwrites *all* user hotkey modifications.
Public Sub CopyDefaultHotkeysToMainHotkeys()
    
    Hotkeys.EraseHotkeyCollection
    
    Dim i As Long
    For i = 0 To m_NumOfDefaultHotkeys - 1
        AddHotkey m_DefaultHotkeys(i).hkKeyCode, m_DefaultHotkeys(i).hkShiftState, m_DefaultHotkeys(i).hkAction, False
    Next i
    
End Sub

'Initialize all hotkeys.  If the user has previously customized hotkeys, this will pull their customized list
' in from file; otherwise, a default set of hotkeys will be initialized.
Public Sub InitializeHotkeys()
    
    'Start by initializing PhotoDemon's default hotkey collection.  We'll only use this if the user hasn't
    ' customized hotkeys previously, but the hotkey editor needs it (so it can restore defaults, as necessary).
    InitializeDefaultHotkeys
    
    Dim listExistsOnDisk As Boolean
    listExistsOnDisk = Files.FileExists(UserPrefs.GetPresetPath() & HOTKEY_FILENAME)
    
    If listExistsOnDisk Then
        'TODO: load from disk
    
    'If the user hasn't customized hotkeys before, that's fine!  Copy PD's default hotkey list over the
    ' current hotkey list.
    Else
        CopyDefaultHotkeysToMainHotkeys
    End If
    
    'All hotkeys - either PD's default ones, or user-customized ones from a saved file - have now been
    ' added to a central hotkey collection.
    
    'Activate hotkey detection on the main form.
    ' (FormMain has a specialized user control that actually detects hotkey presses,
    ' then forwards relevant key combinations to us for translation and execution.)
    FormMain.HotkeyManager.Enabled = True
    FormMain.HotkeyManager.ActivateHook
    
    'Relay all hotkey assignments to the menu manager.  (It needs to generate matching hotkey text
    ' and display it alongside tagged menus.)
    CacheCommonTranslations
    
    Dim i As Long
    For i = 0 To m_NumOfHotkeys - 1
        Menus.NotifyMenuHotkey m_Hotkeys(i).hkAction, i
    Next i
    
End Sub

'Returns a list of all *currently* active hotkeys.  These may originate from an internal default list,
' or the user may have customized them.
'
'Returns: number of hotkeys stored to the destination array.  The array's dimensions are *not* guaranteed
' to exactly match the number of hotkeys returned.
Public Function GetCopyOfAllHotkeys(ByRef dstHotkeys() As PD_Hotkey, Optional ByVal getDefaultHotkeysOnly As Boolean = False) As Long
    
    Dim i As Long
    
    If getDefaultHotkeysOnly Then
    
        If (m_NumOfDefaultHotkeys > 0) Then
            ReDim dstHotkeys(0 To m_NumOfDefaultHotkeys - 1) As PD_Hotkey
            For i = 0 To m_NumOfDefaultHotkeys - 1
                dstHotkeys(i) = m_DefaultHotkeys(i)
            Next i
        End If
        
        GetCopyOfAllHotkeys = m_NumOfDefaultHotkeys
        
    Else
        
        If (m_NumOfHotkeys > 0) Then
            ReDim dstHotkeys(0 To m_NumOfHotkeys - 1) As PD_Hotkey
            For i = 0 To m_NumOfHotkeys - 1
                dstHotkeys(i) = m_Hotkeys(i)
            Next i
        End If
        
        GetCopyOfAllHotkeys = m_NumOfHotkeys
        
    End If
    
End Function

'Erase all current hotkeys.  (Please update the hotkey collection with new hotkeys afterward!)
' Note that this does *not* touch the default hotkey collection - those exist in their own
' m_DefaultHotkeys() array.
Public Sub EraseHotkeyCollection()
    m_NumOfHotkeys = 0
    ReDim m_Hotkeys(0 To INITIAL_HOTKEY_LIST_SIZE - 1) As PD_Hotkey
End Sub

'If a menu has a hotkey associated with it, you can use this function to update the language-specific
' text representation of the hotkey. (This text is appended to the menu caption automatically.)
Public Function GetHotkeyText(ByVal hkID As Long) As String
    
    'Validate ID (which is really just an index into the menu array)
    If (hkID >= 0) And (hkID < m_NumOfHotkeys) Then
        
        With m_Hotkeys(hkID)
            
            GetHotkeyText = vbNullString
            If (.hkShiftState And vbCtrlMask) Then GetHotkeyText = GetHotkeyText & m_CommonMenuText(cmt_Ctrl) & "+"
            If (.hkShiftState And vbAltMask) Then GetHotkeyText = GetHotkeyText & m_CommonMenuText(cmt_Alt) & "+"
            If (.hkShiftState And vbShiftMask) Then GetHotkeyText = GetHotkeyText & m_CommonMenuText(cmt_Shift) & "+"
            
            'Processing the string itself takes a bit of extra work, as some keyboard keys don't automatically map to a
            ' string equivalent.  (Also, translations need to be considered.)
            Dim sChar As String
            
            Const USE_API_FOR_CHAR_TRANSLATION As Boolean = True
            If USE_API_FOR_CHAR_TRANSLATION Then
                sChar = GetCharFromKeyCode(.hkKeyCode)
                
            Else
                
                Select Case .hkKeyCode
                
                    Case vbKeyAdd, VK_OEM_PLUS
                        sChar = "+"
                    
                    Case vbKeySubtract, VK_OEM_MINUS
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
                    Case VK_OEM_COMMA
                        sChar = ","
                        
                    Case VK_OEM_PERIOD
                        sChar = "."
                    
                    Case VK_OEM_1
                        sChar = ";"
                        
                    Case VK_OEM_4
                        sChar = "["
                        
                    Case VK_OEM_6
                        sChar = "]"
                        
                    Case VK_OEM_7
                        sChar = "'"
                    
                    'This is a stupid hack; APIs need to be used instead, although their results may be "unpredictable".
                    ' See https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-mapvirtualkeyw
                    Case Else
                        sChar = UCase$(ChrW$(.hkKeyCode))
                    
                End Select
                
            End If
            
        End With
        
        GetHotkeyText = GetHotkeyText & sChar
        
    '/invalid hotkey ID
    Else
        GetHotkeyText = vbNullString
    End If
    
End Function

'Convert a virtual key-code to a UTF-8 string.
' Automatically returns the extended key name, if one exists.  The caller can pass optional byref strings and bools
' to retrieve detailed pass/fail success for either key name.
Public Function GetCharFromKeyCode(ByVal srcKeyCode As Long, Optional ByRef outKeyName As String, Optional ByRef outKeyNameExists As Boolean, Optional ByRef outKeyNameExtended As String, Optional ByRef outKeyNameExtendedExists As Boolean) As String
    
    Select Case srcKeyCode
        
        'Some unreadable chars have to be manually entered
        Case 8
            outKeyNameExists = True
            outKeyName = g_Language.TranslateMessage("Backspace")
        Case 9
            outKeyNameExists = True
            outKeyName = g_Language.TranslateMessage("Tab")
        Case &H1B
            outKeyNameExists = True
            outKeyName = g_Language.TranslateMessage("Escape")
        
        'Other ones can be pulled from the keyboard driver
        Case Else
            
            'Convert the keycode to a scancode
            Dim retCode As Long, retCode2 As Long
            Const MAPVK_VK_TO_VSC As Long = 0
            retCode = MapVirtualKeyW(srcKeyCode, MAPVK_VK_TO_VSC)
            
            Dim finalScanCode As Long
            finalScanCode = (retCode And &HFFFF&) * 65536
            
            'Use the scan code to pull an actual key name.  Note that we're gonna do this twice:
            ' 1) as a non-extended key
            ' 2) as an extended key
            '
            'If the two results differ, we will use the extended key name as it's generally more intuitive
            ' (e.g. on my laptop, this returns "page down" instead of "num 3").  Note that an extended key
            ' version is *not* guaranteed to exist for all keys, however - on my keyboard, function keys
            ' return nothing for their "extended" version despite those being presented as media keys.
            '
            'Testing this working theory across other locales remains TBD!
            outKeyNameExists = GetKeyName_Normal(finalScanCode, outKeyName)
            outKeyNameExtendedExists = GetKeyName_Extended(finalScanCode, outKeyNameExtended)
            
    End Select
    
    If outKeyNameExtendedExists Then
        GetCharFromKeyCode = outKeyNameExtended
    Else
        GetCharFromKeyCode = outKeyName
    End If
    
End Function

'Thin wrappers to GetKeyNameTextW
Private Function GetKeyName_Normal(ByVal srcScanCode As Long, ByRef outKeyName As String) As Boolean
    
    Const NAME_BUFF_SIZE_IN_CHARS As Long = 32
    outKeyName = String$(NAME_BUFF_SIZE_IN_CHARS, 0)
    
    Dim retCode As Long
    retCode = GetKeyNameTextW(srcScanCode, StrPtr(outKeyName), NAME_BUFF_SIZE_IN_CHARS)   'Buffer length *includes* terminating null
    GetKeyName_Normal = (retCode > 0)
    If GetKeyName_Normal Then outKeyName = Trim$(Strings.TrimNull(Left$(outKeyName, retCode))) Else outKeyName = vbNullString
    
End Function

Private Function GetKeyName_Extended(ByVal srcScanCode As Long, ByRef outKeyName As String) As Boolean
    
    Const NAME_BUFF_SIZE_IN_CHARS As Long = 32
    outKeyName = String$(NAME_BUFF_SIZE_IN_CHARS, 0)
    
    Dim retCode As Long
    retCode = GetKeyNameTextW(srcScanCode Or (2 ^ 24), StrPtr(outKeyName), NAME_BUFF_SIZE_IN_CHARS)  'Buffer length *includes* terminating null
    GetKeyName_Extended = (retCode > 0)
    If GetKeyName_Extended Then outKeyName = Trim$(Strings.TrimNull(Left$(outKeyName, retCode))) Else outKeyName = vbNullString

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

Public Function GetGenericMenuText(ByVal srcID As PD_CommonMenuText) As String
    GetGenericMenuText = m_CommonMenuText(srcID)
End Function

Public Sub UpdateHotkeyLocalization()
    CacheCommonTranslations
End Sub
