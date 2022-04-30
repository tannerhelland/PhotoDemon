Attribute VB_Name = "Hotkeys"
'***************************************************************************
'PhotoDemon Custom Hotkey handler
'Copyright 2015-2022 by Tanner Helland and contributors
'Created: 06/November/15 (formally split off from a heavily modified vbaIHookControl by Steve McMahon)
'Last updated: 04/October/21
'Last update: create this new module to host actual hotkey management; the old pdAccelerator control
'             on FormMain is still used to for actual key-hooking and raising hotkey events, but it
'             is no longer responsible for *any* element of actual hotkey storage and management.
'
'In 2022 (hopefully), PhotoDemon *finally* (dramatic breath) provides a way for users to specify custom hotkeys.
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

'Each hotkey must store a key code, shift state (can be 0), and action ID string.  The action ID string
' will be blindly forwarded to the Actions.LaunchAction_ByName() function, so make sure the spelling
' (and case! always lowercase!) match the action as it is declared there.
Private Type PD_Hotkey
    hkKeyCode As Long
    hkShiftState As ShiftConstants
    hkAction As String
End Type

'The list of hotkeys is stored in a basic array.  This makes it easy to set/retrieve values using
' built-in VB functions, and because the list of keys is short, performance isn't in issue.
Private m_Hotkeys() As PD_Hotkey
Private m_NumOfHotkeys As Long
Private Const INITIAL_HOTKEY_LIST_SIZE As Long = 16&

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

Private m_CommonMenuText() As String

'Add a new hotkey to the collection.  While the final parameter is marked as OPTIONAL, that's purely to
' allow the preceding constant (shift modifiers, which are often null) to be optional.  I plan to reorder
' parameters in the future to make it abundantly clear that the hotkey action is 100% MANDATORY lol.
'
'RETURNS: the ID (index) of the added hotkey
Public Function AddHotkey(ByVal vKeyCode As KeyCodeConstants, Optional ByVal Shift As ShiftConstants = 0&, Optional ByVal hotKeyAction As String = vbNullString) As Long
    
    'If this hotkey already exists in the collection, we will overwrite it with the new hotkey target.
    ' (This works well for overwriting PD's default hotkeys with new ones specified by the user.)
    Const PRINT_WARNING_ON_HOTKEY_DUPE As Boolean = True
    
    Dim idxHotkey As Long
    idxHotkey = Hotkeys.GetHotkeyIndex(vKeyCode, Shift)
    If (idxHotkey >= 0) Then
        'TODO: notify old menu here, so it can remove hotkey info??
        If PRINT_WARNING_ON_HOTKEY_DUPE Then PDDebug.LogAction "WARNING: duplicate hotkey: " & Hotkeys.GetHotKeyAction(idxHotkey)
    
    'If this is a novel entry, enlarge the list accordingly
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
    
    'Add the new entry (or overwrite the previous one, doesn't matter)
    With m_Hotkeys(idxHotkey)
        .hkKeyCode = vKeyCode
        .hkShiftState = Shift
        .hkAction = hotKeyAction
    End With
    
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

'Initialize a default set of program hotkeys.  For the most part, these attempt to mimic hotkey
' "conventions" in popular photo editors (with a strong emphasis on Photoshop).  Order does not
' matter when adding hotkeys.  Duplicate hotkeys are also okay; the second instance will simply
' overwrite the previous instance, if any.  Action strings need to ALWAYS be full lowercase,
' and identical to their corresponding action in the Menus module.  (This is how hotkeys get
' matched up to a corresponding menu, which is important since that's the primary mechanism for
' discovery!)
Public Sub InitializeDefaultHotkeys()
    
    'Special hotkeys
    Hotkeys.AddHotkey vbKeyF, vbCtrlMask, "tool_search"
    
    'Tool hotkeys (e.g. keys not associated with menus)
    Hotkeys.AddHotkey vbKeyH, , "tool_hand"
    Hotkeys.AddHotkey vbKeyZ, , "tool_zoom"
    Hotkeys.AddHotkey vbKeyM, , "tool_move"
    Hotkeys.AddHotkey vbKeyI, , "tool_colorselect"
    
    'Note that some hotkeys do double-duty in tool selection; you can press some of these shortcuts multiple times
    ' to toggle between similar tools (e.g. rectangular and elliptical selections).  Details can be found in
    ' FormMain.pdHotkey event handlers.
    Hotkeys.AddHotkey vbKeyS, , "tool_select_rect"
    Hotkeys.AddHotkey vbKeyL, , "tool_select_lasso"
    Hotkeys.AddHotkey vbKeyW, , "tool_select_wand"
    Hotkeys.AddHotkey vbKeyT, , "tool_text_basic"
    Hotkeys.AddHotkey vbKeyP, , "tool_pencil"
    Hotkeys.AddHotkey vbKeyB, , "tool_paintbrush"
    Hotkeys.AddHotkey vbKeyE, , "tool_erase"
    Hotkeys.AddHotkey vbKeyC, , "tool_clone"
    Hotkeys.AddHotkey vbKeyF, , "tool_paintbucket"
    Hotkeys.AddHotkey vbKeyG, , "tool_gradient"
    
    'File menu
    Hotkeys.AddHotkey vbKeyN, vbCtrlMask, "file_new"
    Hotkeys.AddHotkey vbKeyO, vbCtrlMask, "file_open"
    
        'Most-recently used files.  Note that we cannot automatically associate these with a menu,
        ' as these menus may not exist at run-time.  (They are created dynamically.)
        Hotkeys.AddHotkey vbKey1, vbCtrlMask Or vbShiftMask, COMMAND_FILE_OPEN_RECENT & "0"
        Hotkeys.AddHotkey vbKey2, vbCtrlMask Or vbShiftMask, COMMAND_FILE_OPEN_RECENT & "1"
        Hotkeys.AddHotkey vbKey3, vbCtrlMask Or vbShiftMask, COMMAND_FILE_OPEN_RECENT & "2"
        Hotkeys.AddHotkey vbKey4, vbCtrlMask Or vbShiftMask, COMMAND_FILE_OPEN_RECENT & "3"
        Hotkeys.AddHotkey vbKey5, vbCtrlMask Or vbShiftMask, COMMAND_FILE_OPEN_RECENT & "4"
        Hotkeys.AddHotkey vbKey6, vbCtrlMask Or vbShiftMask, COMMAND_FILE_OPEN_RECENT & "5"
        Hotkeys.AddHotkey vbKey7, vbCtrlMask Or vbShiftMask, COMMAND_FILE_OPEN_RECENT & "6"
        Hotkeys.AddHotkey vbKey8, vbCtrlMask Or vbShiftMask, COMMAND_FILE_OPEN_RECENT & "7"
        Hotkeys.AddHotkey vbKey9, vbCtrlMask Or vbShiftMask, COMMAND_FILE_OPEN_RECENT & "8"
        Hotkeys.AddHotkey vbKey0, vbCtrlMask Or vbShiftMask, COMMAND_FILE_OPEN_RECENT & "9"
        
        'File -> Import submenu
        Hotkeys.AddHotkey vbKeyI, vbCtrlMask Or vbShiftMask Or vbAltMask, "file_import_scanner"
        Hotkeys.AddHotkey vbKeyD, vbCtrlMask Or vbShiftMask, "file_import_web"
        Hotkeys.AddHotkey vbKeyI, vbCtrlMask Or vbAltMask, "file_import_screenshot"
    
    Hotkeys.AddHotkey vbKeyW, vbCtrlMask, "file_close"
    Hotkeys.AddHotkey vbKeyW, vbCtrlMask Or vbAltMask, "file_closeall"
    Hotkeys.AddHotkey vbKeyS, vbCtrlMask, "file_save"
    Hotkeys.AddHotkey vbKeyS, vbCtrlMask Or vbAltMask Or vbShiftMask, "file_savecopy"
    Hotkeys.AddHotkey vbKeyS, vbCtrlMask Or vbShiftMask, "file_saveas"
    Hotkeys.AddHotkey vbKeyF12, 0, "file_revert"
    Hotkeys.AddHotkey vbKeyB, vbCtrlMask, "file_batch_process"
    Hotkeys.AddHotkey vbKeyP, vbCtrlMask, "file_print"
    Hotkeys.AddHotkey vbKeyQ, vbCtrlMask, "file_quit"
        
    'Edit menu
    Hotkeys.AddHotkey vbKeyZ, vbCtrlMask, "edit_undo"
    Hotkeys.AddHotkey vbKeyY, vbCtrlMask, "edit_redo"
    
    Hotkeys.AddHotkey vbKeyX, vbCtrlMask, "edit_cutlayer"
    Hotkeys.AddHotkey vbKeyC, vbCtrlMask, "edit_copylayer"
    Hotkeys.AddHotkey vbKeyC, vbCtrlMask Or vbShiftMask, "edit_copymerged"
    Hotkeys.AddHotkey vbKeyV, vbCtrlMask, "edit_pasteaslayer"
    Hotkeys.AddHotkey vbKeyV, vbCtrlMask Or vbAltMask, "edit_pastetocursor"
    Hotkeys.AddHotkey vbKeyV, vbCtrlMask Or vbShiftMask, "edit_pasteasimage"
    
    Hotkeys.AddHotkey vbKeyF5, vbCtrlMask Or vbShiftMask, "edit_contentawarefill"
    
    'Image menu
    Hotkeys.AddHotkey vbKeyA, vbCtrlMask Or vbShiftMask, "image_duplicate"
    Hotkeys.AddHotkey vbKeyR, vbCtrlMask, "image_resize"
    Hotkeys.AddHotkey vbKeyR, vbCtrlMask Or vbAltMask, "image_canvassize"
    Hotkeys.AddHotkey vbKeyX, vbCtrlMask Or vbShiftMask, "image_crop"
    Hotkeys.AddHotkey vbKeyX, vbCtrlMask Or vbAltMask, "image_trim"
    
        'Image -> Rotate submenu
        Hotkeys.AddHotkey vbKeyR, vbCtrlMask Or vbShiftMask Or vbAltMask, "image_rotatearbitrary"
    
    Hotkeys.AddHotkey vbKeyE, vbCtrlMask Or vbShiftMask, "image_mergevisible"
    Hotkeys.AddHotkey vbKeyF, vbCtrlMask Or vbShiftMask, "image_flatten"
    
    'Layer Menu
    Hotkeys.AddHotkey vbKeyN, vbCtrlMask Or vbShiftMask, "layer_addbasic"
    Hotkeys.AddHotkey vbKeyJ, vbCtrlMask, "layer_addviacopy"
    Hotkeys.AddHotkey vbKeyJ, vbCtrlMask Or vbShiftMask, "layer_addviacut"
    Hotkeys.AddHotkey vbKeyPageUp, vbCtrlMask Or vbAltMask, "layer_gotop"
    Hotkeys.AddHotkey vbKeyPageUp, vbAltMask, "layer_goup"
    Hotkeys.AddHotkey vbKeyPageDown, vbAltMask, "layer_godown"
    Hotkeys.AddHotkey vbKeyPageDown, vbCtrlMask Or vbAltMask, "layer_gobottom"
    Hotkeys.AddHotkey vbKeyE, vbCtrlMask, "layer_mergedown"
    
    'Select Menu
    Hotkeys.AddHotkey vbKeyA, vbCtrlMask, "select_all"
    Hotkeys.AddHotkey vbKeyD, vbCtrlMask, "select_none"
    Hotkeys.AddHotkey vbKeyI, vbCtrlMask Or vbShiftMask, "select_invert"
    Hotkeys.AddHotkey VK_OEM_6, vbCtrlMask Or vbAltMask, "select_grow"      'VK_OEM_6 = }]
    Hotkeys.AddHotkey VK_OEM_4, vbCtrlMask Or vbAltMask, "select_shrink"    'VK_OEM_4 = {[  (next to the letter P)
    Hotkeys.AddHotkey vbKeyD, vbCtrlMask Or vbAltMask, "select_feather"
    
    'Adjustments Menu
    
    'Adjustments top shortcut menu
    Hotkeys.AddHotkey vbKeyL, vbCtrlMask Or vbShiftMask, "adj_autocorrect"
    Hotkeys.AddHotkey vbKeyL, vbCtrlMask Or vbShiftMask Or vbAltMask, "adj_autoenhance"
    Hotkeys.AddHotkey vbKeyU, vbCtrlMask Or vbShiftMask, "adj_blackandwhite"
    Hotkeys.AddHotkey vbKeyB, vbCtrlMask Or vbShiftMask, "adj_bandc"
    Hotkeys.AddHotkey vbKeyC, vbCtrlMask Or vbAltMask, "adj_colorbalance"
    Hotkeys.AddHotkey vbKeyM, vbCtrlMask, "adj_curves"
    Hotkeys.AddHotkey vbKeyL, vbCtrlMask, "adj_levels"
    Hotkeys.AddHotkey vbKeyH, vbCtrlMask Or vbShiftMask, "adj_sandh"
    Hotkeys.AddHotkey vbKeyAdd, vbCtrlMask Or vbAltMask, "adj_vibrance"
    
        'Color adjustments
        Hotkeys.AddHotkey vbKeyH, vbCtrlMask, "adj_hsl"
        Hotkeys.AddHotkey vbKeyP, vbCtrlMask Or vbAltMask, "adj_photofilters"
        Hotkeys.AddHotkey vbKeyT, vbCtrlMask, "adj_temperature"
        
        'Lighting adjustments
        Hotkeys.AddHotkey vbKeyE, vbCtrlMask Or vbAltMask, "adj_exposure"
        Hotkeys.AddHotkey vbKeyG, vbCtrlMask, "adj_gamma"
        
        'Adjustments -> Invert submenu
        Hotkeys.AddHotkey vbKeyI, vbCtrlMask, "adj_invertrgb"
        
        'Adjustments -> Monochrome submenu
        Hotkeys.AddHotkey vbKeyB, vbCtrlMask Or vbAltMask Or vbShiftMask, "adj_colortomonochrome"
        
    'Tools menu
    Hotkeys.AddHotkey 190, vbCtrlMask Or vbAltMask, "tools_playmacro"   'KeyCode 190 = >.  (two keys to the right of the M letter key)
    Hotkeys.AddHotkey vbKeyK, vbCtrlMask, "tools_options"
    Hotkeys.AddHotkey vbKeyM, vbCtrlMask Or vbAltMask, "tools_3rdpartylibs"
    
    'View menu
    Hotkeys.AddHotkey vbKey0, vbCtrlMask, "view_fit"
    Hotkeys.AddHotkey vbKeyAdd, vbCtrlMask, "view_zoomin"
    Hotkeys.AddHotkey vbKeySubtract, vbCtrlMask, "view_zoomout"
    Hotkeys.AddHotkey vbKey5, vbCtrlMask, "zoom_16_1"
    Hotkeys.AddHotkey vbKey4, vbCtrlMask, "zoom_8_1"
    Hotkeys.AddHotkey vbKey3, vbCtrlMask, "zoom_4_1"
    Hotkeys.AddHotkey vbKey2, vbCtrlMask, "zoom_2_1"
    Hotkeys.AddHotkey vbKey1, vbCtrlMask, "zoom_actual"
    Hotkeys.AddHotkey vbKey2, vbShiftMask, "zoom_1_2"
    Hotkeys.AddHotkey vbKey3, vbShiftMask, "zoom_1_4"
    Hotkeys.AddHotkey vbKey4, vbShiftMask, "zoom_1_8"
    Hotkeys.AddHotkey vbKey5, vbShiftMask, "zoom_1_16"
    
    'Window menu
    Hotkeys.AddHotkey vbKeyPageDown, , "window_next"
    Hotkeys.AddHotkey vbKeyPageUp, , "window_previous"
    
    'All default hotkeys have now been added to the collection.
    
    'Activate hotkey detection on the main form.  (FormMain has a specialized user control that
    ' actually detects hotkey presses, then forwards the key combinations to us for translation
    ' and execution.)
    FormMain.HotkeyManager.Enabled = True
    FormMain.HotkeyManager.ActivateHook
    
    'Before exiting, relay all hotkey assignments to the menu manager; it will generate matching hotkey text
    ' and display it alongside appropriate menu items.
    CacheCommonTranslations
    
    Dim i As Long
    For i = 0 To m_NumOfHotkeys - 1
        Menus.NotifyMenuHotkey m_Hotkeys(i).hkAction, i
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

Public Function GetGenericMenuText(ByVal srcID As PD_CommonMenuText) As String
    GetGenericMenuText = m_CommonMenuText(srcID)
End Function

Public Sub UpdateHotkeyLocalization()
    CacheCommonTranslations
End Sub
