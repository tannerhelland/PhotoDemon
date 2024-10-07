VERSION 5.00
Begin VB.Form FormHotkeys 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   " Keyboard shortcuts"
   ClientHeight    =   7470
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11760
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   498
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   784
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdCheckBox chkAutoCapture 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   5160
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   661
      Caption         =   "allow auto hotkey capture"
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   375
      Left            =   120
      Top             =   5640
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   661
      Caption         =   "key modifiers"
      FontSize        =   12
   End
   Begin PhotoDemon.pdDropDown ddKey 
      Height          =   855
      Left            =   6000
      TabIndex        =   6
      Top             =   5640
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   1508
      Caption         =   "key"
      FontSize        =   11
   End
   Begin PhotoDemon.pdCheckBox chkModifier 
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   6120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Caption         =   "Ctrl"
      FontSize        =   11
      Value           =   0   'False
   End
   Begin PhotoDemon.pdTextBox txtHotkey 
      Height          =   735
      Left            =   8520
      TabIndex        =   2
      Top             =   2640
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1296
      FontSize        =   12
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   6735
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   1296
      HideRandomizeButton=   -1  'True
   End
   Begin PhotoDemon.pdTreeviewOD tvMenus 
      Height          =   4935
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   9340
   End
   Begin PhotoDemon.pdCheckBox chkModifier 
      Height          =   375
      Index           =   1
      Left            =   2160
      TabIndex        =   4
      Top             =   6120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Caption         =   "Alt"
      FontSize        =   11
      Value           =   0   'False
   End
   Begin PhotoDemon.pdCheckBox chkModifier 
      Height          =   375
      Index           =   2
      Left            =   4080
      TabIndex        =   5
      Top             =   6120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Caption         =   "Shift"
      FontSize        =   11
      Value           =   0   'False
   End
End
Attribute VB_Name = "FormHotkeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Customizable hotkeys dialog
'Copyright 2024-2024 by Tanner Helland
'Created: 09/September/24
'Last updated: 09/September/24
'Last update: initial build
'
'This dialog allows the user to customize hotkeys.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Objects retrieved from other engines
Private m_Menus() As PD_MenuEntry, m_NumOfMenus As Long
Private m_Hotkeys() As PD_Hotkey, m_NumOfHotkeys As Long

'To simplify lookup of menus (required for pairing children IDs against parent IDs),
' we use a hash table.
Private m_MenuHash As pdVariantHash

'Menu and hotkey data are merged into this local struct, which is far more convenient for this UI
Private Type PD_HotkeyUI
    hk_TextEn As String
    hk_TextLocalized As String
    hk_ActionID As String
    hk_ParentID As String
    hk_HasChildren As Boolean
    hk_SubmenuLevel As Integer
    hk_NumParents As Long
    hk_KeyCode As Long
    hk_ShiftState As ShiftConstants
    hk_HotkeyText As String
    
    'If the user cancels this dialog (or reverts changes), we can use these backup copies of original hotkey data
    ' to revert everything to a pristine, untouched state.
    hk_BackupKeyCode As Long
    hk_BackupShiftState As Long
    hk_BackupHotkeyText As String
    
End Type

'Menu and hotkey information gets merged into this local array, which is much easier to manage
' against the UI of this dialog
Private m_Items() As PD_HotkeyUI, m_numItems As Long

'Height of each list item in the custom-drawn treeview, in pixels, at 96 DPI
Private Const BLOCKHEIGHT As Long = 32

'Two font objects; one for menus that are allowed to have hotkeys, and one for menus that are not
' (e.g. top-level menus or parent-only menus).
Private m_FontAllowed As pdFont, m_FontDisallowed As pdFont, m_FontHotkey As pdFont

'All rendering is suspended until the form is loaded
Private m_RenderingOK As Boolean

'Retaining hotkey text in an edit box is non-trivial.  Store the last hotkey text here in WM_KEYDOWN,
' and manually restore it in Change/WM_KEYUP
Private m_backupHotkeyText As String, m_backupHotkeyShift As Long, m_backupHotkeyVKCode As Long
Private m_inAutoUpdate As Boolean

'List of all possible hotkeys.  Used to fill the key dropdown in the right-side selector.
Private Type PD_PossibleHotkey
    ph_VKCode As Long
    ph_KeyName As String
    ph_KeyComments As String    'Comments from MSDN; used for debugging only!
End Type

Private m_numPossibleHotkeys As Long, m_idxOtherHotkey As Long
Private m_possibleHotkeys() As PD_PossibleHotkey

'Last-edited hotkey
Private m_idxLastHotkey As Long

Private Sub chkModifier_Click(Index As Integer)
    UpdateHotkeyManually
End Sub

Private Sub UpdateHotkeyManually()
    
    If (Not m_inAutoUpdate) And (tvMenus.ListIndex >= 0) And (ddKey.ListIndex >= 0) Then
        
        Debug.Print "UpdateHotkeyManually", tvMenus.ListIndex
        
        m_inAutoUpdate = True
        
        Dim newShiftState As Long
        If chkModifier(0).Value Then newShiftState = newShiftState Or vbCtrlMask
        If chkModifier(1).Value Then newShiftState = newShiftState Or vbAltMask
        If chkModifier(2).Value Then newShiftState = newShiftState Or vbShiftMask
        
        With m_Items(tvMenus.ListIndex)
            m_backupHotkeyShift = newShiftState
            .hk_ShiftState = m_backupHotkeyShift
            m_backupHotkeyVKCode = m_possibleHotkeys(ddKey.ListIndex).ph_VKCode
            .hk_KeyCode = m_backupHotkeyVKCode
            m_backupHotkeyText = AutoTextKeyChange(newShiftState, .hk_KeyCode)
            .hk_HotkeyText = m_backupHotkeyText
        End With
        
        'Redraw the listview to reflect the new hotkey
        tvMenus.RequestListRedraw
        
        m_inAutoUpdate = False
        
    End If

End Sub

Private Sub cmdBar_AddCustomPresetData()
    'TODO: save entire hotkey list - this would let the user swap between e.g. "GIMP" and "Photoshop" presets?
End Sub

Private Sub cmdBar_OKClick()
    
    'TODO
    
End Sub

Private Sub cmdBar_ResetClick()
    'TODO
End Sub

Private Sub ddKey_Click()
    UpdateHotkeyManually
End Sub

Private Sub Form_Load()
    
    'No hotkeys have been edited yet
    m_idxLastHotkey = -1
    
    'Retrieve a copy of all menus (including hierarchies and attributes) from the menu manager
    m_NumOfMenus = Menus.GetCopyOfAllMenus(m_Menus)
    
    'Add all menus (by their ID) to a hash table so we can quickly move between IDs and array indices.
    Set m_MenuHash = New pdVariantHash
    
    'Similarly, retrieve a copy of all hotkeys from the hotkey manager
    m_NumOfHotkeys = Hotkeys.GetCopyOfAllHotkeys(m_Hotkeys)
    
    'There will (typically? always?) be fewer hotkeys than there are menu/action targets.  To simplify
    ' correlating between action IDs and hotkey indices, build a quick dictionary.
    Dim cHotkeys As pdVariantHash
    Set cHotkeys = New pdVariantHash
    
    Dim i As Long
    For i = 0 To m_NumOfHotkeys - 1
        cHotkeys.AddItem m_Hotkeys(i).hkAction, i
    Next i
    
    'Turn off automatic redraws in the treeview object
    tvMenus.SetAutomaticRedraws False
    tvMenus.ListItemHeight = BLOCKHEIGHT
    
    'Iterate the menu collection, and pair each menu with a hotkey against its relevant hotkey partner
    ReDim m_Items(0 To m_NumOfMenus - 1) As PD_HotkeyUI
    m_numItems = 0
    
    For i = 0 To m_NumOfMenus - 1
        
        'Ignore separators
        If (m_Menus(i).me_Name <> "-") Then
            
            With m_Items(m_numItems)
                
                'Before doing anything with this menu, add it to a hash table (so we can quickly correlate between
                ' menu positions *in the menu bar* (which is hierarchical) and menu positions *in this array*).
                Dim mnuID As String
                mnuID = GetMenuPositionID(i)
                m_MenuHash.AddItem mnuID, i
                
                'Start by copying over the menu data we can use as-is (like localizations)
                .hk_ActionID = m_Menus(i).me_Name
                .hk_HasChildren = m_Menus(i).me_HasChildren
                .hk_TextEn = m_Menus(i).me_TextEn
                .hk_TextLocalized = m_Menus(i).me_TextTranslated
                
                'If a hotkey exists for this menu's action, retrieve it and add it
                ' (and make backups of these *original* hotkeys, so we can revert them if the user doesn't like later changes)
                Dim idxHotkey As Variant
                If cHotkeys.GetItemByKey(.hk_ActionID, idxHotkey) Then
                    .hk_KeyCode = m_Hotkeys(idxHotkey).hkKeyCode
                    .hk_BackupKeyCode = .hk_KeyCode
                    .hk_ShiftState = m_Hotkeys(idxHotkey).hkShiftState
                    .hk_BackupShiftState = .hk_ShiftState
                    .hk_HotkeyText = m_Menus(i).me_HotKeyTextTranslated
                    .hk_BackupHotkeyText = .hk_HotkeyText
                End If
                
                'Finally, if this is not a top-level menu, retrieve the ID of this menu's *parent* menu
                If (m_Menus(i).me_SubMenu >= 0) Then
                    Dim idxParent As Variant
                    m_MenuHash.GetItemByKey GetMenuParentPositionID(i), idxParent
                    .hk_ParentID = m_Menus(idxParent).me_Name
                    .hk_NumParents = 1
                    If (m_Menus(i).me_SubSubMenu >= 0) Then .hk_NumParents = 2
                End If
                
                'Debug.Print .hk_ActionID, .hk_ParentID, .hk_HasChildren
                
                'Add this menu item to the treeview
                tvMenus.AddItem .hk_ActionID, .hk_TextLocalized, .hk_ParentID, False
                
                'Advance to the next mappable menu index
                m_numItems = m_numItems + 1
                
            End With
            
        '/Ignore separators
        End If
        
    Next i
    
    'Initialize font renderers for the custom treeview
    Set m_FontAllowed = New pdFont
    m_FontAllowed.SetFontBold True
    m_FontAllowed.SetFontSize 12
    m_FontAllowed.CreateFontObject
    m_FontAllowed.SetTextAlignment vbLeftJustify
    
    Set m_FontDisallowed = New pdFont
    m_FontDisallowed.SetFontBold False
    m_FontDisallowed.SetFontSize 12
    m_FontDisallowed.CreateFontObject
    m_FontDisallowed.SetTextAlignment vbLeftJustify
    
    Set m_FontHotkey = New pdFont
    m_FontHotkey.SetFontBold False
    m_FontHotkey.SetFontSize 12
    m_FontHotkey.CreateFontObject
    m_FontHotkey.SetTextAlignment vbLeftJustify
    
    'Add all possible hotkeys to the dropdown
    GeneratePossibleHotkeys
    
    'Apply custom themes
    Interface.ApplyThemeAndTranslations Me
    
    '*Now* allow the treeview to render itself
    m_RenderingOK = True
    tvMenus.SetAutomaticRedraws True, True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Interface.ReleaseFormTheming Me
End Sub

'Return a unique hash table ID for a given menu
Private Function GetMenuPositionID(ByVal idxMenu As Long) As String
    
    Const ID_SEPARATOR As String = "-"
    
    With m_Menus(idxMenu)
        
        If (.me_TopMenu >= 0) Then
            GetMenuPositionID = Trim$(Str$((.me_TopMenu)))
            If (.me_SubMenu >= 0) Then GetMenuPositionID = GetMenuPositionID & ID_SEPARATOR & Trim$(Str$((.me_SubMenu)))
            If (.me_SubSubMenu >= 0) Then GetMenuPositionID = GetMenuPositionID & ID_SEPARATOR & Trim$(Str$((.me_SubSubMenu)))
        Else
            GetMenuPositionID = vbNullString
        End If
        
    End With
    
End Function

'Return a unique hash table ID for a given menu's parent (if one exists).  Returns a null-string if no parent exists.
Private Function GetMenuParentPositionID(ByVal idxMenu As Long) As String

    Const ID_SEPARATOR As String = "-"
    
    With m_Menus(idxMenu)
        If (.me_TopMenu >= 0) Then
            If (.me_SubMenu >= 0) Then GetMenuParentPositionID = Trim$(Str$((.me_TopMenu)))
            If (.me_SubSubMenu >= 0) Then GetMenuParentPositionID = GetMenuParentPositionID & ID_SEPARATOR & Trim$(Str$((.me_SubMenu)))
        End If
    End With
    
End Function

Private Sub tvMenus_Click()
    
    'Failsafe only
    If (tvMenus.ListIndex < 0) Then
        HideEditBox
        Exit Sub
    End If
    
    'Update the previously edited hotkey, if any
    If (m_idxLastHotkey >= 0) And (m_idxLastHotkey <> tvMenus.ListIndex) Then StoreUpdatedHotkey m_idxLastHotkey
    m_inAutoUpdate = True
    
    'Do not allow hotkeys on menu items with children
    Dim hotkeyEditingAllowed As Boolean
    hotkeyEditingAllowed = (Not m_Items(tvMenus.ListIndex).hk_HasChildren)
    
    'Manual hotkey controls at the bottom of the screen mirror the availability of hotkeys for this command
    If (Not hotkeyEditingAllowed) Then
        chkModifier(0).Value = False
        chkModifier(1).Value = False
        chkModifier(2).Value = False
        ddKey.ListIndex = m_idxOtherHotkey
    End If
    
    chkModifier(0).Enabled = hotkeyEditingAllowed
    chkModifier(1).Enabled = hotkeyEditingAllowed
    chkModifier(2).Enabled = hotkeyEditingAllowed
    ddKey.Enabled = hotkeyEditingAllowed
    
    'Note that the user can also disallow auto key capture
    If hotkeyEditingAllowed Then
        
        'Ensure the hotkey editors at the bottom of the screen reflect this item's current hotkey (if any)
        Dim curShiftState As Long, curKeyCode As Long
        curShiftState = m_Items(tvMenus.ListIndex).hk_ShiftState
        curKeyCode = m_Items(tvMenus.ListIndex).hk_KeyCode
        chkModifier(0).Value = (curShiftState And vbCtrlMask) = vbCtrlMask
        chkModifier(1).Value = (curShiftState And vbAltMask) = vbAltMask
        chkModifier(2).Value = (curShiftState And vbShiftMask) = vbShiftMask
        
        Dim i As Long, keyFound As Boolean
        For i = 0 To m_idxOtherHotkey - 1
            If (m_possibleHotkeys(i).ph_VKCode = curKeyCode) Then
                keyFound = True
                ddKey.ListIndex = i
                Exit For
            End If
        Next i
        
        If (Not keyFound) Then ddKey.ListIndex = m_idxOtherHotkey
        
        'Automatic hotkey capture can be toggled by the user
        If Me.chkAutoCapture.Value Then
            
            'To figure out where to position the text box, we need to query the underlying tree support object for details
            Dim tmpTreeSupport As pdTreeSupport
            Set tmpTreeSupport = tvMenus.AccessUnderlyingTreeSupport()
            
            '...including where its child treeview_view is positioned
            Dim lbViewRectF As RectF
            CopyMemoryStrict VarPtr(lbViewRectF), tvMenus.GetListBoxRectFPtr, LenB(lbViewRectF)
            
            '...and the selected treeview item itself
            Dim tmpTreeItem As PD_TreeItem, tmpScrollX As Long, tmpScrollY As Long
            tmpTreeSupport.GetRenderingItem tvMenus.ListIndex, tmpTreeItem, tmpScrollX, tmpScrollY
            
            'Use data from these to figure out where the edit box should go
            Dim ebRectF As RectF
            ebRectF.Left = (tvMenus.GetLeft + ebRectF.Left + tmpTreeItem.captionRect.Left + tmpTreeItem.captionRect.Width) - Interface.FixDPI(200)
            ebRectF.Top = tvMenus.GetTop + ebRectF.Top + tmpTreeItem.captionRect.Top + Interface.FixDPI(3) - tmpScrollY
            ebRectF.Width = Interface.FixDPI(192)
            ebRectF.Height = tmpTreeItem.captionRect.Height - Interface.FixDPI(4)
            
            'Position it and fill it with the hotkey for the current tree item.
            ' (Note that the backup hotkey text *must* be set first - see the edit box _Change event for details)
            m_backupHotkeyText = m_Items(tvMenus.ListIndex).hk_HotkeyText
            Me.txtHotkey.Text = m_backupHotkeyText
            Me.txtHotkey.SetPositionAndSize ebRectF.Left, ebRectF.Top, ebRectF.Width, ebRectF.Height
            Me.txtHotkey.Visible = True
            Me.txtHotkey.ZOrder 0
            Me.txtHotkey.SetFocusToEditBox True
            
        End If
        
        'Note this as the last-edited hotkey, and update all data backups to match
        m_idxLastHotkey = tvMenus.ListIndex
        m_backupHotkeyText = m_Items(tvMenus.ListIndex).hk_HotkeyText
        m_backupHotkeyVKCode = m_Items(tvMenus.ListIndex).hk_KeyCode
        m_backupHotkeyShift = m_Items(tvMenus.ListIndex).hk_ShiftState
        
    End If
    m_inAutoUpdate = False
    
End Sub

'Render an item into the treeview
Private Sub tvMenus_DrawListEntry(ByVal bufferDC As Long, ByVal itemIndex As Long, ByRef itemID As String, ByVal itemIsSelected As Boolean, ByVal itemIsHovered As Boolean, ByVal ptrToItemRectF As Long, ByVal ptrToCaptionRectF As Long, ByVal ptrToControlRectF As Long)
    
    If (bufferDC = 0) Then Exit Sub
    If (Not m_RenderingOK) Then Exit Sub
    
    'Retrieve the boundary region for this list entry
    Dim tmpRectF As RectF
    CopyMemoryStrict VarPtr(tmpRectF), ptrToCaptionRectF, 16&
    
    Dim offsetY As Single, offsetX As Single
    offsetX = tmpRectF.Left
    offsetY = tmpRectF.Top + Interface.FixDPI(1)
    
    'Hotkeys get a fixed (at 96-dpi) 192 pixels to display their key combo.  If menu text overflows this boundary,
    ' it will be truncated with ellipses.
    Dim leftOffsetHotkey As Long
    leftOffsetHotkey = tmpRectF.Left + tmpRectF.Width - (Interface.FixDPI(192))
    
    'If this item has been selected, draw the background with the system's current selection color
    Dim curFont As pdFont
    If m_Items(itemIndex).hk_HasChildren Then Set curFont = m_FontDisallowed Else Set curFont = m_FontAllowed
    
    If itemIsSelected Then
        curFont.SetFontColor g_Themer.GetGenericUIColor(UI_TextClickableSelected)
        m_FontHotkey.SetFontColor g_Themer.GetGenericUIColor(UI_TextClickableSelected)
    Else
        curFont.SetFontColor g_Themer.GetGenericUIColor(UI_TextClickableUnselected, , , itemIsHovered)
        m_FontHotkey.SetFontColor g_Themer.GetGenericUIColor(UI_TextClickableUnselected, , , itemIsHovered)
    End If
    
    'Prepare the rendering text
    Dim drawString As String
    drawString = m_Items(itemIndex).hk_TextLocalized
    
    'Render the text
    If (LenB(drawString) <> 0) Then
        curFont.AttachToDC bufferDC
        curFont.FastRenderTextWithClipping offsetX, offsetY + Interface.FixDPI(4), leftOffsetHotkey - tmpRectF.Left, tmpRectF.Height, drawString, True, False, False
        curFont.ReleaseFromDC
    End If
    
    'Next, solve for the on-screen size of the hotkey text
    Dim keyCodeToUse As Long, keyComboText As String
    keyCodeToUse = m_Items(itemIndex).hk_KeyCode
    keyComboText = m_Items(itemIndex).hk_HotkeyText
    
    If (keyCodeToUse <> 0) Then
        
        'Right-align the hotkey text in the drop-down area, with a little padding
        m_FontHotkey.AttachToDC bufferDC
        m_FontHotkey.FastRenderText leftOffsetHotkey, offsetY + Interface.FixDPI(4), keyComboText
        m_FontHotkey.ReleaseFromDC
        
    End If
    
    'Still TODO:
    ' - figure out where to position hotkey text/input area
    
End Sub

Private Sub tvMenus_ScrollOccurred()
    HideEditBox
End Sub

Private Sub txtHotkey_Change()
    If (txtHotkey.Text <> m_backupHotkeyText) Then txtHotkey.Text = m_backupHotkeyText
End Sub

Private Sub txtHotkey_KeyDown(ByVal Shift As ShiftConstants, ByVal vKey As Long, preventFurtherHandling As Boolean)
    
    m_inAutoUpdate = True
    
    Dim newText As String
    newText = vbNullString
    
    'Build a string for Ctrl/Alt/Shift, and ensure the checkboxes at the bottom reflect the current state
    m_backupHotkeyText = AutoTextKeyChange(Shift, vKey)
    m_backupHotkeyShift = Shift
    m_backupHotkeyVKCode = vKey
    txtHotkey.Text = m_backupHotkeyText
    preventFurtherHandling = True
    
End Sub

Private Sub txtHotkey_KeyUp(ByVal Shift As ShiftConstants, ByVal vKey As Long, preventFurtherHandling As Boolean)
    
    'TODO: DO WE EVEN NEED KEYUP???
    
    'If no modifier keys are down, save this as the current hotkey
    If (Shift And vbCtrlMask = 0) And (Shift And vbAltMask = 0) And (Shift And vbShiftMask = 0) Then
        If (m_backupHotkeyText <> m_Items(tvMenus.ListIndex).hk_HotkeyText) Then
            'm_hotkeys(i).hkScanCode =
        End If
    End If
    
    preventFurtherHandling = True
    m_inAutoUpdate = False
    
End Sub

Private Sub StoreUpdatedHotkey(ByVal idxTarget As Long)
    Debug.Print "storing updated hotkey for: " & m_Items(idxTarget).hk_ActionID
    With m_Items(idxTarget)
        .hk_KeyCode = m_backupHotkeyVKCode
        .hk_HotkeyText = m_backupHotkeyText
        .hk_ShiftState = m_backupHotkeyShift
    End With
End Sub

'Update the bottom "manual" controls to reflect current keystate.
' Returns: a string reflecting the current key state
Private Function AutoTextKeyChange(ByVal Shift As ShiftConstants, ByVal vKey As Long) As String
    
    Dim newText As String
    
    If ((Shift And vbCtrlMask) <> 0) Then
        newText = newText & Hotkeys.GetGenericMenuText(cmt_Ctrl) & "+"
        Me.chkModifier(0).Value = True
    Else
        Me.chkModifier(0).Value = False
    End If
    If ((Shift And vbAltMask) <> 0) Then
        newText = newText & Hotkeys.GetGenericMenuText(cmt_Alt) & "+"
        Me.chkModifier(1).Value = True
    Else
        Me.chkModifier(1).Value = False
    End If
    If ((Shift And vbShiftMask) <> 0) Then
        newText = newText & Hotkeys.GetGenericMenuText(cmt_Shift) & "+"
        Me.chkModifier(2).Value = True
    Else
        Me.chkModifier(2).Value = False
    End If
    
    'Retrieve the system name for this key, but *only* if it's not a Ctrl/Alt/Shift modifier.
    If (vKey <> VK_SHIFT) And (vKey <> VK_ALT) And (vKey <> VK_CONTROL) Then
        
        Dim keyName As String
        keyName = Hotkeys.GetCharFromKeyCode(vKey)
        newText = newText & keyName
        
        Dim i As Long, keyFound As Boolean
        For i = 0 To m_idxOtherHotkey - 1
            If (m_possibleHotkeys(i).ph_VKCode = vKey) Then
                keyFound = True
                Exit For
            End If
        Next i
        
        If keyFound Then Me.ddKey.ListIndex = i Else Me.ddKey.ListIndex = m_idxOtherHotkey
        
    End If
    
    AutoTextKeyChange = newText
    
End Function

Private Sub txtHotkey_LostFocusAPI()
    If txtHotkey.Visible Then txtHotkey.Visible = False
End Sub

'Hide the hotkey edit box (if visible) and optionally, commit any pending hotkey changes the user has entered
Private Sub HideEditBox(Optional ByVal commitChangesFirst As Boolean = False)
    
    'Ignore if the edit box is already invisible (note also that this *skips* committing changes)
    If (Not txtHotkey.Visible) Then Exit Sub
    
    If commitChangesFirst Then
        'TODO
    End If
    
    txtHotkey.Visible = False
    
End Sub

Private Sub GeneratePossibleHotkeys()
    
    'This list is manually generated from https://learn.microsoft.com/en-us/windows/win32/inputdev/virtual-key-codes
    AddPossibleHotkey &H8, "BACKSPACE key"
    AddPossibleHotkey &H9, "TAB key"
    AddPossibleHotkey &HC, "CLEAR key"
    AddPossibleHotkey &HD, "ENTER key"
    AddPossibleHotkey &H13, "PAUSE key"
    AddPossibleHotkey &H14, "CAPS LOCK key"
    AddPossibleHotkey &H1B, "ESC key"
    AddPossibleHotkey &H20, "SPACEBAR"
    AddPossibleHotkey &H21, "PAGE UP key"
    AddPossibleHotkey &H22, "PAGE DOWN key"
    AddPossibleHotkey &H23, "END key"
    AddPossibleHotkey &H24, "HOME key"
    AddPossibleHotkey &H25, "LEFT ARROW key"
    AddPossibleHotkey &H26, "UP ARROW key"
    AddPossibleHotkey &H27, "RIGHT ARROW key"
    AddPossibleHotkey &H28, "DOWN ARROW key"
    AddPossibleHotkey &H29, "SELECT key"
    AddPossibleHotkey &H2A, "PRINT key"
    AddPossibleHotkey &H2B, "EXECUTE key"
    AddPossibleHotkey &H2C, "PRINT SCREEN key"
    AddPossibleHotkey &H2D, "INS key"
    AddPossibleHotkey &H2E, "DEL key"
    AddPossibleHotkey &H2F, "HELP key"
    AddPossibleHotkey &H30, "0 key"
    AddPossibleHotkey &H31, "1 key"
    AddPossibleHotkey &H32, "2 key"
    AddPossibleHotkey &H33, "3 key"
    AddPossibleHotkey &H34, "4 key"
    AddPossibleHotkey &H35, "5 key"
    AddPossibleHotkey &H36, "6 key"
    AddPossibleHotkey &H37, "7 key"
    AddPossibleHotkey &H38, "8 key"
    AddPossibleHotkey &H39, "9 key"
    AddPossibleHotkey &H41, "A key"
    AddPossibleHotkey &H42, "B key"
    AddPossibleHotkey &H43, "C key"
    AddPossibleHotkey &H44, "D key"
    AddPossibleHotkey &H45, "E key"
    AddPossibleHotkey &H46, "F key"
    AddPossibleHotkey &H47, "G key"
    AddPossibleHotkey &H48, "H key"
    AddPossibleHotkey &H49, "I key"
    AddPossibleHotkey &H4A, "J key"
    AddPossibleHotkey &H4B, "K key"
    AddPossibleHotkey &H4C, "L key"
    AddPossibleHotkey &H4D, "M key"
    AddPossibleHotkey &H4E, "N key"
    AddPossibleHotkey &H4F, "O key"
    AddPossibleHotkey &H50, "P key"
    AddPossibleHotkey &H51, "Q key"
    AddPossibleHotkey &H52, "R key"
    AddPossibleHotkey &H53, "S key"
    AddPossibleHotkey &H54, "T key"
    AddPossibleHotkey &H55, "U key"
    AddPossibleHotkey &H56, "V key"
    AddPossibleHotkey &H57, "W key"
    AddPossibleHotkey &H58, "X key"
    AddPossibleHotkey &H59, "Y key"
    AddPossibleHotkey &H5A, "Z key"
    AddPossibleHotkey &H5B, "Left Windows key"
    AddPossibleHotkey &H5C, "Right Windows key"
    AddPossibleHotkey &H5D, "Applications key"
    AddPossibleHotkey &H5F, "Computer Sleep key"
    AddPossibleHotkey &H60, "Numeric keypad 0 key"
    AddPossibleHotkey &H61, "Numeric keypad 1 key"
    AddPossibleHotkey &H62, "Numeric keypad 2 key"
    AddPossibleHotkey &H63, "Numeric keypad 3 key"
    AddPossibleHotkey &H64, "Numeric keypad 4 key"
    AddPossibleHotkey &H65, "Numeric keypad 5 key"
    AddPossibleHotkey &H66, "Numeric keypad 6 key"
    AddPossibleHotkey &H67, "Numeric keypad 7 key"
    AddPossibleHotkey &H68, "Numeric keypad 8 key"
    AddPossibleHotkey &H69, "Numeric keypad 9 key"
    AddPossibleHotkey &H6A, "Multiply key"
    AddPossibleHotkey &H6B, "Add key"
    AddPossibleHotkey &H6C, "Separator key"
    AddPossibleHotkey &H6D, "Subtract key"
    AddPossibleHotkey &H6E, "Decimal key"
    AddPossibleHotkey &H6F, "Divide key"
    AddPossibleHotkey &H70, "F1 key"
    AddPossibleHotkey &H71, "F2 key"
    AddPossibleHotkey &H72, "F3 key"
    AddPossibleHotkey &H73, "F4 key"
    AddPossibleHotkey &H74, "F5 key"
    AddPossibleHotkey &H75, "F6 key"
    AddPossibleHotkey &H76, "F7 key"
    AddPossibleHotkey &H77, "F8 key"
    AddPossibleHotkey &H78, "F9 key"
    AddPossibleHotkey &H79, "F10 key"
    AddPossibleHotkey &H7A, "F11 key"
    AddPossibleHotkey &H7B, "F12 key"
    AddPossibleHotkey &H7C, "F13 key"
    AddPossibleHotkey &H7D, "F14 key"
    AddPossibleHotkey &H7E, "F15 key"
    AddPossibleHotkey &H7F, "F16 key"
    AddPossibleHotkey &H80, "F17 key"
    AddPossibleHotkey &H81, "F18 key"
    AddPossibleHotkey &H82, "F19 key"
    AddPossibleHotkey &H83, "F20 key"
    AddPossibleHotkey &H84, "F21 key"
    AddPossibleHotkey &H85, "F22 key"
    AddPossibleHotkey &H86, "F23 key"
    AddPossibleHotkey &H87, "F24 key"
    AddPossibleHotkey &H90, "NUM LOCK key"
    AddPossibleHotkey &H91, "SCROLL LOCK key"
    AddPossibleHotkey &H92, "OEM specific"
    AddPossibleHotkey &H93, "OEM specific"
    AddPossibleHotkey &H94, "OEM specific"
    AddPossibleHotkey &H95, "OEM specific"
    AddPossibleHotkey &H96, "OEM specific"
    AddPossibleHotkey &HA6, "Browser Back key"
    AddPossibleHotkey &HA7, "Browser Forward key"
    AddPossibleHotkey &HA8, "Browser Refresh key"
    AddPossibleHotkey &HA9, "Browser Stop key"
    AddPossibleHotkey &HAA, "Browser Search key"
    AddPossibleHotkey &HAB, "Browser Favorites key"
    AddPossibleHotkey &HAC, "Browser Start and Home key"
    AddPossibleHotkey &HAD, "Volume Mute key"
    AddPossibleHotkey &HAE, "Volume Down key"
    AddPossibleHotkey &HAF, "Volume Up key"
    AddPossibleHotkey &HB0, "Next Track key"
    AddPossibleHotkey &HB1, "Previous Track key"
    AddPossibleHotkey &HB2, "Stop Media key"
    AddPossibleHotkey &HB3, "Play/Pause Media key"
    AddPossibleHotkey &HB4, "Start Mail key"
    AddPossibleHotkey &HB5, "Select Media key"
    AddPossibleHotkey &HB6, "Start Application 1 key"
    AddPossibleHotkey &HB7, "Start Application 2 key"
    AddPossibleHotkey &HBA, "Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the ;: key"
    AddPossibleHotkey &HBB, "For any country/region, the + key"
    AddPossibleHotkey &HBC, "For any country/region, the , key"
    AddPossibleHotkey &HBD, "For any country/region, the - key"
    AddPossibleHotkey &HBE, "For any country/region, the . key"
    AddPossibleHotkey &HBF, "Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the /? key"
    AddPossibleHotkey &HC0, "Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the `~ key"
    AddPossibleHotkey &HDB, "Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the [{ key"
    AddPossibleHotkey &HDC, "Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the \\| key"
    AddPossibleHotkey &HDD, "Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the ]} key"
    AddPossibleHotkey &HDE, "Used for miscellaneous characters; it can vary by keyboard. For the US standard keyboard, the '""key"
    AddPossibleHotkey &HDF, "Used for miscellaneous characters; it can vary by keyboard."
    AddPossibleHotkey &HE1, "OEM specific"
    AddPossibleHotkey &HE2, "The <> keys on the US standard keyboard, or the \\| key on the non-US 102-key keyboard"
    AddPossibleHotkey &HE3, "OEM specific"
    AddPossibleHotkey &HE4, "OEM specific"
    AddPossibleHotkey &HE6, "OEM specific"
    AddPossibleHotkey &HE9, "OEM specific"
    AddPossibleHotkey &HEA, "OEM specific"
    AddPossibleHotkey &HEB, "OEM specific"
    AddPossibleHotkey &HEC, "OEM specific"
    AddPossibleHotkey &HED, "OEM specific"
    AddPossibleHotkey &HEE, "OEM specific"
    AddPossibleHotkey &HEF, "OEM specific"
    AddPossibleHotkey &HF1, "OEM specific"
    AddPossibleHotkey &HF2, "OEM specific"
    AddPossibleHotkey &HF3, "OEM specific"
    AddPossibleHotkey &HF4, "OEM specific"
    AddPossibleHotkey &HF5, "OEM specific"
    AddPossibleHotkey &HF6, "Attn key"
    AddPossibleHotkey &HF7, "CrSel key"
    AddPossibleHotkey &HF8, "ExSel key"
    AddPossibleHotkey &HFA, "Play key"
    AddPossibleHotkey &HFB, "Zoom key"
    AddPossibleHotkey &HFD, "PA1 key"
    AddPossibleHotkey &HFE, "Clear key"
    
    'Do a quick insertion sort.  NAmes Points are likely to be somewhat close to sorted, as e.g. A-Z are added in order.
    Dim tmpSortKey As PD_PossibleHotkey, searchCont As Boolean
    
    Dim i As Long, j As Long
    i = 1
    
    Do While (i < m_numPossibleHotkeys)
        tmpSortKey = m_possibleHotkeys(i)
        j = i - 1
        
        'Because VB6 doesn't short-circuit And statements, we split this check into separate parts.
        searchCont = False
        If (j >= 0) Then searchCont = (Strings.StrCompSortPtr(StrPtr(m_possibleHotkeys(j).ph_KeyName), StrPtr(tmpSortKey.ph_KeyName)) > 0)
        
        Do While searchCont
            m_possibleHotkeys(j + 1) = m_possibleHotkeys(j)
            j = j - 1
            searchCont = False
            If (j >= 0) Then searchCont = (Strings.StrCompSortPtr(StrPtr(m_possibleHotkeys(j).ph_KeyName), StrPtr(tmpSortKey.ph_KeyName)) > 0)
        Loop
        
        m_possibleHotkeys(j + 1) = tmpSortKey
        i = i + 1
        
    Loop
    
    'Manually add an "other" key to the list, which we'll use for miscellaneous keypresses that the keyboard driver can't name
    AddPossibleHotkey &HFF, "(other)", "(other)"
    m_idxOtherHotkey = m_numPossibleHotkeys - 1
    
    ReDim Preserve m_possibleHotkeys(0 To m_numPossibleHotkeys - 1) As PD_PossibleHotkey
    
    'Add all items to the on-screen dropdown
    For i = 0 To m_numPossibleHotkeys - 1
        ddKey.AddItem m_possibleHotkeys(i).ph_KeyName
    Next i
    
    'Select the (other) entry
    ddKey.ListIndex = m_idxOtherHotkey
    
End Sub

Private Sub AddPossibleHotkey(ByVal vkCode As Long, Optional ByRef keyComments As String = vbNullString, Optional ByRef manualKeyName As String = vbNullString)
    
    If (m_numPossibleHotkeys = 0) Then
        Const INIT_POSSIBLE_HOTKEYS As Long = 64
        ReDim m_possibleHotkeys(0 To INIT_POSSIBLE_HOTKEYS) As PD_PossibleHotkey
    End If
    
    If (g_Language Is Nothing) Then Exit Sub
    
    'See if this key exists on this keyboard (null names mean key doesn't exist, typically)
    Dim keyName As String, keyNameExtended As String
    If (LenB(manualKeyName) > 0) Then
        keyName = manualKeyName
    Else
        
        Select Case vkCode
            
            'Some unreadable chars have to be manually entered
            Case 8
                keyName = g_Language.TranslateMessage("Backspace")
            Case 9
                keyName = g_Language.TranslateMessage("Tab")
            Case &H1B
                keyName = g_Language.TranslateMessage("Escape")
            
            'Other ones can be pulled from the keyboard driver
            Case Else
                Hotkeys.GetCharFromKeyCode vkCode, outKeyName:=keyName, outKeyNameExtended:=keyNameExtended
                
        End Select
        
    End If
    
    'Ignore blank names (those are likely keys that do not exist)
    If (LenB(keyName) > 0) Then
        
        'Iterate previous entries and skip duplicates.  (OEMs may use OEM-specific keycodes to duplicate standard keycodes.)
        If (m_numPossibleHotkeys > 0) Then
            Dim i As Long
            For i = 0 To m_numPossibleHotkeys - 1
                If Strings.StringsEqual(keyName, m_possibleHotkeys(i).ph_KeyName) Then
                    m_possibleHotkeys(i).ph_VKCode = PDMath.Min2Int(vkCode, m_possibleHotkeys(i).ph_VKCode)
                    Exit Sub
                End If
            Next i
        End If
        
        With m_possibleHotkeys(m_numPossibleHotkeys)
            .ph_VKCode = vkCode
            .ph_KeyName = keyName
            .ph_KeyComments = keyComments
        End With
        
        m_numPossibleHotkeys = m_numPossibleHotkeys + 1
        If (m_numPossibleHotkeys > UBound(m_possibleHotkeys)) Then ReDim Preserve m_possibleHotkeys(0 To m_numPossibleHotkeys * 2 - 1) As PD_PossibleHotkey
        
    End If
    
    'Repeat previous steps for extended key name
    If (LenB(keyNameExtended) > 0) Then
        
        If (m_numPossibleHotkeys > 0) Then
            For i = 0 To m_numPossibleHotkeys - 1
                If Strings.StringsEqual(keyNameExtended, m_possibleHotkeys(i).ph_KeyName) Then
                    m_possibleHotkeys(i).ph_VKCode = PDMath.Min2Int(vkCode, m_possibleHotkeys(i).ph_VKCode)
                    Exit Sub
                End If
            Next i
        End If
        
        With m_possibleHotkeys(m_numPossibleHotkeys)
            .ph_VKCode = vkCode
            .ph_KeyName = keyNameExtended
            .ph_KeyComments = keyComments
        End With
        
        m_numPossibleHotkeys = m_numPossibleHotkeys + 1
        If (m_numPossibleHotkeys > UBound(m_possibleHotkeys)) Then ReDim Preserve m_possibleHotkeys(0 To m_numPossibleHotkeys * 2 - 1) As PD_PossibleHotkey
        
    End If
    
End Sub
