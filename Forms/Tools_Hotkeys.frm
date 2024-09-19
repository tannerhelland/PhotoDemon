VERSION 5.00
Begin VB.Form FormHotkeys 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   " Keyboard shortcuts"
   ClientHeight    =   6045
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
   ScaleHeight     =   403
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   784
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdTreeviewOD tvMenus 
      Height          =   5055
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   8916
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   5310
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   1296
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
End Type

'Menu and hotkey information gets merged into this local array, which is much easier to manage
' against the UI of this dialog
Private m_Items() As PD_HotkeyUI, m_numItems As Long

'Height of each list item in the custom-drawn treeview, in pixels, at 96 DPI
Private Const BLOCKHEIGHT As Long = 32

'Two font objects; one for menus that are allowed to have hotkeys, and one for menus that are not
' (e.g. top-level menus or parent-only menus).
Private m_FontAllowed As pdFont, m_FontDisallowed As pdFont

'All rendering is suspended until the form is loaded
Private m_RenderingOK As Boolean

Private Sub cmdBar_AddCustomPresetData()
    'TODO: save entire hotkey list - this would let the user swap between e.g. "GIMP" and "Photoshop" presets?
End Sub

Private Sub cmdBar_OKClick()
    
    'TODO
    
End Sub

Private Sub cmdBar_ResetClick()
    'TODO
End Sub

Private Sub Form_Load()
    
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
                Dim idxHotkey As Variant
                If cHotkeys.GetItemByKey(.hk_ActionID, idxHotkey) Then
                    .hk_KeyCode = m_Hotkeys(idxHotkey).hkKeyCode
                    .hk_ShiftState = m_Hotkeys(idxHotkey).hkShiftState
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
    
    'If this item has been selected, draw the background with the system's current selection color
    Dim curFont As pdFont
    If m_Items(itemIndex).hk_HasChildren Then Set curFont = m_FontDisallowed Else Set curFont = m_FontAllowed
    
    If itemIsSelected Then
        curFont.SetFontColor g_Themer.GetGenericUIColor(UI_TextClickableSelected)
    Else
        curFont.SetFontColor g_Themer.GetGenericUIColor(UI_TextClickableUnselected, , , itemIsHovered)
    End If
    
    'Prepare the rendering text
    Dim drawString As String
    drawString = m_Items(itemIndex).hk_TextLocalized
    
    'Render the text
    If (LenB(drawString) <> 0) Then
        curFont.AttachToDC bufferDC
        curFont.FastRenderText offsetX, offsetY + Interface.FixDPI(4), drawString
        curFont.ReleaseFromDC
    End If
    
    'Still TODO:
    ' - figure out where to position hotkey text/input area
    
End Sub
