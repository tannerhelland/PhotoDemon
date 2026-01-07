VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdPopupMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Pop-up Menu Manager
'Copyright 2019-2026 by Tanner Helland
'Created: 29/May/19
'Last updated: 16/December/23
'Last update: total overhaul (added submenu support and refactored pretty much everything)
'
'Like all other built-in VB6 UI elements, popup menus do not natively support Unicode.  This class provides a
' custom solution by interacting directly with WAPI (user32), specifically TrackPopupMenu.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Public Event MenuClicked(ByRef clickedMenuID As String, ByVal idxMenuTop As Long, ByVal idxMenuSub As Long)

Private Type PopupMenuItem
    pmCaption As String             'It is the caller's responsibility to translate captions *before* adding them
    pmID As String                  'User-passed ID, typically maps to an internal PD action ID
    pmEnabled As Boolean            'Is this menu item enabled?
    pmChecked As Boolean            'Does this menu display a checkmark?
    pmIdxTop As Long                'Top-level menu index
    pmIdxSub As Long                'Submenu index (0-based, initializes to -1 for top-level items)
    hMenuIfChildrenExist As Long    'Non-zero value if this top-level item pops up a submenu (populated automatically!)
End Type

Private Enum WapiMenuFlags
    MF_BITMAP = &H4&
    MF_CHECKED = &H8&
    MF_DISABLED = &H2&
    MF_ENABLED = &H0&
    MF_GRAYED = &H1&
    MF_MENUBARBREAK = &H20&
    MF_MENUBREAK = &H40&
    MF_OWNERDRAW = &H100&
    MF_POPUP = &H10&
    MF_SEPARATOR = &H800&
    MF_STRING = &H0&
    MF_UNCHECKED = &H0&
End Enum

#If False Then
    Private Const MF_BITMAP = &H4&, MF_CHECKED = &H8&, MF_DISABLED = &H2&, MF_ENABLED = &H0&, MF_GRAYED = &H1&, MF_MENUBARBREAK = &H20&, MF_MENUBREAK = &H40&, MF_OWNERDRAW = &H100&, MF_POPUP = &H10&, MF_SEPARATOR = &H800&, MF_STRING = &H0&, MF_UNCHECKED = &H0&
#End If

Private Declare Function AppendMenuW Lib "user32" (ByVal hMenu As Long, ByVal uFlags As WapiMenuFlags, ByVal uIDNewItem As Long, ByVal lpNewItem As Long) As Long
Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal srcHWnd As Long, ByVal prcRect As Long) As Long

Private Const MENU_NONE As Long = -1

Private m_NumMenuItems As Long
Private m_MenuItems() As PopupMenuItem
Private Const NUM_MENU_ITEMS_START As Long = 8

'Add a new menu item to this popup menu.
' - mnuCaptionTranslated: per the name, the item caption ALREADY LOCALIZED
' - [optional] menuID: unique string identifying this item (will be returned by MenuClicked event)
' - [optional] idxMenuTop: top-level index which does *not* need to be passed; passed order is assumed as menu order when missing
' - [optional] idxMenuSub: sub-menu index; assumed to be a top-level menu item when missing
' - [optional] menuIsEnabled: self-explanatory
Friend Sub AddMenuItem(ByRef menuCaptionTranslated As String, Optional ByRef menuID As String = vbNullString, Optional ByVal idxMenuTop As Long = MENU_NONE, Optional ByVal idxMenuSub As Long = MENU_NONE, Optional ByVal menuIsEnabled As Boolean = True, Optional ByVal menuIsChecked As Boolean = False)
    If (m_NumMenuItems > UBound(m_MenuItems)) Then ReDim Preserve m_MenuItems(0 To m_NumMenuItems * 2 - 1) As PopupMenuItem
    With m_MenuItems(m_NumMenuItems)
        .pmCaption = menuCaptionTranslated
        .pmID = menuID
        If (idxMenuTop = MENU_NONE) Then
            .pmIdxTop = m_NumMenuItems
            .pmIdxSub = MENU_NONE       'Sub-menu index is always ignored when order is implicit
        Else
            .pmIdxTop = idxMenuTop
            .pmIdxSub = idxMenuSub
        End If
        .pmEnabled = menuIsEnabled
        .pmChecked = menuIsChecked
    End With
    m_NumMenuItems = m_NumMenuItems + 1
End Sub

Friend Sub Reset()
    m_NumMenuItems = 0
    ReDim m_MenuItems(0 To NUM_MENU_ITEMS_START - 1) As PopupMenuItem
End Sub

'Display the menu.  Listen for a MenuClicked() event to retrieve the clicked entry (if any).
Friend Sub ShowMenu(ByVal srcHWnd As Long, ByVal srcX As Long, ByVal srcY As Long)
    
    'Translate the source x/y coords into screen coords
    Dim srcPoint As PointAPI
    srcPoint.x = srcX
    srcPoint.y = srcY
    g_WindowManager.GetClientToScreen srcHWnd, srcPoint
    
    'Create an API menu object.  The returned value is the parent handle that contains all other submenus.
    Dim hMenu As Long
    hMenu = CreateMenu()
    
    'Raise the menu and wait for a response
    Const TPM_RETURNCMD As Long = &H100&, TPM_NONOTIFY As Long = &H80&
    
    Dim hRet As Long
    hRet = TrackPopupMenu(hMenu, TPM_NONOTIFY Or TPM_RETURNCMD, srcPoint.x, srcPoint.y, 0&, srcHWnd, 0&)
    If (hRet <> 0) Then
        
        'We add 1 to all IDs to differentiate them from a null return
        hRet = hRet - 1
        
        'Return the clicked ID and associated indices
        RaiseEvent MenuClicked(m_MenuItems(hRet).pmID, m_MenuItems(hRet).pmIdxTop, m_MenuItems(hRet).pmIdxSub)
        
    End If
    
    'Destroy the menu object before exiting, and note that we don't need to destroy child menus per
    ' https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-destroymenu, which says...
    ' "DestroyMenu is recursive, that is, it will destroy the menu and all its submenus."
    DestroyMenu hMenu

End Sub

'Create a novel API menu using current menu settings
Private Function CreateMenu() As Long
    
    'Menu separators use a specific caption
    Const MENU_SEPARATOR As String = "-"
    Const MSG_APPEND_FAIL As String = "WARNING: pdPopupMenu.CreateMenu failed to append menu #"
    
    'Construct the initial menu handle
    Dim hMenu As Long
    hMenu = CreatePopupMenu()
    
    'Next, we need to build submenus.  Start by scanning the menu list and looking for menu entries with a parent menu.
    ' Build a list of these, and we will construct each submenu separately.
    Dim listOfSubmenus As pdStack
    Set listOfSubmenus = New pdStack
    
    Dim i As Long, j As Long
    For i = 0 To m_NumMenuItems - 1
        If (m_MenuItems(i).pmIdxSub >= 0) Then
            If (listOfSubmenus.DoesIntExist(m_MenuItems(i).pmIdxTop) < 0) Then listOfSubmenus.AddInt m_MenuItems(i).pmIdxTop
        End If
    Next i
    
    'Menu item states like enabled, checked, etc are set via flags
    Dim mnuFlags As WapiMenuFlags
    
    'Only construct submenus as necessary
    If (listOfSubmenus.GetNumOfInts > 0) Then
        
        'Handle each top-level menu one-at-a-time
        For i = 0 To listOfSubmenus.GetNumOfInts - 1
            
            'Reset menu flags
            mnuFlags = 0
            
            'Translate the index to the top-level menu to an index into the main menu item array
            Dim idxParentInMenu As Long, idxParentInArray As Long
            idxParentInMenu = listOfSubmenus.GetInt(i)
            
            For j = 0 To m_NumMenuItems - 1
                If (m_MenuItems(j).pmIdxTop = idxParentInMenu) And (m_MenuItems(j).pmIdxSub = MENU_NONE) Then
                    idxParentInArray = j
                    Exit For
                End If
            Next j
            
            'Get an hMenu handle for the parent menu
            Dim hMenuParent As Long
            hMenuParent = CreatePopupMenu()
            m_MenuItems(idxParentInArray).hMenuIfChildrenExist = hMenuParent
            
            'Find all child items of this submenu and append them in turn
            For j = 0 To m_NumMenuItems - 1
                If (m_MenuItems(j).pmIdxTop = idxParentInMenu) And (m_MenuItems(j).pmIdxSub >= 0) Then
                    
                    'Separator bars are handled manually
                    If (m_MenuItems(j).pmCaption = MENU_SEPARATOR) Then
                        If (AppendMenuW(hMenuParent, MF_SEPARATOR, j + 1, 0) = 0) Then PDDebug.LogAction MSG_APPEND_FAIL & i
                    
                    'Regular string menu item
                    Else
                        If m_MenuItems(j).pmEnabled Then
                            mnuFlags = MF_ENABLED
                            If m_MenuItems(j).pmChecked Then mnuFlags = mnuFlags Or MF_CHECKED Else mnuFlags = mnuFlags Or MF_UNCHECKED
                        Else
                            mnuFlags = MF_GRAYED Or MF_UNCHECKED
                        End If
                        mnuFlags = mnuFlags Or MF_STRING
                        If (AppendMenuW(hMenuParent, mnuFlags, j + 1, StrPtr(m_MenuItems(j).pmCaption)) = 0) Then PDDebug.LogAction MSG_APPEND_FAIL & i
                    End If
                    
                End If
            Next j
            
        Next i
        
    End If
    
    'Add each menu item in turn
    For i = 0 To m_NumMenuItems - 1
        
        'We only want to add top-level menu items here
        If (m_MenuItems(i).pmIdxSub = MENU_NONE) Then
            
            'Reset menu flags
            mnuFlags = 0
            
            'Look for items that raise submenus
            If (m_MenuItems(i).hMenuIfChildrenExist <> 0) Then
                
                If m_MenuItems(i).pmEnabled Then mnuFlags = MF_ENABLED Else mnuFlags = MF_GRAYED
                mnuFlags = mnuFlags Or MF_POPUP
                If (AppendMenuW(hMenu, mnuFlags, m_MenuItems(i).hMenuIfChildrenExist, StrPtr(m_MenuItems(i).pmCaption)) = 0) Then PDDebug.LogAction MSG_APPEND_FAIL & i
                
            'This is a normal menu item
            Else
                
                'Separator bars are handled manually
                If (m_MenuItems(i).pmCaption = MENU_SEPARATOR) Then
                    If (AppendMenuW(hMenu, MF_SEPARATOR, i + 1, 0) = 0) Then PDDebug.LogAction MSG_APPEND_FAIL & i
                
                'Regular string menu item
                Else
                    If m_MenuItems(i).pmEnabled Then
                        mnuFlags = MF_ENABLED
                        If m_MenuItems(i).pmChecked Then mnuFlags = mnuFlags Or MF_CHECKED Else mnuFlags = mnuFlags Or MF_UNCHECKED
                    Else
                        mnuFlags = MF_GRAYED Or MF_UNCHECKED
                    End If
                    mnuFlags = mnuFlags Or MF_STRING
                    If (AppendMenuW(hMenu, mnuFlags, i + 1, StrPtr(m_MenuItems(i).pmCaption)) = 0) Then PDDebug.LogAction MSG_APPEND_FAIL & i
                End If
            
            '/end submenu parent vs normal menu item branch
            End If
        
        'Ignore submenu items completely (no else required)
        'Else
        End If
        
    Next i
    
    CreateMenu = hMenu
    
End Function

Private Sub Class_Initialize()
    Me.Reset
End Sub
