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
'Copyright 2019-2022 by Tanner Helland
'Created: 29/May/19
'Last updated: 29/May/19
'Last update: initial build
'
'As with nearly all built-in VB6 UI elements, popup menus do not support Unicode.  A custom solution
' is required.
'
'This class is designed to make Unicode-aware popup menus easier.  I do not guarantee that it will
' always rely on built-in WAPI menu functions.  It may (someday) be migrated to a custom, owner-drawn
' version instead.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Public Event MenuClicked(ByVal mnuIndex As Long, ByRef clickedMenuCaption As String)

Private Type PopupMenuItem
    pmCaption As String     'It is the caller's responsibility to translate captions *before* adding them
    pmEnabled As Boolean
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
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal srcHwnd As Long, ByVal prcRect As Long) As Long

Private m_NumMenuItems As Long
Private m_MenuItems() As PopupMenuItem
Private Const NUM_MENU_ITEMS_START As Long = 8

Friend Sub AddMenuItem(ByRef mnuCaptionTranslated As String, Optional ByVal mnuIsEnabled As Boolean = True)
    If (m_NumMenuItems > UBound(m_MenuItems)) Then ReDim Preserve m_MenuItems(0 To m_NumMenuItems * 2 - 1) As PopupMenuItem
    With m_MenuItems(m_NumMenuItems)
        .pmCaption = mnuCaptionTranslated
        .pmEnabled = mnuIsEnabled
    End With
    m_NumMenuItems = m_NumMenuItems + 1
End Sub

Friend Sub Reset()
    m_NumMenuItems = 0
    ReDim m_MenuItems(0 To NUM_MENU_ITEMS_START - 1) As PopupMenuItem
End Sub

'Display the menu.  Listen for a MenuClicked() event to retrieve the clicked entry (if any).
Friend Sub ShowMenu(ByVal srcHwnd As Long, ByVal srcX As Long, ByVal srcY As Long)
    
    'Translate the source x/y coords into screen coords
    Dim srcPoint As PointAPI
    srcPoint.x = srcX
    srcPoint.y = srcY
    g_WindowManager.GetClientToScreen srcHwnd, srcPoint
    
    'Create an API menu object
    Dim hMenu As Long
    hMenu = CreateMenu()
    
    'Raise the menu and wait for a response
    Const TPM_RETURNCMD As Long = &H100&, TPM_NONOTIFY As Long = &H80&
    
    Dim hRet As Long
    hRet = TrackPopupMenu(hMenu, TPM_NONOTIFY Or TPM_RETURNCMD, srcPoint.x, srcPoint.y, 0&, srcHwnd, 0&)
    If (hRet <> 0) Then RaiseEvent MenuClicked(hRet - 1, m_MenuItems(hRet - 1).pmCaption)
    
    'Destroy the menu object before exiting
    DestroyMenu hMenu

End Sub

'Create a novel API menu using current menu settings
Private Function CreateMenu() As Long

    'Construct the initial menu handle
    Dim hMenu As Long
    hMenu = CreatePopupMenu()
    
    'Add each menu item in turn
    Dim i As Long
    For i = 0 To m_NumMenuItems - 1
        
        'Separator bars are handled manually
        If (m_MenuItems(i).pmCaption = "-") Then
            If (AppendMenuW(hMenu, MF_SEPARATOR, i + 1, 0) = 0) Then PDDebug.LogAction "WARNING: pdPopupMenu.ShowMenu failed to append menu #" & i
        
        'Regular string menu item
        Else
            Dim mnuFlags As WapiMenuFlags
            If m_MenuItems(i).pmEnabled Then mnuFlags = MF_ENABLED Else mnuFlags = MF_GRAYED
            mnuFlags = mnuFlags Or MF_STRING Or MF_UNCHECKED
            If (AppendMenuW(hMenu, mnuFlags, i + 1, StrPtr(m_MenuItems(i).pmCaption)) = 0) Then PDDebug.LogAction "WARNING: pdPopupMenu.ShowMenu failed to append menu #" & i
        End If
        
    Next i
    
    CreateMenu = hMenu
    
End Function

Private Sub Class_Initialize()
    Me.Reset
End Sub
