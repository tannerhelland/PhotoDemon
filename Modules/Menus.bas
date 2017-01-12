Attribute VB_Name = "Menus"
'***************************************************************************
'Specialized Math Routines
'Copyright 2017-2017 by Tanner Helland
'Created: 11/January/17
'Last updated: 11/January/17
'Last update: initial build
'
'PhotoDemon has an extensive menu system.  Managing all those menus is a cumbersome task.  This module exists
' to tackle the worst parts of run-time maintenance, so other functions don't need to.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************


Option Explicit

Private Type PD_MenuEntry
    ME_Name As String
    ME_ResImage As String
    ME_TopMenu As Long
    ME_SubMenu As Long
    ME_SubSubMenu As Long
End Type

Private m_Menus() As PD_MenuEntry
Private m_NumOfMenus As Long

'Early in the PD load process, we initialize the default set of menus.  In the future, it may be nice to let
' users customize this to match their favorite software (e.g. PhotoShop), but that's a ways off as I've yet to
' build a menu control capable of that level of customization support.
Public Sub InitializeMenus()
    
    'File Menu
    AddMenuItem "file_new", 0, 0, , "file_new"              'New
    AddMenuItem "file_open", 0, 1, , "file_open"            'Open Image
    AddMenuItem "file_openrecent", 0, 2                     'Open recent
    AddMenuItem "file_import", 0, 3                         'Import
        
        '--> Import sub-menu
        AddMenuItem "file_import_paste", 0, 3, 0, "file_importclipboard"    'From Clipboard (Paste as New Image)
        AddMenuItem "file_import_scanner", 0, 3, 2, "file_importscanner"    'Scan Image
        AddMenuItem "file_import_selectscanner", 0, 3, 3                    'Select Scanner
        AddMenuItem "file_import_web", 0, 3, 5, "file_importweb"            'Online Image
        AddMenuItem "file_import_screenshot", 0, 3, 7, "file_importscreen"  'Screen Capture
    
    AddMenuItem "file_close", 0, 5, , "file_close"          'Close
    AddMenuItem "file_save", 0, 8, , "file_save"            'Save
    AddMenuItem "file_savecopy", 0, 9, , "file_savedup"     'Save copy
    AddMenuItem "file_saveas", 0, 10, , "file_saveas"       'Save As...
    AddMenuItem "file_revert", 0, 11                        'Revert
    AddMenuItem "file_batch", 0, 13, , "file_batch"         'Batch operations
    
        '--> Batch sub-menu
        AddMenuItem "file_batch_process", 0, 13, 0, "file_batch"   'Batch process
        AddMenuItem "file_batch_repair", 0, 13, 1, "file_repair"   'Batch repair
        
    AddMenuItem "file_print", 0, 15, , "file_print"         'Print
    AddMenuItem "file_quit", 0, 17                          'Exit

End Sub

Private Sub AddMenuItem(ByRef menuName As String, ByVal topMenuID As Long, Optional ByVal subMenuID As Long = -1, Optional ByVal subSubMenuID As Long = -1, Optional ByRef menuImageName As String = vbNullString)
    
    'Make sure a sufficiently large buffer exists for this menu item
    Const INITIAL_MENU_COLLECTION_SIZE As Long = 64
    If (m_NumOfMenus = 0) Then
        ReDim m_Menus(0 To INITIAL_MENU_COLLECTION_SIZE - 1) As PD_MenuEntry
    Else
        If (m_NumOfMenus > UBound(m_Menus)) Then ReDim Preserve m_Menus(0 To m_NumOfMenus * 2 - 1) As PD_MenuEntry
    End If
    
    With m_Menus(m_NumOfMenus)
        .ME_Name = menuName
        .ME_TopMenu = topMenuID
        .ME_SubMenu = subMenuID
        .ME_SubSubMenu = subSubMenuID
        .ME_ResImage = menuImageName
    End With
    
    m_NumOfMenus = m_NumOfMenus + 1

End Sub

'*After* all menus have been initialized, you can call this function to apply their associated icons (if any)
' to the respective menu objects.
Public Sub ApplyIconsToMenus()

    Dim i As Long
    For i = 0 To m_NumOfMenus - 1
        If (Len(m_Menus(i).ME_ResImage) <> 0) Then
            With m_Menus(i)
                Icons_and_Cursors.AddMenuIcon .ME_ResImage, .ME_TopMenu, .ME_SubMenu, .ME_SubSubMenu
            End With
        End If
    Next i

End Sub
