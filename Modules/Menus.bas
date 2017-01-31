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
    
    'Edit menu
    AddMenuItem "edit_undo", 1, 0, , "edit_undo"            'Undo
    AddMenuItem "edit_redo", 1, 1, , "edit_redo"            'Redo
    AddMenuItem "edit_history", 1, 2, , "edit_history"      'Undo history browser
    
    AddMenuItem "edit_repeat", 1, 4, , "edit_repeat"        'Repeat previous action
    AddMenuItem "edit_fade", 1, 5                           'Fade previous action...
    
    AddMenuItem "edit_cut", 1, 7, , "edit_cut"              'Cut
    AddMenuItem "edit_cutlayer", 1, 8                       'Cut from layer
    AddMenuItem "edit_copy", 1, 9, , "edit_copy"            'Copy
    AddMenuItem "edit_copylayer", 1, 10                     'Copy from layer
    AddMenuItem "edit_pasteasimage", 1, 11, , "edit_paste"  'Paste as new image
    AddMenuItem "edit_pasteaslayer", 1, 12                  'Paste as new layer
    AddMenuItem "edit_emptyclipboard", 1, 14                'Empty Clipboard
    
    'View Menu
    AddMenuItem "zoom_fit", 2, 0, , "zoom_fit"              'Fit on Screen
    
    AddMenuItem "zoom_in", 2, 2, , "zoom_in"                'Zoom In
    AddMenuItem "zoom_out", 2, 3, , "zoom_out"              'Zoom Out
    
    AddMenuItem "zoom_16_1", 2, 5                           'Zoom 16:1
    AddMenuItem "zoom_8_1", 2, 6                            'Zoom 8:1
    AddMenuItem "zoom_4_1", 2, 7                            'Zoom 4:1
    AddMenuItem "zoom_2_1", 2, 8                            'Zoom 2:1
    AddMenuItem "zoom_actual", 2, 9, , "zoom_actual"        'Zoom 100%
    AddMenuItem "zoom_1_2", 2, 10                           'Zoom 1:2
    AddMenuItem "zoom_1_4", 2, 11                           'Zoom 1:4
    AddMenuItem "zoom_1_8", 2, 12                           'Zoom 1:8
    AddMenuItem "zoom_1_16", 2, 13                          'Zoom 1:16
    
    'Tools Menu
    AddMenuItem "tools_language", 8, 0, , "tools_language"  'Languages
    AddMenuItem "tools_languageeditor", 8, 1                'Language editor
    
    AddMenuItem "tools_macrotop", 8, 3, , "macro_record"    'Macros
    
        '--> Macro sub-menu
        AddMenuItem "tools_recordmacro", 8, 3, 0, "macro_record" 'Start Recording
        AddMenuItem "tools_stopmacro", 8, 3, 1, "macro_stop"     'Stop Recording
        
    AddMenuItem "tools_playmacro", 8, 4, , "macro_play"     'Play saved macro
    AddMenuItem "tools_recentmacros", 8, 5                  'Recent macros
    
    AddMenuItem "tools_options", 8, 7, , "pref_advanced"    'Options (Preferences)
    AddMenuItem "tools_plugins", 8, 8, , "tools_plugin"     'Plugin Manager
    
    'Window Menu
    AddMenuItem "window_next", 9, 5, , "generic_next"          'Next image
    AddMenuItem "window_previous", 9, 6, , "generic_previous"  'Previous image
    
    'Help Menu
    AddMenuItem "help_donate", 10, 0, , "help_heart"        'Donate
    AddMenuItem "help_checkupdates", 10, 2, , "help_update" 'Check for updates
    AddMenuItem "help_contact", 10, 3, , "help_contact"     'Submit Feedback
    AddMenuItem "help_reportbug", 10, 4, , "help_reportbug" 'Submit Bug
    AddMenuItem "help_website", 10, 6, , "help_website"     'Visit the PhotoDemon website
    AddMenuItem "help_sourcecode", 10, 7, , "help_github"   'Download source code
    AddMenuItem "help_license", 10, 8, , "help_license"     'License
    AddMenuItem "help_about", 10, 10, , "help_about"        'About PD
    
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
