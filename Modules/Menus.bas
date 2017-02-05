Attribute VB_Name = "Menus"
'***************************************************************************
'Specialized Math Routines
'Copyright 2017-2017 by Tanner Helland
'Created: 11/January/17
'Last updated: 05/February/17
'Last update: overhaul menu initialization to prepare for owner-drawn menus
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
    ME_TextEn As String
    ME_TextTranslated As String
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
    AddMenuItem "File", "file_top", 0
    AddMenuItem "New...", "file_new", 0, 0, , "file_new"
    AddMenuItem "Open...", "file_open", 0, 1, , "file_open"
    AddMenuItem "Open recent", "file_openrecent", 0, 2
    AddMenuItem "Import", "file_import", 0, 3
        AddMenuItem "From clipboard", "file_import_paste", 0, 3, 0, "file_importclipboard"
        AddMenuItem "From scanner or camera...", "file_import_scanner", 0, 3, 2, "file_importscanner"
        AddMenuItem "Select which scanner or camera to use...", "file_import_selectscanner", 0, 3, 3
        AddMenuItem "Online image...", "file_import_web", 0, 3, 5, "file_importweb"
        AddMenuItem "Screenshot", "file_import_screenshot", 0, 3, 7, "file_importscreen"
    AddMenuItem "Close", "file_close", 0, 5, , "file_close"
    AddMenuItem "Close all", "file_closeall", 0, 6
    AddMenuItem "Save", "file_save", 0, 8, , "file_save"
    AddMenuItem "Save copy (lossless)", "file_savecopy", 0, 9, , "file_savedup"
    AddMenuItem "Save as...", "file_saveas", 0, 10, , "file_saveas"
    AddMenuItem "Revert", "file_revert", 0, 11
    AddMenuItem "Batch operations", "file_batch", 0, 13
        AddMenuItem "Process...", "file_batch_process", 0, 13, 0, "file_batch"
        AddMenuItem "Repair...", "file_batch_repair", 0, 13, 1, "file_repair"
    AddMenuItem "Print...", "file_print", 0, 15, , "file_print"
    AddMenuItem "Exit", "file_quit", 0, 17
    
    
    'Edit menu
    AddMenuItem "Edit", "edit_top", 1
    AddMenuItem "Undo", "edit_undo", 1, 0, , "edit_undo"
    AddMenuItem "Redo", "edit_redo", 1, 1, , "edit_redo"
    AddMenuItem "Undo history...", "edit_history", 1, 2, , "edit_history"
    AddMenuItem "Repeat", "edit_repeat", 1, 4, , "edit_repeat"
    AddMenuItem "Fade...", "edit_fade", 1, 5
    AddMenuItem "Cut", "edit_cut", 1, 7, , "edit_cut"
    AddMenuItem "Cut from layer", "edit_cutlayer", 1, 8
    AddMenuItem "Copy", "edit_copy", 1, 9, , "edit_copy"
    AddMenuItem "Copy from layer", "edit_copylayer", 1, 10
    AddMenuItem "Paste as new image", "edit_pasteasimage", 1, 11, , "edit_paste"
    AddMenuItem "Paste as new layer", "edit_pasteaslayer", 1, 12
    AddMenuItem "Empty clipboard", "edit_emptyclipboard", 1, 14
    
    
    'View Menu
    AddMenuItem "View", "view_top", 2
    AddMenuItem "Fit image on screen", "zoom_fit", 2, 0, , "zoom_fit"
    AddMenuItem "Zoom in", "zoom_in", 2, 2, , "zoom_in"
    AddMenuItem "Zoom out", "zoom_out", 2, 3, , "zoom_out"
    AddMenuItem "16:1 (1600%)", "zoom_16_1", 2, 5
    AddMenuItem "8:1 (800%)", "zoom_8_1", 2, 6
    AddMenuItem "4:1 (400%)", "zoom_4_1", 2, 7
    AddMenuItem "2:1 (200%)", "zoom_2_1", 2, 8
    AddMenuItem "1:1 (actual size, 100%)", "zoom_actual", 2, 9, , "zoom_actual"
    AddMenuItem "1:2 (50%)", "zoom_1_2", 2, 10
    AddMenuItem "1:4 (25%)", "zoom_1_4", 2, 11
    AddMenuItem "1:8 (12.5%)", "zoom_1_8", 2, 12
    AddMenuItem "1:16 (6.25%)", "zoom_1_16", 2, 13
    
    
    'Image Menu
    AddMenuItem "Image", "image_top", 3
    AddMenuItem "Duplicate", "image_duplicate", 3, 0, , "edit_copy"
    AddMenuItem "Resize...", "image_resize", 3, 2, , "image_resize"
    AddMenuItem "Content-aware resize...", "image_contentawareresize", 3, 3
    AddMenuItem "Canvas size...", "image_canvassize", 3, 5, , "image_canvassize"
    AddMenuItem "Fit canvas to active layer", "image_fittolayer", 3, 6
    AddMenuItem "Fit canvas around all layers", "image_fitalllayers", 3, 7
    AddMenuItem "Crop to selection", "image_crop", 3, 9, , "image_crop"
    AddMenuItem "Trim empty borders", "image_trim", 3, 10
    AddMenuItem "Rotate", "image_rotate", 3, 12
        AddMenuItem "Straighten", "image_straighten", 3, 12, 0
        AddMenuItem "90 clockwise", "image_rotate90", 3, 12, 2, "generic_rotateright"
        AddMenuItem "90 counter-clockwise", "image_rotate270", 3, 12, 3, "generic_rotateleft"
        AddMenuItem "180", "image_rotate180", 3, 12, 4
        AddMenuItem "Arbitrary...", "image_rotatearbitrary", 3, 12, 5
    AddMenuItem "Flip horizontal", "image_fliphorizontal", 3, 13, , "image_fliphorizontal"
    AddMenuItem "Flip vertical", "image_flipvertical", 3, 14, , "image_flipvertical"
    AddMenuItem "Metadata", "image_metadata", 3, 16
        AddMenuItem "Edit metadata...", "image_editmetadata", 3, 16, 0, "image_metadata"
        AddMenuItem "Count unique colors", "image_countcolors", 3, 16, 2
        AddMenuItem "Map photo location...", "image_maplocation", 3, 16, 3, "image_maplocation"
    
    
    'Layer menu
    AddMenuItem "Layer", "layer_top", 4
    AddMenuItem "Add", "layer_add", 4, 0
        AddMenuItem "Blank layer", "layer_addblank", 4, 0, 0
        AddMenuItem "Duplicate of current layer", "layer_duplicate", 4, 0, 1, "edit_copy"
        AddMenuItem "From clipboard", "layer_addfromclipboard", 4, 0, 3, "edit_paste"
        AddMenuItem "From file...", "layer_addfromfile", 4, 0, 4, "file_open"
    AddMenuItem "Delete", "layer_delete", 4, 1
        AddMenuItem "Current layer", "layer_deletecurrent", 4, 1, 0, "generic_trash"
        AddMenuItem "Hidden layers", "layer_deletehidden", 4, 1, 1, "generic_invisible"
    AddMenuItem "Merge up", "layer_mergeup", 4, 3, , "layer_mergeup"
    AddMenuItem "Merge down", "layer_mergedown", 4, 4, , "layer_mergedown"
    AddMenuItem "Order", "layer_order", 4, 5
        AddMenuItem "Raise layer", "layer_up", 4, 5, 0, "layer_up"
        AddMenuItem "Lower layer", "layer_down", 4, 5, 1, "layer_down"
        AddMenuItem "Layer to top", "layer_totop", 4, 5, 3
        AddMenuItem "Layer to bottom", "layer_tobottom", 4, 5, 4
    AddMenuItem "Orientation", "layer_orientation", 4, 7
        AddMenuItem "Straighten...", "layer_straighten", 4, 7, 0
        AddMenuItem "Rotate 90 clockwise", "layer_rotate90", 4, 7, 2, "generic_rotateright"
        AddMenuItem "Rotate 90 counter-clockwise", "layer_rotate270", 4, 7, 3, "generic_rotateleft"
        AddMenuItem "Rotate 180", "layer_rotate180", 4, 7, 4
        AddMenuItem "Rotate arbitrary...", "layer_rotatearbitrary", 4, 7, 5
        AddMenuItem "Flip horizontal", "layer_fliphorizontal", 4, 7, 7, "image_fliphorizontal"
        AddMenuItem "Flip vertical", "layer_flipvertical", 4, 7, 8, "image_flipvertical"
    AddMenuItem "Size", "layer_resize", 4, 8
        AddMenuItem "Reset to actual size", "layer_resetsize", 4, 8, 0, "generic_reset"
        AddMenuItem "Resize...", "layer_resize", 4, 8, 2, "image_resize"
        AddMenuItem "Content-aware resize...", "layer_contentawareresize", 4, 8, 3
    AddMenuItem "Crop to selection", "layer_crop", 4, 9, , "image_crop"
    AddMenuItem "Transparency", "layer_transparency", 4, 11
        AddMenuItem "Make color transparent", "layer_colortoalpha", 4, 11, 0
        AddMenuItem "Remove transparency...", "layer_removealpha", 4, 11, 1, "generic_trash"
    AddMenuItem "Rasterize", "layer_rasterize", 4, 13
        AddMenuItem "Current layer", "layer_rasterizecurrent", 4, 13, 0
        AddMenuItem "All layers", "layer_rasterizeall", 4, 13, 1
    AddMenuItem "Flatten image...", "layer_flatten", 4, 15, , "layer_flatten"
    AddMenuItem "Merge visible layers", "layer_mergevisible", 4, 16, , "generic_visible"
   
   
    'Select Menu
    AddMenuItem "Select", "select_top", 5
    AddMenuItem "All", "select_all", 5, 0
    AddMenuItem "None", "select_none", 5, 1
    AddMenuItem "Invert", "select_invert", 5, 2
    AddMenuItem "Grow...", "select_grow", 5, 4
    AddMenuItem "Shrink...", "select_shrink", 5, 5
    AddMenuItem "Border...", "select_border", 5, 6
    AddMenuItem "Feather...", "select_feather", 5, 7
    AddMenuItem "Sharpen...", "select_sharpen", 5, 8
    AddMenuItem "Erase selected area", "select_erasearea", 5, 10
    AddMenuItem "Load selection...", "select_load", 5, 12, , "file_open"
    AddMenuItem "Save current selection...", "select_save", 5, 13, , "file_save"
    AddMenuItem "Export", "select_export", 5, 14
        AddMenuItem "Selected area as image...", "select_exportarea", 5, 14, 0
        AddMenuItem "Selection mask as image...", "select_exportmask", 5, 14, 1
        
    
    'Tools Menu
    AddMenuItem "Tools", "tools_top", 8
    AddMenuItem "Language", "tools_language", 8, 0, , "tools_language"
    AddMenuItem "Language editor...", "tools_languageeditor", 8, 1
    AddMenuItem "Record macro", "tools_macrotop", 8, 3, , "macro_record"
        AddMenuItem "Start recording", "tools_recordmacro", 8, 3, 0, "macro_record"
        AddMenuItem "Stop recording...", "tools_stopmacro", 8, 3, 1, "macro_stop"
    AddMenuItem "Play macro...", "tools_playmacro", 8, 4, , "macro_play"
    AddMenuItem "Recent macros", "tools_recentmacros", 8, 5
    AddMenuItem "Options...", "tools_options", 8, 7, , "pref_advanced"
    AddMenuItem "Plugin manager...", "tools_plugins", 8, 8, , "tools_plugin"
    
    
    'Window Menu
    AddMenuItem "Window", "window_top", 9
    AddMenuItem "Toolbox", "window_toolbox", 9, 0
        AddMenuItem "Display toolbox", "window_displaytoolbox", 9, 0, 0
        AddMenuItem "Display tool category titles", "window_displaytoolcategories", 9, 0, 2
        AddMenuItem "Small buttons", "window_smalltoolbuttons", 9, 0, 4
        AddMenuItem "Normal buttons", "window_normaltoolbuttons", 9, 0, 5
        AddMenuItem "Large buttons", "window_largetoolbuttons", 9, 0, 6
    AddMenuItem "Tool options", "window_tooloptions", 9, 1
    AddMenuItem "Layers", "window_layers", 9, 2
    AddMenuItem "Image tabstrip", "window_imagetabstrip", 9, 3
        AddMenuItem "Always show", "window_imagetabstrip_alwaysshow", 9, 3, 0
        AddMenuItem "Show when multiple images are loaded", "window_imagetabstrip_shownormal", 9, 3, 1
        AddMenuItem "Never show", "window_imagetabstrip_nevershow", 9, 3, 2
        AddMenuItem "Left", "window_imagetabstrip_alignleft", 9, 3, 4
        AddMenuItem "Top", "window_imagetabstrip_aligntop", 9, 3, 5
        AddMenuItem "Right", "window_imagetabstrip_alignright", 9, 3, 6
        AddMenuItem "Bottom", "window_imagetabstrip_alignbottom", 9, 3, 7
    AddMenuItem "Next image", "window_next", 9, 5, , "generic_next"
    AddMenuItem "Previous image", "window_previous", 9, 6, , "generic_previous"
    
    
    'Help Menu
    AddMenuItem "Help", "help_top", 10
    AddMenuItem "Support us with a small donation (thank you!)", "help_donate", 10, 0, , "help_heart"
    AddMenuItem "Check for updates", "help_checkupdates", 10, 2, , "help_update"
    AddMenuItem "Submit feedback...", "help_contact", 10, 3, , "help_contact"
    AddMenuItem "Submit bug report...", "help_reportbug", 10, 4, , "help_reportbug"
    AddMenuItem "Visit PhotoDemon website", "help_website", 10, 6, , "help_website"
    AddMenuItem "Download PhotoDemon source code", "help_sourcecode", 10, 7, , "help_github"
    AddMenuItem "Read license and terms of use", "help_license", 10, 8, , "help_license"
    AddMenuItem "About", "help_about", 10, 10, , "help_about"
    
End Sub

'Internal helper function for adding a menu entry to the running collection.  Note that PD menus support a number of non-standard properties,
' all of which must be cached early in the load process so we can properly support things like UI themes and language translations.
Private Sub AddMenuItem(ByRef menuTextEn As String, ByRef menuName As String, ByVal topMenuID As Long, Optional ByVal subMenuID As Long = -1, Optional ByVal subSubMenuID As Long = -1, Optional ByRef menuImageName As String = vbNullString)
    
    'Make sure a sufficiently large buffer exists for this menu item
    Const INITIAL_MENU_COLLECTION_SIZE As Long = 64
    If (m_NumOfMenus = 0) Then
        ReDim m_Menus(0 To INITIAL_MENU_COLLECTION_SIZE - 1) As PD_MenuEntry
    Else
        If (m_NumOfMenus > UBound(m_Menus)) Then ReDim Preserve m_Menus(0 To m_NumOfMenus * 2 - 1) As PD_MenuEntry
    End If
    
    With m_Menus(m_NumOfMenus)
        .ME_Name = menuName
        .ME_TextEn = menuTextEn
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
