Attribute VB_Name = "Menus"
'***************************************************************************
'Specialized Math Routines
'Copyright 2017-2017 by Tanner Helland
'Created: 11/January/17
'Last updated: 15/August/17
'Last update: implement manual handling of Unicode menu captions
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
    ME_TopMenu As Long                      'Top-level index of this menu
    ME_SubMenu As Long                      'Sub-menu index of this menu (if any)
    ME_SubSubMenu As Long                   'Sub-sub-menu index of this menu (if any)
    ME_HotKeyCode As KeyCodeConstants       'Hotkey, if any, associated with this menu
    ME_HotKeyShift As ShiftConstants        'Hotkey shift modifiers, if any, associated with this menu
    ME_HotKeyTextTranslated As String       'Hotkey text, with translations (if any) always applied.
    ME_Name As String                       'Name of this menu (must be unique)
    ME_ResImage As String                   'Name of this menu's image, as stored in PD's central resource file
    ME_TextEn As String                     'Text of this menu, in English
    ME_TextTranslated As String             'Text of this menu, as translated by the current language
    ME_TextFinal As String                  'Final on-screen appearance of the text, with translations and accelerator applied
End Type

'https://msdn.microsoft.com/en-us/library/windows/desktop/ms647578(v=vs.85).aspx
Private Enum MIIM_fMask
    MIIM_BITMAP = &H80
    MIIM_CHECKMARKS = &H8
    MIIM_DATA = &H20
    MIIM_FTYPE_ = &H100
    MIIM_ID = &H2
    MIIM_STATE = &H1
    MIIM_STRING = &H40
    MIIM_SUBMENU = &H4
End Enum

Private Enum MIIM_fType
    MFT_BITMAP = &H4&           'Displays the menu item using a bitmap. The low-order word of the dwTypeData member is the bitmap handle, and the cch member is ignored. (MFT_BITMAP is replaced by MIIM_BITMAP and hbmpItem.)
    MFT_MENUBARBREAK = &H20&    'Places the menu item on a new line (for a menu bar) or in a new column (for a drop-down menu, submenu, or shortcut menu). For a drop-down menu, submenu, or shortcut menu, a vertical line separates the new column from the old.
    MFT_MENUBREAK = &H40&       'Places the menu item on a new line (for a menu bar) or in a new column (for a drop-down menu, submenu, or shortcut menu). For a drop-down menu, submenu, or shortcut menu, the columns are not separated by a vertical line.
    MFT_OWNERDRAW = &H100&      'Assigns responsibility for drawing the menu item to the window that owns the menu. The window receives a WM_MEASUREITEM message before the menu is displayed for the first time, and a WM_DRAWITEM message whenever the appearance of the menu item must be updated. If this value is specified, the dwTypeData member contains an application-defined value.
    MFT_RADIOCHECK = &H200&     'Displays selected menu items using a radio-button mark instead of a check mark if the hbmpChecked member is NULL.
    MFT_RIGHTJUSTIFY = &H4000&  'Right-justifies the menu item and any subsequent items. This value is valid only if the menu item is in a menu bar.
    MFT_RIGHTORDER = &H2000&    'Specifies that menus cascade right-to-left (the default is left-to-right). This is used to support right-to-left languages, such as Arabic and Hebrew.
    MFT_SEPARATOR = &H800&      'Specifies that the menu item is a separator. A menu item separator appears as a horizontal dividing line. The dwTypeData and cch members are ignored. This value is valid only in a drop-down menu, submenu, or shortcut menu.
    MFT_STRING = &H0&           'Displays the menu item using a text string. The dwTypeData member is the pointer to a null-terminated string, and the cch member is the length of the string.  (MFT_STRING is replaced by MIIM_STRING.)
End Enum

Private Enum MIIM_fState
    MFS_CHECKED = &H8&          'Checks the menu item. For more information about selected menu items, see the hbmpChecked member.
    MFS_DEFAULT = &H1000&       'Specifies that the menu item is the default. A menu can contain only one default menu item, which is displayed in bold.
    MFS_DISABLED = &H3&         'Disables the menu item and grays it so that it cannot be selected. This is equivalent to MFS_GRAYED.
    MFS_ENABLED = &H0&          'Enables the menu item so that it can be selected. This is the default state.
    MFS_GRAYED = &H3&           'Disables the menu item and grays it so that it cannot be selected. This is equivalent to MFS_DISABLED.
    MFS_HILITE = &H80&          'Highlights the menu item.
    MFS_UNCHECKED = &H0&        'Unchecks the menu item. For more information about clear menu items, see the hbmpChecked member.
    MFS_UNHILITE = &H0&         'Removes the highlight from the menu item. This is the default state.
End Enum

Private Enum Win32_MenuStateFlags
    MF_BYCOMMAND = &H0&         'Indicates that the uId parameter gives the identifier of the menu item. The MF_BYCOMMAND flag is the default if neither the MF_BYCOMMAND nor MF_BYPOSITION flag is specified.
    MF_BYPOSITION = &H400&      'Indicates that the uId parameter gives the zero-based relative position of the menu item.
    MF_CHECKED = &H8&           'A check mark is placed next to the item (for drop-down menus, submenus, and shortcut menus only).
    MF_DISABLED = &H2&          'The item is disabled.
    MF_GRAYED = &H1&            'The item is disabled and grayed.
    MF_HILITE = &H80&           'The item is highlighted.
    MF_MENUBARBREAK = &H20&     'This is the same as the MF_MENUBREAK flag, except for drop-down menus, submenus, and shortcut menus, where the new column is separated from the old column by a vertical line.
    MF_MENUBREAK = &H40&        'The item is placed on a new line (for menu bars) or in a new column (for drop-down menus, submenus, and shortcut menus) without separating columns.
    MF_OWNERDRAW = &H100&       'The item is owner-drawn.
    MF_POPUP = &H10&            'Menu item is a submenu.
    MF_SEPARATOR = &H800&       'There is a horizontal dividing line (for drop-down menus, submenus, and shortcut menus only).
End Enum

Private Type Win32_MenuItemInfoW
    cbSize          As Long
    fMask           As MIIM_fMask
    fType           As MIIM_fType
    fState          As MIIM_fState
    wID             As Long
    hSubMenu        As Long
    hbmpChecked     As Long
    hbmpUnchecked   As Long
    dwItemData      As Long
    dwTypeData      As Long
    cch             As Long
    hbmpItem        As Long
End Type

'When modifying menus, special ID values can be used to restrict operations
Private Const IGNORE_MENU_ID As Long = -10
Private Const ALL_MENU_SUBITEMS As Long = -9
Private Const MENU_NONE As Long = -1

'A number of menu features require us to interact directly with the API
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetMenuItemInfoW Lib "user32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Boolean, ByRef dstMenuItemInfo As Win32_MenuItemInfoW) As Boolean
Private Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal uId As Long, ByVal uFlags As Win32_MenuStateFlags) As Win32_MenuStateFlags
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemInfoW Lib "user32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Boolean, ByRef srcMenuItemInfo As Win32_MenuItemInfoW) As Boolean

'Primary menu collection
Private m_Menus() As PD_MenuEntry
Private m_NumOfMenus As Long

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


'Early in the PD load process, we initialize the default set of menus.  In the future, it may be nice to let
' users customize this to match their favorite software (e.g. PhotoShop), but that's a ways off as I've yet to
' build a menu control capable of that level of customization support.
Public Sub InitializeMenus()
    
    'File Menu
    AddMenuItem "&File", "file_top", 0
    AddMenuItem "&New...", "file_new", 0, 0, , "file_new"
    AddMenuItem "&Open...", "file_open", 0, 1, , "file_open"
    AddMenuItem "Open &recent", "file_openrecent", 0, 2
    AddMenuItem "&Import", "file_import", 0, 3
        AddMenuItem "From clipboard", "file_import_paste", 0, 3, 0, "file_importclipboard"
        AddMenuItem "-", "-", 0, 3, 1
        AddMenuItem "From scanner or camera...", "file_import_scanner", 0, 3, 2, "file_importscanner"
        AddMenuItem "Select which scanner or camera to use...", "file_import_selectscanner", 0, 3, 3
        AddMenuItem "-", "-", 0, 3, 4
        AddMenuItem "Online image...", "file_import_web", 0, 3, 5, "file_importweb"
        AddMenuItem "-", "-", 0, 3, 6
        AddMenuItem "Screenshot", "file_import_screenshot", 0, 3, 7, "file_importscreen"
    AddMenuItem "-", "-", 0, 4
    AddMenuItem "&Close", "file_close", 0, 5, , "file_close"
    AddMenuItem "Close all", "file_closeall", 0, 6
    AddMenuItem "-", "-", 0, 7
    AddMenuItem "&Save", "file_save", 0, 8, , "file_save"
    AddMenuItem "Save copy (&lossless)", "file_savecopy", 0, 9, , "file_savedup"
    AddMenuItem "Save &as...", "file_saveas", 0, 10, , "file_saveas"
    AddMenuItem "Revert", "file_revert", 0, 11
    AddMenuItem "-", "-", 0, 12
    AddMenuItem "&Batch operations", "file_batch", 0, 13
        AddMenuItem "Process...", "file_batch_process", 0, 13, 0, "file_batch"
        AddMenuItem "Repair...", "file_batch_repair", 0, 13, 1, "file_repair"
    AddMenuItem "-", "-", 0, 14
    AddMenuItem "&Print...", "file_print", 0, 15, , "file_print"
    AddMenuItem "-", "-", 0, 16
    AddMenuItem "E&xit", "file_quit", 0, 17
    
    
    'Edit menu
    AddMenuItem "&Edit", "edit_top", 1
    AddMenuItem "&Undo", "edit_undo", 1, 0, , "edit_undo"
    AddMenuItem "&Redo", "edit_redo", 1, 1, , "edit_redo"
    AddMenuItem "Undo history...", "edit_history", 1, 2, , "edit_history"
    AddMenuItem "-", "-", 1, 3
    AddMenuItem "Repeat", "edit_repeat", 1, 4, , "edit_repeat"
    AddMenuItem "Fade...", "edit_fade", 1, 5
    AddMenuItem "-", "-", 1, 6
    AddMenuItem "Cu&t", "edit_cut", 1, 7, , "edit_cut"
    AddMenuItem "Cut from layer", "edit_cutlayer", 1, 8
    AddMenuItem "&Copy", "edit_copy", 1, 9, , "edit_copy"
    AddMenuItem "Copy from layer", "edit_copylayer", 1, 10
    AddMenuItem "&Paste as new image", "edit_pasteasimage", 1, 11, , "edit_paste"
    AddMenuItem "Paste as new layer", "edit_pasteaslayer", 1, 12
    AddMenuItem "-", "-", 1, 13
    AddMenuItem "&Empty clipboard", "edit_emptyclipboard", 1, 14
    
    
    'View Menu
    AddMenuItem "&View", "view_top", 2
    AddMenuItem "&Fit image on screen", "zoom_fit", 2, 0, , "zoom_fit"
    AddMenuItem "-", "-", 2, 1
    AddMenuItem "Zoom &in", "zoom_in", 2, 2, , "zoom_in"
    AddMenuItem "Zoom &out", "zoom_out", 2, 3, , "zoom_out"
    AddMenuItem "-", "-", 2, 4
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
    AddMenuItem "&Image", "image_top", 3
    AddMenuItem "&Duplicate", "image_duplicate", 3, 0, , "edit_copy"
    AddMenuItem "-", "-", 3, 1
    AddMenuItem "Resize...", "image_resize", 3, 2, , "image_resize"
    AddMenuItem "Content-aware resize...", "image_contentawareresize", 3, 3
    AddMenuItem "-", "-", 3, 4
    AddMenuItem "Canvas size...", "image_canvassize", 3, 5, , "image_canvassize"
    AddMenuItem "Fit canvas to active layer", "image_fittolayer", 3, 6
    AddMenuItem "Fit canvas around all layers", "image_fitalllayers", 3, 7
    AddMenuItem "-", "-", 3, 8
    AddMenuItem "Crop to selection", "image_crop", 3, 9, , "image_crop"
    AddMenuItem "Trim empty borders", "image_trim", 3, 10
    AddMenuItem "-", "-", 3, 11
    AddMenuItem "Rotate", "image_rotate", 3, 12
        AddMenuItem "Straighten", "image_straighten", 3, 12, 0
        AddMenuItem "-", "-", 3, 12, 1
        AddMenuItem "90 clockwise", "image_rotate90", 3, 12, 2, "generic_rotateright"
        AddMenuItem "90 counter-clockwise", "image_rotate270", 3, 12, 3, "generic_rotateleft"
        AddMenuItem "180", "image_rotate180", 3, 12, 4
        AddMenuItem "Arbitrary...", "image_rotatearbitrary", 3, 12, 5
    AddMenuItem "Flip horizontal", "image_fliphorizontal", 3, 13, , "image_fliphorizontal"
    AddMenuItem "Flip vertical", "image_flipvertical", 3, 14, , "image_flipvertical"
    AddMenuItem "-", "-", 3, 15
    AddMenuItem "Metadata", "image_metadata", 3, 16
        AddMenuItem "Edit metadata...", "image_editmetadata", 3, 16, 0, "image_metadata"
        AddMenuItem "-", "-", 3, 16, 1
        AddMenuItem "Count unique colors", "image_countcolors", 3, 16, 2
        AddMenuItem "Map photo location...", "image_maplocation", 3, 16, 3, "image_maplocation"
    
    
    'Layer menu
    AddMenuItem "&Layer", "layer_top", 4
    AddMenuItem "Add", "layer_add", 4, 0
        AddMenuItem "Basic layer...", "layer_addbasic", 4, 0, 0
        AddMenuItem "Blank layer", "layer_addblank", 4, 0, 1
        AddMenuItem "Duplicate of current layer", "layer_duplicate", 4, 0, 2, "edit_copy"
        AddMenuItem "-", "-", 4, 0, 3
        AddMenuItem "From clipboard", "layer_addfromclipboard", 4, 0, 4, "edit_paste"
        AddMenuItem "From file...", "layer_addfromfile", 4, 0, 5, "file_open"
        AddMenuItem "From visible layers", "layer_addfromvisiblelayers", 4, 0, 6
    AddMenuItem "Delete", "layer_delete", 4, 1
        AddMenuItem "Current layer", "layer_deletecurrent", 4, 1, 0, "generic_trash"
        AddMenuItem "Hidden layers", "layer_deletehidden", 4, 1, 1, "generic_invisible"
    AddMenuItem "-", "-", 4, 2
    AddMenuItem "Merge up", "layer_mergeup", 4, 3, , "layer_mergeup"
    AddMenuItem "Merge down", "layer_mergedown", 4, 4, , "layer_mergedown"
    AddMenuItem "Order", "layer_order", 4, 5
        AddMenuItem "Raise layer", "layer_up", 4, 5, 0, "layer_up"
        AddMenuItem "Lower layer", "layer_down", 4, 5, 1, "layer_down"
        AddMenuItem "-", "-", 4, 5, 2
        AddMenuItem "Layer to top", "layer_totop", 4, 5, 3
        AddMenuItem "Layer to bottom", "layer_tobottom", 4, 5, 4
    AddMenuItem "-", "-", 4, 6
    AddMenuItem "Orientation", "layer_orientation", 4, 7
        AddMenuItem "Straighten...", "layer_straighten", 4, 7, 0
        AddMenuItem "-", "-", 4, 7, 1
        AddMenuItem "Rotate 90 clockwise", "layer_rotate90", 4, 7, 2, "generic_rotateright"
        AddMenuItem "Rotate 90 counter-clockwise", "layer_rotate270", 4, 7, 3, "generic_rotateleft"
        AddMenuItem "Rotate 180", "layer_rotate180", 4, 7, 4
        AddMenuItem "Rotate arbitrary...", "layer_rotatearbitrary", 4, 7, 5
        AddMenuItem "-", "-", 4, 7, 6
        AddMenuItem "Flip horizontal", "layer_fliphorizontal", 4, 7, 7, "image_fliphorizontal"
        AddMenuItem "Flip vertical", "layer_flipvertical", 4, 7, 8, "image_flipvertical"
    AddMenuItem "-", "-", 4, 7
    AddMenuItem "Size", "layer_resize", 4, 8
        AddMenuItem "Reset to actual size", "layer_resetsize", 4, 8, 0, "generic_reset"
        AddMenuItem "-", "-", 4, 8, 1
        AddMenuItem "Resize...", "layer_resize", 4, 8, 2, "image_resize"
        AddMenuItem "Content-aware resize...", "layer_contentawareresize", 4, 8, 3
    AddMenuItem "-", "-", 4, 8
    AddMenuItem "Crop to selection", "layer_crop", 4, 9, , "image_crop"
    AddMenuItem "-", "-", 4, 10
    AddMenuItem "Transparency", "layer_transparency", 4, 11
        AddMenuItem "Make color transparent", "layer_colortoalpha", 4, 11, 0
        AddMenuItem "Remove transparency...", "layer_removealpha", 4, 11, 1, "generic_trash"
    AddMenuItem "-", "-", 4, 12
    AddMenuItem "Rasterize", "layer_rasterize", 4, 13
        AddMenuItem "Current layer", "layer_rasterizecurrent", 4, 13, 0
        AddMenuItem "All layers", "layer_rasterizeall", 4, 13, 1
        AddMenuItem "-", "-", 4, 14
    AddMenuItem "Flatten image...", "layer_flatten", 4, 15, , "layer_flatten"
    AddMenuItem "Merge visible layers", "layer_mergevisible", 4, 16, , "generic_visible"
   
   
    'Select Menu
    AddMenuItem "&Select", "select_top", 5
    AddMenuItem "All", "select_all", 5, 0
    AddMenuItem "None", "select_none", 5, 1
    AddMenuItem "Invert", "select_invert", 5, 2
    AddMenuItem "-", "-", 5, 3
    AddMenuItem "Grow...", "select_grow", 5, 4
    AddMenuItem "Shrink...", "select_shrink", 5, 5
    AddMenuItem "Border...", "select_border", 5, 6
    AddMenuItem "Feather...", "select_feather", 5, 7
    AddMenuItem "Sharpen...", "select_sharpen", 5, 8
    AddMenuItem "-", "-", 5, 9
    AddMenuItem "Erase selected area", "select_erasearea", 5, 10
    AddMenuItem "-", "-", 5, 11
    AddMenuItem "Load selection...", "select_load", 5, 12, , "file_open"
    AddMenuItem "Save current selection...", "select_save", 5, 13, , "file_save"
    AddMenuItem "Export", "select_export", 5, 14
        AddMenuItem "Selected area as image...", "select_exportarea", 5, 14, 0
        AddMenuItem "Selection mask as image...", "select_exportmask", 5, 14, 1
        
    
    'Adjustments Menu
    AddMenuItem "&Adjustments", "adj_top", 6
    AddMenuItem "Auto correct", "adj_autocorrect", 6, 0
        AddMenuItem "Color", "adj_autocorrectcolor", 6, 0, 0
        AddMenuItem "Contrast", "adj_autocorrectcontrast", 6, 0, 1
        AddMenuItem "Lighting", "adj_autocorrectlighting", 6, 0, 2
        AddMenuItem "Shadows and highlights", "adj_autocorrectsandh", 6, 0, 3
    AddMenuItem "Auto enhance", "adj_autoenhance", 6, 1
        AddMenuItem "Color", "adj_autoenhancecolor", 6, 1, 0
        AddMenuItem "Contrast", "adj_autoenhancecontrast", 6, 1, 1
        AddMenuItem "Lighting", "adj_autoenhancelighting", 6, 1, 2
        AddMenuItem "Shadows and highlights", "adj_autoenhancesandh", 6, 1, 3
    AddMenuItem "-", "-", 6, 2
    AddMenuItem "Black and white...", "adj_blackandwhite", 6, 3
    AddMenuItem "Brightness and contrast...", "adj_bandc", 6, 4
    AddMenuItem "Color balance...", "adj_colorbalance", 6, 5
    AddMenuItem "Curves...", "adj_curves", 6, 6
    AddMenuItem "Levels...", "adj_levels", 6, 7
    AddMenuItem "Shadows and highlights...", "adj_sandh", 6, 8
    AddMenuItem "Vibrance...", "adj_vibrance", 6, 9
    AddMenuItem "White balance...", "adj_whitebalance", 6, 10
    AddMenuItem "-", "-", 6, 11
    AddMenuItem "Channels", "adj_channels", 6, 12
        AddMenuItem "Channel mixer...", "adj_channelmixer", 6, 12, 0
        AddMenuItem "Rechannel...", "adj_rechannel", 6, 12, 1
        AddMenuItem "-", "-", 6, 12, 2
        AddMenuItem "Maximum channel", "adj_maxchannel", 6, 12, 3
        AddMenuItem "Minimum channel", "adj_minchannel", 6, 12, 4
        AddMenuItem "-", "-", 6, 12, 5
        AddMenuItem "Shift left", "adj_shiftchannelsleft", 6, 12, 6
        AddMenuItem "Shift right", "adj_shiftchannelsright", 6, 12, 7
    AddMenuItem "Color", "adj_color", 6, 13
        AddMenuItem "Color balance...", "adj_colorbalance", 6, 13, 0
        AddMenuItem "White balance...", "adj_whitebalance", 6, 13, 1
        AddMenuItem "-", "-", 6, 13, 2
        AddMenuItem "Hue and saturation...", "adj_hsl", 6, 13, 3
        AddMenuItem "Temperature...", "adj_temperature", 6, 13, 4
        AddMenuItem "Tint...", "adj_tint", 6, 13, 5
        AddMenuItem "Vibrance...", "adj_vibrance", 6, 13, 6
        AddMenuItem "-", "-", 6, 13, 7
        AddMenuItem "Black and white...", "adj_blackandwhite", 6, 13, 8
        AddMenuItem "Colorize...", "adj_colorize", 6, 13, 9
        AddMenuItem "Replace color...", "adj_replacecolor", 6, 13, 10
        AddMenuItem "Sepia...", "adj_sepia", 6, 13, 11
    AddMenuItem "Histogram", "adj_histogram", 6, 14
        AddMenuItem "Display...", "adj_histogramdisplay", 6, 14, 0
        AddMenuItem "-", "-", 6, 14, 1
        AddMenuItem "Equalize...", "adj_histogramequalize", 6, 14, 2
        AddMenuItem "Stretch", "adj_histogramstretch", 6, 14, 3
    AddMenuItem "Invert", "adj_invert", 6, 15
        AddMenuItem "CMYK (film negative)", "adj_invertcmyk", 6, 15, 0
        AddMenuItem "Hue", "adj_inverthue", 6, 15, 1
        AddMenuItem "RGB", "adj_invertrgb", 6, 15, 2
    AddMenuItem "Lighting", "adj_lighting", 6, 16
        AddMenuItem "Brightness and contrast...", "adj_bandc", 6, 16, 0
        AddMenuItem "Curves...", "adj_curves", 6, 16, 1
        AddMenuItem "Gamma...", "adj_gamma", 6, 16, 2
        AddMenuItem "Levels...", "adj_levels", 6, 16, 3
        AddMenuItem "Shadows and highlights...", "adj_sandh", 6, 16, 4
    AddMenuItem "Monochrome", "adj_monochrome", 6, 17
        AddMenuItem "Color to monochrome...", "adj_colortomonochrome", 6, 17, 0
        AddMenuItem "Monochrome to gray...", "adj_monochrometogray", 6, 17, 1
    AddMenuItem "Photography", "adj_photo", 6, 18
        AddMenuItem "Exposure...", "adj_exposure", 6, 18, 0
        AddMenuItem "HDR...", "adj_hdr", 6, 18, 1
        AddMenuItem "Photo filters...", "adj_photofilters", 6, 18, 2
        AddMenuItem "Red-eye removal...", "adj_redeyeremoval", 6, 18, 3
        AddMenuItem "Split toning...", "adj_splittone", 6, 18, 4
    
    
    'Effects (Filters) Menu
    AddMenuItem "Effe&cts", "effects_top", 7
    AddMenuItem "Artistic", "effects_artistic", 7, 0
        AddMenuItem "Colored pencil...", "effects_colorpencil", 7, 0, 0
        AddMenuItem "Comic book...", "effects_comicbook", 7, 0, 1
        AddMenuItem "Figured glass (dents)...", "effects_figuredglass", 7, 0, 2
        AddMenuItem "Film noir...", "effects_filmnoir", 7, 0, 3
        AddMenuItem "Glass tiles...", "effects_glasstiles", 7, 0, 4
        AddMenuItem "Kaleidoscope...", "effects_kaleidoscope", 7, 0, 5
        AddMenuItem "Modern art...", "effects_modernart", 7, 0, 6
        AddMenuItem "Oil painting...", "effects_oilpainting", 7, 0, 7
        AddMenuItem "Plastic wrap...", "effects_plasticwrap", 7, 0, 8
        AddMenuItem "Posterize...", "effects_posterize", 7, 0, 9
        AddMenuItem "Relief...", "effects_relief", 7, 0, 10
        AddMenuItem "Stained glass...", "effects_stainedglass", 7, 0, 11
    AddMenuItem "Blur", "effects_blur", 7, 1
        AddMenuItem "Box blur...", "effects_boxblur", 7, 1, 0
        AddMenuItem "Gaussian blur...", "effects_gaussianblur", 7, 1, 1
        AddMenuItem "Surface blur...", "effects_surfaceblur", 7, 1, 2
        AddMenuItem "-", "-", 7, 1, 3
        AddMenuItem "Motion blur...", "effects_motionblur", 7, 1, 4
        AddMenuItem "Radial blur...", "effects_radialblur", 7, 1, 5
        AddMenuItem "Zoom blur...", "effects_zoomblur", 7, 1, 6
        AddMenuItem "-", "-", 7, 1, 7
        AddMenuItem "Kuwahara filter...", "effects_kuwahara", 7, 1, 8
        AddMenuItem "Symmetric nearest-neighbor...", "effects_snn", 7, 1, 9
    AddMenuItem "Distort", "effects_distort", 7, 2
        AddMenuItem "Correct existing distortion...", "effects_fixlensdistort", 7, 2, 0
        AddMenuItem "-", "-", 7, 2, 1
        AddMenuItem "Donut...", "effects_donut", 7, 2, 2
        AddMenuItem "Lens...", "effects_lens", 7, 2, 3
        AddMenuItem "Pinch and whirl...", "effects_pinchandwhirl", 7, 2, 4
        AddMenuItem "Poke...", "effects_poke", 7, 2, 5
        AddMenuItem "Ripple...", "effects_ripple", 7, 2, 6
        AddMenuItem "Squish...", "effects_squish", 7, 2, 7
        AddMenuItem "Swirl...", "effects_swirl", 7, 2, 8
        AddMenuItem "Waves...", "effects_waves", 7, 2, 9
        AddMenuItem "-", "-", 7, 2, 10
        AddMenuItem "Miscellaneous...", "effects_miscdistort", 7, 2, 11
    AddMenuItem "Edge", "effects_edges", 7, 3
        AddMenuItem "Emboss...", "effects_emboss", 7, 3, 0
        AddMenuItem "Enhance edges...", "effects_enhanceedges", 7, 3, 1
        AddMenuItem "Find edges...", "effects_findedges", 7, 3, 2
        AddMenuItem "Range filter...", "effects_rangefilter", 7, 3, 3
        AddMenuItem "Trace contour...", "effects_tracecontour", 7, 3, 4
    AddMenuItem "Light and shadow", "effects_lightandshadow", 7, 4
        AddMenuItem "Black light...", "effects_blacklight", 7, 4, 0
        AddMenuItem "Cross-screen...", "effects_crossscreen", 7, 4, 1
        AddMenuItem "Rainbow...", "effects_rainbow", 7, 4, 2
        AddMenuItem "Sunshine...", "effects_sunshine", 7, 4, 3
        AddMenuItem "-", "-", 7, 4, 4
        AddMenuItem "Dilate...", "effects_dilate", 7, 4, 5
        AddMenuItem "Erode...", "effects_erode", 7, 4, 6
    AddMenuItem "Natural", "effects_natural", 7, 5
        AddMenuItem "Atmosphere...", "effects_atmosphere", 7, 5, 0
        AddMenuItem "Fog...", "effects_fog", 7, 5, 1
        AddMenuItem "Ignite...", "effects_ignite", 7, 5, 2
        AddMenuItem "Lava...", "effects_lava", 7, 5, 3
        AddMenuItem "Metal...", "effects_metal", 7, 5, 4
        AddMenuItem "Snow...", "effects_snow", 7, 5, 5
        AddMenuItem "Underwater...", "effects_underwater", 7, 5, 6
    AddMenuItem "Noise", "effects_noise", 7, 6
        AddMenuItem "Add film grain...", "effects_filmgrain", 7, 6, 0
        AddMenuItem "Add RGB noise...", "effects_rgbnoise", 7, 6, 1
        AddMenuItem "-", "-", 7, 6, 2
        AddMenuItem "Anisotropic diffusion...", "effects_anisotropic", 7, 6, 3
        AddMenuItem "Bilateral filter...", "effects_bilateral", 7, 6, 4
        AddMenuItem "Harmonic mean...", "effects_harmonicmean", 7, 6, 5
        AddMenuItem "Mean shift...", "effects_meanshift", 7, 6, 6
        AddMenuItem "Median...", "effects_median", 7, 6, 7
    AddMenuItem "Pixelate", "effects_pixelate", 7, 7
        AddMenuItem "Color halftone...", "effects_colorhalftone", 7, 7, 0
        AddMenuItem "Crystallize...", "effects_crystallize", 7, 7, 1
        AddMenuItem "Fragment...", "effects_fragment", 7, 7, 2
        AddMenuItem "Mezzotint...", "effects_mezzotint", 7, 7, 3
        AddMenuItem "Mosaic...", "effects_mosaic", 7, 7, 4
    AddMenuItem "Sharpen", "effects_sharpentop", 7, 8
        AddMenuItem "Sharpen...", "effects_sharpen", 7, 8, 0
        AddMenuItem "Unsharp masking...", "effects_unsharp", 7, 8, 1
    AddMenuItem "Stylize", "effects_stylize", 7, 9
        AddMenuItem "Antique", "effects_antique", 7, 9, 0
        AddMenuItem "Diffuse...", "effects_diffuse", 7, 9, 1
        AddMenuItem "Outline...", "effects_outline", 7, 9, 2
        AddMenuItem "Palettize...", "effects_palettize", 7, 9, 3
        AddMenuItem "Portrait glow...", "effects_portraitglow", 7, 9, 4
        AddMenuItem "Solarize...", "effects_solarize", 7, 9, 5
        AddMenuItem "Twins...", "effects_twins", 7, 9, 6
        AddMenuItem "Vignetting...", "effects_vignetting", 7, 9, 7
    AddMenuItem "Transform", "effects_transform", 7, 10
        AddMenuItem "Pan and zoom...", "effects_panandzoom", 7, 10, 0
        AddMenuItem "Perspective...", "effects_perspective", 7, 10, 1
        AddMenuItem "Polar conversion...", "effects_polarconversion", 7, 10, 2
        AddMenuItem "Rotate...", "effects_rotate", 7, 10, 3
        AddMenuItem "Shear...", "effects_shear", 7, 10, 4
        AddMenuItem "Spherize...", "effects_spherize", 7, 10, 5
    AddMenuItem "-", "-", 7, 11
    AddMenuItem "Custom filter...", "effects_customfilter", 7, 12
    
    
    'Tools Menu
    AddMenuItem "&Tools", "tools_top", 8
    AddMenuItem "Language", "tools_language", 8, 0, , "tools_language"
    AddMenuItem "Language editor...", "tools_languageeditor", 8, 1
    AddMenuItem "-", "-", 8, 2
    AddMenuItem "Theme...", "tools_theme", 8, 3
    AddMenuItem "-", "-", 8, 4
    AddMenuItem "Record macro", "tools_macrotop", 8, 5, , "macro_record"
        AddMenuItem "Start recording", "tools_recordmacro", 8, 5, 0, "macro_record"
        AddMenuItem "Stop recording...", "tools_stopmacro", 8, 5, 1, "macro_stop"
    AddMenuItem "Play macro...", "tools_playmacro", 8, 6, , "macro_play"
    AddMenuItem "Recent macros", "tools_recentmacros", 8, 7
    AddMenuItem "-", "-", 8, 8
    AddMenuItem "Options...", "tools_options", 8, 9, , "pref_advanced"
    AddMenuItem "Plugin manager...", "tools_plugins", 8, 10, , "tools_plugin"
    
    Dim debugMenuVisibility As Boolean
    debugMenuVisibility = (PD_BUILD_QUALITY <> PD_PRODUCTION) And (PD_BUILD_QUALITY <> PD_BETA)
    If debugMenuVisibility Then
        AddMenuItem "-", "-", 8, 11
        AddMenuItem "Developers", "tools_developers", 8, 12
            AddMenuItem "Theme editor...", "tools_themeeditor", 8, 12, 0
            AddMenuItem "Build theme package...", "tools_themepackage", 8, 12, 1
        AddMenuItem "Test", "effects_developertest", 8, 13
    End If
    
    'Window Menu
    AddMenuItem "&Window", "window_top", 9
    AddMenuItem "Toolbox", "window_toolbox", 9, 0
        AddMenuItem "Display toolbox", "window_displaytoolbox", 9, 0, 0
        AddMenuItem "-", "-", 9, 0, 1
        AddMenuItem "Display tool category titles", "window_displaytoolcategories", 9, 0, 2
        AddMenuItem "-", "-", 9, 0, 3
        AddMenuItem "Small buttons", "window_smalltoolbuttons", 9, 0, 4
        AddMenuItem "Normal buttons", "window_normaltoolbuttons", 9, 0, 5
        AddMenuItem "Large buttons", "window_largetoolbuttons", 9, 0, 6
    AddMenuItem "Tool options", "window_tooloptions", 9, 1
    AddMenuItem "Layers", "window_layers", 9, 2
    AddMenuItem "Image tabstrip", "window_imagetabstrip", 9, 3
        AddMenuItem "Always show", "window_imagetabstrip_alwaysshow", 9, 3, 0
        AddMenuItem "Show when multiple images are loaded", "window_imagetabstrip_shownormal", 9, 3, 1
        AddMenuItem "Never show", "window_imagetabstrip_nevershow", 9, 3, 2
        AddMenuItem "-", "-", 9, 3, 3
        AddMenuItem "Left", "window_imagetabstrip_alignleft", 9, 3, 4
        AddMenuItem "Top", "window_imagetabstrip_aligntop", 9, 3, 5
        AddMenuItem "Right", "window_imagetabstrip_alignright", 9, 3, 6
        AddMenuItem "Bottom", "window_imagetabstrip_alignbottom", 9, 3, 7
    AddMenuItem "-", "-", 9, 4
    AddMenuItem "Next image", "window_next", 9, 5, , "generic_next"
    AddMenuItem "Previous image", "window_previous", 9, 6, , "generic_previous"
    
    
    'Help Menu
    AddMenuItem "&Help", "help_top", 10
    AddMenuItem "Support us with a small donation (thank you!)", "help_donate", 10, 0, , "help_heart"
    AddMenuItem "-", "-", 10, 1
    AddMenuItem "Check for &updates", "help_checkupdates", 10, 2, , "help_update"
    AddMenuItem "Submit feedback...", "help_contact", 10, 3, , "help_contact"
    AddMenuItem "Submit bug report...", "help_reportbug", 10, 4, , "help_reportbug"
    AddMenuItem "-", "-", 10, 5
    AddMenuItem "&Visit PhotoDemon website", "help_website", 10, 6, , "help_website"
    AddMenuItem "Download PhotoDemon source code", "help_sourcecode", 10, 7, , "help_github"
    AddMenuItem "Read license and terms of use", "help_license", 10, 8, , "help_license"
    AddMenuItem "-", "-", 10, 9
    AddMenuItem "&About", "help_about", 10, 10, , "help_about"
    
End Sub

'Internal helper function for adding a menu entry to the running collection.  Note that PD menus support a number of non-standard properties,
' all of which must be cached early in the load process so we can properly support things like UI themes and language translations.
Private Sub AddMenuItem(ByRef menuTextEn As String, ByRef menuName As String, ByVal topMenuID As Long, Optional ByVal subMenuID As Long = MENU_NONE, Optional ByVal subSubMenuID As Long = MENU_NONE, Optional ByRef menuImageName As String = vbNullString)
    
    'Make sure a sufficiently large buffer exists
    Const INITIAL_MENU_COLLECTION_SIZE As Long = 128
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
                IconsAndCursors.AddMenuIcon .ME_ResImage, .ME_TopMenu, .ME_SubMenu, .ME_SubSubMenu
            End With
        End If
    Next i

End Sub

'If you need to update a menu caption, this function supports Unicode captions.  (Note that Unicode captions can be
' necessary in non-obvious places, like filenames in Recent XYZ menus - so always use this function instead of the
' built-in VB ones, unless you're 100% certain you don't need Unicode!)
Public Sub RequestCaptionChange_ByName(ByVal menuName As String, ByVal newCaptionEn As String, Optional ByVal captionIsAlreadyTranslated As Boolean = False)

    'Resolve the menu name into one or more indices
    Dim numOfMenus As Long, menuIndices() As Long
    GetAllMatchingMenuIndices menuName, numOfMenus, menuIndices
    
    If (numOfMenus > 0) Then
    
        Dim i As Long
        For i = 0 To numOfMenus - 1
            
            With m_Menus(menuIndices(i))
            
                'Store the new caption and apply translations as necessary
                If captionIsAlreadyTranslated Or (g_Language Is Nothing) Then
                    .ME_TextTranslated = newCaptionEn
                Else
                    .ME_TextEn = newCaptionEn
                    .ME_TextTranslated = g_Language.TranslateMessage(newCaptionEn)
                End If
                
                'Deal with trailing accelerator text, if any
                If (Len(.ME_HotKeyTextTranslated) <> 0) Then
                    .ME_TextFinal = .ME_TextTranslated & vbTab & .ME_HotKeyTextTranslated
                Else
                    .ME_TextFinal = .ME_TextTranslated
                End If
                
                'Relay the changed text to the API copy of our menu
                UpdateMenuText_ByIndex menuIndices(i)
                
            End With
            
        Next i
    
    Else
        InternalMenuWarning "RequestCaptionChange_ByName", "no matching menus found"
    End If

End Sub

'Some menu-related text is accessed very frequently (e.g. "Ctrl" for hotkey text), so when a translation
' is active, we want to just cache the translations locally instead of regenerating them over and over.
Private Sub CacheCommonTranslations()
    ReDim m_CommonMenuText(0 To cmt_NumEntries - 1) As String
    If (Not g_Language Is Nothing) Then
        m_CommonMenuText(cmt_Ctrl) = g_Language.TranslateMessage("Ctrl")
        m_CommonMenuText(cmt_Alt) = g_Language.TranslateMessage("Alt")
        m_CommonMenuText(cmt_Shift) = g_Language.TranslateMessage("Shift")
    Else
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  g_Language isn't available, so hotkey captions won't be correct."
        #End If
    End If
End Sub

'After the active language changes, you must call this menu to translate all menu captions.
Public Sub UpdateAgainstCurrentTheme(Optional ByVal redrawMenuBar As Boolean = True)
    
    'Before proceeding, cache some common menu terms (so we don't have to keep translating them)
    CacheCommonTranslations
    
    Dim i As Long
    
    If g_Language.TranslationActive Then
        
        For i = 0 To m_NumOfMenus - 1
        
            With m_Menus(i)
                
                'Ignore separator entries, obviously
                If Strings.StringsNotEqual(.ME_Name, "-", False) Then
                
                    'Update the actual caption text
                    If (Len(.ME_TextEn) <> 0) Then
                    
                        .ME_TextTranslated = g_Language.TranslateMessage(.ME_TextEn)
                    
                        'Update the appended hotkey text, if any
                        If (.ME_HotKeyCode <> 0) Then
                            .ME_HotKeyTextTranslated = GetHotkeyText(.ME_HotKeyCode, .ME_HotKeyShift)
                            .ME_TextFinal = .ME_TextTranslated & vbTab & .ME_HotKeyTextTranslated
                        Else
                            .ME_TextFinal = .ME_TextTranslated
                        End If
                        
                    Else
                        .ME_TextTranslated = vbNullString
                        .ME_TextFinal = vbNullString
                    End If
                    
                Else
                    .ME_TextFinal = vbNullString
                End If
                    
            End With
            
        Next i
        
    Else
    
        For i = 0 To m_NumOfMenus - 1
        
            With m_Menus(i)
            
                If Strings.StringsNotEqual(.ME_Name, "-", False) Then
                
                    .ME_TextTranslated = .ME_TextEn
                    
                    If (.ME_HotKeyCode <> 0) Then
                        .ME_HotKeyTextTranslated = GetHotkeyText(.ME_HotKeyCode, .ME_HotKeyShift)
                        .ME_TextFinal = .ME_TextTranslated & vbTab & .ME_HotKeyTextTranslated
                    Else
                        .ME_TextFinal = .ME_TextTranslated
                    End If
                    
                Else
                    .ME_TextFinal = vbNullString
                End If
                
            End With
            
        Next i
    
    End If
    
    'With all menu captions updated, we now need to relay those changes to the underlying API menu struct
    For i = 0 To m_NumOfMenus - 1
        If (Len(m_Menus(i).ME_TextFinal) <> 0) Then UpdateMenuText_ByIndex i
    Next i
    
    'Some special menus must be dealt with now; note that some menus are already handled by dedicated callers
    ' (e.g. the "Languages" menu), while others must be handled here.
    Menus.UpdateSpecialMenu_RecentFiles
    Menus.UpdateSpecialMenu_RecentMacros
    
    If redrawMenuBar Then DrawMenuBar FormMain.hWnd
    
End Sub

'When the hotkey associated with a menu changes, call this sub to update our internal hotkey trackers.
' (We need to know hotkeys so we can render them with the menu captions.)
'
'NOTE: this function doesn't update the hotkey text associated with this menu, unless requested.
Public Sub NotifyMenuHotkey(ByRef menuID As String, ByVal vKeyCode As KeyCodeConstants, ByVal Shift As ShiftConstants)
    
    'Resolve the menuID into a list of indices.  (Note that menus can share the same name, meaning there can be more
    ' than one physical menu associated with a given hotkey.)
    Dim listOfMatches() As Long, numOfMatches As Long
    GetAllMatchingMenuIndices menuID, numOfMatches, listOfMatches
    
    'Before we enter the loop, generate a translated text representation of this hotkey
    Dim hotkeyText As String
    hotkeyText = GetHotkeyText(vKeyCode, Shift)
    
    Dim i As Long
    For i = 0 To numOfMatches - 1
        With m_Menus(listOfMatches(i))
            .ME_HotKeyCode = vKeyCode
            .ME_HotKeyShift = Shift
            .ME_HotKeyTextTranslated = hotkeyText
        End With
    Next i

End Sub

'Helper check for resolving menu enablement by menu name.  Note that PD *does not* enforce unique menu names; in fact, they are
' specifically allowed by design.  As such, this function only returns the *first* matching entry, with the assumption that
' same-named menus are enabled and disabled as a group.
Public Function IsMenuEnabled(ByRef mnuName As String) As Boolean

    'Resolve the menu name into an index into our menu collection
    Dim i As Long
    Dim mnuIndex As Long: mnuIndex = -1
    
    For i = 0 To m_NumOfMenus - 1
        If Strings.StringsEqual(mnuName, m_Menus(i).ME_Name, True) Then
            mnuIndex = i
            Exit For
        End If
    Next i
    
    If (mnuIndex >= 0) Then
    
        'Get an hMenu and index for this entry
        Dim hMenu As Long, hIndex As Long
        hMenu = GetHMenu_FromIndex(mnuIndex, True)
        hIndex = GetHMenuIndex(mnuIndex)
        
        If (hMenu <> 0) Then
            Dim mFlags As Win32_MenuStateFlags
            mFlags = GetMenuState(hMenu, hIndex, MF_BYPOSITION)
            IsMenuEnabled = Not ((mFlags And (MF_DISABLED Or MF_GRAYED)) <> 0)
        End If
        
    Else
        InternalMenuWarning "IsMenuEnabled", "no matching menu found - check your menu name!"
    End If

End Function

'Until I have a better place to stick this, hotkeys are handled here, by the menu module.  This is primarily done because there is
' fairly tight integration between hotkeys and menu captions, and both need to be handled together while accounting for the usual
' nightmares (like language translations).
Public Sub InitializeAllHotkeys()
    
    With FormMain.pdHotkeys
    
        .Enabled = True
    
        'File menu
        .AddAccelerator vbKeyN, vbCtrlMask, "New image", "file_new", True, False, True, UNDO_Nothing
        .AddAccelerator vbKeyO, vbCtrlMask, "Open", "file_open", True, False, True, UNDO_Nothing
        .AddAccelerator vbKeyF4, vbCtrlMask, "Close", "file_close", True, True, True, UNDO_Nothing
        .AddAccelerator vbKeyF4, vbCtrlMask Or vbShiftMask, "Close all", "file_closeall", True, True, True, UNDO_Nothing
        .AddAccelerator vbKeyS, vbCtrlMask, "Save", "file_save", True, True, True, UNDO_Nothing
        .AddAccelerator vbKeyS, vbCtrlMask Or vbAltMask Or vbShiftMask, "Save copy", "file_savecopy", True, False, True, UNDO_Nothing
        .AddAccelerator vbKeyS, vbCtrlMask Or vbShiftMask, "Save as", "file_saveas", True, True, True, UNDO_Nothing
        .AddAccelerator vbKeyF12, 0, "Revert", "file_revert", True, True, False, UNDO_Nothing
        .AddAccelerator vbKeyB, vbCtrlMask, "Batch wizard", "file_batch_process", True, False, True, UNDO_Nothing
        .AddAccelerator vbKeyP, vbCtrlMask, "Print", "file_print", True, True, True, UNDO_Nothing
        .AddAccelerator vbKeyQ, vbCtrlMask, "Exit program", "file_quit", True, False, True, UNDO_Nothing
        
            'File -> Import submenu
            .AddAccelerator vbKeyI, vbCtrlMask Or vbShiftMask Or vbAltMask, "Scan image", "file_import_scanner", True, False, True, UNDO_Nothing
            .AddAccelerator vbKeyD, vbCtrlMask Or vbShiftMask, "Internet import", "file_import_web", True, False, True, UNDO_Nothing
            .AddAccelerator vbKeyI, vbCtrlMask Or vbAltMask, "Screen capture", "file_import_screenshot", True, False, True, UNDO_Nothing
        
            'Most-recently used files.  Note that we cannot automatically associate these with a menu, as these menus may not
            ' exist at run-time.  (They are dynamically created as necessary.)
            .AddAccelerator vbKey0, vbCtrlMask, "MRU_0"
            .AddAccelerator vbKey1, vbCtrlMask, "MRU_1"
            .AddAccelerator vbKey2, vbCtrlMask, "MRU_2"
            .AddAccelerator vbKey3, vbCtrlMask, "MRU_3"
            .AddAccelerator vbKey4, vbCtrlMask, "MRU_4"
            .AddAccelerator vbKey5, vbCtrlMask, "MRU_5"
            .AddAccelerator vbKey6, vbCtrlMask, "MRU_6"
            .AddAccelerator vbKey7, vbCtrlMask, "MRU_7"
            .AddAccelerator vbKey8, vbCtrlMask, "MRU_8"
            .AddAccelerator vbKey9, vbCtrlMask, "MRU_9"
            
        'Edit menu
        .AddAccelerator vbKeyZ, vbCtrlMask, "Undo", "edit_undo", True, True, False, UNDO_Nothing
        .AddAccelerator vbKeyY, vbCtrlMask, "Redo", "edit_redo", True, True, False, UNDO_Nothing
        
        .AddAccelerator vbKeyF, vbCtrlMask, "Repeat last action", "edit_repeat", True, True, False, UNDO_Image
        
        .AddAccelerator vbKeyX, vbCtrlMask, "Cut", "edit_cut", True, True, False, UNDO_Image
        .AddAccelerator vbKeyX, vbCtrlMask Or vbShiftMask, "Cut from layer", "edit_cutlayer", True, True, False, UNDO_Layer
        .AddAccelerator vbKeyC, vbCtrlMask, "Copy", "edit_copy", True, True, False, UNDO_Nothing
        .AddAccelerator vbKeyC, vbCtrlMask Or vbShiftMask, "Copy from layer", "edit_copylayer", True, True, False, UNDO_Nothing
        .AddAccelerator vbKeyV, vbCtrlMask, "Paste as new image", "edit_pasteasimage", True, False, False, UNDO_Nothing
        .AddAccelerator vbKeyV, vbCtrlMask Or vbShiftMask, "Paste as new layer", "edit_pasteaslayer", True, False, False, UNDO_Image_VectorSafe
        
        'View menu
        .AddAccelerator vbKey0, 0, "FitOnScreen", "zoom_fit", False, True, False, UNDO_Nothing
        '.AddAccelerator vbKeyAdd, 0, "Zoom_In", "zoom_in", False, True, False, UNDO_NOTHING
        '.AddAccelerator VK_OEM_PLUS, 0, "Zoom_In", , False, True, False, UNDO_NOTHING
        '.AddAccelerator vbKeySubtract, 0, "Zoom_Out", "zoom_out", False, True, False, UNDO_NOTHING
        '.AddAccelerator VK_OEM_MINUS, 0, "Zoom_Out", , False, True, False, UNDO_NOTHING
        .AddAccelerator vbKey5, 0, "Zoom_161", "zoom_16_1", False, True, False, UNDO_Nothing
        .AddAccelerator vbKey4, 0, "Zoom_81", "zoom_8_1", False, True, False, UNDO_Nothing
        .AddAccelerator vbKey3, 0, "Zoom_41", "zoom_4_1", False, True, False, UNDO_Nothing
        .AddAccelerator vbKey2, 0, "Zoom_21", "zoom_2_1", False, True, False, UNDO_Nothing
        .AddAccelerator vbKey1, 0, "Actual_Size", "zoom_actual", False, True, False, UNDO_Nothing
        .AddAccelerator vbKey2, vbShiftMask, "Zoom_12", "zoom_1_2", False, True, False, UNDO_Nothing
        .AddAccelerator vbKey3, vbShiftMask, "Zoom_14", "zoom_1_4", False, True, False, UNDO_Nothing
        .AddAccelerator vbKey4, vbShiftMask, "Zoom_18", "zoom_1_8", False, True, False, UNDO_Nothing
        .AddAccelerator vbKey5, vbShiftMask, "Zoom_116", "zoom_1_16", False, True, False, UNDO_Nothing
        
        'Image menu
        .AddAccelerator vbKeyA, vbCtrlMask Or vbShiftMask, "Duplicate image", "image_duplicate", True, True, False, UNDO_Nothing
        .AddAccelerator vbKeyR, vbCtrlMask, "Resize image", "image_resize", True, True, True, UNDO_Image
        .AddAccelerator vbKeyR, vbCtrlMask Or vbAltMask, "Canvas size", "image_canvassize", True, True, True, UNDO_ImageHeader
        '.AddAccelerator vbKeyX, vbCtrlMask Or vbShiftMask, "Crop", FormMain.MnuImage(8), True, True, False, UNDO_IMAGE
        .AddAccelerator vbKeyX, vbCtrlMask Or vbAltMask, "Trim empty borders", "image_trim", True, True, False, UNDO_ImageHeader
        
            'Image -> Rotate submenu
            .AddAccelerator vbKeyR, 0, "Rotate image 90 clockwise", "image_rotate90", True, True, False, UNDO_Image
            .AddAccelerator vbKeyL, 0, "Rotate image 90 counter-clockwise", "image_rotate270", True, True, False, UNDO_Image
            .AddAccelerator vbKeyR, vbCtrlMask Or vbShiftMask Or vbAltMask, "Arbitrary image rotation", "image_rotatearbitrary", True, True, True, UNDO_Nothing
        
        'Layer Menu
        '(none yet)
        
        
        'Select Menu
        .AddAccelerator vbKeyA, vbCtrlMask, "Select all", "select_all", True, True, False, UNDO_Selection
        .AddAccelerator vbKeyD, vbCtrlMask, "Remove selection", "select_none", False, True, False, UNDO_Selection
        .AddAccelerator vbKeyI, vbCtrlMask Or vbShiftMask, "Invert selection", "select_invert", True, True, False, UNDO_Selection
        'KeyCode VK_OEM_4 = {[  (next to the letter P), VK_OEM_6 = }]
        .AddAccelerator VK_OEM_6, vbCtrlMask Or vbAltMask, "Grow selection", "select_grow", True, True, True, UNDO_Nothing
        .AddAccelerator VK_OEM_4, vbCtrlMask Or vbAltMask, "Shrink selection", "select_shrink", True, True, True, UNDO_Nothing
        .AddAccelerator vbKeyD, vbCtrlMask Or vbAltMask, "Feather selection", "select_feather", True, True, True, UNDO_Nothing
        
        'Adjustments Menu
        
        'Adjustments top shortcut menu
        .AddAccelerator vbKeyU, vbCtrlMask Or vbShiftMask, "Black and white", "adj_blackandwhite", True, True, True, UNDO_Nothing
        .AddAccelerator vbKeyB, vbCtrlMask Or vbShiftMask, "Brightness and contrast", "adj_bandc", True, True, True, UNDO_Nothing
        .AddAccelerator vbKeyC, vbCtrlMask Or vbAltMask, "Color balance", "adj_colorbalance", True, True, True, UNDO_Nothing
        .AddAccelerator vbKeyM, vbCtrlMask, "Curves", "adj_curves", True, True, True, UNDO_Nothing
        .AddAccelerator vbKeyL, vbCtrlMask, "Levels", "adj_levels", True, True, True, UNDO_Nothing
        .AddAccelerator vbKeyH, vbCtrlMask Or vbShiftMask, "Shadow and highlight", "adj_sandh", True, True, True, UNDO_Nothing
        .AddAccelerator vbKeyAdd, vbCtrlMask Or vbAltMask, "Vibrance", "adj_vibrance", True, True, True, UNDO_Nothing
        .AddAccelerator VK_OEM_PLUS, vbCtrlMask Or vbAltMask, "Vibrance", , True, True, True, UNDO_Nothing
        .AddAccelerator vbKeyW, vbCtrlMask, "White balance", "adj_whitebalance", True, True, True, UNDO_Nothing
        
            'Color adjustments
            .AddAccelerator vbKeyH, vbCtrlMask, "Hue and saturation", "adj_hsl", True, True, True, UNDO_Nothing
            .AddAccelerator vbKeyT, vbCtrlMask, "Temperature", "adj_temperature", True, True, True, UNDO_Nothing
            
            'Lighting adjustments
            .AddAccelerator vbKeyG, vbCtrlMask, "Gamma", "adj_gamma", True, True, True, UNDO_Nothing
            
            'Adjustments -> Invert submenu
            .AddAccelerator vbKeyI, vbCtrlMask, "Invert RGB", "adj_invertRGB", True, True, False, UNDO_Layer
            
            'Adjustments -> Monochrome submenu
            .AddAccelerator vbKeyB, vbCtrlMask Or vbAltMask Or vbShiftMask, "Color to monochrome", "adj_colortomonochrome", True, True, True, UNDO_Nothing
            
            'Adjustments -> Photography submenu
            .AddAccelerator vbKeyE, vbCtrlMask Or vbAltMask, "Exposure", "adj_exposure", True, True, True, UNDO_Nothing
            .AddAccelerator vbKeyP, vbCtrlMask Or vbAltMask, "Photo filter", "adj_photofilters", True, True, True, UNDO_Nothing
            
        
        'Effects Menu
        '.AddAccelerator vbKeyZ, vbCtrlMask Or vbAltMask Or vbShiftMask, "Add RGB noise", FormMain.MnuNoise(1), True, True, True, False
        '.AddAccelerator vbKeyG, vbCtrlMask Or vbAltMask Or vbShiftMask, "Gaussian blur", FormMain.MnuBlurFilter(1), True, True, True, False
        '.AddAccelerator vbKeyY, vbCtrlMask Or vbAltMask Or vbShiftMask, "Correct lens distortion", FormMain.MnuDistortEffects(1), True, True, True, False
        '.AddAccelerator vbKeyU, vbCtrlMask Or vbAltMask Or vbShiftMask, "Unsharp mask", FormMain.MnuSharpen(1), True, True, True, False
        
        'Tools menu
        'KeyCode 190 = >.  (two keys to the right of the M letter key)
        .AddAccelerator 190, vbCtrlMask Or vbAltMask, "Play macro", "tools_playmacro", True, True, True, UNDO_Nothing
        
        .AddAccelerator vbKeyReturn, vbAltMask, "Preferences", "tools_options", False, False, True, UNDO_Nothing
        .AddAccelerator vbKeyM, vbCtrlMask Or vbAltMask, "Plugin manager", "tools_plugins", False, False, True, UNDO_Nothing
        
        
        'Window menu
        .AddAccelerator vbKeyPageDown, 0, "Next_Image", "window_next", False, True, False, UNDO_Nothing
        .AddAccelerator vbKeyPageUp, 0, "Prev_Image", "window_previous", False, True, False, UNDO_Nothing
        
        'Activate hotkey detection
        .ActivateHook
        
    End With
    
    'Before exiting, notify the menu manager of all menu changes
    Dim i As Long
    
    CacheCommonTranslations
    
    With FormMain.pdHotkeys
        For i = 0 To .Count - 1
            If .HasMenu(i) Then Menus.NotifyMenuHotkey .GetMenuName(i), .GetKeyCode(i), .GetShift(i)
        Next i
    End With
    
End Sub

'If a menu has a hotkey associated with it, you can use this function to update the language-specific text representation of the hotkey.
' (This text is appended to the menu caption automatically.)
Private Function GetHotkeyText(ByVal vKeyCode As KeyCodeConstants, ByVal Shift As ShiftConstants) As String
    
    Dim tmpString As String
    If (Shift And vbCtrlMask) Then tmpString = m_CommonMenuText(cmt_Ctrl) & "+"
    If (Shift And vbAltMask) Then tmpString = tmpString & m_CommonMenuText(cmt_Alt) & "+"
    If (Shift And vbShiftMask) Then tmpString = tmpString & m_CommonMenuText(cmt_Shift) & "+"
    
    'Processing the string itself takes a bit of extra work, as some keyboard keys don't automatically map to a
    ' string equivalent.  (Also, translations need to be considered.)
    Select Case vKeyCode
    
        Case vbKeyAdd
            tmpString = tmpString & "+"
        
        Case vbKeySubtract
            tmpString = tmpString & "-"
        
        Case vbKeyReturn
            tmpString = tmpString & g_Language.TranslateMessage("Enter")
        
        Case vbKeyPageUp
            tmpString = tmpString & g_Language.TranslateMessage("Page Up")
        
        Case vbKeyPageDown
            tmpString = tmpString & g_Language.TranslateMessage("Page Down")
            
        Case vbKeyF1 To vbKeyF16
            tmpString = tmpString & "F" & (vKeyCode - 111)
        
        'In the future I would like to enumerate virtual key bindings properly, using the data at this link:
        ' http://msdn.microsoft.com/en-us/library/windows/desktop/dd375731%28v=vs.85%29.aspx
        ' At the moment, however, they're implemented as magic numbers.
        Case 188
            tmpString = tmpString & ","
            
        Case 190
            tmpString = tmpString & "."
            
        Case 219
            tmpString = tmpString & "["
            
        Case 221
            tmpString = tmpString & "]"
            
        Case Else
            tmpString = tmpString & UCase$(ChrW$(vKeyCode))
        
    End Select
    
    GetHotkeyText = tmpString
    
End Function

Private Sub GetAllMatchingMenuIndices(ByRef menuID As String, ByRef numOfMenus As Long, ByRef menuArray() As Long)

    'At present, there will never be more than two menus matching a given ID; this can be revisited in the future
    Const MAX_MENU_MATCHES As Long = 2
    If (Not VBHacks.IsArrayInitialized(menuArray)) Then
        ReDim menuArray(0 To MAX_MENU_MATCHES - 1) As Long
    Else
        If (UBound(menuArray) < MAX_MENU_MATCHES - 1) Or (LBound(menuArray) <> 0) Then ReDim menuArray(0 To MAX_MENU_MATCHES - 1) As Long
    End If
    
    Dim i As Long, curIndex As Long
    For i = 0 To m_NumOfMenus - 1
        If Strings.StringsEqual(menuID, m_Menus(i).ME_Name, True) Then
            menuArray(curIndex) = i
            curIndex = curIndex + 1
            If (curIndex >= MAX_MENU_MATCHES) Then Exit For
        End If
    Next i
    
    numOfMenus = curIndex
    
End Sub

'Some menus in PD (like the Recent Files menu, or the Tools > Languages menu) are directly modified at run-time.  In PD,
' it is easiest to wipe these entire menus dynamically, than rebuild them from scratch.
'
'IMPORTANT NOTE: to erase an entire submenu, pass ALL_MENU_SUBITEMS as the subMenuID or subSubMenuID, whichever is relevant.
'                ALL_MENU_SUBITEMS indicates "erase everything that matches the two preceding entries, except for the
'                top-level menu itself".
'IMPORTANT NOTE: this function will erase all submenus of the selected menu, by design.
Private Sub EraseMenu(ByVal topMenuID As Long, Optional ByVal subMenuID As Long = IGNORE_MENU_ID, Optional ByVal subSubMenuID As Long = IGNORE_MENU_ID)
    
    'Removed menus are flagged; we traverse the collection in two passes to make it faster to remove large menu subtrees
    Const REMOVED_MENU_ID As Long = -999
    
    Dim i As Long
    For i = 0 To m_NumOfMenus - 1
    
        'Top menus are always matched
        If (m_Menus(i).ME_TopMenu = topMenuID) Then
            
            'Submenu IDs are only matched if the user specifically requests it
            If (subMenuID <> IGNORE_MENU_ID) Then
                
                'Match the submenu ID
                If (m_Menus(i).ME_SubMenu = subMenuID) Then
                    
                    'Match the subsubmenu ID
                    If (subSubMenuID <> IGNORE_MENU_ID) Then
                        
                        If (m_Menus(i).ME_SubSubMenu = subSubMenuID) Then
                            m_Menus(i).ME_TopMenu = REMOVED_MENU_ID
                        ElseIf (subSubMenuID = ALL_MENU_SUBITEMS) And (m_Menus(i).ME_SubSubMenu >= 0) Then
                            m_Menus(i).ME_TopMenu = REMOVED_MENU_ID
                        End If
                    
                    Else
                        m_Menus(i).ME_TopMenu = REMOVED_MENU_ID
                    End If
                    
                ElseIf (subMenuID = ALL_MENU_SUBITEMS) And (m_Menus(i).ME_SubMenu >= 0) Then
                    m_Menus(i).ME_TopMenu = REMOVED_MENU_ID
                End If
            
            Else
                m_Menus(i).ME_TopMenu = REMOVED_MENU_ID
            End If
            
        End If
    Next i
    
    'All menus to be removed have now been properly flagged.  Iterate through the list and fill all empty spots.
    Dim moveOffset As Long
    moveOffset = 0
    
    For i = 0 To m_NumOfMenus - 1
    
        'If this item is set to be deleted, increment our move counter
        If (m_Menus(i).ME_TopMenu = REMOVED_MENU_ID) Then
            moveOffset = moveOffset + 1
        
        'If this is a valid item, shift it downward in the list
        Else
            If (moveOffset > 0) Then m_Menus(i - moveOffset) = m_Menus(i)
        End If
        
    Next i
    
    'Change the menu item count to reflect any/all moved entries
    If (moveOffset = 0) Then InternalMenuWarning "EraseMenu", "no menus erased - were the passed indices valid?"
    m_NumOfMenus = m_NumOfMenus - moveOffset
    
End Sub

'Some of PD's menus obey special rules.  (For example, menus that add/remove entries at run-time.)  These menus have their own
' helper update functions that can be called on demand, separate from other menus in the project.
Public Sub UpdateSpecialMenu_Language(ByVal numOfLanguages As Long, ByRef availableLanguages() As PDLanguageFile)

    'Retrieve handles to the parent menu
    Dim hMenu As Long
    hMenu = GetMenu(FormMain.hWnd)
    hMenu = GetSubMenu(hMenu, 8)
    hMenu = GetSubMenu(hMenu, 0)
    
    'Prepare a MenuItemInfo struct
    Dim tmpMII As Win32_MenuItemInfoW
    tmpMII.cbSize = LenB(tmpMII)
    tmpMII.fMask = MIIM_STRING
    
    If (hMenu <> 0) Then
        
        'Add anew captions for all the current menu entries.  (Note that the language manager has handled the actual creation
        ' of these menu objects; we use VB itself for that.)
        Dim i As Long
        For i = 0 To numOfLanguages - 1
            tmpMII.dwTypeData = StrPtr(availableLanguages(i).LangName)
            SetMenuItemInfoW hMenu, i, 1&, tmpMII
        Next i
        
    Else
        InternalMenuWarning "UpdateSpecialMenu_Language", "null hMenu"
    End If
    
    
End Sub

Public Sub UpdateSpecialMenu_RecentFiles()

    'Whenever the "File > Open Recent" menu is modified, we need to modify our internal list of recent file items.
    ' (We manually track this menu so we can handle translations correctly for the items at the bottom of the menu,
    ' e.g. "Load all" and "Clear list".)
    
    'Start by retrieving a handle to the menu in question
    If (Not g_RecentFiles Is Nothing) Then
    
        Dim hMenu As Long
        hMenu = GetMenu(FormMain.hWnd)
        hMenu = GetSubMenu(hMenu, 0&)
        hMenu = GetSubMenu(hMenu, 2&)
        
        'Prepare a MenuItemInfo struct
        Dim tmpMII As Win32_MenuItemInfoW
        tmpMII.cbSize = LenB(tmpMII)
        tmpMII.fMask = MIIM_STRING
        
        If (hMenu <> 0) Then
            
            'Retrieve the number of MRU files currently being displayed
            Dim numOfMRUFiles As Long
            numOfMRUFiles = g_RecentFiles.GetNumOfItems()
            
            'It is possible for there to be "0" files, in which case a blank "empty" indicator will be shown.
            ' Note that this messes with our ordinal positioning, however, so we need to manually account for
            ' this case.
            Dim listIsEmpty As Boolean
            listIsEmpty = (numOfMRUFiles = 0)
            
            'The position of the "load all" and "erase all" icons are hard-coded, relative to the number of displayed MRU files
            Dim tmpString As String
            tmpString = g_Language.TranslateMessage("Open all recent images")
            tmpMII.dwTypeData = StrPtr(tmpString)
            If (Not listIsEmpty) Then SetMenuItemInfoW hMenu, numOfMRUFiles + 1, 1&, tmpMII
            
            tmpString = g_Language.TranslateMessage("Clear recent image list")
            tmpMII.dwTypeData = StrPtr(tmpString)
            If (Not listIsEmpty) Then SetMenuItemInfoW hMenu, numOfMRUFiles + 2, 1&, tmpMII
                
            'Finally, manually place the captions for all recent file filenames, while handling the special
            ' case of an empty list.
            If listIsEmpty Then
                
                tmpString = g_Language.TranslateMessage("Empty")
                tmpMII.dwTypeData = StrPtr(tmpString)
                SetMenuItemInfoW hMenu, 0&, 1&, tmpMII
                
            Else
                
                'If actual MRU paths exist, note that we apply them *without* translations, obviously.
                Dim i As Long
                For i = 0 To numOfMRUFiles - 1
                    
                    tmpString = g_RecentFiles.GetMenuCaption(i)
                    
                    'Entries under "10" get a free accelerator of the form "Ctrl+i"
                    If (i < 10) Then tmpString = tmpString & vbTab & g_Language.TranslateMessage("Ctrl") & "+" & i
                    
                    tmpMII.dwTypeData = StrPtr(tmpString)
                    SetMenuItemInfoW hMenu, i, 1&, tmpMII
                    
                Next i
                
            End If
            
        Else
            InternalMenuWarning "UpdateSpecialMenu_RecentFiles", "hMenu was null"
        End If
        
    End If
    
End Sub

Public Sub UpdateSpecialMenu_RecentMacros()

    'Whenever the "Tools > Open Recent Macro" menu is modified, we need to modify our internal list of recent
    ' macro items.  (We manually track this menu so we can handle translations correctly for the item at the
    ' bottom of the menu, e.g. "Clear list".)
    
    'Start by retrieving a handle to the menu in question
    If (Not g_RecentMacros Is Nothing) Then
    
        Dim hMenu As Long
        hMenu = GetMenu(FormMain.hWnd)
        hMenu = GetSubMenu(hMenu, 8&)
        hMenu = GetSubMenu(hMenu, 7&)
        
        'Prepare a MenuItemInfo struct
        Dim tmpMII As Win32_MenuItemInfoW
        tmpMII.cbSize = LenB(tmpMII)
        tmpMII.fMask = MIIM_STRING
        
        If (hMenu <> 0) Then
            
            'Retrieve the number of MRU files currently being displayed
            Dim numOfMRUFiles As Long
            numOfMRUFiles = g_RecentMacros.MRU_ReturnCount()
            
            'It is possible for there to be "0" files, in which case a blank "empty" indicator will be shown.
            ' Note that this messes with our ordinal positioning, however, so we need to manually account for
            ' this case.
            Dim listIsEmpty As Boolean
            listIsEmpty = (numOfMRUFiles = 0)
            
            'The position of the "clear list" icon is hard-coded, relative to the number of displayed MRU files
            Dim tmpString As String
            
            tmpString = g_Language.TranslateMessage("Clear recent macro list")
            tmpMII.dwTypeData = StrPtr(tmpString)
            If (Not listIsEmpty) Then SetMenuItemInfoW hMenu, numOfMRUFiles + 1, 1&, tmpMII Else SetMenuItemInfoW hMenu, 2, 1&, tmpMII
                
            'Finally, manually place the captions for all recent file filenames, while handling the special
            ' case of an empty list.
            If listIsEmpty Then
                
                tmpString = g_Language.TranslateMessage("Empty")
                tmpMII.dwTypeData = StrPtr(tmpString)
                SetMenuItemInfoW hMenu, 0&, 1&, tmpMII
                
            Else
                
                'If actual MRU paths exist, note that we apply them *without* translations, obviously.
                Dim i As Long
                For i = 0 To numOfMRUFiles - 1
                    tmpString = g_RecentMacros.GetSpecificMRUCaption(i)
                    tmpMII.dwTypeData = StrPtr(tmpString)
                    SetMenuItemInfoW hMenu, i, 1&, tmpMII
                Next i
                
            End If
            
        Else
            InternalMenuWarning "UpdateSpecialMenu_RecentMacros", "hMenu was null"
        End If
        
    End If
    
End Sub

'Given an index into our menu collection, retrieve a matching hMenu for use with APIs.
'NOTE!  Validate your menu index before passing it to this function.  For performance reasons, no extra validation is
' applied to the incoming index.
Private Function GetHMenu_FromIndex(ByVal mnuIndex As Long, Optional ByVal getParentMenu As Boolean = False) As Long

    'We always start by retrieving the menu handle for the primary form
    Dim curHMenu As Long, hMenuParent As Long
    curHMenu = GetMenu(FormMain.hWnd)
    hMenuParent = curHMenu
    
    'Next, iterate through submenus until we arrive at the entry we want
    curHMenu = GetSubMenu(curHMenu, m_Menus(mnuIndex).ME_TopMenu)
    If (m_Menus(mnuIndex).ME_SubMenu <> MENU_NONE) Then
        
        hMenuParent = curHMenu
        curHMenu = GetSubMenu(curHMenu, m_Menus(mnuIndex).ME_SubMenu)
        
        If (m_Menus(mnuIndex).ME_SubSubMenu <> MENU_NONE) Then
            hMenuParent = curHMenu
            curHMenu = GetSubMenu(curHMenu, m_Menus(mnuIndex).ME_SubSubMenu)
        End If
        
    End If
    
    If getParentMenu Then GetHMenu_FromIndex = hMenuParent Else GetHMenu_FromIndex = curHMenu
    
End Function

'When working with APIs, you typically pass the hMenu of the parent menu, and then a simple itemIndex to address the
' child item in a given menu.  Use this function to simplify the handling of hMenu indices.
Private Function GetHMenuIndex(ByVal mnuIndex As Long) As Long

    If (m_Menus(mnuIndex).ME_SubMenu = MENU_NONE) Then
        GetHMenuIndex = m_Menus(mnuIndex).ME_TopMenu
    Else
        If (m_Menus(mnuIndex).ME_SubSubMenu = MENU_NONE) Then
            GetHMenuIndex = m_Menus(mnuIndex).ME_SubMenu
        Else
            GetHMenuIndex = m_Menus(mnuIndex).ME_SubSubMenu
        End If
    End If

End Function

'Update a given menu's text caption.  By design, this function does *not* trigger a DrawMenuBar call.
Private Function UpdateMenuText_ByIndex(ByVal mnuIndex As Long)

    'Get an hMenu for the specified index
    Dim hMenu As Long
    hMenu = GetHMenu_FromIndex(mnuIndex, True)
    
    If (hMenu <> 0) Then
        
        'Populate a MenuItemInfo struct
        Dim tmpMII As Win32_MenuItemInfoW
        tmpMII.cbSize = LenB(tmpMII)
        tmpMII.fMask = MIIM_STRING
        tmpMII.dwTypeData = StrPtr(m_Menus(mnuIndex).ME_TextFinal)
        
        SetMenuItemInfoW hMenu, GetHMenuIndex(mnuIndex), 1&, tmpMII
        
    Else
        InternalMenuWarning "UpdateMenuText_ByIndex", "null hMenu (" & mnuIndex & ")"
    End If

End Function

Private Sub InternalMenuWarning(ByVal funcName As String, ByVal errMsg As String)
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "WARNING!  Menus." & funcName & " reported: " & errMsg
    #Else
        Debug.Print "Menus." & funcName & " warns: " & errMsg
    #End If
End Sub
