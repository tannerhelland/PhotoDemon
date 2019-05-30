Attribute VB_Name = "Menus"
'***************************************************************************
'PhotoDemon Menu Manager
'Copyright 2017-2019 by Tanner Helland
'Created: 11/January/17
'Last updated: 03/May/19
'Last update: large overhaul to handle action redirection via menu caption and/or name
'
'PhotoDemon has an extensive menu system.  Managing all those menus is a cumbersome task.  This module exists
' to tackle the worst parts of run-time maintenance, so other functions don't need to.
'
'Because the menus provide a nice hierarchical collection of program features, this module also handles
' some module-adjacent tasks, like the ProcessDefaultAction-prefixed functions.  You can pass these functions
' either the name or caption of a menu, and they will automatically initiate the corresponding program action.
' (FormMain makes extensive use of this, obviously.)
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************


Option Explicit

Private Type PD_MenuEntry
    me_TopMenu As Long                    'Top-level index of this menu
    me_SubMenu As Long                    'Sub-menu index of this menu (if any)
    me_SubSubMenu As Long                 'Sub-sub-menu index of this menu (if any)
    me_HotKeyCode As KeyCodeConstants     'Hotkey, if any, associated with this menu
    me_HotKeyShift As ShiftConstants      'Hotkey shift modifiers, if any, associated with this menu
    me_HotKeyTextTranslated As String     'Hotkey text, with translations (if any) always applied.
    me_Name As String                     'Name of this menu (must be unique)
    me_ResImage As String                 'Name of this menu's image, as stored in PD's central resource file
    me_TextEn As String                   'Text of this menu, in English
    me_TextTranslated As String           'Text of this menu, as translated by the current language
    me_TextFinal As String                'Final on-screen appearance of the text, with translations and accelerator
    me_TextSearchable As String           'Localized string for search results.  Uses "TopMenu > ChildMenu > MyMenuName" format.
    me_HasChildren As Boolean             'Is this a non-clickable menu (e.g. it only exists to open a child menu?)
    me_DoNotIncludeInSearch As Boolean    'If TRUE, this menu will not appear in search results.  (Used for checkbox menus.)
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
Private Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal uId As Long, ByVal uFlags As Win32_MenuStateFlags) As Win32_MenuStateFlags
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemInfoW Lib "user32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, ByRef srcMenuItemInfo As Win32_MenuItemInfoW) As Boolean

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
        AddMenuItem "Screenshot...", "file_import_screenshot", 0, 3, 7, "file_importscreen"
    AddMenuItem "-", "-", 0, 4
    AddMenuItem "&Close", "file_close", 0, 5, , "file_close"
    AddMenuItem "Close all", "file_closeall", 0, 6
    AddMenuItem "-", "-", 0, 7
    AddMenuItem "&Save", "file_save", 0, 8, , "file_save"
    AddMenuItem "Save copy (&lossless)", "file_savecopy", 0, 9, , "file_savedup"
    AddMenuItem "Save &as...", "file_saveas", 0, 10, , "file_saveas"
    AddMenuItem "Revert", "file_revert", 0, 11
    AddMenuItem "Export", "file_export", 0, 12
        AddMenuItem "Color profile...", "file_export_colorprofile", 0, 12, 0
        AddMenuItem "Palette...", "file_export_palette", 0, 12, 1
    AddMenuItem "-", "-", 0, 13
    AddMenuItem "&Batch operations", "file_batch", 0, 14
        AddMenuItem "Process...", "file_batch_process", 0, 14, 0, "file_batch"
        AddMenuItem "Repair...", "file_batch_repair", 0, 14, 1, "file_repair"
    AddMenuItem "-", "-", 0, 15
    AddMenuItem "&Print...", "file_print", 0, 16, , "file_print"
    AddMenuItem "-", "-", 0, 17
    AddMenuItem "E&xit", "file_quit", 0, 18
    
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
    
    'Image Menu
    AddMenuItem "&Image", "image_top", 2
    AddMenuItem "&Duplicate", "image_duplicate", 2, 0, , "edit_copy"
    AddMenuItem "-", "-", 2, 1
    AddMenuItem "Resize...", "image_resize", 2, 2, , "image_resize"
    AddMenuItem "Content-aware resize...", "image_contentawareresize", 2, 3
    AddMenuItem "-", "-", 2, 4
    AddMenuItem "Canvas size...", "image_canvassize", 2, 5, , "image_canvassize"
    AddMenuItem "Fit canvas to active layer", "image_fittolayer", 2, 6
    AddMenuItem "Fit canvas around all layers", "image_fitalllayers", 2, 7
    AddMenuItem "-", "-", 2, 8
    AddMenuItem "Crop to selection", "image_crop", 2, 9, , "image_crop"
    AddMenuItem "Trim empty borders", "image_trim", 2, 10
    AddMenuItem "-", "-", 2, 11
    AddMenuItem "Rotate", "image_rotate", 2, 12
        AddMenuItem "Straighten...", "image_straighten", 2, 12, 0
        AddMenuItem "-", "-", 2, 12, 1
        AddMenuItem "90 clockwise", "image_rotate90", 2, 12, 2, "generic_rotateright"
        AddMenuItem "90 counter-clockwise", "image_rotate270", 2, 12, 3, "generic_rotateleft"
        AddMenuItem "180", "image_rotate180", 2, 12, 4
        AddMenuItem "Arbitrary...", "image_rotatearbitrary", 2, 12, 5
    AddMenuItem "Flip horizontal", "image_fliphorizontal", 2, 13, , "image_fliphorizontal"
    AddMenuItem "Flip vertical", "image_flipvertical", 2, 14, , "image_flipvertical"
    AddMenuItem "-", "-", 2, 15
    AddMenuItem "Metadata", "image_metadata", 2, 16
        AddMenuItem "Edit metadata...", "image_editmetadata", 2, 16, 0, "image_metadata"
        AddMenuItem "Remove all metadata", "image_removemetadata", 2, 16, 1
        AddMenuItem "-", "-", 2, 16, 2
        AddMenuItem "Count unique colors", "image_countcolors", 2, 16, 3
        AddMenuItem "Map photo location...", "image_maplocation", 2, 16, 4, "image_maplocation"
    
    'Layer menu
    AddMenuItem "&Layer", "layer_top", 3
    AddMenuItem "Add", "layer_add", 3, 0
        AddMenuItem "Basic layer...", "layer_addbasic", 3, 0, 0
        AddMenuItem "Blank layer", "layer_addblank", 3, 0, 1
        AddMenuItem "Duplicate of current layer", "layer_duplicate", 3, 0, 2, "edit_copy"
        AddMenuItem "-", "-", 3, 0, 3
        AddMenuItem "From clipboard", "layer_addfromclipboard", 3, 0, 4, "edit_paste"
        AddMenuItem "From file...", "layer_addfromfile", 3, 0, 5, "file_open"
        AddMenuItem "From visible layers", "layer_addfromvisiblelayers", 3, 0, 6
    AddMenuItem "Delete", "layer_delete", 3, 1
        AddMenuItem "Current layer", "layer_deletecurrent", 3, 1, 0, "generic_trash"
        AddMenuItem "Hidden layers", "layer_deletehidden", 3, 1, 1, "generic_invisible"
    AddMenuItem "-", "-", 3, 2
    AddMenuItem "Merge up", "layer_mergeup", 3, 3, , "layer_mergeup"
    AddMenuItem "Merge down", "layer_mergedown", 3, 4, , "layer_mergedown"
    AddMenuItem "Order", "layer_order", 3, 5
        AddMenuItem "Raise layer", "layer_up", 3, 5, 0, "layer_up"
        AddMenuItem "Lower layer", "layer_down", 3, 5, 1, "layer_down"
        AddMenuItem "-", "-", 3, 5, 2
        AddMenuItem "Layer to top", "layer_totop", 3, 5, 3
        AddMenuItem "Layer to bottom", "layer_tobottom", 3, 5, 4
    AddMenuItem "-", "-", 3, 6
    AddMenuItem "Orientation", "layer_orientation", 3, 7
        AddMenuItem "Straighten...", "layer_straighten", 3, 7, 0
        AddMenuItem "-", "-", 3, 7, 1
        AddMenuItem "Rotate 90 clockwise", "layer_rotate90", 3, 7, 2, "generic_rotateright"
        AddMenuItem "Rotate 90 counter-clockwise", "layer_rotate270", 3, 7, 3, "generic_rotateleft"
        AddMenuItem "Rotate 180", "layer_rotate180", 3, 7, 4
        AddMenuItem "Rotate arbitrary...", "layer_rotatearbitrary", 3, 7, 5
        AddMenuItem "-", "-", 3, 7, 6
        AddMenuItem "Flip horizontal", "layer_fliphorizontal", 3, 7, 7, "image_fliphorizontal"
        AddMenuItem "Flip vertical", "layer_flipvertical", 3, 7, 8, "image_flipvertical"
    AddMenuItem "-", "-", 3, 7
    AddMenuItem "Size", "layer_resize", 3, 8
        AddMenuItem "Reset to actual size", "layer_resetsize", 3, 8, 0, "generic_reset"
        AddMenuItem "-", "-", 3, 8, 1
        AddMenuItem "Resize...", "layer_resize", 3, 8, 2, "image_resize"
        AddMenuItem "Content-aware resize...", "layer_contentawareresize", 3, 8, 3
    AddMenuItem "-", "-", 3, 8
    AddMenuItem "Crop to selection", "layer_crop", 3, 9, , "image_crop"
    AddMenuItem "-", "-", 3, 10
    AddMenuItem "Transparency", "layer_transparency", 3, 11
        AddMenuItem "Make color transparent...", "layer_colortoalpha", 3, 11, 0
        AddMenuItem "Remove transparency...", "layer_removealpha", 3, 11, 1, "generic_trash"
    AddMenuItem "-", "-", 3, 12
    AddMenuItem "Rasterize", "layer_rasterize", 3, 13
        AddMenuItem "Current layer", "layer_rasterizecurrent", 3, 13, 0
        AddMenuItem "All layers", "layer_rasterizeall", 3, 13, 1
        AddMenuItem "-", "-", 3, 14
    AddMenuItem "Merge visible layers", "layer_mergevisible", 3, 15, , "generic_visible"
    AddMenuItem "Flatten image...", "layer_flatten", 3, 16, , "layer_flatten"
    
    'Select Menu
    AddMenuItem "&Select", "select_top", 4
    AddMenuItem "All", "select_all", 4, 0
    AddMenuItem "None", "select_none", 4, 1
    AddMenuItem "Invert", "select_invert", 4, 2
    AddMenuItem "-", "-", 4, 3
    AddMenuItem "Grow...", "select_grow", 4, 4
    AddMenuItem "Shrink...", "select_shrink", 4, 5
    AddMenuItem "Border...", "select_border", 4, 6
    AddMenuItem "Feather...", "select_feather", 4, 7
    AddMenuItem "Sharpen...", "select_sharpen", 4, 8
    AddMenuItem "-", "-", 4, 9
    AddMenuItem "Erase selected area", "select_erasearea", 4, 10
    AddMenuItem "-", "-", 4, 11
    AddMenuItem "Load selection...", "select_load", 4, 12, , "file_open"
    AddMenuItem "Save current selection...", "select_save", 4, 13, , "file_save"
    AddMenuItem "Export", "select_export", 4, 14
        AddMenuItem "Selected area as image...", "select_exportarea", 4, 14, 0
        AddMenuItem "Selection mask as image...", "select_exportmask", 4, 14, 1
        
    'Adjustments Menu
    AddMenuItem "&Adjustments", "adj_top", 5
    AddMenuItem "Auto correct", "adj_autocorrect", 5, 0
    AddMenuItem "Auto enhance", "adj_autoenhance", 5, 1
    AddMenuItem "-", "-", 5, 2
    AddMenuItem "Black and white...", "adj_blackandwhite", 5, 3
    AddMenuItem "Brightness and contrast...", "adj_bandc", 5, 4
    AddMenuItem "Color balance...", "adj_colorbalance", 5, 5
    AddMenuItem "Curves...", "adj_curves", 5, 6
    AddMenuItem "Levels...", "adj_levels", 5, 7
    AddMenuItem "Shadows and highlights...", "adj_sandh", 5, 8
    AddMenuItem "Vibrance...", "adj_vibrance", 5, 9
    AddMenuItem "White balance...", "adj_whitebalance", 5, 10
    AddMenuItem "-", "-", 5, 11
    AddMenuItem "Channels", "adj_channels", 5, 12
        AddMenuItem "Channel mixer...", "adj_channelmixer", 5, 12, 0
        AddMenuItem "Rechannel...", "adj_rechannel", 5, 12, 1
        AddMenuItem "-", "-", 5, 12, 2
        AddMenuItem "Maximum channel", "adj_maxchannel", 5, 12, 3
        AddMenuItem "Minimum channel", "adj_minchannel", 5, 12, 4
        AddMenuItem "-", "-", 5, 12, 5
        AddMenuItem "Shift left", "adj_shiftchannelsleft", 5, 12, 6
        AddMenuItem "Shift right", "adj_shiftchannelsright", 5, 12, 7
    AddMenuItem "Color", "adj_color", 5, 13
        AddMenuItem "Color balance...", "adj_colorbalance", 5, 13, 0
        AddMenuItem "White balance...", "adj_whitebalance", 5, 13, 1
        AddMenuItem "-", "-", 5, 13, 2
        AddMenuItem "Hue and saturation...", "adj_hsl", 5, 13, 3
        AddMenuItem "Temperature...", "adj_temperature", 5, 13, 4
        AddMenuItem "Tint...", "adj_tint", 5, 13, 5
        AddMenuItem "Vibrance...", "adj_vibrance", 5, 13, 6
        AddMenuItem "-", "-", 5, 13, 7
        AddMenuItem "Black and white...", "adj_blackandwhite", 5, 13, 8
        AddMenuItem "Colorize...", "adj_colorize", 5, 13, 9
        AddMenuItem "Replace color...", "adj_replacecolor", 5, 13, 10
        AddMenuItem "Sepia", "adj_sepia", 5, 13, 11
    AddMenuItem "Histogram", "adj_histogram", 5, 14
        AddMenuItem "Display...", "adj_histogramdisplay", 5, 14, 0
        AddMenuItem "-", "-", 5, 14, 1
        AddMenuItem "Equalize...", "adj_histogramequalize", 5, 14, 2
        AddMenuItem "Stretch", "adj_histogramstretch", 5, 14, 3
    AddMenuItem "Invert", "adj_invert", 5, 15
        AddMenuItem "CMYK (film negative)", "adj_invertcmyk", 5, 15, 0
        AddMenuItem "Hue", "adj_inverthue", 5, 15, 1
        AddMenuItem "RGB", "adj_invertrgb", 5, 15, 2
    AddMenuItem "Lighting", "adj_lighting", 5, 16
        AddMenuItem "Brightness and contrast...", "adj_bandc", 5, 16, 0
        AddMenuItem "Curves...", "adj_curves", 5, 16, 1
        AddMenuItem "Gamma...", "adj_gamma", 5, 16, 2
        AddMenuItem "Levels...", "adj_levels", 5, 16, 3
        AddMenuItem "Shadows and highlights...", "adj_sandh", 5, 16, 4
    AddMenuItem "Monochrome", "adj_monochrome", 5, 17
        AddMenuItem "Color to monochrome...", "adj_colortomonochrome", 5, 17, 0
        AddMenuItem "Monochrome to gray...", "adj_monochrometogray", 5, 17, 1
    AddMenuItem "Photography", "adj_photo", 5, 18
        AddMenuItem "Exposure...", "adj_exposure", 5, 18, 0
        AddMenuItem "HDR...", "adj_hdr", 5, 18, 1
        AddMenuItem "Photo filters...", "adj_photofilters", 5, 18, 2
        AddMenuItem "Red-eye removal...", "adj_redeyeremoval", 5, 18, 3
        AddMenuItem "Split toning...", "adj_splittone", 5, 18, 4
    
    'Effects (Filters) Menu
    AddMenuItem "Effe&cts", "effects_top", 6
    AddMenuItem "Artistic", "effects_artistic", 6, 0
        AddMenuItem "Colored pencil...", "effects_colorpencil", 6, 0, 0
        AddMenuItem "Comic book...", "effects_comicbook", 6, 0, 1
        AddMenuItem "Figured glass (dents)...", "effects_figuredglass", 6, 0, 2
        AddMenuItem "Film noir...", "effects_filmnoir", 6, 0, 3
        AddMenuItem "Glass tiles...", "effects_glasstiles", 6, 0, 4
        AddMenuItem "Kaleidoscope...", "effects_kaleidoscope", 6, 0, 5
        AddMenuItem "Modern art...", "effects_modernart", 6, 0, 6
        AddMenuItem "Oil painting...", "effects_oilpainting", 6, 0, 7
        AddMenuItem "Plastic wrap...", "effects_plasticwrap", 6, 0, 8
        AddMenuItem "Posterize...", "effects_posterize", 6, 0, 9
        AddMenuItem "Relief...", "effects_relief", 6, 0, 10
        AddMenuItem "Stained glass...", "effects_stainedglass", 6, 0, 11
    AddMenuItem "Blur", "effects_blur", 6, 1
        AddMenuItem "Box blur...", "effects_boxblur", 6, 1, 0
        AddMenuItem "Gaussian blur...", "effects_gaussianblur", 6, 1, 1
        AddMenuItem "Surface blur...", "effects_surfaceblur", 6, 1, 2
        AddMenuItem "-", "-", 6, 1, 3
        AddMenuItem "Motion blur...", "effects_motionblur", 6, 1, 4
        AddMenuItem "Radial blur...", "effects_radialblur", 6, 1, 5
        AddMenuItem "Zoom blur...", "effects_zoomblur", 6, 1, 6
        AddMenuItem "-", "-", 6, 1, 7
        AddMenuItem "Kuwahara filter...", "effects_kuwahara", 6, 1, 8
        AddMenuItem "Symmetric nearest-neighbor...", "effects_snn", 6, 1, 9
    AddMenuItem "Distort", "effects_distort", 6, 2
        AddMenuItem "Correct existing distortion...", "effects_fixlensdistort", 6, 2, 0
        AddMenuItem "-", "-", 6, 2, 1
        AddMenuItem "Donut...", "effects_donut", 6, 2, 2
        AddMenuItem "Lens...", "effects_lens", 6, 2, 3
        AddMenuItem "Pinch and whirl...", "effects_pinchandwhirl", 6, 2, 4
        AddMenuItem "Poke...", "effects_poke", 6, 2, 5
        AddMenuItem "Ripple...", "effects_ripple", 6, 2, 6
        AddMenuItem "Squish...", "effects_squish", 6, 2, 7
        AddMenuItem "Swirl...", "effects_swirl", 6, 2, 8
        AddMenuItem "Waves...", "effects_waves", 6, 2, 9
        AddMenuItem "-", "-", 6, 2, 10
        AddMenuItem "Miscellaneous...", "effects_miscdistort", 6, 2, 11
    AddMenuItem "Edge", "effects_edges", 6, 3
        AddMenuItem "Emboss...", "effects_emboss", 6, 3, 0
        AddMenuItem "Enhance edges...", "effects_enhanceedges", 6, 3, 1
        AddMenuItem "Find edges...", "effects_findedges", 6, 3, 2
        AddMenuItem "Range filter...", "effects_rangefilter", 6, 3, 3
        AddMenuItem "Trace contour...", "effects_tracecontour", 6, 3, 4
    AddMenuItem "Light and shadow", "effects_lightandshadow", 6, 4
        AddMenuItem "Black light...", "effects_blacklight", 6, 4, 0
        AddMenuItem "Cross-screen...", "effects_crossscreen", 6, 4, 1
        AddMenuItem "Rainbow...", "effects_rainbow", 6, 4, 2
        AddMenuItem "Sunshine...", "effects_sunshine", 6, 4, 3
        AddMenuItem "-", "-", 6, 4, 4
        AddMenuItem "Dilate...", "effects_dilate", 6, 4, 5
        AddMenuItem "Erode...", "effects_erode", 6, 4, 6
    AddMenuItem "Natural", "effects_natural", 6, 5
        AddMenuItem "Atmosphere...", "effects_atmosphere", 6, 5, 0
        AddMenuItem "Fog...", "effects_fog", 6, 5, 1
        AddMenuItem "Ignite...", "effects_ignite", 6, 5, 2
        AddMenuItem "Lava...", "effects_lava", 6, 5, 3
        AddMenuItem "Metal...", "effects_metal", 6, 5, 4
        AddMenuItem "Snow...", "effects_snow", 6, 5, 5
        AddMenuItem "Underwater...", "effects_underwater", 6, 5, 6
    AddMenuItem "Noise", "effects_noise", 6, 6
        AddMenuItem "Add film grain...", "effects_filmgrain", 6, 6, 0
        AddMenuItem "Add RGB noise...", "effects_rgbnoise", 6, 6, 1
        AddMenuItem "-", "-", 6, 6, 2
        AddMenuItem "Anisotropic diffusion...", "effects_anisotropic", 6, 6, 3
        AddMenuItem "Bilateral filter...", "effects_bilateral", 6, 6, 4
        AddMenuItem "Harmonic mean...", "effects_harmonicmean", 6, 6, 5
        AddMenuItem "Mean shift...", "effects_meanshift", 6, 6, 6
        AddMenuItem "Median...", "effects_median", 6, 6, 7
    AddMenuItem "Pixelate", "effects_pixelate", 6, 7
        AddMenuItem "Color halftone...", "effects_colorhalftone", 6, 7, 0
        AddMenuItem "Crystallize...", "effects_crystallize", 6, 7, 1
        AddMenuItem "Fragment...", "effects_fragment", 6, 7, 2
        AddMenuItem "Mezzotint...", "effects_mezzotint", 6, 7, 3
        AddMenuItem "Mosaic...", "effects_mosaic", 6, 7, 4
    AddMenuItem "Sharpen", "effects_sharpentop", 6, 8
        AddMenuItem "Sharpen...", "effects_sharpen", 6, 8, 0
        AddMenuItem "Unsharp masking...", "effects_unsharp", 6, 8, 1
    AddMenuItem "Stylize", "effects_stylize", 6, 9
        AddMenuItem "Antique...", "effects_antique", 6, 9, 0
        AddMenuItem "Diffuse...", "effects_diffuse", 6, 9, 1
        AddMenuItem "Outline...", "effects_outline", 6, 9, 2
        AddMenuItem "Palettize...", "effects_palettize", 6, 9, 3
        AddMenuItem "Portrait glow...", "effects_portraitglow", 6, 9, 4
        AddMenuItem "Solarize...", "effects_solarize", 6, 9, 5
        AddMenuItem "Twins...", "effects_twins", 6, 9, 6
        AddMenuItem "Vignetting...", "effects_vignetting", 6, 9, 7
    AddMenuItem "Transform", "effects_transform", 6, 10
        AddMenuItem "Pan and zoom...", "effects_panandzoom", 6, 10, 0
        AddMenuItem "Perspective...", "effects_perspective", 6, 10, 1
        AddMenuItem "Polar conversion...", "effects_polarconversion", 6, 10, 2
        AddMenuItem "Rotate...", "effects_rotate", 6, 10, 3
        AddMenuItem "Shear...", "effects_shear", 6, 10, 4
        AddMenuItem "Spherize...", "effects_spherize", 6, 10, 5
    AddMenuItem "-", "-", 6, 11
    AddMenuItem "Custom filter...", "effects_customfilter", 6, 12
    
    'Tools Menu
    AddMenuItem "&Tools", "tools_top", 7
    AddMenuItem "Language", "tools_language", 7, 0, , "tools_language"
    AddMenuItem "Language editor...", "tools_languageeditor", 7, 1
    AddMenuItem "-", "-", 7, 2
    AddMenuItem "Theme...", "tools_theme", 7, 3
    AddMenuItem "-", "-", 7, 4
    AddMenuItem "Create macro", "tools_macrocreatetop", 7, 5
        AddMenuItem "From history...", "tools_macrofromhistory", 7, 5, 0, "edit_history"
        AddMenuItem "-", "-", 7, 5, 1
        AddMenuItem "Start recording", "tools_recordmacro", 7, 5, 2, "macro_record"
        AddMenuItem "Stop recording...", "tools_stopmacro", 7, 5, 3, "macro_stop"
    AddMenuItem "Play macro...", "tools_playmacro", 7, 6, , "macro_play"
    AddMenuItem "Recent macros", "tools_recentmacros", 7, 7
    AddMenuItem "-", "-", 7, 8
    AddMenuItem "Options...", "tools_options", 7, 9, , "pref_advanced"
    AddMenuItem "Plugin manager...", "tools_plugins", 7, 10, , "tools_plugin"
    
    Dim debugMenuVisibility As Boolean
    debugMenuVisibility = (PD_BUILD_QUALITY <> PD_PRODUCTION) And (PD_BUILD_QUALITY <> PD_BETA)
    If debugMenuVisibility Then
        AddMenuItem "-", "-", 7, 11
        AddMenuItem "Developers", "tools_developers", 7, 12
            AddMenuItem "Theme editor...", "tools_themeeditor", 7, 12, 0, , False
            AddMenuItem "Build theme package...", "tools_themepackage", 7, 12, 1, , False
            AddMenuItem "-", "-", 7, 12, 2
            AddMenuItem "Build standalone package...", "tools_standalonepackage", 7, 12, 3, , False
        AddMenuItem "Test", "effects_developertest", 7, 13
    End If
    
    'View Menu
    AddMenuItem "&View", "view_top", 8
    AddMenuItem "&Fit image on screen", "view_fit", 8, 0, , "zoom_fit"
    AddMenuItem "-", "-", 8, 1
    AddMenuItem "Zoom &in", "view_zoomin", 8, 2, , "zoom_in"
    AddMenuItem "Zoom &out", "view_zoomout", 8, 3, , "zoom_out"
    AddMenuItem "Zoom to value", "view_zoomtop", 8, 4
        AddMenuItem "16:1 (1600%)", "zoom_16_1", 8, 4, 0
        AddMenuItem "8:1 (800%)", "zoom_8_1", 8, 4, 1
        AddMenuItem "4:1 (400%)", "zoom_4_1", 8, 4, 2
        AddMenuItem "2:1 (200%)", "zoom_2_1", 8, 4, 3
        AddMenuItem "1:1 (actual size, 100%)", "zoom_actual", 8, 4, 4, "zoom_actual"
        AddMenuItem "1:2 (50%)", "zoom_1_2", 8, 4, 5
        AddMenuItem "1:4 (25%)", "zoom_1_4", 8, 4, 6
        AddMenuItem "1:8 (12.5%)", "zoom_1_8", 8, 4, 7
        AddMenuItem "1:16 (6.25%)", "zoom_1_16", 8, 4, 8
    AddMenuItem "-", "-", 8, 5
    AddMenuItem "Show rulers", "view_rulers", 8, 6
    AddMenuItem "Show status bar", "view_statusbar", 8, 7
    
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
        AddMenuItem "Always show", "window_imagetabstrip_alwaysshow", 9, 3, 0, , False
        AddMenuItem "Show when multiple images are loaded", "window_imagetabstrip_shownormal", 9, 3, 1, , False
        AddMenuItem "Never show", "window_imagetabstrip_nevershow", 9, 3, 2, , False
        AddMenuItem "-", "-", 9, 3, 3
        AddMenuItem "Left", "window_imagetabstrip_alignleft", 9, 3, 4
        AddMenuItem "Top", "window_imagetabstrip_aligntop", 9, 3, 5
        AddMenuItem "Right", "window_imagetabstrip_alignright", 9, 3, 6
        AddMenuItem "Bottom", "window_imagetabstrip_alignbottom", 9, 3, 7
    AddMenuItem "-", "-", 9, 4
    AddMenuItem "Reset all toolboxes", "window_resetsettings", 9, 5
    
    AddMenuItem "-", "-", 9, 6
    AddMenuItem "Next image", "window_next", 9, 7, , "generic_next"
    AddMenuItem "Previous image", "window_previous", 9, 8, , "generic_previous"
    
    'Help Menu
    AddMenuItem "&Help", "help_top", 10
    AddMenuItem "Support us on Patreon...", "help_patreon", 10, 0, , "help_heart"
    AddMenuItem "Support us with a one-time donation...", "help_donate", 10, 1, , "help_heart"
    AddMenuItem "-", "-", 10, 2
    AddMenuItem "Check for &updates...", "help_checkupdates", 10, 3, , "help_update"
    AddMenuItem "Submit bug report or feedback...", "help_reportbug", 10, 4, , "help_reportbug"
    AddMenuItem "-", "-", 10, 5
    AddMenuItem "&Visit PhotoDemon website...", "help_website", 10, 6, , "help_website"
    AddMenuItem "Download PhotoDemon source code...", "help_sourcecode", 10, 7, , "help_github"
    AddMenuItem "Read license and terms of use...", "help_license", 10, 8, , "help_license"
    AddMenuItem "-", "-", 10, 9
    AddMenuItem "&About...", "help_about", 10, 10, , "help_about"
    
    'After all menu items have been added, we need to manually go through and fill the "has children" boolean
    ' for each menu entry.  (This is important because we use it when producing a searchable list of menu items,
    ' as we don't want to return un-clickable menu names in that list.)
    FinalizeMenuProperties
    
End Sub

'Internal helper function for adding a menu entry to the running collection.  Note that PD menus support a number
' of non-standard properties, all of which must be cached early in the load process so we can properly support things
' like UI themes and language translations.
Private Sub AddMenuItem(ByRef menuTextEn As String, ByRef menuName As String, ByVal topMenuID As Long, Optional ByVal subMenuID As Long = MENU_NONE, Optional ByVal subSubMenuID As Long = MENU_NONE, Optional ByRef menuImageName As String = vbNullString, Optional ByVal allowInSearches As Boolean = True)
    
    'Make sure a sufficiently large buffer exists
    Const INITIAL_MENU_COLLECTION_SIZE As Long = 128
    If (m_NumOfMenus = 0) Then
        ReDim m_Menus(0 To INITIAL_MENU_COLLECTION_SIZE - 1) As PD_MenuEntry
    Else
        If (m_NumOfMenus > UBound(m_Menus)) Then ReDim Preserve m_Menus(0 To m_NumOfMenus * 2 - 1) As PD_MenuEntry
    End If
    
    With m_Menus(m_NumOfMenus)
        .me_Name = menuName
        .me_TextEn = menuTextEn
        .me_TopMenu = topMenuID
        .me_SubMenu = subMenuID
        .me_SubSubMenu = subSubMenuID
        .me_ResImage = menuImageName
        .me_DoNotIncludeInSearch = Not allowInSearches
    End With
    
    m_NumOfMenus = m_NumOfMenus + 1

End Sub

'After adding all menu items to the master table, call this function to iterate through the final list and
' auto-populate some helpful menu properties (e.g. the "has children" bool)
Private Sub FinalizeMenuProperties()

    'First, we want to determine whether each menu has child menus.  This is important for producing
    ' menu search results, as we don't want to return "un-clickable" menus - e.g. top-level menus,
    ' or second-level menus that only exist as parents for a child menu.
    Dim i As Long
    For i = 0 To m_NumOfMenus - 1
    
        'Ensure we have an (i+1) menu index available
        If (i < m_NumOfMenus - 1) Then
        
            With m_Menus(i)
            
                'Look for top-level menus first.  They always have child menus.
                If (.me_SubMenu = MENU_NONE) Then
                    .me_HasChildren = True
                
                'Look for second-level menus with children.
                Else
                    If (.me_SubSubMenu = MENU_NONE) Then
                        If (m_Menus(i + 1).me_SubSubMenu <> MENU_NONE) Then .me_HasChildren = True
                    End If
                
                'Third-level menus are *always* clickable in PhotoDemon, so we don't need an additional Else.
                End If
            
            End With
        
        'The last menu is always Help > About.  It's clickable.
        Else
            m_Menus(i).me_HasChildren = False
        End If
    
    Next i
    
    'All child flags have been successfully marked.  Note that we can't generate menu search strings here;
    ' that can only happen *after* localization has been applied.

End Sub

'*After* all menus have been initialized, you can call this function to apply their associated icons (if any)
' to the respective menu objects.
Public Sub ApplyIconsToMenus()

    Dim i As Long
    For i = 0 To m_NumOfMenus - 1
        If (LenB(m_Menus(i).me_ResImage) <> 0) Then
            With m_Menus(i)
                IconsAndCursors.AddMenuIcon .me_ResImage, .me_TopMenu, .me_SubMenu, .me_SubSubMenu
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
                    .me_TextTranslated = newCaptionEn
                Else
                    .me_TextEn = newCaptionEn
                    .me_TextTranslated = g_Language.TranslateMessage(newCaptionEn)
                End If
                
                'Deal with trailing accelerator text, if any
                If (LenB(.me_HotKeyTextTranslated) <> 0) Then
                    .me_TextFinal = .me_TextTranslated & vbTab & .me_HotKeyTextTranslated
                Else
                    .me_TextFinal = .me_TextTranslated
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
        PDDebug.LogAction "WARNING!  g_Language isn't available, so hotkey captions won't be correct."
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
                If (.me_Name <> "-") Then
                
                    'Update the actual caption text
                    If (LenB(.me_TextEn) <> 0) Then
                    
                        .me_TextTranslated = g_Language.TranslateMessage(.me_TextEn)
                    
                        'Update the appended hotkey text, if any
                        If (.me_HotKeyCode <> 0) Then
                            .me_HotKeyTextTranslated = GetHotkeyText(.me_HotKeyCode, .me_HotKeyShift)
                            .me_TextFinal = .me_TextTranslated & vbTab & .me_HotKeyTextTranslated
                        Else
                            .me_TextFinal = .me_TextTranslated
                        End If
                        
                    Else
                        .me_TextTranslated = vbNullString
                        .me_TextFinal = vbNullString
                    End If
                    
                Else
                    .me_TextFinal = vbNullString
                End If
                    
            End With
            
        Next i
        
    Else
    
        For i = 0 To m_NumOfMenus - 1
        
            With m_Menus(i)
            
                If (.me_Name <> "-") Then
                
                    .me_TextTranslated = .me_TextEn
                    
                    If (.me_HotKeyCode <> 0) Then
                        .me_HotKeyTextTranslated = GetHotkeyText(.me_HotKeyCode, .me_HotKeyShift)
                        .me_TextFinal = .me_TextTranslated & vbTab & .me_HotKeyTextTranslated
                    Else
                        .me_TextFinal = .me_TextTranslated
                    End If
                    
                Else
                    .me_TextFinal = vbNullString
                End If
                
            End With
            
        Next i
    
    End If
    
    'With all menu captions updated, we now need to relay those changes to the underlying API menu struct
    For i = 0 To m_NumOfMenus - 1
        If (LenB(m_Menus(i).me_TextFinal) <> 0) Then UpdateMenuText_ByIndex i
    Next i
    
    'We also need to update search strings for all menus
    Dim mnuNameLvl1 As String, mnuNameLvl2 As String, mnuNameFinal As String, lastLvl2Index As Long
    For i = 0 To m_NumOfMenus - 1
        
        With m_Menus(i)
            
            'If this menu has children, we need to update our parent menu name trackers,
            ' but we *don't* need to produce a search string (as we don't want to return
            ' un-clickable parent menus in search results).
            If .me_HasChildren Then
                
                .me_TextSearchable = vbNullString
                
                lastLvl2Index = .me_SubMenu
                If (.me_SubMenu = MENU_NONE) Then
                    mnuNameLvl1 = Replace$(.me_TextTranslated, "&", vbNullString)
                    mnuNameLvl2 = vbNullString
                Else
                    mnuNameLvl2 = Replace$(.me_TextTranslated, "&", vbNullString)
                End If
            
            'This menu does not have children, meaning it's clickable.
            Else
                
                'If this menu doesn't match the last level-2 menu index, reset the level 2 string.
                ' (This ensures that items following a sub-menu - e.g. File > Print, which comes after
                ' File > Batch Operations - don't mistakenly pick-up the second-level name of a
                ' preview menu.)
                If (.me_SubMenu <> lastLvl2Index) Then mnuNameLvl2 = vbNullString
                
                'Make sure this isn't just a menu separator
                If (.me_Name <> "-") Then
                
                    'Append first- and second-level menu names, if any
                    If (LenB(mnuNameLvl1) <> 0) Then mnuNameFinal = mnuNameLvl1 & " > "
                    If (LenB(mnuNameLvl2) <> 0) Then mnuNameFinal = mnuNameFinal & mnuNameLvl2 & " > "
                    mnuNameFinal = mnuNameFinal & Replace$(.me_TextTranslated, "&", vbNullString)
                    .me_TextSearchable = mnuNameFinal
                    
                Else
                    .me_TextSearchable = vbNullString
                End If
            
            End If
            
        End With
        
    Next i
    
    'Some special menus must be dealt with now; note that some menus are already handled by dedicated callers
    ' (e.g. the "Languages" menu), while others must be handled here.
    Menus.UpdateSpecialMenu_RecentFiles
    Menus.UpdateSpecialMenu_RecentMacros
    
    If redrawMenuBar Then DrawMenuBar FormMain.hWnd
    
End Sub

'Given a menu name, return the corresponding menu caption (localized, with accelerator)
Public Function GetCaptionFromName(ByRef mnuName As String, Optional ByVal returnTranslation As Boolean = True) As String

    'Resolve the menu name into an index into our menu collection
    Dim i As Long
    Dim mnuIndex As Long: mnuIndex = -1
    
    For i = 0 To m_NumOfMenus - 1
        If Strings.StringsEqual(mnuName, m_Menus(i).me_Name, True) Then
            mnuIndex = i
            Exit For
        End If
    Next i
    
    If (mnuIndex >= 0) Then
        If returnTranslation Then
            GetCaptionFromName = m_Menus(mnuIndex).me_TextTranslated
        Else
            GetCaptionFromName = m_Menus(mnuIndex).me_TextEn
        End If
    End If
    
End Function

'Return a list of searchable menu strings.  Matches to this list can then be passed back to this module and
' matched against their respective menu(s).
Public Function GetSearchableMenuList(ByRef dstStack As pdStringStack, Optional ByVal ignoreDisabledMenus As Boolean = True, Optional ByVal restrictToThisTopMenuIndex As Long = -1) As Boolean
    
    Set dstStack = New pdStringStack
    
    Dim i As Long, allowedToAdd As Boolean
    For i = 0 To m_NumOfMenus - 1
        If (LenB(m_Menus(i).me_TextSearchable) <> 0) Then
            
            allowedToAdd = True
            If ignoreDisabledMenus Then allowedToAdd = Menus.IsMenuEnabled(m_Menus(i).me_Name)
            If (restrictToThisTopMenuIndex >= 0) Then allowedToAdd = allowedToAdd And (m_Menus(i).me_TopMenu = restrictToThisTopMenuIndex)
            allowedToAdd = allowedToAdd And (Not m_Menus(i).me_DoNotIncludeInSearch)
            
            If allowedToAdd Then dstStack.AddString m_Menus(i).me_TextSearchable
            
        End If
    Next i
    
    GetSearchableMenuList = (dstStack.GetNumOfStrings > 0)
    
End Function

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
            .me_HotKeyCode = vKeyCode
            .me_HotKeyShift = Shift
            .me_HotKeyTextTranslated = hotkeyText
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
        If Strings.StringsEqual(mnuName, m_Menus(i).me_Name, True) Then
            mnuIndex = i
            Exit For
        End If
    Next i
    
    If (mnuIndex >= 0) Then
    
        'We now need to check all parent menus in turn (because they may be disabled, which effectively
        ' means *we're* disabled too - but the API doesn't calculate that for us.)
        
        'Start by getting an hMenu for our top-level parent
        Dim hMenu As Long, hIndex As Long, mFlags As Win32_MenuStateFlags
        hMenu = GetMenu(FormMain.hWnd)
        hIndex = m_Menus(mnuIndex).me_TopMenu
        
        If (hMenu <> 0) And (hIndex >= 0) Then
            mFlags = GetMenuState(hMenu, hIndex, MF_BYPOSITION)
            IsMenuEnabled = Not ((mFlags And (MF_DISABLED Or MF_GRAYED)) <> 0)
        End If
        
        'If our top-level menu is enabled, check sub-menus, if any
        If IsMenuEnabled And (m_Menus(mnuIndex).me_SubMenu <> MENU_NONE) Then
            
            hMenu = GetSubMenu(hMenu, m_Menus(mnuIndex).me_TopMenu)
            hIndex = m_Menus(mnuIndex).me_SubMenu
            
            If (hMenu <> 0) And (hIndex >= 0) Then
                mFlags = GetMenuState(hMenu, hIndex, MF_BYPOSITION)
                IsMenuEnabled = Not ((mFlags And (MF_DISABLED Or MF_GRAYED)) <> 0)
            End If
            
            'If our sub-level menu parent is enabled, check us last (as necessary)
            If IsMenuEnabled And (m_Menus(mnuIndex).me_SubSubMenu <> MENU_NONE) Then
            
                hMenu = GetSubMenu(hMenu, m_Menus(mnuIndex).me_SubMenu)
                hIndex = m_Menus(mnuIndex).me_SubSubMenu
                
                If (hMenu <> 0) And (hIndex >= 0) Then
                    mFlags = GetMenuState(hMenu, hIndex, MF_BYPOSITION)
                    IsMenuEnabled = Not ((mFlags And (MF_DISABLED Or MF_GRAYED)) <> 0)
                End If
            
            End If
            
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
        
        'Special hotkeys
        .AddAccelerator vbKeyF, vbCtrlMask, "tool_search", , False, False, False
        
        'Tool hotkeys (e.g. keys not associated with menus)
        .AddAccelerator vbKeyH, , "tool_activate_hand", , , , False
        .AddAccelerator vbKeyM, , "tool_activate_move", , , , False
        .AddAccelerator vbKeyI, , "tool_activate_colorpicker", , , , False
        
        'Note that some hotkeys do double-duty in tool selection; you can press some of these shortcuts multiple times
        ' to toggle between similar tools (e.g. rectangular and elliptical selections).  Details can be found in
        ' FormMain.pdHotkey event handlers.
        .AddAccelerator vbKeyS, , "tool_activate_selectrect", , , , False
        .AddAccelerator vbKeyL, , "tool_activate_selectlasso", , , , False
        .AddAccelerator vbKeyW, , "tool_activate_selectwand", , , , False
        .AddAccelerator vbKeyT, , "tool_activate_text", , , , False
        .AddAccelerator vbKeyP, , "tool_activate_pencil", , , , False
        .AddAccelerator vbKeyB, , "tool_activate_brush", , , , False
        .AddAccelerator vbKeyE, , "tool_activate_eraser", , , , False
        .AddAccelerator vbKeyF, , "tool_activate_fill", , , , False
        .AddAccelerator vbKeyG, , "tool_activate_gradient", , , , False
        
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
        
        .AddAccelerator vbKeyF, vbCtrlMask Or vbShiftMask, "Repeat last action", "edit_repeat", True, True, False, UNDO_Image
        
        .AddAccelerator vbKeyX, vbCtrlMask, "Cut", "edit_cut", True, True, False, UNDO_Image
        
        'This "cut from layer" hotkey combination is used as "crop to selection" in other software; as such,
        ' I am suspending this instance for now.
        '.AddAccelerator vbKeyX, vbCtrlMask Or vbShiftMask, "Cut from layer", "edit_cutlayer", True, True, False, UNDO_Layer
        .AddAccelerator vbKeyC, vbCtrlMask, "Copy", "edit_copy", True, True, False, UNDO_Nothing
        .AddAccelerator vbKeyC, vbCtrlMask Or vbShiftMask, "Copy from layer", "edit_copylayer", True, True, False, UNDO_Nothing
        .AddAccelerator vbKeyV, vbCtrlMask, "Paste as new image", "edit_pasteasimage", True, False, False, UNDO_Nothing
        .AddAccelerator vbKeyV, vbCtrlMask Or vbShiftMask, "Paste as new layer", "edit_pasteaslayer", True, False, False, UNDO_Image_VectorSafe
        
        'Image menu
        .AddAccelerator vbKeyA, vbCtrlMask Or vbShiftMask, "Duplicate image", "image_duplicate", True, True, False, UNDO_Nothing
        .AddAccelerator vbKeyR, vbCtrlMask, "Resize image", "image_resize", True, True, True, UNDO_Image
        .AddAccelerator vbKeyR, vbCtrlMask Or vbAltMask, "Canvas size", "image_canvassize", True, True, True, UNDO_ImageHeader
        .AddAccelerator vbKeyX, vbCtrlMask Or vbShiftMask, "Crop", "image_crop", True, True, False, UNDO_Image
        .AddAccelerator vbKeyX, vbCtrlMask Or vbAltMask, "Trim empty borders", "image_trim", True, True, False, UNDO_ImageHeader
        
            'Image -> Rotate submenu
            '.AddAccelerator vbKeyR, 0, "Rotate image 90 clockwise", "image_rotate90", True, True, False, UNDO_Image
            '.AddAccelerator vbKeyL, 0, "Rotate image 90 counter-clockwise", "image_rotate270", True, True, False, UNDO_Image
            .AddAccelerator vbKeyR, vbCtrlMask Or vbShiftMask Or vbAltMask, "Arbitrary image rotation", "image_rotatearbitrary", True, True, True, UNDO_Nothing
        
        'Layer Menu
        .AddAccelerator vbKeyE, vbCtrlMask Or vbShiftMask, "Merge visible layers", "layer_mergevisible", True, True, False, UNDO_Image
        .AddAccelerator vbKeyF, vbCtrlMask Or vbShiftMask, "Flatten image", "layer_flatten", True, True, True, UNDO_Nothing
        
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
        If Strings.StringsEqual(menuID, m_Menus(i).me_Name, True) Then
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
        If (m_Menus(i).me_TopMenu = topMenuID) Then
            
            'Submenu IDs are only matched if the user specifically requests it
            If (subMenuID <> IGNORE_MENU_ID) Then
                
                'Match the submenu ID
                If (m_Menus(i).me_SubMenu = subMenuID) Then
                    
                    'Match the subsubmenu ID
                    If (subSubMenuID <> IGNORE_MENU_ID) Then
                        
                        If (m_Menus(i).me_SubSubMenu = subSubMenuID) Then
                            m_Menus(i).me_TopMenu = REMOVED_MENU_ID
                        ElseIf (subSubMenuID = ALL_MENU_SUBITEMS) And (m_Menus(i).me_SubSubMenu >= 0) Then
                            m_Menus(i).me_TopMenu = REMOVED_MENU_ID
                        End If
                    
                    Else
                        m_Menus(i).me_TopMenu = REMOVED_MENU_ID
                    End If
                    
                ElseIf (subMenuID = ALL_MENU_SUBITEMS) And (m_Menus(i).me_SubMenu >= 0) Then
                    m_Menus(i).me_TopMenu = REMOVED_MENU_ID
                End If
            
            Else
                m_Menus(i).me_TopMenu = REMOVED_MENU_ID
            End If
            
        End If
    Next i
    
    'All menus to be removed have now been properly flagged.  Iterate through the list and fill all empty spots.
    Dim moveOffset As Long
    moveOffset = 0
    
    For i = 0 To m_NumOfMenus - 1
    
        'If this item is set to be deleted, increment our move counter
        If (m_Menus(i).me_TopMenu = REMOVED_MENU_ID) Then
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

'Given a menu caption, apply the corresponding default processor action.
Public Sub ProcessDefaultAction_ByCaption(ByRef srcMenuCaption As String)
    
    'Search the menu list for a menu caption matching the passed one
    Dim i As Long
    For i = 0 To m_NumOfMenus - 1
    
        'If the captions match, trigger the corresponding default action, then exit immediately
        If Strings.StringsEqual(srcMenuCaption, m_Menus(i).me_TextEn, True) Then
            ProcessDefaultAction_ByName m_Menus(i).me_Name
            Exit Sub
        End If
    
    Next i
    
    'If the previous loop found no matches, something went horribly wrong
    PDDebug.LogAction "WARNING!  Menus.ProcessDefaultAction_ByCaption couldn't find a match for: " & srcMenuCaption

End Sub

'Given a menu search string, apply the corresponding default processor action.
Public Sub ProcessDefaultAction_BySearch(ByRef srcSearchText As String)
    
    'Search the menu list for a menu caption matching the passed one
    Dim i As Long
    For i = 0 To m_NumOfMenus - 1
    
        'If the captions match, trigger the corresponding default action, then exit immediately
        If Strings.StringsEqual(srcSearchText, m_Menus(i).me_TextSearchable, True) Then
            ProcessDefaultAction_ByName m_Menus(i).me_Name
            Exit Sub
        End If
    
    Next i
    
    'If the previous loop found no matches, something went horribly wrong
    PDDebug.LogAction "WARNING!  Menus.ProcessDefaultAction_BySearch couldn't find a match for: " & srcSearchText

End Sub

'Given a menu name, apply the corresponding default processor action.
Public Sub ProcessDefaultAction_ByName(ByRef srcMenuName As String)
    
    'Helper functions exist for each main menu; once a command is located, we can stop searching.
    Dim cmdFound As Boolean: cmdFound = False
    
    'Search each menu group in turn
    If (Not cmdFound) Then cmdFound = PDA_ByName_MenuFile(srcMenuName)
    If (Not cmdFound) Then cmdFound = PDA_ByName_MenuEdit(srcMenuName)
    If (Not cmdFound) Then cmdFound = PDA_ByName_MenuImage(srcMenuName)
    If (Not cmdFound) Then cmdFound = PDA_ByName_MenuLayer(srcMenuName)
    If (Not cmdFound) Then cmdFound = PDA_ByName_MenuSelect(srcMenuName)
    If (Not cmdFound) Then cmdFound = PDA_ByName_MenuAdjustments(srcMenuName)
    If (Not cmdFound) Then cmdFound = PDA_ByName_MenuEffects(srcMenuName)
    If (Not cmdFound) Then cmdFound = PDA_ByName_MenuTools(srcMenuName)
    If (Not cmdFound) Then cmdFound = PDA_ByName_MenuView(srcMenuName)
    If (Not cmdFound) Then cmdFound = PDA_ByName_MenuWindow(srcMenuName)
    If (Not cmdFound) Then cmdFound = PDA_ByName_MenuHelp(srcMenuName)
    If (Not cmdFound) Then cmdFound = PDA_ByName_NonMenu(srcMenuName)
    
    'Failsafe check to make sure we found *something*
    If (Not cmdFound) Then PDDebug.LogAction "WARNING: Menus.ProcessDefaultAction_ByName received an unknown request: " & srcMenuName
    
End Sub

Private Function PDA_ByName_MenuFile(ByRef srcMenuName As String) As Boolean

    Dim cmdFound As Boolean: cmdFound = True
    
    Select Case srcMenuName
    
        Case "file_new"
            Process "New image", True
            
        Case "file_open"
            Process "Open", True
            
        Case "file_openrecent"
            If (LenB(g_RecentFiles.GetFullPath(0)) <> 0) Then Loading.LoadFileAsNewImage g_RecentFiles.GetFullPath(0)
        
        Case "file_import"
            Case "file_import_paste"
                Process "Paste as new image", False, , UNDO_Nothing, , False
                
            Case "file_import_scanner"
                Process "Scan image", True
                
            Case "file_import_selectscanner"
                Process "Select scanner or camera", True
                
            Case "file_import_web"
                Process "Internet import", True
                
            Case "file_import_screenshot"
                Process "Screen capture", True
                
        Case "file_close"
            Process "Close", True
            
        Case "file_closeall"
            Process "Close all", True
            
        Case "file_save"
            Process "Save", True
            
        Case "file_savecopy"
            Process "Save copy", True
            
        Case "file_saveas"
            Process "Save as", True
            
        Case "file_revert"
            Process "Revert", False, , UNDO_Everything
            
        Case "file_export"
            Case "file_export_colorprofile"
                Process "Export color profile", True
                
            Case "file_export_palette"
                Process "Export palette", True
                
        Case "file_batch"
            Case "file_batch_process"
                Process "Batch wizard", True
                
            Case "file_batch_repair"
                ShowPDDialog vbModal, FormBatchRepair
                
        Case "file_print"
            Process "Print", True
            
        Case "file_quit"
            Process "Exit program", True
            
        Case Else
            cmdFound = False
        
    End Select
    
    PDA_ByName_MenuFile = cmdFound
    
End Function

Private Function PDA_ByName_MenuEdit(ByRef srcMenuName As String) As Boolean

    Dim cmdFound As Boolean: cmdFound = True
    
    Select Case srcMenuName
    
        Case "edit_undo"
            Process "Undo", False
            
        Case "edit_redo"
            Process "Redo", False
            
        Case "edit_history"
            Process "Undo history", True
            
        'TODO: figure out Undo handling for "Repeat last action"... can we always reuse the undo type of
        ' the previous action?  Could this have unforeseen consequences?
        Case "edit_repeat"
            Process "Repeat last action", False, , UNDO_Image
            
        Case "edit_fade"
            Process "Fade", True
            
        Case "edit_cut"
            Process "Cut", False, , UNDO_Image, , True
        
        'If a selection is active, the Undo/Redo engine can simply back up the current layer contents.
        ' If, however, no selection is active, we will delete the entire layer.  That requires a backup
        ' of the full layer stack.
        Case "edit_cutlayer"
            If PDImages.GetActiveImage.IsSelectionActive Then
                Process "Cut from layer", False, , UNDO_Layer, , True
            Else
                Process "Cut from layer", False, , UNDO_Image, , True
            End If
            
        Case "edit_copy"
            Process "Copy", False, , UNDO_Nothing, , False
            
        Case "edit_copylayer"
            Process "Copy from layer", False, , UNDO_Nothing, , False
            
        Case "edit_pasteasimage"
            Process "Paste as new image", False, , UNDO_Nothing, , False
            
        Case "edit_pasteaslayer"
            Process "Paste as new layer", False, , UNDO_Image_VectorSafe, , False
            
        Case "edit_emptyclipboard"
            Process "Empty clipboard", False, , UNDO_Nothing, , False
            
        Case Else
            cmdFound = False
            
    End Select
    
    PDA_ByName_MenuEdit = cmdFound
    
End Function

Private Function PDA_ByName_MenuImage(ByRef srcMenuName As String) As Boolean

    Dim cmdFound As Boolean: cmdFound = True
    
    Select Case srcMenuName
    
        Case "image_duplicate"
            Process "Duplicate image", , , UNDO_Nothing
            
        Case "image_resize"
            Process "Resize image", True
            
        Case "image_contentawareresize"
            Process "Content-aware image resize", True
            
        Case "image_canvassize"
            Process "Canvas size", True
            
        Case "image_fittolayer"
            Process "Fit canvas to layer", False, BuildParamList("targetlayer", PDImages.GetActiveImage.GetActiveLayerIndex), UNDO_ImageHeader
            
        Case "image_fitalllayers"
            Process "Fit canvas to all layers", False, , UNDO_ImageHeader
            
        Case "image_crop"
            Process "Crop", True
            
        Case "image_trim"
            Process "Trim empty borders", , , UNDO_ImageHeader
            
        Case "image_rotate"
            Case "image_straighten"
                Process "Straighten image", True
                
            Case "image_rotate90"
                Process "Rotate image 90 clockwise", , , UNDO_Image
                
            Case "image_rotate270"
                Process "Rotate image 90 counter-clockwise", , , UNDO_Image
                
            Case "image_rotate180"
                Process "Rotate image 180", , , UNDO_Image
                
            Case "image_rotatearbitrary"
                Process "Arbitrary image rotation", True
                
        Case "image_fliphorizontal"
            Process "Flip image horizontally", , , UNDO_Image
            
        Case "image_flipvertical"
            Process "Flip image vertically", , , UNDO_Image
            
        Case "image_metadata"
            Case "image_editmetadata"
                Process "Edit metadata", True
                
            Case "image_removemetadata"
                Process "Remove all metadata", False, , UNDO_ImageHeader
                
            Case "image_countcolors"
                Process "Count image colors", True
                
            Case "image_maplocation"
                Web.MapImageLocation
                
        Case Else
            cmdFound = False
                
    End Select
    
    PDA_ByName_MenuImage = cmdFound
    
End Function

Private Function PDA_ByName_MenuLayer(ByRef srcMenuName As String) As Boolean

    Dim cmdFound As Boolean: cmdFound = True
    
    Select Case srcMenuName
    
        Case "layer_add"
            Case "layer_addbasic"
                Process "Add new layer", True
                
            Case "layer_addblank"
                Process "Add blank layer", False, BuildParamList("targetlayer", PDImages.GetActiveImage.GetActiveLayerIndex), UNDO_Image_VectorSafe
                
            Case "layer_duplicate"
                Process "Duplicate Layer", False, BuildParamList("targetlayer", PDImages.GetActiveImage.GetActiveLayerIndex), UNDO_Image_VectorSafe
                
            Case "layer_addfromclipboard"
                Process "Paste as new layer", False, , UNDO_Image_VectorSafe, , False
                
            Case "layer_addfromfile"
                Process "New layer from file", True
                
            Case "layer_addfromvisiblelayers"
                Process "New layer from visible layers", False, , UNDO_Image_VectorSafe
                
        Case "layer_delete"
            Case "layer_deletecurrent"
                Process "Delete layer", False, BuildParamList("layerindex", PDImages.GetActiveImage.GetActiveLayerIndex), UNDO_Image_VectorSafe
                
            Case "layer_deletehidden"
                Process "Delete hidden layers", False, , UNDO_Image_VectorSafe
                
        Case "layer_mergeup"
            Process "Merge layer up", False, BuildParamList("layerindex", PDImages.GetActiveImage.GetActiveLayerIndex), UNDO_Image
            
        Case "layer_mergedown"
            Process "Merge layer down", False, BuildParamList("layerindex", PDImages.GetActiveImage.GetActiveLayerIndex), UNDO_Image
            
        Case "layer_order"
            Case "layer_up"
                Process "Raise layer", False, BuildParamList("layerindex", PDImages.GetActiveImage.GetActiveLayerIndex), UNDO_ImageHeader
                
            Case "layer_down"
                Process "Lower layer", False, BuildParamList("layerindex", PDImages.GetActiveImage.GetActiveLayerIndex), UNDO_ImageHeader
                
            Case "layer_totop"
                Process "Raise layer to top", False, BuildParamList("layerindex", PDImages.GetActiveImage.GetActiveLayerIndex), UNDO_ImageHeader
                
            Case "layer_tobottom"
                Process "Lower layer to bottom", False, BuildParamList("layerindex", PDImages.GetActiveImage.GetActiveLayerIndex), UNDO_ImageHeader
                
        Case "layer_orientation"
            Case "layer_straighten"
                Process "Straighten layer", True
                
            Case "layer_rotate90"
                Process "Rotate layer 90 clockwise", , , UNDO_Layer
                
            Case "layer_rotate270"
                Process "Rotate layer 90 counter-clockwise", , , UNDO_Layer
                
            Case "layer_rotate180"
                Process "Rotate layer 180", , , UNDO_Layer
                
            Case "layer_rotatearbitrary"
                Process "Arbitrary layer rotation", True
                
            Case "layer_fliphorizontal"
                Process "Flip layer horizontally", , , UNDO_Layer
                
            Case "layer_flipvertical"
                Process "Flip layer vertically", , , UNDO_Layer
                
        Case "layer_resize"
            Case "layer_resetsize"
                Process "Reset layer size", False, BuildParamList("layerindex", PDImages.GetActiveImage.GetActiveLayerIndex), UNDO_LayerHeader
                
            Case "layer_resize"
                Process "Resize layer", True
                
            Case "layer_contentawareresize"
                Process "Content-aware layer resize", True
                
        Case "layer_crop"
            Process "Crop layer to selection", , , UNDO_Layer
            
        Case "layer_transparency"
            Case "layer_colortoalpha"
                Process "Color to alpha", True
                
            Case "layer_removealpha"
                Process "Remove alpha channel", True
                
        Case "layer_rasterize"
            Case "layer_rasterizecurrent"
                Process "Rasterize layer", , , UNDO_Layer
                
            Case "layer_rasterizeall"
                Process "Rasterize all layers", , , UNDO_Image
                
        Case "layer_mergevisible"
            Process "Merge visible layers", , , UNDO_Image
            
        Case "layer_flatten"
            Process "Flatten image", True
            
        Case Else
            cmdFound = False
            
    End Select
    
    PDA_ByName_MenuLayer = cmdFound
    
End Function

Private Function PDA_ByName_MenuSelect(ByRef srcMenuName As String) As Boolean

    Dim cmdFound As Boolean: cmdFound = True
    
    Select Case srcMenuName
    
        Case "select_all"
            Process "Select all", , , UNDO_Selection
            
        Case "select_none"
            Process "Remove selection", , , UNDO_Selection
            
        Case "select_invert"
            Process "Invert selection", , , UNDO_Selection
            
        Case "select_grow"
            Process "Grow selection", True
            
        Case "select_shrink"
            Process "Shrink selection", True
            
        Case "select_border"
            Process "Border selection", True
            
        Case "select_feather"
            Process "Feather selection", True
            
        Case "select_sharpen"
            Process "Sharpen selection", True
            
        Case "select_erasearea"
            Process "Erase selected area", False, BuildParamList("targetlayer", PDImages.GetActiveImage.GetActiveLayerIndex), UNDO_Layer
            
        Case "select_load"
            Process "Load selection", True
            
        Case "select_save"
            Process "Save selection", True
            
        Case "select_export"
            Case "select_exportarea"
                Process "Export selected area as image", True
                
            Case "select_exportmask"
                Process "Export selection mask as image", True
                
        Case Else
            cmdFound = False
                
    End Select
    
    PDA_ByName_MenuSelect = cmdFound
    
End Function

Private Function PDA_ByName_MenuAdjustments(ByRef srcMenuName As String) As Boolean

    Dim cmdFound As Boolean: cmdFound = True
    
    Select Case srcMenuName
    
        Case "adj_autocorrect"
            Process "Auto correct", False, , UNDO_Layer
            
        Case "adj_autoenhance"
            Process "Auto enhance", False, , UNDO_Layer
            
        Case "adj_blackandwhite"
            Process "Black and white", True
            
        Case "adj_bandc"
            Process "Brightness and contrast", True
            
        Case "adj_colorbalance"
            Process "Color balance", True
            
        Case "adj_curves"
            Process "Curves", True
            
        Case "adj_levels"
            Process "Levels", True
            
        Case "adj_sandh"
            Process "Shadow and highlight", True
            
        Case "adj_vibrance"
            Process "Vibrance", True
            
        Case "adj_whitebalance"
            Process "White balance", True
            
        Case "adj_channels"
            Case "adj_channelmixer"
                Process "Channel mixer", True
                
            Case "adj_rechannel"
                Process "Rechannel", True
                
            Case "adj_maxchannel"
                Process "Maximum channel", , , UNDO_Layer
                
            Case "adj_minchannel"
                Process "Minimum channel", , , UNDO_Layer
                
            Case "adj_shiftchannelsleft"
                Process "Shift colors (left)", , , UNDO_Layer
                
            Case "adj_shiftchannelsright"
                Process "Shift colors (right)", , , UNDO_Layer
                
        Case "adj_color"
        
            'Case "adj_colorbalance"    'Covered by parent menu
            
            'Case "adj_whitebalance"    'Covered by parent menu
            
            Case "adj_hsl"
                Process "Hue and saturation", True
                
            Case "adj_temperature"
                Process "Temperature", True
                
            Case "adj_tint"
                Process "Tint", True
                
            'Case "adj_vibrance"        'Covered by parent menu
            
            'Case "adj_blackandwhite"   'Covered by parent menu
            
            Case "adj_colorize"
                Process "Colorize", True
                
            Case "adj_replacecolor"
                Process "Replace color", True
                
            Case "adj_sepia"
                Process "Sepia", , , UNDO_Layer
                
        Case "adj_histogram"
            Case "adj_histogramdisplay"
                ShowPDDialog vbModal, FormHistogram
                
            Case "adj_histogramequalize"
                Process "Equalize", True
                
            Case "adj_histogramstretch"
                Process "Stretch histogram", , , UNDO_Layer
                
        Case "adj_invert"
            Case "adj_invertcmyk"
                Process "Film negative", , , UNDO_Layer
                
            Case "adj_inverthue"
                Process "Invert hue", , , UNDO_Layer
                
            Case "adj_invertrgb"
                Process "Invert RGB", , , UNDO_Layer
                
        Case "adj_lighting"
            'Case "adj_bandc"   'Covered by parent menu
            
            'Case "adj_curves"  'Covered by parent menu
            
            Case "adj_gamma"
                Process "Gamma", True
                
            'Case "adj_levels"  'Covered by parent menu
            
            'Case "adj_sandh"   'Covered by parent menu
            
        Case "adj_monochrome"
            Case "adj_colortomonochrome"
                Process "Color to monochrome", True
                
            Case "adj_monochrometogray"
                Process "Monochrome to gray", True
                
        Case "adj_photo"
            Case "adj_exposure"
                Process "Exposure", True
                
            Case "adj_hdr"
                Process "HDR", True
                
            Case "adj_photofilters"
                Process "Photo filter", True
                
            Case "adj_redeyeremoval"
                Process "Red-eye removal", True
                
            Case "adj_splittone"
                Process "Split toning", True
                
        Case Else
            cmdFound = False
                
    End Select
    
    PDA_ByName_MenuAdjustments = cmdFound
    
End Function

Private Function PDA_ByName_MenuEffects(ByRef srcMenuName As String) As Boolean

    Dim cmdFound As Boolean: cmdFound = True
    
    Select Case srcMenuName
    
        Case "effects_artistic"
            Case "effects_colorpencil"
                Process "Colored pencil", True
                
            Case "effects_comicbook"
                Process "Comic book", True
                
            Case "effects_figuredglass"
                Process "Figured glass", True
                
            Case "effects_filmnoir"
                Process "Film noir", True
                
            Case "effects_glasstiles"
                Process "Glass tiles", True
                
            Case "effects_kaleidoscope"
                Process "Kaleidoscope", True
                
            Case "effects_modernart"
                Process "Modern art", True
                
            Case "effects_oilpainting"
                Process "Oil painting", True
                
            Case "effects_plasticwrap"
                Process "Plastic wrap", True
                
            Case "effects_posterize"
                Process "Posterize", True
                
            Case "effects_relief"
                Process "Relief", True
                
            Case "effects_stainedglass"
                Process "Stained glass", True
                
        Case "effects_blur"
            Case "effects_boxblur"
                Process "Box blur", True
                
            Case "effects_gaussianblur"
                Process "Gaussian blur", True
                
            Case "effects_surfaceblur"
                Process "Surface blur", True
                
            Case "effects_motionblur"
                Process "Motion blur", True
                
            Case "effects_radialblur"
                Process "Radial blur", True
                
            Case "effects_zoomblur"
                Process "Zoom blur", True
                
            Case "effects_kuwahara"
                Process "Kuwahara filter", True
                
            Case "effects_snn"
                Process "Symmetric nearest-neighbor", True
                
        Case "effects_distort"
            Case "effects_fixlensdistort"
                Process "Correct lens distortion", True
                
            Case "effects_donut"
                Process "Donut", True
                
            Case "effects_lens"
                Process "Apply lens distortion", True
                
            Case "effects_pinchandwhirl"
                Process "Pinch and whirl", True
                
            Case "effects_poke"
                Process "Poke", True
                
            Case "effects_ripple"
                Process "Ripple", True
                
            Case "effects_squish"
                Process "Squish", True
                
            Case "effects_swirl"
                Process "Swirl", True
                
            Case "effects_waves"
                Process "Waves", True
                
            Case "effects_miscdistort"
                Process "Miscellaneous distort", True
                
        Case "effects_edges"
            Case "effects_emboss"
                Process "Emboss", True
                
            Case "effects_enhanceedges"
                Process "Enhance edges", True
                
            Case "effects_findedges"
                Process "Find edges", True
                
            Case "effects_rangefilter"
                Process "Range filter", True
                
            Case "effects_tracecontour"
                Process "Trace contour", True
                
        Case "effects_lightandshadow"
            Case "effects_blacklight"
                Process "Black light", True
                
            Case "effects_crossscreen"
                Process "Cross-screen", True
            
            Case "effects_rainbow"
                Process "Rainbow", True
                
            Case "effects_sunshine"
                Process "Sunshine", True
                
            Case "effects_dilate"
                Process "Dilate (maximum rank)", True
                
            Case "effects_erode"
                Process "Erode (minimum rank)", True
                
        Case "effects_natural"
            Case "effects_atmosphere"
                Process "Atmosphere", True
                
            Case "effects_fog"
                Process "Fog", True
                
            Case "effects_ignite"
                Process "Ignite", True
                
            Case "effects_lava"
                Process "Lava", True
                
            Case "effects_metal"
                Process "Metal", True
                
            Case "effects_snow"
                Process "Snow", True
                
            Case "effects_underwater"
                Process "Water", True
                
        Case "effects_noise"
            Case "effects_filmgrain"
                Process "Add film grain", True
                
            Case "effects_rgbnoise"
                Process "Add RGB noise", True
                
            Case "effects_anisotropic"
                Process "Anisotropic diffusion", True
                
            Case "effects_bilateral"
                Process "Bilateral smoothing", True
                
            Case "effects_harmonicmean"
                Process "Harmonic mean", True
                
            Case "effects_meanshift"
                Process "Mean shift", True
                
            Case "effects_median"
                Process "Median", True
                
        Case "effects_pixelate"
            Case "effects_colorhalftone"
                Process "Color halftone", True
                
            Case "effects_crystallize"
                Process "Crystallize", True
                
            Case "effects_fragment"
                Process "Fragment", True
                
            Case "effects_mezzotint"
                Process "Mezzotint", True
                
            Case "effects_mosaic"
                Process "Mosaic", True
                
        Case "effects_sharpentop"
            Case "effects_sharpen"
                Process "Sharpen", True
                
            Case "effects_unsharp"
                Process "Unsharp mask", True
                
        Case "effects_stylize"
            Case "effects_antique"
                Process "Antique", True
                
            Case "effects_diffuse"
                Process "Diffuse", True
                
            Case "effects_outline"
                Process "Outline", True
                
            Case "effects_palettize"
                Process "Palettize", True
                
            Case "effects_portraitglow"
                Process "Portrait glow", True
                
            Case "effects_solarize"
                Process "Solarize", True
                
            Case "effects_twins"
                Process "Twins", True
                
            Case "effects_vignetting"
                Process "Vignetting", True
                
        Case "effects_transform"
            Case "effects_panandzoom"
                Process "Pan and zoom", True
                
            Case "effects_perspective"
                Process "Perspective", True
                
            Case "effects_polarconversion"
                Process "Polar conversion", True
                
            Case "effects_rotate"
                Process "Rotate", True
                
            Case "effects_shear"
                Process "Shear", True
                
            Case "effects_spherize"
                Process "Spherize", True
                
        Case "effects_customfilter"
            Process "Custom filter", True
            
        Case Else
            cmdFound = False
            
    End Select
    
    PDA_ByName_MenuEffects = cmdFound
    
End Function

Private Function PDA_ByName_MenuTools(ByRef srcMenuName As String) As Boolean

    Dim cmdFound As Boolean: cmdFound = True
    
    Select Case srcMenuName
    
        Case "tools_language"
        
        Case "tools_languageeditor"
            If (Not FormLanguageEditor.Visible) Then
                FormMain.pdHotkeys.Enabled = False
                ShowPDDialog vbModal, FormLanguageEditor
                FormMain.pdHotkeys.Enabled = True
            End If
            
        Case "tools_theme"
            DialogManager.PromptUITheme
            
        Case "tools_macrocreatetop"
            Case "tools_macrofromhistory"
                ShowPDDialog vbModal, FormMacroSession
                
            Case "tools_recordmacro"
                Process "Start macro recording", , , UNDO_Nothing
                
            Case "tools_stopmacro"
                Process "Stop macro recording", True
                
        Case "tools_playmacro"
            Process "Play macro", True
            
        Case "tools_recentmacros"
        
        Case "tools_options"
            ShowPDDialog vbModal, FormOptions
            
        Case "tools_plugins"
            ShowPDDialog vbModal, FormPluginManager
            
        Case "tools_developers"
            Case "tools_themeeditor"
                ShowPDDialog vbModal, FormThemeEditor
                
            Case "tools_themepackage"
                g_Themer.BuildThemePackage
                
            Case "tools_standalonepackage"
                ShowPDDialog vbModal, FormPackage
                
        Case "effects_developertest"
        
        Case Else
            cmdFound = False
        
    End Select
    
    PDA_ByName_MenuTools = cmdFound
    
End Function

Private Function PDA_ByName_MenuView(ByRef srcMenuName As String) As Boolean

    Dim cmdFound As Boolean: cmdFound = True
    
    Select Case srcMenuName
    
        Case "view_fit"
            CanvasManager.FitOnScreen
            
        Case "view_zoomin"
            If FormMain.MainCanvas(0).IsZoomEnabled Then
                If (FormMain.MainCanvas(0).GetZoomDropDownIndex > 0) Then FormMain.MainCanvas(0).SetZoomDropDownIndex g_Zoom.GetNearestZoomInIndex(FormMain.MainCanvas(0).GetZoomDropDownIndex)
            End If
            
        Case "view_zoomout"
            If FormMain.MainCanvas(0).IsZoomEnabled Then
                If (FormMain.MainCanvas(0).GetZoomDropDownIndex <> g_Zoom.GetZoomCount) Then FormMain.MainCanvas(0).SetZoomDropDownIndex g_Zoom.GetNearestZoomOutIndex(FormMain.MainCanvas(0).GetZoomDropDownIndex)
            End If
            
        Case "view_zoomtop"
            Case "zoom_16_1"
                If FormMain.MainCanvas(0).IsZoomEnabled Then FormMain.MainCanvas(0).SetZoomDropDownIndex 2
                
            Case "zoom_8_1"
                If FormMain.MainCanvas(0).IsZoomEnabled Then FormMain.MainCanvas(0).SetZoomDropDownIndex 4
                
            Case "zoom_4_1"
                If FormMain.MainCanvas(0).IsZoomEnabled Then FormMain.MainCanvas(0).SetZoomDropDownIndex 8
                
            Case "zoom_2_1"
                If FormMain.MainCanvas(0).IsZoomEnabled Then FormMain.MainCanvas(0).SetZoomDropDownIndex 10
                
            Case "zoom_actual"
                If FormMain.MainCanvas(0).IsZoomEnabled Then FormMain.MainCanvas(0).SetZoomDropDownIndex g_Zoom.GetZoom100Index
                
            Case "zoom_1_2"
                If FormMain.MainCanvas(0).IsZoomEnabled Then FormMain.MainCanvas(0).SetZoomDropDownIndex 14
                
            Case "zoom_1_4"
                If FormMain.MainCanvas(0).IsZoomEnabled Then FormMain.MainCanvas(0).SetZoomDropDownIndex 16
                
            Case "zoom_1_8"
                If FormMain.MainCanvas(0).IsZoomEnabled Then FormMain.MainCanvas(0).SetZoomDropDownIndex 19
                
            Case "zoom_1_16"
                If FormMain.MainCanvas(0).IsZoomEnabled Then FormMain.MainCanvas(0).SetZoomDropDownIndex 21
                
        Case "view_rulers"
            Dim newRulerState As Boolean
            newRulerState = Not FormMain.MainCanvas(0).GetRulerVisibility()
            FormMain.MnuView(6).Checked = newRulerState
            FormMain.MainCanvas(0).SetRulerVisibility newRulerState
            
        Case "view_statusbar"
            Dim newStatusBarState As Boolean
            newStatusBarState = Not FormMain.MainCanvas(0).GetStatusBarVisibility()
            FormMain.MnuView(7).Checked = newStatusBarState
            FormMain.MainCanvas(0).SetStatusBarVisibility newStatusBarState
            
        Case Else
            cmdFound = False
        
    End Select
    
    PDA_ByName_MenuView = cmdFound
    
End Function

Private Function PDA_ByName_MenuWindow(ByRef srcMenuName As String) As Boolean

    Dim cmdFound As Boolean: cmdFound = True
    
    Select Case srcMenuName
    
        Case "window_toolbox"
            Case "window_displaytoolbox"
                Toolboxes.ToggleToolboxVisibility PDT_LeftToolbox
                
            Case "window_displaytoolcategories"
                toolbar_Toolbox.ToggleToolCategoryLabels
                
            Case "window_smalltoolbuttons"
                toolbar_Toolbox.UpdateButtonSize tbs_Small
                
            Case "window_normaltoolbuttons"
                toolbar_Toolbox.UpdateButtonSize tbs_Medium
                
            Case "window_largetoolbuttons"
                toolbar_Toolbox.UpdateButtonSize tbs_Large
                
        Case "window_tooloptions"
            Toolboxes.ToggleToolboxVisibility PDT_BottomToolbox
            
        Case "window_layers"
            Toolboxes.ToggleToolboxVisibility PDT_RightToolbox
            
        Case "window_imagetabstrip"
            Case "window_imagetabstrip_alwaysshow"
                Interface.ToggleImageTabstripVisibility 0
                
            Case "window_imagetabstrip_shownormal"
                Interface.ToggleImageTabstripVisibility 1
                
            Case "window_imagetabstrip_nevershow"
                Interface.ToggleImageTabstripVisibility 2
                
            Case "window_imagetabstrip_alignleft"
                Interface.ToggleImageTabstripAlignment vbAlignLeft
                
            Case "window_imagetabstrip_aligntop"
                Interface.ToggleImageTabstripAlignment vbAlignTop
                
            Case "window_imagetabstrip_alignright"
                Interface.ToggleImageTabstripAlignment vbAlignRight
                
            Case "window_imagetabstrip_alignbottom"
                Interface.ToggleImageTabstripAlignment vbAlignBottom
                
        Case "window_resetsettings"
            Toolboxes.ResetAllToolboxSettings
            
        Case "window_next"
            PDImages.MoveToNextImage True
            
        Case "window_previous"
            PDImages.MoveToNextImage False
            
        Case Else
            cmdFound = False
        
    End Select
    
    PDA_ByName_MenuWindow = cmdFound
    
End Function

Private Function PDA_ByName_MenuHelp(ByRef srcMenuName As String) As Boolean

    Dim cmdFound As Boolean: cmdFound = True
    
    Select Case srcMenuName
    
        Case "help_patreon"
            Web.OpenURL "https://www.patreon.com/photodemon/overview"
            
        Case "help_donate"
            Web.OpenURL "https://photodemon.org/donate"
            
        Case "help_checkupdates"
            
            'Initiate an asynchronous download of the standard PD update file (currently hosted @ GitHub).
            ' When the asynchronous download completes, the downloader will place the completed update file in the /Data/Updates subfolder.
            ' On exit (or subsequent program runs), PD will check for the presence of that file, then proceed accordingly.
            Message "Checking for software updates..."
            FormMain.RequestAsynchronousDownload "PROGRAM_UPDATE_CHECK_USER", "https://raw.githubusercontent.com/tannerhelland/PhotoDemon-Updates/master/summary/pdupdate.xml", , vbAsyncReadForceUpdate, UserPrefs.GetUpdatePath & "updates.xml"
            
        Case "help_reportbug"
            Web.OpenURL "https://github.com/tannerhelland/PhotoDemon/issues/"
            
        Case "help_website"
            Web.OpenURL "http://www.photodemon.org"
            
        Case "help_sourcecode"
            Web.OpenURL "https://github.com/tannerhelland/PhotoDemon"
            
        Case "help_license"
            Web.OpenURL "https://photodemon.org/license/"
            
        Case "help_about"
            ShowPDDialog vbModal, FormAbout
            
        Case Else
            cmdFound = False
        
    End Select
    
    PDA_ByName_MenuHelp = cmdFound
    
End Function

Private Function PDA_ByName_NonMenu(ByRef srcMenuName As String) As Boolean

    Dim cmdFound As Boolean: cmdFound = True
    
    Select Case srcMenuName
        
        Case "tool_hand"
            toolbar_Toolbox.SelectNewTool NAV_DRAG, True, True
        
        Case "tool_move"
            toolbar_Toolbox.SelectNewTool NAV_MOVE, True, True
        
        Case "tool_colorselect"
            toolbar_Toolbox.SelectNewTool COLOR_PICKER, True, True
        
        Case "tool_measure"
            toolbar_Toolbox.SelectNewTool ND_MEASURE, True, True
        
        Case "tool_select_rect"
            toolbar_Toolbox.SelectNewTool SELECT_RECT, True, True
        
        Case "tool_select_ellipse"
            toolbar_Toolbox.SelectNewTool SELECT_CIRC, True, True
        
        Case "tool_select_line"
            toolbar_Toolbox.SelectNewTool SELECT_LINE, True, True
        
        Case "tool_select_polygon"
            toolbar_Toolbox.SelectNewTool SELECT_POLYGON, True, True
        
        Case "tool_select_lasso"
            toolbar_Toolbox.SelectNewTool SELECT_LASSO, True, True
        
        Case "tool_select wand"
            toolbar_Toolbox.SelectNewTool SELECT_WAND, True, True
        
        Case "tool_text"
            toolbar_Toolbox.SelectNewTool VECTOR_TEXT, True, True
        
        Case "tool_typography"
            toolbar_Toolbox.SelectNewTool VECTOR_FANCYTEXT, True, True
        
        Case "tool_pencil"
            toolbar_Toolbox.SelectNewTool PAINT_BASICBRUSH, True, True
        
        Case "tool_paintbrush"
            toolbar_Toolbox.SelectNewTool PAINT_SOFTBRUSH, True, True
        
        Case "tool_erase"
            toolbar_Toolbox.SelectNewTool PAINT_ERASER, True, True
        
        Case "tool_paintbucket"
            toolbar_Toolbox.SelectNewTool PAINT_FILL, True, True
        
        Case "tool_gradient"
            toolbar_Toolbox.SelectNewTool PAINT_GRADIENT, True, True
        
        Case Else
            cmdFound = False
            
    End Select
    
    PDA_ByName_NonMenu = cmdFound

End Function

'Some of PD's menus obey special rules.  (For example, menus that add/remove entries at run-time.)  These menus have their own
' helper update functions that can be called on demand, separate from other menus in the project.
Public Sub UpdateSpecialMenu_Language(ByVal numOfLanguages As Long, ByRef availableLanguages() As PDLanguageFile)

    'Retrieve handles to the parent menu
    Dim hMenu As Long
    hMenu = GetMenu(FormMain.hWnd)
    hMenu = GetSubMenu(hMenu, 7)
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
        hMenu = GetSubMenu(hMenu, 7&)
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
'NOTE!  Validate your menu index before passing it to this function.  For performance reasons,
' no extra validation is applied to the incoming index.
Private Function GetHMenu_FromIndex(ByVal mnuIndex As Long, Optional ByVal getParentMenu As Boolean = False) As Long

    'We always start by retrieving the menu handle for the primary form
    Dim curHMenu As Long, hMenuParent As Long
    curHMenu = GetMenu(FormMain.hWnd)
    hMenuParent = curHMenu

    'Next, iterate through submenus until we arrive at the entry we want
    curHMenu = GetSubMenu(curHMenu, m_Menus(mnuIndex).me_TopMenu)
    If (m_Menus(mnuIndex).me_SubMenu <> MENU_NONE) Then
        
        hMenuParent = curHMenu
        curHMenu = GetSubMenu(curHMenu, m_Menus(mnuIndex).me_SubMenu)
            
        If (m_Menus(mnuIndex).me_SubSubMenu <> MENU_NONE) Then
            hMenuParent = curHMenu
            curHMenu = GetSubMenu(curHMenu, m_Menus(mnuIndex).me_SubSubMenu)
        End If
        
    End If
    
    If getParentMenu Then GetHMenu_FromIndex = hMenuParent Else GetHMenu_FromIndex = curHMenu
    
End Function

'When working with APIs, you typically pass the hMenu of the parent menu, and then a simple itemIndex to address the
' child item in a given menu.  Use this function to simplify the handling of hMenu indices.
Private Function GetHMenuIndex(ByVal mnuIndex As Long) As Long

    If (m_Menus(mnuIndex).me_SubMenu = MENU_NONE) Then
        GetHMenuIndex = m_Menus(mnuIndex).me_TopMenu
    Else
        If (m_Menus(mnuIndex).me_SubSubMenu = MENU_NONE) Then
            GetHMenuIndex = m_Menus(mnuIndex).me_SubMenu
        Else
            GetHMenuIndex = m_Menus(mnuIndex).me_SubSubMenu
        End If
    End If

End Function

'Update a given menu's text caption.  By design, this function does *not* trigger a DrawMenuBar call.
Private Sub UpdateMenuText_ByIndex(ByVal mnuIndex As Long)

    'Get an hMenu for the specified index
    Dim hMenu As Long
    hMenu = GetHMenu_FromIndex(mnuIndex, True)
    
    If (hMenu <> 0) Then
        
        'Populate a MenuItemInfo struct
        Dim tmpMII As Win32_MenuItemInfoW
        tmpMII.cbSize = LenB(tmpMII)
        tmpMII.fMask = MIIM_STRING
        tmpMII.dwTypeData = StrPtr(m_Menus(mnuIndex).me_TextFinal)
        
        SetMenuItemInfoW hMenu, GetHMenuIndex(mnuIndex), 1&, tmpMII
        
    Else
        InternalMenuWarning "UpdateMenuText_ByIndex", "null hMenu (" & mnuIndex & ")"
    End If

End Sub

Private Sub InternalMenuWarning(ByRef funcName As String, ByRef errMsg As String)
    PDDebug.LogAction "WARNING!  Menus." & funcName & " reported: " & errMsg
End Sub
