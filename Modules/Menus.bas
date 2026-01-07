Attribute VB_Name = "Menus"
'***************************************************************************
'PhotoDemon Menu Manager
'Copyright 2017-2026 by Tanner Helland
'Created: 11/January/17
'Last updated: 08/January/25
'Last update: move the Tools > Test menu to the Developer submenu
'
'PhotoDemon has an extensive menu system.  Managing all those menus is cumbersome.
' This module handles the worst parts of run-time maintenance.
'
'Because PD's menus provide an organized collection of program features, this module also handles
' some module-adjacent tasks, like the ProcessDefaultAction-prefixed functions.  You can pass these
' functions a menu name (or caption), and they will automatically initiate the corresponding
' program action.  The long-term goal is to use these links to handle run-time hotkey mapping.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Public Type PD_MenuEntry
    me_TopMenu As Long                    'Top-level index of this menu
    me_SubMenu As Long                    'Sub-menu index of this menu (if any)
    me_SubSubMenu As Long                 'Sub-sub-menu index of this menu (if any)
    me_HotKeyID As Long                   'Hotkey, if any, associated with this menu.  (-1 = no hotkey)
    me_HotKeyTextTranslated As String     'Hotkey text, with translations (if any) always applied.  Hotkeys module generates this.
    me_Name As String                     'Name of this menu (must be unique)
    me_ResImage As String                 'Name of this menu's image, as stored in PD's central resource file
    me_TextEn As String                   'Text of this menu, in English
    me_TextTranslated As String           'Text of this menu, as translated by the current language
    me_TextWithMnemonics As String        'Text of this menu, translated, with a mnemonic char (&) added
    me_TextFinal As String                'Final on-screen appearance of the text, with translations, mnemonics, and accelerator (if any)
    me_TextSearchable As String           'Localized string for search results.  Uses "TopMenu > ChildMenu > MyMenuName" format.  No mnemonics or hotkey.
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

Private Enum Win32_EnableMenuItem
    MFE_BYPOSITION = &H400&      'Indicates that uIDEnableItem gives the zero-based relative position of the menu item.
    MFE_DISABLED = &H2&          'Indicates that the menu item is disabled, but not grayed, so it cannot be selected.
    MFE_ENABLED = &H0&           'Indicates that the menu item is enabled and restored from a grayed state so that it can be selected.
    MFE_GRAYED = &H1&            'Indicates that the menu item is disabled and grayed so that it cannot be selected.
End Enum

#If False Then
    Private Const MFE_BYPOSITION = &H400&, MFE_DISABLED = &H2&, MFE_ENABLED = &H0&, MFE_GRAYED = &H1&
#End If

Private Type Win32_MENUBARINFO
    cbSize As Long
    rcBar As RectL
    hMenu As Long
    hwndMenu As Long
    fFlags As Long
End Type

'When modifying menus, special ID values can be used to restrict operations
Private Const IGNORE_MENU_ID As Long = -10
Private Const ALL_MENU_SUBITEMS As Long = -9
Private Const MENU_NONE As Long = -1

'A special ID is used to flag menus with no associated hotkey.  (Menus need to track this so they can
' display associated hotkey text, if any.)
Private Const NO_MENU_HOTKEY As Long = -1

'A number of menu features require us to interact directly with the API
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal uIDEnabledItem As Long, ByVal uEnable As Win32_EnableMenuItem) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetMenuBarInfo Lib "user32" (ByVal hWnd As Long, ByVal idObject As Long, ByVal idItem As Long, ByVal pMBI As Long) As Long
Private Declare Function GetMenuItemInfoW Lib "user32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, ByRef srcMenuItemInfo As Win32_MenuItemInfoW) As Long
Private Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal uId As Long, ByVal uFlags As Win32_MenuStateFlags) As Win32_MenuStateFlags
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function IsCharAlphaW Lib "user32" (ByVal wChar As Integer) As Long
Private Declare Function SetMenuItemInfoW Lib "user32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, ByRef srcMenuItemInfo As Win32_MenuItemInfoW) As Long
Private Declare Function VkKeyScanW Lib "user32" (ByVal wChar As Integer) As Integer

'Like the built-in VB6 menu editor, PD uses "-" as the caption for menu separator bars
Private Const MENU_SEPARATOR As String = "-"

'Primary menu collection
Private m_Menus() As PD_MenuEntry
Private m_NumOfMenus As Long

'In languages that do not map cleanly to keyboard keys (i.e. languages with IME-driven input),
' PD automatically uses the original English text for mnemonics.  These mnemonics are displayed in parentheses
' to the right of the localized menu text.  As of v2025.4, the user can switch this behavior between "auto/on/off".
Private m_WideMenuMnemonics As PD_BOOL

'Early in the PD load process, we initialize the default set of menus.  In the future, it may be nice to let
' users customize this to match their favorite software (e.g. PhotoShop), but that's a ways off as I've yet to
' build a menu control capable of that level of customization support.
Public Sub InitializeMenus()
        
    'First, load menu mnemonic behavior
    Menus.SetMnemonicsBehavior UserPrefs.GetPref_Long("Interface", "display-mnemonics", PD_BOOL_AUTO)
    
    'File Menu
    AddMenuItem "File", "file_top", 0
    AddMenuItem "New...", "file_new", 0, 0, , "file_new"
    AddMenuItem "Open...", "file_open", 0, 1, , "file_open"
    AddMenuItem "Open recent", "file_openrecent", 0, 2
    AddMenuItem "Import", "file_import", 0, 3
        AddMenuItem "From clipboard", "file_import_paste", 0, 3, 0, "file_importclipboard"
        AddMenuItem "-", "-", 0, 3, 1
        AddMenuItem "From scanner or camera...", "file_import_scanner", 0, 3, 2, "file_importscanner"
        AddMenuItem "Select which scanner or camera to use...", "file_import_selectscanner", 0, 3, 3
        AddMenuItem "-", "-", 0, 3, 4
        AddMenuItem "Online image...", "file_import_web", 0, 3, 5, "file_importweb"
        AddMenuItem "-", "-", 0, 3, 6
        AddMenuItem "Screenshot...", "file_import_screenshot", 0, 3, 7, "file_importscreen"
    AddMenuItem "-", "-", 0, 4
    AddMenuItem "Close", "file_close", 0, 5, , "file_close"
    AddMenuItem "Close all", "file_closeall", 0, 6
    AddMenuItem "-", "-", 0, 7
    AddMenuItem "Save", "file_save", 0, 8, , "file_save"
    AddMenuItem "Save copy (lossless)", "file_savecopy", 0, 9, , "file_savedup"
    AddMenuItem "Save as...", "file_saveas", 0, 10, , "file_saveas"
    AddMenuItem "Revert", "file_revert", 0, 11
    AddMenuItem "Export", "file_export", 0, 12
        AddMenuItem "Image to file...", "file_export_image", 0, 12, 0
        AddMenuItem "Layers to files...", "file_export_layers", 0, 12, 1
        AddMenuItem "-", "-", 0, 12, 2
        AddMenuItem "Animation...", "file_export_animation", 0, 12, 3
        AddMenuItem "-", "-", 0, 12, 4
        AddMenuItem "Color lookup...", "file_export_colorlookup", 0, 12, 5
        AddMenuItem "Color profile...", "file_export_colorprofile", 0, 12, 6
        AddMenuItem "Palette...", "file_export_palette", 0, 12, 7
    AddMenuItem "-", "-", 0, 13
    AddMenuItem "Batch operations", "file_batch", 0, 14
        AddMenuItem "Process...", "file_batch_process", 0, 14, 0, "file_batch"
        AddMenuItem "Repair...", "file_batch_repair", 0, 14, 1, "file_repair"
    AddMenuItem "-", "-", 0, 15
    AddMenuItem "Print...", "file_print", 0, 16, , "file_print"
    AddMenuItem "-", "-", 0, 17
    AddMenuItem "Exit", "file_quit", 0, 18
    
    'Edit menu
    AddMenuItem "Edit", "edit_top", 1
    AddMenuItem "Undo", "edit_undo", 1, 0, , "edit_undo"
    AddMenuItem "Redo", "edit_redo", 1, 1, , "edit_redo"
    AddMenuItem "Undo history...", "edit_history", 1, 2, , "edit_history"
    AddMenuItem "-", "-", 1, 3
    AddMenuItem "Repeat", "edit_repeat", 1, 4, , "edit_repeat"
    AddMenuItem "Fade...", "edit_fade", 1, 5
    AddMenuItem "-", "-", 1, 6
    AddMenuItem "Cut", "edit_cutlayer", 1, 7, , "edit_cut"
    AddMenuItem "Cut merged", "edit_cutmerged", 1, 8
    AddMenuItem "Copy", "edit_copylayer", 1, 9, , "edit_copy"
    AddMenuItem "Copy merged", "edit_copymerged", 1, 10
    AddMenuItem "Paste", "edit_pasteaslayer", 1, 11, , "edit_paste"
    AddMenuItem "Paste to new image", "edit_pasteasimage", 1, 12
    AddMenuItem "Special", "edit_specialtop", 1, 13
    AddMenuItem "Cut special...", "edit_specialcut", 1, 13, 0, "edit_cut"
    AddMenuItem "Copy special...", "edit_specialcopy", 1, 13, 1, "edit_copy"
    AddMenuItem "Paste special...", "edit_specialpaste", 1, 13, 2, "edit_paste"
    AddMenuItem "-", "-", 1, 13, 3
    AddMenuItem "Empty clipboard", "edit_emptyclipboard", 1, 13, 4
    AddMenuItem "-", "-", 1, 14
    AddMenuItem "Clear", "edit_clear", 1, 15, , "paint_erase"
    AddMenuItem "Content-aware fill...", "edit_contentawarefill", 1, 16, , "edit_contentawarefill"
    AddMenuItem "Fill...", "edit_fill", 1, 17, , "paint_fill"
    AddMenuItem "Stroke...", "edit_stroke", 1, 18, , "paint_softbrush"
    
    'Image Menu
    AddMenuItem "Image", "image_top", 2
    AddMenuItem "Duplicate", "image_duplicate", 2, 0, , "edit_copy"
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
        AddMenuItem "Rotate 90 clockwise", "image_rotate90", 2, 12, 2, "generic_rotateright"
        AddMenuItem "Rotate 90 counter-clockwise", "image_rotate270", 2, 12, 3, "generic_rotateleft"
        AddMenuItem "Rotate 180", "image_rotate180", 2, 12, 4
        AddMenuItem "Rotate arbitrary...", "image_rotatearbitrary", 2, 12, 5
    AddMenuItem "Flip horizontal", "image_fliphorizontal", 2, 13, , "image_fliphorizontal"
    AddMenuItem "Flip vertical", "image_flipvertical", 2, 14, , "image_flipvertical"
    AddMenuItem "-", "-", 2, 15
    AddMenuItem "Merge visible layers", "image_mergevisible", 2, 16, , "generic_visible"
    AddMenuItem "Flatten image...", "image_flatten", 2, 17, , "layer_flatten"
    AddMenuItem "-", "-", 2, 18
    AddMenuItem "Animation...", "image_animation", 2, 19, , "animation"
    AddMenuItem "Compare", "image_compare", 2, 20
        AddMenuItem "Create color lookup...", "image_createlut", 2, 20, 0
        AddMenuItem "Similarity...", "image_similarity", 2, 20, 1
    AddMenuItem "Metadata", "image_metadata", 2, 21
        AddMenuItem "Edit metadata...", "image_editmetadata", 2, 21, 0, "image_metadata"
        AddMenuItem "Remove all metadata", "image_removemetadata", 2, 21, 1
        AddMenuItem "-", "-", 2, 21, 2
        AddMenuItem "Count unique colors", "image_countcolors", 2, 21, 3
        AddMenuItem "Map photo location...", "image_maplocation", 2, 21, 4, "image_maplocation"
    AddMenuItem "Show in file manager...", "image_showinexplorer", 2, 22
    
    'Layer menu
    AddMenuItem "Layer", "layer_top", 3
    AddMenuItem "Add", "layer_add", 3, 0
        AddMenuItem "Basic layer...", "layer_addbasic", 3, 0, 0
        AddMenuItem "Blank layer", "layer_addblank", 3, 0, 1
        AddMenuItem "Duplicate of current layer", "layer_duplicate", 3, 0, 2, "edit_copy"
        AddMenuItem "-", "-", 3, 0, 3
        AddMenuItem "From clipboard", "layer_addfromclipboard", 3, 0, 4, "edit_paste"
        AddMenuItem "From file...", "layer_addfromfile", 3, 0, 5, "file_open"
        AddMenuItem "From visible layers", "layer_addfromvisiblelayers", 3, 0, 6
        AddMenuItem "-", "-", 3, 0, 7
        AddMenuItem "Layer via copy", "layer_addviacopy", 3, 0, 8, "edit_copy"
        AddMenuItem "Layer via cut", "layer_addviacut", 3, 0, 9, "edit_cut"
    AddMenuItem "Delete", "layer_delete", 3, 1
        AddMenuItem "Current layer", "layer_deletecurrent", 3, 1, 0, "generic_trash"
        AddMenuItem "Hidden layers", "layer_deletehidden", 3, 1, 1, "generic_invisible"
    AddMenuItem "Replace", "layer_replace", 3, 2
        AddMenuItem "From clipboard", "layer_replacefromclipboard", 3, 2, 0, "edit_paste"
        AddMenuItem "From file...", "layer_replacefromfile", 3, 2, 1, "file_open"
        AddMenuItem "From visible layers", "layer_replacefromvisiblelayers", 3, 2, 2
    AddMenuItem "-", "-", 3, 3
    AddMenuItem "Merge up", "layer_mergeup", 3, 4, , "layer_mergeup"
    AddMenuItem "Merge down", "layer_mergedown", 3, 5, , "layer_mergedown"
    AddMenuItem "Order", "layer_order", 3, 6
        AddMenuItem "Go to top layer", "layer_gotop", 3, 6, 0
        AddMenuItem "Go to layer above", "layer_goup", 3, 6, 1
        AddMenuItem "Go to layer below", "layer_godown", 3, 6, 2
        AddMenuItem "Go to bottom layer", "layer_gobottom", 3, 6, 3
        AddMenuItem "-", "-", 3, 6, 4
        AddMenuItem "Move layer to top", "layer_movetop", 3, 6, 5
        AddMenuItem "Move layer up", "layer_moveup", 3, 6, 6, "layer_up"
        AddMenuItem "Move layer down", "layer_movedown", 3, 6, 7, "layer_down"
        AddMenuItem "Move layer to bottom", "layer_movebottom", 3, 6, 8
        AddMenuItem "-", "-", 3, 6, 9
        AddMenuItem "Reverse", "layer_reverse", 3, 6, 10
    AddMenuItem "Visibility", "layer_visibility", 3, 7
        AddMenuItem "Show this layer", "layer_show", 3, 7, 0
        AddMenuItem "-", "-", 3, 7, 1
        AddMenuItem "Show only this layer", "layer_showonly", 3, 7, 2
        AddMenuItem "Hide only this layer", "layer_hideonly", 3, 7, 3
        AddMenuItem "-", "-", 3, 7, 4
        AddMenuItem "Show all layers", "layer_showall", 3, 7, 5
        AddMenuItem "Hide all layers", "layer_hideall", 3, 7, 6
    AddMenuItem "-", "-", 3, 8
    AddMenuItem "Crop", "layer_crop", 3, 9
        AddMenuItem "Crop to selection", "layer_cropselection", 3, 9, 0, "image_crop"
        AddMenuItem "-", "-", 3, 9, 1
        AddMenuItem "Fit to canvas", "layer_pad", 3, 9, 2
        AddMenuItem "Trim empty borders", "layer_trim", 3, 9, 3
    AddMenuItem "Orientation", "layer_orientation", 3, 10
        AddMenuItem "Straighten...", "layer_straighten", 3, 10, 0
        AddMenuItem "-", "-", 3, 10, 1
        AddMenuItem "Rotate 90 clockwise", "layer_rotate90", 3, 10, 2, "generic_rotateright"
        AddMenuItem "Rotate 90 counter-clockwise", "layer_rotate270", 3, 10, 3, "generic_rotateleft"
        AddMenuItem "Rotate 180", "layer_rotate180", 3, 10, 4
        AddMenuItem "Rotate arbitrary...", "layer_rotatearbitrary", 3, 10, 5
        AddMenuItem "-", "-", 3, 10, 6
        AddMenuItem "Flip horizontal", "layer_fliphorizontal", 3, 10, 7, "image_fliphorizontal"
        AddMenuItem "Flip vertical", "layer_flipvertical", 3, 10, 8, "image_flipvertical"
    AddMenuItem "Size", "layer_resize", 3, 11
        AddMenuItem "Reset to actual size", "layer_resetsize", 3, 11, 0, "generic_reset"
        AddMenuItem "-", "-", 3, 11, 1
        AddMenuItem "Resize...", "layer_resize", 3, 11, 2, "image_resize"
        AddMenuItem "Content-aware resize...", "layer_contentawareresize", 3, 11, 3
        AddMenuItem "-", "-", 3, 11, 4
        AddMenuItem "Fit to image", "layer_fittoimage", 3, 11, 5
    AddMenuItem "-", "-", 3, 12
    AddMenuItem "Transparency", "layer_transparency", 3, 13
        AddMenuItem "From color (chroma key)...", "layer_colortoalpha", 3, 13, 0
        AddMenuItem "From luminance...", "layer_luminancetoalpha", 3, 13, 1
        AddMenuItem "-", "-", 3, 13, 2
        AddMenuItem "Remove transparency...", "layer_removealpha", 3, 13, 3, "generic_trash"
        AddMenuItem "Threshold...", "layer_thresholdalpha", 3, 13, 4
    AddMenuItem "-", "-", 3, 14
    AddMenuItem "Rasterize", "layer_rasterize", 3, 15
        AddMenuItem "Current layer", "layer_rasterizecurrent", 3, 15, 0
        AddMenuItem "All layers", "layer_rasterizeall", 3, 15, 1
    AddMenuItem "Split", "layer_split", 3, 16
        AddMenuItem "Current layer into standalone image", "layer_splitlayertoimage", 3, 16, 0
        AddMenuItem "All layers into standalone images", "layer_splitalllayerstoimages", 3, 16, 1
        AddMenuItem "-", "-", 3, 16, 2
        AddMenuItem "Other open images into this image (as layers)...", "layer_splitimagestolayers", 3, 16, 3
    
    'Select Menu
    AddMenuItem "Select", "select_top", 4
    AddMenuItem "All", "select_all", 4, 0, , "select_all"
    AddMenuItem "None", "select_none", 4, 1, , "select_none"
    AddMenuItem "Invert", "select_invert", 4, 2
    AddMenuItem "-", "-", 4, 3
    AddMenuItem "Grow...", "select_grow", 4, 4
    AddMenuItem "Shrink...", "select_shrink", 4, 5
    AddMenuItem "Border...", "select_border", 4, 6
    AddMenuItem "Feather...", "select_feather", 4, 7
    AddMenuItem "Sharpen...", "select_sharpen", 4, 8
    AddMenuItem "-", "-", 4, 9
    AddMenuItem "Erase selected area", "select_erasearea", 4, 10, , "paint_erase"
    AddMenuItem "Fill selected area...", "select_fill", 4, 11, , "paint_fill"
    AddMenuItem "Heal selected area...", "select_heal", 4, 12, , "edit_contentawarefill"
    AddMenuItem "Stroke selection outline...", "select_stroke", 4, 13, , "paint_softbrush"
    AddMenuItem "-", "-", 4, 14
    AddMenuItem "Load selection...", "select_load", 4, 15, , "file_open"
    AddMenuItem "Save current selection...", "select_save", 4, 16, , "file_save"
    AddMenuItem "Export", "select_export", 4, 17
        AddMenuItem "Selected area as image...", "select_exportarea", 4, 17, 0
        AddMenuItem "Selection mask as image...", "select_exportmask", 4, 17, 1
        
    'Adjustments Menu
    AddMenuItem "Adjustments", "adj_top", 5
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
        AddMenuItem "Color lookup...", "adj_colorlookup", 5, 13, 9
        AddMenuItem "Colorize...", "adj_colorize", 5, 13, 10
        AddMenuItem "Photo filter...", "adj_photofilters", 5, 13, 11
        AddMenuItem "Replace color...", "adj_replacecolor", 5, 13, 12
        AddMenuItem "Sepia...", "adj_sepia", 5, 13, 13
        AddMenuItem "Split toning...", "adj_splittone", 5, 13, 14
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
        AddMenuItem "Dehaze...", "adj_dehaze", 5, 16, 2
        AddMenuItem "Exposure...", "adj_exposure", 5, 16, 3
        AddMenuItem "Gamma...", "adj_gamma", 5, 16, 4
        AddMenuItem "HDR...", "adj_hdr", 5, 16, 5
        AddMenuItem "Levels...", "adj_levels", 5, 16, 6
        AddMenuItem "Shadows and highlights...", "adj_sandh", 5, 16, 7
    AddMenuItem "Map", "adj_map", 5, 17
        AddMenuItem "Gradient map...", "adj_gradientmap", 5, 17, 0
        AddMenuItem "Palette map...", "adj_palettemap", 5, 17, 1
    AddMenuItem "Monochrome", "adj_monochrome", 5, 18
        AddMenuItem "Color to monochrome...", "adj_colortomonochrome", 5, 18, 0
        AddMenuItem "Monochrome to gray...", "adj_monochrometogray", 5, 18, 1
        
    'Effects (Filters) Menu
    AddMenuItem "Effects", "effects_top", 6
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
    AddMenuItem "Distort", "effects_distort", 6, 2
        AddMenuItem "Correct existing distortion...", "effects_fixlensdistort", 6, 2, 0
        AddMenuItem "-", "-", 6, 2, 1
        AddMenuItem "Donut...", "effects_donut", 6, 2, 2
        AddMenuItem "Droste...", "effects_droste", 6, 2, 3
        AddMenuItem "Lens...", "effects_lens", 6, 2, 4
        AddMenuItem "Pinch and whirl...", "effects_pinchandwhirl", 6, 2, 5
        AddMenuItem "Poke...", "effects_poke", 6, 2, 6
        AddMenuItem "Ripple...", "effects_ripple", 6, 2, 7
        AddMenuItem "Squish...", "effects_squish", 6, 2, 8
        AddMenuItem "Swirl...", "effects_swirl", 6, 2, 9
        AddMenuItem "Waves...", "effects_waves", 6, 2, 10
        AddMenuItem "-", "-", 6, 2, 11
        AddMenuItem "Miscellaneous...", "effects_miscdistort", 6, 2, 12
    AddMenuItem "Edge", "effects_edges", 6, 3
        AddMenuItem "Emboss...", "effects_emboss", 6, 3, 0
        AddMenuItem "Enhance edges...", "effects_enhanceedges", 6, 3, 1
        AddMenuItem "Find edges...", "effects_findedges", 6, 3, 2
        AddMenuItem "Gradient flow...", "effects_gradientflow", 6, 3, 3
        AddMenuItem "Range filter...", "effects_rangefilter", 6, 3, 4
        AddMenuItem "Trace contour...", "effects_tracecontour", 6, 3, 5
    AddMenuItem "Light and shadow", "effects_lightandshadow", 6, 4
        AddMenuItem "Black light...", "effects_blacklight", 6, 4, 0
        AddMenuItem "Bump map...", "effects_bumpmap", 6, 4, 1
        AddMenuItem "Cross-screen...", "effects_crossscreen", 6, 4, 2
        AddMenuItem "Rainbow...", "effects_rainbow", 6, 4, 3
        AddMenuItem "Sunshine...", "effects_sunshine", 6, 4, 4
        AddMenuItem "-", "-", 6, 4, 5
        AddMenuItem "Dilate...", "effects_dilate", 6, 4, 6
        AddMenuItem "Erode...", "effects_erode", 6, 4, 7
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
        AddMenuItem "Dust and scratches...", "effects_dustandscratches", 6, 6, 4
        AddMenuItem "Harmonic mean...", "effects_harmonicmean", 6, 6, 5
        AddMenuItem "Mean shift...", "effects_meanshift", 6, 6, 6
        AddMenuItem "Median...", "effects_median", 6, 6, 7
        AddMenuItem "Symmetric nearest-neighbor...", "effects_snn", 6, 6, 8
    AddMenuItem "Pixelate", "effects_pixelate", 6, 7
        AddMenuItem "Color halftone...", "effects_colorhalftone", 6, 7, 0
        AddMenuItem "Crystallize...", "effects_crystallize", 6, 7, 1
        AddMenuItem "Fragment...", "effects_fragment", 6, 7, 2
        AddMenuItem "Mezzotint...", "effects_mezzotint", 6, 7, 3
        AddMenuItem "Mosaic...", "effects_mosaic", 6, 7, 4
        AddMenuItem "Pointillize...", "effects_pointillize", 6, 7, 5
    AddMenuItem "Render", "effects_render", 6, 8
        AddMenuItem "Clouds...", "effects_clouds", 6, 8, 0
        AddMenuItem "Fibers...", "effects_fibers", 6, 8, 1
        AddMenuItem "Truchet...", "effects_truchet", 6, 8, 2
    AddMenuItem "Sharpen", "effects_sharpentop", 6, 9
        AddMenuItem "Sharpen...", "effects_sharpen", 6, 9, 0
        AddMenuItem "Unsharp mask...", "effects_unsharp", 6, 9, 1
    AddMenuItem "Stylize", "effects_stylize", 6, 10
        AddMenuItem "Antique...", "effects_antique", 6, 10, 0
        AddMenuItem "Diffuse...", "effects_diffuse", 6, 10, 1
        AddMenuItem "Kuwahara...", "effects_kuwahara", 6, 10, 2
        AddMenuItem "Outline...", "effects_outline", 6, 10, 3
        AddMenuItem "Palette...", "effects_palette", 6, 10, 4
        AddMenuItem "Portrait glow...", "effects_portraitglow", 6, 10, 5
        AddMenuItem "Solarize...", "effects_solarize", 6, 10, 6
        AddMenuItem "Twins...", "effects_twins", 6, 10, 7
        AddMenuItem "Vignetting...", "effects_vignetting", 6, 10, 8
    AddMenuItem "Transform", "effects_transform", 6, 11
        AddMenuItem "Offset and zoom...", "effects_panandzoom", 6, 11, 0
        AddMenuItem "Perspective...", "effects_perspective", 6, 11, 1
        AddMenuItem "Polar conversion...", "effects_polarconversion", 6, 11, 2
        AddMenuItem "Rotate...", "effects_rotate", 6, 11, 3
        AddMenuItem "Shear...", "effects_shear", 6, 11, 4
        AddMenuItem "Spherize...", "effects_spherize", 6, 11, 5
    AddMenuItem "-", "-", 6, 12
    AddMenuItem "Animation...", "effects_animation", 6, 13
        AddMenuItem "Background...", "effects_animation_background", 6, 13, 0
        AddMenuItem "Foreground...", "effects_animation_foreground", 6, 13, 1
        AddMenuItem "Playback speed...", "effects_animation_speed", 6, 13, 2
    AddMenuItem "Custom filter...", "effects_customfilter", 6, 14
    AddMenuItem "Photoshop (8bf) plugin...", "effects_8bf", 6, 15
    
    'Tools Menu
    AddMenuItem "Tools", "tools_top", 7
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
    AddMenuItem "Animated screen capture...", "tools_screenrecord", 7, 9, , "file_importscreen"
    AddMenuItem "-", "-", 7, 10
    AddMenuItem "Keyboard shortcuts...", "tools_hotkeys", 7, 11, , "keyboard"
    AddMenuItem "Options...", "tools_options", 7, 12, , "pref_advanced"
    
    Dim debugMenuVisibility As Boolean
    debugMenuVisibility = ((PD_BUILD_QUALITY <> PD_PRODUCTION) And (PD_BUILD_QUALITY <> PD_BETA)) Or (Not OS.IsProgramCompiled)
    If debugMenuVisibility Then
        AddMenuItem "-", "-", 7, 13
        AddMenuItem "Developers", "tools_developers", 7, 14
            AddMenuItem "View debug log for this session...", "tools_viewdebuglog", 7, 14, 0, , False
            AddMenuItem "-", "-", 7, 14, 1, , False
            AddMenuItem "Theme editor...", "tools_themeeditor", 7, 14, 2, , False
            AddMenuItem "Build theme package...", "tools_themepackage", 7, 14, 3, , False
            AddMenuItem "-", "-", 7, 14, 4
            AddMenuItem "Build standalone package...", "tools_standalonepackage", 7, 14, 5, , False
            AddMenuItem "Test", "effects_developertest", 7, 14, 6, , False
    End If
    
    'View Menu
    AddMenuItem "View", "view_top", 8
    AddMenuItem "Fit image on screen", "view_fit", 8, 0, , "zoom_fit"
    AddMenuItem "Center image in viewport", "view_center_on_screen", 8, 1, , "zoom_center"
    AddMenuItem "-", "-", 8, 2
    AddMenuItem "Zoom in", "view_zoomin", 8, 3, , "zoom_in"
    AddMenuItem "Zoom out", "view_zoomout", 8, 4, , "zoom_out"
    AddMenuItem "Zoom to value", "view_zoomtop", 8, 5
        AddMenuItem "16:1 (1600%)", "zoom_16_1", 8, 5, 0
        AddMenuItem "8:1 (800%)", "zoom_8_1", 8, 5, 1
        AddMenuItem "4:1 (400%)", "zoom_4_1", 8, 5, 2
        AddMenuItem "2:1 (200%)", "zoom_2_1", 8, 5, 3
        AddMenuItem "1:1 (actual size, 100%)", "zoom_actual", 8, 5, 4, "zoom_actual"
        AddMenuItem "1:2 (50%)", "zoom_1_2", 8, 5, 5
        AddMenuItem "1:4 (25%)", "zoom_1_4", 8, 5, 6
        AddMenuItem "1:8 (12.5%)", "zoom_1_8", 8, 5, 7
        AddMenuItem "1:16 (6.25%)", "zoom_1_16", 8, 5, 8
    AddMenuItem "-", "-", 8, 6
    AddMenuItem "Show rulers", "view_rulers", 8, 7
    AddMenuItem "Show status bar", "view_statusbar", 8, 8
    AddMenuItem "Show extras", "show_extrastop", 8, 9, allowInSearches:=False
    AddMenuItem "Layer edges", "show_layeredges", 8, 9, 0
    AddMenuItem "Smart guides", "show_smartguides", 8, 9, 1
    AddMenuItem "-", "-", 8, 10
    AddMenuItem "Snap", "snap_global", 8, 11
    AddMenuItem "Snap to", "snap_top", 8, 12, allowInSearches:=False
    AddMenuItem "Canvas edges", "snap_canvasedge", 8, 12, 0
    AddMenuItem "Centerlines", "snap_centerline", 8, 12, 1
    AddMenuItem "Layers", "snap_layer", 8, 12, 2
    AddMenuItem "-", "-", 8, 12, 3
    AddMenuItem "Angle 90", "snap_angle_90", 8, 12, 4
    AddMenuItem "Angle 45", "snap_angle_45", 8, 12, 5
    AddMenuItem "Angle 30", "snap_angle_30", 8, 12, 6
    
    'Window Menu
    AddMenuItem "Window", "window_top", 9
    AddMenuItem "Toolbox", "window_toolbox", 9, 0
        AddMenuItem "Display toolbox", "window_displaytoolbox", 9, 0, 0
        AddMenuItem "-", "-", 9, 0, 1
        AddMenuItem "Display tool category titles", "window_displaytoolcategories", 9, 0, 2
        AddMenuItem "-", "-", 9, 0, 3
        AddMenuItem "Small buttons", "window_smalltoolbuttons", 9, 0, 4
        AddMenuItem "Medium buttons", "window_mediumtoolbuttons", 9, 0, 5
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
    AddMenuItem "Help", "help_top", 10
    AddMenuItem "Support us on Patreon...", "help_patreon", 10, 0, , "help_heart"
    AddMenuItem "Support us with a one-time donation...", "help_donate", 10, 1, , "help_heart"
    AddMenuItem "-", "-", 10, 2
    AddMenuItem "Ask a question...", "help_forum", 10, 3, , "generic_question"
    AddMenuItem "Check for updates...", "help_checkupdates", 10, 4, , "help_update"
    AddMenuItem "Submit bug report or feedback...", "help_reportbug", 10, 5, , "help_reportbug"
    AddMenuItem "-", "-", 10, 6
    AddMenuItem "PhotoDemon forum...", "help_website", 10, 7, , "help_forum"
    AddMenuItem "PhotoDemon license and terms of use...", "help_forum", 10, 8, , "help_license"
    AddMenuItem "PhotoDemon source code...", "help_sourcecode", 10, 9, , "help_github"
    AddMenuItem "PhotoDemon website...", "help_website", 10, 10, , "help_website"
    AddMenuItem "-", "-", 10, 11
    AddMenuItem "Third-party libraries...", "help_3rdpartylibs", 10, 12, , "tools_plugin"
    AddMenuItem "-", "-", 10, 13
    AddMenuItem "About...", "help_about", 10, 14, , "help_about"
    
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
    If (m_NumOfMenus = 0) Then
        Const INITIAL_MENU_COLLECTION_SIZE As Long = 128
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
        .me_HotKeyID = NO_MENU_HOTKEY
    End With
    
    m_NumOfMenus = m_NumOfMenus + 1

End Sub

'After adding all menu items to the central menu table, call this function to iterate through
' the final list and auto-populate a few helpful menu properties (e.g. the "has children" bool)
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
    Dim numOfMenus As Long, menuIndices() As Long, changedMenus As pdStack
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
                
            End With
            
            'Now comes some messy refresh business.  After changing one caption in a menu,
            ' we need to re-calculate mnemonics for sibling menus.  We have a function that
            ' takes care of that for us:
            DetermineMnemonics_SingleMenu m_Menus(menuIndices(i)).me_Name, changedMenus
            
            'Because multiple menus may have received updated mnemonics as a result of this
            ' caption change, we may need to update multiple menus (not just the one we were
            ' originally passed).
            If (Not changedMenus Is Nothing) Then
                If (changedMenus.GetNumOfInts > 0) Then
                    
                    Dim j As Long, targetMenu As Long
                    For j = 0 To changedMenus.GetNumOfInts() - 1
                        
                        targetMenu = changedMenus.GetInt(j)
                        
                        With m_Menus(targetMenu)
                            
                            'Combine caption, mnemonics, and hotkeys (if any) into a single, final, display-ready string
                            If (.me_HotKeyID >= 0) Then
                                .me_TextFinal = .me_TextWithMnemonics & vbTab & .me_HotKeyTextTranslated
                            Else
                                .me_TextFinal = .me_TextWithMnemonics
                            End If
                            
                            'Relay our changes to the underlying API menu struct
                            UpdateMenuText_ByIndex targetMenu
                            
                        End With
                        
                    Next j
                    
                End If
            End If
            
        Next i
        
        DrawMenuBar FormMain.hWnd
    
    Else
        InternalMenuWarning "RequestCaptionChange_ByName", "no matching menus found"
    End If

End Sub

'After the active language changes, you must call this menu to translate all menu captions.
Public Sub UpdateAgainstCurrentTheme(Optional ByVal redrawMenuBar As Boolean = True)
    
    'Failsafe only
    Dim useTranslations As Boolean
    If (g_Language Is Nothing) Then
        useTranslations = False
    Else
        useTranslations = g_Language.TranslationActive
    End If
    
    'Before proceeding, cache any hotkey-related translations (so we don't have to keep translating them)
    Hotkeys.UpdateHotkeyLocalization
    
    Dim i As Long
    
    'Next, translate menu captions from English to the currently active language
    If useTranslations Then
        
        For i = 0 To m_NumOfMenus - 1
            
            'Ignoring separators and null-length captions, localize the current English caption
            With m_Menus(i)
                If (.me_Name <> MENU_SEPARATOR) Then
                    If (LenB(.me_TextEn) <> 0) Then
                        .me_TextTranslated = g_Language.TranslateMessage(.me_TextEn, SPECIAL_TRANSLATION_OBJECT_PREFIX & .me_Name)
                    End If
                End If
            End With
            
        Next i
    
    'English is active.  Simply mirror the English text to the localized field.
    Else
        For i = 0 To m_NumOfMenus - 1
            With m_Menus(i)
                If (.me_Name <> MENU_SEPARATOR) Then
                    If (LenB(.me_TextEn) <> 0) Then .me_TextTranslated = .me_TextEn
                End If
            End With
        Next i
    End If
    
    'Mnemonics need to be recalculated after caption changes
    DetermineMnemonics
    
    'Generate localized text for all hotkeys
    For i = 0 To m_NumOfMenus - 1
        With m_Menus(i)
            
            'Failsafe checks for separator and null menus (VB doesn't short-circuit, so we do it manually)
            If (.me_Name <> MENU_SEPARATOR) Then
            If (LenB(.me_TextEn) <> 0) Then
                If (.me_HotKeyID >= 0) Then .me_HotKeyTextTranslated = Hotkeys.GetHotkeyText(.me_HotKeyID)
            End If
            End If
            
        End With
    Next i
    
    'For non-separator, non-zero-length menus, combine caption, mnemonics, and hotkeys (if any)
    ' into a single, final, display-ready string
    For i = 0 To m_NumOfMenus - 1
    
        With m_Menus(i)
            If (.me_Name <> MENU_SEPARATOR) Then
                If (LenB(.me_TextEn) <> 0) Then
                    If (.me_HotKeyID >= 0) Then
                        .me_TextFinal = .me_TextWithMnemonics & vbTab & .me_HotKeyTextTranslated
                    Else
                        .me_TextFinal = .me_TextWithMnemonics
                    End If
                Else
                    .me_TextFinal = vbNullString
                End If
            Else
                .me_TextFinal = vbNullString
            End If
        End With
        
    Next i
    
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
                    mnuNameLvl1 = .me_TextTranslated
                    mnuNameLvl2 = vbNullString
                Else
                    mnuNameLvl2 = .me_TextTranslated
                End If
            
            'This menu does not have children, meaning it's clickable.
            Else
                
                'If this menu doesn't match the last level-2 menu index, reset the level 2 string.
                ' (This ensures that items following a sub-menu - e.g. File > Print, which comes after
                ' File > Batch Operations - don't mistakenly pick-up the second-level name of a
                ' preview menu.)
                If (.me_SubMenu <> lastLvl2Index) Then mnuNameLvl2 = vbNullString
                
                'Make sure this isn't just a menu separator
                If (.me_Name <> MENU_SEPARATOR) Then
                
                    'Append first- and second-level menu names, if any
                    If (LenB(mnuNameLvl1) <> 0) Then mnuNameFinal = mnuNameLvl1 & " > "
                    If (LenB(mnuNameLvl2) <> 0) Then mnuNameFinal = mnuNameFinal & mnuNameLvl2 & " > "
                    mnuNameFinal = mnuNameFinal & .me_TextTranslated
                    .me_TextSearchable = mnuNameFinal
                    
                Else
                    .me_TextSearchable = vbNullString
                End If
            
            End If
            
            'If we generated text for this menu, strip any trailing "...".  IMO ellipses do not add value to
            ' search results and simply clutter up the search list, but you could probably make a case either way
            ' based on current HIG (for example, see https://stackoverflow.com/a/637708/3511152)
            Const STR_ELLIPSES As String = "..."
            If (LenB(.me_TextSearchable) > 3) Then
                If (Right$(.me_TextSearchable, 3) = STR_ELLIPSES) Then .me_TextSearchable = Left$(.me_TextSearchable, Len(.me_TextSearchable) - 3)
            End If
            
        End With
        
    Next i
    
    'Some special menus must be dealt with now; note that some menus are already handled by dedicated callers
    ' (e.g. the "Languages" menu), while others must be handled here.
    Menus.UpdateSpecialMenu_RecentFiles
    Menus.UpdateSpecialMenu_RecentMacros
    Menus.UpdateSpecialMenu_WindowsOpen
    
    If redrawMenuBar Then DrawMenuBar FormMain.hWnd
    
End Sub

'Automatically determine new mnemonics for *all* menu captions.
Private Sub DetermineMnemonics()
    
    'Mnemonics use a (somewhat convoluted) automatic generation strategy,
    ' roughly akin to the way humans would manually create mnemonics.
    
    'First, recognize that all mnemonic decisions are made among sibling menus.
    ' Child menus (and menus with different parents) can and will reuse mnemonic characters.
    
    'Mnemonics are calculated as follows:
    '1) Ignore all non-alpha characters.  (Punctuation, numbers, whitespace are never used for mnemonics.)
    '2) If the language uses wide chars, use the original English text for the mnemonic search.
    '   (Otherwise, use localized text.)
    '3) If the first letter of the caption is unused, use that as the mnemonic.
    '4) If the caption is multi-word, start from the first letter of the *last* word in the caption
    '   and work backward. If you find an unused first-letter, use that as the mnemonic.
    '5) If the above strategies fail, start searching linearly through the caption until you either...
    '   - Find an unused letter (use that).
    '   - No letters match (no mnemonic for this entry).
    
    'If you were a true masochist, you could devise a way to shuffle around previous mnemonics to try
    ' and find a mnemonic for *all* captions, but PD isn't that insane.  The above steps make for a
    ' good effort that covers 99% of cases, and requires 0 effort on my part (or on volunteer
    ' translator's parts).
    
    'For each level (top, child-1, child-2), we must maintain a list of string characters that have
    ' already been used as mnemonics; duplicates obviously aren't valid.
    Dim mnLvl0 As String, mnLvl1 As String, mnLvl2 As String, mnTarget As String
    Dim mnChar As String, noKeys As Boolean
    
    Dim mnPos As Long
    mnPos = 0
    
    Dim i As Long
    For i = 0 To m_NumOfMenus - 1
        
        With m_Menus(i)
            
            'By default, set the mnemonic text to match the translated text.
            ' (If we can't find a valid mnemonic char, we just want to use the
            ' localized text as-is.)
            .me_TextWithMnemonics = .me_TextTranslated
            
            'First, we need to establish hierarchy for this menu item.  When we move *up* a level,
            ' we can erase the mnenomics list for the previous depth.
            
            'First, determine menu depth; this affects which target string we'll use for
            ' checking already-used mnemonics.
            If (.me_SubMenu = MENU_NONE) Then
                mnTarget = mnLvl0
                mnLvl1 = vbNullString
                mnLvl2 = vbNullString
            ElseIf (.me_SubSubMenu = MENU_NONE) Then
                mnTarget = mnLvl1
                mnLvl2 = vbNullString
            Else
                mnTarget = mnLvl2
            End If
            
        End With
        
        'With depth correctly set, we now want to skip any separator or null-length menus
        If (m_Menus(i).me_TextEn = MENU_SEPARATOR) Or (LenB(m_Menus(i).me_TextEn) = 0) Then GoTo NextMenuEntry
        
        'Defer to the mnemonic analyzer to solve for an actual mnemonic character (and character position)
        If EvaluateMnemonics(mnTarget, i, noKeys, mnPos, mnChar) Then
            
            'If a valid mnemonic index was found, mark the corresponding character with
            ' a leading ampersand (unless the string is e.g. Chinese, in which case we
            ' append the character to the end of the string, inside parentheses)
            If (mnPos > 0) Then
                
                With m_Menus(i)
                    
                    'Figure out whether to use the English char or a localized one
                    If noKeys Then
                        mnChar = Mid$(.me_TextEn, mnPos, 1)
                    Else
                        mnChar = Mid$(.me_TextTranslated, mnPos, 1)
                    End If
                    
                    'Check for e.g. Chinese strings
                    If (noKeys And (m_WideMenuMnemonics = PD_BOOL_AUTO)) Or (m_WideMenuMnemonics = PD_BOOL_TRUE) Then
                        
                        Const STR_ELLIPSES As String = "..."
                        Const STR_AMPERSAND As String = "&"
                        Const STR_MNEMONIC_START As String = " (&"
                        Const STR_MNEMONIC_END As String = ")"
                        
                        'Append the original English character to the end of the localized string
                        ' (if ellipses aren't used, or immediately before the ellipsis)
                        Dim posEllipsis As Long
                        posEllipsis = InStr(1, .me_TextTranslated, STR_ELLIPSES, vbBinaryCompare)
                        If (posEllipsis = 0) Then
                            .me_TextWithMnemonics = .me_TextTranslated & STR_MNEMONIC_START & UCase$(mnChar) & STR_MNEMONIC_END
                        Else
                            .me_TextWithMnemonics = Left$(.me_TextTranslated, posEllipsis - 1) & STR_MNEMONIC_START & UCase$(mnChar) & STR_MNEMONIC_END & Right$(.me_TextTranslated, Len(.me_TextTranslated) - (posEllipsis - 1))
                        End If
                        
                    'Place the marker directly inside the localized caption if we can
                    Else
                        If (Not noKeys) Then
                            If (mnPos > 1) Then
                                .me_TextWithMnemonics = Left$(.me_TextTranslated, mnPos - 1) & STR_AMPERSAND & Right$(.me_TextTranslated, Len(.me_TextTranslated) - (mnPos - 1))
                            Else
                                .me_TextWithMnemonics = STR_AMPERSAND & .me_TextTranslated
                            End If
                        End If
                    End If
                    
                    'Append the mnemonics character to our running tracker, so we don't reuse it
                    mnChar = LCase$(mnChar)
                    If (.me_SubMenu = MENU_NONE) Then
                        mnLvl0 = mnLvl0 & mnChar
                    ElseIf (.me_SubSubMenu = MENU_NONE) Then
                        mnLvl1 = mnLvl1 & mnChar
                    Else
                        mnLvl2 = mnLvl2 & mnChar
                    End If
                    
                End With
                
            End If
            
        End If

NextMenuEntry:

    Next i

End Sub

'Evaluate mnemonics for a target menu item.
' IMPORTANTLY, if this function returns FALSE, you *must* skip ahead to the next menu item.
' (ByRef returns aren't guaranteed to be accurate or relevant if this function returns FALSE.)
Private Function EvaluateMnemonics(ByRef mnTarget As String, ByVal mnuIndex As Long, ByRef noKeys As Boolean, ByRef mnPos As Long, ByRef mnChar As String) As Boolean
    
    EvaluateMnemonics = False
    
    'Reset trackers
    mnChar = vbNullString
    mnPos = 0
    
    'We're now going to repeat a series of tests.
    
    'First, we want to test to see if *any* characters in the string are alphabetical.
    ' (If they're not, this menu isn't a candidate for mnemonics.)
    Dim srcString As String
    srcString = m_Menus(mnuIndex).me_TextTranslated
    If (Not AtLeastOneAlphabetic(srcString)) Then
        
        'Try again with en-US
        srcString = m_Menus(mnuIndex).me_TextEn
        If (Not AtLeastOneAlphabetic(srcString)) Then
            EvaluateMnemonics = False
            Exit Function
        Else
            srcString = m_Menus(mnuIndex).me_TextTranslated
        End If
        
    End If
    
    'At least one character in the string is alphabetical.  Next, we want to see if
    ' *any* characters in the string map to a keyboard key.  If they don't, this is
    ' likely a Unicode string from a language like Chinese.
    noKeys = (Not AtLeastOnePhysicalKey(srcString))
    
    'If this is a string with no characters that map to physical keys, we actually want
    ' to use the *original English text* as our mnemonic.
    If noKeys Then
        srcString = m_Menus(mnuIndex).me_TextEn
        If (Not AtLeastOneAlphabetic(srcString)) Then
            EvaluateMnemonics = False
            Exit Function
        End If
    End If
    
    'We now know which string to test for mnemonics.
    
    'We're now going to repeat the above test on each character in the target string,
    ' using a (somewhat?) specialized strategy.
    
    'First, we always want to test the first character in the string.  It's the best
    ' candidate for a mnemonic, assuming that char is available at this menu level.
    If (IsMnemonicCandidate(srcString, 1) And IsMnemonicCharAvailable(mnTarget, srcString, 1)) Then
        mnPos = 1
        EvaluateMnemonics = True
        Exit Function
    End If
    
    'If we're still here, the first character in this caption is already being used by one
    ' of our sibling menus.
    
    'Next, see if this caption is multi-word.  If it is, we want to check the first letter
    ' of other words in the caption.
    If (InStr(1, srcString, " ", vbBinaryCompare) <> 0) Then
        
        'This is a multi-word caption.  Split the text into words.
        Dim listOfWords() As String
        listOfWords = Split(srcString, " ")
        
        Dim curWord As Long, prevWords As Long
        For curWord = UBound(listOfWords) To LBound(listOfWords) Step -1
            
            If (LenB(listOfWords(curWord)) > 0) Then
                
                'Find the first letter in this "word" and check it against our mnemonic list
                Dim j As Long
                For j = 1 To Len(listOfWords(curWord))
                    If IsMnemonicCandidate(listOfWords(curWord), j) Then
                        
                        'This is an alphabetic character that maps to a physical keyboard key.
                        ' Check it against our current mnemonic list - and IMPORTANTLY -
                        ' IF IT FAILS, do *not* check this word further.  (Instead, skip to
                        ' the first letter of the *next* word.)
                        If IsMnemonicCharAvailable(mnTarget, listOfWords(curWord), j) Then
                            
                            'We have the position of this character relative to the start of this word,
                            ' but we need the position relative to the start of the *original* string.
                            ' Add up the length of all preceding words, and assume one space between
                            ' each of them.  (This assumption may not technically be valid if a
                            ' translator inserts multiple spaces somewhere... I'm not sure what to
                            ' do in that case.)
                            mnPos = 0
                            For prevWords = LBound(listOfWords) To curWord - 1
                                mnPos = mnPos + Len(listOfWords(prevWords)) + 1
                            Next prevWords
                            mnPos = mnPos + j
                            
                            'We have what we need!  Exit immediately
                            EvaluateMnemonics = True
                            Exit Function
                            
                        Else
                            GoTo NextWord
                        End If
                        
                    End If

                Next j
                
            End If
NextWord:
        Next curWord
    
    End If
    
    'If we're still here, checking multiple words didn't work.
    
    'Do one last linear search through the string (starting at position 2), looking for an unused character.
    If (Len(srcString) >= 2) Then
        For j = 2 To Len(srcString)
            If IsMnemonicCandidate(srcString, j) Then
                If IsMnemonicCharAvailable(mnTarget, srcString, j) Then
                    mnPos = j
                    EvaluateMnemonics = True
                    Exit Function
                End If
            End If
        Next j
    End If
    
End Function

'Automatically determine new mnemonics for a single menu level.  This function also requires you to pass
' a stack object; the function will populate it with indices of menus whose mnemonic changes as a result
' of this function.  It is CRITICAL that you update those menu's captions to reflect the new mnemonics!
Private Sub DetermineMnemonics_SingleMenu(ByRef mnuName As String, ByRef dstStack As pdStack)
    
    Set dstStack = New pdStack
    
    'For detailed comments, please refer to the parent DetermineMnemonics sub.  This is just a
    ' one-off version designed for menus whose caption changes dynamically at run-time
    ' (e.g. "Undo" can become "Undo [operation]")
    
    'Start by resolving the menu name to an index.
    Dim mnuIndex As Long
    If (Not GetIndexFromName(mnuName, mnuIndex)) Then Exit Sub
    If (mnuIndex < 0) Then Exit Sub
    
    'The passed menu name was successfully mapped to an index in our menu table.
    ' We want to figure out how to identify *siblings* of this menu - these are
    ' the menus that can't share a mnemonic char with the target menu.
    Dim targetTopLevel As Long, targetSubLevel As Long, targetSubSubLevel As Long
    targetTopLevel = m_Menus(mnuIndex).me_TopMenu
    targetSubLevel = m_Menus(mnuIndex).me_SubMenu
    targetSubSubLevel = m_Menus(mnuIndex).me_SubSubMenu
    
    'Lots of trackers are required to assemble mnemonics accurately.
    Dim mnTarget As String, doThisItem As Boolean
    Dim noKeys As Boolean
    
    Dim mnPos As Long, mnChar As String, newText As String
    mnPos = 0
    
    Dim i As Long
    For i = 0 To m_NumOfMenus - 1
        
        'Ignore any separator or null-length menus, regardless of hierarchy
        If (m_Menus(i).me_TextEn = MENU_SEPARATOR) Or (LenB(m_Menus(i).me_TextEn) = 0) Then GoTo NextMenuEntry
        
        'Next, figure out if this menu exists in the same hierarchy
        ' as the target menu.
        With m_Menus(i)
            
            'Top-level menus have no sub or sub-sub menu IDs
            If (targetSubLevel = MENU_NONE) Then
                doThisItem = ((.me_SubMenu = MENU_NONE) And (.me_SubSubMenu = MENU_NONE))
                
            'Sub-menus need to be matched against sub-menus with the same parent menu
            ElseIf (targetSubSubLevel = MENU_NONE) Then
                doThisItem = ((.me_TopMenu = targetTopLevel) And (.me_SubSubMenu = MENU_NONE))
            
            'Sub-sub menus need to be matched against sub-sub-menus with the same parent and sub-parent
            Else
                doThisItem = ((.me_TopMenu = targetTopLevel) And (.me_SubMenu = targetSubLevel))
            End If
            
        End With
        
        'Ignore any menus not meeting the above criteria
        If doThisItem Then
        
            'By default, set the mnemonic text to match the translated text.
            ' (If we can't find a valid mnemonic char, we just want to use the
            ' localized text as-is.)
            m_Menus(i).me_TextWithMnemonics = m_Menus(i).me_TextTranslated
            
            'This menu needs to be analyzed!  Hand it off to the dedicated mnemonics analyzer.
            If EvaluateMnemonics(mnTarget, i, noKeys, mnPos, mnChar) Then
            
                'The analyzer determined a valid mnemonic for this menu.  Now we need to assign it.
                If (mnPos > 0) Then
                    
                    With m_Menus(i)
                        
                        'Figure out whether to use the English char or a localized one
                        If noKeys Then
                            mnChar = Mid$(.me_TextEn, mnPos, 1)
                        Else
                            mnChar = Mid$(.me_TextTranslated, mnPos, 1)
                        End If
                        
                        'Check for e.g. Chinese strings
                        If (noKeys And (m_WideMenuMnemonics = PD_BOOL_AUTO)) Or (m_WideMenuMnemonics = PD_BOOL_TRUE) Then
                            
                            Const STR_ELLIPSES As String = "..."
                            Const STR_AMPERSAND As String = "&"
                            Const STR_MNEMONIC_START As String = " (&"
                            Const STR_MNEMONIC_END As String = ")"
                            
                            'Append the original English character to the end of the localized string
                            ' (if ellipses aren't used, or immediately before the ellipsis)
                            Dim posEllipsis As Long
                            posEllipsis = InStr(1, .me_TextTranslated, STR_ELLIPSES, vbBinaryCompare)
                            If (posEllipsis = 0) Then
                                newText = .me_TextTranslated & STR_MNEMONIC_START & UCase$(mnChar) & STR_MNEMONIC_END
                            Else
                                newText = Left$(.me_TextTranslated, posEllipsis - 1) & STR_MNEMONIC_START & UCase$(mnChar) & STR_MNEMONIC_END & Right$(.me_TextTranslated, Len(.me_TextTranslated) - (posEllipsis - 1))
                            End If
                            
                        'Place the marker directly inside the localized caption
                        Else
                            If (mnPos > 1) Then
                                newText = Left$(.me_TextTranslated, mnPos - 1) & STR_AMPERSAND & Right$(.me_TextTranslated, Len(.me_TextTranslated) - (mnPos - 1))
                            ElseIf (mnPos = 1) Then
                                newText = STR_AMPERSAND & .me_TextTranslated
                            Else
                                newText = .me_TextTranslated
                            End If
                        End If
                        
                        'If this menu's caption changed as a result of the new mnemonic, add this
                        ' menu index to the passed stack.
                        If Strings.StringsNotEqual(.me_TextWithMnemonics, newText, False) Then
                            dstStack.AddInt i
                            .me_TextWithMnemonics = newText
                        End If
                        
                        'Append the mnemonics character to our running tracker, so we don't reuse it
                        mnChar = LCase$(mnChar)
                        mnTarget = mnTarget & mnChar
                        
                    End With
                    
                End If
                
            End If
            
        End If
        
NextMenuEntry:

    Next i

End Sub

'Returns TRUE if at least one character in the string is alphabetic
Private Function AtLeastOneAlphabetic(ByRef srcString As String) As Boolean
    
    AtLeastOneAlphabetic = False
    
    If (LenB(srcString) > 0) Then
    
        Dim i As Long
        For i = 1 To Len(srcString)
            AtLeastOneAlphabetic = (IsCharAlphaW(AscW(Mid$(srcString, i, 1))) <> 0)
            If AtLeastOneAlphabetic Then Exit Function
        Next i
        
    End If
    
End Function

'Returns TRUE if at least one character in the string maps to a physical keyboard key
Private Function AtLeastOnePhysicalKey(ByRef srcString As String) As Boolean

    AtLeastOnePhysicalKey = False
    
    If (LenB(srcString) > 0) Then
    
        Dim i As Long
        For i = 1 To Len(srcString)
            
            'Only check alphabetic keys (e.g. "..." maps to a physical key, but isn't valid for this purpose)
            If (IsCharAlphaW(AscW(Mid$(srcString, i, 1))) <> 0) Then
                AtLeastOnePhysicalKey = (VkKeyScanW(AscW(Mid$(srcString, i, 1))) <> &HFFFF)
                If AtLeastOnePhysicalKey Then Exit Function
            End If
            
        Next i
        
    End If
    
End Function

'Returns TRUE if the character at position [charIndex] is...
' 1) Alpha (non-numeric, punctuation, whitespace, etc)
' 2) Maps to a hardware key
Private Function IsMnemonicCandidate(ByRef srcString As String, ByVal charIndex As Long) As Boolean
    
    IsMnemonicCandidate = False
    On Error GoTo CharFail
    
    If (charIndex > 0) And (charIndex <= Len(srcString)) Then
        
        Dim testChar As String
        testChar = Mid$(srcString, charIndex, 1)
        
        'Test for alphabetic characters (no punctuation or numbers)
        If (IsCharAlphaW(AscW(testChar)) <> 0) Then
            
            'See if the character maps to a virtual key; per MSDN (https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-vkkeyscanw):
            ' "If the function finds no key that translates to the passed character code,
            '  both the low-order and high-order bytes contain –1."
            IsMnemonicCandidate = (VkKeyScanW(AscW(testChar)) <> &HFFFF)
            
        End If
        
    End If
    
    Exit Function
    
CharFail:
    IsMnemonicCandidate = False
    
End Function

'Returns TRUE if the character at index targetCharIndex of srcString does *not* appear in the
' character list passed via srcCharList (which must be maintained in LCASE state for fast compares).
Private Function IsMnemonicCharAvailable(ByRef srcCharList As String, ByRef srcString As String, ByVal targetCharIndex As Long) As Boolean
    
    IsMnemonicCharAvailable = True
    
    On Error GoTo BadCharIndex
    
    If (LenB(srcCharList) > 0) Then
        
        Dim targetChar As String
        targetChar = LCase$(Mid$(srcString, targetCharIndex, 1))
        
        Dim i As Long
        For i = 1 To Len(srcCharList)
            IsMnemonicCharAvailable = (InStr(1, srcCharList, targetChar, vbBinaryCompare) = 0)
            If (Not IsMnemonicCharAvailable) Then Exit Function
        Next i
        
    End If
    
    Exit Function
    
BadCharIndex:
    PDDebug.LogAction "WARNING!  Menus.IsMnemonicCharAvailable was passed a bad char index"
    IsMnemonicCharAvailable = False
    
End Function

'Given a menu name, return the corresponding menu caption (localized, with accelerator)
Public Function GetCaptionFromName(ByRef mnuName As String, Optional ByVal returnTranslation As Boolean = True) As String

    'Resolve the menu name into an index into our menu collection
    Dim mnuIndex As Long
    If GetIndexFromName(mnuName, mnuIndex) Then
        
        If (mnuIndex >= 0) Then
            If returnTranslation Then
                GetCaptionFromName = m_Menus(mnuIndex).me_TextTranslated
            Else
                GetCaptionFromName = m_Menus(mnuIndex).me_TextEn
            End If
        End If
        
    End If
    
End Function

'Given a menu name, return the corresponding index into the local m_Menus() collection.
Private Function GetIndexFromName(ByRef mnuName As String, ByRef dstIndex As Long) As Boolean

    'Resolve the menu name into an index into our menu collection
    Dim i As Long
    dstIndex = -1
    
    For i = 0 To m_NumOfMenus - 1
        If Strings.StringsEqual(mnuName, m_Menus(i).me_Name, True) Then
            dstIndex = i
            Exit For
        End If
    Next i
    
    GetIndexFromName = (dstIndex >= 0)
    
    'Debugging missing menu names can be helpful when adding new tools:
    'If (Not GetIndexFromName) Then InternalMenuWarning "GetIndexFromName", "no match found for name: " & mnuName
    
End Function

'Return a list of all menus and menu attributes.  This is used by the customize hotkeys dialog to both display
' menu names and attributes, and to correlate menu text against a list of canonical menu IDs (which is how the
' hotkey engine associates hotkeys <-> actions <-> menus).
'
'Returns the number of menus in the list, with a guarantee that the target list is resized to [0, numMenus-1]
Public Function GetCopyOfAllMenus(ByRef dstMenuList() As PD_MenuEntry) As Long
    
    'Still TODO: does our list of menus need to be curated before sending it externally?
    ' IDK - but there may be menus that we don't want associated with hotkeys, and this could be
    ' where we strip them out of the menu list.
    ReDim dstMenuList(0 To m_NumOfMenus - 1) As PD_MenuEntry
    
    Dim i As Long
    For i = 0 To m_NumOfMenus - 1
        dstMenuList(i) = m_Menus(i)
    Next i
    
    GetCopyOfAllMenus = m_NumOfMenus
    
End Function

'Given a menu's search text, return the corresponding menu name.  (Used by PD's right-hand search box.)
Public Function GetNameFromSearchText(ByRef srcSearchText As String) As String
    
    If (LenB(srcSearchText) > 0) Then
        
        'Search the menu list for a menu search text matching the passed one.  (Search text includes
        ' menu hierarchy, e.g. "Effects > Blur > Gaussian".)
        Dim i As Long
        For i = 0 To m_NumOfMenus - 1
            If Strings.StringsEqual(srcSearchText, m_Menus(i).me_TextSearchable, True) Then
                GetNameFromSearchText = m_Menus(i).me_Name
                Exit Function
            End If
        Next i
        
    End If
        
    'If the previous loop found no matches, something went horribly wrong
    PDDebug.LogAction "WARNING!  Menus.LaunchAction_BySearch couldn't find a match for: " & srcSearchText

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
'NOTE: this function doesn't actually render the new menu text; the assumption is that new hotkeys
' are assigned in bulk, and for performance reasons you'll want to do a single full-menu-system refresh
' after the bulk assignments finish.
Public Sub NotifyMenuHotkey(ByRef actionID As String, Optional ByVal hkID As Long = NO_MENU_HOTKEY)
    
    'Resolve the menuID into a list of indices.  (Note that menus can share the same ID, meaning there can be more
    ' than one physical menu associated with a given hotkey; this is used to display correct hotkeys for menus like
    ' Import > From Clipboard and Edit > Paste which appear in two places for usability reasons, but ultimately map
    ' to the same action.)
    Dim listOfMatches() As Long, numOfMatches As Long
    GetAllMatchingMenuIndices actionID, numOfMatches, listOfMatches
    
    'No menus may match the current hotkey.  That's okay - some hotkeys trigger actions (like "select tool")
    ' that don't have a menu equivalent.  They'll still work just fine, but we don't have to paint those
    ' hotkey combinations to any corresponding menu.
    If (numOfMatches > 0) Then
        
        'Before we enter the loop, generate a translated text representation of this hotkey
        Dim hotkeyText As String
        hotkeyText = Hotkeys.GetHotkeyText(hkID)
        
        Dim i As Long
        For i = 0 To numOfMatches - 1
            With m_Menus(listOfMatches(i))
                .me_HotKeyID = hkID
                .me_HotKeyTextTranslated = hotkeyText
            End With
        Next i
        
    End If

End Sub

'After the user edits hotkeys, this class needs to be notified so it can erase existing hotkey data.
' (After calling this function, you obviously need to manually update hotkey data to match!)
Public Sub NotifyHotkeysChanged()
    Dim i As Long
    For i = 0 To m_NumOfMenus - 1
        m_Menus(i).me_HotKeyID = -1
        m_Menus(i).me_HotKeyTextTranslated = vbNullString
    Next i
End Sub

'This function is part of an incredibly unpleasant workaround for ensuring that menu navigation
' still works.  When a hotkey like Alt+F is used to send focus to PD's main menu, the active window
' does not actually receive a WM_KILLFOCUS or WM_ACTIVATE (with relevant flags) message, so it will
' unknowingly eat keypresses (arrow or otherwise).  We can work around this by checking for menu
' drop state on the main window (FormMain in PD) and suspending key tracking if the menu is in
' dropped/focused state.  Note that this function also returns TRUE if PD's system menu is dropped.
Public Function IsMainMenuActive() As Boolean

    Dim tmpMBI As Win32_MENUBARINFO
    tmpMBI.cbSize = LenB(tmpMBI)
    
    'Check main menu first
    Const OBJID_MENU As Long = &HFFFFFFFD
    If (GetMenuBarInfo(FormMain.hWnd, OBJID_MENU, 0&, VarPtr(tmpMBI)) <> 0) Then
        IsMainMenuActive = ((tmpMBI.fFlags And 3&) <> 0&)
        
        'Check system menu next
        If (Not IsMainMenuActive) Then
            Const OBJID_SYSMENU As Long = &HFFFFFFFF
            If (GetMenuBarInfo(FormMain.hWnd, OBJID_SYSMENU, 0&, VarPtr(tmpMBI)) <> 0) Then
                IsMainMenuActive = ((tmpMBI.fFlags And 3&) <> 0&)
            End If
        End If
        
    End If
    
End Function

'Helper check for resolving menu enablement by menu name.  Note that PD *does not* enforce unique menu names; in fact, they are
' specifically allowed by design.  As such, this function only returns the *first* matching entry, with the assumption that
' same-named menus are enabled and disabled as a group.
Public Function IsMenuChecked(ByRef mnuName As String) As Boolean

    'Resolve the menu name into an index into our menu collection
    Dim mnuIndex As Long
    If GetIndexFromName(mnuName, mnuIndex) Then
    
        Dim hMenu As Long, hMenuIndex As Long
        hMenu = GetHMenu_FromIndex(mnuIndex, True)
        
        If (m_Menus(mnuIndex).me_SubMenu < 0) Then
            hMenuIndex = m_Menus(mnuIndex).me_TopMenu
        ElseIf (m_Menus(mnuIndex).me_SubSubMenu < 0) Then
            hMenuIndex = m_Menus(mnuIndex).me_SubMenu
        Else
            hMenuIndex = m_Menus(mnuIndex).me_SubSubMenu
        End If
        
        'Fill a MENUITEMINFO struct
        Dim tmpMii As Win32_MenuItemInfoW
        tmpMii.cbSize = LenB(tmpMii)
        tmpMii.fMask = MIIM_STATE
        If (GetMenuItemInfoW(hMenu, hMenuIndex, 1, tmpMii) <> 0) Then
            IsMenuChecked = ((tmpMii.fState And MFS_CHECKED) <> 0)
        Else
            InternalMenuWarning "IsMenuChecked", "GetMenuItemInfoW failed: " & Err.LastDllError
        End If
        
    End If
    
End Function

'Helper check for resolving menu enablement by menu name.  Note that PD *does not* enforce unique menu names;
' in fact, duplicates are specifically allowed by design.  As such, this function only returns the *first*
' matching entry, with the assumption that same-named menus are enabled and disabled as a group.
'
'Optionally, you can test to see if the specified menu even exists in the collection using the optional
' mnuDoesntExist parameter.  This is helpful if you want to prevent an action from running if its
' corresponding menu is disabled; if an action doesn't HAVE a corresponding menu, you can ignore the return
' of this function entirely, and simply execute that action as-is (assuming the action self-validates).
Public Function IsMenuEnabled(ByRef mnuName As String, Optional ByRef dstMenuDoesntExist As Boolean = False) As Boolean

    'Resolve the menu name into an index into our menu collection
    Dim mnuIndex As Long
    If GetIndexFromName(mnuName, mnuIndex) Then
    
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
        
    'If we don't find a menu matching this name, return it via the optional parameter
    Else
        dstMenuDoesntExist = True
    End If

End Function

'Helper check for resolving menu enablement by menu name.  Note that PD *does not* enforce unique menu names; in fact, they are
' specifically allowed by design.  As such, this function only returns the *first* matching entry, with the assumption that
' same-named menus are enabled and disabled as a group.
Public Sub SetMenuChecked(ByRef mnuName As String, Optional ByVal isChecked As Boolean = True)
    
    'Avoid redundant calls
    If (Menus.IsMenuChecked(mnuName) = isChecked) Then Exit Sub
    
    'Resolve the menu name into an index into our menu collection
    Dim mnuIndex As Long
    If GetIndexFromName(mnuName, mnuIndex) Then
    
        Dim hMenu As Long, hMenuIndex As Long
        hMenu = GetHMenu_FromIndex(mnuIndex, True)
        
        If (m_Menus(mnuIndex).me_SubMenu < 0) Then
            hMenuIndex = m_Menus(mnuIndex).me_TopMenu
        ElseIf (m_Menus(mnuIndex).me_SubSubMenu < 0) Then
            hMenuIndex = m_Menus(mnuIndex).me_SubMenu
        Else
            hMenuIndex = m_Menus(mnuIndex).me_SubSubMenu
        End If
        
        'Fill a MENUITEMINFO struct and retrieve current menu state
        Dim tmpMii As Win32_MenuItemInfoW
        tmpMii.cbSize = LenB(tmpMii)
        tmpMii.fMask = MIIM_STATE
        If (GetMenuItemInfoW(hMenu, hMenuIndex, 1, tmpMii) = 0) Then InternalMenuWarning "SetMenuChecked", "GetMenuItemInfoW failed: " & Err.LastDllError
        
        'Set the checked bit flag, then report the change to the API
        If isChecked Then
            tmpMii.fState = tmpMii.fState Or MFS_CHECKED
        Else
            tmpMii.fState = tmpMii.fState And (Not MFS_CHECKED)
        End If
        
        If (SetMenuItemInfoW(hMenu, hMenuIndex, 1, tmpMii) = 0) Then InternalMenuWarning "SetMenuChecked", "SetMenuItemInfoW failed: " & Err.LastDllError
        
    End If
    
End Sub

Public Sub SetMenuEnabled(ByRef mnuName As String, Optional ByVal isEnabled As Boolean = True)
    
    'Avoid redundant calls
    If (Menus.IsMenuEnabled(mnuName) = isEnabled) Then Exit Sub
    
    'Resolve the menu name into an index into our menu collection
    Dim mnuIndex As Long
    If GetIndexFromName(mnuName, mnuIndex) Then
    
        Dim hMenu As Long, hMenuIndex As Long
        hMenu = GetHMenu_FromIndex(mnuIndex, True)
        
        If (m_Menus(mnuIndex).me_SubMenu < 0) Then
            hMenuIndex = m_Menus(mnuIndex).me_TopMenu
        ElseIf (m_Menus(mnuIndex).me_SubSubMenu < 0) Then
            hMenuIndex = m_Menus(mnuIndex).me_SubMenu
        Else
            hMenuIndex = m_Menus(mnuIndex).me_SubSubMenu
        End If
        
        If isEnabled Then
            EnableMenuItem hMenu, hMenuIndex, MFE_BYPOSITION Or MFE_ENABLED
        Else
            EnableMenuItem hMenu, hMenuIndex, MFE_BYPOSITION Or MFE_DISABLED Or MFE_GRAYED
        End If
        
        'Top-level menus need to be redrawn immediately; other ones do not
        If (m_Menus(mnuIndex).me_SubMenu < 0) Then DrawMenuBar FormMain.hWnd
        
    End If
    
End Sub

Public Sub SetMnemonicsBehavior(ByVal newBehavior As PD_BOOL)
    m_WideMenuMnemonics = newBehavior
    If (m_WideMenuMnemonics <> PD_BOOL_TRUE) And (m_WideMenuMnemonics <> PD_BOOL_FALSE) And (m_WideMenuMnemonics <> PD_BOOL_AUTO) Then
        m_WideMenuMnemonics = PD_BOOL_AUTO
    End If
End Sub

Private Sub GetAllMatchingMenuIndices(ByRef menuID As String, ByRef numOfMenus As Long, ByRef menuArray() As Long)

    'At present, there will never be more than two menus matching a given ID; this can be revisited in the future,
    ' but imposing an upper limit improves performance (because we can exit immediately if a second match is found)
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

'Some menus in PD (like the Recent Files menu, or the Tools > Languages menu) are directly
' modified at run-time.  In PD, it is easier to wipe these entire menus dynamically rather
' than rebuild them from scratch.
'
'IMPORTANT NOTE: to erase an entire submenu, pass ALL_MENU_SUBITEMS as the subMenuID or subSubMenuID,
'                whichever is relevant. ALL_MENU_SUBITEMS indicates "erase everything that matches the
'                two preceding entries, except for the top-level menu itself".
'
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

'Some of PD's menus obey special rules.  (For example, menus that add/remove entries at run-time.)  These menus have their own
' helper update functions that can be called on demand, separate from other menus in the project.
Public Sub UpdateSpecialMenu_Language(ByVal numOfLanguages As Long, ByRef availableLanguages() As PDLanguageFile)

    'Retrieve handles to the parent menu
    Dim hMenu As Long
    hMenu = GetMenu(FormMain.hWnd)
    hMenu = GetSubMenu(hMenu, 7)
    hMenu = GetSubMenu(hMenu, 0)
    
    'Prepare a MenuItemInfo struct
    Dim tmpMii As Win32_MenuItemInfoW
    tmpMii.cbSize = LenB(tmpMii)
    tmpMii.fMask = MIIM_STRING
    
    If (hMenu <> 0) Then
        
        'Add anew captions for all the current menu entries.  (Note that the language manager has handled the actual creation
        ' of these menu objects; we use VB itself for that.)
        Dim i As Long
        For i = 0 To numOfLanguages - 1
            tmpMii.dwTypeData = StrPtr(availableLanguages(i).LangName)
            SetMenuItemInfoW hMenu, i, 1&, tmpMii
        Next i
        
    Else
        InternalMenuWarning "UpdateSpecialMenu_Language", "null hMenu"
    End If
    
End Sub

'The main form's Window menu displays a list of open images.  This list must be updated whenever...
' 1) An image is loaded
' 2) An image is unloaded
' 3) A different image is "activated" (e.g. selected for editing)
' 4) The current image is saved to a different filename
Public Sub UpdateSpecialMenu_WindowsOpen()
    
    Dim i As Long
    
    'Quick branch for "no open images" state (it's much easier to handle)
    If (PDImages.GetNumOpenImages > 0) Then
        
        'Images are potentially stored non-sequentially.  Retrieve a list of active image IDs from the
        ' central image manager.
        Dim listOfOpenImages As pdStack
        PDImages.GetListOfActiveImageIDs listOfOpenImages
        
        'Ensure the correct number of menus are available.  (This may involve freeing existing menus
        ' when an image is closed, or adding new menus when an image is opened.)
        
        'This limit is effectively arbitrary, but useability of this menu is kind of pointless
        ' past a certain point.  (Windows may have its own menu count limit as well, idk; I deliberately
        ' prefer to stay well beneath that amount.)
        Const MAX_NUM_MENU_ENTRIES As Long = 64
        
        Dim numImagesAllowed As Long
        numImagesAllowed = PDMath.Min2Int(MAX_NUM_MENU_ENTRIES, listOfOpenImages.GetNumOfInts)
        
        If (FormMain.MnuWindowOpen.Count > numImagesAllowed) Then
            For i = FormMain.MnuWindowOpen.Count - 1 To numImagesAllowed - 1 Step -1
                If (i <> 0) Then Unload FormMain.MnuWindowOpen(i)
            Next i
        End If
        
        Dim curMenuCount As Long
        curMenuCount = FormMain.MnuWindowOpen.Count
        If (curMenuCount < numImagesAllowed) Then
            For i = curMenuCount To numImagesAllowed - 1
                Load FormMain.MnuWindowOpen(i)
            Next i
        End If
        
        'The correct number of menu entries are now available.
        
        'To ensure offsets retrieved from API menu calls are valid, we need to ensure the separator bar
        ' above the open window section is visible *before* interacting with items beneath it.
        ' (FYI: individual menus start at index 10 (9 is the separator bar above the first open image entry))
        Const MENU_OFFSET As Long = 10
        FormMain.MnuWindow(MENU_OFFSET - 1).Visible = True
        
        'We now need to set all menu captions to match the filename of each open image.  (As part of setting
        ' the correct names, we'll also set visible/enabled/checked state.)
        
        'Menu bar itself
        Dim hMenu As Long
        hMenu = GetMenu(FormMain.hWnd)
        
        'Window menu
        hMenu = GetSubMenu(hMenu, 9&)
        
        'Prepare a MenuItemInfo struct
        Dim tmpMii As Win32_MenuItemInfoW
        tmpMii.cbSize = LenB(tmpMii)
        tmpMii.fMask = MIIM_STRING
        
        'Note that we have to use WAPI to do this, because filenames may have Unicode chars.
        For i = 0 To numImagesAllowed - 1
            
            'Use VB to set the rest of the parameters; this will also trigger a DrawMenuBar call
            With FormMain.MnuWindowOpen(i)
                .Visible = True
                .Enabled = True
                .Checked = (listOfOpenImages.GetInt(i) = PDImages.GetActiveImageID)
            End With
            
            'Retrieve the caption (which should be the location on-disk, unless the image hasn't been saved
            ' in which case the loader will have assigned the image a "suggested" filename)
            Dim tmpCaption As String
            If PDImages.GetImageByID(listOfOpenImages.GetInt(i)).ImgStorage.DoesKeyExist("CurrentLocationOnDisk") Then
                tmpCaption = Files.FileGetName(PDImages.GetImageByID(listOfOpenImages.GetInt(i)).ImgStorage.GetEntry_String("CurrentLocationOnDisk"), False)
            End If
            
            If (LenB(tmpCaption) = 0) Then
                tmpCaption = Files.FileGetName(PDImages.GetImageByID(listOfOpenImages.GetInt(i)).ImgStorage.GetEntry_String("OriginalFileName"), False)
            End If
            
            'Assign the caption via WAPI to preserve Unicode chars
            tmpMii.dwTypeData = StrPtr(tmpCaption)
            SetMenuItemInfoW hMenu, MENU_OFFSET + i, 1&, tmpMii
            
        Next i
        
        'Use the API to trigger a state change for any new captions
        DrawMenuBar FormMain.hWnd
        
    'No open images.  Unload all menu items and hide the separator bar at the top of this section.
    Else
    
        If (FormMain.MnuWindowOpen.Count > 1) Then
            For i = 1 To FormMain.MnuWindowOpen.Count - 1
                Unload FormMain.MnuWindowOpen(i)
            Next i
        End If
        
        'Hide the final instance and the separator bar above it
        FormMain.MnuWindow(9).Visible = False
        FormMain.MnuWindowOpen(0).Visible = False
    
    End If
    
End Sub

'Whenever the "File > Open Recent" menu is modified, we need to modify our internal list of recent file items.
' (We manually track this menu so we can handle translations correctly for the items at the bottom of the menu,
' e.g. "Load all" and "Clear list", as well as menu captions for recent files with Unicode filenames.)
Public Sub UpdateSpecialMenu_RecentFiles()

    'Start by retrieving a handle to the menu in question
    If (Not g_RecentFiles Is Nothing) Then
        
        Dim hMenu As Long
        hMenu = GetMenu(FormMain.hWnd)
        hMenu = GetSubMenu(hMenu, 0&)
        hMenu = GetSubMenu(hMenu, 2&)
        
        'Prepare a MenuItemInfo struct
        Dim tmpMii As Win32_MenuItemInfoW
        tmpMii.cbSize = LenB(tmpMii)
        tmpMii.fMask = MIIM_STRING
        
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
            tmpMii.dwTypeData = StrPtr(tmpString)
            If (Not listIsEmpty) Then SetMenuItemInfoW hMenu, numOfMRUFiles + 1, 1&, tmpMii
            
            tmpString = g_Language.TranslateMessage("Clear recent image list")
            tmpMii.dwTypeData = StrPtr(tmpString)
            If (Not listIsEmpty) Then SetMenuItemInfoW hMenu, numOfMRUFiles + 2, 1&, tmpMii
                
            'Finally, manually place the captions for all recent file filenames, while handling the special
            ' case of an empty list.
            If listIsEmpty Then
                tmpString = g_Language.TranslateMessage("empty")
                tmpMii.dwTypeData = StrPtr(tmpString)
                SetMenuItemInfoW hMenu, 0&, 1&, tmpMii
            Else
                
                'If actual MRU paths exist, note that we apply them *without* translations, obviously.
                Dim i As Long, hotkeyNumber As Long
                For i = 0 To numOfMRUFiles - 1
                    
                    tmpString = g_RecentFiles.GetMenuCaption(i)
                    
                    'Entries under "10" get a free accelerator of the form "Ctrl+i"
                    If (i < 10) Then
                        hotkeyNumber = i + 1
                        If (i = 9) Then hotkeyNumber = 0
                        tmpString = tmpString & vbTab & g_Language.TranslateMessage("Ctrl") & "+" & g_Language.TranslateMessage("Shift") & "+" & hotkeyNumber
                    End If
                    
                    tmpMii.dwTypeData = StrPtr(tmpString)
                    SetMenuItemInfoW hMenu, i, 1&, tmpMii
                    
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
        Dim tmpMii As Win32_MenuItemInfoW
        tmpMii.cbSize = LenB(tmpMii)
        tmpMii.fMask = MIIM_STRING
        
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
            tmpMii.dwTypeData = StrPtr(tmpString)
            If (Not listIsEmpty) Then SetMenuItemInfoW hMenu, numOfMRUFiles + 1, 1&, tmpMii Else SetMenuItemInfoW hMenu, 2, 1&, tmpMii
                
            'Finally, manually place the captions for all recent file filenames, while handling the special
            ' case of an empty list.
            If listIsEmpty Then
                
                tmpString = g_Language.TranslateMessage("empty")
                tmpMii.dwTypeData = StrPtr(tmpString)
                SetMenuItemInfoW hMenu, 0&, 1&, tmpMii
                
            Else
                
                'If actual MRU paths exist, note that we apply them *without* translations, obviously.
                Dim i As Long
                For i = 0 To numOfMRUFiles - 1
                    tmpString = g_RecentMacros.GetSpecificMRUCaption(i)
                    tmpMii.dwTypeData = StrPtr(tmpString)
                    SetMenuItemInfoW hMenu, i, 1&, tmpMii
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
        Dim tmpMii As Win32_MenuItemInfoW
        tmpMii.cbSize = LenB(tmpMii)
        tmpMii.fMask = MIIM_STRING
        tmpMii.dwTypeData = StrPtr(m_Menus(mnuIndex).me_TextFinal)
        
        SetMenuItemInfoW hMenu, GetHMenuIndex(mnuIndex), 1&, tmpMii
        
    Else
        InternalMenuWarning "UpdateMenuText_ByIndex", "null hMenu: " & mnuIndex
    End If

End Sub

Private Sub InternalMenuWarning(ByRef funcName As String, ByRef errMsg As String)
    PDDebug.LogAction "WARNING!  Menus." & funcName & " reported: " & errMsg
End Sub
