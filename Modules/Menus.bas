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
        AddMenuItem "Basic layer...", "layer_addbasic", 4, 0, 0
        AddMenuItem "Blank layer", "layer_addblank", 4, 0, 1
        AddMenuItem "Duplicate of current layer", "layer_duplicate", 4, 0, 2, "edit_copy"
        AddMenuItem "From clipboard", "layer_addfromclipboard", 4, 0, 4, "edit_paste"
        AddMenuItem "From file...", "layer_addfromfile", 4, 0, 5, "file_open"
        AddMenuItem "From visible layers", "layer_addfromvisiblelayers", 4, 0, 6
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
        
    
    'Adjustments Menu
    AddMenuItem "Adjustments", "adj_top", 6
    AddMenuItem "Auto-correct", "adj_autocorrect", 6, 0
        AddMenuItem "Color", "adj_autocorrectcolor", 6, 0, 0
        AddMenuItem "Contrast", "adj_autocorrectcontrast", 6, 0, 1
        AddMenuItem "Lighting", "adj_autocorrectlighting", 6, 0, 2
        AddMenuItem "Shadows and highlights", "adj_autocorrectsandh", 6, 0, 3
    AddMenuItem "Auto-enhance", "adj_autoenhance", 6, 1
        AddMenuItem "Color", "adj_autoenhancecolor", 6, 1, 0
        AddMenuItem "Contrast", "adj_autoenhancecontrast", 6, 1, 1
        AddMenuItem "Lighting", "adj_autoenhancelighting", 6, 1, 2
        AddMenuItem "Shadows and highlights", "adj_autoenhancesandh", 6, 1, 3
    AddMenuItem "Black and white...", "adj_blackandwhite", 6, 3
    AddMenuItem "Brightness and contrast...", "adj_bandc", 6, 4
    AddMenuItem "Color balance...", "adj_colorbalance", 6, 5
    AddMenuItem "Curves...", "adj_curves", 6, 6
    AddMenuItem "Levels...", "adj_levels", 6, 7
    AddMenuItem "Shadows and highlights...", "adj_sandh", 6, 8
    AddMenuItem "Vibrance...", "adj_vibrance", 6, 9
    AddMenuItem "White balance...", "adj_whitebalance", 6, 10
    
    AddMenuItem "Channels", "adj_channels", 6, 12
        AddMenuItem "Channel mixer...", "adj_channelmixer", 6, 12, 0
        AddMenuItem "Rechannel...", "adj_rechannel", 6, 12, 1
        AddMenuItem "Maximum channel", "adj_maxchannel", 6, 12, 3
        AddMenuItem "Minimum channel", "adj_minchannel", 6, 12, 4
        AddMenuItem "Shift channels left", "adj_shiftchannelsleft", 6, 12, 6
        AddMenuItem "Shift channels right", "adj_shiftchannelsright", 6, 12, 7
    AddMenuItem "Color", "adj_color", 6, 13
        AddMenuItem "Color balance...", "adj_colorbalance", 6, 13, 0
        AddMenuItem "White balance...", "adj_whitebalance", 6, 13, 1
        AddMenuItem "Hue and saturation...", "adj_hsl", 6, 13, 3
        AddMenuItem "Temperature...", "adj_temperature", 6, 13, 4
        AddMenuItem "Tint...", "adj_tint", 6, 13, 5
        AddMenuItem "Vibrance...", "adj_vibrance", 6, 13, 6
        AddMenuItem "Black and white...", "adj_blackandwhite", 6, 13, 8
        AddMenuItem "Colorize...", "adj_colorize", 6, 13, 9
        AddMenuItem "Replace color...", "adj_replacecolor", 6, 13, 10
        AddMenuItem "Sepia...", "adj_sepia", 6, 13, 11
    AddMenuItem "Histogram", "adj_histogram", 6, 14
        AddMenuItem "Display histogram...", "adj_histogramdisplay", 6, 14, 0
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
    AddMenuItem "Effects", "effects_top", 7
    AddMenuItem "Artistic", "effects_artistic", 7, 0
        AddMenuItem "Colored pencil...", "effects_colorpencil", 7, 0, 0
        AddMenuItem "Comic book...", "effects_comicbook", 7, 0, 1
        AddMenuItem "Figured glass (dents)...", "effects_figuredglass", 7, 0, 2
        AddMenuItem "Film noir...", "effects_filmnoir", 7, 0, 3
        AddMenuItem "Glass tiles...", "effects_glasstiles", 7, 0, 4
        AddMenuItem "Kaleidoscope...", "effects_kaleidoscope", 7, 0, 5
        AddMenuItem "Modern art...", "effects_modernart", 7, 0, 6
        AddMenuItem "Oil painting...", "effects_oilpainting", 7, 0, 7
        AddMenuItem "Posterize...", "effects_posterize", 7, 0, 8
        AddMenuItem "Relief...", "effects_relief", 7, 0, 9
        AddMenuItem "Stained glass...", "effects_stainedglass", 7, 0, 10
    AddMenuItem "Blur", "effects_blur", 7, 1
        AddMenuItem "Box blur...", "effects_boxblur", 7, 1, 0
        AddMenuItem "Gaussian blur...", "effects_gaussianblur", 7, 1, 1
        AddMenuItem "Surface blur...", "effects_surfaceblur", 7, 1, 2
        AddMenuItem "Motion blur...", "effects_motionblur", 7, 1, 4
        AddMenuItem "Radial blur...", "effects_radialblur", 7, 1, 5
        AddMenuItem "Zoom blur...", "effects_zoomblur", 7, 1, 6
        AddMenuItem "Kuwahara filter...", "effects_kuwahara", 7, 1, 8
        AddMenuItem "Symmetric nearest-neighbor...", "effects_snn", 7, 1, 9
    AddMenuItem "Distort", "effects_distort", 7, 2
        AddMenuItem "Correct existing distortion...", "effects_fixlensdistort", 7, 2, 0
        AddMenuItem "Donut...", "effects_donut", 7, 2, 2
        AddMenuItem "Lens...", "effects_lens", 7, 2, 3
        AddMenuItem "Pinch and whirl...", "effects_pinchandwhirl", 7, 2, 4
        AddMenuItem "Poke...", "effects_poke", 7, 2, 5
        AddMenuItem "Ripple...", "effects_ripple", 7, 2, 6
        AddMenuItem "Squish...", "effects_squish", 7, 2, 7
        AddMenuItem "Swirl...", "effects_swirl", 7, 2, 8
        AddMenuItem "Waves...", "effects_waves", 7, 2, 9
        AddMenuItem "Miscellaneous...", "effects_miscdistort", 7, 2, 11
    AddMenuItem "Edges", "effects_edges", 7, 3
        AddMenuItem "Emboss...", "effects_emboss", 7, 3, 0
        AddMenuItem "Enhance edges...", "effects_enhanceedges", 7, 3, 1
        AddMenuItem "Find edges...", "effects_findedges", 7, 3, 2
        AddMenuItem "Range filter...", "effects_rangefilter", 7, 3, 4
        AddMenuItem "Trace contour...", "effects_tracecontour", 7, 3, 4
    AddMenuItem "Light and shadow", "effects_lightandshadow", 7, 4
        AddMenuItem "Black light...", "effects_blacklight", 7, 4, 0
        AddMenuItem "Cross-screen...", "effects_crossscreen", 7, 4, 1
        AddMenuItem "Rainbow...", "effects_rainbow", 7, 4, 2
        AddMenuItem "Sunshine...", "effects_sunshine", 7, 4, 3
        AddMenuItem "Dilate...", "effects_dilate", 7, 4, 5
        AddMenuItem "Erode...", "effects_erode", 7, 4, 6
    AddMenuItem "Natural", 7, 5
        AddMenuItem "Atmosphere...", "effects_atmosphere", 7, 5, 0
        AddMenuItem "Fog...", "effects_fog", 7, 5, 1
        AddMenuItem "Freeze", "effects_freeze", 7, 5, 2
        AddMenuItem "Ignite", "effects_ignite", 7, 5, 3
        AddMenuItem "Lava", "effects_lava", 7, 5, 4
        AddMenuItem "Metal...", "effects_metal", 7, 5, 5
        AddMenuItem "Underwater", "effects_underwater", 7, 5, 6
    AddMenuItem "Noise", "effects_noise", 7, 6
        AddMenuItem "Add film grain...", "effects_filmgrain", 7, 6, 0
        AddMenuItem "Add RGB noise...", "effects_rgbnoise", 7, 6, 1
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
    AddMenuItem "Custom filter...", "effects_customfilter", 7, 12
    AddMenuItem "Test (developers only)", "effects_developertest", 7, 13
    
    
    'Tools Menu
    AddMenuItem "Tools", "tools_top", 8
    AddMenuItem "Language", "tools_language", 8, 0, , "tools_language"
    AddMenuItem "Language editor...", "tools_languageeditor", 8, 1
    AddMenuItem "Record macro", "tools_macrotop", 8, 5, , "macro_record"
        AddMenuItem "Start recording", "tools_recordmacro", 8, 5, 0, "macro_record"
        AddMenuItem "Stop recording...", "tools_stopmacro", 8, 5, 1, "macro_stop"
    AddMenuItem "Play macro...", "tools_playmacro", 8, 6, , "macro_play"
    AddMenuItem "Recent macros", "tools_recentmacros", 8, 7
    AddMenuItem "Options...", "tools_options", 8, 9, , "pref_advanced"
    AddMenuItem "Plugin manager...", "tools_plugins", 8, 10, , "tools_plugin"
    
    
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
                IconsAndCursors.AddMenuIcon .ME_ResImage, .ME_TopMenu, .ME_SubMenu, .ME_SubSubMenu
            End With
        End If
    Next i

End Sub
