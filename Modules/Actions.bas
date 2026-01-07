Attribute VB_Name = "Actions"
'***************************************************************************
'Action Handler
'Copyright 2001-2026 by Tanner Helland
'Created: 07/October/21
'Last updated: 19/August/22
'Last update: build a rudimentary "action database" at run-time, which stores action names and corresponding
'             process IDs, and some attributes (like whether this action can be repeated or faded)
'
'Want to execute a program operation?  Call this module.
'
'Why does this module exist when PhotoDemon already has the Processor module (which seems to do the
' same thing)?  Well, they don't actually do the same thing.  PD's Processor module is a very low-level
' interface for executing program commands.  It has to manage a ton of special per-function details
' like branching for "show dialog" vs "execute action related to dialog".  It has to manage multiple
' varieties of Undo/Redo creation.  It has to record/play macro data.  It has to turn on and off a
' bunch of UI elements based on program state.
'
'But this module?  This module is for triggering just the default behavior for a named action.
' For adjustments and effects, that means displaying their dialog.  This module also launches actions
' with no direct processor equivalent, like "switch to tool [x]".
'
'PhotoDemon's menus, hotkeys, and search bar all rely on this module for their high-level redirection.
' Any new tool (including adjustments, effects, etc) needs to be accessible through this interface,
' so that users can do things like bind hotkeys to that action.
'
'As you might expect, this module relies heavily on the Menus module for correct behavior.
' (For example, actions that can be launched by a menu will query the menu's enabled state before
' launching.)  Check the Menus module for additional details.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'I am not generally in favor of public constants like this, but it's better than redeclaring the same
' constant across a dozen different files.  Because PD uses this module to forward centralized "commands"
' elsewhere in the project (e.g. hotkey commands are sent here first, for validation), it is helpful to
' tag some unique command IDs so that they can be reused elsewhere.
Public Const COMMAND_FILE_OPEN_RECENT As String = "file_open_recent_"
Public Const COMMAND_TOOLS_MACRO_RECENT As String = "tools_macro_recent_"

'PhotoDemon actions can be triggered by different places: menu clicks, hotkeys, or searches.  Some actions
' behave slightly differently depending on source.  (For example, "Paste to cursor" only works if the
' source is a hotkey; if it's a menu or search, a normal Paste action needs to be used instead, because we
' don't have a reliable cursor position.)  When calling the Action module, please pass the correct source
' so the router can handle any special source-related details.
Public Enum PD_ActionSource
    pdas_Hotkey
    pdas_Menu
    pdas_Search
End Enum

#If False Then
    Private Const pdas_Hotkey = 0, pdas_Menu = 0, pdas_Search = 0
#End If

'This table maps between internal action names and their associated Process module calls.  The plan
' (post 9.0 release) is to rework PD's central processor so that it operates directly on action names.
' This is an extremely messy overhaul, however, because every function in the program (literally *every*
' function) calls into the Process module, and all of those calls will need to be reworked to use
' action names instead.  PD's localization generator also relies on "Process [x]" commands to find
' action names that need to localized, so that's another concern.  (Localized action names are necessary
' for populating run-time commands like "Undo [x]", "Repeat [x]", etc.)
'
'Anyway, for the 9.0 release this table is used to map between action and processor names, and I've also
' decided to use it to collect generic action attributes like "should this action be reflected in the
' Edit > Repeat Last Action menu"
Private Type PD_Action
    actionName As String
    processName As String
    isRepeatable As Boolean
    isFadeable As Boolean
End Type

Private m_numActions As Long, m_Actions() As PD_Action

'Given a menu search string, apply the corresponding default processor action.
Public Function LaunchAction_BySearch(ByRef srcSearchText As String) As Boolean
    LaunchAction_BySearch = Actions.LaunchAction_ByName(Menus.GetNameFromSearchText(srcSearchText), pdas_Search)
End Function

'Given an action name, apply the corresponding default processor action.
'
'This function is referenced in many places throughout PD (e.g. the program's menus pretty much all reference this!)
' and it is distinct from PD's Processor module because it validates actions before executing them.  For example -
' if you request an operation associated with a menu, this function won't apply that action if the associated menu is
' disabled.  Similarly, if you request an operation that requires an open image, this function will ensure an image
' is open before actually applying that command.  PhotoDemon's central processor does not handle validation (but it
' handles a ton of other complex tasks, like Undo/Redo behavior) so for operations that need to be safely validated,
' call *this* function instead.
'
'Under very specific circumstances, you may want to skip the automatic action validation that this function provides.
' (For example, right-clicking an inactive layer in the layer toolbox may lead to initiating a "Merge Down" action.
' The "Merge Down" menu always validates against the *active* layer, not an arbitrary one, so the resulting command
' will always fail if the bottom layer in the image is the currently active one.)  If you are 100% certain that an
' action is valid, you can skip the automatic validation steps by passing skipValidation as TRUE.  Be very careful
' when doing this.
'
'Actions that appear in the Layer menu can also be applied to any arbitrary layer.  Just pass the target layer index
' via targetLayerIndex.  (If no index is passed, actions will default to the currently active layer.)
Public Function LaunchAction_ByName(ByRef srcMenuName As String, Optional ByVal actionSource As PD_ActionSource = pdas_Menu, Optional ByVal skipValidation As Boolean = False, Optional ByVal targetLayerIndex As Long = -1) As Boolean
    
    LaunchAction_ByName = False
    
    'Failsafe check for other actions already processing in the background
    If Processor.IsProgramBusy() Then Exit Function
    
    'Failsafe check to see if the menu associated with an action is enabled; if it isn't, that's an
    ' excellent surrogate for "do not allow this operation to proceed".  (Note that this is only
    ' useful for actions with a menu surrogate.  If an action doesn't have a menu surrogate, we ignore
    ' the return from this function.)
    Dim mnuDoesntExist As Boolean
    
    'The user *can* choose to skip validation (see comments at the top of this function).
    If (Not skipValidation) Then
        If (Not Menus.IsMenuEnabled(srcMenuName, mnuDoesntExist)) Then
            If (Not mnuDoesntExist) Then
                
                'Check for some known exceptions to this rule.  These are primarily convenience functions,
                ' which automatically remap to a similar task when the requested one isn't available.
                ' (For example, Ctrl+V is "Paste as new layer", but if no image is open, we silently remap
                ' to "Paste as new image".)
                If (Not Strings.StringsEqualAny(srcMenuName, True, "edit_pasteaslayer")) Then
                    Exit Function
                End If
                
            End If
        End If
    End If
    
    'Helper functions exist for each main menu.  Once a command is located, we can stop searching.
    Dim cmdFound As Boolean: cmdFound = False
    
    'Before searching menu items, perform a "quick" search for UI-specific tool actions
    If (Not cmdFound) Then cmdFound = Launch_ByName_NonMenu(srcMenuName, actionSource)
    
    'Search each menu group in turn
    If (Not cmdFound) Then cmdFound = Launch_ByName_MenuFile(srcMenuName, actionSource)
    If (Not cmdFound) Then cmdFound = Launch_ByName_MenuEdit(srcMenuName, actionSource)
    If (Not cmdFound) Then cmdFound = Launch_ByName_MenuImage(srcMenuName, actionSource)
    If (Not cmdFound) Then cmdFound = Launch_ByName_MenuLayer(srcMenuName, actionSource, targetLayerIndex)
    If (Not cmdFound) Then cmdFound = Launch_ByName_MenuSelect(srcMenuName, actionSource)
    If (Not cmdFound) Then cmdFound = Launch_ByName_MenuAdjustments(srcMenuName, actionSource)
    If (Not cmdFound) Then cmdFound = Launch_ByName_MenuEffects(srcMenuName, actionSource)
    If (Not cmdFound) Then cmdFound = Launch_ByName_MenuTools(srcMenuName, actionSource)
    If (Not cmdFound) Then cmdFound = Launch_ByName_MenuView(srcMenuName, actionSource)
    If (Not cmdFound) Then cmdFound = Launch_ByName_MenuWindow(srcMenuName, actionSource)
    If (Not cmdFound) Then cmdFound = Launch_ByName_MenuHelp(srcMenuName, actionSource)
    If (Not cmdFound) Then cmdFound = Launch_ByName_Misc(srcMenuName, actionSource)
    
    LaunchAction_ByName = cmdFound
    
    'Before exiting, report a debug note if we found *no* matches.
    '
    'NOTE 2021: this can be useful when adding a new feature to the program (to make sure all triggers for it
    ' execute correctly), but it is *not* useful in day-to-day usage because the menu searcher may find a command,
    ' but choose not to execute it because certain safety conditions aren't met (e.g. Ctrl+S is pressed, but no
    ' image is open).  Many of these validation checks occur at the top of a group of related commands -
    ' e.g. nothing in the Effects category will trigger without an open image - and some of those validation
    ' checks will prevent menu-matching from even occurring.  This will report a "no match found", but only
    ' because large chunks of the search were short-circuited because a validation condition wasn't met.
    'If (Not cmdFound) Then PDDebug.LogAction "WARNING: Actions.LaunchAction_ByName received an unknown request: " & srcMenuName
    
End Function

Private Function Launch_ByName_MenuFile(ByRef srcMenuName As String, Optional ByVal actionSource As PD_ActionSource = pdas_Menu) As Boolean

    Dim cmdFound As Boolean: cmdFound = True
    
    Select Case srcMenuName
    
        Case "file_new"
            Process "New image", True
            
        Case "file_open"
            Process "Open", True
            
        Case "file_openrecent"
            'Top-level menu only; see the end of this function for handling actual recent file actions.
            ' (Note that the search bar does present this term, and if clicked, we will simply load the
            ' *top* item in the Recent Files list.)
            If (actionSource = pdas_Search) Or (actionSource = pdas_Hotkey) Then
                If (LenB(g_RecentFiles.GetFullPath(0)) <> 0) Then Loading.LoadFileAsNewImage g_RecentFiles.GetFullPath(0)
            End If
            
            Case "file_open_allrecent"
                Loading.LoadAllRecentFiles
            
            Case "file_open_clearrecent"
                If (Not g_RecentFiles Is Nothing) Then g_RecentFiles.ClearList
            
        Case "file_import"
            Case "file_import_paste"
                Process "Paste to new image", False, , UNDO_Nothing, , False
                
            Case "file_import_scanner"
                Process "Scan image", True
                
            Case "file_import_selectscanner"
                Process "Select scanner or camera", True
                
            Case "file_import_web"
                Process "Internet import", True
                
            Case "file_import_screenshot"
                Process "Screen capture", True
                
        Case "file_close"
            If (Not PDImages.IsImageActive()) Then Exit Function
            Process "Close", True
            
        Case "file_closeall"
            If (Not PDImages.IsImageActive()) Then Exit Function
            Process "Close all", True
            
        Case "file_save"
            If (Not PDImages.IsImageActive()) Then Exit Function
            Process "Save", True
            
        Case "file_savecopy"
            If (Not PDImages.IsImageActive()) Then Exit Function
            Process "Save copy", True
            
        Case "file_saveas"
            If (Not PDImages.IsImageActive()) Then Exit Function
            Process "Save as", True
            
        Case "file_revert"
            If (Not PDImages.IsImageActive()) Then Exit Function
            Process "Revert", False, , UNDO_Everything
            
        Case "file_export"
            Case "file_export_image"
                If (Not PDImages.IsImageActive()) Then Exit Function
                Process "Export image", True
            
            Case "file_export_layers"
                If (Not PDImages.IsImageActive()) Then Exit Function
                Process "Export layers", True
                
            Case "file_export_animation"
                If (Not PDImages.IsImageActive()) Then Exit Function
                Process "Export animation", True
                
            Case "file_export_colorlookup"
                If (Not PDImages.IsImageActive()) Then Exit Function
                Process "Export color lookup", True
                
            Case "file_export_colorprofile"
                If (Not PDImages.IsImageActive()) Then Exit Function
                Process "Export color profile", True
                
            Case "file_export_palette"
                If (Not PDImages.IsImageActive()) Then Exit Function
                Process "Export palette", True
                
        Case "file_batch"
            Case "file_batch_process"
                Process "Batch wizard", True
                
            Case "file_batch_repair"
                ShowPDDialog vbModal, FormBatchRepair
                
        Case "file_print"
            If (Not PDImages.IsImageActive()) Then Exit Function
            Process "Print", True
            
        Case "file_quit"
            Process "Exit program", True
            
        Case Else
            cmdFound = False
        
    End Select
    
    'If we haven't found a match, look for commands related to the Recent Files menu;
    ' these are preceded by the unique "file_open_recent_[n]" command, where [n] is the index of
    ' the recent file to open (0-based).
    If (Not cmdFound) Then
    
        cmdFound = Strings.StringsEqualLeft(srcMenuName, COMMAND_FILE_OPEN_RECENT, True)
        If cmdFound Then
        
            '(Attempt to) load the target file
            Dim targetIndex As Long
            targetIndex = Val(Right$(srcMenuName, Len(srcMenuName) - Len(COMMAND_FILE_OPEN_RECENT)))
            If (LenB(g_RecentFiles.GetFullPath(targetIndex)) <> 0) Then Loading.LoadFileAsNewImage g_RecentFiles.GetFullPath(targetIndex)
            
        End If
        
    End If
    
    Launch_ByName_MenuFile = cmdFound
    
End Function

Private Function Launch_ByName_MenuEdit(ByRef srcMenuName As String, Optional ByVal actionSource As PD_ActionSource = pdas_Menu) As Boolean
    
    '*Almost* all actions in this menu require an open image.  The few outliers that do not can be
    ' checked here, in advance.
    If (Not PDImages.IsImageActive()) Then
        
        'Note that "edit_pasteaslayer" is a weird exception here, as PD's processor will silently forward it
        ' to "edit_pasteasimage" if no images are open.  (This simplifies use of Ctrl+V by beginners.)
        If (Not Strings.StringsEqualAny(srcMenuName, True, "edit_pasteaslayer", "edit_pasteasimage", "edit_specialpaste", "edit_emptyclipboard")) Then
            Exit Function
        End If
        
    End If
    
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
        
        'If a selection is active, the Undo/Redo engine can simply back up the current layer contents.
        ' If, however, no selection is active, we will delete the entire layer.  That requires a backup
        ' of the full layer stack.
        Case "edit_cutlayer"
            If PDImages.GetActiveImage.IsSelectionActive Then
                Process "Cut", False, , UNDO_Layer
            Else
                Process "Cut", False, , UNDO_Image
            End If
        
        Case "edit_cutmerged"
            Process "Cut merged", False, , UNDO_Image
            
        Case "edit_copylayer"
            Process "Copy", False, , UNDO_Nothing
            
        Case "edit_copymerged"
            Process "Copy merged", False, , UNDO_Nothing
            
        Case "edit_pasteaslayer"
            If PDImages.IsImageActive Then
                Process "Paste", False, , UNDO_Image_VectorSafe
            Else
                Process "Paste to new image", False, , UNDO_Nothing, , False
            End If
            
        Case "edit_pastetocursor"
            If (actionSource = pdas_Hotkey) Then
                Process "Paste to cursor", False, BuildParamList("canvas-mouse-x", FormMain.MainCanvas(0).GetLastMouseX(), "canvas-mouse-y", FormMain.MainCanvas(0).GetLastMouseY()), UNDO_Image_VectorSafe
            Else
                Process "Paste", False, , UNDO_Image_VectorSafe
            End If
            
        Case "edit_pasteasimage"
            Process "Paste to new image", False, , UNDO_Nothing, , False
            
        'The cut/copy/paste special menus allow the user to specify the format used for cut/copy/paste
        Case "edit_specialcut"
            Process "Cut special", True
        
        Case "edit_specialcopy"
            Process "Copy special", True
        
        Case "edit_specialpaste"
            Process "Paste special", True
        
        'Empty clipboard is always available
        Case "edit_emptyclipboard"
            Process "Empty clipboard", False, vbNullString, UNDO_Nothing, recordAction:=False
        
        Case "edit_clear"
            Process "Clear", True
            
        Case "edit_contentawarefill"
            Process "Content-aware fill", True
            
        Case "edit_fill"
            Process "Fill", True
            
        Case "edit_stroke"
            Process "Stroke", True
        
        Case Else
            cmdFound = False
            
    End Select
    
    Launch_ByName_MenuEdit = cmdFound
    
End Function

Private Function Launch_ByName_MenuImage(ByRef srcMenuName As String, Optional ByVal actionSource As PD_ActionSource = pdas_Menu) As Boolean
    
    'All actions in this category require an open image.  If no images are open, do not apply the requested action.
    If (Not PDImages.IsImageActive()) Then Exit Function
    
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
            Process "Fit canvas to active layer", False, BuildParamList("targetlayer", PDImages.GetActiveImage.GetActiveLayerIndex), UNDO_ImageHeader
            
        Case "image_fitalllayers"
            Process "Fit canvas around all layers", False, , UNDO_ImageHeader
            
        Case "image_crop"
            Process "Crop", True
            
        Case "image_trim"
            Process "Trim empty image borders", , , UNDO_ImageHeader
            
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
            
        Case "image_mergevisible"
            Process "Merge visible layers", , , UNDO_Image
            
        Case "image_flatten"
            Process "Flatten image", True
        
        Case "image_animation"
            Process "Animation options", True
        
        Case "image_compare"
            Case "image_createlut"
                Process "Create color lookup", True
            
            Case "image_similarity"
                Process "Compare similarity", True
        
        Case "image_metadata"
            Case "image_editmetadata"
                Process "Edit metadata", True
                
            Case "image_removemetadata"
                Process "Remove all metadata", False, , UNDO_ImageHeader
                
            Case "image_countcolors"
                Process "Count unique colors", True
                
            Case "image_maplocation"
                Web.MapImageLocation
        
        Case "image_showinexplorer"
            Interface.ShowActiveImageFileInExplorer
            
        Case Else
            cmdFound = False
                
    End Select
    
    Launch_ByName_MenuImage = cmdFound
    
End Function

'To specify a target layer for a given action, set targetLayerIndex to a value >= 0.  This index *is* validated,
' and will be set to the active layer if an invalid index is passed.
Private Function Launch_ByName_MenuLayer(ByRef srcMenuName As String, Optional ByVal actionSource As PD_ActionSource = pdas_Menu, Optional ByVal targetLayerIndex As Long = -1) As Boolean

    'All actions in this category require an open image.  If no images are open, do not apply the requested action.
    If (Not PDImages.IsImageActive()) Then Exit Function
    
    If (targetLayerIndex < 0) Then
        targetLayerIndex = PDImages.GetActiveImage.GetActiveLayerIndex
    ElseIf (targetLayerIndex >= PDImages.GetActiveImage.GetNumOfLayers) Then
        targetLayerIndex = PDImages.GetActiveImage.GetActiveLayerIndex
    End If
    
    Dim cmdFound As Boolean: cmdFound = True
    
    Select Case srcMenuName
    
        Case "layer_add"
            Case "layer_addbasic"
                Process "Add new layer", True
                
            Case "layer_addblank"
                Process "Add blank layer", False, BuildParamList("targetlayer", targetLayerIndex), UNDO_Image_VectorSafe
                
            Case "layer_duplicate"
                Process "Duplicate Layer", False, BuildParamList("targetlayer", targetLayerIndex), UNDO_Image_VectorSafe
                
            Case "layer_addfromclipboard"
                Process "Paste", False, , UNDO_Image_VectorSafe
                
            Case "layer_addfromfile"
                Process "New layer from file", True
                
            Case "layer_addfromvisiblelayers"
                Process "New layer from visible layers", False, , UNDO_Image_VectorSafe
                
            Case "layer_addviacopy"
                Process "Layer via copy", False, BuildParamList("targetlayer", targetLayerIndex), UNDO_Image_VectorSafe
                
            Case "layer_addviacut"
                Process "Layer via cut", False, BuildParamList("targetlayer", targetLayerIndex), UNDO_Image
                
        Case "layer_delete"
            Case "layer_deletecurrent"
                Process "Delete layer", False, BuildParamList("layerindex", targetLayerIndex), UNDO_Image_VectorSafe
                
            Case "layer_deletehidden"
                Process "Delete hidden layers", False, , UNDO_Image_VectorSafe
        
        Case "layer_replace"
            Case "layer_replacefromclipboard"
                Process "Replace layer from clipboard", False, createUndo:=UNDO_Layer
                
            Case "layer_replacefromfile"
                Process "Replace layer from file", True
                
            Case "layer_replacefromvisiblelayers"
                Process "Replace layer from visible layers", False, createUndo:=UNDO_Layer
                
        Case "layer_mergeup"
            Process "Merge layer up", False, BuildParamList("layerindex", targetLayerIndex), UNDO_Image
            
        Case "layer_mergedown"
            Process "Merge layer down", False, BuildParamList("layerindex", targetLayerIndex), UNDO_Image
            
        Case "layer_order"
            Case "layer_gotop"
                Process "Go to top layer", False, vbNullString, UNDO_Nothing
                
            Case "layer_goup"
                Process "Go to layer above", False, vbNullString, UNDO_Nothing
                
            Case "layer_godown"
                Process "Go to layer below", False, vbNullString, UNDO_Nothing
                
            Case "layer_gobottom"
                Process "Go to bottom layer", False, vbNullString, UNDO_Nothing
            
            Case "layer_movetop"
                Process "Raise layer to top", False, BuildParamList("layerindex", targetLayerIndex), UNDO_ImageHeader
                
            Case "layer_moveup"
                Process "Raise layer", False, BuildParamList("layerindex", targetLayerIndex), UNDO_ImageHeader
                
            Case "layer_movedown"
                Process "Lower layer", False, BuildParamList("layerindex", targetLayerIndex), UNDO_ImageHeader
                
            Case "layer_movebottom"
                Process "Lower layer to bottom", False, BuildParamList("layerindex", targetLayerIndex), UNDO_ImageHeader
            
            Case "layer_reverse"
                Process "Reverse layer order", False, vbNullString, UNDO_Image
        
        Case "layer_visibility"
            Case "layer_show"
                Process "Toggle layer visibility", False, BuildParamList("layerindex", targetLayerIndex), UNDO_LayerHeader
                
            Case "layer_showonly"
                Process "Show only this layer", False, BuildParamList("layerindex", targetLayerIndex), UNDO_ImageHeader
                
            Case "layer_hideonly"
                Process "Hide only this layer", False, BuildParamList("layerindex", targetLayerIndex), UNDO_ImageHeader
                
            Case "layer_showall"
                Process "Show all layers", False, vbNullString, UNDO_ImageHeader
                
            Case "layer_hideall"
                Process "Hide all layers", False, vbNullString, UNDO_ImageHeader
        
        Case "layer_crop"
            Case "layer_cropselection"
                Process "Crop layer to selection", , , UNDO_Layer
            
            Case "layer_pad"
                Process "Pad layer to image size", , , UNDO_Layer
                
            Case "layer_trim"
                Process "Trim empty layer borders", , , UNDO_Layer
            
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
                
        Case "layer_size"
            Case "layer_resetsize"
                Process "Reset layer size", False, BuildParamList("layerindex", targetLayerIndex), UNDO_LayerHeader
                
            Case "layer_resize"
                Process "Resize layer", True
                
            Case "layer_contentawareresize"
                Process "Content-aware layer resize", True
                
            Case "layer_fittoimage"
                Process "Fit layer to image", False, BuildParamList("layerindex", targetLayerIndex), UNDO_LayerHeader
                
        Case "layer_transparency"
            Case "layer_colortoalpha"
                Process "Color to alpha", True
                
            Case "layer_luminancetoalpha"
                Process "Luminance to alpha", True
                
            Case "layer_removealpha"
                Process "Remove alpha channel", True
            
            Case "layer_thresholdalpha"
                Process "Threshold alpha", True
        
        Case "layer_rasterize"
            Case "layer_rasterizecurrent"
                Process "Rasterize layer", False, BuildParamList("layerindex", targetLayerIndex), UNDO_Layer
                
            Case "layer_rasterizeall"
                Process "Rasterize all layers", , , UNDO_Image
        
        Case "layer_split"
            Case "layer_splitlayertoimage"
                Process "Split layer into image", True
                
            Case "layer_splitalllayerstoimages"
                Process "Split layers into images", True
            
            Case "layer_splitimagestolayers"
                Process "Split images into layers", True
                
        Case Else
            cmdFound = False
            
    End Select
    
    Launch_ByName_MenuLayer = cmdFound
    
End Function

Private Function Launch_ByName_MenuSelect(ByRef srcMenuName As String, Optional ByVal actionSource As PD_ActionSource = pdas_Menu) As Boolean

    'All actions in this category require an open image.  If no images are open, do not apply the requested action.
    If (Not PDImages.IsImageActive()) Then Exit Function
    
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
            
        Case "select_fill"
            Process "Fill selected area", True
            
        Case "select_heal"
            Process "Heal selected area", True
            
        Case "select_stroke"
            Process "Stroke selection outline", True
        
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
    
    Launch_ByName_MenuSelect = cmdFound
    
End Function

Private Function Launch_ByName_MenuAdjustments(ByRef srcMenuName As String, Optional ByVal actionSource As PD_ActionSource = pdas_Menu) As Boolean

    'All actions in this category require an open image.  If no images are open, do not apply the requested action.
    If (Not PDImages.IsImageActive()) Then Exit Function
    
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
            Process "Shadows and highlights", True
            
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
            
            Case "adj_colorlookup"
                Process "Color lookup", True
                
            Case "adj_colorize"
                Process "Colorize", True
                
            Case "adj_photofilters"
                Process "Photo filter", True
                
            Case "adj_replacecolor"
                Process "Replace color", True
                
            Case "adj_sepia"
                Process "Sepia", True
                
            Case "adj_splittone"
                Process "Split toning", True
                
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
            
            Case "adj_dehaze"
                Process "Dehaze", True
            
            Case "adj_exposure"
                Process "Exposure", True
            
            Case "adj_gamma"
                Process "Gamma", True
                
            Case "adj_hdr"
                Process "HDR", True
                
            'Case "adj_levels"  'Covered by parent menu
            'Case "adj_sandh"   'Covered by parent menu
            
        Case "adj_map"
            Case "adj_gradientmap"
                Process "Gradient map", True
                
            Case "adj_palettemap"
                Process "Palette map", True
            
        Case "adj_monochrome"
            Case "adj_colortomonochrome"
                Process "Color to monochrome", True
                
            Case "adj_monochrometogray"
                Process "Monochrome to gray", True
            
        Case Else
            cmdFound = False
                
    End Select
    
    Launch_ByName_MenuAdjustments = cmdFound
    
End Function

Private Function Launch_ByName_MenuEffects(ByRef srcMenuName As String, Optional ByVal actionSource As PD_ActionSource = pdas_Menu) As Boolean

    'All actions in this category require an open image.  If no images are open, do not apply the requested action.
    If (Not PDImages.IsImageActive()) Then Exit Function
    
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
                
        Case "effects_distort"
            Case "effects_fixlensdistort"
                Process "Correct lens distortion", True
                
            Case "effects_donut"
                Process "Donut", True
            
            Case "effects_droste"
                Process "Droste", True
                
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
                
            Case "effects_gradientflow"
                Process "Gradient flow", True
                
            Case "effects_rangefilter"
                Process "Range filter", True
                
            Case "effects_tracecontour"
                Process "Trace contour", True
                
        Case "effects_lightandshadow"
            Case "effects_blacklight"
                Process "Black light", True
                
            Case "effects_bumpmap"
                Process "Bump map", True
                
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
            
            'For legacy macros, only; bilateral has been replaced by Blur > Surface Blur
            Case "effects_bilateral"
                Process "Surface blur", True
                
            Case "effects_dustandscratches"
                Process "Dust and scratches", True
                
            Case "effects_harmonicmean"
                Process "Harmonic mean", True
                
            Case "effects_meanshift"
                Process "Mean shift", True
                
            Case "effects_median"
                Process "Median", True
            
            Case "effects_snn"
                Process "Symmetric nearest-neighbor", True
                
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
                
            Case "effects_pointillize"
                Process "Pointillize", True
        
        Case "effects_render"
            Case "effects_clouds"
                Process "Clouds", True
                
            Case "effects_fibers"
                Process "Fibers", True
            
            Case "effects_truchet"
                Process "Truchet", True
            
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
            
            Case "effects_kuwahara"
                Process "Kuwahara filter", True
                
            Case "effects_outline"
                Process "Outline", True
                
            Case "effects_palette"
                Process "Palette", True
                
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
                Process "Offset and zoom", True
                
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
                
        Case "effects_animation"
            Case "effects_animation_background"
                Process "Animation background", True
                
            Case "effects_animation_foreground"
                Process "Animation foreground", True
            
            Case "effects_animation_speed"
                Process "Animation playback speed", True
                
        Case "effects_customfilter"
            Process "Custom filter", True
        
        Case "effects_8bf"
            Process "Photoshop (8bf) plugin", True
        
        Case Else
            cmdFound = False
            
    End Select
    
    Launch_ByName_MenuEffects = cmdFound
    
End Function

Private Function Launch_ByName_MenuTools(ByRef srcMenuName As String, Optional ByVal actionSource As PD_ActionSource = pdas_Menu) As Boolean

    Dim cmdFound As Boolean: cmdFound = True
    
    Select Case srcMenuName
    
        Case "tools_language"
        
        Case "tools_languageeditor"
            If (Not FormLanguageEditor.Visible) Then
                FormMain.HotkeyManager.Enabled = False
                ShowPDDialog vbModal, FormLanguageEditor
                FormMain.HotkeyManager.Enabled = True
            End If
            
        Case "tools_theme"
            Dialogs.PromptUITheme
            
        Case "tools_macrocreatetop"
            Case "tools_macrofromhistory"
                If (Not PDImages.IsImageActive()) Then Exit Function
                ShowPDDialog vbModal, FormMacroSession
                
            Case "tools_recordmacro"
                If (Not PDImages.IsImageActive()) Then Exit Function
                Process "Start macro recording", , , UNDO_Nothing
                
            Case "tools_stopmacro"
                If (Not PDImages.IsImageActive()) Then Exit Function
                Process "Stop macro recording", True
                
        Case "tools_playmacro"
            If (Not PDImages.IsImageActive()) Then Exit Function
            Process "Play macro", True
            
        Case "tools_recentmacros"
        
        Case "tools_screenrecord"
            ShowPDDialog vbModal, FormScreenVideoPrefs
        
        Case "tools_hotkeys"
            FormMain.HotkeyManager.Enabled = False
            ShowPDDialog vbModal, FormHotkeys
            FormMain.HotkeyManager.Enabled = True
            
        Case "tools_options"
            ShowPDDialog vbModal, FormOptions
            
        Case "tools_developers"
            Case "tools_viewdebuglog"
                If (UserPrefs.GenerateDebugLogs() And (LenB(PDDebug.GetDebugLogFilename()) <> 0)) Then Web.OpenURL PDDebug.GetDebugLogFilename()
            
            Case "tools_themeeditor"
                ShowPDDialog vbModal, FormThemeEditor
                
            Case "tools_themepackage"
                g_Themer.BuildThemePackage
                
            Case "tools_standalonepackage"
                ShowPDDialog vbModal, FormPackage
                
        Case "effects_developertest"
            'Handled directly in FormMain (for legacy reasons)
        
        Case Else
            cmdFound = False
        
    End Select
    
    'If we haven't found a match, look for commands related to the Recent Macros menu;
    ' these are preceded by the unique "tools_macro_recent_[n]" command, where [n] is the index of
    ' the recent macro to open (0-based).
    If (Not cmdFound) And PDImages.IsImageActive() Then
    
        cmdFound = Strings.StringsEqualLeft(srcMenuName, COMMAND_TOOLS_MACRO_RECENT, True)
        If cmdFound Then
        
            '(Attempt to) play the target macro
            Dim targetIndex As Long
            targetIndex = Val(Right$(srcMenuName, Len(srcMenuName) - Len(COMMAND_TOOLS_MACRO_RECENT)))
            If (LenB(g_RecentMacros.GetSpecificMRU(targetIndex)) <> 0) Then Macros.PlayMacroFromFile g_RecentMacros.GetSpecificMRU(targetIndex)
        
        End If
        
    End If
    
    Launch_ByName_MenuTools = cmdFound
    
End Function

Private Function Launch_ByName_MenuView(ByRef srcMenuName As String, Optional ByVal actionSource As PD_ActionSource = pdas_Menu) As Boolean

    'All actions in this category require an open image.  If no images are open, do not apply the requested action.
    If (Not PDImages.IsImageActive()) Then Exit Function
    
    Dim cmdFound As Boolean: cmdFound = True
    Dim newState As Boolean
    
    Select Case srcMenuName
    
        Case "view_fit"
            CanvasManager.FitOnScreen
        
        Case "view_center_on_screen"
            CanvasManager.CenterOnScreen
        
        Case "view_zoomin"
            If FormMain.MainCanvas(0).IsZoomEnabled Then
                If (FormMain.MainCanvas(0).GetZoomDropDownIndex > 0) Then FormMain.MainCanvas(0).SetZoomDropDownIndex Zoom.GetNearestZoomInIndex(FormMain.MainCanvas(0).GetZoomDropDownIndex)
            End If
            
        Case "view_zoomout"
            If FormMain.MainCanvas(0).IsZoomEnabled Then
                If (FormMain.MainCanvas(0).GetZoomDropDownIndex <> Zoom.GetZoomCount) Then FormMain.MainCanvas(0).SetZoomDropDownIndex Zoom.GetNearestZoomOutIndex(FormMain.MainCanvas(0).GetZoomDropDownIndex)
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
                If FormMain.MainCanvas(0).IsZoomEnabled Then FormMain.MainCanvas(0).SetZoomDropDownIndex Zoom.GetZoom100Index
                
            Case "zoom_1_2"
                If FormMain.MainCanvas(0).IsZoomEnabled Then FormMain.MainCanvas(0).SetZoomDropDownIndex Zoom.GetZoom100Index() + 3
                
            Case "zoom_1_4"
                If FormMain.MainCanvas(0).IsZoomEnabled Then FormMain.MainCanvas(0).SetZoomDropDownIndex Zoom.GetZoom100Index() + 5
                
            Case "zoom_1_8"
                If FormMain.MainCanvas(0).IsZoomEnabled Then FormMain.MainCanvas(0).SetZoomDropDownIndex Zoom.GetZoom100Index() + 8
                
            Case "zoom_1_16"
                If FormMain.MainCanvas(0).IsZoomEnabled Then FormMain.MainCanvas(0).SetZoomDropDownIndex Zoom.GetZoom100Index() + 10
                
        Case "view_rulers"
            newState = Not FormMain.MainCanvas(0).GetRulerVisibility()
            Menus.SetMenuChecked "view_rulers", newState
            FormMain.MainCanvas(0).SetRulerVisibility newState
            
        Case "view_statusbar"
            newState = Not FormMain.MainCanvas(0).GetStatusBarVisibility()
            Menus.SetMenuChecked "view_statusbar", newState
            FormMain.MainCanvas(0).SetStatusBarVisibility newState
        
        Case "show_layeredges"
            Drawing.ToggleShowOptions pdst_LayerEdges
            Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage, FormMain.MainCanvas(0)
            
        Case "show_smartguides"
            Drawing.ToggleShowOptions pdst_SmartGuides
            
        Case "snap_global"
            Snap.ToggleSnapOptions pdst_Global
            
        Case "snap_canvasedge"
            Snap.ToggleSnapOptions pdst_CanvasEdge
            
        Case "snap_centerline"
            Snap.ToggleSnapOptions pdst_Centerline
            
        Case "snap_layer"
            Snap.ToggleSnapOptions pdst_Layer
            
        Case "snap_angle_90"
            Snap.ToggleSnapOptions pdst_Angle90
            
        Case "snap_angle_45"
            Snap.ToggleSnapOptions pdst_Angle45
            
        Case "snap_angle_30"
            Snap.ToggleSnapOptions pdst_Angle30
            
        Case Else
            cmdFound = False
        
    End Select
    
    Launch_ByName_MenuView = cmdFound
    
End Function

Private Function Launch_ByName_MenuWindow(ByRef srcMenuName As String, Optional ByVal actionSource As PD_ActionSource = pdas_Menu) As Boolean

    Dim cmdFound As Boolean: cmdFound = True
    
    Select Case srcMenuName
    
        Case "window_toolbox"
            Case "window_displaytoolbox"
                Toolboxes.ToggleToolboxVisibility PDT_LeftToolbox
                
            Case "window_displaytoolcategories"
                toolbar_Toolbox.ToggleToolCategoryLabels
                
            Case "window_smalltoolbuttons"
                toolbar_Toolbox.UpdateButtonSize tbs_Small
                
            Case "window_mediumtoolbuttons"
                toolbar_Toolbox.UpdateButtonSize tbs_Medium
                
            Case "window_largetoolbuttons"
                toolbar_Toolbox.UpdateButtonSize tbs_Large
                
        Case "window_tooloptions"
            Toolboxes.ToggleToolboxVisibility PDT_TopToolbox
            
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
    
    Launch_ByName_MenuWindow = cmdFound
    
End Function

Private Function Launch_ByName_MenuHelp(ByRef srcMenuName As String, Optional ByVal actionSource As PD_ActionSource = pdas_Menu) As Boolean

    Dim cmdFound As Boolean: cmdFound = True
    
    Select Case srcMenuName
    
        Case "help_patreon"
            Web.OpenURL "https://www.patreon.com/photodemon/overview"
            
        Case "help_donate"
            Web.OpenURL "https://photodemon.org/donate"
            
        Case "help_forum"
            Web.OpenURL "https://github.com/tannerhelland/PhotoDemon/discussions"
            
        Case "help_checkupdates"
            
            'Initiate an asynchronous download of the standard PD update file (currently hosted @ GitHub).
            ' When the asynchronous download completes, the downloader will place the completed update file in the /Data/Updates subfolder.
            ' On exit (or subsequent program runs), PD will check for the presence of that file, then proceed accordingly.
            Message "Checking for software updates..."
            FormMain.RequestAsynchronousDownload "PROGRAM_UPDATE_CHECK_USER", "https://tannerhelland.github.io/PhotoDemon-Updates-v2/", , vbAsyncReadForceUpdate, UserPrefs.GetUpdatePath & "updates.xml"
            
        Case "help_reportbug"
            Web.OpenURL "https://github.com/tannerhelland/PhotoDemon/issues/new/choose"
            
        Case "help_license"
            Web.OpenURL "https://photodemon.org/license/"
            
        Case "help_sourcecode"
            Web.OpenURL "https://github.com/tannerhelland/PhotoDemon"
            
        Case "help_website"
            Web.OpenURL "https://photodemon.org"
        
        Case "help_3rdpartylibs"
            ShowPDDialog vbModal, FormPluginManager
            
        Case "help_about"
            ShowPDDialog vbModal, FormAbout
            
        Case Else
            cmdFound = False
        
    End Select
    
    Launch_ByName_MenuHelp = cmdFound
    
End Function

Private Function Launch_ByName_NonMenu(ByRef srcMenuName As String, Optional ByVal actionSource As PD_ActionSource = pdas_Menu) As Boolean
    
    Dim cmdFound As Boolean: cmdFound = True
    
    Select Case srcMenuName
        
        'Give priority to "quick actions" related to tools - typically UI-specific operations that
        ' don't involve Undo/Redo.
        '
        '(Note also that these settings are fairly involved because they need to work on *any* relevant tool,
        ' so we typically relay their commands elsewhere.
        Case "tool_active_hardnessdown"
            Tools.QuickToolAction_HardnessDown
        Case "tool_active_hardnessup"
            Tools.QuickToolAction_HardnessUp
        Case "tool_active_sizedown"
            Tools.QuickToolAction_SizeDown
        Case "tool_active_sizeup"
            Tools.QuickToolAction_SizeUp
        Case "tool_active_togglecursor"
            Viewport.Stage4_FlipBufferAndDrawUI PDImages.GetActiveImage, FormMain.MainCanvas(0)
        
        'Activate various tools
        Case "tool_hand"
            toolbar_Toolbox.SelectNewTool NAV_DRAG, (actionSource = pdas_Search), True
        
        Case "tool_zoom"
            toolbar_Toolbox.SelectNewTool NAV_ZOOM, (actionSource = pdas_Search), True
        
        Case "tool_move"
            toolbar_Toolbox.SelectNewTool NAV_MOVE, (actionSource = pdas_Search), True
        
        'When using hotkeys to activate a tool, we use a slightly different strategy.  Some hotkeys are double-assigned
        ' to neighboring tools.  If one of the tools that share a hotkey already has focus, pressing that hotkey will
        ' toggle focus to the other tool in that group.
        Case "tool_colorselect"
            If (actionSource = pdas_Hotkey) Then
                If (g_CurrentTool = COLOR_PICKER) Then toolbar_Toolbox.SelectNewTool ND_MEASURE Else toolbar_Toolbox.SelectNewTool COLOR_PICKER
            Else
                toolbar_Toolbox.SelectNewTool COLOR_PICKER, (actionSource = pdas_Search), True
            End If
        
        Case "tool_measure"
            toolbar_Toolbox.SelectNewTool ND_MEASURE, (actionSource = pdas_Search), True
        
        Case "tool_crop"
            toolbar_Toolbox.SelectNewTool ND_CROP, (actionSource = pdas_Search), True
        
        Case "tool_select_rect"
            If (actionSource = pdas_Hotkey) Then
                If (g_CurrentTool = SELECT_RECT) Then toolbar_Toolbox.SelectNewTool SELECT_CIRC Else toolbar_Toolbox.SelectNewTool SELECT_RECT
            Else
                toolbar_Toolbox.SelectNewTool SELECT_RECT, (actionSource = pdas_Search), True
            End If
        
        Case "tool_select_ellipse"
            toolbar_Toolbox.SelectNewTool SELECT_CIRC, (actionSource = pdas_Search), True
        
        Case "tool_select_polygon"
            toolbar_Toolbox.SelectNewTool SELECT_POLYGON, (actionSource = pdas_Search), True
        
        Case "tool_select_lasso"
            If (actionSource = pdas_Hotkey) Then
                If (g_CurrentTool = SELECT_LASSO) Then toolbar_Toolbox.SelectNewTool SELECT_POLYGON Else toolbar_Toolbox.SelectNewTool SELECT_LASSO
            Else
                toolbar_Toolbox.SelectNewTool SELECT_LASSO, (actionSource = pdas_Search), True
            End If
        
        Case "tool_select_wand"
            toolbar_Toolbox.SelectNewTool SELECT_WAND, (actionSource = pdas_Search), True
        
        Case "tool_text_basic"
            If (actionSource = pdas_Hotkey) Then
                If (g_CurrentTool = TEXT_BASIC) Then toolbar_Toolbox.SelectNewTool TEXT_ADVANCED Else toolbar_Toolbox.SelectNewTool TEXT_BASIC
            Else
                toolbar_Toolbox.SelectNewTool TEXT_BASIC, (actionSource = pdas_Search), True
            End If
        
        Case "tool_text_advanced"
            toolbar_Toolbox.SelectNewTool TEXT_ADVANCED, (actionSource = pdas_Search), True
        
        Case "tool_pencil"
            toolbar_Toolbox.SelectNewTool PAINT_PENCIL, (actionSource = pdas_Search), True
        
        Case "tool_paintbrush"
            toolbar_Toolbox.SelectNewTool PAINT_SOFTBRUSH, (actionSource = pdas_Search), True
        
        Case "tool_erase"
            toolbar_Toolbox.SelectNewTool PAINT_ERASER, (actionSource = pdas_Search), True
        
        Case "tool_clone"
            toolbar_Toolbox.SelectNewTool PAINT_CLONE, (actionSource = pdas_Search), True
        
        Case "tool_paintbucket"
            toolbar_Toolbox.SelectNewTool PAINT_FILL, (actionSource = pdas_Search), True
        
        Case "tool_gradient"
            toolbar_Toolbox.SelectNewTool PAINT_GRADIENT, (actionSource = pdas_Search), True
        
        'Open the search panel and set focus to the search box
        Case "tool_search"
            toolbar_Layers.SetFocusToSearchBox
            
        Case Else
            cmdFound = False
            
    End Select
    
    Launch_ByName_NonMenu = cmdFound

End Function

Private Function Launch_ByName_Misc(ByRef srcMenuName As String, Optional ByVal actionSource As PD_ActionSource = pdas_Menu) As Boolean
    
    Dim cmdFound As Boolean: cmdFound = True
    
    'Image and macro paths can be supplied here.  Check these states up-front, by validating a hard-coded prefix
    ' (and extension, in the case of macros) and then verifying file existence.
    Dim targetFile As String
    If (LCase$(Left$(srcMenuName, 11)) = "image-file:") Then
        targetFile = Right$(srcMenuName, Len(srcMenuName) - 11)
        If Files.FileExists(targetFile) Then Loading.LoadFileAsNewImage targetFile
    ElseIf (LCase$(Left$(srcMenuName, 11)) = "macro-file:") Then
        targetFile = Right$(srcMenuName, Len(srcMenuName) - 11)
        If Files.FileExists(targetFile) And PDImages.IsImageActive() Then Macros.PlayMacroFromFile targetFile
    End If
    
    Launch_ByName_Misc = cmdFound
    
End Function

'PD's search bar aims to be a versatile tool.  It calls this function to retrieve search targets that don't
' fit nicely into the "menu" or "tools" category.
Public Sub GetMiscellaneousSearchActions(ByRef dstNames As pdStringStack, ByRef dstActions As pdStringStack)
    
    Set dstNames = New pdStringStack
    Set dstActions = New pdStringStack
    
    'This list is not managed automatically.  Stick any interesting and/or "hard-to-categorize" search results here.
    ' Just remember that you have to also supply an action trigger elsewhere in the module that actually executes
    ' the passed action!
    
    'These first few items are written like this for the localization engine.  We don't want to produce
    ' new terms for localization (because these terms already exist in the Menus module), but for complex
    ' technical reasons, the menu manager does not manage certain menus.  (Usually ones whose position
    ' changes at run-time.)
    dstNames.AddString g_Language.TranslateMessage("File") & " > " & g_Language.TranslateMessage("Open recent") & " > " & g_Language.TranslateMessage("Open all recent images")
    dstActions.AddString "file_open_allrecent"
    
    dstNames.AddString g_Language.TranslateMessage("File") & " > " & g_Language.TranslateMessage("Open recent") & " > " & g_Language.TranslateMessage("Clear recent image list")
    dstActions.AddString "file_open_clearrecent"
    
    'Next, add all the user's recent files to the list.
    Const MAX_LEN_IN_CHARS As Long = 40&
    Dim i As Long
    If (Not g_RecentFiles Is Nothing) Then
        If (g_RecentFiles.GetNumOfItems > 0) Then
            For i = 0 To g_RecentFiles.GetNumOfItems() - 1
                dstNames.AddString Files.PathCompact(Files.FileGetName(g_RecentFiles.GetFullPath(i), False), MAX_LEN_IN_CHARS)
                dstActions.AddString "image-file:" & g_RecentFiles.GetFullPath(i)
            Next i
        End If
    End If
    
    '...followed by all the user's recent macros
    If (Not g_RecentMacros Is Nothing) Then
        If (g_RecentMacros.MRU_ReturnCount > 0) Then
            For i = 0 To g_RecentMacros.MRU_ReturnCount() - 1
                dstNames.AddString Files.PathCompact(Files.FileGetName(g_RecentMacros.GetSpecificMRU(i), False), MAX_LEN_IN_CHARS)
                dstActions.AddString "macro-file:" & g_RecentMacros.GetSpecificMRU(i)
            Next i
        End If
    End If
    
End Sub

'Build a list of all internal action names and their basic attributes
Public Sub BuildActionDatabase()
    
    m_numActions = 0
    ReDim m_Actions(0 To 511) As PD_Action  'This upper limit will automatically be resized, as necessary
    
    AddAction "file_new", "New image"
    AddAction "file_open", "Open"
    'AddAction "file_openrecent"
    'AddAction "file_open_allrecent"
    'AddAction "file_open_clearrecent"
    'AddAction "file_import"
    AddAction "file_import_paste", "Paste to new image"
    AddAction "file_import_scanner", "Scan image"
    AddAction "file_import_selectscanner", "Select scanner or camera"
    AddAction "file_import_web", "Internet import"
    AddAction "file_import_screenshot", "Screen capture"
    AddAction "file_close", "Close"
    AddAction "file_closeall", "Close all"
    AddAction "file_save", "Save"
    AddAction "file_savecopy", "Save copy"
    AddAction "file_saveas", "Save as"
    AddAction "file_revert", "Revert"
    'AddAction "file_export"
    AddAction "file_export_image", "Export image to file"
    AddAction "file_export_layers", "Export layers to files"
    AddAction "file_export_animation", "Export animation"
    AddAction "file_export_colorlookup", "Export color lookup"
    AddAction "file_export_colorprofile", "Export color profile"
    AddAction "file_export_palette", "Export palette"
    'AddAction "file_batch"
    AddAction "file_batch_process", "Batch wizard"
    'AddAction "file_batch_repair"
    AddAction "file_print", "Print"
    AddAction "file_quit", "Exit program"
    
    AddAction "edit_undo", "Undo"
    AddAction "edit_redo", "Redo"
    AddAction "edit_history", "Undo history"
    AddAction "edit_repeat", "Repeat last action"
    AddAction "edit_fade", "Fade"
    AddAction "edit_cutlayer", "Cut"
    AddAction "edit_cutmerged", "Cut merged"
    AddAction "edit_copylayer", "Copy"
    AddAction "edit_copymerged", "Copy merged"
    AddAction "edit_pasteaslayer", "Paste"
    AddAction "edit_pastetocursor", "Paste to cursor"
    AddAction "edit_pasteasimage", "Paste to new image"
    AddAction "edit_specialcut", "Cut special"
    AddAction "edit_specialcopy", "Copy special"
    AddAction "edit_specialpaste", "Paste special"
    AddAction "edit_emptyclipboard", "Empty clipboard"
    AddAction "edit_clear", "Clear", True, True
    AddAction "edit_contentawarefill", "Content-aware fill", True, True
    AddAction "edit_fill", "Fill", True, True
    AddAction "edit_stroke", "Stroke", True, True
    
    AddAction "image_duplicate", "Duplicate image"
    AddAction "image_resize", "Resize image", True
    AddAction "image_contentawareresize", "Content-aware image resize", True
    AddAction "image_canvassize", "Canvas size", True
    AddAction "image_fittolayer", "Fit canvas to active layer"
    AddAction "image_fitalllayers", "Fit canvas around all layers"
    AddAction "image_crop", "Crop"
    AddAction "image_trim", "Trim empty image borders"
    'AddAction "image_rotate"
    AddAction "image_straighten", "Straighten image", True
    AddAction "image_rotate90", "Rotate image 90 clockwise", True
    AddAction "image_rotate270", "Rotate image 90 counter-clockwise", True
    AddAction "image_rotate180", "Rotate image 180", True
    AddAction "image_rotatearbitrary", "Arbitrary image rotation", True
    AddAction "image_fliphorizontal", "Flip image horizontally", True
    AddAction "image_flipvertical", "Flip image vertically", True
    AddAction "image_mergevisible", "Merge visible layers", True
    AddAction "image_flatten", "Flatten image"
    AddAction "image_animation", "Animation options"
    'AddAction "image_compare"
    AddAction "image_createlut", "Create color lookup"
    AddAction "image_similarity", "Compare similarity"
    'AddAction "image_metadata"
    AddAction "image_editmetadata", "Edit metadata"
    AddAction "image_removemetadata", "Remove all metadata"
    AddAction "image_countcolors", "Count unique colors"
    AddAction "image_maplocation", vbNullString
    AddAction "image_showinexplorer", vbNullString
    
    'AddAction "layer_add"
    AddAction "layer_addbasic", "Add new layer", True
    AddAction "layer_addblank", "Add blank layer", True
    AddAction "layer_duplicate", "Duplicate Layer", True
    AddAction "layer_addfromclipboard", "Paste"
    AddAction "layer_addfromfile", "New layer from file"
    AddAction "layer_addfromvisiblelayers", "New layer from visible layers", True
    AddAction "layer_addviacopy", "Layer via copy", True
    AddAction "layer_addviacut", "Layer via cut", True
    'AddAction "layer_delete"
    AddAction "layer_deletecurrent", "Delete layer", True
    AddAction "layer_deletehidden", "Delete hidden layers"
    'AddAction "layer_replace"
    AddAction "layer_replacefromclipboard", "Replace layer from clipboard"
    AddAction "layer_replacefromfile", "Replace layer from file"
    AddAction "layer_replacefromvisiblelayers", "Replace layer from visible layers"
    AddAction "layer_mergeup", "Merge layer up", True
    AddAction "layer_mergedown", "Merge layer down", True
    'AddAction "layer_order"
    AddAction "layer_gotop", "Go to top layer"
    AddAction "layer_goup", "Go to layer above"
    AddAction "layer_godown", "Go to layer below"
    AddAction "layer_gobottom", "Go to bottom layer"
    AddAction "layer_movetop", "Raise layer to top"
    AddAction "layer_moveup", "Raise layer"
    AddAction "layer_movedown", "Lower layer"
    AddAction "layer_movebottom", "Lower layer to bottom"
    AddAction "layer_reverse", "Reverse layer order"
    'AddAction "layer_visibility"
    AddAction "layer_show", "Toggle layer visibility"
    AddAction "layer_showonly", "Show only this layer"
    AddAction "layer_hideonly", "Hide only this layer"
    AddAction "layer_showall", "Show all layers"
    AddAction "layer_hideall", "Hide all layers"
    'AddAction "layer_crop"
    AddAction "layer_cropselection", "Crop layer to selection"
    AddAction "layer_pad", "Pad layer to image size"
    AddAction "layer_trim", "Trim empty layer borders"
    'AddAction "layer_orientation"
    AddAction "layer_straighten", "Straighten layer", True
    AddAction "layer_rotate90", "Rotate layer 90 clockwise", True
    AddAction "layer_rotate270", "Rotate layer 90 counter-clockwise", True
    AddAction "layer_rotate180", "Rotate layer 180", True
    AddAction "layer_rotatearbitrary", "Arbitrary layer rotation", True
    AddAction "layer_fliphorizontal", "Flip layer horizontally", True
    AddAction "layer_flipvertical", "Flip layer vertically", True
    'AddAction "layer_size"
    AddAction "layer_resetsize", "Reset layer size"
    AddAction "layer_resize", "Resize layer", True
    AddAction "layer_contentawareresize", "Content-aware layer resize", True
    AddAction "layer_fittoimage", "Fit layer to image"
    'AddAction "layer_transparency"
    AddAction "layer_colortoalpha", "Color to alpha"
    AddAction "layer_luminancetoalpha", "Luminance to alpha"
    AddAction "layer_removealpha", "Remove alpha channel"
    AddAction "layer_thresholdalpha", "Threshold alpha"
    'AddAction "layer_rasterize"
    AddAction "layer_rasterizecurrent", "Rasterize layer"
    AddAction "layer_rasterizeall", "Rasterize all layers"
    'AddAction "layer_split"
    AddAction "layer_splitlayertoimage", "Split layer into image"
    AddAction "layer_splitalllayerstoimages", "Split layers into images"
    AddAction "layer_splitimagestolayers", "Split images into layers"
    
    AddAction "select_all", "Select all"
    AddAction "select_none", "Remove selection"
    AddAction "select_invert", "Invert selection", True
    AddAction "select_grow", "Grow selection", True
    AddAction "select_shrink", "Shrink selection", True
    AddAction "select_border", "Border selection", True
    AddAction "select_feather", "Feather selection", True
    AddAction "select_sharpen", "Sharpen selection", True
    AddAction "select_erasearea", "Erase selected area", True
    AddAction "select_fill", "Fill selected area", True, True
    AddAction "select_heal", "Heal selected area", True, True
    AddAction "select_stroke", "Stroke selection outline", True, True
    AddAction "select_load", "Load selection"
    AddAction "select_save", "Save selection"
    'AddAction "select_export"
    AddAction "select_exportarea", "Export selected area as image"
    AddAction "select_exportmask", "Export selection mask as image"
    
    AddAction "adj_autocorrect", "Auto correct", True, True
    AddAction "adj_autoenhance", "Auto enhance", True, True
    AddAction "adj_blackandwhite", "Black and white", True, True
    AddAction "adj_bandc", "Brightness and contrast", True, True
    AddAction "adj_colorbalance", "Color balance", True, True
    AddAction "adj_curves", "Curves", True, True
    AddAction "adj_levels", "Levels", True, True
    AddAction "adj_sandh", "Shadows and highlights", True, True
    AddAction "adj_vibrance", "Vibrance", True, True
    AddAction "adj_whitebalance", "White balance", True, True
    'AddAction "adj_channels"
    AddAction "adj_channelmixer", "Channel mixer", True, True
    AddAction "adj_rechannel", "Rechannel", True, True
    AddAction "adj_maxchannel", "Maximum channel", True, True
    AddAction "adj_minchannel", "Minimum channel", True, True
    AddAction "adj_shiftchannelsleft", "Shift colors (left)", True, True
    AddAction "adj_shiftchannelsright", "Shift colors (right)", True, True
    'AddAction "adj_color"
    AddAction "adj_hsl", "Hue and saturation", True, True
    AddAction "adj_temperature", "Temperature", True, True
    AddAction "adj_tint", "Tint", True, True
    AddAction "adj_colorlookup", "Color lookup", True, True
    AddAction "adj_colorize", "Colorize", True, True
    AddAction "adj_photofilters", "Photo filter", True, True
    AddAction "adj_replacecolor", "Replace color", True, True
    AddAction "adj_sepia", "Sepia", True, True
    AddAction "adj_splittone", "Split toning", True, True
    'AddAction "adj_histogram"
    'AddAction "adj_histogramdisplay"
    AddAction "adj_histogramequalize", "Equalize", True, True
    AddAction "adj_histogramstretch", "Stretch histogram", True, True
    'AddAction "adj_invert
    AddAction "adj_invertcmyk", "Film negative", True, True
    AddAction "adj_inverthue", "Invert hue", True, True
    AddAction "adj_invertrgb", "Invert RGB", True, True
    'AddAction "adj_lighting"
    AddAction "adj_dehaze", "Dehaze", True, True
    AddAction "adj_exposure", "Exposure", True, True
    AddAction "adj_gamma", "Gamma", True, True
    AddAction "adj_hdr", "HDR", True, True
    'AddAction "adj_map"
    AddAction "adj_gradientmap", "Gradient map", True, True
    AddAction "adj_palettemap", "Palette map", True, True
    'AddAction "adj_monochrome"
    AddAction "adj_colortomonochrome", "Color to monochrome", True, True
    AddAction "adj_monochrometogray", "Monochrome to gray", True, True
    
    'AddAction "effects_artistic"
    AddAction "effects_colorpencil", "Colored pencil", True, True
    AddAction "effects_comicbook", "Comic book", True, True
    AddAction "effects_figuredglass", "Figured glass", True, True
    AddAction "effects_filmnoir", "Film noir", True, True
    AddAction "effects_glasstiles", "Glass tiles", True, True
    AddAction "effects_kaleidoscope", "Kaleidoscope", True, True
    AddAction "effects_modernart", "Modern art", True, True
    AddAction "effects_oilpainting", "Oil painting", True, True
    AddAction "effects_plasticwrap", "Plastic wrap", True, True
    AddAction "effects_posterize", "Posterize", True, True
    AddAction "effects_relief", "Relief", True, True
    AddAction "effects_stainedglass", "Stained glass", True, True
    'AddAction "effects_blur"
    AddAction "effects_boxblur", "Box blur", True, True
    AddAction "effects_gaussianblur", "Gaussian blur", True, True
    AddAction "effects_surfaceblur", "Surface blur", True, True
    AddAction "effects_motionblur", "Motion blur", True, True
    AddAction "effects_radialblur", "Radial blur", True, True
    AddAction "effects_zoomblur", "Zoom blur", True, True
    'AddAction "effects_distort"
    AddAction "effects_fixlensdistort", "Correct lens distortion", True, True
    AddAction "effects_donut", "Donut", True, True
    AddAction "effects_droste", "Droste", True, True
    AddAction "effects_lens", "Apply lens distortion", True, True
    AddAction "effects_pinchandwhirl", "Pinch and whirl", True, True
    AddAction "effects_poke", "Poke", True, True
    AddAction "effects_ripple", "Ripple", True, True
    AddAction "effects_squish", "Squish", True, True
    AddAction "effects_swirl", "Swirl", True, True
    AddAction "effects_waves", "Waves", True, True
    AddAction "effects_miscdistort", "Miscellaneous distort", True, True
    'AddAction "effects_edges"
    AddAction "effects_emboss", "Emboss", True, True
    AddAction "effects_enhanceedges", "Enhance edges", True, True
    AddAction "effects_findedges", "Find edges", True, True
    AddAction "effects_gradientflow", "Gradient flow", True, True
    AddAction "effects_rangefilter", "Range filter", True, True
    AddAction "effects_tracecontour", "Trace contour", True, True
    'AddAction "effects_lightandshadow"
    AddAction "effects_blacklight", "Black light", True, True
    AddAction "effects_bumpmap", "Bump map", True, True
    AddAction "effects_crossscreen", "Cross-screen", True, True
    AddAction "effects_rainbow", "Rainbow", True, True
    AddAction "effects_sunshine", "Sunshine", True, True
    AddAction "effects_dilate", "Dilate (maximum rank)", True, True
    AddAction "effects_erode", "Erode (minimum rank)", True, True
    'AddAction "effects_natural"
    AddAction "effects_atmosphere", "Atmosphere", True, True
    AddAction "effects_fog", "Fog", True, True
    AddAction "effects_ignite", "Ignite", True, True
    AddAction "effects_lava", "Lava", True, True
    AddAction "effects_metal", "Metal", True, True
    AddAction "effects_snow", "Snow", True, True
    AddAction "effects_underwater", "Water", True, True
    'AddAction "effects_noise"
    AddAction "effects_filmgrain", "Add film grain", True, True
    AddAction "effects_rgbnoise", "Add RGB noise", True, True
    AddAction "effects_anisotropic", "Anisotropic diffusion", True, True
    AddAction "effects_bilateral", "Surface blur", True, True 'For legacy macros only; bilateral has been replaced by Blur > Surface Blur
    AddAction "effects_dustandscratches", "Dust and scratches", True, True
    AddAction "effects_harmonicmean", "Harmonic mean", True, True
    AddAction "effects_meanshift", "Mean shift", True, True
    AddAction "effects_median", "Median", True, True
    AddAction "effects_snn", "Symmetric nearest-neighbor", True, True
    'AddAction "effects_pixelate"
    AddAction "effects_colorhalftone", "Color halftone", True, True
    AddAction "effects_crystallize", "Crystallize", True, True
    AddAction "effects_fragment", "Fragment", True, True
    AddAction "effects_mezzotint", "Mezzotint", True, True
    AddAction "effects_mosaic", "Mosaic", True, True
    AddAction "effects_pointillize", "Pointillize", True, True
    'AddAction "effects_render"
    AddAction "effects_clouds", "Clouds", True, True
    AddAction "effects_fibers", "Fibers", True, True
    AddAction "effects_truchet", "Truchet", True, True
    '"effects_sharpentop"
    AddAction "effects_sharpen", "Sharpen", True, True
    AddAction "effects_unsharp", "Unsharp mask", True, True
    'AddAction "effects_stylize"
    AddAction "effects_antique", "Antique", True, True
    AddAction "effects_diffuse", "Diffuse", True, True
    AddAction "effects_kuwahara", "Kuwahara filter", True, True
    AddAction "effects_outline", "Outline", True, True
    AddAction "effects_palette", "Palette", True, True
    AddAction "effects_portraitglow", "Portrait glow", True, True
    AddAction "effects_solarize", "Solarize", True, True
    AddAction "effects_twins", "Twins", True, True
    AddAction "effects_vignetting", "Vignetting", True, True
    '"effects_transform"
    AddAction "effects_panandzoom", "Offset and zoom", True, True
    AddAction "effects_perspective", "Perspective", True, True
    AddAction "effects_polarconversion", "Polar conversion", True, True
    AddAction "effects_rotate", "Rotate", True, True
    AddAction "effects_shear", "Shear", True, True
    AddAction "effects_spherize", "Spherize", True, True
    'AddAction "effects_animation"
    AddAction "effects_animation_background", "Animation background"
    AddAction "effects_animation_foreground", "Animation foreground"
    AddAction "effects_animation_speed", "Animation playback speed"
    AddAction "effects_customfilter", "Custom filter", True, True
    AddAction "effects_8bf", "Photoshop (8bf) plugin", True, True
    
    'AddAction "tools_language"
    AddAction "tools_languageeditor", vbNullString
    AddAction "tools_theme", vbNullString
    'AddAction "tools_macrocreatetop"
    AddAction "tools_macrofromhistory", vbNullString
    AddAction "tools_recordmacro", "Start macro recording"
    AddAction "tools_stopmacro", "Stop macro recording"
    AddAction "tools_playmacro""Play macro"
    'AddAction "tools_recentmacros"
    AddAction "tools_screenrecord", vbNullString
    AddAction "tools_options", vbNullString
    'AddAction "tools_developers"
    'AddAction "tools_themeeditor", vbNullString
    'AddAction "tools_themepackage"
    'AddAction "tools_standalonepackage"
    'AddAction "effects_developertest"
    
    AddAction "view_fit", vbNullString
    AddAction "view_center_on_screen", vbNullString
    AddAction "view_zoomin", vbNullString
    AddAction "view_zoomout", vbNullString
    'AddAction "view_zoomtop"
    AddAction "zoom_16_1", vbNullString
    AddAction "zoom_8_1", vbNullString
    AddAction "zoom_4_1", vbNullString
    AddAction "zoom_2_1", vbNullString
    AddAction "zoom_actual", vbNullString
    AddAction "zoom_1_2", vbNullString
    AddAction "zoom_1_4", vbNullString
    AddAction "zoom_1_8", vbNullString
    AddAction "zoom_1_16", vbNullString
    AddAction "view_rulers", vbNullString
    AddAction "view_statusbar", vbNullString
    AddAction "show_layeredges", vbNullString
    AddAction "show_smartguides", vbNullString
    AddAction "snap_global", vbNullString
    AddAction "snap_canvasedge", vbNullString
    AddAction "snap_centerline", vbNullString
    AddAction "snap_layer", vbNullString
    AddAction "snap_angle_90", vbNullString
    AddAction "snap_angle_45", vbNullString
    AddAction "snap_angle_30", vbNullString
    
    'AddAction "window_toolbox"
    AddAction "window_displaytoolbox", vbNullString
    AddAction "window_displaytoolcategories", vbNullString
    AddAction "window_smalltoolbuttons", vbNullString
    AddAction "window_mediumtoolbuttons", vbNullString
    AddAction "window_largetoolbuttons", vbNullString
    AddAction "window_tooloptions", vbNullString
    AddAction "window_layers", vbNullString
    'AddAction "window_imagetabstrip"
    AddAction "window_imagetabstrip_alwaysshow", vbNullString
    AddAction "window_imagetabstrip_shownormal", vbNullString
    AddAction "window_imagetabstrip_nevershow", vbNullString
    AddAction "window_imagetabstrip_alignleft", vbNullString
    AddAction "window_imagetabstrip_aligntop", vbNullString
    AddAction "window_imagetabstrip_alignright", vbNullString
    AddAction "window_imagetabstrip_alignbottom", vbNullString
    AddAction "window_resetsettings", vbNullString
    AddAction "window_next", vbNullString
    AddAction "window_previous", vbNullString
    
    AddAction "help_patreon", vbNullString
    AddAction "help_donate", vbNullString
    AddAction "help_forum", vbNullString
    AddAction "help_checkupdates", vbNullString
    AddAction "help_reportbug", vbNullString
    AddAction "help_license", vbNullString
    AddAction "help_sourcecode", vbNullString
    AddAction "help_website", vbNullString
    AddAction "help_3rdpartylibs", vbNullString
    AddAction "help_about", vbNullString
    
    'Tool and brush actions follow.  These typically do *not* have corresponding processor actions,
    ' and are not recorded by macros.  (Many are UI-specific actions associated with hotkeys,
    ' e.g. "increase brush size".)
    AddAction "tool_active_hardnessdown"
    AddAction "tool_active_hardnessup"
    AddAction "tool_active_sizedown"
    AddAction "tool_active_sizeup"
    
    PDDebug.LogAction CStr(m_numActions) & " actions registered this session."
    
End Sub

Private Sub AddAction(ByVal actionName As String, Optional ByRef processName As String = vbNullString, Optional ByVal isRepeatable As Boolean = False, Optional ByVal isFadeable As Boolean = False)
    
    If (m_numActions > UBound(m_Actions)) Then ReDim Preserve m_Actions(0 To m_numActions * 2 - 1) As PD_Action
    
    With m_Actions(m_numActions)
        .actionName = actionName
        .processName = processName
        .isRepeatable = isRepeatable
        .isFadeable = isFadeable
    End With
    
    m_numActions = m_numActions + 1
    
End Sub

Private Function GetActionIndexFromName(ByRef actionName As String, Optional ByVal nameIsProcessName As Boolean = False) As Long
    
    GetActionIndexFromName = -1
    
    If (m_numActions > 0) And (LenB(actionName) > 0) Then
        
        Dim i As Long
        If nameIsProcessName Then
            For i = 0 To m_numActions - 1
                If Strings.StringsEqual(actionName, m_Actions(i).processName, True) Then
                    GetActionIndexFromName = i
                    Exit For
                End If
            Next i
        Else
            For i = 0 To m_numActions - 1
                If Strings.StringsEqual(actionName, m_Actions(i).actionName, True) Then
                    GetActionIndexFromName = i
                    Exit For
                End If
            Next i
        End If
        
    End If
    
End Function

Public Function IsActionRepeatable(ByRef actionName As String, Optional ByVal nameIsProcessName As Boolean = False) As Boolean
    Dim idxAction As Long
    idxAction = GetActionIndexFromName(actionName, nameIsProcessName)
    If (idxAction >= 0) Then IsActionRepeatable = m_Actions(idxAction).isRepeatable Else IsActionRepeatable = False
End Function

Public Function IsActionFadeable(ByRef actionName As String, Optional ByVal nameIsProcessName As Boolean = False) As Boolean
    Dim idxAction As Long
    idxAction = GetActionIndexFromName(actionName, nameIsProcessName)
    If (idxAction >= 0) Then IsActionFadeable = m_Actions(idxAction).isFadeable Else IsActionFadeable = False
End Function
