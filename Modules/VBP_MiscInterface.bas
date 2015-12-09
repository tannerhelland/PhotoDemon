Attribute VB_Name = "Interface"
'***************************************************************************
'Miscellaneous Functions Related to the User Interface
'Copyright 2001-2015 by Tanner Helland
'Created: 6/12/01
'Last updated: 20/June/14
'Last update: add interface-syncing functions for non-destructive edit tools
'
'Miscellaneous routines related to rendering and handling PhotoDemon's interface.  As the program's complexity has
' increased, so has the need for specialized handling of certain UI elements.
'
'Many of the functions in this module rely on subclassing, either directly or through things like PD's window manager.
' As such, some functions may operate differently (or not at all) while in the IDE.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************


Option Explicit

Private Declare Function GetWindowRect Lib "user32" (ByVal hndWindow As Long, ByRef lpRect As winRect) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hndWindow As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

'Used to measure the expected length of a string
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, ByRef lpSize As POINTAPI) As Long

'These constants are used to toggle visibility of display elements.
Public Const VISIBILITY_TOGGLE As Long = 0
Public Const VISIBILITY_FORCEDISPLAY As Long = 1
Public Const VISIBILITY_FORCEHIDE As Long = 2

'These values are used to remember the user's current font smoothing setting.  We try to be polite and restore
' the original setting when the application terminates.
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uiAction As Long, ByVal uiParam As Long, ByRef pvParam As Long, ByVal fWinIni As Long) As Long

Private Const SPI_GETFONTSMOOTHING As Long = &H4A
Private Const SPI_SETFONTSMOOTHING As Long = &H4B
Private Const SPI_GETFONTSMOOTHINGTYPE As Long = &H200A
Private Const SPI_SETFONTSMOOTHINGTYPE As Long = &H200B
Private Const SmoothingClearType As Long = &H2
Private Const SmoothingStandardType As Long = &H1
Private Const SmoothingNone As Long = &H0
Private Const SPI_GETKEYBOARDDELAY As Long = &H16
Private Const SPI_GETKEYBOARDSPEED As Long = &HA

'Constants that define single meta-level actions that require certain controls to be en/disabled.  (For example, use tSave to disable
' the File -> Save menu, file toolbar Save button, and Ctrl+S hotkey.)  Constants are listed in roughly the order they appear in the
' main menu.
Public Enum metaInitializer
     tSave
     tSaveAs
     tClose
     tUndo
     tRedo
     tCopy
     tPaste
     tView
     tImageOps
     tMetadata
     tGPSMetadata
     tMacro
     tSelection
     tSelectionTransform
     tZoom
     tLayerTools
     tNonDestructiveFX
End Enum

#If False Then
    Private Const tSave = 0, tSaveAs = 0, tClose = 0, tUndo = 0, tRedo = 0, tCopy = 0, tPaste = 0, tView = 0, tImageOps = 0, tMetadata = 0, tGPSMetadata = 0
    Private Const tMacro = 0, tSelection = 0, tSelectionTransform = 0, tZoom = 0, tLayerTools = 0, tNonDestructiveFX = 0
#End If

'If PhotoDemon enabled font smoothing where there was none previously, it will restore the original setting upon exit.  This variable
' can contain the following values:
' 0: did not have to change smoothing, as ClearType is already enabled
' 1: had to change smoothing type from Standard to ClearType
' 2: had to turn on smoothing, as it was originally turned off
Private hadToChangeSmoothing As Long

'PhotoDemon is designed against pixels at an expected screen resolution of 96 DPI.  Other DPI settings mess up our calculations.
' To remedy this, we dynamically modify all pixels measurements at run-time, using the current screen resolution as our guide.
Private dpiRatio As Double

'When a modal dialog is displayed, a reference to it is saved in this variable.  If subsequent modal dialogs are displayed (for example,
' if a tool dialog displays a color selection dialog), the previous modal dialog is given ownership over the new dialog.
Private currentDialogReference As Form
Private isSecondaryDialog As Boolean

'When a message is displayed to the user in the message portion of the status bar, we automatically cache the message's contents.
' If a subsequent request is raised with the exact same text, we can skip the whole message display process.
Private m_PrevMessage As String

'System DPI is used frequently for UI positioning calculations.  Because it's costly to constantly retrieve it via APIs, this module
' prefers to cache it only when the value changes.  Call the CacheSystemDPI() sub to update the value when appropriate, and the
' corresponding GetSystemDPI() function to retrieve the cached value.
Private m_CurrentSystemDPI As Single

Public Sub CacheSystemDPI(ByVal newDPI As Single)
    m_CurrentSystemDPI = newDPI
End Sub

Public Function GetSystemDPI() As Single
    GetSystemDPI = m_CurrentSystemDPI
End Function

'Previously, various PD functions had to manually enable/disable button and menu state based on their actions.  This is no longer necessary.
' Simply call this function whenever an action has done something that will potentially affect the interface, and this function will iterate
' through all potential image/interface interactions, dis/enabling buttons and menus as necessary.
'
'TODO: look at having an optional "layerID" parameter, so we can skip certain steps if only a single layer is affected by a change.
Public Sub SyncInterfaceToCurrentImage()
    
    Dim i As Long
    
    'Interface dis/enabling falls into two rough categories: stuff that changes based on the current image (e.g. Undo), and stuff that changes
    ' based on the *total* number of available images (e.g. visibility of the Effects menu).
    
    'Start by breaking our interface decisions into two broad categories: "no images are loaded" and "one or more images are loaded".
    
    'If no images are loaded, we can disable a whole swath of controls
    If g_OpenImageCount = 0 Then
    
        MetaToggle tSave, False
        MetaToggle tSaveAs, False
        MetaToggle tClose, False
        MetaToggle tCopy, False
        MetaToggle tView, False
        MetaToggle tImageOps, False
        MetaToggle tSelection, False
        MetaToggle tMacro, False
        MetaToggle tZoom, False
        MetaToggle tLayerTools, False
        MetaToggle tNonDestructiveFX, False
        
        'Disable various layer-related commands as well
        FormMain.MnuLayerSize(0).Enabled = False
        toolpanel_MoveSize.cmdLayerMove(0).Enabled = False
        toolpanel_MoveSize.cmdLayerMove(1).Enabled = False
        toolpanel_MoveSize.cmdLayerAngleReset.Enabled = False
        toolpanel_MoveSize.cmdLayerShearReset(0).Enabled = False
        toolpanel_MoveSize.cmdLayerShearReset(1).Enabled = False
        toolpanel_MoveSize.cmdLayerAffinePermanent.Enabled = False
        
        'Reset all Undo/Redo and related menus as well
        SyncUndoRedoInterfaceElements True
                
        'All relevant menu icons can now be redrawn.  (This must be redone after menu captions change, as icons are associated
        ' with captions.)
        resetMenuIcons
        
        '"Paste as new layer" is disabled when no images are loaded (but "Paste as new image" remains active)
        FormMain.MnuEdit(8).Enabled = False
                        
        Message "Please load or import an image to begin editing."
        
        'Assign a generic caption to the main window
        If Not (g_WindowManager Is Nothing) Then
            g_WindowManager.SetWindowCaptionW FormMain.hWnd, getPhotoDemonNameAndVersion()
        Else
            FormMain.Caption = getPhotoDemonNameAndVersion()
        End If
        
        'Erase the main viewport's status bar
        FormMain.mainCanvas(0).displayImageSize Nothing, True
        FormMain.mainCanvas(0).drawStatusBarIcons False
        
        'Because dynamic icons are enabled, restore the main program icon and clear the custom image icon cache
        setNewTaskbarIcon origIcon32, FormMain.hWnd
        setNewAppIcon origIcon16, origIcon32
        Icon_and_Cursor_Handler.DestroyAllIcons
        
        'If no images are currently open, but images were open in the past, release any memory associated with those images.
        ' This helps minimize PD's memory usage.
        If g_NumOfImagesLoaded > 1 Then
        
            'Loop through all pdImage objects and make sure they've been deactivated
            For i = 0 To UBound(pdImages)
                If (Not pdImages(i) Is Nothing) Then
                    pdImages(i).deactivateImage
                    Set pdImages(i) = Nothing
                End If
            Next i
            
            'Reset all window tracking variables
            g_NumOfImagesLoaded = 0
            g_CurrentImage = 0
            g_OpenImageCount = 0
            
        End If
                
        'Erase any remaining viewport buffer.  (This is temporarily disabled because the RAM gain is small, and it potentially introduces
        ' errors into functions that expect an activate viewport pipeline.  Further investigation TBD.)
        'eraseViewportBuffers
    
    'If one or more images are loaded, our job is trickier.  Some controls (such as Copy to Clipboard) are enabled no matter what,
    ' while others (Undo and Redo) are only enabled if the current image requires it.
    Else
        
        If Not pdImages(g_CurrentImage) Is Nothing Then
        
            'Start by enabling actions that are always available if one or more images are loaded.
            MetaToggle tSaveAs, True
            MetaToggle tClose, True
            MetaToggle tCopy, True
            
            MetaToggle tView, True
            MetaToggle tZoom, True
            MetaToggle tImageOps, True
            MetaToggle tMacro, True
            MetaToggle tLayerTools, True
            
            'Paste as new layer is always available if one (or more) images are loaded
            If Not FormMain.MnuEdit(9).Enabled Then FormMain.MnuEdit(9).Enabled = True
            
            'Display this image's path in the title bar.
            If Not (g_WindowManager Is Nothing) Then
                g_WindowManager.SetWindowCaptionW FormMain.hWnd, GetWindowCaption(pdImages(g_CurrentImage))
            Else
                FormMain.Caption = GetWindowCaption(pdImages(g_CurrentImage))
            End If
            
            'Draw icons onto the main viewport's status bar
            FormMain.mainCanvas(0).drawStatusBarIcons True
            
            'Next, attempt to enable controls whose state depends on the current image - e.g. "Save", which is only enabled if
            ' the image has not already been saved in its current state.
            
            'Note that all of these functions rely on the g_CurrentImage value to function.
            
            'Reset all Undo/Redo and related menus.  (Note that this also controls the SAVE BUTTON, as the image's save state is modified
            ' by PD's Undo/Redo engine.)
            SyncUndoRedoInterfaceElements True
            
            'Because those changes may modify menu captions, menu icons need to be reset (as they are tied to menu captions)
            resetMenuIcons
            
            'Determine whether metadata is present, and dis/enable metadata menu items accordingly
            If Not pdImages(g_CurrentImage).imgMetadata Is Nothing Then
                MetaToggle tMetadata, pdImages(g_CurrentImage).imgMetadata.hasXMLMetadata
                MetaToggle tGPSMetadata, pdImages(g_CurrentImage).imgMetadata.hasGPSMetadata()
            Else
                MetaToggle tMetadata, False
                MetaToggle tGPSMetadata, False
            End If
            
            'Display the size of this image in the status bar
            If pdImages(g_CurrentImage).Width <> 0 Then DisplaySize pdImages(g_CurrentImage)
            
            'Update the form's icon to match the current image; if a custom icon is not available, use the stock PD one
            If pdImages(g_CurrentImage).curFormIcon32 <> 0 Then
                
                'If images are docked, they do not have their own taskbar entries.  Change the main program icon to match this image.
                setNewTaskbarIcon pdImages(g_CurrentImage).curFormIcon32, FormMain.hWnd
                setNewAppIcon pdImages(g_CurrentImage).curFormIcon16, pdImages(g_CurrentImage).curFormIcon32
                
            Else
                setNewTaskbarIcon origIcon32, FormMain.hWnd
            End If
                        
            'Check the image's color depth, and check/uncheck the matching Image Mode setting
            'If Not (pdImages(g_CurrentImage).getActiveLayer() Is Nothing) Then
            '    If pdImages(g_CurrentImage).getCompositeImageColorDepth() = 32 Then metaToggle tImgMode32bpp, True Else metaToggle tImgMode32bpp, False
            'End If
            
            'Restore the zoom value for this particular image (again, only if the form has been initialized)
            If pdImages(g_CurrentImage).Width <> 0 Then
                g_AllowViewportRendering = False
                FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex = pdImages(g_CurrentImage).currentZoomValue
                g_AllowViewportRendering = True
            End If
            
            'If a selection is active on this image, update the text boxes to match
            If pdImages(g_CurrentImage).selectionActive And (Not pdImages(g_CurrentImage).mainSelection Is Nothing) Then
                MetaToggle tSelection, True
                MetaToggle tSelectionTransform, pdImages(g_CurrentImage).mainSelection.isTransformable
                syncTextToCurrentSelection g_CurrentImage
            Else
                MetaToggle tSelection, False
                MetaToggle tSelectionTransform, False
            End If
            
            'Update all layer menus; some will be disabled depending on just how many layers are available, how many layers
            ' are visible, and other criteria.
            If pdImages(g_CurrentImage).getNumOfLayers > 0 Then
                
                'First, set some parameters contingent on the current layer's options
                If Not pdImages(g_CurrentImage).getActiveLayer Is Nothing Then
                
                    'First, determine if the current layer is using any form of non-destructive resizing
                    Dim nonDestructiveResizeActive As Boolean
                    nonDestructiveResizeActive = False
                    If (pdImages(g_CurrentImage).getActiveLayer.getLayerCanvasXModifier <> 1) Then nonDestructiveResizeActive = True
                    If (pdImages(g_CurrentImage).getActiveLayer.getLayerCanvasYModifier <> 1) Then nonDestructiveResizeActive = True
                    
                    'If non-destructive resizing is active, the "reset layer size" menu (and corresponding Move Tool button) must be enabled.
                    FormMain.MnuLayerSize(0).Enabled = nonDestructiveResizeActive
                    toolpanel_MoveSize.cmdLayerMove(0).Enabled = nonDestructiveResizeActive
                    toolpanel_MoveSize.cmdLayerMove(1).Enabled = pdImages(g_CurrentImage).getActiveLayer.affineTransformsActive(True)
                    
                    'Similar logic is used for other non-destructive affine transforms
                    toolpanel_MoveSize.cmdLayerAngleReset.Enabled = CBool(pdImages(g_CurrentImage).getActiveLayer.getLayerAngle <> 0)
                    toolpanel_MoveSize.cmdLayerShearReset(0).Enabled = CBool(pdImages(g_CurrentImage).getActiveLayer.getLayerShearX <> 0)
                    toolpanel_MoveSize.cmdLayerShearReset(1).Enabled = CBool(pdImages(g_CurrentImage).getActiveLayer.getLayerShearY <> 0)
                    toolpanel_MoveSize.cmdLayerAffinePermanent.Enabled = pdImages(g_CurrentImage).getActiveLayer.affineTransformsActive(True)
                    
                    'If non-destructive FX are active on the current layer, update the non-destructive tool enablement to match
                    MetaToggle tNonDestructiveFX, True
                    
                    'Layer rasterization depends on the current layer type
                    FormMain.MnuLayerRasterize(0).Enabled = pdImages(g_CurrentImage).getActiveLayer.isLayerVector
                    FormMain.MnuLayerRasterize(1).Enabled = CBool(pdImages(g_CurrentImage).getNumOfVectorLayers > 0)
                    
                End If
                
                'If only one layer is present, a number of layer menu items (Delete, Flatten, Merge, Order) will be disabled.
                If pdImages(g_CurrentImage).getNumOfLayers = 1 Then
                
                    'Delete
                    FormMain.MnuLayer(1).Enabled = False
                
                    'Merge up/down
                    FormMain.MnuLayer(3).Enabled = False
                    FormMain.MnuLayer(4).Enabled = False
                    
                    'Layer order
                    FormMain.MnuLayer(5).Enabled = False
                    
                    'Flatten
                    FormMain.MnuLayer(15).Enabled = False
                    
                    'Merge visible
                    FormMain.MnuLayer(16).Enabled = False
                    
                'This image contains multiple layers.  Enable many menu items (if they aren't already).
                Else
                
                    'Delete
                    If Not FormMain.MnuLayer(1).Enabled Then FormMain.MnuLayer(1).Enabled = True
                    
                    'Delete hidden layers is only available if one or more layers are hidden, but not ALL layers are hidden.
                    If (pdImages(g_CurrentImage).getNumOfHiddenLayers > 0) And (pdImages(g_CurrentImage).getNumOfHiddenLayers < pdImages(g_CurrentImage).getNumOfLayers) Then
                        FormMain.MnuLayerDelete(1).Enabled = True
                    Else
                        FormMain.MnuLayerDelete(1).Enabled = False
                    End If
                
                    'Merge up/down are not available for layers at the top and bottom of the image
                    If isLayerAllowedToMergeAdjacent(pdImages(g_CurrentImage).getActiveLayerIndex, False) <> -1 Then
                        FormMain.MnuLayer(3).Enabled = True
                    Else
                        FormMain.MnuLayer(3).Enabled = False
                    End If
                    
                    If isLayerAllowedToMergeAdjacent(pdImages(g_CurrentImage).getActiveLayerIndex, True) <> -1 Then
                        FormMain.MnuLayer(4).Enabled = True
                    Else
                        FormMain.MnuLayer(4).Enabled = False
                    End If
                    
                    'Order is always available if more than one layer exists in the image
                    If Not FormMain.MnuLayer(5).Enabled Then FormMain.MnuLayer(5).Enabled = True
                    
                    'Within the order menu, certain items are disabled based on layer position.  Note that "move up" and
                    ' "move to top" are both disabled for top images (similarly for bottom images and "move down/bottom"),
                    ' so we can mirror the same enabled state for both options.
                    If pdImages(g_CurrentImage).getActiveLayerIndex < pdImages(g_CurrentImage).getNumOfLayers - 1 Then
                        FormMain.MnuLayerOrder(0).Enabled = True
                    Else
                        FormMain.MnuLayerOrder(0).Enabled = False
                    End If
                    
                    If pdImages(g_CurrentImage).getActiveLayerIndex > 0 Then
                        FormMain.MnuLayerOrder(1).Enabled = True
                    Else
                        FormMain.MnuLayerOrder(1).Enabled = False
                    End If
                    
                    'Mirror "raise to top" and "lower to bottom" against the state of "raise layer" and "lower layer"
                    FormMain.MnuLayerOrder(3).Enabled = FormMain.MnuLayerOrder(0).Enabled
                    FormMain.MnuLayerOrder(4).Enabled = FormMain.MnuLayerOrder(1).Enabled
                                    
                    'Adding transparency to a layer is always permitted, but removing it is invalid if an image is already 24bpp.
                    ' Note that at present, this may have unintended consequences - use with caution!
                    If Not pdImages(g_CurrentImage).getActiveDIB Is Nothing Then
                        If pdImages(g_CurrentImage).getActiveDIB.getDIBColorDepth = 24 Then
                            FormMain.MnuLayerTransparency(3).Enabled = False
                        Else
                            If Not FormMain.MnuLayerTransparency(3).Enabled Then FormMain.MnuLayerTransparency(3).Enabled = True
                        End If
                    End If
                    
                    'Flatten is only available if one or more layers are visible
                    If pdImages(g_CurrentImage).getNumOfVisibleLayers > 0 Then
                        If Not FormMain.MnuLayer(15).Enabled Then FormMain.MnuLayer(15).Enabled = True
                    Else
                        FormMain.MnuLayer(15).Enabled = False
                    End If
                    
                    'Merge visible is only available if two or more layers are visible
                    If pdImages(g_CurrentImage).getNumOfVisibleLayers > 1 Then
                        If Not FormMain.MnuLayer(16).Enabled Then FormMain.MnuLayer(16).Enabled = True
                    Else
                        FormMain.MnuLayer(16).Enabled = False
                    End If
                    
                End If
                
                'If at least one layer is available, enable a number of layer options
                If Not FormMain.MnuLayer(7).Enabled Then FormMain.MnuLayer(7).Enabled = True
                If Not FormMain.MnuLayer(8).Enabled Then FormMain.MnuLayer(8).Enabled = True
                If Not FormMain.MnuLayer(11).Enabled Then FormMain.MnuLayer(11).Enabled = True
            
            Else
            
                'Most layer menus are disabled if an image does not contain layers.  PD isn't designed to allow 0-layer images,
                ' so this is primarily included as a fail-safe.
                FormMain.MnuLayer(1).Enabled = False
                FormMain.MnuLayer(3).Enabled = False
                FormMain.MnuLayer(4).Enabled = False
                FormMain.MnuLayer(5).Enabled = False
                FormMain.MnuLayer(7).Enabled = False
                FormMain.MnuLayer(8).Enabled = False
                FormMain.MnuLayer(9).Enabled = False
                FormMain.MnuLayer(11).Enabled = False
                FormMain.MnuLayer(13).Enabled = False
                FormMain.MnuLayer(15).Enabled = False
                FormMain.MnuLayer(16).Enabled = False
                MetaToggle tNonDestructiveFX, False
            
            End If
                    
        End If
        
        'Finally, synchronize various tool settings.  I've optimized this so that only the settings relative to the current tool
        ' are updated; others will be modified if/when the active tool is changed.
        Tool_Support.syncToolOptionsUIToCurrentLayer
        
        'Finally, if the histogram window is open, redraw it.  (This isn't needed at present, but could be useful in the future)
        'If FormHistogram.Visible And pdImages(g_CurrentImage).loadedSuccessfully Then
        '    FormHistogram.TallyHistogramValues
        '    FormHistogram.DrawHistogram
        'End If
        
    End If
        
    'Perform a special check for the image tabstrip.  Its appearance is contingent on a setting provided by the user, coupled
    ' with the number of presently open images.
    
    'A setting of 2 equates to index 2 in the menu, specifically "Never show image tabstrip".  Hide the tabstrip.
    If g_UserPreferences.GetPref_Long("Core", "Image Tabstrip Visibility", 1) = 2 Then
        g_WindowManager.SetWindowVisibility toolbar_ImageTabs.hWnd, False
    Else
        
        'A setting of 1 equates to index 1 in the menu, specifically "Show for 2+ loaded images".  Check image count and
        ' set visibility accordingly.
        If g_UserPreferences.GetPref_Long("Core", "Image Tabstrip Visibility", 1) = 1 Then
            
            If g_OpenImageCount > 1 Then
                g_WindowManager.SetWindowVisibility toolbar_ImageTabs.hWnd, True
            Else
                g_WindowManager.SetWindowVisibility toolbar_ImageTabs.hWnd, False
            End If
        
        'A setting of 0 equates to index 0 in the menu, specifically "always show tabstrip".
        Else
        
            If g_OpenImageCount > 0 Then
                g_WindowManager.SetWindowVisibility toolbar_ImageTabs.hWnd, True
            Else
                g_WindowManager.SetWindowVisibility toolbar_ImageTabs.hWnd, False
            End If
        
        End If
    
    End If
        
    'Perform a special check if 2 or more images are loaded; if that is the case, enable a few additional controls, like
    ' the "Next/Previous" Window menu items.
    If g_OpenImageCount >= 2 Then
        FormMain.MnuWindow(5).Enabled = True
        FormMain.MnuWindow(6).Enabled = True
    Else
        FormMain.MnuWindow(5).Enabled = False
        FormMain.MnuWindow(6).Enabled = False
    End If
        
    'Redraw the layer box
    toolbar_Layers.NotifyLayerChange
        
End Sub

'Some non-destructive actions need to synchronize *only* Undo/Redo buttons and menus (and their related counterparts, e.g. "Fade").
' To make these actions snappier, I have pulled all Undo/Redo UI sync code out of syncInterfaceToImage, and into this separate sub,
' which can be called on-demand as necessary.
'
'If the caller will be calling resetMenuIcons after using this function, make sure to pass the optional suspendAssociatedRedraws as TRUE
' to prevent unnecessary redraws.
Public Sub SyncUndoRedoInterfaceElements(Optional ByVal suspendAssociatedRedraws As Boolean = False)

    If g_OpenImageCount = 0 Then
    
        MetaToggle tUndo, False, True
        MetaToggle tRedo, False, True
        
        'Undo history is disabled when no images are loaded
        FormMain.MnuEdit(2).Enabled = False
        
        '"Repeat..." and "Fade..." in the Edit menu are disabled when no images are loaded
        FormMain.MnuEdit(4).Enabled = False
        FormMain.MnuEdit(5).Enabled = False
        toolbar_Toolbox.cmdFile(FILE_FADE).Enabled = False
        
        FormMain.MnuEdit(4).Caption = g_Language.TranslateMessage("Repeat")
        FormMain.MnuEdit(5).Caption = g_Language.TranslateMessage("Fade...")
        toolbar_Toolbox.cmdFile(FILE_FADE).AssignTooltip g_Language.TranslateMessage("Fade last action")
        
        'Redraw menu icons as requested
        If Not suspendAssociatedRedraws Then resetMenuIcons
    
    Else
    
        'Save is a bit funny, because if the image HAS been saved to file, we DISABLE the save button.
        MetaToggle tSave, Not pdImages(g_CurrentImage).getSaveState(pdSE_AnySave)
        
        'Undo, Redo, Repeat and Fade are all closely related
        If Not (pdImages(g_CurrentImage).undoManager Is Nothing) Then
        
            MetaToggle tUndo, pdImages(g_CurrentImage).undoManager.getUndoState, True
            MetaToggle tRedo, pdImages(g_CurrentImage).undoManager.getRedoState, True
            
            'Undo history is enabled if either Undo or Redo is active
            If pdImages(g_CurrentImage).undoManager.getUndoState Or pdImages(g_CurrentImage).undoManager.getRedoState Then
                FormMain.MnuEdit(2).Enabled = True
            Else
                FormMain.MnuEdit(2).Enabled = False
            End If
            
            '"Edit > Repeat..." and "Edit > Fade..." are also handled by the current image's undo manager (as it
            ' maintains the list of changes applied to the image, and links to copies of previous image state DIBs).
            Dim tmpDIB As pdDIB, tmpLayerIndex As Long, tmpActionName As String
            
            'See if the "Find last relevant layer action" function in the Undo manager returns TRUE or FALSE.  If it returns TRUE,
            ' enable both Repeat and Fade, and rename each menu caption so the user knows what is being repeated/faded.
            If pdImages(g_CurrentImage).undoManager.fillDIBWithLastUndoCopy(tmpDIB, tmpLayerIndex, tmpActionName, True) Then
                FormMain.MnuEdit(4).Caption = g_Language.TranslateMessage("Repeat: %1", g_Language.TranslateMessage(tmpActionName))
                FormMain.MnuEdit(5).Caption = g_Language.TranslateMessage("Fade: %1...", g_Language.TranslateMessage(tmpActionName))
                toolbar_Toolbox.cmdFile(FILE_FADE).AssignTooltip pdImages(g_CurrentImage).undoManager.getUndoProcessID, "Fade last action"
                
                toolbar_Toolbox.cmdFile(FILE_FADE).Enabled = True
                FormMain.MnuEdit(4).Enabled = True
                FormMain.MnuEdit(5).Enabled = True
            Else
                FormMain.MnuEdit(4).Caption = g_Language.TranslateMessage("Repeat")
                FormMain.MnuEdit(5).Caption = g_Language.TranslateMessage("Fade...")
                toolbar_Toolbox.cmdFile(FILE_FADE).AssignTooltip "Fade last action"
                
                toolbar_Toolbox.cmdFile(FILE_FADE).Enabled = False
                FormMain.MnuEdit(4).Enabled = False
                FormMain.MnuEdit(5).Enabled = False
            End If
            
            'Because these changes may modify menu captions, menu icons need to be reset (as they are tied to menu captions)
            If Not suspendAssociatedRedraws Then resetMenuIcons
        
        End If
    
    End If

End Sub

'metaToggle enables or disables a swath of controls related to a simple keyword (e.g. "Undo", which affects multiple menu items
' and toolbar buttons)
Public Sub MetaToggle(ByVal metaItem As metaInitializer, ByVal newState As Boolean, Optional ByVal suspendAssociatedRedraws As Boolean = False)
    
    Dim i As Long
    
    Select Case metaItem
            
        'Save (left-hand panel button AND menu item)
        Case tSave
            If FormMain.MnuFile(8).Enabled <> newState Then
                
                toolbar_Toolbox.cmdFile(FILE_SAVE).Enabled = newState
                
                FormMain.MnuFile(8).Enabled = newState
                
                'The File -> Revert menu is also tied to Save state (if the image has not been saved in its current state,
                ' we allow the user to revert to the last save state).
                FormMain.MnuFile(11).Enabled = newState
                
            End If
            
        'Save As (menu item only).  Note that Save Copy is also tied to Save As functionality, because they use the same rules
        ' for enablement (e.g. disabled if no images are loaded, always enabled otherwise)
        Case tSaveAs
            If FormMain.MnuFile(10).Enabled <> newState Then
                toolbar_Toolbox.cmdFile(FILE_SAVEAS_LAYERS).Enabled = newState
                toolbar_Toolbox.cmdFile(FILE_SAVEAS_FLAT).Enabled = newState
                
                FormMain.MnuFile(9).Enabled = newState
                FormMain.MnuFile(10).Enabled = newState
            End If
            
        'Close and Close All
        Case tClose
            If FormMain.MnuFile(5).Enabled <> newState Then
                FormMain.MnuFile(5).Enabled = newState
                FormMain.MnuFile(6).Enabled = newState
                toolbar_Toolbox.cmdFile(FILE_CLOSE).Enabled = newState
            End If
        
        'Undo (left-hand panel button AND menu item).  Undo toggles also control the "Fade last action" button, because that
        ' action requires Undo data to operate.
        Case tUndo
        
            If FormMain.MnuEdit(0).Enabled <> newState Then
                toolbar_Toolbox.cmdFile(FILE_UNDO).Enabled = newState
                FormMain.MnuEdit(0).Enabled = newState
            End If
            
            'If Undo is being enabled, change the text to match the relevant action that created this Undo file
            If newState Then
                toolbar_Toolbox.cmdFile(FILE_UNDO).AssignTooltip pdImages(g_CurrentImage).undoManager.getUndoProcessID, "Undo"
                FormMain.MnuEdit(0).Caption = g_Language.TranslateMessage("Undo:") & " " & g_Language.TranslateMessage(pdImages(g_CurrentImage).undoManager.getUndoProcessID) & vbTab & g_Language.TranslateMessage("Ctrl") & "+Z"
            Else
                toolbar_Toolbox.cmdFile(FILE_UNDO).AssignTooltip "Undo last action"
                FormMain.MnuEdit(0).Caption = g_Language.TranslateMessage("Undo") & vbTab & g_Language.TranslateMessage("Ctrl") & "+Z"
            End If
            
            'When changing menu text, icons must be reapplied.
            If Not suspendAssociatedRedraws Then resetMenuIcons
        
        'Redo (left-hand panel button AND menu item)
        Case tRedo
            If FormMain.MnuEdit(1).Enabled <> newState Then
                toolbar_Toolbox.cmdFile(FILE_REDO).Enabled = newState
                FormMain.MnuEdit(1).Enabled = newState
            End If
            
            'If Redo is being enabled, change the menu text to match the relevant action that created this Undo file
            If newState Then
                toolbar_Toolbox.cmdFile(FILE_REDO).AssignTooltip pdImages(g_CurrentImage).undoManager.getRedoProcessID, "Redo"
                FormMain.MnuEdit(1).Caption = g_Language.TranslateMessage("Redo:") & " " & g_Language.TranslateMessage(pdImages(g_CurrentImage).undoManager.getRedoProcessID) & vbTab & g_Language.TranslateMessage("Ctrl") & "+Y"
            Else
                toolbar_Toolbox.cmdFile(FILE_REDO).AssignTooltip "Redo previous action"
                FormMain.MnuEdit(1).Caption = g_Language.TranslateMessage("Redo") & vbTab & g_Language.TranslateMessage("Ctrl") & "+Y"
            End If
            
            'When changing menu text, icons must be reapplied.
            If Not suspendAssociatedRedraws Then resetMenuIcons
            
        'Copy (menu item only)
        Case tCopy
            If FormMain.MnuEdit(7).Enabled <> newState Then FormMain.MnuEdit(7).Enabled = newState
            If FormMain.MnuEdit(8).Enabled <> newState Then FormMain.MnuEdit(8).Enabled = newState
            If FormMain.MnuEdit(9).Enabled <> newState Then FormMain.MnuEdit(9).Enabled = newState
            If FormMain.MnuEdit(10).Enabled <> newState Then FormMain.MnuEdit(10).Enabled = newState
            If FormMain.MnuEdit(12).Enabled <> newState Then FormMain.MnuEdit(12).Enabled = newState
            
        'View (top-menu level)
        Case tView
            If FormMain.MnuView.Enabled <> newState Then FormMain.MnuView.Enabled = newState
        
        'ImageOps is all Image-related menu items; it enables/disables the Image, Layer, Select, Color, and Print menus
        Case tImageOps
            If FormMain.MnuImageTop.Enabled <> newState Then
                FormMain.MnuImageTop.Enabled = newState
                
                'Use this same command to disable other menus
                
                'File -> Print
                FormMain.MnuFile(15).Enabled = newState
                
                'Layer menu
                FormMain.MnuLayerTop.Enabled = newState
                
                'Select menu
                FormMain.MnuSelectTop.Enabled = newState
                
                'Adjustments menu
                FormMain.MnuAdjustmentsTop.Enabled = newState
                
                'Effects menu
                FormMain.MnuEffectsTop.Enabled = newState
                
            End If
            
        'Macro (within the Tools menu)
        Case tMacro
            If FormMain.mnuTool(3).Enabled <> newState Then
                FormMain.mnuTool(3).Enabled = newState
                FormMain.mnuTool(4).Enabled = newState
                FormMain.mnuTool(5).Enabled = newState
            End If
        
        'Selections in general
        Case tSelection
            
            'If selections are not active, clear all the selection value textboxes
            If Not newState Then
                For i = 0 To toolpanel_Selections.tudSel.Count - 1
                    toolpanel_Selections.tudSel(i).Value = 0
                Next i
            End If
            
            'Set selection text boxes to enable only when a selection is active.  Other selection controls can remain active
            ' even without a selection present; this allows the user to set certain parameters in advance, so when they actually
            ' draw a selection, it already has the attributes they want.
            For i = 0 To toolpanel_Selections.tudSel.Count - 1
                toolpanel_Selections.tudSel(i).Enabled = newState
            Next i
            
            'En/disable all selection menu items that rely on an existing selection to operate
            If FormMain.MnuSelect(2).Enabled <> newState Then
                
                'Select none, invert selection
                FormMain.MnuSelect(1).Enabled = newState
                FormMain.MnuSelect(2).Enabled = newState
                
                'Grow/shrink/border/feather/sharpen selection
                For i = 4 To 8
                    FormMain.MnuSelect(i).Enabled = newState
                Next i
                
                'Erase selected area
                FormMain.MnuSelect(10).Enabled = newState
                
                'Save selection
                FormMain.MnuSelect(13).Enabled = newState
                
                'Export selection top-level menu
                FormMain.MnuSelect(14).Enabled = newState
                
            End If
                                    
            'Selection enabling/disabling also affects the two Crop to Selection commands (one in the Image menu, one in the Layer menu)
            If FormMain.MnuImage(9).Enabled <> newState Then FormMain.MnuImage(9).Enabled = newState
            If FormMain.MnuLayer(9).Enabled <> newState Then FormMain.MnuLayer(9).Enabled = newState
            
        'Transformable selection controls specifically
        Case tSelectionTransform
        
            'Under certain circumstances, it is desirable to disable only the selection location boxes
            For i = 0 To toolpanel_Selections.tudSel.Count - 1
                If (Not newState) Then toolpanel_Selections.tudSel(i).Value = 0
                toolpanel_Selections.tudSel(i).Enabled = newState
            Next i
                
        'If the ExifTool plugin is not available, metadata will ALWAYS be disabled.  (We do not currently have a separate fallback for
        ' reading/browsing/writing metadata.)
        Case tMetadata
        
            If g_ExifToolEnabled Then
                If FormMain.MnuMetadata(0).Enabled <> newState Then FormMain.MnuMetadata(0).Enabled = newState
            Else
                If FormMain.MnuMetadata(0).Enabled Then FormMain.MnuMetadata(0).Enabled = False
            End If
        
        'GPS metadata is its own sub-category, and its activation is contigent upon an image having embedded GPS data
        Case tGPSMetadata
        
            If g_ExifToolEnabled Then
                If FormMain.MnuMetadata(3).Enabled <> newState Then FormMain.MnuMetadata(3).Enabled = newState
            Else
                If FormMain.MnuMetadata(3).Enabled Then FormMain.MnuMetadata(3).Enabled = False
            End If
        
        'Zoom controls not just the drop-down zoom box, but the zoom in, zoom out, and zoom fit buttons as well
        Case tZoom
            If FormMain.mainCanvas(0).getZoomDropDownReference().Enabled <> newState Then
                FormMain.mainCanvas(0).getZoomDropDownReference().Enabled = newState
                FormMain.mainCanvas(0).enableZoomIn newState
                FormMain.mainCanvas(0).enableZoomOut newState
                FormMain.mainCanvas(0).enableZoomFit newState
            End If
            
            'When disabling zoom controls, reset the zoom drop-down to 100%
            If Not newState Then FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex = g_Zoom.getZoom100Index
        
        'Various layer-related tools (move, etc) are exposed on the tool options dialog.  For consistency, we disable those UI elements
        ' when no images are loaded.
        Case tLayerTools
            
            'Because we're dealing with text up/downs, we need to set hard limits relative to the current image's size.
            ' I'm currently using the "rule of three" - max/min values are the current dimensions of the image, x3.
            Dim minLayerUIValue_Width As Long, maxLayerUIValue_Width As Long
            Dim minLayerUIValue_Height As Long, maxLayerUIValue_Height As Long
            
            If newState Then
                maxLayerUIValue_Width = pdImages(g_CurrentImage).Width * 3
                maxLayerUIValue_Height = pdImages(g_CurrentImage).Height * 3
            Else
                maxLayerUIValue_Width = 0
                maxLayerUIValue_Height = 0
            End If
            
            'Make sure width/height values are non-zero
            If maxLayerUIValue_Width = 0 Then maxLayerUIValue_Width = 1
            If maxLayerUIValue_Height = 0 Then maxLayerUIValue_Height = 1
            
            'Minimum values are simply the negative of the max values
            minLayerUIValue_Width = -1 * maxLayerUIValue_Width
            minLayerUIValue_Height = -1 * maxLayerUIValue_Height
            
            'Mark the tool engine as busy; this prevents control changes from triggering viewport redraws
            Tool_Support.setToolBusyState True
            
            'Enable/disable all UI elements as necessary
            For i = 0 To toolpanel_MoveSize.tudLayerMove.Count - 1
                If toolpanel_MoveSize.tudLayerMove(i).Enabled <> newState Then toolpanel_MoveSize.tudLayerMove(i).Enabled = newState
            Next i
            
            'Where relevant, also update control bounds
            If newState Then
            
                For i = 0 To toolpanel_MoveSize.tudLayerMove.Count - 1
                    
                    'Even-numbered indices correspond to width; odd-numbered to height
                    If i Mod 2 = 0 Then
                        
                        If toolpanel_MoveSize.tudLayerMove(i).Min <> minLayerUIValue_Width Then
                            toolpanel_MoveSize.tudLayerMove(i).Min = minLayerUIValue_Width
                            toolpanel_MoveSize.tudLayerMove(i).Max = maxLayerUIValue_Width
                        End If
                        
                    Else
                    
                        If toolpanel_MoveSize.tudLayerMove(i).Min <> minLayerUIValue_Height Then
                            toolpanel_MoveSize.tudLayerMove(i).Min = minLayerUIValue_Height
                            toolpanel_MoveSize.tudLayerMove(i).Max = maxLayerUIValue_Height
                        End If
                    
                    End If
                Next i
            
            End If
            
            'Free the tool engine
            Tool_Support.setToolBusyState False
        
        'Non-destructive FX are effects that the user can apply to a layer, without permanently modifying the layer
        Case tNonDestructiveFX
        
            If newState Then
                
                'Start by enabling all non-destructive FX controls
                For i = 0 To toolpanel_NDFX.sltQuickFix.Count - 1
                    If Not toolpanel_NDFX.sltQuickFix(i).Enabled Then toolpanel_NDFX.sltQuickFix(i).Enabled = True
                Next i
                
                'Quick fix buttons are only relevant if the current image has some non-destructive events applied
                If Not pdImages(g_CurrentImage).getActiveLayer Is Nothing Then
                
                    If pdImages(g_CurrentImage).getActiveLayer.getLayerNonDestructiveFXState() Then
                        
                        For i = 0 To toolpanel_NDFX.cmdQuickFix.Count - 1
                            toolpanel_NDFX.cmdQuickFix(i).Enabled = True
                        Next i
                        
                    Else
                        
                        For i = 0 To toolpanel_NDFX.cmdQuickFix.Count - 1
                            toolpanel_NDFX.cmdQuickFix(i).Enabled = False
                        Next i
                        
                    End If
                    
                End If
                
                'Disable automatic NDFX syncing, then update all sliders to match the current layer's values
                With toolpanel_NDFX
                    .setNDFXControlState False
                    
                    'The index of sltQuickFix controls aligns exactly with PD's constants for non-destructive effects.  This is by design.
                    If Not pdImages(g_CurrentImage).getActiveLayer Is Nothing Then
                        For i = 0 To .sltQuickFix.Count - 1
                            .sltQuickFix(i) = pdImages(g_CurrentImage).getActiveLayer.getLayerNonDestructiveFXValue(i)
                        Next i
                    Else
                        For i = 0 To .sltQuickFix.Count - 1
                            .sltQuickFix(i) = 0
                        Next i
                    End If
                    
                    .setNDFXControlState True
                End With
                
            Else
                
                'Disable all non-destructive FX controls
                For i = 0 To toolpanel_NDFX.sltQuickFix.Count - 1
                    If toolpanel_NDFX.sltQuickFix(i).Enabled Then toolpanel_NDFX.sltQuickFix(i).Enabled = False
                Next i
                
                For i = 0 To toolpanel_NDFX.cmdQuickFix.Count - 1
                    If toolpanel_NDFX.cmdQuickFix(i).Enabled Then toolpanel_NDFX.cmdQuickFix(i).Enabled = False
                Next i
                
            End If
            
    End Select
    
End Sub


'For best results, any modal form should be shown via this function.  This function will automatically center the form over the main window,
' while also properly assigning ownership so that the dialog is truly on top of any active windows.  It also handles deactivation of
' other windows (to prevent click-through), and dynamic top-most behavior to ensure that the program doesn't steal focus if the user switches
' to another program while a modal dialog is active.
Public Sub ShowPDDialog(ByRef dialogModality As FormShowConstants, ByRef dialogForm As Form, Optional ByVal doNotUnload As Boolean = False)

    On Error GoTo showPDDialogError
    
    g_ModalDialogActive = True
    
    'Start by loading the form and hiding it
    dialogForm.Visible = False
    
    'Store a reference to this dialog; if subsequent dialogs are loaded, this dialog will be given ownership over them
    If (currentDialogReference Is Nothing) Then
        
        'This is a regular modal dialog, and the main form should be its owner
        isSecondaryDialog = False
        Set currentDialogReference = dialogForm
                
    Else
    
        'We already have a reference to a modal dialog - that means a modal dialog is raising *another* modal dialog.  Give the previous
        ' modal dialog ownership over this new dialog!
        isSecondaryDialog = True
        
    End If
    
    'Retrieve and cache the hWnd; we need access to this even if the form is unloaded, so we can properly deregister it
    ' with the window manager.
    Dim dialogHwnd As Long
    dialogHwnd = dialogForm.hWnd
    
    'Get the rect of the main form, which we will use to calculate a center position
    Dim ownerRect As winRect
    GetWindowRect FormMain.hWnd, ownerRect
    
    'Determine the center of that rect
    Dim centerX As Long, centerY As Long
    centerX = ownerRect.x1 + (ownerRect.x2 - ownerRect.x1) \ 2
    centerY = ownerRect.y1 + (ownerRect.y2 - ownerRect.y1) \ 2
    
    'Get the rect of the child dialog
    Dim dialogRect As winRect
    GetWindowRect dialogHwnd, dialogRect
    
    'Determine an upper-left point for the dialog based on its size
    Dim newLeft As Long, newTop As Long
    newLeft = centerX - ((dialogRect.x2 - dialogRect.x1) \ 2)
    newTop = centerY - ((dialogRect.y2 - dialogRect.y1) \ 2)
    
    'Move the dialog into place, but do not repaint it (that will be handled in a moment by the .Show event)
    MoveWindow dialogHwnd, newLeft, newTop, dialogRect.x2 - dialogRect.x1, dialogRect.y2 - dialogRect.y1, 0
    
    'Use VB to actually display the dialog.  Note that the sub will pause here until the form is closed.
    dialogForm.Show dialogModality, FormMain    'getModalOwner()
    
    'Release our reference to this dialog
    If isSecondaryDialog Then
        isSecondaryDialog = False
    Else
        Set currentDialogReference = Nothing
    End If
    
    'If the form has not been unloaded, unload it now
    If (Not (dialogForm Is Nothing)) And (Not doNotUnload) Then
        Unload dialogForm
        Set dialogForm = Nothing
    End If
    
    g_ModalDialogActive = False
    
    Exit Sub
    
'For reasons I can't yet ascertain, this function will sometimes fail, claiming that a modal window is already active.  If that happens,
' we can just exit.
showPDDialogError:

    g_ModalDialogActive = False

End Sub

'When a modal dialog needs to be raised, we want to set its ownership to the top-most (relevant) window in the program, which may or may
' not be the main form.  This function should be called to determine the proper owner of any modal dialog box.
'
'If the caller knows in advance that a modal dialog is owned by another modal dialog (for example, a tool dialog displaying a color
' selection dialog), it can explicitly mark the assumeSecondaryDialog function as TRUE.
Public Function GetModalOwner(Optional ByVal assumeSecondaryDialog As Boolean = False) As Form

    'If a modal dialog is already active, it gets ownership over subsequent dialogs
    If isSecondaryDialog Or assumeSecondaryDialog Then
        Set GetModalOwner = currentDialogReference
        
    'No modal dialog is active, making this the only one.  Give the main form ownership.
    Else
        
        Set GetModalOwner = FormMain
        
    End If
    
End Function

'Return the system keyboard delay, in seconds.  This isn't an exact science because the delay is actually hardware dependent
' (e.g. the system returns a value from 0 to 3), but we can use a "good enough" approximation.
Public Function GetKeyboardDelay() As Double
    Dim keyDelayIndex As Long
    SystemParametersInfo SPI_GETKEYBOARDDELAY, 0, keyDelayIndex, 0
    GetKeyboardDelay = (keyDelayIndex + 1) * 0.25
End Function

'Return the system keyboard repeat rate, in seconds.  This isn't an exact science because the delay is actually hardware dependent
' (e.g. the system returns a value from 0 to 31), but we can use a "good enough" approximation.
Public Function GetKeyboardRepeatRate() As Double
    
    Dim keyRepeatIndex As Long
    SystemParametersInfo SPI_GETKEYBOARDSPEED, 0, keyRepeatIndex, 0
    
    'Per MSDN (http://msdn.microsoft.com/en-us/library/windows/desktop/ms724947%28v=vs.85%29.aspx#Input)
    ' "Retrieves the keyboard repeat-speed setting, which is a value in the range from 0 (approximately 2.5 repetitions per second)
    ' through 31 (approximately 30 repetitions per second). The actual repeat rates are hardware-dependent and may vary from a linear
    ' scale by as much as 20%. The pvParam parameter must point to a DWORD variable that receives the setting."
    '
    'The formula below mimics this behavior pretty closely (35 repetitions per second at the high-end, but identical at the low end)
    GetKeyboardRepeatRate = (400 - (keyRepeatIndex * 12)) / 1000
    
End Function

Public Sub ToggleImageTabstripAlignment(ByVal newAlignment As AlignConstants, Optional ByVal suppressInterfaceSync As Boolean = False, Optional ByVal suppressPrefUpdate As Boolean = False)
    
    'Reset the menu checkmarks
    Dim curMenuIndex As Long
    
    Select Case newAlignment
    
        Case vbAlignLeft
            curMenuIndex = 4
        
        Case vbAlignTop
            curMenuIndex = 5
        
        Case vbAlignRight
            curMenuIndex = 6
        
        Case vbAlignBottom
            curMenuIndex = 7
        
    End Select
    
    Dim i As Long
    For i = 4 To 7
        If i = curMenuIndex Then
            FormMain.MnuWindowTabstrip(i).Checked = True
        Else
            FormMain.MnuWindowTabstrip(i).Checked = False
        End If
    Next i
    
    'Write the preference out to file.
    If Not suppressPrefUpdate Then g_UserPreferences.SetPref_Long "Core", "Image Tabstrip Alignment", CLng(newAlignment)
    
    'Notify the window manager of the change
    g_WindowManager.SetImageTabstripAlignment newAlignment
    
    If Not suppressInterfaceSync Then
    
        '...and force the tabstrip to redraw itself (which it may not if the tabstrip's size hasn't changed, e.g. if Left and Right layout is toggled)
        toolbar_ImageTabs.forceRedraw
    
        'Refresh the current image viewport (which may be positioned differently due to the tabstrip moving)
        FormMain.refreshAllCanvases
        
    End If
    
End Sub

'The image tabstrip can set to appear under a variety of circumstances.  Use this sub to change the current setting; it will
' automatically handle syncing with the preferences file.
Public Sub ToggleImageTabstripVisibility(ByVal newSetting As Long, Optional ByVal suppressInterfaceSync As Boolean = False, Optional ByVal suppressPrefUpdate As Boolean = False)

    'Start by synchronizing menu checkmarks to the selected option
    Dim i As Long
    For i = 0 To 2
        If newSetting = i Then
            FormMain.MnuWindowTabstrip(i).Checked = True
        Else
            FormMain.MnuWindowTabstrip(i).Checked = False
        End If
    Next i

    'Write the matching preference out to file
    If Not suppressPrefUpdate Then g_UserPreferences.SetPref_Long "Core", "Image Tabstrip Visibility", newSetting
    
    If Not suppressInterfaceSync Then
    
        'Refresh the current image viewport (which may be positioned differently due to the tabstrip moving)
        FormMain.refreshAllCanvases
    
        'Synchronize the interface to match; note that this will handle showing/hiding the tabstrip based on the number of
        ' currently open images.
        SyncInterfaceToCurrentImage
        
    End If
    
    'If images are loaded, we may need to redraw their viewports because the available client area may have changed.
    If (g_NumOfImagesLoaded > 0) Then
        Viewport_Engine.Stage1_InitializeBuffer pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If

End Sub

'Toolbars can be dynamically shown/hidden by a variety of processes (e.g. clicking an entry in the Window menu, clicking the X in a
' toolbar's command box, etc).  All those operations should wrap this singular function.
Public Sub ToggleToolbarVisibility(ByVal whichToolbar As pdToolbarType)

    Select Case whichToolbar
    
        Case FILE_TOOLBOX
            FormMain.MnuWindowToolbox(0).Checked = Not FormMain.MnuWindowToolbox(0).Checked
            g_UserPreferences.SetPref_Boolean "Core", "Show File Toolbox", FormMain.MnuWindowToolbox(0).Checked
            g_WindowManager.SetWindowVisibility toolbar_Toolbox.hWnd, FormMain.MnuWindowToolbox(0).Checked
            
        Case TOOLS_TOOLBOX
            FormMain.MnuWindow(1).Checked = Not FormMain.MnuWindow(1).Checked
            g_UserPreferences.SetPref_Boolean "Core", "Show Selections Toolbox", FormMain.MnuWindow(1).Checked
            
            'Because this toolbox's visibility is also tied to the current tool, we wrap a different functions.  This function
            ' will show/hide the toolbox as necessary.
            toolbar_Toolbox.resetToolButtonStates
            
        Case LAYER_TOOLBOX
            FormMain.MnuWindow(2).Checked = Not FormMain.MnuWindow(2).Checked
            g_UserPreferences.SetPref_Boolean "Core", "Show Layers Toolbox", FormMain.MnuWindow(2).Checked
            g_WindowManager.SetWindowVisibility toolbar_Layers.hWnd, FormMain.MnuWindow(2).Checked
        
        Case DEBUG_TOOLBOX
            FormMain.MnuDevelopers(0).Checked = Not FormMain.MnuDevelopers(0).Checked
            g_UserPreferences.SetPref_Boolean "Core", "Show Debug Window", FormMain.MnuDevelopers(0).Checked
            g_WindowManager.SetWindowVisibility toolbar_Debug.hWnd, FormMain.MnuDevelopers(0).Checked
    
    End Select
    
    'Redraw the primary image viewport, as the available client area may have changed.
    If g_NumOfImagesLoaded > 0 Then FormMain.refreshAllCanvases
    
End Sub

Public Function FixDPI(ByVal pxMeasurement As Long) As Long

    'The first time this function is called, dpiRatio will be 0.  Calculate it.
    If dpiRatio = 0# Then
    
        'There are 1440 twips in one inch.  (Twips are resolution-independent.)  Use that knowledge to calculate DPI.
        dpiRatio = 1440 / TwipsPerPixelXFix
        
        'FYI: if the screen resolution is 96 dpi, this function will return the original pixel measurement, due to
        ' this calculation.
        dpiRatio = dpiRatio / 96
    
    End If
    
    FixDPI = CLng(dpiRatio * CDbl(pxMeasurement))
    
End Function

Public Function FixDPIFloat(ByVal pxMeasurement As Long) As Double

    'The first time this function is called, dpiRatio will be 0.  Calculate it.
    If dpiRatio = 0# Then
    
        'There are 1440 twips in one inch.  (Twips are resolution-independent.)  Use that knowledge to calculate DPI.
        dpiRatio = 1440 / TwipsPerPixelXFix
        
        'FYI: if the screen resolution is 96 dpi, this function will return the original pixel measurement, due to
        ' this calculation.
        dpiRatio = dpiRatio / 96
    
    End If
    
    FixDPIFloat = dpiRatio * CDbl(pxMeasurement)
    
End Function

'Fun fact: there are 15 twips per pixel at 96 DPI.  Not fun fact: at 200% DPI (e.g. 192 DPI), VB's internal
' TwipsPerPixelXFix will return 7, when actually we need the value 7.5.  This causes problems when resizing
' certain controls (like SmartCheckBox) because the size will actually come up short due to rounding errors!
' So whenever TwipsPerPixelXFix/Y is required, use these functions instead.
Public Function TwipsPerPixelXFix() As Double
    
    If m_CurrentSystemDPI = 0 Then
    
        If Screen.TwipsPerPixelX = 7 Then
            TwipsPerPixelXFix = 7.5
        Else
            TwipsPerPixelXFix = Screen.TwipsPerPixelX
        End If
        
    Else
        TwipsPerPixelXFix = 15# / m_CurrentSystemDPI
    End If

End Function

Public Function TwipsPerPixelYFix() As Double
    
    If m_CurrentSystemDPI = 0 Then
    
        If Screen.TwipsPerPixelY = 7 Then
            TwipsPerPixelYFix = 7.5
        Else
            TwipsPerPixelYFix = Screen.TwipsPerPixelY
        End If
        
    Else
        TwipsPerPixelYFix = 15# / m_CurrentSystemDPI
    End If

End Function

'ScaleX and ScaleY functions do not work when converting from pixels to twips, thanks to the 15 / 2 <> 7
' bug described above.  Instead of using ScaleX/Y functions, use these wrapper.
Public Function PXToTwipsX(ByVal srcPixelWidth As Long) As Long
    PXToTwipsX = srcPixelWidth * TwipsPerPixelXFix
End Function

Public Function PXToTwipsY(ByVal srcPixelHeight As Long) As Long
    PXToTwipsY = srcPixelHeight * TwipsPerPixelYFix
End Function

Public Sub DisplayWaitScreen(ByVal waitTitle As String, ByRef ownerForm As Form)
    
    FormWait.Visible = False
    
    FormWait.lblWaitTitle.Caption = waitTitle
    FormWait.lblWaitTitle.Visible = True
    FormWait.lblWaitTitle.Refresh
    
    Screen.MousePointer = vbHourglass
    
    FormWait.Show vbModeless, ownerForm
    FormWait.Refresh
    DoEvents
    
End Sub

Public Sub HideWaitScreen()
    Screen.MousePointer = vbDefault
    Unload FormWait
End Sub

'Given a wordwrap label with a set size, attempt to fit the label's text inside it
Public Sub FitWordwrapLabel(ByRef srcLabel As Label, ByRef srcForm As Form)

    'We will use a pdFont object to help us measure the label in question
    Dim tmpFont As pdFont
    Set tmpFont = New pdFont
    tmpFont.SetFontBold srcLabel.FontBold
    tmpFont.SetFontItalic srcLabel.FontItalic
    tmpFont.SetFontFace srcLabel.fontName
    tmpFont.SetFontSize srcLabel.FontSize
    tmpFont.CreateFontObject
    tmpFont.SetTextAlignment srcLabel.Alignment
    tmpFont.AttachToDC srcForm.hDC
    
    'Retrieve the height from the pdFont class
    Dim lblHeight As Long
    lblHeight = tmpFont.GetHeightOfWordwrapString(srcLabel.Caption, srcLabel.Width - 1)
    
    Dim curFontSize As Long
    curFontSize = srcLabel.FontSize
    
    'If the text is too tall, shrink the font until an acceptable size is found.  Note that the reported text value tends to be
    ' smaller than the space actually required.  I do not know why this happens.  To account for it, I cut a further 10% from
    ' the requested height, just to be safe.
    If (lblHeight > srcLabel.Height * 0.85) Then
            
        'Try shrinking the font size until an acceptable width is found
        Do While (lblHeight > srcLabel.Height * 0.85) And (curFontSize >= 8)
        
            curFontSize = curFontSize - 1
            
            tmpFont.ReleaseFromDC
            tmpFont.SetFontSize curFontSize
            tmpFont.CreateFontObject
            tmpFont.AttachToDC srcForm.hDC
            lblHeight = tmpFont.GetHeightOfWordwrapString(srcLabel.Caption, srcLabel.Width)
            
        Loop
            
    End If
    
    tmpFont.ReleaseFromDC
    
    'When an acceptable size is found, set it and exit.
    srcLabel.FontSize = curFontSize
    srcLabel.Refresh

End Sub

'Because VB6 apps look terrible on modern version of Windows, I do a bit of beautification to every form upon at load-time.
' This routine is nice because every form calls it at least once, so I can make centralized changes without having to rewrite
' code in every individual form.  This is also where run-time translation occurs.
Public Sub MakeFormPretty(ByRef tForm As Form, Optional ByVal useDoEvents As Boolean = False)
    
    If Not g_IsProgramRunning Then Exit Sub
    
    'Before doing anything else, make sure the form's default cursor is set to an arrow
    tForm.MouseIcon = LoadPicture("")
    tForm.MousePointer = 0

    'FORM STEP 1: Enumerate through every control on the form.  We will be making changes on-the-fly on a per-control basis.
    Dim eControl As Control
    
    For Each eControl In tForm.Controls
        
        'STEP 1: give all clickable controls a hand icon instead of the default pointer.
        ' (Note: this code will set all command buttons, scroll bars, option buttons, check boxes,
        ' list boxes, combo boxes, and file/directory/drive boxes to use the system hand cursor)
        If ((TypeOf eControl Is CommandButton) Or (TypeOf eControl Is HScrollBar) Or (TypeOf eControl Is VScrollBar) Or (TypeOf eControl Is OptionButton) Or (TypeOf eControl Is CheckBox) Or (TypeOf eControl Is ListBox) Or (TypeOf eControl Is ComboBox) Or (TypeOf eControl Is FileListBox) Or (TypeOf eControl Is DirListBox) Or (TypeOf eControl Is DriveListBox)) And (Not TypeOf eControl Is PictureBox) Then
            setHandCursor eControl
        End If
        
        'STEP 2: if the current system is Vista or later, and the user has requested modern typefaces via Edit -> Preferences,
        ' redraw all control fonts using Segoe UI.
        If ((TypeOf eControl Is TextBox) Or (TypeOf eControl Is CommandButton) Or (TypeOf eControl Is OptionButton) Or (TypeOf eControl Is CheckBox) Or (TypeOf eControl Is ListBox) Or (TypeOf eControl Is ComboBox) Or (TypeOf eControl Is FileListBox) Or (TypeOf eControl Is DirListBox) Or (TypeOf eControl Is DriveListBox) Or (TypeOf eControl Is Label)) And (Not TypeOf eControl Is PictureBox) Then
            eControl.fontName = g_InterfaceFont
        End If
        
        'TODO: integrate font handling directly into smartResize
        If (TypeOf eControl Is smartResize) Then
            eControl.Font.Name = g_InterfaceFont
        End If
        
        'PhotoDemon's custom controls now provide universal support for an UpdateAgainstCurrentTheme function.  This updates two things:
        ' 1) The control's visual appearance (to reflect any changes to visual themes)
        ' 2) The translated caption, or other text (to reflect any changes to the active language)
        If (TypeOf eControl Is smartOptionButton) Or (TypeOf eControl Is smartCheckBox) Then eControl.UpdateAgainstCurrentTheme
        If (TypeOf eControl Is buttonStrip) Or (TypeOf eControl Is buttonStripVertical) Then eControl.UpdateAgainstCurrentTheme
        If (TypeOf eControl Is pdButton) Or (TypeOf eControl Is pdButtonToolbox) Then eControl.UpdateAgainstCurrentTheme
        If (TypeOf eControl Is pdLabel) Or (TypeOf eControl Is pdHyperlink) Then eControl.UpdateAgainstCurrentTheme
        If (TypeOf eControl Is sliderTextCombo) Or (TypeOf eControl Is textUpDown) Then eControl.UpdateAgainstCurrentTheme
        If (TypeOf eControl Is pdComboBox) Or (TypeOf eControl Is pdComboBox_Font) Or (TypeOf eControl Is pdComboBox_Hatch) Then eControl.UpdateAgainstCurrentTheme
        If (TypeOf eControl Is pdCanvas) Or (TypeOf eControl Is pdScrollBar) Then eControl.UpdateAgainstCurrentTheme
        If (TypeOf eControl Is brushSelector) Or (TypeOf eControl Is gradientSelector) Or (TypeOf eControl Is penSelector) Then eControl.UpdateAgainstCurrentTheme
        If (TypeOf eControl Is pdColorVariants) Or (TypeOf eControl Is pdColorWheel) Then eControl.UpdateAgainstCurrentTheme
        
        'STEP 3: remove TabStop from each picture box.  They should never receive focus, but I often forget to change this
        ' at design-time.
        If (TypeOf eControl Is PictureBox) Then eControl.TabStop = False
        
        'STEP 4: make common control drop-down boxes display their full drop-down contents, without a scroll bar.
        '         (This behavior requires a manifest, so useless in the IDE.)
        If (TypeOf eControl Is ComboBox) Then SendMessage eControl.hWnd, CB_SETMINVISIBLE, CLng(eControl.ListCount), ByVal 0&
        
        'Optionally, DoEvents can be called after each change.  This slows the process, but it allows external progress
        ' bars to be automatically refreshed.
        If useDoEvents Then DoEvents
        
    Next
    
    'FORM STEP 2: translate the form (and all controls on it)
    If g_Language.translationActive And tForm.Enabled Then
        g_Language.applyTranslations tForm, useDoEvents
    End If
    
    'Refresh all non-MDI forms after making the changes above
    If tForm.Name <> "FormMain" Then
        tForm.Refresh
    Else
        'The main from is a bit different - if it has been translated or changed, it needs menu icons reassigned.
        If FormMain.Visible Then applyAllMenuIcons
    End If
    
End Sub

'Used to enable font smoothing if currently disabled.
Public Sub HandleClearType(ByVal startingProgram As Boolean)
    
    'At start-up, activate ClearType.  At shutdown, restore the original setting (as necessary).
    If startingProgram Then
    
        hadToChangeSmoothing = 0
    
        'Get current font smoothing setting
        Dim pv As Long
        SystemParametersInfo SPI_GETFONTSMOOTHING, 0, pv, 0
        
        'If font smoothing is disabled, mark it
        If pv = 0 Then hadToChangeSmoothing = 2
        
        'If font smoothing is enabled but set to Standard instead of ClearType, mark it
        If pv <> 0 Then
            SystemParametersInfo SPI_GETFONTSMOOTHINGTYPE, 0, pv, 0
            If pv = SmoothingStandardType Then hadToChangeSmoothing = 1
        End If
        
        Select Case hadToChangeSmoothing
        
            'ClearType is enabled, no changes necessary
            Case 0
            
            'Standard smoothing is enabled; switch it to ClearType for the duration of the program
            Case 1
                SystemParametersInfo SPI_SETFONTSMOOTHINGTYPE, 0, ByVal SmoothingClearType, 0
                
            'No smoothing is enabled; turn it on and activate ClearType for the duration of the program
            Case 2
                SystemParametersInfo SPI_SETFONTSMOOTHING, 1, pv, 0
                SystemParametersInfo SPI_SETFONTSMOOTHINGTYPE, 0, ByVal SmoothingClearType, 0
            
        End Select
    
    Else
        
        Select Case hadToChangeSmoothing
        
            'ClearType was enabled, no action necessary
            Case 0
            
            'Standard smoothing was enabled; restore it now
            Case 1
                SystemParametersInfo SPI_SETFONTSMOOTHINGTYPE, 0, ByVal SmoothingStandardType, 0
                
            'No smoothing was enabled; restore that setting now
            Case 2
                SystemParametersInfo SPI_SETFONTSMOOTHING, 0, pv, 0
                SystemParametersInfo SPI_SETFONTSMOOTHINGTYPE, 0, ByVal SmoothingNone, 0
        
        End Select
    
    End If
    
End Sub

'When a themed form is unloaded, it may be desirable to release certain changes made to it - or in our case, unsubclass it.
' This function should be called when any themed form is unloaded.
Public Sub ReleaseFormTheming(ByRef tForm As Object)
    'g_Themer.releaseContainerSubclass tForm.hWnd
    Set tForm = Nothing
End Sub

'Given a pdImage object, generate an appropriate caption for the main PhotoDemon window.
Private Function GetWindowCaption(ByRef srcImage As pdImage) As String

    Dim captionBase As String
    captionBase = ""
    
    'Start by seeing if this image has some kind of filename.  This field should always be populated by the load function,
    ' but better safe than sorry!
    If Len(srcImage.originalFileNameAndExtension) <> 0 Then
    
        'This image has a filename!  Next, check the user's preference for long or short window captions
        If g_UserPreferences.GetPref_Long("Interface", "Window Caption Length", 0) = 0 Then
            
            'The user prefers short captions.  Use just the filename and extension (no folders ) as the base.
            captionBase = srcImage.originalFileNameAndExtension
        Else
        
            'The user prefers long captions.  Make sure this image has such a location; if they do not, fallback
            ' and use just the filename.
            If Len(srcImage.locationOnDisk) <> 0 Then
                captionBase = srcImage.locationOnDisk
            Else
                captionBase = srcImage.originalFileNameAndExtension
            End If
            
        End If
    
    'This image does not have a filename.  Assign it a default title.
    Else
        captionBase = g_Language.TranslateMessage("[untitled image]")
    End If
    
    'Append the current PhotoDemon version number and exit
    GetWindowCaption = captionBase & "  -  " & getPhotoDemonNameAndVersion()

End Function


'Display the specified size in the main form's status bar
Public Sub DisplaySize(ByRef srcImage As pdImage)
    
    If Not (srcImage Is Nothing) Then
        FormMain.mainCanvas(0).displayImageSize srcImage
    End If
    
    'Size is only displayed when it is changed, so if any controls have a maximum value linked to the size of the image,
    ' now is an excellent time to update them.
    Dim limitingSize As Long
    
    If srcImage.Width < srcImage.Height Then
        limitingSize = srcImage.Width
        toolpanel_Selections.sltCornerRounding.Max = srcImage.Width
        toolpanel_Selections.sltSelectionLineWidth.Max = srcImage.Height
    Else
        limitingSize = srcImage.Height
        toolpanel_Selections.sltCornerRounding.Max = srcImage.Height
        toolpanel_Selections.sltSelectionLineWidth.Max = srcImage.Width
    End If
    
    Dim i As Long
    For i = 0 To toolpanel_Selections.sltSelectionBorder.Count - 1
        toolpanel_Selections.sltSelectionBorder(i).Max = limitingSize
    Next i
    
End Sub

'This wrapper is used in place of the standard MsgBox function.  At present it's just a wrapper around MsgBox, but
' in the future I may replace the dialog function with something custom.
Public Function PDMsgBox(ByVal pMessage As String, ByVal pButtons As VbMsgBoxStyle, ByVal pTitle As String, ParamArray ExtraText() As Variant) As VbMsgBoxResult

    Dim newMessage As String, newTitle As String
    newMessage = pMessage
    newTitle = pTitle

    'All messages are translatable, but we don't want to translate them if the translation object isn't ready yet
    If (Not (g_Language Is Nothing)) Then
        If g_Language.readyToTranslate Then
            If g_Language.translationActive Then
                newMessage = g_Language.TranslateMessage(pMessage)
                newTitle = g_Language.TranslateMessage(pTitle)
            End If
        End If
    End If
    
    'Once the message is translated, we can add back in any optional parameters
    If UBound(ExtraText) >= LBound(ExtraText) Then
    
        Dim i As Long
        For i = LBound(ExtraText) To UBound(ExtraText)
            newMessage = Replace$(newMessage, "%" & i + 1, CStr(ExtraText(i)))
        Next i
    
    End If

    PDMsgBox = MsgBox(newMessage, pButtons, newTitle)

End Function

'This popular function is used to display a message in the main form's status bar.
' INPUTS:
' 1) the message to be displayed (mString)
' *2) any values that must be calculated at run-time, which are labeled in the message string by "%n", e.g. "Download time remaining: %1", timeRemaining
Public Sub Message(ByVal mString As String, ParamArray ExtraText() As Variant)

    Dim i As Long

    'Before doing anything else, check for a duplicate message request.  They are automatically ignored.
    Dim tmpDupeCheckString As String
    tmpDupeCheckString = mString
    
    If UBound(ExtraText) >= LBound(ExtraText) Then
        
        For i = LBound(ExtraText) To UBound(ExtraText)
            If StrComp(UCase(ExtraText(i)), "DONOTLOG", vbBinaryCompare) <> 0 Then
                tmpDupeCheckString = Replace$(tmpDupeCheckString, "%" & CStr(i + 1), CStr(ExtraText(i)))
            End If
        Next i
        
    End If
    
    'If the message request is for a novel string (e.g. one that differs from the previous message request), display it.
    ' Otherwise, exit now.
    If StrComp(m_PrevMessage, tmpDupeCheckString, vbBinaryCompare) <> 0 Then
        
        'In debug mode, mirror the message output to PD's central Debugger.  Note that this behavior can be overridden by
        ' supplying the string "DONOTLOG" as the final entry in the ParamArray.
        #If DEBUGMODE = 1 Then
            If UBound(ExtraText) < LBound(ExtraText) Then
                pdDebug.LogAction tmpDupeCheckString, PDM_USER_MESSAGE
            Else
            
                'Check the last param passed.  If it's the string "DONOTLOG", do not log this entry.  (PD sometimes uses this
                ' to avoid logging useless data, like layer hover events or download updates.)
                If StrComp(UCase$(CStr(ExtraText(UBound(ExtraText)))), "DONOTLOG", vbBinaryCompare) <> 0 Then
                    pdDebug.LogAction tmpDupeCheckString, PDM_USER_MESSAGE
                End If
            
            End If
        #End If
        
        'Cache the contents of the untranslated message, so we can check for duplicates on the next message request
        m_PrevMessage = tmpDupeCheckString
                
        Dim newString As String
        newString = mString
    
        'All messages are translatable, but we don't want to translate them if the translation object isn't ready yet.
        ' This only happens for a few messages when the program is first loaded, and at some point, I will eventually getting
        ' around to removing them entirely.
        If (Not (g_Language Is Nothing)) Then
            If g_Language.readyToTranslate Then
                If g_Language.translationActive Then newString = g_Language.TranslateMessage(mString)
            End If
        End If
        
        'Once the message is translated, we can add back in any optional text supplied in the ParamArray
        If UBound(ExtraText) >= LBound(ExtraText) Then
        
            For i = LBound(ExtraText) To UBound(ExtraText)
                newString = Replace$(newString, "%" & i + 1, CStr(ExtraText(i)))
            Next i
        
        End If
        
        'While macros are active, append a "Recording" message to help orient the user
        If MacroStatus = MacroSTART Then newString = newString & " {-" & g_Language.TranslateMessage("Recording") & "-}"
        
        'Post the message to the screen
        If MacroStatus <> MacroBATCH Then
            
            'If the window is disabled, it will not refresh when new messages are posted.  We can work around this limitation
            ' by toggling the state immediately prior to updating, then restoring the state afterward
            
            'Make a backup of the current form state
            Dim curMainFormState As Boolean, curMainFormCursor As Long
            curMainFormState = FormMain.Enabled
            curMainFormCursor = Screen.MousePointer
            
            'Display the message
            If (Not curMainFormState) And (Not g_ModalDialogActive) Then FormMain.Enabled = True
            FormMain.mainCanvas(0).displayCanvasMessage newString
            If (Not curMainFormState) Then Replacement_DoEvents FormMain.mainCanvas(0).hWnd
            
            'Restore original form state (only relevant if we had to change state to display the message)
            If (Not curMainFormState) And (Not g_ModalDialogActive) Then
                FormMain.Enabled = False
                Screen.MousePointer = curMainFormCursor
            End If
            
        End If
        
        'Update the global "previous message" string, so external functions can access it.
        g_LastPostedMessage = newString
        
    End If
    
End Sub

'Pass AutoSelectText a text box and it will select all text currently in the text box
Public Function AutoSelectText(ByRef tBox As TextBox)
    If Not tBox.Visible Then Exit Function
    If Not tBox.Enabled Then Exit Function
    tBox.SetFocus
    tBox.SelStart = 0
    tBox.SelLength = Len(tBox.Text)
End Function

'When the mouse is moved outside the primary image, clear the image coordinates display
Public Sub ClearImageCoordinatesDisplay()
    FormMain.mainCanvas(0).displayCanvasCoordinates 0, 0, True
End Sub

'Populate the passed combo box with options related to distort filter edge-handle options.  Also, select the specified method by default.
Public Sub PopDistortEdgeBox(ByRef cmbEdges As ComboBox, Optional ByVal defaultEdgeMethod As EDGE_OPERATOR)

    cmbEdges.Clear
    cmbEdges.AddItem " clamp them to the nearest available pixel"
    cmbEdges.AddItem " reflect them across the nearest edge"
    cmbEdges.AddItem " wrap them around the image"
    cmbEdges.AddItem " erase them"
    cmbEdges.AddItem " ignore them"
    cmbEdges.ListIndex = defaultEdgeMethod
    
End Sub

'Populate the passed button strip with options related to convolution kernel shape.  The caller can also specify which method they
' want set as the default.
Public Sub PopKernelShapeButtonStrip(ByRef srcBTS As buttonStrip, Optional ByVal defaultShape As PD_PIXEL_REGION_SHAPE = PDPRS_Rectangle)
    
    srcBTS.AddItem "Square", 0
    srcBTS.AddItem "Circle", 1
    srcBTS.ListIndex = defaultShape
    
End Sub

'Return the width (and below, height) of a string, in pixels, according to the font assigned to fontContainerDC
Public Function GetPixelWidthOfString(ByVal srcString As String, ByVal fontContainerDC As Long) As Long
    Dim txtSize As POINTAPI
    GetTextExtentPoint32 fontContainerDC, srcString, Len(srcString), txtSize
    GetPixelWidthOfString = txtSize.x
End Function

Public Function GetPixelHeightOfString(ByVal srcString As String, ByVal fontContainerDC As Long) As Long
    Dim txtSize As POINTAPI
    GetTextExtentPoint32 fontContainerDC, srcString, Len(srcString), txtSize
    GetPixelHeightOfString = txtSize.y
End Function

'Use whenever you want the user to not be allowed to interact with the primary PD window.  Make sure that call "enableUserInput", below,
' when you are done processing!
Public Sub DisableUserInput()

    'Set the "input disabled" flag, which individual functions can use to modify their own behavior
    g_DisableUserInput = True
    
    'Forcibly disable the main form
    FormMain.Enabled = False

End Sub

'Sister function to "disableUserInput", above
Public Sub EnableUserInput()
    
    'Start a countdown timer on the main form.  When it terminates, user input will be restored.  A timer is required because
    ' we need a certain amount of "dead time" to elapse between double-clicks on a top-level dialog (like a common dialog)
    ' which may be incorrectly passed through to the main form.  (I know, this seems like a ridiculous solution, but I tried
    ' a thousand others before settling on this.  It's the least of many evils.)
    FormMain.tmrCountdown.Enabled = True
    
    'Re-enable the main form
    FormMain.Enabled = True

End Sub

'Given a combo box, populate it with all currently supported blend modes
Public Sub PopulateBlendModeComboBox(ByRef dstCombo As pdComboBox, Optional ByVal blendIndex As LAYER_BLENDMODE = BL_NORMAL)
    
    dstCombo.Clear
    
    dstCombo.AddItem "Normal", 0, True
    dstCombo.AddItem "Darken"
    dstCombo.AddItem "Multiply"
    dstCombo.AddItem "Color burn"
    dstCombo.AddItem "Linear burn", , True
    dstCombo.AddItem "Lighten"
    dstCombo.AddItem "Screen"
    dstCombo.AddItem "Color dodge"
    dstCombo.AddItem "Linear dodge", , True
    dstCombo.AddItem "Overlay"
    dstCombo.AddItem "Soft light"
    dstCombo.AddItem "Hard light"
    dstCombo.AddItem "Vivid light"
    dstCombo.AddItem "Linear light"
    dstCombo.AddItem "Pin light"
    dstCombo.AddItem "Hard mix", , True
    dstCombo.AddItem "Difference"
    dstCombo.AddItem "Exclusion"
    dstCombo.AddItem "Subtract"
    dstCombo.AddItem "Divide", , True
    dstCombo.AddItem "Hue"
    dstCombo.AddItem "Saturation"
    dstCombo.AddItem "Color"
    dstCombo.AddItem "Luminosity", , True
    dstCombo.AddItem "Grain extract"
    dstCombo.AddItem "Grain merge"
    
    dstCombo.ListIndex = blendIndex
    
End Sub

'In an attempt to better serve high-DPI users, some of PD's stock UI icons are now generated at runtime.
' Note that the requested size is in PIXELS, so it is up to the caller to determine the proper size IN PIXELS of
' any requested UI elements.  This value will be automatically scaled to the current DPI, so make sure the passed
' pixel value is relevant to 100% DPI only (96 DPI).
Public Function GetRuntimeUIDIB(ByVal dibType As PD_RUNTIME_UI_DIB, Optional ByVal dibSize As Long = 16, Optional ByVal dibPadding As Long = 0, Optional ByVal BackColor As Long = 0) As pdDIB

    'Adjust the dib size and padding to account for DPI
    dibSize = FixDPI(dibSize)
    dibPadding = FixDPI(dibPadding)

    'Create the target DIB
    Set GetRuntimeUIDIB = New pdDIB
    GetRuntimeUIDIB.createBlank dibSize, dibSize, 32, BackColor, 0
    GetRuntimeUIDIB.setInitialAlphaPremultiplicationState True
    
    Dim paintColor As Long
    
    'Dynamically create the requested icon
    Select Case dibType
    
        'Red, green, and blue channel icons are all created similarly.
        Case PDRUID_CHANNEL_RED, PDRUID_CHANNEL_GREEN, PDRUID_CHANNEL_BLUE
            
            If dibType = PDRUID_CHANNEL_RED Then
                paintColor = g_Themer.GetThemeColor(PDTC_CHANNEL_RED)
            ElseIf dibType = PDRUID_CHANNEL_GREEN Then
                paintColor = g_Themer.GetThemeColor(PDTC_CHANNEL_GREEN)
            ElseIf dibType = PDRUID_CHANNEL_BLUE Then
                paintColor = g_Themer.GetThemeColor(PDTC_CHANNEL_BLUE)
            End If
            
            'Draw a colored circle just within the bounds of the DIB
            GDI_Plus.GDIPlusFillEllipseToDC GetRuntimeUIDIB.getDIBDC, dibPadding, dibPadding, dibSize - dibPadding * 2, dibSize - dibPadding * 2, paintColor, True
        
        'The RGB DIB is a triad of the individual RGB circles
        Case PDRUID_CHANNEL_RGB
        
            'Draw the red, green, and blue circles, with slight overlap toward the middle
            Dim circleSize As Long
            circleSize = (dibSize - dibPadding) * 0.55
            
            GDI_Plus.GDIPlusFillEllipseToDC GetRuntimeUIDIB.getDIBDC, dibSize - circleSize - dibPadding, dibSize - circleSize - dibPadding, circleSize, circleSize, g_Themer.GetThemeColor(PDTC_CHANNEL_BLUE), True, 210
            GDI_Plus.GDIPlusFillEllipseToDC GetRuntimeUIDIB.getDIBDC, dibPadding, dibSize - circleSize - dibPadding, circleSize, circleSize, g_Themer.GetThemeColor(PDTC_CHANNEL_GREEN), True, 210
            GDI_Plus.GDIPlusFillEllipseToDC GetRuntimeUIDIB.getDIBDC, dibSize \ 2 - circleSize \ 2, dibPadding, circleSize, circleSize, g_Themer.GetThemeColor(PDTC_CHANNEL_RED), True, 210
    
    End Select
    
    'If the user requested any padding, apply it now
    If dibPadding > 0 Then padDIB GetRuntimeUIDIB, dibPadding
    
End Function

'New test functions to (hopefully) help address high-DPI issues where VB's internal scale properties report false values
Public Function APIWidth(ByVal srcHwnd As Long) As Long
    Dim tmpRect As winRect
    GetWindowRect srcHwnd, tmpRect
    APIWidth = tmpRect.x2 - tmpRect.x1
End Function

Public Function APIHeight(ByVal srcHwnd As Long) As Long
    Dim tmpRect As winRect
    GetWindowRect srcHwnd, tmpRect
    APIHeight = tmpRect.y2 - tmpRect.y1
End Function

'Program shutting down?  Call this function to release any interface-related resources stored by this module
Public Sub ReleaseResources()
    Set currentDialogReference = Nothing
End Sub
