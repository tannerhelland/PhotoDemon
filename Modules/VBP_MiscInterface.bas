Attribute VB_Name = "Interface"
'***************************************************************************
'Miscellaneous Functions Related to the User Interface
'Copyright ©2001-2014 by Tanner Helland
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
     tNonDestructiveFX
End Enum

#If False Then
    Private Const tSave = 0, tSaveAs = 0, tClose = 0, tUndo = 0, tRedo = 0, tCopy = 0, tPaste = 0, tView = 0, tImageOps = 0
    Private Const tMetadata = 0, tGPSMetadata = 0, tMacro = 0, tSelection = 0, tSelectionTransform = 0, tZoom = 0, tNonDestructiveFX = 0
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

'Previously, various PD functions had to manually enable/disable button and menu state based on their actions.  This is no longer necessary.
' Simply call this function whenever an action has done something that will potentially affect the interface, and this function will iterate
' through all potential image/interface interactions, dis/enabling buttons and menus as necessary.
Public Sub syncInterfaceToCurrentImage()
    
    Dim i As Long
    
    'Interface dis/enabling falls into two rough categories: stuff that changes based on the current image (e.g. Undo), and stuff that changes
    ' based on the *total* number of available images (e.g. visibility of the Effects menu).
    
    'Start by breaking our interface decisions into two broad categories: "no images are loaded" and "one or more images are loaded".
    
    'If no images are loaded, we can disable a whole swath of controls
    If g_OpenImageCount = 0 Then
    
        metaToggle tSave, False
        metaToggle tSaveAs, False
        metaToggle tClose, False
        metaToggle tUndo, False, True
        metaToggle tRedo, False, True
        metaToggle tCopy, False
        metaToggle tView, False
        metaToggle tImageOps, False
        metaToggle tSelection, False
        metaToggle tMacro, False
        metaToggle tZoom, False
        metaToggle tNonDestructiveFX, False
        
        'Undo history is disabled when no images are loaded
        FormMain.MnuEdit(2).Enabled = False
        
        '"Repeat..." and "Fade..." in the Edit menu are disabled when no images are loaded
        FormMain.MnuEdit(4).Enabled = False
        FormMain.MnuEdit(5).Enabled = False
        
        FormMain.MnuEdit(4).Caption = g_Language.TranslateMessage("Repeat")
        FormMain.MnuEdit(5).Caption = g_Language.TranslateMessage("Fade...")
        
        'All relevant menu icons can now be redrawn.  (This must be redone after menu captions change, as icons are associated
        ' with captions.)
        resetMenuIcons
        
        '"Paste as new layer" is disabled when no images are loaded (but "Paste as new image" remains active)
        FormMain.MnuEdit(8).Enabled = False
                        
        Message "Please load or import an image to begin editing."
        
        'Assign a generic caption to the main window
        FormMain.Caption = getPhotoDemonNameAndVersion()
        
        'Erase the main viewport's status bar
        FormMain.mainCanvas(0).displayImageSize Nothing, True
        FormMain.mainCanvas(0).drawStatusBarIcons False
        
        'Because dynamic icons are enabled, restore the main program icon and clear the custom image icon cache
        destroyAllIcons
        setNewTaskbarIcon origIcon32, FormMain.hWnd
        setNewAppIcon origIcon16, origIcon32
        
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
                
        'Erase any remaining viewport buffer
        'eraseViewportBuffers
    
    'If one or more images are loaded, our job is trickier.  Some controls (such as Copy to Clipboard) are enabled no matter what,
    ' while others (Undo and Redo) are only enabled if the current image requires it.
    Else
        
        If Not pdImages(g_CurrentImage) Is Nothing Then
        
            'Start by enabling actions that are always available if one or more images are loaded.
            metaToggle tSaveAs, True
            metaToggle tClose, True
            metaToggle tCopy, True
            
            metaToggle tView, True
            metaToggle tZoom, True
            metaToggle tImageOps, True
            metaToggle tMacro, True
            
            'Paste as new layer is always available if one (or more) images are loaded
            If Not FormMain.MnuEdit(9).Enabled Then FormMain.MnuEdit(9).Enabled = True
            
            'Display this image's path in the title bar.
            FormMain.Caption = getWindowCaption(pdImages(g_CurrentImage))
            
            'Draw icons onto the main viewport's status bar
            FormMain.mainCanvas(0).drawStatusBarIcons True
            
            'Next, attempt to enable controls whose state depends on the current image - e.g. "Save", which is only enabled if
            ' the image has not already been saved in its current state.
            
            'Note that all of these functions rely on the g_CurrentImage value to function.
            
            'Save is a bit funny, because if the image HAS been saved to file, we DISABLE the save button.
            metaToggle tSave, Not pdImages(g_CurrentImage).getSaveState(pdSE_AnySave)
            
            'Undo, Redo, Repeat and Fade are all closely related
            If Not (pdImages(g_CurrentImage).undoManager Is Nothing) Then
            
                metaToggle tUndo, pdImages(g_CurrentImage).undoManager.getUndoState, True
                metaToggle tRedo, pdImages(g_CurrentImage).undoManager.getRedoState, True
                
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
                    
                    FormMain.MnuEdit(4).Enabled = True
                    FormMain.MnuEdit(5).Enabled = True
                Else
                    FormMain.MnuEdit(4).Caption = g_Language.TranslateMessage("Repeat")
                    FormMain.MnuEdit(5).Caption = g_Language.TranslateMessage("Fade...")
                    
                    FormMain.MnuEdit(4).Enabled = False
                    FormMain.MnuEdit(5).Enabled = False
                End If
                
                'Because these changes may modify menu captions, menu icons need to be reset (as they are tied to menu captions)
                resetMenuIcons
            
            End If
            
            'Determine whether metadata is present, and dis/enable metadata menu items accordingly
            If Not pdImages(g_CurrentImage).imgMetadata Is Nothing Then
                metaToggle tMetadata, pdImages(g_CurrentImage).imgMetadata.hasXMLMetadata
                metaToggle tGPSMetadata, pdImages(g_CurrentImage).imgMetadata.hasGPSMetadata()
            Else
                metaToggle tMetadata, False
                metaToggle tGPSMetadata, False
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
                metaToggle tSelection, True
                metaToggle tSelectionTransform, pdImages(g_CurrentImage).mainSelection.isTransformable
                syncTextToCurrentSelection g_CurrentImage
            Else
                metaToggle tSelection, False
                metaToggle tSelectionTransform, False
            End If
            
            'Update all layer menus; some will be disabled depending on just how many layers are available, how many layers
            ' are visible, and other criteria.
            If pdImages(g_CurrentImage).getNumOfLayers > 0 Then
            
                'If non-destructive FX are active on the current layer, update the non-destructive tools to match
                If Not pdImages(g_CurrentImage).getActiveLayer Is Nothing Then
                    metaToggle tNonDestructiveFX, True
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
                    FormMain.MnuLayer(12).Enabled = False
                    
                    'Merge visible
                    FormMain.MnuLayer(13).Enabled = False
                    
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
                        If Not FormMain.MnuLayer(12).Enabled Then FormMain.MnuLayer(12).Enabled = True
                    Else
                        FormMain.MnuLayer(12).Enabled = False
                    End If
                    
                    'Merge visible is only available if two or more layers are visible
                    If pdImages(g_CurrentImage).getNumOfVisibleLayers > 1 Then
                        If Not FormMain.MnuLayer(13).Enabled Then FormMain.MnuLayer(13).Enabled = True
                    Else
                        FormMain.MnuLayer(13).Enabled = False
                    End If
                    
                End If
                
                'If at least one layer is available, enable a number of layer options
                If Not FormMain.MnuLayer(7).Enabled Then FormMain.MnuLayer(7).Enabled = True
                If Not FormMain.MnuLayer(8).Enabled Then FormMain.MnuLayer(8).Enabled = True
                If Not FormMain.MnuLayer(10).Enabled Then FormMain.MnuLayer(10).Enabled = True
            
            Else
            
                'Most layer menus are disabled if an image does not contain layers.  PD isn't setup to allow 0-layer images,
                ' so this is primarily included as a fail-safe.
                FormMain.MnuLayer(1).Enabled = False
                FormMain.MnuLayer(3).Enabled = False
                FormMain.MnuLayer(4).Enabled = False
                FormMain.MnuLayer(5).Enabled = False
                FormMain.MnuLayer(7).Enabled = False
                FormMain.MnuLayer(8).Enabled = False
                FormMain.MnuLayer(10).Enabled = False
                FormMain.MnuLayer(12).Enabled = False
                FormMain.MnuLayer(13).Enabled = False
                metaToggle tNonDestructiveFX, False
            
            End If
                    
        End If
        
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
        g_WindowManager.setWindowVisibility toolbar_ImageTabs.hWnd, False
    Else
        
        'A setting of 1 equates to index 1 in the menu, specifically "Show for 2+ loaded images".  Check image count and
        ' set visibility accordingly.
        If g_UserPreferences.GetPref_Long("Core", "Image Tabstrip Visibility", 1) = 1 Then
            
            If g_OpenImageCount > 1 Then
                g_WindowManager.setWindowVisibility toolbar_ImageTabs.hWnd, True
            Else
                g_WindowManager.setWindowVisibility toolbar_ImageTabs.hWnd, False
            End If
        
        'A setting of 0 equates to index 0 in the menu, specifically "always show tabstrip".
        Else
        
            If g_OpenImageCount > 0 Then
                g_WindowManager.setWindowVisibility toolbar_ImageTabs.hWnd, True
            Else
                g_WindowManager.setWindowVisibility toolbar_ImageTabs.hWnd, False
            End If
        
        End If
    
    End If
        
    'Perform a special check if 2 or more images are loaded; if that is the case, enable a few additional controls, like
    ' the "Next/Previous" Window menu items.
    If g_OpenImageCount >= 2 Then
        FormMain.MnuWindow(7).Enabled = True
        FormMain.MnuWindow(8).Enabled = True
    Else
        FormMain.MnuWindow(7).Enabled = False
        FormMain.MnuWindow(8).Enabled = False
    End If
        
    'Redraw the layer box
    toolbar_Layers.forceRedraw
        
End Sub

'metaToggle enables or disables a swath of controls related to a simple keyword (e.g. "Undo", which affects multiple menu items
' and toolbar buttons)
Public Sub metaToggle(ByVal metaItem As metaInitializer, ByVal NewState As Boolean, Optional ByVal suspendAssociatedRedraws As Boolean = False)
    
    Dim i As Long
    
    Select Case metaItem
            
        'Save (left-hand panel button AND menu item)
        Case tSave
            If FormMain.MnuFile(7).Enabled <> NewState Then
                toolbar_File.cmdSave.Enabled = NewState
                FormMain.MnuFile(7).Enabled = NewState
                
                'The File -> Revert menu is also tied to Save state (if the image has not been saved in its current state,
                ' we allow the user to revert to the last save state).
                FormMain.MnuFile(9).Enabled = NewState
                
            End If
            
        'Save As (menu item only)
        Case tSaveAs
            If FormMain.MnuFile(8).Enabled <> NewState Then
                toolbar_File.cmdSaveAs.Enabled = NewState
                FormMain.MnuFile(8).Enabled = NewState
            End If
            
        'Close and Close All
        Case tClose
            If FormMain.MnuFile(4).Enabled <> NewState Then
                FormMain.MnuFile(4).Enabled = NewState
                FormMain.MnuFile(5).Enabled = NewState
                toolbar_File.cmdClose.Enabled = NewState
            End If
        
        'Undo (left-hand panel button AND menu item)
        Case tUndo
        
            If FormMain.MnuEdit(0).Enabled <> NewState Then
                toolbar_File.cmdUndo.Enabled = NewState
                FormMain.MnuEdit(0).Enabled = NewState
            End If
            
            'If Undo is being enabled, change the text to match the relevant action that created this Undo file
            If NewState Then
                toolbar_File.cmdUndo.ToolTip = g_Language.TranslateMessage(pdImages(g_CurrentImage).undoManager.getUndoProcessID)
                FormMain.MnuEdit(0).Caption = g_Language.TranslateMessage("Undo:") & " " & g_Language.TranslateMessage(pdImages(g_CurrentImage).undoManager.getUndoProcessID) & vbTab & "Ctrl+Z"
            Else
                toolbar_File.cmdUndo.ToolTip = ""
                FormMain.MnuEdit(0).Caption = g_Language.TranslateMessage("Undo") & vbTab & "Ctrl+Z"
            End If
            
            'When changing menu text, icons must be reapplied.
            If Not suspendAssociatedRedraws Then resetMenuIcons
        
        'Redo (left-hand panel button AND menu item)
        Case tRedo
            If FormMain.MnuEdit(1).Enabled <> NewState Then
                toolbar_File.cmdRedo.Enabled = NewState
                FormMain.MnuEdit(1).Enabled = NewState
            End If
            
            'If Redo is being enabled, change the menu text to match the relevant action that created this Undo file
            If NewState Then
                toolbar_File.cmdRedo.ToolTip = g_Language.TranslateMessage(pdImages(g_CurrentImage).undoManager.getRedoProcessID)
                FormMain.MnuEdit(1).Caption = g_Language.TranslateMessage("Redo:") & " " & g_Language.TranslateMessage(pdImages(g_CurrentImage).undoManager.getRedoProcessID) & vbTab & "Ctrl+Y"
            Else
                toolbar_File.cmdRedo.ToolTip = ""
                FormMain.MnuEdit(1).Caption = g_Language.TranslateMessage("Redo") & vbTab & "Ctrl+Y"
            End If
            
            'When changing menu text, icons must be reapplied.
            If Not suspendAssociatedRedraws Then resetMenuIcons
            
        'Copy (menu item only)
        Case tCopy
            If FormMain.MnuEdit(7).Enabled <> NewState Then FormMain.MnuEdit(7).Enabled = NewState
            If FormMain.MnuEdit(8).Enabled <> NewState Then FormMain.MnuEdit(8).Enabled = NewState
            If FormMain.MnuEdit(9).Enabled <> NewState Then FormMain.MnuEdit(9).Enabled = NewState
            If FormMain.MnuEdit(10).Enabled <> NewState Then FormMain.MnuEdit(10).Enabled = NewState
            If FormMain.MnuEdit(12).Enabled <> NewState Then FormMain.MnuEdit(12).Enabled = NewState
            
        'View (top-menu level)
        Case tView
            If FormMain.MnuView.Enabled <> NewState Then FormMain.MnuView.Enabled = NewState
        
        'ImageOps is all Image-related menu items; it enables/disables the Image, Layer, Select, Color, and Print menus
        Case tImageOps
            If FormMain.MnuImageTop.Enabled <> NewState Then
                FormMain.MnuImageTop.Enabled = NewState
                
                'Use this same command to disable other menus
                
                'File -> Print
                FormMain.MnuFile(13).Enabled = NewState
                
                'Layer menu
                FormMain.MnuLayerTop.Enabled = NewState
                
                'Select menu
                FormMain.MnuSelectTop.Enabled = NewState
                
                'Adjustments menu
                FormMain.MnuAdjustmentsTop.Enabled = NewState
                
                'Effects menu
                FormMain.MnuEffectsTop.Enabled = NewState
                
            End If
            
        'Macro (within the Tools menu)
        Case tMacro
            If FormMain.mnuTool(3).Enabled <> NewState Then FormMain.mnuTool(3).Enabled = NewState
        
        'Selections in general
        Case tSelection
            
            'If selections are not active, clear all the selection value textboxes
            If Not NewState Then
                For i = 0 To toolbar_Tools.tudSel.Count - 1
                    toolbar_Tools.tudSel(i).Value = 0
                Next i
            End If
            
            'Set selection text boxes to enable only when a selection is active.  Other selection controls can remain active
            ' even without a selection present; this allows the user to set certain parameters in advance, so when they actually
            ' draw a selection, it already has the attributes they want.
            For i = 0 To toolbar_Tools.tudSel.Count - 1
                toolbar_Tools.tudSel(i).Enabled = NewState
            Next i
            
            'En/disable all selection menu items that rely on an existing selection to operate
            If FormMain.MnuSelect(2).Enabled <> NewState Then
                
                'Select none, invert selection
                FormMain.MnuSelect(1).Enabled = NewState
                FormMain.MnuSelect(2).Enabled = NewState
                
                'Grow/shrink/border/feather/sharpen selection
                For i = 4 To 8
                    FormMain.MnuSelect(i).Enabled = NewState
                Next i
                
                'Erase selected area
                FormMain.MnuSelect(10).Enabled = NewState
                
                'Save selection
                FormMain.MnuSelect(13).Enabled = NewState
                
                'Export selection top-level menu
                FormMain.MnuSelect(14).Enabled = NewState
                
            End If
                                    
            'Selection enabling/disabling also affects the Crop to Selection command
            If FormMain.MnuImage(9).Enabled <> NewState Then FormMain.MnuImage(9).Enabled = NewState
            
        'Transformable selection controls specifically
        Case tSelectionTransform
        
            'Under certain circumstances, it is desirable to disable only the selection location boxes
            For i = 0 To toolbar_Tools.tudSel.Count - 1
                If (Not NewState) Then toolbar_Tools.tudSel(i).Value = 0
                toolbar_Tools.tudSel(i).Enabled = NewState
            Next i
                
        'If the ExifTool plugin is not available, metadata will ALWAYS be disabled.  (We do not currently have a separate fallback for
        ' reading/browsing/writing metadata.)
        Case tMetadata
        
            If g_ExifToolEnabled Then
                If FormMain.MnuMetadata(0).Enabled <> NewState Then FormMain.MnuMetadata(0).Enabled = NewState
            Else
                If FormMain.MnuMetadata(0).Enabled Then FormMain.MnuMetadata(0).Enabled = False
            End If
        
        'GPS metadata is its own sub-category, and its activation is contigent upon an image having embedded GPS data
        Case tGPSMetadata
        
            If g_ExifToolEnabled Then
                If FormMain.MnuMetadata(3).Enabled <> NewState Then FormMain.MnuMetadata(3).Enabled = NewState
            Else
                If FormMain.MnuMetadata(3).Enabled Then FormMain.MnuMetadata(3).Enabled = False
            End If
        
        'Zoom controls not just the drop-down zoom box, but the zoom in, zoom out, and zoom fit buttons as well
        Case tZoom
            If FormMain.mainCanvas(0).getZoomDropDownReference().Enabled <> NewState Then
                FormMain.mainCanvas(0).getZoomDropDownReference().Enabled = NewState
                FormMain.mainCanvas(0).enableZoomIn NewState
                FormMain.mainCanvas(0).enableZoomOut NewState
                FormMain.mainCanvas(0).enableZoomFit NewState
            End If
            
            'When disabling zoom controls, reset the zoom drop-down to 100%
            If Not NewState Then FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex = g_Zoom.getZoom100Index
            
        'Non-destructive FX are effects that the user can apply to a layer, without permanently modifying the layer
        Case tNonDestructiveFX
        
            If NewState Then
                
                'Start by enabling all non-destructive FX controls
                For i = 0 To toolbar_Tools.sltQuickFix.Count - 1
                    If Not toolbar_Tools.sltQuickFix(i).Enabled Then toolbar_Tools.sltQuickFix(i).Enabled = True
                Next i
                
                'Quick fix buttons are only relevant if the current image has some non-destructive events applied
                If Not pdImages(g_CurrentImage).getActiveLayer Is Nothing Then
                
                    If pdImages(g_CurrentImage).getActiveLayer.getLayerNonDestructiveFXState() Then
                        
                        For i = 0 To toolbar_Tools.cmdQuickFix.Count - 1
                            toolbar_Tools.cmdQuickFix(i).Enabled = True
                        Next i
                        
                    Else
                        
                        For i = 0 To toolbar_Tools.cmdQuickFix.Count - 1
                            toolbar_Tools.cmdQuickFix(i).Enabled = False
                        Next i
                        
                    End If
                    
                End If
                
                'Disable automatic NDFX syncing, then update all sliders to match the current layer's values
                With toolbar_Tools
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
                For i = 0 To toolbar_Tools.sltQuickFix.Count - 1
                    If toolbar_Tools.sltQuickFix(i).Enabled Then toolbar_Tools.sltQuickFix(i).Enabled = False
                Next i
                
                For i = 0 To toolbar_Tools.cmdQuickFix.Count - 1
                    If toolbar_Tools.cmdQuickFix(i).Enabled Then toolbar_Tools.cmdQuickFix(i).Enabled = False
                Next i
                
            End If
            
    End Select
    
End Sub


'For best results, any modal form should be shown via this function.  This function will automatically center the form over the main window,
' while also properly assigning ownership so that the dialog is truly on top of any active windows.  It also handles deactivation of
' other windows (to prevent click-through), and dynamic top-most behavior to ensure that the program doesn't steal focus if the user switches
' to another program while a modal dialog is active.
Public Sub showPDDialog(ByRef dialogModality As FormShowConstants, ByRef dialogForm As Form, Optional ByVal doNotUnload As Boolean = False)

    On Error GoTo showPDDialogError

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
    
    'Register the window with the window manager, which will also make it a top-most window
    If g_WindowManager.getFloatState(TOOLBAR_WINDOW) Then g_WindowManager.requestTopmostWindow dialogHwnd, getModalOwner().hWnd
    
    'Use VB to actually display the dialog
    dialogForm.Show dialogModality, FormMain 'getModalOwner()
    
    'De-register this hWnd with the window manager
    If g_WindowManager.getFloatState(TOOLBAR_WINDOW) Then g_WindowManager.requestTopmostWindow dialogHwnd, 0, True
    
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
    
    Exit Sub
    
'For reasons I can't yet ascertain, this function will sometimes fail, claiming that a modal window is already active.  If that happens,
' we can just exit.
showPDDialogError:

End Sub

'When a modal dialog needs to be raised, we want to set its ownership to the top-most (relevant) window in the program, which may or may
' not be the main form.  This function should be called to determine the proper owner of any modal dialog box.
'
'If the caller knows in advance that a modal dialog is owned by another modal dialog (for example, a tool dialog displaying a color
' selection dialog), it can explicitly mark the assumeSecondaryDialog function as TRUE.
Public Function getModalOwner(Optional ByVal assumeSecondaryDialog As Boolean = False) As Form

    'If a modal dialog is already active, it gets ownership over subsequent dialogs
    If isSecondaryDialog Or assumeSecondaryDialog Then
        Set getModalOwner = currentDialogReference
        
    'No modal dialog is active, making this the only one.  Give the main form ownership.
    Else
        
        Set getModalOwner = FormMain
        
    End If
    
End Function

'Return the system keyboard delay, in seconds.  This isn't an exact science because the delay is actually hardware dependent
' (e.g. the system returns a value from 0 to 3), but we can use a "good enough" approximation.
Public Function getKeyboardDelay() As Double
    Dim keyDelayIndex As Long
    SystemParametersInfo SPI_GETKEYBOARDDELAY, 0, keyDelayIndex, 0
    getKeyboardDelay = (keyDelayIndex + 1) * 0.25
End Function

Public Sub toggleImageTabstripAlignment(ByVal newAlignment As AlignConstants, Optional ByVal suppressInterfaceSync As Boolean = False, Optional ByVal suppressPrefUpdate As Boolean = False)
    
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
    g_WindowManager.setImageTabstripAlignment newAlignment
    
    If Not suppressInterfaceSync Then
    
        '...and force the tabstrip to redraw itself (which it may not if the tabstrip's size hasn't changed, e.g. if Left and Right layout is toggled)
        toolbar_ImageTabs.forceRedraw
    
        'Refresh the current image viewport (which may be positioned differently due to the tabstrip moving)
        FormMain.refreshAllCanvases
        
    End If
    
End Sub

'The image tabstrip can set to appear under a variety of circumstances.  Use this sub to change the current setting; it will
' automatically handle syncing with the preferences file.
Public Sub toggleImageTabstripVisibility(ByVal newSetting As Long, Optional ByVal suppressInterfaceSync As Boolean = False, Optional ByVal suppressPrefUpdate As Boolean = False)

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
        syncInterfaceToCurrentImage
        
    End If
    
    'If images are loaded, we may need to redraw their viewports because the available client area may have changed.
    If (g_NumOfImagesLoaded > 0) Then
        PrepareViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0), "Image tabstrip visibility toggled"
    End If

End Sub

'Both toolbars and image windows can be floated or docked.  Because some behind-the-scenes maintenance has to be applied whenever
' this setting is changed, all float toggle operations should wrap this singular function.
Public Sub toggleWindowFloating(ByVal whichWindowType As pdWindowType, ByVal floatStatus As Boolean, Optional ByVal suspendMenuRefresh As Boolean = False, Optional ByVal suppressPrefUpdate As Boolean = False)

    'Make a note of the currently active image
    Dim backupCurrentImage As Long
    backupCurrentImage = g_CurrentImage
    
    Select Case whichWindowType
    
        Case TOOLBAR_WINDOW
            FormMain.MnuWindow(5).Checked = floatStatus
            If Not suppressPrefUpdate Then g_UserPreferences.SetPref_Boolean "Core", "Floating Toolbars", floatStatus
            g_WindowManager.setFloatState TOOLBAR_WINDOW, floatStatus, suspendMenuRefresh
            
            'If image windows are docked, we need to redraw all their windows, because the available client area will have changed.
            If Not suspendMenuRefresh Then FormMain.refreshAllCanvases
            
    End Select
    
End Sub

'Toolbars can be dynamically shown/hidden by a variety of processes (e.g. clicking an entry in the Window menu, clicking the X in a
' toolbar's command box, etc).  All those operations should wrap this singular function.
Public Sub toggleToolbarVisibility(ByVal whichToolbar As pdToolbarType)

    Select Case whichToolbar
    
        Case FILE_TOOLBOX
            FormMain.MnuWindow(0).Checked = Not FormMain.MnuWindow(0).Checked
            g_UserPreferences.SetPref_Boolean "Core", "Show File Toolbox", FormMain.MnuWindow(0).Checked
            g_WindowManager.setWindowVisibility toolbar_File.hWnd, FormMain.MnuWindow(0).Checked
            
        Case LAYER_TOOLBOX
            FormMain.MnuWindow(1).Checked = Not FormMain.MnuWindow(1).Checked
            g_UserPreferences.SetPref_Boolean "Core", "Show Layers Toolbox", FormMain.MnuWindow(1).Checked
            g_WindowManager.setWindowVisibility toolbar_Layers.hWnd, FormMain.MnuWindow(1).Checked
    
        Case TOOLS_TOOLBOX
            FormMain.MnuWindow(2).Checked = Not FormMain.MnuWindow(2).Checked
            g_UserPreferences.SetPref_Boolean "Core", "Show Selections Toolbox", FormMain.MnuWindow(2).Checked
            g_WindowManager.setWindowVisibility toolbar_Tools.hWnd, FormMain.MnuWindow(2).Checked
    
    End Select
    
    'Redraw the primary image viewport, as the available client area may have changed.
    If g_NumOfImagesLoaded > 0 Then FormMain.refreshAllCanvases
    
End Sub

Public Function fixDPI(ByVal pxMeasurement As Long) As Long

    'The first time this function is called, dpiRatio will be 0.  Calculate it.
    If dpiRatio = 0# Then
    
        'There are 1440 twips in one inch.  (Twips are resolution-independent.)  Use that knowledge to calculate DPI.
        dpiRatio = 1440 / TwipsPerPixelXFix
        
        'FYI: if the screen resolution is 96 dpi, this function will return the original pixel measurement, due to
        ' this calculation.
        dpiRatio = dpiRatio / 96
    
    End If
    
    fixDPI = CLng(dpiRatio * CDbl(pxMeasurement))
    
End Function

Public Function fixDPIFloat(ByVal pxMeasurement As Long) As Double

    'The first time this function is called, dpiRatio will be 0.  Calculate it.
    If dpiRatio = 0# Then
    
        'There are 1440 twips in one inch.  (Twips are resolution-independent.)  Use that knowledge to calculate DPI.
        dpiRatio = 1440 / TwipsPerPixelXFix
        
        'FYI: if the screen resolution is 96 dpi, this function will return the original pixel measurement, due to
        ' this calculation.
        dpiRatio = dpiRatio / 96
    
    End If
    
    fixDPIFloat = dpiRatio * CDbl(pxMeasurement)
    
End Function

'Fun fact: there are 15 twips per pixel at 96 DPI.  Not fun fact: at 200% DPI (e.g. 192 DPI), VB's internal
' TwipsPerPixelXFix will return 7, when actually we need the value 7.5.  This causes problems when resizing
' certain controls (like SmartCheckBox) because the size will actually come up short due to rounding errors!
' So whenever TwipsPerPixelXFix/Y is required, use these functions instead.
Public Function TwipsPerPixelXFix() As Double

    If Screen.TwipsPerPixelX = 7 Then
        TwipsPerPixelXFix = 7.5
    Else
        TwipsPerPixelXFix = Screen.TwipsPerPixelX
    End If

End Function

Public Function TwipsPerPixelYFix() As Double

    If Screen.TwipsPerPixelY = 7 Then
        TwipsPerPixelYFix = 7.5
    Else
        TwipsPerPixelYFix = Screen.TwipsPerPixelY
    End If

End Function

Public Sub displayWaitScreen(ByVal waitTitle As String, ByRef ownerForm As Form)
    
    FormWait.Visible = False
    
    FormWait.lblWaitTitle.Caption = waitTitle
    FormWait.lblWaitTitle.Visible = True
    FormWait.lblWaitTitle.Refresh
    
    Screen.MousePointer = vbHourglass
    
    FormWait.Show vbModeless, ownerForm
    FormWait.Refresh
    DoEvents
    
End Sub

Public Sub hideWaitScreen()
    Screen.MousePointer = vbDefault
    Unload FormWait
End Sub

'Given a wordwrap label with a set size, attempt to fit the label's text inside it
Public Sub fitWordwrapLabel(ByRef srcLabel As Label, ByRef srcForm As Form)

    'We will use a pdFont object to help us measure the label in question
    Dim tmpFont As pdFont
    Set tmpFont = New pdFont
    tmpFont.setFontBold srcLabel.FontBold
    tmpFont.setFontItalic srcLabel.FontItalic
    tmpFont.setFontFace srcLabel.FontName
    tmpFont.setFontSize srcLabel.FontSize
    tmpFont.createFontObject
    tmpFont.setTextAlignment srcLabel.Alignment
    tmpFont.attachToDC srcForm.hDC
    
    'Retrieve the height from the pdFont class
    Dim lblHeight As Long
    lblHeight = tmpFont.getHeightOfWordwrapString(srcLabel.Caption, srcLabel.Width - 1)
    
    Dim curFontSize As Long
    curFontSize = srcLabel.FontSize
    
    'If the text is too tall, shrink the font until an acceptable size is found.  Note that the reported text value tends to be
    ' smaller than the space actually required.  I do not know why this happens.  To account for it, I cut a further 10% from
    ' the requested height, just to be safe.
    If (lblHeight > srcLabel.Height * 0.85) Then
            
        'Try shrinking the font size until an acceptable width is found
        Do While (lblHeight > srcLabel.Height * 0.85) And (curFontSize >= 8)
        
            curFontSize = curFontSize - 1
            
            tmpFont.setFontSize curFontSize
            tmpFont.createFontObject
            tmpFont.attachToDC srcForm.hDC
            lblHeight = tmpFont.getHeightOfWordwrapString(srcLabel.Caption, srcLabel.Width)
            
        Loop
            
    End If
    
    'When an acceptable size is found, set it and exit.
    srcLabel.FontSize = curFontSize
    srcLabel.Refresh

End Sub

'Because VB6 apps look terrible on modern version of Windows, I do a bit of beautification to every form upon at load-time.
' This routine is nice because every form calls it at least once, so I can make centralized changes without having to rewrite
' code in every individual form.  This is also where run-time translation occurs.
Public Sub makeFormPretty(ByRef tForm As Form, Optional ByRef customTooltips As clsToolTip, Optional ByVal tooltipsAlreadyInitialized As Boolean = False, Optional ByVal useDoEvents As Boolean = False)

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
            eControl.FontName = g_InterfaceFont
        End If
        
        If ((TypeOf eControl Is jcbutton) Or (TypeOf eControl Is sliderTextCombo) Or (TypeOf eControl Is textUpDown) Or (TypeOf eControl Is commandBar) Or (TypeOf eControl Is smartResize)) Then
            eControl.Font.Name = g_InterfaceFont
        End If
        
        If (TypeOf eControl Is smartOptionButton) Or (TypeOf eControl Is smartCheckBox) Then
            eControl.updateAgainstCurrentTheme
        End If
                        
        'STEP 3: remove TabStop from each picture box.  They should never receive focus, but I often forget to change this
        ' at design-time.
        If (TypeOf eControl Is PictureBox) Then eControl.TabStop = False
        
        'Optionally, DoEvents can be called after each change.  This slows the process, but it allows external progress
        ' bars to be automatically refreshed.
        If useDoEvents Then DoEvents
                
    Next
    
    'FORM STEP 2: subclass this form and force controls to render transparent borders properly.
    g_Themer.requestContainerSubclass tForm.hWnd
    
    'FORM STEP 3: translate the form (and all controls on it)
    If g_Language.translationActive And tForm.Enabled Then
        g_Language.applyTranslations tForm, useDoEvents
    End If
    
    'FORM STEP 4: if a custom tooltip handler was passed in, activate and populate it now.
    If Not (customTooltips Is Nothing) Then
        
        'In rare cases, the custom tooltip handler passed to this function may already be initialized.  Some forms
        ' do this if they need to handle multiline tooltips (as VB will not handle them properly).  If the class has
        ' NOT been initialized, we can do so now - otherwise, trust that it was already created correctly.
        If Not tooltipsAlreadyInitialized Then
            customTooltips.Create tForm
            customTooltips.MaxTipWidth = PD_MAX_TOOLTIP_WIDTH
            customTooltips.DelayTime(ttDelayShow) = 10000
        End If
        
        'Once again, enumerate every control on the form and copy their tooltips into this object.  (This allows
        ' for things like automatic multiline support, unsupported characters, theming, and displaying tooltips
        ' on the correct monitor of a multimonitor setup.)
        Dim tmpTooltip As String
        For Each eControl In tForm.Controls
            
            If (TypeOf eControl Is CommandButton) Or (TypeOf eControl Is CheckBox) Or (TypeOf eControl Is OptionButton) Or (TypeOf eControl Is PictureBox) Or (TypeOf eControl Is TextBox) Or (TypeOf eControl Is ListBox) Or (TypeOf eControl Is ComboBox) Or (TypeOf eControl Is colorSelector) Then
                If (Trim(eControl.ToolTipText) <> "") Then
                    tmpTooltip = eControl.ToolTipText
                    eControl.ToolTipText = ""
                    customTooltips.AddTool eControl, tmpTooltip
                End If
            End If
            
            'Optionally, DoEvents can be called after each change.  This slows the process, but it allows external progress
            ' bars to be automatically refreshed.
            If useDoEvents Then DoEvents
            
        Next
                
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
Public Sub handleClearType(ByVal startingProgram As Boolean)
    
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
    g_Themer.releaseContainerSubclass tForm.hWnd
    Set tForm = Nothing
End Sub

'Given a pdImage object, generate an appropriate caption for the main PhotoDemon window.
Private Function getWindowCaption(ByRef srcImage As pdImage) As String

    Dim captionBase As String
    captionBase = ""
    
    'Start by seeing if this image has some kind of filename.  This field should always be populated by the load function,
    ' but better safe than sorry!
    If Len(srcImage.originalFileNameAndExtension) > 0 Then
    
        'This image has a filename!  Next, check the user's preference for long or short window captions
        If g_UserPreferences.GetPref_Long("Interface", "Window Caption Length", 0) = 0 Then
            
            'The user prefers short captions.  Use just the filename and extension (no folders ) as the base.
            captionBase = srcImage.originalFileNameAndExtension
        Else
        
            'The user prefers long captions.  Make sure this image has such a location; if they do not, fallback
            ' and use just the filename.
            If Len(srcImage.locationOnDisk) > 0 Then
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
    getWindowCaption = captionBase & "  -  " & getPhotoDemonNameAndVersion()

End Function


'Display the specified size in the main form's status bar
Public Sub DisplaySize(ByRef srcImage As pdImage)
    
    If Not (srcImage Is Nothing) Then
        FormMain.mainCanvas(0).displayImageSize srcImage
    End If
    
    'Size is only displayed when it is changed, so if any controls have a maximum value linked to the size of the image,
    ' now is an excellent time to update them.
    If srcImage.Width < srcImage.Height Then
        toolbar_Tools.sltSelectionBorder.Max = srcImage.Width
        toolbar_Tools.sltCornerRounding.Max = srcImage.Width
        toolbar_Tools.sltSelectionLineWidth.Max = srcImage.Height
    Else
        toolbar_Tools.sltSelectionBorder.Max = srcImage.Height
        toolbar_Tools.sltCornerRounding.Max = srcImage.Height
        toolbar_Tools.sltSelectionLineWidth.Max = srcImage.Width
    End If
    
End Sub

'This wrapper is used in place of the standard MsgBox function.  At present it's just a wrapper around MsgBox, but
' in the future I may replace the dialog function with something custom.
Public Function pdMsgBox(ByVal pMessage As String, ByVal pButtons As VbMsgBoxStyle, ByVal pTitle As String, ParamArray ExtraText() As Variant) As VbMsgBoxResult

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
    If Not IsMissing(ExtraText) Then
    
        Dim i As Long
        For i = LBound(ExtraText) To UBound(ExtraText)
            newMessage = Replace$(newMessage, "%" & i + 1, CStr(ExtraText(i)))
        Next i
    
    End If

    pdMsgBox = MsgBox(newMessage, pButtons, newTitle)

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
    
    If Not IsMissing(ExtraText) Then
                    
        For i = LBound(ExtraText) To UBound(ExtraText)
            tmpDupeCheckString = Replace$(tmpDupeCheckString, "%" & i + 1, CStr(ExtraText(i)))
        Next i
        
    End If
    
    'If the message request is for a novel string (e.g. one that differs from the previous message request), display it.
    ' Otherwise, exit now.
    If StrComp(m_PrevMessage, tmpDupeCheckString, vbBinaryCompare) <> 0 Then
        
        'In debug mode, mirror the message output to PD's central Debugger.  Note that this behavior can be overridden by
        ' supplying the string "DONOTLOG" as the final entry in the ParamArray.
        #If DEBUGMODE = 1 Then
            If IsMissing(ExtraText) Then
                pdDebug.LogAction tmpDupeCheckString, PDM_USER_MESSAGE
            Else
            
                'Check the last param passed.  If it's the string "DONOTLOG", do not log this entry.  (PD sometimes uses this
                ' to avoid logging useless data, like layer hover events or download updates.)
                If StrComp(CStr(ExtraText(UBound(ExtraText))), "DONOTLOG", vbTextCompare) <> 0 Then
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
        If Not IsMissing(ExtraText) Then
        
            For i = LBound(ExtraText) To UBound(ExtraText)
                newString = Replace$(newString, "%" & i + 1, CStr(ExtraText(i)))
            Next i
        
        End If
        
        'While macros are active, append a "Recording" message to help orient the user
        If MacroStatus = MacroSTART Then newString = newString & " {-" & g_Language.TranslateMessage("Recording") & "-}"
        
        'Post the message to the screen
        If MacroStatus <> MacroBATCH Then FormMain.mainCanvas(0).displayCanvasMessage newString
        
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
    'If g_OpenImageCount > 0 Then
    FormMain.mainCanvas(0).displayCanvasCoordinates 0, 0, True
    'End If
End Sub

'Populate the passed combo box with options related to distort filter edge-handle options.  Also, select the specified method by default.
Public Sub popDistortEdgeBox(ByRef cmbEdges As ComboBox, Optional ByVal defaultEdgeMethod As EDGE_OPERATOR)

    cmbEdges.Clear
    cmbEdges.AddItem " clamp them to the nearest available pixel"
    cmbEdges.AddItem " reflect them across the nearest edge"
    cmbEdges.AddItem " wrap them around the image"
    cmbEdges.AddItem " erase them"
    cmbEdges.AddItem " ignore them"
    cmbEdges.ListIndex = defaultEdgeMethod
    
End Sub

'Return the width (and below, height) of a string, in pixels, according to the font assigned to fontContainerDC
Public Function getPixelWidthOfString(ByVal srcString As String, ByVal fontContainerDC As Long) As Long
    Dim txtSize As POINTAPI
    GetTextExtentPoint32 fontContainerDC, srcString, Len(srcString), txtSize
    getPixelWidthOfString = txtSize.x
End Function

Public Function getPixelHeightOfString(ByVal srcString As String, ByVal fontContainerDC As Long) As Long
    Dim txtSize As POINTAPI
    GetTextExtentPoint32 fontContainerDC, srcString, Len(srcString), txtSize
    getPixelHeightOfString = txtSize.y
End Function

'Use whenever you want the user to not be allowed to interact with the primary PD window.  Make sure that call "enableUserInput", below,
' when you are done processing!
Public Sub disableUserInput()

    'Set the "input disabled" flag, which individual functions can use to modify their own behavior
    g_DisableUserInput = True
    
    'Forcibly disable the main form
    FormMain.Enabled = False

End Sub

'Sister function to "disableUserInput", above
Public Sub enableUserInput()
    
    'Start a countdown timer on the main form.  When it terminates, user input will be restored.  A timer is required because
    ' we need a certain amount of "dead time" to elapse between double-clicks on a top-level dialog (like a common dialog)
    ' which may be incorrectly passed through to the main form.  (I know, this seems like a ridiculous solution, but I tried
    ' a thousand others before settling on this.  It's the least of many evils.)
    FormMain.tmrCountdown.Enabled = True
    
    'Re-enable the main form
    FormMain.Enabled = True

End Sub

'Given a combo box, populate it with all currently supported blend modes
Public Sub populateBlendModeComboBox(ByRef dstCombo As ComboBox, Optional ByVal blendIndex As LAYER_BLENDMODE = BL_NORMAL)
    
    dstCombo.Clear
    
    dstCombo.AddItem "Normal", 0
    dstCombo.AddItem "Darken", 1
    dstCombo.AddItem "Multiply", 2
    dstCombo.AddItem "Color burn", 3
    dstCombo.AddItem "Linear burn", 4
    dstCombo.AddItem "Lighten", 5
    dstCombo.AddItem "Screen", 6
    dstCombo.AddItem "Color dodge", 7
    dstCombo.AddItem "Linear dodge", 8
    dstCombo.AddItem "Overlay", 9
    dstCombo.AddItem "Soft light", 10
    dstCombo.AddItem "Hard light", 11
    dstCombo.AddItem "Vivid light", 12
    dstCombo.AddItem "Linear light", 13
    dstCombo.AddItem "Pin light", 14
    dstCombo.AddItem "Hard mix", 15
    dstCombo.AddItem "Difference", 16
    dstCombo.AddItem "Exclusion", 17
    dstCombo.AddItem "Subtract", 18
    dstCombo.AddItem "Divide", 19
    dstCombo.AddItem "Hue", 20
    dstCombo.AddItem "Saturation", 21
    dstCombo.AddItem "Color", 22
    dstCombo.AddItem "Luminosity", 23
    
    dstCombo.ListIndex = blendIndex
    
End Sub
