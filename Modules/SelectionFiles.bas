Attribute VB_Name = "SelectionFiles"
'***************************************************************************
'Selection Tools: File I/O
'Copyright 2013-2026 by Tanner Helland
'Created: 21/June/13
'Last updated: 17/June/26
'Last update: route exports through the new "MenuExportImage" wrapper, which handles additional user behaviors
'             (like hand-typing a file extension that doesn't match the selected format dropdown)
'
'This module should only contain functions for writing/reading selection data to file.  Note that these
' functions will be used primarily for PD's Undo/Redo engine, so performance considerations are paramount.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Export the currently selected area as an image.  This is provided as a convenience to the user, so that they do not
' have to crop or copy-paste the selected area in order to save it.  The selected area is also checked for bit-depth;
' 24bpp is recommended as JPEG, while 32bpp is recommended as PNG (but the user can select any supported PD save format
' from the common dialog).
Public Function ExportSelectedAreaAsImageFile() As Boolean
    
    'If a selection is not active, it should be impossible to select this menu item.
    ' (Just in case, check for that state and exit if necessary.)
    If (Not PDImages.GetActiveImage.IsSelectionActive()) Then
        ExportSelectedAreaAsImageFile = False
        Exit Function
    End If
    
    'Because we're going to export a full image file (not just a pixel buffer), prepare a temporary
    ' pdImage object; it will house the current selection mask and any other image-specific properties.
    Dim tmpImage As pdImage
    Set tmpImage = New pdImage
    
    'Copy the current selection DIB into a temporary DIB.
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    PDImages.GetActiveImage.RetrieveProcessedSelection tmpDIB, True, True
    
    'In the temporary pdImage object, initialize a new layer using the contents of the temporary
    ' selection mask pixel buffer.
    Dim newLayerID As Long
    newLayerID = tmpImage.CreateBlankLayer
    tmpImage.GetLayerByID(newLayerID).InitializeNewLayer PDL_Image, vbNullString, tmpDIB
    
    'Ensure the parent image inherits the selection mask's pixel dimensions
    tmpImage.UpdateSize
    
    'Give the selection a basic filename (this isn's especially relevant, since the user can/will override it)
    tmpImage.ImgStorage.AddEntry "OriginalFileName", g_Language.TranslateMessage("PhotoDemon selection")
    
    'Use the exporter's original logic to determine how to export the image.
    ' (Typically, this will default to the last-used format from an export tool.)
    tmpImage.SetCurrentFileFormat PDIF_UNKNOWN
    
    'Let the central exporter handle the actual export flow from this point on
    ExportSelectedAreaAsImageFile = FileMenu.MenuExportImage(tmpImage)
    
    'Release our temporary pixel buffer and parent image container
    Set tmpDIB = Nothing
    Set tmpImage = Nothing
    
End Function

'Export the current selection mask as an image.  PNG is recommended by default, but the user can choose from any of PD's available formats.
Public Function ExportSelectionMaskAsImage() As Boolean
    
    'If a selection is not active, it should be impossible to select this menu item.
    ' (Just in case, check for that state and exit if necessary.)
    If Not PDImages.GetActiveImage.IsSelectionActive() Then
        ExportSelectionMaskAsImage = False
        Exit Function
    End If
    
    'Prepare a temporary pdImage object to house the current selection mask
    Dim tmpImage As pdImage
    Set tmpImage = New pdImage
    
    'Create a temporary DIB, then retrieve the current selection into it
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    tmpDIB.CreateFromExistingDIB PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB
    
    'Selections use a "white = selected, transparent = unselected" strategy.  Composite against a black background now
    ' (but leave the DIB in 32-bpp format)
    tmpDIB.CompositeBackgroundColor 0, 0, 0
    
    'In a temporary pdImage object, create a blank layer; this will receive the processed DIB
    Dim newLayerID As Long
    newLayerID = tmpImage.CreateBlankLayer
    tmpImage.GetLayerByID(newLayerID).InitializeNewLayer PDL_Image, vbNullString, tmpDIB
    tmpImage.UpdateSize
    
    'Give the selection a basic filename
    tmpImage.ImgStorage.AddEntry "OriginalFileName", g_Language.TranslateMessage("PhotoDemon selection")
    
    'Use the exporter's original logic to determine how to export the image.
    ' (Typically, this will default to the last-used format from an export tool.)
    tmpImage.SetCurrentFileFormat PDIF_UNKNOWN
    
    'Let the central exporter handle the actual export flow from this point on
    ExportSelectionMaskAsImage = FileMenu.MenuExportImage(tmpImage)
    
    'Release our temporary pixel buffer and parent image container
    Set tmpDIB = Nothing
    Set tmpImage = Nothing
    
End Function

'Export the currently selection mask as a grayscale layer in the current image.  This allows the user to interact
' with the selection using any of PD's raster tools.  (Later, they can convert the layer back into a selection
' using the ImportLayerAsSelectionMask tool.)
Public Function ExportSelectionMaskAsLayer() As Boolean
    
    'If a selection is not active, it should be impossible to select this menu item.
    ' (Just in case, check for that state and exit if necessary.)
    If (Not PDImages.GetActiveImage.IsSelectionActive()) Then
        ExportSelectionMaskAsLayer = False
        Exit Function
    End If
    
    'Create a temporary DIB, then copy the current selection mask into it.
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    tmpDIB.CreateFromExistingDIB PDImages.GetActiveImage.MainSelection.GetCompositeMaskDIB
    
    'Selections use a "white = selected, transparent = unselected" strategy.
    ' Composite against a black background (but leave the DIB in 32-bpp format).
    tmpDIB.CompositeBackgroundColor 0, 0, 0
    
    'We now need to add this DIB to the active image as a new layer.
    Dim newLayerID As Long
    newLayerID = PDImages.GetActiveImage.CreateBlankLayer()
    PDImages.GetActiveImage.GetLayerByID(newLayerID).InitializeNewLayer PDL_Image, g_Language.TranslateMessage("Selection mask to layer"), tmpDIB
    
    'Make the blank layer the new active layer
    PDImages.GetActiveImage.SetActiveLayerByID newLayerID
    
    'Notify the parent of the change
    PDImages.GetActiveImage.NotifyImageChanged UNDO_Image_VectorSafe
    
    'Redraw the layer box, and note that thumbnails need to be re-cached
    toolbar_Layers.NotifyLayerChange
    
    'Render the new image to screen (not technically necessary, but doesn't hurt)
    Viewport.Stage1_InitializeBuffer PDImages.GetActiveImage(), FormMain.MainCanvas(0)
            
    'Synchronize the interface to the new image
    SyncInterfaceToCurrentImage
    
End Function

'Import a new selection mask from a grayscale layer in the current image.  This allows the user to create a selection
' using whatever mechanism they want, then use the results as if it were a selection mask.
Public Function ImportSelectionMaskFromLayer() As Boolean
    
    'For now, this function uses the contents of the *active* layer.
    
    'Make a temporary copy of the active layer (including pixel contents)
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer
    tmpLayer.CopyExistingLayer PDImages.GetActiveImage.GetActiveLayer
    
    'Ensure the layer's contents are the same size as the parent image
    tmpLayer.ConvertToNullPaddedLayer PDImages.GetActiveImage.Width, PDImages.GetActiveImage.Height, True
    
    'Retrieve a grayscale map of the pixel data
    Dim grayBytes() As Byte
    DIBs.GetDIBGrayscaleMap tmpLayer.GetLayerDIB, grayBytes, False
    
    'We now have everything we need from the temporary layer.  We're now going to create a blank white
    ' copy of the layer, then apply the grayscale to it as a transparency layer.  (This is how "normal"
    ' selection data behaves in PD.)
    tmpLayer.GetLayerDIB.ResetDIB 255
    DIBs.ApplyTransparencyTable tmpLayer.GetLayerDIB, grayBytes
    
    'The temporary DIB now has everything it needs to be treated as selection data.
    ' Hand it off to the current image's selection manager; it'll handle the rest!
    PDImages.GetActiveImage.MainSelection.ReadSelectionFromDIB tmpLayer.GetLayerDIB
    
    'Ensure the user didn't give us a fully black or transparent mask (which is effectively a "null" selection).
    If PDImages.GetActiveImage.MainSelection.FindNewBoundsManually() Then
    
        'At least one valid selection pixel exists.  Activate it as the "current" selection.
        
        'Lock in this selection
        PDImages.GetActiveImage.MainSelection.LockIn
        PDImages.GetActiveImage.SetSelectionActive True
        
        'Draw the new selection to the screen
        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
    'No selection pixels exist.  Unload the selection mask, and treat this operation as a "remove selection" one.
    Else
        PDDebug.LogAction "No bounds found; removing selection."
        Selections.RemoveCurrentSelection
    End If
    
    'Whatever happened, synchronize the interface to the new image.  This ensures that menus are correctly
    ' dis/enabled according to the new selection state.
    SyncInterfaceToCurrentImage
    
End Function

'Load a previously saved selection.  Note that this function also handles creation and display of the relevant common dialog.
Public Sub LoadSelectionFromFile(ByVal displayDialog As Boolean, Optional ByVal loadSettings As String = vbNullString)

    If displayDialog Then
    
        'Disable user input until the dialog closes
        Interface.DisableUserInput
    
        'Simple open dialog
        Dim openDialog As pdOpenSaveDialog
        Set openDialog = New pdOpenSaveDialog
        
        Dim sFile As String
        
        Dim cdFilter As String
        cdFilter = g_Language.TranslateMessage("PhotoDemon selection") & " (.pds)|*.pds|"
        cdFilter = cdFilter & g_Language.TranslateMessage("All files") & "|*.*"
        
        Dim cdTitle As String
        cdTitle = g_Language.TranslateMessage("Load a previously saved selection")
                
        If openDialog.GetOpenFileName(sFile, vbNullString, True, False, cdFilter, 1, UserPrefs.GetSelectionPath, cdTitle, , GetModalOwner().hWnd) Then
            
            'Use a temporary selection object to validate the requested selection file
            Dim tmpSelection As pdSelection
            Set tmpSelection = New pdSelection
            tmpSelection.SetParentReference PDImages.GetActiveImage()
            
            If tmpSelection.ReadSelectionFromFile(sFile, True) Then
                
                'Save the new directory as the default path for future usage
                UserPrefs.SetSelectionPath sFile
                
                'Call this function again, but with displayDialog set to FALSE and the path of the requested selection file
                Process "Load selection", False, BuildParamList("selectionpath", sFile), UNDO_Selection
                    
            Else
                PDMsgBox "An error occurred while attempting to load %1.  Please verify that the file is a valid PhotoDemon selection file.", vbOKOnly Or vbExclamation, "Error", sFile
            End If
            
            'Release the temporary selection object
            tmpSelection.SetParentReference Nothing
            Set tmpSelection = Nothing
            
        End If
        
        'Re-enable user input
        Interface.EnableUserInput
        
    Else
        
        Dim cParams As pdSerialize
        Set cParams = New pdSerialize
        cParams.SetParamString loadSettings
        
        Message "Loading selection..."
        PDImages.GetActiveImage.MainSelection.ReadSelectionFromFile cParams.GetString("selectionpath")
        PDImages.GetActiveImage.SetSelectionActive True
        
        'Synchronize all user-facing controls to match
        SyncTextToCurrentSelection PDImages.GetActiveImageID()
                
        'Draw the new selection to the screen
        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
        Message "Selection loaded successfully"
    
    End If
        
End Sub

'Save the current selection to file.  Note that this function also handles creation and display of the relevant common dialog.
Public Sub SaveSelectionToFile()

    'Simple save dialog
    Dim saveDialog As pdOpenSaveDialog
    Set saveDialog = New pdOpenSaveDialog
    
    Dim sFile As String
    
    Dim cdFilter As String
    cdFilter = g_Language.TranslateMessage("PhotoDemon selection") & " (.pds)|*.pds"
    
    Dim cdTitle As String
    cdTitle = g_Language.TranslateMessage("Save the current selection")
        
    If saveDialog.GetSaveFileName(sFile, , True, cdFilter, 1, UserPrefs.GetSelectionPath, cdTitle, ".pds", GetModalOwner().hWnd) Then
        
        'Save the new directory as the default path for future usage
        UserPrefs.SetSelectionPath sFile
        
        'Write out the selection file
        Dim cmpLevel As Long
        cmpLevel = Compression.GetDefaultCompressionLevel(cf_Zstd)
        If PDImages.GetActiveImage.MainSelection.WriteSelectionToFile(sFile, cf_Zstd, cmpLevel, cf_Zstd, cmpLevel) Then
            Message "Selection saved."
        Else
            Message "Unknown error occurred.  Selection was not saved.  Please try again."
        End If
        
    End If
        
End Sub

