Attribute VB_Name = "SelectionFiles"
'***************************************************************************
'Selection Tools: File I/O
'Copyright 2013-2026 by Tanner Helland
'Created: 21/June/13
'Last updated: 07/September/21
'Last update: split selection file I/O into its own module
'
'This module should only contain functions for writing/reading selection data to file.  Note that these
' functions will be used primarily for PD's Undo/Redo engine, so performance considerations are paramount.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Export the currently selected area as an image.  This is provided as a convenience to the user, so that they do not have to crop
' or copy-paste the selected area in order to save it.  The selected area is also checked for bit-depth; 24bpp is recommended as
' JPEG, while 32bpp is recommended as PNG (but the user can select any supported PD save format from the common dialog).
Public Function ExportSelectedAreaAsImage() As Boolean
    
    'If a selection is not active, it should be impossible to select this menu item.  Just in case, check for that state and exit if necessary.
    If (Not PDImages.GetActiveImage.IsSelectionActive) Then
        Message "This action requires an active selection.  Please create a selection before continuing."
        ExportSelectedAreaAsImage = False
        Exit Function
    End If
    
    'Prepare a temporary pdImage object to house the current selection mask
    Dim tmpImage As pdImage
    Set tmpImage = New pdImage
    
    'Copy the current selection DIB into a temporary DIB.
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    PDImages.GetActiveImage.RetrieveProcessedSelection tmpDIB, True, True
    
    'In the temporary pdImage object, create a blank layer; this will receive the processed DIB
    Dim newLayerID As Long
    newLayerID = tmpImage.CreateBlankLayer
    tmpImage.GetLayerByID(newLayerID).InitializeNewLayer PDL_Image, , tmpDIB
    tmpImage.UpdateSize
    
    'Give the selection a basic filename
    tmpImage.ImgStorage.AddEntry "OriginalFileName", "PhotoDemon selection"
    
    'Get the last "save image" path from the preferences file
    Dim tempPathString As String
    tempPathString = UserPrefs.GetPref_String("Paths", "Save Image", vbNullString)
    
    'By default, recommend JPEG for 24bpp selections, and PNG for 32bpp selections
    Dim saveFormat As Long
    If DIBs.IsDIBAlphaBinary(tmpDIB, False) Then
        saveFormat = ImageFormats.GetIndexOfOutputPDIF(PDIF_JPEG) + 1
    Else
        saveFormat = ImageFormats.GetIndexOfOutputPDIF(PDIF_PNG) + 1
    End If
    
    'Now it's time to prepare a standard Save Image common dialog
    Dim saveDialog As pdOpenSaveDialog
    Set saveDialog = New pdOpenSaveDialog
    
    'Provide a string to the common dialog; it will fill this with the user's chosen path + filename
    Dim sFile As String
    sFile = tempPathString & IncrementFilename(tempPathString, tmpImage.ImgStorage.GetEntry_String("OriginalFileName", vbNullString), ImageFormats.GetOutputFormatExtension(saveFormat - 1))
    
    'Present a common dialog to the user
    If saveDialog.GetSaveFileName(sFile, , True, ImageFormats.GetCommonDialogOutputFormats, saveFormat, tempPathString, g_Language.TranslateMessage("Export selection as image"), ImageFormats.GetCommonDialogDefaultExtensions, FormMain.hWnd) Then
        
        'Store the selected file format to the image object
        tmpImage.SetCurrentFileFormat ImageFormats.GetOutputPDIF(saveFormat - 1)
        
        'Transfer control to the core SaveImage routine, which will handle color depth analysis and actual saving
        ExportSelectedAreaAsImage = PhotoDemon_SaveImage(tmpImage, sFile, True)
        
    Else
        ExportSelectedAreaAsImage = False
    End If
        
    'Release our temporary image
    Set tmpDIB = Nothing
    Set tmpImage = Nothing
    
End Function

'Export the current selection mask as an image.  PNG is recommended by default, but the user can choose from any of PD's available formats.
Public Function ExportSelectionMaskAsImage() As Boolean
    
    'If a selection is not active, it should be impossible to select this menu item.  Just in case, check for that state and exit if necessary.
    If Not PDImages.GetActiveImage.IsSelectionActive Then
        Message "This action requires an active selection.  Please create a selection before continuing."
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
    tmpImage.GetLayerByID(newLayerID).InitializeNewLayer PDL_Image, , tmpDIB
    tmpImage.UpdateSize
    
    'Give the selection a basic filename
    tmpImage.ImgStorage.AddEntry "OriginalFileName", g_Language.TranslateMessage("PhotoDemon selection")
        
    'Get the last "save image" path from the preferences file
    Dim tempPathString As String
    tempPathString = UserPrefs.GetPref_String("Paths", "Save Image", vbNullString)
    
    'By default, recommend PNG as the save format
    Dim saveFormat As Long
    saveFormat = ImageFormats.GetIndexOfOutputPDIF(PDIF_PNG) + 1
    
    'Provide a string to the common dialog; it will fill this with the user's chosen path + filename
    Dim sFile As String
    sFile = tempPathString & IncrementFilename(tempPathString, tmpImage.ImgStorage.GetEntry_String("OriginalFileName", vbNullString), "png")
    
    'Now it's time to prepare a standard Save Image common dialog
    Dim saveDialog As pdOpenSaveDialog
    Set saveDialog = New pdOpenSaveDialog
    
    'Present a common dialog to the user
    If saveDialog.GetSaveFileName(sFile, , True, ImageFormats.GetCommonDialogOutputFormats, saveFormat, tempPathString, g_Language.TranslateMessage("Export selection as image"), ImageFormats.GetCommonDialogDefaultExtensions, FormMain.hWnd) Then
                
        'Store the selected file format to the image object
        tmpImage.SetCurrentFileFormat ImageFormats.GetOutputPDIF(saveFormat - 1)
                                
        'Transfer control to the core SaveImage routine, which will handle color depth analysis and actual saving
        ExportSelectionMaskAsImage = PhotoDemon_SaveImage(tmpImage, sFile, True)
        
    Else
        ExportSelectionMaskAsImage = False
    End If
    
    'Release our temporary image
    Set tmpImage = Nothing

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

