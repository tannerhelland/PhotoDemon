Attribute VB_Name = "File_Menu"
'***************************************************************************
'File Menu Handler
'Copyright 2001-2015 by Tanner Helland
'Created: 15/Apr/01
'Last updated: 15/August/15
'Last update: convert the old cCommonDialog references to the newer, lighter pdOpenSaveDialog instance
'
'Functions for controlling standard file menu options.  Currently only handles "open image" and "save image".
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'This subroutine loads an image - note that the interesting stuff actually happens in PhotoDemon_OpenImageDialog, below
Public Sub MenuOpen()
    
    'String returned from the common dialog wrapper
    Dim sFile() As String
    
    If PhotoDemon_OpenImageDialog(sFile, GetModalOwner().hWnd) Then LoadFileAsNewImage sFile
    
End Sub

'Pass this function a string array, and it will fill it with a list of files selected by the user.
' The commondialog filters are automatically set according to image formats supported by the program.
Public Function PhotoDemon_OpenImageDialog(ByRef listOfFiles() As String, ByVal ownerHwnd As Long) As Boolean

    'Disable user input until the dialog closes
    Interface.DisableUserInput
        
    'Get the last "open image" path from the preferences file
    Dim tempPathString As String
    tempPathString = g_UserPreferences.GetPref_String("Paths", "Open Image", "")
    
    'Prep a common dialog interface
    Dim openDialog As pdOpenSaveDialog
    Set openDialog = New pdOpenSaveDialog
        
    Dim sFileList As String
        
    'Retrieve one (or more) files to open
    If openDialog.GetOpenFileName(sFileList, , True, True, g_ImageFormats.getCommonDialogInputFormats, g_LastOpenFilter, tempPathString, g_Language.TranslateMessage("Open an image"), , ownerHwnd) Then
        
        'Message "Preparing to load image..."
        
        'Take the return string (a null-delimited list of filenames) and split it out into a string array
        listOfFiles = Split(sFileList, vbNullChar)
        
        Dim i As Long
        
        'Due to the buffering required by the API call, uBound(listOfFiles) should ALWAYS > 0 but
        ' let's check it anyway (just to be safe)
        If UBound(listOfFiles) > 0 Then
        
            'Remove all empty strings from the array (which are a byproduct of the aforementioned buffering)
            For i = UBound(listOfFiles) To 0 Step -1
                If listOfFiles(i) <> "" Then Exit For
            Next
            
            'With all the empty strings removed, all that's left is legitimate file paths
            ReDim Preserve listOfFiles(0 To i) As String
            
        End If
        
        'If multiple files were selected, we need to do some additional processing to the array
        If UBound(listOfFiles) > 0 Then
        
            'The common dialog function returns a unique array. Index (0) contains the folder path (without a
            ' trailing backslash), so first things first - add a trailing backslash
            Dim imagesPath As String
            imagesPath = FixPath(listOfFiles(0))
            
            'The remaining indices contain a filename within that folder.  To get the full filename, we must
            ' append the path from (0) to the start of each filename.  This will relieve the burden on
            ' whatever function called us - it can simply loop through the full paths, loading files as it goes
            For i = 1 To UBound(listOfFiles)
                listOfFiles(i - 1) = imagesPath & listOfFiles(i)
            Next i
            
            ReDim Preserve listOfFiles(0 To UBound(listOfFiles) - 1)
            
            'Save the new directory as the default path for future usage
            g_UserPreferences.SetPref_String "Paths", "Open Image", imagesPath
            
        'If there is only one file in the array (e.g. the user only opened one image), we don't need to do all
        ' that extra processing - just save the new directory to the preferences file
        Else
        
            'Save the new directory as the default path for future usage
            tempPathString = listOfFiles(0)
            StripDirectory tempPathString
        
            g_UserPreferences.SetPref_String "Paths", "Open Image", tempPathString
            
        End If
        
        'Also, remember the file filter for future use (in case the user tends to use the same filter repeatedly)
        g_UserPreferences.SetPref_Long "Core", "Last Open Filter", g_LastOpenFilter
        
        'All done!
        PhotoDemon_OpenImageDialog = True
        
    'If the user cancels the commondialog box, simply exit out.
    Else
        PhotoDemon_OpenImageDialog = False
    End If
    
    'Re-enable user input
    Interface.EnableUserInput
        
End Function

'Provide a common dialog that allows the user to retrieve a single image filename, which the calling function can
' then use as it pleases.
Public Function PhotoDemon_OpenImageDialog_Simple(ByRef userImagePath As String, ByVal ownerHwnd As Long) As Boolean

    'Disable user input until the dialog closes
    Interface.DisableUserInput
    
    'Common dialog interface
    Dim openDialog As pdOpenSaveDialog
    Set openDialog = New pdOpenSaveDialog
    
    'Get the last "open image" path from the preferences file
    Dim tempPathString As String
    tempPathString = g_UserPreferences.GetPref_String("Paths", "Open Image", "")
        
    'Use Steve McMahon's excellent Common Dialog class to launch a dialog (this way, no OCX is required)
    If openDialog.GetOpenFileName(userImagePath, , True, False, g_ImageFormats.getCommonDialogInputFormats, g_LastOpenFilter, tempPathString, g_Language.TranslateMessage("Select an image"), , ownerHwnd) Then
        
        'Save the new directory as the default path for future usage
        tempPathString = userImagePath
        StripDirectory tempPathString
        g_UserPreferences.SetPref_String "Paths", "Open Image", tempPathString
        
        'Also, remember the file filter for future use (in case the user tends to use the same filter repeatedly)
        g_UserPreferences.SetPref_Long "Core", "Last Open Filter", g_LastOpenFilter
        
        'All done!
        PhotoDemon_OpenImageDialog_Simple = True
        
    'If the user cancels the common dialog box, simply exit out
    Else
        
        PhotoDemon_OpenImageDialog_Simple = False
        
    End If
        
    'Re-enable user input
    Interface.EnableUserInput
    
End Function

'Subroutine for saving an image to file.  This function assumes the image already exists on disk and is simply
' being replaced; if the file does not exist on disk, this routine will automatically transfer control to Save As...
' The imageToSave is a reference to an ID in the pdImages() array.  It can be grabbed from the form.Tag value as well.
Public Function MenuSave(ByVal imageID As Long) As Boolean

    If Len(pdImages(imageID).locationOnDisk) = 0 Then
    
        'This image hasn't been saved before.  Launch the Save As... dialog
        MenuSave = MenuSaveAs(imageID)
        
    Else
    
        'This image has been saved before.
        
        Dim dstFilename As String
                
        'If the user has requested that we only save copies of current images, we need to come up with a new filename
        If g_UserPreferences.GetPref_Long("Saving", "Overwrite Or Copy", 0) = 0 Then
            dstFilename = pdImages(imageID).locationOnDisk
        Else
        
            'Determine the destination directory
            Dim tempPathString As String
            tempPathString = pdImages(imageID).locationOnDisk
            StripDirectory tempPathString
            
            'Perform a failsafe check for a filename of some sort.  If this parameter is missing, the common dialog request will fail.
            If Len(pdImages(imageID).originalFileName) = 0 Then
                pdImages(imageID).originalFileName = g_Language.TranslateMessage("New image")
            End If
            
            'Next, determine the target filename
            Dim tempFilename As String
            tempFilename = pdImages(imageID).originalFileName
            
            'Finally, determine the target file extension
            Dim tempExtension As String
            tempExtension = GetExtension(pdImages(imageID).locationOnDisk)
            
            'Now, call the incrementFilename function to find a unique filename of the "filename (n+1)" variety
            dstFilename = tempPathString & incrementFilename(tempPathString, tempFilename, tempExtension) & "." & tempExtension
        
        End If
        
        'Check to see if the image is in a format that potentially provides an "additional settings" prompt.
        ' If it is, the user needs to be prompted at least once for those settings.
        
        'JPEG
        If (pdImages(imageID).currentFileFormat = FIF_JPEG) And (Not pdImages(imageID).imgStorage.getEntry_Boolean("hasSeenJPEGPrompt")) Then
            MenuSave = PhotoDemon_SaveImage(pdImages(imageID), dstFilename, imageID, True)
        
        'JPEG-2000
        ElseIf (pdImages(imageID).currentFileFormat = FIF_JP2) And (Not pdImages(imageID).imgStorage.getEntry_Boolean("hasSeenJP2Prompt")) Then
            MenuSave = PhotoDemon_SaveImage(pdImages(imageID), dstFilename, imageID, True)
            
        'WebP
        ElseIf (pdImages(imageID).currentFileFormat = FIF_WEBP) And (Not pdImages(imageID).imgStorage.getEntry_Boolean("hasSeenWebPPrompt")) Then
            MenuSave = PhotoDemon_SaveImage(pdImages(imageID), dstFilename, imageID, True)
        
        'JXR
        ElseIf (pdImages(imageID).currentFileFormat = FIF_WEBP) And (Not pdImages(imageID).imgStorage.getEntry_Boolean("hasSeenJXRPrompt")) Then
            MenuSave = PhotoDemon_SaveImage(pdImages(imageID), dstFilename, imageID, True)
        
        'All other formats
        Else
            MenuSave = PhotoDemon_SaveImage(pdImages(imageID), dstFilename, imageID, False, pdImages(imageID).saveParameters)
            
        End If
    End If

End Function

'Subroutine for displaying a commondialog save box, then saving an image to the specified file
Public Function MenuSaveAs(ByVal imageID As Long) As Boolean
    
    Dim saveFileDialog As pdOpenSaveDialog
    Set saveFileDialog = New pdOpenSaveDialog
    
    'Get the last "save image" path from the preferences file
    Dim tempPathString As String
    tempPathString = g_UserPreferences.GetPref_String("Paths", "Save Image", "")
        
    'g_LastSaveFilter will be set to "-1" if the user has never saved a file before.  If that happens, default to JPEG
    If g_LastSaveFilter = -1 Then
    
        g_LastSaveFilter = g_ImageFormats.getIndexOfOutputFIF(FIF_JPEG) + 1
    
    'Otherwise, set g_LastSaveFilter to this image's current file format, or optionally the last-used format
    Else
    
        'There is a user preference for defaulting to either:
        ' 1) The current image's format (standard behavior)
        ' 2) The last format the user specified in the Save As screen (my preferred behavior)
        ' Use that preference to determine which save filter we select.
        If g_UserPreferences.GetPref_Long("Saving", "Suggested Format", 0) = 0 Then
        
            g_LastSaveFilter = g_ImageFormats.getIndexOfOutputFIF(pdImages(imageID).currentFileFormat) + 1
    
            'The user may have loaded a file format where INPUT is supported but OUTPUT is not.  If this happens,
            ' we need to suggest an alternative format.  Use the color-depth of the current image as our guide.
            If g_LastSaveFilter = 0 Then
                
                '24bpp DIBs default to JPEG
                If pdImages(g_CurrentImage).getCompositeImageColorDepth() = 24 Then
                    g_LastSaveFilter = g_ImageFormats.getIndexOfOutputFIF(FIF_JPEG) + 1
                
                '32bpp DIBs default to PNG
                Else
                    g_LastSaveFilter = g_ImageFormats.getIndexOfOutputFIF(FIF_PNG) + 1
                End If
            
            End If
                    
        'Note that we don't need an "Else" here - the g_LastSaveFilter value will already be present
        End If
    
    End If
    
    'Perform a failsafe check for a filename of some sort.  If this parameter is missing, the common dialog request will fail.
    If Len(pdImages(imageID).originalFileName) = 0 Then
        pdImages(imageID).originalFileName = g_Language.TranslateMessage("New image")
    End If
    
    'Check to see if an image with this filename appears in the save location. If it does, use the incrementFilename
    ' function to append ascending numbers (of the format "_(#)") to the filename until a unique filename is found.
    Dim sFile As String
    sFile = tempPathString & incrementFilename(tempPathString, pdImages(imageID).originalFileName, g_ImageFormats.getOutputFormatExtension(g_LastSaveFilter - 1))
        
    If saveFileDialog.GetSaveFileName(sFile, , True, g_ImageFormats.getCommonDialogOutputFormats, g_LastSaveFilter, tempPathString, g_Language.TranslateMessage("Save an image"), g_ImageFormats.getCommonDialogDefaultExtensions, FormMain.hWnd) Then
                
        'Store the selected file format to the image object
        pdImages(imageID).currentFileFormat = g_ImageFormats.getOutputFIF(g_LastSaveFilter - 1)
        
        'Save the new directory as the default path for future usage
        tempPathString = sFile
        StripDirectory tempPathString
        g_UserPreferences.SetPref_String "Paths", "Save Image", tempPathString
        
        'Also, remember the file filter for future use (in case the user tends to use the same filter repeatedly)
        g_UserPreferences.SetPref_Long "Core", "Last Save Filter", g_LastSaveFilter
                        
        'Transfer control to the core SaveImage routine, which will handle color depth analysis and actual saving
        MenuSaveAs = PhotoDemon_SaveImage(pdImages(imageID), sFile, imageID, True)
        
    Else
        MenuSaveAs = False
    End If
            
End Function

'Save a lossless copy of the current image.  I've debated a lot of small details about how to best implement this (e.g. how to
' "most intuitively" implement this), and I've settled on the following:
' 1) Save the copy to the same folder as the current image (if available).  If it's not available, we have no choice but to
'     prompt for a folder.
' 2) Save the image in PDI format.
' 3) Update the Recent Files list with the saved copy.  If we don't do this, the user has no way of knowing what save settings
'     we've used (filename, location, etc)
' 4) Increment the filename automatically.  Saving a copy does not overwrite old copies.  This is important.
Public Function MenuSaveLosslessCopy(ByVal imageID As Long) As Boolean

    'First things first: see if the image currently exists on-disk.  If it doesn't, we have no choice but to provide a save
    ' prompt.
    If Len(pdImages(imageID).locationOnDisk) = 0 Then
        
        'TODO: make this a dialog with a "check to remember" option.  I'm waiting on this because I want a generic solution
        '       for these types of dialogs, because they would be helpful in many places throughout PD.
        PDMsgBox "Before lossless copies can be saved, you must save this image at least once." & vbCrLf & vbCrLf & "Lossless copies will be saved to the same folder as this initial image save.", vbInformation + vbOKOnly + vbApplicationModal, "Initial save required"
        
        'This image hasn't been saved before.  Launch the Save As... dialog, and wait for it to return.
        MenuSaveLosslessCopy = MenuSaveAs(imageID)
        
        'If the user canceled, abandon ship
        If Not MenuSaveLosslessCopy Then Exit Function
        
    End If
    
    'If we made it here, this image has been saved before.  That gives us a folder where we can place our lossless copies.
    Dim dstFilename As String, tmpPathString As String
    
    'Determine the destination directory now
    tmpPathString = pdImages(imageID).locationOnDisk
    StripDirectory tmpPathString
    
    'Next, let's determine the target filename.  This is the current filename, auto-incremented to whatever number is
    ' available next.
    Dim tmpFilename As String
    tmpFilename = pdImages(imageID).originalFileName
    
    'Now, call the incrementFilename function to find a unique filename of the "filename (n+1)" variety, with the PDI
    ' file extension forcibly applied.
    dstFilename = tmpPathString & incrementFilename(tmpPathString, tmpFilename, "pdi") & "." & "pdi"
    
    'dstFilename now contains the full path and filename where our image copy should go.  Save it!
    If g_ZLibEnabled Then
        Saving.beginSaveProcess
        MenuSaveLosslessCopy = SavePhotoDemonImage(pdImages(imageID), dstFilename, , , , , , True)
    Else
    
        'If zLib doesn't exist...
        PDMsgBox "The zLib compression library (zlibwapi.dll) was marked as missing or disabled upon program initialization." & vbCrLf & vbCrLf & "To enable PDI saving, please allow %1 to download plugin updates by going to the Tools -> Options menu, and selecting the 'offer to download core plugins' check box.", vbExclamation + vbOKOnly + vbApplicationModal, " PDI Interface Error", PROGRAMNAME
        Message "No %1 encoder found. Save aborted.", "PDI"
        Saving.endSaveProcess
        MenuSaveLosslessCopy = False
        
        Exit Function
        
    End If
        
    'At this point, it's safe to re-enable the main form and restore the default cursor
    Saving.endSaveProcess
    
    'MenuSaveLosslessCopy should only be true if the save was successful
    If MenuSaveLosslessCopy Then
        
        'Add this file to the MRU list
        g_RecentFiles.MRU_AddNewFile dstFilename, pdImages(imageID)
        
        'Return SUCCESS!
        MenuSaveLosslessCopy = True
        
    Else
        
        Message "Save canceled."
        PDMsgBox "An unspecified error occurred when attempting to save this image.  Please try saving the image to an alternate format." & vbCrLf & vbCrLf & "If the problem persists, please report it to the PhotoDemon developers via photodemon.org/contact", vbCritical Or vbApplicationModal Or vbOKOnly, "Image save error"
        MenuSaveLosslessCopy = False
        
    End If

End Function

'Close the active image
Public Sub MenuClose()
    
    'Make sure the correct flag is set so that the MDI Child QueryUnload behaves properly (e.g. note that we
    ' are not closing ALL images - just this one.)
    g_ClosingAllImages = False
    FullPDImageUnload g_CurrentImage
    
    'Reset any relevant parameters
    g_ClosingAllImages = False
    g_DealWithAllUnsavedImages = False

End Sub

'Close all active images
Public Sub MenuCloseAll()

    'Note that the user has opted to close ALL open images; this is used by the central image handler to know what kind
    ' of "Unsaved changes" dialog to display.
    g_ClosingAllImages = True

    'Loop through each image object and close their associated forms
    Dim i As Long
    For i = LBound(pdImages) To UBound(pdImages)
    
        Message "Unloading image %1 of %2", i, g_NumOfImagesLoaded
    
        If Not (pdImages(i) Is Nothing) Then
            If pdImages(i).IsActive Then
                FullPDImageUnload i, False
            End If
        End If
        
        'If the user presses "cancel" at some point in the unload chain, obey their request immediately
        ' (e.g. stop unloading images)
        If Not g_ClosingAllImages Then Exit For
        
    Next i
    
    'Redraw the screen to match the new program state
    toolbar_ImageTabs.forceRedraw
    SyncInterfaceToCurrentImage
    
    'Reset the "closing all images" flags
    g_ClosingAllImages = False
    g_DealWithAllUnsavedImages = False

End Sub

'Create a new, blank image from scratch
Public Function CreateNewImage(ByVal imgWidth As Long, ByVal imgHeight As Long, ByVal imgDPI As Long, ByVal defaultBackground As Long, ByVal backgroundColor As Long)

    'Display a busy cursor
    If Screen.MousePointer <> vbHourglass Then Screen.MousePointer = vbHourglass
    
    'To prevent re-entry problems, forcibly disable the main form
    FormMain.Enabled = False
    
    'Create a new entry in the pdImages() array.  This will update g_CurrentImage as well.
    CreateNewPDImage
    pdImages(g_CurrentImage).loadedSuccessfully = True
    
    'We can now address our new image via pdImages(g_CurrentImage).  Create a blank layer.
    Dim newLayerID As Long
    newLayerID = pdImages(g_CurrentImage).createBlankLayer()
    
    'The parameters passed to the new DIB vary according to layer type.  Use the specified type to determine how we
    ' initialize the new layer.
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    Select Case defaultBackground
    
        'Transparent (blank)
        Case 0
            tmpDIB.createBlank imgWidth, imgHeight, 32, 0, 0
        
        'Black
        Case 1
            tmpDIB.createBlank imgWidth, imgHeight, 32, vbBlack, 255
        
        'White
        Case 2
            tmpDIB.createBlank imgWidth, imgHeight, 32, vbWhite, 255
        
        'Custom color
        Case 3
            tmpDIB.createBlank imgWidth, imgHeight, 32, backgroundColor, 255
    
    End Select
    
    'Assign the newly created DIB to the layer object
    pdImages(g_CurrentImage).getLayerByID(newLayerID).InitializeNewLayer PDL_IMAGE, g_Language.TranslateMessage("Background"), tmpDIB
    
    'Make the newly created layer the active layer
    setActiveLayerByID newLayerID, False
    
    'Update the pdImage container to be the same size as its (newly created) base layer
    pdImages(g_CurrentImage).updateSize
    
    'Assign the requested DPI to the new image
    pdImages(g_CurrentImage).setDPI imgDPI, imgDPI, False
    
    'Disable viewport rendering, then reset the main viewport
    g_AllowViewportRendering = False
    FormMain.mainCanvas(0).setScrollValue PD_BOTH, 0
    
    'By default, set this image to use the program's default metadata setting (settable from Tools -> Options).
    ' The user may override this setting later, but by default we always start with the user's program-wide setting.
    pdImages(g_CurrentImage).imgMetadata.setMetadataExportPreference g_UserPreferences.GetPref_Long("Saving", "Metadata Export", 1)
    
    'Default to JPEGs, for convenience.  Note that a different format will be suggested at save-time, contingent on the image's state,
    pdImages(g_CurrentImage).originalFileFormat = FIF_JPEG
    pdImages(g_CurrentImage).currentFileFormat = pdImages(g_CurrentImage).originalFileFormat
    pdImages(g_CurrentImage).originalColorDepth = 32
    
    'Because this image does not exist on the user's hard-drive, we will force use of a full Save As dialog in the future.
    ' (PD detects this state if a pdImage object does not supply a location on disk)
    pdImages(g_CurrentImage).locationOnDisk = ""
    pdImages(g_CurrentImage).originalFileNameAndExtension = g_Language.TranslateMessage("New image")
    pdImages(g_CurrentImage).originalFileName = pdImages(g_CurrentImage).originalFileNameAndExtension
    pdImages(g_CurrentImage).setSaveState False, pdSE_AnySave
    
    'Create an icon-sized version of this image, which we will use as form's taskbar icon
    createCustomFormIcon pdImages(g_CurrentImage)
    
    'Register this image with the image tab bar
    toolbar_ImageTabs.registerNewImage g_CurrentImage
    
    'Just to be safe, update the color management profile of the current monitor
    CheckParentMonitor True
    
    'If the user wants us to resize the image to fit on-screen, do that now
    If g_AutozoomLargeImages = 0 Then FitImageToViewport True
    
    'g_AllowViewportRendering may have been reset by this point (by the FitImageToViewport sub, among others), so set it back to False, then
    ' update the zoom combo box to match the zoom assigned by the window-fit function.
    g_AllowViewportRendering = False
    FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex = pdImages(g_CurrentImage).currentZoomValue

    'Now that the image's window has been fully sized and moved around, use Viewport_Engine.Stage1_InitializeBuffer to set up any scrollbars and a back-buffer
    g_AllowViewportRendering = True
    Viewport_Engine.Stage1_InitializeBuffer pdImages(g_CurrentImage), FormMain.mainCanvas(0), VSR_ResetToZero
    
    'Reflow any image-window-specific display elements on the actual image form (status bar, rulers, etc)
    FormMain.mainCanvas(0).fixChromeLayout
    
    'Force an immediate Undo/Redo write to file.  This serves multiple purposes: it is our baseline for calculating future
    ' Undo/Redo diffs, and it can be used to recover the original file if something goes wrong before the user performs a
    ' manual save (e.g. AutoSave).
    pdImages(g_CurrentImage).undoManager.createUndoData g_Language.TranslateMessage("Original image"), "", UNDO_EVERYTHING
    
    'Re-enable the main form
    FormMain.Enabled = True
    
    'Synchronize all interface elements to match the newly created image
    SyncInterfaceToCurrentImage
    toolbar_ImageTabs.forceRedraw
    
    'Restore the default cursor
    Screen.MousePointer = vbNormal
    
    'Report success!
    CreateNewImage = True

End Function
