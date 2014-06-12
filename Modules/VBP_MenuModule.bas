Attribute VB_Name = "File_Menu"
'***************************************************************************
'File Menu Handler
'Copyright ©2001-2014 by Tanner Helland
'Created: 15/Apr/01
'Last updated: 22/May/14
'Last update: add a failsafe check for an image filename prior to requesting a common dialog; unbeknownst to me,
'             passing a blank filename will cause the common dialog request to fail!
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
    
    If PhotoDemon_OpenImageDialog(sFile, getModalOwner().hWnd) Then LoadFileAsNewImage sFile

    Erase sFile

End Sub

'Pass this function a string array, and it will fill it with a list of files selected by the user.
' The commondialog filters are automatically set according to image formats supported by the program.
Public Function PhotoDemon_OpenImageDialog(ByRef listOfFiles() As String, ByVal ownerHwnd As Long) As Boolean

    'Disable user input until the dialog closes
    Interface.disableUserInput
    
    'Common dialog interface
    Dim CC As cCommonDialog
    
    'Get the last "open image" path from the preferences file
    Dim tempPathString As String
    tempPathString = g_UserPreferences.GetPref_String("Paths", "Open Image", "")
    
    Set CC = New cCommonDialog
        
    Dim sFileList As String
    
    'Remove top-most status from any/all windows (toolbars in floating mode, primarily).  If we don't do this, they may
    ' appear over the top of the common dialog.
    g_WindowManager.resetTopmostForAllWindows False
    
    'Use Steve McMahon's excellent Common Dialog class to launch a dialog (this way, no OCX is required)
    If CC.VBGetOpenFileName(sFileList, , True, True, False, True, g_ImageFormats.getCommonDialogInputFormats, g_LastOpenFilter, tempPathString, g_Language.TranslateMessage("Open an image"), , ownerHwnd, 0) Then
        
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
        
    'If the user cancels the commondialog box, simply exit out
    Else
        
        If CC.ExtendedError <> 0 Then pdMsgBox "An error occurred: %1", vbCritical + vbOKOnly + vbApplicationModal, "Common dialog error", CC.ExtendedError
    
        PhotoDemon_OpenImageDialog = False
    End If
    
    'Re-enable user input
    Interface.enableUserInput
    
    'Restore window status
    g_WindowManager.resetTopmostForAllWindows True
    
    'Release the common dialog object
    Set CC = Nothing

End Function

'Provide a common dialog that allows the user to retrieve a single image filename, which the calling function can
' then use as it pleases.
Public Function PhotoDemon_OpenImageDialog_Simple(ByRef userImagePath As String, ByVal ownerHwnd As Long) As Boolean

    'Disable user input until the dialog closes
    Interface.disableUserInput
    
    'Common dialog interface
    Dim CC As cCommonDialog
    
    'Get the last "open image" path from the preferences file
    Dim tempPathString As String
    tempPathString = g_UserPreferences.GetPref_String("Paths", "Open Image", "")
    
    Set CC = New cCommonDialog
    
    'Remove top-most status from any/all windows (toolbars in floating mode, primarily).  If we don't do this, they may
    ' appear over the top of the common dialog.
    g_WindowManager.resetTopmostForAllWindows False
    
    'Use Steve McMahon's excellent Common Dialog class to launch a dialog (this way, no OCX is required)
    If CC.VBGetOpenFileName(userImagePath, , True, False, False, True, g_ImageFormats.getCommonDialogInputFormats, g_LastOpenFilter, tempPathString, g_Language.TranslateMessage("Select an image"), , ownerHwnd, 0) Then
        
        'Because the returned string will be null-padded, we must manually trim it down to only the relevant bits
        userImagePath = TrimNull(userImagePath)
        
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
        
        If CC.ExtendedError <> 0 Then pdMsgBox "An error occurred: %1", vbCritical + vbOKOnly + vbApplicationModal, "Common dialog error", CC.ExtendedError
    
        PhotoDemon_OpenImageDialog_Simple = False
        
    End If
    
    'Restore window status
    g_WindowManager.resetTopmostForAllWindows True
    
    'Re-enable user input
    Interface.enableUserInput
    
    'Release the common dialog object
    Set CC = Nothing

End Function

'Subroutine for saving an image to file.  This function assumes the image already exists on disk and is simply
' being replaced; if the file does not exist on disk, this routine will automatically transfer control to Save As...
' The imageToSave is a reference to an ID in the pdImages() array.  It can be grabbed from the form.Tag value as well.
Public Function MenuSave(ByVal imageID As Long) As Boolean

    If pdImages(imageID).locationOnDisk = "" Then
    
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
        If (pdImages(imageID).currentFileFormat = FIF_JPEG) And (Not pdImages(imageID).imgStorage.Item("hasSeenJPEGPrompt")) Then
            MenuSave = PhotoDemon_SaveImage(pdImages(imageID), dstFilename, imageID, True)
        
        'JPEG-2000
        ElseIf (pdImages(imageID).currentFileFormat = FIF_JP2) And (Not pdImages(imageID).imgStorage.Item("hasSeenJP2Prompt")) Then
            MenuSave = PhotoDemon_SaveImage(pdImages(imageID), dstFilename, imageID, True)
            
        'WebP
        ElseIf (pdImages(imageID).currentFileFormat = FIF_WEBP) And (Not pdImages(imageID).imgStorage.Item("hasSeenWebPPrompt")) Then
            MenuSave = PhotoDemon_SaveImage(pdImages(imageID), dstFilename, imageID, True)
        
        'JXR
        ElseIf (pdImages(imageID).currentFileFormat = FIF_WEBP) And (Not pdImages(imageID).imgStorage.Item("hasSeenJXRPrompt")) Then
            MenuSave = PhotoDemon_SaveImage(pdImages(imageID), dstFilename, imageID, True)
        
        'All other formats
        Else
            MenuSave = PhotoDemon_SaveImage(pdImages(imageID), dstFilename, imageID, False, pdImages(imageID).saveParameters)
            
        End If
    End If

End Function

'Subroutine for displaying a commondialog save box, then saving an image to the specified file
Public Function MenuSaveAs(ByVal imageID As Long) As Boolean
    
    Dim CC As cCommonDialog
    Set CC = New cCommonDialog
    
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
            If g_LastSaveFilter = -1 Then
                
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
    
    'Remove top-most status from any/all windows (toolbars in floating mode, primarily).  If we don't do this, they may
    ' appear over the top of the common dialog.
    g_WindowManager.resetTopmostForAllWindows False
    
    If CC.VBGetSaveFileName(sFile, , True, g_ImageFormats.getCommonDialogOutputFormats, g_LastSaveFilter, tempPathString, g_Language.TranslateMessage("Save an image"), g_ImageFormats.getCommonDialogDefaultExtensions, FormMain.hWnd, 0) Then
                
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
    
    'Restore window status
    g_WindowManager.resetTopmostForAllWindows True
    
    'Release the common dialog object
    Set CC = Nothing
    
End Function

'Close the active image
Public Sub MenuClose()
    
    'Make sure the correct flag is set so that the MDI Child QueryUnload behaves properly (e.g. note that we
    ' are not closing ALL images - just this one.)
    g_ClosingAllImages = False
    fullPDImageUnload g_CurrentImage

End Sub

'Close all active images
Public Sub MenuCloseAll()

    'Note that the user has opted to close ALL open images; this is used by the MDI children to know what kind
    ' of "Unsaved changes" dialog to display.
    g_ClosingAllImages = True

    'Loop through each image object and close their associated forms
    Dim i As Long
    For i = 0 To g_NumOfImagesLoaded
    
        If Not (pdImages(i) Is Nothing) Then
            If pdImages(i).IsActive Then fullPDImageUnload i
        End If
        
        'If the user presses "cancel" at some point in the unload chain, obey their request immediately
        ' (e.g. stop unloading images)
        If Not g_ClosingAllImages Then Exit For
        
    Next i

    'Reset the "closing all images" flags
    g_ClosingAllImages = False
    g_DealWithAllUnsavedImages = False

End Sub
