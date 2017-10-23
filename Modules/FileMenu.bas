Attribute VB_Name = "FileMenu"
'***************************************************************************
'File Menu Handler
'Copyright 2001-2017 by Tanner Helland
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

    Dim listOfFiles As pdStringStack
    If PhotoDemon_OpenImageDialog(listOfFiles, GetModalOwner().hWnd) Then
        
        If (listOfFiles.GetNumOfStrings > 1) Then
            Loading.LoadMultipleImageFiles listOfFiles
        Else
            Loading.LoadFileAsNewImage listOfFiles.GetString(0)
        End If
        
    End If
    
End Sub

'Pass this function a string array, and it will fill it with a list of files selected by the user.
' The commondialog filters are automatically set according to image formats supported by the program.
Public Function PhotoDemon_OpenImageDialog(ByRef dstStringStack As pdStringStack, ByVal ownerHwnd As Long) As Boolean
    
    If (dstStringStack Is Nothing) Then Set dstStringStack = New pdStringStack
    
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
    If openDialog.GetOpenFileName(sFileList, , True, True, g_ImageFormats.GetCommonDialogInputFormats, g_LastOpenFilter, tempPathString, g_Language.TranslateMessage("Open an image"), , ownerHwnd) Then
        
        'Message "Preparing to load image..."
        
        'Take the return string (a null-delimited list of filenames) and split it out into a string array
        Dim listOfFiles() As String
        listOfFiles = Split(sFileList, vbNullChar)
        
        Dim i As Long
        
        'Due to the buffering required by the API call, uBound(listOfFiles) should ALWAYS > 0 but
        ' let's check it anyway (just to be safe)
        If UBound(listOfFiles) > 0 Then
        
            'Remove all empty strings from the array (which are a byproduct of the aforementioned buffering)
            For i = UBound(listOfFiles) To 0 Step -1
                If Len(listOfFiles(i)) <> 0 Then Exit For
            Next
            
            'With all the empty strings removed, all that's left is legitimate file paths
            ReDim Preserve listOfFiles(0 To i) As String
            
        End If
        
        'If multiple files were selected, we need to do some additional processing to the array
        If UBound(listOfFiles) > 0 Then
        
            'The common dialog function returns a unique array. Index (0) contains the folder path (without a
            ' trailing backslash), so first things first - add a trailing backslash
            Dim imagesPath As String
            imagesPath = Files.PathAddBackslash(listOfFiles(0))
            
            'The remaining indices contain a filename within that folder.  To get the full filename, we must
            ' append the path from (0) to the start of each filename.  This will relieve the burden on
            ' whatever function called us - it can simply loop through the full paths, loading files as it goes
            For i = 1 To UBound(listOfFiles)
                dstStringStack.AddString imagesPath & listOfFiles(i)
            Next i
            
            'Save the new directory as the default path for future usage
            g_UserPreferences.SetPref_String "Paths", "Open Image", imagesPath
            
        'If there is only one file in the array (e.g. the user only opened one image), we don't need to do all
        ' that extra processing - just save the new directory to the preferences file
        Else
        
            'Save the new directory as the default path for future usage
            tempPathString = Files.FileGetPath(listOfFiles(0))
            g_UserPreferences.SetPref_String "Paths", "Open Image", tempPathString
            
            dstStringStack.AddString listOfFiles(0)
            
        End If
        
        'Copy the raw string array into an iteration-friendly string stack
        
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
    If openDialog.GetOpenFileName(userImagePath, , True, False, g_ImageFormats.GetCommonDialogInputFormats, g_LastOpenFilter, tempPathString, g_Language.TranslateMessage("Select an image"), , ownerHwnd) Then
        
        'Save the new directory as the default path for future usage
        tempPathString = Files.FileGetPath(userImagePath)
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
Public Function MenuSave(ByRef srcImage As pdImage) As Boolean
    
    'Certain criteria make is impossible to blindly save an image to disk (such as the image being loaded from a
    ' non-disk source, like the clipbord).  When this happens, we'll silently switch to a Save As... dialog.
    If Saving.IsCommonDialogRequired(srcImage) Then
        MenuSave = MenuSaveAs(srcImage)
    
    'This image has been saved before, meaning it already exists on disk.
    Else
        
        Dim dstFilename As String
        
        'PhotoDemon supports two different save modes (controlled via the Tools > Options dialog):
        ' 1) Default mode.  When the user clicks "save", overwrite the copy on disk.
        ' 2) "Safe" mode.  When the user clicks "save", save a new copy of the image, auto-incremented with a trailing number.
        '    (e.g. old copies are never overwritten).
        Dim safeSaveModeActive As Boolean
        safeSaveModeActive = CBool(g_UserPreferences.GetPref_Long("Saving", "Overwrite Or Copy", 0) <> 0)
        
        If safeSaveModeActive Then
        
            'File name incrementation requires help from an outside function.  We must pass it the folder, filename, and extension
            ' we want it to search against.
            Dim tmpFolder As String, tmpFilename As String, tmpExtension As String
            tmpFolder = Files.FileGetPath(srcImage.ImgStorage.GetEntry_String("CurrentLocationOnDisk", vbNullString))
            If Len(srcImage.ImgStorage.GetEntry_String("OriginalFileName", vbNullString)) = 0 Then srcImage.ImgStorage.AddEntry "OriginalFileName", g_Language.TranslateMessage("New image")
            tmpFilename = srcImage.ImgStorage.GetEntry_String("OriginalFileName", vbNullString)
            tmpExtension = srcImage.ImgStorage.GetEntry_String("OriginalFileExtension", vbNullString)
            
            'Now, call the incrementFilename function to find a unique filename of the "filename (n+1)" variety
            dstFilename = tmpFolder & Files.IncrementFilename(tmpFolder, tmpFilename, tmpExtension) & "." & tmpExtension
        
        Else
            dstFilename = srcImage.ImgStorage.GetEntry_String("CurrentLocationOnDisk", vbNullString)
        End If
        
        'New to v7.0 is the way save option dialogs work.  PD's primary save function is now responsible for displaying save dialogs.
        ' (We can forcibly request a dialog, as we do in the "Save As" function, but in this function, we leave it up to the primary
        ' save function to determine if a dialog is necessary.)
        MenuSave = PhotoDemon_SaveImage(srcImage, dstFilename, False)
        
    End If

End Function

'Subroutine for displaying a commondialog save box, then saving an image to the specified file
Public Function MenuSaveAs(ByRef srcImage As pdImage) As Boolean
    
    Dim saveFileDialog As pdOpenSaveDialog
    Set saveFileDialog = New pdOpenSaveDialog
    
    'Prior to showing the "save image" dialog, we need to determine three things:
    ' 1) An initial folder
    ' 2) What file format to suggest
    ' 3) What filename to suggest (*without* a file extension)
    ' 4) What filename + extension to suggest, based on the results of 2 and 3
    
    'Each of these will be handled in turn
    
    '1) Determine an initial folder.  This is easy, as we will just grab the last "save image" path from the preferences file.
    '   (The preferences engine will automatically pass us the user's Pictures folder if no "last path" entry exists.)
    Dim initialSaveFolder As String
    initialSaveFolder = g_UserPreferences.GetPref_String("Paths", "Save Image", "")
    
    '2) What file format to suggest.  There is a user preference for persistently defaulting not to the current image's suggested format,
    '   but to the last format used in the Save screen.  (This is useful when mass-converting RAW files to JPEG, for example.)
    '   If that preference is selected, it takes precedence, unless the user has not yet saved any images, in which case we default to
    '   the standard method (of using heuristics on the current image, and suggesting the most appropriate format accordingly).
    Dim cdFormatIndex As Long
    Dim suggestedSaveFormat As PD_IMAGE_FORMAT, suggestedFileExtension As String
    
    If (g_UserPreferences.GetPref_Long("Saving", "Suggested Format", 0) = 1) And (g_LastSaveFilter <> -1) Then
        cdFormatIndex = g_LastSaveFilter
        suggestedSaveFormat = g_ImageFormats.GetOutputPDIF(cdFormatIndex - 1)
        suggestedFileExtension = g_ImageFormats.GetExtensionFromPDIF(suggestedSaveFormat)
        
    'The user's preference is the default value (0) or no previous saves have occurred, meaning we need to suggest a Save As format based
    ' on the current image contents.  This is a fairly complex process, so we offload it to a separate function.
    Else
        suggestedSaveFormat = GetSuggestedSaveFormatAndExtension(srcImage, suggestedFileExtension)
        
        'Now that we have a suggested save format, we need to convert that into its matching Common Dialog filter index.
        ' (Note that the common dialog filter is 1-based, so we manually increment the retrieved index.)
        cdFormatIndex = g_ImageFormats.GetIndexOfOutputPDIF(suggestedSaveFormat) + 1
    End If
    
    '3) What filename to suggest.  This value is pulled from the image storage object; if this file came from a non-file location
    '   (like the clipboard), that function will have supplied a meaningful name at load-time.  Note that we have to supply a non-null
    '   string to the common dialog function for it to work, so some kind of name needs to be suggested.
    Dim suggestedFilename As String
    suggestedFilename = srcImage.ImgStorage.GetEntry_String("OriginalFileName", vbNullString)
    If (Len(suggestedFilename) = 0) Then suggestedFilename = g_Language.TranslateMessage("New image")
    
    '4) What filename + extension to suggest, based on the results of 2 and 3.  Most programs would just toss together the
    ' calculated filename + extension, but I like PD to be a bit smarter.  What we're going to do next is scan the default output
    ' folder to see if any files already match this name and extension.  If they do, we're going to append a number to the end of
    ' the filename, e.g. "New Image (2)", and we're going to auto-increment that number until we find a number that isn't in use.
    ' (If auto-incrementing isn't necessary, this function will return the filename we pass it, as-is.)
    Dim sFile As String
    sFile = initialSaveFolder & IncrementFilename(initialSaveFolder, suggestedFilename, suggestedFileExtension)
    
    'With all our inputs complete, we can finally raise the damn common dialog
    If saveFileDialog.GetSaveFileName(sFile, , True, g_ImageFormats.GetCommonDialogOutputFormats, cdFormatIndex, initialSaveFolder, g_Language.TranslateMessage("Save an image"), g_ImageFormats.GetCommonDialogDefaultExtensions, FormMain.hWnd) Then
        
        'The common dialog results affect two different objects:
        ' 1) the current image (which needs to store things like the format the user chose)
        ' 2) the global user preferences manager (which needs to remember things like the output folder, so we can remember it)
        
        'Store all image-level attributes
        srcImage.SetCurrentFileFormat g_ImageFormats.GetOutputPDIF(cdFormatIndex - 1)
        
        'Store all global-preference attributes
        g_LastSaveFilter = cdFormatIndex
        g_UserPreferences.SetPref_Long "Core", "Last Save Filter", g_LastSaveFilter
        g_UserPreferences.SetPref_String "Paths", "Save Image", Files.FileGetPath(sFile)
        
        'Our work here is done!  Transfer control to the core SaveImage routine, which will handle the actual export process.
        MenuSaveAs = PhotoDemon_SaveImage(srcImage, sFile, True)
        
    Else
        MenuSaveAs = False
    End If
    
End Function

Private Function GetSuggestedSaveFormatAndExtension(ByRef srcImage As pdImage, ByRef dstSuggestedExtension As String) As PD_IMAGE_FORMAT
    
    'First, see if the image has a file format already.  If it does, we need to suggest that preferentially
    GetSuggestedSaveFormatAndExtension = srcImage.GetCurrentFileFormat
    If (GetSuggestedSaveFormatAndExtension = PDIF_UNKNOWN) Then
    
        'This image must have come from a source where the best save format isn't clear (like a generic clipboard DIB).
        ' As such, we need to suggest an appropriate format.
        
        'Start with the most obvious criteria: does the image have multiple layers?  If so, PDI is best.
        If srcImage.GetNumOfLayers > 1 Then
            GetSuggestedSaveFormatAndExtension = PDIF_PDI
        Else
        
            'Query the only layer in the image.  If it has meaningful alpha values, we'll suggest PNG; otherwise, JPEG.
            If DIBs.IsDIBAlphaBinary(srcImage.GetActiveDIB, False) Then
                GetSuggestedSaveFormatAndExtension = PDIF_JPEG
            Else
                GetSuggestedSaveFormatAndExtension = PDIF_PNG
            End If
        
        End If
        
        'Also return a proper extension that matches the selected format
        dstSuggestedExtension = g_ImageFormats.GetExtensionFromPDIF(GetSuggestedSaveFormatAndExtension)
        
    'If the image already has a format, let's reuse its existing file extension instead of suggesting a new one.  This is relevant
    ' for formats with ill-defined extensions, like JPEG (e.g. JPE, JPG, JPEG)
    Else
        dstSuggestedExtension = srcImage.ImgStorage.GetEntry_String("OriginalFileExtension")
        If Len(dstSuggestedExtension) = 0 Then dstSuggestedExtension = g_ImageFormats.GetExtensionFromPDIF(GetSuggestedSaveFormatAndExtension)
    End If
            
End Function

'TODO 7.0: review this function to make sure it works with the new save engine!
'Save a lossless copy of the current image.  I've debated a lot of small details about how to best implement this (e.g. how to
' "most intuitively" implement this), and I've settled on the following:
' 1) Save the copy to the same folder as the current image (if available).  If it's not available, we have no choice but to
'     prompt for a folder.
' 2) Save the image in PDI format.
' 3) Update the Recent Files list with the saved copy.  If we don't do this, the user has no way of knowing what save settings
'     we've used (filename, location, etc)
' 4) Increment the filename automatically.  Saving a copy does not overwrite old copies.  This is important.
Public Function MenuSaveLosslessCopy(ByRef srcImage As pdImage) As Boolean

    'First things first: see if the image currently exists on-disk.  If it doesn't, we have no choice but to provide a save
    ' prompt.
    If Len(srcImage.ImgStorage.GetEntry_String("CurrentLocationOnDisk", vbNullString)) = 0 Then
        
        'TODO: make this a dialog with a "check to remember" option.  I'm waiting on this because I want a generic solution
        '       for these types of dialogs, because they would be helpful in many places throughout PD.
        PDMsgBox "Before lossless copies can be saved, you must save this image at least once." & vbCrLf & vbCrLf & "Lossless copies will be saved to the same folder as this initial image save.", vbExclamation Or vbOKOnly, "Initial save required"
        
        'This image hasn't been saved before.  Launch the Save As... dialog, and wait for it to return.
        MenuSaveLosslessCopy = MenuSaveAs(srcImage)
        
        'If the user canceled, abandon ship
        If (Not MenuSaveLosslessCopy) Then Exit Function
        
    End If
    
    'If we made it here, this image has been saved before.  That gives us a folder where we can place our lossless copies.
    Dim dstFilename As String, tmpPathString As String
    
    'Determine the destination directory now
    tmpPathString = Files.FileGetPath(srcImage.ImgStorage.GetEntry_String("CurrentLocationOnDisk", vbNullString))
    
    'Next, let's determine the target filename.  This is the current filename, auto-incremented to whatever number is
    ' available next.
    Dim tmpFilename As String
    tmpFilename = srcImage.ImgStorage.GetEntry_String("OriginalFileName", vbNullString)
    
    'Now, call the incrementFilename function to find a unique filename of the "filename (n+1)" variety, with the PDI
    ' file extension forcibly applied.
    dstFilename = tmpPathString & IncrementFilename(tmpPathString, tmpFilename, "pdi") & "." & "pdi"
    
    'dstFilename now contains the full path and filename where our image copy should go.  Save it!
    Saving.BeginSaveProcess
    MenuSaveLosslessCopy = SavePhotoDemonImage(srcImage, dstFilename, , , , False, True)
        
    'At this point, it's safe to re-enable the main form and restore the default cursor
    Saving.EndSaveProcess
    
    'MenuSaveLosslessCopy should only be true if the save was successful
    If MenuSaveLosslessCopy Then
        
        'Add this file to the MRU list
        g_RecentFiles.AddFileToList dstFilename, srcImage
        
        'Return SUCCESS!
        MenuSaveLosslessCopy = True
        
    Else
        
        Message "Save canceled."
        PDMsgBox "An unspecified error occurred when attempting to save this image.  Please try saving the image to an alternate format." & vbCrLf & vbCrLf & "If the problem persists, please report it to the PhotoDemon developers via photodemon.org/contact", vbCritical Or vbOKOnly, "Image save error"
        MenuSaveLosslessCopy = False
        
    End If

End Function

'Close the active image
Public Sub MenuClose()
    
    'Just in case, reset the "user is trying to close all images" flag
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
    
    Dim numOfImagesToUnload As Long, numImagesActuallyUnloaded As Long
    numImagesActuallyUnloaded = 0
    numOfImagesToUnload = g_OpenImageCount
    
    'Loop through each image object and close their associated forms.  Note that we want to start from the current image and
    ' work our way down.
    Dim i As Long, startingIndex As Long
    startingIndex = g_CurrentImage
    i = startingIndex
    
    Dim keepUnloading As Boolean: keepUnloading = True
    
    Do
    
        If (Not pdImages(i) Is Nothing) Then
            If pdImages(i).IsActive Then
                numImagesActuallyUnloaded = numImagesActuallyUnloaded + 1
                Message "Unloading image %1 of %2", numImagesActuallyUnloaded, numOfImagesToUnload
                FullPDImageUnload i, False
            End If
        End If
        
        'If the user presses "cancel" at some point in the unload chain, obey their request immediately
        ' (e.g. stop unloading images)
        If (Not g_ClosingAllImages) Then
            If (g_OpenImageCount <> 0) Then Message ""
            Exit Do
            
        'If the unload process hasn't been canceled, move to the next image
        Else
        
            i = i + 1
            If (i > UBound(pdImages)) Then i = LBound(pdImages)
            If (i = startingIndex) Then keepUnloading = False
        
        End If
    
    Loop While keepUnloading
    
    'Redraw the screen to match the new program state
    SyncInterfaceToCurrentImage
    
    'Reset the "closing all images" flags
    g_ClosingAllImages = False
    g_DealWithAllUnsavedImages = False

End Sub

'Create a new, blank image from scratch.  Incoming parameters must be assembled as XML (via pdParamXML, typically)
Public Function CreateNewImage(Optional ByVal newImageParameters As String)
    
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString newImageParameters
    
    Dim newWidth As Long, newHeight As Long, newDPI As Long
    Dim newBackgroundType As Long, newBackgroundColor As Long
    
    With cParams
        newWidth = .GetLong("WidthInPixels", g_Displays.GetDesktopWidth)
        newHeight = .GetLong("HeightInPixels", g_Displays.GetDesktopHeight)
        newDPI = .GetLong("DPI", 96&)
        newBackgroundType = .GetLong("BackgroundType", 0)
        newBackgroundColor = .GetLong("OptionalBackcolor", vbBlack)
    End With
    
    'Display a busy cursor and disable user input
    Processor.MarkProgramBusyState True, True
    
    'Create a new entry in the pdImages() array.  This will update g_CurrentImage as well.
    Dim newImage As pdImage
    CanvasManager.GetDefaultPDImageObject newImage
    
    'We can now address our new image via pdImages(g_CurrentImage).  Create a blank layer.
    Dim newLayerID As Long
    newLayerID = newImage.CreateBlankLayer()
    
    'The parameters passed to the new DIB vary according to layer type.  Use the specified type to determine how we
    ' initialize the new layer.
    Dim newBackColor As Long, newBackAlpha As Long
    Select Case newBackgroundType
    
        'Transparent (blank)
        Case 0
            newBackColor = vbBlack
            newBackAlpha = 0
            
        'Black
        Case 1
            newBackColor = vbBlack
            newBackAlpha = 255
        
        'White
        Case 2
            newBackColor = vbWhite
            newBackAlpha = 255
        
        'Custom color
        Case 3
            newBackColor = newBackgroundColor
            newBackAlpha = 255
    
    End Select
    
    'Create a matching DIB
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    If tmpDIB.CreateBlank(newWidth, newHeight, 32, newBackColor, newBackAlpha) Then
    
        'Assign the newly created DIB to the layer object
        tmpDIB.SetInitialAlphaPremultiplicationState True
        newImage.GetLayerByID(newLayerID).InitializeNewLayer PDL_IMAGE, g_Language.TranslateMessage("Background"), tmpDIB
        
        'Make the newly created layer the active layer
        CanvasManager.AddImageToMasterCollection newImage
        Layers.SetActiveLayerByID newLayerID, False, False
        
        'Update the pdImage container to be the same size as its (newly created) base layer
        newImage.UpdateSize
        
        'Assign the requested DPI to the new image
        newImage.SetDPI newDPI, newDPI, False
        
        'Disable viewport rendering, then reset the main viewport
        ViewportEngine.DisableRendering
        FormMain.mainCanvas(0).SetScrollValue PD_BOTH, 0
        
        'Reset the file format markers; at save-time engine, PD will run heuristics on the image's contents and suggest a better format accordingly.
        newImage.SetOriginalFileFormat PDIF_UNKNOWN
        newImage.SetCurrentFileFormat PDIF_UNKNOWN
        newImage.SetOriginalColorDepth 32
        
        'Because this image does not exist on the user's hard-drive, we will force use of a full Save As dialog in the future.
        ' (PD detects this state if a pdImage object does not supply a location on disk)
        newImage.ImgStorage.AddEntry "CurrentLocationOnDisk", ""
        newImage.ImgStorage.AddEntry "OriginalFileName", g_Language.TranslateMessage("New image")
        newImage.ImgStorage.AddEntry "OriginalFileExtension", ""
        newImage.SetSaveState False, pdSE_AnySave
        
        'Make any interface changes related to the presence of a new image
        Interface.NotifyImageAdded g_CurrentImage
        
        'Just to be safe, update the color management profile of the current monitor
        ColorManagement.CheckParentMonitor True
        
        'If the user wants us to resize the image to fit on-screen, do that now
        If (g_AutozoomLargeImages = 0) Then CanvasManager.FitImageToViewport True
        
        'Viewport rendering may have been reset by this point (by the FitImageToViewport sub, among others), so disable it again, then
        ' update the zoom combo box to match the zoom assigned by the window-fit function.
        ViewportEngine.DisableRendering
        FormMain.mainCanvas(0).SetZoomDropDownIndex newImage.GetZoom
        
        'Now that the image's window has been fully sized and moved around, use ViewportEngine.Stage1_InitializeBuffer to set up any scrollbars and a back-buffer
        ViewportEngine.EnableRendering
        ViewportEngine.Stage1_InitializeBuffer newImage, FormMain.mainCanvas(0), VSR_ResetToZero
        
        'Reflow any image-window-specific display elements on the actual image form (status bar, rulers, etc)
        FormMain.mainCanvas(0).UpdateCanvasLayout
        
        'Force an immediate Undo/Redo write to file.  This serves multiple purposes: it is our baseline for calculating future
        ' Undo/Redo diffs, and it can be used to recover the original file if something goes wrong before the user performs a
        ' manual save (e.g. AutoSave).
        newImage.UndoManager.CreateUndoData g_Language.TranslateMessage("Original image"), "", UNDO_Everything
        
        'Report success!
        CreateNewImage = True
        
    Else
        CreateNewImage = False
        PDMsgBox "Unfortunately, this PC does not have enough memory to create a %1x%2 image.  Please reduce the requested size and try again.", vbExclamation Or vbOKOnly, "Image too large", newWidth, newHeight
    End If
    
    'Re-enable the main form
    Processor.MarkProgramBusyState False
    
    'Synchronize all interface elements to match the newly created image
    Interface.SyncInterfaceToCurrentImage
    
End Function
