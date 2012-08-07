Attribute VB_Name = "File_Menu"
'***************************************************************************
'File Menu Handler
'Copyright �2000-2012 by Tanner Helland
'Created: 15/Apr/01
'Last updated: 14/June/12
'Last update: moved the bulk of the Open dialog to a standalone routine, so that it can be used elsewhere
'             in the project (specifically, the Batch Convert form). This way it can benefit from the
'             addition of new image filters, without me having to cut-and-paste the code to that form.
'
'Module for controlling standard file menu options.  Currently only handles
'"open image" and "save image".
'
'***************************************************************************

Option Explicit

'This subroutine loads an image - note that the interesting stuff actually happens in PhotoDemon_OpenImageDialog, below
Public Sub MenuOpen()
    
    'String returned from the common dialog wrapper
    Dim sFile() As String
    
    If PhotoDemon_OpenImageDialog(sFile, FormMain.hWnd) Then PreLoadImage sFile

End Sub

'Pass this function a string array, and it will fill it with a list of files selected by the user.
' The commondialog filters are automatically set according to image formats supported by the program.
Public Function PhotoDemon_OpenImageDialog(ByRef listOfFiles() As String, ByVal ownerHWnd As Long) As Boolean

    'Common dialog interface
    Dim CC As cCommonDialog
    
    'Get the last "open image" path from the INI file
    Dim tempPathString As String
    tempPathString = GetFromIni("Program Paths", "MainOpen")
    
    Set CC = New cCommonDialog
    Dim cdfStr As String
    cdfStr = "All Compatible Images|*.bmp;*.jpg;*.jpeg;*.gif;*.wmf;*.emf;*.ico;*.pcx;*.tga;*.rle"
    
    'Only allow PDI loading if the zLib dll was detected at program load
    If zLibEnabled = True Then cdfStr = cdfStr & ";*.pdi"
    
    'Only allow FreeImage file loading if the FreeImage dll was detected
    If FreeImageEnabled = True Then cdfStr = cdfStr & ";*.png;*.lbm;*.pbm;*.iff;*.jif;*.jfif;*.psd;*.tif;*.tiff;*.wbmp;*.wbm;*.pgm;*.ppm;*.jng;*.mng;*.koa;*.pcd;*.ras;*.dds;*.pict;*.pct;*.pic;*.sgi;*.rgb;*.rgba;*.bw;*.int;*.inta"
    cdfStr = cdfStr & "|"
    cdfStr = cdfStr & "BMP - OS/2 or Windows Bitmap|*.bmp"
    
    If FreeImageEnabled = True Then cdfStr = cdfStr & "|DDS - DirectDraw Surface|*.dds"
    
    cdfStr = cdfStr & "|EMF - Windows Enhanced Meta File|*.emf|GIF - Compuserve|*.gif|ICO - Windows Icon|*.ico"
    
    If FreeImageEnabled = True Then cdfStr = cdfStr & "|IFF - Amiga Interchange Format|*.iff|JNG - JPEG Network Graphics|*.jng"
    
    cdfStr = cdfStr & "|JPG/JPEG - Joint Photographic Experts Group|*.jpg;*.jpeg;*.jif;*.jfif"
    
    If FreeImageEnabled = True Then cdfStr = cdfStr & "|KOA/KOALA - Commodore 64|*.koa;*.koala|LBM - Deluxe Paint|*.lbm|MNG - Multiple Network Graphics|*.mng|PBM - Portable Bitmap|*.pbm|PCD - Kodak PhotoCD|*.pcd"
    
    cdfStr = cdfStr & "|PCX - Zsoft Paintbrush|*.pcx"
    
    'Only allow PDI (PhotoDemon's native file format) loading if the zLib dll has been properly detected
    If zLibEnabled = True Then cdfStr = cdfStr & "|PDI - PhotoDemon Image|*.pdi"
    
    If FreeImageEnabled = True Then cdfStr = cdfStr & "|PGM - Portable Greymap|*.pgm|PIC/PICT - Macintosh Picture|*.pict;*.pct;*.pic"
    
    cdfStr = cdfStr & "|PNG - Portable Network Graphic|*.png"
    
    If FreeImageEnabled = True Then cdfStr = cdfStr & "|PPM - Portable Pixmap|*.ppm|PSD - Adobe Photoshop|*.psd|RAS - Sun Raster File|*.ras"
    
    cdfStr = cdfStr & "|RLE - Compuserve or Windows|*.rle"
    
    If FreeImageEnabled = True Then cdfStr = cdfStr & "|SGI/RGB/BW - Silicon Graphics Image|*.sgi;*.rgb;*.rgba;*.bw;*.int;*.inta|TGA - Truevision Targa|*.tga|TIF/TIFF - Tagged Image File Format|*.tif;*.tiff|WBMP - Wireless Bitmap|*.wbmp;*.wbm"
    
    cdfStr = cdfStr & "|WMF - Windows Metafile|*.wmf|All files|*.*"
    
    Dim sFileList As String
    
    'Use Steve McMahon's excellent Common Dialog class to launch a dialog (this way, no OCX is required)
    If CC.VBGetOpenFileName(sFileList, , True, True, False, True, cdfStr, LastOpenFilter, tempPathString, "Open an image", , ownerHWnd, 0) Then
        
        Message "Preparing to load image..."
        
        'Take the return string (a null-delimited list of filenames) and split it out into a string array
        listOfFiles = Split(sFileList, vbNullChar)
        
        'Due to the buffering required by the API call, uBound(listOfFiles) should ALWAYS > 0 but
        ' let's check it anyway (just to be safe)
        If UBound(listOfFiles) > 0 Then
        
            'Remove all empty strings from the array (which are a byproduct of the aforementioned buffering)
            For x = UBound(listOfFiles) To 0 Step -1
                If listOfFiles(x) <> "" Then Exit For
            Next
            
            'With all the empty strings removed, all that's left is legitimate file paths
            ReDim Preserve listOfFiles(0 To x) As String
            
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
            For x = 1 To UBound(listOfFiles)
                listOfFiles(x - 1) = imagesPath & listOfFiles(x)
            Next x
            
            ReDim Preserve listOfFiles(0 To UBound(listOfFiles) - 1)
            
            'Save the new directory as the default path for future usage
            WriteToIni "Program Paths", "MainOpen", imagesPath
            
        'If there is only one file in the array (e.g. the user only opened one image), we don't need to do all
        ' that extra processing - just save the new directory to the INI file
        Else
        
            'Save the new directory as the default path for future usage
            tempPathString = listOfFiles(0)
            StripDirectory tempPathString
        
            WriteToIni "Program Paths", "MainOpen", tempPathString
            
        End If
        
        'Also, remember the file filter for future use (in case the user tends to use the same filter repeatedly)
        WriteToIni "File Formats", "LastOpenFilter", CStr(LastOpenFilter)
        
        'All done!
        PhotoDemon_OpenImageDialog = True
        
    'If the user cancels the commondialog box, simply exit out
    Else
        
        If CC.ExtendedError <> 0 Then MsgBox "An error occurred: " & CC.ExtendedError
    
        PhotoDemon_OpenImageDialog = False
    End If
    
    'Release the common dialog object
    Set CC = Nothing

End Function

'Subroutine for saving an image to file.  This function assumes the image already exists on disk and is simply
' being replaced; if the file does not exist on disk, this routine will automatically transfer control to Save As...
' The imageToSave is a reference to an ID in the pdImages() array.  It can be grabbed from the form.Tag value as well.
Public Function MenuSave(ByVal ImageID As Long) As Boolean

    If pdImages(ImageID).LocationOnDisk = "" Then
        'This image hasn't been saved before.  Launch the Save As... dialog
        MenuSave = MenuSaveAs(ImageID)
    Else
        'This image has been saved before.  Overwrite it.
        MenuSave = PhotoDemon_SaveImage(ImageID, pdImages(ImageID).LocationOnDisk, False, pdImages(ImageID).saveFlag0, pdImages(ImageID).saveFlag1)
    End If

End Function

'Subroutine for displaying a commondialog save box, then saving an image to the specified file
Public Function MenuSaveAs(ByVal ImageID As Long) As Boolean

    Dim CC As cCommonDialog
    Set CC = New cCommonDialog
    
    'Get the last "save image" path from the INI file
    Dim tempPathString As String
    tempPathString = GetFromIni("Program Paths", "MainSave")
    
    Dim cdfStr As String
    
    cdfStr = "BMP - Windows Bitmap|*.bmp"
    
    If FreeImageEnabled = True Then cdfStr = cdfStr & "|GIF - Graphics Interchange Format|*.gif"
    
    cdfStr = cdfStr & "|JPG - JPEG - JFIF Compliant|*.jpg|PCX - Zsoft Paintbrush|*.pcx"
    
    If zLibEnabled = True Then cdfStr = cdfStr & "|PDI - PhotoDemon Image (Lossless)|*.pdi"
    
    If FreeImageEnabled = True Then cdfStr = cdfStr & "|PNG - Portable Network Graphic|*.png|PPM - Portable Pixel Map (ASCII)|*.ppm|TGA - Truevision Targa|*.tga|TIFF - Tagged Image File Format|*.tif"
    
    cdfStr = cdfStr & "|All files|*.*"
    
    Dim sFile As String
    sFile = pdImages(ImageID).OriginalFileName
    
    'This next chunk of code checks to see if an image with this filename appears in the download location.
    ' If it does, PhotoDemon will append ascending numbers (of the format "_(#)") to the filename until it
    ' finds a unique name.
    If FileExist(tempPathString & sFile & "." & getExtensionFromFilterIndex(LastSaveFilter)) Then
    
        Dim numToAppend As Long
        numToAppend = 2
        
        Do While FileExist(tempPathString & sFile & " (" & numToAppend & ")" & "." & getExtensionFromFilterIndex(LastSaveFilter))
            numToAppend = numToAppend + 1
        Loop
        
        'If the loop has terminated, a unique filename has been found.  Make that the recommended filename.
        sFile = sFile & " (" & numToAppend & ")"
    
    End If
    
    
    If CC.VBGetSaveFileName(sFile, , True, cdfStr, LastSaveFilter, tempPathString, "Save an image", ".bmp|.gif|.jpg|.pcx|.pdi|.png|.ppm|.tga|.tif|.*", FormMain.hWnd, 0) Then
        
        'Save the new directory as the default path for future usage
        tempPathString = sFile
        StripDirectory tempPathString
        WriteToIni "Program Paths", "MainSave", tempPathString
        
        'Also, remember the file filter for future use (in case the user tends to use the same filter repeatedly)
        WriteToIni "File Formats", "LastSaveFilter", CStr(LastSaveFilter)
        
        pdImages(ImageID).containingForm.Caption = sFile
        SaveFileName = sFile
        
        'Transfer control to the core SaveImage routine, which will handle file extension analysis and actual saving
        MenuSaveAs = PhotoDemon_SaveImage(ImageID, sFile, True)
        
    Else
        MenuSaveAs = False
    End If
    
    'Release the common dialog object
    Set CC = Nothing
    
End Function

'This routine will blindly save the image from the form containing pdImages(imageID) to dstPath.  It is up to
' the calling routine to make sure this is what is wanted. (Note: this routine will erase any existing image
' at dstPath, so BE VERY CAREFUL with what you send here!)
Public Function PhotoDemon_SaveImage(ByVal ImageID As Long, ByVal dstPath As String, Optional ByVal loadRelevantForm As Boolean = False, Optional ByVal optionalSaveParameter0 As Long = -1, Optional ByVal optionalSaveParameter1 As Long = -1) As Boolean

    'Used to determine what kind of file the user is trying to open
    Dim FileExtension As String
    FileExtension = UCase(GetExtension(dstPath))
    
    'Only update the MRU if 1) no form is shown (because the user may cancel it), and 2) no errors occured
    Dim updateMRU As Boolean
    updateMRU = False

    If FileExtension = "JPG" Or FileExtension = "JPEG" Or FileExtension = "JPE" Then
        If loadRelevantForm = True Then
            FormJPEG.Show 1, FormMain
            'If the dialog was canceled, note it
            PhotoDemon_SaveImage = Not saveDialogCanceled
        Else
            'Remember the JPEG quality so we don't have to pester the user for it if they save again
            pdImages(ImageID).saveFlag0 = optionalSaveParameter0
            
            'I implement two separate save functions for JPEG images.  While I greatly appreciate John Korejwa's native
            ' VB JPEG encoder, it's not nearly as robust, or fast, or tested as the FreeImage implementation.  As such,
            ' PhotoDemon will default to FreeImage if it's available; otherwise John's JPEG class will be used.
            If FreeImageEnabled = True Then
                SaveJPEGImageUsingFreeImage ImageID, dstPath, optionalSaveParameter0
            Else
                SaveJPEGImageUsingVB ImageID, dstPath, optionalSaveParameter0
            End If
            updateMRU = True
        End If
    ElseIf FileExtension = "PDI" Then
        If zLibEnabled = True Then
            SavePhotoDemonImage ImageID, dstPath
            updateMRU = True
        Else
        'If zLib doesn't exist...
            MsgBox "The zLib compression library (zlibwapi.dll) was marked as missing or corrupted upon program initialization." & vbCrLf & vbCrLf & "To enable PDI saving, please allow " & PROGRAMNAME & " to download plugin updates by going to the Edit Menu -> Program Preferences, and selecting the 'offer to download core plugins' check box.", vbCritical + vbOKOnly + vbApplicationModal, PROGRAMNAME & " PDI Interface Error"
            Message "PDI saving disabled."
        End If
    ElseIf FileExtension = "GIF" Then
        SaveGIFImage ImageID, dstPath
        updateMRU = True
    ElseIf FileExtension = "PNG" Then
        If optionalSaveParameter0 = -1 Then
            SavePNGImage ImageID, dstPath
        Else
            SavePNGImage ImageID, dstPath, optionalSaveParameter0
        End If
        updateMRU = True
    ElseIf FileExtension = "PPM" Then
        SavePPMImage ImageID, dstPath
        updateMRU = True
    ElseIf FileExtension = "TGA" Then
        SaveTGAImage ImageID, dstPath
        updateMRU = True
    ElseIf FileExtension = "TIF" Then
        SaveTIFImage ImageID, dstPath
        updateMRU = True
    ElseIf FileExtension = "PCX" Then
        If loadRelevantForm = True Then
            FormPCX.Show 1, FormMain
        Else
            'Remember the PCX settings so we don't have to pester the user for it if they save again
            pdImages(ImageID).saveFlag0 = optionalSaveParameter0
            pdImages(ImageID).saveFlag1 = optionalSaveParameter1
            SavePCXImage ImageID, dstPath, optionalSaveParameter0, optionalSaveParameter1
            updateMRU = True
        End If
    Else
        SaveBMP ImageID, dstPath
        updateMRU = True
    End If
    
    'UpdateMRU should only be true if the save was successful
    If updateMRU = True Then
        'Add this file to the MRU list
        MRU_AddNewFile dstPath
    
        'Remember the file's location for future saves
        pdImages(ImageID).LocationOnDisk = dstPath
        
        'Remember the file's filename
        Dim tmpFileName As String
        tmpFileName = dstPath
        StripFilename tmpFileName
        pdImages(ImageID).OriginalFileNameAndExtension = tmpFileName
        StripOffExtension tmpFileName
        pdImages(ImageID).OriginalFileName = tmpFileName
        
        'Mark this file as having been saved
        pdImages(ImageID).UpdateSaveState True
        
        PhotoDemon_SaveImage = True
    
    Else
        'Was a save dialog called?  If it was, use that value to return "success" or not
        If loadRelevantForm <> True Then PhotoDemon_SaveImage = False
    End If

End Function

'Return a string containing the expected file extension of the supplied commondialog filter index
Private Function getExtensionFromFilterIndex(ByVal FilterIndex As Long) As String

    Select Case FilterIndex
        Case 1
            getExtensionFromFilterIndex = "bmp"
        Case 2
            getExtensionFromFilterIndex = "gif"
        Case 3
            getExtensionFromFilterIndex = "jpg"
        Case 4
            getExtensionFromFilterIndex = "pcx"
        Case 5
            getExtensionFromFilterIndex = "pdi"
        Case 6
            getExtensionFromFilterIndex = "png"
        Case 7
            getExtensionFromFilterIndex = "ppm"
        Case 8
            getExtensionFromFilterIndex = "tga"
        Case 9
            getExtensionFromFilterIndex = "tif"
        Case Else
            getExtensionFromFilterIndex = "undefined"
    End Select
    
End Function
