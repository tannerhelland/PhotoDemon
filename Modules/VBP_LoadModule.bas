Attribute VB_Name = "Loading"
'***************************************************************************
'Program/File Loading Handler
'Copyright ©2000-2012 by Tanner Helland
'Created: 4/15/01
'Last updated: 03/September/12
'Last update: completely rewrote everything against the new layer class.
'
'Module for handling any and all program loading.  This includes the program itself,
'files, and anything else the program needs to take from the hard drive.
'
'***************************************************************************

Option Explicit

'IT ALL BEGINS HERE (after Sub Main, that is)
Public Sub LoadTheProgram()
    
    'Before we can display the splash screen, we need to paint the logo to it.  (This is done for several reasons; it allows
    ' us to keep just one copy of the logo in the project, and it guarantees proper painting regardless of screen DPI.)
    Dim logoWidth As Long, logoHeight As Long
    Dim logoAspectRatio As Double
    
    logoWidth = FormMain.picLogo.ScaleWidth
    logoHeight = FormMain.picLogo.ScaleHeight
    logoAspectRatio = CDbl(logoWidth) / CDbl(logoHeight)
    
    FormSplash.Visible = False
    SetStretchBltMode FormSplash.hDC, STRETCHBLT_HALFTONE
    StretchBlt FormSplash.hDC, 0, 0, FormSplash.ScaleWidth, FormSplash.ScaleWidth / logoAspectRatio, FormMain.picLogo.hDC, 0, 0, logoWidth, logoHeight, vbSrcCopy
    FormSplash.Picture = FormSplash.Image
    
    'With that done, we can now display the splash screen. That form will determine whether we're running in the IDE or as a
    ' standalone EXE.  It will also determine the appropriate program path, and from that the plug-in path.
    FormSplash.Show 0
    DoEvents
    
    'First things first: figure out where this .exe was launched from
    ProgramPath = App.Path
    If Right(ProgramPath, 1) <> "\" Then ProgramPath = ProgramPath & "\"
    
    'Now, before doing anything else, load the INI file and corresponding data (via the INIProcessor module)
    LoadINI
    
    'Check for plug-ins (we do this early, because other routines rely on this knowledge)
    ' (Note that this is also the routine that checks GDI+ availability, despite it not really being a "plugin")
    LoadMessage "Loading plugins..."
    LoadPlugins
    
    'Set default variables
    LoadMessage "Initializing all variables..."
    
    'No custom filters have been created yet
    HasCreatedFilter = False
    
    'Mark the Macro recorder as "not recording"
    MacroStatus = MacroSTOP
    
    'Set the default common dialog filters
    LastOpenFilter = CLng(GetFromIni("File Formats", "LastOpenFilter"))
    LastSaveFilter = CLng(GetFromIni("File Formats", "LastSaveFilter"))
    
    'No images have been loaded yet
    NumOfImagesLoaded = 0
    'Set the default MDI window index to 0
    CurrentImage = 0
    'Set the number of open image windows to 0
    NumOfWindows = 0
    
    'Set the default emboss/engrave color
    EmbossEngraveColor = &HFF8080
    
    'Initialize our current zoom method
    LoadMessage "Initializing zoom processor..."
    
    'Create the first half of the zoom combo box values (zoomed out)
    For x = 16 To 1 Step -1
        If x <> 1 Then FormMain.CmbZoom.AddItem (Format((CDbl(1) / CDbl(x)) * CDbl(100), "##.0") & "%") Else FormMain.CmbZoom.AddItem "100%"
    Next x
    
    'Create the second half of the zoom combo box values (zoomed in)
    For x = 2 To 16 Step 1
        FormMain.CmbZoom.AddItem (CInt((CDbl(x) / CDbl(1)) * CDbl(100)) & "%")
    Next x

    'Set the global "zoom object"'s # of available zoom values
    Zoom.ZoomCount = FormMain.CmbZoom.ListCount - 1
    ReDim Zoom.ZoomArray(0 To Zoom.ZoomCount) As Double
    ReDim Zoom.ZoomFactor(0 To Zoom.ZoomCount) As Byte
    
    'Store zoom coefficients (such as .5, .3333, .25) within a global ZoomArray, whose indices correspond
    ' with the matching combo box values; ZoomFactor stores whole-number values of the zoom ratio, and it
    ' is up to the zoom routine to remember that < index 16 is zoomed out values, while > index 15 is
    ' zoomed in values.
    For x = 0 To 15
        Zoom.ZoomArray(x) = 1 / (16 - x)
        Zoom.ZoomFactor(x) = 16 - x
    Next x
    For x = 2 To 16
        Zoom.ZoomArray(x + 14) = x
        Zoom.ZoomFactor(x + 14) = x
    Next x
    
    'Set the zoom box to display "100%"
    FormMain.CmbZoom.ListIndex = 15
    
    'Get the auto-zoom preference from the INI file
    AutosizeLargeImages = CLng(GetFromIni("General Preferences", "AutosizeLargeImages"))
    
    'Render various aspects of the UI
    LoadMessage "Initializing user interface..."
    
    'Create all manual shortcuts (ones VB isn't capable of generating itself)
    LoadMenuShortcuts
    
    'Load the most-recently-used file list (MRU)
    MRU_LoadFromINI
    
    'Use the API to give PhotoDemon's main form a 32-bit icon (VB doesn't support that bit-depth)
    LoadMessage "Fixing icon..."
    SetIcon FormMain.HWnd, "AAA", True
    
    'Load and draw the menu icons
    LoadMessage "Generating menu icons..."
    LoadMenuIcons
    
    'Initialize the hand cursor we use for all clickable objects
    initHandCursor
    
    'Look in the MDIWindow module for this code - it enables/disables various control and menus based on
    ' whether or not images have been loaded
    UpdateMDIStatus
    
    'Set up our main progress bar control
    Set cProgBar = New cProgressBar
    
    With cProgBar
        .DrawObject = FormMain.picProgBar
        .Min = 0
        .Max = 100
        .XpStyle = True
        .TextAlignX = EVPRGcenter
        .TextAlignY = EVPRGcenter
        .ShowText = True
        .Text = "Please load an image.  (The large 'Open Image' button at the top-left should do the trick!)"
        .Draw
    End With
    
    'Clear the newly built progress bar
    SetProgBarVal 0
    
    'Set up the program's title bar
    FormMain.Caption = App.Title & " v" & App.Major & "." & App.Minor
    
    'Initialize the custom MDI child form icon handler
    initializeIconHandler
    
    'Check the command line to see if the user is attempting to load an image
    LoadMessage "Checking command line..."
    
    If CommandLine <> "" Then
        
        LoadMessage "Loading image(s)..."
        
        'NOTE: Windows will pass multiple filenames in a command line parameter, but its behavior is idiotic.
        ' Specifically, it will place quotation marks around filenames IFF they contain a space, otherwise that filename
        ' will be from its neighboring filenames by a space.  This creates a problem when passing a mixture of filenames
        ' with spaces and filenames without, because Windows will switch between using and not using quotation marks to
        ' delimit the filenames.  Stupid, isn't it?
        
        'At any rate, this means we must perform some specialized parsing of the command line.
        
        'This array will ultimately contain each filename to be loaded (one filename per index)
        Dim sFile() As String
        
        'First, check the command line for quotation marks
        If InStr(CommandLine, Chr(34)) = 0 Then
        
            'If there aren't any, our work is simple - simply split the array using the "space" character as the delimiter
            sFile = Split(CommandLine, Chr(32))
            
        'If there are quotation marks, things get a bit messier.
        Else
        
            Dim inQuotes As Boolean
            inQuotes = False
            Dim tChar As String
            
            'Scan the command line one character at a time
            For x = 1 To Len(CommandLine)
            
                tChar = Mid(CommandLine, x, 1)
                
                'If the current character is a quotation mark, change inQuotes to specify that we are either inside
                ' or outside a SET of quotation marks (note: they will always occur in pairs, per the rules of
                ' how Windows handles command line parameters)
                If tChar = Chr(34) Then inQuotes = Not inQuotes
                
                'If the current character is a space...
                If tChar = Chr(32) Then
                    
                    '...check to see if we are inside quotation marks.  If we are, that means this space is part of a
                    ' filename and NOT a delimiter.  Replace it with an asterisk.
                    If inQuotes = True Then
                        CommandLine = Left(CommandLine, x - 1) & "*" & Right(CommandLine, Len(CommandLine) - x)
                    End If
                    
                End If
            Next x
            
            'At this point, spaces that are parts of filenames have been replaced by asterisks.  That means we can use
            ' Split() to fill our filename array, because the only spaces remaining in the command line are delimiters
            ' between filenames.
            sFile = Split(CommandLine, Chr(32))
            
            'Now that our filenames are successfully inside the sFile() array, go back and replace our asterisk placeholders
            ' with spaces.  Also, remove any quotation marks (since those aren't technically part of the filename).
            For x = 0 To UBound(sFile)
                sFile(x) = Replace$(sFile(x), Chr(42), Chr(32))
                sFile(x) = Replace$(sFile(x), Chr(34), "")
            Next x
        
        End If
        
        'Finally, pass the array of filenames to the loading routine
        PreLoadImage sFile
        
    End If

    'Set a generic message and start unloading and loading forms
    LoadMessage "Preparing software processor..."
    
    'Display the main form and hide the splash form
    FormMain.Show
    
    Unload FormSplash
    
    DoEvents
    
End Sub

'Loading an image begins here.  This routine examines a given file's extension and re-routes control based on that.
Public Sub PreLoadImage(ByRef sFile() As String, Optional ByVal ToUpdateMRU As Boolean = True, Optional ByVal imgFormTitle As String = "", Optional ByVal imgName As String = "", Optional ByVal isThisPrimaryImage As Boolean = True, Optional ByRef targetImage As pdImage, Optional ByRef targetLayer As pdLayer)
    
    Dim thisImage As Long
    
    'Because this routine accepts an array of images, we have to be prepared for the possibility that more than
    ' one image file is being opened.  This loop will execute until all files are loaded.
    For thisImage = 0 To UBound(sFile)
    
        'First, ensure that the image file actually exists
        Message "Verifying that file exists..."
    
        If FileExist(sFile(thisImage)) = False Then
            Message "File not found. Image load canceled."
            MsgBox "Unfortunately, the image '" & sFile(thisImage) & "' could not be found.  If it was originally located on removable media (DVD, USB drive, etc), please re-insert or re-attach the media and try again.", vbApplicationModal + vbCritical + vbOKOnly, "File not found"
            Exit Sub
        End If
    
        'If this is a standard load (e.g. loading an image via File -> Open), prepare a blank form to receive the image.
        If isThisPrimaryImage Then
            
            Message "Image found. Initializing blank form..."

            CreateNewImageForm
        
            Set targetImage = pdImages(CurrentImage)
            Set targetLayer = pdImages(CurrentImage).mainLayer
        
            FixScrolling = False
        
            FormMain.ActiveForm.HScroll.Value = 0
            FormMain.ActiveForm.VScroll.Value = 0
        
            'Prepare the user interface for a new image
            tInit tSaveAs, True
            tInit tCopy, True
            tInit tPaste, True
            tInit tUndo, False
            tInit tRedo, False
            tInit tImageOps, True
            tInit tFilter, True
            tInit tHistogram, True
            
        End If
            
        Message "Determining filetype..."
        
        Dim FileExtension As String
        FileExtension = UCase(GetExtension(sFile(thisImage)))
        
        'Add this file to the MRU list (unless specifically told not to)
        If ToUpdateMRU = True Then MRU_AddNewFile sFile(thisImage)
        
        Dim loadSuccessful As Boolean
        loadSuccessful = True
        
        'Dependent on the file's extension, load the appropriate image decoding routine
        Select Case FileExtension
            Case "BMP"
                'Bitmaps are preferentially loaded by GDI+ if available (which can handle 32bpp), then FreeImage (which is
                ' unpredictable with 32bpp), then default VB (which simply fails with 32bpp but will load other depths).
                If GDIPlusEnabled Then
                    LoadGDIPlusImage sFile(thisImage), targetLayer
                ElseIf FreeImageEnabled Then
                    loadSuccessful = LoadFreeImageV3(sFile(thisImage), targetLayer)
                Else
                    LoadBMP sFile(thisImage), targetLayer
                End If
            Case "GIF"
                'GIF is preferentially loaded by FreeImage, then GDI+ if available, then default VB.
                If FreeImageEnabled Then
                    loadSuccessful = LoadFreeImageV3(sFile(thisImage), targetLayer)
                ElseIf GDIPlusEnabled Then
                    LoadGDIPlusImage sFile(thisImage), targetLayer
                Else
                    LoadBMP sFile(thisImage), targetLayer
                End If
            Case "EMF", "WMF"
                'Metafiles are preferentially loaded by GDI+ if available, then default VB.
                If GDIPlusEnabled Then
                    LoadGDIPlusImage sFile(thisImage), targetLayer
                Else
                    LoadBMP sFile(thisImage), targetLayer
                End If
            Case "ICO"
                'Icons are preferentially loaded by FreeImage, then GDI+ if available, then default VB.
                If FreeImageEnabled Then
                    loadSuccessful = LoadFreeImageV3(sFile(thisImage), targetLayer)
                ElseIf GDIPlusEnabled Then
                    LoadGDIPlusImage sFile(thisImage), targetLayer
                Else
                    LoadBMP sFile(thisImage), targetLayer
                End If
            Case "JIF", "JPG", "JPEG", "JPE"
                'JPEGs are preferentially loaded by FreeImage, then GDI+ if available, then default VB.
                If FreeImageEnabled Then
                    loadSuccessful = LoadFreeImageV3(sFile(thisImage), targetLayer)
                ElseIf GDIPlusEnabled Then
                    LoadGDIPlusImage sFile(thisImage), targetLayer
                Else
                    LoadBMP sFile(thisImage), targetLayer
                End If
            Case "PDI"
                'PDI images require zLib, and are only loaded via a custom routine (obviously, since they are PhotoDemon's native format)
                LoadPhotoDemonImage sFile(thisImage), targetLayer
            Case "PNG"
                'FreeImage has a more robust .png implementation than GDI+, so use it if available
                If FreeImageEnabled = True Then
                    loadSuccessful = LoadFreeImageV3(sFile(thisImage), targetLayer)
                Else
                    LoadGDIPlusImage sFile(thisImage), targetLayer
                End If
            Case "TIF", "TIFF"
                'FreeImage has a more robust (and reliable) TIFF implementation than GDI+, so use it if available
                If FreeImageEnabled = True Then
                    loadSuccessful = LoadFreeImageV3(sFile(thisImage), targetLayer)
                Else
                    LoadGDIPlusImage sFile(thisImage), targetLayer
                End If
            Case "TMP"
                'TMP files are internal files used by PhotoDemon.  VB's internal LoadPicture is fine for these.
                LoadBMP sFile(thisImage), targetLayer
            'Every other file type must be loaded by FreeImage.  Unfortunately, we can't be guaranteed that FreeImage exists.
            Case Else
                If FreeImageEnabled = True Then
                    loadSuccessful = LoadFreeImageV3(sFile(thisImage), targetLayer)
                Else
                    MsgBox "Unfortunately, the FreeImage plugin (FreeImage.dll) was marked as missing or corrupted upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please allow " & PROGRAMNAME & " to download a fresh copy of FreeImage by going to the Edit -> Program Preferences menu and enabling the option called:" & vbCrLf & vbCrLf & """If core plugins cannot be located, offer to download them""" & vbCrLf & vbCrLf & "Once this is enabled, restart " & PROGRAMNAME & " and it will proceed to download this plugin for you.", vbCritical + vbOKOnly + vbApplicationModal, PROGRAMNAME & " FreeImage Interface Error"
                    Message "Image load canceled."
                    pdImages(CurrentImage).isActive = False
                    Unload FormMain.ActiveForm
                    GoTo PreloadMoreImages
                End If
        
        End Select
        
        'Double-check to make sure the image was loaded successfully
        If loadSuccessful = False Then
            MsgBox "Unfortunately, PhotoDemon was unable to load the following image:" & vbCrLf & vbCrLf & sFile(thisImage) & vbCrLf & vbCrLf & "Please use another program to save this image in a generic format (such as JPEG or PNG) before loading it into PhotoDemon.  Thanks!", vbCritical + vbOKOnly + vbApplicationModal, "PhotoDemon Import Failed"
            Message "Image load canceled."
            targetImage.isActive = False
            If isThisPrimaryImage Then Unload FormMain.ActiveForm
            GoTo PreloadMoreImages
        End If
        
        'Store important data about the image to the pdImages array
        targetImage.updateSize
        targetImage.OriginalFileSize = FileLen(sFile(thisImage))
        
        'If this is a primary image, it needs to be rendered to the screen
        If isThisPrimaryImage Then
            
            'If the form isn't maximized or minimized then set its dimensions to just slightly bigger than the image size
            Message "Resizing image to fit screen..."
    
            'If the user wants us to resize the image to fit on-screen, do that now
            If AutosizeLargeImages = 0 Then FitImageToWindow True
                    
            'If the window is not maximized or minimized, fit the form around the picture box
            If FormMain.ActiveForm.WindowState = 0 Then FitWindowToImage True
            
        End If
        
        'If a different image name has been specified, we can assume the calling routine is NOT loading a file
        ' from disk (e.g. it's a scan, or Internet download, or screen capture, etc.).  Therefore, set the
        ' file name as requested but leave the .LocationOnDisk blank so that a Save command will trigger
        ' the necessary Save As... dialog.
        Dim tmpFileName As String
        
        If imgName = "" Then
            'The calling routine hasn't specified an image name, so assume this is a normal load situation.
            ' That means pulling the filename from the file itself.
            targetImage.LocationOnDisk = sFile(thisImage)
            
            tmpFileName = sFile(thisImage)
            StripFilename tmpFileName
            targetImage.OriginalFileNameAndExtension = tmpFileName
            StripOffExtension tmpFileName
            targetImage.OriginalFileName = tmpFileName
            
            'Disable the save button, because this file exists on disk
            targetImage.UpdateSaveState True
            
        Else
            'The calling routine has specified a file name.  Assume this is a special case, and force a Save As...
            ' dialog in the future by not specifying a location on disk
            targetImage.LocationOnDisk = ""
            targetImage.OriginalFileNameAndExtension = imgName
            
            tmpFileName = imgName
            StripOffExtension tmpFileName
            targetImage.OriginalFileName = tmpFileName
            
            'Similarly, enable the save button
            targetImage.UpdateSaveState False
            
        End If
            
        'More UI-related updates are necessary if this is a primary image
        If isThisPrimaryImage Then
        
            'Update relevant user interface controls
            DisplaySize targetImage.Width, targetImage.Height
            If imgFormTitle = "" Then FormMain.ActiveForm.Caption = sFile(thisImage) Else FormMain.ActiveForm.Caption = imgFormTitle
            
            'FixScrolling may have been reset by this point (by the FitImageToWindow sub, among others), so MAKE SURE it's false
            FixScrolling = False
            FormMain.CmbZoom.ListIndex = targetImage.CurrentZoomValue
        
            'Now that the image is loaded, allow PrepareViewport to set up the scrollbars and buffer
            FixScrolling = True
        
            PrepareViewport FormMain.ActiveForm, "PreLoadImage"
            
            'Render an icon-sized version of this image as the MDI child form's icon
            CreateCustomFormIcon FormMain.ActiveForm
            
            'Note the window state, as it may be important in the future
            targetImage.WindowState = FormMain.ActiveForm.WindowState
            
            'The form has been hiding off-screen this entire time, and now it's finally time to bring it to the forefront
            If FormMain.ActiveForm.WindowState = 0 Then
                FormMain.ActiveForm.Left = targetImage.WindowLeft
                FormMain.ActiveForm.Top = targetImage.WindowTop
            End If
        
        End If
        
        'If we made it all the way here, the image loaded successfully.
        targetImage.loadedSuccessfully = True
        
        Message "Image loaded successfully."
        
PreloadMoreImages:

    'If we have more images to process, now's the time to do it!
    Next thisImage
        
End Sub

'Load any file that hasn't explicitly been sent elsewhere.  FreeImage will automatically determine filetype.
Public Function LoadFreeImageV3(ByVal sFile As String, ByRef dstLayer As pdLayer) As Boolean

    LoadFreeImageV3 = LoadFreeImageV3_Advanced(sFile, dstLayer)
    
End Function

'PDI loading.  "PhotoDemon Image" files are basically just bitmap files ran through zLib compression.
Public Sub LoadPhotoDemonImage(ByVal PDIPath As String, ByRef dstLayer As pdLayer)
    
    'Decompress the current PDI file
    DecompressFile PDIPath
    
    'Load the decompressed bitmap into a temporary StdPicture object
    Dim tmpPicture As StdPicture
    Set tmpPicture = New StdPicture
    Set tmpPicture = LoadPicture(PDIPath)
    
    'Copy the image into the current pdImage object
    dstLayer.CreateFromPicture tmpPicture
    
    'Recompress the file back to its original state (I know, it's a terrible way to load these files - but since no one
    ' uses them at present (because there is literally zero advantage to them) I'm not going to optimize it further.)
    CompressFile PDIPath

End Sub

'Use GDI+ to load an image.  This does very minimal error checking (which is a no-no with GDI+) but because it's only a
' fallback when FreeImage can't be found, I'm postponing further debugging for now.
'Used for PNG and TIFF files if FreeImage cannot be located.
Public Sub LoadGDIPlusImage(ByVal imagePath As String, ByRef dstLayer As pdLayer)
        
    'Copy the image returned by GDI+ into the current pdImage object
    dstLayer.CreateFromPicture GDIPlusLoadPicture(imagePath)
    
End Sub

'BITMAP loading
Public Sub LoadBMP(ByVal BMPFile As String, ByRef dstLayer As pdLayer)
    
    'Create a temporary StdPicture object that will be used to load the image
    Dim tmpPicture As StdPicture
    Set tmpPicture = New StdPicture
    Set tmpPicture = LoadPicture(BMPFile)
    
    'Copy the image into the current pdImage object
    dstLayer.CreateFromPicture tmpPicture
    
End Sub

'UNDO loading
Public Sub LoadUndo(ByVal UndoFile As String)
    
    'The layer handles the actual loading of the undo data
    pdImages(CurrentImage).mainLayer.createFromFile UndoFile
    
    'Update the displayed size
    pdImages(CurrentImage).updateSize
    DisplaySize pdImages(CurrentImage).mainLayer.getLayerWidth, pdImages(CurrentImage).mainLayer.getLayerHeight
    
    'Render the image to the screen
    PrepareViewport FormMain.ActiveForm, "LoadUndo"
    
    Message "Undo restored successfully."
    
End Sub

'This routine sets the message on the splash screen (used only when the program is first started)
Public Sub LoadMessage(ByVal sMsg As String)
    FormSplash.lblMessage = sMsg
    FormSplash.lblMessage.Refresh
    DoEvents
End Sub

'Generates all shortcuts that VB can't; many thanks to Steve McMahon for his accelerator class, which helps a great deal
Public Sub LoadMenuShortcuts()

    With FormMain.ctlAccelerator
        .AddAccelerator vbKeyS, vbCtrlMask Or vbShiftMask, "Save_As"
        .AddAccelerator vbKeyI, vbCtrlMask Or vbShiftMask, "Internet_Import"
        .AddAccelerator vbKeyI, vbCtrlMask Or vbAltMask, "Screen_Capture"
        .AddAccelerator vbKeyI, vbCtrlMask Or vbAltMask Or vbShiftMask, "Import_FRX"
        .AddAccelerator vbKeyReturn, vbAltMask, "Preferences"
        .AddAccelerator vbKeyZ, vbCtrlMask Or vbAltMask, "Redo"
        .AddAccelerator vbKeyZ, vbCtrlMask Or vbShiftMask, "Repeat_Last"
        .AddAccelerator vbKeyX, vbCtrlMask Or vbShiftMask, "Empty_Clipboard"
        .AddAccelerator vbKeyAdd, vbCtrlMask, "Zoom_In"
        .AddAccelerator vbKeySubtract, vbCtrlMask, "Zoom_Out"
        .AddAccelerator vbKeyEscape, 0, "Escape"
        
        'Most-recently used files
        .AddAccelerator vbKey0, vbCtrlMask, "MRU_0"
        .AddAccelerator vbKey1, vbCtrlMask, "MRU_1"
        .AddAccelerator vbKey2, vbCtrlMask, "MRU_2"
        .AddAccelerator vbKey3, vbCtrlMask, "MRU_3"
        .AddAccelerator vbKey4, vbCtrlMask, "MRU_4"
        .AddAccelerator vbKey5, vbCtrlMask, "MRU_5"
        .AddAccelerator vbKey6, vbCtrlMask, "MRU_6"
        .AddAccelerator vbKey7, vbCtrlMask, "MRU_7"
        .AddAccelerator vbKey8, vbCtrlMask, "MRU_8"
        .AddAccelerator vbKey9, vbCtrlMask, "MRU_9"
        
        'Next/previous image
        .AddAccelerator vbKeyPageUp, 0, "Prev_Image"
        .AddAccelerator vbKeyPageDown, 0, "Next_Image"
        
        'Brightness/Contrast
        .AddAccelerator vbKeyC, vbCtrlMask Or vbShiftMask, "Bright_Contrast"
        
        .Enabled = True
    End With

    'File menu
    FormMain.MnuSaveAs.Caption = FormMain.MnuSaveAs.Caption & vbTab & "Ctrl+Shift+S"
    FormMain.MnuImportFromInternet.Caption = FormMain.MnuImportFromInternet.Caption & vbTab & "Ctrl+Shift+I"
    FormMain.MnuScreenCapture.Caption = FormMain.MnuScreenCapture.Caption & vbTab & "Ctrl+Alt+I"
    FormMain.MnuImportFrx.Caption = FormMain.MnuImportFrx.Caption & vbTab & "Ctrl+Alt+Shift+I"

    'Edit menu
    FormMain.MnuPreferences.Caption = FormMain.MnuPreferences.Caption & vbTab & "Alt+Enter"
    FormMain.MnuRedo.Caption = FormMain.MnuRedo.Caption & vbTab & "Ctrl+Alt+Z"
    FormMain.MnuRepeatLast.Caption = FormMain.MnuRepeatLast.Caption & vbTab & "Ctrl+Shift+Z"
    FormMain.MnuEmptyClipboard.Caption = FormMain.MnuEmptyClipboard.Caption & vbTab & "Ctrl+Shift+X"
    
    'Color menu
    FormMain.MnuBrightness.Caption = FormMain.MnuBrightness.Caption & vbTab & "Ctrl+Shift+C"
    
    'NOTE: Drawing of MRU shortcuts is handled in the MRU module
    
End Sub

'This subroutine handles the detection of the three core plugins strongly recommended for an optimal PhotoDemon
' experience: zLib, EZTwain32, and FreeImage.  For convenience' sake, it also checks for GDI+ availability.
Public Sub LoadPlugins()
    
    'Use the path the program was launched from to determine plug-in folder
    PluginPath = ProgramPath & "Plugins\"
    
    'Check for image scanning
    'First, make sure we have our dll file
    If FileExist(PluginPath & "EZTW32.dll") = False Then
        'If we can't find the DLL, hide the menu options and internally disable scanning
        '(perhaps overkill, but it acts as a safeguard to prevent bad DLL-based crashes)
        ScanEnabled = False
        FormMain.MnuScanImage.Visible = False
        FormMain.MnuSelectScanner.Visible = False
        FormMain.MnuImportSepBar0.Visible = False
    Else
        ScanEnabled = True
        
        'If we do find the DLL, use it to check that TWAIN32 support is available on this machine.
        ' If TWAIN32 support isn't available, hide the scanning options, but leave ScanEnabled as True
        ' so our automatic plugin downloader doesn't mistakenly try to download the DLL again.
        If EnableScanner() = False Then
            FormMain.MnuScanImage.Visible = False
            FormMain.MnuSelectScanner.Visible = False
            FormMain.MnuImportSepBar0.Visible = False
        Else
            FormMain.MnuScanImage.Visible = True
            FormMain.MnuSelectScanner.Visible = True
            FormMain.MnuImportSepBar0.Visible = True
        End If
    End If
    
    'Check for zLib compression capabilities
    If FileExist(PluginPath & "zlibwapi.dll") = False Then
        zLibEnabled = False
    Else
        zLibEnabled = True
    End If
    
    'Check for FreeImage file interface
    If FileExist(PluginPath & "FreeImage.dll") = False Then
        FreeImageEnabled = False
    Else
        FreeImageEnabled = True
    End If
    
    'Finally, check GDI+ availability
    If isGDIPlusAvailable() Then
        GDIPlusEnabled = True
    Else
        GDIPlusEnabled = False
    End If
    
End Sub

'Make a copy of the current image.  Thanks to PSC user "Achmad Junus" for this suggestion.
Public Sub DuplicateCurrentImage()
    
    Message "Duplicating current image..."
    
    'First, make a note of the currently active form
    Dim imageToBeDuplicated As Long
    imageToBeDuplicated = CurrentImage
    
    CreateNewImageForm
        
    FixScrolling = False
        
    FormMain.ActiveForm.HScroll.Value = 0
    FormMain.ActiveForm.VScroll.Value = 0
        
    'Prepare the user interface for a new image
    tInit tSaveAs, True
    tInit tCopy, True
    tInit tPaste, True
    tInit tUndo, False
    tInit tRedo, False
    tInit tImageOps, True
    tInit tFilter, True
    tInit tHistogram, True
        
    'Copy the picture from the previous form to this new one
    pdImages(CurrentImage).mainLayer.createFromExistingLayer pdImages(imageToBeDuplicated).mainLayer

    'Store important data about the image to the pdImages array
    pdImages(CurrentImage).updateSize
    pdImages(CurrentImage).OriginalFileSize = pdImages(imageToBeDuplicated).OriginalFileSize
    pdImages(CurrentImage).LocationOnDisk = ""
            
    'Get the original file's extension and filename, then append " - Copy" to it
    Dim originalExtension As String
    originalExtension = GetExtension(pdImages(imageToBeDuplicated).OriginalFileNameAndExtension)
            
    Dim newFilename As String
    newFilename = pdImages(imageToBeDuplicated).OriginalFileName & " - Copy"
    pdImages(CurrentImage).OriginalFileName = newFilename
    pdImages(CurrentImage).OriginalFileNameAndExtension = newFilename & "." & originalExtension
            
    'Because this image hasn't been saved to disk, mark its save state as "false"
    pdImages(CurrentImage).UpdateSaveState False
    
    'Fit the window to the newly duplicated image
    Message "Resizing image to fit screen..."
    
    'If the user wants us to resize the image to fit on-screen, do that now
    If AutosizeLargeImages = 0 Then FitImageToWindow True
                
    'If the window is not maximized or minimized, fit the form around the picture box
    If FormMain.ActiveForm.WindowState = 0 Then FitWindowToImage True
        
    'Note the image dimensions and display them on the left-hand pane
    DisplaySize pdImages(CurrentImage).Width, pdImages(CurrentImage).Height
    
    'Update the current caption to match
    FormMain.ActiveForm.Caption = pdImages(CurrentImage).OriginalFileNameAndExtension
        
    'FixScrolling may have been reset by this point (by the FitImageToWindow sub, among others), so MAKE SURE it's false
    FixScrolling = False
    FormMain.CmbZoom.ListIndex = pdImages(CurrentImage).CurrentZoomValue
        
    Message "Image duplication complete."
    
    'Now that the image is loaded, allow PrepareViewport to set up the scrollbars and buffer
    FixScrolling = True
    
    PrepareViewport FormMain.ActiveForm, "DuplicateImage"
        
    'Render an icon-sized version of this image as the MDI child form's icon
    CreateCustomFormIcon FormMain.ActiveForm
        
    'Note the window state, as it may be important in the future
    pdImages(CurrentImage).WindowState = FormMain.ActiveForm.WindowState
        
    'The form has been hiding off-screen this entire time, and now it's finally time to bring it to the forefront
    If FormMain.ActiveForm.WindowState = 0 Then
        FormMain.ActiveForm.Left = pdImages(CurrentImage).WindowLeft
        FormMain.ActiveForm.Top = pdImages(CurrentImage).WindowTop
    End If
    
    'If we made it all the way here, the image was successfully duplicated.
    pdImages(CurrentImage).loadedSuccessfully = True
        
End Sub
