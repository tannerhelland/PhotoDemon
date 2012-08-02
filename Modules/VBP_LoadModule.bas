Attribute VB_Name = "Loading"
'***************************************************************************
'Program/File Loading Handler
'Copyright ©2000-2012 by Tanner Helland
'Created: 4/15/01
'Last updated: 03/July/12
'Last update: Implemented bad file format error handling in the FreeImage load routine.
'
'Module for handling any and all program loading.  This includes the program itself,
'files, and anything else the program needs to take from the hard drive.
'
'***************************************************************************

Option Explicit

'IT ALL BEGINS HERE (after Sub Main, that is)
Public Sub LoadTheProgram()
    
    'Load the splash screen and display it; that form will determine whether
    'we're running in the IDE or as a standalone EXE.  It will also determine
    'the appropriate program path, and from that the plug-in path.
    FormSplash.Show 0
    DoEvents
    
    'First things first: figure out where this .exe was launched from
    ProgramPath = App.Path
    If Right(ProgramPath, 1) <> "\" Then ProgramPath = ProgramPath & "\"
    
    'Now, before doing anything else, load the INI file and corresponding data (via the INIProcessor module)
    LoadINI
    
    'Check for plug-ins (we do this early, because other routies rely on this knowledge)
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
    
    'Set up the toolbar
    LoadMessage "Initializing user interface..."
    
    'Look in the MDIWindow module for this code - it enables/disables various control and menus based on
    ' whether or not images have been loaded
    UpdateMDIStatus
    
    'Set up our main progress bar control
    Set cProgBar = New cProgressBar
    cProgBar.DrawObject = FormMain.picProgBar
    cProgBar.Min = 0
    cProgBar.Max = 100
    cProgBar.XpStyle = True
    cProgBar.TextAlignX = EVPRGcenter
    cProgBar.TextAlignY = EVPRGcenter
    cProgBar.ShowText = True
    cProgBar.Text = "Please load an image.  (The large 'Open Image' button at the top-left should do the trick!)"
    cProgBar.Draw
    
    'Set up GUI defaults
    FormMain.Caption = App.Title & " v" & App.Major & "." & App.Minor
    
    'Clear the progress bar
    SetProgBarVal 0
    
    'Create all manual shortcuts (ones VB isn't capable of generating itself)
    LoadMenuShortcuts
    
    'Load the most-recently-used file list (MRU)
    MRU_LoadFromINI
    
    LoadMessage "Fixing icon..."
    SetIcon FormMain.HWnd, "AAA", True
    
    'Menu icons are on hold until I can figure out how I want to do it.  I'm not thrilled about implementing
    ' a full-on owner-drawn menu system, but that seems preferable to a mess of third-party dependencies
    ' that do their own subclassing... this will take some time to sort out.  :/
    LoadMessage "Generating menu icons..."
    LoadMenuIcons
    
    'Check the command line to see if the user is attempting to load an image
    LoadMessage "Checking command line..."
    
    If CommandLine <> "" Then
        
        LoadMessage "Loading image(s)..."
        
        'Parse the command line for multiple files, then pass the array over to PreLoadImage
        Dim sFile() As String
        sFile = Split(Mid$(CommandLine, 2, Len(CommandLine) - 2), """ """)
        
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
Public Sub PreLoadImage(ByRef sFile() As String, Optional ByVal ToUpdateMRU As Boolean = True, Optional ByVal imgFormTitle As String = "", Optional ByVal imgName As String = "")
    
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
    
        Message "Image found. Initializing blank form..."

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
        Unload FormHistogram
        
        Message "Determining filetype..."
        
        Dim FileExtension As String
        FileExtension = UCase(GetExtension(sFile(thisImage)))
        
        'Add this file to the MRU list (unless specifically told not to)
        If ToUpdateMRU = True Then MRU_AddNewFile sFile(thisImage)
        
        'Dependent on the file's extension, load the appropriate image decoding routine
        If FileExtension = "PCX" Then
            LoadPCXImage (sFile(thisImage))
        ElseIf FileExtension = "PDI" Then
            LoadPhotoDemonImage (sFile(thisImage))
        ElseIf FileExtension = "PNG" Then
            'FreeImage has a more robust .png implementation than our VB-only solution, so use it if available
            If FreeImageEnabled = True Then
                LoadFreeImageV3 (sFile(thisImage))
            Else
                LoadPNGImage (sFile(thisImage))
            End If
        ElseIf FileExtension = "ICO" Then
            'FreeImage has a more robust .ico implementation than VB's LoadPicture, so use it if available
            If FreeImageEnabled = True Then
                LoadFreeImageV3 sFile(thisImage)
            Else
                LoadBMP sFile(thisImage)
            End If
        ElseIf FileExtension = "JIF" Or FileExtension = "JPG" Or FileExtension = "JPEG" Or FileExtension = "JPE" Then
            'FreeImage has a more robust .jpeg implementation than VB's LoadPicture, so use it if available
            If FreeImageEnabled = True Then
                LoadFreeImageV3 sFile(thisImage)
            Else
                LoadBMP sFile(thisImage)
            End If
        ElseIf FileExtension = "GIF" Or FileExtension = "WMF" Or FileExtension = "EMF" Or FileExtension = "BMP" Or FileExtension = "RLE" Or FileExtension = "TMP" Then
            LoadBMP sFile(thisImage)
        'Every other file type must be loaded by FreeImage.  Unfortunately, we can't be guaranteed that FreeImage exists.
        Else
            If FreeImageEnabled = True Then
                LoadFreeImageV3 sFile(thisImage)
            Else
                MsgBox "Unfortunately. the FreeImage interface plugin (FreeImage.dll) was marked as missing or corrupted upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy this file into the plugin directory and reload " & PROGRAMNAME & ".", vbCritical + vbOKOnly + vbApplicationModal, PROGRAMNAME & " FreeImage Interface Error"
                Exit Sub
            End If
        End If
        
        'If the form isn't maximized or minimized then set its dimensions to just slightly bigger than the image size
        Message "Resizing image to fit screen..."
    
        'If the user wants us to resize the image to fit on-screen, do that now
        If AutosizeLargeImages = 0 Then FitImageToWindow True
                
        'If the window is not maximized or minimized, fit the form around the picture box
        If FormMain.ActiveForm.WindowState = 0 Then FitWindowToImage True
        
        'Store important data about the image to the pdImages array
        pdImages(CurrentImage).PicWidth = GetImageWidth()
        pdImages(CurrentImage).PicHeight = GetImageHeight()
        pdImages(CurrentImage).OriginalFileSize = FileLen(sFile(thisImage))
        
        'If a different image name has been specified, we can assume the calling routine is NOT loading a file
        ' from disk (e.g. it's a scan, or Internet download, or screen capture, etc.).  Therefore, set the
        ' file name as requested but leave the .LocationOnDisk blank so that a Save command will trigger
        ' the Save As... dialog
        Dim tmpFileName As String
        
        If imgName = "" Then
            'The calling routine hasn't specified an image name, so assume this is a normal load situation.
            ' That means pulling the filename from the file itself.
            pdImages(CurrentImage).LocationOnDisk = sFile(thisImage)
            
            tmpFileName = sFile(thisImage)
            StripFilename tmpFileName
            pdImages(CurrentImage).OriginalFileNameAndExtension = tmpFileName
            StripOffExtension tmpFileName
            pdImages(CurrentImage).OriginalFileName = tmpFileName
            
            'Disable the save button, because this file exists on disk
            pdImages(CurrentImage).UpdateSaveState True
            
        Else
            'The calling routine has specified a file name.  Assume this is a special case, and force a Save As...
            ' dialog in the future by not specifying a location on disk
            pdImages(CurrentImage).LocationOnDisk = ""
            pdImages(CurrentImage).OriginalFileNameAndExtension = imgName
            
            tmpFileName = imgName
            StripOffExtension tmpFileName
            pdImages(CurrentImage).OriginalFileName = tmpFileName
            
            'Similarly, enable the save button
            pdImages(CurrentImage).UpdateSaveState False
            
        End If
            
        If imgFormTitle = "" Then FormMain.ActiveForm.Caption = sFile(thisImage) Else FormMain.ActiveForm.Caption = imgFormTitle
        
        'Finally, remember the image dimensions and display them on the left-hand pane
        GetImageData
        DisplaySize FormMain.ActiveForm.BackBuffer.ScaleWidth, FormMain.ActiveForm.BackBuffer.ScaleHeight
        
        'FixScrolling may have been reset by this point (by the FitImageToWindow sub, among others), so MAKE SURE it's false
        FixScrolling = False
        FormMain.CmbZoom.ListIndex = pdImages(CurrentImage).CurrentZoomValue
        
        Message "Image loaded successfully."
    
        'Now that the image is loaded, allow PrepareViewport to set up the scrollbars and buffer
        FixScrolling = True
    
        PrepareViewport FormMain.ActiveForm, "PreLoadImage"
        
        'Note the window state, as it may be important in the future
        pdImages(CurrentImage).WindowState = FormMain.ActiveForm.WindowState
        
        'The form has been hiding off-screen this entire time, and now it's finally time to bring it to the forefront
        If FormMain.ActiveForm.WindowState = 0 Then
            FormMain.ActiveForm.Left = pdImages(CurrentImage).WindowLeft
            FormMain.ActiveForm.Top = pdImages(CurrentImage).WindowTop
        End If
        
    'If we have more images to process, now's the time to do it!
    Next thisImage
    
    Exit Sub
    
End Sub

'Load any file that hasn't explicitly been sent elsewhere.  FreeImage will automatically determine filetype.
Public Sub LoadFreeImageV3(ByVal sFile As String)

    On Error GoTo FreeImageError

    'Make sure we found the plug-in when we loaded the program
    If FreeImageEnabled = False Then
        MsgBox "The FreeImage interface plug-in (FreeImage.dll) was marked as missing or corrupted upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please copy the FreeImage.dll file (downloadable from http://freeimage.sourceforge.net/download.html) into the plug-in directory and reload " & PROGRAMNAME & ".", vbCritical + vbOKOnly + vbApplicationModal, "FreeImage Interface Error"
        Unload FormMain.ActiveForm
        Exit Sub
    End If
    
    'Load the FreeImage library from the plugin directory
    Dim hFreeImgLib As Long
    hFreeImgLib = LoadLibrary(PluginPath & "FreeImage.dll")
    
    'Continue with loading the image...
    FormMain.ActiveForm.BackBuffer.AutoSize = True
    
    'We'll use some customized load strings for certain filetypes, so we need to see what kind of file we're opening
    Dim FileExtension As String
    FileExtension = UCase(GetExtension(sFile))
    
    'When loading JPEGs, opt for accuracy and quality over load speed
    If FileExtension = "JIF" Or FileExtension = "JPG" Or FileExtension = "JPEG" Or FileExtension = "JPE" Then
        FormMain.ActiveForm.BackBuffer.Picture = LoadPictureEx(sFile, FILO_JPEG_ACCURATE)
    Else
        FormMain.ActiveForm.BackBuffer.Picture = LoadPictureEx(sFile)
    End If
    
    FormMain.ActiveForm.BackBuffer.Picture = FormMain.ActiveForm.BackBuffer.Image
    FormMain.ActiveForm.BackBuffer.Refresh
    
    'Release the FreeImage library
    FreeLibrary hFreeImgLib
    
    Exit Sub
    
FreeImageError:

    'Reset the mouse pointer
    FormMain.MousePointer = vbDefault

    'We'll use this string to hold additional error data
    Dim AddInfo As String
    
    'This variable stores the message box type
    Dim mType As VbMsgBoxStyle
    
    'Tracks the user input from the message box
    Dim MsgReturn As VbMsgBoxResult

    'FreeImage throws Error #5 if an invalid image is loaded
    If Err.Number = 5 Then
        AddInfo = "You have attempted to load an invalid picture.  This can happen if a file does not contain image data, or if it contains image data in an unsupported format." & vbCrLf & vbCrLf & "- If you downloaded this image from the Internet, the download may have terminated prematurely.  Please try downloading the image again." & vbCrLf & vbCrLf & "- If this image file came from a digital camera, scanner, or other image editing program, it's possible that " & PROGRAMNAME & " simply doesn't understand this particular file format.  Please save the image in a generic format (such as bitmap or JPEG), then reload it."
        Message "Invalid picture.  Image load cancelled."
        mType = vbCritical + vbOKOnly
    End If
    
    'Create the message box to return the error information
    MsgReturn = MsgBox(PROGRAMNAME & " has experienced an error.  Details on the problem include:" & vbCrLf & vbCrLf & _
    "Error number " & Err.Number & vbCrLf & _
    "Description: " & Err.Description & vbCrLf & vbCrLf & _
    AddInfo & vbCrLf & vbCrLf & _
    "Sorry for the inconvenience," & vbCrLf & _
    "-Tanner Helland" & vbCrLf & PROGRAMNAME & " Developer" & vbCrLf & _
    "www.tannerhelland.com/contact", mType, PROGRAMNAME & " Error Handler: #" & Err.Number)
    
    'If an invalid picture was loaded, unload the active form (since it will just be empty and pictureless)
    If Err.Number = 5 Then Unload FormMain.ActiveForm
    
End Sub

'PDI loading.  "PhotoDemon Image" files are basically just bitmap files ran through zLib compression.
Public Sub LoadPhotoDemonImage(ByVal PDIPath As String)
    
    'Continue with loading the image...
    FormMain.ActiveForm.BackBuffer.AutoSize = True
    DecompressFile PDIPath
    BitmapSize = FileLen(PDIPath)
    FormMain.ActiveForm.BackBuffer.Picture = LoadPicture(PDIPath)
    FormMain.ActiveForm.BackBuffer.Picture = FormMain.ActiveForm.BackBuffer.Image
    FormMain.ActiveForm.BackBuffer.Refresh
    CompressFile PDIPath

End Sub

'PNG loading
Public Sub LoadPNGImage(ByVal PNGPath As String)
    
    Dim pngFile As New LoadPNG

    FormMain.ActiveForm.BackBuffer.AutoSize = False
    FormMain.ActiveForm.BackBuffer.Picture = LoadPicture("")
    FormMain.ActiveForm.ScaleMode = 1
    pngFile.PicBox = FormMain.ActiveForm.BackBuffer
    pngFile.SetToBkgrnd False, 0, 0
    pngFile.BackgroundPicture = FormMain.ActiveForm.BackBuffer
    pngFile.SetAlpha = True
    pngFile.SetTrans = True
    pngFile.OpenPNG PNGPath
    FormMain.ActiveForm.ScaleMode = 3
    FormMain.ActiveForm.BackBuffer.Picture = FormMain.ActiveForm.BackBuffer.Image
    FormMain.ActiveForm.Refresh
    
    'We can't be guaranteed that the PNG loading code will resize the main picture box properly, so do it manually
    FormMain.ActiveForm.BackBuffer.Width = pngFile.Width + 2
    FormMain.ActiveForm.BackBuffer.Height = pngFile.Height + 2
    FormMain.ActiveForm.BackBuffer.AutoSize = True
    DoEvents

End Sub

'PCX loading
Public Sub LoadPCXImage(ByVal PCXPath As String)
    
    Dim pcxFile As New LoadPCX

    pcxFile.Autoscale = True
    pcxFile.LoadPCX PCXPath
    If pcxFile.IsPCX = True Then
        pcxFile.ScaleMode = 1
        FormMain.ActiveForm.BackBuffer.AutoSize = False
        FormMain.ActiveForm.BackBuffer.Picture = LoadPicture("")
        FormMain.ActiveForm.BackBuffer.AutoSize = True
        FormMain.ActiveForm.ScaleMode = 1
        DoEvents
        pcxFile.DrawPCX FormMain.ActiveForm.BackBuffer
        FormMain.ActiveForm.ScaleMode = 3
        FormMain.ActiveForm.BackBuffer.Picture = FormMain.ActiveForm.BackBuffer.Image
        FormMain.ActiveForm.Refresh
    Else
        Message "PCX data corrupted.  Image load aborted."
        Unload FormMain.ActiveForm
    End If
    
End Sub

'BITMAP loading
Public Sub LoadBMP(ByVal BMPFile As String)
    
    'Continue with loading the image...
    FormMain.ActiveForm.BackBuffer.AutoSize = True
    FormMain.ActiveForm.BackBuffer.Picture = LoadPicture(BMPFile)
    FormMain.ActiveForm.BackBuffer.Picture = FormMain.ActiveForm.BackBuffer.Image
    FormMain.ActiveForm.BackBuffer.Refresh

End Sub

'UNDO loading
Public Sub LoadUndo(ByVal UndoFile As String)
    Message "Loading undo/redo data from file..."
       
    'Continue with loading the image...
    FormMain.ActiveForm.BackBuffer.AutoSize = True
    FormMain.ActiveForm.BackBuffer.Picture = LoadPicture("")
    FormMain.ActiveForm.BackBuffer.Picture = LoadPicture(UndoFile)
    FormMain.ActiveForm.BackBuffer.Picture = FormMain.ActiveForm.BackBuffer.Image
    FormMain.ActiveForm.BackBuffer.Refresh
    
    'This will autopopulate things like width, height, etc
    GetImageData
    
    DisplaySize FormMain.ActiveForm.BackBuffer.ScaleWidth, FormMain.ActiveForm.BackBuffer.ScaleHeight
    
    PrepareViewport FormMain.ActiveForm, "Undo"
    
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
    
    'NOTE: Drawing of MRU shortcuts is handled in the MRU module
    
End Sub
'This subroutine handles the detection of the three core plugins strongly recommended for an optimal PhotoDemon
' experience: zLib, EZTwain32, and FreeImage.
Public Sub LoadPlugins()
    
    'Use the path the program was launched from to determine plug-in folder, etc.
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
End Sub


