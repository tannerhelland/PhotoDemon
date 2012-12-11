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
    
    'Next, detect the version of Windows we're running on.  PhotoDemon is only concerned with "Vista or later", which lets it
    ' know that certain features are guaranteed to be available.
    isVistaOrLater = getVistaOrLaterStatus
    
    'Initialize a preferences and settings handler
    Set userPreferences = New pdPreferences
    
    'Initialize an image format handler
    Set imageFormats = New pdFormats
    
    'Ask the new preferences handler to generate key program folders.  (If these folders don't exist, the handler will create them)
    LoadMessage "Initializing all program directories..."
    userPreferences.initializePaths
    
    'Now, ask the preferences handler to load all other user settings from the INI file
    LoadMessage "Loading all user settings..."
    userPreferences.loadUserSettings
        
    'Check for plugins (we do this early, because other routines rely on this knowledge)
    ' (Note that this is also the routine that checks GDI+ availability, despite it not really being a "plugin")
    LoadMessage "Loading plugins..."
    LoadPlugins
    
    'Based on the list of available plugins, initialize the program's image format handler
    LoadMessage "Loading import/export libraries..."
    imageFormats.generateInputFormats
    imageFormats.generateOutputFormats
    
    'Set default variables
    LoadMessage "Initializing all user settings..."
    
    'No custom filters have been created yet
    HasCreatedFilter = False
    
    'Mark the Macro recorder as "not recording"
    MacroStatus = MacroSTOP
    
    'Set the default common dialog filters
    LastOpenFilter = userPreferences.GetPreference_Long("File Formats", "LastOpenFilter", 1)
    LastSaveFilter = userPreferences.GetPreference_Long("File Formats", "LastSaveFilter", 3)
    
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
    
    'This list of zoom values is effectively arbitrary.  I've based this list off similar lists (Paint.NET, GIMP) while
    ' including a few extra values for convenience's sake
    Zoom.ZoomCount = 25
    ReDim Zoom.ZoomArray(0 To Zoom.ZoomCount) As Double
    ReDim Zoom.ZoomFactor(0 To Zoom.ZoomCount) As Double
    FormMain.CmbZoom.AddItem "3200%", 0
        Zoom.ZoomArray(0) = 32
        Zoom.ZoomFactor(0) = 32
    FormMain.CmbZoom.AddItem "2400%", 1
        Zoom.ZoomArray(1) = 24
        Zoom.ZoomFactor(1) = 24
    FormMain.CmbZoom.AddItem "1600%", 2
        Zoom.ZoomArray(2) = 16
        Zoom.ZoomFactor(2) = 16
    FormMain.CmbZoom.AddItem "1200%", 3
        Zoom.ZoomArray(3) = 12
        Zoom.ZoomFactor(3) = 12
    FormMain.CmbZoom.AddItem "800%", 4
        Zoom.ZoomArray(4) = 8
        Zoom.ZoomFactor(4) = 8
    FormMain.CmbZoom.AddItem "700%", 5
        Zoom.ZoomArray(5) = 7
        Zoom.ZoomFactor(5) = 7
    FormMain.CmbZoom.AddItem "600%", 6
        Zoom.ZoomArray(6) = 6
        Zoom.ZoomFactor(6) = 6
    FormMain.CmbZoom.AddItem "500%", 7
        Zoom.ZoomArray(7) = 5
        Zoom.ZoomFactor(7) = 5
    FormMain.CmbZoom.AddItem "400%", 8
        Zoom.ZoomArray(8) = 4
        Zoom.ZoomFactor(8) = 4
    FormMain.CmbZoom.AddItem "300%", 9
        Zoom.ZoomArray(9) = 3
        Zoom.ZoomFactor(9) = 3
    FormMain.CmbZoom.AddItem "200%", 10
        Zoom.ZoomArray(10) = 2
        Zoom.ZoomFactor(10) = 2
    FormMain.CmbZoom.AddItem "100%", 11
        Zoom.ZoomArray(11) = 1
        Zoom.ZoomFactor(11) = 1
    FormMain.CmbZoom.AddItem "75%", 12
        Zoom.ZoomArray(12) = 3 / 4
        Zoom.ZoomFactor(12) = 4 / 3
    FormMain.CmbZoom.AddItem "67%", 13
        Zoom.ZoomArray(13) = 2 / 3
        Zoom.ZoomFactor(13) = 3 / 2
    FormMain.CmbZoom.AddItem "50%", 14
        Zoom.ZoomArray(14) = 0.5
        Zoom.ZoomFactor(14) = 2
    FormMain.CmbZoom.AddItem "33%", 15
        Zoom.ZoomArray(15) = 1 / 3
        Zoom.ZoomFactor(15) = 3
    FormMain.CmbZoom.AddItem "25%", 16
        Zoom.ZoomArray(16) = 0.25
        Zoom.ZoomFactor(16) = 4
    FormMain.CmbZoom.AddItem "20%", 17
        Zoom.ZoomArray(17) = 0.2
        Zoom.ZoomFactor(17) = 5
    FormMain.CmbZoom.AddItem "16%", 18
        Zoom.ZoomArray(18) = 0.16
        Zoom.ZoomFactor(18) = 100 / 16
    FormMain.CmbZoom.AddItem "12%", 19
        Zoom.ZoomArray(19) = 0.12
        Zoom.ZoomFactor(19) = 100 / 12
    FormMain.CmbZoom.AddItem "8%", 20
        Zoom.ZoomArray(20) = 0.08
        Zoom.ZoomFactor(20) = 100 / 8
    FormMain.CmbZoom.AddItem "6%", 21
        Zoom.ZoomArray(21) = 0.06
        Zoom.ZoomFactor(21) = 100 / 6
    FormMain.CmbZoom.AddItem "4%", 22
        Zoom.ZoomArray(22) = 0.04
        Zoom.ZoomFactor(22) = 25
    FormMain.CmbZoom.AddItem "3%", 23
        Zoom.ZoomArray(23) = 0.03
        Zoom.ZoomFactor(23) = 100 / 0.03
    FormMain.CmbZoom.AddItem "2%", 24
        Zoom.ZoomArray(24) = 0.02
        Zoom.ZoomFactor(24) = 50
    FormMain.CmbZoom.AddItem "1%", 25
        Zoom.ZoomArray(25) = 0.01
        Zoom.ZoomFactor(25) = 100
    
    'Set the zoom box to display "100%"
    FormMain.CmbZoom.ListIndex = zoomIndex100
        
    'Initialize the selection box next
    LoadMessage "Initializing selection tool..."
    FormMain.cmbSelRender.AddItem "Lightbox", 0
    FormMain.cmbSelRender.AddItem "Highlight (Blue)", 1
    FormMain.cmbSelRender.AddItem "Highlight (Red)", 2
    FormMain.cmbSelRender.ListIndex = 0
    selectionRenderPreference = 0
    
    'Analyze the current monitor arrangement to make sure we handle multimonitor setups properly
    LoadMessage "Analyzing current monitor setup..."
    Set cMonitors = New clsMonitors
    cMonitors.Refresh
    
    'Render various aspects of the UI
    LoadMessage "Initializing user interface..."
        
    'Manually create multi-line tooltips
    FormMain.cmdOpen.ToolTip = "Open one or more images for editing." & vbCrLf & vbCrLf & "(Another way to open images is dragging them from your desktop" & vbCrLf & " or Windows Explorer and dropping them onto PhotoDemon.)"
    If ConfirmClosingUnsaved Then
        FormMain.cmdClose.ToolTip = "Close the current image." & vbCrLf & vbCrLf & "If the current image has not been saved, you will" & vbCrLf & " receive a prompt to save it before it closes."
    Else
        FormMain.cmdClose.ToolTip = "Close the current image." & vbCrLf & vbCrLf & "Because you have turned off save prompts (via Edit -> Preferences)," & vbCrLf & " you WILL NOT receive a prompt to save this image before it closes."
    End If
    FormMain.cmdSave.ToolTip = "Save the current image." & vbCrLf & vbCrLf & "WARNING: this will overwrite the current image file." & vbCrLf & " To save to a different file, use the ""Save As"" button."
    FormMain.cmdSaveAs.ToolTip = "Save the current image to a new file."
    
    'Load the most-recently-used file list (MRU)
    MRU_LoadFromINI
    
    'Create all manual shortcuts (ones VB isn't capable of generating itself)
    LoadMenuShortcuts
            
    'Initialize the drop shadow engine
    Set canvasShadow = New pdShadow
    canvasShadow.initializeSquareShadow PD_CANVASSHADOWSIZE, PD_CANVASSHADOWSTRENGTH, CanvasBackground
    
    'Use the API to give PhotoDemon's main form a 32-bit icon (VB doesn't support that bit-depth)
    LoadMessage "Fixing icon..."
    SetIcon FormMain.hWnd, "AAA", True
        
    'Load and draw the menu icons
    ' (Note: as a bonus, this function also checks the current Windows version and updates the "isVistaOrLater" public variable accordingly)
    LoadMessage "Generating menu icons..."
    LoadMenuIcons
    
    'Initialize the hand cursor we use for all clickable objects
    InitAllCursors
    
    'Look in the MDIWindow module for this code - it enables/disables various control and menus based on
    ' whether or not images have been loaded
    LoadMessage "Enabling user interface..."
    UpdateMDIStatus
        
    'Set up our main progress bar control
    LoadMessage "Initializing progress bar..."
    Set cProgBar = New cProgressBar
    
    With cProgBar
        .DrawObject = FormMain.picProgBar
        .BarColor = RGB(48, 117, 255)
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
    LoadMessage "Captioning main form..."
    FormMain.Caption = App.Title & " v" & App.Major & "." & App.Minor
    
    'Initialize the custom MDI child form icon handler
    LoadMessage "Preparing custom child icon handler..."
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
    LoadMessage "All preparations successful. Loading final interface..."
    
    'Display the main form and hide the splash form
    FormMain.Show
    
    Unload FormSplash
    
    DoEvents
    
End Sub

'Loading an image begins here.  This routine examines a given file's extension and re-routes control based on that.
Public Sub PreLoadImage(ByRef sFile() As String, Optional ByVal ToUpdateMRU As Boolean = True, Optional ByVal imgFormTitle As String = "", Optional ByVal imgName As String = "", Optional ByVal isThisPrimaryImage As Boolean = True, Optional ByRef targetImage As pdImage, Optional ByRef targetLayer As pdLayer, Optional ByVal pageNumber As Long = 0)
        
    'Display a busy cursor
    If Screen.MousePointer <> vbHourglass Then Screen.MousePointer = vbHourglass
            
    'One of the things we'll be doing in this routine is establishing an original color depth for this image.
    ' FreeImage will return this automatically; GDI+ may not.  Use this tracking variable to notify us that
    ' a manual color count needs to be performed.
    Dim mustCountColors As Boolean
    Dim colorCountCheck As Long
            
    'Because this routine accepts an array of images, we have to be prepared for the possibility that more than
    ' one image file is being opened.  This loop will execute until all files are loaded.
    Dim thisImage As Long
    For thisImage = 0 To UBound(sFile)
    
        'Before doing anything else, reset the multipage checker
        imageHasMultiplePages = False
        
        '...and reset the "need to check colors" variable
        mustCountColors = False
    
        'Next, ensure that the image file actually exists
        Message "Verifying that file exists..."
    
        If FileExist(sFile(thisImage)) = False Then
            Message "File not found. Image load canceled."
            MsgBox "Unfortunately, the image '" & sFile(thisImage) & "' could not be found." & vbCrLf & vbCrLf & "If this image was originally located on removable media (DVD, USB drive, etc), please re-insert or re-attach the media and try again.", vbApplicationModal + vbExclamation + vbOKOnly, "File not found"
            GoTo PreloadMoreImages
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
            
        End If
            
        Message "Determining filetype..."
        
        'Initially, set the filetype of the target image to "unknown".  If the load is successful, this value will
        ' be changed to something >= 0. (Note: if FreeImage is used to load the file, this value will be set by the
        ' LoadFreeImageV3 function.)
        targetImage.OriginalFileFormat = -1
        
        Dim fileExtension As String
        fileExtension = UCase(GetExtension(sFile(thisImage)))
                
        Dim loadSuccessful As Boolean
        loadSuccessful = True
        
        'Dependent on the file's extension, load the appropriate image decoding routine
        Select Case fileExtension
        
            'BMP
            Case "BMP"
                'Bitmaps are preferentially loaded by FreeImage (which loads 32bpp bitmaps correctly), then GDI+ (which
                ' loads 32bpp bitmaps but may ignore the alpha channel), then default VB (which fails with 32bpp but will
                ' load other depths).
                If imageFormats.FreeImageEnabled Then
                    loadSuccessful = LoadFreeImageV3(sFile(thisImage), targetLayer, targetImage)
                Else
                    
                    If imageFormats.GDIPlusEnabled Then
                        loadSuccessful = LoadGDIPlusImage(sFile(thisImage), targetLayer)
                    Else
                        loadSuccessful = LoadVBImage(sFile(thisImage), targetLayer)
                    End If
                    
                    targetImage.OriginalFileFormat = FIF_BMP
                    mustCountColors = True
                    
                End If
                
            'GIF
            Case "GIF"
                'GIF is preferentially loaded by FreeImage, then GDI+ if available, then default VB.
                If imageFormats.FreeImageEnabled Then
                    loadSuccessful = LoadFreeImageV3(sFile(thisImage), targetLayer, targetImage, pageNumber)
                Else
                
                    If imageFormats.GDIPlusEnabled Then
                        loadSuccessful = LoadGDIPlusImage(sFile(thisImage), targetLayer)
                    Else
                        loadSuccessful = LoadVBImage(sFile(thisImage), targetLayer)
                    End If
                    
                    targetImage.OriginalFileFormat = FIF_GIF
                    targetImage.OriginalColorDepth = 8
                    
                End If
                
            'ICONS
            Case "ICO"
                'Icons are preferentially loaded by FreeImage, then GDI+ if available, then default VB.
                loadSuccessful = False
                If imageFormats.FreeImageEnabled Then
                    loadSuccessful = LoadFreeImageV3(sFile(thisImage), targetLayer, targetImage)
                End If
                
                'If FreeImage failed (not likely, but not impossible) or is otherwise unavailable, attempt to load
                ' the icon file with GDI+ or VB.
                If loadSuccessful = False Then
                    
                    If imageFormats.GDIPlusEnabled Then
                        loadSuccessful = LoadGDIPlusImage(sFile(thisImage), targetLayer)
                    Else
                        loadSuccessful = LoadVBImage(sFile(thisImage), targetLayer)
                    End If
                    
                    targetImage.OriginalFileFormat = FIF_ICO
                    mustCountColors = True
                    
                End If
                
            'JPEG
            Case "JIF", "JPG", "JPEG", "JPE"
            
                'JPEGs are preferentially loaded by FreeImage, then GDI+ if available, then default VB, unless we are in the
                ' midst of a batch conversion - in that case, use GDI+ first because it is significantly faster as it doesn't
                ' need to make a copy of the image before operating on it.
                
                'Yes, this system is complicated - but it yields quite good results, including for edge cases like CMYK-encoded JPEGs.
                If MacroStatus = MacroBATCH Then
                
                    If imageFormats.GDIPlusEnabled Then
                        loadSuccessful = LoadGDIPlusImage(sFile(thisImage), targetLayer)
                        targetImage.OriginalFileFormat = FIF_JPEG
                        targetImage.OriginalColorDepth = 24
                    ElseIf imageFormats.FreeImageEnabled Then
                        loadSuccessful = LoadFreeImageV3(sFile(thisImage), targetLayer, targetImage)
                    Else
                        loadSuccessful = LoadVBImage(sFile(thisImage), targetLayer)
                        targetImage.OriginalFileFormat = FIF_JPEG
                        targetImage.OriginalColorDepth = 24
                    End If
                    
                Else
                
                    If imageFormats.FreeImageEnabled Then
                        loadSuccessful = LoadFreeImageV3(sFile(thisImage), targetLayer, targetImage)
                    Else
                        If imageFormats.GDIPlusEnabled Then
                            loadSuccessful = LoadGDIPlusImage(sFile(thisImage), targetLayer)
                        Else
                            loadSuccessful = LoadVBImage(sFile(thisImage), targetLayer)
                        End If
                        
                        targetImage.OriginalFileFormat = FIF_JPEG
                        targetImage.OriginalColorDepth = 24
                            
                    End If
                    
                End If
                    
            'Internal PhotoDemon format
            Case "PDI"
            
                'PDI images require zLib, and are only loaded via a custom routine (obviously, since they are PhotoDemon's native format)
                loadSuccessful = LoadPhotoDemonImage(sFile(thisImage), targetLayer)
                targetImage.OriginalFileFormat = 100
                mustCountColors = True
                
            Case "PNG"
            
                'FreeImage has a more robust (and reliable) PNG implementation than GDI+, so use it if available
                loadSuccessful = False
                If imageFormats.FreeImageEnabled Then
                    loadSuccessful = LoadFreeImageV3(sFile(thisImage), targetLayer, targetImage)
                End If
                
                'If FreeImage fails for some reason (such as it being a 1bpp PNG), offload the image to GDI+
                If loadSuccessful = False Then
                    
                    loadSuccessful = LoadGDIPlusImage(sFile(thisImage), targetLayer)
                    targetImage.OriginalFileFormat = FIF_PNG
                    mustCountColors = True
                    
                End If
                
            Case "TIF", "TIFF"
            
                'FreeImage has a more robust (and reliable) TIFF implementation than GDI+, so use it if available
                If imageFormats.FreeImageEnabled Then
                    loadSuccessful = LoadFreeImageV3(sFile(thisImage), targetLayer, targetImage, pageNumber)
                Else
                    loadSuccessful = LoadGDIPlusImage(sFile(thisImage), targetLayer)
                    targetImage.OriginalFileFormat = FIF_TIFF
                    mustCountColors = True
                End If
                
            Case "TMP"
            
                'TMP files are internal files (BMP format) used by PhotoDemon.  GDI+ is preferable, but .LoadPicture works too
                If imageFormats.GDIPlusEnabled Then
                    loadSuccessful = LoadGDIPlusImage(sFile(thisImage), targetLayer)
                Else
                    loadSuccessful = LoadVBImage(sFile(thisImage), targetLayer)
                End If
                
                targetImage.OriginalFileFormat = FIF_BMP
                mustCountColors = True
                
            'Every other file type must be loaded by FreeImage.  Unfortunately, we can't be guaranteed that FreeImage exists.
            Case Else
            
                If imageFormats.FreeImageEnabled = True Then
                    loadSuccessful = LoadFreeImageV3(sFile(thisImage), targetLayer, targetImage)
                Else
                    MsgBox "Unfortunately, the FreeImage plugin (FreeImage.dll) was marked as missing or corrupted upon program initialization." & vbCrLf & vbCrLf & "To enable support for this image format, please allow " & PROGRAMNAME & " to download a fresh copy of FreeImage by going to the Edit -> Program Preferences menu and enabling the option called:" & vbCrLf & vbCrLf & """If core plugins cannot be located, offer to download them""" & vbCrLf & vbCrLf & "Once this is enabled, restart " & PROGRAMNAME & " and it will download this plugin for you.", vbExclamation + vbOKOnly + vbApplicationModal, PROGRAMNAME & " FreeImage Interface Error"
                    Message "Image load canceled."
                    pdImages(CurrentImage).IsActive = False
                    Unload FormMain.ActiveForm
                    GoTo PreloadMoreImages
                End If
        
        End Select
        
        'Double-check to make sure the image was loaded successfully
        If loadSuccessful = False Then
            Message "Image load canceled."
            MsgBox "Unfortunately, PhotoDemon was unable to load the following image:" & vbCrLf & vbCrLf & sFile(thisImage) & vbCrLf & vbCrLf & "Please use another program to save this image in a generic format (such as JPEG or PNG) before loading it into PhotoDemon.  Thanks!", vbExclamation + vbOKOnly + vbApplicationModal, "PhotoDemon Import Failed"
            targetImage.IsActive = False
            If isThisPrimaryImage Then Unload FormMain.ActiveForm
            GoTo PreloadMoreImages
        Else
            Message "Image data loaded successfully."
        End If
        
        'Before continuing, if the image is 32bpp, verify the alpha channel.  If the alpha channel is all 0's or all 255's,
        ' we can conserve on resources by transparently converting it to 24bpp.
        If targetImage.mainLayer.getLayerColorDepth = 32 Then
            
            'Make sure the user hasn't disabled this capability
            If userPreferences.GetPreference_Boolean("General Preferences", "ValidateAlphaChannels", True) Then
            
                Message "Verfiying alpha channel..."
            
                'Verify the alpha channel.  If this function returns FALSE, the alpha channel is unnecessary.
                If targetImage.mainLayer.verifyAlphaChannel = False Then
                
                    Message "Alpha channel deemed unnecessary.  Converting image to 24bpp..."
                
                    'Transparently convert the main layer to 24bpp
                    targetImage.mainLayer.convertTo24bpp
                
                Else
                    Message "Alpha channel verified.  Leaving image in 32bpp mode."
                End If
                
            Else
                Message "Alpha channel validation ignored at user's request."
            End If
        
        End If
        
        'Store important data about the image to the pdImages array
        targetImage.updateSize
        targetImage.OriginalFileSize = FileLen(sFile(thisImage))
        targetImage.CurrentFileFormat = targetImage.OriginalFileFormat
                
        'At this point, we now have loaded image data in 24 or 32bpp format.  For future reference, let's count
        ' the number of colors present in the image (if the user has allowed it).  If the user HASN'T allowed
        ' it, we have no choice but to rely on whatever color depth was returned by FreeImage or GDI+ (or was
        ' inferred by us for this format, e.g. we know that GIFs are 8bpp).
        
        If userPreferences.GetPreference_Boolean("General Preferences", "VerifyInitialColorDepth", True) Or mustCountColors Then
            
            colorCountCheck = getQuickColorCount(targetImage)
        
            'If 256 or less colors were found in the image, mark it as 8bpp.  Otherwise, mark it as 24 or 32bpp.
            targetImage.OriginalColorDepth = getColorDepthFromColorCount(colorCountCheck, targetImage.mainLayer)
            
            Message "Color count successful (" & targetImage.OriginalColorDepth & " BPP)"
            
        End If
                
        'If this is a primary image, it needs to be rendered to the screen
        If isThisPrimaryImage Then
            
            'If the form isn't maximized or minimized then set its dimensions to just slightly bigger than the image size
            Message "Resizing image to fit screen..."
    
            'If the user wants us to resize the image to fit on-screen, do that now
            If AutosizeLargeImages = 0 Then FitImageToViewport True
                    
            'If the window is not maximized or minimized, fit the form around the picture box
            If FormMain.ActiveForm.WindowState = 0 Then FitWindowToImage True
            
        End If
        
        'If a different image name has been specified, we can assume the calling routine is NOT loading a file
        ' from disk (e.g. it's a scan, or Internet download, or screen capture, etc.).  Therefore, set the
        ' file name as requested but leave the .LocationOnDisk blank so that a Save command will trigger
        ' the necessary Save As... dialog.
        Dim tmpFilename As String
        
        If imgName = "" Then
            'The calling routine hasn't specified an image name, so assume this is a normal load situation.
            ' That means pulling the filename from the file itself.
            targetImage.LocationOnDisk = sFile(thisImage)
            
            tmpFilename = sFile(thisImage)
            StripFilename tmpFilename
            targetImage.OriginalFileNameAndExtension = tmpFilename
            StripOffExtension tmpFilename
            targetImage.OriginalFileName = tmpFilename
            
            'Disable the save button, because this file exists on disk
            targetImage.UpdateSaveState True
            
        Else
            'The calling routine has specified a file name.  Assume this is a special case, and force a Save As...
            ' dialog in the future by not specifying a location on disk
            targetImage.LocationOnDisk = ""
            targetImage.OriginalFileNameAndExtension = imgName
            
            tmpFilename = imgName
            StripOffExtension tmpFilename
            targetImage.OriginalFileName = tmpFilename
            
            'Similarly, enable the save button
            targetImage.UpdateSaveState False
            
        End If
            
        'More UI-related updates are necessary if this is a primary image
        If isThisPrimaryImage Then
        
            'Update relevant user interface controls
            DisplaySize targetImage.Width, targetImage.Height
            If imgFormTitle = "" Then
                If userPreferences.GetPreference_Long("General Preferences", "ImageCaptionSize", 0) = 0 Then
                    FormMain.ActiveForm.Caption = getFilename(sFile(thisImage))
                Else
                    FormMain.ActiveForm.Caption = sFile(thisImage)
                End If
            Else
                FormMain.ActiveForm.Caption = imgFormTitle
            End If
            
            'Check the image's color depth, and check/uncheck the matching Image Mode setting
            If targetImage.mainLayer.getLayerColorDepth() = 32 Then tInit tImgMode32bpp, True Else tInit tImgMode32bpp, False
            
            'FixScrolling may have been reset by this point (by the FitImageToViewport sub, among others), so MAKE SURE it's false
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
            
            'Finally, if the image has not been resized to fit on screen, check its viewport to make sure the right and
            ' bottom edges don't fall outside the MDI client area
            'If the user wants us to resize the image to fit on-screen, do that now
            If AutosizeLargeImages = 1 Then FitWindowToViewport
        
            'Finally, add this file to the MRU list (unless specifically told not to)
            If (ToUpdateMRU = True) And (pageNumber = 0) Then MRU_AddNewFile sFile(thisImage), targetImage
        
        End If
        
        'If we made it all the way here, the image loaded successfully.
        targetImage.loadedSuccessfully = True
        
        Message "Image loaded successfully."
        
        'Before continuing on to the next image (if any), see if the just-loaded image was in multipage format.  If it was, the user
        ' may have requested that we load all frames from this image.
        If imageHasMultiplePages Then
        
            Dim pageTracker As Long
            
            Dim tmpStringArray(0) As String
            tmpStringArray(0) = sFile(thisImage)
            
            'Call PreLoadImage again for each individual frame in the multipage file
            For pageTracker = 1 To imagePageCount
                If GetExtension(sFile(thisImage)) = "gif" Then
                    PreLoadImage tmpStringArray, False, targetImage.OriginalFileName & " (frame " & (pageTracker + 1) & ")." & GetExtension(sFile(thisImage)), targetImage.OriginalFileName & " (frame " & (pageTracker + 1) & ")." & GetExtension(sFile(thisImage)), , , , pageTracker
                Else
                    PreLoadImage tmpStringArray, False, targetImage.OriginalFileName & " (page " & (pageTracker + 1) & ")." & GetExtension(sFile(thisImage)), targetImage.OriginalFileName & " (page " & (pageTracker + 1) & ")." & GetExtension(sFile(thisImage)), , , , pageTracker
                End If
            Next pageTracker
        
        End If
        
PreloadMoreImages:

    'If we have more images to process, now's the time to do it!
    Next thisImage
        
    If pageNumber = 0 Then Screen.MousePointer = vbNormal
        
End Sub

'Load any file that hasn't explicitly been sent elsewhere.  FreeImage will automatically determine filetype.
Public Function LoadFreeImageV3(ByVal sFile As String, ByRef dstLayer As pdLayer, ByRef dstImage As pdImage, Optional ByVal pageNumber As Long = 0) As Boolean

    LoadFreeImageV3 = LoadFreeImageV3_Advanced(sFile, dstLayer, dstImage, pageNumber)
    
End Function

'PDI loading.  "PhotoDemon Image" files are basically just bitmap files ran through zLib compression.
Public Function LoadPhotoDemonImage(ByVal PDIPath As String, ByRef dstLayer As pdLayer) As Boolean
    
    'Decompress the current PDI file
    DecompressFile PDIPath
    
    'Load the decompressed bitmap into a temporary StdPicture object
    Dim tmpPicture As StdPicture
    Set tmpPicture = New StdPicture
    Set tmpPicture = LoadPicture(PDIPath)
    
    If tmpPicture.Width = 0 Or tmpPicture.Height = 0 Then
        LoadPhotoDemonImage = False
        Exit Function
    End If
    
    'Copy the image into the current pdImage object
    dstLayer.CreateFromPicture tmpPicture
    
    'Recompress the file back to its original state (I know, it's a terrible way to load these files - but since no one
    ' uses them at present (because there is literally zero advantage to them) I'm not going to optimize it further.)
    CompressFile PDIPath
    
    LoadPhotoDemonImage = True

End Function

'Use GDI+ to load an image.  This does very minimal error checking (which is a no-no with GDI+) but because it's only a
' fallback when FreeImage can't be found, I'm postponing further debugging for now.
'Used for PNG and TIFF files if FreeImage cannot be located.
Public Function LoadGDIPlusImage(ByVal imagePath As String, ByRef dstLayer As pdLayer) As Boolean

    Dim tmpPicture As StdPicture
    Set tmpPicture = New StdPicture
            
    Dim verifyGDISuccess As Boolean
    
    verifyGDISuccess = GDIPlusLoadPicture(imagePath, tmpPicture)
    
    If verifyGDISuccess And (tmpPicture.Width <> 0) And (tmpPicture.Height <> 0) Then
    
        'Copy the image returned by GDI+ into the current pdImage object
        LoadGDIPlusImage = dstLayer.CreateFromPicture(tmpPicture)
        
        'If the load was successful and the image contains an alpha channel, remove the effects of a premultiplied alpha channel
        ' (which is the GDI+ default)
        If LoadGDIPlusImage And dstLayer.getLayerColorDepth = 32 Then dstLayer.fixPremultipliedAlpha
                
    Else
        LoadGDIPlusImage = False
    End If
    
End Function

'BITMAP loading
Public Function LoadVBImage(ByVal imagePath As String, ByRef dstLayer As pdLayer) As Boolean
    
    'Create a temporary StdPicture object that will be used to load the image
    Dim tmpPicture As StdPicture
    Set tmpPicture = New StdPicture
    Set tmpPicture = LoadPicture(imagePath)
    
    If tmpPicture.Width = 0 Or tmpPicture.Height = 0 Then
        LoadVBImage = False
        Exit Function
    End If
    
    'Copy the image into the current pdImage object
    dstLayer.CreateFromPicture tmpPicture
    
    LoadVBImage = True
    
End Function

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

    'Don't allow custom shortcuts in the IDE, as they require subclassing and might crash
    If Not isProgramCompiled Then Exit Sub

    With FormMain.ctlAccelerator
    
        'File menu
        .AddAccelerator vbKeyS, vbCtrlMask Or vbShiftMask, "Save_As"
        .AddAccelerator vbKeyI, vbCtrlMask Or vbShiftMask, "Internet_Import"
        .AddAccelerator vbKeyI, vbCtrlMask Or vbAltMask, "Screen_Capture"
        .AddAccelerator vbKeyI, vbCtrlMask Or vbAltMask Or vbShiftMask, "Import_FRX"
        
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
            
        'Edit menu
        .AddAccelerator vbKeyReturn, vbAltMask, "Preferences"
        .AddAccelerator vbKeyZ, vbCtrlMask Or vbAltMask, "Redo"
        .AddAccelerator vbKeyX, vbCtrlMask Or vbShiftMask, "Empty_Clipboard"
        
        'View menu
        .AddAccelerator vbKey0, 0, "FitOnScreen"
        .AddAccelerator vbKeyAdd, 0, "Zoom_In"
        .AddAccelerator vbKeySubtract, 0, "Zoom_Out"
        .AddAccelerator vbKey5, 0, "Zoom_161"
        .AddAccelerator vbKey4, 0, "Zoom_81"
        .AddAccelerator vbKey3, 0, "Zoom_41"
        .AddAccelerator vbKey2, 0, "Zoom_21"
        .AddAccelerator vbKey1, 0, "Actual_Size"
        .AddAccelerator vbKey2, vbShiftMask, "Zoom_12"
        .AddAccelerator vbKey3, vbShiftMask, "Zoom_14"
        .AddAccelerator vbKey4, vbShiftMask, "Zoom_18"
        .AddAccelerator vbKey5, vbShiftMask, "Zoom_116"
        
        'Image menu
        .AddAccelerator vbKeyL, 0, "Rotate_Left"
        .AddAccelerator vbKeyR, 0, "Rotate_Right"
        .AddAccelerator vbKeyX, vbCtrlMask Or vbShiftMask, "Crop_Selection"
        
        'Color Menu
        .AddAccelerator vbKeyC, vbCtrlMask Or vbShiftMask, "Bright_Contrast"
        
        'Window menu
        .AddAccelerator vbKeyPageUp, 0, "Prev_Image"
        .AddAccelerator vbKeyPageDown, 0, "Next_Image"
        
        'No equivalent menu
        .AddAccelerator vbKeyEscape, 0, "Escape"
        
        .Enabled = True
    End With

    'File menu
    FormMain.MnuSaveAs.Caption = FormMain.MnuSaveAs.Caption & vbTab & "Ctrl+Shift+S"
    FormMain.MnuImportFromInternet.Caption = FormMain.MnuImportFromInternet.Caption & vbTab & "Ctrl+Shift+I"
    FormMain.MnuScreenCapture.Caption = FormMain.MnuScreenCapture.Caption & vbTab & "Ctrl+Alt+I"
    FormMain.MnuImportFrx.Caption = FormMain.MnuImportFrx.Caption & vbTab & "Ctrl+Alt+Shift+I"
    FormMain.MnuImportClipboard.Caption = FormMain.MnuImportClipboard.Caption & vbTab & "Ctrl+V"
    
    'NOTE: Drawing of MRU shortcuts is handled in the MRU module

    'Edit menu
    FormMain.MnuPreferences.Caption = FormMain.MnuPreferences.Caption & vbTab & "Alt+Enter"
    'This Redo shortcut remains, but it is hidden; the Windows convention of Ctrl+Y is displayed instead.
    'FormMain.MnuRedo.Caption = FormMain.MnuRedo.Caption & vbTab & "Ctrl+Alt+Z"
    FormMain.MnuEmptyClipboard.Caption = FormMain.MnuEmptyClipboard.Caption & vbTab & "Ctrl+Shift+X"
    
    'View menu
    FormMain.MnuFitOnScreen.Caption = FormMain.MnuFitOnScreen.Caption & vbTab & "0"
    FormMain.MnuZoomIn.Caption = FormMain.MnuZoomIn.Caption & vbTab & " +"
    FormMain.MnuZoomOut.Caption = FormMain.MnuZoomOut.Caption & vbTab & "-"
    FormMain.MnuActualSize.Caption = FormMain.MnuActualSize.Caption & vbTab & "1"
    FormMain.MnuZoom161.Caption = FormMain.MnuZoom161.Caption & vbTab & "5"
    FormMain.MnuZoom81.Caption = FormMain.MnuZoom81.Caption & vbTab & "4"
    FormMain.MnuZoom41.Caption = FormMain.MnuZoom41.Caption & vbTab & "3"
    FormMain.MnuZoom21.Caption = FormMain.MnuZoom21.Caption & vbTab & "2"
    FormMain.MnuZoom12.Caption = FormMain.MnuZoom12.Caption & vbTab & "Shift+2"
    FormMain.MnuZoom14.Caption = FormMain.MnuZoom14.Caption & vbTab & "Shift+3"
    FormMain.MnuZoom18.Caption = FormMain.MnuZoom18.Caption & vbTab & "Shift+4"
    FormMain.MnuZoom116.Caption = FormMain.MnuZoom116.Caption & vbTab & "Shift+5"
        
    'Image menu
    FormMain.MnuRotateClockwise.Caption = FormMain.MnuRotateClockwise.Caption & vbTab & "R"
    FormMain.MnuRotate270Clockwise.Caption = FormMain.MnuRotate270Clockwise.Caption & vbTab & "L"
    FormMain.MnuCropSelection.Caption = FormMain.MnuCropSelection.Caption & vbTab & "Ctrl+Shift+X"
    
    'Color menu
    FormMain.MnuBrightness.Caption = FormMain.MnuBrightness.Caption & vbTab & "Ctrl+Shift+C"
    
    'Window menu
    FormMain.MnuNextImage.Caption = FormMain.MnuNextImage.Caption & vbTab & "Page Down"
    FormMain.MnuPreviousImage.Caption = FormMain.MnuPreviousImage.Caption & vbTab & "Page Up"
    
    
End Sub

'This subroutine handles the detection of the three core plugins strongly recommended for an optimal PhotoDemon
' experience: zLib, EZTwain32, and FreeImage.  For convenience' sake, it also checks for GDI+ availability.
Public Sub LoadPlugins()
    
    'Plugin files are located in the \Data\Plugins subdirectory
    PluginPath = userPreferences.getDataPath & "Plugins\"
    
    'Make sure the plugin path exists
    If Not DirectoryExist(PluginPath) Then MkDir PluginPath
    
    'Old versions of PhotoDemon kept plugins in a different directory. Check the old location,
    ' and if plugin-related files are found, copy them to the new directory
    On Error Resume Next
    Dim tmpPluginPath As String
    tmpPluginPath = userPreferences.getProgramPath & "Plugins\"
    
    If DirectoryExist(tmpPluginPath) Then
        LoadMessage "Copying plugin files to new \Data\Plugins subdirectory"
        
        Dim pluginName As String
        pluginName = "EZTW32.dll"
        If FileExist(tmpPluginPath & pluginName) Then
            FileCopy tmpPluginPath & pluginName, PluginPath & pluginName
            Kill tmpPluginPath & pluginName
        End If
        
        pluginName = "EZTWAIN_README.TXT"
        If FileExist(tmpPluginPath & pluginName) Then
            FileCopy tmpPluginPath & pluginName, PluginPath & pluginName
            Kill tmpPluginPath & pluginName
        End If
        
        pluginName = "FreeImage.dll"
        If FileExist(tmpPluginPath & pluginName) Then
            FileCopy tmpPluginPath & pluginName, PluginPath & pluginName
            Kill tmpPluginPath & pluginName
        End If
        
        pluginName = "license-fi.txt"
        If FileExist(tmpPluginPath & pluginName) Then
            FileCopy tmpPluginPath & pluginName, PluginPath & pluginName
            Kill tmpPluginPath & pluginName
        End If
        
        pluginName = "license-gplv2.txt"
        If FileExist(tmpPluginPath & pluginName) Then
            FileCopy tmpPluginPath & pluginName, PluginPath & pluginName
            Kill tmpPluginPath & pluginName
        End If
        
        pluginName = "license-gplv3.txt"
        If FileExist(tmpPluginPath & pluginName) Then
            FileCopy tmpPluginPath & pluginName, PluginPath & pluginName
            Kill tmpPluginPath & pluginName
        End If
        
        pluginName = "zlibwapi.dll"
        If FileExist(tmpPluginPath & pluginName) Then
            FileCopy tmpPluginPath & pluginName, PluginPath & pluginName
            Kill tmpPluginPath & pluginName
        End If
        
        'After all files have been removed, kill the old Plugin directory
        RmDir tmpPluginPath
        
    End If
        
    'Check for image scanning
    'First, make sure we have our dll file
    If FileExist(PluginPath & "EZTW32.dll") = False Then
        'If we can't find the DLL, hide the menu options and internally disable scanning
        '(perhaps overkill, but it acts as a safeguard to prevent bad DLL-based crashes)
        ScanEnabled = False
        FormMain.MnuScanImage.Visible = False
        FormMain.MnuSelectScanner.Visible = False
        FormMain.MnuImportSepBar1.Visible = False
    Else
        ScanEnabled = True
        
        'If we do find the DLL, use it to check that TWAIN32 support is available on this machine.
        ' If TWAIN32 support isn't available, hide the scanning options, but leave ScanEnabled as True
        ' so our automatic plugin downloader doesn't mistakenly try to download the DLL again.
        If EnableScanner() = False Then
            FormMain.MnuScanImage.Visible = False
            FormMain.MnuSelectScanner.Visible = False
            FormMain.MnuImportSepBar1.Visible = False
        Else
            FormMain.MnuScanImage.Visible = True
            FormMain.MnuSelectScanner.Visible = True
            FormMain.MnuImportSepBar1.Visible = True
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
        imageFormats.FreeImageEnabled = False
        FormMain.MnuRotateArbitrary.Visible = False
    Else
        imageFormats.FreeImageEnabled = True
    End If
    
    'Finally, check GDI+ availability
    If isGDIPlusAvailable() Then
        imageFormats.GDIPlusEnabled = True
    Else
        imageFormats.GDIPlusEnabled = False
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
    If AutosizeLargeImages = 0 Then
        FitImageToViewport True
    Else
        FitWindowToViewport True
    End If
                
    'If the window is not maximized or minimized, fit the form around the picture box
    If FormMain.ActiveForm.WindowState = 0 Then FitWindowToImage True
        
    'Note the image dimensions and display them on the left-hand pane
    DisplaySize pdImages(CurrentImage).Width, pdImages(CurrentImage).Height
    
    'Update the current caption to match
    FormMain.ActiveForm.Caption = pdImages(CurrentImage).OriginalFileNameAndExtension
        
    'FixScrolling may have been reset by this point (by the FitImageToViewport sub, among others), so MAKE SURE it's false
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
