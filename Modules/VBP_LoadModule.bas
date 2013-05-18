Attribute VB_Name = "Loading"
'***************************************************************************
'Program/File Loading Handler
'Copyright ©2001-2013 by Tanner Helland
'Created: 4/15/01
'Last updated: 23/January/13
'Last update: began implementing translation support
'
'Module for handling any and all program loading.  This includes the program itself,
' files, and anything else the program needs to take from the hard drive.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'IT ALL BEGINS HERE (after Sub Main, that is).
' Note that this function is called AFTER FormMain has been loaded.  FormMain is loaded - but not visible - so it can be operated
' on by functions called from this routine.  (It is necessary to load the main form first, since a number of these operations -
' like loading all PNG menu icons from the resource file - operate on the main form.)
Public Sub LoadTheProgram()
    
    '*************************************************************************************************************************************
    ' Prepare the splash screen (but don't display it yet)
    '*************************************************************************************************************************************
    
    'Load FormSplash into memory, but don't make it visible.  Then ask it to prepare itself.
    FormSplash.Visible = False
    FormSplash.prepareSplash
    
    'Check the environment.  If inside the the IDE, the splash needs to be modified slightly.
    CheckLoadingEnvironment
    
    
    
    '*************************************************************************************************************************************
    ' Determine which version of Windows the user is running (as other load functions rely on this)
    '*************************************************************************************************************************************
    
    LoadMessage "Detecting Windows® version..."
    
    'Note that PhotoDemon is only concerned with "Vista or later", which lets it know that certain features are
    ' guaranteed to be available (such as the Segoe UI font, which may not exist on XP installs).
    g_IsVistaOrLater = getVistaOrLaterStatus
    
    
    
    '*************************************************************************************************************************************
    ' Initialize the user preferences (settings) handler
    '*************************************************************************************************************************************
    
    Set g_UserPreferences = New pdPreferences
    
    'Ask the preferences handler to generate key program folders.  (If these folders don't exist, the handler will create them.)
    LoadMessage "Initializing all program directories..."
    
    g_UserPreferences.initializePaths
    
    'Now, ask the preferences handler to load all other user settings from the INI file
    LoadMessage "Loading all user settings..."
    
    g_UserPreferences.loadUserSettings
            
    'While here, also initialize the image format handler (as plugins and other load functions interact with it)
    Set g_ImageFormats = New pdFormats
    
    
    
    '*************************************************************************************************************************************
    ' PhotoDemon works very well with multiple monitors.  Check for such a situation now.
    '*************************************************************************************************************************************
    
    LoadMessage "Analyzing current monitor setup..."
    
    Set g_cMonitors = New clsMonitors
    g_cMonitors.Refresh
    
    
    
    '*************************************************************************************************************************************
    ' Now we have what we need to properly display the splash screen.  Do so now.
    '*************************************************************************************************************************************
        
    'Determine the program's previous on-screen location.  We need that to determine where to display the splash screen.
    Dim wRect As RECT
    wRect.Left = g_UserPreferences.GetPreference_Long("General Preferences", "LastWindowLeft", 1)
    wRect.Top = g_UserPreferences.GetPreference_Long("General Preferences", "LastWindowTop", 1)
    wRect.Right = wRect.Left + g_UserPreferences.GetPreference_Long("General Preferences", "LastWindowWidth", 1)
    wRect.Bottom = wRect.Top + g_UserPreferences.GetPreference_Long("General Preferences", "LastWindowHeight", 1)
    g_cMonitors.CenterFormOnMonitor FormSplash, , wRect.Left, wRect.Right, wRect.Top, wRect.Bottom
            
    'Display the splash screen, centered on whichever monitor the user previously used the program on.
    FormSplash.Show
            
    
    
    '*************************************************************************************************************************************
    ' Initialize the translation (language) engine
    '*************************************************************************************************************************************
    
    'Initialize a new language engine.
    Set g_Language = New pdTranslate
        
    LoadMessage "Scanning for language files..."
    
    'Before doing anything else, check to see what languages are available in the language folder.
    ' (Note that this function will also populate the Languages menu, though it won't place a checkmark next to an entry yet.)
    g_Language.CheckAvailableLanguages
        
    LoadMessage "Determining which language to use..."
        
    'Next, determine which language to use.  (This function will take into account the system language at first-run, so it can
    ' estimate which language to present to the user.)
    g_Language.DetermineLanguage
    
    LoadMessage "Applying selected language..."
    
    'Apply that language to the program.  This involves loading the translation file into memory, which can take a bit of time,
    ' but it only needs to be done once.  From that point forward, any text requests will operate on the in-memory copy of the file.
    g_Language.ApplyLanguage
    
    
    
    '*************************************************************************************************************************************
    ' Check for the presence of plugins (as other functions rely on these to initialize themselves)
    '*************************************************************************************************************************************
    
    LoadMessage "Loading plugins..."
    
    LoadPlugins
    
    '(Note that LoadPlugins also checks GDI+ availability, despite GDI+ not really being a "plugin")
    
    
    
    '*************************************************************************************************************************************
    ' Based on available plugins, determine which image formats PhotoDemon can handle
    '*************************************************************************************************************************************
        
    LoadMessage "Loading import/export libraries..."
        
    g_ImageFormats.generateInputFormats
    g_ImageFormats.generateOutputFormats
    
    
    
    '*************************************************************************************************************************************
    ' Get the viewport engine ready
    '*************************************************************************************************************************************
    
    'Initialize our current zoom method
    LoadMessage "Initializing viewport engine..."
    
    initializeViewportEngine
    
    
        
    '*************************************************************************************************************************************
    ' Set all default tool values
    '*************************************************************************************************************************************
        
    LoadMessage "Initializing image tools..."
    
    Dim i As Long
    
    'INITIALIZE ALL SELECTION TOOLS
    
        'Selection visual styles (currently lightbox or highlight)
        For i = 0 To FormMain.cmbSelRender.Count - 1
            FormMain.cmbSelRender(i).AddItem "Lightbox", 0
            FormMain.cmbSelRender(i).AddItem "Highlight (Blue)", 1
            FormMain.cmbSelRender(i).AddItem "Highlight (Red)", 2
            FormMain.cmbSelRender(i).ListIndex = 0
            g_selectionRenderPreference = 0
        Next i
        
        'Selection types (currently interior, exterior, border)
        FormMain.cmbSelType(0).AddItem "Interior", 0
        FormMain.cmbSelType(0).AddItem "Exterior", 1
        FormMain.cmbSelType(0).AddItem "Border", 2
        FormMain.cmbSelType(0).ListIndex = 0
        
        'Selection smoothing (currently none, antialiased, fully feathered)
        FormMain.cmbSelSmoothing(0).AddItem "None", 0
        FormMain.cmbSelSmoothing(0).AddItem "Antialiased", 1
        FormMain.cmbSelSmoothing(0).AddItem "Feathered", 2
        FormMain.cmbSelSmoothing(0).ListIndex = 1
    
    g_UserPreferences.loadToolSettings
            
        
        
    '*************************************************************************************************************************************
    ' PhotoDemon draws its own drop shadows.  Prepare the engine that handles such shadows.
    '*************************************************************************************************************************************
    
    LoadMessage "Initializing drop shadow renderer..."
            
    'Initialize the drop shadow engine
    Set g_CanvasShadow = New pdShadow
    g_CanvasShadow.initializeSquareShadow PD_CANVASSHADOWSIZE, PD_CANVASSHADOWSTRENGTH, g_CanvasBackground
    
    
    
    '*************************************************************************************************************************************
    ' PhotoDemon's complex interface requires a lot of things to be generated at run-time.
    '*************************************************************************************************************************************
    
    LoadMessage "Initializing user interface..."
                
    'Display or hide the main form's tool panes according to the saved setting in the INI file
    If g_UserPreferences.GetPreference_Boolean("General Preferences", "HideLeftPanel", False) Then
        ChangeLeftPane VISIBILITY_FORCEHIDE
    Else
        ChangeLeftPane VISIBILITY_FORCEDISPLAY
    End If
    
    If g_UserPreferences.GetPreference_Boolean("General Preferences", "HideRightPanel", False) Then
        ChangeRightPane VISIBILITY_FORCEHIDE
    Else
        ChangeRightPane VISIBILITY_FORCEDISPLAY
    End If
                
    'Manually create multi-line tooltips for some command buttons
    FormMain.cmdOpen.ToolTip = g_Language.TranslateMessage("Open one or more images for editing." & vbCrLf & vbCrLf & "(Another way to open images is dragging them from your desktop" & vbCrLf & " or Windows Explorer and dropping them onto PhotoDemon.)")
    If g_ConfirmClosingUnsaved Then
        FormMain.cmdClose.ToolTip = g_Language.TranslateMessage("Close the current image." & vbCrLf & vbCrLf & "If the current image has not been saved, you will" & vbCrLf & " receive a prompt to save it before it closes.")
    Else
        FormMain.cmdClose.ToolTip = g_Language.TranslateMessage("Close the current image." & vbCrLf & vbCrLf & "Because you have turned off save prompts (via Edit -> Preferences)," & vbCrLf & " you WILL NOT receive a prompt to save this image before it closes.")
    End If
    FormMain.cmdSave.ToolTip = g_Language.TranslateMessage("Save the current image." & vbCrLf & vbCrLf & "WARNING: this will overwrite the current image file." & vbCrLf & " To save to a different file, use the ""Save As"" button.")
    FormMain.cmdSaveAs.ToolTip = g_Language.TranslateMessage("Save the current image to a new file.")
                        
    'Use the API to give PhotoDemon's main form a 32-bit icon (VB is too old to support 32bpp icons)
    SetIcon FormMain.hWnd, "AAA", True
    
    'Initialize all system cursors we rely on (hand, busy, resizing, etc)
    InitAllCursors
    
    'Set up the program's title bar
    FormMain.Caption = App.Title & " v" & App.Major & "." & App.Minor
    
    'PhotoDemon renders many of its own icons dynamically.  Initialize that engine now.
    initializeIconHandler
    
    'Before displaying the main window, see if the user wants to restore last-used window location.
    If g_UserPreferences.GetPreference_Boolean("General Preferences", "RememberWindowLocation", True) Then restoreMainWindowLocation
    
    'Finish applying visual styles
    makeFormPretty FormMain
    
    'If Segoe UI is in use, the zoom buttons need to be adjusted to match the combo box
    If g_UseFancyFonts Then
        FormMain.cmdZoomIn.Height = FormMain.cmdZoomIn.Height + 1
        FormMain.cmdZoomOut.Height = FormMain.cmdZoomOut.Height + 1
    End If
    
    'Allow drag-and-drop operations
    g_AllowDragAndDrop = True
    
    
    '*************************************************************************************************************************************
    ' The program's menus support many features that VB can't do natively (like icons and custom shortcuts).  Load such things now.
    '*************************************************************************************************************************************
    
    LoadMessage "Preparing program menus..."
    
    'If inside the IDE, disable the "Effects" -> "Test" menu
    If g_IsProgramCompiled Then FormMain.MnuTest.Visible = False Else FormMain.MnuTest.Visible = True
    
    'Load the most-recently-used file list (MRU)
    MRU_LoadFromINI
    
    'Create all manual shortcuts (ones VB isn't capable of generating itself)
    LoadMenuShortcuts
            
    'Load and draw all menu icons
    LoadMenuIcons
    
    'Look in the MDIWindow module for this code - it enables/disables additional menus based on whether or not images have been loaded.
    ' At this point, it mostly disables all image-related menu items (as no images have been loaded yet)
    UpdateMDIStatus
    
    
    
    '*************************************************************************************************************************************
    ' To avoid relying on OCXs, we use a custom progress bar control.  Initialize it now.
    '*************************************************************************************************************************************
    
    LoadMessage "Initializing progress bar..."
    
    Set g_ProgBar = New cProgressBar
    
    With g_ProgBar
        .DrawObject = FormMain.picProgBar
        .BarColor = RGB(48, 117, 255)
        .Min = 0
        .Max = 100
        .xpStyle = True
        .TextAlignX = EVPRGcenter
        .TextAlignY = EVPRGcenter
        .ShowText = True
        .Text = g_Language.TranslateMessage("Please load an image.  (The large 'Open Image' button at the top-left should do the trick!)")
        .Draw
    End With
    
    'Clear the newly built progress bar
    SetProgBarVal 0
    
    
    
    '*************************************************************************************************************************************
    ' Finally, before loading the final interface, analyze the command line and load any image files (if present).
    '*************************************************************************************************************************************
    
    LoadMessage "Checking command line..."
    
    If g_CommandLine <> "" Then
        LoadMessage "Loading images..."
        FormSplash.Visible = False
        LoadImagesFromCommandLine
    Else
        LoadMessage "All systems go!  Launching main window..."
    End If
    
    
    
    '*************************************************************************************************************************************
    ' Unload the splash screen and present the main form
    '*************************************************************************************************************************************
        
    FormMain.Show
    
    Unload FormSplash
        
End Sub

'If files are present in the command line, this sub will load them
Private Sub LoadImagesFromCommandLine()

    LoadMessage "Loading image(s)..."
        
    'NOTE: Windows can pass the program multiple filenames via the command line, but it does so in a confusing and overly complex way.
    ' Specifically, quotation marks are placed around filenames IFF they contain a space; otherwise, file names are separated from
    ' neighboring filenames by a space.  This creates a problem when passing a mixture of filenames with spaces and filenames without,
    ' because Windows will switch between using and not using quotation marks to delimit the filenames.  Thus, we must perform complex,
    ' specialized parsing of the command line.
        
    'This array will ultimately contain each filename to be loaded (one filename per index)
    Dim sFile() As String
        
    'First, check the command line for quotation marks
    If InStr(g_CommandLine, Chr(34)) = 0 Then
        
        'If there aren't any, our work is simple - simply split the array using the "space" character as the delimiter
        sFile = Split(g_CommandLine, Chr(32))
        
    'If there are quotation marks, things get a lot messier.
    Else
        
        Dim inQuotes As Boolean
        inQuotes = False
        
        Dim tChar As String
        
        'Scan the command line one character at a time
        Dim x As Long
        For x = 1 To Len(g_CommandLine)
            
            tChar = Mid(g_CommandLine, x, 1)
                
            'If the current character is a quotation mark, change inQuotes to specify that we are either inside
            ' or outside a SET of quotation marks (note: they will always occur in pairs, per the rules of
            ' how Windows handles command line parameters)
            If tChar = Chr(34) Then inQuotes = Not inQuotes
                
            'If the current character is a space...
            If tChar = Chr(32) Then
                    
                '...check to see if we are inside quotation marks.  If we are, that means this space is part of a
                ' filename and NOT a delimiter.  Replace it with an asterisk.
                If inQuotes = True Then g_CommandLine = Left(g_CommandLine, x - 1) & "*" & Right(g_CommandLine, Len(g_CommandLine) - x)
                    
            End If
            
        Next x
            
        'At this point, spaces that are parts of filenames have been replaced by asterisks.  That means we can use
        ' Split() to fill our filename array, because the only spaces remaining in the command line are delimiters
        ' between filenames.
        sFile = Split(g_CommandLine, Chr(32))
            
        'Now that our filenames are successfully inside the sFile() array, go back and replace our asterisk placeholders
        ' with spaces.  Also, remove any quotation marks (since those aren't technically part of the filename).
        For x = 0 To UBound(sFile)
            sFile(x) = Replace$(sFile(x), Chr(42), Chr(32))
            sFile(x) = Replace$(sFile(x), Chr(34), "")
        Next x
        
    End If
        
    'Finally, pass the array of filenames to the image loading routine
    PreLoadImage sFile

End Sub

'Loading an image begins here.  This routine examines a given file's extension and re-routes control based on that.
Public Sub PreLoadImage(ByRef sFile() As String, Optional ByVal ToUpdateMRU As Boolean = True, Optional ByVal imgFormTitle As String = "", Optional ByVal imgName As String = "", Optional ByVal isThisPrimaryImage As Boolean = True, Optional ByRef targetImage As pdImage, Optional ByRef targetLayer As pdLayer, Optional ByVal pageNumber As Long = 0)
        
    '*************************************************************************************************************************************
    ' Prepare all variables related to image loading
    '*************************************************************************************************************************************
        
    'Display a busy cursor
    If Screen.MousePointer <> vbHourglass Then Screen.MousePointer = vbHourglass
            
    'One of the things we'll be doing in this routine is establishing an original color depth for this image.
    ' FreeImage will return this automatically; GDI+ may not.  Use this tracking variable to notify us that
    ' a manual color count needs to be performed.
    Dim mustCountColors As Boolean
    Dim colorCountCheck As Long
            
    'To improve load time, declare a variety of other variables outside the image load loop
    Dim FileExtension As String
    Dim loadSuccessful As Boolean
    
    Dim loadedByOtherMeans As Boolean
    loadedByOtherMeans = False
            
    'If multiple files are being loaded, we want to suppress all warnings and errors until the very end.
    Dim multipleFilesLoading As Boolean
    If UBound(sFile) > 0 Then multipleFilesLoading = True Else multipleFilesLoading = False
    
    Dim missingFiles As String
    missingFiles = ""
    
    Dim brokenFiles As String
    brokenFiles = ""
            
            
            
    '*************************************************************************************************************************************
    ' Loop through each entry in the sFile() array, loading images as we go
    '*************************************************************************************************************************************
            
    'Because this routine accepts an array of images, we have to be prepared for the possibility that more than
    ' one image file is being opened.  This loop will execute until all files are loaded.
    Dim thisImage As Long
    
    For thisImage = 0 To UBound(sFile)
    
    
    
        '*************************************************************************************************************************************
        ' Reset all variables used on a per-image level
        '*************************************************************************************************************************************
    
        'Before doing anything else, reset the multipage checker
        imageHasMultiplePages = False
        
        '...and reset the "need to check colors" variable
        mustCountColors = False
    
    
    
        '*************************************************************************************************************************************
        ' Before attempting to load this image, make sure it exists
        '*************************************************************************************************************************************
    
        If isThisPrimaryImage Then Message "Verifying that file exists..."
    
        If isThisPrimaryImage And (Not FileExist(sFile(thisImage))) Then
            Message "File not found (%1). Image load canceled.", sFile(thisImage)
            
            'If multiple files are being loaded, suppress any errors until the end
            If multipleFilesLoading Then
                missingFiles = missingFiles & getFilename(sFile(thisImage)) & vbCrLf
            Else
                pdMsgBox "Unfortunately, the image '%1' could not be found." & vbCrLf & vbCrLf & "If this image was originally located on removable media (DVD, USB drive, etc), please re-insert or re-attach the media and try again.", vbApplicationModal + vbExclamation + vbOKOnly, "File not found", sFile(thisImage)
            End If
            
            GoTo PreloadMoreImages
        End If
        
        
        
        '*************************************************************************************************************************************
        ' If the image being loaded is a primary image (e.g. one opened normally), prepare a blank form to receive it
        '*************************************************************************************************************************************
        
        'If this is a standard load (e.g. loading an image via File -> Open), prepare a blank form to receive the image.
        If isThisPrimaryImage Then
            
            Message "Image found. Initializing blank form..."

            CreateNewImageForm
        
            Set targetImage = pdImages(CurrentImage)
            Set targetLayer = pdImages(CurrentImage).mainLayer
        
            g_FixScrolling = False
        
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
            
            
            
        '*************************************************************************************************************************************
        ' Based on what type of image this is, call the most appropriate load function for it (FreeImage, GDI+, or VB's LoadPicture)
        '*************************************************************************************************************************************
            
        If isThisPrimaryImage Then Message "Determining filetype..."
        
        'Initially, set the filetype of the target image to "unknown".  If the load is successful, this value will
        ' be changed to something >= 0. (Note: if FreeImage is used to load the file, this value will be set by the
        ' LoadFreeImageV3 function.)
        If Not (targetImage Is Nothing) Then targetImage.OriginalFileFormat = -1
        
        'Strip the extension from the file
        FileExtension = UCase(GetExtension(sFile(thisImage)))
        
        loadSuccessful = False
        loadedByOtherMeans = False
            
        'Depending on the file's extension, load the image using the most appropriate image decoding routine
        Select Case FileExtension
        
            'PhotoDemon's custom file format must be handled specially (as obviously, FreeImage and GDI+ won't handle it)
            Case "PDI"
            
                'PDI images require zLib, and are only loaded via a custom routine (obviously, since they are PhotoDemon's native format)
                loadSuccessful = LoadPhotoDemonImage(sFile(thisImage), targetLayer)
                
                targetImage.OriginalFileFormat = 100
                mustCountColors = True
        
            'TMP files are internal files (BMP format) used by PhotoDemon.  GDI+ is preferable, but .LoadPicture works too.
            Case "TMP"
            
                If g_ImageFormats.GDIPlusEnabled Then loadSuccessful = LoadGDIPlusImage(sFile(thisImage), targetLayer)
                
                If (Not g_ImageFormats.GDIPlusEnabled) Or (Not loadSuccessful) Then loadSuccessful = LoadVBImage(sFile(thisImage), targetLayer)
                
                targetImage.OriginalFileFormat = FIF_BMP
                mustCountColors = True
        
            'All other formats follow a prescribed behavior - try to load via FreeImage (if available), then GDI+, then finally
            ' VB's internal LoadPicture function.
            Case Else
                                
                If g_ImageFormats.FreeImageEnabled Then loadSuccessful = LoadFreeImageV3(sFile(thisImage), targetLayer, targetImage, pageNumber, isThisPrimaryImage)
                
                If loadSuccessful Then loadedByOtherMeans = False
                
                'If FreeImage fails for some reason, offload the image to GDI+ - UNLESS the image is a WMF or EMF, which can cause
                ' GDI+ to experience a silent fail, thus bringing down the entire program.
                If (Not loadSuccessful) And g_ImageFormats.GDIPlusEnabled And ((FileExtension <> "EMF") And (FileExtension <> "WMF")) Then
                    
                    If isThisPrimaryImage Then Message "FreeImage refused to load image.  Dropping back to GDI+ and trying again..."
                    loadSuccessful = LoadGDIPlusImage(sFile(thisImage), targetLayer)
                    
                    'If GDI+ loaded the image successfully, note that we have to count available colors ourselves
                    If loadSuccessful Then
                        loadedByOtherMeans = True
                        mustCountColors = True
                    End If
                        
                End If
                
                'If both FreeImage and GDI+ failed, give the image one last try with VB's LoadPicture
                If (Not loadSuccessful) Then
                    
                    If isThisPrimaryImage Then Message "GDI+ refused to load image.  Dropping back to internal routines and trying again..."
                    loadSuccessful = LoadVBImage(sFile(thisImage), targetLayer)
                
                    'If VB managed to load the image successfully, note that we have to count available colors ourselves
                    If loadSuccessful Then
                        loadedByOtherMeans = True
                        mustCountColors = True
                    End If
                
                End If
                    
        End Select
        
        
        
        '*************************************************************************************************************************************
        ' Make sure the image data was loaded successfully
        '*************************************************************************************************************************************
        
        'Double-check to make sure the image was loaded successfully
        If ((Not loadSuccessful) Or (targetImage.mainLayer.getLayerWidth = 0) Or (targetImage.mainLayer.getLayerHeight = 0)) And isThisPrimaryImage Then
            Message "Failed to load %1", sFile(thisImage)
            
            'If multiple files are being loaded, suppress any errors until the end
            If multipleFilesLoading Then
                brokenFiles = brokenFiles & getFilename(sFile(thisImage)) & vbCrLf
            Else
                If MacroStatus <> MacroBATCH Then pdMsgBox "Unfortunately, PhotoDemon was unable to load the following image:" & vbCrLf & vbCrLf & "%1" & vbCrLf & vbCrLf & "Please use another program to save this image in a generic format (such as JPEG or PNG) before loading it into PhotoDemon.  Thanks!", vbExclamation + vbOKOnly + vbApplicationModal, "Image Import Failed", sFile(thisImage)
            End If
            
            targetImage.deactivateImage
            If isThisPrimaryImage Then Unload FormMain.ActiveForm
            GoTo PreloadMoreImages
            
        Else
            If isThisPrimaryImage Then Message "Image data loaded successfully."
        End If
        
        
        
        '*************************************************************************************************************************************
        ' If GDI+ or LoadPicture was used to grab the image data, populate some related fields manually (filetype, color depth, etc)
        '*************************************************************************************************************************************
        
        If loadedByOtherMeans Then
        
            Select Case FileExtension
                
                Case "GIF"
                    targetImage.OriginalFileFormat = FIF_GIF
                    targetImage.OriginalColorDepth = 8
                    
                Case "ICO"
                    targetImage.OriginalFileFormat = FIF_ICO
                
                Case "JIF", "JPG", "JPEG", "JPE"
                    targetImage.OriginalFileFormat = FIF_JPEG
                    targetImage.OriginalColorDepth = 24
                    
                Case "PNG"
                    targetImage.OriginalFileFormat = FIF_PNG
                
                Case "TIF", "TIFF"
                    targetImage.OriginalFileFormat = FIF_TIFF
                    
                'Treat anything else as a BMP file
                Case Else
                    targetImage.OriginalFileFormat = FIF_BMP
                    
            End Select
        
        End If
        
        
        
        '*************************************************************************************************************************************
        ' If the image has an alpha channel, verify it.  If it contains all 0 or all 255, convert it to 24bpp to conserve resources.
        '*************************************************************************************************************************************
        
        If targetImage.mainLayer.getLayerColorDepth = 32 Then
            
            'Make sure the user hasn't disabled this capability
            If g_UserPreferences.GetPreference_Boolean("General Preferences", "ValidateAlphaChannels", True) Then
            
                If isThisPrimaryImage Then Message "Verfiying alpha channel..."
            
                'Verify the alpha channel.  If this function returns FALSE, the alpha channel is unnecessary.
                If targetImage.mainLayer.verifyAlphaChannel = False Then
                
                    If isThisPrimaryImage Then Message "Alpha channel deemed unnecessary.  Converting image to 24bpp..."
                
                    'Transparently convert the main layer to 24bpp
                    targetImage.mainLayer.convertTo24bpp
                
                Else
                    If isThisPrimaryImage Then Message "Alpha channel verified.  Leaving image in 32bpp mode."
                End If
                
            Else
                If isThisPrimaryImage Then Message "Alpha channel validation ignored at user's request."
            End If
        
        End If
        
        
        
        '*************************************************************************************************************************************
        ' Store some universally important information to the target image object
        '*************************************************************************************************************************************
        
        targetImage.updateSize
        If FileExist(sFile(thisImage)) Then targetImage.OriginalFileSize = fileLen(sFile(thisImage))
        targetImage.CurrentFileFormat = targetImage.OriginalFileFormat
                
                
                
        '*************************************************************************************************************************************
        ' If requested by the user, manually count the number of unique colors in the image (to determine absolutely accurate color depth)
        '*************************************************************************************************************************************
                
        'At this point, we now have loaded image data in 24 or 32bpp format.  For future reference, let's count
        ' the number of colors present in the image (if the user has allowed it).  If the user HASN'T allowed
        ' it, we have no choice but to rely on whatever color depth was returned by FreeImage or GDI+ (or was
        ' inferred by us for this format, e.g. we know that GIFs are 8bpp).
        
        If isThisPrimaryImage And (g_UserPreferences.GetPreference_Boolean("General Preferences", "VerifyInitialColorDepth", True) Or mustCountColors) Then
            
            colorCountCheck = getQuickColorCount(targetImage, CurrentImage)
        
            'If 256 or less colors were found in the image, mark it as 8bpp.  Otherwise, mark it as 24 or 32bpp.
            targetImage.OriginalColorDepth = getColorDepthFromColorCount(colorCountCheck, targetImage.mainLayer)
            
            If g_IsImageGray Then
                Message "Color count successful (%1 BPP, grayscale)", targetImage.OriginalColorDepth
            Else
                Message "Color count successful (%1 BPP, indexed color)", targetImage.OriginalColorDepth
            End If
                        
        End If
        
        
                
        '*************************************************************************************************************************************
        ' Determine a name for this image
        '*************************************************************************************************************************************
        
        If isThisPrimaryImage Then Message "Determining image title..."
        
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
            
        
        
        '*************************************************************************************************************************************
        ' If this is a primary image, render it to the screen and update all relevant interface elements
        '*************************************************************************************************************************************
                
        'If this is a primary image, it needs to be rendered to the screen
        If isThisPrimaryImage Then
            
            'If the form isn't maximized or minimized then set its dimensions to just slightly bigger than the image size
            Message "Resizing image to fit screen..."
    
            'If the user wants us to resize the image to fit on-screen, do that now
            If g_AutosizeLargeImages = 0 Then FitImageToViewport True
                    
            'If the window is not maximized or minimized, fit the form around the picture box
            If FormMain.ActiveForm.WindowState = 0 Then FitWindowToImage True, True
            
            'Update relevant user interface controls
            DisplaySize targetImage.Width, targetImage.Height
            
            If imgFormTitle = "" Then
                If g_UserPreferences.GetPreference_Long("General Preferences", "ImageCaptionSize", 0) = 0 Then
                    FormMain.ActiveForm.Caption = getFilename(sFile(thisImage))
                Else
                    FormMain.ActiveForm.Caption = sFile(thisImage)
                End If
            Else
                FormMain.ActiveForm.Caption = imgFormTitle
            End If
            
            'Check the image's color depth, and check/uncheck the matching Image Mode setting
            If targetImage.mainLayer.getLayerColorDepth() = 32 Then tInit tImgMode32bpp, True Else tInit tImgMode32bpp, False
            
            'g_FixScrolling may have been reset by this point (by the FitImageToViewport sub, among others), so MAKE SURE it's false
            g_FixScrolling = False
            FormMain.CmbZoom.ListIndex = targetImage.CurrentZoomValue
        
            'Now that the image is loaded, allow PrepareViewport to set up the scrollbars and buffer
            g_FixScrolling = True
        
            PrepareViewport FormMain.ActiveForm, "PreLoadImage"
            
            'Render an icon-sized version of this image as the MDI child form's icon
            If MacroStatus <> MacroBATCH Then CreateCustomFormIcon FormMain.ActiveForm
            
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
            If g_AutosizeLargeImages = 1 Then FitWindowToViewport
        
            'Finally, add this file to the MRU list (unless specifically told not to)
            If ToUpdateMRU And (pageNumber = 0) And (MacroStatus <> MacroBATCH) Then MRU_AddNewFile sFile(thisImage), targetImage
        
        End If
        
        
        
        '*************************************************************************************************************************************
        ' Image loaded successfully.
        '*************************************************************************************************************************************
        
        targetImage.loadedSuccessfully = True
        
        If isThisPrimaryImage Then Message "Image loaded successfully."
        
        
        
        '*************************************************************************************************************************************
        ' If the just-loaded image was in a multipage format (icon, animated GIF, multipage TIFF), perform a few extra checks.
        '*************************************************************************************************************************************
        
        'Before continuing on to the next image (if any), see if the just-loaded image was in multipage format.  If it was, the user
        ' may have requested that we load all frames from this image.
        If imageHasMultiplePages Then
            
            Dim pageTracker As Long
            
            Dim tmpStringArray(0) As String
            tmpStringArray(0) = sFile(thisImage)
            
            'Call PreLoadImage again for each individual frame in the multipage file
            For pageTracker = 1 To imagePageCount
                If UCase(GetExtension(sFile(thisImage))) = "GIF" Then
                    PreLoadImage tmpStringArray, False, targetImage.OriginalFileName & " (" & g_Language.TranslateMessage("frame") & " " & (pageTracker + 1) & ")." & GetExtension(sFile(thisImage)), targetImage.OriginalFileName & " (" & g_Language.TranslateMessage("frame") & " " & (pageTracker + 1) & ")." & GetExtension(sFile(thisImage)), , , , pageTracker
                ElseIf UCase(GetExtension(sFile(thisImage))) = "ICO" Then
                    PreLoadImage tmpStringArray, False, targetImage.OriginalFileName & " (" & g_Language.TranslateMessage("icon") & " " & (pageTracker + 1) & ")." & GetExtension(sFile(thisImage)), targetImage.OriginalFileName & " (" & g_Language.TranslateMessage("icon") & " " & (pageTracker + 1) & ")." & GetExtension(sFile(thisImage)), , , , pageTracker
                Else
                    PreLoadImage tmpStringArray, False, targetImage.OriginalFileName & " (" & g_Language.TranslateMessage("page") & " " & (pageTracker + 1) & ")." & GetExtension(sFile(thisImage)), targetImage.OriginalFileName & " (" & g_Language.TranslateMessage("page") & " " & (pageTracker + 1) & ")." & GetExtension(sFile(thisImage)), , , , pageTracker
                End If
            Next pageTracker
        
        End If
        
        
        
    '*************************************************************************************************************************************
    ' Move on to the next image.
    '*************************************************************************************************************************************
        
PreloadMoreImages:

    'If we have more images to process, now's the time to do it!
    Next thisImage
        
    
    
    '*************************************************************************************************************************************
    ' Before finishing, display any relevant load problems (missing files, invalid formats, etc)
    '*************************************************************************************************************************************
    
    'Restore the screen cursor if necessary
    If pageNumber <= 0 Then Screen.MousePointer = vbNormal
    
    'If multiple images were loaded and everything went well, display a success message
    If multipleFilesLoading And (Len(missingFiles) = 0) And (Len(brokenFiles) = 0) And isThisPrimaryImage Then Message "All images loaded successfully."
        
    'Finally, if we were loading multiple images and something went wrong (missing files, broken files), let the user know about them.
    If multipleFilesLoading And (Len(missingFiles) > 0) Then
        Message "All images loaded, except for those that could not be found."
        pdMsgBox "Unfortunately, PhotoDemon was unable to find the following image(s):" & vbCrLf & vbCrLf & "%1" & vbCrLf & vbCrLf & "If these images were originally located on removable media (DVD, USB drive, etc), please re-insert or re-attach the media and try again.", vbApplicationModal + vbExclamation + vbOKOnly, "Image files missing", missingFiles
    End If
        
    If multipleFilesLoading And (Len(brokenFiles) > 0) Then
        Message "All images loaded, except for those in invalid formats."
        pdMsgBox "Unfortunately, PhotoDemon was unable to load the following image(s):" & vbCrLf & vbCrLf & "%1" & vbCrLf & vbCrLf & "Please use another program to save these images in a generic format (such as JPEG or PNG) before loading them into PhotoDemon. Thanks!", vbExclamation + vbOKOnly + vbApplicationModal, "Image Formats Not Supported", brokenFiles
    End If
        
End Sub

'Load any file that hasn't explicitly been sent elsewhere.  FreeImage will automatically determine filetype.
Public Function LoadFreeImageV3(ByVal sFile As String, ByRef dstLayer As pdLayer, ByRef dstImage As pdImage, Optional ByVal pageNumber As Long = 0, Optional ByVal showMessages As Boolean = True) As Boolean

    LoadFreeImageV3 = LoadFreeImageV3_Advanced(sFile, dstLayer, dstImage, pageNumber, showMessages)
    
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
    
    On Error GoTo LoadVBImageFail
    
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
    Exit Function
    
LoadVBImageFail:

    LoadVBImage = False
    Exit Function
    
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

    'Load messages are translatable, but we don't want to translate them if the translation object isn't ready yet
    If (Not (g_Language Is Nothing)) Then
        If g_Language.readyToTranslate Then
            If g_Language.translationActive Then sMsg = g_Language.TranslateMessage(sMsg)
        End If
    End If
    
    If FormSplash.Visible Then
        FormSplash.lblMessage = sMsg
        FormSplash.lblMessage.Refresh
    End If
    'DoEvents
    
End Sub

'Generates all shortcuts that VB can't; many thanks to Steve McMahon for his accelerator class, which helps a great deal
Public Sub LoadMenuShortcuts()

    'Don't allow custom shortcuts in the IDE, as they require subclassing and might crash
    If Not g_IsProgramCompiled Then Exit Sub

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
        .AddAccelerator vbKeyB, vbCtrlMask Or vbShiftMask, "Bright_Contrast"
        .AddAccelerator vbKeyC, vbCtrlMask Or vbShiftMask, "Color_Balance"
        .AddAccelerator vbKeyH, vbCtrlMask Or vbShiftMask, "Shadow_Highlight"
        
        'Window menu
        .AddAccelerator vbKeyPageUp, 0, "Prev_Image"
        .AddAccelerator vbKeyPageDown, 0, "Next_Image"
        
        'No equivalent menu
        .AddAccelerator vbKeyEscape, 0, "Escape"
        
        .Enabled = True
    End With

    DrawMenuShortcuts
    
End Sub

'After all menu shortcuts (accelerators) are loaded above, the custom shortcuts need to be added to the menu entries themselves.
' If we don't do this, the user won't know how to trigger the shortcuts!
Public Sub DrawMenuShortcuts()

    'Don't allow custom shortcuts in the IDE, as they require subclassing and might crash
    If Not g_IsProgramCompiled Then Exit Sub

    'File menu
    FormMain.MnuSaveAs.Caption = FormMain.MnuSaveAs.Caption & vbTab & "Ctrl+Shift+S"
    FormMain.MnuImportFromInternet.Caption = FormMain.MnuImportFromInternet.Caption & vbTab & "Ctrl+Shift+I"
    FormMain.MnuScreenCapture.Caption = FormMain.MnuScreenCapture.Caption & vbTab & "Ctrl+Alt+I"
    FormMain.MnuImportFrx.Caption = FormMain.MnuImportFrx.Caption & vbTab & "Ctrl+Alt+Shift+I"
    FormMain.MnuImportClipboard.Caption = FormMain.MnuImportClipboard.Caption & vbTab & "Ctrl+V"
    
    'NOTE: Drawing of MRU shortcuts is handled in the MRU module

    'Edit menu
    'This Redo shortcut remains, but it is hidden; the Windows convention of Ctrl+Y is displayed instead.
    'FormMain.MnuRedo.Caption = FormMain.MnuRedo.Caption & vbTab & "Ctrl+Alt+Z"
    FormMain.MnuEmptyClipboard.Caption = FormMain.MnuEmptyClipboard.Caption & vbTab & "Ctrl+Shift+X"
    
    'View menu
    FormMain.MnuFitOnScreen.Caption = FormMain.MnuFitOnScreen.Caption & vbTab & "0"
    FormMain.MnuZoomIn.Caption = FormMain.MnuZoomIn.Caption & vbTab & " +"
    FormMain.MnuZoomOut.Caption = FormMain.MnuZoomOut.Caption & vbTab & "-"
    FormMain.MnuSpecificZoom(0).Caption = FormMain.MnuSpecificZoom(0).Caption & vbTab & "5"
    FormMain.MnuSpecificZoom(1).Caption = FormMain.MnuSpecificZoom(1).Caption & vbTab & "4"
    FormMain.MnuSpecificZoom(2).Caption = FormMain.MnuSpecificZoom(2).Caption & vbTab & "3"
    FormMain.MnuSpecificZoom(3).Caption = FormMain.MnuSpecificZoom(3).Caption & vbTab & "2"
    FormMain.MnuSpecificZoom(4).Caption = FormMain.MnuSpecificZoom(4).Caption & vbTab & "1"
    FormMain.MnuSpecificZoom(5).Caption = FormMain.MnuSpecificZoom(5).Caption & vbTab & "Shift+2"
    FormMain.MnuSpecificZoom(6).Caption = FormMain.MnuSpecificZoom(6).Caption & vbTab & "Shift+3"
    FormMain.MnuSpecificZoom(7).Caption = FormMain.MnuSpecificZoom(7).Caption & vbTab & "Shift+4"
    FormMain.MnuSpecificZoom(8).Caption = FormMain.MnuSpecificZoom(8).Caption & vbTab & "Shift+5"
        
    'Image menu
    FormMain.MnuRotateClockwise.Caption = FormMain.MnuRotateClockwise.Caption & vbTab & "R"
    FormMain.MnuRotate270Clockwise.Caption = FormMain.MnuRotate270Clockwise.Caption & vbTab & "L"
    FormMain.MnuCropSelection.Caption = FormMain.MnuCropSelection.Caption & vbTab & "Ctrl+Shift+X"
    
    'Color menu
    FormMain.MnuColor(0).Caption = FormMain.MnuColor(0).Caption & vbTab & "Ctrl+Shift+B"
    FormMain.MnuColor(1).Caption = FormMain.MnuColor(1).Caption & vbTab & "Ctrl+Shift+C"
    FormMain.MnuColor(5).Caption = FormMain.MnuColor(5).Caption & vbTab & "Ctrl+Shift+H"
    
    'Tools menu
    FormMain.mnuTool(4).Caption = FormMain.mnuTool(4).Caption & vbTab & "Alt+Enter"     'Options (Preferences)
    
    'Window menu (Next Image and Previous Image, specifically)
    FormMain.MnuWindow(0).Caption = FormMain.MnuWindow(0).Caption & vbTab & "Page Down"
    FormMain.MnuWindow(1).Caption = FormMain.MnuWindow(1).Caption & vbTab & "Page Up"
    
End Sub

'This subroutine handles the detection of the three core plugins strongly recommended for an optimal PhotoDemon
' experience: zLib, EZTwain32, and FreeImage.  For convenience' sake, it also checks for GDI+ availability.
Public Sub LoadPlugins()
    
    'Plugin files are located in the \Data\Plugins subdirectory
    g_PluginPath = g_UserPreferences.getAppPath & "Plugins\"
    
    'Make sure the plugin path exists
    If Not DirectoryExist(g_PluginPath) Then MkDir g_PluginPath
    
    'Old versions of PhotoDemon kept plugins in a different directory. Check the old location,
    ' and if plugin-related files are found, copy them to the new directory
    On Error Resume Next
    Dim tmpg_PluginPath As String
    tmpg_PluginPath = g_UserPreferences.getDataPath & "Plugins\"
    
    If DirectoryExist(tmpg_PluginPath) Then
        LoadMessage "Moving plugins to updated folder location..."
        
        Dim pluginName As String
        pluginName = "EZTW32.dll"
        If FileExist(tmpg_PluginPath & pluginName) Then
            FileCopy tmpg_PluginPath & pluginName, g_PluginPath & pluginName
            Kill tmpg_PluginPath & pluginName
        End If
        
        pluginName = "EZTWAIN_README.TXT"
        If FileExist(tmpg_PluginPath & pluginName) Then
            FileCopy tmpg_PluginPath & pluginName, g_PluginPath & pluginName
            Kill tmpg_PluginPath & pluginName
        End If
        
        pluginName = "FreeImage.dll"
        If FileExist(tmpg_PluginPath & pluginName) Then
            FileCopy tmpg_PluginPath & pluginName, g_PluginPath & pluginName
            Kill tmpg_PluginPath & pluginName
        End If
        
        pluginName = "license-fi.txt"
        If FileExist(tmpg_PluginPath & pluginName) Then
            FileCopy tmpg_PluginPath & pluginName, g_PluginPath & pluginName
            Kill tmpg_PluginPath & pluginName
        End If
        
        pluginName = "license-freeimage.txt"
        If FileExist(tmpg_PluginPath & pluginName) Then
            FileCopy tmpg_PluginPath & pluginName, g_PluginPath & pluginName
            Kill tmpg_PluginPath & pluginName
        End If
        
        pluginName = "license-gplv2.txt"
        If FileExist(tmpg_PluginPath & pluginName) Then
            FileCopy tmpg_PluginPath & pluginName, g_PluginPath & pluginName
            Kill tmpg_PluginPath & pluginName
        End If
        
        pluginName = "license-gplv3.txt"
        If FileExist(tmpg_PluginPath & pluginName) Then
            FileCopy tmpg_PluginPath & pluginName, g_PluginPath & pluginName
            Kill tmpg_PluginPath & pluginName
        End If
        
        pluginName = "zlibwapi.dll"
        If FileExist(tmpg_PluginPath & pluginName) Then
            FileCopy tmpg_PluginPath & pluginName, g_PluginPath & pluginName
            Kill tmpg_PluginPath & pluginName
        End If
        
        pluginName = "pngnq-s9.exe"
        If FileExist(tmpg_PluginPath & pluginName) Then
            FileCopy tmpg_PluginPath & pluginName, g_PluginPath & pluginName
            Kill tmpg_PluginPath & pluginName
        End If
        
        pluginName = "PNGNQ-S9-LICENSE"
        If FileExist(tmpg_PluginPath & pluginName) Then
            FileCopy tmpg_PluginPath & pluginName, g_PluginPath & pluginName
            Kill tmpg_PluginPath & pluginName
        End If
        
        pluginName = "PNGNQ-S9-LICENSE.txt"
        If FileExist(tmpg_PluginPath & pluginName) Then
            FileCopy tmpg_PluginPath & pluginName, g_PluginPath & pluginName
            Kill tmpg_PluginPath & pluginName
        End If
        
        pluginName = "PNGNQ-S9-LICENCE.txt"
        If FileExist(tmpg_PluginPath & pluginName) Then
            FileCopy tmpg_PluginPath & pluginName, g_PluginPath & pluginName
            Kill tmpg_PluginPath & pluginName
        End If
        
        'After all files have been removed, kill the old Plugin directory
        RmDir tmpg_PluginPath
        
    End If
        
    'Check for image scanning
    'First, make sure we have our dll file
    If isEZTwainAvailable Then
                
        'If we do find the DLL, check to see if EZTwain has been forcibly disabled by the user.
        If g_UserPreferences.GetPreference_Boolean("Plugin Preferences", "ForceEZTwainDisable", False) Then
            g_ScanEnabled = False
        Else
            g_ScanEnabled = True
        End If
        
    Else
        
        'If we can't find the DLL, hide the menu options and internally disable scanning
        '(perhaps overkill, but it acts as a safeguard to prevent bad DLL-based crashes)
        g_ScanEnabled = False
        
    End If
    
        'Additionally related to EZTwain - enable/disable the various scanner options contigent on EZTwain's enabling
        FormMain.MnuScanImage.Visible = g_ScanEnabled
        FormMain.MnuSelectScanner.Visible = g_ScanEnabled
        FormMain.MnuImportSepBar1.Visible = g_ScanEnabled
    
    'Check for zLib compression capabilities
    If isZLibAvailable Then
    
        'Check to see if zLib has been forcibly disabled.
        If g_UserPreferences.GetPreference_Boolean("Plugin Preferences", "ForceZLibDisable", False) Then
            g_ZLibEnabled = False
        Else
            g_ZLibEnabled = True
        End If
        
    Else
        g_ZLibEnabled = False
    End If
    
    'Check for FreeImage file interface
    If isFreeImageAvailable Then
        
        'Check to see if FreeImage has been forcibly disabled
        If g_UserPreferences.GetPreference_Boolean("Plugin Preferences", "ForceFreeImageDisable", False) Then
            g_ImageFormats.FreeImageEnabled = False
        Else
            g_ImageFormats.FreeImageEnabled = True
        End If
        
    Else
        g_ImageFormats.FreeImageEnabled = False
    End If
    
        'Additionally related to FreeImage - enable/disable the arbitrary rotation option contingent on FreeImage's enabling
        FormMain.MnuRotateArbitrary.Visible = g_ImageFormats.FreeImageEnabled
    
    'Check for pngnq interface
    If isPngnqAvailable Then
        
        'Check to see if pngnq-s9 has been forcibly disabled
        If g_UserPreferences.GetPreference_Boolean("Plugin Preferences", "ForcePngnqDisable", False) Then
            g_ImageFormats.pngnqEnabled = False
        Else
            g_ImageFormats.pngnqEnabled = True
        End If
        
    Else
        g_ImageFormats.pngnqEnabled = False
    End If
    
    'Finally, check GDI+ availability
    If isGDIPlusAvailable() Then
        g_ImageFormats.GDIPlusEnabled = True
    Else
        g_ImageFormats.GDIPlusEnabled = False
    End If
    
    
    
End Sub

'Make a copy of the current image.  Thanks to PSC user "Achmad Junus" for this suggestion.
Public Sub DuplicateCurrentImage()
    
    Message "Duplicating current image..."
    
    'First, make a note of the currently active form
    Dim imageToBeDuplicated As Long
    imageToBeDuplicated = CurrentImage
    
    CreateNewImageForm
        
    g_FixScrolling = False
        
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
    newFilename = pdImages(imageToBeDuplicated).OriginalFileName & " - " & g_Language.TranslateMessage("Copy")
    pdImages(CurrentImage).OriginalFileName = newFilename
    pdImages(CurrentImage).OriginalFileNameAndExtension = newFilename & "." & originalExtension
            
    'Because this image hasn't been saved to disk, mark its save state as "false"
    pdImages(CurrentImage).UpdateSaveState False
    
    'Fit the window to the newly duplicated image
    Message "Resizing image to fit screen..."
    
    'If the user wants us to resize the image to fit on-screen, do that now
    If g_AutosizeLargeImages = 0 Then
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
        
    'g_FixScrolling may have been reset by this point (by the FitImageToViewport sub, among others), so MAKE SURE it's false
    g_FixScrolling = False
    FormMain.CmbZoom.ListIndex = pdImages(CurrentImage).CurrentZoomValue
        
    Message "Image duplication complete."
    
    'Now that the image is loaded, allow PrepareViewport to set up the scrollbars and buffer
    g_FixScrolling = True
    
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

'Check for IDE or compiled EXE, and set program parameters accordingly
Private Sub CheckLoadingEnvironment()
    
    'Check the run-time environment.
    
    'App is compiled:
    If App.LogMode = 1 Then
        
        g_IsProgramCompiled = True
        
        'Determine the version automatically from the EXE information
        FormSplash.lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
                
    'App is not compiled:
    Else
    
        g_IsProgramCompiled = False

        'Add a gentle reminder to compile the program
        FormSplash.lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision & " - please compile!"
        
    End If
    
End Sub
