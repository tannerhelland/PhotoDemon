Attribute VB_Name = "Loading"
'***************************************************************************
'Program/File Loading Handler
'Copyright ©2001-2013 by Tanner Helland
'Created: 4/15/01
'Last updated: 16/October/13
'Last update: register newly loaded images with the image thumbnail control
'
'Module for handling any and all program loading.  This includes the program itself,
' plugins, files, and anything else the program needs to take from the hard drive.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Use these to ensure that the splash shows for a certain amount of time
Dim m_LoadTime As Double, m_StartTime As Double

'IT ALL BEGINS HERE (after Sub Main, that is).
' Note that this function is called AFTER FormMain has been loaded.  FormMain is loaded - but not visible - so it can be operated
' on by functions called from this routine.  (It is necessary to load the main form first, since a number of these operations -
' like loading all PNG menu icons from the resource file - operate on the main form.)
Public Sub LoadTheProgram()
    
    m_StartTime = Timer
    
    '*************************************************************************************************************************************
    ' Prepare the splash screen (but don't display it yet)
    '*************************************************************************************************************************************
    
    'We need GDI+ to extract a JPEG from the resource file and convert it in-memory.  (Yes, there are other ways to do this.  No, I don't
    ' care about using them.)  Check it's availability.
    If isGDIPlusAvailable() Then
    
        'Load FormSplash into memory, but don't make it visible.  Then ask it to prepare itself.
        FormSplash.Visible = False
        FormSplash.prepareSplash
        
    End If
        
    'Check the environment.  If inside the the IDE, the splash needs to be modified slightly.
    CheckLoadingEnvironment
    
    If g_GDIPlusAvailable Then
        If g_IsProgramCompiled Then m_LoadTime = 1.5 Else m_LoadTime = 1#
    Else
        m_LoadTime = 0#
    End If
    
    '*************************************************************************************************************************************
    ' Determine which version of Windows the user is running (as other load functions rely on this)
    '*************************************************************************************************************************************
    
    LoadMessage "Detecting Windows® version..."
    
    'Certain features are OS-version dependent.  We must determine the OS version early in the load process to ensure that all features
    ' are loaded correctly.
    g_IsVistaOrLater = getVistaOrLaterStatus
    g_IsWin7OrLater = getWin7OrLaterStatus
    
    
    
    '*************************************************************************************************************************************
    ' If the user doesn't have font smoothing enabled, enable it now.  PD's interface looks much better with some form of antialiasing.
    '*************************************************************************************************************************************
    
    handleClearType True
    
    
    
    '*************************************************************************************************************************************
    ' Initialize the user preferences (settings) handler
    '*************************************************************************************************************************************
    
    Set g_UserPreferences = New pdPreferences
    
    'Ask the preferences handler to generate key program folders.  (If these folders don't exist, the handler will create them.)
    LoadMessage "Initializing all program directories..."
    
    g_UserPreferences.initializePaths
    
    'Now, ask the preferences handler to load all other user settings from the preferences file
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
    wRect.Left = g_UserPreferences.GetPref_Long("Core", "Last Window Left", 1)
    wRect.Top = g_UserPreferences.GetPref_Long("Core", "Last Window Top", 1)
    wRect.Right = wRect.Left + g_UserPreferences.GetPref_Long("Core", "Last Window Width", 1)
    wRect.Bottom = wRect.Top + g_UserPreferences.GetPref_Long("Core", "Last Window Height", 1)
    g_cMonitors.CenterFormOnMonitor FormSplash, , wRect.Left, wRect.Right, wRect.Top, wRect.Bottom
    
    'Make the splash screen's message display match the rest of PD
    If g_IsVistaOrLater And g_UseFancyFonts Then
        g_InterfaceFont = "Segoe UI"
    Else
        g_InterfaceFont = "Tahoma"
    End If
    FormSplash.lblMessage.FontName = g_InterfaceFont
    
    'Display the splash screen, centered on whichever monitor the user previously used the program on.
    FormSplash.Show vbModeless
    
    
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
    ' Initialize the visual themes engine
    '*************************************************************************************************************************************
    
    'Because this class subclasses all forms in the project, it must be loaded very early in the start process
    LoadMessage "Initializing theme engine..."
    
    Set g_Themer = New pdVisualThemes
    
    
    
    '*************************************************************************************************************************************
    ' Get the viewport engine ready
    '*************************************************************************************************************************************
    
    'Initialize our current zoom method
    LoadMessage "Initializing viewport engine..."
    
    initializeViewportEngine
    
    
    
    '*************************************************************************************************************************************
    ' Initialize the window manager (the class that synchronizes all toolbox and image window positions)
    '*************************************************************************************************************************************
    
    LoadMessage "Initializing window manager..."
    Set g_WindowManager = New pdWindowManager
    
    'Register the main form
    g_WindowManager.registerParentForm FormMain
    
    'Load all tool windows.  Even though they may not be visible (as the user can elect to hide them), we still want them loaded,
    ' so we can interact with them as necessary (e.g. "enable Undo button", etc).
    Load toolbar_File
    Load toolbar_ImageTabs
    Load toolbar_Selections
        
    'Retrieve floating window status from the preferences file, mark their menus, and pass their values to the window manager
    toggleWindowFloating TOOLBAR_WINDOW, g_UserPreferences.GetPref_Boolean("Core", "Floating Toolbars", False), True
    toggleWindowFloating IMAGE_WINDOW, g_UserPreferences.GetPref_Boolean("Core", "Floating Image Windows", False), True
    
    'Retrieve visibility and mark those menus as well
    FormMain.MnuWindow(0).Checked = g_UserPreferences.GetPref_Boolean("Core", "Show File Toolbox", True)
    FormMain.MnuWindow(1).Checked = g_UserPreferences.GetPref_Boolean("Core", "Show Selections Toolbox", True)
    
    
    
    '*************************************************************************************************************************************
    ' Set all default tool values
    '*************************************************************************************************************************************
        
    LoadMessage "Initializing image tools..."
        
    'Note that selection tools are initialized in the selection toolbar's Form_Load event
            
        
        
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
    
    'Manually create multi-line tooltips for some command buttons
    toolbar_File.cmdOpen.ToolTip = g_Language.TranslateMessage("Open one or more images for editing." & vbCrLf & vbCrLf & "(Another way to open images is dragging them from your desktop or Windows Explorer and dropping them onto PhotoDemon.)")
    If g_ConfirmClosingUnsaved Then
        toolbar_File.cmdClose.ToolTip = g_Language.TranslateMessage("Close the current image." & vbCrLf & vbCrLf & "If the current image has not been saved, you will receive a prompt to save it before it closes.")
    Else
        toolbar_File.cmdClose.ToolTip = g_Language.TranslateMessage("Close the current image." & vbCrLf & vbCrLf & "Because you have turned off save prompts (via Edit -> Preferences), you WILL NOT receive a prompt to save this image before it closes.")
    End If
    toolbar_File.cmdSave.ToolTip = g_Language.TranslateMessage("Save the current image." & vbCrLf & vbCrLf & "WARNING: this will overwrite the current image file.  To save to a different file, use the ""Save As"" button.")
    toolbar_File.cmdSaveAs.ToolTip = g_Language.TranslateMessage("Save the current image to a new file.")
                        
    'Use the API to give PhotoDemon's main form a 32-bit icon (VB is too old to support 32bpp icons)
    SetIcon FormMain.hWnd, "AAA", True
    
    'Initialize all system cursors we rely on (hand, busy, resizing, etc)
    initAllCursors
    
    'Set up the program's title bar.  Odd-numbered releases are development releases.  Even-numbered releases are formal builds.
    If (CLng(App.Minor) Mod 2 = 0) Then
        FormMain.Caption = App.Title & " " & App.Major & "." & App.Minor
    Else
        'Odd-numbered development releases of the pattern X.9 are betas for the next major version, e.g. (X+1).0
        If App.Minor = 9 Then
            FormMain.Caption = App.Title & " " & (App.Major + 1) & ".0 beta (build " & App.Revision & ")"
        Else
            FormMain.Caption = App.Title & " " & App.Major & "." & CStr(App.Minor + 1) & " beta (build " & App.Revision & ")"
        End If
    End If
    
    'PhotoDemon renders many of its own icons dynamically.  Initialize that engine now.
    initializeIconHandler
    
    'Before displaying the main window, see if the user wants to restore last-used window location.
    restoreMainWindowLocation
        
    'If Segoe UI is in use, the zoom buttons need to be adjusted to match the combo box
    If g_UseFancyFonts Then
        toolbar_File.cmdZoomIn.Height = toolbar_File.cmdZoomIn.Height + 1
        toolbar_File.cmdZoomOut.Height = toolbar_File.cmdZoomOut.Height + 1
    End If
    
    'Allow drag-and-drop operations
    g_AllowDragAndDrop = True
    
    'Apply visual styles
    FormMain.requestMakeFormPretty
    
    
    '*************************************************************************************************************************************
    ' The program's menus support many features that VB can't do natively (like icons and custom shortcuts).  Load such things now.
    '*************************************************************************************************************************************
    
    LoadMessage "Preparing program menus..."
    
    'If inside the IDE, disable the "Effects" -> "Test" menu
    If g_IsProgramCompiled Then FormMain.MnuTest.Visible = False Else FormMain.MnuTest.Visible = True
    
    'Create all manual shortcuts (ones VB isn't capable of generating itself)
    LoadAccelerators
            
    'Initialize the Recent Files manager and load the most-recently-used file list (MRU)
    Set g_RecentFiles = New pdRecentFiles
    g_RecentFiles.MRU_LoadFromFile
            
    'Load and draw all menu icons
    loadMenuIcons
    'resetMenuIcons
    
    'Synchronize all other interface elements to match the current program state (e.g. no images loaded).
    syncInterfaceToCurrentImage
    
    
    
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
        .Font.Name = g_InterfaceFont
        .Font.Size = 9
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
        
    'Display the splash screen for at least a second or two
    If Timer - m_StartTime < m_LoadTime Then
        Do While Timer - m_StartTime < m_LoadTime
        Loop
    End If
        
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
        Dim i As Long
        For i = 1 To Len(g_CommandLine)
            
            tChar = Mid(g_CommandLine, i, 1)
                
            'If the current character is a quotation mark, change inQuotes to specify that we are either inside
            ' or outside a SET of quotation marks (note: they will always occur in pairs, per the rules of
            ' how Windows handles command line parameters)
            If tChar = Chr(34) Then inQuotes = Not inQuotes
                
            'If the current character is a space...
            If tChar = Chr(32) Then
                    
                '...check to see if we are inside quotation marks.  If we are, that means this space is part of a
                ' filename and NOT a delimiter.  Replace it with an asterisk.
                If inQuotes = True Then g_CommandLine = Left(g_CommandLine, i - 1) & "*" & Right(g_CommandLine, Len(g_CommandLine) - i)
                    
            End If
            
        Next i
            
        'At this point, spaces that are parts of filenames have been replaced by asterisks.  That means we can use
        ' Split() to fill our filename array, because the only spaces remaining in the command line are delimiters
        ' between filenames.
        sFile = Split(g_CommandLine, Chr(32))
            
        'Now that our filenames are successfully inside the sFile() array, go back and replace our asterisk placeholders
        ' with spaces.  Also, remove any quotation marks (since those aren't technically part of the filename).
        For i = 0 To UBound(sFile)
            sFile(i) = Replace$(sFile(i), Chr(42), Chr(32))
            sFile(i) = Replace$(sFile(i), Chr(34), "")
        Next i
        
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
            
    'One of the things we'll be doing in this routine is establishing an original color depth for this image. FreeImage will return
    ' this automatically; GDI+ may not.  Use a tracking variable to determine if a manual color count needs to be performed.
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
        g_imageHasMultiplePages = False
        
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
        
        If isThisPrimaryImage Then
            
            Message "Image found. Initializing blank form..."

            CreateNewImageForm
        
            Set targetImage = pdImages(g_CurrentImage)
            Set targetLayer = pdImages(g_CurrentImage).mainLayer
        
            g_AllowViewportRendering = False
        
            pdImages(g_CurrentImage).containingForm.HScroll.value = 0
            pdImages(g_CurrentImage).containingForm.VScroll.value = 0
        
        End If
            
            
            
        '*************************************************************************************************************************************
        ' Call the most appropriate load function for this image's format (FreeImage, GDI+, or VB's LoadPicture)
        '*************************************************************************************************************************************
            
        If isThisPrimaryImage Then Message "Determining filetype..."
        
        'Initially, set the filetype of the target image to "unknown".  If the load is successful, this value will
        ' be changed to something >= 0. (Note: if FreeImage is used to load the file, this value will be set by the
        ' LoadFreeImageV3 function.)
        If Not (targetImage Is Nothing) Then targetImage.originalFileFormat = -1
        
        'Strip the extension from the file
        FileExtension = UCase(GetExtension(sFile(thisImage)))
        
        loadSuccessful = False
        loadedByOtherMeans = False
            
        'Depending on the file's extension, load the image using the most appropriate image decoding routine
        Select Case FileExtension
        
            'PhotoDemon's custom file format must be handled specially (as obviously, FreeImage and GDI+ won't handle it)
            Case "PDI"
            
                'PDI images require zLib, and are only loaded via a custom routine (obviously, since they are PhotoDemon's native format)
                loadSuccessful = LoadPhotoDemonImage(sFile(thisImage), targetLayer, targetImage)
                
                targetImage.originalFileFormat = 100
                mustCountColors = True
        
            'TMP files are internal files (BMP format) used by PhotoDemon.  GDI+ is preferable, but .LoadPicture works too.
            Case "TMP"
            
                If g_ImageFormats.GDIPlusEnabled Then loadSuccessful = LoadGDIPlusImage(sFile(thisImage), targetLayer, targetImage)
                
                If (Not g_ImageFormats.GDIPlusEnabled) Or (Not loadSuccessful) Then loadSuccessful = LoadVBImage(sFile(thisImage), targetLayer, targetImage)
                
                targetImage.originalFileFormat = FIF_JPEG
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
                    loadSuccessful = LoadGDIPlusImage(sFile(thisImage), targetLayer, targetImage)
                    
                    'If GDI+ loaded the image successfully, note that we have to determine color depth manually
                    If loadSuccessful Then
                        loadedByOtherMeans = True
                        mustCountColors = True
                    End If
                        
                End If
                
                'If both FreeImage and GDI+ failed, give the image one last try with VB's LoadPicture
                If (Not loadSuccessful) Then
                    
                    If isThisPrimaryImage Then Message "GDI+ refused to load image.  Dropping back to internal routines and trying again..."
                    loadSuccessful = LoadVBImage(sFile(thisImage), targetLayer, targetImage)
                
                    'If VB managed to load the image successfully, note that we have to deteremine color depth manually
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
            
            'Deactivating the image will remove the reference to the containing form - this is desired behavior, because VB counts object references,
            ' and doesn't free an object until all references are resolved.  To work around this, we make a copy of the reference prior to deactivation.
            ' We then use this reference copy to unload the form.
            Dim failedImageForm As Form
            Set failedImageForm = targetImage.containingForm
            failedImageForm.Visible = False
            
            'If multiple files are being loaded, suppress any errors until the end
            If multipleFilesLoading Then
                brokenFiles = brokenFiles & getFilename(sFile(thisImage)) & vbCrLf
            Else
                If MacroStatus <> MacroBATCH Then pdMsgBox "Unfortunately, PhotoDemon was unable to load the following image:" & vbCrLf & vbCrLf & "%1" & vbCrLf & vbCrLf & "Please use another program to save this image in a generic format (such as JPEG or PNG) before loading it into PhotoDemon.  Thanks!", vbExclamation + vbOKOnly + vbApplicationModal, "Image Import Failed", sFile(thisImage)
            End If
            
            targetImage.deactivateImage
            
            If isThisPrimaryImage Then
                'Notify the window manager that this hWnd will soon be dead, so we can stop subclassing it
                g_WindowManager.unregisterForm failedImageForm
                Unload failedImageForm
                Set failedImageForm = Nothing
            End If
            
            'Update the interface to reflect the images currently loaded
            syncInterfaceToCurrentImage
            
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
                    targetImage.originalFileFormat = FIF_GIF
                    targetImage.originalColorDepth = 8
                    
                Case "ICO"
                    targetImage.originalFileFormat = FIF_ICO
                
                Case "JIF", "JPG", "JPEG", "JPE"
                    targetImage.originalFileFormat = FIF_JPEG
                    targetImage.originalColorDepth = 24
                    
                Case "PNG"
                    targetImage.originalFileFormat = FIF_PNG
                
                Case "TIF", "TIFF"
                    targetImage.originalFileFormat = FIF_TIFF
                    
                'Treat anything else as a BMP file
                Case Else
                    targetImage.originalFileFormat = FIF_BMP
                    
            End Select
        
        End If
        
        
        
        '*************************************************************************************************************************************
        ' If the image has an alpha channel, verify it.  If it contains all 0 or all 255, convert it to 24bpp to conserve resources.
        '*************************************************************************************************************************************
        
        If targetImage.mainLayer.getLayerColorDepth = 32 Then
            
            'Make sure the user hasn't disabled this capability
            If g_UserPreferences.GetPref_Boolean("Transparency", "Validate Alpha Channels", True) Then
            
                If isThisPrimaryImage Then Message "Verfiying alpha channel..."
            
                'Verify the alpha channel.  If this function returns FALSE, the alpha channel is unnecessary.
                If Not targetImage.mainLayer.verifyAlphaChannel Then
                
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
        ' Store some universally important information in the target image object
        '*************************************************************************************************************************************
        
        targetImage.updateSize
        If FileExist(sFile(thisImage)) Then targetImage.originalFileSize = FileLen(sFile(thisImage))
        targetImage.currentFileFormat = targetImage.originalFileFormat
                
                
        
        '*************************************************************************************************************************************
        ' If the image contains an embedded ICC profile, apply it now (before counting colors, etc).
        '*************************************************************************************************************************************
        
        If targetImage.ICCProfile.hasICCData Then targetImage.ICCProfile.applyICCtoParentImage targetImage
        
        
        
        '*************************************************************************************************************************************
        ' If requested by the user, manually count the number of unique colors in the image (to determine absolutely accurate color depth)
        '*************************************************************************************************************************************
                
        'At this point, we now have loaded image data in 24 or 32bpp format.  For future reference, let's count
        ' the number of colors present in the image (if the user has allowed it).  If the user HASN'T allowed
        ' it, we have no choice but to rely on whatever color depth was returned by FreeImage or GDI+ (or was
        ' inferred by us for this format, e.g. we know that GIFs are 8bpp).
        
        If isThisPrimaryImage And (g_UserPreferences.GetPref_Boolean("Loading", "Verify Initial Color Depth", True) Or mustCountColors) Then
            
            colorCountCheck = getQuickColorCount(targetImage, g_CurrentImage)
        
            'If 256 or less colors were found in the image, mark it as 8bpp.  Otherwise, mark it as 24 or 32bpp.
            targetImage.originalColorDepth = getColorDepthFromColorCount(colorCountCheck, targetImage.mainLayer)
            
            If g_IsImageGray Then
                Message "Color count successful (%1 BPP, grayscale)", targetImage.originalColorDepth
            Else
                Message "Color count successful (%1 BPP, indexed color)", targetImage.originalColorDepth
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
            targetImage.locationOnDisk = sFile(thisImage)
            
            tmpFilename = sFile(thisImage)
            StripFilename tmpFilename
            targetImage.originalFileNameAndExtension = tmpFilename
            StripOffExtension tmpFilename
            targetImage.originalFileName = tmpFilename
            
            'Disable the save button, because this file exists on disk
            targetImage.setSaveState True
            
        Else
            'The calling routine has specified a file name.  Assume this is a special case, and force a Save As...
            ' dialog in the future by not specifying a location on disk
            targetImage.locationOnDisk = ""
            targetImage.originalFileNameAndExtension = imgName
            
            tmpFilename = imgName
            StripOffExtension tmpFilename
            targetImage.originalFileName = tmpFilename
            
            'Similarly, enable the save button
            targetImage.setSaveState False
            
        End If
        
        
        
        '*************************************************************************************************************************************
        ' If the ExifTool plugin is available, and user preferences allow, extract any possible metadata from the image file
        '*************************************************************************************************************************************
        
        Set targetImage.imgMetadata = New pdMetadata
        
        If g_ExifToolEnabled And isThisPrimaryImage Then
                
            'A user preference determines whether we loda metadata now, or suspend it until requested (e.g. save or browse time)
            If g_UserPreferences.GetPref_Boolean("Loading", "Automatically Load Metadata", True) Then
                Message "Compiling metadata..."
                targetImage.imgMetadata.loadAllMetadata sFile(thisImage), targetImage.originalFileFormat
            End If
            
        End If
        
        
        '*************************************************************************************************************************************
        ' If this is a primary image, update all relevant interface elements (image size display, 24/32bpp options, custom form icon, etc)
        '*************************************************************************************************************************************
                
        'If this is a primary image, it needs to be rendered to the screen
        If isThisPrimaryImage Then
            
            If imgFormTitle = "" Then
                If g_UserPreferences.GetPref_Long("Interface", "Window Caption Length", 0) = 0 Then
                    g_WindowManager.requestWindowCaptionChange targetImage.containingForm, getFilename(sFile(thisImage))
                Else
                    g_WindowManager.requestWindowCaptionChange targetImage.containingForm, sFile(thisImage)
                End If
            Else
                g_WindowManager.requestWindowCaptionChange targetImage.containingForm, imgFormTitle
            End If
            
            'Create an icon-sized version of this image, which we will use as form's taskbar icon
            If MacroStatus <> MacroBATCH Then createCustomFormIcon targetImage.containingForm
            
            'Synchronize all other interface elements to match the newly loaded image
            syncInterfaceToCurrentImage
        
        '*************************************************************************************************************************************
        ' If this is a primary image, register the new window with the window manager, then display the image on-screen
        '*************************************************************************************************************************************
            
            
            'Register this window with PhotoDemon's window manager.  This will do things like set the proper border state depending on whether
            ' image windows are docked or floating, which we need before doing things like auto-zoom or window placement.
            g_WindowManager.registerChildForm targetImage.containingForm, IMAGE_WINDOW, , , g_CurrentImage
            
            'Also register this image with the image tab bar
            toolbar_ImageTabs.registerNewImage g_CurrentImage
            
            Message "Resizing image to fit screen..."
    
            'If the user wants us to resize the image to fit on-screen, do that now
            If g_AutozoomLargeImages = 0 Then FitImageToViewport True
                    
            'If image windows are not docked, fit the window around the (now properly-zoomed) image
            If g_WindowManager.getFloatState(IMAGE_WINDOW) Then FitWindowToImage True, True
            
            'g_AllowViewportRendering may have been reset by this point (by the FitImageToViewport sub, among others), so set it back to False, then
            ' update the zoom combo box to match the zoom assigned by the window-fit function.
            g_AllowViewportRendering = False
            toolbar_File.CmbZoom.ListIndex = targetImage.currentZoomValue
        
            'Now that the image's window has been fully sized and moved around, use PrepareViewport to set up any scrollbars and a back-buffer
            g_AllowViewportRendering = True
            PrepareViewport targetImage.containingForm, "PreLoadImage"
            
            'Note the window state, as it may be important in the future
            targetImage.WindowState = targetImage.containingForm.WindowState
            
            'If image windows are floating, move the window into place now using values previously calculated by FitToScreen
            If g_WindowManager.getFloatState(IMAGE_WINDOW) Then
                targetImage.containingForm.Move targetImage.WindowLeft * Screen.TwipsPerPixelX, targetImage.WindowTop * Screen.TwipsPerPixelY
                g_WindowManager.requestWindowResync targetImage.indexInWindowManager
                
            'If image windows are docked, the window manager will have already positioned the window for us.
            Else
                
            End If
            
            'Finally, if the image has not been resized to fit on screen, check its viewport to make sure the right and
            ' bottom edges don't fall outside the MDI client area
            'If the user wants us to resize the image to fit on-screen, do that now
            If g_AutozoomLargeImages = 1 Then FitWindowToViewport
                        
            'Add this file to the MRU list (unless specifically told not to)
            If ToUpdateMRU And (pageNumber = 0) And (MacroStatus <> MacroBATCH) Then g_RecentFiles.MRU_AddNewFile sFile(thisImage), targetImage
            
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
        If g_imageHasMultiplePages Then
            
            Dim pageTracker As Long
            
            Dim tmpStringArray(0) As String
            tmpStringArray(0) = sFile(thisImage)
            
            'Call PreLoadImage again for each individual frame in the multipage file
            For pageTracker = 1 To g_imagePageCount
                If UCase(GetExtension(sFile(thisImage))) = "GIF" Then
                    PreLoadImage tmpStringArray, False, targetImage.originalFileName & " (" & g_Language.TranslateMessage("frame") & " " & (pageTracker + 1) & ")." & GetExtension(sFile(thisImage)), targetImage.originalFileName & " (" & g_Language.TranslateMessage("frame") & " " & (pageTracker + 1) & ")." & GetExtension(sFile(thisImage)), , , , pageTracker
                ElseIf UCase(GetExtension(sFile(thisImage))) = "ICO" Then
                    PreLoadImage tmpStringArray, False, targetImage.originalFileName & " (" & g_Language.TranslateMessage("icon") & " " & (pageTracker + 1) & ")." & GetExtension(sFile(thisImage)), targetImage.originalFileName & " (" & g_Language.TranslateMessage("icon") & " " & (pageTracker + 1) & ")." & GetExtension(sFile(thisImage)), , , , pageTracker
                Else
                    PreLoadImage tmpStringArray, False, targetImage.originalFileName & " (" & g_Language.TranslateMessage("page") & " " & (pageTracker + 1) & ")." & GetExtension(sFile(thisImage)), targetImage.originalFileName & " (" & g_Language.TranslateMessage("page") & " " & (pageTracker + 1) & ")." & GetExtension(sFile(thisImage)), , , , pageTracker
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
Public Function LoadPhotoDemonImage(ByVal PDIPath As String, ByRef dstLayer As pdLayer, ByRef dstImage As pdImage) As Boolean
    
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
Public Function LoadGDIPlusImage(ByVal imagePath As String, ByRef dstLayer As pdLayer, ByRef dstImage As pdImage) As Boolean

    Dim tmpPicture As StdPicture
    Set tmpPicture = New StdPicture
            
    Dim verifyGDISuccess As Boolean
    
    verifyGDISuccess = GDIPlusLoadPicture(imagePath, tmpPicture, dstImage)
    
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
Public Function LoadVBImage(ByVal imagePath As String, ByRef dstLayer As pdLayer, ByRef dstImage As pdImage) As Boolean
    
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
Public Sub LoadUndo(ByVal undoFile As String, ByVal undoType As Long, Optional ByVal isRedoData As Boolean = False)

    'Several Undo Types are supported
    'Select Case undoType
    
        'Pixel data
        'Case 1
        
            'The layer handles the actual loading of the undo data
            pdImages(g_CurrentImage).mainLayer.createFromFile undoFile
            pdImages(g_CurrentImage).updateSize
            
        'Selection data
        'Case 2
        
            'Load the previous selection from file
            pdImages(g_CurrentImage).mainSelection.readSelectionFromFile undoFile & ".selection"
            
            'Activate the selection as necessary
            pdImages(g_CurrentImage).selectionActive = pdImages(g_CurrentImage).mainSelection.isLockedIn
            
            'Synchronize the text boxes as necessary
            syncTextToCurrentSelection g_CurrentImage
        
    'End Select
    
    'If a selection is active, request a redraw of the selection mask before rendering the image to the screen.  (If we are
    ' "undoing" an action that changed the image's size, the selection mask will be out of date.  Thus we need to re-render
    ' it before rendering the image or OOB errors may occur.)
    If pdImages(g_CurrentImage).selectionActive Then pdImages(g_CurrentImage).mainSelection.requestNewMask
        
    'Render the image to the screen
    PrepareViewport pdImages(g_CurrentImage).containingForm, "LoadUndo"
    
End Sub

'This routine sets the message on the splash screen (used only when the program is first started)
Public Sub LoadMessage(ByVal sMsg As String)

    Dim warnIDE As String
    warnIDE = "(IDE NOT RECOMMENDED - PLEASE COMPILE)"
    
    'Load messages are translatable, but we don't want to translate them if the translation object isn't ready yet
    If (Not (g_Language Is Nothing)) Then
        If g_Language.readyToTranslate Then
            If g_Language.translationActive Then
                sMsg = g_Language.TranslateMessage(sMsg)
                warnIDE = g_Language.TranslateMessage("(IDE NOT RECOMMENDED - PLEASE COMPILE)")
            End If
        End If
    End If
    
    If Not g_IsProgramCompiled Then sMsg = sMsg & "  " & warnIDE
    
    If FormSplash.Visible Then
        FormSplash.lblMessage = sMsg
        FormSplash.lblMessage.Refresh
    End If
    'DoEvents
    
End Sub

'Generates all shortcuts that VB can't; many thanks to Steve McMahon for his accelerator class, which helps a great deal
Public Sub LoadAccelerators()

    'Don't allow custom shortcuts in the IDE, as they require subclassing and might crash
    'If Not g_IsProgramCompiled Then Exit Sub

    With FormMain.ctlAccelerator
    
        'File menu
        .AddAccelerator vbKeyO, vbCtrlMask, "Open", FormMain.MnuFile(0), True, False, True, False
        .AddAccelerator vbKeyF4, vbCtrlMask, "Close", FormMain.MnuFile(4), True, True, True, False
        .AddAccelerator vbKeyF4, vbCtrlMask Or vbShiftMask, "Close all", FormMain.MnuFile(5), True, True, True, False
        .AddAccelerator vbKeyS, vbCtrlMask, "Save", FormMain.MnuFile(7), True, True, True, False
        .AddAccelerator vbKeyS, vbCtrlMask Or vbShiftMask, "Save as", FormMain.MnuFile(8), True, True, True, False
        .AddAccelerator vbKeyF12, 0, "Revert", FormMain.MnuFile(9), True, True, False, False
        .AddAccelerator vbKeyB, vbCtrlMask, "Batch wizard", FormMain.MnuFile(11), True, True, True
        .AddAccelerator vbKeyP, vbCtrlMask, "Print", FormMain.MnuFile(13), True, True, True
        .AddAccelerator vbKeyQ, vbCtrlMask, "Exit program", FormMain.MnuFile(15), True, False, True, False
        
            'File -> Import submenu
            .AddAccelerator vbKeyI, vbCtrlMask Or vbShiftMask Or vbAltMask, "Scan image", FormMain.MnuScanImage, True, False, True, False
            .AddAccelerator vbKeyD, vbCtrlMask Or vbShiftMask, "Internet import", FormMain.MnuImportFromInternet, True, True, True, False
            .AddAccelerator vbKeyI, vbCtrlMask Or vbAltMask, "Screen capture", FormMain.MnuScreenCapture, True, True, True, False
        
            'Most-recently used files.  Note that we cannot automatically associate these with a menu, as these menus may not
            ' exist at run-time.  (They are dynamically created as necessary.)
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
        .AddAccelerator vbKeyZ, vbCtrlMask, "Undo", FormMain.MnuEdit(0), True, True, False, False
        .AddAccelerator vbKeyY, vbCtrlMask, "Redo", FormMain.MnuEdit(1), True, True, False, False
        .AddAccelerator vbKeyF, vbCtrlMask, "Repeat last action", FormMain.MnuEdit(2), True, True, False, True
        
        .AddAccelerator vbKeyC, vbCtrlMask, "Copy to clipboard", FormMain.MnuEdit(4), True, True, False, False
        .AddAccelerator vbKeyV, vbCtrlMask, "Paste as new image", FormMain.MnuEdit(5), True, False, False, False
        
        'View menu
        .AddAccelerator vbKey0, 0, "FitOnScreen", FormMain.MnuFitOnScreen, False, True, False, False
        .AddAccelerator vbKeyAdd, 0, "Zoom_In", FormMain.MnuZoomIn, False, True, False, False
        .AddAccelerator vbKeySubtract, 0, "Zoom_Out", FormMain.MnuZoomOut, False, True, False, False
        .AddAccelerator vbKey5, 0, "Zoom_161", FormMain.MnuSpecificZoom(0), False, True, False, False
        .AddAccelerator vbKey4, 0, "Zoom_81", FormMain.MnuSpecificZoom(1), False, True, False, False
        .AddAccelerator vbKey3, 0, "Zoom_41", FormMain.MnuSpecificZoom(2), False, True, False, False
        .AddAccelerator vbKey2, 0, "Zoom_21", FormMain.MnuSpecificZoom(3), False, True, False, False
        .AddAccelerator vbKey1, 0, "Actual_Size", FormMain.MnuSpecificZoom(4), False, True, False, False
        .AddAccelerator vbKey2, vbShiftMask, "Zoom_12", FormMain.MnuSpecificZoom(5), False, True, False, False
        .AddAccelerator vbKey3, vbShiftMask, "Zoom_14", FormMain.MnuSpecificZoom(6), False, True, False, False
        .AddAccelerator vbKey4, vbShiftMask, "Zoom_18", FormMain.MnuSpecificZoom(7), False, True, False, False
        .AddAccelerator vbKey5, vbShiftMask, "Zoom_116", FormMain.MnuSpecificZoom(8), False, True, False, False
        
        'Image menu
        .AddAccelerator vbKeyA, vbCtrlMask Or vbShiftMask, "Duplicate image", FormMain.MnuImage(0), True, True, False, 0
        .AddAccelerator vbKeyR, vbCtrlMask, "Resize", FormMain.MnuImage(4), True, True, True, 0
        .AddAccelerator vbKeyR, vbCtrlMask Or vbAltMask, "Canvas resize", FormMain.MnuImage(5), True, True, True, 0
        .AddAccelerator vbKeyX, vbCtrlMask Or vbShiftMask, "Crop", FormMain.MnuImage(7), True, True, False, 1
        .AddAccelerator vbKeyX, vbCtrlMask Or vbAltMask, "Autocrop", FormMain.MnuImage(8), True, True, False, 1
        'KeyCode 188 = <,  (next to the letter M)
        .AddAccelerator 188, vbCtrlMask Or vbAltMask, "Reduce colors", FormMain.MnuImage(15), True, True, False, 0
        
            'Image -> Rotate submenu
            .AddAccelerator vbKeyR, 0, "Rotate 90° clockwise", FormMain.MnuRotate(0), True, True, False, 1
            .AddAccelerator vbKeyL, 0, "Rotate 90° counter-clockwise", FormMain.MnuRotate(1), True, True, False, 1
            .AddAccelerator vbKeyR, vbCtrlMask Or vbShiftMask Or vbAltMask, "Arbitrary rotation", FormMain.MnuRotate(3), True, True, True, False
        
        'Select Menu
        .AddAccelerator vbKeyA, vbCtrlMask, "Select all", FormMain.MnuSelect(0), True, True, False, 2
        .AddAccelerator vbKeyD, vbCtrlMask, "Remove selection", FormMain.MnuSelect(1), False, True, False, 2
        .AddAccelerator vbKeyI, vbCtrlMask Or vbShiftMask, "Invert selection", FormMain.MnuSelect(2), True, True, False, 2
        'KeyCode 219 = {[  (next to the letter P), 221 = }]
        .AddAccelerator 221, vbCtrlMask Or vbAltMask, "Grow selection", FormMain.MnuSelect(4), True, True, True, False
        .AddAccelerator 219, vbCtrlMask Or vbAltMask, "Shrink selection", FormMain.MnuSelect(5), True, True, True, False
        .AddAccelerator vbKeyD, vbCtrlMask Or vbAltMask, "Feather selection", FormMain.MnuSelect(7), True, True, True, False
        
        'Adjustments Menu
        
        'Adjustments top shortcut menu
        .AddAccelerator vbKeyU, vbCtrlMask Or vbShiftMask, "Black and white", FormMain.MnuAdjustments(0), True, True, True, False
        .AddAccelerator vbKeyB, vbCtrlMask Or vbShiftMask, "Brightness and contrast", FormMain.MnuAdjustments(1), True, True, True, False
        .AddAccelerator vbKeyC, vbCtrlMask Or vbShiftMask, "Color balance", FormMain.MnuAdjustments(2), True, True, True, False
        .AddAccelerator vbKeyM, vbCtrlMask, "Curves", FormMain.MnuAdjustments(3), True, True, True, False
        .AddAccelerator vbKeyL, vbCtrlMask, "Levels", FormMain.MnuAdjustments(4), True, True, True, False
        .AddAccelerator vbKeyAdd, vbCtrlMask Or vbAltMask, "Vibrance", FormMain.MnuAdjustments(5), True, True, True, False
        .AddAccelerator vbKeyW, vbCtrlMask, "White balance", FormMain.MnuAdjustments(6), True, True, True, False
        
            'Color adjustments
            .AddAccelerator vbKeyH, vbCtrlMask, "Hue and saturation", FormMain.MnuColor(3), True, True, True, False
            .AddAccelerator vbKeyP, vbCtrlMask Or vbAltMask, "Photo filter", FormMain.MnuColor(4), True, True, True, False
            
            'Lighting adjustments
            .AddAccelerator vbKeyE, vbCtrlMask Or vbAltMask, "Exposure", FormMain.MnuLighting(2), True, True, True, False
            .AddAccelerator vbKeyG, vbCtrlMask, "Gamma", FormMain.MnuLighting(3), True, True, True, False
            .AddAccelerator vbKeyH, vbCtrlMask Or vbShiftMask, "Shadows and highlights", FormMain.MnuLighting(5), True, True, True, False
            .AddAccelerator vbKeyT, vbCtrlMask, "Temperature", FormMain.MnuLighting(6), True, True, True, False
            
            'Adjustments -> Invert submenu
            .AddAccelerator vbKeyI, vbCtrlMask, "Invert RGB", FormMain.mnuInvert, True, True, False, 1
            
            'Adjustments -> Monochrome submenu
            .AddAccelerator vbKeyB, vbCtrlMask Or vbAltMask Or vbShiftMask, "Color to monochrome", FormMain.MnuMonochrome(0), True, True, True, False
        
        'Effects Menu
        '.AddAccelerator vbKeyZ, vbCtrlMask Or vbAltMask Or vbShiftMask, "Add RGB noise", FormMain.MnuNoise(1), True, True, True, False
        '.AddAccelerator vbKeyG, vbCtrlMask Or vbAltMask Or vbShiftMask, "Gaussian blur", FormMain.MnuBlurFilter(1), True, True, True, False
        '.AddAccelerator vbKeyY, vbCtrlMask Or vbAltMask Or vbShiftMask, "Correct lens distortion", FormMain.MnuDistortEffects(1), True, True, True, False
        '.AddAccelerator vbKeyU, vbCtrlMask Or vbAltMask Or vbShiftMask, "Unsharp mask", FormMain.MnuSharpen(1), True, True, True, False
        
        'Tools menu
        .AddAccelerator vbKeyReturn, vbAltMask, "Preferences", FormMain.mnuTool(5), False, False, True, False
        .AddAccelerator vbKeyM, vbCtrlMask Or vbAltMask, "Plugin manager", FormMain.mnuTool(6), False, False, True, False
        'KeyCode 190 = >.  (two over from the letter M)
        .AddAccelerator 190, vbCtrlMask Or vbAltMask, "Play macro", FormMain.MnuPlayMacroRecording, True, True, True, False
        
        'Window menu
        .AddAccelerator vbKeyPageDown, 0, "Next_Image", FormMain.MnuWindow(6), False, True, False, False
        .AddAccelerator vbKeyPageUp, 0, "Prev_Image", FormMain.MnuWindow(7), False, True, False, False
                
        'No equivalent menu
        .AddAccelerator vbKeyEscape, 0, "Escape"
        
        .Enabled = True
    End With

    DrawAccelerators
    
End Sub

'After all menu shortcuts (accelerators) are loaded above, the custom shortcuts need to be added to the menu entries themselves.
' If we don't do this, the user won't know how to trigger the shortcuts!
Public Sub DrawAccelerators()

    Dim i As Long
    
    For i = 1 To FormMain.ctlAccelerator.Count
        With FormMain.ctlAccelerator
            If .hasMenu(i) Then
                .associatedMenu(i).Caption = .associatedMenu(i).Caption & vbTab & .stringRep(i)
            End If
        End With
    Next i

    'A few menu shortcuts must be drawn manually.
    
    'Because the Import -> From Clipboard menu shares the same shortcut as Edit -> Paste, we must manually add
    ' its shortcut (as only the Edit -> Paste will be automatically handled).
    FormMain.MnuImportClipboard.Caption = FormMain.MnuImportClipboard.Caption & vbTab & "Ctrl+V"
    
    'NOTE: Drawing of MRU shortcuts is handled in the MRU module
    
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
        If g_UserPreferences.GetPref_Boolean("Plugins", "Force EZTwain Disable", False) Then
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
        If g_UserPreferences.GetPref_Boolean("Plugins", "Force ZLib Disable", False) Then
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
        If g_UserPreferences.GetPref_Boolean("Plugins", "Force FreeImage Disable", False) Then
            g_ImageFormats.FreeImageEnabled = False
        Else
            g_ImageFormats.FreeImageEnabled = True
        End If
        
    Else
        g_ImageFormats.FreeImageEnabled = False
    End If
    
        'Additionally related to FreeImage - enable/disable the arbitrary rotation option contingent on FreeImage's enabling
        FormMain.MnuRotate(3).Visible = g_ImageFormats.FreeImageEnabled
    
    'Check for pngnq interface
    If isPngnqAvailable Then
        
        'Check to see if pngnq-s9 has been forcibly disabled
        If g_UserPreferences.GetPref_Boolean("Plugins", "Force Pngnq Disable", False) Then
            g_ImageFormats.pngnqEnabled = False
        Else
            g_ImageFormats.pngnqEnabled = True
        End If
        
    Else
        g_ImageFormats.pngnqEnabled = False
    End If
    
    'Check for ExifTool metadata interface
    If isExifToolAvailable Then
        
        'Check to see if ExifTool has been forcibly disabled
        If g_UserPreferences.GetPref_Boolean("Plugins", "Force ExifTool Disable", False) Then
            g_ExifToolEnabled = False
        Else
            g_ExifToolEnabled = True
        End If
        
    Else
        g_ExifToolEnabled = False
    End If
    
    'Finally, check GDI+ availability
    If g_GDIPlusAvailable Then
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
    imageToBeDuplicated = g_CurrentImage
    
    CreateNewImageForm
        
    g_AllowViewportRendering = False
        
    pdImages(g_CurrentImage).containingForm.HScroll.value = 0
    pdImages(g_CurrentImage).containingForm.VScroll.value = 0
        
    'Copy the picture from the previous form to this new one
    pdImages(g_CurrentImage).mainLayer.createFromExistingLayer pdImages(imageToBeDuplicated).mainLayer

    'Store important data about the image to the pdImages array
    pdImages(g_CurrentImage).updateSize
    pdImages(g_CurrentImage).originalFileSize = pdImages(imageToBeDuplicated).originalFileSize
    pdImages(g_CurrentImage).locationOnDisk = ""
            
    'Get the original file's extension and filename, then append " - Copy" to it
    Dim originalExtension As String
    originalExtension = GetExtension(pdImages(imageToBeDuplicated).originalFileNameAndExtension)
            
    Dim newFilename As String
    newFilename = pdImages(imageToBeDuplicated).originalFileName & " - " & g_Language.TranslateMessage("Copy")
    pdImages(g_CurrentImage).originalFileName = newFilename
    If Len(originalExtension) > 0 Then
        pdImages(g_CurrentImage).originalFileNameAndExtension = newFilename & "." & originalExtension
    Else
        pdImages(g_CurrentImage).originalFileNameAndExtension = newFilename
    End If
            
    'Because this image hasn't been saved to disk, mark its save state as "false"
    pdImages(g_CurrentImage).setSaveState False
    
    'Fit the window to the newly duplicated image
    Message "Resizing image to fit screen..."
    
    'Update the current caption to match
    g_WindowManager.requestWindowCaptionChange pdImages(g_CurrentImage).containingForm, pdImages(g_CurrentImage).originalFileNameAndExtension
        
    'Register this window with PhotoDemon's window manager.  This will do things like set the proper border state depending on whether
    ' image windows are docked or floating, which we need before doing things like auto-zoom or window placement.
    g_WindowManager.registerChildForm pdImages(g_CurrentImage).containingForm, IMAGE_WINDOW, , , g_CurrentImage
            
    'Also register this image with the image tab bar
    createCustomFormIcon pdImages(g_CurrentImage).containingForm
    toolbar_ImageTabs.registerNewImage g_CurrentImage
    
    'If the user wants us to resize the image to fit on-screen, do that now
    If g_AutozoomLargeImages = 0 Then FitImageToViewport True
                    
    'If image windows are not docked, fit the window around the (now properly-zoomed) image
    If g_WindowManager.getFloatState(IMAGE_WINDOW) Then FitWindowToImage True, True
            
    'g_AllowViewportRendering may have been reset by this point (by the FitImageToViewport sub, among others), so set it back to False, then
    ' update the zoom combo box to match the zoom assigned by the window-fit function.
    g_AllowViewportRendering = False
    toolbar_File.CmbZoom.ListIndex = pdImages(g_CurrentImage).currentZoomValue
        
    'Now that the image's window has been fully sized and moved around, use PrepareViewport to set up any scrollbars and a back-buffer
    g_AllowViewportRendering = True
    PrepareViewport pdImages(g_CurrentImage).containingForm, "Duplicate image"
            
    'Note the window state, as it may be important in the future
    pdImages(g_CurrentImage).WindowState = pdImages(g_CurrentImage).containingForm.WindowState
            
    'If image windows are floating, move the window into place now using values previously calculated by FitToScreen
    If g_WindowManager.getFloatState(IMAGE_WINDOW) Then
        pdImages(g_CurrentImage).containingForm.Move pdImages(g_CurrentImage).WindowLeft * Screen.TwipsPerPixelX, pdImages(g_CurrentImage).WindowTop * Screen.TwipsPerPixelY
        g_WindowManager.requestWindowResync pdImages(g_CurrentImage).indexInWindowManager
        
    'If image windows are docked, the window manager will have already positioned the window for us.
    Else
            
    End If
            
    'Finally, if the image has not been resized to fit on screen, check its viewport to make sure the right and
    ' bottom edges don't fall outside the MDI client area
    'If the user wants us to resize the image to fit on-screen, do that now
    If g_AutozoomLargeImages = 1 Then FitWindowToViewport
    
    'Synchronize the interface to match the newly created image's settings
    syncInterfaceToCurrentImage
    
    Message "Image duplication complete."
    
    'If we made it all the way here, the image was successfully duplicated.
    pdImages(g_CurrentImage).loadedSuccessfully = True
        
End Sub

'Check for IDE or compiled EXE, and set program parameters accordingly
Private Sub CheckLoadingEnvironment()
    If App.logMode = 1 Then
        g_IsProgramCompiled = True
    Else
        g_IsProgramCompiled = False
    End If
End Sub
