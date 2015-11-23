Attribute VB_Name = "Loading"
'***************************************************************************
'Program/File Loading Handler
'Copyright 2001-2015 by Tanner Helland
'Created: 4/15/01
'Last updated: 28/April/15
'Last update: thanks to the new pdGlyphCollection class, PD now caches a list of all fonts, not just TrueType ones.
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
Private m_LoadTime As Double, m_StartTime As Double

'This constant is the number of "discrete" loading steps involved in loading the program.  It is relevant for displaying a progress bar on the
' initial splash screen.
Private Const NUMBER_OF_LOADING_STEPS As Long = 14

'PHOTODEMON STARTS HERE (after Sub Main, that is).

'Note that this function is called AFTER FormMain has been loaded.  FormMain is loaded - but not visible - so it can be operated
' on by functions called from this routine.  (It is necessary to load the main window first, since a number of load operations -
' like decoding PNG menu icons from the resource file, then applying them to program menus - operate directly on the main window.)
Public Sub LoadTheProgram()
    
    '*************************************************************************************************************************************
    ' Check the state of this build (alpha, beta, production, etc) and activate debug code as necessary
    '*************************************************************************************************************************************
    
    'Current build state is stored in the public const "PD_BUILD_QUALITY".  For non-production builds, a number of program-wide
    ' parameters are automatically set.
    
    'If the program is in pre-alpha or alpha state, enable timing reports.
    If (PD_BUILD_QUALITY = PD_PRE_ALPHA) Or (PD_BUILD_QUALITY = PD_ALPHA) Then g_DisplayTimingReports = True
    
    'Regardless of debug mode or not, we instantiate a pdDebug instance.  It will only be interacted with if the program is compiled
    ' with DEBUGMODE = 1, however.
    Set pdDebug = New pdDebugger
    
    'During development, I find it helpful to profile PhotoDemon's startup process.  Timing functions like this can be commented out
    ' without harming anything.
    Dim perfCheck As pdProfiler
    Set perfCheck = New pdProfiler
    
    #If DEBUGMODE = 1 Then
        perfCheck.startProfiling "PhotoDemon Startup", True
    #End If
    
    
    
    '*************************************************************************************************************************************
    ' With the debugger initialized, prep a few crucial variables
    '*************************************************************************************************************************************
    
    'Most importantly, we need to create a default pdImages() array, as some initialization functions may attempt to access that array
    ReDim pdImages(0 To 3) As pdImage
    
    
    
    '*************************************************************************************************************************************
    ' Prepare the splash screen (but don't display it yet)
    '*************************************************************************************************************************************
    
    #If DEBUGMODE = 1 Then
        perfCheck.markEvent "Prepare splash screen"
    #End If
    
    m_StartTime = Timer
    
    'We need GDI+ to extract a JPEG from the resource file and convert it in-memory.  (Yes, there are other ways to do this.  No, I don't
    ' care about using them.)  Check its availability.
    If isGDIPlusAvailable() Then
    
        'Load FormSplash into memory, but don't make it visible.
        FormSplash.Visible = False
        
    End If
        
    'Check the environment.  If inside the the IDE, the splash needs to be modified slightly.
    CheckLoadingEnvironment
    
    If g_GDIPlusAvailable Then
        If g_IsProgramCompiled Then m_LoadTime = 1# Else m_LoadTime = 0.5
    Else
        m_LoadTime = 0#
    End If
    
    'Retrieve a Unicode-friendly copy of any command line parameters, and store them publicly
    Dim cUnicode As pdUnicode
    Set cUnicode = New pdUnicode
    g_CommandLine = cUnicode.CommandW()
    
    
    '*************************************************************************************************************************************
    ' Determine which version of Windows the user is running (as other load functions rely on this)
    '*************************************************************************************************************************************
    
    #If DEBUGMODE = 1 Then
        perfCheck.markEvent "Check Windows version"
    #End If
    
    LoadMessage "Detecting Windows® version..."
    
    'Certain features are OS-version dependent.  We must determine the OS version early in the load process to ensure that all features
    ' are loaded correctly.
    Dim cSysInfo As pdSystemInfo
    Set cSysInfo = New pdSystemInfo
    
    g_IsVistaOrLater = cSysInfo.IsOSVistaOrLater
    g_IsWin7OrLater = cSysInfo.IsOSWin7OrLater
    g_IsWin8OrLater = cSysInfo.IsOSWin8OrLater
    g_IsWin81OrLater = cSysInfo.IsOSWin81OrLater
    g_IsWin10OrLater = cSysInfo.IsOSWin10OrLater
    
    'If we are on Windows 7, prepare some Win7-specific features (like taskbar progress bars)
    If g_IsWin7OrLater Then prepWin7Features
    
    
    
    '*************************************************************************************************************************************
    ' If the user doesn't have font smoothing enabled, enable it now.  PD's interface looks much better with some form of antialiasing.
    '*************************************************************************************************************************************
    
    #If DEBUGMODE = 1 Then
        perfCheck.markEvent "ClearType check"
    #End If
    
    HandleClearType True
    
    
    
    '*************************************************************************************************************************************
    ' Initialize the user preferences (settings) handler
    '*************************************************************************************************************************************
    
    #If DEBUGMODE = 1 Then
        perfCheck.markEvent "Initialize preferences engine"
    #End If
    
    'Before initializing the preference engine, generate a unique session ID for this PhotoDemo instance.  This ID will be used to
    ' separate the temp files for this program instance from any other simultaneous instances.
    g_SessionID = cSysInfo.GetUniqueSessionID()
    
    Set g_UserPreferences = New pdPreferences
    
    'Ask the preferences handler to generate key program folders.  (If these folders don't exist, the handler will create them.)
    LoadMessage "Initializing all program directories..."
    
    g_UserPreferences.initializePaths
        
    'Now, ask the preferences handler to load all other user settings from the preferences file
    LoadMessage "Loading all user settings..."
    
    g_UserPreferences.loadUserSettings
        
    'Mark the Macro recorder as "not recording"
    MacroStatus = MacroSTOP
    
    'Note that no images have been loaded yet
    g_NumOfImagesLoaded = 0
    
    'Set the default active image index to 0
    g_CurrentImage = 0
    
    'Set the number of open image windows to 0
    g_OpenImageCount = 0
    
    'While here, also initialize the image format handler (as plugins and other load functions interact with it)
    Set g_ImageFormats = New pdFormats
    
    
    
    '*************************************************************************************************************************************
    ' Initialize the translation (language) engine
    '*************************************************************************************************************************************
    
    #If DEBUGMODE = 1 Then
        perfCheck.markEvent "Initialize translation engine"
    #End If
    
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
    g_Language.ApplyLanguage False
    
    
    
    
    '*************************************************************************************************************************************
    ' PhotoDemon works very well with multiple monitors.  Check for such a situation now.
    '*************************************************************************************************************************************
    
    #If DEBUGMODE = 1 Then
        perfCheck.markEvent "Detect monitors"
    #End If
    
    LoadMessage "Analyzing current monitor setup..."
    
    Set g_Displays = New pdDisplays
    g_Displays.RefreshDisplays
    
    'While here, also cache various display-related settings; this is faster than constantly retrieving them via APIs
    Color_Management.CacheCurrentSystemColorProfile
    Interface.CacheSystemDPI g_Displays.GetWindowsDPI
    
    
    '*************************************************************************************************************************************
    ' Now we have what we need to properly display the splash screen.  Do so now.
    '*************************************************************************************************************************************
    
    #If DEBUGMODE = 1 Then
        perfCheck.markEvent "Display splash screen"
    #End If
    
    'Determine the program's previous on-screen location.  We need that to determine where to display the splash screen.
    Dim wRect As RECTL
    With wRect
        .Left = g_UserPreferences.GetPref_Long("Core", "Last Window Left", 1)
        .Top = g_UserPreferences.GetPref_Long("Core", "Last Window Top", 1)
        .Right = .Left + g_UserPreferences.GetPref_Long("Core", "Last Window Width", 1)
        .Bottom = .Top + g_UserPreferences.GetPref_Long("Core", "Last Window Height", 1)
    End With
    
    'Center the splash screen on whichever monitor the user previously used.
    g_Displays.CenterFormViaReferenceRect FormSplash, wRect
    
    'If Segoe UI is available, we prefer to use it instead of Tahoma.  On XP this is not guaranteed, however, so we have to check.
    Dim tmpFontCheck As pdFont
    Set tmpFontCheck = New pdFont
    
    'If Segoe exists, we mark two variables: a String (which user controls use to create their own font objects), and a Boolean
    ' (which some dialogs use to slightly modify their layout for better alignments).
    If tmpFontCheck.DoesFontExist("Segoe UI") Then
        g_InterfaceFont = "Segoe UI"
        g_UseFancyFonts = True
    Else
        g_InterfaceFont = "Tahoma"
        g_UseFancyFonts = False
    End If
    
    Set tmpFontCheck = Nothing
        
    'Make the splash screen's message display match the rest of PD
    FormSplash.lblMessage.fontName = g_InterfaceFont
    
    'Ask the splash screen to finish whatever initializing it needs prior to displaying itself
    FormSplash.prepareSplashLogo NUMBER_OF_LOADING_STEPS
    FormSplash.prepareRestOfSplash
    
    'Display the splash screen, centered on whichever monitor the user previously used the program on.
    FormSplash.Show vbModeless
    
    
    
    '*************************************************************************************************************************************
    ' Check for the presence of plugins (as other functions rely on these to initialize themselves)
    '*************************************************************************************************************************************
    
    #If DEBUGMODE = 1 Then
        perfCheck.markEvent "Load plugins"
    #End If
    
    LoadMessage "Loading plugins..."
    
    Plugin_Management.LoadAllPlugins
    
    
    'If ExifTool was enabled successfully, ask it to double-check that its tag database has been created
    ' successfully at some point in the past.  If it hasn't, generate a new copy now.
    '
    'Now that this has been thoroughly tested, I'm postponing actual enabling of it until PD supports metadata editing.
    'If g_ExifToolEnabled Then writeTagDatabase
    
    
    
    '*************************************************************************************************************************************
    ' If this is not a production build, initialize PhotoDemon's central debugger
    '*************************************************************************************************************************************
    
    'We wait until after the translation and plugin engines are initialized; this allows us to report their information in the debug log
    #If DEBUGMODE = 1 Then
        pdDebug.InitializeDebugger True
    #End If
    
    
    
    '*************************************************************************************************************************************
    ' Based on available plugins, determine which image formats PhotoDemon can handle
    '*************************************************************************************************************************************
    
    #If DEBUGMODE = 1 Then
        perfCheck.markEvent "Load import and export libraries"
    #End If
    
    LoadMessage "Loading import/export libraries..."
    
    'The FreeImage.dll plugin provides most of PD's advanced image format support, but we can also fall back on GDI+.
    ' Prior to generating a list of supported formats, notify the image format class of GDI+ availability
    ' (which was determined earlier in this function, prior to loading the splash screen).
    g_ImageFormats.GDIPlusEnabled = g_GDIPlusAvailable
    
    'Generate a list of currently supported input/output formats, which may vary based on plugin version and availability
    g_ImageFormats.generateInputFormats
    g_ImageFormats.generateOutputFormats
    
    
    
    '*************************************************************************************************************************************
    ' Initialize the visual themes engine
    '*************************************************************************************************************************************
    
    #If DEBUGMODE = 1 Then
        perfCheck.markEvent "Initialize theme engine"
    #End If
    
    'Because this class controls the visual appearance of all forms in the project, it must be loaded early in the boot process
    LoadMessage "Initializing theme engine..."
    
    Set g_Themer = New pdVisualThemes
    
    
    
    '*************************************************************************************************************************************
    ' Build a font cache for this system
    '*************************************************************************************************************************************
    
    #If DEBUGMODE = 1 Then
        perfCheck.markEvent "Build font cache"
    #End If
    
    LoadMessage "Building font cache..."
        
    'PD currently builds two font caches:
    ' 1) A name-only list of all fonts currently installed.  This is used to populate font dropdown boxes.
    ' 2) An pdFont-based cache of the current UI font, at various requested sizes.  This cache spares individual controls from needing
    '     to do their own font management; instead, they can simply request a matching object from the Font_Management module.
    Font_Management.BuildFontCaches
    
    'Next, build a list of font properties, like supported scripts
    Font_Management.BuildFontCacheProperties
    
    
    
    '*************************************************************************************************************************************
    ' Initialize PD's central clipboard manager
    '*************************************************************************************************************************************
    
    #If DEBUGMODE = 1 Then
        perfCheck.markEvent "Initialize pdClipboardMain"
    #End If
    
    LoadMessage "Initializing clipboard interface..."
    
    Set g_Clipboard = New pdClipboardMain
    
    
    '*************************************************************************************************************************************
    ' Get the viewport engine ready
    '*************************************************************************************************************************************
    
    #If DEBUGMODE = 1 Then
        perfCheck.markEvent "Initialize viewport engine"
    #End If
    
    'Initialize our current zoom method
    LoadMessage "Initializing viewport engine..."
    
    'Create the program's primary zoom handler
    Set g_Zoom = New pdZoom
    g_Zoom.initializeViewportEngine
    
    'Populate the main form's zoom drop-down
    g_Zoom.populateZoomComboBox FormMain.mainCanvas(0).getZoomDropDownReference()
    
    'Populate the main canvas's size unit dropdown
    FormMain.mainCanvas(0).PopulateSizeUnits
    
    
    
    '*************************************************************************************************************************************
    ' Initialize the window manager (the class that synchronizes all toolbox and image window positions)
    '*************************************************************************************************************************************
    
    #If DEBUGMODE = 1 Then
        perfCheck.markEvent "Initialize window manager"
    #End If
    
    LoadMessage "Initializing window manager..."
    Set g_WindowManager = New pdWindowManager
    
    'Register the main form
    g_WindowManager.RegisterParentForm FormMain
    
    'Load all tool windows.  Even though they may not be visible (as the user can elect to hide them), we still want them loaded,
    ' so we can interact with them as necessary (e.g. "enable Undo button", etc).
    Load toolbar_Toolbox
    Load toolbar_ImageTabs
    Load toolbar_Options
    
    'Retrieve tool window visibility and mark those menus as well
    FormMain.MnuWindowToolbox(0).Checked = g_UserPreferences.GetPref_Boolean("Core", "Show File Toolbox", True)
    FormMain.MnuWindow(1).Checked = g_UserPreferences.GetPref_Boolean("Core", "Show Selections Toolbox", True)
    FormMain.MnuWindow(2).Checked = g_UserPreferences.GetPref_Boolean("Core", "Show Layers Toolbox", True)
    
    #If DEBUGMODE = 1 Then
        FormMain.MnuDevelopers(0).Checked = g_UserPreferences.GetPref_Boolean("Core", "Show Debug Window", False)
    #End If
    
    'Retrieve two additional settings for the image tabstrip menu: when to display the image tabstrip...
    ToggleImageTabstripVisibility g_UserPreferences.GetPref_Long("Core", "Image Tabstrip Visibility", 1), True, True
    
    '...and the alignment of the tabstrip
    ToggleImageTabstripAlignment g_UserPreferences.GetPref_Long("Core", "Image Tabstrip Alignment", vbAlignTop), True, True
    
    'The primary toolbox has some options of its own.  Load them now.
    FormMain.MnuWindowToolbox(2).Checked = g_UserPreferences.GetPref_Boolean("Core", "Show Toolbox Category Labels", True)
    toolbar_Toolbox.updateButtonSize g_UserPreferences.GetPref_Long("Core", "Toolbox Button Size", 1), True
    
    
    
    '*************************************************************************************************************************************
    ' Set all default tool values
    '*************************************************************************************************************************************
    
    #If DEBUGMODE = 1 Then
        perfCheck.markEvent "Initialize tools"
    #End If
    
    LoadMessage "Initializing image tools..."
    
    'As of May 2015, tool panels are now loaded on-demand.  This improves the program's startup performance, and it saves a bit of memory
    ' if a user doesn't use a tool during a given session.
    
    'Also, while here, prep the specialized non-destructive tool handler in the central processor
    Processor.InitializeProcessor
    
    
    '*************************************************************************************************************************************
    ' PhotoDemon's complex interface requires a lot of things to be generated at run-time.
    '*************************************************************************************************************************************
    
    #If DEBUGMODE = 1 Then
        perfCheck.markEvent "Initialize UI"
    #End If
    
    LoadMessage "Initializing user interface..."
    
    'Use the API to give PhotoDemon's main form a 32-bit icon (VB is too old to support 32bpp icons)
    SetIcon FormMain.hWnd, "AAA", True
    
    'Initialize all system cursors we rely on (hand, busy, resizing, etc)
    initAllCursors
    
    'Set up the program's title bar.  Odd-numbered releases are development releases.  Even-numbered releases are formal builds.
    If Not (g_WindowManager Is Nothing) Then
        g_WindowManager.SetWindowCaptionW FormMain.hWnd, getPhotoDemonNameAndVersion()
    Else
        FormMain.Caption = getPhotoDemonNameAndVersion()
    End If
    
    'PhotoDemon renders many of its own icons dynamically.  Initialize that engine now.
    initializeIconHandler
    
    'Prepare a checkerboard pattern, which will be used behind any transparent objects.  Caching this is much more efficient.
    ' than re-creating it every time it's needed.
    Set g_CheckerboardPattern = New pdDIB
    Drawing.createAlphaCheckerboardDIB g_CheckerboardPattern
    
    'Allow drag-and-drop operations
    g_AllowDragAndDrop = True
    
    'Set the main canvas background color
    FormMain.mainCanvas(0).BackColor = g_CanvasBackground
    
    'Clear the main canvas coordinate and size displays
    FormMain.mainCanvas(0).displayCanvasCoordinates 0, 0, True
    FormMain.mainCanvas(0).displayImageSize Nothing, True
    
    'Throughout the program, g_MouseAccuracy is used to determine how close the mouse cursor must be to a point of interest to
    ' consider it "over" that point.  DPI must be accounted for when calculating this value (as it's calculated in pixels).
    g_MouseAccuracy = FixDPIFloat(6)
    
    'Apply visual styles
    FormMain.UpdateAgainstCurrentTheme False
    
    
    
    '*************************************************************************************************************************************
    ' The program's menus support many features that VB can't do natively (like icons and custom shortcuts).  Load such things now.
    '*************************************************************************************************************************************
    
    #If DEBUGMODE = 1 Then
        perfCheck.markEvent "Prep menus"
    #End If
    
    LoadMessage "Preparing program menus..."
    
    'In debug modes, certain developer and experimental options can be enabled.
    If (PD_BUILD_QUALITY <> PD_PRODUCTION) And (PD_BUILD_QUALITY <> PD_BETA) Then
        FormMain.MnuTest.Visible = True
        FormMain.mnuTool(9).Visible = True
        FormMain.mnuTool(10).Visible = True
    Else
        FormMain.MnuTest.Visible = False
        FormMain.mnuTool(9).Visible = False
        FormMain.mnuTool(10).Visible = False
    End If
        
    'Create all manual shortcuts (ones VB isn't capable of generating itself)
    LoadAccelerators
            
    'Initialize the Recent Files manager and load the most-recently-used file list (MRU)
    ' CHANGING: Using pdMRUManager instead of pdRecentFiles
    Set g_RecentFiles = New pdMRUManager
    g_RecentFiles.InitList New pdMRURecentFiles
    g_RecentFiles.MRU_LoadFromFile
    
    Set g_RecentMacros = New pdMRUManager
    g_RecentMacros.InitList New pdMRURecentMacros
    g_RecentMacros.MRU_LoadFromFile
            
    'Load and draw all menu icons
    loadMenuIcons
    
    'Synchronize all other interface elements to match the current program state (e.g. no images loaded).
    SyncInterfaceToCurrentImage
    
    
    
    '*************************************************************************************************************************************
    ' Unload the splash screen and present the main form
    '*************************************************************************************************************************************
    
    'While in debug mode, copy a timing report of program startup to the debug folder
    #If DEBUGMODE = 1 Then
        perfCheck.stopProfiling
        perfCheck.generateProfileReport True
    #End If
    
    'Display the splash screen for at least a second or two
    If Timer - m_StartTime < m_LoadTime Then
        Do While Timer - m_StartTime < m_LoadTime
        Loop
    End If
    
    'If this is the first time the user has run PhotoDemon, resize the window a bit to make the default position nice.
    ' (If this is *not* the first time, the window manager will automatically restore the window's last-known position and state.)
    If g_IsFirstRun Then g_WindowManager.SetFirstRunMainWindowPosition
    
    'In debug mode, make a baseline memory reading here, before the main form is displayed.
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "LoadTheProgram() function complete.  Baseline memory reading:"
        pdDebug.LogAction "", PDM_MEM_REPORT
        pdDebug.LogAction "Proceeding to load main window..."
    #End If
    
    'Display the main form
    'FormMain.Show
    
    Unload FormSplash
        
End Sub

'If files are present in the command line, this sub will load them
Public Sub LoadImagesFromCommandLine()

    Message "Loading image(s)..."
        
    'NOTE: Windows will pass multiple filenames via the command line, but it does so in a confusing and overly complex way.
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
    LoadFileAsNewImage sFile

End Sub

'Loading an image begins here.  This routine examines a given file's extension and re-routes control based on that.
Public Sub LoadFileAsNewImage(ByRef sFile() As String, Optional ByVal ToUpdateMRU As Boolean = True, Optional ByVal imgFormTitle As String = "", Optional ByVal imgName As String = "", Optional ByVal isThisPrimaryImage As Boolean = True, Optional ByRef targetImage As pdImage, Optional ByRef targetDIB As pdDIB, Optional ByVal pageNumber As Long = 0, Optional ByVal fillDIBWithCompositePDI As Boolean = False, Optional ByVal suspendWarnings As Boolean = False)
    
    'NOTE ABOUT DOEVENTS:
    ' Normally, PD avoids DoEvents for all the obvious reasons.  You'll notice that this function, however, uses DoEvents liberally.  Why?
    ' While this function is busy loading the image in question, the ExifTool plugin is running asynchronously, parsing image metadata
    ' and forwarding it to the main form's ShellPipe control as it proceeds.  By using DoEvents throughout this function, we yield control
    ' to that ShellPipe control, allowing it to periodically clear stdout so ExifTool can continue pushing metadata through.  That said,
    ' a LOT of precautions are taken to make sure DoEvents doesn't cause reentry and other issues, so don't try to mimic this behavior
    ' in your own software unless you understand the many repercussions!
    
    'If debug mode is active, image loading is a place where a lot of things can go wrong - bad files, corrupt formats, heavy RAM usage,
    ' incompatible color formats, and about a bazillion other things.  Make a special note in the debug log, to help narrow down issues.
    #If DEBUGMODE = 1 Then
        Dim startTime As Double
        startTime = Timer
    
        pdDebug.LogAction "Preparing to load one or more images.  Baseline memory reading:"
        pdDebug.LogAction "", PDM_MEM_REPORT
    #End If
    
    '*************************************************************************************************************************************
    ' Prepare all variables related to image loading
    '*************************************************************************************************************************************
    
    'Display a busy cursor
    If Screen.MousePointer <> vbHourglass Then Screen.MousePointer = vbHourglass
    
    'Additional file interactions are handled via pdFSO
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    'One of the things we'll be doing in this routine is establishing an original color depth for this image. FreeImage and GDI+ will
    ' return this automatically; VB's LoadPicture will not.  We use a tracking variable to determine if a manual color count needs to
    ' be performed.
    Dim mustCountColors As Boolean
    Dim colorCountCheck As Long
    
    'File extension determines how an image is loaded; certain formats require separate plugins (e.g. FreeImage)
    Dim FileExtension As String
    Dim loadSuccessful As Boolean
    
    Dim loadedByOtherMeans As Boolean
    loadedByOtherMeans = False
    
    'Individual image files might contain multiple layers.  If such an image is found, this will be set to TRUE.
    Dim imageHasMultiplePages As Boolean
    Dim numOfPages As Long
    
    'If multiple files are being loaded, we want to suppress all warnings and errors until the very end.
    Dim multipleFilesLoading As Boolean
    If UBound(sFile) > 0 Then multipleFilesLoading = True Else multipleFilesLoading = False
    
    Dim missingFiles As String
    missingFiles = ""
    
    Dim brokenFiles As String
    brokenFiles = ""
    
    'Some layers may receive extra information in their name.  (For example, when loading .ICO files with multiple icons inside,
    ' PD will automatically add the name and original bit-depth to each layer, as relevant.)
    Dim layerNameBase As String, layerNameAddon As String
    
    'Some behavior varies based on the image decoding engine used.  PD uses a fairly complex cascading system for image decoders;
    ' if one fails, we continue trying alternates until either the load succeeds, or all known decoders have been exhausted.
    Dim decoderUsed As PD_IMAGE_DECODER_ENGINE
            
    
    '*************************************************************************************************************************************
    'Before actually loading anything, perform a one-time check to make sure the metadata engine isn't still busy
    ' processing an initial database build.
    '*************************************************************************************************************************************
    
    If g_ExifToolEnabled And isDatabaseModeActive Then
        
        'Wait for metadata parsing to finish...
        If Not isMetadataFinished Then
        
            Message "Finishing final program initialization steps..."
        
            'Forcibly disable the main form to avoid DoEvents allowing click-through
            FormMain.Enabled = False
        
            'Pause for 1/10 second
            Do
                PauseProgram 0.1
                
                'If the user shuts down the program while we are still waiting for input, exit immediately
                If g_ProgramShuttingDown Then Exit Sub
                
            Loop While (Not isMetadataFinished)
            
            'Re-enable the main form
            FormMain.Enabled = True
            
        End If
        
    End If
        
        
    '*************************************************************************************************************************************
    ' To prevent re-entry problems, forcibly disable the main form
    '*************************************************************************************************************************************
    
    FormMain.Enabled = False
    
    
            
    '*************************************************************************************************************************************
    ' Loop through each entry in the sFile() array, loading images as we go
    '*************************************************************************************************************************************
            
    'Because this routine accepts an array of images, we have to be prepared for the possibility that more than
    ' one image file is being opened.  This loop will execute until all files are loaded.  If a file fails to
    ' load, it will automatically move on to the next one, and an error message will be displayed after all
    ' files have been processed.
    Dim thisImage As Long
    
    For thisImage = 0 To UBound(sFile)
    
        'If debug mode is active, post some helpful debugging information
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "Image load requested for """ & GetFilename(sFile(thisImage)) & """"
        #End If
    
        '*************************************************************************************************************************************
        ' Reset all variables used on a per-image level
        '*************************************************************************************************************************************
    
        'Reset the multipage checker (which is now handled on a per-image basis)
        imageHasMultiplePages = False
        numOfPages = 0
        
        '...and reset the "need to check colors" variable.  If FreeImage or GDI+ is used, color depth of the source file is retrieved
        ' automatically.  If another source is used, we manually calculate a bit-depth for incoming images.
        mustCountColors = False
        
    
        '*************************************************************************************************************************************
        ' Before attempting to load this image, make sure it exists
        '*************************************************************************************************************************************
    
        'If isThisPrimaryImage Then Message "Verifying that file exists..."
    
        If isThisPrimaryImage And (Not cFile.FileExist(sFile(thisImage))) Then
            
            'If multiple files are being loaded, suppress any errors until the end
            If multipleFilesLoading Then
                missingFiles = missingFiles & GetFilename(sFile(thisImage)) & vbCrLf
            Else
                If Not suspendWarnings Then
                    PDMsgBox "Unfortunately, the image '%1' could not be found." & vbCrLf & vbCrLf & "If this image was originally located on removable media (DVD, USB drive, etc), please re-insert or re-attach the media and try again.", vbApplicationModal + vbExclamation + vbOKOnly, "File not found", sFile(thisImage)
                End If
            End If
            
            'If the missing image was part of a list of images, try loading the next entry in the list
            GoTo PreloadMoreImages
            
        End If
        
        
        
        '*************************************************************************************************************************************
        ' If the image being loaded is a primary image (e.g. one opened normally), prepare a blank pdImage object to receive it
        '*************************************************************************************************************************************
        
        If isThisPrimaryImage Then
            
            If UBound(sFile) > 0 Then
                Message "Loading image %1 of %2...", thisImage + 1, UBound(sFile) + 1
            Else
                Message "Loading image..."
            End If
            
            CreateNewPDImage
            
            'If this is a primary image, we will automatically set the targetImage and targetDIB parameters.  If this is NOT a primary image,
            ' the calling function must have specified this for us.
            Set targetImage = pdImages(g_CurrentImage)
            
            'Create a blank layer in the receiving image, and retrieve a pointer to it
            Dim newLayerID As Long
            newLayerID = pdImages(g_CurrentImage).createBlankLayer
            
            Set targetDIB = New pdDIB
            
            g_AllowViewportRendering = False
            
            'Reset the main viewport's scroll bars
            FormMain.mainCanvas(0).setScrollValue PD_BOTH, 0
            
        End If
        
        
        
        '*************************************************************************************************************************************
        ' If the ExifTool plugin is available, initiate a separate thread for metadata extraction
        '*************************************************************************************************************************************
        
        'Note that metadata extraction is handled asynchronously (e.g. in parallel to the core image loading process), which is
        ' why we launch it so early in the load process.  If the image load fails, we simply ignore any received metadata.
        ' ExifTool is extremely robust, so any errors it experiences during processing will not affect PD.
        If g_ExifToolEnabled And isThisPrimaryImage Then
            
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "Starting separate metadata extraction thread..."
            #End If
            
            startMetadataProcessing sFile(thisImage), targetImage.originalFileFormat, targetImage.imageID
        End If

        'By default, set this image to use the program's default metadata setting (settable from Tools -> Options).
        ' The user may override this setting later, but by default we always start with the user's program-wide setting.
        targetImage.imgMetadata.setMetadataExportPreference g_UserPreferences.GetPref_Long("Saving", "Metadata Export", 1)

        
            
        '*************************************************************************************************************************************
        ' Call the most appropriate load function for this image's format (FreeImage, GDI+, or VB's LoadPicture)
        '*************************************************************************************************************************************
            
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "Determining filetype..."
        #End If
        
        'Initially, set the filetype of the target image to "unknown".  If the load is successful, this value will
        ' be changed to something >= 0. (Note: if FreeImage is used to load the file, this value will be set by the
        ' LoadFreeImageV4 function.)
        If Not (targetImage Is Nothing) Then targetImage.originalFileFormat = -1
        
        'Strip the extension from the file
        FileExtension = UCase(GetExtension(sFile(thisImage)))
        
        loadSuccessful = False
        loadedByOtherMeans = False
        
        'Note that FreeImage may raise additional dialogs (e.g. for HDR/RAW images), so it does not return a binary pass/fail.
        ' If the function fails due to user cancellation, we will suppress subsequent error message boxes.
        Dim freeImage_Return As PD_OPERATION_OUTCOME
        freeImage_Return = PD_FAILURE_GENERIC
            
        'Depending on the file's extension, load the image using the most appropriate image decoding routine
        Select Case FileExtension
        
            'PhotoDemon's custom file format must be handled specially (as obviously, FreeImage and GDI+ won't handle it!)
            Case "PDI", "PDTMP"
            
                'PDI images require zLib, and are only loaded via a custom routine (obviously, since they are PhotoDemon's native format)
                loadSuccessful = LoadPhotoDemonImage(sFile(thisImage), targetDIB, targetImage)
                
                targetImage.originalFileFormat = FIF_PDI
                targetImage.currentFileFormat = FIF_PDI
                targetImage.originalColorDepth = 32
                targetImage.notifyImageChanged UNDO_EVERYTHING
                mustCountColors = False
                
                decoderUsed = PDIDE_INTERNAL
            
            'TMPDIB files are raw pdDIB objects dumped directly to file.  In some cases, this is faster and easier for PD than wrapping
            ' the pdDIB object inside a pdPackage layer (e.g. during clipboard interactions, since we start with a raw pdDIB object
            ' after selections and such are applied to the base layer/image, so we may as well just use the raw pdDIB data we've cached).
            Case "TMPDIB", "PDTMPDIB"
            
                'These raw pdDIB objects may require zLib for parsing (compression is optional), so it is possible for the load function
                ' to fail if zLib goes missing.
                loadSuccessful = LoadRawImageBuffer(sFile(thisImage), targetDIB, targetImage)
                
                targetImage.originalFileFormat = FIF_JPEG
                targetImage.currentFileFormat = FIF_JPEG
                targetImage.originalColorDepth = 32
                targetImage.notifyImageChanged UNDO_EVERYTHING
                mustCountColors = False
                
                decoderUsed = PDIDE_INTERNAL
            
            'Straight TMP files are internal files (BMP, typically) used by PhotoDemon.  A standard flow of load engines is used,
            ' but
            Case "TMP"
                
                If g_ImageFormats.FreeImageEnabled Then
                    pageNumber = 0
                    loadSuccessful = CBool(LoadFreeImageV4(sFile(thisImage), targetDIB, pageNumber, isThisPrimaryImage) = PD_SUCCESS)
                    If loadSuccessful Then
                        decoderUsed = PDIDE_FREEIMAGE
                        targetImage.setDPI targetDIB.getDPI, targetDIB.getDPI
                        targetImage.originalColorDepth = targetDIB.getOriginalColorDepth
                    End If
                End If
                
                If g_ImageFormats.GDIPlusEnabled And (Not loadSuccessful) Then
                    loadSuccessful = LoadGDIPlusImage(sFile(thisImage), targetDIB)
                    If loadSuccessful Then
                        decoderUsed = PDIDE_GDIPLUS
                        targetImage.setDPI targetDIB.getDPI, targetDIB.getDPI
                        targetImage.originalColorDepth = targetDIB.getOriginalColorDepth
                    End If
                End If
                
                If (Not loadSuccessful) Then
                    loadSuccessful = LoadVBImage(sFile(thisImage), targetDIB)
                    If loadSuccessful Then decoderUsed = PDIDE_VBLOADPICTURE
                End If
                
                'Lie and say that the original file format of this image was JPEG.  We do this because tmp images are typically images
                ' captured via non-traditional means (screenshots, scans), and when the user tries to save the file, they should not
                ' be prompted to save it as a BMP.
                targetImage.originalFileFormat = FIF_JPEG
                mustCountColors = True
                
                #If DEBUGMODE = 1 Then
                    If Not loadSuccessful Then
                        pdDebug.LogAction "WARNING!  LoadFileAsNewImage failed on an internal file; both GDI+ and VB failed to handle " & sFile(thisImage) & " correctly."
                    End If
                #End If
                            
            'All other formats follow a set pattern: try to load them via FreeImage (if available), then GDI+, then finally
            ' VB's internal LoadPicture function.
            Case Else
                                
                'If FreeImage is available, we first use it to try and load the image.
                If g_ImageFormats.FreeImageEnabled Then
                
                    'Start by seeing if the image file contains multiple pages.  If it does, we will load each page as a separate layer.
                    If isMultiImage(sFile(thisImage)) > 0 Then
                        
                        'TODO: preferences or prompt for how to handle such files
                        
                        'Mark the image as having multiple pages
                        imageHasMultiplePages = True
                        numOfPages = isMultiImage(sFile(thisImage))
                        
                        'Start by loading just the first page
                        pageNumber = 0
                        loadSuccessful = LoadFreeImageV4(sFile(thisImage), targetDIB, pageNumber, isThisPrimaryImage)
                     
                    'The image only has one page.  Load it!
                    Else
                        pageNumber = 0
                        freeImage_Return = LoadFreeImageV4(sFile(thisImage), targetDIB, pageNumber, isThisPrimaryImage)
                        loadSuccessful = CBool(freeImage_Return = PD_SUCCESS)
                    End If
                    
                    'FreeImage worked!  Copy any relevant information from the DIB to the parent pdImage object (such as file format),
                    ' then continue with the load process.
                    If loadSuccessful Then
                    
                        loadedByOtherMeans = False
                        
                        decoderUsed = PDIDE_FREEIMAGE
                        
                        'Mirror the determined file format from the DIB to the parent pdImage object
                        targetImage.originalFileFormat = targetDIB.getOriginalFormat
                        
                        'Mirror the discovered resolution, if any, from the DIB
                        targetImage.setDPI targetDIB.getDPI, targetDIB.getDPI
                        
                        'Mirror the original file's color depth
                        targetImage.originalColorDepth = targetDIB.getOriginalColorDepth
                        
                        'Finally, copy the background color (if any) from the DIB
                        If (targetImage.originalFileFormat = FIF_PNG) And (targetDIB.getBackgroundColor <> -1) Then
                            targetImage.imgStorage.AddEntry "pngBackgroundColor", targetDIB.getBackgroundColor
                        End If
                        
                    End If
                    
                End If
                
                'If FreeImage fails for some reason, offload the image to GDI+.
                If (Not loadSuccessful) And (freeImage_Return <> PD_FAILURE_USER_CANCELED) And g_ImageFormats.GDIPlusEnabled Then
                    
                    #If DEBUGMODE = 1 Then
                        pdDebug.LogAction "FreeImage refused to load image.  Dropping back to GDI+ and trying again..."
                    #End If
                    
                    loadSuccessful = LoadGDIPlusImage(sFile(thisImage), targetDIB)
                    
                    'If GDI+ loaded the image successfully, note that we have to determine color depth manually.  (There may be a way
                    ' to retrieve that info from GDI+, but I haven't bothered to look!)
                    If loadSuccessful Then
                    
                        loadedByOtherMeans = False
                        
                        decoderUsed = PDIDE_GDIPLUS
                        
                        'Mirror the determined file format from the DIB to the parent pdImage object
                        targetImage.originalFileFormat = targetDIB.getOriginalFormat
                        
                        'Mirror the discovered resolution, if any, from the DIB
                        targetImage.setDPI targetDIB.getDPI, targetDIB.getDPI
                        
                        'Mirror the original file's color depth
                        targetImage.originalColorDepth = targetDIB.getOriginalColorDepth
                        
                    End If
                        
                End If
                
                'If both FreeImage and GDI+ failed, give the image one last try with VB's LoadPicture - UNLESS the image is a WMF or EMF,
                ' which if malformed can cause LoadPicture to experience a silent fail, bringing down the entire program.
                If (Not loadSuccessful) And (freeImage_Return <> PD_FAILURE_USER_CANCELED) And ((FileExtension <> "EMF") And (FileExtension <> "WMF")) Then
                    
                    #If DEBUGMODE = 1 Then
                        Message "GDI+ refused to load image.  Dropping back to internal routines and trying again..."
                    #End If
                    
                    loadSuccessful = LoadVBImage(sFile(thisImage), targetDIB)
                
                    'If VB managed to load the image successfully, note that we have to deteremine color depth manually
                    If loadSuccessful Then
                    
                        decoderUsed = PDIDE_VBLOADPICTURE
                        loadedByOtherMeans = True
                        mustCountColors = True
                        
                    End If
                
                End If
                    
        End Select
        
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "Format-specific parsing complete.  Running a few failsafe checks on the new pdImage object..."
        #End If
        
        'Because ExifTool is sending us data in the background, we periodically yield for metadata piping.
        If targetImage.originalFileFormat <> FIF_PDI Then DoEvents
                
        '*************************************************************************************************************************************
        ' Run a few checks to confirm that the image data was loaded successfully
        '*************************************************************************************************************************************
        
        'Sometimes, our image load functions will think the image loaded correctly, but they will return a blank image.  Check for
        ' non-zero width and height before continuing.
        If ((Not loadSuccessful) Or (targetDIB.getDIBWidth = 0) Or (targetDIB.getDIBHeight = 0)) And isThisPrimaryImage Then
            
            Message "Failed to load %1", sFile(thisImage)
            
            'If multiple files are being loaded, suppress any errors until the end
            If multipleFilesLoading Then
                brokenFiles = brokenFiles & GetFilename(sFile(thisImage)) & vbCrLf
            Else
                If (MacroStatus <> MacroBATCH) And (Not suspendWarnings) And (freeImage_Return <> PD_FAILURE_USER_CANCELED) Then
                    PDMsgBox "Unfortunately, PhotoDemon was unable to load the following image:" & vbCrLf & vbCrLf & "%1" & vbCrLf & vbCrLf & "Please use another program to save this image in a generic format (such as JPEG or PNG) before loading it into PhotoDemon.  Thanks!", vbExclamation + vbOKOnly + vbApplicationModal, "Image Import Failed", sFile(thisImage)
                End If
            End If
            
            'Deactivate the (now useless) pdImage object, and forcibly unload whatever resources it has claimed
            targetImage.deactivateImage
            FullPDImageUnload targetImage.imageID, False
            
            GoTo PreloadMoreImages

        End If
        
        'Because ExifTool is sending us data in the background, we periodically yield for metadata piping.
        If targetImage.originalFileFormat <> FIF_PDI Then DoEvents
        
        'If debug mode is active, post some helpful debugging information
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "Debug note: image load appeared to be successful.  Summary forthcoming."
        #End If
        
        
        '*************************************************************************************************************************************
        ' If the loaded image was in PDI format (PhotoDemon's internal format), skip a number of additional processing steps.
        '*************************************************************************************************************************************
        
        If targetImage.originalFileFormat <> FIF_PDI Then
            
            
            '*************************************************************************************************************************************
            ' If GDI+ or VB's LoadPicture was used to load the file, populate some data fields manually (filetype, color depth, etc)
            '*************************************************************************************************************************************
            
            If loadedByOtherMeans Then
            
                Select Case FileExtension
                    
                    Case "GIF"
                        targetImage.originalFileFormat = FIF_GIF
                        targetImage.originalColorDepth = 8
                        
                    Case "ICO"
                        targetImage.originalFileFormat = FIF_ICO
                    
                    Case "JIF", "JFIF", "JPG", "JPEG", "JPE"
                        targetImage.originalFileFormat = FIF_JPEG
                        targetImage.originalColorDepth = 24
                        
                    Case "PNG"
                        targetImage.originalFileFormat = FIF_PNG
                    
                    Case "TIF", "TIFF"
                        targetImage.originalFileFormat = FIF_TIFF
                    
                    Case "PDI", "TMP", "PDTMP", "TMPDIB", "PDTMPDIB"
                        targetImage.originalFileFormat = FIF_JPEG
                        targetImage.originalColorDepth = 24
                    
                    'Treat anything else as a BMP file
                    Case Else
                        targetImage.originalFileFormat = FIF_BMP
                        
                End Select
            
            End If
            
            DoEvents
           
            
            
            '*************************************************************************************************************************************
            ' If the image contained an embedded ICC profile, apply it now (before counting colors, etc).
            '*************************************************************************************************************************************
            
            'Note that we now need to see if the ICC profile has already been applied.  For CMYK images, the ICC profile will be applied by
            ' the image load function.  If we don't do this, we'll be left with a 32bpp image that contains CMYK data instead of RGBA!
            If targetDIB.ICCProfile.hasICCData And (Not targetDIB.ICCProfile.hasProfileBeenApplied) And (Not targetImage.imgStorage.doesKeyExist("Tone-mapping")) Then
                
                '32bpp images must be un-premultiplied before the transformation
                If targetDIB.getDIBColorDepth = 32 Then targetDIB.SetAlphaPremultiplication False
                
                'Apply the ICC transform
                targetDIB.ICCProfile.applyICCtoSelf targetDIB
                
                '32bpp images must be re-premultiplied after the transformation
                If targetDIB.getDIBColorDepth = 32 Then targetDIB.SetAlphaPremultiplication True
                
            End If
            
            DoEvents
            
            
            
            '*************************************************************************************************************************************
            ' If the incoming image is 24bpp, convert it to 32bpp.  PD assumes an available alpha channel for all layers.
            '*************************************************************************************************************************************
            
            If targetDIB.getDIBColorDepth = 24 Then
            
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "Original image was 24bpp.  Converting to 32bpp now..."
                #End If
                
                If g_GDIPlusAvailable Then
                    GDI_Plus.GDIPlusConvertDIB24to32 targetDIB
                Else
                    targetDIB.convertTo32bpp
                End If
                
            End If
            
            
            '*************************************************************************************************************************************
            ' The target DIB has been loaded successfully, so copy its contents into the main layer of the targetImage
            '*************************************************************************************************************************************
            
            If isThisPrimaryImage Then
                
                'Assemble a base name for this layer
                If Len(imgName) = 0 Then
                    layerNameBase = getFilenameWithoutExtension(sFile(thisImage))
                Else
                    layerNameBase = imgName
                End If
                
                'Images with multiple pages/frames/icons receive special layer naming considering
                If imageHasMultiplePages Or (targetImage.originalFileFormat = FIF_ICO) Then
                
                    layerNameAddon = ""
                    
                    Select Case targetImage.originalFileFormat
                    
                        'GIFs are called "frames" instead of pages
                        Case FIF_GIF
                            layerNameAddon = g_Language.TranslateMessage("frame %1", "1")
                            layerNameAddon = " (" & layerNameAddon & ")"
                        
                        'Icons have their actual dimensions added to the layer name
                        Case FIF_ICO
                            
                            If targetDIB.getOriginalFreeImageColorDepth = 0 Then
                                layerNameAddon = g_Language.TranslateMessage("icon (%1x%2)", CStr(targetDIB.getDIBWidth), CStr(targetDIB.getDIBHeight))
                            Else
                                layerNameAddon = g_Language.TranslateMessage("icon (%1x%2, %3 bpp)", CStr(targetDIB.getDIBWidth), CStr(targetDIB.getDIBHeight), CStr(targetDIB.getOriginalFreeImageColorDepth))
                            End If
                            
                            layerNameAddon = " " & layerNameAddon
                            
                        'Any other format is treated as "pages"
                        Case Else
                            layerNameAddon = g_Language.TranslateMessage("page %1", "1")
                            layerNameAddon = " (" & layerNameAddon & ")"
                        
                    End Select
                    
                    'Merge this newly created add-on string with the original name
                    layerNameBase = layerNameBase & layerNameAddon
                    
                End If
                
                'Create the layer now, and assign our assembled name
                targetImage.getLayerByID(newLayerID).InitializeNewLayer PDL_IMAGE, layerNameBase, targetDIB, targetImage
                
            End If
            
            'Update the pdImage container to be the same size as its (newly created) base layer
            targetImage.updateSize
            
            DoEvents
            
            
            '*************************************************************************************************************************************
            ' If requested by the user, manually count the number of unique colors in the image (to accurately determine color depth)
            '*************************************************************************************************************************************
            
            'At this point, we now have loaded image data in 24 or 32bpp format.  For future reference, let's count
            ' the number of colors present in the image (if the user has allowed it).  If the user HASN'T allowed
            ' it, we have no choice but to rely on whatever color depth was returned by FreeImage or GDI+ (or was
            ' inferred by us for this format, e.g. we know that GIFs are 8bpp).
            
            If isThisPrimaryImage And (g_UserPreferences.GetPref_Boolean("Loading", "Verify Initial Color Depth", True) Or mustCountColors) Then
                
                colorCountCheck = getQuickColorCount(targetDIB, g_CurrentImage)
            
                'If 256 or less colors were found in the image, mark it as 8bpp.  Otherwise, mark it as 24 or 32bpp.
                targetImage.originalColorDepth = getColorDepthFromColorCount(colorCountCheck, targetDIB)
                
                #If DEBUGMODE = 1 Then
                    If g_IsImageGray Then
                        pdDebug.LogAction "Color count successful (" & targetImage.originalColorDepth & " BPP, grayscale)"
                    Else
                        pdDebug.LogAction "Color count successful (" & targetImage.originalColorDepth & " BPP, color)"
                    End If
                #End If
                            
            End If
            
            DoEvents
        
        
        'If the image is in PDI format, the following ELSE branch will be triggered
        Else
        
            'If the caller wants a copy of the image (perhaps for previewing purposes), they can mark "fillDIBWithCompositePDI" as TRUE.
            If fillDIBWithCompositePDI Then targetImage.getCompositedImage targetDIB
            
        End If
                
        '*************************************************************************************************************************************
        ' Determine a name for this image, and store it (along with any other relevant bits) inside the parent pdImage object
        '*************************************************************************************************************************************
        
        
        'Note: this is where PDI format processing picks up again
PDI_Load_Continuation:

        
        'Mark the original file size and file format of the image
        If cFile.FileExist(sFile(thisImage)) Then targetImage.originalFileSize = cFile.FileLenW(sFile(thisImage))
        targetImage.currentFileFormat = targetImage.originalFileFormat
        
        'If Debug Mode is active, supply a basic image summary
        #If DEBUGMODE = 1 Then
        
            pdDebug.LogAction "~ Summary of image """ & GetFilename(sFile(thisImage)) & """ follows ~", , True
            pdDebug.LogAction vbTab & "Image ID: " & targetImage.imageID, , True
            
            Select Case decoderUsed
                
                Case PDIDE_INTERNAL
                    pdDebug.LogAction vbTab & "Load engine: Internal PhotoDemon decoder", , True
                
                Case PDIDE_FREEIMAGE
                    pdDebug.LogAction vbTab & "Load engine: FreeImage plugin", , True
                
                Case PDIDE_GDIPLUS
                    pdDebug.LogAction vbTab & "Load engine: GDI+", , True
                
                Case PDIDE_VBLOADPICTURE
                    pdDebug.LogAction vbTab & "Load engine: VB's LoadPicture() function", , True
                
            End Select
            
            pdDebug.LogAction vbTab & "Detected format: " & g_ImageFormats.getInputFormatDescription(g_ImageFormats.getIndexOfInputFIF(targetImage.originalFileFormat)), , True
            pdDebug.LogAction vbTab & "Image dimensions: " & targetImage.Width & "x" & targetImage.Height, , True
            pdDebug.LogAction vbTab & "Image size (original file): " & Format(CStr(targetImage.originalFileSize), "###,###,###,###") & " Bytes", , True
            pdDebug.LogAction vbTab & "Image size (as loaded, approximate): " & Format(CStr(targetImage.estimateRAMUsage), "###,###,###,###") & " Bytes", , True
            pdDebug.LogAction vbTab & "Original color depth: " & targetImage.originalColorDepth, , True
            pdDebug.LogAction vbTab & "Grayscale: " & CStr(g_IsImageGray), , True
            pdDebug.LogAction vbTab & "ICC profile embedded: " & targetDIB.ICCProfile.hasICCData, , True
            pdDebug.LogAction vbTab & "Multiple pages embedded: " & CStr(imageHasMultiplePages), , True
            pdDebug.LogAction vbTab & "Number of layers: " & targetImage.getNumOfLayers, , True
            pdDebug.LogAction "~ End of image summary ~", , True
            
        #End If
        
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "Determining image title..."
        #End If
        
        'If a different image name has been specified, we can assume the calling routine is NOT loading a file
        ' from disk (e.g. it's a scan, or Internet download, or screen capture, etc.).  Therefore, set the
        ' file name as requested but leave the .LocationOnDisk blank so that a Save command will trigger
        ' the necessary Save As... dialog.
        Dim tmpFilename As String
        
        'Autosaved images are handled differently from normal images.  In order to preserve their original data,
        ' we load certain image data from a standalone XML file.
        If FileExtension = "PDTMP" Then
            
            targetImage.locationOnDisk = sFile(thisImage)
            
            'Ask the AutoSave engine to retrieve this image's data from the matching XML autosave file
            Autosave_Handler.alignLoadedImageWithAutosave targetImage
            
            'This is a bit wacky, but - the MRU engine will automatically update this entry based on its location
            ' on disk (per PD convention) AS STORED IN THE sFile ARRAY.  But as this file's location on disk is
            ' a temp file, we need to rewrite its sFile entry mid-loading!
            sFile(thisImage) = targetImage.locationOnDisk
        
        'This is a non-autosave (normal!) image.
        Else
        
            If Len(imgName) = 0 Then
                'The calling routine hasn't specified an image name, so assume this is a normal load situation.
                ' That means pulling the filename from the file itself.
                targetImage.locationOnDisk = sFile(thisImage)
                
                tmpFilename = sFile(thisImage)
                StripFilename tmpFilename
                targetImage.originalFileNameAndExtension = tmpFilename
                StripOffExtension tmpFilename
                targetImage.originalFileName = tmpFilename
                
                'Disable the save button, because this file exists on disk
                If targetImage.currentFileFormat = FIF_PDI Then
                    targetImage.setSaveState True, pdSE_SavePDI
                Else
                    targetImage.setSaveState True, pdSE_SaveFlat
                End If
                
            Else
            
                'The calling routine has specified a file name.  Assume this is a special case, and force a Save As...
                ' dialog in the future by not specifying a location on disk
                targetImage.locationOnDisk = ""
                targetImage.originalFileNameAndExtension = imgName
                
                tmpFilename = imgName
                StripOffExtension tmpFilename
                targetImage.originalFileName = tmpFilename
                
                'Similarly, enable the save button
                targetImage.setSaveState False, pdSE_AnySave
                
            End If
        
        End If
        
        'Because ExifTool is sending us data in the background, we periodically yield for metadata piping.
        If targetImage.originalFileFormat <> FIF_PDI Then DoEvents
        
        
        '*************************************************************************************************************************************
        ' If this is a primary image, update all relevant interface elements (image size display, 24/32bpp options, custom form icon, etc)
        '*************************************************************************************************************************************
        
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "Finalizing image details..."
        #End If
        
        'If this is a primary image, it needs to be rendered to the screen
        If isThisPrimaryImage Then
            
            'Create an icon-sized version of this image, which we will use as form's taskbar icon
            If MacroStatus <> MacroBATCH Then createCustomFormIcon targetImage
            
            'Register this image with the image tab bar
            toolbar_ImageTabs.registerNewImage g_CurrentImage
            
            'Just to be safe, update the color management profile of the current monitor
            CheckParentMonitor True
            
            'If the user wants us to resize the image to fit on-screen, do that now
            If g_AutozoomLargeImages = 0 Then FitImageToViewport True
            
            'g_AllowViewportRendering may have been reset by this point (by the FitImageToViewport sub, among others), so set it back to False, then
            ' update the zoom combo box to match the zoom assigned by the window-fit function.
            g_AllowViewportRendering = False
            FormMain.mainCanvas(0).getZoomDropDownReference().ListIndex = targetImage.currentZoomValue
        
            'Now that the image's window has been fully sized and moved around, use Viewport_Engine.Stage1_InitializeBuffer to set up any scrollbars and a back-buffer
            g_AllowViewportRendering = True
            Viewport_Engine.Stage1_InitializeBuffer targetImage, FormMain.mainCanvas(0), VSR_ResetToZero
                                    
            'Add this file to the MRU list (unless specifically told not to)
            If ToUpdateMRU And (pageNumber = 0) And (MacroStatus <> MacroBATCH) Then g_RecentFiles.MRU_AddNewFile sFile(thisImage), targetImage
            
            'Reflow any image-window-specific display elements on the actual image form (status bar, rulers, etc)
            FormMain.mainCanvas(0).fixChromeLayout
            
            'Because ExifTool is sending us data in the background, we periodically yield for metadata piping.
            If targetImage.originalFileFormat <> FIF_PDI Then DoEvents
            
        End If
        
        
        
        '*************************************************************************************************************************************
        ' If the just-loaded image was in a multipage format (icon, animated GIF, multipage TIFF), perform a few extra checks.
        '*************************************************************************************************************************************
        
        'Before continuing on to the next image (if any), see if the just-loaded image contains multiple pages within the file.
        ' If it does, load each page into its own layer.
        If imageHasMultiplePages Then
            
            Dim pageTracker As Long
            
            'Call LoadFileAsNewImage again for each individual frame in the multipage file
            For pageTracker = 1 To numOfPages - 1
                
                'To load each page as its own image, use the code below
                'If UCase(GetExtension(sFile(thisImage))) = "GIF" Then
                '    LoadFileAsNewImage tmpStringArray, False, targetImage.originalFileName & " (" & g_Language.TranslateMessage("frame") & " " & (pageTracker + 1) & ")." & GetExtension(sFile(thisImage)), targetImage.originalFileName & " (" & g_Language.TranslateMessage("frame") & " " & (pageTracker + 1) & ")." & GetExtension(sFile(thisImage)), , , , pageTracker
                'ElseIf UCase(GetExtension(sFile(thisImage))) = "ICO" Then
                '    LoadFileAsNewImage tmpStringArray, False, targetImage.originalFileName & " (" & g_Language.TranslateMessage("icon") & " " & (pageTracker + 1) & ")." & GetExtension(sFile(thisImage)), targetImage.originalFileName & " (" & g_Language.TranslateMessage("icon") & " " & (pageTracker + 1) & ")." & GetExtension(sFile(thisImage)), , , , pageTracker
                'Else
                '    LoadFileAsNewImage tmpStringArray, False, targetImage.originalFileName & " (" & g_Language.TranslateMessage("page") & " " & (pageTracker + 1) & ")." & GetExtension(sFile(thisImage)), targetImage.originalFileName & " (" & g_Language.TranslateMessage("page") & " " & (pageTracker + 1) & ")." & GetExtension(sFile(thisImage)), , , , pageTracker
                'End If
                
                
                'To load each page to its own layer, use the code below
                
                'Create a blank layer in the receiving image, and retrieve a pointer to it
                newLayerID = pdImages(g_CurrentImage).createBlankLayer
                
                'Clear the temporary DIB
                Set targetDIB = New pdDIB
                
                'Load the next page into the temporary DIB
                loadSuccessful = LoadFreeImageV4(sFile(thisImage), targetDIB, pageTracker, isThisPrimaryImage)
                
                'If the load was successful, copy the DIB into place
                If loadSuccessful Then
                
                    'Convert 24bpp layers to 32bpp
                    If targetDIB.getDIBColorDepth = 24 Then
                    
                        If g_GDIPlusAvailable Then
                            GDI_Plus.GDIPlusConvertDIB24to32 targetDIB
                        Else
                            targetDIB.convertTo32bpp
                        End If
                    
                    End If
                    
                    'Determine a name for each layer, contingent on its size and type
                    layerNameAddon = ""
                    
                    Select Case targetImage.originalFileFormat
                    
                        'GIFs are called "frames" instead of pages
                        Case FIF_GIF
                            layerNameAddon = g_Language.TranslateMessage("frame")
                            layerNameAddon = " (" & layerNameAddon & " " & CStr(pageTracker + 1) & ")"
                        
                        'Icons have their actual dimensions added to the layer name
                        Case FIF_ICO
                            
                            If targetDIB.getOriginalFreeImageColorDepth = 0 Then
                                layerNameAddon = g_Language.TranslateMessage("icon (%1x%2)", CStr(targetDIB.getDIBWidth), CStr(targetDIB.getDIBHeight))
                            Else
                                layerNameAddon = g_Language.TranslateMessage("icon (%1x%2, %3 bpp)", CStr(targetDIB.getDIBWidth), CStr(targetDIB.getDIBHeight), CStr(targetDIB.getOriginalFreeImageColorDepth))
                            End If
                            
                            layerNameAddon = " " & layerNameAddon
                            
                        'Any other format is treated as "pages"
                        Case Else
                            layerNameAddon = g_Language.TranslateMessage("page")
                            layerNameAddon = " (" & layerNameAddon & " " & CStr(pageTracker + 1) & ")"
                        
                    End Select
                    
                    
                    'Copy the DIB into the layer, with a relevant name attached
                    If Len(imgName) = 0 Then
                        targetImage.getLayerByID(newLayerID).InitializeNewLayer PDL_IMAGE, getFilenameWithoutExtension(sFile(thisImage)) & layerNameAddon, targetDIB, targetImage
                    Else
                        targetImage.getLayerByID(newLayerID).InitializeNewLayer PDL_IMAGE, imgName & layerNameAddon, targetDIB, targetImage
                    End If
                    
                    'Redraw the main viewport
                    Viewport_Engine.Stage1_InitializeBuffer targetImage, FormMain.mainCanvas(0), VSR_ResetToZero
                
                'If the load was unsuccessful, delete the blank layer we created
                Else
                    targetImage.deleteLayerByIndex pdImages(g_CurrentImage).getLayerIndexFromID(newLayerID)
                End If
            
            'Continue on with the next page
            Next pageTracker
            
            'Now, as a convenience, make all but the first frame invisible.
            If targetImage.getNumOfLayers > 1 Then
            
                For pageTracker = 1 To targetImage.getNumOfLayers - 1
                    targetImage.getLayerByIndex(pageTracker).setLayerVisibility False
                Next pageTracker
        
            End If
        
        End If
        
        
        
        
        '*************************************************************************************************************************************
        ' Hopefully metadata processing has finished, but if it hasn't, start a timer on the main form, which will wait for it to complete.
        '*************************************************************************************************************************************
        
        'Ask the metadata handler if it has finished parsing the image
        If g_ExifToolEnabled And isThisPrimaryImage And (targetImage.originalFileFormat <> FIF_PDI) Then

            'Wait for metadata parsing to finish...
            If isMetadataFinished Then
            
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "Metadata retrieved successfully."
                #End If
            
                targetImage.imgMetadata.loadAllMetadata retrieveMetadataString, targetImage.imageID
            
            Else
                
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "Metadata parsing hasn't finished; switching to asynchronous wait mode..."
                #End If
                
                If Not FormMain.tmrMetadata.Enabled Then FormMain.tmrMetadata.Enabled = True
            
            End If

            'Next, retrieve any specific metadata-related entries that may be useful to further processing, like image resolution
            Dim xResolution As Double, yResolution As Double
            If targetImage.imgMetadata.getResolution(xResolution, yResolution) Then
                targetImage.setDPI xResolution, yResolution
            End If
        

        End If
        
        
        '*************************************************************************************************************************************
        ' As of 2014, the new Active Undo/Redo engine requires a base pdImage copy as the starting point for Undo/Redo diffs.
        '*************************************************************************************************************************************
        
        'If this is a primary image, force an immediate Undo/Redo write to file.  This serves multiple purposes: it is our
        ' baseline for calculating future Undo/Redo diffs, and it can be used to recover the original file if something
        ' goes wrong before the user performs a manual save (e.g. AutoSave).
        '
        '(Note that all Undo behavior is disabled during batch processing, to improve performance, so we can skip this step.)
        If isThisPrimaryImage And (MacroStatus <> MacroBATCH) Then
            
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "Creating initial auto-save entry (this may take a moment)..."
            #End If
            
            targetImage.undoManager.createUndoData g_Language.TranslateMessage("Original image"), "", UNDO_EVERYTHING
            
        End If
        
        
        '*************************************************************************************************************************************
        ' Image loaded successfully.  Carry on.
        '*************************************************************************************************************************************
        
        targetImage.loadedSuccessfully = True
        
        'In debug mode, note the new memory baseline, post-load
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "targetImage.loadedSuccessfully set to TRUE"
            pdDebug.LogAction "New memory report after loading image """ & GetFilename(sFile(thisImage)) & """:"
            pdDebug.LogAction "", PDM_MEM_REPORT
            
            'Also report an estimated memory delta, based on the pdImage object's self-reported memory usage.
            ' This provides a nice baseline for making sure PD's memory usage isn't out of whack for a given image.
            pdDebug.LogAction "(FYI, expected delta was approximately " & Format(CStr(targetImage.estimateRAMUsage \ 1000), "###,###,###,###") & " K)"
        #End If
        
        
    '*************************************************************************************************************************************
    ' Move on to the next image.
    '*************************************************************************************************************************************
        
PreloadMoreImages:

    'If we have more images to process, now's the time to do it!
    Next thisImage
        
    
    '*************************************************************************************************************************************
    ' As all images have now loaded, re-enable the main form
    '*************************************************************************************************************************************
    
    FormMain.Enabled = True
    
    'Synchronize all interface elements to match the newly loaded image(s)
    SyncInterfaceToCurrentImage
    toolbar_ImageTabs.forceRedraw
    
    
    '*************************************************************************************************************************************
    ' Before finishing, display any relevant load problems (missing files, invalid formats, etc)
    '*************************************************************************************************************************************
    
    'Restore the screen cursor if necessary
    If pageNumber <= 0 Then Screen.MousePointer = vbNormal
        
    'If multiple images were loaded and everything went well, display a success message
    If multipleFilesLoading Then
        If (Len(missingFiles) = 0) And (Len(brokenFiles) = 0) And isThisPrimaryImage Then Message "All images loaded successfully."
    Else
        If isThisPrimaryImage And Not (targetImage Is Nothing) Then
            If targetImage.loadedSuccessfully Then Message "Image loaded successfully."
        End If
    End If
        
    'Finally, if we were loading multiple images and something went wrong (missing files, broken files), let the user know about them.
    If multipleFilesLoading And (Len(missingFiles) <> 0) Then
        Message "All images loaded, except for those that could not be found."
        If Not suspendWarnings Then
            PDMsgBox "Unfortunately, PhotoDemon was unable to find the following image(s):" & vbCrLf & vbCrLf & "%1" & vbCrLf & vbCrLf & "If these images were originally located on removable media (DVD, USB drive, etc), please re-insert or re-attach the media and try again.", vbApplicationModal + vbExclamation + vbOKOnly, "Image files missing", missingFiles
        End If
    End If
        
    If multipleFilesLoading And (Len(brokenFiles) <> 0) Then
        Message "All images loaded, except for those in invalid formats."
        If Not suspendWarnings Then
            PDMsgBox "Unfortunately, PhotoDemon was unable to load the following image(s):" & vbCrLf & vbCrLf & "%1" & vbCrLf & vbCrLf & "Please use another program to save these images in a generic format (such as JPEG or PNG) before loading them into PhotoDemon. Thanks!", vbExclamation + vbOKOnly + vbApplicationModal, "Image Formats Not Supported", brokenFiles
        End If
    End If
    
    #If DEBUGMODE = 1 Then
        'The line below can be uncommented to report image load times.
        pdDebug.LogAction "Image loaded in %1 seconds", Format$((Timer - startTime), "0.000")
    #End If
        
End Sub

'Quick and dirty function for loading an image file to a containing DIB.  This function provides none of the extra scans or features
' that the more advanced LoadFileAsNewImage does; instead, it is assumed that the calling function will handle any extra work.
' (Note that things like metadata will not be processed *at all* for the image file.)
'
'That said, FreeImage/GDI+ are still used intelligently.  The function will return TRUE if successful.
Public Function QuickLoadImageToDIB(ByVal imagePath As String, ByRef targetDIB As pdDIB) As Boolean
    
    'Even though this function is designed to operate as quickly as possible, some images may take a long time to load.
    ' Display a busy cursor
    If Screen.MousePointer <> vbHourglass Then Screen.MousePointer = vbHourglass
            
    'To improve load time, declare a variety of other variables outside the image load loop
    Dim FileExtension As String
    Dim loadSuccessful As Boolean
    
    'To prevent re-entry problems, forcibly disable the main form until loading is complete
    FormMain.Enabled = False
    
    'Before attempting to load an image, make sure it exists
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    If Not cFile.FileExist(imagePath) Then
        PDMsgBox "Unfortunately, the image '%1' could not be found." & vbCrLf & vbCrLf & "If this image was originally located on removable media (DVD, USB drive, etc), please re-insert or re-attach the media and try again.", vbApplicationModal + vbExclamation + vbOKOnly, "File not found", imagePath
        QuickLoadImageToDIB = False
        FormMain.Enabled = True
        Screen.MousePointer = vbNormal
        Exit Function
    End If
        
    'Prepare a dummy pdImage object, which some external functions may require
    Dim tmpPDImage As pdImage
    Set tmpPDImage = New pdImage
    
    'Determine the most appropriate load function for this image's format (FreeImage, GDI+, or VB's LoadPicture).  Note that FreeImage does not
    ' return a generic pass/fail value.
    Dim freeImageReturn As PD_OPERATION_OUTCOME
    freeImageReturn = PD_FAILURE_GENERIC
    
    'Start by stripping the extension from the file path
    FileExtension = UCase$(cFile.GetFileExtension(imagePath))
    loadSuccessful = False
    
    'Depending on the file's extension, load the image using the most appropriate image decoding routine
    Select Case FileExtension
    
        'PhotoDemon's custom file format must be handled specially (as obviously, FreeImage and GDI+ won't handle it!)
        Case "PDI"
        
            'PDI images require zLib, and are only loaded via a custom routine (obviously, since they are PhotoDemon's native format)
            loadSuccessful = LoadPhotoDemonImage(imagePath, targetDIB, tmpPDImage)
            
            'Retrieve a copy of the fully composited image
            tmpPDImage.getCompositedImage targetDIB
            
        'TMP files are internal PD temp files generated from a wide variety of use-cases (Clipboard is one example).  These are
        ' typically in BMP format, but this is not contractual.  A standard cascade of load functions is used.
        Case "TMP"
            If g_ImageFormats.FreeImageEnabled Then loadSuccessful = CBool(LoadFreeImageV4(imagePath, targetDIB, , False) = PD_SUCCESS)
            If g_ImageFormats.GDIPlusEnabled And (Not loadSuccessful) Then loadSuccessful = LoadGDIPlusImage(imagePath, targetDIB)
            If (Not loadSuccessful) Then loadSuccessful = LoadVBImage(imagePath, targetDIB)
            If (Not loadSuccessful) Then loadSuccessful = LoadRawImageBuffer(imagePath, targetDIB, tmpPDImage)
            
        'TMPDIB files are raw pdDIB objects dumped directly to file.  In some cases, this is faster and easier for PD than wrapping
        ' the pdDIB object inside a pdPackage layer (especially if this function is going to be used, since we're just going to
        ' decode the saved file into a pdDIB anyway).
        Case "TMPDIB", "PDTMPDIB"
            loadSuccessful = LoadRawImageBuffer(imagePath, targetDIB, tmpPDImage)
            
        'PDTMP files are custom PD-format files saved ONLY during Undo/Redo or Autosaving.  As such, they have some weirdly specific
        ' parsing criteria during the master load function, but for quick-loading, we can simply grab the raw image buffer portion.
        Case "PDTMP"
            loadSuccessful = LoadRawImageBuffer(imagePath, targetDIB, tmpPDImage)
            
        'All other formats follow a set pattern: try to load them via FreeImage (if it's available), then GDI+, then finally
        ' VB's internal LoadPicture function.
        Case Else
            
            'If FreeImage is available, use it to try and load the image.
            If g_ImageFormats.FreeImageEnabled Then
                freeImageReturn = LoadFreeImageV4(imagePath, targetDIB, 0, False)
                If freeImageReturn = PD_SUCCESS Then loadSuccessful = True Else loadSuccessful = False
            End If
                
            'If FreeImage fails for some reason, offload the image to GDI+
            If (Not loadSuccessful) And g_ImageFormats.GDIPlusEnabled Then loadSuccessful = LoadGDIPlusImage(imagePath, targetDIB)
            
            'If both FreeImage and GDI+ failed, give the image one last try with VB's LoadPicture - UNLESS the image is a WMF or EMF,
            ' which can cause LoadPicture to experience a silent fail, thus bringing down the entire program.
            If (Not loadSuccessful) And ((FileExtension <> "EMF") And (FileExtension <> "WMF")) Then loadSuccessful = LoadVBImage(imagePath, targetDIB)
                    
    End Select
    
    
    'Sometimes, our image load functions will think the image loaded correctly, but they will return a blank image.  Check for
    ' non-zero width and height before continuing.
    If (Not loadSuccessful) Or (targetDIB.getDIBWidth = 0) Or (targetDIB.getDIBHeight = 0) Then
        
        'Only display an error dialog if the import wasn't canceled by the user
        If freeImageReturn <> PD_FAILURE_USER_CANCELED Then
            Message "Failed to load %1", imagePath
            PDMsgBox "Unfortunately, PhotoDemon was unable to load the following image:" & vbCrLf & vbCrLf & "%1" & vbCrLf & vbCrLf & "Please use another program to save this image in a generic format (such as JPEG or PNG) before loading it into PhotoDemon.  Thanks!", vbExclamation + vbOKOnly + vbApplicationModal, "Image Import Failed", imagePath
        Else
            Message "Layer import canceled."
        End If
        
        'Deactivate the (now useless) DIB
        targetDIB.eraseDIB
        
        'Re-enable the main interface
        FormMain.Enabled = True
        Screen.MousePointer = vbNormal
        
        'Exit with failure status
        QuickLoadImageToDIB = False
        
        Exit Function
        
    End If
    
    'If the loaded image contains alpha data, verify it.  If the alpha channel is blank (e.g. all 0 or all 255), convert it to 24bpp
    If targetDIB.getDIBColorDepth = 32 Then
        
        'Make sure the user hasn't disabled alpha channel validation
        If g_UserPreferences.GetPref_Boolean("Transparency", "Validate Alpha Channels", True) Then
            
            'Verify the alpha channel.  If this function returns FALSE, the current alpha channel is unnecessary.
            If Not DIB_Handler.verifyDIBAlphaChannel(targetDIB) Then targetDIB.convertTo24bpp
            
        End If
        
    End If
    
    'If the image contained an embedded ICC profile, apply it now.
    '
    'Note that we need to check if the ICC profile has already been applied.  For CMYK images, the ICC profile will be applied by
    ' the image load function.  (If we don't do this, we'll be left with a 32bpp image that contains CMYK data instead of RGBA!)
    If targetDIB.ICCProfile.hasICCData And (Not targetDIB.ICCProfile.hasProfileBeenApplied) Then
        
        '32bpp images must be un-premultiplied before the transformation
        If targetDIB.getDIBColorDepth = 32 Then targetDIB.SetAlphaPremultiplication False
        
        'Apply the ICC transform
        targetDIB.ICCProfile.applyICCtoSelf targetDIB
        
        '32bpp images must be re-premultiplied after the transformation
        If targetDIB.getDIBColorDepth = 32 Then targetDIB.SetAlphaPremultiplication True
    
    End If

    'Restore the main interface
    FormMain.Enabled = True
    Screen.MousePointer = vbNormal

    'If we made it all the way here, the image file was loaded successfully!
    QuickLoadImageToDIB = True

End Function

'PDI loading.  "PhotoDemon Image" files are the only format PD supports for saving layered images.  PDI to PhotoDemon is like
' PSD to PhotoShop, or XCF to Gimp.
'
'Note the unique "sourceIsUndoFile" parameter for this load function.  PDI files are used to store undo/redo data, and when one of their
' kind is loaded as part of an Undo/Redo action, we must ignore certain elements stored in the file (e.g. settings like "LastSaveFormat"
' which we do not want to Undo/Redo).  This parameter is passed to the pdImage initializer, and it tells it to ignore certain settings.
Public Function LoadPhotoDemonImage(ByVal PDIPath As String, ByRef dstDIB As pdDIB, ByRef dstImage As pdImage, Optional ByVal sourceIsUndoFile As Boolean = False) As Boolean
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "PDI file identified.  Starting pdPackage decompression..."
    #End If
    
    On Error GoTo LoadPDIFail
    
    'First things first: create a pdPackage instance.  It will handle all the messy business of extracting individual data bits
    ' from the source file.
    Dim pdiReader As pdPackager
    Set pdiReader = New pdPackager
    pdiReader.init_ZLib "", True, g_ZLibEnabled
    
    'Load the file into the pdPackager instance.  It will cache the file contents, so we only have to do this once.
    ' Note that this step will also validate the incoming file.
    If pdiReader.readPackageFromFile(PDIPath, PD_IMAGE_IDENTIFIER) Then
    
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "pdPackage successfully read and initialized.  Starting package parsing..."
        #End If
    
        'First things first: extract the pdImage header, which will be in Node 0.  (We could double-check this by searching
        ' for the node entry by name, but since there is no variation, it's faster to access it directly.)
        Dim retBytes() As Byte, retString As String
        
        If pdiReader.getNodeDataByIndex(0, True, retBytes, sourceIsUndoFile) Then
            
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "Initial PDI node retrieved.  Initializing corresponding pdImage object..."
            #End If
            
            'Copy the received bytes into a string
            If pdiReader.getPDPackageVersion >= PDPACKAGE_UNICODE_FRIENDLY_VERSION Then
                retString = Space$((UBound(retBytes) + 1) \ 2)
                CopyMemory ByVal StrPtr(retString), ByVal VarPtr(retBytes(0)), UBound(retBytes) + 1
            Else
                retString = StrConv(retBytes, vbUnicode)
            End If
            
            'Pass the string to the target pdImage, which will read the XML data and initialize itself accordingly
            dstImage.readExternalData retString, True, sourceIsUndoFile
        
        'Bytes could not be read, or alternately, checksums didn't match for the first node.
        Else
            Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
        End If
        
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "pdImage created successfully.  Moving on to individual layers..."
        #End If
        
        'With the main pdImage now assembled, the next task is to populate all layers with two pieces of information:
        ' 1) The layer header, which contains stuff like layer name, opacity, blend mode, etc
        ' 2) Layer-specific information, which varies by layer type.  For DIBs, this will be a raw stream of bytes
        '    containing the layer DIB's raster data.  For text or other vector layers, this is an XML stream containing
        '    whatever information is necessary to construct the layer from scratch.
        
        Dim i As Long
        For i = 0 To dstImage.getNumOfLayers - 1
        
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "Retrieving layer header " & i & "..."
            #End If
        
            'First, retrieve the layer's header
            If pdiReader.getNodeDataByIndex(i + 1, True, retBytes, sourceIsUndoFile) Then
            
                'Copy the received bytes into a string
                If pdiReader.getPDPackageVersion >= PDPACKAGE_UNICODE_FRIENDLY_VERSION Then
                    retString = Space$((UBound(retBytes) + 1) \ 2)
                    CopyMemory ByVal StrPtr(retString), ByVal VarPtr(retBytes(0)), UBound(retBytes) + 1
                Else
                    retString = StrConv(retBytes, vbUnicode)
                End If
                
                'Pass the string to the target layer, which will read the XML data and initialize itself accordingly
                If Not dstImage.getLayerByIndex(i).CreateNewLayerFromXML(retString) Then
                    Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
                End If
                
            Else
                Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
            End If
            
            'How we extract the rest of the layer's data varies by layer type.  Raster layers can skip the need for a temporary buffer,
            ' because we've already created a DIB with a built-in buffer for the pixel data.
            '
            'Other layer types (e.g. vector layers) are tiny so a temporary buffer does not matter; also, unlike raster buffers, we cannot
            ' easily predict the buffer size in advance, so we rely on pdPackage to do it for us
            Dim nodeLoadedSuccessfully As Boolean
            nodeLoadedSuccessfully = False
            
            'Image (raster) layers
            If dstImage.getLayerByIndex(i).isLayerRaster Then
                
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "Raster layer identified.  Retrieving pixel bits..."
                #End If
                
                'We are going to load the node data directly into the DIB, completely bypassing the need for a temporary array.
                Dim tmpDIBPointer As Long, tmpDIBLength As Long
                dstImage.getLayerByIndex(i).layerDIB.retrieveDIBPointerAndSize tmpDIBPointer, tmpDIBLength
                
                'At present, all pdPackage layers will contain premultiplied alpha, so force the corresponding state now
                dstImage.getLayerByIndex(i).layerDIB.setInitialAlphaPremultiplicationState True
                
                nodeLoadedSuccessfully = pdiReader.getNodeDataByIndex_UnsafeDstPointer(i + 1, False, tmpDIBPointer, sourceIsUndoFile)
            
            'Text and other vector layers
            ElseIf dstImage.getLayerByIndex(i).isLayerVector Then
                
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "Vector layer identified.  Retrieving layer XML..."
                #End If
                
                If pdiReader.getNodeDataByIndex(i + 1, False, retBytes, sourceIsUndoFile) Then
                
                    'Convert the byte array to a Unicode string.  Note that we do not need an ASCII branch for old versions,
                    ' as vector layers were implemented after pdPackager gained full Unicode compatibility.
                    retString = Space$((UBound(retBytes) + 1) \ 2)
                    CopyMemory ByVal StrPtr(retString), ByVal VarPtr(retBytes(0)), UBound(retBytes) + 1
                    
                    'Pass the string to the target layer, which will read the XML data and initialize itself accordingly
                    If dstImage.getLayerByIndex(i).CreateVectorDataFromXML(retString) Then
                        nodeLoadedSuccessfully = True
                    Else
                        Err.Raise PDP_GENERIC_ERROR, , "PDI Node (vector type) could not be read; data invalid or checksums did not match."
                    End If
                    
                Else
                    Err.Raise PDP_GENERIC_ERROR, , "PDI Node (vector type) could not be read; data invalid or checksums did not match."
                End If
                    
            'In the future, additional layer types can be handled here
            Else
                Debug.Print "WARNING! Unknown layer type exists in this PDI file: " & dstImage.getLayerByIndex(i).getLayerType
            
            End If
            
            'If successful, notify the parent of the change
            If nodeLoadedSuccessfully Then
                dstImage.notifyImageChanged UNDO_LAYER, i
            Else
                Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
            End If
        
        Next i
        
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "All layers loaded.  Looking for remaining non-essential PDI data..."
        #End If
        
        'Finally, check to see if the PDI image has a metadata entry.  If it does, load that data now.
        If pdiReader.getNodeDataByName("pdMetadata_Raw", True, retBytes, sourceIsUndoFile) Then
        
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "Raw metadata chunk found.  Retrieving now..."
            #End If
        
            'Copy the received bytes into a string
            If pdiReader.getPDPackageVersion >= PDPACKAGE_UNICODE_FRIENDLY_VERSION Then
                retString = Space$((UBound(retBytes) + 1) \ 2)
                CopyMemory ByVal StrPtr(retString), ByVal VarPtr(retBytes(0)), UBound(retBytes) + 1
            Else
                retString = StrConv(retBytes, vbUnicode)
            End If
            
            'Pass the string to the parent image's metadata handler, which will parse the XML data and prepare a matching
            ' internal metadata struct.
            If Not dstImage.imgMetadata.loadAllMetadata(retString, dstImage.imageID) Then
                
                'For invalid metadata, do not reject the rest of the PDI file.  Instead, just warn the user and carry on.
                Debug.Print "PDI Metadata Node rejected by metadata parser."
                
            End If
        
        End If
        
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "PDI parsing complete.  Returning control to main image loader..."
        #End If
        
        'Funny quirk: this function has no use for the dstDIB parameter, but if that DIB returns a width/height of zero,
        ' the upstream load function will think the load process failed.  Because of that, we must initialize the DIB to *something*.
        If dstDIB Is Nothing Then Set dstDIB = New pdDIB
        dstDIB.createBlank 16, 16, 32, 0
        
        'That's all there is to it!  Mark the load as successful and carry on.
        LoadPhotoDemonImage = True
    
    Else
    
        'If we made it to this block, the first stage of PDI validation failed, meaning this file isn't in PDI format.
        Message "Selected file is not in PDI format.  Load abandoned."
        LoadPhotoDemonImage = False
    
    End If
    
    Exit Function
    
LoadPDIFail:
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "WARNING!  LoadPDIFail error routine reached.  Checking for known error states..."
    #End If
    
    'Before falling back to a generic error message, check for a couple known problem states.
    
    'Case 1: zLib is required for this file, but the user doesn't have the zLib plugin
    If pdiReader.getPackageFlag(PDP_FLAG_ZLIB_REQUIRED, PDP_LOCATION_ANY) And (Not g_ZLibEnabled) Then
        PDMsgBox "The PDI file ""%1"" contains compressed data, but the zLib plugin is missing or disabled." & vbCrLf & vbCrLf & "To enable support for compressed PDI files, click Help > Check for Updates, and when prompted, allow PhotoDemon to download all recommended plugins.", vbInformation + vbOKOnly + vbApplicationModal, "zLib plugin missing", GetFilename(PDIPath)
        Exit Function
    End If

    Select Case Err.Number
    
        Case PDP_GENERIC_ERROR
            Message "PDI node could not be read; file may be invalid or corrupted.  Load abandoned."
            
        Case Else
            Message "An error has occurred (#%1 - %2).  PDI load abandoned.", Err.Number, Err.Description
        
    End Select
    
    LoadPhotoDemonImage = False
    Exit Function

End Function

'Load just the layer stack from a standard PDI file, and non-destructively align our current layer stack to match.
' At present, this function is only used internally by the Undo/Redo engine.
Public Function LoadPhotoDemonImageHeaderOnly(ByVal PDIPath As String, ByRef dstImage As pdImage) As Boolean
    
    On Error GoTo LoadPDIHeaderFail
    
    'First things first: create a pdPackage instance.  It will handle all the messy business of extracting individual data bits
    ' from the source file.
    Dim pdiReader As pdPackager
    Set pdiReader = New pdPackager
    pdiReader.init_ZLib "", True, g_ZLibEnabled
    
    'Load the file into the pdPackager instance.  It will cache the file contents, so we only have to do this once.
    ' Note that this step will also validate the incoming file.
    If pdiReader.readPackageFromFile(PDIPath, PD_IMAGE_IDENTIFIER) Then
    
        'First things first: extract the pdImage header, which will be in Node 0.  (We could double-check this by searching
        ' for the node entry by name, but since there is no variation, it's faster to access it directly.)
        Dim retBytes() As Byte, retString As String
        
        If pdiReader.getNodeDataByIndex(0, True, retBytes, True) Then
        
            'Copy the received bytes into a string
            If pdiReader.getPDPackageVersion >= PDPACKAGE_UNICODE_FRIENDLY_VERSION Then
                retString = Space$((UBound(retBytes) + 1) \ 2)
                CopyMemory ByVal StrPtr(retString), ByVal VarPtr(retBytes(0)), UBound(retBytes) + 1
            Else
                retString = StrConv(retBytes, vbUnicode)
            End If
            
            'Pass the string to the target pdImage, which will read the XML data and initialize itself accordingly
            dstImage.readExternalData retString, True, True, True
        
        'Bytes could not be read, or alternately, checksums didn't match for the first node.
        Else
            Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
        End If
        
        'With the main pdImage now assembled, the next task is to populate all layer headers.  This is a bit more
        ' confusing than a regular PDI load, because we have to maintain existing layer DIB data (ugh!).
        ' So basically, we must:
        ' 1) Extract each layer header from file, in turn
        ' 2) See if the current pdImage copy of this layer is in the proper position in the layer stack; if it isn't,
        '    move it into the location specified by the PDI file.
        ' 3) Ask the layer to non-destructively overwrite its header with the header from the PDI file (e.g. don't
        '    touch its DIB or vector-specific contents).
        
        Dim layerNodeName As String, layerNodeID As Long, layerNodeType As Long
        
        Dim i As Long
        For i = 0 To dstImage.getNumOfLayers - 1
        
            'Before doing anything else, retrieve the ID of the node at this position.  (Retrieve the rest of the node
            ' header too, although we don't actually have a use for those values at present.)
            pdiReader.getNodeInfo i + 1, layerNodeName, layerNodeID, layerNodeType
            
            'We now know what layer ID is supposed to appear at this position in the layer stack.  If that layer ID
            ' is *not* in its proper position, move it now.
            If dstImage.getLayerIndexFromID(layerNodeID) <> i Then dstImage.swapTwoLayers dstImage.getLayerIndexFromID(layerNodeID), i
            
            'Now that the node is in place, we can retrieve its header.
            If pdiReader.getNodeDataByIndex(i + 1, True, retBytes, True) Then
            
                'Copy the received bytes into a string
                If pdiReader.getPDPackageVersion >= PDPACKAGE_UNICODE_FRIENDLY_VERSION Then
                    retString = Space$((UBound(retBytes) + 1) \ 2)
                    CopyMemory ByVal StrPtr(retString), ByVal VarPtr(retBytes(0)), UBound(retBytes) + 1
                Else
                    retString = StrConv(retBytes, vbUnicode)
                End If
                
                'Pass the string to the target layer, which will read the XML data and initialize itself accordingly
                If Not dstImage.getLayerByIndex(i).CreateNewLayerFromXML(retString, , True) Then
                    Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
                End If
                
            Else
                Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
            End If
            
            'Normally we would load the layer's DIB data here, but we don't care about that when loading just the headers!
            ' Continue to the next layer.
        
        Next i
        
        'That's all there is to it!  Mark the load as successful and carry on.
        LoadPhotoDemonImageHeaderOnly = True
    
    Else
    
        'If we made it to this block, the first stage of PDI validation failed, meaning this file isn't in PDI format.
        Message "Selected file is not in PDI format.  Load abandoned."
        LoadPhotoDemonImageHeaderOnly = False
    
    End If
    
    Exit Function
    
LoadPDIHeaderFail:

    Select Case Err.Number
    
        Case PDP_GENERIC_ERROR
            Message "PDI node could not be read; file may be invalid or corrupted.  Load abandoned."
            
        Case Else

        Message "An error has occurred (#" & Err.Number & " - " & Err.Description & ").  PDI load abandoned."
        
    End Select
    
    LoadPhotoDemonImageHeaderOnly = False
    Exit Function

End Function

'Load a single layer from a standard PDI file.
' At present, this function is only used internally by the Undo/Redo engine.  If the nearest diff to a layer-specific change is a
' full pdImage stack, this function is used to extract only the relevant layer (or layer header) from the PDI file.
Public Function LoadSingleLayerFromPDI(ByVal PDIPath As String, ByRef dstLayer As pdLayer, ByVal targetLayerID As Long, Optional ByVal loadHeaderOnly As Boolean = False) As Boolean
    
    On Error GoTo LoadLayerFromPDIFail
    
    'First things first: create a pdPackage instance.  It will handle all the messy business of extracting individual data bits
    ' from the source file.
    Dim pdiReader As pdPackager
    Set pdiReader = New pdPackager
    pdiReader.init_ZLib "", True, g_ZLibEnabled
    
    'Load the file into the pdPackager instance.  It will cache the file contents, so we only have to do this once.
    ' Note that this step will also validate the incoming file.
    If pdiReader.readPackageFromFile(PDIPath, PD_IMAGE_IDENTIFIER) Then
    
        'PDI files all follow a standard format: a pdImage node at the top, which contains the full pdImage header,
        ' followed by individual nodes for each layer.  Layers are stored in stack order, which makes it very fast and easy
        ' to reconstruct the layer stack.
        
        'Unfortunately, stack order is not helpful in this function, because the target layer's position may have changed
        ' since the time this pdImage file was created.  To work around that, we must located the layer using its cardinal
        ' ID value, which is helpfully stored as the node ID parameter for a given layer node.
        
        Dim retBytes() As Byte, retString As String
        
        If pdiReader.getNodeDataByID(targetLayerID, True, retBytes, True) Then
        
            'Copy the received bytes into a string
            retString = Space$((UBound(retBytes) + 1) \ 2)
            CopyMemory ByVal StrPtr(retString), ByVal VarPtr(retBytes(0)), UBound(retBytes) + 1
            
            'Pass the string to the target layer, which will read the XML data and initialize itself accordingly.
            ' Note that we also pass along the loadHeaderOnly flag, which will instruct the layer to erase its current
            ' DIB as necessary.
            If Not dstLayer.CreateNewLayerFromXML(retString, , loadHeaderOnly) Then
                Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
            End If
        
        'Bytes could not be read, or alternately, checksums didn't match for the first node.
        Else
            Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
        End If
        
        'If this is not a header-only operation, repeat the above steps, but for the layer DIB this time
        If Not loadHeaderOnly Then
        
            'How we extract this data varies by layer type.  Raster layers can skip the need for a temporary buffer, because we've
            ' already created a DIB with a built-in buffer for the pixel data.
            '
            'Other layer types (e.g. vector layers) are tiny so a temporary buffer does not matter; also, unlike raster buffers, we cannot
            ' easily predict the buffer size in advance, so we rely on pdPackage to do it for us
            Dim nodeLoadedSuccessfully As Boolean
            nodeLoadedSuccessfully = False
            
            'Image (raster) layers
            If dstLayer.isLayerRaster Then
                
                'We are going to load the node data directly into the DIB, completely bypassing the need for a temporary array.
                Dim tmpDIBPointer As Long, tmpDIBLength As Long
                dstLayer.layerDIB.retrieveDIBPointerAndSize tmpDIBPointer, tmpDIBLength
                
                nodeLoadedSuccessfully = pdiReader.getNodeDataByID_UnsafeDstPointer(targetLayerID, False, tmpDIBPointer, True)
                    
            'Text and other vector layers
            ElseIf dstLayer.isLayerVector Then
                
                If pdiReader.getNodeDataByID(targetLayerID, False, retBytes, True) Then
                
                    'Convert the byte array to a Unicode string.  Note that we do not need an ASCII branch for old versions,
                    ' as vector layers were implemented after pdPackager was given Unicode compatibility.
                    retString = Space$((UBound(retBytes) + 1) \ 2)
                    CopyMemory ByVal StrPtr(retString), ByVal VarPtr(retBytes(0)), UBound(retBytes) + 1
                    
                    'Pass the string to the target layer, which will read the XML data and initialize itself accordingly
                    If dstLayer.CreateVectorDataFromXML(retString) Then
                        nodeLoadedSuccessfully = True
                    Else
                        Err.Raise PDP_GENERIC_ERROR, , "PDI Node (vector type) could not be read; data invalid or checksums did not match."
                    End If
                
                Else
                    Err.Raise PDP_GENERIC_ERROR, , "PDI Node (vector type) could not be read; data invalid or checksums did not match."
                End If
            
            'In the future, additional layer types can be handled here
            Else
                Debug.Print "WARNING! Unknown layer type exists in this PDI file: " & dstLayer.getLayerType
            
            End If
                
            'If successful, notify the target layer that its DIB data has been changed; the layer will use this to regenerate various internal caches
            If nodeLoadedSuccessfully Then
                dstLayer.notifyOfDestructiveChanges
                
            'Bytes could not be read, or alternately, checksums didn't match for the first node.
            Else
                Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
            End If
                
        End If
        
        'That's all there is to it!  Mark the load as successful and carry on.
        LoadSingleLayerFromPDI = True
    
    Else
    
        'If we made it to this block, the first stage of PDI validation failed, meaning this file isn't in PDI format.
        Message "Selected file is not in PDI format.  Load abandoned."
        LoadSingleLayerFromPDI = False
    
    End If
    
    Exit Function
    
LoadLayerFromPDIFail:

    Select Case Err.Number
    
        Case PDP_GENERIC_ERROR
            Message "PDI node could not be read; file may be invalid or corrupted.  Load abandoned."
            
        Case Else

        Message "An error has occurred (#" & Err.Number & " - " & Err.Description & ").  PDI load abandoned."
        
    End Select
    
    LoadSingleLayerFromPDI = False
    Exit Function

End Function

'Load a single PhotoDemon layer from a standalone pdLayer file (which is really just a modified PDI file).
' At present, this function is only used internally by the Undo/Redo engine.  Its counterpart is SavePhotoDemonLayer in
' the Saving module; any changes there should be mirrored here.
Public Function LoadPhotoDemonLayer(ByVal PDIPath As String, ByRef dstLayer As pdLayer, Optional ByVal loadHeaderOnly As Boolean = False) As Boolean
    
    On Error GoTo LoadPDLayerFail
    
    'First things first: create a pdPackage instance.  It will handle all the messy business of extracting individual data bits
    ' from the source file.
    Dim pdiReader As pdPackager
    Set pdiReader = New pdPackager
    pdiReader.init_ZLib "", True, g_ZLibEnabled
    
    'Load the file into the pdPackager instance.  pdPackager It will cache the file contents, so we only have to do this once.
    ' Note that this step will also validate the incoming file.
    If pdiReader.readPackageFromFile(PDIPath, PD_LAYER_IDENTIFIER) Then
    
        'Layer variants of PDI files contain a single node.  The layer's header is stored to the node's header chunk
        ' (in XML format, as expected).  The layer's DIB data is stored to the node's data chunk (in binary format, as expected).
        'First things first: extract the pdImage header, which will be in Node 0.  (We could double-check this by searching
        ' for the node entry by name, but since there is no variation, it's faster to access it directly.)
        Dim retBytes() As Byte, retString As String
        
        If pdiReader.getNodeDataByIndex(0, True, retBytes, True) Then
        
            'Copy the received bytes into a string
            retString = Space$((UBound(retBytes) + 1) \ 2)
            CopyMemory ByVal StrPtr(retString), ByVal VarPtr(retBytes(0)), UBound(retBytes) + 1
            
            'Pass the string to the target layer, which will read the XML data and initialize itself accordingly.
            ' Note that we pass the loadHeaderOnly request to this function; if this is a header-only load, the target
            ' layer must retain its current DIB.  This functionality is used by PD's Undo/Redo engine.
            dstLayer.CreateNewLayerFromXML retString, , loadHeaderOnly
            
        'Bytes could not be read, or alternately, checksums didn't match.  (Note that checksums are currently disabled
        ' for this function, for performance reasons, but I'm leaving this check in case we someday decide to re-enable them.)
        Else
            Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
        End If
        
        'Unless a header-only load was requested, we will now repeat the steps above, but for layer-specific data
        ' (a raw DIB stream for raster layers, or an XML string for vector/text layers)
        If Not loadHeaderOnly Then
        
            'How we extract this data varies by layer type.  Raster layers can skip the need for a temporary buffer, because we've
            ' already created a DIB with a built-in buffer for the pixel data.
            '
            'Other layer types (e.g. vector layers) are tiny so a temporary buffer does not matter; also, unlike raster buffers, we cannot
            ' easily predict the buffer size in advance, so we rely on pdPackage to do it for us
            Dim nodeLoadedSuccessfully As Boolean
            nodeLoadedSuccessfully = False
            
            'Image (raster) layers
            If dstLayer.isLayerRaster Then
                
                'We are going to load the node data directly into the DIB, completely bypassing the need for a temporary array.
                Dim tmpDIBPointer As Long, tmpDIBLength As Long
                dstLayer.layerDIB.retrieveDIBPointerAndSize tmpDIBPointer, tmpDIBLength
                
                nodeLoadedSuccessfully = pdiReader.getNodeDataByIndex_UnsafeDstPointer(0, False, tmpDIBPointer, True)
                
            'Text and other vector layers
            ElseIf dstLayer.isLayerVector Then
                
                If pdiReader.getNodeDataByIndex(0, False, retBytes, True) Then
                
                    'Convert the byte array to a Unicode string.  Note that we do not need an ASCII branch for old versions,
                    ' as vector layers were implemented after pdPackager was given Unicode compatibility.
                    retString = Space$((UBound(retBytes) + 1) \ 2)
                    CopyMemory ByVal StrPtr(retString), ByVal VarPtr(retBytes(0)), UBound(retBytes) + 1
                    
                    'Pass the string to the target layer, which will read the XML data and initialize itself accordingly
                    If dstLayer.CreateVectorDataFromXML(retString) Then
                        nodeLoadedSuccessfully = True
                    Else
                        Err.Raise PDP_GENERIC_ERROR, , "PDI Node (vector type) could not be read; data invalid or checksums did not match."
                    End If
                    
                Else
                    Err.Raise PDP_GENERIC_ERROR, , "PDI Node (vector type) could not be read; data invalid or checksums did not match."
                End If
            
            'In the future, additional layer types can be handled here
            Else
                Debug.Print "WARNING! Unknown layer type exists in this PDI file: " & dstLayer.getLayerType
            
            End If
                
            'If the load was successful, notify the target layer that its DIB data has been changed; the layer will use this to
            ' regenerate various internal caches.
            If nodeLoadedSuccessfully Then
                dstLayer.notifyOfDestructiveChanges
                
            'Failure means package bytes could not be read, or alternately, checksums didn't match.  (Note that checksums are currently
            ' disabled for this function, for performance reasons, but I'm leaving this check in case we someday decide to re-enable them.)
            Else
                Err.Raise PDP_GENERIC_ERROR, , "PDI Node could not be read; data invalid or checksums did not match."
            End If
            
        End If
        
        'That's all there is to it!  Mark the load as successful and carry on.
        LoadPhotoDemonLayer = True
    
    Else
    
        'If we made it to this block, the first stage of PDI validation failed, meaning this file isn't in PDI format.
        Message "Selected file is not in PDI format.  Load abandoned."
        LoadPhotoDemonLayer = False
    
    End If
    
    Exit Function
    
LoadPDLayerFail:

    Select Case Err.Number
    
        Case PDP_GENERIC_ERROR
            Message "PDI node could not be read; file may be invalid or corrupted.  Load abandoned."
            
        Case Else

        Message "An error has occurred (#" & Err.Number & " - " & Err.Description & ").  PDI load abandoned."
        
    End Select
    
    LoadPhotoDemonLayer = False
    Exit Function

End Function

'Use GDI+ to load an image.  This does very minimal error checking (which is a no-no with GDI+) but because it's only a
' fallback when FreeImage can't be found, I'm postponing further debugging for now.
'Used for PNG and TIFF files if FreeImage cannot be located.
Public Function LoadGDIPlusImage(ByVal imagePath As String, ByRef dstDIB As pdDIB) As Boolean
            
    Dim verifyGDISuccess As Boolean
    
    verifyGDISuccess = GDIPlusLoadPicture(imagePath, dstDIB)
    
    If verifyGDISuccess And (dstDIB.getDIBWidth <> 0) And (dstDIB.getDIBHeight <> 0) Then
        LoadGDIPlusImage = True
    Else
        LoadGDIPlusImage = False
    End If
    
End Function

'BITMAP loading
Public Function LoadVBImage(ByVal imagePath As String, ByRef dstDIB As pdDIB) As Boolean
    
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
    dstDIB.CreateFromPicture tmpPicture
    
    LoadVBImage = True
    Exit Function
    
LoadVBImageFail:

    LoadVBImage = False
    Exit Function
    
End Function

'Load data from a PD-generated Undo file.  This function is fairly complex, on account of PD's new diff-based Undo engine.
' Note that two types of Undo data must be specified: the Undo type of the file requested (because this function has no
' knowledge of that, by design), and what type of Undo data the caller wants extracted from the file.
'
'New as of 11 July '14 is the ability to specify a custom layer destination, for layer-relevant load operations.  If this value is NOTHING,
' the function will automatically load the data to the relevant layer in the parent pdImage object.  If this layer is supplied, however,
' the supplied layer reference will be used instead.
Public Sub LoadUndo(ByVal undoFile As String, ByVal undoTypeOfFile As Long, ByVal undoTypeOfAction As Long, Optional ByVal targetLayerID As Long = -1, Optional ByVal suspendRedraw As Boolean = False, Optional ByRef customLayerDestination As pdLayer = Nothing)
    
    'Certain load functions require access to a DIB, so declare a generic one in advance
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    'If selection data was loaded as part of this diff, this value will be set to TRUE.  We check it at the end of
    ' the load function, and activate various selection-related items as necessary.
    Dim selectionDataLoaded As Boolean
    selectionDataLoaded = False
    
    'Depending on the Undo data requested, we may end up loading one or more diff files at this location
    Select Case undoTypeOfAction
    
        'UNDO_EVERYTHING: a full copy of both the pdImage stack and all selection data is wanted
        Case UNDO_EVERYTHING
            Loading.LoadPhotoDemonImage undoFile, tmpDIB, pdImages(g_CurrentImage), True
            pdImages(g_CurrentImage).mainSelection.readSelectionFromFile undoFile & ".selection"
            selectionDataLoaded = True
            
        'UNDO_IMAGE, UNDO_IMAGE_VECTORSAFE: a full copy of the pdImage stack is wanted
        '             Because the underlying file data must be of type UNDO_EVERYTHING or UNDO_IMAGE/_VECTORSAFE, we
        '             don't have to do any special processing to the file - just load the whole damn thing.
        Case UNDO_IMAGE, UNDO_IMAGE_VECTORSAFE
            Loading.LoadPhotoDemonImage undoFile, tmpDIB, pdImages(g_CurrentImage), True
            
            'Once the full image has been loaded, we now know that at least the *existence* of all layers is correct.
            ' Unfortunately, subsequent changes to the pdImage header (or individual layers/layer headers) still need
            ' to be manually reconstructed, because they may have changed between the last full pdImage write and the
            ' current image state.  This step is handled by the Undo/Redo engine, which will call this LoadUndo function
            ' as many times as necessary to reconstruct each individual layer against its most recent diff.
        
        'UNDO_IMAGEHEADER: a full copy of the pdImage stack is wanted, but with all DIB data ignored (if present)
        '             For UNDO_IMAGEHEADER requests, we know the underlying file data is a PDI file.  We don't actually
        '             care if it has DIB data or not, because we'll just ignore it - but a special load function is
        '             required, due to the messy business of non-destructively aligning the current layer stack with
        '             the layer stack described by the file.
        Case UNDO_IMAGEHEADER
            Loading.LoadPhotoDemonImageHeaderOnly undoFile, pdImages(g_CurrentImage)
            
            'Once the full image has been loaded, we now know that at least the *existence* of all layers is correct.
            ' Unfortunately, subsequent changes to the pdImage header (or individual layers/layer headers) still need
            ' to be manually reconstructed, because they may have changed between the last full pdImage write and the
            ' current image state.  This step is handled by the Undo/Redo engine, which will call this LoadUndo function
            ' as many times as necessary to reconstruct each individual layer against its most recent diff.
        
        'UNDO_LAYER, UNDO_LAYER_VECTORSAFE: a full copy of the saved layer data at this position.
        '             Because the underlying file data can be different types (layer data can be loaded from standalone layer saves,
        '             or from a full pdImage stack save), we must check the undo type of the saved file, and modify our load
        '             behavior accordingly.
        Case UNDO_LAYER, UNDO_LAYER_VECTORSAFE
            
            'New as of 11 July '14 is the ability for the caller to supply their own destination layer for layer-specific Undo data.
            ' Check this optional parameter, and if it is NOT supplied, point it at the relevant layer in the parent pdImage object.
            If (customLayerDestination Is Nothing) Then Set customLayerDestination = pdImages(g_CurrentImage).getLayerByID(targetLayerID)
            
            'Layer data can appear in multiple types of Undo files
            Select Case undoTypeOfFile
            
                'The underlying save file is a standalone layer entry.  Simply overwrite the target layer with the data from the file.
                Case UNDO_LAYER, UNDO_LAYER_VECTORSAFE
                    Loading.LoadPhotoDemonLayer undoFile & ".layer", customLayerDestination, False
            
                'The underlying save file is a full pdImage stack.  Extract only the relevant layer data from the stack.
                Case UNDO_EVERYTHING, UNDO_IMAGE, UNDO_IMAGE_VECTORSAFE
                    Loading.LoadSingleLayerFromPDI undoFile, customLayerDestination, targetLayerID, False
                
            End Select
        
        'UNDO_LAYERHEADER: a full copy of the saved layer header data at this position.  Layer DIB data is ignored.
        '             Because the underlying file data can be many different types (layer data header can be loaded from
        '             standalone layer header saves, or full layer saves, or even a full pdImage stack), we must check the
        '             undo type of the saved file, and modify our load behavior accordingly.
        Case UNDO_LAYERHEADER
            
            'Layer header data can appear in multiple types of Undo files
            Select Case undoTypeOfFile
            
                'The underlying save file is a standalone layer entry.  Simply overwrite the target layer header with the
                ' header data from this file.
                Case UNDO_LAYER, UNDO_LAYER_VECTORSAFE, UNDO_LAYERHEADER
                    Loading.LoadPhotoDemonLayer undoFile & ".layer", pdImages(g_CurrentImage).getLayerByID(targetLayerID), True
            
                'The underlying save file is a full pdImage stack.  Extract only the relevant layer data from the stack.
                Case UNDO_EVERYTHING, UNDO_IMAGE, UNDO_IMAGE_VECTORSAFE, UNDO_IMAGEHEADER
                    Loading.LoadSingleLayerFromPDI undoFile, pdImages(g_CurrentImage).getLayerByID(targetLayerID), targetLayerID, True
                
            End Select
        
        'UNDO_SELECTION: a full copy of the saved selection data is wanted
        '                 Because the underlying file data must be of type UNDO_EVERYTHING or UNDO_SELECTION, we don't have to do
        '                 any special processing.
        Case UNDO_SELECTION
            pdImages(g_CurrentImage).mainSelection.readSelectionFromFile undoFile & ".selection"
            selectionDataLoaded = True
            
            
        'For now, any unhandled Undo types result in a request for the full pdImage stack.  This line can be removed when
        ' all Undo types finally have their own custom handling implemented.
        Case Else
            Loading.LoadPhotoDemonImage undoFile, tmpDIB, pdImages(g_CurrentImage), True
            
        
    End Select
    
    'If a selection was loaded, activate all selection-related stuff now
    If selectionDataLoaded Then
    
        'Activate the selection as necessary
        pdImages(g_CurrentImage).selectionActive = pdImages(g_CurrentImage).mainSelection.isLockedIn
        
        'Synchronize the text boxes as necessary
        syncTextToCurrentSelection g_CurrentImage
    
    End If
    
    'If a selection is active, request a redraw of the selection mask before rendering the image to the screen.  (If we are
    ' "undoing" an action that changed the image's size, the selection mask will be out of date.  Thus we need to re-render
    ' it before rendering the image or OOB errors may occur.)
    If pdImages(g_CurrentImage).selectionActive Then pdImages(g_CurrentImage).mainSelection.requestNewMask
        
    'Render the image to the screen, if requested
    If Not suspendRedraw Then Viewport_Engine.Stage1_InitializeBuffer pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

'Load a raw pdDIB file dump into the destination image and DIB.  (Note that pdDIB may have applied zLib compression during the save,
' depending on the parameters it was passed, so it is possible for this function to fail if zLib goes missing.)
Public Function LoadRawImageBuffer(ByVal imagePath As String, ByRef dstDIB As pdDIB, ByRef dstImage As pdImage) As Boolean

    On Error GoTo LoadRawImageBufferFail
    
    'Ask the destination DIB to create itself using the raw image buffer data
    LoadRawImageBuffer = dstDIB.CreateFromFile(imagePath)
    
    Exit Function
    
LoadRawImageBufferFail:

    LoadRawImageBuffer = False
    Exit Function

End Function

'This routine sets the message on the splash screen (used only when the program is first started)
Public Sub LoadMessage(ByVal sMsg As String)
    
    Static loadProgress As Long
        
    'In debug mode, mirror message output to PD's central Debugger
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction sMsg, PDM_USER_MESSAGE
    #End If
    
    'Load messages are translatable, but we don't want to translate them if the translation object isn't ready yet
    If (Not (g_Language Is Nothing)) Then
        If g_Language.readyToTranslate Then
            If g_Language.translationActive Then sMsg = g_Language.TranslateMessage(sMsg)
        End If
    End If
    
    'Previously, the current load text would be displayed to the user at this point.  As of version 6.6, this step is skipped in favor
    ' of a more minimalist splash screen.
    ' TODO BY 6.8's RELEASE: revisit this function entirely, and consider removing it if applicable
    If FormSplash.Visible Then FormSplash.updateLoadProgress loadProgress
    
    loadProgress = loadProgress + 1
    
End Sub

'Loading all hotkeys (accelerators) requires a few different things.  Besides just populating the hotkey collection, we also paint all
' menu captions to match.
Public Sub LoadAccelerators()
    
    With FormMain.pdHotkeys
    
        .Enabled = True
    
        'File menu
        .AddAccelerator vbKeyN, vbCtrlMask, "New image", FormMain.MnuFile(0), True, False, True, UNDO_NOTHING
        .AddAccelerator vbKeyO, vbCtrlMask, "Open", FormMain.MnuFile(1), True, False, True, UNDO_NOTHING
        .AddAccelerator vbKeyF4, vbCtrlMask, "Close", FormMain.MnuFile(5), True, True, True, UNDO_NOTHING
        .AddAccelerator vbKeyF4, vbCtrlMask Or vbShiftMask, "Close all", FormMain.MnuFile(6), True, True, True, UNDO_NOTHING
        .AddAccelerator vbKeyS, vbCtrlMask, "Save", FormMain.MnuFile(8), True, True, True, UNDO_NOTHING
        .AddAccelerator vbKeyS, vbCtrlMask Or vbAltMask Or vbShiftMask, "Save copy", FormMain.MnuFile(9), True, False, True, UNDO_NOTHING
        .AddAccelerator vbKeyS, vbCtrlMask Or vbShiftMask, "Save as", FormMain.MnuFile(10), True, True, True, UNDO_NOTHING
        .AddAccelerator vbKeyF12, 0, "Revert", FormMain.MnuFile(11), True, True, False, UNDO_NOTHING
        .AddAccelerator vbKeyB, vbCtrlMask, "Batch wizard", FormMain.MnuFile(13), True, True, True, UNDO_NOTHING
        .AddAccelerator vbKeyP, vbCtrlMask, "Print", FormMain.MnuFile(15), True, True, True, UNDO_NOTHING
        .AddAccelerator vbKeyQ, vbCtrlMask, "Exit program", FormMain.MnuFile(17), True, False, True, UNDO_NOTHING
        
            'File -> Import submenu
            .AddAccelerator vbKeyI, vbCtrlMask Or vbShiftMask Or vbAltMask, "Scan image", FormMain.MnuScanImage, True, False, True, UNDO_NOTHING
            .AddAccelerator vbKeyD, vbCtrlMask Or vbShiftMask, "Internet import", FormMain.MnuImportFromInternet, True, True, True, UNDO_NOTHING
            .AddAccelerator vbKeyI, vbCtrlMask Or vbAltMask, "Screen capture", FormMain.MnuScreenCapture, True, True, True, UNDO_NOTHING
        
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
        .AddAccelerator vbKeyZ, vbCtrlMask, "Undo", FormMain.MnuEdit(0), True, True, False, UNDO_NOTHING
        .AddAccelerator vbKeyY, vbCtrlMask, "Redo", FormMain.MnuEdit(1), True, True, False, UNDO_NOTHING
        
        .AddAccelerator vbKeyF, vbCtrlMask, "Repeat last action", FormMain.MnuEdit(4), True, True, False, UNDO_IMAGE
        
        .AddAccelerator vbKeyX, vbCtrlMask, "Cut", FormMain.MnuEdit(7), True, True, False, UNDO_IMAGE
        .AddAccelerator vbKeyX, vbCtrlMask Or vbShiftMask, "Cut from layer", FormMain.MnuEdit(8), True, True, False, UNDO_LAYER
        .AddAccelerator vbKeyC, vbCtrlMask, "Copy", FormMain.MnuEdit(9), True, True, False, UNDO_NOTHING
        .AddAccelerator vbKeyC, vbCtrlMask Or vbShiftMask, "Copy from layer", FormMain.MnuEdit(10), True, True, False, UNDO_NOTHING
        .AddAccelerator vbKeyV, vbCtrlMask, "Paste as new image", FormMain.MnuEdit(11), True, False, False, UNDO_NOTHING
        .AddAccelerator vbKeyV, vbCtrlMask Or vbShiftMask, "Paste as new layer", FormMain.MnuEdit(12), True, False, False, UNDO_IMAGE_VECTORSAFE
        
        'View menu
        .AddAccelerator vbKey0, 0, "FitOnScreen", FormMain.MnuFitOnScreen, False, True, False, UNDO_NOTHING
        .AddAccelerator vbKeyAdd, 0, "Zoom_In", FormMain.MnuZoomIn, False, True, False, UNDO_NOTHING
        .AddAccelerator vbKeySubtract, 0, "Zoom_Out", FormMain.MnuZoomOut, False, True, False, UNDO_NOTHING
        .AddAccelerator vbKey5, 0, "Zoom_161", FormMain.MnuSpecificZoom(0), False, True, False, UNDO_NOTHING
        .AddAccelerator vbKey4, 0, "Zoom_81", FormMain.MnuSpecificZoom(1), False, True, False, UNDO_NOTHING
        .AddAccelerator vbKey3, 0, "Zoom_41", FormMain.MnuSpecificZoom(2), False, True, False, UNDO_NOTHING
        .AddAccelerator vbKey2, 0, "Zoom_21", FormMain.MnuSpecificZoom(3), False, True, False, UNDO_NOTHING
        .AddAccelerator vbKey1, 0, "Actual_Size", FormMain.MnuSpecificZoom(4), False, True, False, UNDO_NOTHING
        .AddAccelerator vbKey2, vbShiftMask, "Zoom_12", FormMain.MnuSpecificZoom(5), False, True, False, UNDO_NOTHING
        .AddAccelerator vbKey3, vbShiftMask, "Zoom_14", FormMain.MnuSpecificZoom(6), False, True, False, UNDO_NOTHING
        .AddAccelerator vbKey4, vbShiftMask, "Zoom_18", FormMain.MnuSpecificZoom(7), False, True, False, UNDO_NOTHING
        .AddAccelerator vbKey5, vbShiftMask, "Zoom_116", FormMain.MnuSpecificZoom(8), False, True, False, UNDO_NOTHING
        
        'Image menu
        .AddAccelerator vbKeyA, vbCtrlMask Or vbShiftMask, "Duplicate image", FormMain.MnuImage(0), True, True, False, UNDO_NOTHING
        .AddAccelerator vbKeyR, vbCtrlMask, "Resize image", FormMain.MnuImage(2), True, True, True, UNDO_IMAGE
        .AddAccelerator vbKeyR, vbCtrlMask Or vbAltMask, "Canvas size", FormMain.MnuImage(5), True, True, True, UNDO_IMAGEHEADER
        '.AddAccelerator vbKeyX, vbCtrlMask Or vbShiftMask, "Crop", FormMain.MnuImage(8), True, True, False, UNDO_IMAGE
        .AddAccelerator vbKeyX, vbCtrlMask Or vbAltMask, "Trim empty borders", FormMain.MnuImage(10), True, True, False, UNDO_IMAGEHEADER
        'KeyCode 188 = <,  (next to the letter M)
        .AddAccelerator 188, vbCtrlMask Or vbAltMask, "Reduce colors", FormMain.MnuImage(16), True, True, False, UNDO_IMAGE
        
            'Image -> Rotate submenu
            .AddAccelerator vbKeyR, 0, "Rotate image 90 clockwise", FormMain.MnuRotate(2), True, True, False, UNDO_IMAGE
            .AddAccelerator vbKeyL, 0, "Rotate image 90 counter-clockwise", FormMain.MnuRotate(3), True, True, False, UNDO_IMAGE
            .AddAccelerator vbKeyR, vbCtrlMask Or vbShiftMask Or vbAltMask, "Arbitrary image rotation", FormMain.MnuRotate(5), True, True, True, UNDO_NOTHING
        
        'Layer Menu
        '(none yet)
        
        
        'Select Menu
        .AddAccelerator vbKeyA, vbCtrlMask, "Select all", FormMain.MnuSelect(0), True, True, False, UNDO_SELECTION
        .AddAccelerator vbKeyD, vbCtrlMask, "Remove selection", FormMain.MnuSelect(1), False, True, False, UNDO_SELECTION
        .AddAccelerator vbKeyI, vbCtrlMask Or vbShiftMask, "Invert selection", FormMain.MnuSelect(2), True, True, False, UNDO_SELECTION
        'KeyCode 219 = {[  (next to the letter P), 221 = }]
        .AddAccelerator 221, vbCtrlMask Or vbAltMask, "Grow selection", FormMain.MnuSelect(4), True, True, True, UNDO_NOTHING
        .AddAccelerator 219, vbCtrlMask Or vbAltMask, "Shrink selection", FormMain.MnuSelect(5), True, True, True, UNDO_NOTHING
        .AddAccelerator vbKeyD, vbCtrlMask Or vbAltMask, "Feather selection", FormMain.MnuSelect(7), True, True, True, UNDO_NOTHING
        
        'Adjustments Menu
        
        'Adjustments top shortcut menu
        .AddAccelerator vbKeyU, vbCtrlMask Or vbShiftMask, "Black and white", FormMain.MnuAdjustments(3), True, True, True, UNDO_NOTHING
        .AddAccelerator vbKeyB, vbCtrlMask Or vbShiftMask, "Brightness and contrast", FormMain.MnuAdjustments(4), True, True, True, UNDO_NOTHING
        .AddAccelerator vbKeyC, vbCtrlMask Or vbAltMask, "Color balance", FormMain.MnuAdjustments(5), True, True, True, UNDO_NOTHING
        .AddAccelerator vbKeyM, vbCtrlMask, "Curves", FormMain.MnuAdjustments(6), True, True, True, UNDO_NOTHING
        .AddAccelerator vbKeyL, vbCtrlMask, "Levels", FormMain.MnuAdjustments(7), True, True, True, UNDO_NOTHING
        .AddAccelerator vbKeyH, vbCtrlMask Or vbShiftMask, "Shadow and highlight", FormMain.MnuAdjustments(8), True, True, True, UNDO_NOTHING
        .AddAccelerator vbKeyAdd, vbCtrlMask Or vbAltMask, "Vibrance", FormMain.MnuAdjustments(9), True, True, True, UNDO_NOTHING
        .AddAccelerator vbKeyW, vbCtrlMask, "White balance", FormMain.MnuAdjustments(10), True, True, True, UNDO_NOTHING
        
            'Color adjustments
            .AddAccelerator vbKeyH, vbCtrlMask, "Hue and saturation", FormMain.MnuColor(3), True, True, True, UNDO_NOTHING
            .AddAccelerator vbKeyT, vbCtrlMask, "Temperature", FormMain.MnuColor(4), True, True, True, UNDO_NOTHING
            
            'Lighting adjustments
            .AddAccelerator vbKeyG, vbCtrlMask, "Gamma", FormMain.MnuLighting(2), True, True, True, UNDO_NOTHING
            
            'Adjustments -> Invert submenu
            .AddAccelerator vbKeyI, vbCtrlMask, "Invert RGB", FormMain.mnuInvert, True, True, False, UNDO_LAYER
            
            'Adjustments -> Monochrome submenu
            .AddAccelerator vbKeyB, vbCtrlMask Or vbAltMask Or vbShiftMask, "Color to monochrome", FormMain.MnuMonochrome(0), True, True, True, UNDO_NOTHING
            
            'Adjustments -> Photography submenu
            .AddAccelerator vbKeyE, vbCtrlMask Or vbAltMask, "Exposure", FormMain.MnuAdjustmentsPhoto(0), True, True, True, UNDO_NOTHING
            .AddAccelerator vbKeyP, vbCtrlMask Or vbAltMask, "Photo filter", FormMain.MnuAdjustmentsPhoto(2), True, True, True, UNDO_NOTHING
            
        
        'Effects Menu
        '.AddAccelerator vbKeyZ, vbCtrlMask Or vbAltMask Or vbShiftMask, "Add RGB noise", FormMain.MnuNoise(1), True, True, True, False
        '.AddAccelerator vbKeyG, vbCtrlMask Or vbAltMask Or vbShiftMask, "Gaussian blur", FormMain.MnuBlurFilter(1), True, True, True, False
        '.AddAccelerator vbKeyY, vbCtrlMask Or vbAltMask Or vbShiftMask, "Correct lens distortion", FormMain.MnuDistortEffects(1), True, True, True, False
        '.AddAccelerator vbKeyU, vbCtrlMask Or vbAltMask Or vbShiftMask, "Unsharp mask", FormMain.MnuSharpen(1), True, True, True, False
        
        'Tools menu
        'KeyCode 190 = >.  (two keys to the right of the M letter key)
        .AddAccelerator 190, vbCtrlMask Or vbAltMask, "Play macro", FormMain.mnuTool(4), True, True, True, UNDO_NOTHING
        
        .AddAccelerator vbKeyReturn, vbAltMask, "Preferences", FormMain.mnuTool(7), False, False, True, UNDO_NOTHING
        .AddAccelerator vbKeyM, vbCtrlMask Or vbAltMask, "Plugin manager", FormMain.mnuTool(8), False, False, True, UNDO_NOTHING
        
        
        'Window menu
        .AddAccelerator vbKeyPageDown, 0, "Next_Image", FormMain.MnuWindow(5), False, True, False, UNDO_NOTHING
        .AddAccelerator vbKeyPageUp, 0, "Prev_Image", FormMain.MnuWindow(6), False, True, False, UNDO_NOTHING
        
        'Activate hotkey detection
        .ActivateHook
        
    End With
    
    'Before exiting, paint all shortcut captions to their respective menus
    DrawAccelerators
    
End Sub

'After all menu shortcuts (accelerators) are loaded above, the custom shortcuts need to be painted to their corresponding menus
Public Sub DrawAccelerators()

    Dim i As Long
    
    With FormMain.pdHotkeys
        For i = 0 To .Count - 1
            If .HasMenu(i) Then
                .MenuReference(i).Caption = .MenuReference(i).Caption & vbTab & .StringRepresentation(i)
            End If
        Next i
    End With

    'A few menu shortcuts must be drawn manually.
    
    'Because the Import -> From Clipboard menu shares the same shortcut as Edit -> Paste as new image, we must
    ' manually add its shortcut (as only the Edit -> Paste will be handled automatically).
    FormMain.MnuImportClipboard.Caption = FormMain.MnuImportClipboard.Caption & vbTab & g_Language.TranslateMessage("Ctrl") & "+V"
    
    'Similarly for the Layer -> New -> From Clipboard menu
    FormMain.MnuLayerNew(3).Caption = FormMain.MnuLayerNew(3).Caption & vbTab & g_Language.TranslateMessage("Ctrl") & "+" & g_Language.TranslateMessage("Shift") & "+V"
    
    'NOTE: Drawing of MRU shortcuts is handled in the MRU module
    
End Sub

'Make a copy of the current image.  Thanks to PSC user "Achmad Junus" for this suggestion.
Public Sub DuplicateCurrentImage()
    
    Message "Duplicating current image..."
    
    'Ask the currently active image to write itself out to file
    Dim tmpDuplicationFile As String
    tmpDuplicationFile = g_UserPreferences.GetTempPath & "PDDuplicate.pdi"
    SavePhotoDemonImage pdImages(g_CurrentImage), tmpDuplicationFile, True, True, True, False
    
    'We can now use the standard image load routine to import the temporary file
    Dim sFile() As String, sTitle As String, sFilename As String
    ReDim sFile(0) As String
    sFile(0) = tmpDuplicationFile
    sTitle = pdImages(g_CurrentImage).originalFileName & " - " & g_Language.TranslateMessage("Copy")
    sFilename = sTitle
    
    LoadFileAsNewImage sFile, False, sTitle, sFilename
                    
    'Be polite and remove the temporary file
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    If cFile.FileExist(tmpDuplicationFile) Then cFile.KillFile tmpDuplicationFile
    
    Message "Image duplication complete."
        
End Sub

'Check for IDE or compiled EXE, and set program parameters accordingly
Private Sub CheckLoadingEnvironment()
    If App.logMode = 1 Then
        g_IsProgramCompiled = True
    Else
        g_IsProgramCompiled = False
    End If
End Sub
