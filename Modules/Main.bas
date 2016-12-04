Attribute VB_Name = "MainModule"
'Note: this file has been modified for use within PhotoDemon.

'This module is required for theming via embedded manifest.  Many thanks to LaVolpe for the automated tool that coincides
' with this fine piece of code.  Download it yourself at: http://www.vbforums.com/showthread.php?t=606736

Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32" ()
Private Declare Function InitCommonControlsEx Lib "comctl32" (iccex As InitCommonControlsExStruct) As Boolean
Private Type InitCommonControlsExStruct
    lngSize As Long
    lngICC As Long
End Type

'As of September 2015, reordering the list of files in the master VBP has caused unpredictable crashes when PD closes.
' I've spent two days bisecting commits and I can conclusively nail it down to
' https://github.com/tannerhelland/PhotoDemon/commit/293de1ba4f2d5bc3102304d0263af624e93b6093
'
'I eventually solved the problem by manually unloading all global class instances in a specific order, rather than
' leaving it to VB, but during testing, I found it helpful to suppress the default Windows crash dialog.  In case this
' ever proves useful in the future, I'll leave the declaration here.
Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Private Const SEM_FAILCRITICALERRORS = &H1
Private Const SEM_NOGPFAULTERRORBOX = &H2
Private Const SEM_NOOPENFILEERRORBOX = &H8000&

Private hShellModule As Long

'These can be used to ensure that the splash shows for a minimum amount of time
Private m_LoadTime As Double, m_StartTime As Double

'This constant is the number of "discrete" loading steps involved in loading the program.  It is relevant for displaying
'the progress bar on the initial splash screen; this value is the progress bar's maximum value.
Private Const NUMBER_OF_LOADING_STEPS As Long = 18

'PhotoDemon starts here.  Main() is necessary as a start point (vs a form) to make sure that theming is implemented
' correctly.  Note that this code is irrelevant within the IDE.
Public Sub Main()

    Dim iccex As InitCommonControlsExStruct
    
    'For descriptions of these constants, visit: http://msdn.microsoft.com/en-us/library/bb775507%28VS.85%29.aspx
    iccex.lngSize = LenB(iccex)
    Const ICC_BAR_CLASSES As Long = &H4&
    Const ICC_STANDARD_CLASSES As Long = &H4000&
    Const ICC_WIN95_CLASSES As Long = &HFF&
    iccex.lngICC = ICC_STANDARD_CLASSES Or ICC_BAR_CLASSES Or ICC_WIN95_CLASSES
    
    'InitCommonControlsEx requires IEv3 or above, which shouldn't be a problem on any modern system.  But just in case,
    ' continue loading even if the common control module load fails.
    On Error Resume Next
    
    'The following block of code prevents XP crashes when VB usercontrols are present in a project (as they are in PhotoDemon)
    Dim strShellName As String
    strShellName = "shell32.dll"
    hShellModule = LoadLibrary(StrPtr(strShellName))
    InitCommonControlsEx iccex
    
    'If an error occurs, attempt to initiate the Win9x version
    If Err Then
        InitCommonControls ' try Win9x version
        Err.Clear
    End If
    
    On Error GoTo 0
    
    'Because Ambient.UserMode does not report IDE behavior properly, we use our own UserMode tracker.
    ' Thank you to Kroc of camendesign.com for suggesting this workaround.
    g_IsProgramRunning = True
    
    'FormMain can now be loaded.  (We load it first, because many initialization steps silently interact with it,
    ' like loading menu icons or prepping toolboxes.)  That said, the first step of FormMain's load process is calling
    'the ContinueLoadingProgram sub, below, so look there for the next stages of the load process.
    Load FormMain
    
End Sub

'Note that this function is called AFTER FormMain has been loaded.  FormMain is loaded - but not visible - so it can be operated
' on by functions called from this routine.  (It is necessary to load the main window first, since a number of load operations -
' like decoding PNG menu icons from the resource file, then applying them to program menus - operate directly on the main window.)
Public Sub ContinueLoadingProgram()
    
    '*************************************************************************************************************************************
    ' Check the state of this build (alpha, beta, production, etc) and activate debug code as necessary
    '*************************************************************************************************************************************
    
    'Current build state is stored in the public const "PD_BUILD_QUALITY".  For non-production builds, a number of program-wide
    ' parameters are automatically set.
    
    'If the program is in pre-alpha or alpha state, enable timing reports.
    If (PD_BUILD_QUALITY = PD_PRE_ALPHA) Or (PD_BUILD_QUALITY = PD_ALPHA) Then g_DisplayTimingReports = True
    
    'Enable high-performance timer objects
    VB_Hacks.EnableHighResolutionTimers
    
    'Regardless of debug mode or not, we instantiate a pdDebug instance.  It will only be interacted with if the program is compiled
    ' with DEBUGMODE = 1, however.
    Set pdDebug = New pdDebugger
    
    'During development, I find it helpful to profile PhotoDemon's startup process.  Timing functions like this can be commented out
    ' without harming anything.
    Dim perfCheck As pdProfiler
    Set perfCheck = New pdProfiler
    
    #If DEBUGMODE = 1 Then
        perfCheck.StartProfiling "PhotoDemon Startup", True
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
        perfCheck.MarkEvent "Prepare splash screen"
    #End If
    
    m_StartTime = Timer
    
    'Before doing any 2D rendering, we need to start at least one valid 2D rendering backend.
    ' (At present, only GDI+ is used)
    Interface.InitializeInterfaceBackend
    
    If Drawing2D.StartRenderingEngine(P2_DefaultBackend) Then
        
        #If DEBUGMODE = 1 Then
            Drawing2D.SetLibraryDebugMode True
        #End If
        
        'Load FormSplash into memory, but don't make it visible.
        FormSplash.Visible = False
        
    End If
        
    'Check the environment.  If inside the the IDE, the splash needs to be modified slightly.
    CheckLoadingEnvironment
    
    If Drawing2D.IsRenderingEngineActive(P2_GDIPlusBackend) Then
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
        perfCheck.MarkEvent "Check Windows version"
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
    If g_IsWin7OrLater Then PrepWin7Features
    
    
    
    '*************************************************************************************************************************************
    ' If the user doesn't have font smoothing enabled, enable it now.  PD's interface looks much better with some form of antialiasing.
    '*************************************************************************************************************************************
    
    #If DEBUGMODE = 1 Then
        perfCheck.MarkEvent "ClearType check"
    #End If
    
    HandleClearType True
    
    
    
    '*************************************************************************************************************************************
    ' Initialize the user preferences (settings) handler
    '*************************************************************************************************************************************
    
    #If DEBUGMODE = 1 Then
        perfCheck.MarkEvent "Initialize preferences engine"
    #End If
    
    'Before initializing the preference engine, generate a unique session ID for this PhotoDemo instance.  This ID will be used to
    ' separate the temp files for this program instance from any other simultaneous instances.
    g_SessionID = cSysInfo.GetUniqueSessionID()
    
    Set g_UserPreferences = New pdPreferences
    
    'Ask the preferences handler to generate key program folders.  (If these folders don't exist, the handler will create them.)
    LoadMessage "Initializing all program directories..."
    
    g_UserPreferences.InitializePaths
        
    'Now, ask the preferences handler to load all other user settings from the preferences file
    LoadMessage "Loading all user settings..."
    
    g_UserPreferences.LoadUserSettings
        
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
    ImageImporter.ResetImageImportPreferenceCache
    
    
    '*************************************************************************************************************************************
    ' Initialize the translation (language) engine
    '*************************************************************************************************************************************
    
    #If DEBUGMODE = 1 Then
        perfCheck.MarkEvent "Initialize translation engine"
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
    ' Initialize the visual themes engine
    '*************************************************************************************************************************************
    
    #If DEBUGMODE = 1 Then
        perfCheck.MarkEvent "Initialize theme engine"
    #End If
    
    'Because this class controls the visual appearance of all forms in the project, it must be loaded early in the boot process
    LoadMessage "Initializing theme engine..."
    
    Set g_Themer = New pdVisualThemes
    
    'Load and validate the user's selected theme file
    g_Themer.LoadDefaultPDTheme
    
    'Now that a theme has been loaded, we can initialize additional UI rendering elements
    Drawing.CacheUIPensAndBrushes
    Paintbrush.InitializeBrushEngine
    
    '*************************************************************************************************************************************
    ' PhotoDemon works very well with multiple monitors.  Check for such a situation now.
    '*************************************************************************************************************************************
    
    #If DEBUGMODE = 1 Then
        perfCheck.MarkEvent "Detect monitors"
    #End If
    
    LoadMessage "Analyzing current monitor setup..."
    
    Set g_Displays = New pdDisplays
    g_Displays.RefreshDisplays
    
    'While here, also cache various display-related settings; this is faster than constantly retrieving them via APIs
    Interface.CacheSystemDPI g_Displays.GetWindowsDPI
    
    
    '*************************************************************************************************************************************
    ' Now we have what we need to properly display the splash screen.  Do so now.
    '*************************************************************************************************************************************
    
    #If DEBUGMODE = 1 Then
        perfCheck.MarkEvent "Display splash screen"
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
    Else
        g_InterfaceFont = "Tahoma"
    End If
    
    Set tmpFontCheck = Nothing
    
    'Ask the splash screen to finish whatever initializing it needs prior to displaying itself
    FormSplash.PrepareSplashLogo NUMBER_OF_LOADING_STEPS
    FormSplash.prepareRestOfSplash
    
    'Display the splash screen, centered on whichever monitor the user previously used the program on.
    FormSplash.Show vbModeless
        
    
    '*************************************************************************************************************************************
    ' Check for the presence of plugins (as other functions rely on these to initialize themselves)
    '*************************************************************************************************************************************
    
    #If DEBUGMODE = 1 Then
        perfCheck.MarkEvent "Load plugins"
    #End If
    
    LoadMessage "Loading plugins..."
    
    PluginManager.LoadAllPlugins
    
    
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
        perfCheck.MarkEvent "Load import and export libraries"
    #End If
    
    LoadMessage "Loading import/export libraries..."
    
    'The FreeImage.dll plugin provides most of PD's advanced image format support, but we can also fall back on GDI+.
    ' Prior to generating a list of supported formats, notify the image format class of GDI+ availability
    ' (which was determined earlier in this function, prior to loading the splash screen).
    g_ImageFormats.GDIPlusEnabled = Drawing2D.IsRenderingEngineActive(P2_GDIPlusBackend)
    
    'Generate a list of currently supported input/output formats, which may vary based on plugin version and availability
    g_ImageFormats.GenerateInputFormats
    g_ImageFormats.GenerateOutputFormats
    
    
    '*************************************************************************************************************************************
    ' Build a font cache for this system
    '*************************************************************************************************************************************
    
    #If DEBUGMODE = 1 Then
        perfCheck.MarkEvent "Build font cache"
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
        perfCheck.MarkEvent "Initialize pdClipboardMain"
    #End If
    
    LoadMessage "Initializing clipboard interface..."
    
    Set g_Clipboard = New pdClipboardMain
    
    
    '*************************************************************************************************************************************
    ' Get the viewport engine ready
    '*************************************************************************************************************************************
    
    #If DEBUGMODE = 1 Then
        perfCheck.MarkEvent "Initialize viewport engine"
    #End If
    
    'Initialize our current zoom method
    LoadMessage "Initializing viewport engine..."
    
    'Create the program's primary zoom handler
    Set g_Zoom = New pdZoom
    g_Zoom.InitializeViewportEngine
    
    'Populate the main form's zoom drop-down
    g_Zoom.PopulateZoomComboBox FormMain.mainCanvas(0).GetZoomDropDownReference()
    
    'Populate the main canvas's size unit dropdown
    FormMain.mainCanvas(0).PopulateSizeUnits
    
    
    '*************************************************************************************************************************************
    ' Initialize the window manager (the class that synchronizes all toolbox and image window positions)
    '*************************************************************************************************************************************
    
    #If DEBUGMODE = 1 Then
        perfCheck.MarkEvent "Initialize window manager"
    #End If
    
    LoadMessage "Initializing window manager..."
    Set g_WindowManager = New pdWindowManager
    
    'Register the main form
    g_WindowManager.SetAutoRefreshMode False
    g_WindowManager.RegisterMainForm FormMain
    
    'As of 7.0, all we need to do here is initialize the new, lightweight toolbox handler.  This will load things
    ' like toolbox sizes and visibility from the previous session.
    Toolboxes.LoadToolboxData
    
    'With toolbox data assembled, we can now silently load each tool window.  Even though these windows may not
    ' be visible (as the user can elect to hide them), we still want them loaded so that we can activate them quickly
    ' if/when they are enabled.
    
    #If DEBUGMODE = 1 Then
        perfCheck.MarkEvent "Window manager: load left toolbox"
    #End If
    Load toolbar_Toolbox
    
    #If DEBUGMODE = 1 Then
        perfCheck.MarkEvent "Window manager: load right toolbox"
    #End If
    Load toolbar_Layers
    
    #If DEBUGMODE = 1 Then
        perfCheck.MarkEvent "Window manager: load bottom toolbox"
    #End If
    Load toolbar_Options
    
    'Retrieve tool window visibility and mark those menus as well
    FormMain.MnuWindowToolbox(0).Checked = Toolboxes.GetToolboxVisibilityPreference(PDT_LeftToolbox)
    FormMain.MnuWindow(1).Checked = Toolboxes.GetToolboxVisibilityPreference(PDT_BottomToolbox)
    FormMain.MnuWindow(2).Checked = Toolboxes.GetToolboxVisibilityPreference(PDT_RightToolbox)
    
    'Retrieve two additional settings for the image tabstrip menu: when to display it, and its alignment
    ToggleImageTabstripVisibility g_UserPreferences.GetPref_Long("Core", "Image Tabstrip Visibility", 1), True
    ToggleImageTabstripAlignment g_UserPreferences.GetPref_Long("Core", "Image Tabstrip Alignment", vbAlignTop), True
    
    'The primary toolbox has some options of its own.  Load them now.
    FormMain.MnuWindowToolbox(2).Checked = g_UserPreferences.GetPref_Boolean("Core", "Show Toolbox Category Labels", True)
    toolbar_Toolbox.UpdateButtonSize g_UserPreferences.GetPref_Long("Core", "Toolbox Button Size", 1), True
    
    
    
    '*************************************************************************************************************************************
    ' Set all default tool values
    '*************************************************************************************************************************************
    
    #If DEBUGMODE = 1 Then
        perfCheck.MarkEvent "Initialize tools"
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
        perfCheck.MarkEvent "Initialize UI"
    #End If
    
    LoadMessage "Initializing user interface..."
    
    'Use the API to give PhotoDemon's main form a 32-bit icon (VB is too old to support 32bpp icons)
    Icons_and_Cursors.SetThunderMainIcon
    
    'Initialize all system cursors we rely on (hand, busy, resizing, etc)
    Icons_and_Cursors.InitializeCursors
    
    'Set up the program's title bar.  Odd-numbered releases are development releases.  Even-numbered releases are formal builds.
    If Not (g_WindowManager Is Nothing) Then
        g_WindowManager.SetWindowCaptionW FormMain.hWnd, Update_Support.GetPhotoDemonNameAndVersion()
    Else
        FormMain.Caption = Update_Support.GetPhotoDemonNameAndVersion()
    End If
    
    'PhotoDemon renders many of its own icons dynamically.  Initialize that engine now.
    InitializeIconHandler
    
    'Prepare a checkerboard pattern, which will be used behind any transparent objects.  Caching this is much more efficient.
    ' than re-creating it every time it's needed.  (Note that PD exposes two versions of the checkerboard pattern: a GDI version
    ' and a GDI+ version.)
    Set g_CheckerboardPattern = New pdDIB
    Drawing.CreateAlphaCheckerboardDIB g_CheckerboardPattern
    Set g_CheckerboardBrush = New pd2DBrush
    g_CheckerboardBrush.SetBrushMode P2_BM_Texture
    g_CheckerboardBrush.SetBrushTextureWrapMode P2_WM_Tile
    g_CheckerboardBrush.SetBrushTextureFromDIB g_CheckerboardPattern
    
    'Allow drag-and-drop operations
    g_AllowDragAndDrop = True
    
    'Throughout the program, g_MouseAccuracy is used to determine how close the mouse cursor must be to a point of interest to
    ' consider it "over" that point.  DPI must be accounted for when calculating this value (as it's calculated in pixels).
    g_MouseAccuracy = FixDPIFloat(6)
    
    'Allow main form components to load any control-specific preferences they may utilize
    FormMain.mainCanvas(0).ReadUserPreferences
    
    'Prep the color management pipeline
    ColorManagement.CacheDisplayCMMData
    
    'Apply visual styles
    g_Themer.SynchronizeThemeMenus
    FormMain.UpdateAgainstCurrentTheme False
    
    
    
    '*************************************************************************************************************************************
    ' The program's menus support many features that VB can't do natively (like icons and custom shortcuts).  Load such things now.
    '*************************************************************************************************************************************
    
    #If DEBUGMODE = 1 Then
        perfCheck.MarkEvent "Prep menus"
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
    Icons_and_Cursors.LoadMenuIcons
    
    'Synchronize all other interface elements to match the current program state (e.g. no images loaded).
    SyncInterfaceToCurrentImage
    
    
    
    '*************************************************************************************************************************************
    ' Unload the splash screen and present the main form
    '*************************************************************************************************************************************
    
    'While in debug mode, copy a timing report of program startup to the debug folder
    #If DEBUGMODE = 1 Then
        perfCheck.StopProfiling
        perfCheck.GenerateProfileReport True
    #End If
    
    'Display the splash screen for at least a second or two
    If (Timer - m_StartTime) < m_LoadTime Then
        Do While (Timer - m_StartTime) < m_LoadTime
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
    
    Unload FormSplash
    
End Sub

'Check for IDE or compiled EXE, and set program parameters accordingly
Private Sub CheckLoadingEnvironment()
    g_IsProgramCompiled = CBool(App.logMode = 1)
End Sub

'FormMain's Unload step calls this process as its final action.
Public Sub FinalShutdown()
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "FinalShutdown() reached."
    #End If
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "Manually unloading all remaining public class instances..."
    #End If
    
    Set g_RecentFiles = Nothing
    Set g_RecentMacros = Nothing
    Set g_Themer = Nothing
    Set g_Displays = Nothing
    Set g_CheckerboardPattern = Nothing
    Set g_Zoom = Nothing
    Set g_WindowManager = Nothing
    
    Dim i As Long
    For i = LBound(pdImages) To UBound(pdImages)
        If (Not pdImages(i) Is Nothing) Then
            pdImages(i).DeactivateImage
            Set pdImages(i) = Nothing
        End If
    Next i
    
    'Delete any remaining temp files in the cache
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "Clearing temp file cache..."
    #End If
    
    FileSystem.DeleteTempFiles
    
    'Release each potentially active plugin in turn
    PluginManager.TerminateAllPlugins
    
    'Release any active drawing backends
    Drawing.ReleaseUIPensAndBrushes
    Set g_CheckerboardPattern = Nothing
    Set g_CheckerboardBrush = Nothing
    If Drawing2D.StopRenderingEngine(P2_DefaultBackend) Then
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "GDI+ released"
        #End If
    End If
    
    'NOTE: in the future, any final user-preference actions could be handled here, as g_UserPreferences is still alive.
    Set g_UserPreferences = Nothing
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "Everything we can physically unload has been forcibly unloaded.  Releasing final library reference..."
    #End If
    
    'If the shell32 library was loaded successfully, once FormMain is closed, we need to unload the library handle.
    If (hShellModule <> 0) Then FreeLibrary hShellModule
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "All human-written code complete.  Shutting down pdDebug and exiting gracefully."
        pdDebug.LogAction "Final memory report", PDM_MEM_REPORT
        pdDebug.TerminateDebugger
        Set pdDebug = Nothing
    #End If
    
    g_IsProgramRunning = False
    
    'We have now terminated everything we can physically terminate.
    
    'Suppress any crashes caused by VB herself (which may be possible due to a variety of issues outside our control),
    ' then let the program go...
    SetErrorMode SEM_NOGPFAULTERRORBOX
    
End Sub
