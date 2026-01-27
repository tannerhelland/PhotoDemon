Attribute VB_Name = "PDMain"
'***************************************************************************
'PhotoDemon Startup Module
'Copyright 2014-2026 by Tanner Helland
'Created: 03/March/14
'Last updated: 16/September/20
'Last update: PD can now forward its command-line to an existing PD session then silently terminate,
'             if user preferences request single-session behavior
'
'The Main() sub in this module is the first thing invoked when PD begins (after VB's own internal startup processes,
' obviously).  I've also included some other crucial startup and shutdown functions in this module.
'
'Portions of the Main() process (related to manually initializing shell libraries) were adopted from a vbforums.com
' project by LaVolpe.  You can see his original work here: http://www.vbforums.com/showthread.php?t=606736
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'If critical errors are affecting PD at startup time, you can activate this constant to forcibly write incomplete debug
' data during the program intialization steps.
' (This constant should always be DISABLED, unless you are doing purely local testing on an egregious startup bug.)
Private Const ENABLE_EMERGENCY_DEBUGGER As Boolean = False

Private Declare Sub InitCommonControls Lib "comctl32" ()
Private Declare Function InitCommonControlsEx Lib "comctl32" (ByRef iccex As InitCommonControlsExStruct) As Long
Private Type InitCommonControlsExStruct
    lngSize As Long
    lngICC As Long
End Type

'As of September 2015, reordering the list of files in PhotoDemon.VBP caused unpredictable
' crashes when PD closes. (After the final line of PD code is run, no less.)  I spent two days
' bisecting commits and can conclusively nail the problem down to
' https://github.com/tannerhelland/PhotoDemon/commit/293de1ba4f2d5bc3102304d0263af624e93b6093
'
'I eventually solved the problem by manually unloading all global class instances in a specific order,
' rather than leaving it to VB, but during testing, I still sometimes find it helpful to suppress
' the default Windows crash dialog. In case this proves useful in the future, I'll leave the declaration.
Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Private Const SEM_NOGPFAULTERRORBOX As Long = &H2

Private m_hShellModule As Long

'This constant is the number of "discrete" loading steps involved in loading the program.
' It is relevant for displaying the progress bar on the initial splash screen; this value is the
' progress bar's maximum value.
Private Const NUMBER_OF_LOADING_STEPS As Long = 19

'After Main() has been invoked, this will be set to TRUE.  This is important in VBy as some functions (like those
' inside user controls) will be called during either design-time or compilation-time.  PD relies on this variable,
' accessed via the IsProgramRunning function, to forcibly suspend certain program operations.
Private m_IsProgramRunning As Boolean

'If the program was loaded successfully, this will be set to TRUE.  Various shutdown procedures check this before
' attempting to write data to file.
Private m_ProgramStartupSuccessful As Boolean

'PhotoDemon starts here.  Main() is necessary as a start point (vs a form) to make sure theming is implemented
' correctly.  Note that this code is irrelevant within the IDE.
Public Sub Main()
    
    m_ProgramStartupSuccessful = False
    
    'InitCommonControlsEx requires IEv3 or above, which shouldn't be a problem on any modern system.
    ' But just in case, continue loading even if the common control module load fails.
    On Error GoTo DamnThisPCisOld
    
    'The following block of code prevents XP crashes when VB usercontrols are present in a project (as they are in PhotoDemon)
    
    'Make sure shell32 is loaded
    Dim strShellName As String
    strShellName = "shell32.dll"
    m_hShellModule = VBHacks.LoadLib(strShellName)
    
    'Make sure comctl32 is loaded.  (For details on these constants, visit http://msdn.microsoft.com/en-us/library/bb775507%28VS.85%29.aspx)
    Dim iccex As InitCommonControlsExStruct
    With iccex
        .lngSize = LenB(iccex)
        Const ICC_BAR_CLASSES As Long = &H4&
        Const ICC_STANDARD_CLASSES As Long = &H4000&
        Const ICC_WIN95_CLASSES As Long = &HFF&
        .lngICC = ICC_STANDARD_CLASSES Or ICC_BAR_CLASSES Or ICC_WIN95_CLASSES
    End With
    InitCommonControlsEx iccex

DamnThisPCisOld:
    'If an error occurs, attempt to initiate the Win9x version, then reset error handling
    If Err Then
        InitCommonControls
        Err.Clear
    End If
    
    On Error GoTo 0
    
    'Because Ambient.UserMode can produce unexpected behavior - see, for example, this link:
    ' http://www.vbforums.com/showthread.php?805711-VB6-UserControl-Ambient-UserMode-workaround
    ' - we manually track program run state.  See additional details at the top of this module,
    ' where m_IsProgramRunning is declared.
    m_IsProgramRunning = True
    
    'In the past, PhotoDemon would manually enable DEP in an attempt to satisfy various virus scanners.
    ' (I never tested this empirically to see if it made a difference.)
    '
    'In 2023 I was notified that DEP breaks many legacy 3rd-party Photoshop plugins.  Ideally, I would write a
    ' separate non-DEP .exe that runs Photoshop plugins and then communicates results back to PD itself,
    ' but at present it's much easier for me to simply disable DEP in PD itself as I have no empirical evidence
    ' that manually enabling DEP makes any difference to end-users.
    
    'CURRENTLY DISABLED PER COMMENT ABOVE:
    'On Win 7+, manually enable DEP.  This may decrease issues with low-quality 3rd-party virus
    ' and malware scanners, and PD works just fine with DEP enabled.
    'OS.EnableProcessDEP
    
    'FormMain can now be loaded.  (We load it first, because many initialization steps silently interact with it,
    ' like loading menu icons or prepping toolboxes.)  That said, the first step of FormMain's load process is calling
    'the ContinueLoadingProgram sub, below, so look there for the next stages of the load process.
    On Error GoTo ExitMainImmediately
    If (Not g_ProgramShuttingDown) Then Load FormMain
    
ExitMainImmediately:

End Sub

'Note that this function is called AFTER FormMain has been loaded.  FormMain is loaded - but not visible - so it can be
' operated on by functions called from this routine.  (It is necessary to load the main window first, since a number of
' load operations - like UI theming and localizations - need to directly operate on the main window.)
Public Function ContinueLoadingProgram(Optional ByRef suspendAdditionalMessages As Boolean = False) As Boolean
    
    'We assume that the program will initialize correctly.  If for some reason it doesn't, it will return FALSE, and the
    ' program needs to be shut down accordingly, because it is catastrophically broken.
    ContinueLoadingProgram = True
    
    '*************************************************************************************************************************************
    ' Check the state of this build (alpha, beta, production, etc) and activate debug code as necessary
    '*************************************************************************************************************************************
    
    'Current build state is stored in the public const "PD_BUILD_QUALITY".  For non-production builds, a number of program-wide
    ' parameters are automatically set.
    
    'If the program is in pre-alpha or alpha state, enable timing reports.
    If (PD_BUILD_QUALITY = PD_ALPHA) Then g_DisplayTimingReports = True
    
    'Enable program-wide high-performance timer objects
    VBHacks.EnableHighResolutionTimers
    
    'Regardless of debug mode, we initialize our internal debugger.  (As of 8.0, the user can
    ' also opt-out of debug tracking, or they can manually enable it in stable builds.)
    PDDebug.InitializeDebugger
    
    'During development, I find it helpful to profile PhotoDemon's startup process (so I can watch for obvious regressions).
    ' PD utilizes several different profiler-types; the LT type is "long-term" profiling, where data is written to a persistent
    ' log file and tracked over time.
    Dim perfCheck As pdProfilerLT
    Set perfCheck = New pdProfilerLT
    perfCheck.StartProfiling "PhotoDemon Startup", True
    
    
    '*************************************************************************************************************************************
    ' With the debugger initialized, prep a few crucial variables
    '*************************************************************************************************************************************
    
    'Most importantly, we need to create a default image array, as some initialization functions
    ' may attempt to access it
    PDImages.ResetPDImageCollection
    
    'Also prep a generic ThunderMain listener
    Set g_ThunderMain = New pdThunderMain
    
    '*************************************************************************************************************************************
    ' Prepare the splash screen (but don't display it yet)
    '*************************************************************************************************************************************
    
    perfCheck.MarkEvent "Prepare splash screen"
    
    'Before doing any 2D rendering, we need to start at least one valid 2D rendering backend.
    ' (At present, only GDI+ is required.)
    Interface.InitializeInterfaceBackend
    
    'Load the splash screen (but note it's only LOADED here, not actually displayed... yet)
    If Drawing2D.StartRenderingEngine(P2_DefaultBackend) Then Load FormSplash
    
    
    '*************************************************************************************************************************************
    ' Determine which version of Windows the user is running (as other load functions rely on this)
    '*************************************************************************************************************************************
    
    perfCheck.MarkEvent "Check Windows version"
    LogStartupEvent "Detecting Windows version..."
    
    'If we are on Windows 7, prepare some Win7-specific features (like taskbar progress bars)
    If OS.IsWin7OrLater Then OS.StartWin7PlusFeatures
    
    
    '*************************************************************************************************************************************
    ' Initialize the user preferences (settings) handler
    '*************************************************************************************************************************************
    
    perfCheck.MarkEvent "Initialize preferences engine"
    UserPrefs.StartPrefEngine
    
    'Ask the preferences handler to generate key program folders.  (If these folders don't exist, the handler will create them.)
    ' Similarly, if the user has done something stupid, like unzip PD inside a system folder, the preferences manager will
    ' auto-detect this and silently redirect program settings to the appropriate user folder.  A flag will also be set, so we
    ' can warn the user about this behavior after the program finishes loading.)
    LogStartupEvent "Initializing all program directories..."
    
    'This is one of the few functions where failures force PD to exit immediately.
    ContinueLoadingProgram = UserPrefs.InitializePaths()
    If (Not ContinueLoadingProgram) Then Exit Function
    
    'Now, ask the preferences handler to load all other user settings from the preferences file.
    LogStartupEvent "Loading all user settings..."
    UserPrefs.LoadUserSettings
    
    'Mark the Macro recorder as "not recording"
    Macros.SetMacroStatus MacroSTOP
    
    'While here, also initialize the image format handler (as plugins and other load functions interact with it)
    ImageImporter.ResetImageImportPreferenceCache
    
    
    '*************************************************************************************************************************************
    ' If this is an emergency debug session, write our first log
    '*************************************************************************************************************************************
    
    'Normally, PD logs a bunch of internal data before exporting its first debug log, but if things are really dire,
    ' we can forcibly initialize debugging here.  (Just note that things like plugin data will *not* be accurate,
    ' as they haven't been loaded yet!)
    If ENABLE_EMERGENCY_DEBUGGER Then PDDebug.StartDebugger True, False
    
    
    '*************************************************************************************************************************************
    ' Enforce multi-instancing checks
    '*************************************************************************************************************************************
    
    perfCheck.MarkEvent "Check multi-session status"
    
    'This step requires access to the UserPrefs module, as each PD install location uses a unique key.
    ' (Note that this check *can happen in the IDE*, contingent on a compile-time constant in the
    ' Mutex module.)
    If (Not Mutex.IsThisOnlyInstance) Then
        
        PDDebug.LogAction "This PhotoDemon instance is not unique!  Querying user preferences for session behavior..."
        
        'Check user preference for single-session behavior
        If UserPrefs.GetPref_Boolean("Loading", "Single Instance", False) Then
            
            'The user wants single-session mode.  Forward our command-line (if any) to the already-open
            ' instance, then immediately exit.
            Dim cPipe As pdPipe
            Set cPipe = New pdPipe
            If cPipe.ConnectToExistingPipe(UserPrefs.GetPref_String("Core", "SessionID", vbNullString, False), True, False, True) Then
                
                Dim cArgStack As pdStream
                Set cArgStack = New pdStream
                cArgStack.StartStream PD_SM_MemoryBacked
                
                Dim ourCmdLine As pdStringStack
                If OS.CommandW(ourCmdLine, True) Then
                    
                    'Write out a dummy value (for full packet size; we'll return and fill this value
                    ' with the actual packet length, once we've written the whole thing)
                    cArgStack.WriteLong 0&
                    
                    'Next, write the total number of embedded arguments
                    cArgStack.WriteLong ourCmdLine.GetNumOfStrings()
                    
                    'Extract the stack into a list of arguments and commands
                    Dim i As Long, sToWrite As String
                    For i = 0 To ourCmdLine.GetNumOfStrings() - 1
                        
                        'Pull the next command-line argument and strip any leading or trailing nulls.
                        ' (Trailing null-char behavior appears to be different when dragging onto an .exe vs
                        '  vs launching from a batch file - see https://github.com/tannerhelland/PhotoDemon/issues/729.
                        '  It's also entirely possible this behavior differs in pre-Win-10 versions, but forcibly
                        '  stripping *any* preceding or trailing nulls should cover all possible variations.)
                        sToWrite = Strings.TrimNull(ourCmdLine.GetString(i))
                        
                        'Write the trimmed string out to file (and IMPORTANTLY, auto-preface each string
                        ' with its length, in bytes, when converted to UTF-8).
                        cArgStack.WriteString_UTF8 sToWrite, True
                        
                    Next i
                    
                    'Retreat to the start of the stream and write out the total stream size
                    cArgStack.SetPosition 0, FILE_BEGIN
                    cArgStack.WriteLong cArgStack.GetStreamSize() - 4
                    
                Else
                    'No arguments?  Not sure what to do here; maybe send some "special" signal
                    ' that requests the main app flash or something?
                    cArgStack.WriteLong 0&
                End If
                
                'Send our data to the already-open PD session
                cPipe.WriteDataToPipe cArgStack.Peek_PointerOnly(0, cArgStack.GetStreamSize()), cArgStack.GetStreamSize
                Set cArgStack = Nothing
                Set cPipe = Nothing
                
                suspendAdditionalMessages = True
                ContinueLoadingProgram = False
                Exit Function
                
            Else
                PDDebug.LogAction "WARNING!  Couldn't connect to existing PD session; starting session anyway..."
            End If
        
        '/end user allows multiple sessions
        End If
        
    Else
        PDDebug.LogAction "(Note: this PD instance is unique; no other instances discovered.)"
    End If
    
    'While here, check another start-up related user perference.  Forced system reboots are
    ' an ever-more-annoying issue on modern versions of Window.  PhotoDemon can automatically
    ' recover sessions interrupted by reboots.
    If UserPrefs.GetPref_Boolean("Loading", "RestoreAfterReboot", False) Then OS.SetRestartRestoreBehavior True
    
    
    '*************************************************************************************************************************************
    ' Initialize the plugin manager and load any high-priority plugins (e.g. those required to start the program successfully)
    '*************************************************************************************************************************************
    
    perfCheck.MarkEvent "Load high-priority plugins"
    PluginManager.InitializePluginManager
    PluginManager.LoadPluginGroup True
    
    
    '*************************************************************************************************************************************
    ' Make sure all required plugins loaded successfully.  If they didn't, bail immediately.
    '*************************************************************************************************************************************
    
    Dim corePluginsAvailable As Boolean: corePluginsAvailable = True
    corePluginsAvailable = corePluginsAvailable And PluginManager.IsPluginCurrentlyEnabled(CCP_zstd)
    corePluginsAvailable = corePluginsAvailable And PluginManager.IsPluginCurrentlyEnabled(CCP_lz4)
    corePluginsAvailable = corePluginsAvailable And PluginManager.IsPluginCurrentlyEnabled(CCP_libdeflate)
    corePluginsAvailable = corePluginsAvailable And PluginManager.IsPluginCurrentlyEnabled(CCP_LittleCMS)
    
    If (Not corePluginsAvailable) Then
        
        PDDebug.LogAction "Core plugins missing or broken; PD will now terminate."
        
        'To help with further debugging, dump a tiny debug file with details on the specific library(ies) that failed.
        ' (Unpatched Windows installs have known issues with 3rd-party libraries that require updated mscvrt runtimes,
        '  so if all 3rd-party libs failed to load this points squarely at an unpatched Windows install that I can't
        '  resolve - knowing this in advance saves all of us trouble.)
        Dim tmpDebugLog As pdString
        Set tmpDebugLog = New pdString
        tmpDebugLog.AppendLine "Core plugin status (installed, enabled)"
        tmpDebugLog.AppendLine "libzstd: " & PluginManager.IsPluginCurrentlyInstalled(CCP_zstd) & ", " & PluginManager.IsPluginCurrentlyEnabled(CCP_zstd)
        tmpDebugLog.AppendLine "liblz4: " & PluginManager.IsPluginCurrentlyInstalled(CCP_lz4) & ", " & PluginManager.IsPluginCurrentlyEnabled(CCP_lz4)
        tmpDebugLog.AppendLine "libdeflate: " & PluginManager.IsPluginCurrentlyInstalled(CCP_libdeflate) & ", " & PluginManager.IsPluginCurrentlyEnabled(CCP_libdeflate)
        tmpDebugLog.AppendLine "lcms2: " & PluginManager.IsPluginCurrentlyInstalled(CCP_LittleCMS) & ", " & PluginManager.IsPluginCurrentlyEnabled(CCP_LittleCMS)
        Files.FileSaveAsText tmpDebugLog.ToString(), UserPrefs.GetDebugPath() & "startup-failure.log", True, True
        
        'Translations will not be available yet, so use non-localized strings.
        Dim tmpMsg As pdString, tmpTitle As String
        Set tmpMsg = New pdString
        tmpMsg.AppendLine "This PhotoDemon copy is broken (essential libraries missing)."
        tmpMsg.AppendLineBreak
        
        tmpMsg.AppendLine "This usually means your PhotoDemon download was interrupted, or the program was unzipped incorrectly."
        tmpMsg.AppendLineBreak
        
        tmpMsg.AppendLine "To fix this copy of PhotoDemon, please try the following steps:"
        tmpMsg.AppendLineBreak
        
        tmpMsg.AppendLine "1) Download a fresh copy from photodemon.org/download"
        tmpMsg.AppendLine "2) Extract the zip file's contents to a folder on your PC.  Make sure both PhotoDemon.exe and its /App subfolder are extracted."
        tmpMsg.AppendLine "3) Double-click the new PhotoDemon.exe file."
        tmpMsg.AppendLine "4) Because the program is freshly downloaded, Windows SmartScreen may raise a confirmation window.  You will need to give PhotoDemon permission to run on your PC."
        tmpMsg.AppendLine "5) Enjoy the program!"
        tmpMsg.AppendLineBreak
        tmpMsg.Append "(This copy of PhotoDemon will now exit.)"
        tmpTitle = "Critical error"
        MsgBox tmpMsg.ToString(), vbCritical Or vbOKOnly Or vbApplicationModal, tmpTitle
        
        'Set termination flags, then exit
        suspendAdditionalMessages = True
        ContinueLoadingProgram = False
        
        Exit Function
        
    End If
        
    '*************************************************************************************************************************************
    ' Initialize the internal resources handler and extract default assets
    '*************************************************************************************************************************************
    
    perfCheck.MarkEvent "Initialize resource handler"
    Set g_Resources = New pdResources
    g_Resources.LoadInitialResourceCollection
    
    'Now that resources are available, extract any default program assets.  (This is only done
    ' once for each group of assets; once extracted, PD will never attempt to extract them again,
    ' unless the user does a full preferences reset - this is to avoid overwriting changes the
    ' user may have made to said files.)
    g_Resources.ExtractDefaultAssets
    
    
    '*************************************************************************************************************************************
    ' Initialize our internal menu manager
    '*************************************************************************************************************************************
    
    perfCheck.MarkEvent "Initialize menu manager"
    Menus.InitializeMenus
    
    
    '*************************************************************************************************************************************
    ' Initialize the translation (language) engine
    '*************************************************************************************************************************************
    
    perfCheck.MarkEvent "Initialize translation engine"
    Set g_Language = New pdTranslate
    
    LogStartupEvent "Scanning for language files..."
    
    'Before doing anything else, check to see what languages are available in the language folder.
    ' (Note that this function will also populate the Languages menu, though it won't place a checkmark next to an entry yet.)
    g_Language.CheckAvailableLanguages
        
    LogStartupEvent "Determining which language to use..."
        
    'Next, determine which language to use.  (This function will take into account the system language at first-run, so it can
    ' estimate which language to present to the user.)
    g_Language.DetermineLanguage
    
    LogStartupEvent "Applying selected language..."
    
    'Apply that language to the program.  This involves loading the translation file into memory, which can take a bit of time,
    ' but it only needs to be done once.  From that point forward, any text requests will operate on the in-memory copy of the file.
    g_Language.ApplyLanguage False
    
    
    '*************************************************************************************************************************************
    ' Initialize the visual themes engine
    '*************************************************************************************************************************************
    
    'Because this class controls the visual appearance of all forms in the project, it must be loaded early in the boot process
    perfCheck.MarkEvent "Initialize theme engine"
    LogStartupEvent "Initializing theme engine..."
    
    Set g_Themer = New pdTheme
    
    'Load and validate the user's selected theme file
    g_Themer.LoadDefaultPDTheme
    
    'Now that a theme has been loaded, we can initialize additional UI rendering elements
    g_Resources.NotifyThemeChange
    Drawing.CacheUIPensAndBrushes
    Tools_Paint.InitializeBrushEngine
    Tools_Pencil.InitializeBrushEngine
    Tools_Clone.InitializeBrushEngine
    SelectionUI.InitializeSelectionRendering
    
    '*************************************************************************************************************************************
    ' PhotoDemon works well with multiple monitors.  Check for such a situation now.
    '*************************************************************************************************************************************
    
    perfCheck.MarkEvent "Detect displays"
    LogStartupEvent "Analyzing current monitor setup..."
    
    Set g_Displays = New pdDisplays
    g_Displays.RefreshDisplays
    
    'While here, also cache various display-related settings; this is faster than constantly retrieving them via APIs
    Interface.CacheSystemDPIRatio g_Displays.GetWindowsDPI
    
    
    '*************************************************************************************************************************************
    ' Now we have what we need to properly display the splash screen.  Do so now.
    '*************************************************************************************************************************************
    
    perfCheck.MarkEvent "Calculate splash screen coordinates"
    
    'Determine the program's previous on-screen location.  We need that to determine where to display the splash screen.
    Dim wRect As RectL
    With wRect
        .Left = UserPrefs.GetPref_Long("Core", "Last Window Left", 1)
        .Top = UserPrefs.GetPref_Long("Core", "Last Window Top", 1)
        .Right = .Left + UserPrefs.GetPref_Long("Core", "Last Window Width", 1)
        .Bottom = .Top + UserPrefs.GetPref_Long("Core", "Last Window Height", 1)
    End With
    
    'Center the splash screen on whichever monitor the user previously used.
    g_Displays.CenterFormViaReferenceRect FormSplash, wRect
    
    'If Segoe UI is available, we prefer to use it instead of Tahoma.  On XP this is not guaranteed, however, so we have to check.
    perfCheck.MarkEvent "Confirm UI font exists"
    Fonts.DetermineUIFont
    
    'Ask the splash screen to finish whatever initializing it needs prior to displaying itself
    perfCheck.MarkEvent "Retrieve splash logo"
    FormSplash.PrepareSplashLogo NUMBER_OF_LOADING_STEPS
    perfCheck.MarkEvent "Finalize splash screen"
    FormSplash.PrepareRestOfSplash
    
    'Display the splash screen, centered on whichever monitor the user previously used the program on.
    perfCheck.MarkEvent "Display splash screen"
    If UserPrefs.GetPref_Boolean("Loading", "splash-screen", True) Then FormSplash.Show vbModeless
    
    '*************************************************************************************************************************************
    ' If this is not a production build, initialize PhotoDemon's central debugger
    '*************************************************************************************************************************************
    
    'We wait until after the translation and plugin engines are initialized; this allows us to report their information in the debug log
    perfCheck.MarkEvent "Initialize debugger"
    If UserPrefs.GenerateDebugLogs Then PDDebug.StartDebugger True
    
    
    '*************************************************************************************************************************************
    ' Build a font cache for this system
    '*************************************************************************************************************************************
    
    perfCheck.MarkEvent "Build font cache"
    LogStartupEvent "Building font cache..."
        
    'PD currently builds two font caches:
    ' 1) A name-only list of all fonts currently installed.  This is used to populate font dropdown boxes.
    ' 2) An pdFont-based cache of the current UI font, at various requested sizes.  This cache spares individual controls from needing
    '     to do their own font management; instead, they can simply request a matching object from the Fonts module.
    Fonts.BuildFontCaches
    
    
    '*************************************************************************************************************************************
    ' Initialize PD's central clipboard manager
    '*************************************************************************************************************************************
    
    perfCheck.MarkEvent "Initialize pdClipboardMain"
    LogStartupEvent "Initializing clipboard interface..."
    Set g_Clipboard = New pdClipboardMain
    
    
    '*************************************************************************************************************************************
    ' Get the viewport engine ready
    '*************************************************************************************************************************************
    
    perfCheck.MarkEvent "Initialize viewport engine"
    LogStartupEvent "Initializing viewport engine..."
    
    'The viewport engine is currently compartmentalized into a few different pieces.
    ' The "zoom" portion deals with the UI that translates user-facing zoom attributes
    ' (e.g. mousewheel events or the canvas zoom dropdown) into actual zoom ratios.
    Zoom.InitializeZoomEngine
    Zoom.PopulateZoomDropdown FormMain.MainCanvas(0).GetZoomDropDownReference
    
    'Manually populate the main canvas's size unit dropdown
    FormMain.MainCanvas(0).PopulateSizeUnits
    
    
    '*************************************************************************************************************************************
    ' Finish loading low-priority plugins
    '*************************************************************************************************************************************
    
    perfCheck.MarkEvent "Load low-priority plugins"
    PluginManager.LoadPluginGroup False
    PluginManager.ReportPluginLoadSuccess
    
    
    '*************************************************************************************************************************************
    ' Based on available plugins, determine which image formats PhotoDemon can handle
    '*************************************************************************************************************************************
    
    perfCheck.MarkEvent "Load import and export libraries"
    LogStartupEvent "Loading import/export libraries..."
    
    'Generate a list of currently supported input/output formats, which may vary based on plugin version and availability
    ImageFormats.GenerateInputFormats
    ImageFormats.GenerateOutputFormats
    
    
    '*************************************************************************************************************************************
    ' Load keyboard shortcuts (either PD's default collection, or a standalone file with user edits)
    '*************************************************************************************************************************************
    
    'In late 2024, this step was moved earlier in the load process.  PD needs to have hotkey data
    ' available before toolboxes are loaded, because the toolbox pulls hotkey data from the hotkey
    ' manager in order to show relevant shortcuts in tooltips.
    
    'Initialize hotkeys.  (User-customized hotkeys will be loaded from file; if they don't exist,
    ' this function will generate a default hotkey list.)
    perfCheck.MarkEvent "Initialize hotkey manager"
    LogStartupEvent "Initializing hotkeys..."
    Hotkeys.InitializeHotkeys
    
    
    '*************************************************************************************************************************************
    ' Initialize the window manager (the class that synchronizes all toolbox and image window positions)
    '*************************************************************************************************************************************
    
    perfCheck.MarkEvent "Initialize window manager"
    LogStartupEvent "Initializing window manager..."
    Set g_WindowManager = New pdWindowManager
    
    'Register the main form
    g_WindowManager.SetAutoRefreshMode False
    g_WindowManager.RegisterMainForm FormMain
    
    'As of 7.0, all we need to do here is initialize the new, lightweight toolbox handler.
    ' This will load things like toolbox sizes and visibility from the previous session.
    Toolboxes.LoadToolboxData
    
    'With toolbox data assembled, we can now silently load each tool window.  Even though these
    ' windows may not be visible (as the user can elect to hide them), we still want them loaded
    ' so we can activate them quickly if/when they are enabled.
    perfCheck.MarkEvent "Window manager: load left toolbox"
    Load toolbar_Toolbox
    
    perfCheck.MarkEvent "Window manager: load right toolbox"
    Load toolbar_Layers
    
    perfCheck.MarkEvent "Window manager: load bottom toolbox"
    Load toolbar_Options
    
    'Retrieve tool window visibility and mark those menus as well
    FormMain.MnuWindowToolbox(0).Checked = Toolboxes.GetToolboxVisibilityPreference(PDT_LeftToolbox)
    FormMain.MnuWindow(1).Checked = Toolboxes.GetToolboxVisibilityPreference(PDT_TopToolbox)
    FormMain.MnuWindow(2).Checked = Toolboxes.GetToolboxVisibilityPreference(PDT_RightToolbox)
    
    'Retrieve two additional settings for the image tabstrip menu: when to display it, and its alignment
    ToggleImageTabstripVisibility UserPrefs.GetPref_Long("Core", "Image Tabstrip Visibility", 1), True
    ToggleImageTabstripAlignment UserPrefs.GetPref_Long("Core", "Image Tabstrip Alignment", vbAlignTop), True
    
    'The primary toolbox has some options of its own.  Load them now.
    FormMain.MnuWindowToolbox(2).Checked = UserPrefs.GetPref_Boolean("Core", "Show Toolbox Category Labels", True)
    toolbar_Toolbox.UpdateButtonSize UserPrefs.GetPref_Long("Core", "Toolbox Button Size", tbs_Small), True
    
    
    '*************************************************************************************************************************************
    ' Set all default tool values
    '*************************************************************************************************************************************
    
    perfCheck.MarkEvent "Initialize tools"
    LogStartupEvent "Initializing image tools..."
    
    'As of May 2015, tool panels are now loaded on-demand.  This improves the program's startup performance,
    ' and it saves a bit of memory if a user doesn't use a tool during a given session.  This negates the need
    ' to initialize any tools here.
    
    'But while we're here, let's prep the specialized non-destructive tool handler in the central processor
    Processor.InitializeProcessor
    
    'Similarly, build a "database" of action names and attributes.  This is queried by multiple parts of the
    ' app to determine if an action is e.g. "repeat-able" or "fade-able", etc
    Actions.BuildActionDatabase
    
    '*************************************************************************************************************************************
    ' PhotoDemon's complex interface requires a lot of things to be generated at run-time.
    '*************************************************************************************************************************************
    
    perfCheck.MarkEvent "Initialize UI"
    LogStartupEvent "Initializing user interface..."
    
    'Use the API to give PhotoDemon's main form a 32-bit icon (VB is too old to support 32bpp icons)
    IconsAndCursors.SetThunderMainIcon
    
    'Initialize all system cursors we rely on (hand, busy, resizing, etc)
    IconsAndCursors.InitializeCursors
    
    'Set up the program's title bar.  Odd-numbered releases are development releases.  Even-numbered releases are formal builds.
    If (Not g_WindowManager Is Nothing) Then
        g_WindowManager.SetWindowCaptionW FormMain.hWnd, Updates.GetPhotoDemonNameAndVersion()
    Else
        FormMain.Caption = Updates.GetPhotoDemonNameAndVersion()
    End If
    
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
    
    'Allow main form components to load any control-specific preferences they may utilize
    FormMain.MainCanvas(0).ReadUserPreferences
    
    'Prep the color management pipeline and any associated color management settings
    ColorManagement.CacheDisplayCMMData
    ColorManagement.UpdateColorManagementPreferences
    
    
    '*************************************************************************************************************************************
    ' The program's menus support many features that VB can't do natively (like icons and custom shortcuts).  Load such things now.
    '*************************************************************************************************************************************
    
    perfCheck.MarkEvent "Prep developer menus"
    LogStartupEvent "Preparing program menus..."
    
    'In alpha-build or IDE modes, Tools > Options > Developer is exposed.
    ' (Two menus are toggled for this - the first is just a separator bar.)
    'TODO: when we switch to owner-drawn menus, there is no mechanism for run-time visibility changes.
    ' These menus will simply need to *not* be created at all.
    Dim debugMenuVisibility As Boolean
    debugMenuVisibility = ((PD_BUILD_QUALITY <> PD_PRODUCTION) And (PD_BUILD_QUALITY <> PD_BETA)) Or (Not OS.IsProgramCompiled)
    FormMain.MnuTool(13).Visible = debugMenuVisibility
    FormMain.MnuTool(14).Visible = debugMenuVisibility
    
    'Initialize the Recent Files manager and load the most-recently-used file list (MRU)
    perfCheck.MarkEvent "Prep MRU menus"
    LogStartupEvent "Initializing recent file lists..."
    Set g_RecentFiles = New pdRecentFiles
    g_RecentFiles.LoadListFromFile
    
    Set g_RecentMacros = New pdMRUManager
    g_RecentMacros.InitList New pdMRURecentMacros
    g_RecentMacros.MRU_LoadFromFile
    
    'Load and draw all menu icons
    perfCheck.MarkEvent "Load all menu icons"
    LogStartupEvent "Loading UI icons..."
    IconsAndCursors.LoadMenuIcons False
    
    'Finally, apply all of our various UI features
    perfCheck.MarkEvent "Apply theme/language to FormMain"
    FormMain.UpdateAgainstCurrentTheme
    
    'Synchronize all other interface elements to match the current program state (e.g. no images loaded).
    perfCheck.MarkEvent "Final interface sync"
    LogStartupEvent "Final interface sync before main screen appears..."
    Interface.SyncInterfaceToCurrentImage
    
    'Minimize UI memory usage
    LogStartupEvent "Minimizing sprite cache..."
    UIImages.MinimizeCacheMemory
    
    'If we made it all the way here, startup can be considered successful!
    m_ProgramStartupSuccessful = True
    
    '*************************************************************************************************************************************
    ' Unload the splash screen and present the main form
    '*************************************************************************************************************************************
    
    'While in debug mode, copy a timing report of program startup to the debug folder
    perfCheck.StopProfiling
    If UserPrefs.GenerateDebugLogs Then perfCheck.GenerateProfileReport True
    
    'If this is the first time the user has run PhotoDemon, resize the window a bit to make the default position nice.
    ' (If this is *not* the first time, the window manager will automatically restore the window's last-known position and state.)
    If g_IsFirstRun Then g_WindowManager.SetFirstRunMainWindowPosition
    
    'In debug mode, make a baseline memory reading here, before the main form is displayed.
    PDDebug.LogAction "LoadTheProgram() function complete.  Baseline memory reading:"
    PDDebug.LogAction vbNullString, PDM_Mem_Report
    PDDebug.LogAction "Proceeding to load main window..."
    
    Unload FormSplash
    
End Function

'FormMain's Unload step calls this process as its final action.
Public Sub FinalShutdown()
    
    PDDebug.LogAction "FinalShutdown() reached."
    PDDebug.LogAction "Manually unloading all remaining public class instances..."
    
    Set g_ThunderMain = Nothing
    Set g_RecentFiles = Nothing
    Set g_RecentMacros = Nothing
    Set g_Themer = Nothing
    Set g_Displays = Nothing
    Set g_CheckerboardPattern = Nothing
    Set g_WindowManager = Nothing
    
    'Release any remaining data associated with user-loaded images
    PDImages.ReleaseAllPDImageResources
    
    'Report final profiling data
    Viewport.ReportViewportProfilingData
    If (Not g_Language Is Nothing) Then
        PDDebug.LogAction "Final translation engine time was: " & Format$(g_Language.GetNetTranslationTime() * 1000#, "0.0") & " ms"
        g_Language.PrintAdditionalDebugInfo
    End If
    
    'Free any other resources we're manually managing
    PDDebug.LogAction "Releasing VB-specific hackarounds..."
    VBHacks.ShutdownCleanup
    Mutex.FreeAllMutexes
    
    'Delete any remaining temp files in the cache
    PDDebug.LogAction "Clearing temp file cache..."
    Files.DeleteTempFiles
    
    'Release each potentially active plugin in turn
    PluginManager.TerminateAllPlugins
    
    'Release any active drawing backends
    Drawing.ReleaseUIPensAndBrushes
    Set g_CheckerboardPattern = Nothing
    Set g_CheckerboardBrush = Nothing
    If Drawing2D.StopRenderingEngine(P2_DefaultBackend) Then PDDebug.LogAction "GDI+ released"
    
    'Write all preferences out to file and terminate the XML parser
    PDDebug.LogAction "Writing user preferences to file..."
    UserPrefs.StopPrefEngine
    
    PDDebug.LogAction "Everything we can physically unload has been forcibly unloaded.  Releasing final library reference..."
    
    'If the shell32 library was loaded successfully, once FormMain is closed, we need to unload the library handle.
    VBHacks.FreeLib m_hShellModule
    
    PDDebug.LogAction "All human-written code complete.  Shutting down pdDebug and exiting gracefully."
    PDDebug.LogAction "Final memory report", PDM_Mem_Report
    PDDebug.TerminateDebugger
    
    m_IsProgramRunning = False
    
    'We have now terminated everything we can physically terminate.
    
    'Suppress any crashes caused by VB herself (which can happen due to a variety of issues outside our control),
    ' then let the program go...
    SetErrorMode SEM_NOGPFAULTERRORBOX
    
End Sub

'Returns TRUE if Main() has been invoked
Public Function IsProgramRunning() As Boolean
    IsProgramRunning = m_IsProgramRunning
End Function

'Returns TRUE if PD's startup routines all triggered successfully.
Public Function WasStartupSuccessful() As Boolean
    WasStartupSuccessful = m_ProgramStartupSuccessful
End Function
