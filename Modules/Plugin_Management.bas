Attribute VB_Name = "PluginManager"
'***************************************************************************
'Core Plugin Manager
'Copyright 2014-2018 by Tanner Helland
'Created: 30/August/15
'Last updated: 01/December/16
'Last update: integrate zstd as a permanent plugin
'
'As PD grows, it's more and more difficult to supply the functionality we need through WAPI alone.  To that end,
' a number of third-party plugins are now required for proper program operation.
'
'To simplify the management of these plugins, I've created this singular module.  My hope is that future plugins
' will be easier to add and maintain thanks to this.
'
'When adding a new plugin, please make sure to read the declarations at the top of the class, particularly the
' CORE_PLUGINS enum and associated CORE_PLUGIN_COUNT constant.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Currently supported core plugins.  These values are arbitrary and can be changed without consequence, but THEY MUST
' ALWAYS BE SEQUENTIAL, STARTING WITH ZERO, because the enum is iterated using For loops (e.g. during initialization).
Public Enum CORE_PLUGINS
    CCP_ExifTool = 0
    CCP_EZTwain = 1
    CCP_FreeImage = 2
    CCP_LittleCMS = 3
    CCP_lz4 = 4
    CCP_OptiPNG = 5
    CCP_PNGQuant = 6
    CCP_zLib = 7
    CCP_zstd = 8
End Enum

#If False Then
    Private Const CCP_ExifTool = 0, CCP_EZTwain = 1, CCP_FreeImage = 2, CCP_LittleCMS = 3, CCP_lz4 = 4, CCP_OptiPNG = 5, CCP_PNGQuant = 6, CCP_zLib = 7, CCP_zstd = 8
#End If

'Expected version numbers of plugins.  These are updated at each new PhotoDemon release (if a new version of
' the plugin is available, obviously).
Private Const EXPECTED_EXIFTOOL_VERSION As String = "10.93"
Private Const EXPECTED_EZTWAIN_VERSION As String = "1.18.0"
Private Const EXPECTED_FREEIMAGE_VERSION As String = "3.18.0"
Private Const EXPECTED_LITTLECMS_VERSION As String = "2.9.0"
Private Const EXPECTED_LZ4_VERSION As String = "10801"
Private Const EXPECTED_OPTIPNG_VERSION As String = "0.7.7"
Private Const EXPECTED_PNGQUANT_VERSION As String = "2.5.2"
Private Const EXPECTED_ZLIB_VERSION As String = "1.2.8"
Private Const EXPECTED_ZSTD_VERSION As String = "10304"

'This constant is used to iterate all core plugins (as listed under the CORE_PLUGINS enum), so if you add or remove
' a plugin, make sure to update this!
Private Const CORE_PLUGIN_COUNT As Long = 9

'To simplify handling throughout this module, plugin existence, allowance, and successful initialization are tracked internally.
' Note that not all of these specific states are retrievable externally; in general, callers should use the simplified
' IsPluginCurrentlyEnabled() function to determine if a plugin is usable (e.g. installed, valid version, not forcibly disabled).
Private m_PluginExists() As Boolean
Private m_PluginAllowed() As Boolean
Private m_PluginInitialized() As Boolean

'For high-performance code paths, we specifically track a few plugin states.
Private m_ZlibEnabled As Boolean, m_ZstdEnabled As Boolean, m_lz4Enabled As Boolean
Private m_ExifToolEnabled As Boolean, m_OptiPNGEnabled As Boolean, m_LCMSEnabled As Boolean

'Path to plugin folder.  For security reasons, this is forcibly constructed as an absolute path
' (generally "App.Path/App/PhotoDemon/Plugins"), because we pass it directly to LoadLibrary.
Private m_PluginPath As String

Public Function GetPluginPath() As String
    If (LenB(m_PluginPath) <> 0) Then
        GetPluginPath = m_PluginPath
    Else
        PDDebug.LogAction "WARNING!  PluginManager.GetPluginPath() was called before the plugin manager was initialized!"
    End If
End Function

Public Function GetNumOfPlugins() As Long
    GetNumOfPlugins = CORE_PLUGIN_COUNT
End Function

'Before loading any program plugins, you must first call this function to initialize the plugin manager.
Public Sub InitializePluginManager()

    'Reset all plugin trackers
    ReDim m_PluginExists(0 To CORE_PLUGIN_COUNT - 1) As Boolean
    ReDim m_PluginAllowed(0 To CORE_PLUGIN_COUNT - 1) As Boolean
    ReDim m_PluginInitialized(0 To CORE_PLUGIN_COUNT - 1) As Boolean
    
    'Plugin files are located in the \Data\Plugins subdirectory
    m_PluginPath = UserPrefs.GetAppPath() & "Plugins\"
    
    'Make sure the plugin path exists
    If (Not Files.PathExists(PluginManager.GetPluginPath)) Then Files.PathCreate PluginManager.GetPluginPath, True
    
End Sub

'This subroutine handles the detection and installation of all core plugins. required for an optimal PhotoDemon
' experience: zLib, EZTwain32, and FreeImage.  For convenience' sake, it also checks for GDI+ availability.
Public Sub LoadPluginGroup(Optional ByVal loadHighPriorityPlugins As Boolean = True)
    
    If loadHighPriorityPlugins Then
        PDDebug.LogAction "Initializing high-priority plugins..."
    Else
        PDDebug.LogAction "Initializing low-priority plugins..."
    End If
    
    Dim startTime As Currency
        
    'Plugin loading is handled in a loop.  This loop will call several helper functions, passing each sequential plugin
    ' index as defined by the CORE_PLUGINS enum (and matching CORE_PLUGIN_COUNT const).  Some initialization steps are
    ' shared among all plugins (e.g. checking for the plugin's existence), while some require custom initializations.
    ' This behavior is all carefully documented in the functions called by the initialization loop.
    Dim i As Long
    For i = 0 To CORE_PLUGIN_COUNT - 1
        
        If (loadHighPriorityPlugins = IsPluginHighPriority(i)) Then
            
            VBHacks.GetHighResTime startTime
            
            'Before doing anything else, see if the plugin file actually exists.
            m_PluginExists(i) = DoesPluginFileExist(i)
            
            'If the plugin file exists, see if the user has forcibly disabled it.  If they have, we can skip initialization.
            ' we can initialize it.  (Some plugins may not require this step; that's okay.)
            If m_PluginExists(i) Then m_PluginAllowed(i) = IsPluginAllowed(i)
            
            'If the user has allowed a plugin's use, attempt to initialize it.
            If m_PluginAllowed(i) Then m_PluginInitialized(i) = InitializePlugin(i)
            
            'We now know enough to set global initialization flags.  (This step is technically optional; see comments in the matching sub.)
            SetGlobalPluginFlags i, m_PluginInitialized(i)
                    
            'Finally, if a plugin affects UI or other user-exposed bits, that's the last thing we set before exiting.
            ' (This step is optional; plugins do not need to support it.)
            FinalizePluginInitialization i, m_PluginInitialized(i)
            
            PDDebug.LogAction GetPluginName(i) & " initialized in " & Format$(CStr(VBHacks.GetTimerDifferenceNow(startTime) * 1000#), "#####0") & " ms"
            
        End If
        
    Next i
    
End Sub

'List the initialization state of all plugins.  This is currently only enabled in debug builds, and it is very helpful
' for tracking down obscure 3rd-party library issues.
Public Sub ReportPluginLoadSuccess()

    'Initialization complete!  In debug builds, write out some plugin debug information.
    Dim successfulPluginCount As Long
    successfulPluginCount = 0
    
    Dim i As Long
    For i = 0 To CORE_PLUGIN_COUNT - 1
        If m_PluginInitialized(i) Then
            successfulPluginCount = successfulPluginCount + 1
        Else
            PDDebug.LogAction "WARNING!  Plugin ID#" & i & " (" & GetPluginName(i) & ") was not initialized."
        End If
    Next i
    
    PDDebug.LogAction CStr(successfulPluginCount) & "/" & CStr(CORE_PLUGIN_COUNT) & " plugins initialized successfully."
    
End Sub

'Given a plugin enum value, return a string of the core plugin's filename.  Note that this (obviously) does not include helper files,
' like README or LICENSE files - just the core DLL or EXE for the plugin.
Public Function GetPluginFilename(ByVal pluginEnumID As CORE_PLUGINS) As String
    Select Case pluginEnumID
        Case CCP_ExifTool
            GetPluginFilename = "exiftool.exe"
        Case CCP_EZTwain
            GetPluginFilename = "eztw32.dll"
        Case CCP_FreeImage
            GetPluginFilename = "FreeImage.dll"
        Case CCP_LittleCMS
            GetPluginFilename = "lcms2.dll"
        Case CCP_lz4
            GetPluginFilename = "liblz4.dll"
        Case CCP_OptiPNG
            GetPluginFilename = "optipng.exe"
        Case CCP_PNGQuant
            GetPluginFilename = "pngquant.exe"
        Case CCP_zLib
            GetPluginFilename = "zlibwapi.dll"
        Case CCP_zstd
            GetPluginFilename = "libzstd.dll"
    End Select
End Function

Public Function GetPluginName(ByVal pluginEnumID As CORE_PLUGINS) As String
    Select Case pluginEnumID
        Case CCP_ExifTool
            GetPluginName = "ExifTool"
        Case CCP_EZTwain
            GetPluginName = "EZTwain"
        Case CCP_FreeImage
            GetPluginName = "FreeImage"
        Case CCP_LittleCMS
            GetPluginName = "LittleCMS"
        Case CCP_lz4
            GetPluginName = "LZ4"
        Case CCP_OptiPNG
            GetPluginName = "OptiPNG"
        Case CCP_PNGQuant
            GetPluginName = "PNGQuant"
        Case CCP_zLib
            GetPluginName = "zLib"
        Case CCP_zstd
            GetPluginName = "zstd"
        Case Else
            Debug.Print "WARNING!  PluginManager.GetPluginName was handed an invalid Enum ID."
    End Select
End Function

'Plugin versions can be retrieved via two primary means:
' 1) Just reading the product version metadata from the actual file
' 2) Using some plugin-specific mechanism (typically an exported GetVersion() function of some sort)
'
'If a version cannot be retrieved, this function returns a blank string
Public Function GetPluginVersion(ByVal pluginEnumID As CORE_PLUGINS) As String
    
    GetPluginVersion = vbNullString
    
    Select Case pluginEnumID
        
        'ExifTool can write its version number to stdout
        Case CCP_ExifTool
            If PluginManager.IsPluginCurrentlyInstalled(pluginEnumID) Then GetPluginVersion = ExifTool.GetExifToolVersion()
            
        'EZTwain provides a dedicated version-checking function
        Case CCP_EZTwain
            If PluginManager.IsPluginCurrentlyInstalled(pluginEnumID) Then GetPluginVersion = Plugin_EZTwain.GetEZTwainVersion()
        
        'LittleCMS provides a dedicated version-checking function
        Case CCP_LittleCMS
            If PluginManager.IsPluginCurrentlyInstalled(pluginEnumID) Then GetPluginVersion = LittleCMS.GetLCMSVersion()
        
        'LZ4 provides a dedicated version-checking function
        Case CCP_lz4
            If PluginManager.IsPluginCurrentlyInstalled(pluginEnumID) Then GetPluginVersion = Plugin_lz4.GetLz4Version()
            
        'OptiPNG can write its version number to stdout
        Case CCP_OptiPNG
            If PluginManager.IsPluginCurrentlyInstalled(pluginEnumID) Then GetPluginVersion = Plugin_OptiPNG.GetOptiPNGVersion()
            
        'PNGQuant can write its version number to stdout
        Case CCP_PNGQuant
            If PluginManager.IsPluginCurrentlyInstalled(pluginEnumID) Then GetPluginVersion = Plugin_PNGQuant.GetPngQuantVersion()
        
        'zstd provides a dedicated version-checking function
        Case CCP_zstd
            If PluginManager.IsPluginCurrentlyInstalled(pluginEnumID) Then GetPluginVersion = Plugin_zstd.GetZstdVersion()
        
        'All other plugins pull their version info directly from file metadata
        Case Else
            Dim cFSO As pdFSO
            Set cFSO = New pdFSO
            cFSO.FileGetVersionAsString PluginManager.GetPluginPath & PluginManager.GetPluginFilename(pluginEnumID), GetPluginVersion, True
            
    End Select
    
End Function

'Given a plugin enum value, return a string stack of any non-essential files associated with the plugin.  This includes things like
' README or LICENSE files, and it can be EMPTY if no helper files exist.
'
'Returns TRUE if one or more helper files exist; FALSE if none exist.  This should make it easier for the caller to know if the
' string stack needs to be processed further.
Private Function GetNonEssentialPluginFiles(ByVal pluginEnumID As CORE_PLUGINS, ByRef dstStringStack As pdStringStack) As Boolean
    
    If dstStringStack Is Nothing Then Set dstStringStack = New pdStringStack
    dstStringStack.ResetStack
    
    Select Case pluginEnumID
    
        Case CCP_ExifTool
            dstStringStack.AddString "exiftool-README.txt"
                    
        Case CCP_EZTwain
            dstStringStack.AddString "eztwain-README.txt"
        
        Case CCP_FreeImage
            dstStringStack.AddString "freeimage-LICENSE.txt"
        
        Case CCP_LittleCMS
            dstStringStack.AddString "lcms2-LICENSE.txt"
            
        Case CCP_lz4
            dstStringStack.AddString "liblz4-LICENSE.txt"
        
        Case CCP_OptiPNG
            dstStringStack.AddString "optipng-LICENSE.txt"
            
        Case CCP_PNGQuant
            dstStringStack.AddString "pngquant-README.txt"
        
        Case CCP_zLib
            dstStringStack.AddString "zlib-README.txt"
            
        Case CCP_zstd
            dstStringStack.AddString "libzstd-LICENSE.txt"
            
    End Select
    
    GetNonEssentialPluginFiles = (dstStringStack.GetNumOfStrings <> 0)
    
End Function

'The Plugin Manager dialog allows the user to forcibly disable plugins.  This can be very helpful when testing bugs and crashes,
' but generally isn't relevant for a casual user.  Regardless, the plugin loader will check this value prior to initializing a plugin.
Private Function IsPluginAllowed(ByVal pluginEnumID As CORE_PLUGINS) As Boolean
    IsPluginAllowed = Not UserPrefs.GetPref_Boolean("Plugins", "Force " & PluginManager.GetPluginName(pluginEnumID) & " Disable", False)
End Function

'Simplified function to forcibly disable a plugin via the user's preference file.  Note that this *will not take affect for
' this session* by design; you must subsequently call the SetPluginEnablement() function to live-change the setting.
Public Sub SetPluginAllowed(ByVal pluginEnumID As CORE_PLUGINS, ByVal newEnabledState As Boolean)
    UserPrefs.SetPref_Boolean "Plugins", "Force " & PluginManager.GetPluginName(pluginEnumID) & " Disable", Not newEnabledState
End Sub

'Simplified function to detect if a given plugin is currently enabled.  (Plugins can be disabled for a variety of reasons,
' including forcible disablement by the user, bugs, missing files, etc; this catch-all function returns a binary "enabled" state.)
Public Function IsPluginCurrentlyEnabled(ByVal pluginEnumID As CORE_PLUGINS) As Boolean
    Select Case pluginEnumID
        Case CCP_ExifTool
            IsPluginCurrentlyEnabled = m_ExifToolEnabled
        Case CCP_EZTwain
            IsPluginCurrentlyEnabled = Plugin_EZTwain.IsScannerAvailable
        Case CCP_FreeImage
            IsPluginCurrentlyEnabled = g_ImageFormats.FreeImageEnabled
        Case CCP_LittleCMS
            IsPluginCurrentlyEnabled = m_LCMSEnabled
        Case CCP_lz4
            IsPluginCurrentlyEnabled = m_lz4Enabled
        Case CCP_OptiPNG
            IsPluginCurrentlyEnabled = m_OptiPNGEnabled
        Case CCP_PNGQuant
            IsPluginCurrentlyEnabled = g_ImageFormats.pngQuantEnabled
        Case CCP_zLib
            IsPluginCurrentlyEnabled = m_ZlibEnabled
        Case CCP_zstd
            IsPluginCurrentlyEnabled = m_ZstdEnabled
    End Select
End Function

'Simplified function to forcibly en/disable a given plugin for this session.  Note that this has program-wide repercussions,
' including UI states that may no longer be valid - as such, this function should *not* be changed by anything except the
' Plugin Manager dialog or the plugin initialization functions.
Public Sub SetPluginEnablement(ByVal pluginEnumID As CORE_PLUGINS, ByVal newEnabledState As Boolean)
    Select Case pluginEnumID
        Case CCP_ExifTool
            m_ExifToolEnabled = newEnabledState
        Case CCP_EZTwain
            Plugin_EZTwain.ForciblySetScannerAvailability newEnabledState
        Case CCP_FreeImage
            g_ImageFormats.FreeImageEnabled = newEnabledState
        Case CCP_LittleCMS
            m_LCMSEnabled = newEnabledState
        Case CCP_lz4
            m_lz4Enabled = newEnabledState
        Case CCP_OptiPNG
            m_OptiPNGEnabled = newEnabledState
        Case CCP_PNGQuant
            g_ImageFormats.pngQuantEnabled = newEnabledState
        Case CCP_zLib
            m_ZlibEnabled = newEnabledState
        Case CCP_zstd
            m_ZstdEnabled = newEnabledState
    End Select
End Sub

'Simplified function to detect if a given plugin is currently installed in PD's plugin folder.  This is (obviously) separate
' from a plugin actually being *enabled*, as that requires initialization and other steps.
Public Function IsPluginCurrentlyInstalled(ByVal pluginEnumID As CORE_PLUGINS) As Boolean
    IsPluginCurrentlyInstalled = Files.FileExists(PluginManager.GetPluginPath & GetPluginFilename(pluginEnumID))
End Function

'PD loads plugins in two waves.  Before the splash screen appears, "high-priority" plugins are loaded.  These include the
' decompression plugins required to decompress things like the splash screen image.  Much later in the load process,
' we load the rest of the program's core plugins.  This function determines which wave a plugin is loaded during.
Public Function IsPluginHighPriority(ByVal pluginEnumID As CORE_PLUGINS) As Boolean
    Select Case pluginEnumID
        Case CCP_ExifTool
            IsPluginHighPriority = False
        Case CCP_EZTwain
            IsPluginHighPriority = False
        Case CCP_FreeImage
            IsPluginHighPriority = True
        Case CCP_LittleCMS
            IsPluginHighPriority = True
        Case CCP_lz4
            IsPluginHighPriority = True
        Case CCP_OptiPNG
            IsPluginHighPriority = False
        Case CCP_PNGQuant
            IsPluginHighPriority = False
        Case CCP_zLib
            IsPluginHighPriority = True
        Case CCP_zstd
            IsPluginHighPriority = True
    End Select
End Function

'Simplified function to return the expected version number of a plugin.  These numbers change with each PD release, and they can
' be helpful for seeing if a user has manually updated a plugin file to some new version (which is generally okay!)
Public Function ExpectedPluginVersion(ByVal pluginEnumID As CORE_PLUGINS) As String
    Select Case pluginEnumID
        Case CCP_ExifTool
            ExpectedPluginVersion = EXPECTED_EXIFTOOL_VERSION
        Case CCP_EZTwain
            ExpectedPluginVersion = EXPECTED_EZTWAIN_VERSION
        Case CCP_FreeImage
            ExpectedPluginVersion = EXPECTED_FREEIMAGE_VERSION
        Case CCP_LittleCMS
            ExpectedPluginVersion = EXPECTED_LITTLECMS_VERSION
        Case CCP_lz4
            ExpectedPluginVersion = EXPECTED_LZ4_VERSION
        Case CCP_OptiPNG
            ExpectedPluginVersion = EXPECTED_OPTIPNG_VERSION
        Case CCP_PNGQuant
            ExpectedPluginVersion = EXPECTED_PNGQUANT_VERSION
        Case CCP_zLib
            ExpectedPluginVersion = EXPECTED_ZLIB_VERSION
        Case CCP_zstd
            ExpectedPluginVersion = EXPECTED_ZSTD_VERSION
    End Select
End Function

'Simplified function for retrieving the homepage URL for a given plugin
Public Function GetPluginHomepage(ByVal pluginEnumID As CORE_PLUGINS) As String
    Select Case pluginEnumID
        Case CCP_ExifTool
            GetPluginHomepage = "http://www.sno.phy.queensu.ca/~phil/exiftool/"
        Case CCP_EZTwain
            GetPluginHomepage = "http://eztwain.com/eztwain1.htm"
        Case CCP_FreeImage
            GetPluginHomepage = "http://freeimage.sourceforge.net/"
        Case CCP_LittleCMS
            GetPluginHomepage = "http://www.littlecms.com"
        Case CCP_lz4
            GetPluginHomepage = "http://www.lz4.org"
        Case CCP_OptiPNG
            GetPluginHomepage = "http://optipng.sourceforge.net/"
        Case CCP_PNGQuant
            GetPluginHomepage = "https://pngquant.org"
        Case CCP_zLib
            GetPluginHomepage = "http://zlib.net"
        Case CCP_zstd
            GetPluginHomepage = "http://www.zstd.net"
    End Select
End Function

'Simplified function for retrieving the license name for a given plugin
Public Function GetPluginLicenseName(ByVal pluginEnumID As CORE_PLUGINS) As String
    Select Case pluginEnumID
        Case CCP_ExifTool
            GetPluginLicenseName = g_Language.TranslateMessage("artistic license")
        Case CCP_EZTwain
            GetPluginLicenseName = g_Language.TranslateMessage("public domain")
        Case CCP_FreeImage
            GetPluginLicenseName = g_Language.TranslateMessage("FreeImage public license")
        Case CCP_LittleCMS
            GetPluginLicenseName = g_Language.TranslateMessage("MIT license")
        Case CCP_lz4
            GetPluginLicenseName = g_Language.TranslateMessage("BSD license")
        Case CCP_OptiPNG
            GetPluginLicenseName = g_Language.TranslateMessage("zLib license")
        Case CCP_PNGQuant
            GetPluginLicenseName = g_Language.TranslateMessage("GNU GPLv3")
        Case CCP_zLib
            GetPluginLicenseName = g_Language.TranslateMessage("zLib license")
        Case CCP_zstd
            GetPluginLicenseName = g_Language.TranslateMessage("BSD license")
    End Select
End Function

'Simplified function for retrieving the license URL for a given plugin
Public Function GetPluginLicenseURL(ByVal pluginEnumID As CORE_PLUGINS) As String
    Select Case pluginEnumID
        Case CCP_ExifTool
            GetPluginLicenseURL = "http://dev.perl.org/licenses/artistic.html"
        Case CCP_EZTwain
            GetPluginLicenseURL = "http://eztwain.com/ezt1faq.htm"
        Case CCP_FreeImage
            GetPluginLicenseURL = "http://freeimage.sourceforge.net/freeimage-license.txt"
        Case CCP_LittleCMS
            GetPluginLicenseURL = "http://www.opensource.org/licenses/mit-license.php"
        Case CCP_lz4
            GetPluginLicenseURL = "https://github.com/lz4/lz4/blob/dev/lib/LICENSE"
        Case CCP_OptiPNG
            GetPluginLicenseURL = "http://optipng.sourceforge.net/license.txt"
        Case CCP_PNGQuant
            GetPluginLicenseURL = "https://github.com/pornel/pngquant/blob/master/COPYRIGHT"
        Case CCP_zLib
            GetPluginLicenseURL = "http://zlib.net/zlib_license.html"
        Case CCP_zstd
            GetPluginLicenseURL = "https://github.com/facebook/zstd/blob/dev/LICENSE"
    End Select
End Function

'If a function requires specialized initialization steps, add them here.  Do NOT add any user-facing interactions (e.g. UI) here,
' and DO NOT account for user preferences here.  (User preferences are handled via ism_PluginAllowed(), above).
'
'This step purely exists to handle custom initialization of plugins, when the master plugin file is known to exist in the
 'official plugin folder, and the user has not forcibly disabled a given plugin.
'
'Returns TRUE if the plugin was initialized successfully; FALSE otherwise.
Private Function InitializePlugin(ByVal pluginEnumID As CORE_PLUGINS) As Boolean
    
    'Because this function has variable complexity (depending on the plugin), an intermediary value is used to track success/failure.
    ' At the end of the function, the function return will simply copy this value, so make sure it's set correctly before the
    ' function ends.
    Dim initializationSuccessful As Boolean
    
    Select Case pluginEnumID
        
        'Unlike most plugins, ExifTool is an .exe file.  Because we interact with it asynchronously, we start it now, then leave
        ' it in "wait" mode.
        Case CCP_ExifTool
            
            'Crashes (or IDE stop button use) can result in stranded ExifTool instances.  As a convenience to the caller, we attempt
            ' to kill any stranded instances before starting new ones.
            If (Not Autosaves.PeekLastShutdownClean) Then
                PDDebug.LogAction "Previous PhotoDemon session terminated unexpectedly.  Performing plugin clean-up..."
                ExifTool.KillStrandedExifToolInstances
            End If
            
            'Attempt to shell a new ExifTool instance.  If shell fails (for whatever reason), the function will return FALSE.
            initializationSuccessful = ExifTool.StartExifTool()
                    
        Case CCP_EZTwain
            initializationSuccessful = Plugin_EZTwain.InitializeEZTwain()
        
        'FreeImage maintains a program-wide handle for the life of the program, which we attempt to generate now.
        Case CCP_FreeImage
            initializationSuccessful = Plugin_FreeImage.InitializeFreeImage()
        
        'LittleCMS maintains a program-wide handle for the life of the program, which we attempt to generate now.
        Case CCP_LittleCMS
            initializationSuccessful = LittleCMS.InitializeLCMS()
        
        'LZ4 maintains a program-wide handle for the life of the program, which we attempt to generate now.
        Case CCP_lz4
            initializationSuccessful = Compression.InitializeCompressionEngine(PD_CE_Lz4, PluginManager.GetPluginPath)
            
        'OptiPNG and PNGQuant are loaded on-demand, as they may not be used in every session
        Case CCP_OptiPNG
            initializationSuccessful = True
            
        Case CCP_PNGQuant
            initializationSuccessful = True
        
        'zLib maintains a program-wide handle for the life of the program, which we attempt to generate now.
        Case CCP_zLib
            initializationSuccessful = Compression.InitializeCompressionEngine(PD_CE_ZLib, PluginManager.GetPluginPath)
        
        'zstd maintains a program-wide handle for the life of the program, which we attempt to generate now.
        Case CCP_zstd
            initializationSuccessful = Compression.InitializeCompressionEngine(PD_CE_Zstd, PluginManager.GetPluginPath)
            
    End Select

    InitializePlugin = initializationSuccessful

End Function

'Most plugins provide a single global "is plugin available" flag, which spares the program from having to plow through
' all these verification steps when it needs to do something with a plugin.
Private Sub SetGlobalPluginFlags(ByVal pluginEnumID As CORE_PLUGINS, ByVal pluginState As Boolean)
    
    Select Case pluginEnumID
    
        Case CCP_ExifTool
            m_ExifToolEnabled = pluginState
                    
        Case CCP_EZTwain
            Plugin_EZTwain.ForciblySetScannerAvailability pluginState
        
        Case CCP_FreeImage
            g_ImageFormats.FreeImageEnabled = pluginState
        
        Case CCP_LittleCMS
            m_LCMSEnabled = pluginState
        
        Case CCP_lz4
            m_lz4Enabled = pluginState
        
        Case CCP_OptiPNG
            m_OptiPNGEnabled = pluginState
        
        Case CCP_PNGQuant
            g_ImageFormats.pngQuantEnabled = pluginState
        
        Case CCP_zLib
            m_ZlibEnabled = pluginState
            
        Case CCP_zstd
            m_ZstdEnabled = pluginState
            
    End Select
    
End Sub

'This final plugin initialization step is OPTIONAL.
'
'It provides a catch-all for custom initialization behavior (e.g. modifying UI bits to reflect plugin-related features).
' New plugins do not need to make use of this functionality.
Private Sub FinalizePluginInitialization(ByVal pluginEnumID As CORE_PLUGINS, ByVal pluginState As Boolean)

    Select Case pluginEnumID
                
        Case CCP_EZTwain
            'EZTwain is currently the only supported method for scanners.  (I hope to fix this in the future.)
            ' As such, availability of the scanner UI is based on EZTwain's successful initialization.
            FormMain.MnuScanImage.Visible = pluginState
            FormMain.MnuSelectScanner.Visible = pluginState
            FormMain.MnuImportSepBar1.Visible = pluginState
        
        Case CCP_FreeImage
            'As of v6.4, PD uses a dedicated callback function to track and report any internal FreeImage errors.
            If pluginState Then Outside_FreeImageV3.FreeImage_InitErrorHandler
            
        Case Else
        
    End Select
    
End Sub

'NOTE: the following function is PLUGIN AGNOSTIC.  You do not need to modify it when adding a new plugin to the program.
'
'This function performs several tasks:
' 1) If the requested plugin file exists in the target folder, great; it returns TRUE and exits.
' 2) If the requested plugin file does NOT exist in the target folder, it scans the program folder to see if it can find it there.
' 3) If it finds a missing plugin in the program folder, it will automatically move the file to the plugin folder, including any
'     helper files (README, LICENSE, etc).
' 4) If the move is successful, it will return TRUE and exit.
Private Function DoesPluginFileExist(ByVal pluginEnumID As CORE_PLUGINS) As Boolean
    
    'Start by getting the filename of the plugin in question
    Dim pluginFilename As String
    pluginFilename = GetPluginFilename(pluginEnumID)
    
    'pdFSO is used for all file interactions
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    'See if the file exists.  If it does, great!  We can exit immediately.
    If Files.FileExists(PluginManager.GetPluginPath & pluginFilename) Then
        DoesPluginFileExist = True
    
    'The plugin file is missing.  Let's see if we can find it.
    Else
    
        PDDebug.LogAction "WARNING!  Plugin ID#" & pluginEnumID & " (" & GetPluginFilename(pluginEnumID) & ") is missing.  Scanning alternate folders..."
    
        Dim extraFiles As pdStringStack
        Set extraFiles = New pdStringStack
    
        'See if the plugin file exists in the base PD folder.  This can happen if a user unknowingly extracts the PD .zip without
        ' folders preserved.
        If Files.FileExists(UserPrefs.GetProgramPath() & pluginFilename) Then
            
            PDDebug.LogAction "UPDATE!  Plugin ID#" & pluginEnumID & " (" & GetPluginFilename(pluginEnumID) & ") was found in the base PD folder.  Attempting to relocate..."
            
            'Move the plugin file to the proper folder
            If cFile.FileCopyW(UserPrefs.GetProgramPath() & pluginFilename, PluginManager.GetPluginPath & pluginFilename) Then
                
                PDDebug.LogAction "UPDATE!  Plugin ID#" & pluginEnumID & " (" & GetPluginFilename(pluginEnumID) & ") was relocated successfully."
                
                'Kill the old plugin instance
                cFile.FileDelete UserPrefs.GetProgramPath() & pluginFilename
                
                'Finally, move any associated files to their new home in the plugin folder
                If GetNonEssentialPluginFiles(pluginEnumID, extraFiles) Then
                    
                    Dim tmpFilename As String
                    
                    Do While extraFiles.PopString(tmpFilename)
                        If cFile.FileCopyW(UserPrefs.GetProgramPath() & tmpFilename, PluginManager.GetPluginPath & tmpFilename) Then
                            cFile.FileDelete UserPrefs.GetProgramPath() & tmpFilename
                        End If
                    Loop
                    
                End If
                
                'Return success!
                DoesPluginFileExist = True
            
            'The file couldn't be moved.  There's probably write issues with the folder structure, in which case the program
            ' as a whole is pretty much doomed.  Exit now.
            Else
                PDDebug.LogAction "WARNING!  Plugin ID#" & pluginEnumID & " (" & GetPluginFilename(pluginEnumID) & ") could not be relocated.  Initialization abandoned."
                DoesPluginFileExist = False
            End If
        
        'If the plugin file doesn't exist in the base folder either, we're SOL.  Exit now.
        Else
            PDDebug.LogAction "WARNING!  Plugin ID#" & pluginEnumID & " (" & GetPluginFilename(pluginEnumID) & ") wasn't found in alternate locations.  Initialization abandoned."
            DoesPluginFileExist = False
        End If
    
    End If
    
End Function

'Convenience wrapper for mass plugin termination.  This function *will* release each plugin's handle, making them unavailable
' for further use.  As such, do not call this until PD is shutting down (and even then, be careful about timing).
'
'Note also that some plugins don't need to be released this way; for example, pngquant and OptiPNG are initialized
' conditionally, if the user needs them, and they are immediately freed after use.
Public Sub TerminateAllPlugins()
    
    'Plugins are released in the order of "how much do we use them", with the most-used plugins being saved for last.
    ' (There's not really a reason for this, except as a failsafe against asynchronous actions happening in the background.)
    Plugin_EZTwain.ReleaseEZTwain
    PDDebug.LogAction "EZTwain released"
    
    Plugin_FreeImage.ReleaseFreeImage
    g_ImageFormats.FreeImageEnabled = False
    PDDebug.LogAction "FreeImage released"
    
    LittleCMS.ReleaseLCMS
    PDDebug.LogAction "LittleCMS released"
    
    If m_ZlibEnabled Then
        Compression.ShutDownCompressionEngine PD_CE_ZLib
        m_ZlibEnabled = False
        PDDebug.LogAction "zLib released"
    End If
    
    If m_ZstdEnabled Then
        Compression.ShutDownCompressionEngine PD_CE_Zstd
        m_ZstdEnabled = False
        PDDebug.LogAction "zstd released"
    End If
    
    If m_lz4Enabled Then
        Compression.ShutDownCompressionEngine PD_CE_Lz4
        m_lz4Enabled = False
        PDDebug.LogAction "lz4 released"
    End If
    
End Sub
