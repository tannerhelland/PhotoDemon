Attribute VB_Name = "PluginManager"
'***************************************************************************
'Core Plugin Manager
'Copyright 2014-2016 by Tanner Helland
'Created: 30/August/15
'Last updated: 30/August/15
'Last update: migrate a ton of scattered plugin management code to this singular module
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
    CCP_OptiPNG = 4
    CCP_PNGQuant = 5
    CCP_zLib = 6
End Enum

#If False Then
    Private Const CCP_ExifTool = 0, CCP_EZTwain = 1, CCP_FreeImage = 2, CCP_LittleCMS = 3, CCP_OptiPNG = 4, CCP_PNGQuant = 5, CCP_zLib = 6
#End If

'Expected version numbers of plugins.  These are updated at each new PhotoDemon release (if a new version of
' the plugin is available, obviously).
Private Const EXPECTED_EXIFTOOL_VERSION As String = "10.12"
Private Const EXPECTED_EZTWAIN_VERSION As String = "1.18.0"
Private Const EXPECTED_FREEIMAGE_VERSION As String = "3.18.0"
Private Const EXPECTED_LITTLECMS_VERSION As String = "2.8.0"
Private Const EXPECTED_OPTIPNG_VERSION As String = "0.7.6"
Private Const EXPECTED_PNGQUANT_VERSION As String = "2.5.2"
Private Const EXPECTED_ZLIB_VERSION As String = "1.2.8"

'This constant is used to iterate all core plugins (as listed under the CORE_PLUGINS enum), so if you add or remove
' a plugin, make sure to update this!
Private Const CORE_PLUGIN_COUNT As Long = 7

'Much of the version-checking code used in this module was derived from http://allapi.mentalis.org/apilist/GetFileVersionInfo.shtml
' Many thanks to those authors for their work on demystifying obscure API calls
Private Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersionl As Integer     ' e.g. = &h0000 = 0
    dwStrucVersionh As Integer     ' e.g. = &h0042 = .42
    dwFileVersionMSl As Integer    ' e.g. = &h0003 = 3
    dwFileVersionMSh As Integer    ' e.g. = &h0075 = .75
    dwFileVersionLSl As Integer    ' e.g. = &h0000 = 0
    dwFileVersionLSh As Integer    ' e.g. = &h0031 = .31
    dwProductVersionMSl As Integer ' e.g. = &h0003 = 3
    dwProductVersionMSh As Integer ' e.g. = &h0010 = .1
    dwProductVersionLSl As Integer ' e.g. = &h0000 = 0
    dwProductVersionLSh As Integer ' e.g. = &h0031 = .31
    dwFileFlagsMask As Long        ' = &h3F for version "0.42"
    dwFileFlags As Long            ' e.g. VFF_DEBUG Or VFF_PRERELEASE
    dwFileOS As Long               ' e.g. VOS_DOS_WINDOWS16
    dwFileType As Long             ' e.g. VFT_DRIVER
    dwFileSubtype As Long          ' e.g. VFT2_DRV_KEYBOARD
    dwFileDateMS As Long           ' e.g. 0
    dwFileDateLS As Long           ' e.g. 0
End Type
Private Declare Function GetFileVersionInfo Lib "Version" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, ByVal Source As Long, ByVal Length As Long)

'To simplify handling throughout this module, plugin existence, allowance, and successful initialization are tracked internally.
' Note that these values ARE NOT EXTERNALLY AVAILABLE, by design; external callers should use the global plugin trackers
' (e.g. g_ZLibEnabled, g_ExifToolEnabled, etc) because those trackers encompass the combination of all these factors, and are
' thus preferable for high-performance code paths.
Private m_PluginExists() As Boolean
Private m_PluginAllowed() As Boolean
Private m_PluginInitialized() As Boolean

Public Function GetNumOfPlugins() As Long
    GetNumOfPlugins = CORE_PLUGIN_COUNT
End Function

'This subroutine handles the detection and installation of all core plugins. required for an optimal PhotoDemon
' experience: zLib, EZTwain32, and FreeImage.  For convenience' sake, it also checks for GDI+ availability.
Public Sub LoadAllPlugins()
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "LoadAllPlugins() called.  Attempting to initialize core plugins..."
    #End If
    
    'Reset all plugin trackers
    ReDim m_PluginExists(0 To CORE_PLUGIN_COUNT - 1) As Boolean
    ReDim m_PluginAllowed(0 To CORE_PLUGIN_COUNT - 1) As Boolean
    ReDim m_PluginInitialized(0 To CORE_PLUGIN_COUNT - 1) As Boolean
    
    'Plugin files are located in the \Data\Plugins subdirectory
    g_PluginPath = g_UserPreferences.GetAppPath & "Plugins\"
    
    'Make sure the plugin path exists
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    If Not cFile.FolderExist(g_PluginPath) Then cFile.CreateFolder g_PluginPath, True
        
    'Plugin loading is handled in a loop.  This loop will call several helper functions, passing each sequential plugin
    ' index as defined by the CORE_PLUGINS enum (and matching CORE_PLUGIN_COUNT const).  Some initialization steps are
    ' shared among all plugins (e.g. checking for the plugin's existence), while some require custom initializations.
    ' This behavior is all carefully documented in the functions called by the initialization loop.
    Dim i As Long
    For i = 0 To CORE_PLUGIN_COUNT - 1
    
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
        
    Next i
    
    'Initialization complete!  In debug builds, write out some plugin debug information.
    #If DEBUGMODE = 1 Then
        
        Dim successfulPluginCount As Long
        successfulPluginCount = 0
        
        For i = 0 To CORE_PLUGIN_COUNT - 1
            If m_PluginInitialized(i) Then
                successfulPluginCount = successfulPluginCount + 1
            Else
                pdDebug.LogAction "WARNING!  Plugin ID#" & i & " (" & GetPluginName(i) & ") was not initialized."
            End If
        Next i
        
        pdDebug.LogAction CStr(successfulPluginCount) & "/" & CStr(CORE_PLUGIN_COUNT) & " plugins initialized successfully."
        
    #End If
    
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
        Case CCP_OptiPNG
            GetPluginFilename = "optipng.exe"
        Case CCP_PNGQuant
            GetPluginFilename = "pngquant.exe"
        Case CCP_zLib
            GetPluginFilename = "zlibwapi.dll"
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
        Case CCP_OptiPNG
            GetPluginName = "OptiPNG"
        Case CCP_PNGQuant
            GetPluginName = "PNGQuant"
        Case CCP_zLib
            GetPluginName = "zLib"
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
        
        'OptiPNG can write its version number to stdout
        Case CCP_OptiPNG
            If PluginManager.IsPluginCurrentlyInstalled(pluginEnumID) Then GetPluginVersion = Plugin_OptiPNG.GetOptiPNGVersion()
            
        'PNGQuant can write its version number to stdout
        Case CCP_PNGQuant
            If PluginManager.IsPluginCurrentlyInstalled(pluginEnumID) Then GetPluginVersion = Plugin_PNGQuant.GetPngQuantVersion()
        
        'All other plugins pull their version info directly from file metadata
        Case Else
            GetPluginVersion = RetrieveGenericVersionString(g_PluginPath & PluginManager.GetPluginFilename(pluginEnumID))
            
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
        
        Case CCP_OptiPNG
            dstStringStack.AddString "optipng-LICENSE.txt"
            
        Case CCP_PNGQuant
            dstStringStack.AddString "pngquant-README.txt"
        
        Case CCP_zLib
            dstStringStack.AddString "zlib-README.txt"
    
    End Select
    
    GetNonEssentialPluginFiles = CBool(dstStringStack.GetNumOfStrings <> 0)
    
End Function

'The Plugin Manager dialog allows the user to forcibly disable plugins.  This can be very helpful when testing bugs and crashes,
' but generally isn't relevant for a casual user.  Regardless, the plugin loader will check this value prior to initializing a plugin.
Private Function IsPluginAllowed(ByVal pluginEnumID As CORE_PLUGINS) As Boolean
    IsPluginAllowed = Not g_UserPreferences.GetPref_Boolean("Plugins", "Force " & PluginManager.GetPluginName(pluginEnumID) & " Disable", False)
End Function

'Simplified function to forcibly disable a plugin via the user's preference file.  Note that this *will not take affect for
' this session* by design; you must subsequently call the SetPluginEnablement() function to live-change the setting.
Public Sub SetPluginAllowed(ByVal pluginEnumID As CORE_PLUGINS, ByVal newEnabledState As Boolean)
    g_UserPreferences.SetPref_Boolean "Plugins", "Force " & PluginManager.GetPluginName(pluginEnumID) & " Disable", Not newEnabledState
End Sub

'Simplified function to detect if a given plugin is currently enabled.  (Plugins can be disabled for a variety of reasons,
' including forcible disablement by the user, bugs, missing files, etc; this catch-all function returns a binary "enabled" state.)
Public Function IsPluginCurrentlyEnabled(ByVal pluginEnumID As CORE_PLUGINS) As Boolean
    Select Case pluginEnumID
        Case CCP_ExifTool
            IsPluginCurrentlyEnabled = g_ExifToolEnabled
        Case CCP_EZTwain
            IsPluginCurrentlyEnabled = Plugin_EZTwain.IsScannerAvailable
        Case CCP_FreeImage
            IsPluginCurrentlyEnabled = g_ImageFormats.FreeImageEnabled
        Case CCP_LittleCMS
            IsPluginCurrentlyEnabled = g_LCMSEnabled
        Case CCP_OptiPNG
            IsPluginCurrentlyEnabled = g_OptiPNGEnabled
        Case CCP_PNGQuant
            IsPluginCurrentlyEnabled = g_ImageFormats.pngQuantEnabled
        Case CCP_zLib
            IsPluginCurrentlyEnabled = g_ZLibEnabled
    End Select
End Function

'Simplified function to forcibly en/disable a given plugin for this session.  Note that this has program-wide repercussions,
' including UI states that may no longer be valid - as such, this function should *not* be changed by anything except the
' Plugin Manager dialog or the plugin initialization functions.
Public Sub SetPluginEnablement(ByVal pluginEnumID As CORE_PLUGINS, ByVal newEnabledState As Boolean)
    Select Case pluginEnumID
        Case CCP_ExifTool
            g_ExifToolEnabled = newEnabledState
        Case CCP_EZTwain
            Plugin_EZTwain.ForciblySetScannerAvailability newEnabledState
        Case CCP_FreeImage
            g_ImageFormats.FreeImageEnabled = newEnabledState
        Case CCP_LittleCMS
            g_LCMSEnabled = newEnabledState
        Case CCP_OptiPNG
            g_OptiPNGEnabled = newEnabledState
        Case CCP_PNGQuant
            g_ImageFormats.pngQuantEnabled = newEnabledState
        Case CCP_zLib
            g_ZLibEnabled = newEnabledState
    End Select
End Sub

'Simplified function to detect if a given plugin is currently installed in PD's plugin folder.  This is (obviously) separate
' from a plugin actually being *enabled*, as that requires initialization and other steps.
Public Function IsPluginCurrentlyInstalled(ByVal pluginEnumID As CORE_PLUGINS) As Boolean
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    IsPluginCurrentlyInstalled = cFile.FileExist(g_PluginPath & GetPluginFilename(pluginEnumID))
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
        Case CCP_OptiPNG
            ExpectedPluginVersion = EXPECTED_OPTIPNG_VERSION
        Case CCP_PNGQuant
            ExpectedPluginVersion = EXPECTED_PNGQUANT_VERSION
        Case CCP_zLib
            ExpectedPluginVersion = EXPECTED_ZLIB_VERSION
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
            GetPluginHomepage = "http://www.littlecms.com/"
        Case CCP_OptiPNG
            GetPluginHomepage = "http://optipng.sourceforge.net/"
        Case CCP_PNGQuant
            GetPluginHomepage = "https://pngquant.org/"
        Case CCP_zLib
            GetPluginHomepage = "http://zlib.net/"
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
        Case CCP_OptiPNG
            GetPluginLicenseName = g_Language.TranslateMessage("zLib license")
        Case CCP_PNGQuant
            GetPluginLicenseName = g_Language.TranslateMessage("GNU GPLv3")
        Case CCP_zLib
            GetPluginLicenseName = g_Language.TranslateMessage("zLib license")
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
        Case CCP_OptiPNG
            GetPluginLicenseURL = "http://optipng.sourceforge.net/license.txt"
        Case CCP_PNGQuant
            GetPluginLicenseURL = "https://raw.githubusercontent.com/pornel/pngquant/master/COPYRIGHT"
        Case CCP_zLib
            GetPluginLicenseURL = "http://zlib.net/zlib_license.html"
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
            If Not peekLastShutdownClean Then
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "Previous PhotoDemon session terminated unexpectedly.  Performing plugin clean-up..."
                #End If
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
        
        'TODO!
        Case CCP_OptiPNG
            initializationSuccessful = True
            
        Case CCP_PNGQuant
            initializationSuccessful = True
        
        Case CCP_zLib
            'zLib maintains a program-wide handle for the life of the program, which we attempt to generate now.
            initializationSuccessful = Plugin_zLib_Interface.InitializeZLib()
            
    End Select

    InitializePlugin = initializationSuccessful

End Function

'Most plugins provide a single global "is plugin available" flag, which spares the program from having to plow through all these
' verification steps when it needs to do something with a plugin.  This step is technically optional, although I prefer global flags
' because they let me use plugins in performance-sensitive areas without worry.
Private Sub SetGlobalPluginFlags(ByVal pluginEnumID As CORE_PLUGINS, ByVal pluginState As Boolean)
    
    Select Case pluginEnumID
    
        Case CCP_ExifTool
            g_ExifToolEnabled = pluginState
                    
        Case CCP_EZTwain
            Plugin_EZTwain.ForciblySetScannerAvailability pluginState
        
        Case CCP_FreeImage
            g_ImageFormats.FreeImageEnabled = pluginState
        
        Case CCP_LittleCMS
            g_LCMSEnabled = pluginState
            
        Case CCP_OptiPNG
            g_OptiPNGEnabled = pluginState
        
        Case CCP_PNGQuant
            g_ImageFormats.pngQuantEnabled = pluginState
        
        Case CCP_zLib
            g_ZLibEnabled = pluginState
            
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
            If pluginState Then
                #If DEBUGMODE = 1 Then
                    Outside_FreeImageV3.FreeImage_InitErrorHandler
                #End If
            End If
            
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
    If cFile.FileExist(g_PluginPath & pluginFilename) Then
        DoesPluginFileExist = True
    
    'The plugin file is missing.  Let's see if we can find it.
    Else
    
        pdDebug.LogAction "WARNING!  Plugin ID#" & pluginEnumID & " (" & GetPluginFilename(pluginEnumID) & ") is missing.  Scanning alternate folders..."
    
        Dim extraFiles As pdStringStack
        Set extraFiles = New pdStringStack
    
        'See if the plugin file exists in the base PD folder.  This can happen if a user unknowingly extracts the PD .zip without
        ' folders preserved.
        If cFile.FileExist(g_UserPreferences.GetProgramPath & pluginFilename) Then
            
            pdDebug.LogAction "UPDATE!  Plugin ID#" & pluginEnumID & " (" & GetPluginFilename(pluginEnumID) & ") was found in the base PD folder.  Attempting to relocate..."
            
            'Move the plugin file to the proper folder
            If cFile.CopyFile(g_UserPreferences.GetProgramPath & pluginFilename, g_PluginPath & pluginFilename) Then
                
                pdDebug.LogAction "UPDATE!  Plugin ID#" & pluginEnumID & " (" & GetPluginFilename(pluginEnumID) & ") was relocated successfully."
                
                'Kill the old plugin instance
                cFile.KillFile g_UserPreferences.GetProgramPath & pluginFilename
                
                'Finally, move any associated files to their new home in the plugin folder
                If GetNonEssentialPluginFiles(pluginEnumID, extraFiles) Then
                    
                    Dim tmpFilename As String
                    
                    Do While extraFiles.PopString(tmpFilename)
                        If cFile.CopyFile(g_UserPreferences.GetProgramPath & tmpFilename, g_PluginPath & tmpFilename) Then
                            cFile.KillFile g_UserPreferences.GetProgramPath & tmpFilename
                        End If
                    Loop
                    
                End If
                
                'Return success!
                DoesPluginFileExist = True
            
            'The file couldn't be moved.  There's probably write issues with the folder structure, in which case the program
            ' as a whole is pretty much doomed.  Exit now.
            Else
                pdDebug.LogAction "WARNING!  Plugin ID#" & pluginEnumID & " (" & GetPluginFilename(pluginEnumID) & ") could not be relocated.  Initialization abandoned."
                DoesPluginFileExist = False
            End If
        
        'If the plugin file doesn't exist in the base folder either, we're SOL.  Exit now.
        Else
            pdDebug.LogAction "WARNING!  Plugin ID#" & pluginEnumID & " (" & GetPluginFilename(pluginEnumID) & ") wasn't found in alternate locations.  Initialization abandoned."
            DoesPluginFileExist = False
        End If
    
    End If
    
End Function

'Given an arbitrary filename, return a string with that file's version (as retrieved from file metadata).
Private Function RetrieveGenericVersionString(ByVal FullFileName As String) As String
    
    'Start by retrieving the required version buffer size (and bail if there's no version info)
    Dim lBufferLen As Long, tmpLong As Long
    lBufferLen = GetFileVersionInfoSize(FullFileName, tmpLong)
    If lBufferLen < 1 Then Exit Function
    
    'Pull the version info into a dedicated struct
    Dim sBuffer() As Byte
    ReDim sBuffer(0 To lBufferLen - 1) As Byte
    tmpLong = GetFileVersionInfo(FullFileName, 0&, lBufferLen, sBuffer(0))
    
    Dim lVerPointer As Long, lVerbufferLen As Long
    tmpLong = VerQueryValue(sBuffer(0), "\", lVerPointer, lVerbufferLen)
    
    Dim udtVerBuffer As VS_FIXEDFILEINFO
    MoveMemory udtVerBuffer, lVerPointer, Len(udtVerBuffer)
    
    'If it proves helpful in the future, here's code for retrieving versioning of the file itself
    'Dim FileVer As String
    'FileVer = Trim(Format$(udtVerBuffer.dwFileVersionMSh)) & "." & Trim(Format$(udtVerBuffer.dwFileVersionMSl)) & "." & Trim(Format$(udtVerBuffer.dwFileVersionLSh)) & "." & Trim(Format$(udtVerBuffer.dwFileVersionLSl))
    
    '...but right now, we're only concerned with product versioning
    RetrieveGenericVersionString = Trim$(Format$(udtVerBuffer.dwProductVersionMSh)) & "." & Trim$(Format$(udtVerBuffer.dwProductVersionMSl)) & "." & Trim$(Format$(udtVerBuffer.dwProductVersionLSh)) & "." & Trim$(Format$(udtVerBuffer.dwProductVersionLSl))
    
End Function
