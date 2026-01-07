Attribute VB_Name = "UserPrefs"
'***************************************************************************
'PhotoDemon User Preferences Manager
'Copyright 2012-2026 by Tanner Helland
'Created: 03/November/12
'Last updated: 06/October/25
'Last update: perform additional validation when loading the user preferences file (count for matched tags),
'             and if validation fails, do a hard reset on the pref file.  (See https://github.com/tannerhelland/PhotoDemon/issues/700)
'
'This is the modern incarnation of PD's old "INI file" module.  It is responsible for managing all
' persistent user settings.
'
'By default, user settings are stored in an XML-ish file in the \Data\ subfolder.  This class will
' generate a default settings file on first run.
'
'Because the settings XML file may receive new settings with each new version, all setting
' interaction functions require the caller to specify a default value (which will be used if
' that setting is requested, but it doesn't exist in the XML).  Also note that if you attempt to
' write a setting, but that setting name or section does not exist, it will automatically be
' appended as a "new" setting at the end of its respective section.
'
'Finally, outside functions should *never* interact with the central XML settings file directly.
' Always pass read/writes through this class.  I cannot guarantee that the XML format or style
' will be consistent between versions, but as long as you use the wrapping functions in this class,
' settings will always behave correctly.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'To make PhotoDemon compatible with the PortableApps spec (http://portableapps.com/), several sub-folders
' are necessary.  These include:
'  /App subfolder, which contains information ESSENTIAL and UNIVERSAL for each PhotoDemon install
'        (e.g. plugin DLLs, language files)
'  /Data subfolder, which contains information that is OPTIONAL and UNIQUE for each PhotoDemon install
'        (e.g. user prefs, saved macros, last-used dialog settings)
Private m_ProgramPath As String
Private m_AppPath As String
Private m_DataPath As String

'Within the /App and /Data folders are additional subfolders, whose purposes should be obvious from their titles

'/App subfolders come first.  These folders should already exist in the downloaded PD .zip, and we will create them
' if they do not exist.
Private m_ThemePath As String
Private m_LanguagePath As String

'/Data subfolders come next.  Note that some of these can be modified at run-time by user behavior -
' e.g. PhotoDemon does not currently ship with a prebuilt palette collection, so its palette path
' changes as the user loads/saves palettes from standalone palette files.
'
'Similarly, some features support additional user paths - like 8bf plugins, which have a default folder
' in the /Data subfolder, but users can also add their own paths through the 8bf dialog; those paths
' are tracked and stored separate from this module (which just exists to ensure a core set of default
' folders exist on every PD install).
Private m_PreferencesPath As String, m_TempPath As String
Private m_MacroPath As String, m_IconPath As String
Private m_ColorProfilePath As String, m_UserLanguagePath As String
Private m_LUTPathDefault As String, m_LUTPathUser As String
Private m_GradientPathDefault As String, m_GradientPathUser As String
Private m_PalettePath As String, m_SelectionPath As String
Private m_8bfPath As String, m_HotkeyPath As String
Private m_FontPath As String    'While PD has its own local font folder, the user can add more folders via Tools > Options

Private m_PresetPath As String        'This folder is a bit different; it is used to store last-used and user-created presets for each tool dialog
Private m_DebugPath As String         'If the user is running a nightly or beta buid, a Debug folder will be created.  Debug and performance dumps
                                    ' are automatically placed here.
Private m_UserThemePath As String     '6.6 nightly builds added prelimianary theme support.  These are currently handled in-memory only, but in
                                    ' the future, themes may be extracted into this (or a matching /Data/) folder.
Private m_UpdatesPath As String       '6.6 greatly improved update support.  Update check and temp files are now stored in a dedicated folder.

'XML engine for reading/writing preference values from file
Private m_XMLEngine As pdXML

'Some preferences are used in performance-sensitive areas.  These preferences are cached internally to improve responsiveness.
' Outside callers can retrieve them via their dedicated functions.
Private m_ThumbnailPerformance As PD_PerformanceSetting, m_ThumbnailInterpolation As GP_InterpolationMode
Private m_CanvasColor As Long

Public Enum PD_DebugLogBehavior
    dbg_Auto = 0
    dbg_False = 1
    dbg_True = 2
End Enum

#If False Then
    Private Const dbg_Auto = 0, dbg_False = 1, dbg_True = 2
#End If

Private m_GenerateDebugLogs As PD_DebugLogBehavior, m_EmergencyDebug As Boolean
Private m_UIFontName As String
Private m_ZoomWithWheel As Boolean

'Prior to v7.0, each dialog stored its preset data to a unique XML file.
' This causes a lot of HDD thrashing as each main window panel retrieves its preset data separately.
' To improve performance, we now use a single central preset file, and individual windows rely on
' this module to manage persistence.
Private m_XMLPresets As pdXML, m_CentralPresetFile As String

'PD runs in portable mode by default, with all data folders assumed present in the same folder
' as PD itself.  If for some reason this is *not* the case, this variable will be flagged.
Private m_NonPortableModeActive As Boolean

'Helper functions for performance-sensitive preferences.
Public Function GetCanvasColor() As Long
    GetCanvasColor = m_CanvasColor
End Function

Public Sub SetCanvasColor(ByVal newColor As Long)
    m_CanvasColor = newColor
End Sub

Public Function GetThumbnailInterpolationPref() As GP_InterpolationMode
    GetThumbnailInterpolationPref = m_ThumbnailInterpolation
End Function

Public Function GetThumbnailPerformancePref() As PD_PerformanceSetting
    GetThumbnailPerformancePref = m_ThumbnailPerformance
End Function

Public Sub SetThumbnailPerformancePref(ByVal newSetting As PD_PerformanceSetting)
    m_ThumbnailPerformance = newSetting
    If (newSetting = PD_PERF_BESTQUALITY) Then
        m_ThumbnailInterpolation = GP_IM_HighQualityBicubic
    ElseIf (newSetting = PD_PERF_BALANCED) Then
        m_ThumbnailInterpolation = GP_IM_Bilinear
    ElseIf (newSetting = PD_PERF_FASTEST) Then
        m_ThumbnailInterpolation = GP_IM_NearestNeighbor
    End If
End Sub

'Each PD "install" (e.g. each preferences file) gets assigned a random ID.  This is useful for
' e.g. instancing purposes.
Public Function GetUniqueAppID() As String
    GetUniqueAppID = UserPrefs.GetPref_String("Core", "UniqueAppID", OS.GetArbitraryGUID(False), True)
End Function

Public Function GenerateDebugLogs() As Boolean
    If (m_GenerateDebugLogs = dbg_Auto) Then
        GenerateDebugLogs = ((PD_BUILD_QUALITY <> PD_PRODUCTION) Or m_EmergencyDebug) And PDMain.IsProgramRunning()
    ElseIf (m_GenerateDebugLogs = dbg_False) Then
        GenerateDebugLogs = False
    Else
        GenerateDebugLogs = True
    End If
End Function

Public Function GetDebugLogPreference() As PD_DebugLogBehavior
    GetDebugLogPreference = m_GenerateDebugLogs
End Function

Public Sub SetDebugLogPreference(ByVal newPref As PD_DebugLogBehavior)
    If (newPref <> m_GenerateDebugLogs) Then
        m_GenerateDebugLogs = newPref
        UserPrefs.SetPref_Long "Core", "GenerateDebugLogs", m_GenerateDebugLogs
    End If
End Sub

Public Sub SetEmergencyDebugger(ByVal newState As Boolean)
    m_EmergencyDebug = newState
End Sub

'Non-portable mode means PD has been extracted to an access-restricted folder.  The program (should) still run normally,
' with silent redirection to the local appdata folder, but in-place automatic upgrades will be disabled (as we don't
' have write access to our own folder, alas).
Public Function IsNonPortableModeActive() As Boolean
    IsNonPortableModeActive = m_NonPortableModeActive
End Function

'Get the current Theme path.  Note that there are /App (program default) and /Data (userland) variants of this folder.
Public Function GetThemePath(Optional ByVal getUserThemePathInstead As Boolean = False) As String
    If getUserThemePathInstead Then GetThemePath = m_UserThemePath Else GetThemePath = m_ThemePath
End Function

'Get/set subfolders from the user's /Data folder.  These paths may not exist at run-time, so always ensure that code
' works even if these paths are not available.
'
'Similarly, not all folders support a SetXYZPath partner.  This is a deliberate choice, and a detailed explanation
' varies according to path type.  (Some paths are internal PD ones that never vary.  Other paths are fallbacks only,
' and the user has mechanisms for adding additional or different paths through dialogs other than the Preferences one.)
Public Function Get8bfPath() As String
    Get8bfPath = m_8bfPath
End Function

Public Function GetDebugPath() As String
    GetDebugPath = m_DebugPath
End Function

Public Function GetColorProfilePath() As String
    GetColorProfilePath = m_ColorProfilePath
End Function

Public Sub SetColorProfilePath(ByRef newPath As String)
    m_ColorProfilePath = Files.PathAddBackslash(Files.FileGetPath(newPath))
    SetPref_String "Paths", "ColorProfiles", m_ColorProfilePath
End Sub

Public Function GetFontPath() As String
    GetFontPath = m_FontPath
End Function

Public Function GetGradientPath(Optional ByVal useDefaultLocation As Boolean = False) As String
    If useDefaultLocation Then
        GetGradientPath = m_GradientPathDefault
    Else
        GetGradientPath = m_GradientPathUser
    End If
End Function

'There are two gradient paths at present; a "default" one that marks PD's gradient collection folder
' (hard-coded to /Data/Gradients), and a user-editable one that auto-updates when individual gradients
' are imported/exported from standalone files - e.g. the equivalent of a "last-used gradient" path.
'
'This function sets the "last-used gradient" path.
Public Sub SetGradientPath(ByRef newPath As String)
    m_GradientPathUser = Files.PathAddBackslash(Files.FileGetPath(newPath))
    SetPref_String "Paths", "Gradients", m_GradientPathUser
End Sub

Public Function GetHotkeyPath() As String
    GetHotkeyPath = m_HotkeyPath
End Function

Public Sub SetHotkeyPath(ByRef newPath As String)
    m_HotkeyPath = Files.PathAddBackslash(Files.FileGetPath(newPath))
    SetPref_String "Paths", "Hotkeys", m_HotkeyPath
End Sub

Public Function GetLUTPath(Optional ByVal useDefaultLocation As Boolean = False) As String
    If useDefaultLocation Then
        GetLUTPath = m_LUTPathDefault
    Else
        GetLUTPath = m_LUTPathUser
    End If
End Function

'There are two 3DLUT paths at present; a "default" one that marks PD's 3DLUT collection folder
' (hard-coded to /Data/3DLUTs), and a user-editable one that auto-updates when individual LUTs
' are imported/exported from standalone files - e.g. the equivalent of a "last-used LUT" path.
'
'This function sets the "last-used LUT" path.
Public Sub SetLUTPath(ByRef newPath As String)
    m_LUTPathUser = Files.PathAddBackslash(Files.FileGetPath(newPath))
    SetPref_String "Paths", "LUTs", m_LUTPathUser
End Sub

Public Function GetPalettePath() As String
    GetPalettePath = m_PalettePath
End Function

Public Sub SetPalettePath(ByRef newPath As String)
    m_PalettePath = Files.PathAddBackslash(Files.FileGetPath(newPath))
    SetPref_String "Paths", "Palettes", m_PalettePath
End Sub

Public Function GetPresetPath() As String
    GetPresetPath = m_PresetPath
End Function

'Get/set the current Selection directory
Public Function GetSelectionPath() As String
    GetSelectionPath = m_SelectionPath
End Function

Public Sub SetSelectionPath(ByRef newSelectionPath As String)
    m_SelectionPath = Files.PathAddBackslash(Files.FileGetPath(newSelectionPath))
    SetPref_String "Paths", "Selections", m_SelectionPath
End Sub

'Return the current Language directory
Public Function GetLanguagePath(Optional ByVal getUserLanguagePathInstead As Boolean = False) As String
    If getUserLanguagePathInstead Then GetLanguagePath = m_UserLanguagePath Else GetLanguagePath = m_LanguagePath
End Function

'Return the current temporary directory, as specified by the user's preferences.  (Note that this may not be the
' current Windows system temp path, if the user has opted to change it.)
Public Function GetTempPath() As String
    GetTempPath = m_TempPath
End Function

'Set the current temp directory
Public Sub SetTempPath(ByVal newTempPath As String)
    
    'If the folder exists and is writable as-is, great: save it and exit
    Dim doesFolderExist As Boolean
    doesFolderExist = Files.PathExists(newTempPath, True)
    If (Not doesFolderExist) Then doesFolderExist = Files.PathExists(Files.PathAddBackslash(newTempPath), True)
    
    If doesFolderExist Then
        m_TempPath = Files.PathAddBackslash(newTempPath)
        
    'If it doesn't exist, make sure the user didn't do something weird, like supply a file instead of a folder
    Else
    
        newTempPath = Files.PathAddBackslash(Files.FileGetPath(newTempPath))
        
        'Test the path again
        doesFolderExist = Files.PathExists(newTempPath, True)
        If (Not doesFolderExist) Then doesFolderExist = Files.PathExists(Files.PathAddBackslash(newTempPath), True)
    
        If doesFolderExist Then
            m_TempPath = Files.PathAddBackslash(newTempPath)
            
        'If it still fails, revert to the default system temp path
        Else
            m_TempPath = OS.SystemTempPath()
        End If
    
    End If
    
    'Write the final path out to file
    SetPref_String "Paths", "TempFiles", m_TempPath
    
End Sub

'Return the current program directory
Public Function GetProgramPath() As String
    GetProgramPath = m_ProgramPath
End Function

'Return the current app data directory
Public Function GetAppPath() As String
    GetAppPath = m_AppPath
End Function

'Return the current user data directory
Public Function GetDataPath() As String
    GetDataPath = m_DataPath
End Function

'Return the current macro directory
Public Function GetMacroPath() As String
    GetMacroPath = m_MacroPath
End Function

'Set the current macro directory
Public Sub SetMacroPath(ByRef newMacroPath As String)
    m_MacroPath = Files.PathAddBackslash(Files.FileGetPath(newMacroPath))
    SetPref_String "Paths", "Macro", m_MacroPath
End Sub

'Return the current MRU icon directory
Public Function GetIconPath() As String
    GetIconPath = m_IconPath
End Function

'Return the current update-specific temp path
Public Function GetUpdatePath() As String
    GetUpdatePath = m_UpdatesPath
End Function

'Get the user's preferred UI font name (if any; this defaults to "Segoe UI")
Public Function GetUIFontName() As String
    GetUIFontName = m_UIFontName
End Function

'By default, Ctrl+Mousewheel zooms.  The user can change this behavior from the Tools > Options > Interface panel.
Public Function GetZoomWithWheel() As Boolean
    GetZoomWithWheel = m_ZoomWithWheel
End Function

Public Sub SetZoomWithWheel(ByVal newValue As Boolean)
    m_ZoomWithWheel = newValue
End Sub

'Initialize key program directories.  If this function fails, PD will fail to load.
Public Function InitializePaths() As Boolean
    
    InitializePaths = True
    
    'First things first: figure out where this .exe was launched from
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    m_ProgramPath = cFile.AppPathW
    
    'If this is the first time PhotoDemon is run, we need to create a series of data folders.
    ' Because PD is a portable application, we default to creating those folders in our own
    ' application folder.  Unfortunately, some users do dumb things like put PD inside protected
    ' system folders, which causes this step to fail.  We try to handle this situation gracefully,
    ' by redirecting those folders to the current user's AppData folder.
    m_NonPortableModeActive = False
    
    'Anyway, before doing anything else, let's make sure we actually have write access to our own
    ' damn folder; if we don't, we can activate what I call "non-portable" mode.
    Dim localAppDataPath As String
    localAppDataPath = OS.SpecialFolder(CSIDL_LOCAL_APPDATA)
    
    Dim baseFolder As String
    
    Dim tmpFileWrite As String, tmpHandle As Long
    tmpFileWrite = m_ProgramPath & "tmp.tmp"
    m_NonPortableModeActive = (Not cFile.FileCreateHandle(tmpFileWrite, tmpHandle, True, True, OptimizeNone))
    cFile.FileCloseHandle tmpHandle
    cFile.FileDelete tmpFileWrite
        
    If m_NonPortableModeActive Then
        
        PDDebug.LogAction "WARNING!  Portable mode has been deactivated due to folder rights.  Attempting to salvage session..."
        
        'Because we don't have access to our own folder, we need a plan B for PD's expected user data folders.
        ' (Note that we still need access to required plugin DLLs, which must exist *somewhere* we can access.)
        
        'Our only real option is to silently redirect the program's settings subfolder to known-good folder, in this case
        ' the standard local app storage folder.
        baseFolder = localAppDataPath & "PhotoDemon\"
        If (Not Files.PathExists(baseFolder)) Then
            
            If (Not Files.PathCreate(baseFolder)) Then
            
                'Something has gone horrifically wrong.  I'm not sure what to do except let the program
                ' crash and burn.
                InitializePaths = False
                Exit Function
            
            End If
            
        End If
        
        'If we're still here, we were able to create a data folder in a backup location.
        ' Try to proceed with the load process.
        PDDebug.LogAction "Non-portable mode activated successfully.  Continuing program initialization..."
        
    'This is a normal portable session.  The base folder is the same as PD's app path.
    Else
        baseFolder = m_ProgramPath
    End If
    
    'Ensure we have access to an "App" subfolder - this is where essential application files (like plugins)
    ' are stored.  (In portable mode, we can create this folder as necessary; this typically only happens
    ' when a user uses a shitty 3rd-party zip program that doesn't preserve zip folder structure, and
    ' everything gets dumped into the base folder.)
    m_AppPath = m_ProgramPath & "App\"
    If (Not Files.PathExists(m_AppPath)) Then InitializePaths = Files.PathCreate(m_AppPath)
    If (Not InitializePaths) Then Exit Function
    
    m_AppPath = m_AppPath & "PhotoDemon\"
    If (Not Files.PathExists(m_AppPath)) Then InitializePaths = Files.PathCreate(m_AppPath)
    If (Not InitializePaths) Then Exit Function
    
    'Within the App\PhotoDemon\ folder, create a folder for any available OFFICIAL translations.  (User translations go in the Data folder.)
    m_LanguagePath = m_AppPath & "Languages\"
    If (Not Files.PathExists(m_LanguagePath)) Then Files.PathCreate m_LanguagePath
    
    'Within the App\PhotoDemon\ folder, create a folder for any available OFFICIAL themes.  (User themes go in the Data folder.)
    m_ThemePath = m_AppPath & "Themes\"
    If (Not Files.PathExists(m_ThemePath)) Then Files.PathCreate m_ThemePath
    
    'We are now guaranteed an /App subfolder ready for use.  (Note that we haven't checked for plugins - that's handled
    ' by the separate PluginManager module.)
    
    'Next, create a "Data" path based off the base folder we determined above.  (In non-portable mode,
    ' this points at the user's Local folder.)  This is where the preferences file and any other user-specific
    ' files (saved filters, macros) get stored.
    m_DataPath = baseFolder & "Data\"
    If (Not Files.PathExists(m_DataPath)) Then
        
        Dim needToCreateDataFolder As Boolean
        needToCreateDataFolder = True
        
        'If the data path is missing, there are two possible explanations:
        ' 1) This is the first time the user has run PD, meaning the data folder simply needs to be created.
        ' 2) This is *not* the first time the user has run PD, but they've moved PD to a new location.
        '    Normally this isn't a problem - *unless* they've orphaned a data folder somewhere else.
        '    (While this is a rare possibility, PD auto-detects installs to system folders, and it actively
        '     encourages the user to move the application somewhere else - so if the user follows our good advice,
        '     we want to reward them by redirecting the data folder to its original location, so they don't lose
        '     any of their settings or recently-used lists.)
        
        'If we're running in portable mode, look for an existing (orphaned) data folder in local app storage.
        If (Not m_NonPortableModeActive) Then
        
            If Files.PathExists(localAppDataPath & "PhotoDemon\Data\") Then
                m_DataPath = localAppDataPath & "PhotoDemon\Data\"
                needToCreateDataFolder = False
            End If
        
        'In portable mode, we have write-access to our own folder, so create the data folder and carry on!
        End If
        
        'If we didn't find a data folder in a non-standard location, go ahead and create it wherever the
        ' current base folder points.  (In a portable install, this will be PD's application path;
        ' otherwise, it will be the standard local app storage folder inside the \Users folder.)
        If needToCreateDataFolder Then Files.PathCreate m_DataPath
        
    End If
    
    PDDebug.LogAction "PD base folder is " & m_ProgramPath
    PDDebug.LogAction "PD data folder points at " & m_DataPath
    
    'Within the \Data subfolder, check for additional user folders - saved macros, filters, selections, etc...
    m_8bfPath = m_DataPath & "8bfPlugins\"
    If (Not Files.PathExists(m_8bfPath)) Then Files.PathCreate m_8bfPath
    
    m_ColorProfilePath = m_DataPath & "ColorProfiles\"
    If (Not Files.PathExists(m_ColorProfilePath)) Then Files.PathCreate m_ColorProfilePath
    
    m_DebugPath = m_DataPath & "Debug\"
    If (Not Files.PathExists(m_DebugPath)) Then Files.PathCreate m_DebugPath
    
    m_FontPath = m_DataPath & "Fonts\"
    If (Not Files.PathExists(m_FontPath)) Then Files.PathCreate m_FontPath
    
    m_GradientPathDefault = m_DataPath & "Gradients\"
    m_GradientPathUser = m_GradientPathDefault  'This will be overwritten with the user's current path, if any, in a subsequent step
    If (Not Files.PathExists(m_GradientPathDefault)) Then Files.PathCreate m_GradientPathDefault
    
    m_HotkeyPath = m_DataPath & "Hotkeys\"
    If (Not Files.PathExists(m_HotkeyPath)) Then Files.PathCreate m_HotkeyPath
    
    m_LUTPathDefault = m_DataPath & "3DLUTs\"
    m_LUTPathUser = m_LUTPathDefault  'This will be overwritten with the user's current path, if any, in a subsequent step
    If (Not Files.PathExists(m_LUTPathDefault)) Then Files.PathCreate m_LUTPathDefault
    
    m_IconPath = m_DataPath & "Icons\"
    If (Not Files.PathExists(m_IconPath)) Then Files.PathCreate m_IconPath
    
    m_UserLanguagePath = m_DataPath & "Languages\"
    If (Not Files.PathExists(m_UserLanguagePath)) Then Files.PathCreate m_UserLanguagePath
    
    m_MacroPath = m_DataPath & "Macros\"
    If (Not Files.PathExists(m_MacroPath)) Then Files.PathCreate m_MacroPath
    
    m_PalettePath = m_DataPath & "Palettes\"
    If (Not Files.PathExists(m_PalettePath)) Then Files.PathCreate m_PalettePath
    
    m_PresetPath = m_DataPath & "Presets\"
    If (Not Files.PathExists(m_PresetPath)) Then Files.PathCreate m_PresetPath
    
    m_SelectionPath = m_DataPath & "Selections\"
    If (Not Files.PathExists(m_SelectionPath)) Then Files.PathCreate m_SelectionPath
    
    m_UserThemePath = m_DataPath & "Themes\"
    If (Not Files.PathExists(m_UserThemePath)) Then Files.PathCreate m_UserThemePath
    
    m_UpdatesPath = m_DataPath & "Updates\"
    If (Not Files.PathExists(m_UpdatesPath)) Then Files.PathCreate m_UpdatesPath
    
    'After all paths have been validated, we sometimes need to perform path clean-up.  This is required if new builds
    ' change where key PhotoDemon files are stored, or renames key files.  (Without this, we risk leaving behind
    ' duplicate files between builds.)
    PerformPathCleanup
    
    'The user preferences file is also located in the \Data folder.  We don't actually load it yet; this is handled
    ' by the (rather large) LoadUserSettings() function.
    m_PreferencesPath = m_DataPath & "PhotoDemon_settings.xml"
    
    'Last-used dialog settings are also located in the \Presets subfolder; this file *is* loaded now, if it exists.
    m_CentralPresetFile = m_PresetPath & "MainPanels.xml"
    If (m_XMLPresets Is Nothing) Then Set m_XMLPresets = New pdXML
    
    If Files.FileExists(m_CentralPresetFile) Then
        If m_XMLPresets.LoadXMLFile(m_CentralPresetFile) Then
            If (Not m_XMLPresets.IsPDDataType("Presets")) Then m_XMLPresets.PrepareNewXML "Presets"
        End If
    Else
        m_XMLPresets.PrepareNewXML "Presets"
    End If
        
End Function

Private Sub PerformPathCleanup()
    
    'This step is pointless if we are in non-portable mode
    If m_NonPortableModeActive Then Exit Sub
    
    '****
    '6.6 > 7.0 RELEASE CLEANUP
    
    'In PD 7.0, I switched to distributing text files like README.txt as markdown files (README.md).
    ' This spares me from maintaining duplicate copies, and it ensures that the actual README used
    ' on GitHub is identical to the one downloaded from photodemon.org.
    
    'To prevent duplicate copies, check for any leftover.txt instances and remove them.
    ' (Note that we explicitly check file size to avoid removing files that are not ours.)
    Dim targetFilename As String, replaceFilename As String
    targetFilename = UserPrefs.GetProgramPath & "README.txt"
    replaceFilename = UserPrefs.GetProgramPath & "README.md"
    If (Files.FileExists(targetFilename) And Files.FileExists(replaceFilename)) Then
    
        'Check filesize.  This uses magic numbers taken from the official 6.6 release file sizes.
        If (Files.FileLenW(targetFilename) = 13364&) Then Files.FileDelete targetFilename
        
    End If
    
    'Repeat above steps for LICENSE.md
    targetFilename = UserPrefs.GetProgramPath & "LICENSE.txt"
    replaceFilename = UserPrefs.GetProgramPath & "LICENSE.md"
    If Files.FileExists(targetFilename) And Files.FileExists(replaceFilename) Then If (Files.FileLenW(targetFilename) = 30659&) Then Files.FileDelete targetFilename
    
    'END 6.6 > 7.0 RELEASE CLEANUP
    '****

End Sub

'Load all user settings from file
Public Sub LoadUserSettings()
    
    'Ensure the preferences file...
    ' 1) exists (it will be created as necessary), and...
    ' 2) has a valid header, and...
    ' 3) is explicitly tagged as a PD settings file, and...
    ' 4) looks to have "mostly" valid XML syntax (PD wants left/right angle bracket count to match)
    If LoadAndValidatePrefFile() Then
        
        'The user preferences file loaded and validated OK.
        
        'Pull the temp file path from the preferences file and make sure it exists.
        ' (If it doesn't, transparently set it to the system temp path.)
        m_TempPath = GetPref_String("Paths", "TempFiles", vbNullString)
        If (Not Files.PathExists(m_TempPath)) Then
            m_TempPath = OS.SystemTempPath()
            SetPref_String "Paths", "TempFiles", m_TempPath
        End If
        
        'Pull all other stored paths
        m_8bfPath = GetPref_String("Paths", "8bf", m_8bfPath)
        m_ColorProfilePath = GetPref_String("Paths", "ColorProfiles", m_ColorProfilePath)
        m_GradientPathUser = GetPref_String("Paths", "Gradients", m_GradientPathDefault)
        m_LUTPathUser = GetPref_String("Paths", "LUTs", m_LUTPathDefault)
        m_MacroPath = GetPref_String("Paths", "Macro", m_MacroPath)
        m_PalettePath = GetPref_String("Paths", "Palettes", m_PalettePath)
        m_SelectionPath = GetPref_String("Paths", "Selections", m_SelectionPath)
            
        'Check if the user wants us to prompt them about closing unsaved images
        g_ConfirmClosingUnsaved = GetPref_Boolean("Saving", "ConfirmClosingUnsaved", True)
        
        'Grab the last-used common dialog filters
        g_LastOpenFilter = GetPref_Long("Core", "LastOpenFilter", 1)    'Default to "All compatible images"
        g_LastSaveFilter = GetPref_Long("Core", "LastSaveFilter", PD_USER_PREF_UNKNOWN)
        
        'For performance reasons, cache any performance-related settings.
        ' (This is much faster than reading preferences from file every time they're needed.)
        g_InterfacePerformance = UserPrefs.GetPref_Long("Performance", "InterfaceDecorationPerformance", PD_PERF_BALANCED)
        UserPrefs.SetThumbnailPerformancePref UserPrefs.GetPref_Long("Performance", "ThumbnailPerformance", PD_PERF_BALANCED)
        g_ViewportPerformance = UserPrefs.GetPref_Long("Performance", "ViewportRenderPerformance", PD_PERF_BALANCED)
        g_UndoCompressionLevel = UserPrefs.GetPref_Long("Performance", "UndoCompression", 1)
        
        m_GenerateDebugLogs = UserPrefs.GetPref_Long("Core", "GenerateDebugLogs", 0)
        Tools.SetToolSetting_HighResMouse UserPrefs.GetPref_Boolean("Tools", "HighResMouseInput", True)
        m_CanvasColor = Colors.GetRGBLongFromHex(UserPrefs.GetPref_String("Interface", "CanvasColor", "#a0a0a0"))
        
        Drawing.ToggleShowOptions pdst_LayerEdges, True, UserPrefs.GetPref_Boolean("Interface", "show-layeredges", False)
        Drawing.ToggleShowOptions pdst_SmartGuides, True, UserPrefs.GetPref_Boolean("Interface", "show-smartguides", True)
        Snap.ToggleSnapOptions pdst_Global, True, UserPrefs.GetPref_Boolean("Interface", "snap-global", True)
        Snap.ToggleSnapOptions pdst_CanvasEdge, True, UserPrefs.GetPref_Boolean("Interface", "snap-canvas-edge", True)
        Snap.ToggleSnapOptions pdst_Centerline, True, UserPrefs.GetPref_Boolean("Interface", "snap-centerline", False)
        Snap.ToggleSnapOptions pdst_Layer, True, UserPrefs.GetPref_Boolean("Interface", "snap-layer", True)
        Snap.ToggleSnapOptions pdst_Angle90, True, UserPrefs.GetPref_Boolean("Interface", "snap-angle-90", True)
        Snap.ToggleSnapOptions pdst_Angle45, True, UserPrefs.GetPref_Boolean("Interface", "snap-angle-45", True)
        Snap.ToggleSnapOptions pdst_Angle30, True, UserPrefs.GetPref_Boolean("Interface", "snap-angle-30", False)
        Snap.SetSnap_Degrees UserPrefs.GetPref_Float("Interface", "snap-degrees", 7.5)
        Snap.SetSnap_Distance UserPrefs.GetPref_Long("Interface", "snap-distance", 8&)
        
        m_ZoomWithWheel = UserPrefs.GetPref_Boolean("Interface", "wheel-zoom", False)
        
        m_UIFontName = UserPrefs.GetPref_String("Interface", "UIFont", vbNullString, False)
        
    Else
        PDDebug.LogAction "WARNING! UserPrefs.LoadUserSettings() failed.  This session is unrecoverable."
    End If
                
End Sub

'Load the user's preference file, then perform a few basic validations on it.
'
'*IF* the current preference file fails validations, this function will forcibly re-create it from scratch
' (if it can) then attempt to re-validate the fresh copy.
'
'If this function returns FALSE, the preference file is not just messed up - re-creating it also failed,
' which points to system-level problems.  Sessions with failed preference loading are likely to experience
' instability and/or outright crashes.
Private Function LoadAndValidatePrefFile() As Boolean
    
    LoadAndValidatePrefFile = False
    
    'If no preferences file exists, construct a default one
    If (Not Files.FileExists(m_PreferencesPath)) Then
        PDDebug.LogAction "WARNING!  UserPrefs.LoadUserSettings couldn't find a pref file.  Creating a new one now..."
        CreateNewPreferencesFile
    End If
    
    'Load the file (requires read access only)
    Dim prefFileOK As Boolean
    prefFileOK = m_XMLEngine.LoadXMLFile(m_PreferencesPath)
    If prefFileOK Then
        
        'Check the XML type and ensure it is "user preferences"
        prefFileOK = m_XMLEngine.IsPDDataType("User Preferences")
    
    'If the file couldn't be loaded, it's probably an access issue (which we can't resolve programmatically).
    Else
        
        'Attempt to create the file one last time
        PDDebug.LogAction "WARNING!  UserPrefs.LoadUserSettings *still* couldn't find a pref file.  Attempting to re-create..."
        CreateNewPreferencesFile
        prefFileOK = m_XMLEngine.LoadXMLFile(m_PreferencesPath)
        If (Not prefFileOK) Then
            PDDebug.LogAction "Re-creating pref file failed.  Session borked."
            Exit Function
        Else
            prefFileOK = m_XMLEngine.IsPDDataType("User Preferences")
        End If
        
    End If
    
    'If the preferences file exists *and* has a correct header and data type, perform basic XML validation.
    If prefFileOK Then
    
        'Perform a secondary check for mismatched left/right tag counts.
        ' (User-edited files may contain errors, and errors in parsing the settings file
        '  can break PD in unpredictable ways.)
        prefFileOK = m_XMLEngine.BasicXMLValidation()
        If prefFileOK Then
            
            'Success!  Settings look good - proceed normally.
            LoadAndValidatePrefFile = True
            Exit Function
        
        'If the pref file has unclosed tags, reset it to allow this session to load at all,
        ' then repeat all previous safety checks.
        Else
            
            PDDebug.LogAction "WARNING: pref file reported mismatched tag counts; resetting for safety..."
            CreateNewPreferencesFile
            
            prefFileOK = m_XMLEngine.LoadXMLFile(m_PreferencesPath) And m_XMLEngine.IsPDDataType("User Preferences")
            If prefFileOK Then
                PDDebug.LogAction "Resetting seemed to work OK.  Continuing with load..."
            Else
                PDDebug.LogAction "WARNING: reset still produced an invalid file.  Session borked."
            End If
            
            If prefFileOK Then
                prefFileOK = m_XMLEngine.BasicXMLValidation()
                If prefFileOK Then
                    PDDebug.LogAction "Fresh reset passed validation.  Session should proceed normally."
                    LoadAndValidatePrefFile = True
                    Exit Function
                Else
                    PDDebug.LogAction "WARNING: pref file *still* has 1+ unclosed tags.  Session borked."
                End If
            End If
            
        End If
        
    Else
        PDDebug.LogAction "WARNING: pref file (still) doesn't exist or has wrong PD data type; session borked."
    End If
    
End Function

'Reset the preferences file to its default state.  (Basically, delete any existing file, then create a new one from scratch.)
Public Sub ResetPreferences()
    PDDebug.LogAction "WARNING!  pdPreferences.ResetPreferences() has been called.  Any previous settings will now be erased."
    Files.FileDeleteIfExists m_PreferencesPath
    CreateNewPreferencesFile
    LoadUserSettings
    Fonts.DetermineUIFont
End Sub

'Create a new preferences XML file from scratch.  When new preferences are added to the preferences dialog, they should also be
' added to this function, to ensure that the most intelligent preference is selected by default.
Private Sub CreateNewPreferencesFile()

    'This function is used to determine whether PhotoDemon is being run for the first time.  Why do it here?
    ' 1) When first downloaded, PhotoDemon doesn't come with a prefs file.  Thus this routine MUST be called.
    ' 2) When preferences are reset, this file is deleted.  That is an appropriate time to mark the program as
    '     "first run", to ensure that any first-run dialogs are also reset.
    ' 3) If the user moves PhotoDemon but leaves behind the old prefs file.  There's no easy way to check this,
    '     but treating the program like it's being run for the first time is as good a plan as any.
    g_IsFirstRun = True
    
    'As a failsafe against data corruption, if this is determined to be a first run, we also delete some
    ' settings-related files in the Presets folder (if they exist).
    If g_IsFirstRun Then Files.FileDeleteIfExists m_PresetPath & "Program_WindowLocations.xml"
    
    'Reset our XML engine
    With m_XMLEngine
        
        .PrepareNewXML "User Preferences"
        .WriteBlankLine
    
        'Write out a comment marking the date and build of this preferences code; this can be helpful when debugging
        .WriteComment "This preferences file was created on " & Format$(Now, "dd-mmm-yyyy") & " by app version " & Updates.GetPhotoDemonVersionCanonical()
        .WriteBlankLine
        
        'New in v8.0 are auto-constructed assets for various tools.  These are just folders of
        ' standalone files that populate various "collections" in the program - e.g. the default
        ' gradient files that ship for the gradient tool.  To avoid overwriting user changes,
        ' we only attempt to extract these once, when the program is run for the first time
        ' (or if the preference is encountered for the first time).  After that point,
        ' the assets are never extracted again.
        .WriteTag "Assets", vbNullString, True
            .WriteTag "ExtractedGradients", "False"
        .CloseTag "Assets"
        .WriteBlankLine
        
        .WriteTag "BatchProcess", vbNullString, True
            .WriteTag "InputFolder", OS.SpecialFolder(CSIDL_MYPICTURES)
            .WriteTag "ListFolder", OS.SpecialFolder(CSIDL_MYPICTURES)
            .WriteTag "OutputFolder", OS.SpecialFolder(CSIDL_MYPICTURES)
        .CloseTag "BatchProcess"
        .WriteBlankLine
    
        'Write out the "color management" block of preferences:
        .WriteTag "ColorManagement", vbNullString, True
            .WriteTag "DisplayCMMode", Trim$(Str$(DCM_NoManagement))
            .WriteTag "DisplayBPC", "True"
            .WriteTag "DisplayRenderingIntent", Trim$(Str$(INTENT_PERCEPTUAL))
        .CloseTag "ColorManagement"
        .WriteBlankLine
        
        'Write out the "core" block of preferences.  These are preferences that PD uses internally.  These are never directly
        ' exposed to the user (e.g. the user cannot toggle these from the Preferences dialog).
        .WriteTag "Core", vbNullString, True
            .WriteTag "DisplayIDEWarning", "True"
            .WriteTag "GenerateDebugLogs", "0"     'Default to "automatic" debug log behavior
            .WriteTag "HasGitHubAccount", vbNullString
            .WriteTag "LastOpenFilter", "1"        'Default to "All Compatible Graphics" filter for loading
            .WriteTag "LastPreferencesPage", "0"
            .WriteTag "LastSaveFilter", Trim$(Str$(PD_USER_PREF_UNKNOWN))  'Mark the last-used save filter as "unknown"
            .WriteTag "LastWindowState", "0"
            .WriteTag "LastWindowLeft", "1"
            .WriteTag "LastWindowTop", "1"
            .WriteTag "LastWindowWidth", "1"
            .WriteTag "LastWindowHeight", "1"
            .WriteTag "SessionsSinceLastCrash", "-1"
        .CloseTag "Core"
        .WriteBlankLine
        
        'Write out a blank "dialogs" block.  Dialogs that offer to remember the user's current choice will store the given choice here.
        ' We don't prepopulate it with all possible choices; instead, choices are added as the user encounters those dialogs.
        .WriteTag "Dialogs", vbNullString, True
        .CloseTag "Dialogs"
        .WriteBlankLine
        
        .WriteTag "Interface", vbNullString, True
            .WriteTag "MRUCaptionLength", "0"
            .WriteTag "RecentFilesLimit", "10"
            .WriteTag "WindowCaptionLength", "0"
            .WriteTag "CanvasColor", "#a0a0a0"
        .CloseTag "Interface"
        .WriteBlankLine
        
        .WriteTag "Language", vbNullString, True
            .WriteTag "CurrentLanguageFile", vbNullString
        .CloseTag "Language"
        .WriteBlankLine
        
        .WriteTag "Loading", vbNullString, True
            .WriteTag "ExifAutoRotate", "True"
            .WriteTag "MetadataEstimateJPEG", "True"
            .WriteTag "MetadataExtractBinary", "False"
            .WriteTag "MetadataExtractUnknown", "False"
            .WriteTag "MetadataHideDuplicates", "True"
            .WriteTag "ToneMappingPrompt", "True"
        .CloseTag "Loading"
        .WriteBlankLine
        
        .WriteTag "Paths", vbNullString, True
            .WriteTag "8bf", m_8bfPath
            .WriteTag "Macro", m_MacroPath
            .WriteTag "OpenImage", OS.SpecialFolder(CSIDL_MYPICTURES)
            .WriteTag "Palettes", m_DataPath & "Palettes\"
            .WriteTag "SaveImage", OS.SpecialFolder(CSIDL_MYPICTURES)
            .WriteTag "Selections", m_SelectionPath
            .WriteTag "TempFiles", OS.SystemTempPath()
        .CloseTag "Paths"
        .WriteBlankLine
        
        .WriteTag "Performance", vbNullString, True
            .WriteTag "InterfaceDecorationPerformance", "1"
            .WriteTag "ThumbnailPerformance", "1"
            .WriteTag "ViewportRenderPerformance", "1"
            .WriteTag "UndoCompression", "1"
        .CloseTag "Performance"
        .WriteBlankLine
        
        .WriteTag "Plugins", vbNullString, True
            .WriteTag "LastPluginPreferencesPage", "0"
        .CloseTag "Plugins"
        .WriteBlankLine
        
        .WriteTag "Saving", vbNullString, True
            .WriteTag "ConfirmClosingUnsaved", "True"
            .WriteTag "HasSavedAFile", "False"
            .WriteTag "MetadataListPD", "True"
            .WriteTag "OverwriteOrCopy", "0"
            .WriteTag "save-as-autoincrement", "False"
            .WriteTag "SuggestedFormat", "0"
            .WriteTag "UseLastFolder", "False"
        .CloseTag "Saving"
        .WriteBlankLine
        
        .WriteTag "Themes", vbNullString, True
            .WriteTag "CurrentTheme", "Dark"
            .WriteTag "CurrentAccent", "Blue"
            .WriteTag "HasSeenThemeDialog", "False"
            .WriteTag "MonochromeIcons", "False"
        .CloseTag "Themes"
        .WriteBlankLine
        
        'Toolbox settings are automatically filled-in by the Toolboxes module
        .WriteTag "Toolbox", vbNullString, True
        .CloseTag "Toolbox"
        .WriteBlankLine
        
        .WriteTag "Tools", vbNullString, True
            .WriteTag "ClearSelectionAfterCrop", "True"
            .WriteTag "HighResMouseInput", "True"
        .CloseTag "Tools"
        .WriteBlankLine
        
        .WriteTag "Transparency", vbNullString, True
            .WriteTag "AlphaCheckMode", "0"
            .WriteTag "AlphaCheckOne", Trim$(Str$(RGB(255, 255, 255)))
            .WriteTag "AlphaCheckTwo", Trim$(Str$(RGB(204, 204, 204)))
            .WriteTag "AlphaCheckSize", "1"
        .CloseTag "Transparency"
        .WriteBlankLine
        
        .WriteTag "Updates", vbNullString, True
            .WriteTag "LastUpdateCheck", vbNullString
            .WriteTag "UpdateFrequency", PDUF_WEEKLY
            
            'The current update track is set according to the hard-coded build ID of this .exe instance.
            Select Case PD_BUILD_QUALITY
            
                'Technically, I would like to default to nightly updates for alpha versions.  However, I sometimes refer
                ' casual users to the nightly builds to address specific bugs they've experienced.  They likely don't
                ' want to be bothered by myriad updates, so I've changed the default to beta builds only.  Advanced users
                ' can always opt for a faster update frequency.
                Case PD_ALPHA
                    .WriteTag "UpdateTrack", ut_Developer
                    
                Case PD_BETA
                    .WriteTag "UpdateTrack", ut_Beta
                    
                Case PD_PRODUCTION
                    .WriteTag "UpdateTrack", ut_Stable
            
            End Select
            
            .WriteTag "UpdateNotifications", True
            
        .CloseTag "Updates"
        .WriteBlankLine
        
    End With
    
    'With all tags successfully written, forcibly write the result out to file
    ' (so we have a "clean" file copy that mirrors our current settings, just like a normal session).
    ForceWriteToFile
    
End Sub

'Get a Boolean-type value from the preferences file.  (A default value must be supplied; this is used if no such value exists.)
Public Function GetPref_Boolean(ByRef preferenceSection As String, ByRef preferenceName As String, ByVal prefDefaultValue As Boolean) As Boolean

    'Request the value (as a string)
    Dim tmpString As String
    tmpString = GetPreference(preferenceSection, preferenceName)
    
    'If the requested value DOES NOT exist, return the default value as supplied by the user
    If (LenB(tmpString) = 0) Then
        
        'To prevent future blank results, write out a default value
        'Debug.Print "Requested preference " & preferenceSection & ":" & preferenceName & " was not found.  Writing out a default value of " & Trim$(Str$(prefDefaultValue))
        UserPrefs.SetPref_Boolean preferenceSection, preferenceName, prefDefaultValue
        GetPref_Boolean = prefDefaultValue
            
    'If the requested value DOES exist, convert it to boolean type and return it
    Else
        
        If (tmpString = "False") Or (tmpString = "0") Then
            GetPref_Boolean = False
        Else
            GetPref_Boolean = True
        End If
    
    End If

End Function

'Write a Boolean-type value to the preferences file.
Public Sub SetPref_Boolean(ByRef preferenceSection As String, ByRef preferenceName As String, ByVal boolVal As Boolean)
    If boolVal Then
        UserPrefs.WritePreference preferenceSection, preferenceName, "True"
    Else
        UserPrefs.WritePreference preferenceSection, preferenceName, "False"
    End If
End Sub

'Get a Long-type value from the preference file.  (A default value must be supplied; this is used if no such value exists.)
Public Function GetPref_Long(ByRef preferenceSection As String, ByRef preferenceName As String, ByVal prefDefaultValue As Long) As Long

    'Get the value (as a string) from the INI file
    Dim tmpString As String
    tmpString = GetPreference(preferenceSection, preferenceName)
    
    'If the requested value DOES NOT exist, return the default value as supplied by the user
    If (LenB(tmpString) = 0) Then
    
        'To prevent future blank results, write out a default value
        'Debug.Print "Requested preference " & preferenceSection & ":" & preferenceName & " was not found.  Writing out a default value of " & Trim$(Str$(prefDefaultValue ))
        UserPrefs.SetPref_Long preferenceSection, preferenceName, prefDefaultValue
        GetPref_Long = prefDefaultValue
    
    'If the requested value DOES exist, convert it to Long type and return it
    Else
        GetPref_Long = CLng(tmpString)
    End If

End Function

'Set a Long-type value to the preferences file.
Public Sub SetPref_Long(ByRef preferenceSection As String, ByRef preferenceName As String, ByVal longVal As Long)
    UserPrefs.WritePreference preferenceSection, preferenceName, Trim$(Str$(longVal))
End Sub

'Get a Float-type value from the preference file.  (A default value must be supplied; this is used if no such value exists.)
Public Function GetPref_Float(ByRef preferenceSection As String, ByRef preferenceName As String, ByVal prefDefaultValue As Double) As Double

    'Get the value (as a string) from the INI file
    Dim tmpString As String
    tmpString = GetPreference(preferenceSection, preferenceName)
    
    'If the requested value DOES NOT exist, return the default value as supplied by the user
    If (LenB(tmpString) = 0) Then
    
        'To prevent future blank results, write out a default value
        UserPrefs.SetPref_Float preferenceSection, preferenceName, prefDefaultValue
        GetPref_Float = prefDefaultValue
    
    'If the requested value DOES exist, convert it to Long type and return it
    Else
        GetPref_Float = CDblCustom(tmpString)
    End If

End Function

'Set a Float-type value to the preferences file.
Public Sub SetPref_Float(ByRef preferenceSection As String, ByRef preferenceName As String, ByVal floatVal As Double)
    UserPrefs.WritePreference preferenceSection, preferenceName, Trim$(Str$(floatVal))
End Sub

'Get a String-type value from the preferences file.  (A default value must be supplied; this is used if no such value exists.)
Public Function GetPref_String(ByRef preferenceSection As String, ByRef preferenceName As String, Optional ByVal prefDefaultValue As String = vbNullString, Optional ByVal writeIfMissing As Boolean = True) As String

    'Get the requested value from the preferences file
    Dim tmpString As String
    tmpString = GetPreference(preferenceSection, preferenceName)
    
    'If the requested value DOES NOT exist, return the default value as supplied by the user
    If (LenB(tmpString) = 0) Then
        
        'To prevent future blank results, write out a default value
        'Debug.Print "Requested preference " & preferenceSection & ":" & preferenceName & " was not found.  Writing out a default value of " & prefDefaultValue
        If writeIfMissing Then UserPrefs.SetPref_String preferenceSection, preferenceName, prefDefaultValue
        GetPref_String = prefDefaultValue
    
    'If the requested value DOES exist, convert it to Long type and return it
    Else
        GetPref_String = tmpString
    End If

End Function

'Set a String-type value to the INI file.
Public Sub SetPref_String(ByRef preferenceSection As String, ByRef preferenceName As String, ByRef stringVal As String)
    UserPrefs.WritePreference preferenceSection, preferenceName, stringVal
End Sub

'Sometimes we want to know if a value exists at all.  This function handles that.
Public Function DoesValueExist(ByRef preferenceSection As String, ByRef preferenceName As String) As Boolean
    Dim tmpString As String
    tmpString = GetPreference(preferenceSection, preferenceName)
    DoesValueExist = (LenB(tmpString) <> 0)
End Function

'Read a value from the preferences file and return it (as a string)
Private Function GetPreference(ByRef strSectionHeader As String, ByRef strVariableName As String) As String
    
    'Failsafe only
    If (m_XMLEngine Is Nothing) Then Exit Function
    
    'I find it helpful to give preference strings names with spaces, to improve readability.  However, XML doesn't allow tags to have
    ' spaces in the name.  So remove any spaces before interacting with the XML file.
    Const SPACE_CHAR As String = " "
    If InStr(1, strSectionHeader, SPACE_CHAR, vbBinaryCompare) Then strSectionHeader = Replace$(strSectionHeader, SPACE_CHAR, vbNullString, , , vbBinaryCompare)
    If InStr(1, strVariableName, SPACE_CHAR, vbBinaryCompare) Then strVariableName = Replace$(strVariableName, SPACE_CHAR, vbNullString, , , vbBinaryCompare)
    
    'Read the associated preference
    GetPreference = m_XMLEngine.GetUniqueTag_String(strVariableName, , , strSectionHeader)
    
End Function

'Write a string value to the preferences file
Public Function WritePreference(ByVal strSectionHeader As String, ByVal strVariableName As String, ByVal strValue As String) As Boolean
    
    'Failsafe only
    If (m_XMLEngine Is Nothing) Then Exit Function
    
    'I find it helpful to give preference strings names with spaces, to improve readability.  However, XML doesn't allow tags to have
    ' spaces in the name.  So remove any spaces before interacting with the XML file.
    Const SPACE_CHAR As String = " "
    strSectionHeader = Replace$(strSectionHeader, SPACE_CHAR, vbNullString)
    strVariableName = Replace$(strVariableName, SPACE_CHAR, vbNullString)
    
    'Check for a few necessary tags, just to make sure this is actually a PhotoDemon preferences file
    If m_XMLEngine.IsPDDataType("User Preferences") And m_XMLEngine.ValidateLoadedXMLData("Paths") Then
    
        'Update the requested tag, and if it does not exist, write it out as a new tag at the end of the specified section
        WritePreference = m_XMLEngine.UpdateTag(strVariableName, strValue, strSectionHeader)
        
        'Tag updates will fail if the requested preferences section doesn't exist
        ' (which may happen after the user upgrades from an old PhotoDemon version,
        ' but retains their existing preferences file).  To prevent the problem from recurring,
        ' add this section to the current preferences file.
        If (Not WritePreference) Then
            WritePreference = m_XMLEngine.WriteNewSection(strSectionHeader)
            If WritePreference Then WritePreference = m_XMLEngine.UpdateTag(strVariableName, strValue, strSectionHeader)
        End If
        
    End If
    
End Function

'Return the XML parameter list for a given dialog ID (constructed by the last-used settings class).
' Returns: TRUE if a preset exists for that ID; FALSE otherwise.
Public Function GetDialogPresets(ByRef dialogID As String, ByRef dstXMLString As String) As Boolean

    If m_XMLPresets.DoesTagExist(dialogID) Then
        dstXMLString = m_XMLPresets.GetUniqueTag_String(dialogID, vbNullString)
        GetDialogPresets = True
    Else
        dstXMLString = vbNullString
        GetDialogPresets = False
    End If

End Function

'Set an XML parameter list for a given dialog ID (constructed by the last-used settings class).
Public Function SetDialogPresets(ByRef dialogID As String, ByRef srcXMLString As String) As Boolean
    If Not m_XMLPresets.UpdateTag(dialogID, srcXMLString) Then
        PDDebug.LogAction "UserPrefs.SetDialogPresets() failed to update ID " & dialogID
    End If
End Function

Public Sub StartPrefEngine()
    
    'Initialize two preference engines: one for saved presets (shared across certain windows and tools), and another for
    ' the core PD user preferences file.
    Set m_XMLPresets = New pdXML
    Set m_XMLEngine = New pdXML
    m_XMLEngine.SetTextCompareMode vbBinaryCompare
    
    'Note that XML data is *not actually loaded* until the InitializePaths function is called.  (That function determines
    ' where PD's user settings file is actually stored, as it can be in several places depending on folder rights of
    ' whereever the user unzipped us.)
    
End Sub

Public Sub StopPrefEngine()
    UserPrefs.ForceWriteToFile
    Set m_XMLEngine = Nothing
    Set m_XMLPresets = Nothing
End Sub

Public Function IsReady() As Boolean
    IsReady = Not (m_XMLPresets Is Nothing)
End Function

'In rare cases, we may want to forcibly copy all current user preferences out to file
' (e.g. after the Tools > Options dialog is closed via OK button).  This function forces an
' immediate dump to disk, but note that it will only work if...
' 1) the preferences engine has been successfully initialized, and...
' 2) the central preset file path has already been validated
Public Sub ForceWriteToFile(Optional ByVal alsoWritePresets As Boolean = True)
    If ((Not m_XMLEngine Is Nothing) And (LenB(m_PreferencesPath) <> 0)) Then m_XMLEngine.WriteXMLToFile m_PreferencesPath
    If alsoWritePresets Then
        If ((Not m_XMLPresets Is Nothing) And (LenB(m_CentralPresetFile) <> 0)) Then m_XMLPresets.WriteXMLToFile m_CentralPresetFile
    End If
End Sub
