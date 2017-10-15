VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdPreferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Preferences Interface
'Copyright 2012-2017 by Tanner Helland
'Created: 03/November/12
'Last updated: 12/October/17
'Last update: add detection for "non-portable" mode, which automatically activates when the user extracts
'             PD to an access-restricted folder (like Program Files).
'
'This is the modern incarnation of PD's old "INI file" module.  It is responsible for managing all interaction
' with persistent user settings.
'
'By default, user settings are stored in an XML file in the \Data\ subfolder.  This class will generate a default
' settings file on first run.
'
'Because the settings XML file may receive new settings with each new version, all setting interaction functions
' require the caller to specify a default value, which will be used if the requested setting does not exist in the
' XML file.  Also note that if code attempts to write a setting value, but that setting value does not exist,
' it will automatically be appended as a "new" setting at the end of its respective section.
'
'Finally, let me state again that you should *never* interact with the XML settings file directly.  Always pass
' setting read/writes through this class.  I cannot guarantee that the XML format or style will remain consistent
' between versions, but as long as you stick to the wrapping functions in this class, settings will be handled correctly.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'To make PhotoDemon compatible with the PortableApps spec (http://portableapps.com/), several sub-folders are necessary.  These include:
'  /App/ subfolder, which contains information ESSENTIAL and UNIVERSAL for each PhotoDemon install (e.g. plugin DLLs, master language files)
'  /Data/ subfolder, which contains information that is OPTIONAL and UNIQUE for each PhotoDemon install (e.g. user prefs, saved macros)
Private m_ProgramPath As String
Private m_AppPath As String
Private m_DataPath As String

'Within the /App and /Data folders are additional subfolders, whose purposes should be obvious from their titles

'/App subfolders come first.  These folders should already exist in the downloaded PD .zip, and we will create them
' if they do not exist.
Private m_ThemePath As String
Private m_LanguagePath As String

'/Data subfolders come next.
Private m_MacroPath As String
Private m_PreferencesPath As String
Private m_TempPath As String
Private m_IconPath As String

Private m_UserLanguagePath As String
Private m_SelectionPath As String
Private m_PresetPath As String        'This folder is a bit different; it is used to store last-used and user-created presets for each tool dialog
Private m_DebugPath As String         'If the user is running a nightly or beta buid, a Debug folder will be created.  Debug and performance dumps
                                    ' are automatically placed here.
Private m_UserThemePath As String     '6.6 nightly builds added prelimianary theme support.  These are currently handled in-memory only, but in
                                    ' the future, themes may be extracted into this (or a matching /Data/) folder.
Private m_UpdatesPath As String       '6.6 greatly improved update support.  Update check and temp files are now stored in a dedicated folder.

'XML engine for reading/writing preference values from file
Private m_XMLEngine As pdXML

'To improve performance when saving many preferences at once (as we do when closing the Preferences dialog), this class can be placed in
' "batch preference mode".  While active, a single XML class instance will be persistently used to handle all preference updates, and
' when the mode is released, all updates will be dumped to file.
Private m_BatchModeActive As Boolean

'Some preferences are used in performance-sensitive areas.  These preferences are cached internally to improve responsiveness.
' Outside callers can retrieve them via their dedicated functions.
Private m_ThumbnailPerformance As PD_PerformanceSetting, m_ThumbnailInterpolation As GP_InterpolationMode

'Prior to v7.0, each dialog stored its preset data to a unique XML file.  This causes a lot of HDD thrashing as each
' main window panel retrieves its preset data separately.  To improve performance, we now use a single master preset
' file, and individual windows rely on this module to manage the file for them.
Private m_XMLPresets As pdXML, m_MasterPresetFile As String

'PD runs in portable mode by default, with all data folders assumed present in the same folder
' as PD itself.  If for some reason this is *not* the case, this variable will be flagged.
Private m_NonPortableModeActive As Boolean

'Helper functions for performance-sensitive preferences.
Friend Function GetThumbnailInterpolationPref() As GP_InterpolationMode
    GetThumbnailInterpolationPref = m_ThumbnailInterpolation
End Function

Friend Function GetThumbnailPerformancePref() As PD_PerformanceSetting
    GetThumbnailPerformancePref = m_ThumbnailPerformance
End Function

Friend Sub SetThumbnailPerformancePref(ByVal newSetting As PD_PerformanceSetting)
    m_ThumbnailPerformance = newSetting
    If (newSetting = PD_PERF_BESTQUALITY) Then
        m_ThumbnailInterpolation = GP_IM_HighQualityBicubic
    ElseIf (newSetting = PD_PERF_BALANCED) Then
        m_ThumbnailInterpolation = GP_IM_Bilinear
    ElseIf (newSetting = PD_PERF_FASTEST) Then
        m_ThumbnailInterpolation = GP_IM_NearestNeighbor
    End If
End Sub

'Non-portable mode means PD has been extracted to an access-restricted folder.  The program (should) still run normally,
' with silent redirection to the local appdata folder, but in-place automatic upgrades will be disabled (as we don't
' have write access to our own folder, alas).
Friend Function IsNonPortableModeActive() As Boolean
    IsNonPortableModeActive = m_NonPortableModeActive
End Function

'Initialize batch preference mode.  This is used by the preferences dialog to allow us to write a whole swath of preferences in one fell swoop,
' without writing the updates out to file after each updated.
' NOTE!  You must remember to deactivate this mode when finished; otherwise, preferences may not get written out to file for that session.
Friend Sub StartBatchPreferenceMode()

    m_BatchModeActive = True

    'Initialize an XML object
    Set m_XMLEngine = New pdXML
    m_XMLEngine.LoadXMLFile m_PreferencesPath

End Sub

'When you have finished updating multiple preferences, use this function to terminate batch mode.  The updated XML file will be
' immediately written to file.
Friend Sub EndBatchPreferenceMode()
    m_BatchModeActive = False
    m_XMLEngine.WriteXMLToFile m_PreferencesPath
End Sub

'Get the current Theme path.  Note that there are /App (program default) and /Data (userland) variants of this folder.
Friend Function GetThemePath(Optional ByVal getUserThemePathInstead As Boolean = False) As String
    If getUserThemePathInstead Then GetThemePath = m_UserThemePath Else GetThemePath = m_ThemePath
End Function

'Get the current Debug path
Friend Function GetDebugPath() As String
    GetDebugPath = m_DebugPath
End Function

'Get the current Preset path
Friend Function GetPresetPath() As String
    GetPresetPath = m_PresetPath
End Function

'Get/set the current Selection directory
Friend Function GetSelectionPath() As String
    GetSelectionPath = m_SelectionPath
End Function

Friend Sub SetSelectionPath(ByVal newSelectionPath As String)
    m_SelectionPath = Files.PathAddBackslash(Files.FileGetPath(newSelectionPath))
    SetPref_String "Paths", "Selections", m_SelectionPath
End Sub

'Return the current Language directory
Friend Function GetLanguagePath(Optional ByVal getUserLanguagePathInstead As Boolean = False) As String
    If getUserLanguagePathInstead Then GetLanguagePath = m_UserLanguagePath Else GetLanguagePath = m_LanguagePath
End Function

'Return the current temporary directory, as specified by the user's preferences.  (Note that this may not be the
' current Windows system temp path, if the user has opted to change it.)
Friend Function GetTempPath() As String
    GetTempPath = m_TempPath
End Function

'Set the current temp directory
Friend Sub SetTempPath(ByVal newTempPath As String)
    
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
Friend Function GetProgramPath() As String
    GetProgramPath = m_ProgramPath
End Function

'Return the current app data directory
Friend Function GetAppPath() As String
    GetAppPath = m_AppPath
End Function

'Return the current user data directory
Friend Function GetDataPath() As String
    GetDataPath = m_DataPath
End Function

'Return the current macro directory
Friend Function GetMacroPath() As String
    GetMacroPath = m_MacroPath
End Function

'Set the current macro directory
Friend Sub SetMacroPath(ByVal newMacroPath As String)
    m_MacroPath = Files.PathAddBackslash(Files.FileGetPath(newMacroPath))
    SetPref_String "Paths", "Macro", m_MacroPath
End Sub

'Return the current MRU icon directory
Friend Function GetIconPath() As String
    GetIconPath = m_IconPath
End Function

'Return the current update-specific temp path
Friend Function GetUpdatePath() As String
    GetUpdatePath = m_UpdatesPath
End Function

'Initialize key program directories.  If this function fails, PD will fail to load.
Friend Function InitializePaths() As Boolean
    
    InitializePaths = True
    
    'First things first: figure out where this .exe was launched from
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    m_ProgramPath = cFile.AppPathW
    
    'If this is the first time PhotoDemon is run, we need to create a series of data folders.
    ' Because PD is a portable application, we default to creating those folders in our own
    ' application folder.  Unfortunately, some users do dumb things like put PD inside protected
    ' system folders, which causes this step to fail.  We try to handle this situation gracefully,
    ' by redirecting those folders to the active user folder.
    m_NonPortableModeActive = False
    
    'Anyway, before doing anything else, let's make sure we actually have write access to our own
    ' damn folder; if we don't, we can fall activate non-portable mode.
    Dim localAppDataPath As String
    localAppDataPath = OS.SpecialFolder(CSIDL_LOCAL_APPDATA)
    
    Dim baseFolder As String
    
    Dim tmpFileWrite As String, tmpHandle As Long
    tmpFileWrite = m_ProgramPath & "tmp.tmp"
    m_NonPortableModeActive = (Not cFile.FileCreateHandle(tmpFileWrite, tmpHandle, True, True, OptimizeNone))
    cFile.FileCloseHandle tmpHandle
    cFile.FileDelete tmpFileWrite
        
    If m_NonPortableModeActive Then
        
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  Portable mode has been deactivated due to folder rights.  Attempting to salvage session..."
        #End If
        
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
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "Non-portable mode activated successfully.  Continuing program initialization..."
        #End If
    
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
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "PD data folder points at " & m_DataPath
    #End If
    
    'Within the \Data subfolder, check for additional user folders - saved macros, filters, selections, etc...
    m_MacroPath = m_DataPath & "Macros\"
    If (Not Files.PathExists(m_MacroPath)) Then Files.PathCreate m_MacroPath
            
    m_SelectionPath = m_DataPath & "Selections\"
    If (Not Files.PathExists(m_SelectionPath)) Then Files.PathCreate m_SelectionPath
    
    m_UserLanguagePath = m_DataPath & "Languages\"
    If (Not Files.PathExists(m_UserLanguagePath)) Then Files.PathCreate m_UserLanguagePath
    
    m_PresetPath = m_DataPath & "Presets\"
    If (Not Files.PathExists(m_PresetPath)) Then Files.PathCreate m_PresetPath
    
    m_IconPath = m_DataPath & "Icons\"
    If (Not Files.PathExists(m_IconPath)) Then Files.PathCreate m_IconPath
        
    m_DebugPath = m_DataPath & "Debug\"
    If (Not Files.PathExists(m_DebugPath)) Then Files.PathCreate m_DebugPath
    
    m_UserThemePath = m_DataPath & "Themes\"
    If (Not Files.PathExists(m_UserThemePath)) Then Files.PathCreate m_UserThemePath
    
    m_UpdatesPath = m_DataPath & "Updates\"
    If (Not Files.PathExists(m_UpdatesPath)) Then Files.PathCreate m_UpdatesPath
    
    'The user preferences file is also located in the \Data folder
    m_PreferencesPath = m_DataPath & PROGRAMNAME & "_settings.xml"
    
    'Last-used dialog settings are also located in the \Presets subfolder; retrieve the master file now, if it exists.
    m_MasterPresetFile = m_PresetPath & "MainPanels.xml"
    If (m_XMLPresets Is Nothing) Then Set m_XMLPresets = New pdXML
    
    If Files.FileExists(m_MasterPresetFile) Then
        If m_XMLPresets.LoadXMLFile(m_MasterPresetFile) Then
            If (Not m_XMLPresets.IsPDDataType("Presets")) Then m_XMLPresets.PrepareNewXML "Presets"
        End If
    Else
        m_XMLPresets.PrepareNewXML ("Presets")
    End If
        
End Function

'Load all user settings from file
Friend Sub LoadUserSettings(Optional ByVal endBatchModeWhenFinished As Boolean = True)
    
    'If no preferences file exists, build a default one
    If (Not Files.FileExists(m_PreferencesPath)) Then CreateNewPreferencesFile
    
    'Start batch processing mode for preferences, which is much faster (as we don't have to go out to the hard drive
    ' for each preference access).
    Me.StartBatchPreferenceMode
        
    'Pull the temp file path from the preferences file and make sure it exists. (If it doesn't, transparently set it to
    ' the system temp path.)
    m_TempPath = GetPref_String("Paths", "TempFiles", "")
    If (Not Files.PathExists(m_TempPath)) Then
        m_TempPath = OS.SystemTempPath()
        SetPref_String "Paths", "TempFiles", m_TempPath
    End If
        
    'Pull all other stored paths
    m_MacroPath = GetPref_String("Paths", "Macro", m_MacroPath)
    m_SelectionPath = GetPref_String("Paths", "Selections", m_SelectionPath)
        
    'Check if the user wants us to prompt them about closing unsaved images
    g_ConfirmClosingUnsaved = GetPref_Boolean("Saving", "ConfirmClosingUnsaved", True)
    
    'Check if the user wants large images automatically zoomed when loaded
    g_AutozoomLargeImages = GetPref_Long("Loading", "InitialImageZoom", 0)
    
    'Grab the last-used common dialog filters
    g_LastOpenFilter = GetPref_Long("Core", "LastOpenFilter", 1)
    g_LastSaveFilter = GetPref_Long("Core", "LastSaveFilter", 3)
    
    'For performance reasons, cache any performance-related settings.  (This is much faster than reading the preferences from file
    ' every time they're needed.)
    g_InterfacePerformance = g_UserPreferences.GetPref_Long("Performance", "InterfaceDecorationPerformance", PD_PERF_BALANCED)
    Me.SetThumbnailPerformancePref g_UserPreferences.GetPref_Long("Performance", "ThumbnailPerformance", PD_PERF_BALANCED)
    g_ViewportPerformance = g_UserPreferences.GetPref_Long("Performance", "ViewportRenderPerformance", PD_PERF_BALANCED)
    g_UndoCompressionLevel = g_UserPreferences.GetPref_Long("Performance", "UndoCompression", 1)
    
    Tools.SetToolSetting_HighResMouse g_UserPreferences.GetPref_Boolean("Tools", "HighResMouseInput", True)
    
    'Initialize a few other settings to default values, despite them not coming from the preferences file
    
    'Terminate batch preference mode before exiting
    If endBatchModeWhenFinished Then Me.EndBatchPreferenceMode
                
End Sub

'Reset the preferences file to its default state.  (Basically, delete any existing file, then create a new one from scratch.)
Friend Sub ResetPreferences()
    Files.FileDeleteIfExists m_PreferencesPath
    CreateNewPreferencesFile
    LoadUserSettings
End Sub

'Create a new preferences XML file from scratch.  When new preferences are added to the preferences dialog, they should also be
' added to this function, to ensure that the most intelligent preference is selected by default.
Private Sub CreateNewPreferencesFile()

    'This function is used to determine whether PhotoDemon is being run for the first time.  Why do it here?
    ' 1) When first downloaded, PhotoDemon doesn't come with a prefs file.  Thus this routine MUST be called.
    ' 2) When preferences are reset, this file is deleted.  That is an appropriate time to mark the program as "first run",
    '     so any first-run type dialogs are reset as well.
    ' 3) If the user moves PhotoDemon but leaves behind the old prefs file.  There's no easy way to check this, but
    '     treating the program as if its being run for the first time is as good a plan as any.
    g_IsFirstRun = True
    
    'As a failsafe against data corruption, if this is determined to be a first run, we also delete some settings-related
    ' files in the Presets folder (if they exist).
    If g_IsFirstRun Then Files.FileDeleteIfExists m_PresetPath & "Program_WindowLocations.xml"
    
    'Create a pdXML class, which will help us assemble the preferences file (in XML format, obviously)
    Set m_XMLEngine = New pdXML
    m_XMLEngine.PrepareNewXML "User Preferences"
    m_XMLEngine.WriteBlankLine
    
    'Write out a comment marking the date and build of this preferences code; this can be helpful when debugging
    m_XMLEngine.WriteComment "This preferences file was created on " & Format$(Now, "dd-mmm-yyyy") & " by version " & App.Major & "." & App.Minor & "." & App.Revision & " of the software."
    m_XMLEngine.WriteBlankLine
    
    'Write out the "advanced" block of preferences:
    m_XMLEngine.WriteTag "Advanced", "", True
        'As of 6.4, there are no longer any Advanced preferences that need to be stored here.  There may be in the future,
        ' so I've left the creation code for this section of the preferences file.
    m_XMLEngine.CloseTag "Advanced"
    m_XMLEngine.WriteBlankLine
    
    m_XMLEngine.WriteTag "BatchProcess", "", True
        m_XMLEngine.WriteTag "DriveBox", ""
        m_XMLEngine.WriteTag "InputFolder", OS.SpecialFolder(CSIDL_MYPICTURES)
        m_XMLEngine.WriteTag "ListFolder", OS.SpecialFolder(CSIDL_PERSONAL)
        m_XMLEngine.WriteTag "OutputFolder", OS.SpecialFolder(CSIDL_MYPICTURES)
    m_XMLEngine.CloseTag "BatchProcess"
    m_XMLEngine.WriteBlankLine
    
    'Write out the "color management" block of preferences:
    m_XMLEngine.WriteTag "ColorManagement", "", True
        m_XMLEngine.WriteTag "DisplayCMMode", Trim$(Str(DCM_NoManagement))
        m_XMLEngine.WriteTag "DisplayRenderingIntent", Trim$(Str(INTENT_PERCEPTUAL))
    m_XMLEngine.CloseTag "ColorManagement"
    m_XMLEngine.WriteBlankLine
    
    'Write out the "core" block of preferences.  These are preferences that PD uses internally.  These are never directly
    ' exposed to the user (e.g. the user cannot toggle these from the Preferences dialog).
    m_XMLEngine.WriteTag "Core", "", True
        m_XMLEngine.WriteTag "DisplayIDEWarning", "True"
        m_XMLEngine.WriteTag "HasGitHubAccount", ""
        m_XMLEngine.WriteTag "LastOpenFilter", "1"        'Default to "All Compatible Graphics" filter for loading
        m_XMLEngine.WriteTag "LastPreferencesPage", "0"
        m_XMLEngine.WriteTag "LastSaveFilter", "-1"       'Mark the last-used save filter as "unknown"
        m_XMLEngine.WriteTag "LastWindowState", "0"
        m_XMLEngine.WriteTag "LastWindowLeft", "1"
        m_XMLEngine.WriteTag "LastWindowTop", "1"
        m_XMLEngine.WriteTag "LastWindowWidth", "1"
        m_XMLEngine.WriteTag "LastWindowHeight", "1"
    m_XMLEngine.CloseTag "Core"
    m_XMLEngine.WriteBlankLine
    
    'Write out a blank "dialogs" block.  Dialogs that offer to remember the user's current choice will store the given choice here.
    ' We don't prepopulate it with all possible choices; instead, choices are added as the user encounters those dialogs.
    m_XMLEngine.WriteTag "Dialogs", "", True
    m_XMLEngine.CloseTag "Dialogs"
    m_XMLEngine.WriteBlankLine
    
    m_XMLEngine.WriteTag "Interface", "", True
        m_XMLEngine.WriteTag "MRUCaptionLength", "0"
        m_XMLEngine.WriteTag "RecentFilesLimit", "10"
        m_XMLEngine.WriteTag "WindowCaptionLength", "0"
    m_XMLEngine.CloseTag "Interface"
    m_XMLEngine.WriteBlankLine
    
    m_XMLEngine.WriteTag "Language", "", True
        m_XMLEngine.WriteTag "CurrentLanguageFile", ""
    m_XMLEngine.CloseTag "Language"
    m_XMLEngine.WriteBlankLine
    
    m_XMLEngine.WriteTag "Loading", "", True
        m_XMLEngine.WriteTag "ExifAutoRotate", "True"
        m_XMLEngine.WriteTag "InitialImageZoom", "0"
        m_XMLEngine.WriteTag "MetadataEstimateJPEG", "True"
        m_XMLEngine.WriteTag "MetadataExtractBinary", "False"
        m_XMLEngine.WriteTag "MetadataExtractUnknown", "False"
        m_XMLEngine.WriteTag "MetadataHideDuplicates", "True"
        m_XMLEngine.WriteTag "ToneMappingPrompt", "True"
    m_XMLEngine.CloseTag "Loading"
    m_XMLEngine.WriteBlankLine
    
    m_XMLEngine.WriteTag "Paths", "", True
        m_XMLEngine.WriteTag "TempFiles", OS.SystemTempPath()
        m_XMLEngine.WriteTag "OpenImage", OS.SpecialFolder(CSIDL_MYPICTURES)
        m_XMLEngine.WriteTag "SaveImage", OS.SpecialFolder(CSIDL_MYPICTURES)
        m_XMLEngine.WriteTag "Macro", m_MacroPath
        m_XMLEngine.WriteTag "Selections", m_SelectionPath
    m_XMLEngine.CloseTag "Paths"
    m_XMLEngine.WriteBlankLine
    
    m_XMLEngine.WriteTag "Performance", "", True
        m_XMLEngine.WriteTag "InterfaceDecorationPerformance", "1"
        m_XMLEngine.WriteTag "ThumbnailPerformance", "1"
        m_XMLEngine.WriteTag "ViewportRenderPerformance", "1"
        m_XMLEngine.WriteTag "UndoCompression", "1"
    m_XMLEngine.CloseTag "Performance"
    m_XMLEngine.WriteBlankLine
    
    m_XMLEngine.WriteTag "Plugins", "", True
        m_XMLEngine.WriteTag "ForceExifToolDisable", "False"
        m_XMLEngine.WriteTag "ForceEZTwainDisable", "False"
        m_XMLEngine.WriteTag "ForceFreeImageDisable", "False"
        m_XMLEngine.WriteTag "ForceLittleCMSDisable", "False"
        m_XMLEngine.WriteTag "ForceOptiPNGDisable", "False"
        m_XMLEngine.WriteTag "ForcePngQuantDisable", "False"
        m_XMLEngine.WriteTag "ForceZLibDisable", "False"
        m_XMLEngine.WriteTag "LastPluginPreferencesPage", "0"
    m_XMLEngine.CloseTag "Plugins"
    m_XMLEngine.WriteBlankLine
    
    m_XMLEngine.WriteTag "Saving", "", True
        m_XMLEngine.WriteTag "ConfirmClosingUnsaved", "True"
        m_XMLEngine.WriteTag "OverwriteOrCopy", "0"
        m_XMLEngine.WriteTag "SuggestedFormat", "0"
        m_XMLEngine.WriteTag "MetadataListPD", "True"
    m_XMLEngine.CloseTag "Saving"
    m_XMLEngine.WriteBlankLine
    
    m_XMLEngine.WriteTag "Themes", "", True
        m_XMLEngine.WriteTag "CurrentTheme", "Dark"
        m_XMLEngine.WriteTag "CurrentAccent", "Blue"
        m_XMLEngine.WriteTag "HasSeenThemeDialog", "False"
        m_XMLEngine.WriteTag "MonochromeIcons", "False"
    m_XMLEngine.CloseTag "Themes"
    m_XMLEngine.WriteBlankLine
    
    'Toolbox settings are automatically filled-in by the Toolboxes module
    m_XMLEngine.WriteTag "Toolbox", "", True
    m_XMLEngine.CloseTag "Toolbox"
    m_XMLEngine.WriteBlankLine
    
    m_XMLEngine.WriteTag "Tools", "", True
        m_XMLEngine.WriteTag "ClearSelectionAfterCrop", "True"
        m_XMLEngine.WriteTag "SelectionRenderMode", "0"
        m_XMLEngine.WriteTag "SelectionHighlightColor", "#FF3A48"
        m_XMLEngine.WriteTag "HighResMouseInput", "True"
    m_XMLEngine.CloseTag "Tools"
    m_XMLEngine.WriteBlankLine
    
    m_XMLEngine.WriteTag "Transparency", "", True
        m_XMLEngine.WriteTag "AlphaCheckMode", "0"
        m_XMLEngine.WriteTag "AlphaCheckOne", Trim$(Str(RGB(255, 255, 255)))
        m_XMLEngine.WriteTag "AlphaCheckTwo", Trim$(Str(RGB(204, 204, 204)))
        m_XMLEngine.WriteTag "AlphaCheckSize", "1"
    m_XMLEngine.CloseTag "Transparency"
    m_XMLEngine.WriteBlankLine
    
    m_XMLEngine.WriteTag "Updates", "", True
        m_XMLEngine.WriteTag "CheckForUpdates", "True"
        m_XMLEngine.WriteTag "LastUpdateCheck", ""
        m_XMLEngine.WriteTag "PromptForPluginDownload", "True"
        m_XMLEngine.WriteTag "UpdateFrequency", PDUF_WEEKLY
        
        'The current update track is set according to the hard-coded build ID of this .exe instance.
        Select Case PD_BUILD_QUALITY
        
            'Technically, I would like to default to nightly updates for alpha versions.  However, I sometimes refer
            ' casual users to the nightly builds to address specific bugs they've experienced.  They likely don't
            ' want to be bothered by myriad updates, so I've changed the default to beta builds only.  Advanced users
            ' can always opt for a faster update frequency.
            Case PD_PRE_ALPHA, PD_ALPHA
                m_XMLEngine.WriteTag "UpdateTrack", PDUT_BETA
                
            Case PD_BETA
                m_XMLEngine.WriteTag "UpdateTrack", PDUT_BETA
                
            Case PD_PRODUCTION
                m_XMLEngine.WriteTag "UpdateTrack", PDUT_STABLE
        
        End Select
        
        m_XMLEngine.WriteTag "UpdateLanguagesIndependently", True
        m_XMLEngine.WriteTag "UpdatePluginsIndependently", True
        m_XMLEngine.WriteTag "UpdateNotifications", True
        
    m_XMLEngine.CloseTag "Updates"
    m_XMLEngine.WriteBlankLine
    
    'With all tags successfully written, we can now close the XML data and write it out to file.
    m_XMLEngine.WriteXMLToFile m_PreferencesPath
    
End Sub

'Get a Boolean-type value from the preferences file.  (A default value must be supplied; this is used if no such value exists.)
Friend Function GetPref_Boolean(ByRef preferenceSection As String, ByRef preferenceName As String, ByVal prefDefaultValue As Boolean) As Boolean

    'Request the value (as a string)
    Dim tmpString As String
    tmpString = GetPreference(preferenceSection, preferenceName)
    
    'If the requested value DOES NOT exist, return the default value as supplied by the user
    If (Len(tmpString) = 0) Then
        
        'To prevent future blank results, write out a default value
        'Debug.Print "Requested preference " & preferenceSection & ":" & preferenceName & " was not found.  Writing out a default value of " & Trim$(Str(prefDefaultValue))
        SetPref_Boolean preferenceSection, preferenceName, prefDefaultValue
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
Friend Sub SetPref_Boolean(ByRef preferenceSection As String, ByRef preferenceName As String, ByVal boolVal As Boolean)

    If boolVal Then
        WritePreference preferenceSection, preferenceName, "True"
    Else
        WritePreference preferenceSection, preferenceName, "False"
    End If

End Sub

'Get a Long-type value from the preference file.  (A default value must be supplied; this is used if no such value exists.)
Friend Function GetPref_Long(ByRef preferenceSection As String, ByRef preferenceName As String, ByVal prefDefaultValue As Long) As Long

    'Get the value (as a string) from the INI file
    Dim tmpString As String
    tmpString = GetPreference(preferenceSection, preferenceName)
    
    'If the requested value DOES NOT exist, return the default value as supplied by the user
    If (Len(tmpString) = 0) Then
    
        'To prevent future blank results, write out a default value
        'Debug.Print "Requested preference " & preferenceSection & ":" & preferenceName & " was not found.  Writing out a default value of " & Trim$(Str(prefDefaultValue ))
        SetPref_Long preferenceSection, preferenceName, prefDefaultValue
        GetPref_Long = prefDefaultValue
    
    'If the requested value DOES exist, convert it to Long type and return it
    Else
        GetPref_Long = CLng(tmpString)
    End If

End Function

'Set a Long-type value to the preferences file.
Friend Sub SetPref_Long(ByRef preferenceSection As String, ByRef preferenceName As String, ByVal longVal As Long)
    WritePreference preferenceSection, preferenceName, Trim$(Str(longVal))
End Sub

'Get a Float-type value from the preference file.  (A default value must be supplied; this is used if no such value exists.)
Friend Function GetPref_Float(ByRef preferenceSection As String, ByRef preferenceName As String, ByVal prefDefaultValue As Double) As Double

    'Get the value (as a string) from the INI file
    Dim tmpString As String
    tmpString = GetPreference(preferenceSection, preferenceName)
    
    'If the requested value DOES NOT exist, return the default value as supplied by the user
    If (Len(tmpString) = 0) Then
    
        'To prevent future blank results, write out a default value
        SetPref_Float preferenceSection, preferenceName, prefDefaultValue
        GetPref_Float = prefDefaultValue
    
    'If the requested value DOES exist, convert it to Long type and return it
    Else
        GetPref_Float = CDblCustom(tmpString)
    End If

End Function

'Set a Float-type value to the preferences file.
Friend Sub SetPref_Float(ByRef preferenceSection As String, ByRef preferenceName As String, ByVal floatVal As Double)
    WritePreference preferenceSection, preferenceName, Trim$(Str(floatVal))
End Sub

'Get a String-type value from the preferences file.  (A default value must be supplied; this is used if no such value exists.)
Friend Function GetPref_String(ByRef preferenceSection As String, ByRef preferenceName As String, Optional ByVal prefDefaultValue As String = vbNullString) As String

    'Get the requested value from the preferences file
    Dim tmpString As String
    tmpString = GetPreference(preferenceSection, preferenceName)
    
    'If the requested value DOES NOT exist, return the default value as supplied by the user
    If (Len(tmpString) = 0) Then
        
        'To prevent future blank results, write out a default value
        'Debug.Print "Requested preference " & preferenceSection & ":" & preferenceName & " was not found.  Writing out a default value of " & prefDefaultValue
        SetPref_String preferenceSection, preferenceName, prefDefaultValue
        GetPref_String = prefDefaultValue
    
    'If the requested value DOES exist, convert it to Long type and return it
    Else
        GetPref_String = tmpString
    End If

End Function

'Set a String-type value to the INI file.
Friend Sub SetPref_String(ByRef preferenceSection As String, ByRef preferenceName As String, ByRef stringVal As String)
    WritePreference preferenceSection, preferenceName, stringVal
End Sub

'Sometimes we want to know if a value exists at all.  This function handles that.
Friend Function DoesValueExist(ByRef preferenceSection As String, ByRef preferenceName As String) As Boolean
    Dim tmpString As String
    tmpString = GetPreference(preferenceSection, preferenceName)
    DoesValueExist = CBool(Len(tmpString) <> 0)
End Function

'Read a value from the preferences file and return it (as a string)
Private Function GetPreference(ByRef strSectionHeader As String, ByRef strVariableName As String) As String
    
    'I find it helpful to give preference strings names with spaces, to improve readability.  However, XML doesn't allow tags to have
    ' spaces in the name.  So remove any spaces before interacting with the XML file.
    If InStr(1, strSectionHeader, " ", vbBinaryCompare) Then strSectionHeader = Replace$(strSectionHeader, " ", "", , , vbBinaryCompare)
    If InStr(1, strVariableName, " ", vbBinaryCompare) Then strVariableName = Replace$(strVariableName, " ", "", , , vbBinaryCompare)
    
    'Create an XML object and load the XML settings file
    Dim prefFileOK As Boolean
    If m_BatchModeActive Then
        prefFileOK = True
    Else
        Set m_XMLEngine = New pdXML
        m_XMLEngine.LoadXMLFile m_PreferencesPath
        prefFileOK = m_XMLEngine.IsPDDataType("User Preferences") And m_XMLEngine.ValidateLoadedXMLData("Paths")
    End If
    
    'Check for a few necessary tags, just to make sure this is actually a PhotoDemon preferences file
    If prefFileOK Then
        GetPreference = m_XMLEngine.GetUniqueTag_String(strVariableName, , , strSectionHeader)
    Else
        GetPreference = vbNullString
    End If
    
End Function

'Write a string value to the preferences file
Friend Function WritePreference(ByVal strSectionHeader As String, ByVal strVariableName As String, ByVal strValue As String) As Boolean

    'I find it helpful to give preference strings names with spaces, to improve readability.  However, XML doesn't allow tags to have
    ' spaces in the name.  So remove any spaces before interacting with the XML file.
    strSectionHeader = Replace$(strSectionHeader, " ", "")
    strVariableName = Replace$(strVariableName, " ", "")
    
    'Create an XML object and load the XML settings file
    If (Not m_BatchModeActive) Then
        Set m_XMLEngine = New pdXML
        m_XMLEngine.LoadXMLFile m_PreferencesPath
    End If
    
    'Check for a few necessary tags, just to make sure this is actually a PhotoDemon preferences file
    If m_XMLEngine.IsPDDataType("User Preferences") And m_XMLEngine.ValidateLoadedXMLData("Paths") Then
    
        'Update the requested tag, and if it does not exist, write it out as a new tag at the end of the specified section
        If m_XMLEngine.UpdateTag(strVariableName, strValue, strSectionHeader) Then
            WritePreference = True
            If (Not m_BatchModeActive) Then m_XMLEngine.WriteXMLToFile m_PreferencesPath
        Else
            
            'Update tag will fail if the requested preferences section doesn't exist (which may happen after the user upgrades from
            ' an old PhotoDemon version, but keeps their old preferences file).  To prevent the problem from recurring, add this
            ' section to the preferences file.
            If m_XMLEngine.WriteNewSection(strSectionHeader) Then
            
                'Try the write again
                If m_XMLEngine.UpdateTag(strVariableName, strValue, strSectionHeader) Then
                    WritePreference = True
                    If (Not m_BatchModeActive) Then m_XMLEngine.WriteXMLToFile m_PreferencesPath
                Else
                    WritePreference = False
                End If
            
            Else
                WritePreference = False
            End If
            
        End If
        
    End If
    
End Function

'Return the XML parameter list for a given dialog ID (constructed by the last-used settings class).
' Returns: TRUE if a preset exists for that ID; FALSE otherwise.
Friend Function GetDialogPresets(ByRef dialogID As String, ByRef dstXMLString As String) As Boolean

    If m_XMLPresets.DoesTagExist(dialogID) Then
        dstXMLString = m_XMLPresets.GetUniqueTag_String(dialogID, vbNullString)
        GetDialogPresets = True
    Else
        dstXMLString = vbNullString
        GetDialogPresets = False
    End If

End Function

'Set an XML parameter list for a given dialog ID (constructed by the last-used settings class).
Friend Function SetDialogPresets(ByRef dialogID As String, ByRef srcXMLString As String) As Boolean
    m_XMLPresets.UpdateTag dialogID, srcXMLString
End Function

Private Sub Class_Initialize()
    Set m_XMLPresets = New pdXML
End Sub

Private Sub Class_Terminate()
    If ((Not m_XMLPresets Is Nothing) And (Len(m_MasterPresetFile) <> 0)) Then m_XMLPresets.WriteXMLToFile m_MasterPresetFile
End Sub
