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
'Copyright 2012-2016 by Tanner Helland
'Created: 03/November/12
'Last updated: 06/February/14
'Last update: added new failsafes against the possibility of settings corruption.  If the program feels compelled to rewrite the
'              settings file from scratch (for example, after a series of unsafe shutdowns), certain key settings-related
'              preset files will also be rewritten.  This should protect against the random bug where the image tabstrip
'              renders as if top-aligned but is clearly positioned as left-aligned.
'
'This class is the newest incarnation of the former "INI Handler" module.  It is responsible for managing all interaction with user
' settings, which are currently stored in an XML file.  This class is also capable of generating a default preferences file
' (necessary for first run), determining system folders (also necessary for first run), and saving/loading any user preferences
' to/from a persistent file on the user's HDD.
'
'Because the preferences XML file may receive new preferences with any new version of the software, this class must also handle the
' possibility of missing values.  As such, any requests for preferences must specify a default value, which will be used if the
' requested preference does not exist in the file.  Also, if a preference update is requested and that preference cannot be found,
' it will automatically be appended to the end of its section.
'
'When PhotoDemon is run for the first time, a default preferences file is created.  To increase usability, this file is automatically
' populated with intelligent preferences - such as setting the user's initial Open/Save folders to the system "Pictures" folder,
' and automatically setting the "Temp" folder to the OS Temp folder.
'
'Finally, please do not interact with the XML file directly.  Always pass preference writes and reads through this class.  I cannot
' guarantee that the XML format or style will remain consistent between versions, but as long as you stick to the wrapping functions
' in this class, preferences will be read/written correctly.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'API for determining certain system folders
Private Declare Function GetSystemTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'API for determining specific system-related folders
Private Declare Function SHGetFolderPath Lib "shfolder" Alias "SHGetFolderPathA" (ByVal hWndOwner As Long, ByVal nFolder As CSIDLs, ByVal hToken As Long, ByVal dwReserved As Long, ByVal lpszPath As String) As Long

'Windows Folder Constants
Public Enum CSIDLs
    CSIDL_MY_DOCUMENTS = &H5
    CSIDL_MYPICTURES = &H27
End Enum

Private Const SHGFP_TYPE_CURRENT = &H0 'current value for user, verify it exists
Private Const MAX_LENGTH = 260
Private Const S_OK = 0

'***

'To make PhotoDemon compatible with the PortableApps spec (http://portableapps.com/), several sub-folders are necessary.  These include:
'  /App/ subfolder, which contains information ESSENTIAL and UNIVERSAL for each PhotoDemon install (e.g. plugin DLLs, master language files)
'  /Data/ subfolder, which contains information that is OPTIONAL and UNIQUE for each PhotoDemon install (e.g. user prefs, saved macros)
Private ProgramPath As String
Private AppPath As String
Private DataPath As String
Private ThemePath As String

'Within the /App and /Data folders are additional subfolders, whose purposes should be obvious from their titles
Private MacroPath As String
Private PreferencesPath As String
Private TempPath As String
Private IconPath As String
Private LanguagePath As String
Private UserLanguagePath As String
Private SelectionPath As String
Private PresetPath As String        'This folder is a bit different; it is used to store last-used and user-created presets for each tool dialog
Private DebugPath As String         'If the user is running a nightly or beta buid, a Debug folder will be created.  Debug and performance dumps
                                    ' are automatically placed here.
Private UserThemePath As String     '6.6 nightly builds added prelimianary theme support.  These are currently handled in-memory only, but in
                                    ' the future, themes may be extracted into this (or a matching /Data/) folder.
Private UpdatesPath As String       '6.6 greatly improved update support.  Update check and temp files are now stored in a dedicated folder.

'XML engine for reading/writing preference values from file
Private xmlEngine As pdXML

'To improve performance when saving many preferences at once (as we do when closing the Preferences dialog), this class can be placed in
' "batch preference mode".  While active, a single XML class instance will be persistently used to handle all preference updates, and
' when the mode is released, all updates will be dumped to file.
Private m_BatchModeActive As Boolean

'Initialize batch preference mode.  This is used by the preferences dialog to allow us to write a whole swath of preferences in one fell swoop,
' without writing the updates out to file after each updated.
' NOTE!  You must remember to deactivate this mode when finished; otherwise, preferences may not get written out to file for that session.
Public Sub StartBatchPreferenceMode()

    m_BatchModeActive = True

    'Initialize an XML object
    Set xmlEngine = New pdXML
    xmlEngine.LoadXMLFile PreferencesPath

End Sub

'When you have finished updating multiple preferences, use this function to terminate batch mode.  The updated XML file will be
' immediately written to file.
Public Sub EndBatchPreferenceMode()
    m_BatchModeActive = False
    xmlEngine.WriteXMLToFile PreferencesPath
End Sub

'Get the current Theme path.  Note that there are /App (program default) and /Data (userland) variants of this folder.
Public Function GetThemePath(Optional ByVal getUserThemePathInstead As Boolean = False) As String
    If getUserThemePathInstead Then
        GetThemePath = UserThemePath
    Else
        GetThemePath = ThemePath
    End If
End Function

'Get the current Debug path
Public Function GetDebugPath() As String
    GetDebugPath = DebugPath
End Function

'Get the current Preset path
Public Function GetPresetPath() As String
    GetPresetPath = PresetPath
End Function

'Get/set the current Selection directory
Public Function GetSelectionPath() As String
    GetSelectionPath = SelectionPath
End Function

Public Sub SetSelectionPath(ByVal newSelectionPath As String)
    StripDirectory newSelectionPath
    newSelectionPath = FixPath(newSelectionPath)
    SelectionPath = newSelectionPath
    SetPref_String "Paths", "Selections", SelectionPath
End Sub

'Return the current Language directory
Public Function GetLanguagePath(Optional ByVal getUserLanguagePathInstead As Boolean = False) As String
    If getUserLanguagePathInstead Then
        GetLanguagePath = UserLanguagePath
    Else
        GetLanguagePath = LanguagePath
    End If
End Function

'Return the current temporary directory, as specified by the user's preferences.  (Note that this is a totally different function from
' this class's GetSystemTemporaryPath function, which returns the current Windows system temp path.)
Public Function GetTempPath() As String
    GetTempPath = TempPath
End Function

'Set the current temp directory
Public Sub SetTempPath(ByVal newTempPath As String)
    
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    'If the folder exists and is writable as-is, great: save it and exit
    If cFile.FolderExist(newTempPath, True) Or cFile.FolderExist(cFile.EnforcePathSlash(newTempPath), True) Then
        TempPath = cFile.EnforcePathSlash(newTempPath)
        
    'If it doesn't exist, make sure the user didn't do something weird, like supply a file instead of a folder
    Else
    
        newTempPath = cFile.EnforcePathSlash(cFile.GetPathOnly(newTempPath))
        
        'Test the path again
        If cFile.FolderExist(newTempPath, True) Or cFile.FolderExist(cFile.EnforcePathSlash(newTempPath), True) Then
            TempPath = cFile.EnforcePathSlash(newTempPath)
            
        'If it still fails, revert to the default system temp path
        Else
            TempPath = GetSystemTemporaryPath()
        End If
    
    End If
    
    'Write the final path out to file
    SetPref_String "Paths", "TempFiles", TempPath
    
End Sub

'Return the current program directory
Public Function GetProgramPath() As String
    GetProgramPath = ProgramPath
End Function

'Return the current app data directory
Public Function GetAppPath() As String
    GetAppPath = AppPath
End Function

'Return the current user data directory
Public Function GetDataPath() As String
    GetDataPath = DataPath
End Function

'Return the current macro directory
Public Function GetMacroPath() As String
    GetMacroPath = MacroPath
End Function

'Set the current macro directory
Public Sub SetMacroPath(ByVal newMacroPath As String)
    StripDirectory newMacroPath
    newMacroPath = FixPath(newMacroPath)
    MacroPath = newMacroPath
    SetPref_String "Paths", "Macro", MacroPath
End Sub

'Return the current MRU icon directory
Public Function GetIconPath() As String
    GetIconPath = IconPath
End Function

'Return the current update-specific temp path
Public Function GetUpdatePath() As String
    GetUpdatePath = UpdatesPath
End Function

'Initialize key program directories
Public Sub InitializePaths()
    
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    'First things first: figure out where this .exe was launched from
    ProgramPath = cFile.AppPathW
    
    'Create an "App" path based off the ProgramPath - this is where essential application files (like plugins) are stored.
    ' Also, make sure the directory exists; if it doesn't, create it.
    AppPath = ProgramPath & "App\"
    If Not cFile.FolderExist(AppPath) Then cFile.CreateFolder AppPath
    
    AppPath = AppPath & "PhotoDemon\"
    If Not cFile.FolderExist(AppPath) Then cFile.CreateFolder AppPath
    
    'Within the App\PhotoDemon\ folder, create a folder for any available OFFICIAL translations.  (User translations go in the Data folder.)
    LanguagePath = AppPath & "Languages\"
    If Not cFile.FolderExist(LanguagePath) Then cFile.CreateFolder LanguagePath
    
    'Within the App\PhotoDemon\ folder, create a folder for any available OFFICIAL themes.  (User themes go in the Data folder.)
    ThemePath = AppPath & "Themes\"
    If Not cFile.FolderExist(ThemePath) Then cFile.CreateFolder ThemePath
    
    'Create a "Data" path based off the ProgramPath - this is where the preferences file and any other user-specific files
    ' (saved filters, macros) will be stored.  Also, make sure the directory exists; if it doesn't, create it.
    DataPath = ProgramPath & "Data\"
    If Not cFile.FolderExist(DataPath) Then cFile.CreateFolder DataPath
    
    'Within the \Data subfolder, check for additional user folders - saved macros, filters, selections, etc...
    MacroPath = DataPath & "Macros\"
    If Not cFile.FolderExist(MacroPath) Then cFile.CreateFolder MacroPath
            
    SelectionPath = DataPath & "Selections\"
    If Not cFile.FolderExist(SelectionPath) Then cFile.CreateFolder SelectionPath
    
    UserLanguagePath = DataPath & "Languages\"
    If Not cFile.FolderExist(UserLanguagePath) Then cFile.CreateFolder UserLanguagePath
    
    PresetPath = DataPath & "Presets\"
    If Not cFile.FolderExist(PresetPath) Then cFile.CreateFolder PresetPath
    
    IconPath = DataPath & "Icons\"
    If Not cFile.FolderExist(IconPath) Then cFile.CreateFolder IconPath
        
    DebugPath = DataPath & "Debug\"
    If Not cFile.FolderExist(DebugPath) Then cFile.CreateFolder DebugPath
    
    UserThemePath = DataPath & "Themes\"
    If Not cFile.FolderExist(UserThemePath) Then cFile.CreateFolder UserThemePath
    
    UpdatesPath = DataPath & "Updates\"
    If Not cFile.FolderExist(UpdatesPath) Then cFile.CreateFolder UpdatesPath
    
    'The user preferences file is also located in the \Data subdirectory
    PreferencesPath = DataPath & PROGRAMNAME & "_settings.xml"
    
    'Historically, language files were stored in a different location.  We need to move any valid language files from
    ' the old language directory to the new language directory.
    ' (NOTE: this is no longer necessary, but in case it becomes necessary to move language files in the future, I have
    '         left this code here as a helpful reminder.)
    
    'On Error Resume Next
    'Dim chkFile As String
    '
    'chkFile = Dir(DataPath & "Languages\" & "*.xml", vbNormal)
    '
    'If chkFile <> "" Then
    '
    '    LoadMessage "Moving language files to updated folder location..."
    '
    '    Do While chkFile <> ""
    '        If FileExist(DataPath & "Languages\" & chkFile) Then
    '            FileCopy DataPath & "Languages\" & chkFile, LanguagePath & chkFile
    '            Kill DataPath & "Languages\" & chkFile
    '        End If
    '        chkFile = Dir
    '    Loop
    '
    '    RmDir DataPath & "Languages\"
    '
    'End If
        
End Sub

'Load all user settings from file
Public Sub LoadUserSettings()
    
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    'If no preferences file exists, build a default one
    If Not cFile.FileExist(PreferencesPath) Then CreateNewPreferencesFile
    
    'Start batch processing mode for preferences, which is much faster (as we don't have to go out to the hard drive
    ' for each preference access).
    StartBatchPreferenceMode
        
    'Pull the temp file path from the preferences file and make sure it exists. (If it doesn't, transparently set it to
    ' the system temp path.)
    TempPath = GetPref_String("Paths", "TempFiles", "")
    If Not cFile.FolderExist(TempPath) Then
        TempPath = GetSystemTemporaryPath
        SetPref_String "Paths", "TempFiles", TempPath
    End If
        
    'Pull all other stored paths
    MacroPath = GetPref_String("Paths", "Macro", MacroPath)
    SelectionPath = GetPref_String("Paths", "Selections", SelectionPath)
        
    'Check if the user wants us to prompt them about closing unsaved images
    g_ConfirmClosingUnsaved = GetPref_Boolean("Saving", "ConfirmClosingUnsaved", True)
    
    'Check if the user wants large images automatically zoomed when loaded
    g_AutozoomLargeImages = GetPref_Long("Loading", "InitialImageZoom", 0)
    
    'Grab the last-used common dialog filters
    g_LastOpenFilter = GetPref_Long("Core", "LastOpenFilter", 1)
    g_LastSaveFilter = GetPref_Long("Core", "LastSaveFilter", 3)
    
    'For performance reasons, cache the system color management setting
    g_UseSystemColorProfile = g_UserPreferences.GetPref_Boolean("Transparency", "Use System Color Profile", True)
    
    'For performance reasons, cache any performance-related settings.  (This is much faster than reading the preferences from file
    ' every time they're needed.)
    g_ColorPerformance = g_UserPreferences.GetPref_Long("Performance", "ColorPerformance", PD_PERF_BALANCED)
    g_InterfacePerformance = g_UserPreferences.GetPref_Long("Performance", "InterfaceDecorationPerformance", PD_PERF_BALANCED)
    g_ThumbnailPerformance = g_UserPreferences.GetPref_Long("Performance", "ThumbnailPerformance", PD_PERF_BALANCED)
    g_ViewportPerformance = g_UserPreferences.GetPref_Long("Performance", "ViewportRenderPerformance", PD_PERF_BALANCED)
    g_UndoCompressionLevel = g_UserPreferences.GetPref_Long("Performance", "UndoCompression", 0)
    g_HighResolutionInput = g_UserPreferences.GetPref_Boolean("Interface", "HighResolutionInput", True)
    
    'Initialize a few other settings to default values, despite them not coming from the preferences file
    
    'Terminate batch preference mode before exiting
    EndBatchPreferenceMode
                
End Sub

'Reset the preferences file to its default state.  (Basically, delete any existing file, then create a new one from scratch.)
Public Sub ResetPreferences()

    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    If cFile.FileExist(PreferencesPath) Then cFile.KillFile PreferencesPath
    
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
    
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    'As a failsafe against data corruption, if this is determined to be a first run, we also delete some settings-related
    ' files in the Presets folder (if they exist).
    If g_IsFirstRun Then
        If cFile.FileExist(PresetPath & "Program_WindowLocations.xml") Then cFile.KillFile PresetPath & "Program_WindowLocations.xml"
    End If
    
    'Create a pdXML class, which will help us assemble the preferences file (in XML format, obviously)
    Set xmlEngine = New pdXML
    xmlEngine.PrepareNewXML "User Preferences"
    xmlEngine.WriteBlankLine
    
    'Write out a comment marking the date and build of this preferences code; this can be helpful when debugging
    xmlEngine.WriteComment "This preferences file was created on " & Format$(Now, "dd-mmm-yyyy") & " by version " & App.Major & "." & App.Minor & "." & App.Revision & " of the software."
    xmlEngine.WriteBlankLine
    
    'Write out the "advanced" block of preferences:
    xmlEngine.WriteTag "Advanced", "", True
        'As of 6.4, there are no longer any Advanced preferences that need to be stored here.  There may be in the future,
        ' so I've left the creation code for this section of the preferences file.
    xmlEngine.CloseTag "Advanced"
    xmlEngine.WriteBlankLine
    
    xmlEngine.WriteTag "BatchProcess", "", True
        xmlEngine.WriteTag "DriveBox", ""
        xmlEngine.WriteTag "InputFolder", GetWindowsFolder(CSIDL_MYPICTURES)
        xmlEngine.WriteTag "ListFolder", GetWindowsFolder(CSIDL_MY_DOCUMENTS)
        xmlEngine.WriteTag "OutputFolder", GetWindowsFolder(CSIDL_MYPICTURES)
    xmlEngine.CloseTag "BatchProcess"
    xmlEngine.WriteBlankLine
    
    'Write out the "core" block of preferences.  These are preferences that PD uses internally.  These are never directly
    ' exposed to the user (e.g. the user cannot toggle these from the Preferences dialog).
    xmlEngine.WriteTag "Core", "", True
        xmlEngine.WriteTag "DisplayIDEWarning", "True"
        xmlEngine.WriteTag "HasGitHubAccount", ""
        xmlEngine.WriteTag "LastOpenFilter", "1"        'Default to "All Compatible Graphics" filter for loading
        xmlEngine.WriteTag "LastPreferencesPage", "0"
        xmlEngine.WriteTag "LastSaveFilter", "-1"       'Mark the last-used save filter as "unknown"
        xmlEngine.WriteTag "LastWindowState", "0"
        xmlEngine.WriteTag "LastWindowLeft", "1"
        xmlEngine.WriteTag "LastWindowTop", "1"
        xmlEngine.WriteTag "LastWindowWidth", "1"
        xmlEngine.WriteTag "LastWindowHeight", "1"
    xmlEngine.CloseTag "Core"
    xmlEngine.WriteBlankLine
    
    'Write out a blank "dialogs" block.  Dialogs that offer to remember the user's current choice will store the given choice here.
    ' We don't prepopulate it with all possible choices; instead, choices are added as the user encounters those dialogs.
    xmlEngine.WriteTag "Dialogs", "", True
    xmlEngine.CloseTag "Dialogs"
    xmlEngine.WriteBlankLine
    
    xmlEngine.WriteTag "Interface", "", True
        xmlEngine.WriteTag "HighResolutionInput", "False"
        xmlEngine.WriteTag "MRUCaptionLength", "0"
        xmlEngine.WriteTag "RecentFilesLimit", "10"
        xmlEngine.WriteTag "WindowCaptionLength", "0"
    xmlEngine.CloseTag "Interface"
    xmlEngine.WriteBlankLine
    
    xmlEngine.WriteTag "Language", "", True
        xmlEngine.WriteTag "CurrentLanguageFile", ""
    xmlEngine.CloseTag "Language"
    xmlEngine.WriteBlankLine
    
    xmlEngine.WriteTag "Loading", "", True
        xmlEngine.WriteTag "ExifAutoRotate", "True"
        xmlEngine.WriteTag "InitialImageZoom", "0"
        xmlEngine.WriteTag "MetadataEstimateJPEG", "True"
        xmlEngine.WriteTag "MetadataExtractBinary", "False"
        xmlEngine.WriteTag "MetadataExtractUnknown", "False"
        xmlEngine.WriteTag "MetadataHideDuplicates", "True"
        xmlEngine.WriteTag "ToneMappingPrompt", "True"
    xmlEngine.CloseTag "Loading"
    xmlEngine.WriteBlankLine
    
    xmlEngine.WriteTag "Paths", "", True
        xmlEngine.WriteTag "TempFiles", GetSystemTemporaryPath
        xmlEngine.WriteTag "OpenImage", GetWindowsFolder(CSIDL_MYPICTURES)
        xmlEngine.WriteTag "SaveImage", GetWindowsFolder(CSIDL_MYPICTURES)
        xmlEngine.WriteTag "Macro", MacroPath
        xmlEngine.WriteTag "Selections", SelectionPath
    xmlEngine.CloseTag "Paths"
    xmlEngine.WriteBlankLine
    
    xmlEngine.WriteTag "Performance", "", True
        xmlEngine.WriteTag "ColorPerformance", "1"
        xmlEngine.WriteTag "InterfaceDecorationPerformance", "1"
        xmlEngine.WriteTag "ThumbnailPerformance", "1"
        xmlEngine.WriteTag "ViewportRenderPerformance", "1"
        xmlEngine.WriteTag "UndoCompression", "0"
    xmlEngine.CloseTag "Performance"
    xmlEngine.WriteBlankLine
    
    xmlEngine.WriteTag "Plugins", "", True
        xmlEngine.WriteTag "ForceExifToolDisable", "False"
        xmlEngine.WriteTag "ForceEZTwainDisable", "False"
        xmlEngine.WriteTag "ForceFreeImageDisable", "False"
        xmlEngine.WriteTag "ForceLittleCMSDisable", "False"
        xmlEngine.WriteTag "ForceOptiPNGDisable", "False"
        xmlEngine.WriteTag "ForcePngQuantDisable", "False"
        xmlEngine.WriteTag "ForceZLibDisable", "False"
        xmlEngine.WriteTag "LastPluginPreferencesPage", "0"
    xmlEngine.CloseTag "Plugins"
    xmlEngine.WriteBlankLine
    
    xmlEngine.WriteTag "Saving", "", True
        xmlEngine.WriteTag "ConfirmClosingUnsaved", "True"
        xmlEngine.WriteTag "OverwriteOrCopy", "0"
        xmlEngine.WriteTag "SuggestedFormat", "0"
        xmlEngine.WriteTag "PreserveMetadata", "True"
        xmlEngine.WriteTag "AnonymizeMetadata", "False"
        xmlEngine.WriteTag "MetadataListPD", "True"
    xmlEngine.CloseTag "Saving"
    xmlEngine.WriteBlankLine
    
    xmlEngine.WriteTag "Themes", "", True
        xmlEngine.WriteTag "CurrentTheme", "Light"
        xmlEngine.WriteTag "CurrentAccent", "Blue"
    xmlEngine.CloseTag "Themes"
    xmlEngine.WriteBlankLine
    
    'Toolbox settings are automatically filled-in by the Toolboxes module
    xmlEngine.WriteTag "Toolbox", "", True
    xmlEngine.CloseTag "Toolbox"
    xmlEngine.WriteBlankLine
    
    xmlEngine.WriteTag "Tools", "", True
        xmlEngine.WriteTag "ClearSelectionAfterCrop", "True"
    xmlEngine.CloseTag "Tools"
    xmlEngine.WriteBlankLine
    
    xmlEngine.WriteTag "Transparency", "", True
        xmlEngine.WriteTag "AlphaCheckMode", "0"
        xmlEngine.WriteTag "AlphaCheckOne", Trim$(Str(RGB(255, 255, 255)))
        xmlEngine.WriteTag "AlphaCheckTwo", Trim$(Str(RGB(204, 204, 204)))
        xmlEngine.WriteTag "AlphaCheckSize", "1"
    xmlEngine.CloseTag "Transparency"
    xmlEngine.WriteBlankLine
    
    xmlEngine.WriteTag "Updates", "", True
        xmlEngine.WriteTag "CheckForUpdates", "True"
        xmlEngine.WriteTag "LastUpdateCheck", ""
        xmlEngine.WriteTag "PromptForPluginDownload", "True"
        xmlEngine.WriteTag "UpdateFrequency", PDUF_EACH_SESSION
        
        'The current update track is set according to the hard-coded build ID of this .exe instance.
        Select Case PD_BUILD_QUALITY
        
            'Technically, I would like to default to nightly updates for alpha versions.  However, I sometimes send casual users
            ' to the nightly builds to fix a specific bug.  As they likely don't want to be bothered by so many updates, I am
            ' changing this to default to beta builds only.
            Case PD_PRE_ALPHA, PD_ALPHA
                xmlEngine.WriteTag "UpdateTrack", PDUT_BETA
                
            Case PD_BETA
                xmlEngine.WriteTag "UpdateTrack", PDUT_BETA
                
            Case PD_PRODUCTION
                xmlEngine.WriteTag "UpdateTrack", PDUT_STABLE
        
        End Select
        
        xmlEngine.WriteTag "UpdateLanguagesIndependently", True
        xmlEngine.WriteTag "UpdatePluginsIndependently", True
        xmlEngine.WriteTag "UpdateNotifications", True
        
    xmlEngine.CloseTag "Updates"
    xmlEngine.WriteBlankLine
    
    'With all tags successfully written, we can now close the XML data and write it out to file.
    xmlEngine.WriteXMLToFile PreferencesPath
    
End Sub

'Get a Boolean-type value from the preferences file.  (A default value must be supplied; this is used if no such value exists.)
Public Function GetPref_Boolean(ByRef preferenceSection As String, ByRef preferenceName As String, ByVal prefDefaultValue As Boolean) As Boolean

    'Request the value (as a string)
    Dim tmpString As String
    tmpString = GetPreference(preferenceSection, preferenceName)
    
    'If the requested value DOES NOT exist, return the default value as supplied by the user
    If Len(tmpString) = 0 Then
        
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
Public Sub SetPref_Boolean(ByRef preferenceSection As String, ByRef preferenceName As String, ByVal boolVal As Boolean)

    If boolVal Then
        WritePreference preferenceSection, preferenceName, "True"
    Else
        WritePreference preferenceSection, preferenceName, "False"
    End If

End Sub

'Get a Long-type value from the preference file.  (A default value must be supplied; this is used if no such value exists.)
Public Function GetPref_Long(ByRef preferenceSection As String, ByRef preferenceName As String, ByVal prefDefaultValue As Long) As Long

    'Get the value (as a string) from the INI file
    Dim tmpString As String
    tmpString = GetPreference(preferenceSection, preferenceName)
    
    'If the requested value DOES NOT exist, return the default value as supplied by the user
    If Len(tmpString) = 0 Then
    
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
Public Sub SetPref_Long(ByRef preferenceSection As String, ByRef preferenceName As String, ByVal longVal As Long)
    WritePreference preferenceSection, preferenceName, Trim$(Str(longVal))
End Sub

'Get a String-type value from the preferences file.  (A default value must be supplied; this is used if no such value exists.)
Public Function GetPref_String(ByRef preferenceSection As String, ByRef preferenceName As String, ByVal prefDefaultValue As String) As String

    'Get the requested value from the preferences file
    Dim tmpString As String
    tmpString = GetPreference(preferenceSection, preferenceName)
    
    'If the requested value DOES NOT exist, return the default value as supplied by the user
    If Len(tmpString) = 0 Then
        
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
Public Sub SetPref_String(ByRef preferenceSection As String, ByRef preferenceName As String, ByRef stringVal As String)
    WritePreference preferenceSection, preferenceName, stringVal
End Sub

'Sometimes we want to know if a value exists at all.  This function handles that.
Public Function DoesValueExist(ByRef preferenceSection As String, ByRef preferenceName As String) As Boolean
    Dim tmpString As String
    tmpString = GetPreference(preferenceSection, preferenceName)
    DoesValueExist = CBool(Len(tmpString) <> 0)
End Function

'Read a value from the preferences file and return it (as a string)
Private Function GetPreference(ByRef strSectionHeader As String, ByRef strVariableName As String) As String
    
    'I find it helpful to give preference strings names with spaces, to improve readability.  However, XML doesn't allow tags to have
    ' spaces in the name.  So remove any spaces before interacting with the XML file.
    strSectionHeader = Replace$(strSectionHeader, " ", "")
    strVariableName = Replace$(strVariableName, " ", "")
    
    'Create an XML object and load the XML settings file
    If (Not m_BatchModeActive) Then
        Set xmlEngine = New pdXML
        xmlEngine.LoadXMLFile PreferencesPath
    End If
    
    'Check for a few necessary tags, just to make sure this is actually a PhotoDemon preferences file
    If xmlEngine.IsPDDataType("User Preferences") And xmlEngine.ValidateLoadedXMLData("Paths") Then
        GetPreference = xmlEngine.GetUniqueTag_String(strVariableName, , , strSectionHeader)
    Else
        GetPreference = ""
    End If
    
End Function

'Write a string value to the preferences file
Public Function WritePreference(strSectionHeader As String, strVariableName As String, strValue As String) As Boolean

    'I find it helpful to give preference strings names with spaces, to improve readability.  However, XML doesn't allow tags to have
    ' spaces in the name.  So remove any spaces before interacting with the XML file.
    strSectionHeader = Replace$(strSectionHeader, " ", "")
    strVariableName = Replace$(strVariableName, " ", "")
    
    'Create an XML object and load the XML settings file
    If Not m_BatchModeActive Then
        Set xmlEngine = New pdXML
        xmlEngine.LoadXMLFile PreferencesPath
    End If
    
    'Check for a few necessary tags, just to make sure this is actually a PhotoDemon preferences file
    If xmlEngine.IsPDDataType("User Preferences") And xmlEngine.ValidateLoadedXMLData("Paths") Then
    
        'Update the requested tag, and if it does not exist, write it out as a new tag at the end of the specified section
        If xmlEngine.UpdateTag(strVariableName, strValue, strSectionHeader) Then
            WritePreference = True
            If Not m_BatchModeActive Then xmlEngine.WriteXMLToFile PreferencesPath
        Else
            
            'Update tag will fail if the requested preferences section doesn't exist (which may happen after the user upgrades from
            ' an old PhotoDemon version, but keeps their old preferences file).  To prevent the problem from recurring, add this
            ' section to the preferences file.
            If xmlEngine.WriteNewSection(strSectionHeader) Then
            
                'Try the write again
                If xmlEngine.UpdateTag(strVariableName, strValue, strSectionHeader) Then
                    WritePreference = True
                    If Not m_BatchModeActive Then xmlEngine.WriteXMLToFile PreferencesPath
                Else
                    WritePreference = False
                End If
            
            Else
                WritePreference = False
            End If
            
        End If
        
    End If
    
End Function

'Get the current Windows-specified temp directory
Private Function GetSystemTemporaryPath() As String
    
    'Create a blank string (as required by the API call)
    Dim sRet As String
    sRet = String(255, 0)
    
    'Fill that string with the temporary path
    Dim lngLen As Long
    lngLen = GetSystemTempPath(255, sRet)
    
    'If something went wrong, raise an error
    If lngLen = 0 Then Err.Raise Err.LastDllError
    
    'Return the temporary path, and ensure that it has a trailing "\"
    GetSystemTemporaryPath = FixPath(Left$(sRet, lngLen))

End Function

'Get a special folder from Windows (as specified by the CSIDL)
Private Function GetWindowsFolder(eFolder As CSIDLs) As String

    Dim iR As Integer
    Dim sPath As String
    
    sPath = String$(MAX_LENGTH, " ")
    If SHGetFolderPath(0&, eFolder, 0&, SHGFP_TYPE_CURRENT, sPath) = S_OK Then
        
        'Find the end of the string
        iR = InStr(1, sPath, vbNullChar) - 1
        
        'Return everything up to the NULL + (Tanner's fix) add a terminating slash
        GetWindowsFolder = FixPath(Left$(sPath, iR))
        
    End If
    
End Function
