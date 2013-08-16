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
'Copyright ©2012-2013 by Tanner Helland
'Created: 03/November/12
'Last updated: 14/August/13
'Last update: removed the last bit of INI code from the project.  PhotoDemon is now 100% XML instead of INI.
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
' in this class, your preferences will be read/written correctly.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'API for determining certain system folders
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'API for determining specific system-related folders
Private Declare Function SHGetFolderPath Lib "shfolder.dll" Alias "SHGetFolderPathA" (ByVal hWndOwner As Long, ByVal nFolder As CSIDLs, ByVal hToken As Long, ByVal dwReserved As Long, ByVal lpszPath As String) As Long

'Windows Folder Constants
Public Enum CSIDLs
    CSIDL_MY_DOCUMENTS = &H5
    CSIDL_MYPICTURES = &H27
End Enum

Private Const CSIDL_FLAG_CREATE As Long = 32768 '&H8000 'combine with CSIDL_ value to force create on SHGetSpecialFolderLocation()
Private Const CSIDL_FLAG_DONT_VERIFY = &H4000 'combine with CSIDL_ value to force create on SHGetSpecialFolderLocation()
Private Const CSIDL_FLAG_MASK = &HFF00 'mask for all possible flag Values
Private Const SHGFP_TYPE_CURRENT = &H0 'current value for user, verify it exists
Private Const SHGFP_TYPE_DEFAULT = &H1
Private Const MAX_LENGTH = 260
Private Const S_OK = 0
Private Const S_FALSE = 1
'***

'To make PhotoDemon compatible with the PortableApps spec (http://portableapps.com/), several sub-folders are necessary.  These include:
'  /App/ subfolder, which contains information ESSENTIAL and UNIVERSAL for each PhotoDemon install (e.g. plugin DLLs, master language files)
'  /Data/ subfolder, which contains information that is OPTIONAL and UNIQUE for each PhotoDemon install (e.g. user prefs, saved macros)
Private ProgramPath As String
Private AppPath As String
Private DataPath As String

'Within the /App and /Data folders are additional subfolders, whose purposes should be obvious from their titles
Private MacroPath As String
Private FilterPath As String
Private PreferencesPath As String
Private TempPath As String
Private IconPath As String
Private LanguagePath As String
Private UserLanguagePath As String
Private SelectionPath As String
Private PresetPath As String        'This folder is a bit different; it is used to store last-used and user-created presets for each tool dialog

'Get the current Preset path
Public Function getPresetPath() As String
    getPresetPath = PresetPath
End Function

'Get/set the current Selection directory
Public Function getSelectionPath() As String
    getSelectionPath = SelectionPath
End Function

Public Sub setSelectionPath(ByVal newSelectionPath As String)
    StripDirectory newSelectionPath
    newSelectionPath = FixPath(newSelectionPath)
    SelectionPath = newSelectionPath
    SetPref_String "Program Paths", "Selections", SelectionPath
End Sub

'Return the current Language directory
Public Function getLanguagePath(Optional ByVal getUserLanguagePathInstead As Boolean = False) As String
    If getUserLanguagePathInstead Then
        getLanguagePath = UserLanguagePath
    Else
        getLanguagePath = LanguagePath
    End If
End Function

'Return the current temporary directory
Public Function getTempPath() As String
    getTempPath = TempPath
End Function

'Set the current temp directory
Public Sub setTempPath(ByVal newTempPath As String)
    
    StripDirectory newTempPath
    newTempPath = FixPath(newTempPath)
    TempPath = newTempPath
    
    'Ensure that this directory exists AND that we have write access to it.  If we do not, substitute the default system temp folder.
    If DirectoryExist(TempPath) = False Then TempPath = GetTemporaryPath()
    
    SetPref_String "Program Paths", "TempFiles", TempPath
    
End Sub

'Return the current program directory
Public Function getProgramPath() As String
    getProgramPath = ProgramPath
End Function

'Return the current app data directory
Public Function getAppPath() As String
    getAppPath = AppPath
End Function

'Return the current user data directory
Public Function getDataPath() As String
    getDataPath = DataPath
End Function

'Return the current macro directory
Public Function getMacroPath() As String
    getMacroPath = MacroPath
End Function

'Set the current macro directory
Public Sub setMacroPath(ByVal newMacroPath As String)
    StripDirectory newMacroPath
    newMacroPath = FixPath(newMacroPath)
    MacroPath = newMacroPath
    SetPref_String "Program Paths", "Macro", MacroPath
End Sub

'Return the current filter directory
Public Function getFilterPath() As String
    getFilterPath = FilterPath
End Function

'Set the current filter directory
Public Sub setFilterPath(ByVal newFilterPath As String)
    StripDirectory newFilterPath
    newFilterPath = FixPath(newFilterPath)
    FilterPath = newFilterPath
    SetPref_String "Program Paths", "Filter", FilterPath
End Sub

'Return the current MRU icon directory
Public Function getIconPath() As String
    getIconPath = IconPath
End Function

'Initialize key program directories
Public Sub initializePaths()
    
    'First things first: figure out where this .exe was launched from
    ProgramPath = App.Path
    ProgramPath = FixPath(ProgramPath)
    
    'Create an "App" path based off the ProgramPath - this is where essential application files (like plugins) are stored.
    ' Also, make sure the directory exists; if it doesn't, create it.
    AppPath = ProgramPath & "App\"
    If Not DirectoryExist(AppPath) Then MkDir AppPath
    
    AppPath = AppPath & "PhotoDemon\"
    If Not DirectoryExist(AppPath) Then MkDir AppPath
    
    'Within the App\PhotoDemon\ folder, create a folder for any available OFFICIAL translations.  (User translations go in the Data folder.)
    LanguagePath = AppPath & "Languages\"
    If Not DirectoryExist(LanguagePath) Then MkDir LanguagePath
        
    'Create a "Data" path based off the ProgramPath - this is where the preferences file and any other user-specific files
    ' (saved filters, macros) will be stored.  Also, make sure the directory exists; if it doesn't, create it.
    DataPath = ProgramPath & "Data\"
    If Not DirectoryExist(DataPath) Then MkDir DataPath
    
    'Within the \Data subfolder, check for additional user folders - saved macros, filters, selections, etc...
    MacroPath = DataPath & "Macros\"
    If Not DirectoryExist(MacroPath) Then MkDir MacroPath
    
    FilterPath = DataPath & "Filters\"
    If Not DirectoryExist(FilterPath) Then MkDir FilterPath
        
    SelectionPath = DataPath & "Selections\"
    If Not DirectoryExist(SelectionPath) Then MkDir SelectionPath
    
    UserLanguagePath = DataPath & "Languages\"
    If Not DirectoryExist(UserLanguagePath) Then MkDir UserLanguagePath
    
    PresetPath = DataPath & "Presets\"
    If Not DirectoryExist(PresetPath) Then MkDir PresetPath
    
    IconPath = DataPath & "Icons\"
    If Not DirectoryExist(IconPath) Then MkDir IconPath
    
    'The user preferences file is also located in the \Data subdirectory
    PreferencesPath = DataPath & PROGRAMNAME & "_settings.xml"
    
    'Historically, language files were stored in a different location.  We need to move any valid language files from
    ' the old language directory to the new language directory.
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
Public Sub loadUserSettings()
        
    'If no preferences file exists, build a default one
    If Not FileExist(PreferencesPath) Then createNewPreferencesFile
            
    'Get the LogProgramMessages preference
    g_LogProgramMessages = GetPref_Boolean("General Preferences", "LogProgramMessages", False)
    
    'If the user wants us to log program messages, open the log file and write out the current date and time
    If g_LogProgramMessages Then
        
        Dim fileNum As Integer
        fileNum = FreeFile
    
        Open DataPath & PROGRAMNAME & "_DebugMessages.log" For Append As #fileNum
            Print #fileNum, vbCrLf
            Print #fileNum, vbCrLf
            Print #fileNum, "**********************************************"
            Print #fileNum, "Date: " & Date
            Print #fileNum, "Time: " & Time
        Close #fileNum
        
    End If
            
    'Pull the temp file path from the preferences file and make sure it exists. (If it doesn't, transparently set it to
    ' the system temp path.)
    TempPath = GetPref_String("Program Paths", "TempFiles", "")
    If Not DirectoryExist(TempPath) Then
        TempPath = GetTemporaryPath
        SetPref_String "Program Paths", "TempFiles", TempPath
    End If
    
    'Pull all other stored paths
    MacroPath = GetPref_String("Program Paths", "Macro", MacroPath)
    FilterPath = GetPref_String("Program Paths", "CustomFilter", FilterPath)
    SelectionPath = GetPref_String("Program Paths", "Selections", SelectionPath)
    
    'Get the canvas background color
    g_CanvasBackground = GetPref_Long("General Preferences", "CanvasBackground", RGB(201, 211, 226))
    
    'Get the canvas drop shadow preference
    g_CanvasDropShadow = GetPref_Boolean("General Preferences", "CanvasDropShadow", False)
    
    'Get the alpha transparency checkerboard color mode (light/moderate/dark/custom)
    g_AlphaCheckMode = GetPref_Long("General Preferences", "AlphaCheckMode", 1)
    
    'Get the alpha transparency checkerboard colors
    g_AlphaCheckOne = GetPref_Long("General Preferences", "AlphaCheckOne", RGB(153, 153, 153))
    g_AlphaCheckTwo = GetPref_Long("General Preferences", "AlphaCheckTwo", RGB(102, 102, 102))
    
    'Get the alpha checkerboard size
    g_AlphaCheckSize = GetPref_Long("General Preferences", "AlphaCheckSize", 1)
    
    'Check if the user wants us to prompt them about closing unsaved images
    g_ConfirmClosingUnsaved = GetPref_Boolean("General Preferences", "ConfirmClosingUnsaved", True)
        
    'Check if the user wants the interface to use modern typefaces if available
    g_UseFancyFonts = GetPref_Boolean("General Preferences", "UseFancyFonts", False)
    
    'Check if the user wants large images automatically zoomed when loaded
    g_AutosizeLargeImages = GetPref_Long("General Preferences", "AutosizeLargeImages", 0)
    
    'Grab the last-used common dialog filters
    g_LastOpenFilter = GetPref_Long("File Formats", "LastOpenFilter", 1)
    g_LastSaveFilter = GetPref_Long("File Formats", "LastSaveFilter", 3)
    
    'Initialize a few other settings to default values, despite them not coming from the preferences file
    
    'No custom filters have been created yet
    g_HasCreatedFilter = False
    
    'Mark the Macro recorder as "not recording"
    MacroStatus = MacroSTOP
    
    'No images have been loaded yet
    NumOfImagesLoaded = 0
    
    'Set the default MDI window index to 0
    CurrentImage = 0
    
    'Set the number of open image windows to 0
    NumOfWindows = 0
    
    'Set a default emboss/engrave color
    g_EmbossEngraveColor = &HFF8080
    
End Sub

'Save and load tool settings - note that this must be done after the various tool controls have been initialized
' on the main form.
Public Sub loadToolSettings()

    'Selection type
    Dim i As Long
    For i = 0 To FormMain.cmbSelRender.Count - 1
        FormMain.cmbSelRender(i).ListIndex = GetPref_Long("Tool Preferences", "LastSelectionShape", 0)
    Next i
    
End Sub

Public Sub saveToolSettings()
    
    'Last-used tool
    SetPref_Long "Tool Preferences", "LastActiveTool", g_CurrentTool
        
End Sub

'Reset the preferences file to its default state.  (Basically, delete any existing file, then create a new one from scratch.)
Public Sub resetPreferences()
    If FileExist(PreferencesPath) Then Kill PreferencesPath
    createNewPreferencesFile
    loadUserSettings
End Sub

'Create a new preferences XML file from scratch.  When new preferences are added to the preferences dialog, they should also be
' added to this function, to ensure that the most intelligent preference is selected by default.
Private Sub createNewPreferencesFile()

    'This function is used to determine whether PhotoDemon is being run for the first time.  Why do it here?
    ' 1) When first downloaded, PhotoDemon doesn't come with a prefs file.  Thus this routine MUST be called.
    ' 2) When preferences are reset, this file is deleted.  That is an appropriate time to mark the program as "first run".
    ' 3) If the user moves PhotoDemon but leaves behind the old prefs file.  There's no easy way to check this, but
    '     treating the program as if its being run for the first time is as good a plan as any.
    g_IsFirstRun = True
    
    'Create a pdXML class, which will help us assemble the preferences file (in XML format, obviously)
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    xmlEngine.prepareNewXML "PhotoDemon User Preferences"
    xmlEngine.writeBlankLine
    
    'Write out a comment marking the date and build of this preferences code; this can be helpful when debugging
    xmlEngine.writeComment "This preferences file was created on " & Format$(Now, "dd-mmm-yyyy") & " by version " & App.Major & "." & App.Minor & "." & App.Revision & " of the software."
    xmlEngine.writeBlankLine
    
    'Write out the first block of settings: program paths
    xmlEngine.writeTag "ProgramPaths", "", True
        xmlEngine.writeTag "TempFiles", GetTemporaryPath
        xmlEngine.writeTag "OpenImage", GetWindowsFolder(CSIDL_MYPICTURES)
        xmlEngine.writeTag "SaveImage", GetWindowsFolder(CSIDL_MYPICTURES)
        xmlEngine.writeTag "CustomFilter", FilterPath
        xmlEngine.writeTag "Macro", MacroPath
        xmlEngine.writeTag "Selections", SelectionPath
    xmlEngine.closeTag "ProgramPaths"
    xmlEngine.writeBlankLine

    'Write out the second block of settings: general preferences
    xmlEngine.writeTag "GeneralPreferences", "", True
        xmlEngine.writeTag "AlphaCheckMode", "0"
        xmlEngine.writeTag "AlphaCheckOne", CStr(RGB(255, 255, 255))
        xmlEngine.writeTag "AlphaCheckTwo", CStr(RGB(204, 204, 204))
        xmlEngine.writeTag "AlphaCheckSize", "1"
        xmlEngine.writeTag "AutosizeLargeImages", "0"
        xmlEngine.writeTag "BitmapRLE", "False"
        xmlEngine.writeTag "CanvasBackground", CStr(vb3DShadow) 'RGB(201, 211, 226)
        xmlEngine.writeTag "CanvasDropShadow", "True"
        xmlEngine.writeTag "CheckForUpdates", "True"
        xmlEngine.writeTag "ConfirmClosingUnsaved", "True"
        xmlEngine.writeTag "CurrentLanguageFile", ""
        xmlEngine.writeTag "DefaultSaveFormat", "0"
        xmlEngine.writeTag "DisplayIDEWarning", "True"
        xmlEngine.writeTag "DynamicTaskbarIcon", "True"
        xmlEngine.writeTag "HasGitHubAccount", ""
        xmlEngine.writeTag "HideLeftPanel", "False"
        xmlEngine.writeTag "HideRightPanel", "False"
        xmlEngine.writeTag "ImageCaptionSize", "0"
        xmlEngine.writeTag "LastUpdateCheck", ""
        xmlEngine.writeTag "LastPreferencesPage", "0"
        xmlEngine.writeTag "LastFilePreferencesPage", "1"
        xmlEngine.writeTag "LogProgramMessages", "False"
        xmlEngine.writeTag "MetadataExport", "1"
        xmlEngine.writeTag "MRUCaptionSize", "0"
        xmlEngine.writeTag "MultipageImagePrompt", "0"
        xmlEngine.writeTag "OutgoingColorDepth", "1"
        xmlEngine.writeTag "PNGBackgroundPreservation", "True"
        xmlEngine.writeTag "PNGCompression", "9"
        xmlEngine.writeTag "PNGInterlacing", "False"
        xmlEngine.writeTag "PPMExportFormat", "0"
        xmlEngine.writeTag "PromptForPluginDownload", "True"
        xmlEngine.writeTag "RememberWindowLocation", "True"
        xmlEngine.writeTag "LastWindowState", "0"
        xmlEngine.writeTag "LastWindowLeft", "1"
        xmlEngine.writeTag "LastWindowTop", "1"
        xmlEngine.writeTag "LastWindowWidth", "1"
        xmlEngine.writeTag "LastWindowHeight", "1"
        xmlEngine.writeTag "SaveBehavior", "0"
        xmlEngine.writeTag "TGARLE", "False"
        xmlEngine.writeTag "TIFFCMYK", "False"
        xmlEngine.writeTag "TIFFCompression", "0"
        'On Vista and later, Segoe UI is used as the interface font instead of Tahoma
        If g_IsVistaOrLater Then xmlEngine.writeTag "UseFancyFonts", "True" Else xmlEngine.writeTag "UseFancyFonts", "False"
        xmlEngine.writeTag "UseToneMapping", "True"
        xmlEngine.writeTag "ValidateAlphaChannels", "True"
        xmlEngine.writeTag "VerifyInitialColorDepth", "True"
    xmlEngine.closeTag "GeneralPreferences"
    xmlEngine.writeBlankLine
    
    'Write out the third block of settings: plugin preferences
    xmlEngine.writeTag "PluginPreferences", "", True
        xmlEngine.writeTag "ForceFreeImageDisable", "False"
        xmlEngine.writeTag "ForceZLibDisable", "False"
        xmlEngine.writeTag "ForceEZTwainDisable", "False"
        xmlEngine.writeTag "ForcePngnqDisable", "False"
        xmlEngine.writeTag "ForceExifToolDisable", "False"
        xmlEngine.writeTag "LastPluginPreferencesPage", "0"
        xmlEngine.writeTag "PngnqAlphaExtenuation", "False"
        xmlEngine.writeTag "PngnqYUV", "True"
        xmlEngine.writeTag "PngnqColorSample", "3"
        xmlEngine.writeTag "PngnqDithering", "5"
    xmlEngine.closeTag "PluginPreferences"
    xmlEngine.writeBlankLine
    
    'Write out batch conversion preferences
    xmlEngine.writeTag "BatchPreferences", "", True
        xmlEngine.writeTag "DriveBox", ""
        xmlEngine.writeTag "InputFolder", GetWindowsFolder(CSIDL_MYPICTURES)
        xmlEngine.writeTag "OutputFolder", GetWindowsFolder(CSIDL_MYPICTURES)
        xmlEngine.writeTag "ListFolder", GetWindowsFolder(CSIDL_MY_DOCUMENTS)
    xmlEngine.closeTag "BatchPreferences"
    xmlEngine.writeBlankLine
        
    'Write out file format preferences
    xmlEngine.writeTag "FileFormats", "", True
        xmlEngine.writeTag "LastOpenFilter", "1"   'Default to "All Compatible Graphics" filter for loading
        xmlEngine.writeTag "LastSaveFilter", "-1"  'Mark the last-used save filter as "unknown"
    xmlEngine.closeTag "FileFormat"
    xmlEngine.writeBlankLine
        
    'Write out tool preferences
    xmlEngine.writeTag "ToolPreferences", "", True
        xmlEngine.writeTag "ClearSelectionAfterCrop", "True"
        xmlEngine.writeTag "LastSelectionShape", "0"
        xmlEngine.writeTag "LastActiveTool", "0"
    xmlEngine.closeTag "ToolPreferences"
    xmlEngine.writeBlankLine
                
    'Finally, write out MRU placeholders
    xmlEngine.writeTag "MRU", "", True
        xmlEngine.writeTag "NumberOfEntries", "0"
        xmlEngine.writeTag "f0", ""
        xmlEngine.writeTag "f1", ""
        xmlEngine.writeTag "f2", ""
        xmlEngine.writeTag "f3", ""
        xmlEngine.writeTag "f4", ""
        xmlEngine.writeTag "f5", ""
        xmlEngine.writeTag "f6", ""
        xmlEngine.writeTag "f7", ""
        xmlEngine.writeTag "f8", ""
    xmlEngine.closeTag "MRU"
    xmlEngine.writeBlankLine
    
    'With all tags successfully written, we can now close the XML data and write it out to file.
    xmlEngine.writeXMLToFile PreferencesPath
    
End Sub

'Get a Boolean-type value from the preferences file.  (A default value must be supplied; this is used if no such value exists.)
Public Function GetPref_Boolean(ByRef preferenceSection As String, ByRef preferenceName As String, ByVal defaultValue As Boolean) As Boolean

    'Request the value (as a string)
    Dim tmpString As String
    tmpString = GetPreference(preferenceSection, preferenceName)
    
    'If the requested value DOES NOT exist, return the default value as supplied by the user
    If Len(tmpString) = 0 Then
        
        'To prevent future blank results, write out a default value
        SetPref_Boolean preferenceSection, preferenceName, defaultValue
        GetPref_Boolean = defaultValue
            
    'If the requested value DOES exist, convert it to boolean type and return it
    Else
        
        If tmpString = "False" Then
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
Public Function GetPref_Long(ByRef preferenceSection As String, ByRef preferenceName As String, ByVal defaultValue As Long) As Long

    'Get the value (as a string) from the INI file
    Dim tmpString As String
    tmpString = GetPreference(preferenceSection, preferenceName)
    
    'If the requested value DOES NOT exist, return the default value as supplied by the user
    If Len(tmpString) = 0 Then
    
        'To prevent future blank results, write out a default value
        SetPref_Long preferenceSection, preferenceName, defaultValue
        GetPref_Long = defaultValue
    
    'If the requested value DOES exist, convert it to Long type and return it
    Else
        GetPref_Long = CLng(tmpString)
    End If

End Function

'Set a Long-type value to the preferences file.
Public Sub SetPref_Long(ByRef preferenceSection As String, ByRef preferenceName As String, ByVal longVal As Long)
    WritePreference preferenceSection, preferenceName, CStr(longVal)
End Sub

'Get a String-type value from the preferences file.  (A default value must be supplied; this is used if no such value exists.)
Public Function GetPref_String(ByRef preferenceSection As String, ByRef preferenceName As String, ByVal defaultValue As String) As String

    'Get the requested value from the preferences file
    Dim tmpString As String
    tmpString = GetPreference(preferenceSection, preferenceName)
    
    'If the requested value DOES NOT exist, return the default value as supplied by the user
    If Len(tmpString) = 0 Then
        
        'To prevent future blank results, write out a default value
        SetPref_String preferenceSection, preferenceName, defaultValue
        GetPref_String = defaultValue
    
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
Public Function doesValueExist(ByRef preferenceSection As String, ByRef preferenceName As String) As Boolean
    Dim tmpString As String
    tmpString = GetPreference(preferenceSection, preferenceName)
    If Len(tmpString) = 0 Then doesValueExist = False Else doesValueExist = True
End Function

'Read a value from the preferences file and return it (as a string)
Private Function GetPreference(ByRef strSectionHeader As String, ByRef strVariableName As String) As String
    
    'I find it helpful to give preference strings names with spaces, to improve readability.  However, XML doesn't allow tags to have
    ' spaces in the name.  So remove any spaces before interacting with the XML file.
    strSectionHeader = Replace$(strSectionHeader, " ", "")
    strVariableName = Replace$(strVariableName, " ", "")
    
    'Create an XML object and load the XML settings file
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    xmlEngine.loadXMLFile PreferencesPath
    
    'Check for a few necessary tags, just to make sure this is actually a PhotoDemon preferences file
    If xmlEngine.isPDDataType("PhotoDemon User Preferences") And xmlEngine.validateLoadedXMLData("GeneralPreferences") Then
    
        'Return the tag at that location
        GetPreference = xmlEngine.getUniqueTag_String(strVariableName, , , strSectionHeader)
        
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
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    xmlEngine.loadXMLFile PreferencesPath
    
    'Check for a few necessary tags, just to make sure this is actually a PhotoDemon preferences file
    If xmlEngine.isPDDataType("PhotoDemon User Preferences") And xmlEngine.validateLoadedXMLData("GeneralPreferences") Then
    
        'Update the requested tag, and if it does not exist, write it out as a new tag at the end of the specified section
        If xmlEngine.updateTag(strVariableName, strValue, strSectionHeader) Then
            WritePreference = True
            xmlEngine.writeXMLToFile PreferencesPath
        Else
            WritePreference = False
        End If
        
    End If
    
End Function

'Get the current Windows-specified temp directory
Private Function GetTemporaryPath() As String
    
    'Create a blank string (as required by the API call)
    Dim sRet As String
    sRet = String(255, 0)
    
    'Fill that string with the temporary path
    Dim lngLen As Long
    lngLen = GetSystemTempPath(255, sRet)
    
    'If something went wrong, raise an error
    If lngLen = 0 Then Err.Raise Err.LastDllError
    
    'Return the temporary path, and ensure that it has a trailing "\"
    GetTemporaryPath = FixPath(Left$(sRet, lngLen))

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
