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
'PhotoDemon Preferences (INI) Interface
'Copyright ©2011-2013 by Tanner Helland
'Created: 03/November/12
'Last updated: 19/November/12
'Last update: provide an externally accessible function for seeing if a key/pair value exists in the INI file
'
'This class is the newest incarnation of the former "INI Handler" module.  It is responsible for managing
' all interaction with user settings, which are currently stored in an INI file.  This class is also capable
' of generating a default INI file (necessary for first run), determining system folders (also necessary for
' first run), and saving/loading any user preferences to and from file.
'
'Because the INI file receives changes with almost every new version of the software, this class must also
' be capable of translating values between INI versions.  Typically these just means having a contingency plan
' for values that do not yet exist in the INI.
'
'When PhotoDemon is run for the first time, a default INI file is created.  To increase usability, I
' autopopulate this file with intelligently selected preferences - such as setting the user's initial Open/Save
' folders to the system "Pictures" directory, and automatically setting the "Temp" directory to the system's
' specified temp directory.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit


'API calls for interfacing with an INI file
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'API call for determining certain system folders
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'***
'API call for determining specific system-related folders
Private Declare Function SHGetFolderPath Lib "shfolder.dll" Alias "SHGetFolderPathA" (ByVal hWndOwner As Long, ByVal nFolder As CSIDLs, ByVal hToken As Long, ByVal dwReserved As Long, ByVal lpszPath As String) As Long

'Windows Folder Constants
Public Enum CSIDLs
    CSIDL_MY_DOCUMENTS = &H5 'My Documents
    ' CSIDL_WINDOWS = &H24 'GetWindowsDirectory()
    ' CSIDL_SYSTEM = &H25 'GetSystemDirectory()
    ' CSIDL_PROGRAM_FILES = &H26 'C:\Program Files
    ' CSIDL_START_MENU = &HB '{user name}\Start Menu
    ' CSIDL_FONTS = &H14 'windows\fonts
    ' CSIDL_DESKTOP = &H0 '{desktop}
    ' CSIDL_INTERNET = &H1 'Internet Explorer (icon on desktop)
    ' CSIDL_PROGRAMS = &H2 'Start Menu\Programs
    ' CSIDL_CONTROLS = &H3 'My Computer\Control Panel
    ' CSIDL_PRINTERS = &H4 'My Computer\Printers
    ' CSIDL_FAVORITES = &H6 '{user name}\Favorites
    ' CSIDL_STARTUP = &H7 'Start Menu\Programs\Startup
    ' CSIDL_RECENT = &H8 '{user name}\Recent
    ' CSIDL_SENDTO = &H9 '{user name}\SendTo
    ' CSIDL_BITBUCKET = &HA '{desktop}\Recycle Bin
    ' CSIDL_DESKTOPDIRECTORY = &H10 '{user name}\Desktop
    ' CSIDL_DRIVES = &H11 'My Computer
    ' CSIDL_NETWORK = &H12 'Network Neighborhood
    ' CSIDL_NETHOOD = &H13 '{user name}\nethood
    ' CSIDL_TEMPLATES = &H15
    ' CSIDL_COMMON_STARTMENU = &H16 'All Users\Start Menu
    ' CSIDL_COMMON_PROGRAMS = &H17 'All Users\Programs
    ' CSIDL_COMMON_STARTUP = &H18 'All Users\Startup
    ' CSIDL_COMMON_DESKTOPDIRECTORY = &H19 'All Users\Desktop
    ' CSIDL_APPDATA = &H1A '{user name}\Application Data
    ' CSIDL_PRINTHOOD = &H1B '{user name}\PrintHood
    ' CSIDL_LOCAL_APPDATA = &H1C '{user name}\Local Settings\Application Data (non roaming)
    ' CSIDL_ALTSTARTUP = &H1D 'non localized startup
    ' CSIDL_COMMON_ALTSTARTUP = &H1E 'non localized common startup
    ' CSIDL_COMMON_FAVORITES = &H1F
    ' CSIDL_INTERNET_CACHE = &H20
    ' CSIDL_COOKIES = &H21
    ' CSIDL_HISTORY = &H22
    ' CSIDL_COMMON_APPDATA = &H23 'All Users\Application Data
    CSIDL_MYPICTURES = &H27 'C:\Program Files\My Pictures
    ' CSIDL_PROFILE = &H28 'USERPROFILE
    ' CSIDL_SYSTEMX86 = &H29 'x86 system directory on RISC
    ' CSIDL_PROGRAM_FILESX86 = &H2A 'x86 C:\Program Files on RISC
    ' CSIDL_PROGRAM_FILES_COMMON = &H2B 'C:\Program Files\Common
    ' CSIDL_PROGRAM_FILES_COMMONX86 = &H2C 'x86 Program Files\Common on RISC
    ' CSIDL_COMMON_TEMPLATES = &H2D 'All Users\Templates
    ' CSIDL_COMMON_DOCUMENTS = &H2E 'All Users\Documents
    ' CSIDL_COMMON_ADMINTOOLS = &H2F 'All Users\Start Menu\Programs\Administrative Tools
    ' CSIDL_ADMINTOOLS = &H30 '{user name}\Start Menu\Programs\Administrative Tools
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

'Location of various relevant program directories
Private ProgramPath As String
Private DataPath As String
Private MacroPath As String
Private FilterPath As String
Private INIPath As String
Private TempPath As String
Private IconPath As String
Private LanguagePath As String

'To make PhotoDemon compatible with the PortableApps spec (http://portableapps.com/), several sub-folders are necessary.  These include:
'  /App/ subfolder, which contains information ESSENTIAL and UNIVERSAL for each PhotoDemon install (e.g. FreeImage, zLib, EZTwain dlls)
'  /Data/ subfolder, which contains information that is OPTIONAL and UNIQUE for each PhotoDemon install (e.g. user's INI file, saved filters)
Private AppPath As String

'Return the current Language directory
Public Function getLanguagePath() As String
    getLanguagePath = LanguagePath
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
    
    SetPreference_String "Program Paths", "TempPath", TempPath
    
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
    SetPreference_String "Program Paths", "Macro", MacroPath
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
    SetPreference_String "Program Paths", "Filter", FilterPath
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
    
    'Create an "App" path based off the ProgramPath - this is where essential plugins are stored.
    ' Also, make sure the directory exists; if it doesn't, create it.
    AppPath = ProgramPath & "App\"
    If Not DirectoryExist(AppPath) Then MkDir AppPath
    
    AppPath = AppPath & "PhotoDemon\"
    If Not DirectoryExist(AppPath) Then MkDir AppPath
    
    'Within the App\PhotoDemon\ folder, create a folder for any available translations
    LanguagePath = AppPath & "Languages\"
    If Not DirectoryExist(LanguagePath) Then MkDir LanguagePath
        
    'Create a "Data" path based off the ProgramPath - this is where the INI file and any other user-specific files (saved filters, macros) will be stored.
    ' Also, make sure the directory exists; if it doesn't, create it.
    DataPath = ProgramPath & "Data\"
    If Not DirectoryExist(DataPath) Then MkDir DataPath
    
    'Within the \Data subfolder, check for three more folders - saved Macros, saved convolution filters...
    MacroPath = DataPath & "Macros\"
    If Not DirectoryExist(MacroPath) Then MkDir MacroPath
    
    FilterPath = DataPath & "Filters\"
    If Not DirectoryExist(FilterPath) Then MkDir FilterPath
        
    'Finally, if it does not exist, create a folder to hold the MRU icons
    IconPath = DataPath & "Icons\"
    If Not DirectoryExist(IconPath) Then MkDir IconPath
    
    'The INI file is also located in the \Data subdirectory
    INIPath = DataPath & PROGRAMNAME & "_settings.ini"
    
    'Historically, language files were stored in a different location.  We need to move any valid language files from
    ' the old language directory to the new language directory.
    On Error Resume Next
    Dim chkFile As String
    
    chkFile = Dir(DataPath & "Languages\" & "*.xml", vbNormal)
        
    If chkFile <> "" Then
        
        LoadMessage "Moving language files to updated folder location..."
        
        Do While chkFile <> ""
            If FileExist(DataPath & "Languages\" & chkFile) Then
                FileCopy DataPath & "Languages\" & chkFile, LanguagePath & chkFile
                Kill DataPath & "Languages\" & chkFile
            End If
            chkFile = Dir
        Loop
                
        RmDir DataPath & "Languages\"
        
    End If
        
End Sub

'Load all user settings from file
Public Sub loadUserSettings()
        
    'If no INI file exists, build a default one
    If FileExist(INIPath) = False Then createNewINIFile
            
    'Get the LogProgramMessages preference
    g_LogProgramMessages = GetPreference_Boolean("General Preferences", "LogProgramMessages", False)
    
    'If the user wants us to log program messages, open the log file and write out the current date and time
    If g_LogProgramMessages Then
        
        Dim fileNum As Integer
        fileNum = FreeFile
    
        Open DataPath & PROGRAMNAME & "_DebugMessages.log" For Append As #fileNum
            Print #fileNum, vbCrLf
            Print #fileNum, vbCrLf
            Print #fileNum, "**********************************************"
            Print #fileNum, "Date: " & Date
            Print #fileNum, "Time: " & time
        Close #fileNum
        
    End If
            
    'Extract the temporary path from the INI and make sure it exists. (If it doesn't,
    ' transparently set it to the system temporary path.)
    TempPath = GetPreference_String("Program Paths", "TempPath", "")
    If Not DirectoryExist(TempPath) Then
        TempPath = GetTemporaryPath
        SetPreference_String "Program Paths", "TempPath", TempPath
    End If
    
    'Extract the macro path from the INI
    MacroPath = GetPreference_String("Program Paths", "Macro", MacroPath)
    
    'Extract the filter path from the INI
    FilterPath = GetPreference_String("Program Paths", "CustomFilter", FilterPath)
    
    'Get the canvas background preference (color vs checkerboard pattern)
    Dim x As Long
    x = GetPreference_Long("General Preferences", "CanvasBackground", RGB(201, 211, 226))
    
    'A value of -1 previously indicated a checkerboard background, which is no longer allowed.
    ' If that value is found, overwrite it with a pleasant shade of blue.
    If x = -1 Then
        x = RGB(201, 211, 226)
        SetPreference_Long "General Preferences", "CanvasBackground", CStr(RGB(201, 211, 226))
    End If
    g_CanvasBackground = x
    
    'Get the canvas drop shadow preference
    g_CanvasDropShadow = GetPreference_Boolean("General Preferences", "CanvasDropShadow", False)
    
    'Get the alpha transparency checkerboard pattern
    g_AlphaCheckMode = GetPreference_Long("General Preferences", "AlphaCheckMode", 1)
    
    'Get the alpha transparency checkerboard colors
    g_AlphaCheckOne = GetPreference_Long("General Preferences", "AlphaCheckOne", RGB(153, 153, 153))
    g_AlphaCheckTwo = GetPreference_Long("General Preferences", "AlphaCheckTwo", RGB(102, 102, 102))
    
    'Get the alpha checkerboard size
    g_AlphaCheckSize = GetPreference_Long("General Preferences", "AlphaCheckSize", 1)
    
    'Check if the user wants us to prompt them about closing unsaved images
    g_ConfirmClosingUnsaved = GetPreference_Boolean("General Preferences", "ConfirmClosingUnsaved", True)
        
    'Check if the user wants the interface to use modern typefaces if available
    g_UseFancyFonts = GetPreference_Boolean("General Preferences", "UseFancyFonts", False)
    
    'Check if the user wants large images automatically zoomed when loaded
    g_AutosizeLargeImages = GetPreference_Long("General Preferences", "AutosizeLargeImages", 0)
    
    'Grab the last-used common dialog filters
    g_LastOpenFilter = GetPreference_Long("File Formats", "LastOpenFilter", 1)
    g_LastSaveFilter = GetPreference_Long("File Formats", "LastSaveFilter", 3)
    
    'Initialize a few other settings to default values, despite them not coming from the INI file
    
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
        FormMain.cmbSelRender(i).ListIndex = GetPreference_Long("Tool Preferences", "LastSelectionShape", 0)
    Next i
    
End Sub

Public Sub saveToolSettings()
    
    'Last-used tool
    SetPreference_Long "Tool Preferences", "LastActiveTool", g_CurrentTool
        
End Sub

'Reset the INI file to its default state.  (Basically, delete any existing file then call createNewINIFile.)
Public Sub resetPreferences()

    If FileExist(INIPath) Then Kill INIPath
    
    createNewINIFile

    loadUserSettings

End Sub

'Create a new INI file from scratch
Private Sub createNewINIFile()

    'This function is used to determine whether PhotoDemon is being run for the first time.  Why do it here?
    ' 1) When first downloaded, PhotoDemon doesn't come with an INI.  Thus this routine MUST be called.
    ' 2) When preferences are reset, this file is deleted.  That is an appropriate time to mark the program as "first run".
    ' 3) If the user moves PhotoDemon but leaves behind the old INI file.  There's no easy way to check this, but
    '     treating the program as if its being run for the first time is as good a plan as any.
    g_IsFirstRun = True

    'This routine needs to open a file.  To prevent "duplicate declarations in current scope," use a free file value
    Dim fileNum As Integer
    fileNum = FreeFile
    
    'Open the file (which will create it when it doesn't exist)
    Open INIPath For Append As #fileNum
    
        'Write out the INI header
        Print #fileNum, "[PhotoDemon Program Specifications]"
        Print #fileNum, "BuildVersion=Beta"
        Print #fileNum, ""
        
        'Write out all program paths
        Print #fileNum, "[Program Paths]"
        Print #fileNum, "TempPath=" & GetTemporaryPath
        Print #fileNum, "MainOpen=" & GetWindowsFolder(CSIDL_MYPICTURES)
        Print #fileNum, "MainSave=" & GetWindowsFolder(CSIDL_MYPICTURES)
        Print #fileNum, "ImportFRX=" & GetWindowsFolder(CSIDL_MY_DOCUMENTS)
        Print #fileNum, "CustomFilter=" & FilterPath
        Print #fileNum, "Macro=" & MacroPath
        Print #fileNum, ""
                
        'Write out general preferences
        Print #fileNum, "[General Preferences]"
        Print #fileNum, "AlphaCheckMode=0"
        Print #fileNum, "AlphaCheckOne=" & RGB(255, 255, 255)
        Print #fileNum, "AlphaCheckTwo=" & RGB(204, 204, 204)
        Print #fileNum, "AlphaCheckSize=1"
        Print #fileNum, "AutosizeLargeImages=0"
        Print #fileNum, "BitmapRLE=0"
        Print #fileNum, "CanvasBackground=" & vb3DShadow 'RGB(201, 211, 226)
        Print #fileNum, "CanvasDropShadow=1"
        Print #fileNum, "CheckForUpdates=1"
        Print #fileNum, "ConfirmClosingUnsaved=1"
        Print #fileNum, "CurrentLanguageFile="
        Print #fileNum, "DefaultSaveFormat=0"
        Print #fileNum, "DisplayIDEWarning=1"
        Print #fileNum, "DynamicTaskbarIcon=1"
        Print #fileNum, "HasGitHubAccount="
        Print #fileNum, "HideLeftPanel=0"
        Print #fileNum, "HideRightPanel=0"
        Print #fileNum, "ImageCaptionSize=0"
        Print #fileNum, "LastUpdateCheck="
        Print #fileNum, "LastPreferencesPage=0"
        Print #fileNum, "LastFilePreferencesPage=1"
        Print #fileNum, "LogProgramMessages=0"
        Print #fileNum, "MRUCaptionSize=0"
        Print #fileNum, "MultipageImagePrompt=0"
        Print #fileNum, "OutgoingColorDepth=1"
        Print #fileNum, "PNGBackgroundPreservation=1"
        Print #fileNum, "PNGCompression=9"
        Print #fileNum, "PNGInterlacing=0"
        Print #fileNum, "PPMExportFormat=0"
        Print #fileNum, "PromptForPluginDownload=1"
        Print #fileNum, "RememberWindowLocation=1"
        Print #fileNum, "LastWindowState=0"
        Print #fileNum, "LastWindowLeft=1"
        Print #fileNum, "LastWindowTop=1"
        Print #fileNum, "LastWindowWidth=1"
        Print #fileNum, "LastWindowHeight=1"
        Print #fileNum, "SaveBehavior=0"
        Print #fileNum, "TGARLE=0"
        Print #fileNum, "TIFFCMYK=0"
        Print #fileNum, "TIFFCompression=0"
        'On Vista and later, Segoe UI is used as the interface font instead of Tahoma
        If g_IsVistaOrLater Then Print #fileNum, "UseFancyFonts=1" Else Print #fileNum, "UseFancyFonts=0"
        Print #fileNum, "UseToneMapping=1"
        Print #fileNum, "ValidateAlphaChannels=1"
        Print #fileNum, "VerifyInitialColorDepth=1"
        Print #fileNum, ""
        
        'Write out plugin preferences
        Print #fileNum, "[Plugin Preferences]"
        Print #fileNum, "ForceFreeImageDisable=0"
        Print #fileNum, "ForceZLibDisable=0"
        Print #fileNum, "ForceEZTwainDisable=0"
        Print #fileNum, "ForcePngnqDisable=0"
        Print #fileNum, "ForceExifToolDisable=0"
        Print #fileNum, "LastPluginPreferencesPage=0"
        Print #fileNum, "PngnqAlphaExtenuation=0"
        Print #fileNum, "PngnqYUV=1"
        Print #fileNum, "PngnqColorSample=3"
        Print #fileNum, "PngnqDithering=5"
        Print #fileNum, ""
        
        'Write out batch conversion preferences
        Print #fileNum, "[Batch Preferences]"
        Print #fileNum, "DriveBox="
        Print #fileNum, "InputFolder=" & GetWindowsFolder(CSIDL_MYPICTURES)
        Print #fileNum, "OutputFolder=" & GetWindowsFolder(CSIDL_MYPICTURES)
        Print #fileNum, "ListFolder=" & GetWindowsFolder(CSIDL_MY_DOCUMENTS)
        Print #fileNum, ""
        
        'Write out file format preferences
        Print #fileNum, "[File Formats]"
        Print #fileNum, "LastOpenFilter=1"   'Default to "All Compatible Graphics" filter for loading
        Print #fileNum, "LastSaveFilter=-1"  'Mark the last-used save filter as "unknown"
        Print #fileNum, ""
        
        'Write out tool preferences
        Print #fileNum, "[Tool Preferences]"
        Print #fileNum, "ClearSelectionAfterCrop=1"
        Print #fileNum, "LastSelectionShape=0"
        Print #fileNum, "LastActiveTool=0"
        Print #fileNum, ""
                
        'Finally, write out MRU placeholders
        Print #fileNum, "[MRU]"
        Print #fileNum, "NumberOfEntries=0"
        Print #fileNum, "f0="
        Print #fileNum, "f1="
        Print #fileNum, "f2="
        Print #fileNum, "f3="
        Print #fileNum, "f4="
        Print #fileNum, "f5="
        Print #fileNum, "f6="
        Print #fileNum, "f7="
        Print #fileNum, "f8="
        
    'Close the newly created INI file
    Close #fileNum
    
End Sub

'Get a Boolean-type value from the INI file.  (A default value must be supplied; this is used if no value exists in the INI file.)
Public Function GetPreference_Boolean(ByRef preferenceSection As String, ByRef preferenceName As String, ByVal defaultValue As Boolean) As Boolean

    'Get the value (as a string) from the INI file
    Dim tmpString As String
    tmpString = GetINIString(preferenceSection, preferenceName)
    
    'If the requested value DOES NOT exist, return the default value as supplied by the user
    If tmpString = "" Then
        
        'To prevent future blank results, write the default value to the INI file
        SetPreference_Boolean preferenceSection, preferenceName, defaultValue
        
        GetPreference_Boolean = defaultValue
            
    'If the requested value DOES exist, convert it to boolean type and return it
    Else
        
        'A string of "0" is treated as FALSE.  All other values are treated as TRUE.
        If tmpString = "0" Then
            GetPreference_Boolean = False
        Else
            GetPreference_Boolean = True
        End If
    
    End If

End Function

'Set a Boolean-type value to the INI file.
Public Sub SetPreference_Boolean(ByRef preferenceSection As String, ByRef preferenceName As String, ByVal boolVal As Boolean)

    'If the boolean value is FALSE, set "0" to the INI file.  Otherwise, set "1".
    If boolVal Then
        WriteINIString preferenceSection, preferenceName, "1"
    Else
        WriteINIString preferenceSection, preferenceName, "0"
    End If

End Sub

'Get a Long-type value from the INI file.  (A default value must be supplied; this is used if no value exists in the INI file.)
Public Function GetPreference_Long(ByRef preferenceSection As String, ByRef preferenceName As String, ByVal defaultValue As Long) As Long

    'Get the value (as a string) from the INI file
    Dim tmpString As String
    tmpString = GetINIString(preferenceSection, preferenceName)
    
    'If the requested value DOES NOT exist, return the default value as supplied by the user
    If tmpString = "" Then
    
        'To prevent future blank results, write the default value to the INI file
        SetPreference_Long preferenceSection, preferenceName, defaultValue
    
        GetPreference_Long = defaultValue
    
    'If the requested value DOES exist, convert it to Long type and return it
    Else
        GetPreference_Long = CLng(tmpString)
    End If

End Function

'Set a Long-type value to the INI file.
Public Sub SetPreference_Long(ByRef preferenceSection As String, ByRef preferenceName As String, ByVal longVal As Long)

    WriteINIString preferenceSection, preferenceName, CStr(longVal)

End Sub

'Get a String-type value from the INI file.  (A default value must be supplied; this is used if no value exists in the INI file.)
Public Function GetPreference_String(ByRef preferenceSection As String, ByRef preferenceName As String, ByVal defaultValue As String) As String

    'Get the requested value from the INI file
    Dim tmpString As String
    tmpString = GetINIString(preferenceSection, preferenceName)
    
    'If the requested value DOES NOT exist, return the default value as supplied by the user
    If tmpString = "" Then
        
        'To prevent future blank results, write the default value to the INI file
        SetPreference_String preferenceSection, preferenceName, defaultValue
    
        GetPreference_String = defaultValue
    
    'If the requested value DOES exist, convert it to Long type and return it
    Else
        GetPreference_String = tmpString
    End If

End Function

'Set a String-type value to the INI file.
Public Sub SetPreference_String(ByRef preferenceSection As String, ByRef preferenceName As String, ByRef stringVal As String)

    WriteINIString preferenceSection, preferenceName, stringVal

End Sub

'Sometimes we want to know if a value exists at all.  This function handles that.
Public Function doesValueExist(ByRef preferenceSection As String, ByRef preferenceName As String) As Boolean
    Dim tmpString As String
    tmpString = GetINIString(preferenceSection, preferenceName)
    If tmpString = "" Then doesValueExist = False Else doesValueExist = True
End Function

'Read a value from an INI file and return it (as a string)
Private Function GetINIString(ByRef strSectionHeader As String, ByRef strVariableName As String) As String
    
    'Blank out the return string (required by the API call)
    Dim strReturn As String
    strReturn = String(255, Chr(0))
    
    'Get the value form the file
    Dim chkReturn As String
    chkReturn = Left$(strReturn, GetPrivateProfileString(strSectionHeader, ByVal strVariableName, "", strReturn, Len(strReturn), INIPath))
    
    'Make sure a non-null string was returned
    If chkReturn <> "" Then
        GetINIString = chkReturn
    
    'If a null string was returned, it may be because this key/value pair doesn't appear in the INI file.
    ' Force a write of this key/pair to ensure that it exists in the file.
    Else
        WritePrivateProfileString strSectionHeader, strVariableName, "", INIPath
        GetINIString = ""
    End If
    
End Function

'Write a string value to an INI file
Public Function WriteINIString(strSectionHeader As String, strVariableName As String, strValue As String) As Long
    WriteINIString = WritePrivateProfileString(strSectionHeader, strVariableName, strValue, INIPath)
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

'The next two functions are used as part of the update process.  Thus they require specifying an arbitrary INI path as well.

'Read values from an arbitrary INI file (e.g. NOT the core PhotoDemon one)
Public Function GetFromArbitraryIni(strINIFile As String, strSectionHeader As String, strVariableName As String) As String
    Dim strReturn As String
    'Blank out the string (required by the API call)
    strReturn = String(255, Chr(0))
    GetFromArbitraryIni = Left$(strReturn, GetPrivateProfileString(strSectionHeader, ByVal strVariableName, "", strReturn, Len(strReturn), strINIFile))
End Function

'Set values into an arbitrary INI file (e.g. NOT the core PhotoDemon one)
Public Function WriteToArbitraryIni(strINIFile As String, strSectionHeader As String, strVariableName As String, strValue As String) As Long
    WriteToArbitraryIni = WritePrivateProfileString(strSectionHeader, strVariableName, strValue, strINIFile)
End Function

