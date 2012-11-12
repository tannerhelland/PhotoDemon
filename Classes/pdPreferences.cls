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
'Copyright ©2000-2012 by Tanner Helland
'Created: 03/November/12
'Last updated: 09/November/12
'Last update: finished initial build
'
'This class is the newest incarnation of the former "INI Handler" module.  It is responsible for managing
' all interaction with user settings, which are currently stored in an INI file.  This class is also capable
' of generating a default INI file (necessary for first run), determining system folders (also necessary for
' first run), and saving/loading any user preferences to and from file.
'
'Because the INI file receives changes with almost every new version of the software, this class must also
' be capable of translating values between INI versions.  Typically these just means having a contingency plan
' for values that do not yet exist in the INI file.
'
'When PhotoDemon is run for the first time, a default INI file is created.  To increase usability, I
' autopopulate this file with intelligently selected preferences - such as setting the user's initial Open/Save
' folders to the system "Pictures" directory, and automatically setting the "Temp" directory to the system's
' specified temp directory.
'
'TODO: give the user a way to reset all preferences (e.g. rewrite the INI file from scratch)
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
Private Declare Function SHGetFolderPath Lib "shfolder.dll" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As CSIDLs, ByVal hToken As Long, ByVal dwReserved As Long, ByVal lpszPath As String) As Long

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

'Return the current temporary directory
Public Function getTempPath() As String
    getTempPath = TempPath
End Function

'Set the current temp directory
Public Sub setTempPath(ByVal newTempPath As String)
    StripDirectory newTempPath
    newTempPath = FixPath(newTempPath)
    TempPath = newTempPath
    SetPreference_String "Program Paths", "TempPath", TempPath
End Sub

'Return the current program directory
Public Function getProgramPath() As String
    getProgramPath = ProgramPath
End Function

'Return the current data directory
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

'Initialize key program directories
Public Sub initializePaths()
    
    'First things first: figure out where this .exe was launched from
    ProgramPath = App.Path
    If Right(ProgramPath, 1) <> "\" Then ProgramPath = ProgramPath & "\"
    
    'Create a "Data" path based off the ProgramPath - this is where plugins, the INI file, Help, and more are stored.
    ' Also, make sure the directory exists; if it doesn't, create it.
    DataPath = ProgramPath & "Data\"
    If Not DirectoryExist(DataPath) Then MkDir DataPath
    
    'Within the \Data subfolder, check for two more folders - one for saved Macros and another for saved convolution filters
    MacroPath = DataPath & "Macros\"
    If Not DirectoryExist(MacroPath) Then MkDir MacroPath
    
    FilterPath = DataPath & "Filters\"
    If Not DirectoryExist(FilterPath) Then MkDir FilterPath
    
    'The INI file is also located in the \Data subdirectory
    INIPath = DataPath & PROGRAMNAME & "_settings.ini"
    
    'Old versions of PhotoDemon kept the INI file in the same directory as the .exe file.  Check this old location,
    ' and if an INI file is found, copy it to the new directory.
    On Error Resume Next
    Dim tmpINIPath As String
    tmpINIPath = ProgramPath & PROGRAMNAME & "_settings.ini"
    If FileExist(tmpINIPath) Then
        FileCopy tmpINIPath, INIPath
        Kill tmpINIPath
    End If
    
End Sub

'Load all user settings from file
Public Sub loadUserSettings()
        
    'If no INI file exists, build a default one
    If FileExist(INIPath) = False Then createNewINIFile
            
    'Get the LogProgramMessages preference
    LogProgramMessages = GetPreference_Boolean("General Preferences", "LogProgramMessages", False)
    
    'If the user wants us to log program messages, open the log file and write out the current date and time
    If LogProgramMessages = True Then
        
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
    CanvasBackground = x
    
    'Get the canvas drop shadow preference
    CanvasDropShadow = GetPreference_Boolean("General Preferences", "CanvasDropShadow", False)
    
    'Get the alpha transparency checkerboard pattern
    AlphaCheckMode = GetPreference_Long("General Preferences", "AlphaCheckMode", 1)
    
    'Get the alpha transparency checkerboard colors
    AlphaCheckOne = GetPreference_Long("General Preferences", "AlphaCheckOne", RGB(153, 153, 153))
    AlphaCheckTwo = GetPreference_Long("General Preferences", "AlphaCheckTwo", RGB(102, 102, 102))
    
    'Get the alpha checkerboard size
    AlphaCheckSize = GetPreference_Long("General Preferences", "AlphaCheckSize", 1)
    
    'Check if the user wants us to prompt them about closing unsaved images
    ConfirmClosingUnsaved = GetPreference_Boolean("General Preferences", "ConfirmClosingUnsaved", True)
        
    'Check if the user wants the interface to use modern typefaces if available
    useFancyFonts = GetPreference_Boolean("General Preferences", "UseFancyFonts", False)
    
    'Check if the user wants large images automatically zoomed when loaded
    AutosizeLargeImages = GetPreference_Long("General Preferences", "AutosizeLargeImages", 0)
    
End Sub

'Reset the INI file to its default state.  (Basically, delete any existing file then call createNewINIFile.)
Public Sub resetPreferences()

    If FileExist(INIPath) Then Kill INIPath
    
    createNewINIFile

    loadUserSettings

End Sub

'Create a new INI file from scratch
Private Sub createNewINIFile()

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
        
        'Write out file format preferences
        Print #fileNum, "[File Formats]"
        Print #fileNum, "LastOpenFilter=1"   'Default to "All Compatible Graphics" filter for loading
        Print #fileNum, "LastSaveFilter=3"   'Default to JPEG for saving
        Print #fileNum, ""
        
        'Write out general preferences
        Print #fileNum, "[General Preferences]"
        Print #fileNum, "AlphaCheckMode=1"
        Print #fileNum, "AlphaCheckOne=" & RGB(153, 153, 153)
        Print #fileNum, "AlphaCheckTwo=" & RGB(102, 102, 102)
        Print #fileNum, "AlphaCheckSize=1"
        Print #fileNum, "AutosizeLargeImages=0"
        Print #fileNum, "CanvasBackground=" & vb3DShadow 'RGB(201, 211, 226)
        Print #fileNum, "CanvasDropShadow=1"
        Print #fileNum, "CheckForUpdates=1"
        Print #fileNum, "ConfirmClosingUnsaved=1"
        Print #fileNum, "LastUpdateCheck="
        Print #fileNum, "LogProgramMessages=0"
        Print #fileNum, "PromptForPluginDownload=1"
        Print #fileNum, "RememberWindowLocation=1"
        Print #fileNum, "LastWindowState=0"
        Print #fileNum, "LastWindowLeft=1"
        Print #fileNum, "LastWindowTop=1"
        Print #fileNum, "LastWindowWidth=1"
        Print #fileNum, "LastWindowHeight=1"
        Print #fileNum, "UseFancyFonts=0"
        Print #fileNum, ""
        
        'Write out batch conversion preferences
        Print #fileNum, "[Batch Preferences]"
        Print #fileNum, "DriveBox="
        Print #fileNum, "InputFolder=" & GetWindowsFolder(CSIDL_MYPICTURES)
        Print #fileNum, "OutputFolder=" & GetWindowsFolder(CSIDL_MYPICTURES)
        Print #fileNum, "ListFolder=" & GetWindowsFolder(CSIDL_MY_DOCUMENTS)
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
    If boolVal = False Then
        WriteINIString preferenceSection, preferenceName, "0"
    Else
        WriteINIString preferenceSection, preferenceName, "1"
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

