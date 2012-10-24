Attribute VB_Name = "Ini_Handler"
'***************************************************************************
'Program INI Handler
'Copyright ©2000-2012 by Tanner Helland
'Created: 26/September/01
'Last updated: 05/October/12
'Last update: store the INI file in the new program data subdirectory
'
'Module for handling the initialization of the program via an INI file.  This
' routine sets program defaults, determines folders, and generally prepares the
' information PhotoDemon requires for successful execution.
'
'***************************************************************************

Option Explicit

'API calls for interfacing with an INI file
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'API call for determining certain system folders
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'***
'Enormous code block for determining special Windows folders
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
Public Const CSIDL_FLAG_CREATE As Long = 32768 '&H8000 'combine with CSIDL_ value to force create on SHGetSpecialFolderLocation()
Public Const CSIDL_FLAG_DONT_VERIFY = &H4000 'combine with CSIDL_ value to force create on SHGetSpecialFolderLocation()
Public Const CSIDL_FLAG_MASK = &HFF00 'mask for all possible flag Values
Public Const SHGFP_TYPE_CURRENT = &H0 'current value for user, verify it exists
Public Const SHGFP_TYPE_DEFAULT = &H1
Public Const MAX_LENGTH = 260
Public Const S_OK = 0
Public Const S_FALSE = 1
'***

'Public variables that hold the path data
Public TempPath As String

'Path of the INI file
Private INIPath As String

'This LoadINI routine is one of the first things PhotoDemon launches. It pulls all user settings out of an INI file, and if an INI file
' cannot be found, it creates one from scratch.
Public Sub LoadINI()

    'Send a nice little message to the load form
    LoadMessage "Loading INI data..."
    
    'The INI file is located in the /Data subdirectory
    INIPath = DataPath & PROGRAMNAME & "_settings.ini"
    
    'Old versions of PhotoDemon kept the INI file in the same directory as the .exe file.  We now use a /Data subdirectory.
    ' Check the old location, and if an INI file is found, copy it to the new directory
    On Error Resume Next
    Dim tmpINIPath As String
    tmpINIPath = ProgramPath & PROGRAMNAME & "_settings.ini"
    
    If FileExist(tmpINIPath) Then
        LoadMessage "Copying old .INI file to new \Data subdirectory"
        FileCopy tmpINIPath, INIPath
        Kill tmpINIPath
    End If
    
    'This routine needs to open files.  To prevent "duplicate declarations in current scope," use a free file value
    Dim fileNum As Integer
    
    'If no INI file exists, let's build one from scratch
    If FileExist(INIPath) = False Then
        
        LoadMessage "INI file could not be located. Generating a new one..."
        
        'Finally, create a default INI file so we don't have to go through this again
        fileNum = FreeFile
    
        Open INIPath For Append As #fileNum
            Print #fileNum, "[PhotoDemon Program Specifications]"
            Print #fileNum, "BuildVersion=Beta"
            Print #fileNum, ""
            Print #fileNum, "[Paths]"
            Print #fileNum, "TempPath=" & GetTemporaryPath
            Print #fileNum, ""
            Print #fileNum, "[Program Paths]"
            Print #fileNum, "MainOpen=" & GetWindowsFolder(CSIDL_MYPICTURES)
            Print #fileNum, "MainSave=" & GetWindowsFolder(CSIDL_MYPICTURES)
            Print #fileNum, "ImportFRX=" & GetWindowsFolder(CSIDL_MY_DOCUMENTS)
            Print #fileNum, "CustomFilter=" & FilterPath
            Print #fileNum, "Macro=" & MacroPath
            Print #fileNum, ""
            Print #fileNum, "[File Formats]"
            Print #fileNum, "LastOpenFilter=1"   'Default to "All Compatible Graphics" filter for loading
            Print #fileNum, "LastSaveFilter=3"   'Default to JPEG for saving
            Print #fileNum, ""
            Print #fileNum, "[General Preferences]"
            Print #fileNum, "AlphaCheckMode=1"
            Print #fileNum, "AlphaCheckOne=" & RGB(153, 153, 153)
            Print #fileNum, "AlphaCheckTwo=" & RGB(102, 102, 102)
            Print #fileNum, "AlphaCheckSize=1"
            Print #fileNum, "AutosizeLargeImages=0"
            Print #fileNum, "CanvasBackground=16777215"
            Print #fileNum, "CheckForUpdates=1"
            Print #fileNum, "ConfirmClosingUnsaved=1"
            Print #fileNum, "LastUpdateCheck="
            Print #fileNum, "LogProgramMessages=0"
            Print #fileNum, "PromptForPluginDownload=1"
            Print #fileNum, "UseFancyFonts=0"
            Print #fileNum, ""
            Print #fileNum, "[Batch Preferences]"
            Print #fileNum, "DriveBox="
            Print #fileNum, "InputFolder=" & GetWindowsFolder(CSIDL_MYPICTURES)
            Print #fileNum, "OutputFolder=" & GetWindowsFolder(CSIDL_MYPICTURES)
            Print #fileNum, "ListFolder=" & GetWindowsFolder(CSIDL_MY_DOCUMENTS)
            Print #fileNum, ""
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
        Close #fileNum
        
    End If
    
    'Extract the temporary path from the INI
    TempPath = GetFromIni("Paths", "TempPath")
    
    'As a backup, make sure the temp path exists (to prevent future ugly errors)
    If Not DirectoryExist(TempPath) Then TempPath = GetTemporaryPath
    
    'Get the LogProgramMessages preference
    Dim tempString As String
    tempString = GetFromIni("General Preferences", "LogProgramMessages")
    x = Val(tempString)
    If x = 0 Then LogProgramMessages = False Else LogProgramMessages = True
    
    'If we're logging program messages, open up a log file and dump the date and time there
    If LogProgramMessages = True Then
        fileNum = FreeFile
    
        Open DataPath & PROGRAMNAME & "_DebugMessages.log" For Append As #fileNum
            Print #fileNum, vbCrLf
            Print #fileNum, vbCrLf
            Print #fileNum, "**********************************************"
            Print #fileNum, "Date: " & Date
            Print #fileNum, "Time: " & time
        Close #fileNum
    End If

    'Get the Canvas background preference (color vs checkerboard pattern)
    tempString = GetFromIni("General Preferences", "CanvasBackground")
    x = Val(tempString)
    CanvasBackground = x
    
    'Get the alpha transparency checkerboard pattern
    tempString = GetFromIni("General Preferences", "AlphaCheckMode")
    If tempString = "" Then AlphaCheckMode = 1 Else AlphaCheckMode = Val(tempString)
    
    'Get the alpha transparency checkerboard colors
    tempString = GetFromIni("General Preferences", "AlphaCheckOne")
    If tempString = "" Then AlphaCheckOne = RGB(153, 153, 153) Else AlphaCheckOne = Val(tempString)
    
    tempString = GetFromIni("General Preferences", "AlphaCheckTwo")
    If tempString = "" Then AlphaCheckTwo = RGB(102, 102, 102) Else AlphaCheckTwo = Val(tempString)
    
    'Get the alpha checkerboard size
    tempString = GetFromIni("General Preferences", "AlphaCheckSize")
    If tempString = "" Then AlphaCheckSize = 1 Else AlphaCheckSize = Val(tempString)
    
    'Check if the user wants us to prompt them about closing unsaved images
    tempString = GetFromIni("General Preferences", "ConfirmClosingUnsaved")
    x = Val(tempString)
    If x = 0 Then ConfirmClosingUnsaved = False Else ConfirmClosingUnsaved = True
    
    'Check if the user wants the interface to use modern typefaces if available
    tempString = GetFromIni("General Preferences", "UseFancyFonts")
    If tempString = "" Then
        useFancyFonts = False
    Else
        x = Val(tempString)
        If x = 0 Then useFancyFonts = False Else useFancyFonts = True
    End If
    
End Sub

'Read values from an INI
Public Function GetFromIni(strSectionHeader As String, strVariableName As String) As String
    
    'Blank out the return string (required by the API call)
    Dim strReturn As String
    strReturn = String(255, Chr(0))
    
    Dim chkReturn As String
    chkReturn = Left$(strReturn, GetPrivateProfileString(strSectionHeader, ByVal strVariableName, "", strReturn, Len(strReturn), INIPath))
    
    'Make sure a non-null string was returned
    If chkReturn <> "" Then
        GetFromIni = chkReturn
    
    'If a null string was returned, it may be because this key/value pair doesn't appear in the INI file.
    ' Force a write of this key/pair to ensure that it exists in the file.
    Else
        WritePrivateProfileString strSectionHeader, strVariableName, "", INIPath
        GetFromIni = ""
    End If
    
End Function

'Set values into an INI
Public Function WriteToIni(strSectionHeader As String, strVariableName As String, strValue As String) As Long
    WriteToIni = WritePrivateProfileString(strSectionHeader, strVariableName, strValue, INIPath)
End Function

'Get the temp directory from Windows
Public Function GetTemporaryPath() As String
    Dim sRet As String, lngLen As Long
    sRet = String(255, 0)
    lngLen = GetTempPath(255, sRet)
    If lngLen = 0 Then Err.Raise Err.LastDllError
    GetTemporaryPath = FixPath(Left$(sRet, lngLen))
End Function

'Get a special folder from Windows (as specified by the CSIDL)
Public Function GetWindowsFolder(eFolder As CSIDLs) As String

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
