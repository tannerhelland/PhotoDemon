Attribute VB_Name = "Plugin_ExifTool_Interface"
'***************************************************************************
'ExifTool Plugin Interface
'Copyright ©2012-2013 by Tanner Helland
'Created: 24/May/13
'Last updated: 27/May/13
'Last update: use custom separators for list-type values; this is much easier than parsing them manually
'
'Module for handling all ExifTool interfacing.  This module is pointless without the accompanying ExifTool plugin,
' which can be found in the App/PhotoDemon/Plugins subdirectory as "exiftool.exe".  The ExifTool plugin will be
' available by default in all versions of PhotoDemon after and including 5.6 (release TBD, estimated summer 2013).
'
'ExifTool is a comprehensive image metadata handler written by Phil Harvey.  No DLL or VB-compatible library is
' available, so PhotoDemon relies on the stock Windows ExifTool executable file for all interfacing.  You can read
' more about ExifTool at its homepage:
'
'http://www.sno.phy.queensu.ca/~phil/exiftool/
'
'stdout is piped so PhotoDemon can read ExifTool output in real-time.  I used a sample VB module as a reference
' while developing this code, courtesy of Michael Wandel:
'
'http://owl.phy.queensu.ca/~phil/exiftool/modExiftool_101.zip
'
'...as well as a modified piping function, derived from code originally written by Joacim Andersson:
'
'http://www.vbforums.com/showthread.php?364219-Classic-VB-How-do-I-shell-a-command-line-program-and-capture-the-output
'
'This project was designed against v9.29 of ExifTool (18 May '13).  It may not work with other versions of the
' software.  Additional documentation regarding the use of ExifTool can be found in the official ExifTool
' package, downloadable from http://www.sno.phy.queensu.ca/~phil/exiftool/Image-ExifTool-9.29.tar.gz
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'A number of API functions are required to pipe stdout
Private Const STARTF_USESHOWWINDOW = &H1
Private Const STARTF_USESTDHANDLES = &H100
Private Const SW_NORMAL = 1
Private Const SW_HIDE = 0
Private Const DUPLICATE_CLOSE_SOURCE = &H1
Private Const DUPLICATE_SAME_ACCESS = &H2

'Potential error codes (not used at present, but could be added in the future)
Private Const ERROR_BROKEN_PIPE = 109

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As Any, ByVal nSize As Long) As Long
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function DuplicateHandle Lib "kernel32" (ByVal hSourceProcessHandle As Long, ByVal hSourceHandle As Long, ByVal hTargetProcessHandle As Long, lpTargetHandle As Long, ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwOptions As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long

'This type is what PhotoDemon uses internally for storing and displaying metadata
Public Type mdItem
    FullGroupAndName As String
    Group As String
    SubGroup As String
    Name As String
    Description As String
    Value As String
    ActualValue As String
    isValueBinary As Boolean
    isValueList As Boolean
    isActualValueBinary As Boolean
    isActualValueList As Boolean
    markedForRemoval As Boolean
End Type


'Is ExifTool available as a plugin?
Public Function isExifToolAvailable() As Boolean
    If FileExist(g_PluginPath & "exiftool.exe") Then isExifToolAvailable = True Else isExifToolAvailable = False
End Function

'Retrieve the ExifTool version.
Public Function getExifToolVersion() As String

    If Not isExifToolAvailable Then
        getExifToolVersion = ""
        Exit Function
    Else
        
        Dim exifPath As String
        exifPath = g_PluginPath & "exiftool.exe"
        
        Dim outputString As String
        If ShellExecuteCapture(exifPath, "exiftool.exe -ver", outputString) Then
        
            'The output string will be a simple version number, e.g. "X.YY", and it will be terminated by a vbCrLf character.
            ' Remove vbCrLf now.
            outputString = Trim$(outputString)
            outputString = Replace(outputString, vbCrLf, "")
            getExifToolVersion = outputString
            
        Else
            getExifToolVersion = ""
        End If
        
    End If
    
End Function

'Given a path to a valid image file, retrieve all metadata into a single (enormous) string
Public Function getMetadata(ByVal srcFile As String, ByVal srcFormat As Long) As String
    
    'Many ExifTool options are delimited by quotation marks (").  Because VB has the worst character escaping scheme ever conceived, I use
    ' a variable to hold the ASCII equivalent of a quotation mark.  This makes things slightly more readable.
    Dim Quotes As String
    Quotes = Chr(34)
    
    'Grab the ExifTool path, which we will shell and pipe in a moment
    Dim appLocation As String
    appLocation = g_PluginPath & "exiftool.exe"
    
    'Next, build a string of command-line parameters.  These will modify ExifTool's behavior to make it compatible with our code.
    Dim cmdParams As String
    
    'Ignore minor errors and warnings (note that exiftool.exe must be included as param [0], per C convention)
    cmdParams = cmdParams & "exiftool.exe -m"
    
    'Output long-format data
    cmdParams = cmdParams & " -l"
        
    'Request ANSI-compatible text (Windows-1252 specifically)
    cmdParams = cmdParams & " -L"
    
    'Request a custom separator for list-type values
    cmdParams = cmdParams & " -sep ;"
        
    'If a translation is active, request descriptions in the current language
    If g_Language.translationActive Then cmdParams = cmdParams & " -lang " & g_Language.getCurrentLanguage()
    
    'Request that binary data be processed.  We have no use for this data within PD, but when it comes time to write
    ' our metadata back out to file, we need to have a copy of it.
    cmdParams = cmdParams & " -b"
    
    'Requesting binary data also means preview and thumbnail images will be processed.  We DEFINITELY don't want these,
    ' so deny them specifically.
    cmdParams = cmdParams & " -x PreviewImage -x ThumbnailImage -x PhotoshopThumbnail"
    
    'Output XML data (a lot more complex, but the only way to retrieve descriptions and names simultaneously)
    cmdParams = cmdParams & " -X"
    
    'Finally, add the image path
    cmdParams = cmdParams & " " & Quotes & srcFile & Quotes
    
    'MsgBox cmdParams
    
    'NOTE: while in the IDE, it's useful to see ExifTool's output, so the command-line window is displayed there.
    
    'Use ExifTool to retrieve this image's metadata
    If Not ShellExecuteCapture(appLocation, cmdParams, getMetadata, Not g_IsProgramCompiled) Then
        Message "Failed to retrieve metadata."
        getMetadata = ""
    End If
    
    'Because we request metadata in XML format, ExifTool escapes disallowed XML characters.  Convert those back to standard characters now.
    If InStr(1, getMetadata, "&amp;") > 0 Then getMetadata = Replace(getMetadata, "&amp;", "&")
    If InStr(1, getMetadata, "&#39;") > 0 Then getMetadata = Replace(getMetadata, "&#39;", "'")
    If InStr(1, getMetadata, "&quot;") > 0 Then getMetadata = Replace(getMetadata, "&quot;", """")
    If InStr(1, getMetadata, "&gt;") > 0 Then getMetadata = Replace(getMetadata, "&gt;", ">")
    If InStr(1, getMetadata, "&lt;") > 0 Then getMetadata = Replace(getMetadata, "&lt;", "<")
    
End Function

'Given a path to a valid metadata file, and a second path to a valid image file, copy the metadata file into the image file.
Public Function writeMetadata(ByVal srcMetadataFile As String, ByVal dstImageFile As String, ByRef srcPDImage As pdImage, Optional ByVal removeGPS As Boolean = False) As String
    
    'Many ExifTool options are delimited by quotation marks (").  Because VB has the worst character escaping scheme ever conceived, I use
    ' a variable to hold the ASCII equivalent of a quotation mark.  This makes things slightly more readable.
    Dim Quotes As String
    Quotes = Chr(34)
    
    'Grab the ExifTool path, which we will shell and pipe in a moment
    Dim appLocation As String
    appLocation = g_PluginPath & "exiftool.exe"
    
    'Build a command-line string that ExifTool can understand
    Dim cmdParams As String
    
    'Ignore minor errors and warnings (note that exiftool.exe must be included as param [0], per C convention)
    cmdParams = cmdParams & "exiftool.exe -m"
        
    'Overwrite the original destination file, but only if the metadata was embedded succesfully
    cmdParams = cmdParams & " -overwrite_original"
    
    'Copy all tags
    cmdParams = cmdParams & " -tagsfromfile " & Quotes & srcMetadataFile & Quotes & " " & Quotes & dstImageFile & Quotes
    
    'Regardless of the type of metadata copy we're performing, we need to alter or remove some tags because their
    ' original values are no longer relevant.
    cmdParams = cmdParams & " --Orientation"
    cmdParams = cmdParams & " --IFD2:ImageWidth --IFD2:ImageHeight"
    cmdParams = cmdParams & " -ImageWidth=" & srcPDImage.Width & " -ExifIFD:ExifImageWidth=" & srcPDImage.Width
    cmdParams = cmdParams & " -ImageHeight=" & srcPDImage.Height & " -ExifIFD:ExifImageHeight=" & srcPDImage.Height
    
    'If we were asked to remove GPS data, do so now
    If removeGPS Then cmdParams = cmdParams & " -gps:all="
    
    'NOTE: while in the IDE, it can be useful to see ExifTool's output, so the console window will be displayed.
    
    'Use ExifTool to write the metadata
    If Not ShellExecuteCapture(appLocation, cmdParams, writeMetadata, Not g_IsProgramCompiled) Then
        Message "Failed to write metadata."
        writeMetadata = ""
    End If
    
End Function

'Capture output from the requested command-line executable and return it as a string
' NOTE: This function is a heavily modified version of code originally written by Joacim Andersson.  A download link to his original
'        version is available at the top of this module.
Public Function ShellExecuteCapture(ByVal sApplicationPath As String, sCommandLineParams As String, ByRef dstString As String, Optional bShowWindow As Boolean = False) As Boolean

    Dim hPipeRead As Long, hPipeWrite As Long
    Dim hCurProcess As Long
    Dim sa As SECURITY_ATTRIBUTES
    Dim si As STARTUPINFO
    Dim PI As PROCESS_INFORMATION
    Dim baOutput() As Byte
    Dim sNewOutput As String
    Dim lBytesRead As Long
    
    Dim lRet As Long

    'This pipe buffer size is effectively arbitrary, but I haven't had any problems with 1024
    Const BUFSIZE = 1024

    dstString = ""
    
    ReDim baOutput(BUFSIZE - 1) As Byte

    With sa
        .nLength = Len(sa)
        .bInheritHandle = 1
    End With

    If CreatePipe(hPipeRead, hPipeWrite, sa, BUFSIZE) = 0 Then
        ShellExecuteCapture = False
        Message "Failed to start plugin service (couldn't create pipe)."
        Exit Function
    End If

    hCurProcess = GetCurrentProcess()

    'Replace the inheritable read handle with a non-inheritable one. (MSDN suggestion)
    DuplicateHandle hCurProcess, hPipeRead, hCurProcess, hPipeRead, 0&, 0&, DUPLICATE_SAME_ACCESS Or DUPLICATE_CLOSE_SOURCE

    With si
        .cb = Len(si)
        .dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
        .hStdOutput = hPipeWrite
        
        'NOTE: calling functions typically request that the shelled window be shown in the IDE but not in the compiled .exe
        If bShowWindow Then .wShowWindow = SW_NORMAL Else .wShowWindow = SW_HIDE
        
    End With
    'MsgBox sApplicationPath & vbCrLf & sCommandLineParams
    If CreateProcess(sApplicationPath, sCommandLineParams, ByVal 0&, ByVal 0&, 1, 0&, ByVal 0&, vbNullString, si, PI) Then

        'Close the thread handle, as we have no use for it
        CloseHandle PI.hThread

        'Also close the pipe's write handle.  This is important, because ReadFile will not return until all write handles
        ' are closed or the buffer is full.
        CloseHandle hPipeWrite
        hPipeWrite = 0
        
        Do
            
            If ReadFile(hPipeRead, baOutput(0), BUFSIZE, lBytesRead, ByVal 0&) = 0 Then Exit Do
            
            sNewOutput = StrConv(baOutput, vbUnicode)
            dstString = dstString & Left$(sNewOutput, lBytesRead)
            
        Loop

        CloseHandle PI.hProcess
        CloseHandle hCurProcess
    Else
        ShellExecuteCapture = False
        Message "Failed to start plugin service (couldn't create process: %1).", Err.LastDllError
        Exit Function
    End If

    CloseHandle hPipeRead
    If hPipeWrite Then CloseHandle hPipeWrite
    If hCurProcess Then CloseHandle hCurProcess
    
    ShellExecuteCapture = True
    
End Function

