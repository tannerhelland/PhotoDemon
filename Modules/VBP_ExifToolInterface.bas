Attribute VB_Name = "Plugin_ExifTool_Interface"
'***************************************************************************
'ExifTool Plugin Interface
'Copyright ©2012-2013 by Tanner Helland
'Created: 24/May/13
'Last updated: 24/May/13
'Last update: initial build
'
'Module for handling all ExifTool interfacing.  This module is pointless without the accompanying ExifTool plugin,
' which can be found in the App/PhotoDemon/Plugins subdirectory as "exiftool.exe".  The ExifTool plugin will be
' available by default in all versions of PhotoDemon after and including 5.6 (release TBD, estimate as summer 2013).
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
'...as well as a modified piping function, derived from code originally downloaded from this link:
'
'http://www.visualbasic.happycodings.com/Graphics_Games_Programming/code3.html
'
'This project was designed against v9.29 of ExifTool (18 May '13).  It may not work with other versions of the
' software.  Additional documentation regarding the use of ExifTool is available as part of the official ExifTool
' package, downloadable from http://www.sno.phy.queensu.ca/~phil/exiftool/Image-ExifTool-9.29.tar.gz
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'A number of API functions are required to pipe stdout
Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As Any, ByVal nSize As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function GetNamedPipeInfo Lib "kernel32" (ByVal hNamedPipe As Long, lType As Long, lLenOutBuf As Long, lLenInBuf As Long, lMaxInstances As Long) As Long
   
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
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
    dwProcessID As Long
    dwThreadID As Long
End Type

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As Any, lpProcessInformation As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private hWritePipe As Long
Private hReadPipe As Long

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
        exifPath = g_PluginPath & "exiftool.exe -ver"
        
        Dim outputString As String
        If ShellExecuteCapture(exifPath, outputString) Then
        
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

'Capture output from the requested command-line executable and return it as a string
Private Function ShellExecuteCapture(ByRef sCommandLine As String, ByRef sReceiveOutput As String, Optional bShowWindow As Boolean = False) As Boolean
    
    Const clReadBytes As Long = 256, INFINITE As Long = &HFFFFFFFF
    Const STARTF_USESHOWWINDOW = &H1, STARTF_USESTDHANDLES = &H100&
    Const SW_HIDE = 0, SW_NORMAL = 1
    Const NORMAL_PRIORITY_CLASS = &H20&
    
    Const PIPE_CLIENT_END = &H0     'The handle refers to the client end of a named pipe instance. This is the default.
    Const PIPE_SERVER_END = &H1     'The handle refers to the server end of a named pipe instance. If this value is not specified, the handle refers to the client end of a named pipe instance.
    Const PIPE_TYPE_BYTE = &H0      'The named pipe is a byte pipe. This is the default.
    Const PIPE_TYPE_MESSAGE = &H4   'The named pipe is a message pipe. If this value is not specified, the pipe is a byte pipe
    
    Dim tProcInfo As PROCESS_INFORMATION, lRetVal As Long, lSuccess As Long
    Dim tStartupInf As STARTUPINFO
    Dim tSecurAttrib As SECURITY_ATTRIBUTES, lhwndReadPipe As Long, lhwndWritePipe As Long
    Dim lBytesRead As Long, sBuffer As String
    Dim lPipeOutLen As Long, lPipeInLen As Long, lMaxInst As Long
    
    tSecurAttrib.nLength = Len(tSecurAttrib)
    tSecurAttrib.bInheritHandle = 1&
    tSecurAttrib.lpSecurityDescriptor = 0&

    lRetVal = CreatePipe(lhwndReadPipe, lhwndWritePipe, tSecurAttrib, 0)
    If lRetVal = 0 Then
        ShellExecuteCapture = False
        Message "Failed to start plugin service (couldn't create pipe)."
        Exit Function
    End If

    tStartupInf.cb = Len(tStartupInf)
    tStartupInf.dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
    tStartupInf.hStdOutput = lhwndWritePipe
    
    'Show or hide the command-line window as requested
    If bShowWindow Then tStartupInf.wShowWindow = SW_NORMAL Else tStartupInf.wShowWindow = SW_HIDE
    
    lRetVal = CreateProcessA(0&, sCommandLine, tSecurAttrib, tSecurAttrib, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, tStartupInf, tProcInfo)
    If lRetVal <> 1 Then
        ShellExecuteCapture = False
        Message "Failed to start plugin service (couldn't create process)."
        Exit Function
    End If
    
    'Process created, wait for completion. Note, this will cause your application
    'to hang indefinitely until the process completes.
    WaitForSingleObject tProcInfo.hProcess, INFINITE
    
    'Determine pipe's contents
    lSuccess = GetNamedPipeInfo(lhwndReadPipe, PIPE_TYPE_BYTE, lPipeOutLen, lPipeInLen, lMaxInst)
    
    If lSuccess Then
        
        'Got pipe info, create buffer
        sBuffer = String(lPipeOutLen, 0)
        
        'Read Output Pipe
        lSuccess = ReadFile(lhwndReadPipe, sBuffer, lPipeOutLen, lBytesRead, 0&)
        
        'Pipe read successfully
        If lSuccess = 1 Then sReceiveOutput = Left$(sBuffer, lBytesRead)
        
        ShellExecuteCapture = True

    Else
        Message "Failed to retrieve plugin output (couldn't find named pipe)."
        ShellExecuteCapture = False
    End If
    
    'Close handles
    Call CloseHandle(tProcInfo.hProcess)
    Call CloseHandle(tProcInfo.hThread)
    Call CloseHandle(lhwndReadPipe)
    Call CloseHandle(lhwndWritePipe)
    
End Function
