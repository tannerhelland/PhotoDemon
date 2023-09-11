VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdPipeSync"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Synchronous Pipe Interface
'Copyright 2017-2023 by Tanner Helland
'Created: 24/October/17
'Last updated: 08/September/23
'Last update: split out from pdPipeAsync, and simplify for synchronous use only (this is preferable for some plugins)
'
'This class provides a simplified way to shell external apps, wait until they finish, then retrieve any
' data they sent to stdread/write/err.
'
'This class was originally split off from pdPipeAsync; look there for additional implementation details,
' including the use of timers for asynchronous interactions with a shelled app.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Const NORMAL_PRIORITY_CLASS As Long = &H20&
Private Const STARTF_USESHOWWINDOW As Long = &H1&
Private Const STARTF_USESTDHANDLES As Long = &H100&
Private Const SW_HIDE As Long = 0&, SW_NORMAL As Long = 1&
Private Const ERROR_BROKEN_PIPE As Long = 109&

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

Private Type PROCESSINFO
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreatePipe Lib "kernel32" (ByRef phReadPipe As Long, ByRef phWritePipe As Long, ByRef lpPipeAttributes As Any, ByVal nSize As Long) As Long
Private Declare Function CreateProcessW Lib "kernel32" (ByVal ptrToApplicationName As Long, ByVal ptrToCommandLine As Long, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, ByVal ptrToStartupInfo As Long, ByRef lpProcessInformation As PROCESSINFO) As Long
Private Declare Function PeekNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long, ByVal lpDstBuffer As Long, ByVal nBufSize As Long, ByRef lpBytesRead As Long, ByRef lpTotalBytesAvail As Long, ByVal lpBytesLeftThisMessage As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuf As Long, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function SetHandleInformation Lib "kernel32" (ByVal hObject As Long, ByVal dwMask As Long, ByVal dwFlags As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hHandle As Long, ByVal uExitCode As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuf As Long, ByVal cToWrite As Long, ByRef cWritten As Long, ByVal lpOverlapped As Any) As Long

Private m_ProcInfo As PROCESSINFO
Private m_PipeSecurityAttributes As SECURITY_ATTRIBUTES

Private m_PipeInRead As Long, m_PipeInWrite As Long
Private m_PipeOutRead As Long, m_PipeOutWrite As Long
Private m_PipeErrRead As Long, m_PipeErrWrite As Long

'pdStream is integral to this class; it grants a ton of performance advantages, especially when doing unsafe reading/writing
' from pointers returned by APIs.
Private m_BufferOut As pdStream
Private m_BufferErr As pdStream

'Track child process state so that we can safely free it in case of error
Private m_ChildProcessActive As Boolean

'Some modern apps can trade UTF-8 data over pipes.  This class works in both modes, but is ANSI by default.
' Call the SetUTF8Mode function to enable.
Private m_UseUTF8 As Boolean

Private Sub CloseHandleSafely(ByRef srcHandle As Long)
    Const FUNC_NAME As String = "CloseHandleSafely"
    If (srcHandle <> 0) Then
        If (CloseHandle(srcHandle) <> 0) Then srcHandle = 0 Else InternalError FUNC_NAME, "CloseHandle failed on handle #" & srcHandle
    End If
End Sub

Friend Function TerminateChildProcess() As Boolean
    
    If (m_ProcInfo.hProcess <> 0) Then
    
        Dim dstExitCode As Long
        
        'Terminate the child process
        If m_ChildProcessActive Then
            TerminateProcess m_ProcInfo.hProcess, dstExitCode
            m_ProcInfo.hProcess = 0
            m_ChildProcessActive = False
        End If
        
        'Clear all buffers and close all pipes
        m_BufferOut.StopStream True
        m_BufferErr.StopStream True
        
        CloseHandleSafely m_PipeOutRead
        CloseHandleSafely m_PipeErrRead
        
        'Release any remaining handles
        If (m_ProcInfo.hThread <> 0) Then CloseHandleSafely m_ProcInfo.hThread
        If (m_ProcInfo.hProcess <> 0) Then CloseHandleSafely m_ProcInfo.hProcess
        
    End If
    
    TerminateChildProcess = True
    
End Function

Friend Function GetStdOutDataAsString() As String
    
    If (m_BufferOut.GetStreamSize > 0) Then
    
        m_BufferOut.SetPosition 0, FILE_BEGIN
        
        'Translate the requested amount of data from our internal buffer into a usable string.
        If m_UseUTF8 Then
            GetStdOutDataAsString = m_BufferOut.ReadString_UTF8(m_BufferOut.GetStreamSize(), False)
            
        'Same, but without UTF-8 translation
        Else
            GetStdOutDataAsString = m_BufferOut.ReadString_ASCII(m_BufferOut.GetStreamSize())
        End If
        
    End If
    
End Function

'See comments for GetDataAsString(), above, if you're curious about how this works
Friend Function GetStdErrDataAsString() As String
    If (m_BufferErr.GetStreamSize > 0) Then
        m_BufferErr.SetPosition 0, FILE_BEGIN
        If m_UseUTF8 Then
            GetStdErrDataAsString = m_BufferErr.ReadString_UTF8(m_BufferErr.GetStreamSize(), False)
        Else
            GetStdErrDataAsString = m_BufferErr.ReadString_ASCII(m_BufferErr.GetStreamSize())
        End If
    End If
End Function

'Run the target application, then hold until it completes.
' After this function returns, you can call the various GET-prefixed functions to retrieve data from individual std-pipes.
Friend Function RunAndCaptureOutput(ByVal childCmdLine As String, Optional ByVal childCmdLineParams As String = vbNullString, Optional bShowWindow As Boolean = False) As Boolean
    
    Const FUNC_NAME As String = "RunAndCaptureOutput"
    
    'Before doing anything else, we need to prep pipe security attributes.  Default settings are used,
    ' and we explicitly make the handle inheritable.
    With m_PipeSecurityAttributes
        .nLength = Len(m_PipeSecurityAttributes)
        .lpSecurityDescriptor = 0&
        .bInheritHandle = 1&
    End With
    
    'Ask for a normal-ish chunk size, but note that this size may not be respected by Windows; for details, see
    ' https://msdn.microsoft.com/en-us/library/windows/desktop/aa365152(v=vs.85).aspx
    Dim readPipeSize As Long
    readPipeSize = 1024
    
    'Create an stdout pipe
    If (CreatePipe(m_PipeOutRead, m_PipeOutWrite, m_PipeSecurityAttributes, readPipeSize) = 0) Then
        InternalError FUNC_NAME, "CreatePipe failed for stdout", Err.LastDllError
    End If
    
    'stderr uses default allocation sizes
    If (CreatePipe(m_PipeErrRead, m_PipeErrWrite, m_PipeSecurityAttributes, 0&) = 0) Then
        InternalError FUNC_NAME, "CreatePipe failed for stderr", Err.LastDllError
    End If
    
    'stdin uses default allocation sizes
    If (CreatePipe(m_PipeInRead, m_PipeInWrite, m_PipeSecurityAttributes, 0&) = 0) Then
        InternalError FUNC_NAME, "CreatePipe failed for stdin", Err.LastDllError
    End If
    
    'Make sure all pipe handles are valid
    If (m_PipeOutRead = 0) Or (m_PipeOutWrite = 0) Or (m_PipeErrRead = 0) Or (m_PipeErrWrite = 0) Or (m_PipeInRead = 0) Or (m_PipeInWrite = 0) Then
        
        InternalError FUNC_NAME, "Because one more handle creation(s) failed, pdPipeSync will not attempt to start the child process."
        RunAndCaptureOutput = False
        
        CloseHandleSafely m_PipeOutRead
        CloseHandleSafely m_PipeOutWrite
        CloseHandleSafely m_PipeErrRead
        CloseHandleSafely m_PipeErrWrite
        CloseHandleSafely m_PipeInRead
        CloseHandleSafely m_PipeInWrite
        
        Exit Function
        
    End If
    
    'If we're still here, our pipes were created successfully.  Convert one of each pipe-pairs to be non-inheritable
    ' (as we want access to three of the pipes, while our child process gets access to the other end of those three pipes)
    Const HANDLE_FLAG_INHERIT As Long = &H1&
    SetHandleInformation m_PipeOutRead, HANDLE_FLAG_INHERIT, 0&
    SetHandleInformation m_PipeErrRead, HANDLE_FLAG_INHERIT, 0&
    SetHandleInformation m_PipeInWrite, HANDLE_FLAG_INHERIT, 0&
    
    'With all pipes read, we now need to prep startup objects for the child process.  Note how our constructed
    ' pipe handles are constructed - remember that the child process gets the *opposite* ends of each pipe.
    Dim siStart As STARTUPINFO
    With siStart
        .cb = Len(siStart)
        .dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
        If bShowWindow Then .wShowWindow = SW_NORMAL Else .wShowWindow = SW_HIDE
        .hStdOutput = m_PipeOutWrite
        .hStdError = m_PipeErrWrite
        .hStdInput = m_PipeInRead
    End With
    
    'PD only ever starts the child process once, so we don't need to clean up m_ProcInfo here.
    
    'Per PD requirements, we must use the Unicode-friendly CreateProcess variety, to ensure Unicode paths
    ' are supported properly.  Note that we once again declare inheritable handles.
    If (CreateProcessW(StrPtr(childCmdLine), StrPtr(childCmdLineParams), 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, VarPtr(siStart), m_ProcInfo) = 0) Then
        
        'ExifTool failed to start.  Dump some (hopefully?) useful debug data before exiting.
        InternalError FUNC_NAME, "CreateProcessW failed with err#" & Err.LastDllError & ".  Relevant paths may include: "
        InternalError FUNC_NAME, "childCmdLine: " & childCmdLine
        InternalError FUNC_NAME, "childCmdLineParams: " & childCmdLineParams
        
        m_ChildProcessActive = False
        RunAndCaptureOutput = False
    
    Else
        
        'Close the thread handle, as we have no use for it
        CloseHandleSafely m_ProcInfo.hThread
        
        'Success!  Close the handles inherited by the child process (as we no longer need them).
        CloseHandleSafely m_PipeOutWrite
        CloseHandleSafely m_PipeErrWrite
        CloseHandleSafely m_PipeInRead
        m_ChildProcessActive = True
        
        'Also close our write handle; this allows the child process to run to completion without waiting for further input
        CloseHandleSafely m_PipeInWrite
        
        'Read from stdout and stderr
        Dim numBytesRead As Long
        
        Do
            numBytesRead = 0
            m_BufferOut.EnsureBufferSpaceAvailable readPipeSize
            If (ReadFile(m_PipeOutRead, m_BufferOut.Peek_PointerOnly(, readPipeSize), readPipeSize, numBytesRead, 0&) = 0) Then
                Exit Do
            Else
                m_BufferOut.SetSizeExternally m_BufferOut.GetPosition + numBytesRead
                m_BufferOut.SetPosition numBytesRead, FILE_CURRENT
            End If
        Loop
        
        Do
            numBytesRead = 0
            m_BufferErr.EnsureBufferSpaceAvailable readPipeSize
            If (ReadFile(m_PipeErrRead, m_BufferErr.Peek_PointerOnly(, readPipeSize), readPipeSize, numBytesRead, 0&) = 0) Then
                Exit Do
            Else
                m_BufferErr.SetSizeExternally m_BufferErr.GetPosition + numBytesRead
                m_BufferErr.SetPosition numBytesRead, FILE_CURRENT
            End If
        Loop
        
        'Ensure child process terminates
        CloseHandleSafely m_ProcInfo.hProcess
        
        RunAndCaptureOutput = True
        
    End If
    
    'Regardless of success or failure, terminate all pipes and other handles safely
    CloseHandleSafely m_ProcInfo.hProcess
    CloseHandleSafely m_PipeOutWrite
    CloseHandleSafely m_PipeErrWrite
    CloseHandleSafely m_PipeInRead
    CloseHandleSafely m_PipeInWrite
    CloseHandleSafely m_PipeOutRead
    CloseHandleSafely m_PipeErrRead
    
    m_ChildProcessActive = False
    
End Function

Friend Sub SetUTF8Mode(ByVal useUTF8 As Boolean)
    m_UseUTF8 = useUTF8
End Sub

Private Sub InternalError(ByVal funcName As String, ByVal errMessage As String, Optional ByVal errNumber As Long = 0)
    PDDebug.LogAction "WARNING!  pdPipeSync error in " & funcName & ": " & errMessage
End Sub

Private Sub Class_Initialize()
    
    'Prep buffers for communication (stdout/err only!)
    Set m_BufferOut = New pdStream
    m_BufferOut.StartStream PD_SM_MemoryBacked, PD_SA_ReadWrite
    
    Set m_BufferErr = New pdStream
    m_BufferErr.StartStream PD_SM_MemoryBacked, PD_SA_ReadWrite
    
    'By default, this class assumes ANSI interop.  Call the SetUTF8Mode function to change this behavior
    ' (but be sure to TEST accordingly - most apps don't require this!)
    m_UseUTF8 = False
    
    m_ChildProcessActive = False
    
End Sub

Private Sub Class_Terminate()
    TerminateChildProcess
End Sub
