VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdPipeAsync"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Async ExifTool Interface
'Copyright 2017-2026 by Tanner Helland
'Created: 24/October/17
'Last updated: 08/September/23
'Last update: finally write code for non-UTF-8 mode(s)
'
'This class manages PD's async connection to the 3rd-party ExifTool library.  Communication is handled over
' stdin/out/err pipes, and special modifications are in place to support a number of ExifTool-specific quirks.
'
'This class was originally implemented as a heavily modified version of the "ShellPipe" project,
' by vbforums user dilettante.  dilettante's very nice project is a much better choice for a general-purpose
' pipe implementation, and you can find his original project here (link good as of Oct 2017):
' http://www.vbforums.com/showthread.php?660014-VB6-ShellPipe-quot-Shell-with-I-O-Redirection-quot-control
'
'dilettante's tool was based around sending and receiving strings, but with ExifTool, this assumption doesn't
' work because we need to interop via UTF-8 bytes (so we can support not just Unicode-aware filenames, but
' Unicode-aware image metadata, too).  This was one of several differing requirements that required me to
' extensively hack up the original UserControl, and at the end of the day, what I really need is not a
' UserControl, but a dedicated class that can live inside the ExifTool module (instead of being sited on a
' random form).
'
'So in October 2017, I fully rewrote PD's ExifTool interface as this standalone class.  Everything here is now
' ultra-specific to ExifTool interop, with all the good and bad that entails.  (For example, events are now
' handled internally, because the ExifTool interface is module-based so we can't easily raise events.)
'
'A lot of unsafe hackery is also in place to minimize memory copies.  Images - particularly in RAW formats -
' can carry metadata that's multiple MB in size, so we don't want to be transferring that to/from VB strings
' any more than absolutely necessary.  As such, we do weird things like pass peeked pointers from pdStream
' objects directly to ReadFile, to bypass the need for a temporary copy of incoming data.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Const NORMAL_PRIORITY_CLASS As Long = &H20&
Private Const STARTF_USESHOWWINDOW As Long = &H1&
Private Const STARTF_USESTDHANDLES As Long = &H100&
Private Const SW_HIDE As Long = 0&
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

Private m_ChildProcessActive As Boolean

'pdStream is integral to this class; it grants a ton of performance advantages, especially when doing unsafe reading/writing
' from pointers returned by APIs.
Private m_BufferIn As pdStream
Private m_BufferOut As pdStream
Private m_BufferErr As pdStream

'Edit by Tanner:
' I trade UTF-8 data with ExifTool, which (obviously) requires some special interop pieces.  As a failsafe against horribly
' breaking this class, I've implemented my changes using these helper variables.
Private m_AssumeUTF8Input As Boolean
Private m_AssumeUTF8Output As Boolean

'Edit by Tanner: use a system timer object instead of a VB one, so we don't require a full underlying usercontrol
Private WithEvents m_Async As pdTimer
Attribute m_Async.VB_VarHelpID = -1

Private Sub ClosePipeIn()
    CloseHandleSafely m_PipeInWrite
End Sub

Private Sub ClosePipeOut()
    CloseHandleSafely m_PipeOutRead
End Sub

Private Sub ClosePipeErr()
    CloseHandleSafely m_PipeErrRead
End Sub

Private Sub CloseHandleSafely(ByRef srcHandle As Long)
    If (srcHandle <> 0) Then
        If (CloseHandle(srcHandle) <> 0) Then srcHandle = 0 Else InternalError "CloseHandle failed on handle #" & srcHandle
    End If
End Sub

Friend Function GetSizeOfInputBuffer() As Long
    GetSizeOfInputBuffer = m_BufferIn.GetStreamSize()
End Function

Friend Sub ResetInputBuffer()
    m_BufferIn.SetPosition 0
    m_BufferIn.SetSizeExternally 1024
End Sub

Friend Function TerminateChildProcess() As Boolean
    
    If (m_ProcInfo.hProcess <> 0) Then
    
        Dim dstExitCode As Long
        
        'Terminate the child process
        If m_ChildProcessActive Then
            TerminateProcess m_ProcInfo.hProcess, dstExitCode
            m_ChildProcessActive = False
            m_Async.StopTimer
        End If
        
        'Clear all buffers and close all pipes
        m_BufferIn.StopStream True
        m_BufferOut.StopStream True
        m_BufferErr.StopStream True
        
        ClosePipeIn
        ClosePipeOut
        ClosePipeErr
        
        'Release any remaining handles
        If (m_ProcInfo.hThread <> 0) Then
            CloseHandle m_ProcInfo.hThread
            m_ProcInfo.hThread = 0
        End If
        
        If (m_ProcInfo.hProcess <> 0) Then
            CloseHandle m_ProcInfo.hProcess
            m_ProcInfo.hProcess = 0
        End If
        
    End If
    
    TerminateChildProcess = True
    
End Function

Friend Function GetDataAsString() As String
    
    If (m_BufferIn.GetStreamSize > 0) Then
    
        'Translate the requested amount of data from our internal buffer into a usable string.
        If m_AssumeUTF8Input Then
            GetDataAsString = Strings.StringFromUTF8Ptr(m_BufferIn.Peek_PointerOnly(0), m_BufferIn.GetStreamSize())
            
        'Same, but without UTF-8 translation
        Else
            GetDataAsString = Strings.StringFromCharPtr(m_BufferIn.Peek_PointerOnly(0), False, m_BufferIn.GetStreamSize(), True)
        End If
        
        'Remove the processed bytes from the stream
        m_BufferIn.DeleteFromStart m_BufferIn.GetStreamSize()
        
    End If
    
End Function

'See comments for GetDataAsString(), above, if you're curious about how this works
Friend Function ErrGetDataAsString() As String
    If (m_BufferErr.GetStreamSize > 0) Then
        If m_AssumeUTF8Input Then
            ErrGetDataAsString = Strings.StringFromUTF8Ptr(m_BufferErr.Peek_PointerOnly(0), m_BufferErr.GetStreamSize())
        Else
            ErrGetDataAsString = Strings.StringFromCharPtr(m_BufferErr.Peek_PointerOnly(0), False, m_BufferErr.GetStreamSize(), True)
        End If
        m_BufferErr.DeleteFromStart m_BufferErr.GetStreamSize()
    End If
End Function

'With ExifTool, we often need to search for a flag like {ready123} that indicates ExifTool has gone idle.  The ExifTool module
' uses this function to peek at the last few bytes in the buffer without actually removing any data.  Note that the requested
' number of bytes may translate to a different number of chars due to UTF-8 conversion.  (Similarly, if there aren't enough
' requested bytes in the buffer, we'll obviously just return as many as we can.)
Friend Function PeekLastNBytes(Optional ByVal maxNumBytes As Long = 16&) As String
    
    'Calculate an offset into the source buffer
    Dim ptrOffset As Long
    ptrOffset = m_BufferIn.GetPosition() - maxNumBytes
    If (ptrOffset < 0) Then
        maxNumBytes = maxNumBytes + ptrOffset
        ptrOffset = 0
    End If
    
    'Retrieve as many bytes as we can
    If (maxNumBytes > 0) Then PeekLastNBytes = Strings.StringFromUTF8Ptr(m_BufferIn.Peek_PointerOnly(ptrOffset), maxNumBytes) Else PeekLastNBytes = vbNullString
    
End Function

Friend Function PeekPointer(Optional ByVal ptrPosition As Long = 0) As Long
    PeekPointer = m_BufferIn.Peek_PointerOnly(ptrPosition)
End Function

'NOTE FROM TANNER: I have modified this function to better work with ExifTool.  Specifically, I have separated out the command line
'                   and command line params into two separate strings, which are then passed SEPARATELY to CreateProcess.  Because
'                   ExifTool requests can require many command line parameters, this helps us avoid MAX_PATH limitations for the
'                   whole command line + params string, and it also makes it easier to deal with spaces in the path name.
Friend Function Run(ByVal childCmdLine As String, Optional ByVal childCmdLineParams As String = vbNullString) As Boolean
    
    'Before doing anything else, we need to prep pipe security attributes.  Default settings are used,
    ' and we explicitly make the handle inheritable.
    With m_PipeSecurityAttributes
        .nLength = Len(m_PipeSecurityAttributes)
        .lpSecurityDescriptor = 0&
        .bInheritHandle = 1&
    End With
    
    'ExifTool may send extremely large chunks of data (1+ MB are possible in a normal session).
    ' To improve performance, we ask for a particularly large chunk size.  Note that the size request
    ' may not be respected by Windows; for details, see
    ' https://msdn.microsoft.com/en-us/library/windows/desktop/aa365152(v=vs.85).aspx
    Dim readPipeSize As Long
    readPipeSize = 131072
    
    'Create an stdout pipe
    If (CreatePipe(m_PipeOutRead, m_PipeOutWrite, m_PipeSecurityAttributes, readPipeSize) = 0) Then
        InternalError "CreatePipe failed for stdout", Err.LastDllError
    End If
    
    'stderr uses default allocation sizes
    If (CreatePipe(m_PipeErrRead, m_PipeErrWrite, m_PipeSecurityAttributes, 0&) = 0) Then
        InternalError "CreatePipe failed for stderr", Err.LastDllError
    End If
    
    'stdin uses default allocation sizes
    If (CreatePipe(m_PipeInRead, m_PipeInWrite, m_PipeSecurityAttributes, 0&) = 0) Then
        InternalError "CreatePipe failed for stdin", Err.LastDllError
    End If
    
    'Make sure all pipe handles are valid
    If (m_PipeOutRead = 0) Or (m_PipeOutWrite = 0) Or (m_PipeErrRead = 0) Or (m_PipeErrWrite = 0) Or (m_PipeInRead = 0) Or (m_PipeInWrite = 0) Then
        
        InternalError "Because one more handle creation(s) failed, pdPipeAsync will not attempt to start the child process."
        Run = False
        
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
        .wShowWindow = SW_HIDE
        .hStdOutput = m_PipeOutWrite
        .hStdError = m_PipeErrWrite
        .hStdInput = m_PipeInRead
    End With
    
    'PD only ever starts the child process once, so we don't need to clean up m_ProcInfo here.
    
    'Per PD requirements, we must use the Unicode-friendly CreateProcess variety, to ensure Unicode paths
    ' are supported properly.  Note that we once again declare inheritable handles.
    If (CreateProcessW(StrPtr(childCmdLine), StrPtr(childCmdLineParams), 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, VarPtr(siStart), m_ProcInfo) = 0) Then
        
        'ExifTool failed to start.  Dump some (hopefully?) useful debug data before exiting.
        InternalError "CreateProcessW failed with err#" & Err.LastDllError & ".  Relevant paths may include: "
        InternalError "childCmdLine: " & childCmdLine
        InternalError "childCmdLineParams: " & childCmdLineParams
        
        m_ChildProcessActive = False
        Run = False
    
    Else
        
        'Success!  Close the handles inherited by the child process (as we no longer need them).
        CloseHandleSafely m_PipeOutWrite
        CloseHandleSafely m_PipeErrWrite
        CloseHandleSafely m_PipeInRead
        m_ChildProcessActive = True
        
        'Start listening for return data
        If PDMain.IsProgramRunning Then m_Async.StartTimer
        
        Run = True
        
    End If
    
End Function

'Send string data to the child process.
Friend Sub SendData(ByRef srcData As String)
    
    If (LenB(srcData) <> 0) Then
        
        'If UTF-8 mode is active, convert incoming strings to UTF-8 before appending them to the buffer
        Dim byteBuffer() As Byte, writeLen As Long
        If m_AssumeUTF8Output Then
            Strings.UTF8FromString srcData, byteBuffer, writeLen
        Else
            byteBuffer = StrConv(srcData, vbFromUnicode)
            writeLen = UBound(byteBuffer) + 1
        End If
        
        'If we actually appended usable data to the out stream, submit it to the child process now
        If (writeLen <> 0) Then
            m_BufferOut.WriteBytesFromPointer VarPtr(byteBuffer(0)), writeLen
            WriteData
        End If
        
    End If
    
End Sub

Private Sub m_Async_Timer()
    
    If m_ChildProcessActive Then
        ReadData
        WriteData
    
    'If the child process is no longer active, perform one last read from the pipe, then terminate this timer
    Else
        ReadData
        m_Async.StopTimer
    End If
    
End Sub

Private Sub ReadData()
    
    'The stdout and stderr pipes are handled in turn, using basically identical code.
    If (ReadSpecificPipe(m_PipeOutRead, m_BufferIn) > 0) Then InternalDataArrival
    If (ReadSpecificPipe(m_PipeErrRead, m_BufferErr) > 0) Then InternalErrDataArrival
    
End Sub

'Read data from a specific pipe, and place it in the requested buffer.  Returns a positive integer if data was placed in the buffer,
' representing the number of bytes placed.  Returns 0 if nothing was placed in the buffer, and returns a negative integer if
' something catastrophic happened.
Private Function ReadSpecificPipe(ByRef srcPipeHandle As Long, ByRef dstBuffer As pdStream) As Long

    'Ensure a pipe exists...
    If (srcPipeHandle <> 0) Then
        
        'If our peek or read operation fails, we'll look for a broken pipe, then shut down our connection accordingly
        Dim failState As Long
        
        'See if the pipe has any bytes for us
        Dim numBytesAvailable As Long, numBytesRead As Long
        If (PeekNamedPipe(srcPipeHandle, 0&, 0&, ByVal 0&, numBytesAvailable, 0&) <> 0) Then
            If (numBytesAvailable > 0) Then
                
                'Grab as many bytes as we can!  (The +4 is arbitrary - it's just to ensure a few extra bytes, "just in case".)
                ' Also - because we're clever, notice how we dump the bytes directly into our buffer, then advance the
                ' pointer manually.
                dstBuffer.EnsureBufferSpaceAvailable numBytesAvailable + 4
                If (ReadFile(srcPipeHandle, dstBuffer.Peek_PointerOnly(), numBytesAvailable, numBytesRead, 0&) = 0) Then
                    
                    failState = Err.LastDllError
                    InternalError "ReadFile (inside .ReadData()) failed with API error #" & failState
                    
                    'Return an arbitrary negative number to indicate failure
                    ReadSpecificPipe = -2
                    
                Else
                    
                    'Make sure we received more than 0 bytes before modifying the stream object
                    If (numBytesRead > 0) Then
                    
                        'Advance the stream size and pointer manually.
                        dstBuffer.SetSizeExternally dstBuffer.GetStreamSize + numBytesRead
                        dstBuffer.SetPosition numBytesRead, FILE_CURRENT
                        ReadSpecificPipe = numBytesRead
                        
                        'Handle the data arrival immediately
                        InternalDataArrival
                        
                    'If we received zero bytes, the child process is done.  A "data finished" notification could be raised
                    ' in the future, if we want.
                    Else
                        CloseHandleSafely srcPipeHandle
                        ReadSpecificPipe = 0
                    End If
                    
                End If
                
            'End bytesAvailable > 0
            End If
            
        'Pipe peek failed
        Else
            failState = Err.LastDllError
            InternalError "PeekNamedPipe (inside .ReadData()) failed with API error #" & failState
            ReadSpecificPipe = -2
            
        'End pipe peek
        End If
        
        'Look for potential failure states we can handle intelligently
        
        'Broken pipes are a terminal issue; close our handle so we stop looking for data
        If (failState = ERROR_BROKEN_PIPE) Then
            InternalError "Target pipe is broken.  Shutting down our end of the connection."
            CloseHandleSafely srcPipeHandle
            ReadSpecificPipe = -2
        End If
    
    End If

End Function

'Internal function, used only if...
' 1) A caller submitted usable data via the SendData function
' 2) A past SendData call wasn't able to submit its full data queue
Private Sub WriteData()
    
    'Make sure the pipe exists, and that we have a non-zero amount of bytes to write
    If (m_PipeInWrite <> 0) And (Not m_BufferOut Is Nothing) Then
        If (m_BufferOut.GetStreamSize > 0) Then
        
            'Try to submit the entire remaining data stream
            Dim writeSuccess As Boolean, numBytesWritten As Long
            writeSuccess = (WriteFile(m_PipeInWrite, m_BufferOut.Peek_PointerOnly(0), m_BufferOut.GetStreamSize, numBytesWritten, 0&) <> 0)
            
            'If we were able to submit at least *some* of our data, remove it from the submission queue
            If writeSuccess Then
                m_BufferOut.DeleteFromStart numBytesWritten
            
            'If the write failed, do not update any internal buffers (so we can try again later)
            Else
                InternalError "WARNING! pdAsync write failure; is the child process still running?", Err.LastDllError
            End If
        End If
    
    'As a failsafe, clear the buffer if a write attempt is made but our pipe has died
    Else
        If (Not m_BufferOut Is Nothing) Then
            If (m_BufferOut.GetStreamSize > 0) Then m_BufferOut.StopStream True
        End If
    End If
    
End Sub

Private Sub InternalError(ByVal errMessage As String, Optional ByVal errNumber As Long = 0)
    'Async errors are not uncommon, especially if the user is running a bunch of PD sessions in parallel.
    ' To avoid overloading debug logs with reports, we simply report error-number-based messages
    ' to the debug window.
    If (errNumber <> 0) Then
        Debug.Print "WARNING!  pdPipeAsync error # " & errNumber & ": " & errMessage
    Else
        PDDebug.LogAction "WARNING!  Unspecified pdPipeAsync error: " & errMessage
    End If
End Sub

'Ideally, a class like this would raise meaningful events, but in PD, we only use pipe interactions with
' one plugin - ExifTool.  Our ExifTool interface is currently sited in a module, which means it can't
' handle events; as such, we currently handle pipe operations internally.
Private Sub InternalDataArrival()
    
    If (m_BufferIn.GetStreamSize() > 0) Then
        PDDebug.LogAction "Asynchronously received " & m_BufferIn.GetStreamSize() & " new characters from ExifTool."
        ExifTool.NewMetadataReceived
    End If
    
End Sub

Private Sub InternalErrDataArrival()
    PDDebug.LogAction "FYI: pdPipeAsync's source returned the following data on stderr: "
    PDDebug.LogAction Me.ErrGetDataAsString()
End Sub

Private Sub Class_Initialize()
    
    'API timer is used to "fake" async
    Set m_Async = New pdTimer
    m_Async.Interval = 34
    
    Set m_BufferIn = New pdStream
    m_BufferIn.StartStream PD_SM_MemoryBacked, PD_SA_ReadWrite
    
    Set m_BufferOut = New pdStream
    m_BufferOut.StartStream PD_SM_MemoryBacked, PD_SA_ReadWrite
    
    Set m_BufferErr = New pdStream
    m_BufferErr.StartStream PD_SM_MemoryBacked, PD_SA_ReadWrite
    
    'Edit by Tanner: we use UTF-8 interop with Exiftool; this behavior may break interactions with other software,
    ' which is why I haven't passed the suggestion upstream.
    m_AssumeUTF8Input = True
    m_AssumeUTF8Output = True
    
End Sub

Private Sub Class_Terminate()
    TerminateChildProcess
End Sub
