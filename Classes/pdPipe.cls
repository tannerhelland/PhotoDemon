VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdPipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'Simple (Named) Pipe class
'Copyright 2020-2026 by Tanner Helland
'Created: 14/September/20
'Last updated: 14/September/20
'Last update: initial build
'
'This class is currently used to implement single-session behavior in PD.  If the user has set
' this preference, parallel PD sessions will detect the original session, and forward their
' command-line parameters via a named pipe implemented through this class.
'
'When this class references MSDN, it typically refers to the API name in question, or to the
' article on named pipes (link good as of September 2020):
' https://docs.microsoft.com/en-us/windows/win32/ipc/named-pipes
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'If TRUE, writes additional pipe creation notifications to the debug log.
' (Turn off in production code!)
Private Const PDPIPE_DEBUG_VERBOSE As Boolean = False

'The type of raised event depends on the pipe's mode (currently only byte mode is implemented).
Public Event BytesArrived(ByVal initStreamPosition As Long, ByVal numOfBytes As Long)

Public Enum PD_PipeOpenMode
    
    'One of these modes *MUST* be specified or CreateNamedPipe will fail.  (MSDN: The function fails if dwOpenMode specifies anything other than 0 or the flags listed in the following tables.  This parameter must specify one of the following pipe access modes. The same mode must be specified for each instance of the pipe.)
    pom_BiDirectional = &H3&   'The pipe is bi-directional; both server and client processes can read from and write to the pipe. This mode gives the server the equivalent of GENERIC_READ and GENERIC_WRITE access to the pipe. The client can specify GENERIC_READ or GENERIC_WRITE, or both, when it connects to the pipe using the CreateFile function.
    pom_ClientToServer = &H1&  'The flow of data in the pipe goes from client to server only. This mode gives the server the equivalent of GENERIC_READ access to the pipe. The client must specify GENERIC_WRITE access when connecting to the pipe. If the client must read pipe settings by calling the GetNamedPipeInfo or GetNamedPipeHandleState functions, the client must specify GENERIC_WRITE and FILE_READ_ATTRIBUTES access when connecting to the pipe.
    pom_ServerToClient = &H2&  'The flow of data in the pipe goes from server to client only. This mode gives the server the equivalent of GENERIC_WRITE access to the pipe. The client must specify GENERIC_READ access when connecting to the pipe. If the client must change pipe settings by calling the SetNamedPipeHandleState function, the client must specify GENERIC_READ and FILE_WRITE_ATTRIBUTES access when connecting to the pipe.
    
    'These parameters are *OPTIONAL*.  (MSDN: This parameter can also include one or more of the following flags, which enable the write-through and overlapped modes. These modes can be different for different instances of the same pipe.)
    pom_FlagFirstPipeInstance = &H80000 'If you attempt to create multiple instances of a pipe with this flag, creation of the first instance succeeds, but creation of the next instance fails with ERROR_ACCESS_DENIED.
    pom_FlagWriteThrough = &H80000000   'Write-through mode is enabled. This mode affects only write operations on byte-type pipes and, then, only when the client and server processes are on different computers. If this mode is enabled, functions writing to a named pipe do not return until the data written is transmitted across the network and is in the pipe's buffer on the remote computer. If this mode is not enabled, the system enhances the efficiency of network operations by buffering data until a minimum number of bytes accumulate or until a maximum time elapses.
    pom_FlagOverlapped = &H40000000     'Overlapped mode is enabled. If this mode is enabled, functions performing read, write, and connect operations that may take a significant time to be completed can return immediately. This mode enables the thread that started the operation to perform other operations while the time-consuming operation executes in the background. For example, in overlapped mode, a thread can handle simultaneous input and output (I/O) operations on multiple instances of a pipe or perform simultaneous read and write operations on the same pipe handle. If overlapped mode is not enabled, functions performing read, write, and connect operations on the pipe handle do not return until the operation is finished. The ReadFileEx and WriteFileEx functions can only be used with a pipe handle in overlapped mode. The ReadFile, WriteFile, ConnectNamedPipe, and TransactNamedPipe functions can execute either synchronously or as overlapped operations.
    
    'These security flags are *OPTIONAL*. (MSDN: This parameter can include any combination of the following security access modes. These modes can be different for different instances of the same pipe.)
    pom_SecurityWriteDAC = &H40000               'The caller will have write access to the named pipe's discretionary access control list (ACL).
    pom_SecurityWriteOwner = &H80000             'The caller will have write access to the named pipe's owner.
    pom_SecurityAccessSystemSecurity = &H1000000 'The caller will have write access to the named pipe's SACL. For more information, see Access-Control Lists (ACLs) and SACL Access Right.
    
End Enum

#If False Then
    Private Const pom_BiDirectional = &H3&, pom_ClientToServer = &H1&, pom_ServerToClient = &H2&
    Private Const pom_FlagFirstPipeInstance = &H80000, pom_FlagWriteThrough = &H80000000, pom_FlagOverlapped = &H40000000
    Private Const pom_SecurityWriteDAC = &H40000, pom_SecurityWriteOwner = &H80000, pom_SecurityAccessSystemSecurity = &H1000000
#End If

Public Enum PD_PipeMode

    'The function fails if dwPipeMode specifies anything other than 0 or the flags listed in the following tables.
    'One of the following type modes can be specified. The same type mode must be specified for each instance of the pipe.
    pm_TypeByte = &H0&    'Data is written to the pipe as a stream of bytes. This mode cannot be used with PIPE_READMODE_MESSAGE. The pipe does not distinguish bytes written during different write operations.
    pm_TypeMessage = &H4& 'Data is written to the pipe as a stream of messages. The pipe treats the bytes written during each write operation as a message unit. The GetLastError function returns ERROR_MORE_DATA when a message is not read completely. This mode can be used with either PIPE_READMODE_MESSAGE or PIPE_READMODE_BYTE.
    
    'One of the following read modes can be specified. Different instances of the same pipe can specify different read modes.
    pm_ReadModeByte = &H0&      'Data is read from the pipe as a stream of bytes. This mode can be used with either PIPE_TYPE_MESSAGE or PIPE_TYPE_BYTE.
    pm_ReadModeMessage = &H2&   'Data is read from the pipe as a stream of messages. This mode can be only used if PIPE_TYPE_MESSAGE is also specified.
    
    'One of the following wait modes can be specified. Different instances of the same pipe can specify different wait modes.
    pm_ModeWait = &H0&      'Blocking mode is enabled. When the pipe handle is specified in the ReadFile, WriteFile, or ConnectNamedPipe function, the operations are not completed until there is data to read, all data is written, or a client is connected. Use of this mode can mean waiting indefinitely in some situations for a client process to perform an action.
    pm_ModeDontWait = &H1&  'Nonblocking mode is enabled. In this mode, ReadFile, WriteFile, and ConnectNamedPipe always return immediately.
    'IMPORTANT: Note that nonblocking mode is supported for compatibility with Microsoft LAN Manager version 2.0 and should not be used to achieve asynchronous I/O with named pipes. For more information on asynchronous pipe I/O, see Synchronous and Overlapped Input and Output.
    
    'One of the following remote-client modes can be specified. Different instances of the same pipe can specify different remote-client modes.
    pm_RemoteClientsAccept = &H0&   'Connections from remote clients can be accepted and checked against the security descriptor for the pipe.
    pm_RemoteClientsReject = &H8&   'Connections from remote clients are automatically rejected.
    
End Enum

#If False Then
    Private Const pm_TypeByte = &H0&, pm_TypeMessage = &H4&
    Private Const pm_ReadModeByte = &H0&, pm_ReadModeMessage = &H2&
    Private Const pm_ModeWait = &H0&, pm_ModeDontWait = &H1&
    Private Const pm_RemoteClientsAccept = &H0&, pm_RemoteClientsReject = &H8&
#End If

Private Const INVALID_HANDLE_VALUE As Long = -1

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'There is no -W CreateNamedPipe, so we use -A wrappers for CreateFile as well
Private Declare Function CreateFileA Lib "kernel32" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CreateNamedPipeA Lib "kernel32" (ByVal lpName As String, ByVal dwOpenMode As PD_PipeOpenMode, ByVal dwPipeMode As Long, ByVal nMaxInstances As Long, ByVal nOutBufferSize As Long, ByVal nInBufferSize As Long, ByVal nDefaultTimeOut As Long, ByVal lpSecurityAttributes As Long) As Long
Private Declare Function DisconnectNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long) As Long
'Private Declare Function GetNamedPipeInfo Lib "kernel32" (ByVal hNamedPipe As Long, ByRef lpFlags As Long, ByRef outBufferSize As Long, ByRef inBufferSize As Long, ByRef nMaxInstances As Long) As Long
Private Declare Function PeekNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long, ByVal lpDstBuffer As Long, ByVal nBufSize As Long, ByVal ptrToOutBytes As Long, ByRef lpTotalBytesAvail As Long, ByVal lpBytesLeftThisMessage As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuf As Long, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
'Private Declare Function WaitNamedPipeA Lib "kernel32" (ByVal lpName As String, ByVal nTimeOut As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal ptrToSourceBuffer As Long, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal ptrToOverlappedStruct As Long) As Long

'non-zero if a pipe is active
Private m_Pipe As Long, m_PipeName As String

'Type of pipe; we use different functions to assess pipe state depending on pipe type
Private m_PipeIsInByteMode As Boolean

'Waiting for pipe responses is currently handled by timer
Private WithEvents m_Timer As pdTimer
Attribute m_Timer.VB_VarHelpID = -1

'Here's the cool thing about doing this inside PD - instead of forcing the caller to manually
' manage a buffer, we instead ask them to pass an instantiated pdStream object to us.  They can
' back the stream however they want (file, memory-mapped file, just memory, whatever), and we'll
' fill the stream with data as it arrives, then raise events to notify the caller of the new data.
' It's an elegant solution to a complex problem.
Private m_Stream As pdStream

'Create a pipe with the given name.
' pipeName: must be passed.  Any Windows-required prefixes are handled automatically, so don't worry about adding them yourself.
Friend Function CreatePipe(ByRef pipeName As String, ByRef instantiatedDstStream As pdStream, ByVal pipeOpenMode As PD_PipeOpenMode, ByVal pipeMode As PD_PipeMode, Optional ByVal maxInstances As Long = 1, Optional ByVal inputBufferSize As Long = 1024, Optional ByVal outputBufferSize As Long = 1024) As Boolean
    
    If (m_Pipe <> 0) Then Me.ClosePipe
    
    m_PipeName = AppendPipeNameBits(pipeName)
    
    'Store pipe mode; this affects some read/write functions
    m_PipeIsInByteMode = ((pipeMode And pm_TypeByte) = pm_TypeByte)
    
    'Basic validation on some parameters
    If (maxInstances < 1) Then maxInstances = 1
    If (maxInstances > 255) Then maxInstances = 255
    
    'Attempt to create the pipe
    m_Pipe = CreateNamedPipeA(m_PipeName, pipeOpenMode, pipeMode, maxInstances, outputBufferSize, inputBufferSize, 0&, 0&)
    CreatePipe = (m_Pipe <> 0) And (m_Pipe <> INVALID_HANDLE_VALUE)
    If CreatePipe Then
        
        If PDPIPE_DEBUG_VERBOSE Then PDDebug.LogAction "Successfully created pipe: " & m_Pipe & ", " & Err.LastDllError
        
        'Point our stream reference at the caller's stream; all pipe data will be placed there
        ' as it arrives.
        Set m_Stream = instantiatedDstStream
        
    Else
        PDDebug.LogAction "WARNING!  pdPipe.CreatePipe failed to create pipe: " & m_PipeName & "; lastDLLError: " & Err.LastDllError
    End If
    
End Function

'Connect to an already existing pipe instance.
' IMPORTANTLY, you need to ensure that *all* pipe details (not just name) match the initial pipe settings.
' This is a Microsoft requirement, not a PD one.  Mismatched settings can result in connection failure
' (or worse).
'
'pipeName: must be passed.  Any Windows-required prefixes are handled automatically, so don't worry about adding them yourself.
Friend Function ConnectToExistingPipe(ByRef pipeName As String, Optional ByVal needWriteAccess As Boolean = True, Optional ByVal needReadAccess As Boolean = True, Optional ByVal waitUntilAvailable As Boolean = True) As Boolean

    m_PipeName = AppendPipeNameBits(pipeName)
    
    'Determine connection flags based on access requirements
    Const GENERIC_READ As Long = &H80000000
    Const GENERIC_WRITE As Long = &H40000000
    Dim cfFlags As Long
    If needReadAccess Then cfFlags = cfFlags Or GENERIC_READ
    If needWriteAccess Then cfFlags = cfFlags Or GENERIC_WRITE
    
    'Attempt to connect
    Const OPEN_EXISTING As Long = 3, FILE_ATTRIBUTE_NORMAL As Long = &H80&
    m_Pipe = CreateFileA(m_PipeName, cfFlags, 0&, 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0&)
    ConnectToExistingPipe = (m_Pipe <> 0) And (m_Pipe <> INVALID_HANDLE_VALUE)
    
    If (Not ConnectToExistingPipe) Then PDDebug.LogAction "WARNING!  pdPipe.ConnectToExistingPipe failed to connect to pipe: " & pipeName & "; lastDLLError: " & Err.LastDllError
    
    'TODO: implement wait for pipe via WaitNamedPipeA?
    
End Function

'Server instances only: disconnect this pipe instance from a connected (or recently connected
' but now closed) pipe client.
Friend Function DisconnectFromClient() As Boolean
    If (m_Pipe <> 0) Then DisconnectFromClient = (DisconnectNamedPipe(m_Pipe) <> 0)
End Function

'Once connected to a pipe, call this to pass binary data.
Friend Function WriteDataToPipe(ByVal ptrToData As Long, ByVal sizeOfData As Long) As Boolean
    
    If (m_Pipe <> 0) Then
        
        Dim numBytesWritten As Long
        WriteDataToPipe = (WriteFile(m_Pipe, ptrToData, sizeOfData, numBytesWritten, 0&) <> 0)
        
        If WriteDataToPipe Then
            
            'Ensure written byte count was the same
            Dim totalBytesWritten As Long
            totalBytesWritten = numBytesWritten
            
            Do While (totalBytesWritten < sizeOfData)
                
                'Write again from the current offset
                numBytesWritten = 0
                WriteDataToPipe = (WriteFile(m_Pipe, ptrToData + totalBytesWritten, sizeOfData - totalBytesWritten, numBytesWritten, 0&) <> 0)
                
                'Quit on failure
                If (Not WriteDataToPipe) Then
                    PDDebug.LogAction "WARNING!  pdPipe.WriteDataToPipe failed to write additional data; " & totalBytesWritten & " of " & sizeOfData & " bytes written"
                    Exit Function
                End If
                
                'Update counts and offsets
                totalBytesWritten = totalBytesWritten + numBytesWritten
                
            Loop
            
            'TODO: add an option to force a flush op here
            
        Else
            PDDebug.LogAction "WARNING!  pdPipe.WriteDataToPipe failed to write."
        End If
        
    Else
        PDDebug.LogAction "WARNING!  pdPipe.WriteDataToPipe can't write to a pipe that doesn't exist!"
    End If
    
End Function

'Once a pipe has been created, call this function to start waiting for a response.
Friend Sub Server_WaitForResponse(Optional ByVal waitIntervalInMS As Long = 1000)
    
    'Start the wait timer
    Set m_Timer = New pdTimer
    m_Timer.Interval = waitIntervalInMS
    m_Timer.StartTimer
    
End Sub

'Manually shutdown our pipe (if one exists)
Friend Sub ClosePipe()
    
    Set m_Stream = Nothing
    Set m_Timer = Nothing
    
    If (m_Pipe <> 0) Then
        If (CloseHandle(m_Pipe) = 0) Then PDDebug.LogAction "WARNING!  pdPipe failed to close its handle; last error was " & Err.LastDllError()
        m_Pipe = 0
        m_PipeName = vbNullString
    End If
    
End Sub

'Convert a bare pipe name to a Windows-formatted local pipe name (with Win 10 workarounds)
Private Function AppendPipeNameBits(ByRef srcName As String) As String
    
    'Per MSDN:
    ' "The entire pipe name string can be up to 256 characters long. Pipe names are not case-sensitive."
    ' "The pipe server cannot create a pipe on another computer, so CreateNamedPipe must use a period for
    ' the server name, as shown in the following example: \\.\pipe\PipeName"
    ' "Windows 10, version 1709:  Pipes are only supported within an app-container; ie, from one
    ' UWP process to another UWP process that's part of the same app. Also, named pipes must use
    ' the syntax "\.\pipe\LOCAL" for the pipe name."
    AppendPipeNameBits = "\\.\pipe\"
    If (OS.GetWin10Build() >= 16299) Or (Not OS.IsProgramCompiled()) Then AppendPipeNameBits = AppendPipeNameBits & "LOCAL\"
    AppendPipeNameBits = AppendPipeNameBits & srcName
    
End Function

Private Sub Class_Terminate()
    ClosePipe
End Sub

Private Sub m_Timer_Timer()
    
    'In the future, handling may need to vary by pipe mode... but for now, we use the same
    ' approach for both message- and byte-based pipes
    Dim numBytesAvailable As Long, peekResult As Long
    peekResult = PeekNamedPipe(m_Pipe, 0&, 0&, 0&, numBytesAvailable, 0&)
    
    'Peek at the pipe
    If (peekResult <> 0) Then
    
        'Anything for us?
        If (numBytesAvailable > 0) Then
            
            Dim initStreamPosition As Long
            
            'Prep the stream to receive the new data
            m_Stream.EnsureBufferSpaceAvailable numBytesAvailable
            initStreamPosition = m_Stream.GetPosition()
            
            Dim numBytesRead As Long
            If (ReadFile(m_Pipe, m_Stream.Peek_PointerOnly(peekLength:=numBytesAvailable), numBytesAvailable, numBytesRead, 0&) <> 0) Then
                
                'Ensure retrieved byte amount was accurate
                If (numBytesAvailable <> numBytesRead) Then PDDebug.LogAction "WARNING!  pdPipe.Timer got a weird result for bytes read."
                
                'Manually increase the stream's size to account for the write,
                ' and move the pointer accordingly.
                m_Stream.SetSizeExternally initStreamPosition + numBytesRead
                m_Stream.SetPosition numBytesRead, FILE_CURRENT
                
                'Notify the caller of the update
                RaiseEvent BytesArrived(initStreamPosition, numBytesRead)
                
            Else
                PDDebug.LogAction "WARNING! pdPipe.Timer failed to read."
            End If
            
        End If
    
    Else
        
        Dim errLast As Long
        errLast = Err.LastDllError
        
        If (errLast = 230) Then
            'No clients connected yet; this is fine, and 99% of the time it's the expected result.
        ElseIf (errLast = 109) Then
            'A client *was* connected, but isn't now; this is also fine, and will occur after
            ' a second (or third, etc) instance connects, then disconnects
        Else
            PDDebug.LogAction "WARNING! pdPipe.Timer failed to peek pipe: " & m_Pipe & ", lastDllError: " & Err.LastDllError
        End If
        
    End If
    
End Sub
