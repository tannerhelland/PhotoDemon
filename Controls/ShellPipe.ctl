VERSION 5.00
Begin VB.UserControl ShellPipe 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   CanGetFocus     =   0   'False
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   360
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   HitBehavior     =   0  'None
   InvisibleAtRuntime=   -1  'True
   PaletteMode     =   4  'None
   Picture         =   "ShellPipe.ctx":0000
   ScaleHeight     =   360
   ScaleWidth      =   360
   ToolboxBitmap   =   "ShellPipe.ctx":01A2
   Windowless      =   -1  'True
   Begin VB.Timer tmrCheck 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   -240
      Top             =   0
   End
End
Attribute VB_Name = "ShellPipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Note: this file has been modified for use within PhotoDemon.

'This code was originally written by vbforums user "dilettante".

'You may download the original version of this code from the following link (good as of November '13):
' http://www.vbforums.com/showthread.php?660014-VB6-ShellPipe-quot-Shell-with-I-O-Redirection-quot-control

'Many thanks to dilettante for this excellent user control, which allows PhotoDemon to asynchronously interact
' with the ExifTool plugin, thus greatly improving responsiveness and performance of metadata handling.


Option Explicit
'
'ShellPipe (version 4)
'=========
'
'Run a console program and communicate with it via the
'Standard I/O streams.
'
'Both StdOut and StdErr are piped to one stream reader here.
'
'NOTES
'-----
'
'Because overlapped I/O isn't available under Win9x, this
'control uses a Timer control and a "polling" model to
'process pipe reads and writes and detect process termination.
'
'Requires SmartBuffer class.
'
'ENUMS
'-----
'
'SP_RESULTS
'        SP_SUCCESS = 0
'        SP_CREATEPIPEFAILED = &H80042B00
'        SP_CREATEPROCFAILED = &H80042B01
'
'SPEOF_TYPES
'        SPEOF_NORMAL = 0
'        SPEOF_BROKEN_PIPE = 109
'
'
'PROPERTIES
'----------
'
'Active  Boolean, read-only.
'
'        Returns a result telling whether or not the child
'        process is currently active.
'
'        If NOT, then FinishChild() should be called as
'        soon as possible.
'
'HasLine Boolean, read-only.
'
'        True if we have a buffered "line" from the child
'        process buffered.
'
'        Set appropriately by every call to GetData() and
'        GetLine() as well as by DataArrival events so it
'        is possible to loop on this property to retrieve
'        lines of text.
'
'Length  Long, read-only.
'
'        Number of characters currently buffered from the
'        child process.
'
'PollInterval Long, read/write.  Default: 50.
'
'        This is a value in milliseconds used to:
'
'        * Look for data or EOF from the child process'
'          OutPipe.
'        * Send pending blocked output to the child
'          process' InPipe.
'        * Check for process termination.
'
'WaitForIdle Long, read/write.  Default: 200.
'
'        This is a value in milliseconds used to wait for
'        the child process to finish initialization.  No
'        waiting takes place for Console child processes.
'
'METHODS
'-------
'
'ClosePipe()
'
'        Closes our output pipe to the child process' StdIn.
'
'FinishChild(Optional ByVal WaitMs As Long = 500, _
'            Optional ByVal KillReturnCode = 666) As Long
'
'        MUST be called after child process terminates
'        (ChildFinished event is raised), or may be called
'        to FORCE (kill) the process.
'
'        Waits WaitMs milliseconds for the child process to
'        complete.  If the child doesn't finish, terminates
'        the child process with KillReturnCode.  Caller
'        may check for KillReturnCode value to determine
'        that the process was killed.
'
'        Closes pipes and cleans up the process.
'
'        Returns:  Child process' return code.  Returns -1
'                  if the child has already been "finished."
'
'GetData(Optional ByVal MaxLen As Long = -1) As String
'
'        Get data from child process' OutPipe.
'
'        Returns MaxLen characters (or as many as are
'        available).  When MaxLen is -1 returns all
'        available characters.  May return an empty string.
'
'GetLine() As String
'
'        Get a line of data from child process' OutPipe.
'
'        Should only be called when HasLine is True.
'        May return an empty string.
'
'        A "line" is defined as text delimited by a CR, but
'        if CRLF occurs the LFs are consumed as well.  Both
'        conventions are used by StdIO programs.
'
'Interrupt(Optional ByVal Break As Boolean = False)
'
'        Attempts to interrupt the child process.  This is
'        only effective if the parent has a console window,
'        which will be inherited by the child.  Note that
'        the parent will also be interrupted, so a null
'        handler must be installed or else the parent will
'        also be terminated!
'
'        Break sends a CTRL-C signal when False or a
'        CTRL-Break signal when True.
'
'Run    (ByVal CommandLine As String, _
'        Optional ByVal CurrentDir As String = vbNullString) _
'        As SP_RESULTS
'
'        Runs the command line provided via CommandLine with
'        standard streams redirected to the ShellPipe control.
'        If the caller doesn't supply a CurrentDir string,
'        the child process inherits the caller's current
'        directory.
'
'        Returns:  SP_SUCCESS
'                  SP_CREATEPIPEFAILED
'                  SP_CREATEPROCFAILED
'
'SendData(ByVal Data As String)
'
'        Pipes Data to child process' StdIn.
'
'SendLine(ByVal Line As String,
'         Optional ByVal UseLFs As Boolean = True)
'
'        Pipes Line and CR or CRLF to child process' StdIn.
'
'EVENTS
'------
'
'ChildFinished()
'
'        Signals that the child process has completed.  The
'        program should call the FinishChild() method as
'        soon as possible to clean up process handles and
'        obtain the child process' return code.
'
'DataArrival(ByVal CharsTotal As Long)
'
'        Signals that input data from the child process'
'        OutPipe is available to be read via GetData().
'
'        CharsTotal is the amount of data available to be
'        read.
'
'EOF    (ByVal EOFType As SPEOF_TYPES)
'
'        Signals that an EOF or BROKEN_PIPE has occurred
'        on the child process' OutPipe.
'
'        EOFType:  SPEOF_NORMAL
'                  SPEOF_BROKEN_PIPE
'
'Error  (ByVal Number As Long, _
'        ByVal Source As String, _
'        ByRef CancelDisplay As Boolean)
'
'        Signals that some sort of error occurred
'        while performing an operation.
'
'        Number is the error number, typically a pipe or
'        other system error.
'
'        Source is a string describing the source of the
'        error, generally some operation internal to
'        ShellPipe.
'
'        CancelDisplay indicates whether to cancel the
'        display. The default is False, which is to display
'        the default error message box. If you do not want
'        to use the default message box, set CancelDisplay
'        to True.
'
'EXCEPTIONS
'----------
'
'&H80042B02 in ShellPipe.PollInterval
'
'        PollInterval value supplied is outside the valid
'        range 10 to 65535.
'

Private Const WIN32NULL As Long = 0&
Private Const WIN32FALSE As Long = 0&
Private Const WIN32TRUE As Long = 1&
Private Const WAIT_OBJECT_0 As Long = 0&
Private Const NORMAL_PRIORITY_CLASS As Long = &H20&
Private Const INFINITE As Long = -1&
Private Const STARTF_USESHOWWINDOW As Long = &H1&
Private Const STARTF_USESTDHANDLES As Long = &H100&
Private Const SW_HIDE As Long = 0&
Private Const STD_INPUT_HANDLE As Long = -10&
Private Const STD_OUTPUT_HANDLE As Long = -11&
Private Const STD_ERROR_HANDLE As Long = -12&
Private Const HANDLE_FLAG_INHERIT As Long = &H1&
Private Const CTRL_C_EVENT As Long = 0&
Private Const CTRL_BREAK_EVENT As Long = 1&
Private Const ERROR_ACCESS_DENIED As Long = 5&
Private Const ERROR_INVALID_HANDLE As Long = 6&
Private Const ERROR_HANDLE_EOF As Long = 38&
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

Private Declare Function CloseHandle Lib "kernel32" _
    (ByVal hObject As Long) As Long

Private Declare Function CreatePipe Lib "kernel32" _
    (ByRef phReadPipe As Long, _
     ByRef phWritePipe As Long, _
     ByRef lpPipeAttributes As Any, _
     ByVal nSize As Long) As Long

Private Declare Function CreateProcessA Lib "kernel32" _
    (ByVal lpApplicationName As String, _
     ByVal lpCommandLine As String, _
     ByVal lpProcessAttributes As Long, _
     ByVal lpThreadAttributes As Long, _
     ByVal bInheritHandles As Long, _
     ByVal dwCreationFlags As Long, _
     ByVal lpEnvironment As Long, _
     ByVal lpCurrentDirectory As String, _
     ByRef lpStartupInfo As STARTUPINFO, _
     ByRef lpProcessInformation As PROCESSINFO) As Long

Private Declare Function GenerateConsoleCtrlEvent Lib "kernel32" _
    (ByVal dwCtrlEvent As Long, _
     ByVal dwProcessGroupId As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" _
    (ByVal hHandle As Long, _
     ByRef lpExitCode As Long) As Long

Private Declare Function PeekNamedPipe Lib "kernel32" _
    (ByVal hNamedPipe As Long, _
     ByVal lpBuf As String, _
     ByVal nBufSize As Long, _
     ByRef lpBytesRead As Long, _
     ByRef lpTotalBytesAvail As Long, _
     ByVal lpBytesLeftThisMessage As Long) As Long

Private Declare Function ReadFile Lib "kernel32" _
    (ByVal hFile As Long, _
     ByVal lpBuf As String, _
     ByVal nNumberOfBytesToRead As Long, _
     ByRef lpNumberOfBytesRead As Long, _
     ByVal lpOverlapped As Any) As Long

Private Declare Function SetHandleInformation Lib "kernel32" _
    (ByVal hObject As Long, _
     ByVal dwMask As Long, _
     ByVal dwFlags As Long) As Long

Private Declare Function TerminateProcess Lib "kernel32" _
    (ByVal hHandle As Long, _
     ByVal uExitCode As Long) As Long

Private Declare Function WaitForInputIdle Lib "user32" ( _
    ByVal hProcess As Long, _
    ByVal dwMilliseconds As Long) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" _
    (ByVal hHandle As Long, _
     ByVal dwMilliseconds As Long) As Long

Private Declare Function WriteFile Lib "kernel32" _
    (ByVal hFile As Long, _
     ByVal lpBuf As String, _
     ByVal cToWrite As Long, _
     ByRef cWritten As Long, _
     ByVal lpOverlapped As Any) As Long

Private piProc As PROCESSINFO
Private saPipe As SECURITY_ATTRIBUTES
Private hChildInPipeRd As Long
Private hChildInPipeWr As Long
Private hChildOutPipeRd As Long
Private hChildOutPipeWr As Long
Private blnFinishedChild As Boolean
Private blnProcessActive As Boolean
Private blnInPipeOpen As Boolean
Private lngWaitForIdle As Long
Private sbInBuffer As SmartBuffer
Private sbOutBuffer As SmartBuffer

Public Enum SPEOF_TYPES
    SPEOF_NORMAL = 0
    SPEOF_BROKEN_PIPE = ERROR_BROKEN_PIPE
End Enum

Public Enum SP_RESULTS
    SP_SUCCESS = 0
    SP_CREATEPIPEFAILED = &H80042B00
    SP_CREATEPROCFAILED = &H80042B01
End Enum

Public Event DataArrival(ByVal CharsTotal As Long)
Public Event EOF(ByVal EOFType As SPEOF_TYPES)
Public Event Error(ByVal Number As Long, _
                   ByVal Source As String, _
                   CancelDisplay As Boolean)
Public Event ChildFinished()

Public Property Get Active() As Boolean
    If blnProcessActive Then 'Last we knew, it was active.
        If WaitForSingleObject(piProc.hProcess, 0&) <> WAIT_OBJECT_0 Then
            Active = True
        Else
            blnProcessActive = False
            Active = False
        End If
    Else
        Active = False
    End If
End Property

Public Sub ClosePipe()
    CloseHandle hChildInPipeWr
    blnInPipeOpen = False
End Sub

Public Function FinishChild(Optional ByVal WaitMs As Long = 500, _
                            Optional ByVal KillReturnCode = 666) As Long
    If blnFinishedChild Then
        FinishChild = -1 'Already "Finished Child."
    Else
        If blnProcessActive Then
            If WaitForSingleObject(piProc.hProcess, WaitMs) <> WAIT_OBJECT_0 Then
                TerminateProcess piProc.hProcess, KillReturnCode
                'WaitForSingleObject piProc.hProcess, INFINITE
            End If
            blnProcessActive = False
            tmrCheck.Enabled = False
        End If
        
        sbInBuffer.Flush
        sbOutBuffer.Flush
        GetExitCodeProcess piProc.hProcess, FinishChild
        CloseHandle hChildOutPipeRd
        If blnInPipeOpen Then ClosePipe
        CloseHandle piProc.hThread
        CloseHandle piProc.hProcess
        blnFinishedChild = True
    End If
End Function

Public Function GetData(Optional ByVal MaxLen As Long = -1) As String
    GetData = sbInBuffer.GetData(MaxLen)
End Function

Public Function GetLine() As String
    GetLine = sbInBuffer.GetLine()
End Function

Public Property Get HasLine() As Boolean
    HasLine = sbInBuffer.HasLine
End Property

Public Sub Interrupt(Optional ByVal Break As Boolean = False)
    Dim lngEvent As Long
    Dim lngError As Long
    Dim blnCancel As Boolean
    
    lngEvent = IIf(Break, CTRL_BREAK_EVENT, CTRL_C_EVENT)
    If GenerateConsoleCtrlEvent(lngEvent, 0&) = 0 Then
        lngError = Err.LastDllError
        RaiseEvent Error(lngError, "ShellPipe.Interrupt.ConsoleCtrlEvent", blnCancel)
        If Not blnCancel Then
            Err.Raise lngError, TypeName(Me), "Interrupt ConsoleCtrlEvent error"
        End If
    End If
End Sub

Public Property Get Length() As Long
    Length = sbInBuffer.Length
End Property

Public Property Get PollInterval() As Long
    PollInterval = tmrCheck.Interval
End Property

Public Property Let PollInterval(ByVal RHS As Long)
    If 5 > RHS Or RHS > 65535 Then
        Err.Raise &H80042B02, TypeName(Me), "PollInterval outside valid range 5-65535"
    End If
    tmrCheck.Interval = RHS
    PropertyChanged "PollInterval"
End Property

Public Function Run(ByVal CommandLine As String, Optional ByVal CommandLineParams As String = "", Optional ByVal CurrentDir As String = vbNullString) As SP_RESULTS
    
    Dim siStart As STARTUPINFO
    
    With saPipe
        .nLength = Len(saPipe)
        .lpSecurityDescriptor = WIN32NULL
        .bInheritHandle = WIN32TRUE
    End With
    
    If CreatePipe(hChildOutPipeRd, hChildOutPipeWr, saPipe, 0&) = WIN32FALSE Then
        Run = SP_CREATEPIPEFAILED
        Exit Function
    End If
    SetHandleInformation hChildOutPipeRd, HANDLE_FLAG_INHERIT, 0&
    
    If CreatePipe(hChildInPipeRd, hChildInPipeWr, saPipe, 0&) = WIN32FALSE Then
        CloseHandle hChildOutPipeRd
        CloseHandle hChildOutPipeWr
        Run = SP_CREATEPIPEFAILED
        Exit Function
    End If
    SetHandleInformation hChildInPipeWr, HANDLE_FLAG_INHERIT, 0&
    
    With siStart
        .cb = Len(siStart)
        .dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
        .wShowWindow = SW_HIDE
        .hStdOutput = hChildOutPipeWr
        .hStdError = hChildOutPipeWr
        .hStdInput = hChildInPipeRd
        'Leave other fields 0/Null.
    End With
    
    'Clear all fields, global UDT and we may have been used more than once.
    With piProc
        .hProcess = 0
        .hThread = 0
        .dwProcessID = 0
        .dwThreadID = 0
    End With
    
    If CreateProcessA(CommandLine, CommandLineParams, WIN32NULL, WIN32NULL, WIN32TRUE, _
                      NORMAL_PRIORITY_CLASS, WIN32NULL, CurrentDir, _
                      siStart, piProc) = WIN32FALSE Then
        blnProcessActive = False
        Run = SP_CREATEPROCFAILED
    Else
        CloseHandle hChildOutPipeWr
        CloseHandle hChildInPipeRd
        blnProcessActive = True
        blnFinishedChild = False
        blnInPipeOpen = True
        If WaitForIdle > 0 Then WaitForInputIdle piProc.hProcess, WaitForIdle
        tmrCheck.Enabled = True
        Run = SP_SUCCESS
    End If
End Function

Public Sub SendData(ByVal Data As String)
    sbOutBuffer.Append Data
    WriteData
End Sub

Public Sub SendLine(ByVal Line As String, Optional ByVal UseLFs As Boolean = True)
    If UseLFs Then
        SendData Line & vbCrLf
    Else
        SendData Line & vbCr
    End If
End Sub

Public Property Get WaitForIdle() As Long
    WaitForIdle = lngWaitForIdle
End Property

Public Property Let WaitForIdle(ByVal RHS As Long)
    If RHS < 0 Then Err.Raise &H80042B03, TypeName(Me), "WaitForIdle must be >= 0"
    lngWaitForIdle = RHS
    PropertyChanged "WaitForIdle"
End Property

Private Sub tmrCheck_Timer()
    If Active Then
        ReadData
        DoEvents 'Let events raised in ReadData be handled.
        WriteData
    Else
        'Last gasp.
        ReadData
        DoEvents 'Let events raised in ReadData be handled.

        tmrCheck.Enabled = False
        RaiseEvent ChildFinished
    End If
End Sub

Private Sub ReadData()
    Dim strBuf As String
    Dim lngAvail As Long
    Dim lngRead As Long
    Dim lngError As Long
    Dim blnCancel As Boolean
    
    If PeekNamedPipe(hChildOutPipeRd, WIN32NULL, 0&, WIN32NULL, lngAvail, WIN32NULL) <> WIN32FALSE Then
        If lngAvail > 0 Then
            strBuf = String$(lngAvail, 0)
            If ReadFile(hChildOutPipeRd, ByVal strBuf, lngAvail, lngRead, WIN32NULL) <> WIN32FALSE Then
                If lngRead > 0 Then
                    sbInBuffer.Append Left$(strBuf, lngRead)
                    RaiseEvent DataArrival(sbInBuffer.Length)
                Else
                    RaiseEvent EOF(SPEOF_NORMAL)
                End If
            Else
                lngError = Err.LastDllError
                If lngError = ERROR_BROKEN_PIPE Then
                    RaiseEvent EOF(SPEOF_BROKEN_PIPE)
                Else
                    RaiseEvent Error(lngError, "ShellPipe.ReadData.ReadFile", blnCancel)
                    If Not blnCancel Then
                        Err.Raise lngError, TypeName(Me), "ReadData ReadFile error"
                    End If
                End If
            End If
        End If
    Else
        lngError = Err.LastDllError
        Select Case lngError
            Case ERROR_BROKEN_PIPE
                RaiseEvent EOF(SPEOF_BROKEN_PIPE)
                
            Case ERROR_ACCESS_DENIED, ERROR_INVALID_HANDLE
                'Ignore as "no input."
                
            Case Else
                RaiseEvent Error(lngError, "ShellPipe.ReadData.PeekNamedPipe", blnCancel)
                If Not blnCancel Then
                    Err.Raise TypeName(Me), "ReadData PeeknamedPipe error"
                End If
        End Select
    End If
End Sub

Private Sub WriteData()
    Dim strBuffer As String
    Dim lngWritten As Long
    Dim lngError As Long
    Dim blnCancel As Boolean
    
    If blnInPipeOpen Then
        If sbOutBuffer.Length > 0 Then
            sbOutBuffer.PeekBuffer strBuffer
            If WriteFile(hChildInPipeWr, ByVal strBuffer, Len(strBuffer), lngWritten, 0&) <> WIN32FALSE Then
                sbOutBuffer.DeleteData lngWritten
            Else
                lngError = Err.LastDllError
                RaiseEvent Error(lngError, "ShellPipe.WriteData.WriteFile", blnCancel)
                If Not blnCancel Then
                    'NOTE FROM TANNER: we don't care about write errors in PD, so just ignore any that may rise
                    'Err.Raise lngError, TypeName(Me), "WriteData WriteFile error"
                End If
            End If
        End If
    Else
        sbOutBuffer.Flush
    End If
End Sub

Private Sub UserControl_Initialize()
    blnFinishedChild = True
    Set sbInBuffer = New SmartBuffer
    Set sbOutBuffer = New SmartBuffer
End Sub

Private Sub UserControl_InitProperties()
    PollInterval = 50
    WaitForIdle = 200
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    PollInterval = PropBag.ReadProperty("PollInterval", 50)
    WaitForIdle = PropBag.ReadProperty("WaitForIdle", 200)
End Sub

Private Sub UserControl_Resize()
    Height = 360
    Width = 360
End Sub

Private Sub UserControl_Terminate()
    FinishChild
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "PollInterval", PollInterval, 50
    PropBag.WriteProperty "WaitForIdle", WaitForIdle, 200
End Sub
