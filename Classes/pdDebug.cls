VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdDebugger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Custom Debug helper class
'Copyright ©2013-2014 by Tanner Helland
'Created: 17/July/14
'Last updated: 17/July/14
'Last update: initial build
'
'As PhotoDemon has grown more complex, debugging has become correspondingly difficult.  Debugging on my local PC is fine
' thanks to the IDE, but a lot of behaviors are different in the compiled .exe, and user PCs are another problem entirely.
'
'To that end, a more comprehensive debugging solution was required.  Enter this class.
'
'I am currently in the process of converting PD's many Debug.Print statements to pdDebug.LogAction, surrounded by
' conditional compilation "#IF DEBUGMODE = 1 // #END IF" statements.  Every time an action is logged this way, it will
' be not only mirrored to the Debug window (same as Debug.Print), it will also be written out to file in the program's
' /Data folder.  This should allow me to retrieve at least basic information from end users who experience hard crashes.
'
'While some elements of this class are PD-specific (such as where it writes its logs to file), it wouldn't take much
' work to change those bits to fit any other project.  Aside from that particular aspect, I've tried to keep the rest
' of the class as generic as possible in case this is helpful to others.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Has this instance been initialized?  This will be set to true if the InitializeDebugger function has executed successfully.
Private m_debuggerActive As Boolean

'Does the user want us writing this data to file?  If so, this will be set to TRUE.
Private m_logDatatoFile As Boolean

'ID of this debug session.  PD allows 10 unique debug logs to be kept.  Once 10 exist, the program will start overwriting
' old debug logs.  This ID value is automatically determined by checking the /Debug subfolder, and scanning for available
' logs.
Private m_debuggerID As Long

'Full path to the log file for this debug session.  This is created by the InitializeDebugger function, and it relies on
' the /Debug path specified by the pdPreferences class.  (Generally this is the /Data/Debug folder of wherever PhotoDemon.exe
' is located.)
Private m_logPath As String

'Handle to the log file.  This value will change over the life of the debugger, as we repeatedly open and close the
' log file to force VB to immediately write its data (otherwise it will wait, which is typically a good thing, but in
' the case of a debugger, we need the file kept as up-to-date as possible in case of crashes).
Private m_FileHandle As Integer

'For technical reasons (how's that for nondescript), the public instance of this class cannot be formally initialized
' until fairly late into PD's load process.  However, there are technical details that would be nice to know prior to
' that point.  To work around this, I allow the class to log messages prior to being formally initialized.  When it
' *is* finally initialized, it will dump all the cached messages to the log file, with their original timestamps.
Dim backupMessages() As String
Dim numBackupMessages As Long

'This specialty Initialize function must be called before attempting to use this class.  It will figure out where to log
' this session's data, among other things, so don't attempt to use the class until this has been called!
' Returns: TRUE if successful, FALSE otherwise.
Public Function InitializeDebugger(Optional ByVal writeLogDataToFile As Boolean = False) As Boolean
    
    If writeLogDataToFile Then
    
        'First things first: we need to make sure a Debug path exists.  Otherwise, we can't write any of our debug data to file.
        m_logPath = g_UserPreferences.getDebugPath
        
        'Make sure the path exists, and make sure we have write access.  If either of these fail, terminate the debugger.
        If Not DirectoryExist(m_logPath) Then
            Debug.Print "Log path invalid.  Debug logs not available for this session."
            m_debuggerActive = False
            InitializeDebugger = False
            Exit Function
        End If
        
        'We now know the Debug path exists.  Retrieve a relevant ID for this file.
        m_debuggerID = getLogID()
        
        'Generate a filename for this log, using that ID.
        m_logPath = m_logPath & "DebugReport_" & m_debuggerID & ".log"
        
        'If a log file already exists at that location, remove it.  (Only 10 log files are allowed, so if we reach 11,
        ' the oldest one will be overwritten.)
        Dim overwroteOldLog As Boolean
        overwroteOldLog = False
        
        If FileExist(m_logPath) Then
            Kill m_logPath
            overwroteOldLog = True
        End If
        
        'Build a new log file, and populate it with some default information.
        m_logDatatoFile = True
        
        m_FileHandle = FreeFile
        
        Open m_logPath For Append As #m_FileHandle
            Print #m_FileHandle, "**********************************************"
            Print #m_FileHandle, ""
            Print #m_FileHandle, "-- PHOTODEMON DEBUG LOG #" & m_debuggerID & " --"
            Print #m_FileHandle, ""
            Print #m_FileHandle, "Date: " & Date
            Print #m_FileHandle, "Time: " & Time
            Print #m_FileHandle, "Session ID: " & g_SessionID
            Print #m_FileHandle, "Program compiled: " & CStr(g_IsProgramCompiled)
            Print #m_FileHandle, "Log is original: " & CStr(Not overwroteOldLog)
            Print #m_FileHandle, ""
            
            Print #m_FileHandle, "-- SYSTEM INFORMATION --"
            Print #m_FileHandle, ""
            Print #m_FileHandle, "OS Version: " & OS_Interactions.getOSVersionAsString
            Print #m_FileHandle, "Processor cores (logical): " & OS_Interactions.getNumLogicalCores
            Print #m_FileHandle, "Processor features: " & OS_Interactions.getProcessorFeatures
            Print #m_FileHandle, "System RAM: " & OS_Interactions.getTotalSystemRAM
            Print #m_FileHandle, "Max memory available to PhotoDemon: " & OS_Interactions.getRAMAvailableToPD
            Print #m_FileHandle, ""
            Print #m_FileHandle, "-- PROGRAM INFORMATION -- "
            Print #m_FileHandle, ""
            Print #m_FileHandle, "Version: " & getPhotoDemonNameAndVersion
            Print #m_FileHandle, "Translations active: " & CStr(g_Language.translationActive())
            Print #m_FileHandle, "Language in use: " & CStr(g_Language.getCurrentLanguage())
            Print #m_FileHandle, "GDI+ available: " & CStr(g_GDIPlusAvailable)
            Print #m_FileHandle, "zLib enabled: " & CStr(g_ZLibEnabled)
            Print #m_FileHandle, "FreeImage enabled: " & CStr(g_ImageFormats.FreeImageEnabled)
            Print #m_FileHandle, "ExifTool enabled: " & CStr(g_ExifToolEnabled)
            Print #m_FileHandle, "EZTwain enabled: " & CStr(g_ScanEnabled)
            Print #m_FileHandle, "PngQuant enabled: " & CStr(g_ImageFormats.pngQuantEnabled)
            Print #m_FileHandle, ""
            Print #m_FileHandle, "**********************************************"
            Print #m_FileHandle, ""
            Print #m_FileHandle, "-- SESSION REPORT --"
            
        'Close the file to force an immediate write
        Close #m_FileHandle
        
    End If
    
    m_debuggerActive = True
    
    'Log an initial event, to note that debug mode was successfully initiated
    Me.LogAction "Debugger initialized successfully"
    
    'If messages were logged prior to this class being formally initialized, dump them now
    If (numBackupMessages > 0) And (writeLogDataToFile) Then
    
        Me.LogAction "(The following " & numBackupMessages & " actions were logged prior to initialization.)"
        Me.LogAction "(They are presented here with their original timestamps.)"
        
        m_FileHandle = FreeFile
        Open m_logPath For Append As #m_FileHandle
        
        Dim i As Long
        For i = 0 To numBackupMessages - 1
            Print #m_FileHandle, backupMessages(i)
        Next i
        
        Close #m_FileHandle
        
        Me.LogAction "(End of pre-initialization data)"
        
        Erase backupMessages
        numBackupMessages = 0
        
    End If
    
    InitializeDebugger = True
    
End Function

'Search the debug folder for existing debug files, sort them by date, and automatically give this log a unique ID on the
' range [0, 9].  If there are already 10 debug files present, steal the ID of the oldest file.
Private Function getLogID() As Long

    'Start by assembling a list of matching files
    Dim logFiles() As String
    ReDim logFiles(0 To 9) As String
    
    Dim numFiles As Long
    numFiles = 0
    
    'Retrieve the first matching file from the folder (if any)
    Dim chkFile As String
    chkFile = Dir(m_logPath & "DebugReport_*.log", vbNormal)
    
    'Continue checking potential autosave XML entries until all have been analyzed
    Do While Len(chkFile) > 0
        
        'Log this file and increment the file counter
        logFiles(numFiles) = m_logPath & chkFile
        numFiles = numFiles + 1
        
        'Check the next file in the folder
        chkFile = Dir
        
    Loop
    
    'logFiles() will now contain a list of all debug logs in the current folder.  If there are already 10 entries, we want
    ' to find the oldest file in the list, and steal its ID number.
    If numFiles = 10 Then
    
        Dim minDate As Date, minID As Long
        
        'Grab the date of the first file.
        minDate = File_And_Path_Handling.FileAnyDateTime(logFiles(0))
        minID = 0
        
        'Loop through all other files; if an earlier date is found, mark that as the minimum date and ID
        Dim i As Long
        For i = 1 To 9
            If File_And_Path_Handling.FileAnyDateTime(logFiles(i)) < minDate Then
                minDate = File_And_Path_Handling.FileAnyDateTime(logFiles(i))
                minID = i
            End If
        Next i
        
        'minID now contains the ID of the oldest debug log entry.  Return it as the log ID we want to use.
        getLogID = minID
    
    Else
    
        'There are not yet 10 log files.  Use the next ID number in line.
        getLogID = numFiles
    
    End If
    
End Function

'Replace Debug.Print with this LogAction sub.  Basically it will mirror the output to the Immediate window, and add
' a new log line to the relevant debug file in the program's /Data folder.
' Input: debug string, and a BOOL indicating whether the message comes from PD's central user-visible "Message()" function
Public Sub LogAction(ByVal actionString As String, Optional ByVal isUserVisibleMessage As Boolean = False)

    'Generate a timestamp for this request
    Dim curTime As String
    curTime = Format$(Now, "ttttt", vbUseSystemDayOfWeek, vbUseSystem)

    'Modify the string to reflect whether it's a DEBUG message or user-visible MESSAGE() message
    If isUserVisibleMessage Then
        actionString = "(M) " & curTime & " | " & actionString
    Else
        actionString = "-D- " & curTime & " | " & actionString
    End If

    'Mirror output to the Immediate window
    Debug.Print actionString
    
    If m_debuggerActive Then
        
        'If file logging is active, also mirror output to this session's log file
        If m_logDatatoFile Then
        
            m_FileHandle = FreeFile
            Open m_logPath For Append As #m_FileHandle
                Print #m_FileHandle, actionString
            Close #m_FileHandle
            
        End If
    
    Else
    
        'As described at the top of this class, I like to cache certain relevant messages before the main loader is able to
        ' formally initialize this class.  When that happens, we cache the messages in a temporary array; when the class is
        ' formally initialized, we'll dump that array out to file.
        backupMessages(numBackupMessages) = actionString
        numBackupMessages = numBackupMessages + 1
        
        If numBackupMessages > UBound(backupMessages) Then ReDim Preserve backupMessages(0 To numBackupMessages * 2 - 1) As String
        
        Debug.Print "You can't use the pdDebugger class's .LogAction function without first initializing the class via InitializeDebugger!"
        
    End If

End Sub

Private Sub Class_Initialize()
    
    m_debuggerActive = False
    m_logDatatoFile = False
    
    ReDim backupMessages(0 To 15) As String
    numBackupMessages = 0

End Sub

Private Sub Class_Terminate()

    'If logging is active, post a final message
    m_FileHandle = FreeFile
    Open m_logPath For Append As #m_FileHandle
        Print #m_FileHandle, "-- END SESSION REPORT --"
    Close #m_FileHandle
    
End Sub
