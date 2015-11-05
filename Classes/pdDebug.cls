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
'Copyright 2014-2015 by Tanner Helland
'Created: 17/July/14
'Last updated: 23/June/15
'Last update: integrated pdSystemInfo for various OS interaction bits
'Dependencies: pdSystemInfo (for retrieving system-level debug data)
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

'Multiple message types can be logged by the class.  While these don't have to be used, they make it much easier to
' quickly scan the final text output.
Public Enum PD_DEBUG_MESSAGES
    PDM_NORMAL = 0
    PDM_USER_MESSAGE = 1
    PDM_MEM_REPORT = 2
    PDM_HDD_REPORT = 3
    PDM_PROCESSOR = 4
    PDM_EXTERNAL_LIB = 5
    PDM_STARTUP_MESSAGE = 6
End Enum

#If False Then
    Const PDM_NORMAL = 0, PDM_USER_MESSAGE = 1, PDM_MEM_REPORT = 2, PDM_HDD_REPORT = 3, PDM_PROCESSOR = 4, PDM_EXTERNAL_LIB = 5, PDM_STARTUP_MESSAGE = 6
#End If

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

'Number of unique events logged this session.
Private m_NumLoggedEvents As Long

'For technical reasons (how's that for nondescript), the public instance of this class cannot be formally initialized
' until fairly late into PD's load process.  However, there are technical details that would be nice to know prior to
' that point.  To work around this, I allow the class to log messages prior to being formally initialized.  When it
' *is* finally initialized, it will dump all the cached messages to the log file, with their original timestamps.
Private backupMessages() As String
Private numBackupMessages As Long

'When the user requests a RAM update, we report the delta between the current update and the previous update.
' This is very helpful for catching memory leaks.
Private m_lastMemCheck As Long, m_lastMemCheckEventNum As Long

'Various OS interactions are handled by pdSystemInfo
Private cSysInfo As pdSystemInfo

'File I/O is handled via pdFSO
Private cFile As pdFSO

'This specialty Initialize function must be called before attempting to use this class.  It will figure out where to log
' this session's data, among other things, so don't attempt to use the class until this has been called!
' Returns: TRUE if successful, FALSE otherwise.
Public Function InitializeDebugger(Optional ByVal writeLogDataToFile As Boolean = False) As Boolean
    
    If writeLogDataToFile Then
    
        'First things first: we need to make sure a Debug path exists.  Otherwise, we can't write any of our debug data to file.
        m_logPath = g_UserPreferences.getDebugPath
        
        'Make sure the path exists, and make sure we have write access.  If either of these fail, terminate the debugger.
        If Not cFile.FolderExist(m_logPath) Then
            
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
        
        If cFile.FileExist(m_logPath) Then
            cFile.KillFile m_logPath
            overwroteOldLog = True
        End If
        
        'Assemble a basic collection of relevant debug data.
        Dim debugHeader As String
            debugHeader = "**********************************************" & vbCrLf
            debugHeader = debugHeader & "-- PHOTODEMON DEBUG LOG #" & CStr(m_debuggerID + 1) & " --" & vbCrLf & vbCrLf
            
            debugHeader = debugHeader & "Date: " & Date & vbCrLf
            debugHeader = debugHeader & "Time: " & Time & vbCrLf
            debugHeader = debugHeader & "Session ID: " & g_SessionID & vbCrLf
            debugHeader = debugHeader & "Compiled: " & CStr(g_IsProgramCompiled) & vbCrLf
            debugHeader = debugHeader & "First run: " & CStr(g_IsFirstRun) & vbCrLf & vbCrLf
            
            debugHeader = debugHeader & "-- SYSTEM INFORMATION --" & vbCrLf & vbCrLf
            
            debugHeader = debugHeader & "OS: " & cSysInfo.GetOSVersionAsString & vbCrLf
            debugHeader = debugHeader & "Processor cores (logical): " & cSysInfo.GetNumLogicalCores & vbCrLf
            debugHeader = debugHeader & "Processor features: " & cSysInfo.GetProcessorFeatures & vbCrLf
            debugHeader = debugHeader & "System RAM: " & cSysInfo.GetTotalSystemRAM & vbCrLf
            debugHeader = debugHeader & "Max memory available to PhotoDemon: " & cSysInfo.GetRAMAvailableToPD & vbCrLf & vbCrLf
            
            debugHeader = debugHeader & "-- PROGRAM INFORMATION -- " & vbCrLf & vbCrLf
            
            debugHeader = debugHeader & "Version: " & getPhotoDemonNameAndVersion & vbCrLf
            debugHeader = debugHeader & "Translations active: " & CStr(g_Language.translationActive()) & vbCrLf
            debugHeader = debugHeader & "Language in use: " & CStr(g_Language.getCurrentLanguage()) & vbCrLf
            debugHeader = debugHeader & "GDI+ available: " & CStr(g_GDIPlusAvailable) & vbCrLf
            debugHeader = debugHeader & "zLib enabled: " & CStr(g_ZLibEnabled) & vbCrLf
            debugHeader = debugHeader & "FreeImage enabled: " & CStr(g_ImageFormats.FreeImageEnabled) & vbCrLf
            debugHeader = debugHeader & "ExifTool enabled: " & CStr(g_ExifToolEnabled) & vbCrLf
            debugHeader = debugHeader & "EZTwain enabled: " & CStr(g_ScanEnabled) & vbCrLf
            debugHeader = debugHeader & "PngQuant enabled: " & CStr(g_ImageFormats.pngQuantEnabled) & vbCrLf & vbCrLf
            
            debugHeader = debugHeader & "**********************************************" & vbCrLf & vbCrLf
            debugHeader = debugHeader & "-- SESSION REPORT --" & vbCrLf
        
        'Build a new log file, and dump the debug string into it
        m_logDatatoFile = True
        
        If Not cFile.SaveStringToTextFile(debugHeader, m_logPath) Then
            Debug.Print "WARNING!  Debug file could not be created!"
        End If
        
    End If
    
    m_debuggerActive = True
    
    'Log an initial event, to note that debug mode was successfully initiated
    Me.LogAction "Debugger initialized successfully"
    
    'Perform an initial memory check; this gives us a nice baseline measurement
    Me.LogAction "", PDM_MEM_REPORT
    
    'If messages were logged prior to this class being formally initialized, dump them now
    If (numBackupMessages > 0) And (writeLogDataToFile) Then
    
        Me.LogAction "(The following " & numBackupMessages & " actions were logged prior to initialization.)"
        Me.LogAction "(They are presented here with their original timestamps.)"
        
        Dim i As Long
        For i = 0 To numBackupMessages - 1
            Me.LogAction backupMessages(i), PDM_STARTUP_MESSAGE, True
        Next i
        
        Me.LogAction "(End of pre-initialization data)"
        
        'We don't need the backup messages any more, so we are free to release them into the ether
        ReDim backupMessages(0) As String
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
    chkFile = cFile.DirW(m_logPath & "DebugReport_*.log")
    
    'Continue checking potential autosave XML entries until all have been analyzed
    Do While Len(chkFile) <> 0
        
        'Log this file and increment the file counter
        logFiles(numFiles) = m_logPath & chkFile
        numFiles = numFiles + 1
        
        'Check the next file in the folder
        chkFile = cFile.DirW
        
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
Public Sub LogAction(Optional ByVal actionString As String = vbNullString, Optional ByVal debugMsgType As PD_DEBUG_MESSAGES = PDM_NORMAL, Optional ByVal suspendMemoryAutoUpdate As Boolean = False)
    
    Dim curTime As String, prefixString As String
    
    'If this message was logged at startup, skip all the usual formalities and proceed directly to writing the file.
    If debugMsgType <> PDM_STARTUP_MESSAGE Then
        
        'Increase the event count
        m_NumLoggedEvents = m_NumLoggedEvents + 1
        
        'Generate a timestamp for this request
        curTime = Format$(Now, "ttttt", vbUseSystemDayOfWeek, vbUseSystem)
        
        'Modify the string to reflect whether it's a DEBUG message or user-visible MESSAGE() message
        Select Case debugMsgType
        
            Case PDM_NORMAL
                prefixString = "-DBG-"
                
            Case PDM_USER_MESSAGE
                prefixString = "(USM)"
            
            Case PDM_MEM_REPORT
                prefixString = "*RAM*"
            
            Case PDM_HDD_REPORT
                prefixString = "^HDD^"
                
            Case PDM_PROCESSOR
                prefixString = "#PRC#"
                
            Case PDM_EXTERNAL_LIB
                prefixString = "!EXT!"
        
        End Select
        
        'For special message types, populate their contents now
        If debugMsgType = PDM_MEM_REPORT Then
        
            m_lastMemCheckEventNum = m_NumLoggedEvents
        
            'The caller wants a RAM update.  Generate one now.
            Dim curMemUsage As Double, maxMemUsage As Double, deltaMem As Double
            curMemUsage = cSysInfo.GetPhotoDemonMemoryUsage(False)
            maxMemUsage = cSysInfo.GetPhotoDemonMemoryUsage(True)
            deltaMem = curMemUsage - m_lastMemCheck
            
            'While here, also grab GDI and user object counts
            Dim curGDIObjects As Long, curUserObjects As Long
            curGDIObjects = cSysInfo.GetPhotoDemonResourceUsage(GdiObjects)
            curUserObjects = cSysInfo.GetPhotoDemonResourceUsage(UserObjects)
            
            'Format the strings into something more readable
            If curMemUsage <> 0 Then
            
                Dim curMemUsageString As String, maxMemUsageString As String, deltaMemString As String
                curMemUsageString = Format$(curMemUsage, "###,###,###,##0") & " K"
                maxMemUsageString = Format$(maxMemUsage, "###,###,###,##0") & " K"
                deltaMemString = Format$(deltaMem, "###,###,###,##0") & " K"
                If deltaMem > 0 Then deltaMemString = "+" & deltaMemString
                
                Dim gdiUsageString As String, userUsageString As String
                gdiUsageString = Format$(curGDIObjects, "###,##0")
                userUsageString = Format$(curUserObjects, "###,##0")
                
                'Ignore the supplied action string, and create our own
                actionString = "CURRENT: " & curMemUsageString & " | DELTA: " & deltaMemString & " | SESSION MAX: " & maxMemUsageString _
                                & " | GDI: " & gdiUsageString & " | USER: " & userUsageString
                
                'Also report some internal program object counts (memory DCs, hFonts, etc)
                Dim dcCountString As String
                dcCountString = Format$(g_DCsCreated - g_DCsDestroyed, "###,##0")
                dcCountString = dcCountString & "(" & Format$(g_DCsCreated, "###,##0") & ":" & Format$(g_DCsDestroyed, "###,##0") & ")"
                actionString = actionString & vbCrLf & Space$(20) & "DCs: " & dcCountString
                
                Dim fontCountString As String
                fontCountString = Format$(g_FontsCreated - g_FontsDestroyed, "###,##0")
                fontCountString = fontCountString & "(" & Format$(g_FontsCreated, "###,##0") & ":" & Format$(g_FontsDestroyed, "###,##0") & ")"
                actionString = actionString & " | FONTS: " & fontCountString
                
                Dim dibCountString As String
                dibCountString = Format$(g_DIBsCreated - g_DIBsDestroyed, "###,##0")
                dibCountString = dibCountString & "(" & Format$(g_DIBsCreated, "###,##0") & ":" & Format$(g_DIBsDestroyed, "###,##0") & ")"
                actionString = actionString & " | DIBS: " & dibCountString
                
            Else
                actionString = "WARNING: PD was unable to measure its own memory usage.  Please investigate."
            End If
            
            'Update the module-level last mem check value
            m_lastMemCheck = curMemUsage
        
        ElseIf debugMsgType = PDM_HDD_REPORT Then
        
        End If
        
        'Add a timestamp
        actionString = prefixString & " " & curTime & " | " & actionString
        
        'Mirror output to the Immediate window
        Debug.Print actionString
        
    End If
    
    'If file logging is active, also mirror output to this session's log file
    If m_debuggerActive Then
        
        If m_logDatatoFile Then
            cFile.AppendTextToFile actionString & vbCrLf, m_logPath
        End If
    
    Else
    
        'As described at the top of this class, I like to cache certain relevant messages before the main loader is able to
        ' formally initialize this class.  When that happens, we cache the messages in a temporary array; when the class is
        ' formally initialized, we'll dump that array out to file.
        backupMessages(numBackupMessages) = "(" & CStr(numBackupMessages + 1) & ") " & actionString
        numBackupMessages = numBackupMessages + 1
        
        If numBackupMessages > UBound(backupMessages) Then ReDim Preserve backupMessages(0 To numBackupMessages * 2 - 1) As String
        
    End If
    
    'For messages that are sent en masse (e.g. when loading a new image), the caller can choose to postpone automatic memory updates,
    ' as it will likely raise its own when relevant.
    If suspendMemoryAutoUpdate Or (debugMsgType = PDM_MEM_REPORT) Then m_lastMemCheckEventNum = m_lastMemCheckEventNum + 1
    
    'If we've gone 25 events without a RAM check, provide one now
    If m_NumLoggedEvents > (m_lastMemCheckEventNum + 25) Then Me.LogAction "", PDM_MEM_REPORT

End Sub

Private Sub Class_Initialize()
    
    m_debuggerActive = False
    m_logDatatoFile = False
    
    ReDim backupMessages(0 To 15) As String
    numBackupMessages = 0
    
    m_NumLoggedEvents = 0
    m_lastMemCheck = 0
    
    Set cSysInfo = New pdSystemInfo
    Set cFile = New pdFSO

End Sub

Public Sub TerminateDebugger()

    'If logging is active, post a final message
    If m_logDatatoFile Then
        cFile.AppendTextToFile "-- END SESSION REPORT --", m_logPath
        m_logDatatoFile = False
    End If
    
    If m_debuggerActive Then m_debuggerActive = False
    
End Sub

Private Sub Class_Terminate()
    TerminateDebugger
End Sub
