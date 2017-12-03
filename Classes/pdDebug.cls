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
'Copyright 2014-2017 by Tanner Helland
'Created: 17/July/14
'Last updated: 28/August/17
'Last update: heavily optimized file writes by using pdFSO and a persistent append handle
'Dependencies: OS module (for retrieving system-level debug data)
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

'By default, memory and object reports are generated every (n) actions.  (Some places, like program startup,
' will also produce these messages manually.)  To reduce the frequency of memory reports, set this number
' to a higher value.
Private Const GAP_BETWEEN_MEMORY_REPORTS As Long = 50

'Multiple message types can be logged by the class.  While these don't have to be used, they make it much easier to
' quickly scan the final text output.
Public Enum PD_DebugMessages
    PDM_Normal = 0
    PDM_User_Message = 1
    PDM_Mem_Report = 2
    PDM_HDD_Report = 3
    PDM_Processor = 4
    PDM_External_Lib = 5
    PDM_Startup_Message = 6
    PDM_Timer_Report = 7
End Enum

#If False Then
    Private Const PDM_Normal = 0, PDM_User_Message = 1, PDM_Mem_Report = 2, PDM_HDD_Report = 3, PDM_Processor = 4, PDM_External_Lib = 5, PDM_Startup_Message = 6, PDM_Timer_Report = 7
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
Private Const INIT_NUM_BACKUP_MESSAGES As Long = 32
Private m_backupMessages() As String
Private m_numBackupMessages As Long

'When the user requests a RAM update, we report the delta between the current update and the previous update.
' This is very helpful for catching memory leaks.
Private m_lastMemCheck As Long, m_lastMemCheckEventNum As Long

'File I/O is handled via pdFSO.  A dedicated append handle is created when the log is first written; just to be safe,
' please check for a null handle before subsequent writes.
Private m_FSO As pdFSO
Private m_LogFileHandle As Long

'Because this class writes out a *ton* of strings - strings that need to be converted to UTF-8 prior to writing -
' we reuse a UTF-8 conversion buffer between writes.  This buffer should only be size-increased if necessary.
Private m_utf8Buffer() As Byte, m_utf8Size As Long

'A string builder is used to cut down on string allocations
Private m_LogString As pdString

'This specialty Initialize function must be called before attempting to use this class.  It will figure out where to log
' this session's data, among other things, so don't attempt to use the class until this has been called!
' Returns: TRUE if successful, FALSE otherwise.
Friend Function InitializeDebugger(Optional ByVal writeLogDataToFile As Boolean = False, Optional ByVal writeHeaderToo As Boolean = True) As Boolean
    
    If writeLogDataToFile Then
        
        Dim i As Long
        
        'First things first: we need to make sure a Debug path exists.  Otherwise, we can't write any of our debug data to file.
        m_logPath = g_UserPreferences.GetDebugPath
        
        'Make sure the path exists, and make sure we have write access.  If either of these fail, terminate the debugger.
        If Files.PathExists(m_logPath, True) Then
        
            'We now know the Debug path exists.  Retrieve a relevant ID for this file.
            m_debuggerID = GetLogID()
            
            'Generate a filename for this log, using that ID.
            m_logPath = m_logPath & "DebugReport_" & m_debuggerID & ".log"
            
            'If a log file already exists at that location, remove it.  (Only 10 log files are allowed, so if we reach 11,
            ' the oldest one will be overwritten.)
            Files.FileDeleteIfExists m_logPath
            
            'Assemble a basic collection of relevant debug data.
            Dim debugHeader As pdString
            Set debugHeader = New pdString
            
            If writeHeaderToo Then
            
                debugHeader.AppendLine "**********************************************" & vbCrLf
                debugHeader.AppendLine "-- PHOTODEMON DEBUG LOG #" & CStr(m_debuggerID + 1) & " --" & vbCrLf
                
                debugHeader.AppendLine "Date: " & Date
                debugHeader.AppendLine "Time: " & Time
                debugHeader.AppendLine "Session ID: " & OS.UniqueSessionID()
                debugHeader.AppendLine "Compiled: " & CStr(OS.IsProgramCompiled)
                debugHeader.AppendLine "First run: " & CStr(g_IsFirstRun) & vbCrLf
                
                debugHeader.AppendLine "-- SYSTEM INFORMATION --" & vbCrLf
                
                debugHeader.AppendLine "OS: " & OS.OSVersionAsString
                debugHeader.AppendLine "Processor cores (logical): " & OS.LogicalCoreCount
                debugHeader.AppendLine "Processor features: " & OS.ProcessorFeatures
                debugHeader.AppendLine "System RAM: " & OS.RAM_SystemTotal
                debugHeader.AppendLine "Max memory available to PhotoDemon: " & OS.RAM_Available
                debugHeader.AppendLine "Memory load at startup: " & OS.RAM_CurrentLoad & vbCrLf
                
                debugHeader.AppendLine "-- PROGRAM INFORMATION -- " & vbCrLf
                
                debugHeader.AppendLine "Version: " & GetPhotoDemonNameAndVersion
                debugHeader.AppendLine "Translation active: " & CStr(g_Language.TranslationActive())
                debugHeader.AppendLine "Language in use: " & CStr(g_Language.GetCurrentLanguage())
                debugHeader.AppendLine "GDI+ available: " & CStr(Drawing2D.IsRenderingEngineActive(P2_GDIPlusBackend)) & vbCrLf
                
                debugHeader.AppendLine "-- PLUGIN INFORMATION -- " & vbCrLf
                
                For i = 0 To PluginManager.GetNumOfPlugins - 1
                    debugHeader.Append PluginManager.GetPluginName(i) & ": "
                    If PluginManager.IsPluginCurrentlyInstalled(i) Then
                        debugHeader.Append "available"
                    Else
                        debugHeader.Append "MISSING"
                    End If
                    debugHeader.AppendLineBreak
                Next i
            
            End If
            
            debugHeader.AppendLine vbCrLf & "**********************************************" & vbCrLf
            debugHeader.AppendLine "-- SESSION REPORT --" & vbCrLf
            
            'Grab a persistent append handle to the log file
            m_logDatatoFile = True
            If (Not m_FSO.FileCreateAppendHandle(m_logPath, m_LogFileHandle)) Then m_LogFileHandle = 0
            
            'Convert the first chunk of debug text to UTF-8, then write the data to file
            WriteDebugStringAsUTF8 debugHeader.ToString()
            
        'File writing is requested, but the log file folder is inaccessible
        Else
        
            Debug.Print "Log path invalid.  Saved debug logs not available for this session."
            
            m_debuggerActive = False
            InitializeDebugger = False
            Exit Function
            
        End If
        
    End If
    
    m_debuggerActive = True
    
    'Log an initial event, to note that debug mode was successfully initiated
    Me.LogAction "Debugger initialized successfully"
    
    'Perform an initial memory check; this gives us a nice baseline measurement
    Me.LogAction "", PDM_Mem_Report
    
    'If messages were logged prior to this class being formally initialized, dump them now
    If (m_numBackupMessages > 0) And (writeLogDataToFile) Then
    
        Me.LogAction "(The following " & m_numBackupMessages & " actions were logged prior to initialization.)"
        Me.LogAction "(They are presented here with their original timestamps.)"
        
        For i = 0 To m_numBackupMessages - 1
            Me.LogAction m_backupMessages(i), PDM_Startup_Message, True
        Next i
        
        Me.LogAction "(End of pre-initialization data)"
        
        'We don't need the backup messages any more, so we are free to release them into the ether
        ReDim m_backupMessages(0) As String
        m_numBackupMessages = 0
        
    End If
    
    InitializeDebugger = True
    
End Function

'Search the debug folder for existing debug files, sort them by date, and automatically give this log a unique ID on the
' range [0, 9].  If there are already 10 debug files present, steal the ID of the oldest file.
Private Function GetLogID() As Long

    'Start by assembling a list of matching files
    Dim logFiles() As String
    ReDim logFiles(0 To 9) As String
    
    Dim numFiles As Long
    numFiles = 0
    
    'Retrieve the first matching file from the folder (if any)
    Dim chkFile As String
    chkFile = m_FSO.DirW(m_logPath & "DebugReport_*.log")
    
    'Continue checking potential autosave XML entries until all have been analyzed
    Do While (Len(chkFile) <> 0)
        
        'Log this file and increment the file counter
        logFiles(numFiles) = m_logPath & chkFile
        numFiles = numFiles + 1
        
        'Check the next file in the folder
        chkFile = m_FSO.DirW
        
    Loop
    
    'logFiles() will now contain a list of all debug logs in the current folder.  If there are already 10 entries, we want
    ' to find the oldest file in the list, and steal its ID number.
    If (numFiles = 10) Then
    
        Dim minDate As Date, minID As Long
        
        'Grab the date of the first file.
        minDate = Files.FileGetTimeAsDate(logFiles(0), PDFT_WriteTime)
        minID = 0
        
        'Loop through all other files; if an earlier date is found, mark that as the minimum date and ID
        Dim i As Long
        For i = 1 To 9
            If (Files.FileGetTimeAsDate(logFiles(i), PDFT_WriteTime) < minDate) Then
                minDate = Files.FileGetTimeAsDate(logFiles(i), PDFT_WriteTime)
                minID = i
            End If
        Next i
        
        'minID now contains the ID of the oldest debug log entry.  Return it as the log ID we want to use.
        GetLogID = minID
    
    Else
    
        'There are not yet 10 log files.  Use whichever ID is missing, starting from position 0.
        For i = 0 To 9
            If (Not Files.FileExists(m_logPath & "DebugReport_" & CStr(i) & ".log")) Then
                GetLogID = i
                Exit For
            End If
        Next i
    
    End If
    
End Function

'Replace Debug.Print with this LogAction sub.  Basically it will mirror the output to the Immediate window, and add
' a new log line to the relevant debug file in the program's /Data folder.
' Input: debug string, and a BOOL indicating whether the message comes from PD's central user-visible "Message()" function
Friend Sub LogAction(Optional ByVal actionString As String = vbNullString, Optional ByVal debugMsgType As PD_DebugMessages = PDM_Normal, Optional ByVal suspendMemoryAutoUpdate As Boolean = False)
    
    m_LogString.Reset
    
    'If this message was logged at startup, skip all the usual formalities and proceed directly to writing the file.
    If (debugMsgType <> PDM_Startup_Message) Then
        
        'Increase the event count
        m_NumLoggedEvents = m_NumLoggedEvents + 1
        
        'Modify the string to reflect whether it's a DEBUG message or user-visible MESSAGE() message
        Select Case debugMsgType
        
            Case PDM_Normal
                m_LogString.Append "-DBG-"
                
            Case PDM_User_Message
                m_LogString.Append "(USM)"
            
            Case PDM_Mem_Report
                m_LogString.Append "*RAM*"
            
            Case PDM_HDD_Report
                m_LogString.Append "^HDD^"
                
            Case PDM_Processor
                m_LogString.Append "#PRC#"
                
            Case PDM_External_Lib
                m_LogString.Append "!EXT!"
                
            Case PDM_Timer_Report
                m_LogString.Append "/TMR/"
        
        End Select
        
        'Generate a timestamp for this request
        m_LogString.Append " | "
        m_LogString.Append Format$(Now, "ttttt", vbUseSystemDayOfWeek, vbUseSystem)
        m_LogString.Append " | "
        m_LogString.Append actionString
        
        'For special message types, populate their contents now
        If (debugMsgType = PDM_Mem_Report) Then
            
            If (Len(actionString) > 0) Then
                m_LogString.AppendLineBreak
                m_LogString.Append Space$(22)
            End If
            
            m_lastMemCheckEventNum = m_NumLoggedEvents
        
            'The caller wants a RAM update.  Generate one now.
            Dim curMemUsage As Double, maxMemUsage As Double, deltaMem As Double
            curMemUsage = OS.AppMemoryUsage(False)
            maxMemUsage = OS.AppMemoryUsage(True)
            deltaMem = curMemUsage - m_lastMemCheck
            
            'While here, also grab GDI and user object counts
            Dim curGDIObjects As Long, curUserObjects As Long, gdiObjectPeak As Long, userObjectPeak As Long
            curGDIObjects = OS.AppResourceUsage(PDGR_GdiObjects)
            curUserObjects = OS.AppResourceUsage(PDGR_UserObjects)
            If OS.IsWin7OrLater Then
                gdiObjectPeak = OS.AppResourceUsage(PDGR_GdiObjectsPeak)
                userObjectPeak = OS.AppResourceUsage(PDGR_UserObjectsPeak)
            End If
            
            'Format the return objects into something readable
            If (curMemUsage <> 0) Then
            
                m_LogString.Append "CURRENT: "
                m_LogString.Append Format$(curMemUsage, "###,###,###,##0") & " K"
                m_LogString.Append " | DELTA: "
                If (deltaMem > 0) Then m_LogString.Append "+"
                m_LogString.Append Format$(deltaMem, "###,###,###,##0") & " K"
                m_LogString.Append " | SESSION MAX: "
                m_LogString.Append Format$(maxMemUsage, "###,###,###,##0") & " K"
                m_LogString.Append " | GDI: "
                m_LogString.Append Format$(curGDIObjects, "###,##0") & " (" & Format$(gdiObjectPeak, "###,##0") & ")"
                m_LogString.Append " | USER: "
                m_LogString.Append Format$(curUserObjects, "###,##0") & " (" & Format$(userObjectPeak, "###,##0") & ")"
                m_LogString.AppendLineBreak
                
                'Also report some internal program object counts (DCs, hWnds, hFonts, etc)
                m_LogString.Append Space$(22) & "DC: "
                m_LogString.Append Format$(g_DCsCreated - g_DCsDestroyed, "###,##0") & " (" & Format$(g_DCsCreated, "###,##0") & ":" & Format$(g_DCsDestroyed, "###,##0") & ")"
                
                Dim apiWindowsCreated As Long, apiWindowsDestroyed As Long, apiWindowsNet As Long
                apiWindowsNet = UserControls.GetAPIWindowCount(apiWindowsCreated, apiWindowsDestroyed)
                m_LogString.Append " | HWND: "
                m_LogString.Append Format$(apiWindowsNet, "###,##0") & " (" & Format$(apiWindowsCreated, "###,##0") & ":" & Format$(apiWindowsDestroyed, "###,##0") & ")"
                
                m_LogString.Append " | FONT: "
                m_LogString.Append Format$(g_FontsCreated - g_FontsDestroyed, "###,##0") & " (" & Format$(g_FontsCreated, "###,##0") & ":" & Format$(g_FontsDestroyed, "###,##0") & ")"
                
                m_LogString.Append " | DIB: "
                m_LogString.Append Format$(g_DIBsCreated - g_DIBsDestroyed, "###,##0") & " (" & Format$(g_DIBsCreated, "###,##0") & ":" & Format$(g_DIBsDestroyed, "###,##0") & ")"
                
                Dim icosNet As Long, icosCreated As Long, icosDestroyed As Long
                icosNet = IconsAndCursors.GetCreatedIconCount(icosCreated, icosDestroyed)
                m_LogString.Append " | ICON: "
                m_LogString.Append Format$(icosNet, "###,##0") & " (" & Format$(icosCreated, "###,##0") & ":" & Format$(icosDestroyed, "###,##0") & ")"
                
                Dim hooksNet As Long, hooksCreated As Long, hooksDestroyed As Long
                hooksNet = UserControls.GetHookCount(hooksCreated, hooksDestroyed)
                m_LogString.AppendLineBreak
                m_LogString.Append Space$(22) & "HOOK: "
                m_LogString.Append Format$(hooksNet, "###,##0") & " (" & Format$(hooksCreated, "###,##0") & ":" & Format$(hooksDestroyed, "###,##0") & ")"
                
                Dim timersNet As Long, timersCreated As Long, timersDestroyed As Long
                timersNet = UserControls.GetTimerCount(timersCreated, timersDestroyed)
                m_LogString.Append " | TIMER: "
                m_LogString.Append Format$(timersNet, "###,##0") & " (" & Format$(timersCreated, "###,##0") & ":" & Format$(timersDestroyed, "###,##0") & ")"
                
                m_LogString.Append " | UC: "
                m_LogString.Append Format$(UserControls.GetPDControlCount(), "###,##0")
                
            Else
                m_LogString.Append "WARNING: PD was unable to measure its own memory usage.  Please investigate."
            End If
            
            'Update the module-level last mem check value
            m_lastMemCheck = curMemUsage
        
        'In the future, it may be helpful to track how much HDD space we use.  This is not yet implemented, though.
        ElseIf (debugMsgType = PDM_HDD_Report) Then
        
        End If
        
        'Mirror output to the Immediate window, then append a final vbCrLf before dumping out to file
        Debug.Print m_LogString.ToString()
        m_LogString.AppendLineBreak
        
    Else
        Debug.Print actionString
        m_LogString.AppendLine actionString
    End If
    
    'If file logging is active, also mirror output to this session's log file
    If m_debuggerActive Then
        If m_logDatatoFile Then WriteDebugStringAsUTF8 m_LogString.ToString()
    Else
    
        'As described at the top of this class, I like to cache certain relevant messages before the main loader is able to
        ' formally initialize this class.  When that happens, we cache the messages in a temporary array; when the class is
        ' formally initialized, we'll dump that array out to file.
        If (m_numBackupMessages > UBound(m_backupMessages)) Then ReDim Preserve m_backupMessages(0 To m_numBackupMessages * 2 - 1) As String
        m_backupMessages(m_numBackupMessages) = "<BCK< | " & Format$(Now, "ttttt", vbUseSystemDayOfWeek, vbUseSystem) & " | " & "(" & CStr(m_numBackupMessages + 1) & ") " & actionString
        m_numBackupMessages = m_numBackupMessages + 1
        
    End If
    
    'For messages that are sent en masse (e.g. when loading a new image), the caller can choose to postpone automatic memory updates,
    ' as it will likely raise its own when relevant.
    If (suspendMemoryAutoUpdate Or (debugMsgType = PDM_Mem_Report)) Then m_lastMemCheckEventNum = m_lastMemCheckEventNum + 1
    
    'If we've gone GAP_BETWEEN_MEMORY_REPORTS events without a RAM report, provide one now
    If (m_NumLoggedEvents > (m_lastMemCheckEventNum + GAP_BETWEEN_MEMORY_REPORTS)) Then Me.LogAction vbNullString, PDM_Mem_Report

End Sub

'Shorcut function for logging timing results
Friend Sub LogTiming(ByRef strDescription As String, ByVal timeTakenRaw As Double)
    Me.LogAction "Timing report: " & strDescription & " - " & Format$(timeTakenRaw * 1000#, "#####0") & " ms", PDM_Timer_Report
End Sub

'Internal helper function that handles the "convert string to UTF-8 and append to file" part of logging
Private Sub WriteDebugStringAsUTF8(ByRef srcString As String)
    If Strings.UTF8FromString(srcString, m_utf8Buffer, m_utf8Size) Then
        If (m_LogFileHandle <> 0) Then m_FSO.FileWriteData m_LogFileHandle, VarPtr(m_utf8Buffer(0)), m_utf8Size
    End If
End Sub

Private Sub Class_Initialize()
    
    Set m_LogString = New pdString
    
    m_debuggerActive = False
    m_logDatatoFile = False
    
    ReDim m_backupMessages(0 To INIT_NUM_BACKUP_MESSAGES - 1) As String
    m_numBackupMessages = 0
    
    m_NumLoggedEvents = 0
    m_lastMemCheck = 0
    
    ReDim m_utf8Buffer(0) As Byte
    
    Set m_FSO = New pdFSO

End Sub

Public Sub TerminateDebugger()

    'If logging is active, post a final message
    If m_logDatatoFile And (m_LogFileHandle <> 0) Then
        WriteDebugStringAsUTF8 "-- END SESSION REPORT --"
        m_FSO.FileCloseHandle m_LogFileHandle
        m_LogFileHandle = 0
        m_logDatatoFile = False
    End If
    
    If m_debuggerActive Then m_debuggerActive = False
    
End Sub

Private Sub Class_Terminate()
    TerminateDebugger
End Sub
