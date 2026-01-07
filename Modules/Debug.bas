Attribute VB_Name = "PDDebug"
'***************************************************************************
'PhotoDemon Custom Debug Logger
'Copyright 2014-2026 by Tanner Helland
'Created: 17/July/14
'Last updated: 08/March/18
'Last update: rework initialization to allow us to arbitrarily start/suspend the debugger during a given session
'Dependencies: OS module (for retrieving system-level debug data)
'
'As PhotoDemon has grown more complex, debugging has become correspondingly difficult.  Debugging on my local PC is fine
' thanks to the IDE, but a lot of behaviors are different in the compiled .exe, and user PCs are another problem entirely.
'
'To that end, a more comprehensive debugging solution was required.  Enter this class.
'
'Throughout PD, you'll see many pdDebug.LogAction() statements in place of Debug.Print.  When actions are logged this way,
' they are mirrored to the Debug window (same as Debug.Print), and also be written out to file in the program's /Data
' folder (if user preferences allow).  This allows me to retrieve relevant information from users who experience crashes.
'
'While some elements of this class are PD-specific (such as where it writes its logs to file), it wouldn't take much
' work to change those bits to fit any other project.  Aside from that particular aspect, I've tried to keep the rest
' of the class as generic as possible in case this is helpful to others.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
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

'As part of aggressive leak prevention, PD tracks creation and destruction of various
' leak-prone resources.
Public Enum PD_ResourceTracker
    PDRT_hDC = 0
    PDRT_hDIB = 1
    PDRT_hFont = 2
End Enum

#If False Then
    Private Const PDRT_hDC = 0, PDRT_hDIB = 1, PDRT_hFont = 2
#End If

Private m_GDI_DIBs As Long
Private m_GDI_Fonts As Long
Private m_GDI_DCs As Long

'Has this instance been initialized?  This will be set to true if the StartDebugger function has executed successfully.
Private m_debuggerActive As Boolean

'Does the user want us writing this data to file?  If so, this will be set to TRUE.
Private m_logDatatoFile As Boolean

'ID of this debug session.  PD allows 10 unique debug logs to be kept.  Once 10 exist, the program will start overwriting
' old debug logs.  This ID value is automatically determined by checking the /Debug subfolder, and scanning for available
' logs.
Private m_debuggerID As Long

'Full path to the log file for this debug session.  This is created by the StartDebugger function, and it relies on
' the /Debug path specified by the pdPreferences class.
'
'(Generally this is the /Data/Debug folder of wherever PhotoDemon.exe is located.)
Private m_logPath As String

'Number of unique events logged this session.
Private m_NumLoggedEvents As Long

'For technical reasons (how's that for nondescript), the public instance of this class cannot be formally initialized
' until fairly late into PD's load process.  (In particular, we need to access PD's central user preferences collection
' to know where to store debug logs.)  To prevent the loss of debug events prior to the class's formal initialization,
' we use a fallback log method during startup phases.  When this class is (finally) initialized completely, we dump all
' cached messages to the log file, with their original timestamps.
Private m_backupMessages As pdStringStack

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

Public Function GetDebugLogFilename() As String
    GetDebugLogFilename = m_logPath
End Function

Public Sub InitializeDebugger()
    
    Set m_LogString = New pdString
    Set m_backupMessages = New pdStringStack
    Set m_FSO = New pdFSO
    
    ReDim m_utf8Buffer(0) As Byte
    
    m_debuggerActive = False
    m_logDatatoFile = False
    m_NumLoggedEvents = 0
    m_lastMemCheck = 0
    
End Sub

'Replace Debug.Print with this LogAction sub.  Basically it will mirror the output to the Immediate window, and add
' a new log line to the relevant debug file in the program's /Data folder.
' Input: debug string, and a BOOL indicating whether the message comes from PD's central user-visible "Message()" function
Public Sub LogAction(Optional ByVal actionString As String = vbNullString, Optional ByVal debugMsgType As PD_DebugMessages = PDM_Normal, Optional ByVal suspendMemoryAutoUpdate As Boolean = False)
    
    'If the debugger has not been initialized, do nothing
    If (m_LogString Is Nothing) Then Exit Sub
    
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
            
            If (LenB(actionString) > 0) Then
                m_LogString.AppendLineBreak
                m_LogString.Append Space$(22)
            End If
            
            m_lastMemCheckEventNum = m_NumLoggedEvents
        
            'The caller wants a RAM update.  Generate one now.
            Dim curMemUsage As Double, maxMemUsage As Double, deltaMem As Double
            curMemUsage = OS.AppMemoryUsageInMB(False)
            maxMemUsage = OS.AppMemoryUsageInMB(True)
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
                
                Const NUMBER_FORMAT_GENERIC As String = "#,#0"
                
                m_LogString.Append "CURRENT: "
                m_LogString.Append Format$(curMemUsage, NUMBER_FORMAT_GENERIC) & " M"
                m_LogString.Append " | DELTA: "
                If (deltaMem > 0) Then m_LogString.Append "+"
                m_LogString.Append Format$(deltaMem, NUMBER_FORMAT_GENERIC) & " M"
                m_LogString.Append " | SESSION MAX: "
                m_LogString.Append Format$(maxMemUsage, NUMBER_FORMAT_GENERIC) & " M"
                m_LogString.Append " | GDI: "
                m_LogString.Append Format$(curGDIObjects, NUMBER_FORMAT_GENERIC) & " (" & Format$(gdiObjectPeak, NUMBER_FORMAT_GENERIC) & ")"
                m_LogString.Append " | USER: "
                m_LogString.Append Format$(curUserObjects, NUMBER_FORMAT_GENERIC) & " (" & Format$(userObjectPeak, NUMBER_FORMAT_GENERIC) & ")"
                m_LogString.AppendLineBreak
                
                'Also report some internal program object counts (DCs, hWnds, hFonts, etc)
                m_LogString.Append Space$(22) & "DC: "
                m_LogString.Append Format$(m_GDI_DCs, NUMBER_FORMAT_GENERIC)
                
                Dim apiWindowsCreated As Long, apiWindowsDestroyed As Long, apiWindowsNet As Long
                apiWindowsNet = UserControls.GetAPIWindowCount(apiWindowsCreated, apiWindowsDestroyed)
                m_LogString.Append " | HWND: "
                m_LogString.Append Format$(apiWindowsNet, NUMBER_FORMAT_GENERIC) & " (" & Format$(apiWindowsCreated, NUMBER_FORMAT_GENERIC) & ":" & Format$(apiWindowsDestroyed, NUMBER_FORMAT_GENERIC) & ")"
                
                m_LogString.Append " | FONT: "
                m_LogString.Append Format$(m_GDI_Fonts, NUMBER_FORMAT_GENERIC)
                
                m_LogString.Append " | DIB: "
                m_LogString.Append Format$(m_GDI_DIBs, NUMBER_FORMAT_GENERIC)
                m_LogString.Append " (MNU: "
                m_LogString.Append Format$(IconsAndCursors.GetMenuImageCount(), NUMBER_FORMAT_GENERIC)
                m_LogString.Append ")"
                
                Dim icosNet As Long, icosCreated As Long, icosDestroyed As Long
                icosNet = IconsAndCursors.GetCreatedIconCount(icosCreated, icosDestroyed)
                m_LogString.Append " | ICON: "
                m_LogString.Append Format$(icosNet, NUMBER_FORMAT_GENERIC) & " (" & Format$(icosCreated, NUMBER_FORMAT_GENERIC) & ":" & Format$(icosDestroyed, NUMBER_FORMAT_GENERIC) & ")"
                
                Dim timersNet As Long, timersCreated As Long, timersDestroyed As Long
                timersNet = UserControls.GetTimerCount(timersCreated, timersDestroyed)
                m_LogString.Append " | TIMER: "
                m_LogString.Append Format$(timersNet, NUMBER_FORMAT_GENERIC) & " (" & Format$(timersCreated, NUMBER_FORMAT_GENERIC) & ":" & Format$(timersDestroyed, NUMBER_FORMAT_GENERIC) & ")"
                
                m_LogString.Append " | UC: "
                m_LogString.Append Format$(UserControls.GetPDControlCount(), NUMBER_FORMAT_GENERIC)
                
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
        If (LenB(actionString) <> 0) Then
            m_backupMessages.AddString "<BCK< | " & Format$(Now, "ttttt", vbUseSystemDayOfWeek, vbUseSystem) & " | " & "(" & CStr(m_backupMessages.GetNumOfStrings + 1) & ") " & actionString
        End If
        
    End If
    
    'For messages that are sent en masse (e.g. when loading a new image), the caller can choose to postpone automatic memory updates,
    ' as it will likely raise its own when relevant.
    If (suspendMemoryAutoUpdate Or (debugMsgType = PDM_Mem_Report)) Then m_lastMemCheckEventNum = m_lastMemCheckEventNum + 1
    
    'If we've gone GAP_BETWEEN_MEMORY_REPORTS events without a RAM report, provide one now
    If (m_NumLoggedEvents > (m_lastMemCheckEventNum + GAP_BETWEEN_MEMORY_REPORTS)) Then PDDebug.LogAction vbNullString, PDM_Mem_Report

End Sub

'Shorcut function for logging timing results
Public Sub LogTiming(ByRef strDescription As String, ByVal timeTakenRaw As Double)
    PDDebug.LogAction "Timing report: " & strDescription & " - " & Format$(timeTakenRaw * 1000#, "#0") & " ms", PDM_Timer_Report
End Sub

'If this is the first session after a hard crash, we want to forcibly activate the debugger (if user preferences allow)
Public Sub NotifyLastSessionState(ByVal lastSessionDidntCrash As Boolean)
    
    Dim sessionsSinceLastCrash As Long
    
    'If the last session was clean, update our "sessions since last crash" tracker.
    If lastSessionDidntCrash Then
        
        'Increment the "time since last crash" tracker
        sessionsSinceLastCrash = UserPrefs.GetPref_Long("Core", "SessionsSinceLastCrash", -1)
        
        'If the program hasn't crashed recently, we don't need to process anything else
        If (sessionsSinceLastCrash <> -1) Then
        
            'The program crashed recently.  Increment our "clean session" crash tracker
            sessionsSinceLastCrash = sessionsSinceLastCrash + 1
            
            'If 10 clean sessions have passed, reset the counter.
            If (sessionsSinceLastCrash > 10) Then
                sessionsSinceLastCrash = -1
                UserPrefs.SetPref_Long "Core", "SessionsSinceLastCrash", -1
            Else
                UserPrefs.SetPref_Long "Core", "SessionsSinceLastCrash", sessionsSinceLastCrash
            End If
        
        End If
        
    'Shit - the last session crashed.  If user preferences allow, let's start the debugger and see if we can't catch
    ' the problem (as users are likely to try the same task again).
    Else
    
        'Notify the preference manager of the problem
        sessionsSinceLastCrash = 0
        UserPrefs.SetPref_Long "Core", "SessionsSinceLastCrash", sessionsSinceLastCrash
    
    End If
    
    'If we experienced a recent crash, and user preferences allow, start the debugger - *even if this is a stable build*!
    If (sessionsSinceLastCrash >= 0) Then
        
        'See if we're already running
        If (Not UserPrefs.GenerateDebugLogs) Then
            
            'We're not running.  See if the preference for debug logs is set to "automatic" mode.
            If (UserPrefs.GetDebugLogPreference = dbg_Auto) Then
                
                'The user has debug logging set to "automatic" - so we're allowed to invoke the debugger!  Start it up.
                UserPrefs.SetEmergencyDebugger True
                PDDebug.StartDebugger True, , False
                PDDebug.LogAction "WARNING!  A recent PD session crashed (" & CStr(sessionsSinceLastCrash) & " session(s) ago)."
                PDDebug.LogAction "          Even though this is a production build, debug logging has been activated as a failsafe."
                
            End If
            
        End If
        
    End If

End Sub

'This specialty Initialize function must be called before attempting to use this class.  It will figure out where to log
' this session's data, among other things, so don't attempt to use the class until this has been called!
' Returns: TRUE if successful, FALSE otherwise.
Public Function StartDebugger(Optional ByVal writeLogDataToFile As Boolean = False, Optional ByVal writeHeaderToo As Boolean = True, Optional ByVal initIsNormal As Boolean = True) As Boolean
    
    If writeLogDataToFile Then
        
        Dim i As Long
        
        'First things first: we need to make sure a Debug path exists.  Otherwise, we can't write any of our debug data to file.
        m_logPath = UserPrefs.GetDebugPath
        
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
                If (Not initIsNormal) Then debugHeader.AppendLine "WARNING: debugger was not initiated by default; session data may be incomplete"
                
                debugHeader.AppendLine "Date: " & Date$
                debugHeader.AppendLine "Time: " & Time$
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
            StartDebugger = False
            Exit Function
            
        End If
        
    End If
    
    m_debuggerActive = True
    
    'Log an initial event, to note that debug mode was successfully initiated
    PDDebug.LogAction "Debugger initialized successfully"
    
    'Perform an initial memory check; this gives us a nice baseline measurement
    PDDebug.LogAction vbNullString, PDM_Mem_Report
    
    'If messages were logged prior to this class being formally initialized, dump them now
    If (Not m_backupMessages Is Nothing) Then
        
        If (m_backupMessages.GetNumOfStrings > 0) Then
        
            PDDebug.LogAction "(The following " & m_backupMessages.GetNumOfStrings & " actions were logged prior to initialization.)"
            PDDebug.LogAction "(They are presented here with their original timestamps.)"
            
            For i = 0 To m_backupMessages.GetNumOfStrings - 1
                PDDebug.LogAction m_backupMessages.GetString(i), PDM_Startup_Message, True
            Next i
            
            PDDebug.LogAction "(End of pre-initialization data)"
            
        End If
        
        'We don't need the backup messages any more, so we are free to release them into the ether
        m_backupMessages.ResetStack 0
        
    End If
    
    StartDebugger = True
    
End Function

Public Sub TerminateDebugger(Optional ByVal terminationIsNormal As Boolean = True)

    'If logging is active, post a final message
    If m_logDatatoFile And (m_LogFileHandle <> 0) Then
        If terminationIsNormal Then WriteDebugStringAsUTF8 "-- END SESSION REPORT --" Else WriteDebugStringAsUTF8 "-- WARNING: DEBUGGER TERMINATED BY USER, SESSION DATA INCOMPLETE --"
        m_FSO.FileCloseHandle m_LogFileHandle
        m_LogFileHandle = 0
        m_logDatatoFile = False
    End If
    
    If m_debuggerActive Then m_debuggerActive = False
    
End Sub

Public Sub UpdateResourceTracker(ByVal resID As PD_ResourceTracker, ByVal resCountChange As Long)
    
    Select Case resID
        Case PDRT_hDC
            m_GDI_DCs = m_GDI_DCs + resCountChange
        Case PDRT_hDIB
            m_GDI_DIBs = m_GDI_DIBs + resCountChange
        Case PDRT_hFont
            m_GDI_Fonts = m_GDI_Fonts + resCountChange
    End Select
    
End Sub

'Search the debug folder for existing debug files, sort them by date, and automatically give this log a unique ID on the
' range [0, 9].  If there are already 10 debug files present, steal the ID of the oldest file.
Private Function GetLogID() As Long

    'Start by assembling a list of existing debug files
    Dim numFiles As Long
    numFiles = 0
    
    Dim logFiles As pdStringStack
    If m_FSO.FileFind(m_logPath & "DebugReport_*.log", logFiles) Then
    
        numFiles = logFiles.GetNumOfStrings()
        
        'logFiles contains a list of all debug logs in the current folder.  If there are already
        ' 10 entries, we want to find the oldest file in the list, and steal its ID number.
        If (numFiles = 10) Then
        
            Dim minDate As Date, minID As Long, tmpDate As Date
            
            'Grab the date of the first file.
            minDate = Files.FileGetTimeAsDate(m_logPath & logFiles.GetString(0), PDFT_WriteTime)
            minID = 0
            
            'Loop through all other files; if an earlier date is found, mark that as the minimum date and ID
            Dim i As Long
            For i = 1 To 9
                tmpDate = Files.FileGetTimeAsDate(m_logPath & logFiles.GetString(i), PDFT_WriteTime)
                If (tmpDate < minDate) Then
                    minDate = tmpDate
                    minID = i
                End If
            Next i
            
            'minID now contains the ID of the oldest debug log entry.  Return it as the log ID we want to use.
            GetLogID = minID
            PDDebug.LogAction "(Reusing debug log file #" & CStr(GetLogID) & " for this session.)"
        
        Else
        
            'There are not yet 10 log files.  Use whichever ID is missing, starting from position 0.
            For i = 0 To 9
                If (Not Files.FileExists(m_logPath & "DebugReport_" & CStr(i) & ".log")) Then
                    GetLogID = i
                    Exit For
                End If
            Next i
            
        End If
    
    'If the search function fails, start over with debug log 0
    Else
        GetLogID = 0
    End If
    
End Function

'Internal helper function that handles the "convert string to UTF-8 and append to file" part of logging
Private Sub WriteDebugStringAsUTF8(ByRef srcString As String)
    If Strings.UTF8FromString(srcString, m_utf8Buffer, m_utf8Size) Then
        If (m_LogFileHandle <> 0) Then m_FSO.FileWriteData m_LogFileHandle, VarPtr(m_utf8Buffer(0)), m_utf8Size
    End If
End Sub
