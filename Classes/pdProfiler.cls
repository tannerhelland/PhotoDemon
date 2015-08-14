VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdProfiler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Performance Profiler
'Copyright 2014-2015 by Tanner Helland
'Created: 19/June/14
'Last updated: 21/July/14
'Last update: add support for timing databases, so we can track performance (and associated metrics, like stddevs)
'              across multiple sessions
'
'As PhotoDemon has grown more complex, I've had to be more careful about overall program performance.  Unfortunately,
' performance is a tough thing to track using default VB code, which is why it became necessary to write a custom
' profiler better tailored to the kind of complex performance analysis PD needs.
'
'pdProfile aims to make profiling complex functions simpler, by allowing the user to break a profiling session into
' discrete "events", each of which gets its own timing report.  When profiling terminates, this class will automatically
' sort each discrete event by time taken, giving me an easy way to know where I can focus my optimization efforts.  As of
' July '14, the class also supports measuring performance across multiple sessions, via use of the "useDatabase" parameter
' in the initial startProfiling function.  If set to TRUE, this class will automatically write an XML database of its timing
' findings, and on subsequent runs, the latest data will be added to the running total.  This allows for tracking historic
' mean timings, standard variation, and coefficient of variation (e.g. how much does timing deviate) for all individual
' events within a single report.
'
'Many thanks to Karl Petersen's "Measuring Optimizations for Classic VB" article (http://vb.mvps.org/articles/vsm20100518.pdf)
' and accompanying "StopWatch" class (http://vb.mvps.org/samples/StopWatch/), which served as my inspiration for this project.
'
'Because this class is meant to be as simple to use as possible *from within PD* (because it's often added for a single
' debugging session, then removed once a problematic area is found), I've added a number of PD-specific bits to it, like
' automatically using the program's known Debug folder for file output, and automatically generating filenames from report
' titles.  These PD-specific instances should be easy to strip out, but as with any class containing a "pd" prefix, I haven't
' gone out of my way to make everything in here project-agnostic.  NB: Use at your own risk!
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit


'The profiler supports sorting by several different criteria.  In addition to single-run timing values, we can also sort data
' by historical mean, standard deviation, and CV (to see which times are most variable).
Private Enum PD_PROFILE_SORT_CRITERIA
    PDPSORT_SESSION_TIME = 0
    PDPSORT_MEAN = 1
    PDPSORT_STDDEV = 2
    PDPSORT_CV = 3
End Enum

#If False Then
    Private Const PDPSORT_SESSION_TIME = 0, PDPSORT_MEAN = 1, PDPSORT_STDDEV = 2, PDPSORT_CV = 3
#End If

'TIMECAPS is used to retrieve the best available resolution (wPeriodMin) for the current system timer.
Private Type TIMECAPS
   wPeriodMin As Long
   wPeriodMax As Long
End Type

'WinMM.dll provides all the interesting features we need for timing
Private Declare Function timeBeginPeriod Lib "winmm" (ByVal uPeriod As Long) As Long
Private Declare Function timeEndPeriod Lib "winmm" (ByVal uPeriod As Long) As Long
Private Declare Function timeGetTime Lib "winmm" () As Long
Private Declare Function timeGetDevCaps Lib "winmm" (lpTimeCaps As TIMECAPS, ByVal uSize As Long) As Long

'When timing is active, this will be set to TRUE
Private m_TimingActive As Boolean

'Minimum available resolution for the requested timer.  Note that timing reports CANNOT be more accurate than this amount.
Private m_MinResolution As Long

'As a convenience, this class allows you to track up to 64 consecutive events within a single timer schedule.  (The allowed event
' count could be handled dynamically, but that would introduce variability into timing, so I avoid it.)  Those events are stored
' in this array, as they are added to the profiler.
Private Type PD_TIMER_EVENT
    Name As String
    Index As Long
    StartTick As Long
    EndTick As Long
    ElapsedTime As Long
    HistoricalMean As Double        'NOTE: this value is only used if database support is active
    HistoricalStdDev As Double      'NOTE: this value is only used if database support is active
    CoefficientVariation As Double  'NOTE: this value is only used if database support is active
    NumHistoricalEntries As Long    'NOTE: this value is only used if database support is active
End Type

Private Const NUM_OF_TIMER_EVENTS As Long = 64
Private m_Events() As PD_TIMER_EVENT

'Current event index; starts at 0 and moves up from there.
Private m_CurrentEventIndex As Long

'I don't know if this variable will actually matter, but I'm curious to track it anyway... basically, when this class has some data
' it needs to crunch, it makes a note of the tick count before crunching, and a note after, and adds the difference to this variable.
' This count can then be removed from the final report total.
Private m_TimeSpentProfiling As Long

'Name of the report.  To simplify the saving process, this will automatically be used as the report's name, saving the user from having
' to specify a report title.
Private m_ReportName As String

'This class now supports the use of custom "databases", which is a fancy way of saying that alongside a human-readable report,
' it can also save its raw data to a small XML file.  If invoked again, it will load data from its previous run, then average
' it all together as a single running total - very helpful for timing data with large variances.
Private m_UseDatabase As Boolean

Private Sub Class_Initialize()

    'Reset all timing variables
    m_TimingActive = False
    
    'Retrieve the current system limit for timer resolution, and store it at module-level
    Dim tmpTimeCaps As TIMECAPS
    timeGetDevCaps tmpTimeCaps, Len(tmpTimeCaps)
    m_MinResolution = tmpTimeCaps.wPeriodMin
            
End Sub

'Call this function to START profile recording.
' Inputs:
' 1) reportName - String that defines the purpose of the report.  Note that it will also be used as the filename of the report.
' 2) useDatabase - Store timing data between runs, and report means, std devs, and CVs in addition to per-session values.
Public Sub startProfiling(Optional ByVal reportName As String = "Generic report", Optional ByVal useDatabase As Boolean = False)

    'Note the use of a database.  (We won't actually do anything with this value until the log session ends, FYI.)
    m_UseDatabase = useDatabase

    'Cache the report name
    m_ReportName = reportName

    'Note that profiling has begun
    m_TimingActive = True
    
    'Reset the current event index, time profiling count, and anything else relevant to a single profile run
    m_CurrentEventIndex = 0
    m_TimeSpentProfiling = 0
    
    'Reset the event tracker array
    ReDim m_Events(0 To NUM_OF_TIMER_EVENTS - 1) As PD_TIMER_EVENT
        
    'Set the timer to its minimum available resolution (presumably 1ms on any modern system).
    timeBeginPeriod m_MinResolution
    
    'Given the timer a moment to accept the new resolution
    Dim checkTime1 As Long, checkTime2 As Long
    Do
        checkTime2 = timeGetTime()
        If (checkTime2 - checkTime1) = m_MinResolution Then Exit Do
        checkTime1 = checkTime2
    Loop
    
    'Make a note of the starting time
    With m_Events(m_CurrentEventIndex)
        .Index = 0
        .Name = reportName
        .StartTick = timeGetTime()
    End With
    
    'Note that EndTick and ElapsedTime are not calculated until all profiling is complete
    
    'Advance the event index
    m_CurrentEventIndex = m_CurrentEventIndex + 1

End Sub

'Use this function to add a new event to the profiler.
Public Sub markEvent(ByRef eventName As String)

    'Retrieve the current tick count
    Dim curTick As Long
    curTick = timeGetTime()
    
    'Fill in this time as the END of the previous tracked event
    m_Events(m_CurrentEventIndex - 1).EndTick = curTick
    
    'Start a new tracked event entry
    With m_Events(m_CurrentEventIndex)
        .Index = m_CurrentEventIndex
        .Name = eventName
        .StartTick = timeGetTime()
    End With
    
    'Make a note of how much time we wasted in this function, which we can then subtract from the final timing report
    m_TimeSpentProfiling = m_TimeSpentProfiling + (timeGetTime() - curTick)
    
    'Advance the event index
    m_CurrentEventIndex = m_CurrentEventIndex + 1
    
End Sub

'Call this function to STOP profile recording.  If the class is terminated before this function is called, no worries - PD will still
' clean everything up just fine, but you (obviously) won't get your timing report.
Public Sub stopProfiling()
    
    'Fill in the final tick count for the last profile entry
    m_Events(m_CurrentEventIndex - 1).EndTick = timeGetTime()
    
    'Note that profiling has ended
    m_TimingActive = False
    
    'Allow the timer to return to its default resolution
    timeEndPeriod m_MinResolution
    
    'Calculate the elapsed time for each individual entry in the profiler array
    Dim i As Long
    For i = 0 To m_CurrentEventIndex - 1
        m_Events(i).ElapsedTime = m_Events(i).EndTick - m_Events(i).StartTick
    Next i
    
End Sub

'Generate a report on this profile instance.  Optional parameters may be added in the future to make the data easier to use.
Public Sub generateProfileReport(Optional ByVal alsoWriteToFile As Boolean = False)
    
    'For convenience, assemble the report into one giant string.  We will decide where to stick the data later.
    Dim reportString As String
    reportString = ""
    
    'Give the report a title
    reportString = reportString & vbCrLf & "-- pdProfile report for " & Chr(34) & m_Events(0).Name & Chr(34) & " --" & vbCrLf
    
    'Note the data as well
    reportString = reportString & "(Report updated at " & Format$(Now, "h:mm:ss AM/PM") & " on " & Format$(Now, "d-mmmm-yy") & ")" & vbCrLf & vbCrLf
    
    'Note the number of events tracked and total time
    Dim totalTime As Long
    totalTime = (m_Events(m_CurrentEventIndex - 1).EndTick - m_Events(0).StartTick) - m_TimeSpentProfiling
    
    reportString = reportString & "Number of events tracked: " & m_CurrentEventIndex & vbCrLf
    reportString = reportString & "Total time taken (this session): " & FormatNumber(totalTime / 1000, , vbTrue, vbFalse, vbUseDefault) & " seconds" & vbCrLf
    reportString = reportString & "Program compiled: " & g_IsProgramCompiled & vbCrLf & vbCrLf
    
    'Next, report each event individually, but sorted from longest to shortest
    reportString = reportString & "Timing for THIS SESSION, with events sorted from longest to shortest:" & vbCrLf & vbCrLf
    
    sortProfileEntries PDPSORT_SESSION_TIME
    
    Dim i As Long
    For i = 0 To m_CurrentEventIndex - 1
        reportString = reportString & vbTab & CStr(i + 1) & ") " & m_Events(i).Name & ", " & FormatNumber(m_Events(i).ElapsedTime, 0, vbTrue, vbFalse, vbFalse) & " ms" & vbCrLf
    Next i
    
    'If the user has enabled database support, and a database file exists, update our recorded times with any
    ' saved values from previous runs.
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    If m_UseDatabase And cFile.FileExist(getDatabaseFilename) Then
    
        'Retrieve values from file
        retrieveDatabase
    
        'Display results by their historical mean time taken
        reportString = reportString & vbCrLf & "Mean time across ALL RECORDED SESSIONS, with events sorted from longest to shortest:" & vbCrLf & vbCrLf
        
        sortProfileEntries PDPSORT_MEAN
        
        Dim uncertaintyValue As Double
        
        For i = 0 To m_CurrentEventIndex - 1
            
            If m_Events(i).HistoricalStdDev = 0 Then
                uncertaintyValue = 0
            Else
                uncertaintyValue = Abs(m_Events(i).HistoricalStdDev / (m_Events(i).NumHistoricalEntries - 1))
            End If
            
            reportString = reportString & vbTab & CStr(i + 1) & ") " & m_Events(i).Name & ", " & FormatNumber(m_Events(i).HistoricalMean, 0, vbTrue, vbFalse, vbFalse) & " ± " & FormatNumber(Sqr(uncertaintyValue), 0, vbTrue, vbFalse, vbFalse) & " ms" & vbCrLf
        Next i
        
        'Display results by their coefficient of deviation, which is a dimensionless number defined as StdDev / Mean;
        ' This tells us which timing values are most variable.
        reportString = reportString & vbCrLf & "Coefficient of variation across ALL RECORDED SESSIONS, with events sorted from highest to lowest magnitude:" & vbCrLf & vbCrLf
        
        sortProfileEntries PDPSORT_CV
        
        For i = 0 To m_CurrentEventIndex - 1
            reportString = reportString & vbTab & CStr(i + 1) & ") " & m_Events(i).Name & ", " & FormatNumber(m_Events(i).CoefficientVariation, 4, vbTrue, vbFalse, vbFalse) & vbCrLf
        Next i
        
    End If
    
    reportString = reportString & vbCrLf & "-- End of timing report --" & vbCrLf
    
    'If an output path was specified, write the timing report out to file now
    If alsoWriteToFile Then
    
        Dim logFile As String
        logFile = getReportFilename()
        
        'Write the data to file using pdFSO
        cFile.SaveStringToTextFile reportString, logFile
        
        'If the user wants us to use a persistent database for timing data, write out the current session's data now.
        If m_UseDatabase Then writeUpdatedDatabase
        
    'If not writing the data out to file, dump it to the debug window instead
    Else
        Debug.Print reportString
    End If
    
End Sub

'After a timing session completes, use this function to write a new database out to file.
Private Sub retrieveDatabase()

    'As always, we'll use an XML engine to retrieve the database from file
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    
    'Load the XML file into memory
    xmlEngine.loadXMLFile getDatabaseFilename()
    
    'Check for a few necessary tags, just to make sure this is actually a PhotoDemon XML file
    If xmlEngine.isPDDataType("Timing Database") And xmlEngine.validateLoadedXMLData("reportName") Then
    
        Dim oldMean As Double, oldStdDev As Double
    
        'Next, loop through all entries in the current timing database.  For each one, try to find a match
        ' in the XML database.
        Dim i As Long, tagPosition As Long
        For i = 0 To m_CurrentEventIndex - 1
            
            'See if the XML engine can locate an entry for this event
            tagPosition = xmlEngine.getLocationOfTagPlusAttribute("profileEvent", "name", Replace(m_Events(i).Name, " ", "_"))
            
            'If an entry was found, retrieve it
            If tagPosition > 0 Then
            
                m_Events(i).NumHistoricalEntries = xmlEngine.getUniqueTag_Long("numRecordedEntries", m_Events(i).NumHistoricalEntries, tagPosition)
                oldMean = xmlEngine.getUniqueTag_Double("eventMean", m_Events(i).HistoricalMean, tagPosition)
                oldStdDev = xmlEngine.getUniqueTag_Double("eventStdDev", m_Events(i).HistoricalStdDev, tagPosition)
                
                'We now need to factor the current timing results into the historical average.  We do this using a very cool
                ' optimization from a 1962 paper by BP Welford (see http://www.johndcook.com/standard_deviation.html).
                ' What makes his optimization so great is not only that it allows us to accurately compute a running standard
                ' deviation without knowing all past entries, but that it allows us to do it *more accurately* when considering
                ' the rounding errors introduced by silicon-based floating-point arithmetic.  Sweet!
                m_Events(i).NumHistoricalEntries = m_Events(i).NumHistoricalEntries + 1
                m_Events(i).HistoricalMean = oldMean + (m_Events(i).ElapsedTime - oldMean) / m_Events(i).NumHistoricalEntries
                m_Events(i).HistoricalStdDev = (oldStdDev + (m_Events(i).ElapsedTime - oldMean) * (m_Events(i).ElapsedTime - m_Events(i).HistoricalMean))
                
                'Calculate a coefficient of variation for the new values
                If m_Events(i).HistoricalMean > 0 Then
                    m_Events(i).CoefficientVariation = Sqr(Abs(m_Events(i).HistoricalStdDev)) / m_Events(i).HistoricalMean
                Else
                    m_Events(i).CoefficientVariation = 0
                End If
                                
                'Update complete!  Doesn't get much easier than this...
                
            'The entry was not found.  Populate it with default values.
            Else
            
                m_Events(i).NumHistoricalEntries = 1
                m_Events(i).HistoricalMean = m_Events(i).ElapsedTime
                m_Events(i).HistoricalStdDev = 0
            
            End If
            
        Next i
    
    End If
    
End Sub

'After a timing session completes, use this function to write a new database out to file.
Private Sub writeUpdatedDatabase()

    'As always, we'll use an XML engine to write the database to file
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    xmlEngine.prepareNewXML "Timing Database"
    
    'Write out the name of the report, so we can check it when the database is loaded in the future
    xmlEngine.writeTag "reportName", m_ReportName
    
    'Not technically necessary, but mark how many events we've recorded
    xmlEngine.writeTag "eventsTracked", m_CurrentEventIndex
    
    xmlEngine.writeBlankLine
    xmlEngine.writeComment "This file is a performance report automatically generated by a non-production PhotoDemon build."
    xmlEngine.writeComment "You can delete this file without consequence; it only exists to help PD's developers track down bugs."
    xmlEngine.writeBlankLine
    
    'Write each timed event to file
    Dim i As Long
    For i = 0 To m_CurrentEventIndex - 1
    
        'Write a header for each entry
        xmlEngine.writeTagWithAttribute "profileEvent", "name", Replace(m_Events(i).Name, " ", "_"), "", True
        
        'Before writing anything, if this is the first time writing the database, update each entry's historical values
        ' to match the current entry.
        If m_Events(i).NumHistoricalEntries = 0 Then
            m_Events(i).NumHistoricalEntries = 1
            m_Events(i).HistoricalMean = m_Events(i).ElapsedTime
            m_Events(i).HistoricalStdDev = 0
        End If
        
        'Dump all relevant data from the type
        With xmlEngine
            .writeTag "numRecordedEntries", m_Events(i).NumHistoricalEntries
            .writeTag "eventMean", m_Events(i).HistoricalMean
            .writeTag "eventStdDev", m_Events(i).HistoricalStdDev
        End With
        
        'Close this entry
        xmlEngine.closeTag "profileEvent"
        xmlEngine.writeBlankLine
            
    Next i
    
    'Write the file
    xmlEngine.writeXMLToFile getDatabaseFilename()

End Sub

'Once a report's name has been set, use the following two functions to retrieve the report's log filename and database filename,
' both of which adhere to strict patterns, which is how we can track them between runs!
Private Function getReportFilename()

    'Generate a default filename, using the report title
    Dim dstFilename As String
    dstFilename = "pdProfileLog_" & m_ReportName
    
    'IDE and compiled .exe are tracked separately
    If Not g_IsProgramCompiled Then dstFilename = dstFilename & "_IDE"
    
    dstFilename = dstFilename & ".log"
    
    'Make sure the report name can be used as a valid filename
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    dstFilename = cFile.MakeValidWindowsFilename(dstFilename)
    If InStr(1, dstFilename, " ") > 0 Then dstFilename = Replace(dstFilename, " ", "_")
    
    'Use the report's name to write the file out to PhotoDemon's default Debug directory; since the profiler is only over used
    ' in debug mode, we know that folder exists.
    getReportFilename = g_UserPreferences.getDebugPath & dstFilename

End Function

Private Function getDatabaseFilename()

    'Generate a default filename, using the report title
    Dim dstFilename As String
    dstFilename = "pdProfileLog_" & m_ReportName
    
    'IDE and compiled .exe are tracked separately
    If Not g_IsProgramCompiled Then dstFilename = dstFilename & "_IDE"
    
    dstFilename = dstFilename & ".xml"
    
    'Make sure the report name can be used as a valid filename
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    dstFilename = cFile.MakeValidWindowsFilename(dstFilename)
    If InStr(1, dstFilename, " ") > 0 Then dstFilename = Replace(dstFilename, " ", "_")
    
    'Use the report's name to write the file out to PhotoDemon's default Debug directory; since the profiler is only over used
    ' in debug mode, we know that folder exists.
    getDatabaseFilename = g_UserPreferences.getDebugPath & dstFilename

End Function

'Sort the profile entries from longest to shortest time taken
Private Sub sortProfileEntries(Optional ByVal sortCriteria As PD_PROFILE_SORT_CRITERIA = PDPSORT_SESSION_TIME)

    Dim i As Long, j As Long
    
    'Loop through all entries in the profiling array, sorting them as we go
    For i = 0 To m_CurrentEventIndex - 1
        For j = 0 To m_CurrentEventIndex - 1
            
            'Compare two profiled events using the supplied criteria; if one entry exceeds the other, swap 'em
            Select Case sortCriteria
                
                Case PDPSORT_SESSION_TIME
                    If m_Events(i).ElapsedTime > m_Events(j).ElapsedTime Then swapEventData m_Events(i), m_Events(j)
                
                Case PDPSORT_MEAN
                    If m_Events(i).HistoricalMean > m_Events(j).HistoricalMean Then swapEventData m_Events(i), m_Events(j)
                
                Case PDPSORT_STDDEV
                    If m_Events(i).HistoricalStdDev > m_Events(j).HistoricalStdDev Then swapEventData m_Events(i), m_Events(j)
                    
                Case PDPSORT_CV
                    If m_Events(i).CoefficientVariation > m_Events(j).CoefficientVariation Then swapEventData m_Events(i), m_Events(j)
                
            End Select
            
        Next j
    Next i

End Sub

'Swap the values of two event profile entries
Private Sub swapEventData(ByRef profileOne As PD_TIMER_EVENT, ByRef profileTwo As PD_TIMER_EVENT)
    Dim tmpProf As PD_TIMER_EVENT
    tmpProf = profileOne
    profileOne = profileTwo
    profileTwo = tmpProf
End Sub

Private Sub Class_Terminate()

    'Failsafe check to end profiling.
    If m_TimingActive Then
        stopProfiling
        Debug.Print "WARNING!  Why are you terminating an active pdProfile instance?  The stopProfiling() function exists for a reason!"
    End If

End Sub
