VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdProfilerLT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Performance Profiler - Long-term Tracking
'Copyright 2014-2018 by Tanner Helland
'Created: 19/June/14
'Last updated: 28/August/17
'Last update: code clean-up in preparation for splitting some functionality into a short-term class
'
'As PhotoDemon has grown more complex, I've had to be more careful about overall program performance.  Unfortunately,
' performance is a tough thing to track using default VB code, which is why it became necessary to write a custom
' profiler better tailored to the kind of complex performance analysis PD needs.
'
'pdProfilerLT aims to solve a specific profiling problem: tracking single-occurrence events across multiple sessions,
' and calculating things like average time taken, standard deviations, and coefficients of variation.  At present, I use
' this to track performance of PD's startup process(es), and the profiler results have been incredibly helpful at
' helping me has isolate and solve a number of perf issues during program initialization.
'
'Because this class is meant to be as simple to use as possible *from within PD*, I've added a number of PD-specific bits
' to it, like automatically using the program's known Debug folder for file output, and automatically generating filenames
' from report titles.  These PD-specific instances should be easy to strip out, but as with any class containing a "pd"
' prefix, I haven't gone out of my way to make everything in here project-agnostic.  NB: Use at your own risk!
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'The profiler supports sorting by several different criteria.  In addition to single-run timing values, we can also sort data
' by historical mean, standard deviation, and CV (to see which times are most variable).
Private Enum PD_ProfilerSort
    pdps_SessionTime = 0
    pdps_Mean = 1
    pdps_StdDev = 2
    pdps_CV = 3
End Enum

#If False Then
    Private Const pdps_SessionTime = 0, pdps_Mean = 1, pdps_StdDev = 2, pdps_CV = 3
#End If

'When timing is active, this will be set to TRUE
Private m_TimingActive As Boolean

'As a convenience, this class allows you to track up to 64 consecutive events within a single timer schedule.  (The allowed event
' count could be handled dynamically, but that would introduce variability into timing, so I avoid it.)  Those events are stored
' in this array, as they are added to the profiler.
Private Type PD_TimedEvent
    Name As String
    Index As Long
    StartTick As Currency
    EndTick As Currency
    ElapsedTime As Double
    
    'The following four values are only filled/used if database support is active
    NumHistoricalEntries As Long
    HistoricalMean As Double
    HistoricalStdDev As Double
    CoefficientVariation As Double
End Type

Private Const NUM_OF_TIMER_EVENTS As Long = 64
Private m_Events() As PD_TimedEvent

'Current event index; starts at 0 and moves up from there.
Private m_CurrentEventIndex As Long

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
            
End Sub

'Call this function to START profile recording.
' Inputs:
' 1) reportName - String that defines the purpose of the report.  Note that it will also be used as the filename of the report.
' 2) useDatabase - Store timing data between runs, and report means, std devs, and CVs in addition to per-session values.
Friend Sub StartProfiling(Optional ByVal reportName As String = "Generic report", Optional ByVal useDatabase As Boolean = False)

    'Note the use of a database.  (We won't actually do anything with this value until the log session ends, FYI.)
    m_UseDatabase = useDatabase

    'Cache the report name
    m_ReportName = reportName

    'Note that profiling has begun
    m_TimingActive = True
    
    'Reset the current event index and anything else relevant to a single profile run
    m_CurrentEventIndex = 0
    
    'Reset the event tracker array
    ReDim m_Events(0 To NUM_OF_TIMER_EVENTS - 1) As PD_TimedEvent
    
    'Make a note of the starting time
    With m_Events(m_CurrentEventIndex)
        .Index = 0
        .Name = reportName
        VBHacks.GetHighResTime .StartTick
    End With
    
    'Note that EndTick and ElapsedTime are not calculated until all profiling is complete
    
    'Advance the event index
    m_CurrentEventIndex = m_CurrentEventIndex + 1

End Sub

'Use this function to add a new event to the profiler.
Friend Sub MarkEvent(ByRef eventName As String)

    'Retrieve the current tick count
    Dim curTick As Currency
    VBHacks.GetHighResTime curTick
    
    'Fill in this time as the END of the previous tracked event
    m_Events(m_CurrentEventIndex - 1).EndTick = curTick
    
    'Start a new tracked event entry
    With m_Events(m_CurrentEventIndex)
        .Index = m_CurrentEventIndex
        .Name = eventName
        VBHacks.GetHighResTime .StartTick
    End With
    
    'Advance the event index
    m_CurrentEventIndex = m_CurrentEventIndex + 1
    
End Sub

'Call this function to STOP profile recording.  If the class is terminated before this function is called, no worries - PD will still
' clean everything up just fine, but you (obviously) won't get your timing report.
Friend Sub StopProfiling()
    
    'Fill in the final tick count for the last profile entry
    VBHacks.GetHighResTime m_Events(m_CurrentEventIndex - 1).EndTick
    
    'Note that profiling has ended
    m_TimingActive = False
    
    'Calculate the elapsed time for each individual entry in the profiler array
    Dim i As Long
    For i = 0 To m_CurrentEventIndex - 1
        m_Events(i).ElapsedTime = VBHacks.GetTimerDifference(m_Events(i).StartTick, m_Events(i).EndTick) * 1000#
    Next i
    
End Sub

'Generate a report on this profile instance.  Optional parameters may be added in the future to make the data easier to use.
Friend Sub GenerateProfileReport(Optional ByVal alsoWriteToFile As Boolean = False)
    
    'For convenience, assemble the report into one giant string.  We will decide where to stick the data later.
    Dim reportString As String
    reportString = vbNullString
    
    'Give the report a title
    reportString = reportString & vbCrLf & "-- pdProfile report for " & Chr(34) & m_Events(0).Name & Chr(34) & " --" & vbCrLf
    
    'Note the data as well
    reportString = reportString & "(Report updated at " & Format$(Now, "h:mm:ss AM/PM") & " on " & Format$(Now, "d-mmmm-yy") & ")" & vbCrLf & vbCrLf
    
    'Note the number of events tracked and total time
    Dim totalTime As Double
    totalTime = VBHacks.GetTimerDifference(m_Events(0).StartTick, m_Events(m_CurrentEventIndex - 1).EndTick) * 1000#
    
    reportString = reportString & "Number of events tracked: " & m_CurrentEventIndex & vbCrLf
    reportString = reportString & "Total time taken (this session): " & FormatNumber(CDbl(totalTime) / 1000#, , vbTrue, vbFalse, vbUseDefault) & " seconds" & vbCrLf
    reportString = reportString & "Program compiled: " & OS.IsProgramCompiled() & vbCrLf & vbCrLf
    
    'Next, report each event individually, but sorted from longest to shortest
    reportString = reportString & "Timing for THIS SESSION, with events sorted from longest to shortest:" & vbCrLf & vbCrLf
    
    SortProfileEntries pdps_SessionTime
    
    Dim i As Long
    For i = 0 To m_CurrentEventIndex - 1
        reportString = reportString & vbTab & CStr(i + 1) & ") " & m_Events(i).Name & ", " & FormatNumber(m_Events(i).ElapsedTime, 0, vbTrue, vbFalse, vbFalse) & " ms" & vbCrLf
    Next i
    
    'If the user has enabled database support, and a database file exists, update our recorded times with any
    ' saved values from previous runs.
    If m_UseDatabase And Files.FileExists(GetDatabaseFilename) Then
    
        'Retrieve values from file
        RetrieveDatabase
    
        'Display results by their historical mean time taken
        reportString = reportString & vbCrLf & "Mean time across ALL RECORDED SESSIONS, with events sorted from longest to shortest:" & vbCrLf & vbCrLf
        
        SortProfileEntries pdps_Mean
        
        Dim uncertaintyValue As Double
        
        For i = 0 To m_CurrentEventIndex - 1
            
            If (m_Events(i).HistoricalStdDev = 0#) Or (m_Events(i).NumHistoricalEntries <= 1) Then
                uncertaintyValue = 0#
            Else
                uncertaintyValue = Abs(m_Events(i).HistoricalStdDev / (m_Events(i).NumHistoricalEntries - 1))
            End If
            
            reportString = reportString & vbTab & CStr(i + 1) & ") " & m_Events(i).Name & ", " & FormatNumber(m_Events(i).HistoricalMean, 0, vbTrue, vbFalse, vbFalse) & " ± " & FormatNumber(Sqr(uncertaintyValue), 0, vbTrue, vbFalse, vbFalse) & " ms" & vbCrLf
            
        Next i
        
        'Display results by their coefficient of deviation, which is a dimensionless number defined as StdDev / Mean;
        ' This tells us which timing values are most variable.
        reportString = reportString & vbCrLf & "Coefficient of variation across ALL RECORDED SESSIONS, with events sorted from highest to lowest magnitude:" & vbCrLf & vbCrLf
        
        SortProfileEntries pdps_CV
        
        For i = 0 To m_CurrentEventIndex - 1
            reportString = reportString & vbTab & CStr(i + 1) & ") " & m_Events(i).Name & ", " & FormatNumber(m_Events(i).CoefficientVariation, 4, vbTrue, vbFalse, vbFalse) & vbCrLf
        Next i
        
    End If
    
    reportString = reportString & vbCrLf & "-- End of timing report --" & vbCrLf
    
    'If an output path was specified, write the timing report out to file now
    If alsoWriteToFile Then
    
        Dim logFile As String
        logFile = GetReportFilename()
        
        'Write the data to file using pdFSO
        Files.FileSaveAsText reportString, logFile
        
        'If the user wants us to use a persistent database for timing data, write out the current session's data now.
        If m_UseDatabase Then WriteUpdatedDatabase
        
    'If not writing the data out to file, dump it to the debug window instead
    Else
        Debug.Print reportString
    End If
    
End Sub

'After a timing session completes, use this function to write a new database out to file.
Private Sub RetrieveDatabase()

    'As always, we'll use an XML engine to retrieve the database from file
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    
    'Load the XML file into memory
    xmlEngine.LoadXMLFile GetDatabaseFilename()
    
    'Check for a few necessary tags, just to make sure this is actually a PhotoDemon XML file
    If xmlEngine.IsPDDataType("Timing Database") And xmlEngine.ValidateLoadedXMLData("reportName") Then
    
        Dim oldMean As Currency, oldStdDev As Double
    
        'Next, loop through all entries in the current timing database.  For each one, try to find a match
        ' in the XML database.
        Dim i As Long, tagPosition As Long
        For i = 0 To m_CurrentEventIndex - 1
            
            'See if the XML engine can locate an entry for this event
            tagPosition = xmlEngine.GetLocationOfTagPlusAttribute("profileEvent", "name", Replace(m_Events(i).Name, " ", "_"))
            
            'If an entry was found, retrieve it
            If (tagPosition > 0) Then
            
                m_Events(i).NumHistoricalEntries = xmlEngine.GetUniqueTag_Long("numRecordedEntries", m_Events(i).NumHistoricalEntries, tagPosition)
                oldMean = xmlEngine.GetUniqueTag_Double("eventMean", m_Events(i).HistoricalMean, tagPosition)
                oldStdDev = xmlEngine.GetUniqueTag_Double("eventStdDev", m_Events(i).HistoricalStdDev, tagPosition)
                
                'We now need to factor the current timing results into the historical average.  We do this using a very cool
                ' optimization from a 1962 paper by BP Welford (see http://www.johndcook.com/standard_deviation.html).
                ' What makes his optimization so great is not only that it allows us to accurately compute a running standard
                ' deviation without knowing all past entries, but that it allows us to do it *more accurately* when considering
                ' the rounding errors introduced by silicon-based floating-point arithmetic.  Sweet!
                m_Events(i).NumHistoricalEntries = m_Events(i).NumHistoricalEntries + 1
                m_Events(i).HistoricalMean = oldMean + (m_Events(i).ElapsedTime - oldMean) / m_Events(i).NumHistoricalEntries
                m_Events(i).HistoricalStdDev = (oldStdDev + (m_Events(i).ElapsedTime - oldMean) * (m_Events(i).ElapsedTime - m_Events(i).HistoricalMean))
                
                'Calculate a coefficient of variation for the new values
                If (m_Events(i).HistoricalMean > 0) Then
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
Private Sub WriteUpdatedDatabase()

    'As always, we'll use an XML engine to write the database to file
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    xmlEngine.PrepareNewXML "Timing Database"
    
    'Write out the name of the report, so we can check it when the database is loaded in the future
    xmlEngine.WriteTag "reportName", m_ReportName
    
    'Not technically necessary, but mark how many events we've recorded
    xmlEngine.WriteTag "eventsTracked", m_CurrentEventIndex
    
    xmlEngine.WriteBlankLine
    xmlEngine.WriteComment "This file is a performance report automatically generated by a non-production PhotoDemon build."
    xmlEngine.WriteComment "You can delete this file without consequence; it only exists to help PD's developers track down bugs."
    xmlEngine.WriteBlankLine
    
    'Write each timed event to file
    Dim i As Long
    For i = 0 To m_CurrentEventIndex - 1
    
        'Write a header for each entry
        xmlEngine.WriteTagWithAttribute "profileEvent", "name", Replace(m_Events(i).Name, " ", "_"), vbNullString, True
        
        'Before writing anything, if this is the first time writing the database, update each entry's historical values
        ' to match the current entry.
        If (m_Events(i).NumHistoricalEntries = 0) Then
            m_Events(i).NumHistoricalEntries = 1
            m_Events(i).HistoricalMean = m_Events(i).ElapsedTime
            m_Events(i).HistoricalStdDev = 0
        End If
        
        'Dump all relevant data from the type
        With xmlEngine
            .WriteTag "numRecordedEntries", m_Events(i).NumHistoricalEntries
            .WriteTag "eventMean", m_Events(i).HistoricalMean
            .WriteTag "eventStdDev", m_Events(i).HistoricalStdDev
        End With
        
        'Close this entry
        xmlEngine.CloseTag "profileEvent"
        xmlEngine.WriteBlankLine
            
    Next i
    
    'Write the file
    xmlEngine.WriteXMLToFile GetDatabaseFilename()

End Sub

'Once a report's name has been set, use the following two functions to retrieve the report's log filename and database filename,
' both of which adhere to strict patterns, which is how we can track them between runs!
Private Function GetReportFilename()

    'Generate a default filename, using the report title
    Dim dstFilename As String
    dstFilename = "pdProfileLog_" & m_ReportName
    
    'IDE and compiled .exe are tracked separately
    If (Not OS.IsProgramCompiled) Then dstFilename = dstFilename & "_IDE"
    
    dstFilename = dstFilename & ".log"
    
    'Make sure the report name can be used as a valid filename
    dstFilename = Files.FileMakeNameValid(dstFilename)
    If (InStr(1, dstFilename, " ") > 0) Then dstFilename = Replace$(dstFilename, " ", "_")
    
    'Use the report's name to write the file out to PhotoDemon's default Debug directory; since the profiler is only over used
    ' in debug mode, we know that folder exists.
    GetReportFilename = g_UserPreferences.GetDebugPath & dstFilename

End Function

Private Function GetDatabaseFilename()

    'Generate a default filename, using the report title
    Dim dstFilename As String
    dstFilename = "pdProfileLog_" & m_ReportName
    
    'IDE and compiled .exe are tracked separately
    If (Not OS.IsProgramCompiled) Then dstFilename = dstFilename & "_IDE"
    
    dstFilename = dstFilename & ".xml"
    
    'Make sure the report name can be used as a valid filename
    dstFilename = Files.FileMakeNameValid(dstFilename)
    If (InStr(1, dstFilename, " ") > 0) Then dstFilename = Replace$(dstFilename, " ", "_")
    
    'Use the report's name to write the file out to PhotoDemon's default Debug directory; since the profiler is only over used
    ' in debug mode, we know that folder exists.
    GetDatabaseFilename = g_UserPreferences.GetDebugPath & dstFilename

End Function

'Sort the profile entries from longest to shortest time taken
Private Sub SortProfileEntries(Optional ByVal sortCriteria As PD_ProfilerSort = pdps_SessionTime)

    Dim i As Long, j As Long
    
    'Loop through all entries in the profiling array, sorting them as we go
    For i = 0 To m_CurrentEventIndex - 1
        For j = 0 To m_CurrentEventIndex - 1
            
            'Compare two profiled events using the supplied criteria; if one entry exceeds the other, swap 'em
            Select Case sortCriteria
                
                Case pdps_SessionTime
                    If (m_Events(i).ElapsedTime > m_Events(j).ElapsedTime) Then SwapEventData m_Events(i), m_Events(j)
                
                Case pdps_Mean
                    If (m_Events(i).HistoricalMean > m_Events(j).HistoricalMean) Then SwapEventData m_Events(i), m_Events(j)
                
                Case pdps_StdDev
                    If (m_Events(i).HistoricalStdDev > m_Events(j).HistoricalStdDev) Then SwapEventData m_Events(i), m_Events(j)
                    
                Case pdps_CV
                    If (m_Events(i).CoefficientVariation > m_Events(j).CoefficientVariation) Then SwapEventData m_Events(i), m_Events(j)
                
            End Select
            
        Next j
    Next i

End Sub

'Swap the values of two event profile entries
Private Sub SwapEventData(ByRef profileOne As PD_TimedEvent, ByRef profileTwo As PD_TimedEvent)
    Dim tmpProf As PD_TimedEvent
    tmpProf = profileOne
    profileOne = profileTwo
    profileTwo = tmpProf
End Sub

Private Sub Class_Terminate()

    'Failsafe check to end profiling.
    If m_TimingActive Then
        StopProfiling
        Debug.Print "WARNING!  Why are you terminating an active pdProfilerLT instance?  The .StopProfiling() function exists for a reason!"
    End If

End Sub
