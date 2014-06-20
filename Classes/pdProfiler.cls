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
'Copyright ©2013-2014 by Tanner Helland
'Created: 19/June/14
'Last updated: 19/June/14
'Last update: initial build
'
'As PhotoDemon has grown more complex, I've had to be more careful about overall program performance.  Unfortunately,
' performance is a tough thing to track using default VB code, hence the existence of this class.
'
'pdProfile aims to make profiling complex functions simpler, by allowing the user to break a profiling session into
' discrete "events", each of which gets its own timing report.  When profiling terminates, this class will automatically
' sort each discrete event by time taken, giving me an easy way to know where I can focus my optimization efforts.
'
'Many thanks to Karl Petersen's "Measuring Optimizations for Classic VB" article (http://vb.mvps.org/articles/vsm20100518.pdf)
' and accompanying "StopWatch" class (http://vb.mvps.org/samples/StopWatch/), which served as my inspiration for this project.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit


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
End Type

Private Const NUM_OF_TIMER_EVENTS As Long = 64
Private m_Events() As PD_TIMER_EVENT

'Current event index; starts at 0 and moves up from there.
Private m_CurrentEventIndex As Long

'I don't know if this variable will actually matter, but I'm curious to track it anyway... basically, when this class has some data
' it needs to crunch, it makes a note of the tick count before crunching, and a note after, and adds the difference to this variable.
' This count can then be removed from the final report total.
Private m_TimeSpentProfiling As Long

Private Sub Class_Initialize()

    'Reset all timing variables
    m_TimingActive = False
    
    'Retrieve the current system limit for timer resolution, and store it at module-level
    Dim tmpTimeCaps As TIMECAPS
    timeGetDevCaps tmpTimeCaps, Len(tmpTimeCaps)
    m_MinResolution = tmpTimeCaps.wPeriodMin
            
End Sub

'Call this function to START profile recording
Public Sub startProfiling(Optional ByVal reportName As String = "Generic report")

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
Public Sub generateProfileReport(Optional ByVal dstTextFile As String = "")

    'For convenience, assemble the report into one giant string.  We can decide where to stick the data later.
    Dim reportString As String
    reportString = ""
    
    'Give the report a title
    reportString = reportString & vbCrLf & "-- pdProfile report for " & Chr(34) & m_Events(0).Name & Chr(34) & " --" & vbCrLf
    
    'Note the data as well
    reportString = reportString & "(Report generated at " & Format$(Now, "h:mm:ss AM/PM") & " on " & Format$(Now, "d-mmmm-yy") & ")" & vbCrLf & vbCrLf
    
    'Note the number of events tracked and total time
    Dim totalTime As Long
    totalTime = (m_Events(m_CurrentEventIndex - 1).EndTick - m_Events(0).StartTick) - m_TimeSpentProfiling
    
    reportString = reportString & "Number of events tracked: " & m_CurrentEventIndex & vbCrLf
    reportString = reportString & "Total time taken: " & FormatNumber(totalTime / 1000, , vbTrue, vbFalse, vbUseDefault) & " seconds" & vbCrLf & vbCrLf
    
    'Next, report each event individually, in original order
    reportString = reportString & "Timing for individual events, in original order:" & vbCrLf & vbCrLf
    
    Dim i As Long
    For i = 0 To m_CurrentEventIndex - 1
        reportString = reportString & vbTab & CStr(m_Events(i).Index) & ") " & m_Events(i).Name & ", " & FormatNumber(m_Events(i).ElapsedTime, 0, vbTrue, vbFalse, vbFalse) & " ms" & vbCrLf
    Next i
    
    reportString = reportString & vbCrLf
    
    'Next, report each event individually, but sorted from longest to shortest
    reportString = reportString & "Timing for individual events, sorted from longest to shortest:" & vbCrLf & vbCrLf
    
    sortProfileEntries
    
    For i = 0 To m_CurrentEventIndex - 1
        reportString = reportString & vbTab & CStr(i) & ") " & m_Events(i).Name & ", " & FormatNumber(m_Events(i).ElapsedTime, 0, vbTrue, vbFalse, vbFalse) & " ms" & vbCrLf
    Next i
    
    reportString = reportString & vbCrLf & "-- End of timing report --" & vbCrLf
    
    'For now, dump the profile to the debug window.
    Debug.Print reportString
    
    'If an output path was specified, write the timing report out to file now
    If Len(dstTextFile) > 0 Then
    
        'If the file already exists, remove it.  (I have also added a DoEvents here after noticing random errors in this
        ' sub - it's important to wait for the file to be deleted, so that the write attempt below does not fail.)
        If FileExist(dstTextFile) Then
            Kill dstTextFile
            DoEvents
        End If
    
        'Open the specified file
        Dim fileNum As Integer
        fileNum = FreeFile
        
        Open dstTextFile For Output As #fileNum
            Print #fileNum, reportString
        Close #fileNum
        
    End If

End Sub

'Sort the profile entries from longest to shortest time taken
Private Sub sortProfileEntries()

    Dim i As Long, j As Long
    
    'Loop through all entries in the profiling array, sorting them as we go
    For i = 0 To m_CurrentEventIndex - 1
        For j = 0 To m_CurrentEventIndex - 1
            
            'Compare two profiled events, and if one took longer than the other, swap them
            If m_Events(i).ElapsedTime > m_Events(j).ElapsedTime Then swapEventData m_Events(i), m_Events(j)
            
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
