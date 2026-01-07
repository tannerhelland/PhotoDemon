VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdTimerCountdown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Countdown Timer (modified wrapper around pdTimer, targeting precise countdowns)
'Copyright 2020-2026 by Tanner Helland
'Created: 24/July/20
'Last updated: 24/July/20
'Last update: split off from pdTimer.  That class should stay as lightweight as possible.
'
'This class is just a minor wrapper around the core pdTimer class.  This is designed to make
' countdowns of some specific duration easier, so clients don't need to wrap a bunch of code
' around timer instances themselves.
'
'Note that the API SetWaitableTimer (with an asynchronous procedure call) is potentially a
' smarter replacement for this class.  I haven't wrapped it as it requires additional support
' code inside a module, which I don't want to assemble right now.  Perhaps in the future I'll
' take a closer look at using that API instad.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Public Event CountdownFinished()
Public Event UpdateTimeRemaining(ByVal timeRemainingInMS As Long)

'A base pdTimer object is used to raise interval events
Private WithEvents m_Timer As pdTimer
Attribute m_Timer.VB_VarHelpID = -1

'Callers get to set two properties: when the countdown finishes, and how often to raise
' intermediary "time remaining" events.  These can be the same value.  If they are *not*
' the sam evalue, the countdown time should be a perfect (or near-perfect) multiple of
' the interval time.
Private m_IntervalInMS As Long, m_CountdownTimeInMS As Long

'Target time, in MS; we set this when the countdown is started
Private m_TargetTimeInMS As Currency

'Last-reported time.  Because timer intervals are not perfect, before raising a
' "time remaining" event, we compare the current time remaining to the last-reported
' time remaining.  An event is only raised if they do not match.
Private m_LastTimeRemaining As Long

Friend Function GetIntervalTime() As Long
    GetIntervalTime = m_IntervalInMS
End Function

Friend Sub SetIntervalTimeInMS(ByVal newInterval As Long)
    m_IntervalInMS = newInterval
End Sub

Friend Function GetCountdownTime() As Long
    GetCountdownTime = m_CountdownTimeInMS
End Function

Friend Sub SetCountdownTimeInMS(ByVal newTime As Long)
    m_CountdownTimeInMS = newTime
End Sub

Friend Sub StartCountdown()
    If (m_CountdownTimeInMS <> 0) And (m_IntervalInMS <> 0) Then
        m_Timer.Interval = m_IntervalInMS
        m_TargetTimeInMS = VBHacks.GetHighResTimeInMSEx() + m_CountdownTimeInMS
        m_Timer.StartTimer
    End If
End Sub

Private Sub Class_Initialize()
    Set m_Timer = New pdTimer
End Sub

Private Sub Class_Terminate()
    If (Not m_Timer Is Nothing) Then m_Timer.StopTimer
End Sub

Private Sub m_Timer_Timer()
    
    'Figure out how much time remains
    Dim curTimeRemaining As Long
    curTimeRemaining = m_TargetTimeInMS - VBHacks.GetHighResTimeInMSEx()
    
    'Crop time remaining to the specified time interval
    curTimeRemaining = Int(curTimeRemaining / m_IntervalInMS + 0.5) * m_IntervalInMS
    
    'If the time remaining is <= 0 (or less than an arbitrary delta away from that),
    ' raise the "time's up" event.
    If (curTimeRemaining < 16) Then
    
        RaiseEvent CountdownFinished
        m_Timer.StopTimer
    
    'Otherwise, provide a new "update" event
    Else
        
        If (curTimeRemaining <> m_LastTimeRemaining) Then
            m_LastTimeRemaining = curTimeRemaining
            RaiseEvent UpdateTimeRemaining(curTimeRemaining)
        End If
        
    End If
    
End Sub
