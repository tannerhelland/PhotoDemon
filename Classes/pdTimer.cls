VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Class-only Timer
'Copyright 2016-2026 by Tanner Helland
'Created: 07/February/16
'Last updated: 08/February/16
'Last update: added support for coalescing timers on Win 8+
'
'Sometimes you want a timer without an attached UserControl (and all the extra objects that entails).
' This class implements a basic timer message handler, meaning it behaves roughly identically to a
' VB timer control.  On Windows 8 or later, a coalescing timer is preferentially used.
' See go.microsoft.com/fwlink/p/?linkid=246618 for details.
'
'Because AddressOf does not work in a module, this class leans on a few helper functions in the
' VBHacks module.  Also, just like a regular VB Timer, this is a bad solution for high-resolution timing
' or anything that needs to be extremely consistent.
'
'Many thanks to Karl Peterson for additional resources related to callback procs inside VB classes
' (retrieved from http://vb.mvps.org/samples/TimerObj/)
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This class raises just one event: a basic timer event
Public Event Timer()

Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function SetCoalescableTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long, ByVal uToleranceDelay As Long) As Long

'Timer handle.  Pass this value to KillTimer to release the timer.
Private m_TimerID As Long

'Current interval.  Do not set to zero.
Private m_Interval As Long

Private Sub Class_Terminate()
    If (m_TimerID <> 0) Then Me.StopTimer
End Sub

Friend Property Get Interval() As Long
    Interval = m_Interval
End Property

Friend Property Let Interval(ByVal newInterval As Long)
    If (newInterval <> m_Interval) Then
        m_Interval = newInterval
        If (m_TimerID <> 0) Then Me.StartTimer
    End If
End Property

Friend Property Get IsActive() As Boolean
    IsActive = (m_TimerID <> 0)
End Property

Friend Sub StartTimer(Optional ByVal coalescingThreshold As Long = &H0&)

    'Previously (as a failsafe), we would kill old timers before starting new ones, but now that this class
    ' is used for certain animation duties, it's actually preferable to just replace the current timer as-is
    ' (and let Windows handle disposal of the old timer for us).  Note that either way, this does *not* clear
    ' out WM_TIMER messages already in the hWnds queue - that would need to be handled manually.
    'If (m_TimerID <> 0) Then StopTimer
    
    If (m_Interval > 0) Then
        
        If OS.IsWin8OrLater Then
            m_TimerID = SetCoalescableTimer(OS.ThunderMainHWnd, ObjPtr(Me), m_Interval, AddressOf VBHacks.StandInTimerProc, coalescingThreshold)
        Else
            m_TimerID = SetTimer(OS.ThunderMainHWnd, ObjPtr(Me), m_Interval, AddressOf VBHacks.StandInTimerProc)
        End If
        
        If (m_TimerID <> 0) Then UserControls.NotifyTimerCreated
        
    Else
        Debug.Print "WARNING!  Interval values passed to pdTimer must be > zero."
    End If
    
End Sub

Friend Sub StopTimer()

    If (m_TimerID <> 0) Then
        
        Dim backupTimerID As Long
        backupTimerID = ObjPtr(Me)
        
        Dim timerReleaseCheck As Boolean
        timerReleaseCheck = (KillTimer(OS.ThunderMainHWnd, ObjPtr(Me)) <> 0)
        m_TimerID = 0
        
        'Purge all messages remaining for this timer
        VBHacks.PurgeTimerMessagesByID backupTimerID
        
        If timerReleaseCheck Then
            UserControls.NotifyTimerDestroyed
        Else
            PDDebug.LogAction "WARNING!  pdTimer object was not released successfully."
        End If
        
    End If
    
End Sub

'This sub will be called by the VBHacks module whenever a new timer event arrives.
' Note that stopping a timer does *not* remove already-posted WM_TIMER messages from the message queue;
' to prevent those from firing, we check the timer ID here.
Friend Sub TimerEventArrived()
    If (m_TimerID <> 0) Then RaiseEvent Timer
End Sub
