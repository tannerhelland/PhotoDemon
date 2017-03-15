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
'Copyright 2016-2017 by Tanner Helland
'Created: 07/February/16
'Last updated: 08/February/16
'Last update: added support for coalescing timers on Win 8+
'
'Sometimes you want a timer without an attached UserControl (and all the extra objects that entails).  This class
' implements a basic timer message handler, meaning it behaves roughly identically to a VB timer control.  On Windows 8
' or later, a coalescing timer is preferentially used.  See go.microsoft.com/fwlink/p/?linkid=246618 for details.
'
'Because AddressOf does not work in a module, this class leans on a few helper functions in the VBHacks module.
'Also, just like a regular VB Timer, this is a bad solution for high-resolution timing or anything that needs to be
' extremely consistent.
'
'Many thanks to Karl Peterson for additional resources related to callback procs inside VB classes
' (retrieved from http://vb.mvps.org/samples/TimerObj/)
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
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

'Copy of the original PD owner window (ThunderMain)
Private m_OwnerHWnd As Long

'Current interval.  Do not set to zero.
Private m_Interval As Long

Private Sub Class_Initialize()
    m_OwnerHWnd = VBHacks.GetThunderMainHWnd()
End Sub

Private Sub Class_Terminate()
    If (m_TimerID <> 0) Then StopTimer
End Sub

Public Property Get Interval() As Long
    Interval = m_Interval
End Property

Public Property Let Interval(ByVal newInterval As Long)
    If (newInterval <> m_Interval) Then
        m_Interval = newInterval
        If (m_TimerID <> 0) Then StartTimer
    End If
End Property

Public Property Get IsActive() As Boolean
    IsActive = CBool(m_TimerID <> 0)
End Property

Public Sub StartTimer()

    If (m_TimerID <> 0) Then StopTimer
    
    If (m_Interval > 0) Then
        
        If g_IsWin8OrLater Then
            m_TimerID = SetCoalescableTimer(m_OwnerHWnd, ObjPtr(Me), m_Interval, AddressOf VBHacks.StandInTimerProc, 0&)
        Else
            m_TimerID = SetTimer(m_OwnerHWnd, ObjPtr(Me), m_Interval, AddressOf VBHacks.StandInTimerProc)
        End If
        
        If (m_TimerID <> 0) Then UserControl_Support.NotifyTimerCreated
        
    Else
        Debug.Print "WARNING!  Interval values passed to pdTimer must be > zero."
    End If
    
End Sub

Public Sub StopTimer()

    If (m_TimerID <> 0) Then
    
        Dim timerReleaseCheck As Boolean
        timerReleaseCheck = CBool(KillTimer(m_OwnerHWnd, ObjPtr(Me)) <> 0)
        m_TimerID = 0
        
        If timerReleaseCheck Then
            UserControl_Support.NotifyTimerDestroyed
        Else
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "WARNING!  pdTimer object was not released successfully."
            #End If
        End If
        
    End If
    
End Sub

'This sub will be called by the VBHacks module whenever a new timer event arrives.  Note that stopping a timer does not
' remove already-posted WM_TIMER messages from the message queue; to prevent those from firing, we check the timer ID here.
Public Sub TimerEventArrived()
    If (m_TimerID <> 0) Then RaiseEvent Timer
End Sub
