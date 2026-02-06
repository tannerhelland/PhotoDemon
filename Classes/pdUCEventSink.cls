VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdUCEventSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'Dummy event sink for PD's user control engine
'Copyright 2016-2026 by Tanner Helland
'Created: 13/October/16
'Last updated: 13/October/16
'Last update: initial build
'
'The UserControls module handles most generic user control management tasks, but because VB modules can't
' use the WithEvents keyword, this dummy class is used to manually wrap events.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private WithEvents m_TTTimer As pdTimer
Attribute m_TTTimer.VB_VarHelpID = -1

Friend Function IsTTTimerActive() As Boolean
    IsTTTimerActive = False
    If (Not m_TTTimer Is Nothing) Then IsTTTimerActive = m_TTTimer.IsActive
End Function

Friend Sub StartTTTimer(Optional ByVal msTimerInterval As Long = 0&)
    
    If (m_TTTimer Is Nothing) Then Set m_TTTimer = New pdTimer
    
    With m_TTTimer
        .Interval = msTimerInterval
        .StartTimer
    End With
    
End Sub

Friend Sub StopTTTimer()
    If (Not m_TTTimer Is Nothing) Then m_TTTimer.StopTimer
End Sub

Private Sub m_TTTimer_Timer()
    If (Not m_TTTimer Is Nothing) Then
        If m_TTTimer.IsActive Then UserControls.TTTimerFired
    End If
End Sub
