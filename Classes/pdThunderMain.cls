VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdThunderMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'ThunderMain listener
'Copyright 2021-2026 by Tanner Helland
'Created: 18/October/21
'Last updated: 18/October/21
'Last update: initial build
'
'This class has a very simple raison d'être: to subclass ThunderMain and listen for useful messages.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This class caches Query/EndSession messages, which is convenient for querying from arbitrary objects
' (without their own Query/Unload events)
Private Const WM_QUERYENDSESSION As Long = &H11&, WM_ENDSESSION As Long = &H16&
Private m_QueryEndSessionReceived As Boolean, m_EndSessionReceived As Boolean

Implements ISubclass

Friend Sub ResetEndSessionFlags()
    m_QueryEndSessionReceived = False
    m_EndSessionReceived = False
End Sub

Friend Function WasEndSessionReceived(Optional ByVal useQueryEndSessionInstead As Boolean = True) As Boolean
    
    'At present, WM_QUERY/ENDSESSION is tracked as part of automatic session restore after a forced
    ' system reboot.  This feature is only available on Win 7+ (possibly Vista too, but I no longer
    ' keep a Vista VM for testing).  On earlier OS versions, we return FALSE unilaterally.
    If OS.IsWin7OrLater Then
        If useQueryEndSessionInstead Then
            WasEndSessionReceived = m_QueryEndSessionReceived
        Else
            WasEndSessionReceived = m_EndSessionReceived
        End If
    Else
        WasEndSessionReceived = False
    End If
    
End Function

Private Sub Class_Initialize()
    VBHacks.StartSubclassing OS.ThunderMainHWnd(), Me
End Sub

Private Sub Class_Terminate()
    VBHacks.StopSubclassing OS.ThunderMainHWnd, Me
End Sub

Private Function ISubclass_WindowMsg(ByVal hWnd As Long, ByVal uiMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
    
    'Cache relevant flags
    If (uiMsg = WM_QUERYENDSESSION) Then
        m_QueryEndSessionReceived = True
    ElseIf (uiMsg = WM_ENDSESSION) Then
        m_EndSessionReceived = True
    End If
    
    ISubclass_WindowMsg = VBHacks.DefaultSubclassProc(hWnd, uiMsg, wParam, lParam)
    
End Function
