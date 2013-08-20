VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMonitors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Note: this file has been modified for use within PhotoDemon.

'This class was originally written by Mike Dixon.

'You may download the original version of this code from the following link (good as of November '12):
' http://support.microsoft.com/kb/194578


Option Explicit

' --------------------------------------------------------------------------
'               Copyright (C) 1998 Microsoft Corporation                   '
' --------------------------------------------------------------------------
' You have a royalty-free right to use, modify, reproduce and distribute   '
' the Sample Application Files (and/or any modified version) in any way    '
' you find useful, provided that you agree that Microsoft has no warranty, '
' obligations or liability for any Sample Application Files.               '
' --------------------------------------------------------------------------
' Written by Mike Dixon (mikedix@microsoft.com)                            '
' --------------------------------------------------------------------------

'Virtual Desktop sizes
Const SM_XVIRTUALSCREEN = 76    'Virtual Left
Const SM_YVIRTUALSCREEN = 77    'Virtual Top
Const SM_CXVIRTUALSCREEN = 78   'Virtual Width
Const SM_CYVIRTUALSCREEN = 79   'Virtual Height

Const SM_CMONITORS = 80         'Get number of monitors

'Constants for the return value when finding a monitor
Const MONITOR_DEFAULTTONEAREST = &H2    'If the monitor is not found, return the nearest monitor

'Rectangle structure, for determining
'monitors at a given position
Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

'Structure for the position of a monitor
Private Type tagMONITORINFO
    cbSize      As Long 'Size of structure
    rcMonitor   As RECT 'Monitor rect
    rcWork      As RECT 'Working area rect
    dwFlags     As Long 'Flags
End Type

Public Monitors As New Collection

Private Declare Function GetSystemMetrics Lib "user32" ( _
        ByVal nIndex As Long) As Long

'These API's are not present in Pre Windows 98 and
'Pre Windows NT 5 operating systems, you will need
'to trap for errors when using them.
'(Err.Number 453 Can't find DLL entry point...
Private Declare Function GetMonitorInfo Lib "user32" _
        Alias "GetMonitorInfoA" ( _
        ByVal hMonitor As Long, _
        MonInfo As tagMONITORINFO) As Long

Private Declare Function MonitorFromRect Lib "user32" ( _
        rc As RECT, _
        ByVal dwFlags As Long) As Long
        
Private Declare Function GetWindowRect Lib "user32.dll" _
          (ByVal hWnd As Long, _
           rc As RECT) As Long

'==================================================================================================
'Public Members
'==================================================================================================
Private Sub Class_Initialize()
    'Load the monitors collection
    Refresh
End Sub

Public Property Get DesktopLeft() As Long
    DesktopLeft = GetSystemMetrics2(SM_XVIRTUALSCREEN, 0)
End Property

Public Property Get DesktopTop() As Long
    DesktopTop = GetSystemMetrics2(SM_YVIRTUALSCREEN, 0)
End Property

Public Property Get DesktopWidth() As Long
    DesktopWidth = GetSystemMetrics2(SM_CXVIRTUALSCREEN, Screen.Width \ Screen.TwipsPerPixelX)
End Property

Public Property Get DesktopHeight() As Long
    DesktopHeight = GetSystemMetrics2(SM_CYVIRTUALSCREEN, Screen.Height \ Screen.TwipsPerPixelY)
End Property

Public Function GetMonitorFromWindow(hWnd As Long, dwFlags As Long) As Long
    '=====================================================
    'Returns a monitor handle that the Window (hwnd) is in
    '=====================================================
    Dim lReturn As Long
    
    On Error GoTo GetMonitorFromWindow_Err
    'lReturn = MonitorFromWindow(hWnd, dwFlags)
    'FIX BY TANNER: MonitorFromWindow is prone to failing (returns null for no good reason).  Using the window
    ' rect instead ensures accurate results.
    Dim fRect As RECT
    GetWindowRect hWnd, fRect
    lReturn = MonitorFromRect(fRect, MONITOR_DEFAULTTONEAREST)
    GetMonitorFromWindow = lReturn
    Exit Function
GetMonitorFromWindow_Err:
    If Err.Number = 453 Then
        'Non-Multimonitor OS, return -1
        GetMonitorFromWindow = -1
    End If
End Function

Public Function GetMonitorFromRect(ByVal wLeft As Long, ByVal wRight As Long, ByVal wTop As Long, ByVal wBottom As Long) As Long
    '=====================================================
    'Returns a monitor handle that the Window (hwnd) is in
    '=====================================================
    Dim lReturn As Long
    
    Dim rc As RECT
    With rc
        .Left = wLeft
        .Right = wRight
        .Top = wTop
        .Bottom = wBottom
    End With
    
    On Error GoTo GetMonitorFromRect_Err
    lReturn = MonitorFromRect(rc, MONITOR_DEFAULTTONEAREST)
    GetMonitorFromRect = lReturn
    Exit Function
GetMonitorFromRect_Err:
    If Err.Number = 453 Then
        'Non-Multimonitor OS, return -1
        GetMonitorFromRect = -1
    End If
End Function

Public Function CenterFormOnMonitor(FormToCenter As Form, Optional ReferenceForm As Variant, Optional ByVal wLeft As Long, Optional ByVal wRight As Long, Optional ByVal wTop As Long, Optional ByVal wBottom As Long) As Boolean
    '====================================================================
    'Centers the FormToCenter on the monitor that the ReferenceForm is on
    'or the primary monitor if the ReferenceForm is ommited
    '====================================================================
    Dim lMonitor        As Long
    Dim lReturn         As Long
    Dim MonitorInfo     As tagMONITORINFO
    Dim lMonitorWidth   As Long
    Dim lMonitorHeight  As Long
    
    On Error GoTo CenterFormOnMonitor_Err
    
    'Get the handle to the monitor that the reference form is on
    If IsMissing(ReferenceForm) Then
        If IsMissing(wLeft) Then
            lMonitor = GetMonitorFromXYPoint(1, 1, MONITOR_DEFAULTTONEAREST)
        Else
            lMonitor = GetMonitorFromRect(wLeft, wRight, wTop, wBottom)
        End If
    Else
        lMonitor = GetMonitorFromWindow(ReferenceForm.hWnd, MONITOR_DEFAULTTONEAREST)
    End If
    
    'If we get a valid lMonitor
    If lMonitor Then
        
        'init the structure
        MonitorInfo.cbSize = Len(MonitorInfo)
        
        'Get the monitor information
        lReturn = GetMonitorInfo(lMonitor, MonitorInfo)
        'If the Call does not fail then center the form over that monitor
        If lReturn Then
            With MonitorInfo
                lMonitorWidth = (.rcWork.Right - .rcWork.Left) * Screen.TwipsPerPixelX
                lMonitorHeight = (.rcWork.Bottom - .rcWork.Top) * Screen.TwipsPerPixelY
                FormToCenter.Move ((lMonitorWidth - FormToCenter.Width) \ 2) + .rcMonitor.Left * Screen.TwipsPerPixelX, ((lMonitorHeight - FormToCenter.Height) \ 2) + MonitorInfo.rcMonitor.Top * Screen.TwipsPerPixelX
            End With
        End If
    Else
        'There was not a monitor found, center on default screen
        FormToCenter.Move (Screen.Width - FormToCenter.Width) \ 2, (Screen.Height - FormToCenter.Height) \ 2
    End If
    Exit Function
CenterFormOnMonitor_Err:
    If Err.Number = 453 Then
        'Non-Multimonitor OS
        FormToCenter.Move (Screen.Width - FormToCenter.Width) \ 2, (Screen.Width - FormToCenter.Width) \ 2
    End If
End Function

Public Function GetMonitorFromXYPoint(x As Long, y As Long, dwFlags As Long) As Long
    '==========================================
    'Gets a monitor handle from the xy point
    'Workaround for the GetMonitorFromPoint API
    'is to use the GetMonitorFromRect API and
    'build a rect instead
    '==========================================
    Dim lReturn As Long
    Dim rcRect As RECT
    
    'Transfer the x y into a rect 1 pixel square
    With rcRect
        .Top = y
        .Left = x
        .Right = x + 1
        .Bottom = y + 1
    End With
    On Error Resume Next
    lReturn = MonitorFromRect(rcRect, dwFlags)
    If Err.Number = 0 Then
        GetMonitorFromXYPoint = lReturn
    Else
        GetMonitorFromXYPoint = -1
    End If
End Function

Public Sub Refresh()
    '=====================================================
    'Iterate through the Virtual Desktop and enumerate the
    'Monitors that intersect each 640x480 grid section
    '=====================================================
    Dim lMonitors       As Long
    Dim cMonitor        As clsMonitor
    Dim lLoop           As Long
    Dim lLoop2          As Long
    Dim lMonitor        As Long
    
    On Error GoTo Refresh_Err
    
    Set Me.Monitors = Nothing
    
    'Find Out How Many monitors there are
    lMonitors = GetSystemMetrics(SM_CMONITORS)
    
    If lMonitors = 0 Then
        'Non multimonitor OS, just do the screen size
        ClearMonitorsCollection
        Set cMonitor = New clsMonitor
        With cMonitor
            .Handle = 0
            .Bottom = Screen.Height \ Screen.TwipsPerPixelY
            .Left = 0
            .Right = Screen.Width \ Screen.TwipsPerPixelX
            .Top = 0
            .WorkBottom = .Bottom
            .WorkLeft = 0
            .WorkRight = .Right
            .WorkTop = 0
            .Width = .Right
            .Height = .Bottom
        End With
        'Add the monitor to the monitors collection
        Monitors.Add Item:=cMonitor, Key:=CStr(0)
    Else
        
        'Loop through an imaginary grid of 640x480 cells across the virtual desktop
        'testing each for the monitor it is on, then try to add that monitor to the
        'collection, if it fails, it is a duplicate, so just keep going.
        For lLoop = DesktopTop To DesktopHeight Step 480
            For lLoop2 = DesktopLeft To DesktopWidth Step 640
                lMonitor = GetMonitorFromXYPoint(lLoop2 + 320, lLoop + 240, 0)
                If lMonitor <> 0 Then
                    Set cMonitor = New clsMonitor
                    Call GetMonitorInformation(lMonitor, cMonitor)
                    Monitors.Add Item:=cMonitor, Key:=CStr(lMonitor)
                End If
            Next
        Next
    End If
    Exit Sub
Refresh_Err:
    'Duplicate in the collection, so
    'just ignore it and look for the next one
    If Err.Number = 457 Then Resume Next
End Sub

'==================================================================================================
'Private Members
'==================================================================================================
Private Function GetSystemMetrics2(lItem As Long, lDefault As Long) As Long
    '===============================================
    'Calls GetSystemMetrics if multi-monitor capable
    'Otherwise return the default value passed in
    '===============================================
    If GetSystemMetrics(SM_CMONITORS) = 0 Then
        'No multi monitor, return default
        GetSystemMetrics2 = lDefault
    Else
        'Get the desired metric
        GetSystemMetrics2 = GetSystemMetrics(lItem)
    End If
End Function

Private Function GetMonitorInformation(hMonitor As Long, cMon As clsMonitor) As Long
    '======================================================
    'Fills in the cMon class passed in with the information
    '======================================================
    Dim MonitorInfo As tagMONITORINFO
    Dim lReturn     As Long
    
    On Error GoTo GetMonitorInformation_Err
    MonitorInfo.cbSize = Len(MonitorInfo)
    lReturn = GetMonitorInfo(hMonitor, MonitorInfo)
    With cMon
        .Handle = hMonitor
        .Left = MonitorInfo.rcMonitor.Left
        .Right = MonitorInfo.rcMonitor.Right
        .Top = MonitorInfo.rcMonitor.Top
        .Bottom = MonitorInfo.rcMonitor.Bottom
        
        .WorkLeft = MonitorInfo.rcWork.Left
        .WorkRight = MonitorInfo.rcWork.Right
        .WorkTop = MonitorInfo.rcWork.Top
        .WorkBottom = MonitorInfo.rcWork.Bottom
        
        .Height = MonitorInfo.rcMonitor.Bottom - MonitorInfo.rcMonitor.Top
        .Width = MonitorInfo.rcMonitor.Right - MonitorInfo.rcMonitor.Left
    End With
    GetMonitorInformation = lReturn
    Exit Function
GetMonitorInformation_Err:
    If Err.Number = 453 Then
        'Non-Multimonitor OS, return -1
        GetMonitorInformation = -1
    End If
End Function

Private Sub ClearMonitorsCollection()
    '==============================
    'Clears the monitors collection
    '==============================
    Dim lCount      As Long
    Dim lLoop       As Long
    
    lCount = Monitors.Count
    On Error Resume Next
    For lLoop = 0 To lCount Step -1
        Monitors.Remove lLoop
    Next
End Sub

