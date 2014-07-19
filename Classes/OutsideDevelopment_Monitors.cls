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
'Note: this file has been heavily modified for use within PhotoDemon.

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
Private Const MONITOR_DEFAULTTONEAREST As Long = &H2    'If the monitor is not found, return the nearest monitor

'Rectangle structure, for determining
'monitors at a given position
Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

'Structure for the position of a monitor
Private Const CCHDEVICENAME As Long = 32

Private Type tagMONITORINFO
    cbSize      As Long 'Size of structure
    rcMonitor   As RECT 'Monitor rect
    rcWork      As RECT 'Working area rect
    dwFlags     As Long 'Flags
    b(0 To CCHDEVICENAME - 1) As Byte 'Monitor name
End Type

Public Monitors As New Collection

'Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

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
           
'ADDITION BY TANNER: this is helpful for retrieving detailed monitor info, like a human-friendly name and description
Private Type DISPLAY_DEVICE
   cb As Long
   DeviceName(0 To 31) As Byte
   DeviceString(0 To 127) As Byte
   StateFlags As Long
   DeviceID(0 To 127) As Byte
   DeviceKey(0 To 127) As Byte
End Type

Private Declare Function EnumDisplayDevices Lib "user32" _
   Alias "EnumDisplayDevicesA" _
  (ByVal lpDevice As Any, _
   ByVal iDevNum As Long, _
   lpDisplayDevice As DISPLAY_DEVICE, _
   ByVal dwFlags As Long) As Long
   
Private Const DD_PRIMARY_DEVICE As Long = &H4

'MORE ADDITIONS BY TANNER: we can use the registry to retrieve a monitor's EDID value, which in turn can be used to
' retrieve things like physical dimensions.
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpszValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByRef lpData As Any, ByRef lpcbData As Long) As Long


'==================================================================================================
'Public Members
'==================================================================================================
Private Sub Class_Initialize()

    'Load the monitors collection
    Refresh
    
    'If this is a non-production release, note our discovered monitors in the debug log; this can be a big help when tracking
    ' down strange issues.
    #If DEBUGMODE = 1 Then
                
        Dim i As Long, prefixText As String
        
        For i = 1 To Monitors.Count
            
            If Monitors(i).isPrimary Then
                prefixText = "Found primary monitor: "
            Else
                prefixText = "Found secondary monitor: "
            End If
            
            pdDebug.LogAction prefixText & Monitors(i).getMonitorSizeAsString & " " & Monitors(i).getMonitorResolutionAsString & " " & Monitors(i).getBestMonitorName & " (powered by " & Monitors(i).Description & ")"
            
        Next i
              
    #End If
                    
End Sub

Public Property Get DesktopLeft() As Long
    DesktopLeft = GetSystemMetrics2(SM_XVIRTUALSCREEN, 0)
End Property

Public Property Get DesktopTop() As Long
    DesktopTop = GetSystemMetrics2(SM_YVIRTUALSCREEN, 0)
End Property

Public Property Get DesktopWidth() As Long
    DesktopWidth = GetSystemMetrics2(SM_CXVIRTUALSCREEN, Screen.Width \ TwipsPerPixelXFix)
End Property

Public Property Get DesktopHeight() As Long
    DesktopHeight = GetSystemMetrics2(SM_CYVIRTUALSCREEN, Screen.Height \ TwipsPerPixelYFix)
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
    Dim monitorInfo     As tagMONITORINFO
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
        monitorInfo.cbSize = Len(monitorInfo)
        
        'Get the monitor information
        lReturn = GetMonitorInfo(lMonitor, monitorInfo)
        'If the Call does not fail then center the form over that monitor
        If lReturn Then
            With monitorInfo
                lMonitorWidth = (.rcWork.Right - .rcWork.Left) * TwipsPerPixelXFix
                lMonitorHeight = (.rcWork.Bottom - .rcWork.Top) * TwipsPerPixelYFix
                FormToCenter.Move ((lMonitorWidth - FormToCenter.Width) \ 2) + .rcMonitor.Left * TwipsPerPixelXFix, ((lMonitorHeight - FormToCenter.Height) \ 2) + monitorInfo.rcMonitor.Top * TwipsPerPixelXFix
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
            .Bottom = Screen.Height \ TwipsPerPixelYFix
            .Left = 0
            .Right = Screen.Width \ TwipsPerPixelXFix
            .Top = 0
            .WorkBottom = .Bottom
            .WorkLeft = 0
            .WorkRight = .Right
            .WorkTop = 0
            .Width = .Right
            .Height = .Bottom
        End With
        'Add the monitor to the monitors collection
        Monitors.Add Item:=cMonitor, Key:=Str(0)
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
                    Monitors.Add Item:=cMonitor, Key:=Str(lMonitor)
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
    Dim monitorInfo As tagMONITORINFO
    Dim lReturn     As Long
    
    On Error GoTo GetMonitorInformation_Err
    monitorInfo.cbSize = Len(monitorInfo)
    lReturn = GetMonitorInfo(hMonitor, monitorInfo)
    With cMon
        .Handle = hMonitor
        .Name = TrimNull(StrConv(monitorInfo.b, vbUnicode))
        .Left = monitorInfo.rcMonitor.Left
        .Right = monitorInfo.rcMonitor.Right
        .Top = monitorInfo.rcMonitor.Top
        .Bottom = monitorInfo.rcMonitor.Bottom
        
        .WorkLeft = monitorInfo.rcWork.Left
        .WorkRight = monitorInfo.rcWork.Right
        .WorkTop = monitorInfo.rcWork.Top
        .WorkBottom = monitorInfo.rcWork.Bottom
        
        .Height = monitorInfo.rcMonitor.Bottom - monitorInfo.rcMonitor.Top
        .Width = monitorInfo.rcMonitor.Right - monitorInfo.rcMonitor.Left
    End With
    
    'ADDED BY TANNER: use the device name to retrieve even more helpful information, like a human-friendly
    ' name and description.
    Dim monitorInfoDetailed As DISPLAY_DEVICE
    monitorInfoDetailed.cb = Len(monitorInfoDetailed)
    
    'Enumerate display devices until we find one that matches this monitor
    Dim displayCount As Long
    displayCount = 0
    
    Do While EnumDisplayDevices(0&, displayCount, monitorInfoDetailed, 0&)
    
        'A match was found; retrieve the display driver's description
        If cMon.Name = TrimNull(StrConv(monitorInfoDetailed.DeviceName, vbUnicode)) Then
            cMon.Description = TrimNull(StrConv(monitorInfoDetailed.DeviceString, vbUnicode))
            
            'Also, if this is the primary monitor, mark it as such
            If (monitorInfoDetailed.StateFlags And DD_PRIMARY_DEVICE) Then
                cMon.isPrimary = True
            Else
                cMon.isPrimary = False
            End If
            
            Exit Do
        End If
        
        displayCount = displayCount + 1
    
    Loop
    
    'Repeat the call a second time, providing the monitor's name, in order to receive even *more* information
    If EnumDisplayDevices(cMon.Name, 0, monitorInfoDetailed, 1&) <> 0 Then
    
        cMon.friendlyName = TrimNull(StrConv(monitorInfoDetailed.DeviceString, vbUnicode))
        
        'If we have not yet retrieved an EDID value for this monitor, attempt to do so now
        If (Not cMon.hasEDID) And g_IsVistaOrLater Then
        
            'We can also retrieve a DeviceID from this second instance, which in turn can be translated to a registry path
            ' containing the monitor's EDID (which contains a *wealth* of information)
            Dim devID As String
            devID = TrimNull(StrConv(monitorInfoDetailed.DeviceID, vbUnicode))
            
            'If successful, the EDID contents will be saved to this byte array
            Dim edidArray() As Byte
            
            'If a device ID was successfully retrieved (this should always happen), attempt to retrieve a matching EDID
            If (Len(devID) > 0) And (InStr(1, devID, "#") > 0) Then
            
                'Convert the device ID to a registry path
                Dim regPath As String, tmpStringArray() As String
                tmpStringArray = Split(devID, "#")
                regPath = "System\CurrentControlSet\Enum\Display\" & tmpStringArray(1) & "\" & tmpStringArray(2) & "\Device Parameters"
                
                'Using that path, retrieve this monitor's EDID from the registry
                Dim regSuccess As Long
                
                regSuccess = retrieveEDIDFromRegistry(edidArray, regPath)
                
                'If successful, copy the contents of the EDID array into this monitor object.
                If regSuccess = 0 Then
                    cMon.setEDID edidArray
                Else
                    cMon.setEDID edidArray, True
                End If
                
            Else
                cMon.setEDID edidArray, True
            End If
        
        End If
        
    End If
        
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

'Given a registry path (created from a monitor's device ID), retrieve the corresponding EDID value.  If the EDID is found, this
' will return 0, otherwise it will return the registry API's error code (or 13 if a non-specific error occurred).
Private Function retrieveEDIDFromRegistry(ByRef dstArray() As Byte, ByVal registryPath As String) As Long

    On Error GoTo registryFetchError

    'Make sure the array has been dimensioned at least once; that way, even if this function fails, the caller will not
    ' experience OOB errors.
    ReDim dstArray(0) As Byte

    'EDID data is stored in the local machine branch; note that WRITING values here requires admin access, but QUERYING does not.
    ' (we will only be querying)
    Const HKEY_LOCAL_MACHINE As Long = &H80000002
    
    'To avoid UAC, we must only use QUERY access rights
    Const KEY_QUERY_VALUE As Long = &H1

    Dim successCheck As Long
    
    'Before we can retrieve the EDID, we first have to open its registry key.  Get a handle to the key (if possible).
    Dim regHandle As Long
    successCheck = RegOpenKeyEx(HKEY_LOCAL_MACHINE, registryPath, 0&, KEY_QUERY_VALUE, regHandle)
    
    'The registry key was opened successfully.  Use it to query the actual value at that location.
    If (successCheck = 0) And (regHandle <> 0) Then
    
        'EDIDs are always stored as BINARY type
        Const REG_BINARY As Long = 3
        
        'Before retrieving the array itself, we need to find out its size.
        Dim edidSize As Long
        successCheck = RegQueryValueEx(regHandle, "EDID", 0&, REG_BINARY, ByVal 0&, edidSize)

        'Size was retrieved successfully.  Dimension the array and receive the EDID data.
        If (successCheck = 0) And (edidSize > 0) Then
        
            ReDim dstArray(0 To edidSize - 1) As Byte
            successCheck = RegQueryValueEx(regHandle, "EDID", 0&, REG_BINARY, dstArray(0), edidSize)
            
        Else
        
            'If the key was queried successfully but the EDID size was 0 (note: this should never happen), return
            ' an arbitrary failure number.
            If successCheck = 0 Then successCheck = 13
            
        End If
    
    'Couldn't open the registry key.  Exit.
    Else
    
        'If the registry key opened successfully but something else went wrong, assign an arbitrary failure number
        If successCheck = 0 Then successCheck = 13
            
    End If
    
    retrieveEDIDFromRegistry = successCheck
    
    Exit Function
    
registryFetchError:

    'Return the API fail state, or an arbitrary number if no fail state occurred but something else went wrong
    If successCheck <> 0 Then retrieveEDIDFromRegistry = successCheck Else retrieveEDIDFromRegistry = 13

End Function
