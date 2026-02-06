VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdDisplays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Display Device Manager (parent class)
'Copyright 2012-2026 by Tanner Helland
'Created: 12/November/2012
'Last updated: 18/September/15
'Last update: add additional failsafes just in case core display APIs fail
'
'As a photo editing program, PD spends a lot of time interacting with display devices.  In early builds, it relied
' on a Microsoft-provided class to manage interactions with multiple displays (http://support.microsoft.com/kb/194578).
' That code was published in 1998 (!!!) so you can imagine that various GDI functions have improved, modified, or
' replaced since then.  That class was also causing problems under Wine, so it needed to be cleaned up and rewritten.
'
'Thus we have pdDisplays.  pdDisplays manages an array of pdDisplay classes.  It generates this list at start-up and
' can be refreshed manually by using the .RefreshDisplays function.  Its goal is simplifying interaction with display
' devices, particularly in a multiple-display environment.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Virtual screen sizes (used with the GetSystemMetrics API)
Private Const SM_XVIRTUALSCREEN As Long = 76
Private Const SM_YVIRTUALSCREEN As Long = 77
Private Const SM_CXVIRTUALSCREEN As Long = 78
Private Const SM_CYVIRTUALSCREEN As Long = 79

'Current system DPI settings (used with the GetDeviceCaps API)
Private Const GDC_LOGPIXELSX As Long = 88
'Private Const GDC_LOGPIXELSY As Long = 90  'Non-square pixels are not currently supported

Private Const DD_PRIMARY_DEVICE As Long = &H4
Private Const CCHDEVICENAME As Long = 32

Private Const MONITOR_DEFAULTTOPRIMARY As Long = 1
Private Const MONITOR_DEFAULTTONEAREST As Long = 2

'GetMonitorInfo struct
Private Type MONITORINFOEX
    cbSize As Long
    rcMonitor As RectL
    rcWork As RectL
    dwFlags As Long
    szDevice(0 To 63) As Byte
End Type

'EnumDisplayDevices struct
'ADDITION BY TANNER: this is helpful for retrieving detailed monitor info, like a human-friendly name and description
Private Type DISPLAY_DEVICEW
   cb As Long
   DeviceName(0 To 63) As Byte
   DeviceString(0 To 255) As Byte
   StateFlags As Long
   DeviceID(0 To 255) As Byte
   DeviceKey(0 To 255) As Byte
End Type

'Some display settings are stored in the registry (e.g. EDID).
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExW" (ByVal hKey As Long, ByVal lpszValueName As Long, ByVal lpReserved As Long, ByRef lpType As Long, ByVal ptrToData As Long, ByRef lpcbData As Long) As Long

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long

Private Declare Function CLSIDFromString Lib "ole32" (ByVal ptrToGuidString As Long, ByVal ptrToByteArray As Long) As Long

Private Type Win32_SP_DEVICE_INTERFACE_DATA
    cbSize As Long
    InterfaceClassGuid(0 To 15) As Byte
    wFlags As Long
    wReserved As Long
End Type

Private Type Win32_SP_DEVINFO_DATA
    cbSize As Long
    ClassGuid(0 To 15) As Byte
    DevInst As Long
    Reserved As Long
End Type

Private Declare Function SetupDiDestroyDeviceInfoList Lib "setupapi" (ByVal hDevInfo As Long) As Long
Private Declare Function SetupDiEnumDeviceInfo Lib "setupapi" (ByVal hDevInfo As Long, ByVal iMemberIndex As Long, ByVal ptrDevInfoData As Long) As Long
Private Declare Function SetupDiEnumDeviceInterfaces Lib "setupapi" (ByVal hDevInfo As Long, ByVal ptrDevInfoData As Long, ByVal ptrInterfaceClassGuid As Long, ByVal iMemberIndex As Long, ByVal ptrDstDevInterfaceData As Long) As Long
Private Declare Function SetupDiGetClassDevs Lib "setupapi" Alias "SetupDiGetClassDevsW" (ByVal ptrClassGuid As Long, ByVal ptrStrEnumerator As Long, ByVal hWndParent As Long, ByVal dwFlags As Long) As Long
Private Declare Function SetupDiGetDeviceInterfaceDetail Lib "setupapi" Alias "SetupDiGetDeviceInterfaceDetailW" (ByVal hDevInfo As Long, ByVal DevInterfaceData As Long, ByVal ptrToDevInterfaceDetailData As Long, ByVal DevInterfaceDetailDataSize As Long, ByRef dstRequiredSize As Long, ByVal DevInfoData As Long) As Long
Private Declare Function SetupDiOpenDevRegKey Lib "setupapi" (ByVal hDevInfo As Long, ByVal ptrDevInfoData As Long, ByVal Scope As Long, ByVal HwProfile As Long, ByVal KeyType As Long, ByVal samDesired As Long) As Long
    
Private Declare Function EnumDisplayDevices Lib "user32" Alias "EnumDisplayDevicesW" (ByVal ptrToDeviceName As Long, ByVal iDevNum As Long, ByRef lpDisplayDevice As DISPLAY_DEVICEW, ByVal dwFlags As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoW" (ByVal hMonitor As Long, ByRef dstMonitorInfo As MONITORINFOEX) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As winRect) As Long
Private Declare Function MonitorFromPoint Lib "user32" (ByVal x As Long, ByVal y As Long, ByVal dwFlags As Long) As Long
Private Declare Function MonitorFromRect Lib "user32" (ByRef srcRect As RectL, ByVal dwFlags As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hndWindow As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

'Number of displays in our collection
Private m_NumOfDisplays As Long

'Display collection.  Each pdDisplay object stores info on its monitor.
Private listOfDisplays() As pdDisplay

Private Sub Class_Initialize()
    
    'This class does not create a list of displays automatically, as the caller may want to perform certain
    ' initialization tasks prior to generating said list.  Thus, you must manually call .RefreshDisplays at
    ' least *once* before attempting to access any child pdDisplay objects.
    m_NumOfDisplays = 0
    ResetDisplays
    
End Sub

'Retrieve virtual desktop dimensions.  These are retrieved on-the-fly, to simplify dealing with changing monitor
' resolutions at runtime.
Friend Function GetDesktopLeft() As Long
    GetDesktopLeft = GetSystemMetrics(SM_XVIRTUALSCREEN)
End Function

Friend Function GetDesktopTop() As Long
    GetDesktopTop = GetSystemMetrics(SM_YVIRTUALSCREEN)
End Function

Friend Function GetDesktopRight() As Long
    GetDesktopRight = Me.GetDesktopLeft + Me.GetDesktopWidth
End Function

Friend Function GetDesktopBottom() As Long
    GetDesktopBottom = Me.GetDesktopTop + Me.GetDesktopHeight
End Function

Friend Function GetDesktopWidth() As Long
    GetDesktopWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN)
End Function

Friend Function GetDesktopHeight() As Long
    GetDesktopHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN)
End Function

Friend Sub GetDesktopRect(ByRef dstRect As RectL)
    
    With dstRect
        .Left = Me.GetDesktopLeft
        .Top = Me.GetDesktopTop
        .Right = Me.GetDesktopLeft + Me.GetDesktopWidth
        .Bottom = Me.GetDesktopTop + Me.GetDesktopHeight
    End With
    
End Sub

'As an absolute final failsafe, VB's Screen object can be queried for a display rect.  This is typically only done under alternative
' environments (e.g. Wine) if all standard display enumeration APIs have failed.
Friend Sub GetVBDesktopRect(ByRef dstRect As RectL, Optional ByVal returnInPixels As Boolean = True)
    
    With dstRect
    
        .Left = 0
        .Top = 0
        
        If returnInPixels Then
            .Right = Screen.Width / InternalTwipsFix
            .Bottom = Screen.Height / InternalTwipsFix
        Else
            .Right = Screen.Width
            .Bottom = Screen.Height
        End If
        
    End With
    
End Sub

'Returns the number of active displays.  Inaccurate until .RefreshDisplays() has been called.
Friend Function GetDisplayCount() As Long
    If (m_NumOfDisplays = 0) Then Me.RefreshDisplays
    GetDisplayCount = m_NumOfDisplays
End Function

'Taskbar height is more complicated in Win 10, since the taskbar can optionally appear across multiple displays.
' At present, PD assumes the taskbar is present on most screens, and it uses this to control things like drop-down height.
Friend Function GetTaskbarHeight() As Long
    Dim primaryDisplayRect As RectL, primaryDisplayWorkingRect As RectL
    Me.PrimaryDisplay.GetRect primaryDisplayRect
    Me.PrimaryDisplay.GetWorkingRect primaryDisplayWorkingRect
    GetTaskbarHeight = (primaryDisplayRect.Bottom - primaryDisplayRect.Top) - (primaryDisplayWorkingRect.Bottom - primaryDisplayWorkingRect.Top)
End Function

'Returns the current screen DPI, *as set in Windows display settings.*  This has no relationship to actual screen DPI,
' which would need to be calculated on a per-monitor basis using the EDID data PD collects.
'
'For convenience, this is returned as a float where...
' 1.0 = 96 DPI (the default Windows setting)
' 2.0 = 200% DPI scaling, etc.
Friend Function GetWindowsDPI() As Double
    
    'Retrieve LogPixelsX via the API; this will be 96 at 100% DPI scaling
    Dim screenDC As Long, logPixelsX As Double
    screenDC = GetDC(0)
    logPixelsX = CDbl(GetDeviceCaps(screenDC, GDC_LOGPIXELSX))
    ReleaseDC 0, screenDC
    
    'Convert that value into a fractional DPI modified (e.g. 1.0 for 100% scaling, 2.0 for 200% scaling)
    If (logPixelsX = 0) Then
        GetWindowsDPI = 1#
        PDDebug.LogAction "WARNING!  System DPI could not be retrieved via pdDisplays.GetWindowsDPI()."
    Else
        GetWindowsDPI = logPixelsX / 96#
    End If
    
End Function

'Erase our current display collection.  Not exposed publicly, as callers should rely on RefreshDisplays, instead.
Private Sub ResetDisplays()
    
    If (m_NumOfDisplays > 0) Then
        
        Dim i As Long
        For i = 0 To m_NumOfDisplays - 1
            Set listOfDisplays(i) = Nothing
        Next i
        
    End If
    
    m_NumOfDisplays = 0
    ReDim listOfDisplays(0) As pdDisplay
    
End Sub

'Refresh the current display list.  Returns the number of displays found.
Friend Function RefreshDisplays() As Long

    'If an existing collection exists, clear it now, then cache the new value
    If (m_NumOfDisplays > 0) Then ResetDisplays
    
    'Retrieve the virtual screen area
    Dim virtualScreenRect As RectL
    Me.GetDesktopRect virtualScreenRect
    
    'We're now going to search the display area, in 320x200 increments, probing for unique monitors.
    ' This is a relatively foolproof way to check for active monitors and their positioning/dimensions.
    Dim x As Long, y As Long, hDisplay As Long
    
    For x = virtualScreenRect.Left To virtualScreenRect.Right Step 320
    For y = virtualScreenRect.Top To virtualScreenRect.Bottom Step 200
        
        'Retrieve a handle to the display containing this point; if no display covers it, revert to the primary display.
        hDisplay = MonitorFromPoint(x, y, MONITOR_DEFAULTTOPRIMARY)
        
        'If the display doesn't exist in our collection, add it now!
        If (hDisplay <> 0) Then
            If (Not DoesDisplayExistInCollection(hDisplay)) Then AddNewDisplay hDisplay
        End If
        
    Next y
    Next x
    
    'If for some reason no displays were found (perhaps possible under Wine or some alternative environment), create a
    ' default display instance.
    If (m_NumOfDisplays = 0) Then
        
        'Get a default screen rect from VB's internal methods
        Dim workingRect As RectL
        Me.GetVBDesktopRect workingRect, True
        
        'Populate basic display features so external functions don't crash
        Set listOfDisplays(0) = New pdDisplay
        With listOfDisplays(0)
            .SetRect workingRect
            .SetWorkingRect workingRect
            .SetAdapterName g_Language.TranslateMessage("Unknown adapter")
            .SetDescription g_Language.TranslateMessage("Unknown display")
            .SetFriendlyName g_Language.TranslateMessage("Unknown display")
        End With
        
        m_NumOfDisplays = 1
        
    End If
    
    'Whenever the display collection is refreshed, add any discoveries to the debug log.
    'If this is a non-production release, note our discovered monitors in the debug log; this can be a big help when tracking
    ' down strange issues.
    If (m_NumOfDisplays > 0) Then
    
        Dim i As Long, prefixText As String
        For i = 0 To m_NumOfDisplays - 1
            
            If listOfDisplays(i).IsPrimary Then
                prefixText = "Found primary monitor: "
            Else
                prefixText = "Found secondary monitor: "
            End If
            
            PDDebug.LogAction prefixText & listOfDisplays(i).GetMonitorSizeAsString & " " & listOfDisplays(i).GetMonitorResolutionAsString & " " & listOfDisplays(i).GetBestMonitorName & " (powered by " & listOfDisplays(i).GetDescription & ")"
            
        Next i
        
    Else
        PDDebug.LogAction "WARNING!  pdDisplays.RefreshDisplays couldn't locate any displays.  This shouldn't be possible - please investigate!"
    End If
    
    PDDebug.LogAction "System-wide DPI currently set to " & Format$(Me.GetWindowsDPI * 100, "#0") & "%"
    
End Function

'Given a display handle (hMonitor), see if it already exists in our display collection.
Private Function DoesDisplayExistInCollection(ByVal hDisplay As Long) As Boolean
        
    If (m_NumOfDisplays = 0) Then
        DoesDisplayExistInCollection = False
    Else
        
        Dim displayFound As Boolean
        displayFound = False
        
        Dim i As Long
        For i = 0 To m_NumOfDisplays - 1
            If (listOfDisplays(i).GetHandle = hDisplay) Then
                displayFound = True
                Exit For
            End If
        Next i
        
        DoesDisplayExistInCollection = displayFound
        
    End If
        
End Function

'Add a new display to the collection.  Do *not* call this without first calling doesDisplayExistInCollection(), above.
Private Sub AddNewDisplay(ByVal hDisplay As Long)

    'Our display collection is 0-based, so we can add new entries immediately, without incrementing our counter.
    Set listOfDisplays(m_NumOfDisplays) = New pdDisplay
    listOfDisplays(m_NumOfDisplays).SetHandle hDisplay
    
    'Retrieve a matching MonitorInfoEx struct for this display, which contains the rect and working rect, among other things.
    Dim monInfo As MONITORINFOEX
    monInfo.cbSize = LenB(monInfo)
    
    If (GetMonitorInfo(hDisplay, monInfo) <> 0) Then
        
        With listOfDisplays(m_NumOfDisplays)
            .SetRect monInfo.rcMonitor
            .SetWorkingRect monInfo.rcWork
        End With
        
        'Extract the adapter name as well
        Dim tmpName As String
        tmpName = Strings.StringFromUTF16_FixedLen(VarPtr(monInfo.szDevice(0)), CCHDEVICENAME * 2, True)
        If (LenB(tmpName) > 0) Then tmpName = Trim$(tmpName)
        listOfDisplays(m_NumOfDisplays).SetAdapterName tmpName
        
        'While here, we want to add some additional information to the display entry, but it's fairly involved, so we're
        ' going to turn control over to a secondary sub for now.
        FillDetailedDisplayInfo m_NumOfDisplays
        
        'Increment the display counter before exiting
        m_NumOfDisplays = m_NumOfDisplays + 1
        If (m_NumOfDisplays > UBound(listOfDisplays)) Then ReDim Preserve listOfDisplays(0 To m_NumOfDisplays * 2 - 1) As pdDisplay
        
    Else
        PDDebug.LogAction "WARNING!  In pdDisplays.addNewDisplay, GetMonitorInfo inexplicably failed on handle " & hDisplay & "."
    End If
    
End Sub

'After adding a new display to the collection, call this to populate detailed info
' (like friendly name, physical dimensions, etc)
Private Sub FillDetailedDisplayInfo(ByVal displayIndex As Long)
    
    On Error GoTo DetailedDisplayInfoBug
    
    'First, we're going to enumerate display devices until we find one that matches the monitor we were passed.
    Dim monitorInfoDetailed As DISPLAY_DEVICEW
    monitorInfoDetailed.cb = LenB(monitorInfoDetailed)
    
    Dim displayCount As Long
    displayCount = 0
    
    Dim tmpName As String, tmpDescription As String
    Do While (EnumDisplayDevices(0&, displayCount, monitorInfoDetailed, 0&) <> 0)
    
        'Extract the monitor's device name (again, it's really the adapter name - thanks, Windows!) and stick
        ' it in a string, so we can compare this device's name against the one we just added.
        tmpName = Strings.StringFromUTF16_FixedLen(VarPtr(monitorInfoDetailed.DeviceName(0)), CCHDEVICENAME * 2, True)
        If (LenB(tmpName) > 0) Then tmpName = Trim$(tmpName)
        
        'If the adapter name of our previous display matches this one, retrieve the extra information
        If Strings.StringsEqual(listOfDisplays(displayIndex).GetAdapterName, tmpName, True) Then
            
            'This is a match!  Grab the description string while we're here.
            tmpDescription = Strings.StringFromUTF16_FixedLen(VarPtr(monitorInfoDetailed.DeviceString(0)), 256, True)
            If (LenB(tmpDescription) > 0) Then tmpDescription = Trim$(tmpDescription)
            listOfDisplays(displayIndex).SetDescription tmpDescription
            
            'Also, if this is the primary monitor, mark it as such
            listOfDisplays(displayIndex).IsPrimary = (monitorInfoDetailed.StateFlags And DD_PRIMARY_DEVICE)
            
            'We've retrieved all we can from this enum.  Exit the do loop now.
            Exit Do
            
        End If
        
        displayCount = displayCount + 1
    
    Loop
    
    'EnumDisplayDevices is a strange API.  You can call it multiple times,
    ' passing in different device-related strings, to retrieve different data.
    ' We're going to do that now, with a string retrieved from the previous enum,
    ' so we can grab a little more data about this display device.
    
    'Repeat the call a second time, providing the monitor's name, in order to receive
    ' even *more* information
    Dim adapterName As String
    adapterName = listOfDisplays(displayIndex).GetAdapterName
    
    Const EDD_GET_DEVICE_INTERFACE_NAME As Long = &H1
    If (EnumDisplayDevices(StrPtr(adapterName), 0&, monitorInfoDetailed, EDD_GET_DEVICE_INTERFACE_NAME) <> 0) Then
        
        'The monitor's "friendly" name should now be stored inside the .DeviceString portion of the struct.
        tmpName = Strings.StringFromUTF16_FixedLen(VarPtr(monitorInfoDetailed.DeviceString(0)), 256, True)
        If (LenB(tmpName) > 0) Then tmpName = Trim$(tmpName)
        listOfDisplays(displayIndex).SetFriendlyName tmpName
        
        'Normally, the .deviceID portion of the monitorInfoDetailed struct is useless,
        ' but MS changed this for Vista.  Now, on a second call, DeviceID is the GUID of the
        ' display and DeviceKey is the registry key.  We can use this to pull an EDID out of
        ' the registry (if one exists), which can then be queried for all kinds of useful info,
        ' like physical screen dimensions.
        
        'Because this process is energy-intensive, we only want to do it once.  At present,
        ' I also limit it to Vista+; rumor has it that the change was introduced at-or-around
        ' XP SP3, but because I can't verify this, it's easier to simply restrict the check
        ' to Vista or later.
        
        'If we have not yet retrieved an EDID value for this monitor, attempt to do so now
        If (Not listOfDisplays(displayIndex).HasEDID) And OS.IsVistaOrLater() Then
        
            'Retrieve and cache the device ID
            Dim devID As String
            devID = Strings.StringFromUTF16_FixedLen(VarPtr(monitorInfoDetailed.DeviceID(0)), 256, True)
            If (LenB(devID) <> 0) Then devID = Trim$(devID)
            listOfDisplays(displayIndex).SetDeviceID devID
            
            'If a device ID was successfully retrieved, use it to try and locate a matching EDID array.
            If (LenB(devID) <> 0) Then
            
                'If successful, the EDID contents will be saved to this byte array
                Dim edidArray() As Byte
                
                'Addendum July 2020: rather than do our own manual registry spelunking, we can use the
                ' SetupAPI family of functions to do the spelunking for us.  Thank you to comments in this
                ' URL for this (much smarter) approach:
                ' https://ofekshilon.com/2014/06/19/reading-specific-monitor-dimensions/
                
                'Importantly, that approach lets us map 1:1 between GDI display objects and their
                ' corresponding setupapi object.
                
                'Start by initializing a device information handle.  GUID comes from MSDN:
                ' https://docs.microsoft.com/en-us/windows-hardware/drivers/install/guid-devinterface-monitor
                Const DIGCF_DEVICEINTERFACE As Long = &H10&
                Dim GUID_DevInterface_Monitor() As Byte
                ReDim GUID_DevInterface_Monitor(0 To 15) As Byte
                CLSIDFromString StrPtr("{E6F07B5F-EE97-4a90-B076-33F57BF4EAA7}"), VarPtr(GUID_DevInterface_Monitor(0))
                
                Dim hDevInfo As Long
                hDevInfo = SetupDiGetClassDevs(VarPtr(GUID_DevInterface_Monitor(0)), 0&, 0&, DIGCF_DEVICEINTERFACE)
                
                'Next, we need to iterate compatible device interfaces.  See MSDN for details:
                ' https://docs.microsoft.com/en-us/windows/win32/api/setupapi/nf-setupapi-setupdienumdeviceinterfaces
                ' GetLastError will return this constant when the list is exhausted:
                'Const ERROR_NO_MORE_ITEMS As Long = &H103&
                '...but we just look for function failure (we don't care why, as this isn't a critical code path)
                Dim tmpInterfaceData As Win32_SP_DEVICE_INTERFACE_DATA
                tmpInterfaceData.cbSize = LenB(tmpInterfaceData)
                
                Dim keyResolved As Boolean
                keyResolved = False
                
                'Interfaces will be manually iterated until we find a match or hit a "no more devices" return
                Dim idxInterface As Long
                idxInterface = 0
                
                Dim retVal As Long
                retVal = SetupDiEnumDeviceInterfaces(hDevInfo, 0&, VarPtr(GUID_DevInterface_Monitor(0)), idxInterface, VarPtr(tmpInterfaceData))
                Do While (retVal <> 0)
                    
                    'Retrieving device info uses a variable-size struct which is inelegant to handle in VB.
                    ' Instead, we just use an array of arbitrary size, with the required buffer size appended
                    ' to the start of the array.
                    
                    'Start by retrieving required buffer size.  (Per MSDN, "This size includes the size of the
                    ' fixed part of the structure plus the number of bytes required for the variable-length device
                    ' path string.)
                    Dim diBufSize As Long
                    diBufSize = 0
                    SetupDiGetDeviceInterfaceDetail hDevInfo, VarPtr(tmpInterfaceData), 0&, 0&, diBufSize, 0&
                    If (diBufSize > 0) Then
                        
                        'Construct the buffer and note the requirements from MSDN; the cbSize member (at the
                        ' head of the struct) needs to contain the *fixed* struct size only - in this case,
                        ' 4 bytes for the size + at least 2 more for a null terminator
                        Dim arbitraryBuffer() As Byte
                        ReDim arbitraryBuffer(0 To diBufSize - 1) As Byte
                        PutMem4 VarPtr(arbitraryBuffer(0)), 6&
                        
                        If (SetupDiGetDeviceInterfaceDetail(hDevInfo, VarPtr(tmpInterfaceData), VarPtr(arbitraryBuffer(0)), diBufSize, 0&, 0&) <> 0) Then
                        
                            'Pull a null-terminated wchar string from the (now-filled) buffer, while obviously
                            ' skipping the DWORD cbSize member at the head
                            Dim tmpString As String
                            tmpString = Strings.StringFromCharPtr(VarPtr(arbitraryBuffer(4)), True)
                            
                            'Look for a match with the device ID we pulled from GDI
                            If Strings.StringsEqual(tmpString, devID, True) Then
                            
                                'Device IDs match!  We can now attempt to pull a corresponding EDID
                                ' from the registry (if one exists; VMs in particular may not supply this).
                                
                                'Before doing anything else, note that we've resolved the device ID.  This breaks
                                ' us out of the setupapi enum loop
                                keyResolved = True
                                
                                Dim tmpDevInfo As Win32_SP_DEVINFO_DATA
                                tmpDevInfo.cbSize = LenB(tmpDevInfo)
                                If (SetupDiEnumDeviceInfo(hDevInfo, idxInterface, VarPtr(tmpDevInfo)) <> 0) Then
                                
                                    'We've got enough info to open a corresponding registry key.
                                    Const DICS_FLAG_GLOBAL As Long = 1
                                    Const DIREG_DEV As Long = 1
                                    Const KEY_READ = (&H20000 Or &H1& Or &H8& Or &H10&) And (Not &H100000)
                                    Dim hEDIDRegKey As Long
                                    hEDIDRegKey = SetupDiOpenDevRegKey(hDevInfo, VarPtr(tmpDevInfo), DICS_FLAG_GLOBAL, 0&, DIREG_DEV, KEY_READ)
                                    
                                    Const INVALID_HANDLE_VALUE As Long = -1
                                    If (hEDIDRegKey <> 0) And (hEDIDRegKey <> INVALID_HANDLE_VALUE) Then
                                    
                                        'Registry key checks out.  Attempt to pull an EDID.
                                        Dim edidSize As Long
                                        If (RegQueryValueEx(hEDIDRegKey, StrPtr("EDID"), 0&, ByVal 0&, ByVal 0&, edidSize) = 0) Then
                                            If (edidSize > 0) Then
                                            
                                                'EDID found.  Attempt to query it.
                                                ReDim edidArray(0 To edidSize - 1) As Byte
                                                If (RegQueryValueEx(hEDIDRegKey, StrPtr("EDID"), 0&, 0&, VarPtr(edidArray(0)), edidSize) = 0) Then
                                                    listOfDisplays(displayIndex).SetEDID edidArray, False
                                                Else
                                                    listOfDisplays(displayIndex).SetEDID edidArray, True
                                                End If
                                                
                                            Else
                                                listOfDisplays(displayIndex).SetEDID edidArray, True
                                                PDDebug.LogAction "WARNING!  pdDisplays found an EDID, but its reported length was zero."
                                            End If
                                        Else
                                            listOfDisplays(displayIndex).SetEDID edidArray, True
                                            PDDebug.LogAction "WARNING!  pdDisplays didn't find an EDID key; this may be a VM."
                                        End If
                                    
                                    Else
                                        PDDebug.LogAction "WARNING!  pdDisplays couldn't open a registry key for this device"
                                        listOfDisplays(displayIndex).SetEDID edidArray, True
                                    End If
                                    
                                    'Regardless of outcome, we must close the initial registry handle before exiting.
                                    ' (Note that this uses regular registry APIs, as there's no special setupapi version:
                                    ' https://docs.microsoft.com/en-us/windows/win32/api/setupapi/nf-setupapi-setupdiopendevregkey)
                                    RegCloseKey hEDIDRegKey
                                
                                Else
                                    PDDebug.LogAction "WARNING!  pdDisplays.FillDetailedDisplayInfo() failed on SetupDiEnumDeviceInfo: " & Err.LastDllError
                                End If
                                
                                'If we're inside this code, it means we aligned our GDI and setupapi objects.
                                ' Exit the iteration immediately.
                                Exit Do
                                
                            End If
                            
                        Else
                            PDDebug.LogAction "WARNING!  pdDisplays.FillDetailedDisplayInfo() failed on SetupDiGetDeviceInterfaceDetail: " & Err.LastDllError
                        End If
                        
                    Else
                        PDDebug.LogAction "WARNING!  pdDisplays.FillDetailedDisplayInfo() failed on SetupDiGetDeviceInterfaceDetail buffer size: " & Err.LastDllError
                    End If
                    
                    'Failsafe key resolution check; should never trigger
                    If keyResolved Then Exit Do
                    
                    'Manually increment interface count and attempt to query again
                    idxInterface = idxInterface + 1
                    retVal = SetupDiEnumDeviceInterfaces(hDevInfo, 0&, VarPtr(GUID_DevInterface_Monitor(0)), idxInterface, VarPtr(tmpInterfaceData))
                    
                Loop
                
                'Regardless of outcome, free all remaining setupapi objects
                If (hDevInfo <> 0) Then SetupDiDestroyDeviceInfoList hDevInfo
                
                'If we didn't find a match, report failure to the debug log
                If (Not keyResolved) Then PDDebug.LogAction "pdDisplays couldn't resolve GDI and setupapi devices.  This PC is likely a VM or an older XP machine."
                
            '/End LenB(devID) <> 0
            End If
        
        '/End previous attempted to cache devID or not Vista+
        End If
        
    '/End second attempt at calling EnumDisplayDevice
    End If
    
    Exit Sub

DetailedDisplayInfoBug:
    PDDebug.LogAction "WARNING!  An error (" & Err.Description & ") occurred in pdDisplays.FillDetailedDisplayInfo.  Some display info may be missing."
    
End Sub

'Shortcut function for those who want to access the primary monitor (which may or may not be index 0 in our collection)
Friend Function PrimaryDisplay() As pdDisplay
    
    If (m_NumOfDisplays > 0) Then
    
        Dim i As Long
        For i = 0 To m_NumOfDisplays - 1
                
            If listOfDisplays(i).IsPrimary Then
                Set PrimaryDisplay = listOfDisplays(i)
                Exit For
            End If
                
        Next i
        
    Else
        Set PrimaryDisplay = Nothing
    End If

End Function

Friend Function GetDisplayByHandle(ByVal srcHMonitor As Long) As pdDisplay
    
    If (m_NumOfDisplays > 0) Then
    
        Dim i As Long
        For i = 0 To m_NumOfDisplays - 1
                
            If (listOfDisplays(i).GetHandle = srcHMonitor) Then
                Set GetDisplayByHandle = listOfDisplays(i)
                Exit For
            End If
                
        Next i
        
    Else
        Set GetDisplayByHandle = Nothing
    End If
    
End Function

Friend Function Displays(ByVal displayIndex As Long) As pdDisplay
    
    If (m_NumOfDisplays > 0) And (displayIndex >= 0) And (displayIndex < m_NumOfDisplays) Then
        Set Displays = listOfDisplays(displayIndex)
    Else
        Set Displays = Nothing
    End If
    
End Function

'Given a rect, find the largest overlapping display device
Friend Function GetHMonitorFromRectL(ByRef srcRect As RectL) As Long
    GetHMonitorFromRectL = MonitorFromRect(srcRect, MONITOR_DEFAULTTONEAREST)
End Function

'Helper function to center a form according to a reference rect.  The reference rect will be used to find the largest overlapping display;
' the form will then be centered on that display.  This is essential on multiple monitor systems, to ensure that a form appears on
' a single monitor, rather than centered on the virtual display (which may lie between two monitors and thus look shitty!).
Friend Function CenterFormViaReferenceRect(ByRef FormToCenter As Form, ByRef srcRect As RectL) As Boolean
    
    On Error GoTo CouldNotCenterForm
    
    'Our goal is to fill this working rect with the working rectangle (e.g. non-chrome) of the monitor containing most of the form.
    Dim workingRect As RectL
    
    'Start by getting the containing display
    Dim hDisplay As Long
    hDisplay = Me.GetHMonitorFromRectL(srcRect)
    
    If (hDisplay <> 0) Then
    
        'Find the matching display in our collection
        Dim srcDisplay As pdDisplay
        Set srcDisplay = GetDisplayByHandle(hDisplay)
        
        'Get the display's working rect
        If (Not srcDisplay Is Nothing) Then
            srcDisplay.GetRect workingRect
        Else
            Me.GetVBDesktopRect workingRect, True
        End If
    
    'hDisplay should always contain a non-zero value, but if it doesn't, we can fall back on internal VB methods as a last resort.
    Else
        Me.GetVBDesktopRect workingRect, True
    End If
    
    'Next, get the window rect of the target form
    Dim srcWinRect As winRect, winWidth As Long, winHeight As Long
    GetWindowRect FormToCenter.hWnd, srcWinRect
    winWidth = srcWinRect.x2 - srcWinRect.x1
    winHeight = srcWinRect.y2 - srcWinRect.y1
    
    'Center the form
    Dim dWidth As Long, dHeight As Long
    With workingRect
        dWidth = (.Right - .Left)
        dHeight = (.Bottom - .Top)
        MoveWindow FormToCenter.hWnd, ((dWidth - winWidth) \ 2) + .Left, ((dHeight - winHeight) \ 2) + .Top, winWidth, winHeight, 0&
    End With
    
    CenterFormViaReferenceRect = True
    Exit Function
    
CouldNotCenterForm:
    CenterFormViaReferenceRect = False
    PDDebug.LogAction "WARNING!  pdDisplays.CenterFormViaReferenceRect() failed for unknown reasons."
    
End Function

'VB truncates its internal TwipsPerPixel measurement, so use this method instead.
Private Function InternalTwipsFix() As Double
    InternalTwipsFix = 15# / Me.GetWindowsDPI()
End Function
