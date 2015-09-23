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
'Copyright 2012-2015 by Tanner Helland
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
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Virtual screen sizes (used with the GetSystemMetrics API)
Private Const SM_XVIRTUALSCREEN As Long = 76
Private Const SM_YVIRTUALSCREEN As Long = 77
Private Const SM_CXVIRTUALSCREEN As Long = 78
Private Const SM_CYVIRTUALSCREEN As Long = 79

'Current system DPI settings (used with the GetDeviceCaps API)
Private Const GDC_VERTRES As Long = 10
Private Const GDC_LOGPIXELSX As Long = 88
Private Const GDC_LOGPIXELSY As Long = 90
Private Const GDC_DESKTOPVERTRES As Long = 117

'GetMonitorInfo struct
Private Const CCHDEVICENAME As Long = 32
Private Type MONITORINFOEX
    cbSize As Long
    rcMonitor As RECTL
    rcWork As RECTL
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

Private Declare Function EnumDisplayDevices Lib "user32" Alias "EnumDisplayDevicesW" (ByVal ptrToDeviceName As Long, ByVal iDevNum As Long, ByRef lpDisplayDevice As DISPLAY_DEVICEW, ByVal dwFlags As Long) As Long
Private Const DD_PRIMARY_DEVICE As Long = &H4

'Monitor interaction APIs
Private Const MONITOR_DEFAULTTONULL As Long = 0
Private Const MONITOR_DEFAULTTOPRIMARY As Long = 1
Private Const MONITOR_DEFAULTTONEAREST As Long = 2
Private Declare Function MonitorFromPoint Lib "user32" (ByVal x As Long, ByVal y As Long, ByVal dwFlags As Long) As Long
Private Declare Function MonitorFromRect Lib "user32" (ByRef srcRect As RECTL, ByVal dwFlags As Long) As Long
Private Declare Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoW" (ByVal hMonitor As Long, ByRef dstMonitorInfo As MONITORINFOEX) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Sub CopyMemoryStrict Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDst As Long, ByVal lpSrc As Long, ByVal byteLength As Long)

'Some display settings are stored in the registry (e.g. EDID).
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExW" (ByVal hKey As Long, ByVal ptrToSubKey As Long, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExW" (ByVal hKey As Long, ByVal lpszValueName As Long, ByVal lpReserved As Long, ByRef lpType As Long, ByVal ptrToData As Long, ByRef lpcbData As Long) As Long

'Number of displays in our collection
Private m_NumOfDisplays As Long

'Display collection.  Each pdDisplay object stores info on its monitor.
Private listOfDisplays() As pdDisplay

'This class now supports Unicode display names, so we inevitably need to work with some Unicode data
Private cUnicode As pdUnicode

Private Sub Class_Initialize()
    
    Set cUnicode = New pdUnicode
    
    'This class does not create a list of monitors automatically, as the caller may want to perform certain
    ' initialization tasks prior to collecting a list of active displays.  Thus, you must manually call
    ' .RefreshDisplays at least once prior to accessing child pdDisplay objects.
    m_NumOfDisplays = 0
    ResetDisplays
    
End Sub

'Retrieve virtual desktop dimensions.  These are retrieved on-the-fly, to simplify dealing with changing monitor
' resolutions at runtime.
Public Function GetDesktopLeft() As Long
    GetDesktopLeft = GetSystemMetrics(SM_XVIRTUALSCREEN)
End Function

Public Function GetDesktopTop() As Long
    GetDesktopTop = GetSystemMetrics(SM_YVIRTUALSCREEN)
End Function

Public Function GetDesktopWidth() As Long
    GetDesktopWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN)
End Function

Public Function GetDesktopHeight() As Long
    GetDesktopHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN)
End Function

Friend Sub GetDesktopRect(ByRef dstRect As RECTL)
    
    With dstRect
        .Left = Me.GetDesktopLeft
        .Top = Me.GetDesktopTop
        .Right = Me.GetDesktopLeft + Me.GetDesktopWidth
        .Bottom = Me.GetDesktopTop + Me.GetDesktopHeight
    End With
    
End Sub

'As an absolute final failsafe, VB's Screen object can be queried for a display rect.  This is typically only done under alternative
' environments (e.g. Wine) if all standard display enumeration APIs have failed.
Friend Sub GetVBDesktopRect(ByRef dstRect As RECTL, Optional ByVal returnInPixels As Boolean = True)
    
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
Public Function GetDisplayCount() As Long
    If m_NumOfDisplays = 0 Then RefreshDisplays
    GetDisplayCount = m_NumOfDisplays
End Function

'Returns the current screen DPI, *as set in Windows display settings.*  This has no relationship to actual screen DPI, which would
' need to be calculated on a per-monitor basis using the EDID data PD collects.
'
'For convenience, this is returned as a float where 1.0 = 96 DPI (the default Windows setting).  2.0 = 200% DPI scaling, etc.
Public Function GetWindowsDPI() As Double
    
    'Retrieve LogPixelsX via the API; this will be 96 at 100% DPI scaling
    Dim screenDC As Long, LogPixelsX As Double
    screenDC = GetDC(0)
    LogPixelsX = CDbl(GetDeviceCaps(screenDC, GDC_LOGPIXELSX))
    ReleaseDC 0, screenDC
    
    'Convert that value into a fractional DPI modified (e.g. 1.0 for 100% scaling, 2.0 for 200% scaling)
    If LogPixelsX = 0 Then
        GetWindowsDPI = 1#
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  System DPI could not be retrieved via pdDisplays.GetWindowsDPI()."
        #End If
    Else
        GetWindowsDPI = LogPixelsX / 96#
    End If
    
End Function

'Erase our current display collection.  Not exposed publicly, as callers should rely on RefreshDisplays, instead.
Private Sub ResetDisplays()
    
    If m_NumOfDisplays > 0 Then
        
        Dim i As Long
        For i = 0 To m_NumOfDisplays - 1
            Set listOfDisplays(i) = Nothing
        Next i
        
    End If
    
    m_NumOfDisplays = 0
    ReDim listOfDisplays(0) As pdDisplay
    
End Sub

'Refresh the current display list.  Returns the number of displays found.
Public Function RefreshDisplays() As Long

    'If an existing collection exists, clear it now, then cache the new value
    If m_NumOfDisplays > 0 Then ResetDisplays
    
    'Retrieve the virtual screen area
    Dim virtualScreenRect As RECTL
    Me.GetDesktopRect virtualScreenRect
    
    'We're now going to search the display area, in 320x200 increments, probing for unique monitors.
    ' This is a relatively foolproof way to check for active monitors and their positioning/dimensions.
    Dim x As Long, y As Long, hDisplay As Long
    
    For x = virtualScreenRect.Left To virtualScreenRect.Right Step 320
    For y = virtualScreenRect.Top To virtualScreenRect.Bottom Step 200
        
        'Retrieve a handle to the display containing this point; if no display covers it, revert to the primary display.
        hDisplay = MonitorFromPoint(x, y, MONITOR_DEFAULTTOPRIMARY)
        
        'If the display doesn't exist in our collection, add it now!
        If hDisplay <> 0 Then
            If Not DoesDisplayExistInCollection(hDisplay) Then AddNewDisplay hDisplay
        End If
        
    Next y
    Next x
    
    'If for some reason no displays were found (perhaps possible under Wine or some alternative environment), create a
    ' default display instance.
    If m_NumOfDisplays = 0 Then
        
        'Get a default screen rect from VB's internal methods
        Dim workingRect As RECTL
        Me.GetVBDesktopRect workingRect, True
        
        'Populate basic display features so external functions don't crash
        Set listOfDisplays(0) = New pdDisplay
        With listOfDisplays(0)
            .SetRect workingRect
            .setWorkingRect workingRect
            .setAdapterName g_Language.TranslateMessage("Unknown adapter")
            .setDescription g_Language.TranslateMessage("Unknown display")
            .setFriendlyName g_Language.TranslateMessage("Unknown display")
        End With
        
        m_NumOfDisplays = 1
        
    End If
    
    'Whenever the display collection is refreshed, add any discoveries to the debug log.
    'If this is a non-production release, note our discovered monitors in the debug log; this can be a big help when tracking
    ' down strange issues.
    #If DEBUGMODE = 1 Then
                
        If m_NumOfDisplays > 0 Then
        
            Dim i As Long, prefixText As String
            For i = 0 To m_NumOfDisplays - 1
                
                If listOfDisplays(i).isPrimary Then
                    prefixText = "Found primary monitor: "
                Else
                    prefixText = "Found secondary monitor: "
                End If
                
                pdDebug.LogAction prefixText & listOfDisplays(i).getMonitorSizeAsString & " " & listOfDisplays(i).getMonitorResolutionAsString & " " & listOfDisplays(i).getBestMonitorName & " (powered by " & listOfDisplays(i).getDescription & ")"
                
            Next i
            
        Else
            pdDebug.LogAction "WARNING!  pdDisplays.RefreshDisplays couldn't locate any displays.  This shouldn't be possible - please investigate!"
        End If
        
        pdDebug.LogAction "System-wide DPI currently set to " & Format$(Me.GetWindowsDPI * 100, "##0") & "%"
              
    #End If
    
End Function

'Given a display handle (hMonitor), see if it already exists in our display collection.
Private Function DoesDisplayExistInCollection(ByVal hDisplay As Long) As Boolean
        
    If m_NumOfDisplays = 0 Then
        DoesDisplayExistInCollection = False
    Else
        
        Dim displayFound As Boolean
        displayFound = False
        
        Dim i As Long
        For i = 0 To m_NumOfDisplays - 1
            If listOfDisplays(i).getHandle = hDisplay Then
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
    listOfDisplays(m_NumOfDisplays).setHandle hDisplay
    
    'Retrieve a matching MonitorInfoEx struct for this display, which contains the rect and working rect, among other things.
    Dim monInfo As MONITORINFOEX
    monInfo.cbSize = LenB(monInfo)
    
    If GetMonitorInfo(hDisplay, monInfo) <> 0 Then
        
        With listOfDisplays(m_NumOfDisplays)
            .SetRect monInfo.rcMonitor
            .setWorkingRect monInfo.rcWork
        End With
        
        'Extract the adapter name as well
        Dim tmpName As String
        tmpName = cUnicode.ConvertUTF16PointerToVBString(VarPtr(monInfo.szDevice(0)), CCHDEVICENAME * 2)
        If Len(tmpName) > 0 Then tmpName = Trim$(tmpName)
        listOfDisplays(m_NumOfDisplays).setAdapterName tmpName
        
        'While here, we want to add some additional information to the display entry, but it's fairly involved, so we're
        ' going to turn control over to a secondary sub for now.
        FillDetailedDisplayInfo m_NumOfDisplays
        
        'Increment the display counter before exiting
        m_NumOfDisplays = m_NumOfDisplays + 1
        If m_NumOfDisplays > UBound(listOfDisplays) Then
            ReDim Preserve listOfDisplays(0 To m_NumOfDisplays * 2 - 1) As pdDisplay
        End If
        
    Else
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  In pdDisplays.addNewDisplay, GetMonitorInfo inexplicably failed on handle " & hDisplay & "."
        #End If
    End If
    
End Sub

'After adding a new display to the collection, call this to populate detailed info (like friendly name, physical dimensions, etc)
Private Sub FillDetailedDisplayInfo(ByVal displayIndex As Long)
    
    On Error GoTo detailedDisplayInfoBug
    
    'First, we're going to enumerate display devices until we find one that matches the monitor we were passed.
    Dim monitorInfoDetailed As DISPLAY_DEVICEW
    monitorInfoDetailed.cb = LenB(monitorInfoDetailed)
    
    Dim displayCount As Long
    displayCount = 0
    
    Dim tmpName As String, tmpDescription As String
    
    Do While EnumDisplayDevices(0&, displayCount, monitorInfoDetailed, 0&) <> 0
    
        'Extract the monitor's device name (again, it's really the adapter name - thanks Windows!) and stick
        ' it in a string, so we can compare this device's name against the one we just added.
        tmpName = cUnicode.ConvertUTF16PointerToVBString(VarPtr(monitorInfoDetailed.DeviceName(0)), CCHDEVICENAME * 2)
        If Len(tmpName) > 0 Then tmpName = Trim$(tmpName)
        
        'If the adapter name of our previous
        If StrComp(LCase(listOfDisplays(displayIndex).getAdapterName), LCase(tmpName), vbBinaryCompare) = 0 Then
            
            'This is a match!  Grab the description string while we're here.
            tmpDescription = cUnicode.ConvertUTF16PointerToVBString(VarPtr(monitorInfoDetailed.DeviceString(0)), 256)
            If Len(tmpDescription) > 0 Then tmpDescription = Trim$(tmpDescription)
            listOfDisplays(displayIndex).setDescription tmpDescription
            
            'Also, if this is the primary monitor, mark it as such
            listOfDisplays(displayIndex).isPrimary = CBool(monitorInfoDetailed.StateFlags And DD_PRIMARY_DEVICE)
            
            'We've retrieved all we can from this enum.  Exit the do loop now.
            Exit Do
            
        End If
        
        displayCount = displayCount + 1
    
    Loop
    
    'EnumDisplayDevices is a strange API.  You can call it multiple times, passing in different device-related strings,
    ' to retrieve different data.  We're going to do that now, with a string retrieved from the previous enum,
    ' so we can grab a little more data about this display device.
    'Repeat the call a second time, providing the monitor's name, in order to receive even *more* information
    Dim adapterName As String
    adapterName = listOfDisplays(displayIndex).getAdapterName
    If EnumDisplayDevices(StrPtr(adapterName), 0&, monitorInfoDetailed, 1&) <> 0 Then
        
        'The monitor's friendly name should now be stored inside the .DeviceString portion of the struct.
        tmpName = cUnicode.ConvertUTF16PointerToVBString(VarPtr(monitorInfoDetailed.DeviceString(0)), 256)
        If Len(tmpName) > 0 Then tmpName = Trim$(tmpName)
        listOfDisplays(displayIndex).setFriendlyName tmpName
        
        'Normally, the .deviceID portion of the monitorInfoDetailed struct is useless, but MS changed this for Vista.
        ' Now, on a second call, DeviceID is the GUID of the display and DeviceKey is the registry key.  We can use
        ' this to pull an EDID out of the registry (if one exists), which can then be queried for all kinds of
        ' useful info, like physical screen dimensions.
        
        'Because this process is energy-intensive, we only want to do it once.  At present, I also limit it to
        ' Vista+; rumor has it that the change was introduced at-or-around XP SP3, but because I can't verify
        ' this, it's easier to simply restrict the functionality to Vista or later.
        
        'If we have not yet retrieved an EDID value for this monitor, attempt to do so now
        If (Not listOfDisplays(displayIndex).hasEDID) And g_IsVistaOrLater Then
        
            'Retrieve and cache the device ID
            Dim devID As String
            devID = cUnicode.ConvertUTF16PointerToVBString(VarPtr(monitorInfoDetailed.DeviceID(0)), 256)
            If Len(devID) <> 0 Then devID = Trim$(devID)
            listOfDisplays(displayIndex).setDeviceID devID
            
            'If a device ID was successfully retrieved, use it to try and locate a matching EDID array.
            If Len(devID) <> 0 Then
            
                'If successful, the EDID contents will be saved to this byte array
                Dim edidArray() As Byte
                
                'We need to parse out various bits of the device ID in order to construct a registry path where the
                ' EDID lies.  The substrings we need are delimited by pound signs, so make sure at least one exists.
                If (InStr(1, devID, "#", vbBinaryCompare) <> 0) Then
                
                    'Parse the device ID into discrete substrings
                    Dim tmpStringArray() As String
                    tmpStringArray = Split(devID, "#")
                    
                    'Make sure we generated enough substrings to correctly generate a registry path
                    If UBound(tmpStringArray) >= 2 Then
                        
                        'Use the substrings to create a path to the EDID location inside the registry
                        Dim regPath As String
                        regPath = "System\CurrentControlSet\Enum\Display\" & tmpStringArray(1) & "\" & tmpStringArray(2) & "\Device Parameters"
                        
                        'Using that path, retrieve this monitor's EDID
                        If RetrieveEDIDFromRegistry(edidArray, regPath) Then
                            listOfDisplays(displayIndex).setEDID edidArray
                        Else
                            listOfDisplays(displayIndex).setEDID edidArray, True
                        End If
                        
                    End If
                    
                Else
                    listOfDisplays(displayIndex).setEDID edidArray, True
                End If
                
            '/End Len(devID) <> 0
            End If
        
        '/End previous attempted to cache devID or not Vista+
        End If
        
    '/End second attempt at calling EnumDisplayDevice
    End If
    
    Exit Sub
    
detailedDisplayInfoBug:

    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "WARNING!  An error (" & Err.Description & ") occurred in pdDisplays.fillDetailedDisplayInfo.  Some display info may be missing."
    #End If
    
End Sub

'Given a registry path (created from a monitor's device ID), retrieve a corresponding EDID.
' Returns: TRUE if the EDID is found, FALSE otherwise.  API errors will be logged to pdDebug.
Private Function RetrieveEDIDFromRegistry(ByRef dstArray() As Byte, ByVal registryPath As String) As Boolean

    On Error GoTo registryFetchError

    'Make sure the array has been dimensioned at least once; that way, even if this function fails, the caller will not
    ' experience OOB errors.
    ReDim dstArray(0) As Byte

    'EDID data is stored in the local machine branch; note that WRITING values here requires admin access, but QUERYING does not.
    ' (We will only be querying.)
    Const HKEY_LOCAL_MACHINE As Long = &H80000002
    
    'To avoid UAC, we must only use QUERY access rights
    Const KEY_QUERY_VALUE As Long = &H1
    
    'Before we can retrieve the EDID, we first have to open its registry key.  Get a handle to the key (if possible).
    Dim regHandle As Long, successCheck As Long
    successCheck = RegOpenKeyEx(HKEY_LOCAL_MACHINE, StrPtr(registryPath), 0&, KEY_QUERY_VALUE, regHandle)
    
    'The registry key was opened successfully.  Use it to query the actual value at that location.
    If (successCheck = 0) And (regHandle <> 0) Then
        
        'Before retrieving the array itself, we need to determine its size.
        Dim regString As String
        regString = "EDID"
        
        Dim edidSize As Long
        successCheck = RegQueryValueEx(regHandle, StrPtr(regString), 0&, ByVal 0&, ByVal 0&, edidSize)

        'Size was retrieved successfully.  Dimension the array and receive the EDID data.
        If (successCheck = 0) And (edidSize > 0) Then
        
            ReDim dstArray(0 To edidSize - 1) As Byte
            successCheck = RegQueryValueEx(regHandle, StrPtr(regString), 0&, 0&, VarPtr(dstArray(0)), edidSize)
            RetrieveEDIDFromRegistry = CBool(successCheck = 0)
            
        Else
        
            'If the key was queried successfully but the EDID size was 0, we don't have a choice but to fail.
            RetrieveEDIDFromRegistry = False
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "WARNING!  pdDisplays.retrieveEDID found the EDID key, but the EDID length was zero."
            #End If
            
        End If
    
    'Couldn't open the registry key.  Exit.
    Else
        
        RetrieveEDIDFromRegistry = False
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  pdDisplays.retrieveEDID couldn't locate the EDID key in the registry.  Sorry."
        #End If
            
    End If
    
    Exit Function
    
registryFetchError:

    'successCheck will contain the last return value of a registry API call; might as well return it in case it's helpful
    RetrieveEDIDFromRegistry = False
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "WARNING!  pdDisplays.retrieveEDID experienced an error.  The last return code from registry APIs was: " & successCheck & ".  The last VB error was: " & Err.Description & "."
    #End If
    
End Function

'Shortcut function for those who want to access the primary monitor (which may or may not be index 0 in our collection)
Public Function PrimaryDisplay() As pdDisplay
    
    If m_NumOfDisplays > 0 Then
    
        Dim i As Long
        For i = 0 To m_NumOfDisplays - 1
                
            If listOfDisplays(i).isPrimary Then
                Set PrimaryDisplay = listOfDisplays(i)
                Exit For
            End If
                
        Next i
        
    Else
        Set PrimaryDisplay = Nothing
    End If

End Function

Public Function GetDisplayByHandle(ByVal srcHMonitor As Long) As pdDisplay
    
    If m_NumOfDisplays > 0 Then
    
        Dim i As Long
        For i = 0 To m_NumOfDisplays - 1
                
            If listOfDisplays(i).getHandle = srcHMonitor Then
                Set GetDisplayByHandle = listOfDisplays(i)
                Exit For
            End If
                
        Next i
        
    Else
        Set GetDisplayByHandle = Nothing
    End If
    
End Function

Public Function Displays(ByVal displayIndex As Long) As pdDisplay
    
    If (m_NumOfDisplays > 0) And (displayIndex >= 0) And (displayIndex < m_NumOfDisplays) Then
        Set Displays = listOfDisplays(displayIndex)
    Else
        Set Displays = Nothing
    End If
    
End Function

'Given a rect, find the largest overlapping display device
Friend Function GetHMonitorFromRectL(ByRef srcRect As RECTL) As Long
    GetHMonitorFromRectL = MonitorFromRect(srcRect, MONITOR_DEFAULTTONEAREST)
End Function

'Helper function to center a form according to a reference rect.  The reference rect will be used to find the largest overlapping display;
' the form will then be centered on that display.  This is essential on multiple monitor systems, to ensure that a form appears on
' a single monitor, rather than centered on the virtual display (which may lie between two monitors and thus look shitty!).
Friend Function CenterFormViaReferenceRect(ByRef FormToCenter As Form, ByRef srcRect As RECTL) As Boolean
    
    On Error GoTo couldNotCenterForm
    
    'Our goal is to fill this working rect with the working rectangle (e.g. non-chrome) of the monitor containing most of the form.
    Dim workingRect As RECTL
    
    'Start by getting the containing display
    Dim hDisplay As Long
    hDisplay = Me.GetHMonitorFromRectL(srcRect)
    
    If hDisplay <> 0 Then
    
        'Find the matching display in our collection
        Dim srcDisplay As pdDisplay
        Set srcDisplay = GetDisplayByHandle(hDisplay)
        
        'Get the display's working rect
        If Not srcDisplay Is Nothing Then
            srcDisplay.getRect workingRect
        Else
            Me.GetVBDesktopRect workingRect, True
        End If
    
    'hDisplay should always contain a non-zero value, but if it doesn't, we can fall back on internal VB methods as a last resort.
    Else
        Me.GetVBDesktopRect workingRect, True
    End If
    
    'Center the form
    Dim dWidth As Long, dHeight As Long
    With workingRect
        dWidth = (.Right - .Left) * InternalTwipsFix
        dHeight = (.Bottom - .Top) * InternalTwipsFix
        FormToCenter.Move ((dWidth - FormToCenter.Width) \ 2) + .Left * InternalTwipsFix, ((dHeight - FormToCenter.Height) \ 2) + .Top * InternalTwipsFix
    End With
    
    CenterFormViaReferenceRect = True
    Exit Function
    
couldNotCenterForm:
    
    CenterFormViaReferenceRect = False
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "WARNING!  pdDisplays.CenterFormViaReferenceRect() failed for unknown reasons."
    #End If
    
End Function

'VB truncates its internal TwipsPerPixel measurement, so use this method instead.
Private Function InternalTwipsFix() As Double
    InternalTwipsFix = 15# / Me.GetWindowsDPI()
End Function
