VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Display Device Manager (child class)
'Copyright 2012-2015 by Tanner Helland
'Created: 12/November/12
'Last updated: 18/September/15
'Last update: add additional failsafes just in case core display APIs fail
'
'pdDisplay manages data for a single display instance.  Look inside the parent class, pdDisplays, for more
' detailed comments on how this class is populated.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Declare Sub CopyMemoryStrict Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDst As Long, ByVal lpSrc As Long, ByVal byteLength As Long)

'The monitor's handle (hMonitor)
Private m_hMonitor As Long

'Is this the primary display device?
Private m_Primary As Boolean

'Various names and descriptors associated with this display device
Private m_AdapterName As String, m_FriendlyName As String, m_Description As String, m_DeviceID As String

'Display rect and working rect.  (Working rect does not include chrome like taskbars.)
Private m_MonitorRect As RECTL, m_WorkingRect As RECTL

'Settings related to the monitor's EDID.  These are only available on Vista+.
Private m_hasEDID As Boolean
Private m_EDID() As Byte
Private m_EDIDWidthMM As Long
Private m_EDIDHeightMM As Long
Private m_EDIDWidthInches As Double
Private m_EDIDHeightInches As Double
Private m_EDIDDiagonalInches As Double
Private m_EDIDModelName As String
Private m_EDIDSerialNumber As String
Private m_EDIDNativeResH As Long
Private m_EDIDNativeResV As Long

Public Function getHandle() As Long
    getHandle = m_hMonitor
End Function

Public Sub setHandle(ByVal newHandle As Long)
    m_hMonitor = newHandle
End Sub

Public Function getAdapterName() As String
    getAdapterName = m_AdapterName
End Function

Public Sub setAdapterName(ByRef newName As String)
    m_AdapterName = newName
End Sub

Public Function getFriendlyName() As String
    getFriendlyName = m_FriendlyName
End Function

Public Sub setFriendlyName(ByRef newName As String)
    m_FriendlyName = newName
End Sub

Public Function getDescription() As String
    getDescription = m_Description
End Function

Public Sub setDescription(ByRef newDescription As String)
    m_Description = newDescription
End Sub

Public Function getDeviceID() As String
    getDeviceID = m_DeviceID
End Function

Public Sub setDeviceID(ByRef newDeviceID As String)
    m_DeviceID = newDeviceID
End Sub

Friend Sub getRect(ByRef dstRect As RECTL)
    dstRect = m_MonitorRect
End Sub

Friend Sub SetRect(ByRef srcRect As RECTL)
    m_MonitorRect = srcRect
End Sub

Friend Sub getWorkingRect(ByRef dstRect As RECTL)
    dstRect = m_WorkingRect
End Sub

Friend Sub setWorkingRect(ByRef srcRect As RECTL)
    m_WorkingRect = srcRect
End Sub

Public Property Get isPrimary() As Boolean
    isPrimary = m_Primary
End Property

Public Property Let isPrimary(lPrimary As Boolean)
    m_Primary = lPrimary
End Property

'Retrieve the best available name for this monitor.  If available, an EDID model name will be used.  If that doesn't
' exist, we'll use the Windows fallback (which is often something dumb like "Generic PnP Monitor").
Public Function getBestMonitorName() As String
    
    If Len(m_EDIDModelName) <> 0 Then
        getBestMonitorName = m_EDIDModelName
    Else
        getBestMonitorName = m_FriendlyName
    End If
    
End Function

'Retrieve the monitor's size, as a string (useful for displaying to the user)
Public Function getMonitorSizeAsString() As String

    If m_hasEDID Then
        getMonitorSizeAsString = Format$(m_EDIDDiagonalInches, "#0.0") & """"
    Else
        getMonitorSizeAsString = ""
    End If

End Function

'Retrieve the monitor's native resolution, as a string (useful for displaying to the user)
Public Function getMonitorResolutionAsString() As String

    If m_hasEDID Then
        getMonitorResolutionAsString = CStr(m_EDIDNativeResH) & "x" & CStr(m_EDIDNativeResV)
    Else
        getMonitorResolutionAsString = CStr(m_MonitorRect.Right - m_MonitorRect.Left) & "x" & CStr(m_MonitorRect.Bottom - m_MonitorRect.Top)
    End If

End Function

Public Function hasEDID() As Boolean
    hasEDID = m_hasEDID
End Function

'At creation time, the clsMonitors class will attempt to retrieve an EDID for this monitor.  Regardless of success
' or failure, it will call this function.
Public Sub setEDID(ByRef edidArray() As Byte, Optional ByVal warningEDIDMissing As Boolean = False)

    'If we already have an EDID, exit.  (This shouldn't be possible, but oh well.)
    If m_hasEDID Then Exit Sub
    
    'If no EDID could be found, mark the failure state and exit.
    If warningEDIDMissing Then
        
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING: Monitor EDID could not be retrieved; physical dimensions unknown for this session."
        #End If
        
        m_hasEDID = False
        Exit Sub
        
    End If
    
    'If we made it here, assume the EDID was gathered successfully.  Make a local copy of it.
    m_hasEDID = True
    
    Dim EDIDLength As Long
    EDIDLength = UBound(edidArray) + 1
    ReDim m_EDID(0 To EDIDLength - 1) As Byte
    CopyMemoryStrict VarPtr(m_EDID(0)), VarPtr(edidArray(0)), EDIDLength
    
    'EDIDs should never be shorter than 128 bytes.  Newer versions of the space may be *longer* than 128, but 128 is the minimum.
    If EDIDLength < 128 Then
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "WARNING!  This monitor's EDID is less than 128 bytes.  This shouldn't be possible - investigate??"
        #End If
    End If
    
    'With the EDID successfully retrieved, attempt to parse it for useful values.
    
    'First, note that we're going to be accessing specific locations inside the EDID.  If these locations don't exist
    ' (which should never happen, but these are hardware manufacturers we're talking about!), we have no choice but
    ' to prematurely bail.
    
    'The first spot we need to access is (68), which gives us the physical dimensions.
    If EDIDLength >= 68 Then
    
        'Start with the monitor's physical dimensions, specifically the monitor's physical size (in mm).
        m_EDIDWidthMM = ((m_EDID(68) And &HF0) * 16) + m_EDID(66)
        m_EDIDHeightMM = ((m_EDID(68) And &HF) * 256) + m_EDID(67)
        
        'For convenience, store a matching size in inches
        m_EDIDWidthInches = (m_EDIDWidthMM / 25.4)
        m_EDIDHeightInches = (m_EDIDHeightMM / 25.4)
        m_EDIDDiagonalInches = Sqr(m_EDIDWidthInches * m_EDIDWidthInches + m_EDIDHeightInches * m_EDIDHeightInches)
        
        'Next, retrieve the monitor's native resolution.  Note that this is only ACTIVE pixels.  CRT monitors may also
        ' include blanking and/or border pixels; these are not retrieved.
        m_EDIDNativeResH = ((m_EDID(58) And &HF0) * 16) + m_EDID(56)
        m_EDIDNativeResV = ((m_EDID(61) And &HF0) * 16) + m_EDID(59)
        
    End If
    
    'Next, we will try to retrieve a monitor name.  Monitor name is an optional descriptor (but provided 99.9% of the time)
    ' found in one of the four VESA Descriptor blocks of the EDID.  Those descriptor blocks are found in set locations:
    ' 54–71, 72–89, 90–107, 108–125
    '
    'Location of the monitor name is not guaranteed in advance, so each block must be parsed individually until found.
    If EDIDLength >= 128 Then
    
        Dim descriptorBlock As Long, startOffset As Long, i As Long
        
        Dim monName As String, monSerial As String, monChar As String
        
        'Search each descriptor block for the "monitor name" identifier
        For descriptorBlock = 0 To 3
        
            'Determine a proper offset based on which descriptor block we're searching
            Select Case descriptorBlock
            
                Case 0
                    startOffset = 54
                
                Case 1
                    startOffset = 72
                
                Case 2
                    startOffset = 90
                
                Case 3
                    startOffset = 108
            
            End Select
            
            'Check the first three bytes of the descriptor.  These must always be 0 (per the EDID spec); if they aren't,
            ' this is the primary timing descriptor.
            If (m_EDID(startOffset) = 0) And (m_EDID(startOffset + 1) = 0) And (m_EDID(startOffset + 2) = 0) Then
            
                'This is not the primary timing block.  Look for the #FC identifier in byte 3, which indicates
                ' a monitor name block.
                If (m_EDID(startOffset + 3) = &HFC) Then
                
                    'Retrieve the monitor name and trim any null bytes
                    monName = ""
                    For i = (startOffset + 4) To (startOffset + 17)
                        If (m_EDID(i) <> 0) Then
                            monChar = Chr(m_EDID(i))
                            If (monChar = vbLf) Or (monChar = vbCr) Then monChar = " "
                            monName = monName & monChar
                        End If
                    Next i
                    
                    'Save the retrieved monitor name now
                    If Len(monName) > 0 Then m_EDIDModelName = Trim$(monName)
                    
                'While we're parsing descriptor blocks, we may as well parse for serial numbers, too.  This could be very helpful
                ' in the future as a unique ID for storing extra color management data on a per-monitor basis, with a better
                ' guarantee of uniqueness vs something like an HMONITOR (which may change as monitors are added/removed).
                ElseIf (m_EDID(startOffset + 3) = &HFF) Then
                
                    'Retrieve the monitor name and trim any null bytes
                    monSerial = ""
                    For i = (startOffset + 4) To (startOffset + 17)
                        If (m_EDID(i) <> 0) Then
                            monChar = Chr(m_EDID(i))
                            If (monChar = vbLf) Or (monChar = vbCr) Then monChar = " "
                            monSerial = monSerial & monChar
                        End If
                    Next i
                    
                    'Save the retrieved monitor name in its parent object
                    If Len(monSerial) > 0 Then m_EDIDSerialNumber = Trim$(monSerial)
                
                End If
            
            End If
        
        Next descriptorBlock
        
    End If
    
End Sub
