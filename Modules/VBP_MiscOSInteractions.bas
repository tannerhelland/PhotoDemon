Attribute VB_Name = "OS_Interactions"
'***************************************************************************
'Miscellaneous OS Interaction Handler
'Copyright ©2011-2014 by Tanner Helland
'Created: 27/November/12
'Last updated: 02/April/14
'Last update: added getArbitraryGUID function, for generating unique strings at run-time.  (PD will use this to
'              create unique session IDs for each running instance.)
'
'Sometimes, PhotoDemon needs to query Windows for OS-specific data - such as the current version of Windows, or the
' available RAM on the system.  This module handles such calls.
'
'Special thanks goes to Mike Raynder, who wrote the original version of the process-specific memory function.  You can
' download a copy of Mike's original code from this link (good as of 27 Nov '12): http://www.xtremevbtalk.com/showthread.php?t=229758
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Type and call necessary for determining the current version of Windows
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    wServicePackMajor  As Integer
    wServicePackMinor  As Integer
    wSuiteMask         As Integer
    wProductType       As Byte
    wReserved          As Byte
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFOEX) As Long

'Type and call for receiving additional OS data (32/64 bit for PD's purposes)
Private Type SYSTEM_INFO
    wProcessorArchitecture        As Integer
    wReserved                     As Integer
    dwPageSize                    As Long
    lpMinimumApplicationAddress   As Long
    lpMaximumApplicationAddress   As Long
    dwActiveProcessorMask         As Long
    dwNumberOfProcessors          As Long
    dwProcessorType               As Long
    dwAllocationGranularity       As Long
    wProcessorLevel               As Integer
    wProcessorRevision            As Integer
End Type

Private Const VER_NT_WORKSTATION As Long = &H1&

Private Declare Sub GetNativeSystemInfo Lib "kernel32" (ByRef lpSystemInfo As SYSTEM_INFO)

'Constants for GetSystemInfo and GetNativeSystemInfo API functions (SYSTEM_INFO structure)
Private Const PROCESSOR_ARCHITECTURE_AMD64      As Long = 9         'x64 (AMD or Intel)
Private Const PROCESSOR_ARCHITECTURE_IA64       As Long = 6         'Intel Itanium Processor Family (IPF)
Private Const PROCESSOR_ARCHITECTURE_INTEL      As Long = 0
Private Const PROCESSOR_ARCHITECTURE_UNKNOWN    As Long = &HFFFF&

'Query for specific processor features
Private Declare Function IsProcessorFeaturePresent Lib "kernel32" (ByVal ProcessorFeature As Long) As Boolean

Private Const PF_3DNOW_INSTRUCTIONS_AVAILABLE As Long = 7
Private Const PF_MMX_INSTRUCTIONS_AVAILABLE As Long = 3
Private Const PF_NX_ENABLED As Long = 12
Private Const PF_SSE3_INSTRUCTIONS_AVAILABLE As Long = 13
Private Const PF_VIRT_FIRMWARE_ENABLED As Long = 21
Private Const PF_XMMI_INSTRUCTIONS_AVAILABLE As Long = 6
Private Const PF_XMMI64_INSTRUCTIONS_AVAILABLE As Long = 10

'Query system memory counts and availability
Private Type MemoryStatusEx
    dwLength As Long
    dwMemoryLoad As Long
    ullTotalPhys As Currency
    ullAvailPhys As Currency
    ullTotalPageFile As Currency
    ullAvailPageFile As Currency
    ullTotalVirtual As Currency
    ullAvailVirtual As Currency
    ullAvailExtendedVirtual As Currency
End Type

Private Declare Function GlobalMemoryStatusEx Lib "kernel32" (ByRef lpBuffer As MemoryStatusEx) As Long

'Types and calls necessary for calculating PhotoDemon's current memory usage
Private Type PROCESS_MEMORY_COUNTERS
   cb As Long
   PageFaultCount As Long
   PeakWorkingSetSize As Long
   WorkingSetSize As Long
   QuotaPeakPagedPoolUsage As Long
   QuotaPagedPoolUsage As Long
   QuotaPeakNonPagedPoolUsage As Long
   QuotaNonPagedPoolUsage As Long
   PagefileUsage As Long
   PeakPagefileUsage As Long
End Type

Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16
Private Const MAX_PATH = 260

Private Declare Function EnumProcesses Lib "psapi" (lpidProcess As Long, ByVal cb As Long, cbNeeded As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi" (ByVal hProcess As Long, lphModule As Long, ByVal cb As Long, lpcbNeeded As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function GetProcessMemoryInfo Lib "psapi" (ByVal hProcess As Long, ppsmemCounters As PROCESS_MEMORY_COUNTERS, ByVal cb As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal Handle As Long) As Long

'Device caps, or "device capabilities", which can be probed using the constants below
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As DeviceChecks) As Long

Public Enum DeviceChecks
    CURVECAPS = 28
    LINECAPS = 30
    POLYGONALCAPS = 32
    TEXTCAPS = 34
    RASTERCAPS = 38
    SHADEBLENDCAPS = 45
    COLORMGMTCAPS = 121
End Enum

#If False Then
    Private Const CURVECAPS = 28, LINECAPS = 30, POLYGONALCAPS = 32, TEXTCAPS = 34, RASTERCAPS = 38, SHADEBLENDCAPS = 45, COLORMGMTCAPS = 121
#End If

'Alpha blend capabilites
Private Const SB_CONST_ALPHA As Long = 1
Private Const SB_PIXEL_ALPHA As Long = 2

'Blt hardware capabilities
Private Const RC_BITBLT As Long = 1
Private Const RC_BANDING As Long = 2
Private Const RC_SCALING As Long = 4
Private Const RC_BITMAP64 As Long = 8
Private Const RC_GDI20_OUTPUT As Long = &H10
Private Const RC_DI_BITMAP As Long = &H80
Private Const RC_PALETTE As Long = &H100
Private Const RC_DIBTODEV As Long = &H200
Private Const RC_STRETCHBLT As Long = &H800
Private Const RC_FLOODFILL As Long = &H1000
Private Const RC_STRETCHDIB As Long = &H2000

'Color management capabilities
Private Const CM_NONE As Long = 0
Private Const CM_DEVICE_ICM As Long = 1
Private Const CM_GAMMA_RAMP As Long = 2
Private Const CM_CMYK_COLOR As Long = 4

'Line drawing capabilities
Private Const LC_NONE As Long = 0
Private Const LC_POLYLINE As Long = 2
Private Const LC_MARKER As Long = 4
Private Const LC_POLYMARKER As Long = 8
Private Const LC_WIDE As Long = 16
Private Const LC_STYLED As Long = 32
Private Const LC_INTERIORS As Long = 128
Private Const LC_WIDESTYLED As Long = 64

'Curve drawing capabilities
Private Const CC_NONE As Long = 0
Private Const CC_CIRCLES As Long = 1
Private Const CC_PIE As Long = 2
Private Const CC_CHORD As Long = 4
Private Const CC_ELLIPSES As Long = 8
Private Const CC_WIDE As Long = 16
Private Const CC_STYLED As Long = 32
Private Const CC_WIDESTYLED As Long = 64
Private Const CC_INTERIORS As Long = 128
Private Const CC_ROUNDRECT As Long = 256

'Polygon drawing capabilities
Private Const PC_NONE As Long = 0
Private Const PC_POLYGON As Long = 1
Private Const PC_RECTANGLE As Long = 2
Private Const PC_WINDPOLYGON As Long = 4
Private Const PC_SCANLINE As Long = 8
Private Const PC_WIDE As Long = 16
Private Const PC_STYLED As Long = 32
Private Const PC_WIDESTYLED As Long = 64
Private Const PC_INTERIORS As Long = 128

'Text drawing capabilities
Private Const TC_OP_CHARACTER As Long = 1
Private Const TC_OP_STROKE As Long = 2
Private Const TC_CP_STROKE As Long = 4
Private Const TC_CR_90 As Long = 8
Private Const TC_CR_ANY As Long = 10
Private Const TC_SF_X_YINDEP As Long = 20
Private Const TC_SA_DOUBLE As Long = 40
Private Const TC_SA_INTEGER As Long = 80
Private Const TC_SA_CONTIN As Long = 100
Private Const TC_EA_DOUBLE As Long = 200
Private Const TC_IA_ABLE As Long = 400
Private Const TC_UA_ABLE As Long = 800
Private Const TC_SO_ABLE As Long = 1000
Private Const TC_RA_ABLE As Long = 2000
Private Const TC_VA_ABLE As Long = 4000
Private Const TC_SCROLLBLT As Long = 10000

'GUID creation
Public Type Guid
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(0 To 7) As Byte
End Type

Private Declare Function CoCreateGuid Lib "ole32" (ByRef pGuid As Guid) As Long
Private Declare Function StringFromGUID2 Lib "ole32" (ByRef rGuid As Any, ByVal lpstrClsId As Long, ByVal cbMax As Long) As Long

'Windows constants for retrieving a unique temporary filename
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

'Return a unique session ID for this PhotoDemon instance.  A session ID is generated by retrieving a random GUID,
' hashing it, then returning the first 16 characters from the hash.  So many random steps are not necessary, but
' they help ensure that the IDs are actually unique.
Public Function getUniqueSessionID() As String

    'Start by retrieving a random GUID
    Dim randomGUID As String
    randomGUID = getArbitraryGUID(True)
    
    'Hash the returned GUID
    Dim cSHA2 As CSHA256
    Set cSHA2 = New CSHA256
        
    Dim hString As String
    hString = cSHA2.SHA256(randomGUID)
            
    'The SHA-256 function returns a 64 character string (256 / 8 = 32 bytes, but 64 characters due to hex representation).
    ' This is too long for a filename, so take only the first sixteen characters of the hash.
    getUniqueSessionID = Left$(hString, 16)

End Function

'Sometimes, a unique string is needed.  Use this function to retrieve an arbitrary GUID from WAPI.
Private Function getArbitraryGUID(Optional ByVal stripNonHexCharacters As Boolean = False) As String

    'Fill a GUID struct with data via WAPI
    Dim tmpGuid As Guid
    CoCreateGuid tmpGuid
    
    'We can convert it into a string manually, but it's much easier to let Windows do it for us
    
    'Prepare an empty byte array
    Dim tmpBytes() As Byte
    Dim lenGuid As Long
    lenGuid = 40
    ReDim tmpBytes(0 To (lenGuid * 2) - 1) As Byte

    'Use the API to fill to the byte array with a string version of the GUID we created.  This function will return
    ' the length of the created string - *including the null terminator*; use that to trim the string.
    Dim guidString As String
    Dim lenGuidString As Long
    lenGuidString = StringFromGUID2(tmpGuid, VarPtr(tmpBytes(0)), lenGuid)
    guidString = Left$(tmpBytes, lenGuidString - 1)
    
    'If the caller wants non-hex characters removed from the String, do so now
    If stripNonHexCharacters Then
        
        'Trim brackets
        guidString = Mid$(guidString, 2, Len(guidString) - 2)
        
        'Trim dividers
        guidString = Replace$(guidString, "-", "")
        
    End If
    
    getArbitraryGUID = guidString

End Function

'Return a unique temporary filename, via the API.  Thank you to this MSDN support doc for the implementation:
' http://support.microsoft.com/kb/195763
Public Function getUniqueTempFilename(Optional ByRef customPrefix As String = "PD_") As String
         
    Dim sTmpPath As String * 512
    Dim sTmpName As String * 576
    Dim nRet As Long

    nRet = GetTempPath(512, sTmpPath)
    If (nRet > 0 And nRet < 512) Then
    
        nRet = GetTempFileName(sTmpPath, customPrefix, 0, sTmpName)
        
        If nRet <> 0 Then
            getUniqueTempFilename = Left$(sTmpName, InStr(sTmpName, vbNullChar) - 1)
        Else
            getUniqueTempFilename = ""
        End If
    
    Else
        getUniqueTempFilename = ""
    End If

End Function

'Given a type of device capability check, return a string that describes the reported capabilities
Public Function getDeviceCapsString() As String

    Dim fullString As String
    fullString = ""
    
    Dim hwYes As String, hwNo As String
    hwYes = ""
    hwNo = ""
    
    Dim supportedCount As Long, totalCount As Long
    supportedCount = 0
    totalCount = 0
    
    Dim gdcReturn As Long
    
    'Start with blitting actions
    startDevCapsSection fullString, gdcReturn, RASTERCAPS, g_Language.TranslateMessage("General image actions")
    
    addToDeviceCapsString gdcReturn, RC_BITBLT, hwYes, hwNo, "BitBlt", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, RC_STRETCHBLT, hwYes, hwNo, "StretchBlt", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, RC_DI_BITMAP, hwYes, hwNo, "DIBs", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, RC_STRETCHDIB, hwYes, hwNo, "StretchDIB", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, RC_DIBTODEV, hwYes, hwNo, "SetDIBitsToDevice", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, RC_BITMAP64, hwYes, hwNo, "64kb+ chunks", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, RC_SCALING, hwYes, hwNo, "general scaling", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, RC_FLOODFILL, hwYes, hwNo, "flood fill", supportedCount, totalCount
    
    endDevCapsSection fullString, hwYes, hwNo
    
    'Alpha blending
    startDevCapsSection fullString, gdcReturn, SHADEBLENDCAPS, g_Language.TranslateMessage("Alpha-blending")
    
    addToDeviceCapsString gdcReturn, SB_CONST_ALPHA, hwYes, hwNo, "simple alpha", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, SB_PIXEL_ALPHA, hwYes, hwNo, "per-pixel alpha", supportedCount, totalCount
    
    endDevCapsSection fullString, hwYes, hwNo
    
    'Color management
    startDevCapsSection fullString, gdcReturn, COLORMGMTCAPS, g_Language.TranslateMessage("Color management")
    
    addToDeviceCapsString gdcReturn, CM_DEVICE_ICM, hwYes, hwNo, "color transformation", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, CM_GAMMA_RAMP, hwYes, hwNo, "gamma ramping", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, CM_CMYK_COLOR, hwYes, hwNo, "CMYK", supportedCount, totalCount
    
    endDevCapsSection fullString, hwYes, hwNo
    
    'Lines
    startDevCapsSection fullString, gdcReturn, LINECAPS, g_Language.TranslateMessage("Lines")
    
    addToDeviceCapsString gdcReturn, LC_POLYLINE, hwYes, hwNo, "polylines", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, LC_MARKER, hwYes, hwNo, "markers", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, LC_POLYMARKER, hwYes, hwNo, "polymarkers", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, LC_INTERIORS, hwYes, hwNo, "interiors", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, LC_WIDE, hwYes, hwNo, "wide", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, LC_STYLED, hwYes, hwNo, "styled", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, LC_WIDESTYLED, hwYes, hwNo, "wide+styled", supportedCount, totalCount
    
    endDevCapsSection fullString, hwYes, hwNo
    
    'Polygons
    startDevCapsSection fullString, gdcReturn, POLYGONALCAPS, g_Language.TranslateMessage("Polygons")
    
    addToDeviceCapsString gdcReturn, PC_RECTANGLE, hwYes, hwNo, "rectangles", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, PC_POLYGON, hwYes, hwNo, "alternate-fill", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, PC_WINDPOLYGON, hwYes, hwNo, "winding-fill", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, PC_INTERIORS, hwYes, hwNo, "interiors", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, PC_WIDE, hwYes, hwNo, "wide", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, PC_STYLED, hwYes, hwNo, "styled", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, PC_WIDESTYLED, hwYes, hwNo, "wide+styled", supportedCount, totalCount
    
    endDevCapsSection fullString, hwYes, hwNo
    
    'Curves
    startDevCapsSection fullString, gdcReturn, CURVECAPS, g_Language.TranslateMessage("Curves")
    
    addToDeviceCapsString gdcReturn, CC_CIRCLES, hwYes, hwNo, "circles", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, CC_ELLIPSES, hwYes, hwNo, "ellipses", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, CC_ROUNDRECT, hwYes, hwNo, "rounded rectangles", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, CC_PIE, hwYes, hwNo, "pie wedges", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, CC_INTERIORS, hwYes, hwNo, "interiors", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, CC_CHORD, hwYes, hwNo, "chords", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, CC_WIDE, hwYes, hwNo, "wide", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, CC_STYLED, hwYes, hwNo, "styled", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, CC_WIDESTYLED, hwYes, hwNo, "wide+styled", supportedCount, totalCount
    
    endDevCapsSection fullString, hwYes, hwNo
    
    'Text
    startDevCapsSection fullString, gdcReturn, TEXTCAPS, g_Language.TranslateMessage("Text")
    
    addToDeviceCapsString gdcReturn, TC_RA_ABLE, hwYes, hwNo, "raster fonts", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, TC_VA_ABLE, hwYes, hwNo, "vector fonts", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, TC_OP_CHARACTER, hwYes, hwNo, "high-precision characters", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, TC_OP_STROKE, hwYes, hwNo, "high-precision strokes", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, TC_CP_STROKE, hwYes, hwNo, "high-precision clipping", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, TC_SA_CONTIN, hwYes, hwNo, "high-precision scaling", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, TC_SF_X_YINDEP, hwYes, hwNo, "independent x/y scaling", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, TC_CR_90, hwYes, hwNo, "90-degree rotation", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, TC_CR_ANY, hwYes, hwNo, "free rotation", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, TC_EA_DOUBLE, hwYes, hwNo, "bold", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, TC_IA_ABLE, hwYes, hwNo, "italics", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, TC_UA_ABLE, hwYes, hwNo, "underline", supportedCount, totalCount
    addToDeviceCapsString gdcReturn, TC_SO_ABLE, hwYes, hwNo, "strikeouts", supportedCount, totalCount
        
    endDevCapsSection fullString, hwYes, hwNo
    
    'Add some summary statistics at the end
    fullString = fullString & g_Language.TranslateMessage("Final results") & vbCrLf
    fullString = fullString & "    " & "Accelerated actions: " & supportedCount & " (" & Format((CDbl(supportedCount) / CDbl(totalCount)), "00.0%") & ")" & vbCrLf
    fullString = fullString & "    " & "Not accelerated actions: " & (totalCount - supportedCount) & " (" & Format((CDbl(totalCount - supportedCount) / CDbl(totalCount)), "00.0%") & ")"
    fullString = fullString & vbCrLf & vbCrLf & g_Language.TranslateMessage("Disclaimer: all hardware acceleration data is provided by the operating system.  It specifically represents GDI acceleration, which is independent from DirectX and OpenGL.  OS version and desktop mode also affect support capabilities.  For best results, please run PhotoDemon on Windows 7 or 8, on an Aero-enabled desktop.")
    fullString = fullString & vbCrLf & vbCrLf & g_Language.TranslateMessage("For more information on GDI hardware acceleration, visit http://msdn.microsoft.com/en-us/library/windows/desktop/ff729480")
    
    getDeviceCapsString = fullString

End Function

'Helper function for getDeviceCapsString, above; used to append text to the start of a new device caps section
Private Sub startDevCapsSection(ByRef srcString As String, ByRef getDevCapsReturn As Long, ByVal gdcSection As DeviceChecks, ByRef sectionTitle As String)
    
    getDevCapsReturn = GetDeviceCaps(GetDC(GetDesktopWindow), gdcSection)
    
    srcString = srcString & sectionTitle & vbCrLf
    
End Sub

'Helper function for getDeviceCapsString, above; used to append text to the end of a device caps section
Private Sub endDevCapsSection(ByRef srcString As String, ByRef supportedCaps As String, ByRef unsupportedCaps As String)
    
    If Len(supportedCaps) = 0 Then supportedCaps = g_Language.TranslateMessage("none")
    If Len(unsupportedCaps) = 0 Then unsupportedCaps = g_Language.TranslateMessage("none")
    
    srcString = srcString & "    " & g_Language.TranslateMessage("accelerated: ") & supportedCaps & vbCrLf
    srcString = srcString & "    " & g_Language.TranslateMessage("not accelerated: ") & unsupportedCaps & vbCrLf
    
    Dim headerLine As String
    headerLine = "---------------------------------------"
    
    srcString = srcString & headerLine & vbCrLf
    
    supportedCaps = ""
    unsupportedCaps = ""
    
End Sub

'Helper function for getDeviceCapsString, above; used to automatically check a given GetDeviceCaps return value, and append the
' results to a user-friendly string
Private Sub addToDeviceCapsString(ByVal devCapsReturn As Long, ByVal paramToCheck As Long, ByRef stringIfSupported As String, ByRef stringIfNotSupported As String, ByRef capName As String, ByRef supportedCount As Long, ByRef totalCount As Long)
    
    totalCount = totalCount + 1
    
    If ((devCapsReturn And paramToCheck) <> 0) Then
        appendCapToString stringIfSupported, capName
        supportedCount = supportedCount + 1
    Else
        appendCapToString stringIfNotSupported, capName
    End If

End Sub

'Helper function for addToDeviceCapsString, above; simply appends text to a list with a comma, as necessary
Private Sub appendCapToString(ByRef oldPart As String, ByRef newPart As String)

    If Len(oldPart) = 0 Then
        oldPart = newPart
    Else
        oldPart = oldPart & ", " & newPart
    End If

End Sub

'Check for a version >= Vista.
Public Function getVistaOrLaterStatus() As Boolean

    Dim tOSVI As OSVERSIONINFOEX
    tOSVI.dwOSVersionInfoSize = Len(tOSVI)
    GetVersionEx tOSVI
    
    getVistaOrLaterStatus = (tOSVI.dwMajorVersion >= 6)

End Function

'Check for a version >= Win 7
Public Function getWin7OrLaterStatus() As Boolean

    Dim tOSVI As OSVERSIONINFOEX
    tOSVI.dwOSVersionInfoSize = Len(tOSVI)
    GetVersionEx tOSVI
    
    getWin7OrLaterStatus = ((tOSVI.dwMajorVersion >= 6) And (tOSVI.dwMinorVersion >= 1))

End Function

'Return the current OS version as a string.  (At present, this data is added to debug logs.)
Public Function getOSVersionAsString() As String
    
    'Retrieve OS version data
    Dim tOSVI As OSVERSIONINFOEX
    tOSVI.dwOSVersionInfoSize = Len(tOSVI)
    GetVersionEx tOSVI
    
    Dim osName As String
    
    Select Case tOSVI.dwMajorVersion
    
        Case 6
            
            Select Case tOSVI.dwMinorVersion
            
                Case 3
                    If (tOSVI.wProductType And VER_NT_WORKSTATION) <> 0 Then
                        osName = "Windows 8.1"
                    Else
                        osName = "Windows Server 2012 R2"
                    End If
                    
                Case 2
                    If (tOSVI.wProductType And VER_NT_WORKSTATION) <> 0 Then
                        osName = "Windows 8"
                    Else
                        osName = "Windows Server 2012"
                    End If
                    
                Case 1
                    If (tOSVI.wProductType And VER_NT_WORKSTATION) <> 0 Then
                        osName = "Windows 7"
                    Else
                        osName = "Windows Server 2008 R2"
                    End If
                
                Case 0
                    If (tOSVI.wProductType And VER_NT_WORKSTATION) <> 0 Then
                        osName = "Windows Vista"
                    Else
                        osName = "Windows Server 2008"
                    End If
            
            End Select
        
        Case 5
            osName = "Windows XP"
    
    End Select
    
    'Retrieve 32/64 bit OS version
    Dim osBitness As String
    
    Dim tSYSINFO As SYSTEM_INFO
    Call GetNativeSystemInfo(tSYSINFO)
    
    Select Case tSYSINFO.wProcessorArchitecture
    
        Case PROCESSOR_ARCHITECTURE_AMD64
            osBitness = " 64-bit "
            
        Case PROCESSOR_ARCHITECTURE_IA64
            osBitness = " Itanium "
            
        Case Else
            osBitness = " 32-bit "
    
    End Select
    
    Dim buildString As String
    buildString = TrimNull(tOSVI.szCSDVersion)
    
    With tOSVI
        getOSVersionAsString = osName & " " & IIf(Len(buildString) > 0, buildString, "") & osBitness & "(" & .dwMajorVersion & "." & .dwMinorVersion & "." & .dwBuildNumber & ")"
    End With

End Function

'Return the number of logical cores on this system
Public Function getNumLogicalCores() As Long
    
    Dim tSYSINFO As SYSTEM_INFO
    Call GetNativeSystemInfo(tSYSINFO)
    
    getNumLogicalCores = tSYSINFO.dwNumberOfProcessors

End Function

'Return a list of PD-relevant processor features, in string format
Public Function getProcessorFeatures() As String

    Dim listFeatures As String
    listFeatures = ""
    
    If IsProcessorFeaturePresent(PF_3DNOW_INSTRUCTIONS_AVAILABLE) Then listFeatures = listFeatures & "3DNow!" & ", "
    If IsProcessorFeaturePresent(PF_MMX_INSTRUCTIONS_AVAILABLE) Then listFeatures = listFeatures & "MMX" & ", "
    If IsProcessorFeaturePresent(PF_XMMI_INSTRUCTIONS_AVAILABLE) Then listFeatures = listFeatures & "SSE" & ", "
    If IsProcessorFeaturePresent(PF_XMMI64_INSTRUCTIONS_AVAILABLE) Then listFeatures = listFeatures & "SSE2" & ", "
    If IsProcessorFeaturePresent(PF_SSE3_INSTRUCTIONS_AVAILABLE) Then listFeatures = listFeatures & "SSE3" & ", "
    If IsProcessorFeaturePresent(PF_NX_ENABLED) Then listFeatures = listFeatures & "DEP" & ", "
    If IsProcessorFeaturePresent(PF_VIRT_FIRMWARE_ENABLED) Then listFeatures = listFeatures & "Virtualization" & ", "
    
    'Trim the trailing comma and blank space
    If Len(listFeatures) > 0 Then
        getProcessorFeatures = Left$(listFeatures, Len(listFeatures) - 2)
    Else
        getProcessorFeatures = "(none)"
    End If
    
End Function

'Query total system RAM
Public Function getTotalSystemRAM() As String

    Dim memStatus As MemoryStatusEx
    memStatus.dwLength = Len(memStatus)
    Call GlobalMemoryStatusEx(memStatus)
    
    getTotalSystemRAM = CStr(Int(CDbl(memStatus.ullTotalPhys / 1024) * 10)) & " MB"
    
End Function

'Query RAM available to PD
Public Function getRAMAvailableToPD() As String

    Dim memStatus As MemoryStatusEx
    memStatus.dwLength = Len(memStatus)
    Call GlobalMemoryStatusEx(memStatus)
    
    Dim tmpString As String
    
    tmpString = CStr(Int(CDbl(memStatus.ullTotalVirtual / 1024) * 10)) & " MB"
    tmpString = tmpString & " (real), "
    tmpString = tmpString & CStr(Int(CDbl(memStatus.ullAvailPageFile / 1024) * 10)) & " MB"
    tmpString = tmpString & " (hypothetical)"
    
    getRAMAvailableToPD = tmpString
    
End Function

'Function for returning PhotoDemon's current memory usage.  This is a modified version of code first published by Mike Raynder.
' You can download a copy of Mike's original code from this link (good as of 27 Nov '12): http://www.xtremevbtalk.com/showthread.php?t=229758

'---------------------------------------------------------------------------------------
' Procedure : GetProcessMemory
' DateTime  : 2 sep 2004
' Author    : Mike Raynder
' Purpose   : Will only work for NT
'
' Tanner's note: if False is passed, returns a process's working memory set, in kilobytes
'                if True is passed, returns the highest memory count the program has hit
'---------------------------------------------------------------------------------------

Public Function GetPhotoDemonMemoryUsage(Optional returnPeakValue As Boolean = False) As Long
  
    Dim lngLength As Long
    Dim strProcessName As String
    
    'Specifies the size, In bytes, of the lpidProcess array
    Dim lngCBSize As Long
    
    'Receives the number of bytes returned
    Dim lngCBSizeReturned As Long
    
    Dim lngNumElements As Long
    Dim lngProcessIDs() As Long
    Dim lngCBSize2 As Long
    Dim lngModules(1 To 200) As Long
    Dim lngReturn As Long
    Dim strModuleName As String
    Dim lngSize As Long
    Dim lnghWndProcess As Long
    Dim lngLoop As Long
    Dim pmc As PROCESS_MEMORY_COUNTERS
    Dim lRet As Long
    Dim strProcName2 As String

    On Error GoTo Memory_Measure_Error:

    'EXEName was original passed into this function as a parameter.  There is no need to require that for use in PhotoDemon only.
    Dim EXEName As String
    If g_IsProgramCompiled Then EXEName = "photodemon.exe" Else EXEName = "vb6.exe"

    EXEName = UCase$(Trim$(EXEName))
    lngLength = Len(EXEName)
    
    lngCBSize = 8 ' Really needs To be 16, but Loop will increment prior to calling API
    lngCBSizeReturned = 96

    Do While lngCBSize <= lngCBSizeReturned
    
        'Increment Size
        lngCBSize = lngCBSize * 2
        
        'Allocate Memory for Array
        ReDim lngProcessIDs(lngCBSize / 4) As Long
        
        'Get Process ID's
        lngReturn = EnumProcesses(lngProcessIDs(1), lngCBSize, lngCBSizeReturned)
  
    Loop

    'Count number of processes returned
    lngNumElements = lngCBSizeReturned / 4
  
    'Loop thru each process
    For lngLoop = 1 To lngNumElements

        'Get a handle to the Process and Open
        lnghWndProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lngProcessIDs(lngLoop))
    
        If lnghWndProcess <> 0 Then
            'Get an array of the module handles for the specified process
            lngReturn = EnumProcessModules(lnghWndProcess, lngModules(1), 200, lngCBSize2)

            'If the Module Array is retrieved, Get the ModuleFileName
            If lngReturn <> 0 Then

                'Buffer with spaces first to allocate memory for byte array
                strModuleName = Space(MAX_PATH)

                'Must be set prior to calling API
                lngSize = 500

                'Get Process Name
                lngReturn = GetModuleFileNameExA(lnghWndProcess, lngModules(1), strModuleName, lngSize)

                'Remove trailing spaces
                strProcessName = Left$(strModuleName, lngReturn)
        
                strProcName2 = GetExeName(strProcessName)

                If strProcName2 = EXEName Then
          
                    'Get the Site of the Memory Structure
                    pmc.cb = LenB(pmc)
          
                    lRet = GetProcessMemoryInfo(lnghWndProcess, pmc, pmc.cb)
              
                    If returnPeakValue Then
                        GetPhotoDemonMemoryUsage = pmc.PeakWorkingSetSize / 1024
                    Else
                        GetPhotoDemonMemoryUsage = pmc.WorkingSetSize / 1024
                    End If

                End If
                
            End If
            
        End If
    
        'Close the handle to this process
        lngReturn = CloseHandle(lnghWndProcess)
        
    Next lngLoop

IsProcessRunning_Exit:

    'Exit early to avoid error handler
    Exit Function
    
Memory_Measure_Error:
  
    'I am not current interested in raising errors for this function, so simply resume if something goes wrong.
    Resume Next
  
End Function

'Used to extract the EXE name from a running process
Private Function GetExeName(ByVal sPath As String) As String
  
    Dim lPos1 As Long
    Dim lPos2 As Long
      
    On Error Resume Next
      
    lPos1 = InStr(1, sPath, Chr$(0))
    lPos2 = InStrRev(sPath, "\")
    
    If lPos1 > 0 Then
        GetExeName = UCase$(Mid$(sPath, lPos2 + 1, lPos1 - lPos2))
    Else
        GetExeName = UCase$(Mid$(sPath, lPos2 + 1))
    End If
  
End Function
