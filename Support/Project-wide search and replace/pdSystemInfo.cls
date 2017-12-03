VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdSystemInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'Generic "System Information" interface
'Copyright 2012-2016 by Tanner Helland
'Created: 27/November/12
'Last updated: 14/September/15
'Last update: greatly simplify the way PD measures its own memory usage
'
'Sometimes, PhotoDemon needs to grab OS-specific data (e.g. the current OS version of Windows, available RAM , etc.)
' This module is designed to make such requests easier.
'
'I am currently working on migrating random bits of code from other places in PD to this class, so please be patient and ignore
' duplicate code entries in the meantime.  :)
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

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function GetProcessMemoryInfo Lib "psapi" (ByVal hProcess As Long, ppsmemCounters As PROCESS_MEMORY_COUNTERS, ByVal cb As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal Handle As Long) As Long

Private Const TH32CS_SNAPPROCESS As Long = 2&
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long

'Similar APIs for retrieving GDI and user objects
Public Enum GUI_RESOURCE_FLAGS
    GdiObjects = 0
    UserObjects = 1
    GdiObjectsPeak = 2
    UserObjectsPeak = 4
End Enum
#If False Then
    Private Const GdiObjects = 0, GdiObjectsPeak = 2, UserObjects = 1, UserObjectsPeak = 4
#End If
Private Declare Function GetGuiResources Lib "user32" (ByVal hProcess As Long, ByVal resourceToCheck As GUI_RESOURCE_FLAGS) As Long

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
Private Type Guid
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(0 To 7) As Byte
End Type

Private Declare Function CoCreateGuid Lib "ole32" (ByRef pGuid As Guid) As Long
Private Declare Function StringFromGUID2 Lib "ole32" (ByRef rguid As Any, ByVal lpstrClsId As Long, ByVal cbMax As Long) As Long

'Windows constants for retrieving a unique temporary filename
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

'Sometimes, a unique string is needed.  Use this function to retrieve an arbitrary GUID from WAPI.
Private Function GetArbitraryGUID(Optional ByVal stripNonHexCharacters As Boolean = False) As String

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
    
    GetArbitraryGUID = guidString

End Function

'Return a unique temporary filename, via the API.  Thank you to this MSDN support doc for the implementation:
' http://support.microsoft.com/kb/195763
Public Function GetUniqueTempFilename(Optional ByRef customPrefix As String = "PD_") As String
         
    Dim sTmpPath As String * 512
    Dim sTmpName As String * 576
    Dim nRet As Long

    nRet = GetTempPath(512, sTmpPath)
    If (nRet > 0 And nRet < 512) Then
    
        nRet = GetTempFileName(sTmpPath, customPrefix, 0, sTmpName)
        
        If nRet <> 0 Then
            GetUniqueTempFilename = Left$(sTmpName, InStr(sTmpName, vbNullChar) - 1)
        Else
            GetUniqueTempFilename = ""
        End If
    
    Else
        GetUniqueTempFilename = ""
    End If

End Function

'Check for a version >= Vista.
Public Function IsOSVistaOrLater() As Boolean

    Dim tOSVI As OSVERSIONINFOEX
    tOSVI.dwOSVersionInfoSize = Len(tOSVI)
    GetVersionEx tOSVI
    
    IsOSVistaOrLater = (tOSVI.dwMajorVersion >= 6)

End Function

'Check for a version >= Win 7
Public Function IsOSWin7OrLater() As Boolean

    Dim tOSVI As OSVERSIONINFOEX
    tOSVI.dwOSVersionInfoSize = Len(tOSVI)
    GetVersionEx tOSVI
    
    IsOSWin7OrLater = (tOSVI.dwMajorVersion > 6) Or ((tOSVI.dwMajorVersion = 6) And (tOSVI.dwMinorVersion >= 1))

End Function

'Check for a version >= Win 8.0
Public Function IsOSWin8OrLater() As Boolean

    Dim tOSVI As OSVERSIONINFOEX
    tOSVI.dwOSVersionInfoSize = Len(tOSVI)
    GetVersionEx tOSVI
    
    IsOSWin8OrLater = (tOSVI.dwMajorVersion > 6) Or ((tOSVI.dwMajorVersion = 6) And (tOSVI.dwMinorVersion >= 2))

End Function

'Check for a version >= Win 8.1
Public Function IsOSWin81OrLater() As Boolean

    Dim tOSVI As OSVERSIONINFOEX
    tOSVI.dwOSVersionInfoSize = Len(tOSVI)
    GetVersionEx tOSVI
    
    IsOSWin81OrLater = (tOSVI.dwMajorVersion > 6) Or ((tOSVI.dwMajorVersion = 6) And (tOSVI.dwMinorVersion >= 3))

End Function

'Check for a version >= Win 10.
' (NOTE: this requires a manifest, so don't rely on it in the IDE.  Also, MS doesn't guarantee that this check will
'  remain valid forever, though it does appear to work in Windows 10 TP builds.)
Public Function IsOSWin10OrLater() As Boolean

    Dim tOSVI As OSVERSIONINFOEX
    tOSVI.dwOSVersionInfoSize = Len(tOSVI)
    GetVersionEx tOSVI
    
    IsOSWin10OrLater = (tOSVI.dwMajorVersion > 6) Or ((tOSVI.dwMajorVersion = 6) And (tOSVI.dwMinorVersion >= 4))

End Function

'Return the current OS version as a string.  (At present, this data is added to debug logs.)
Public Function GetOSVersionAsString() As String
    
    'Retrieve OS version data
    Dim tOSVI As OSVERSIONINFOEX
    tOSVI.dwOSVersionInfoSize = Len(tOSVI)
    GetVersionEx tOSVI
    
    Dim osName As String
    
    Select Case tOSVI.dwMajorVersion
        
        Case 10
            osName = "Windows 10"
        
        Case 6
            
            Select Case tOSVI.dwMinorVersion
                
                Case 4
                    osName = "Windows 10 Technical Preview"
                    
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
                    
                Case Else
                    osName = "(Unknown 6.x variant)"
            
            End Select
        
        Case 5
            osName = "Windows XP"
            
        Case Else
            osName = "(Unknown OS?)"
    
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
    
    Dim cUnicode As pdUnicode
    Set cUnicode = New pdUnicode
    
    Dim buildString As String
    buildString = Trim$(cUnicode.TrimNull(tOSVI.szCSDVersion))
    
    With tOSVI
        GetOSVersionAsString = osName & IIf(Len(buildString) <> 0, " " & buildString, "") & osBitness & "(" & .dwMajorVersion & "." & .dwMinorVersion & "." & .dwBuildNumber & ")"
    End With

End Function

'Return the number of logical cores on this system
Public Function GetNumLogicalCores() As Long
    
    Dim tSYSINFO As SYSTEM_INFO
    Call GetNativeSystemInfo(tSYSINFO)
    
    GetNumLogicalCores = tSYSINFO.dwNumberOfProcessors

End Function

'Return a list of PD-relevant processor features, in string format
Public Function GetProcessorFeatures() As String

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
    If Len(listFeatures) <> 0 Then
        GetProcessorFeatures = Left$(listFeatures, Len(listFeatures) - 2)
    Else
        GetProcessorFeatures = "(none)"
    End If
    
End Function

'Query total system RAM
Public Function GetTotalSystemRAM() As String

    Dim memStatus As MemoryStatusEx
    memStatus.dwLength = Len(memStatus)
    Call GlobalMemoryStatusEx(memStatus)
    
    GetTotalSystemRAM = CStr(Int(CDbl(memStatus.ullTotalPhys / 1024) * 10)) & " MB"
    
End Function
