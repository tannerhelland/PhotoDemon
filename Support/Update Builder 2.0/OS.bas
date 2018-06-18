Attribute VB_Name = "OS"
'***************************************************************************
'OS-specific specific features
'Copyright 2013-2018 by Tanner Helland
'Created: 21/November/13
'Last updated: 18/July/17
'Last update: migrate various OS-specific bits to this module
'
'Newer Windows versions expose some neat features (like progress bars overlaying the taskbar), and PhotoDemon
' tries to make use of them when relevant.  Similarly, some OS-level features are not easily mimicked from within VB
' (like Unicode-aware command-line processing), and I've tried to encapsulate those features here.
'
'Special thanks for this module include:
' - "AndRAY": http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=72856&lngWId=1
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This module is automatically enabled/disabled based on the current OS version.  If you want these
' features disabled even on valid OS versions, you can set this failsafe constant to FALSE.
Private Const WIN7_FEATURES_ALLOWED As Boolean = True

Private Const CLSID_TASKBARLIST As String = "{56FDF344-FD6D-11d0-958A-006097C9A090}"
Private Const IID_ITASKBARLIST3 As String = "{EA1AFB91-9E28-4B86-90E9-9E9F8A5EEFAF}"

Private Const ASM_CALL_REL32 As Byte = &HE8
Private Const ASM_PUSH_IMM32 As Byte = &H68
Private Const UNK_Release As Long = 2

Private Const TH32CS_SNAPPROCESS As Long = 2
Private Const GW_OWNER As Long = 4
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const MAX_PATH As Long = 260
Private Const MEM_COMMIT As Long = &H1000&
Private Const MEM_RELEASE As Long = &H8000&
Private Const MEM_RESERVE As Long = &H2000&
Private Const PAGE_EXECUTE_READWRITE As Long = &H40&
Private Const PROCESS_QUERY_INFORMATION As Long = 1024
Private Const PROCESS_VM_READ As Long = 16
Private Const PROCESSOR_ARCHITECTURE_AMD64 As Long = 9        'x64 (AMD or Intel)
Private Const PROCESSOR_ARCHITECTURE_IA64 As Long = 6         'Intel Itanium Processor Family (IPF)
Private Const PROCESSOR_ARCHITECTURE_INTEL As Long = 0
Private Const PROCESSOR_ARCHITECTURE_UNKNOWN As Long = &HFFFF&
Private Const SHGFP_TYPE_CURRENT As Long = &H0                'current value for user, verify it exists
Private Const VER_NT_WORKSTATION As Long = &H1&

Public Enum OS_CSIDL
    CSIDL_ADMINTOOLS = &H30
    CSIDL_ALTSTARTUP = &H1D
    CSIDL_APPDATA = &H1A
    CSIDL_BITBUCKET = &HA
    CSIDL_CDBURN_AREA = &H3B
    CSIDL_COMMON_ADMINTOOLS = &H2F
    CSIDL_COMMON_ALTSTARTUP = &H1E
    CSIDL_COMMON_APPDATA = &H23
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19
    CSIDL_COMMON_DOCUMENTS = &H2E
    CSIDL_COMMON_FAVORITES = &H1F
    CSIDL_COMMON_MUSIC = &H35
    CSIDL_COMMON_OEM_LINKS = &H3A
    CSIDL_COMMON_PICTURES = &H36
    CSIDL_COMMON_PROGRAMS = &H17
    CSIDL_COMMON_STARTMENU = &H16
    CSIDL_COMMON_STARTUP = &H18
    CSIDL_COMMON_TEMPLATES = &H2D
    CSIDL_COMMON_VIDEO = &H37
    CSIDL_COMPUTERSNEARME = &H3D
    CSIDL_CONNECTIONS = &H31
    CSIDL_CONTROLS = &H3
    CSIDL_COOKIES = &H21
    CSIDL_DESKTOP = &H0
    CSIDL_DESKTOPDIRECTORY = &H10
    CSIDL_DRIVES = &H11
    CSIDL_FAVORITES = &H6
    CSIDL_FLAG_CREATE = &H8000
    CSIDL_FLAG_DONT_VERIFY = &H4000
    CSIDL_FLAG_MASK = &HFF00
    CSIDL_FLAG_NO_ALIAS = &H1000
    CSIDL_FLAG_PER_USER_INIT = &H800
    CSIDL_FONTS = &H14
    CSIDL_HISTORY = &H22
    CSIDL_INTERNET = &H1
    CSIDL_INTERNET_CACHE = &H20
    CSIDL_LOCAL_APPDATA = &H1C
    CSIDL_MYDOCUMENTS = &HC
    CSIDL_MYMUSIC = &HD
    CSIDL_MYPICTURES = &H27
    CSIDL_MYVIDEO = &HE
    CSIDL_NETHOOD = &H13
    CSIDL_NETWORK = &H12
    CSIDL_PERSONAL = &H5
    CSIDL_PRINTERS = &H4
    CSIDL_PRINTHOOD = &H1B
    CSIDL_PROFILE = &H28
    CSIDL_PROGRAM_FILES = &H26
    CSIDL_PROGRAM_FILES_COMMON = &H2B
    CSIDL_PROGRAM_FILES_COMMONX86 = &H2C
    CSIDL_PROGRAM_FILESX86 = &H2A
    CSIDL_PROGRAMS = &H2
    CSIDL_RECENT = &H8
    CSIDL_RESOURCES = &H38
    CSIDL_RESOURCES_LOCALIZED = &H39
    CSIDL_SENDTO = &H9
    CSIDL_STARTMENU = &HB
    CSIDL_STARTUP = &H7
    CSIDL_SYSTEM = &H25
    CSIDL_SYSTEMX86 = &H29
    CSIDL_TEMPLATES = &H15
    CSIDL_WINDOWS = &H24
End Enum

#If False Then
    Private Const CSIDL_ADMINTOOLS = &H30, CSIDL_ALTSTARTUP = &H1D, CSIDL_APPDATA = &H1A, CSIDL_BITBUCKET = &HA, CSIDL_CDBURN_AREA = &H3B, CSIDL_COMMON_ADMINTOOLS = &H2F, CSIDL_COMMON_ALTSTARTUP = &H1E, CSIDL_COMMON_APPDATA = &H23, CSIDL_COMMON_DESKTOPDIRECTORY = &H19, CSIDL_COMMON_DOCUMENTS = &H2E, CSIDL_COMMON_FAVORITES = &H1F, CSIDL_COMMON_MUSIC = &H35, CSIDL_COMMON_OEM_LINKS = &H3A, CSIDL_COMMON_PICTURES = &H36, CSIDL_COMMON_PROGRAMS = &H17, CSIDL_COMMON_STARTMENU = &H16, CSIDL_COMMON_STARTUP = &H18, CSIDL_COMMON_TEMPLATES = &H2D, CSIDL_COMMON_VIDEO = &H37, CSIDL_COMPUTERSNEARME = &H3D, CSIDL_CONNECTIONS = &H31
    Private Const CSIDL_CONTROLS = &H3, CSIDL_COOKIES = &H21, CSIDL_DESKTOP = &H0, CSIDL_DESKTOPDIRECTORY = &H10, CSIDL_DRIVES = &H11, CSIDL_FAVORITES = &H6, CSIDL_FLAG_CREATE = &H8000, CSIDL_FLAG_DONT_VERIFY = &H4000, CSIDL_FLAG_MASK = &HFF00, CSIDL_FLAG_NO_ALIAS = &H1000, CSIDL_FLAG_PER_USER_INIT = &H800, CSIDL_FONTS = &H14, CSIDL_HISTORY = &H22, CSIDL_INTERNET = &H1, CSIDL_INTERNET_CACHE = &H20, CSIDL_LOCAL_APPDATA = &H1C, CSIDL_MYDOCUMENTS = &HC, CSIDL_MYMUSIC = &HD, CSIDL_MYPICTURES = &H27, CSIDL_MYVIDEO = &HE, CSIDL_NETHOOD = &H13, CSIDL_NETWORK = &H12, CSIDL_PERSONAL = &H5, CSIDL_PRINTERS = &H4, CSIDL_PRINTHOOD = &H1B
    Private Const CSIDL_PROFILE = &H28, CSIDL_PROGRAM_FILES = &H26, CSIDL_PROGRAM_FILES_COMMON = &H2B, CSIDL_PROGRAM_FILES_COMMONX86 = &H2C, CSIDL_PROGRAM_FILESX86 = &H2A, CSIDL_PROGRAMS = &H2, CSIDL_RECENT = &H8, CSIDL_RESOURCES = &H38, CSIDL_RESOURCES_LOCALIZED = &H39, CSIDL_SENDTO = &H9, CSIDL_STARTMENU = &HB, CSIDL_STARTUP = &H7, CSIDL_SYSTEM = &H25, CSIDL_SYSTEMX86 = &H29, CSIDL_TEMPLATES = &H15, CSIDL_WINDOWS = &H24
#End If

Public Enum OS_ProcessorFeature
    PF_ARM_64BIT_LOADSTORE_ATOMIC = 25 'The 64-bit load/store atomic instructions are available.
    PF_ARM_DIVIDE_INSTRUCTION_AVAILABLE = 24 'The divide instructions are available.
    PF_ARM_EXTERNAL_CACHE_AVAILABLE = 26 'The external cache is available.
    PF_ARM_FMAC_INSTRUCTIONS_AVAILABLE = 27 'The floating-point multiply-accumulate instruction is available.
    PF_ARM_VFP_32_REGISTERS_AVAILABLE = 18 'The VFP/Neon: 32 x 64bit register bank is present. This flag has the same meaning as PF_ARM_VFP_EXTENDED_REGISTERS.
    PF_3DNOW_INSTRUCTIONS_AVAILABLE = 7 'The 3D-Now instruction set is available.
    PF_CHANNELS_ENABLED = 16 'The processor channels are enabled.
    PF_COMPARE_EXCHANGE_DOUBLE = 2 'The atomic compare and exchange operation (cmpxchg) is available.
    PF_COMPARE_EXCHANGE128 = 14 'The atomic compare and exchange 128-bit operation (cmpxchg16b) is available.
    PF_COMPARE64_EXCHANGE128 = 15 'The atomic compare 64 and exchange 128-bit operation (cmp8xchg16) is available.
    PF_FASTFAIL_AVAILABLE = 23 '_fastfail() is available.
    PF_FLOATING_POINT_EMULATED = 1 'Floating-point operations are emulated using a software emulator.
    PF_FLOATING_POINT_PRECISION_ERRATA = 0 'On a Pentium, a floating-point precision error can occur in rare circumstances.
    PF_MMX_INSTRUCTIONS_AVAILABLE = 3 'The MMX instruction set is available.
    PF_NX_ENABLED = 12 'Data execution prevention is enabled.
    PF_PAE_ENABLED = 9 'The processor is PAE-enabled. For more information, see Physical Address Extension.
    PF_RDTSC_INSTRUCTION_AVAILABLE = 8 'The RDTSC instruction is available.
    PF_RDWRFSGSBASE_AVAILABLE = 22 'RDFSBASE, RDGSBASE, WRFSBASE, and WRGSBASE instructions are available.
    PF_SECOND_LEVEL_ADDRESS_TRANSLATION = 20 'Second Level Address Translation is supported by the hardware.
    PF_SSE3_INSTRUCTIONS_AVAILABLE = 13 'The SSE3 instruction set is available.
    PF_VIRT_FIRMWARE_ENABLED = 21 'Virtualization is enabled in the firmware.
    PF_XMMI_INSTRUCTIONS_AVAILABLE = 6 'The SSE instruction set is available.
    PF_XMMI64_INSTRUCTIONS_AVAILABLE = 10 'The SSE2 instruction set is available.
    PF_XSAVE_ENABLED = 17 'The processor implements the XSAVE and XRSTOR instructions.
End Enum

#If False Then
    Private Const PF_ARM_64BIT_LOADSTORE_ATOMIC = 25, PF_ARM_DIVIDE_INSTRUCTION_AVAILABLE = 24, PF_ARM_EXTERNAL_CACHE_AVAILABLE = 26, PF_ARM_FMAC_INSTRUCTIONS_AVAILABLE = 27, PF_ARM_VFP_32_REGISTERS_AVAILABLE = 18, PF_3DNOW_INSTRUCTIONS_AVAILABLE = 7, PF_CHANNELS_ENABLED = 16, PF_COMPARE_EXCHANGE_DOUBLE = 2, PF_COMPARE_EXCHANGE128 = 14, PF_COMPARE64_EXCHANGE128 = 15, PF_FASTFAIL_AVAILABLE = 23, PF_FLOATING_POINT_EMULATED = 1, PF_FLOATING_POINT_PRECISION_ERRATA = 0, PF_MMX_INSTRUCTIONS_AVAILABLE = 3
    Private Const PF_NX_ENABLED = 12, PF_PAE_ENABLED = 9, PF_RDTSC_INSTRUCTION_AVAILABLE = 8, PF_RDWRFSGSBASE_AVAILABLE = 22, PF_SECOND_LEVEL_ADDRESS_TRANSLATION = 20, PF_SSE3_INSTRUCTIONS_AVAILABLE = 13, PF_VIRT_FIRMWARE_ENABLED = 21, PF_XMMI_INSTRUCTIONS_AVAILABLE = 6, PF_XMMI64_INSTRUCTIONS_AVAILABLE = 10, PF_XSAVE_ENABLED = 17
#End If

'Similar APIs for retrieving GDI and user objects
Public Enum PD_GuiResources
    PDGR_GdiObjects = 0
    PDGR_UserObjects = 1
    PDGR_GdiObjectsPeak = 2
    PDGR_UserObjectsPeak = 4
End Enum

#If False Then
    Private Const PDGR_GdiObjects = 0, PDGR_GdiObjectsPeak = 2, PDGR_UserObjects = 1, PDGR_UserObjectsPeak = 4
#End If

'Possible task bar progress states.  PD is primarily interested in NOPROGRESS and NORMAL
Public Enum PD_TaskBarProgress
    TBP_NoProgress = 0
    TBP_Indeterminate = 1
    TBP_Normal = 2
    TBP_Error = 4
    TBP_Paused = 8
End Enum

#If False Then
    Private Const TBP_NoProgress = 0, TBP_Indeterminate = 1, TBP_Normal = 2, TBP_Error = 4, TBP_Paused = 8
#End If

Private Enum ITaskbarList3Members
                                '/* ITaskbarList methods */
    HrInit_ = 3                 'STDMETHOD( HrInit )( THIS ) PURE;
    AddTab_ = 4                 'STDMETHOD( AddTab )( THIS_ HWND ) PURE;
    DeleteTab_ = 5              'STDMETHOD( DeleteTab )( THIS_ HWND ) PURE;
    ActivateTab_ = 6            'STDMETHOD( ActivateTab )( THIS_ HWND ) PURE;
    SetActiveAlt_ = 7           'STDMETHOD( SetActiveAlt )( THIS_ HWND ) PURE;
                                '/* ITaskbarList2 methods */
    MarkFullscreenWindow_ = 8   'STDMETHOD( MarkFullscreenWindow )( THIS_ HWND, BOOL ) PURE;
                                '/* ITaskbarList3 methods */
    SetProgressValue_ = 9       'STDMETHOD( SetProgressValue )( THIS_ HWND, ULONGLONG, ULONGLONG ) PURE;
    SetProgressState_ = 10      'STDMETHOD( SetProgressState )( THIS_ HWND, TBPFLAG ) PURE;
    RegisterTab_ = 11           'STDMETHOD( RegisterTab )( THIS_ HWND, HWND ) PURE;
    UnregisterTab_ = 12         'STDMETHOD( UnregisterTab )( THIS_ HWND ) PURE;
    SetTabOrder_ = 13           'STDMETHOD( SetTabOrder )( THIS_ HWND, HWND ) PURE;
    SetTabActive_ = 14          'STDMETHOD( SetTabActive )( THIS_ HWND, HWND, DWORD ) PURE;
    ThumbBarAddButtons_ = 15    'STDMETHOD( ThumbBarAddButtons )( THIS_ HWND, UINT, LPTHUMBBUTTON ) PURE;
    ThumbBarUpdateButtons_ = 16 'STDMETHOD( ThumbBarUpdateButtons )( THIS_ HWND, UINT, LPTHUMBBUTTON ) PURE;
    ThumbBarSetImageList_ = 17  'STDMETHOD( ThumbBarSetImageList )( THIS_ HWND, HIMAGELIST ) PURE;
    SetOverlayIcon_ = 18        'STDMETHOD( SetOverlayIcon )( THIS_ HWND, HICON, LPCWSTR ) PURE;
    SetThumbnailTooltip_ = 19   'STDMETHOD( SetThumbnailTooltip )( THIS_ HWND, LPCWSTR ) PURE;
    SetThumbnailClip_ = 20      'STDMETHOD( SetThumbnailClip )( THIS_ HWND, RECT * ) PURE;
'                                '/* ITaskbarList4 methods */
'    SetTabProperties_ = 21      'STDMETHOD( SetTabProperties )( THIS_ HWND, STPFLAG ) PURE;
End Enum

Private Type OS_Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type OS_MemoryStatusEx
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

Private Type OS_ProcessEntry32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile(0 To MAX_PATH * 2 - 1) As Byte
End Type

Private Type OS_ProcessMemoryCounter
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

Private Type OS_SystemInfo
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

Private Type OS_SystemTime
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type OS_VersionInfoEx
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion(0 To 255) As Byte
    wServicePackMajor  As Integer
    wServicePackMinor  As Integer
    wSuiteMask         As Integer
    wProductType       As Byte
    wReserved          As Byte
End Type

Private Declare Function CloseHandle Lib "kernel32" (ByVal Handle As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function GetCommandLineW Lib "kernel32" () As Long
Private Declare Function GetModuleFileNameW Lib "kernel32" (ByVal hModule As Long, ByVal ptrToFileNameBuffer As Long, ByVal nSize As Long) As Long
Private Declare Sub GetNativeSystemInfo Lib "kernel32" (ByRef lpSystemInfo As OS_SystemInfo)
Private Declare Sub GetSystemTimeAsFileTime Lib "kernel32" (ByRef dstTime As Currency)
Private Declare Function GetTempPathW Lib "kernel32" (ByVal nBufferLength As Long, ByVal lpStrBuffer As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExW" (ByVal lpVersionInformation As Long) As Long
Private Declare Function GlobalMemoryStatusEx Lib "kernel32" (ByRef lpBuffer As OS_MemoryStatusEx) As Long
Private Declare Function IsProcessorFeaturePresent Lib "kernel32" (ByVal procFeature As Long) As Boolean
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32FirstW" (ByVal hSnapShot As Long, ByVal ptrToUProcess As Long) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32NextW" (ByVal hSnapShot As Long, ByVal ptrToUProcess As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long

Private Declare Function SysAllocString Lib "oleaut32" (ByVal ptr As Long) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As String, ByRef pClsID As OS_Guid) As Long
Private Declare Function CoCreateGuid Lib "ole32" (ByRef pGuid As OS_Guid) As Long
Private Declare Function CoCreateInstance Lib "ole32" (ByRef rClsID As OS_Guid, ByVal pUnkOuter As Long, ByVal dwClsContext As Long, ByRef rIID As OS_Guid, ByRef ppv As Any) As Long
Private Declare Function IIDFromString Lib "ole32" (ByVal lpsz As String, ByRef lpiid As OS_Guid) As Long
Private Declare Function StringFromGUID2 Lib "ole32" (ByRef rguid As Any, ByVal lpstrClsID As Long, ByVal cbMax As Long) As Long

Private Declare Function GetProcessMemoryInfo Lib "psapi" (ByVal hProcess As Long, ByRef ppsmemCounters As OS_ProcessMemoryCounter, ByVal cb As Long) As Long

Private Declare Function CommandLineToArgvW Lib "shell32" (ByVal lpCmdLine As Long, ByRef pNumArgs As Long) As Long
Private Declare Function SHGetFolderPathW Lib "shfolder" (ByVal hWndOwner As Long, ByVal nFolder As OS_CSIDL, ByVal hToken As Long, ByVal dwReserved As Long, ByVal lpszPath As Long) As Long

Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long 'Used to intercept and process VB-window messages (hence the -A variant)
Private Declare Function FindWindowW Lib "user32" (Optional ByVal lpClassName As Long, Optional ByVal lpWindowName As Long) As Long
Private Declare Function GetGuiResources Lib "user32" (ByVal hProcess As Long, ByVal resourceToCheck As PD_GuiResources) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal uCmd As Long) As Long

Private Declare Function OpenThemeData Lib "uxtheme" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme" (ByVal hTheme As Long) As Long

'A persistent handle to the OLE interface we create for Win7+ taskbar features
Private m_taskbarObjHandle As Long

'Similarly, to improve performance, the first call to GetVersionEx is cached, and subsequent calls just use the cached value.
Private m_OSVI As OS_VersionInfoEx, m_VersionInfoCached As Boolean

'PD marks some data with unique "session ID" values, which lets it distinguish between parallel executions.
' This value is generated once, on first demand, and cached locally.
Private m_SessionID As String

'Various bits of VB hackery require us to interact directly with ThunderMain.  The first time such an action is invoked,
' we cache the relevant hWnd (as it's a pain to retrieve).
Private m_ThunderMainHwnd As Long

'Retrieving PD's process ID is energy-intensive.  Once we've retrieved it for a session, we cache the ID.
' (Also, if we failed to retrieve the ID on a previous attempt, we cache that as well, so we don't waste time
'  trying again.)
Private m_AppProcID As Long, m_TriedToRetrieveID As Boolean

'This module caches whether or not Aero is enabled.  (We require this for various UI interop bits.)
Private Enum PD_ThemingAvailable
    pdta_Unknown = 0
    pdta_False = 1
    pdta_True = 2
End Enum

#If False Then
    Private Const pdta_Unknown = 0, pdta_False = 1, pdta_True = 2
#End If

Private m_ThemingAvailable As PD_ThemingAvailable

'Function for returning this application's current memory usage.  Note that this function will return warped
' values inside the IDE (because the reported total is for *all* of VB6, including the IDE itself).
Public Function AppMemoryUsage(Optional returnPeakValue As Boolean = False) As Long
    
    'Open a handle to this process
    Dim procHandle As Long
    procHandle = AppProcessHandle()
    If (procHandle <> 0) Then
                
        'Attempt to retrieve process memory information
        Dim procMemInfo As OS_ProcessMemoryCounter
        procMemInfo.cb = LenB(procMemInfo)
        
        If (GetProcessMemoryInfo(procHandle, procMemInfo, procMemInfo.cb) <> 0) Then
            
            If returnPeakValue Then
                AppMemoryUsage = procMemInfo.PeakWorkingSetSize / 1024#
            Else
                AppMemoryUsage = procMemInfo.WorkingSetSize / 1024#
            End If
            
        End If
        
        'Release the retrieved handle
        CloseHandle procHandle
        
    Else
        InternalError "OS.AppMemoryUsage() failed to open a handle to this process."
    End If
    
End Function

'Return a read-and-query-access handle to this process.
' NOTE: inside the IDE, this returns a handle to vb6.exe instead.
' ANOTHER NOTE: *the caller is responsible for freeing this handle* when done with it.
Private Function AppProcessHandle() As Long
    
    Dim procID As Long
    procID = AppProcessID()
    
    If (procID <> 0) Then
        AppProcessHandle = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0&, procID)
    Else
        AppProcessHandle = 0
    End If
    
End Function

Private Function AppProcessID() As Long

    'If we've already retrieved a handle this session, return it immediately
    If (m_AppProcID <> 0) Then
        AppProcessID = m_AppProcID
    
    'If we haven't retrieved it, do so now.
    ElseIf (Not m_TriedToRetrieveID) Then
        
        Dim hSnapShot As Long
        Dim uProcess As OS_ProcessEntry32
    
        'Prep a process enumerator.  We're going to search the active process list, looking for PhotoDemon.
        hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
        If (hSnapShot <> INVALID_HANDLE_VALUE) Then
    
            'Attempt to enumerate the first entry in the list
            uProcess.dwSize = LenB(uProcess)
            If (ProcessFirst(hSnapShot, VarPtr(uProcess)) <> 0) Then
            
                'The enumerator is working correctly.  Check each uProcess entry for this application's name.
                ' (NOTE!  Inside the IDE, this only works if PhotoDemon is the only running VB6 window.
                '  Otherwise, another VB6 window may be accidentally grabbed, leading to weird outcomes.)
                Dim processName As String
                
                'MAX_PATH no longer applies, but the docs are unclear on a reasonable buffer size.  Because buffer behavior is sketchy on XP
                ' (see https://msdn.microsoft.com/en-us/library/windows/desktop/ms683197%28v=vs.85%29.aspx) it's easier to just go with a
                ' huge buffer, then manually trim the result.
                Dim TEMPORARY_LARGE_BUFFER As Long
                TEMPORARY_LARGE_BUFFER = 1024
                
                Dim tmpString As String
                tmpString = String$(TEMPORARY_LARGE_BUFFER, 0)
                
                If (GetModuleFileNameW(0&, StrPtr(tmpString), TEMPORARY_LARGE_BUFFER \ 2) <> 0) Then
                    processName = Strings.TrimNull(tmpString)
                    processName = Files.FileGetName(processName)
                Else
                    InternalError "AppProcessID failed to retrieve the current process name."
                End If
                
                Dim testProcessName As String, procFound As Boolean
                procFound = False
                
                Do
                    testProcessName = Strings.StringFromUTF16_FixedLen(VarPtr(uProcess.szExeFile(0)), MAX_PATH * 2, True)
                    
                    If (Len(testProcessName) <> 0) Then
                        If Strings.StringsEqual(testProcessName, processName, True) Then
                            procFound = True
                            Exit Do
                        End If
                    End If
                    
                Loop While (ProcessNext(hSnapShot, VarPtr(uProcess)) <> 0)
                
                'If we found PD's process handle, cache it!
                If procFound Then
                    m_AppProcID = uProcess.th32ProcessID
                Else
                    m_AppProcID = 0
                    InternalError "OS.AppProcessID() failed to locate this process."
                End If
                
            Else
                m_AppProcID = 0
                InternalError "OS.AppProcessID() failed to initiate a ProcessFirst()-based search."
            End If
            
            'Regardless of outcome, close the ToolHelp enumerator when we're done
            CloseHandle hSnapShot
        
        Else
            m_AppProcID = 0
            InternalError "OS.AppProcessID() failed to create a ToolHelp snapshot."
        End If
        
        'Regardless of outcome, note that we've tried to retrieve the process ID
        m_TriedToRetrieveID = True
        AppProcessID = m_AppProcID
        
    Else
        AppProcessID = 0
    End If

End Function

'Return this application's current GDI or User object count.  (On Win 7 or later, peak usage can also be returned.)
Public Function AppResourceUsage(Optional ByVal resourceType As PD_GuiResources = PDGR_GdiObjects) As Long
    
    'Open a handle to this process
    Dim procHandle As Long
    procHandle = AppProcessHandle()
    If (procHandle <> 0) Then
                
        'Attempt to retrieve resource information.  Note that certain resource types are restricted by OS version.
        AppResourceUsage = GetGuiResources(procHandle, resourceType)
        
        'Release our process handle
        CloseHandle procHandle
        
    Else
        InternalError "OS.AppResourceUsage() failed to open a handle to this process."
    End If
    
End Function

'Many places in PD need to know the current Windows version, so they can enable/disable features accordingly.  To avoid
' constantly retrieving that info via APIs, we retrieve it once - at first request - then cache it locally.
Private Sub CacheOSVersion()
    If (Not m_VersionInfoCached) Then
        m_OSVI.dwOSVersionInfoSize = Len(m_OSVI)
        GetVersionEx VarPtr(m_OSVI)
        m_VersionInfoCached = True
    End If
End Sub

'Request an OLE interface from within VB.  I apologize for a lack of comments in this function, but I did not write it.
' For additional details, please see the original project, available here: http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=72856&lngWId=1
Private Function CallInterface(ByVal pInterface As Long, ByVal Member As Long, ByVal ParamsCount As Long, Optional ByVal p1 As Long = 0, Optional ByVal p2 As Long = 0, Optional ByVal p3 As Long = 0, Optional ByVal p4 As Long = 0, Optional ByVal p5 As Long = 0, Optional ByVal p6 As Long = 0, Optional ByVal p7 As Long = 0, Optional ByVal p8 As Long = 0, Optional ByVal p9 As Long = 0, Optional ByVal p10 As Long = 0) As Long
        
    Dim i As Long, t As Long
    Dim hGlobal As Long, hGlobalOffset As Long
    
    If (ParamsCount < 0) Then Err.Raise 5
    If (pInterface = 0) Then Err.Raise 5
    
    'Rewritten by Tanner: VirtualAlloc is required to not make DEP angry
    hGlobal = VirtualAlloc(0&, 5 * ParamsCount + 5 + 5 + 3 + 1, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If (hGlobal = 0) Then
        Debug.Print "WARNING!  Windows.CallInterface() failed to allocate virtual memory.  Exiting prematurely."
        Exit Function
    End If
    hGlobalOffset = hGlobal
    
    If (ParamsCount > 0) Then
        
        t = VarPtr(p1)
        
        For i = ParamsCount - 1 To 0 Step -1
          PutMem2 hGlobalOffset, ASM_PUSH_IMM32
          hGlobalOffset = hGlobalOffset + 1
          GetMem4 t + i * 4, ByVal hGlobalOffset
          hGlobalOffset = hGlobalOffset + 4
        Next
      
    End If
    
    PutMem2 hGlobalOffset, ASM_PUSH_IMM32
    hGlobalOffset = hGlobalOffset + 1
    PutMem4 hGlobalOffset, pInterface
    hGlobalOffset = hGlobalOffset + 4
    
    PutMem2 hGlobalOffset, ASM_CALL_REL32
    hGlobalOffset = hGlobalOffset + 1
    GetMem4 pInterface, ByVal VarPtr(t)
    GetMem4 t + Member * 4, ByVal VarPtr(t)
    PutMem4 hGlobalOffset, t - hGlobalOffset - 4
    hGlobalOffset = hGlobalOffset + 4
      
    PutMem4 hGlobalOffset, &H10C2&
    
    CallInterface = CallWindowProcA(hGlobal, 0, 0, 0, 0)
    
    'Edit by Tanner: match VirtualAlloc(), above
    If (VirtualFree(hGlobal, 0&, MEM_RELEASE) = 0) Then
        Debug.Print "WARNING!  Windows.CallInterface() failed to release virtual memory @" & hGlobal & ".  Please investigate."
    End If
  
End Function

'Return a Unicode-friendly copy of PD's command line params, pre-parsed into individual arguments.
' By default, the standard exe path entry is removed.  (This behavior can be toggled via the "removeExePath" parameter.)
' Returns: TRUE if argument count > 0; FALSE otherwise.
'          If TRUE is returned, dstStringStack is guaranteed to be initialized.
Public Function CommandW(ByRef dstStringStack As pdStringStack, Optional ByVal removeExePath As Boolean = True) As Boolean
    
    Dim fullCmdLine As String
    
    'If inside the IDE, use VB's regular command-line; this allows test params set via Project Properties to still work
    If (Not OS.IsProgramCompiled) Then
        fullCmdLine = Command$
    
    'When compiled, a true Unicode-friendly command line is returned
    Else
        fullCmdLine = Strings.StringFromCharPtr(GetCommandLineW(), True)
    End If
    
    'Next, we want to pre-parse the string into individual arguments using WAPI
    If (Len(fullCmdLine) <> 0) Then
    
        Dim lPtr As Long, numArgs As Long
        lPtr = CommandLineToArgvW(StrPtr(fullCmdLine), numArgs)
        
        'lPtr now points to the first (of potentially many) string pointers, each one a command-line argument.
        ' We want to assume control over each string in turn, and add each to our destination pdStringStack object.
        If (dstStringStack Is Nothing) Then
            Set dstStringStack = New pdStringStack
        Else
            dstStringStack.resetStack
        End If
        
        If (numArgs > 0) Then
        
            Dim i As Long, tmpString As String, tmpPtr As Long
            For i = 0 To numArgs - 1
                
                'Retrieve the next pointer
                CopyMemoryStrict VarPtr(tmpPtr), lPtr + 4 * i, 4&
                
                'Allocate a matching string (over which we have ownership)
                PutMem4 VarPtr(tmpString), SysAllocString(tmpPtr)
                
                'Conditionally add it to the string stack, depending on the removeExePath setting
                If removeExePath Then
                    If (InStr(1, tmpString, "PhotoDemon.exe", vbBinaryCompare) = 0) Then dstStringStack.AddString tmpString
                Else
                    dstStringStack.AddString tmpString
                End If
                
                'Free the temporary string
                tmpString = vbNullString
                
            Next
            
            CommandW = (dstStringStack.getNumOfStrings <> 0)
        
        End If
        
        'Free the original arg pointer (which frees the corresponding system-controlled string references as well)
        ' Details here: https://msdn.microsoft.com/en-us/library/windows/desktop/bb776391%28v=vs.85%29.aspx?f=255&MSPPError=-2147217396
        If (lPtr <> 0) Then LocalFree lPtr
    
    End If
    
End Function

'Sometimes, a unique string is needed.  Use this function to retrieve an arbitrary GUID from WAPI.
Public Function GetArbitraryGUID(Optional ByVal stripNonHexCharacters As Boolean = False) As String

    'Fill a GUID struct with data via WAPI
    Dim tmpGuid As OS_Guid
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
        guidString = Replace$(guidString, "-", vbNullString)
        
    End If
    
    GetArbitraryGUID = guidString

End Function

'Want to retrieve the current system time, using APIs, while auto-translating it from the (terrible) SystemTime
' struct to a more usable longlong?  Use this function.
Public Function GetSystemTimeAsCurrency() As Currency
    GetSystemTimeAsFileTime GetSystemTimeAsCurrency
End Function

'Is Aero enabled (requires Vista+ and classic theme must *not* be in use)
Public Function IsAeroAvailable() As Boolean
    
    'Only check this once; it does not change per-session
    If (m_ThemingAvailable = pdta_Unknown) Then
    
        'Windows XP is always false
        If (Not IsVistaOrLater) Then
            m_ThemingAvailable = pdta_False
        
        'Win 8+ always makes Aero available
        ElseIf IsWin8OrLater Then
            m_ThemingAvailable = pdta_True
            
        'Win Vista/7 must be dynamically detected
        Else
            Dim hTheme As Long, sClass As String
            sClass = "Window"
            hTheme = OpenThemeData(frmUpdate.hWnd, StrPtr(sClass))
            If (hTheme <> 0) Then
                m_ThemingAvailable = pdta_True
                CloseThemeData hTheme
            Else
                m_ThemingAvailable = pdta_False
            End If
        End If
    
    End If
    
    IsAeroAvailable = (m_ThemingAvailable = pdta_True)
    
End Function

'Is this program instance compiled, or running from the IDE?
Public Function IsProgramCompiled() As Boolean
    IsProgramCompiled = (App.LogMode = 1)
End Function

'Check for a version >= the specified version.
Public Function IsVistaOrLater() As Boolean
    If (Not m_VersionInfoCached) Then CacheOSVersion
    IsVistaOrLater = (m_OSVI.dwMajorVersion >= 6)
End Function

Public Function IsWin7OrLater() As Boolean
    If (Not m_VersionInfoCached) Then CacheOSVersion
    IsWin7OrLater = (m_OSVI.dwMajorVersion > 6) Or ((m_OSVI.dwMajorVersion = 6) And (m_OSVI.dwMinorVersion >= 1))
End Function

Public Function IsWin8OrLater() As Boolean
    If (Not m_VersionInfoCached) Then CacheOSVersion
    IsWin8OrLater = (m_OSVI.dwMajorVersion > 6) Or ((m_OSVI.dwMajorVersion = 6) And (m_OSVI.dwMinorVersion >= 2))
End Function

Public Function IsWin81OrLater() As Boolean
    If (Not m_VersionInfoCached) Then CacheOSVersion
    IsWin81OrLater = (m_OSVI.dwMajorVersion > 6) Or ((m_OSVI.dwMajorVersion = 6) And (m_OSVI.dwMinorVersion >= 3))
End Function

' (NOTE: the Win-10 check requires a manifest, so don't rely on it in the IDE.  Also, MS doesn't guarantee that this
' check will remain valid forever, though it does work as of Windows 10-1703.)
Public Function IsWin10OrLater() As Boolean
    If (Not m_VersionInfoCached) Then CacheOSVersion
    IsWin10OrLater = (m_OSVI.dwMajorVersion > 6) Or ((m_OSVI.dwMajorVersion = 6) And (m_OSVI.dwMinorVersion >= 4))
End Function

'Return the number of logical cores on this system
Public Function LogicalCoreCount() As Long
    Dim tmpSysInfo As OS_SystemInfo
    GetNativeSystemInfo tmpSysInfo
    LogicalCoreCount = tmpSysInfo.dwNumberOfProcessors
End Function

'Return the current OS version as a string.  (This is basically a helper function for PD's debug logger.)
Public Function OSVersionAsString() As String
    
    CacheOSVersion
    Dim osName As String
    
    Select Case m_OSVI.dwMajorVersion
        
        Case 10
            osName = "Windows 10"
        
        Case 6
            
            If (m_OSVI.dwMinorVersion = 4) Then
                osName = "Windows 10 Technical Preview"
                    
            ElseIf (m_OSVI.dwMinorVersion = 3) Then
                If ((m_OSVI.wProductType And VER_NT_WORKSTATION) <> 0) Then
                    osName = "Windows 8.1"
                Else
                    osName = "Windows Server 2012 R2"
                End If
                    
            ElseIf (m_OSVI.dwMinorVersion = 2) Then
                If ((m_OSVI.wProductType And VER_NT_WORKSTATION) <> 0) Then
                    osName = "Windows 8"
                Else
                    osName = "Windows Server 2012"
                End If
                    
            ElseIf (m_OSVI.dwMinorVersion = 1) Then
                If (m_OSVI.wProductType And VER_NT_WORKSTATION) <> 0 Then
                    osName = "Windows 7"
                Else
                    osName = "Windows Server 2008 R2"
                End If
                
            ElseIf (m_OSVI.dwMinorVersion = 0) Then
                If ((m_OSVI.wProductType And VER_NT_WORKSTATION) <> 0) Then
                    osName = "Windows Vista"
                Else
                    osName = "Windows Server 2008"
                End If
                    
            Else
                osName = "(Unknown 6.x variant)"
            
            End If
        
        Case 5
            osName = "Windows XP"
            
        Case Else
            osName = "(Unknown OS?)"
    
    End Select
    
    'Retrieve 32/64 bit OS version
    Dim osBitness As String
    
    Dim tSYSINFO As OS_SystemInfo
    GetNativeSystemInfo tSYSINFO
    
    Select Case tSYSINFO.wProcessorArchitecture
    
        Case PROCESSOR_ARCHITECTURE_AMD64
            osBitness = " 64-bit "
            
        Case PROCESSOR_ARCHITECTURE_IA64
            osBitness = " Itanium "
            
        Case Else
            osBitness = " 32-bit "
    
    End Select
    
    Dim buildString As String
    buildString = Trim$(Strings.TrimNull(Strings.StringFromCharPtr(VarPtr(m_OSVI.szCSDVersion(0)), True)))
    
    With m_OSVI
        OSVersionAsString = osName & IIf(Len(buildString) <> 0, " " & buildString, vbNullString) & osBitness & "(" & .dwMajorVersion & "." & .dwMinorVersion & "." & .dwBuildNumber & ")"
    End With

End Function

'Return a list of PD-relevant processor features, in string format.  (At present, this is designed purely for
' debug reporting purposes, as PD does make use of some SSE and SSE2 features in places.)
Public Function ProcessorFeatures() As String

    Dim listFeatures As String
    If IsProcessorFeaturePresent(PF_3DNOW_INSTRUCTIONS_AVAILABLE) Then listFeatures = listFeatures & "3DNow!" & ", "
    If IsProcessorFeaturePresent(PF_MMX_INSTRUCTIONS_AVAILABLE) Then listFeatures = listFeatures & "MMX" & ", "
    If IsProcessorFeaturePresent(PF_XMMI_INSTRUCTIONS_AVAILABLE) Then listFeatures = listFeatures & "SSE" & ", "
    If IsProcessorFeaturePresent(PF_XMMI64_INSTRUCTIONS_AVAILABLE) Then listFeatures = listFeatures & "SSE2" & ", "
    If IsProcessorFeaturePresent(PF_SSE3_INSTRUCTIONS_AVAILABLE) Then listFeatures = listFeatures & "SSE3" & ", "
    If IsProcessorFeaturePresent(PF_NX_ENABLED) Then listFeatures = listFeatures & "DEP" & ", "
    If IsProcessorFeaturePresent(PF_VIRT_FIRMWARE_ENABLED) Then listFeatures = listFeatures & "Virtualization" & ", "
    
    'Trim the trailing comma and blank space before returning
    If (Len(listFeatures) <> 0) Then
        ProcessorFeatures = Left$(listFeatures, Len(listFeatures) - 2)
    Else
        'NOTE: normally we would apply a translation here, but since this is meant for internal debugging
        ' purposes only, en-US is okay
        ProcessorFeatures = "(none)"
    End If
    
End Function

'Query RAM available to PD, as a user-friendly string
Public Function RAM_Available() As String

    Dim memStatus As OS_MemoryStatusEx
    memStatus.dwLength = LenB(memStatus)
    If (GlobalMemoryStatusEx(memStatus) <> 0) Then
    
        Dim tmpString As String
        tmpString = Trim$(Str(Int(CDbl(memStatus.ullTotalVirtual / 1024#) * 10#))) & " MB"
        tmpString = tmpString & " (real), "
        tmpString = tmpString & Trim$(Str(Int(CDbl(memStatus.ullAvailPageFile / 1024#) * 10#))) & " MB"
        tmpString = tmpString & " (hypothetical)"
        
        RAM_Available = tmpString
    
    End If
    
End Function

Public Function RAM_CurrentLoad() As String
    Dim memStatus As OS_MemoryStatusEx
    memStatus.dwLength = LenB(memStatus)
    If (GlobalMemoryStatusEx(memStatus) <> 0) Then
        RAM_CurrentLoad = Format$(CDbl(memStatus.dwMemoryLoad) / 100#, "0%")
    End If
End Function

'Query total installed system RAM, as a user-friendly string
Public Function RAM_SystemTotal() As String

    Dim memStatus As OS_MemoryStatusEx
    memStatus.dwLength = LenB(memStatus)
    If (GlobalMemoryStatusEx(memStatus) <> 0) Then
        RAM_SystemTotal = Trim$(Str(Int(CDbl(memStatus.ullTotalPhys / 1024#) * 10#))) & " MB"
    End If
    
End Function

'If desired, a custom state can be set for the taskbar.  (Normally this is handled by the SetTaskbarProgressValue function.)
Public Function SetTaskbarProgressState(ByVal tbpFlags As PD_TaskBarProgress) As Long
    If WIN7_FEATURES_ALLOWED Then SetTaskbarProgressState = CallInterface(m_taskbarObjHandle, SetProgressState_, 2, frmUpdate.hWnd, tbpFlags)
End Function

Public Function SetTaskbarProgressValue(ByVal amtCompleted As Long, ByVal amtTotal As Long) As Long
    If WIN7_FEATURES_ALLOWED Then
        If (amtCompleted = 0) Then
            SetTaskbarProgressState TBP_NoProgress
        Else
            SetTaskbarProgressState TBP_Normal
            SetTaskbarProgressValue = CallInterface(m_taskbarObjHandle, SetProgressValue_, 5, frmUpdate.hWnd, amtCompleted, 0, amtTotal, 0)
        End If
    End If
End Function

'If the OS is detected as Windows 7+, this function will be called.  It will prepare a handle to the OLE interface
' we use for Win7-specific features.
Public Sub StartWin7PlusFeatures()

    'To disable this functionality (e.g during testing), change this line to FALSE.  It will prevent any further execution of Win7-specific features.
    If WIN7_FEATURES_ALLOWED Then
        Dim clsID As OS_Guid, InterfaceGuid As OS_Guid
        CLSIDFromString StrConv(CLSID_TASKBARLIST, vbUnicode), clsID
        IIDFromString StrConv(IID_ITASKBARLIST3, vbUnicode), InterfaceGuid
        CoCreateInstance clsID, 0, 1, InterfaceGuid, m_taskbarObjHandle
    End If
    
End Sub

'Make sure to release the interface when we're done with it!
Public Sub StopWin7PlusFeatures()
    If WIN7_FEATURES_ALLOWED Then
        If (m_taskbarObjHandle <> 0) Then CallInterface m_taskbarObjHandle, UNK_Release, 0
    End If
End Sub

'Get a special folder from Windows (as specified by the CSIDL)
Public Function SpecialFolder(ByVal folderType As OS_CSIDL) As String
    
    Dim dstPath As String
    dstPath = String$(MAX_PATH, 0)
    
    If (SHGetFolderPathW(0&, folderType, 0&, SHGFP_TYPE_CURRENT, StrPtr(dstPath)) = 0) Then
        SpecialFolder = Files.PathAddBackslash(Strings.TrimNull(dstPath))
    Else
        InternalError "OS.SpecialFolder failed to retrieve the folder with type: " & folderType
    End If
    
End Function

'Return the current Windows-specified temp directory
Public Function SystemTempPath() As String
    
    'Create a blank string (as required by the API call)
    Dim sRet As String
    sRet = String$(261, 0)
    
    'Fill that string with the temporary path
    Dim lngLen As Long
    lngLen = GetTempPathW(261, StrPtr(sRet))
    
    'If something went wrong, raise an error
    If (lngLen <> 0) Then
        SystemTempPath = Files.PathAddBackslash(Left$(sRet, lngLen))
    Else
        InternalError "OS.SystemTempPath() failed to retrieve a valid path from the API (#" & Err.LastDllError & ")"
    End If
    
End Function

'Get PD's master hWnd.  The value is cached after an initial call.  Based on a sample project by the ever-talented Bonnie West
' (http://www.vbforums.com/showthread.php?682474-VB6-ThunderMain-class).
'
'Note that this call *does* work in the IDE, but any subsequent calls that operate on the hWnd may be prone to failure
' and/or difficult-to-replicate bugs.
Public Function ThunderMainHWnd() As Long

    'If we already grabbed the hWnd this session, we can skip right to the end
    If (m_ThunderMainHwnd = 0) Then
        
        'If one or more forms exist, we can retrieve ThunderMain directly by grabbing the owner handle of any open form.
        If Forms.Count Then
            m_ThunderMainHwnd = GetWindow(Forms(0&).hWnd, GW_OWNER)
        
        'If no forms exist, we must retrieve the hWnd manually
        Else
        
            'Cache the current program title
            Dim strPrevTitle As String
            strPrevTitle = App.Title
            
            'Create a unique, temporary program title
            App.Title = OS.GetArbitraryGUID()
            
            'Find the window matching our new, arbitrary title
            If OS.IsProgramCompiled Then
                m_ThunderMainHwnd = FindWindowW(StrPtr("ThunderRT6Main"), StrPtr(App.Title))
            Else
                m_ThunderMainHwnd = FindWindowW(StrPtr("ThunderMain"), StrPtr(App.Title))
            End If
            
            'Restore the original title
            App.Title = strPrevTitle
            
        End If
    End If
    
    ThunderMainHWnd = m_ThunderMainHwnd

End Function

'Return a unique session ID for this PhotoDemon instance.  A session ID is generated by retrieving a random GUID,
' hashing it, then returning the first 16 characters from the hash.  So many random steps are not necessary, but
' they help ensure that the IDs are actually unique.
Public Function UniqueSessionID() As String
    
    If (Len(m_SessionID) = 0) Then
        Dim cCrypto As pdCrypto: Set cCrypto = New pdCrypto
        Dim tmpString As String
        tmpString = GetArbitraryGUID()
        m_SessionID = cCrypto.QuickHashString(tmpString, Len(tmpString))
        If (Len(m_SessionID) > 12) Then m_SessionID = Left$(m_SessionID, 12)
    End If
    
    UniqueSessionID = m_SessionID
    
End Function

'Generate a unique temp file name.  (I use arbitrary GUIDs to generate these, instead of the old GetTempFileName()
' kernel function - this is because that API has some annoying caveats, per https://msdn.microsoft.com/en-us/library/windows/desktop/aa364991(v=vs.85).aspx)
'
'Returns: valid filename (with prepended path and trailing ".tmp") if successful; null-string if unsuccessful
Public Function UniqueTempFilename(Optional ByRef customPrefix As String = "PD_") As String
    
    Dim tmpFolder As String
    
    'Use the current program-level temp folder as the destination for this file.  (And because we're thorough,
    ' use the system temp folder as a failsafe.)
    'If UserPrefs.IsReady Then
    '    tmpFolder = UserPrefs.GetTempPath()
    'Else
        
        tmpFolder = String$(512, 0)
        
        Dim nRet As Long
        nRet = GetTempPathW(512, StrPtr(tmpFolder))
        If (nRet > 0 And nRet < 262) Then tmpFolder = Strings.TrimNull(tmpFolder)
    
    'End If
    
    If (LenB(tmpFolder) <> 0) Then
    
        'Ensure a trailing backslash on the destination folder
        tmpFolder = Files.PathAddBackslash(tmpFolder)
        
        'Construct a temporary filename using a random GUID, and keep iterating GUIDs until we find a unique filename.
        ' (This should always succeed on the first try, but better safe than sorry!)
        Dim finalName As String
        Do
            finalName = tmpFolder & customPrefix & Left$(OS.GetArbitraryGUID(True), 8) & ".tmp"
        Loop While Files.FileExists(finalName)
        
        'We are now guaranteed a file that does not yet exist!  Return it.
        UniqueTempFilename = finalName
        
    Else
        InternalError "UniqueTempFilename failed to retrieve a valid temp folder"
    End If

End Function

'Internal system-related errors are passed here.  PD writes these to a debug log, but only in debug builds;
' you can choose to handle errors differently.
Private Sub InternalError(ByVal errComment As String, Optional ByVal errNumber As Long = 0)
    If (errNumber <> 0) Then
        Debug.Print "WARNING!  VB error in OS module (#" & Err.Number & "): " & Err.Description & " || " & errComment
    Else
        Debug.Print "WARNING!  OS module internal error: " & errComment
    End If
End Sub

