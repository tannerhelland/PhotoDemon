Attribute VB_Name = "Misc_OS_Interactions"
'***************************************************************************
'Miscellaneous OS Interaction Handler
'Copyright ©2011-2013 by Tanner Helland
'Created: 27/November/12
'Last updated: 27/November/12
'Last update: added function for returning PhotoDemon's current memory usage
'
'Sometimes, PhotoDemon needs to query Windows for OS-specific data - such as the current version of Windows, or the
' available RAM on the system.  This module handles such calls.
'
'Special thanks goes to Mike Raynder, who wrote the original version of the process-specific memory function.  You can
' download a copy of Mike's original code from this link (good as of 27 Nov '12): http://www.xtremevbtalk.com/showthread.php?t=229758
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

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long

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

Private Type LARGE_INTEGER
     lowpart As Long
     highpart As Long
End Type

Private Type MEMORYSTATUSEX
  dwLength As Long
  dwMemoryLoad As Long
  ullTotalPhys As LARGE_INTEGER
  ullAvailPhys As LARGE_INTEGER
  ullTotalPageFile As LARGE_INTEGER
  ullAvailPageFile As LARGE_INTEGER
  ullTotalVirtual As LARGE_INTEGER
  ullAvailVirtual As LARGE_INTEGER
  ullAvailExtendedVirtual As LARGE_INTEGER
End Type

Private Declare Function EnumProcesses Lib "PSAPI.DLL" (lpidProcess As Long, ByVal cb As Long, cbNeeded As Long) As Long
Private Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, lphModule As Long, ByVal cb As Long, lpcbNeeded As Long) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "PSAPI.DLL" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function GetProcessMemoryInfo Lib "PSAPI.DLL" (ByVal hProcess As Long, ppsmemCounters As PROCESS_MEMORY_COUNTERS, ByVal cb As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal Handle As Long) As Long

'Check the current Windows version.  In PhotoDemon, we are only concerned with "is it Vista or later?"
Public Function getVistaOrLaterStatus() As Boolean

    Dim tOSVI As OSVERSIONINFO
    tOSVI.dwOSVersionInfoSize = Len(tOSVI)
    GetVersionEx tOSVI
    
    getVistaOrLaterStatus = (tOSVI.dwMajorVersion >= 6)

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
