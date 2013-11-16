Attribute VB_Name = "OS_Interactions"
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

Private Declare Function EnumProcesses Lib "psapi" (lpidProcess As Long, ByVal cb As Long, cbNeeded As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi" (ByVal hProcess As Long, lphModule As Long, ByVal cb As Long, lpcbNeeded As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function GetProcessMemoryInfo Lib "psapi" (ByVal hProcess As Long, ppsmemCounters As PROCESS_MEMORY_COUNTERS, ByVal cb As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal Handle As Long) As Long
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

    Dim tOSVI As OSVERSIONINFO
    tOSVI.dwOSVersionInfoSize = Len(tOSVI)
    GetVersionEx tOSVI
    
    getVistaOrLaterStatus = (tOSVI.dwMajorVersion >= 6)

End Function

'Check for a version >= Win 7
Public Function getWin7OrLaterStatus() As Boolean

    Dim tOSVI As OSVERSIONINFO
    tOSVI.dwOSVersionInfoSize = Len(tOSVI)
    GetVersionEx tOSVI
    
    getWin7OrLaterStatus = ((tOSVI.dwMajorVersion >= 6) And (tOSVI.dwMinorVersion >= 1))

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
