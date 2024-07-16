Attribute VB_Name = "Plugin_Heif"
'***************************************************************************
'libheif Library Interface
'Copyright 2024-2024 by Tanner Helland
'Created: 16/July/24
'Last updated: 16/July/24
'Last update: initial build
'
'Per its documentation (available at https://github.com/strukturag/libheif), libheif is...
'
'"...an ISO/IEC 23008-12:2017 HEIF and AVIF (AV1 Image File Format) file format
' decoder and encoder... HEIF and AVIF are new image file formats employing HEVC (H.265)
' or AV1 image coding, respectively, for the best compression ratios currently possible."
'
'libheif is LGPL-licensed and actively maintained.  PhotoDemon does not use its potential
' AVIF support due to x86 compatibility issues (AVIF support is 64-bit focused and x86 builds
' are not currently feasible for me to self-maintain, so I only compile with HEIF enabled).
'
'Note that all features in this module rely on the libheif binaries that ship with PhotoDemon.
' These features will not work if libheif cannot be located.  Per standard LGPL terms, you can
' supply your own libheif copies in place of PD's default ones, but libheif and all supporting
' libraries obviously need to be built as x86 libraries for this to work (not x64).
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'libheif is built using vcpkg, which limits my control over calling convention.  (I don't want to
' manually build the libraries via CMake - they're complex!)  DispCallFunc is used to work around
' VB6 stdcall limitations.
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef retVAR As Variant) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

'At load-time, we cache a number of proc addresses (required for passing through DispCallFunc).
' This saves us a little time vs calling GetProcAddress on each call.
Private Enum Libheif_ProcAddress
    heif_get_version_number
    [last_address]
End Enum

Private m_ProcAddresses() As Long

'Rather than allocate new memory on each DispCallFunc invoke, just reuse a set of temp arrays declared
' to the maximum relevant size (see InitializeEngine, below).
Private Const MAX_PARAM_COUNT As Long = 8
Private m_vType() As Integer, m_vPtr() As Long

'Library handle will be non-zero if each library is available; you can also forcibly override the
' "availability" state by setting m_LibAvailable to FALSE
Private m_hLibHeif As Long, m_hLibde265 As Long, m_hLibx265 As Long, m_LibAvailable As Boolean

'Forcibly disable plugin interactions at run-time (if newState is FALSE).
' Setting newState to TRUE is not advised; this module will handle state internally based
' on successful library loading.
Public Sub ForciblySetAvailability(ByVal newState As Boolean)
    m_LibAvailable = newState
End Sub

Public Function GetHandle_LibHeif() As Long
    GetHandle_LibHeif = m_hLibHeif
End Function

Public Function GetHandle_Libde265() As Long
    GetHandle_Libde265 = m_hLibde265
End Function

Public Function GetHandle_Libx265() As Long
    GetHandle_Libx265 = m_hLibx265
End Function

Public Function GetVersion() As String
    
    If (m_hLibHeif = 0) Or (Not m_LibAvailable) Then Exit Function
        
    'Byte version numbers get packed into a long
    Dim versionAsInt(0 To 3) As Byte
    
    Dim tmpLong As Long
    tmpLong = CallCDeclW(heif_get_version_number, vbLong)
    PutMem4 VarPtr(versionAsInt(0)), tmpLong
    
    'Want to ensure we retrieved the correct values?  Use this:
    GetVersion = versionAsInt(3) & "." & versionAsInt(2) & "." & versionAsInt(1) & "." & versionAsInt(0)
    
End Function

Public Function InitializeEngine(ByRef pathToDLLFolder As String) As Boolean
    
    'Initialize all required libraries
    Dim strLibPath As String
    
    'Load dependencies first
    strLibPath = pathToDLLFolder & "libde265.dll"
    m_hLibde265 = VBHacks.LoadLib(strLibPath)
    strLibPath = pathToDLLFolder & "libx265.dll"
    m_hLibx265 = VBHacks.LoadLib(strLibPath)
    
    'The main library can now resolve dependencies correctly...
    strLibPath = pathToDLLFolder & "libheif.dll"
    m_hLibHeif = VBHacks.LoadLib(strLibPath)
    m_LibAvailable = (m_hLibHeif <> 0) And (m_hLibde265 <> 0) And (m_hLibx265 <> 0)
    InitializeEngine = m_LibAvailable
    
    'If we initialized the library successfully, preload proc addresses
    If InitializeEngine Then
    
        ReDim m_ProcAddresses(0 To [last_address] - 1) As Long
        m_ProcAddresses(heif_get_version_number) = GetProcAddress(m_hLibHeif, "heif_get_version_number")
        
        'Initialize all module-level arrays
        ReDim m_vType(0 To MAX_PARAM_COUNT - 1) As Integer
        ReDim m_vPtr(0 To MAX_PARAM_COUNT - 1) As Long
        
    Else
        PDDebug.LogAction "WARNING!  LoadLibrary failed to load libheif.  Last DLL error: " & Err.LastDllError
        PDDebug.LogAction "(FYI, the attempted path was: " & strLibPath & ")"
    End If
    
End Function

Public Function IsLibheifEnabled() As Boolean
    IsLibheifEnabled = m_LibAvailable
End Function

Public Sub ReleaseEngine()
    
    'For extra safety, free in reverse order from loading
    If (m_hLibHeif <> 0) Then
        VBHacks.FreeLib m_hLibHeif
        m_hLibHeif = 0
    End If
    If (m_hLibde265 <> 0) Then
        VBHacks.FreeLib m_hLibde265
        m_hLibde265 = 0
    End If
    If (m_hLibde265 <> 0) Then
        VBHacks.FreeLib m_hLibde265
        m_hLibde265 = 0
    End If
    
    m_LibAvailable = False
    
End Sub

'DispCallFunc wrapper originally by Olaf Schmidt, with a few minor modifications; see the top of this class
' for a link to his original, unmodified version
Private Function CallCDeclW(ByVal lProc As Libheif_ProcAddress, ByVal fRetType As VbVarType, ParamArray pa() As Variant) As Variant

    Dim i As Long, vTemp() As Variant, hResult As Long
    
    Dim numParams As Long
    If (UBound(pa) < LBound(pa)) Then numParams = 0 Else numParams = UBound(pa) + 1
    
    If IsMissing(pa) Then
        ReDim vTemp(0) As Variant
    Else
        vTemp = pa 'make a copy of the params, to prevent problems with VT_Byref-Members in the ParamArray
    End If
    
    For i = 0 To numParams - 1
        If VarType(pa(i)) = vbString Then vTemp(i) = StrPtr(pa(i))
        m_vType(i) = VarType(vTemp(i))
        m_vPtr(i) = VarPtr(vTemp(i))
    Next i
    
    Const CC_CDECL As Long = 1
    hResult = DispCallFunc(0, m_ProcAddresses(lProc), CC_CDECL, fRetType, i, m_vType(0), m_vPtr(0), CallCDeclW)
    
End Function
