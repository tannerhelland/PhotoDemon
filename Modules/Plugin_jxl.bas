Attribute VB_Name = "Plugin_jxl"
'***************************************************************************
'JPEG-XL Reference Library (libjxl) Interface
'Copyright 2022-2022 by Tanner Helland
'Created: 28/September/22
'Last updated: 28/September/22
'Last update: initial build
'
'libjxl (available at https://github.com/libjxl/libjxl) is the official reference library implementation
' for the modern JPEG-XL format.  Support for this format was added during the PhotoDemon 10.0 release cycle.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Library handle will be non-zero if libjxl is available; you can also forcibly override the
' "availability" state by setting m_LibAvailable to FALSE
Private m_LibHandle As Long, m_LibAvailable As Boolean

'libjxl has very specific compiler needs in order to produce maximum perf code, so rather than
' compile myself, I stick with the prebuilt Windows binaries and wrap 'em using DispCallFunc
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef retVAR As Variant) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

'At load-time, we cache a number of proc addresses (required for passing through DispCallFunc).
' This saves us a little time vs calling GetProcAddress on each call.
Private Enum LibJXL_ProcAddress
    JxlDecoderVersion
    [last_address]
End Enum

Private m_ProcAddresses() As Long

'Rather than allocate new memory on each DispCallFunc invoke, just reuse a set of temp arrays declared
' to the maximum relevant size (see InitializeEngine, below).
Private Const MAX_PARAM_COUNT As Long = 8
Private m_vType() As Integer, m_vPtr() As Long

'Initialize the library.  Do not call this until you have verified its existence (typically via the PluginManager module)
Public Function InitializeLibJXL(ByRef pathToDLLFolder As String) As Boolean

    'Manually load the DLL from the plugin folder (should be App.Path\Data\Plugins)
    Dim libPath As String
    libPath = pathToDLLFolder & "libjxl.dll"
    m_LibHandle = VBHacks.LoadLib(libPath)
    InitializeLibJXL = (m_LibHandle <> 0)
    m_LibAvailable = InitializeLibJXL
    
    'Initialize all module-level arrays
    ReDim m_ProcAddresses(0 To [last_address] - 1) As Long
    ReDim m_vType(0 To MAX_PARAM_COUNT - 1) As Integer
    ReDim m_vPtr(0 To MAX_PARAM_COUNT - 1) As Long
    
    'If we initialized the library successfully, cache some library-specific data
    If InitializeLibJXL Then
        
        'Pre-load all relevant proc addresses
        m_ProcAddresses(JxlDecoderVersion) = GetProcAddress(m_LibHandle, "JxlDecoderVersion")
        
    Else
        PDDebug.LogAction "WARNING!  LoadLibrary failed to load libjxl.  Last DLL error: " & Err.LastDllError
        PDDebug.LogAction "(FYI, the attempted path was: " & libPath & ")"
    End If
    
End Function

'Forcibly disable library interactions at run-time (if newState is FALSE).
' Setting newState to TRUE is not advised; this module will handle state internally based
' on successful library loading.
Public Sub ForciblySetAvailability(ByVal newState As Boolean)
    m_LibAvailable = newState
End Sub

Public Function GetLibJXLVersion() As String
    
    Dim ptrVersion As Long
    ptrVersion = CallCDeclW(JxlDecoderVersion, vbLong)
    
    'From the docs (https://libjxl.readthedocs.io/en/latest/api_decoder.html):
    ' Returns the decoder library version as an integer:
    ' MAJOR_VERSION * 1000000 + MINOR_VERSION * 1000 + PATCH_VERSION.
    ' (For example, version 1.2.3 would return 1002003.)
    GetLibJXLVersion = Trim$(Str$(ptrVersion \ 1000000)) & "." & Trim$(Str$((ptrVersion \ 1000) Mod 1000)) & "." & Trim$(Str$(ptrVersion Mod 1000)) & ".0"
    
End Function

Public Function IsLibJXLAvailable() As Boolean
    IsLibJXLAvailable = (m_LibHandle <> 0)
End Function

Public Function IsLibJXLEnabled() As Boolean
    IsLibJXLEnabled = m_LibAvailable
End Function

'When PD closes, make sure to release our open library handle
Public Sub ReleaseLibJXL()
    If (m_LibHandle <> 0) Then
        VBHacks.FreeLib m_LibHandle
        m_LibHandle = 0
    End If
End Sub

'DispCallFunc wrapper originally by Olaf Schmidt, with a few minor modifications; see the top of this class
' for a link to his original, unmodified version
Private Function CallCDeclW(ByVal lProc As LibJXL_ProcAddress, ByVal fRetType As VbVarType, ParamArray pa() As Variant) As Variant

    Dim i As Long, vTemp() As Variant, hResult As Long
    
    Dim numParams As Long
    If (UBound(pa) < LBound(pa)) Then numParams = 0 Else numParams = UBound(pa) + 1
    
    If IsMissing(pa) Then
        ReDim vTemp(0) As Variant
    Else
        vTemp = pa 'make a copy of the params to prevent problems with VT_ByRef members in the ParamArray
    End If
    
    For i = 0 To numParams - 1
        m_vType(i) = VarType(vTemp(i))
        m_vPtr(i) = VarPtr(vTemp(i))
    Next i
    
    Const CC_CDECL As Long = 1
    hResult = DispCallFunc(0, m_ProcAddresses(lProc), CC_CDECL, fRetType, i, m_vType(0), m_vPtr(0), CallCDeclW)
    
End Function

Private Sub InternalError(ByVal errString As String, Optional ByVal faultyReturnCode As Long = 256)
    If (faultyReturnCode <> 256) Then
        PDDebug.LogAction "libjxl returned an error code: " & faultyReturnCode, PDM_External_Lib
    Else
        PDDebug.LogAction "libjxl experienced an error; additional explanation may be: " & errString, PDM_External_Lib
    End If
End Sub
