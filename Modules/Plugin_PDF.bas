Attribute VB_Name = "Plugin_PDF"
'***************************************************************************
'Adobe PDF Interface
'Copyright 2024-2024 by Tanner Helland
'Created: 23/February/24
'Last updated: 23/February/24
'Last update: start work on initial build
'
'PhotoDemon uses the pdfium library (https://pdfium.googlesource.com/pdfium/) for all PDF features.
' pdfium is provided under BSD-3 and Apache 2.0 licenses (https://pdfium.googlesource.com/pdfium/+/main/LICENSE).
'
'Support for this format was added during the PhotoDemon 10.0 release cycle.
'
'This module primarily deals with initializing and low-level interfacing with pdfium.  For higher-level
' implementation details, please refer to the pdPDF class (which is designed to hold a single PDF instance).
'
'This wrapper class also uses a shorthand implementation of DispCallFunc originally written by Olaf Schmidt.
' Many thanks to Olaf, whose original version can be found here (link good as of Feb 2019):
' http://www.vbforums.com/showthread.php?781595-VB6-Call-Functions-By-Pointer-(Universall-DLL-Calls)&p=4795471&viewfull=1#post4795471
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Most functions in this header are taken from the current pdfium fpdfview.h header, available here:
' https://pdfium.googlesource.com/pdfium/+/main/public/fpdfview.h

'Process-wide options for initializing the library.
Private Type FPDF_LIBRARY_CONFIG
  
    'Version number of the interface. Currently must be 2.
    ' Support for version 1 will be deprecated in the future.
    fpdf_Version As Long
    
    'Array of paths to scan in place of the defaults when using built-in
    ' FXGE font loading code. The array is terminated by a NULL pointer.
    ' The Array may be NULL itself to use the default paths. May be ignored
    ' entirely depending upon the platform.
    fdpf_ptrToUserFontPaths As Long
    
    'Version 2.
    ' Pointer to the v8::Isolate to use, or NULL to force PDFium to create one.
    fpdf_ptrIsolate As Long
    
    'The embedder data slot to use in the v8::Isolate to store PDFium's
    ' per-isolate data. The value needs to be in the range
    ' [0, |v8::Internals::kNumIsolateDataLots|). Note that 0 is fine for most
    ' embedders.
    fdpf_v8EmbedderSlot As Long
    
    'Version 3 - Experimental.
    ' Pointer to the V8::Platform to use.
    fdpf_ptrPlatform As Long
    
    'Version 4 - Experimental.
    ' Explicit specification of core renderer to use. |m_RendererType| must be
    ' a valid value for |FPDF_LIBRARY_CONFIG| versions of this level or higher,
    ' or else the initialization will fail with an immediate crash.
    ' Note that use of a specified |FPDF_RENDERER_TYPE| value for which the
    ' corresponding render library is not included in the build will similarly
    ' fail with an immediate crash.
    fdpf_RendererType As Long
    
End Type

'APIs listed here for convenience; due to cdecl calling convention, they must be wrapped in a helper function in VB6
'FPDF_EXPORT void FPDF_CALLCONV FPDF_InitLibraryWithConfig(const FPDF_LIBRARY_CONFIG* config);
'FPDF_EXPORT void FPDF_CALLCONV FPDF_DestroyLibrary();

'This library has very specific compiler needs in order to produce maximum perf code, so rather than
' recompile it, I've just grabbed the prebuilt Windows binaries and wrapped 'em using DispCallFunc
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef retVAR As Variant) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

'At load-time, we cache a number of proc addresses (required for passing through DispCallFunc).
' This saves us a little time vs calling GetProcAddress on each call.
Private Enum pdfium_ProcAddress
    FPDF_InitLibraryWithConfig
    FPDF_DestroyLibrary
    [last_address]
End Enum

Private m_ProcAddresses() As Long

'Rather than allocate new memory on each DispCallFunc invoke, just reuse a set of temp arrays declared
' to a maximum relevant size (see InitializeEngine, below).
Private Const MAX_PARAM_COUNT As Long = 8
Private m_vType() As Integer, m_vPtr() As Long

Private m_LibHandle As Long, m_LibAvailable As Boolean
Private m_LibFullPath As String, m_LibVersion As String

Public Sub ForciblySetAvailability(ByVal newState As Boolean)
    m_LibAvailable = newState
End Sub

Public Function GetVersion() As String
    
    'Version string is cached on first access
    If (LenB(m_LibVersion) <> 0) Then
        GetVersion = m_LibVersion
    Else
        
        Const FUNC_NAME As String = "GetVersion"
        
        Dim cFSO As pdFSO
        Set cFSO = New pdFSO
        If (Not cFSO.FileGetVersionAsString(m_LibFullPath, m_LibVersion)) Then
            InternalError FUNC_NAME, "couldn't retrieve version"
            m_LibVersion = "unknown"
        End If
        
        GetVersion = m_LibVersion
        
    End If
        
End Function

Public Function InitializeEngine(ByRef pathToDLLFolder As String) As Boolean
    
    m_LibFullPath = pathToDLLFolder & "pdfium.dll"
    m_LibHandle = VBHacks.LoadLib(m_LibFullPath)
    m_LibAvailable = (m_LibHandle <> 0)
    InitializeEngine = m_LibAvailable
    
    If InitializeEngine Then
        
        'Pre-load all relevant proc addresses
        ReDim m_ProcAddresses(0 To [last_address] - 1) As Long
        m_ProcAddresses(FPDF_InitLibraryWithConfig) = GetProcAddress(m_LibHandle, "FPDF_InitLibraryWithConfig")
        m_ProcAddresses(FPDF_DestroyLibrary) = GetProcAddress(m_LibHandle, "FPDF_DestroyLibrary")
        
        'Initialize all module-level arrays
        ReDim m_vType(0 To MAX_PARAM_COUNT - 1) As Integer
        ReDim m_vPtr(0 To MAX_PARAM_COUNT - 1) As Long
        
        'With the DLL loaded, attempt to initialize pdfium
        Dim pdfInit As FPDF_LIBRARY_CONFIG
        With pdfInit
            .fpdf_Version = 2
        End With
        
        'From the header:
        ' "You have to call this function before you can call any PDF processing functions."
        CallCDeclW FPDF_InitLibraryWithConfig, vbEmpty, VarPtr(pdfInit)
        
    End If
    
    'FPDF_InitLibraryWithConfig VarPtr(pdfInit)
    'FPDF_DestroyLibrary
    
    If (Not InitializeEngine) Then
        PDDebug.LogAction "WARNING!  LoadLibraryW failed to load pdfium.  Last DLL error: " & Err.LastDllError
    End If
    
End Function

Public Function IsPDFiumAvailable() As Boolean
    IsPDFiumAvailable = m_LibAvailable
End Function

Public Sub ReleaseEngine()
    
    'From the header:
    ' "After this function is called, you must not call any PDF processing functions.
    '  Calling this function does not automatically close other objects.
    '  It is recommended to close other objects before closing the library with this function."
    CallCDeclW FPDF_DestroyLibrary, vbEmpty
    
    'Free the library handle
    If (m_LibHandle <> 0) Then
        VBHacks.FreeLib m_LibHandle
        m_LibHandle = 0
    End If
    
End Sub

'DispCallFunc wrapper originally by Olaf Schmidt, with a few minor modifications; see the top of this class
' for a link to his original, unmodified version
Private Function CallCDeclW(ByVal lProc As pdfium_ProcAddress, ByVal fRetType As VbVarType, ParamArray pa() As Variant) As Variant

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

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String)
    If UserPrefs.GenerateDebugLogs Then
        PDDebug.LogAction "Plugin_PDF." & funcName & "() reported an error: " & errDescription
    Else
        Debug.Print "Plugin_PDF." & funcName & "() reported an error: " & errDescription
    End If
End Sub
