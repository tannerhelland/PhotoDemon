Attribute VB_Name = "Plugin_PDF"
'***************************************************************************
'Adobe PDF Interface (via pdfium)
'Copyright 2024-2026 by Tanner Helland
'Created: 23/February/24
'Last updated: 25/March/25
'Last update: delay-load the library to improve performance
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

'Rendering options when rendering a PDF page.  Can be combined via OR.
Public Enum PDFium_RenderOptions
    FPDF_ANNOT = &H1&           'Set if annotations are to be rendered.
    FPDF_LCD_TEXT = &H2&        'Set if using text rendering optimized for LCD display. This flag will only take effect if anti-aliasing is enabled for text.
    FPDF_NO_NATIVETEXT = &H4&   'Don't use the native text output available on some platforms
    FPDF_GRAYSCALE = &H8&       'Grayscale output.
    FPDF_DEBUG_INFO = &H80&     'Obsolete, has no effect, retained for compatibility.
    FPDF_NO_CATCH = &H100&      'Obsolete, has no effect, retained for compatibility.
    FPDF_RENDER_LIMITEDIMAGECACHE = &H200&  'Limit image cache size.
    FPDF_RENDER_FORCEHALFTONE = &H400&      'Always use halftone for image stretching.
    FPDF_PRINTING = &H800&      'Render for printing.
    FPDF_RENDER_NO_SMOOTHTEXT = &H1000&     'Set to disable anti-aliasing on text. This flag will also disable LCD optimization for text rendering.
    FPDF_RENDER_NO_SMOOTHIMAGE = &H2000&    'Set to disable anti-aliasing on images.
    FPDF_RENDER_NO_SMOOTHPATH = &H4000&     'Set to disable anti-aliasing on paths.
    FPDF_REVERSE_BYTE_ORDER = &H10&     'Set whether to render in a reverse Byte order, this flag is only used when rendering to a bitmap.
    FPDF_CONVERT_FILL_TO_STROKE = &H20& 'Set whether fill paths need to be stroked. This flag is only used when FPDF_COLORSCHEME is passed in, since with a single fill color for paths the boundaries of adjacent fill paths are less visible.
End Enum

#If False Then
    Private Const FPDF_ANNOT = &H1&, FPDF_LCD_TEXT = &H2&, FPDF_NO_NATIVETEXT = &H4&, FPDF_GRAYSCALE = &H8&, FPDF_DEBUG_INFO = &H80&, FPDF_NO_CATCH = &H100&, FPDF_RENDER_LIMITEDIMAGECACHE = &H200&, FPDF_RENDER_FORCEHALFTONE = &H400&, FPDF_PRINTING = &H800&, FPDF_RENDER_NO_SMOOTHTEXT = &H1000&, FPDF_RENDER_NO_SMOOTHIMAGE = &H2000&, FPDF_RENDER_NO_SMOOTHPATH = &H4000&, FPDF_REVERSE_BYTE_ORDER = &H10&, FPDF_CONVERT_FILL_TO_STROKE = &H20&
#End If

Public Enum PDFium_Orientation
    FPDF_Normal = 0     '(normal)
    FPDF_Rotate90 = 1   '(rotated 90 degrees clockwise)
    FPDF_Rotate180 = 2  '(rotated 180 degrees)
    FPDF_Rotate270 = 3  '(rotated 90 degrees counter-clockwise)
End Enum

#If False Then
    Private Const FPDF_Normal = 0, FPDF_Rotate90 = 1, FPDF_Rotate180 = 2, FPDF_Rotate270 = 3
#End If

Public Enum PDFium_Boundary
    FSPDF_PAGEBOX_MEDIABOX = 0  'The boundary of the physical medium on which page is to be displayed or printed.
    FSPDF_PAGEBOX_CROPBOX = 1   'The region to which the contents of page are to be clipped (cropped) while displaying or printing.
    FSPDF_PAGEBOX_TRIMBOX = 2   'The region to which the contents of page should be clipped while outputting in a production environment.
    FSPDF_PAGEBOX_ARTBOX = 3    'The intended dimensions of a finished page after trimming.
    FSPDF_PAGEBOX_BLEEDBOX = 4  'The extent of page's meaningful content (including potential white space) as intended by page's creator.
End Enum

#If False Then
    Private Const FSPDF_PAGEBOX_MEDIABOX = 0, FSPDF_PAGEBOX_CROPBOX = 1, FSPDF_PAGEBOX_TRIMBOX = 2, FSPDF_PAGEBOX_ARTBOX = 3, FSPDF_PAGEBOX_BLEEDBOX = 4
#End If

'This library has very specific compiler needs in order to produce maximum perf code, so rather than
' recompile it, I've just grabbed the prebuilt Windows binaries and wrapped 'em using DispCallFunc
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef retVAR As Variant) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

'At load-time, we cache a number of proc addresses (required for passing through DispCallFunc).
' This saves us a little time vs calling GetProcAddress on each call.
Public Enum PDFium_ProcAddress
    FPDF_InitLibraryWithConfig
    FPDF_DestroyLibrary
    FPDF_GetLastError
    FPDF_LoadDocument
    FPDF_CloseDocument
    FPDF_GetPageCount
    
    FPDF_LoadPage
    FPDF_GetPageWidthF
    FPDF_GetPageWidth
    FPDF_GetPageHeightF
    FPDF_GetPageHeight
    FPDF_GetPageBoundingBox
    FPDF_GetPageSizeByIndexF
    FPDF_GetPageSizeByIndex
    FPDF_RenderPage
    FPDF_RenderPageBitmap
    FPDF_RenderPageBitmapWithMatrix
    FPDF_ClosePage
    FPDF_DeviceToPage
    FPDF_PageToDevice
    
    FPDFBitmap_CreateEx
    FPDFBitmap_Destroy
    
    FPDFPage_GetRotation
    
    [last_address]
End Enum

#If False Then
    Private Const FPDF_InitLibraryWithConfig = 0, FPDF_DestroyLibrary = 0, FPDF_GetLastError = 0, FPDF_LoadDocument = 0, FPDF_CloseDocument = 0
    Private Const FPDF_GetPageCount = 0, FPDF_LoadPage = 0, FPDF_GetPageWidthF = 0, FPDF_GetPageWidth = 0, FPDF_GetPageHeightF = 0
    Private Const FPDF_GetPageHeight = 0, FPDF_GetPageBoundingBox = 0, FPDF_GetPageSizeByIndexF = 0, FPDF_GetPageSizeByIndex = 0
    Private Const FPDF_RenderPage = 0, FPDF_RenderPageBitmap = 0, FPDF_RenderPageBitmapWithMatrix = 0, FPDF_ClosePage = 0
    Private Const FPDF_DeviceToPage = 0, FPDF_PageToDevice = 0, FPDFBitmap_CreateEx = 0, FPDFPage_GetRotation = 0
#End If

'Child classes need to retrieve this proc list in order to efficiently interface with pdfium
Private m_ProcAddresses() As PDFium_ProcAddress

'Rather than allocate new memory on each DispCallFunc invoke, just reuse a set of temp arrays declared
' to a maximum relevant size (see InitializeEngine, below).
Private Const MAX_PARAM_COUNT As Long = 8
Private m_vType() As Integer, m_vPtr() As Long

Private m_LibHandle As Long, m_LibAvailable As Boolean
Private m_LibFullPath As String, m_LibVersion As String

Public Sub CopyPDFiumProcAddresses(ByRef dstList() As PDFium_ProcAddress)
    
    'Ensure library is available before proceeding
    If (m_LibHandle = 0) Then Plugin_PDF.InitializeEngine True
    If (Not Plugin_PDF.IsPDFiumAvailable()) Then Exit Sub
    
    ReDim dstList(0 To [last_address] - 1) As PDFium_ProcAddress
    
    Dim i As PDFium_ProcAddress
    For i = 0 To [last_address] - 1
        dstList(i) = m_ProcAddresses(i)
    Next i
    
End Sub

Public Sub ForciblySetAvailability(ByVal newState As Boolean)
    m_LibAvailable = newState
End Sub

'Version is retrieved from file properties, *not* an actual PDF call (so we don't need to initialize first)
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

Public Function InitializeEngine(Optional ByVal actuallyLoadDLL As Boolean = True) As Boolean
    
    'I don't currently know how to build pdfium in an XP-compatible way.
    ' As a result, its support is limited to Win Vista and above.
    If (Not OS.IsVistaOrLater) Then
        InitializeEngine = False
        PDDebug.LogAction "pdfium does not currently work on Windows XP"
        Exit Function
    End If
        
    m_LibFullPath = PluginManager.GetPluginPath() & "pdfium.dll"
    
    If (Not actuallyLoadDLL) Then
        InitializeEngine = Files.FileExists(m_LibFullPath)
        Exit Function
    End If
    
    m_LibHandle = VBHacks.LoadLib(m_LibFullPath)
    m_LibAvailable = (m_LibHandle <> 0)
    InitializeEngine = m_LibAvailable
    
    If InitializeEngine Then
        
        'Pre-load all relevant proc addresses
        ReDim m_ProcAddresses(0 To [last_address] - 1) As PDFium_ProcAddress
        m_ProcAddresses(FPDF_InitLibraryWithConfig) = GetProcAddress(m_LibHandle, "FPDF_InitLibraryWithConfig")
        m_ProcAddresses(FPDF_DestroyLibrary) = GetProcAddress(m_LibHandle, "FPDF_DestroyLibrary")
        m_ProcAddresses(FPDF_GetLastError) = GetProcAddress(m_LibHandle, "FPDF_GetLastError")
        m_ProcAddresses(FPDF_LoadDocument) = GetProcAddress(m_LibHandle, "FPDF_LoadDocument")
        m_ProcAddresses(FPDF_CloseDocument) = GetProcAddress(m_LibHandle, "FPDF_CloseDocument")
        m_ProcAddresses(FPDF_GetPageCount) = GetProcAddress(m_LibHandle, "FPDF_GetPageCount")
        m_ProcAddresses(FPDF_LoadPage) = GetProcAddress(m_LibHandle, "FPDF_LoadPage")
        m_ProcAddresses(FPDF_GetPageWidthF) = GetProcAddress(m_LibHandle, "FPDF_GetPageWidthF")
        m_ProcAddresses(FPDF_GetPageWidth) = GetProcAddress(m_LibHandle, "FPDF_GetPageWidth")
        m_ProcAddresses(FPDF_GetPageHeightF) = GetProcAddress(m_LibHandle, "FPDF_GetPageHeightF")
        m_ProcAddresses(FPDF_GetPageHeight) = GetProcAddress(m_LibHandle, "FPDF_GetPageHeight")
        m_ProcAddresses(FPDF_GetPageBoundingBox) = GetProcAddress(m_LibHandle, "FPDF_GetPageBoundingBox")
        m_ProcAddresses(FPDF_GetPageSizeByIndexF) = GetProcAddress(m_LibHandle, "FPDF_GetPageSizeByIndexF")
        m_ProcAddresses(FPDF_GetPageSizeByIndex) = GetProcAddress(m_LibHandle, "FPDF_GetPageSizeByIndex")
        m_ProcAddresses(FPDF_RenderPage) = GetProcAddress(m_LibHandle, "FPDF_RenderPage")
        m_ProcAddresses(FPDF_RenderPageBitmap) = GetProcAddress(m_LibHandle, "FPDF_RenderPageBitmap")
        m_ProcAddresses(FPDF_RenderPageBitmapWithMatrix) = GetProcAddress(m_LibHandle, "FPDF_RenderPageBitmapWithMatrix")
        m_ProcAddresses(FPDF_ClosePage) = GetProcAddress(m_LibHandle, "FPDF_ClosePage")
        m_ProcAddresses(FPDF_DeviceToPage) = GetProcAddress(m_LibHandle, "FPDF_DeviceToPage")
        m_ProcAddresses(FPDF_PageToDevice) = GetProcAddress(m_LibHandle, "FPDF_PageToDevice")
        m_ProcAddresses(FPDFBitmap_CreateEx) = GetProcAddress(m_LibHandle, "FPDFBitmap_CreateEx")
        m_ProcAddresses(FPDFBitmap_Destroy) = GetProcAddress(m_LibHandle, "FPDFBitmap_Destroy")
        m_ProcAddresses(FPDFPage_GetRotation) = GetProcAddress(m_LibHandle, "FPDFPage_GetRotation")
        
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
    
    If (Not InitializeEngine) Then
        PDDebug.LogAction "WARNING!  LoadLibraryW failed to load pdfium.  Last DLL error: " & Err.LastDllError
    End If
    
End Function

'Test if an arbitrary file is a valid PDF.  This should catch nearly all valid PDF files, regardless of extension.
Public Function IsFileLikelyPDF(ByRef srcFile As String) As Boolean
    
    IsFileLikelyPDF = False
    
    'If the required 3rd-party library isn't available, bail (because we can't handle the file anyway)
    If (Not Plugin_PDF.IsPDFiumAvailable()) Then Exit Function
    
    If Files.FileExists(srcFile) Then
        
        Dim cStream As pdStream
        Set cStream = New pdStream
        If cStream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFile, 1024, optimizeAccess:=OptimizeSequentialAccess) Then
            
            'Pull the first 1024 bytes into a local string
            Dim testHeader As String
            testHeader = cStream.ReadString_ASCII(1024)
            
            'Look for the magic number "%PDF" somewhere in those bytes
            IsFileLikelyPDF = Strings.StrStrI(StrPtr(testHeader), StrPtr("%PDF"))
            
            cStream.StopStream True
            
        End If
        
    End If
    
End Function

Public Function IsPDFiumAvailable() As Boolean
    IsPDFiumAvailable = m_LibAvailable
End Function

Public Sub ReleaseEngine()
    
    If (m_LibHandle <> 0) Then
    
        'From the header:
        ' "After this function is called, you must not call any PDF processing functions.
        '  Calling this function does not automatically close other objects.
        '  It is recommended to close other objects before closing the library with this function."
        CallCDeclW FPDF_DestroyLibrary, vbEmpty
    
        'Free the library handle
        VBHacks.FreeLib m_LibHandle
        
    End If
    
    m_LibHandle = 0
    
End Sub

'DispCallFunc wrapper originally by Olaf Schmidt, with a few minor modifications; see the top of this class
' for a link to his original, unmodified version
Private Function CallCDeclW(ByVal lProc As PDFium_ProcAddress, ByVal fRetType As VbVarType, ParamArray pa() As Variant) As Variant

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
