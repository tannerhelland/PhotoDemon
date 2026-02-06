VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'Individual Adobe PDF File Interface (via pdfium)
'Copyright 2024-2026 by Tanner Helland
'Created: 23/February/24
'Last updated: 06/January/26
'Last update: add support for importing password-protected PDFs
'
'PhotoDemon uses the pdfium library (https://pdfium.googlesource.com/pdfium/) for all PDF features.
' pdfium is provided under BSD-3 and Apache 2.0 licenses (https://pdfium.googlesource.com/pdfium/+/main/LICENSE).
'
'Support for this format was added during the PhotoDemon 10.0 release cycle.
'
'This class relies on pdfium initialization and interface code written in Plugin_PDF.bas.
' You *MUST* call Plugin_PDF.InitializeEngine() with a valid path to pdfium.dll before initializing
' any instances of this class.
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

'Much like Windows itself, error state can be retrieved via a dedicated FPDF_GetLastError() function.
Private Enum FPDF_ErrorCodes
    FPDF_ERR_SUCCESS = 0   ' No error.
    FPDF_ERR_UNKNOWN = 1   ' Unknown error.
    FPDF_ERR_FILE = 2      ' File not found or could not be opened.
    FPDF_ERR_FORMAT = 3    ' File not in PDF format or corrupted.
    FPDF_ERR_PASSWORD = 4  ' Password required or incorrect password.
    FPDF_ERR_SECURITY = 5  ' Unsupported security scheme.
    FPDF_ERR_PAGE = 6      ' Page not found or content error.
    '#ifdef PDF_ENABLE_XFA
    FPDF_ERR_XFALOAD = 7   ' Load XFA error.
    FPDF_ERR_XFALAYOUT = 8 ' Layout XFA error.
    '#endif  // PDF_ENABLE_XFA
End Enum

#If False Then
    Private Const FPDF_ERR_SUCCESS = 0, FPDF_ERR_UNKNOWN = 1, FPDF_ERR_FILE = 2, FPDF_ERR_FORMAT = 3, FPDF_ERR_PASSWORD = 4, FPDF_ERR_SECURITY = 5, FPDF_ERR_PAGE = 6, FPDF_ERR_XFALOAD = 7, FPDF_ERR_XFALAYOUT = 8
#End If

Private Type PDFium_SizeF
    pdfWidth As Single
    pdfHeight As Single
End Type

'This library has very specific compiler needs in order to produce maximum perf code, so rather than
' recompile it, I've just grabbed the prebuilt Windows binaries and wrapped 'em using DispCallFunc
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef retVAR As Variant) As Long

'DLL addresses are only calculated once, by Plugin_PDF (we just borrow its list)
Private m_ProcAddresses() As PDFium_ProcAddress

'Rather than allocate new memory on each DispCallFunc invoke, just reuse a set of temp arrays declared
' to a maximum relevant size (see InitializeEngine, below).
Private Const MAX_PARAM_COUNT As Long = 8
Private m_vType() As Integer, m_vPtr() As Long

'If this class holds a live PDF handle, this will be non-zero
Private m_hPDF As Long

'To reduce load-time, after validating a file as PDF, the caller can leave the file open (because presumably
' a load operation is soon to follow).  After validation, this string will contain the full path of the
' just-validated source file.
Private m_lastValidatedFile As String

'If a page is active and loaded, this will be non-zero.  The open page (if any) needs to be manually closed
' before opening a new page or changing documents.  (Note that pdfium itself allows unlimited open handles,
' but that greatly complicates this class so it enforces serial page access.)
Private m_hPage As Long

'Close the currently open page.  This class handles this automatically, so you don't need to call this
' (unless you want to free up memory or something).
Friend Sub CloseCurrentPage()
    If (m_hPage <> 0) Then
        CallCDeclW FPDF_ClosePage, vbEmpty, m_hPage
        m_hPage = 0
    End If
End Sub

'Close any open PDF file handle
Friend Sub ClosePDF()
    
    'Close the current page, if any
    If (m_hPage <> 0) Then CloseCurrentPage
    
    'Close the current document
    If (m_hPDF <> 0) Then
        
        'From the header:
        '// Function: FPDF_CloseDocument
        '//          Close a loaded PDF document.
        '// Parameters:
        '//          document    -   Handle to the loaded document.
        '// Return value:
        '//          None.
        'FPDF_EXPORT void FPDF_CALLCONV FPDF_CloseDocument(FPDF_DOCUMENT document);
        CallCDeclW FPDF_CloseDocument, vbEmpty, m_hPDF
        
    End If
    
    'Reset any relevant status trackers to "no PDF loaded" state
    m_hPDF = 0
    m_lastValidatedFile = vbNullString
    
End Sub

'Number of pages in the PDF (1-based)
Friend Function GetPageCount() As Long
    
    Const FUNC_NAME As String = "GetPageCount"
    If (m_hPDF = 0) Then
        InternalError FUNC_NAME, "no PDF loaded"
        Exit Function
    End If
    
    'From the header:
    '// Function: FPDF_GetPageCount
    '//          Get total number of pages in the document.
    '// Parameters:
    '//          document    -   Handle to document. Returned by FPDF_LoadDocument.
    '// Return value:
    '//          Total number of pages in the document.
    'FPDF_EXPORT int FPDF_CALLCONV FPDF_GetPageCount(FPDF_DOCUMENT document);
    GetPageCount = CallCDeclW(FPDF_GetPageCount, vbLong, m_hPDF)
    
End Function

'Height of the currently open page, in points (1/72")
Friend Function GetPageHeightInPoints() As Single
    
    Const FUNC_NAME As String = "GetPageHeightInPoints"
    
    'Make sure a valid page is loaded
    If (m_hPage = 0) Then
        InternalError FUNC_NAME, "no page loaded"
        Exit Function
    End If
    
    GetPageHeightInPoints = CallCDeclW(FPDF_GetPageHeightF, vbSingle, m_hPage)
    
End Function

'Height of any arbitrary page, in points (1/72").  Does not validate page index.
Friend Function GetPageHeightInPoints_ByIndex(ByVal idxPage As Long) As Single
    Dim tmpSizeF As PDFium_SizeF
    If (CallCDeclW(FPDF_GetPageSizeByIndexF, vbLong, m_hPDF, idxPage, VarPtr(tmpSizeF)) <> 0) Then
        GetPageHeightInPoints_ByIndex = tmpSizeF.pdfHeight
    End If
End Function

'Orientation of the currently open page
Friend Function GetPageOrientation() As PDFium_Orientation
    
    Const FUNC_NAME As String = "GetPageOrientation"
    
    'Make sure a valid page is loaded
    If (m_hPage = 0) Then
        InternalError FUNC_NAME, "no page loaded"
        Exit Function
    End If
    
    GetPageOrientation = CallCDeclW(FPDFPage_GetRotation, vbLong, m_hPage)
    
End Function

'Width of the currently open page, in points (1/72")
Friend Function GetPageWidthInPoints() As Single

    Const FUNC_NAME As String = "GetPageWidthInPoints"
    
    'Make sure a valid page is loaded
    If (m_hPage = 0) Then
        InternalError FUNC_NAME, "no page loaded"
        Exit Function
    End If
    
    GetPageWidthInPoints = CallCDeclW(FPDF_GetPageWidthF, vbSingle, m_hPage)
    
End Function

'Width of any arbitrary page, in points (1/72").  Does not validate page index.
Friend Function GetPageWidthInPoints_ByIndex(ByVal idxPage As Long) As Single
    Dim tmpSizeF As PDFium_SizeF
    If (CallCDeclW(FPDF_GetPageSizeByIndexF, vbLong, m_hPDF, idxPage, VarPtr(tmpSizeF)) <> 0) Then
        GetPageWidthInPoints_ByIndex = tmpSizeF.pdfWidth
    End If
End Function

'Returns TRUE if this class holds a working PDF instance
Friend Function HasPDF() As Boolean
    HasPDF = (m_hPDF <> 0)
End Function

'Test if an arbitrary file is a valid PDF.  Note that this function will return *FALSE* if the file *is* a PDF
' but the file is also password-protected but the user doesn't supply a valid password.
Friend Function IsFilePDF(ByRef srcFile As String, Optional ByVal leaveHandleOpen As Boolean = True) As Boolean
    
    IsFilePDF = False
    
    'If the required 3rd-party library isn't available, bail
    If (Not Plugin_PDF.IsPDFiumAvailable()) Then Exit Function
    
    'Reset the validation tracker
    m_lastValidatedFile = vbNullString
    
    'Attempt to load the file.  As mentioned above, if the file requires a password but the user does not
    ' supply a valid one, this function will return FALSE.
    IsFilePDF = LoadPDFFromFile(srcFile)
    
    'If the caller doesn't want us to keep the file open for further processing, close it before exiting
    If (m_hPDF <> 0) Then
        m_lastValidatedFile = srcFile
        If (Not leaveHandleOpen) Then ClosePDF
    
    'The file failed validation; clear any relevant trackers
    Else
        m_lastValidatedFile = vbNullString
    End If
    
End Function

'If all pages in the PDF have the same page size, this will be TRUE; FALSE means pages have variable sizes.
Friend Function IsPageSizeUniform() As Boolean
    
    'Assume uniformity.  We'll exit if we encounter a page with differing sizes.
    IsPageSizeUniform = True
    
    'If there's only one page, the answer is always TRUE
    If (Me.GetPageCount() <= 1) Then Exit Function
    
    'Pull the size of the first page
    Dim firstPageSize As PDFium_SizeF, otherPageSize As PDFium_SizeF
    
    Dim apiSuccess As Long
    apiSuccess = CallCDeclW(FPDF_GetPageSizeByIndexF, vbLong, m_hPDF, 0&, VarPtr(firstPageSize))
    
    If (apiSuccess <> 0) Then
        
        Dim i As Long
        For i = 1 To Me.GetPageCount() - 1
            apiSuccess = CallCDeclW(FPDF_GetPageSizeByIndexF, vbLong, m_hPDF, i, VarPtr(otherPageSize))
            If (apiSuccess <> 0) Then
                If (firstPageSize.pdfWidth <> otherPageSize.pdfWidth) Or (firstPageSize.pdfHeight <> otherPageSize.pdfHeight) Then
                    IsPageSizeUniform = False
                    Exit Function
                End If
            End If
        Next i
        
    End If
    
End Function

'Load a given page inside the PDF.  Importantly, PAGES ARE 0-BASED, not 1-based.
' Returns TRUE if successful; FALSE otherwise.
' (This action will unload the previously loaded page, if one exists - this isn't strictly necessary,
' but it makes auto-management of resources much easier on this class.)
Friend Function LoadPage(ByVal idxPage As Long) As Boolean
    
    Const FUNC_NAME As String = "LoadPage"
    
    'Make sure a PDF is loaded
    If (m_hPDF = 0) Then
        InternalError FUNC_NAME, "no PDF loaded"
        Exit Function
    End If
    
    'Validate the page index
    If (idxPage < 0) Or (idxPage >= Me.GetPageCount()) Then
        InternalError FUNC_NAME, "bad page index"
    End If
    
    'Close the current page, if any
    If (m_hPage <> 0) Then CloseCurrentPage
    
    'Load the new page.
    m_hPage = CallCDeclW(FPDF_LoadPage, vbLong, m_hPDF, idxPage)
    LoadPage = (m_hPage <> 0)
    If (Not LoadPage) Then InternalError FUNC_NAME, "FPDF_GetLastError: " & GetLastPDFiumError()
    
End Function

'Attempt to load a PDF from file.  Returns a non-null handle if successful.
Friend Function LoadPDFFromFile(ByRef srcFile As String) As Long
    
    Const FUNC_NAME As String = "LoadPDFFromFile"
    LoadPDFFromFile = False
    
    'If the required 3rd-party library isn't available, bail
    If (Not Plugin_PDF.IsPDFiumAvailable()) Then Exit Function
    
    'If a PDF is already loaded, the caller may have just validated the file and is now proceeding
    ' with a full load.
    If (m_hPDF <> 0) Then
        
        'Check for matching validation strings and exit if the PDF has already been loaded
        If Strings.StringsEqual(srcFile, m_lastValidatedFile, True) Then
            LoadPDFFromFile = True
            Exit Function
        
        'This is a new PDF; close the existing handle
        Else
            m_lastValidatedFile = vbNullString
            ClosePDF
        End If
    
    '/Else: the file hasn't been validated; attempt a full load
    End If
    
    'From the header:
    '// Function: FPDF_LoadDocument
    '//          Open and load a PDF document.
    '// Parameters:
    '//          file_path -  Path to the PDF file (including extension).
    '//          password  -  A string used as the password for the PDF file.
    '//                       If no password is needed, empty or NULL can be used.
    '//                       See comments below regarding the encoding.
    '// Return value:
    '//          A handle to the loaded document, or NULL on failure.
    '// Comments:
    '//          Loaded document can be closed by FPDF_CloseDocument().
    '//          If this function fails, you can use FPDF_GetLastError() to retrieve
    '//          the reason why it failed.
    '//
    '//          The encoding for |file_path| is UTF-8.
    '//
    '//          The encoding for |password| can be either UTF-8 or Latin-1. PDFs,
    '//          depending on the security handler revision, will only accept one or
    '//          the other encoding. If |password|'s encoding and the PDF's expected
    '//          encoding do not match, FPDF_LoadDocument() will automatically
    '//          convert |password| to the other encoding.
    'FPDF_EXPORT FPDF_DOCUMENT FPDF_CALLCONV
    'FPDF_LoadDocument(FPDF_STRING file_path, FPDF_BYTESTRING password);
    
    'Convert the path to a UTF-8 string
    Dim ptrFile As Long, utf8Bytes() As Byte, utf8Len As Long
    Strings.UTF8FromString srcFile, utf8Bytes, utf8Len
    If (utf8Len > 0) Then ptrFile = VarPtr(utf8Bytes(0))
    
    'If a password was supplied, convert that to UTF-8 too.
    ' (First, we'll always try *without* a password.)
    Dim inPassword As String
    
TryAgainWithNewPassword:
    Dim ptrPassword As Long
    ptrPassword = 0
    If (LenB(inPassword) <> 0) Then
        Dim utf8Password() As Byte, utf8PasswordLen As Long
        Strings.UTF8FromString inPassword, utf8Password, utf8PasswordLen
        If (utf8PasswordLen > 0) Then ptrPassword = VarPtr(utf8Password(0))
    End If
    
    'Only proceed if the source path was successfully converted to UTF8
    If (ptrFile <> 0) Then
        m_hPDF = CallCDeclW(FPDF_LoadDocument, vbLong, ptrFile, ptrPassword)
        If (m_hPDF <> 0) Then m_lastValidatedFile = srcFile
    End If
    
    'If the PDF is password-protected, we can prompt the user for a password and try again
    Dim lastErr As FPDF_ErrorCodes
    lastErr = FPDF_ERR_SUCCESS
    If (m_hPDF = 0) Then
        lastErr = GetLastPDFiumError()
        InternalError FUNC_NAME, "FPDF_GetLastError: " & lastErr
    End If
    
    'As of Jan 2026, pdfium is unreliable about returning ERR_PASSWORD when a file is password-protected,
    ' so we'll need an alternate way to determine if a password prompt should be presented to the user.
    ' (Note that password input is *always* disabled during batch processes.)
    Dim passwordRequired As Boolean
    passwordRequired = (m_hPDF = 0) And (lastErr = FPDF_ERR_PASSWORD)
    If (Not passwordRequired) Then
        passwordRequired = (m_hPDF = 0) And IsFileEncrypted(srcFile)
        If passwordRequired Then PDDebug.LogAction "PDF is encrypted; user needs to supply a password..."
    End If
    
    'Prompt the user for a password and if one is supplied, attempt import again.
    ' (If *no* password is supplied, simply quit and report failure.)
    If passwordRequired And (Macros.GetMacroStatus() <> MacroBATCH) Then
        
        Dim dlgTitle As String, dlgText As String
        If (Not g_Language Is Nothing) Then
            dlgTitle = g_Language.TranslateMessage("Enter password")
            dlgText = g_Language.TranslateMessage("This file is password-protected.  Please enter a password:")
        End If
        
        If (Dialogs.PromptInputBox(dlgTitle, dlgText, inPassword) = vbOK) Then
            If (LenB(inPassword) > 0) Then GoTo TryAgainWithNewPassword
        End If
        
    End If
    
    'Return success if a non-null handle was generated
    LoadPDFFromFile = (m_hPDF <> 0)
    
End Function

'Non-PDFium way to see if a PDF file is encrypted.  If a PDF file *is* encrypted and the file fails to load,
' we'll prompt the user for a password.
'
'Note that encryption and password-protection do *not* correlate 1:1; for example, a file can be encrypted solely
' to prevent modifications (but no password is required for opening).
'
'I added this function to PD because as of Jan 2026, FPDF_GetLastError() in x86 builds always reports 0 for reasons
' currently unknown.  While I track down that bug, this provides a "reasonably" good way to distinguish between
' "failed to load PDF because file is malformed" and "failed to load PDF because password is potentially required".
Private Function IsFileEncrypted(ByRef srcFile As String) As Boolean
    
    IsFileEncrypted = False
    
    'Open a stream on the target file.
    Dim cStream As pdStream
    Set cStream = New pdStream
    If cStream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFile) Then
        
        'Load the last ~1 MB of the file as a string.  (We probably need *way* less than that, but this is a safe amount.)
        Const ABOUT_1_MB As Long = 1024000
        
        Dim pos1MBfromEnd As Long, sizeToLoad As Long
        pos1MBfromEnd = PDMath.Max2Int(0, cStream.GetStreamSize() - ABOUT_1_MB)
        sizeToLoad = cStream.GetStreamSize() - pos1MBfromEnd
        cStream.SetPosition pos1MBfromEnd
        
        Dim strTest As String
        strTest = cStream.ReadString_ASCII(sizeToLoad)
        
        'With our test string loaded, we can release the source file
        cStream.StopStream True
        
        'Search backward in the string, looking for the text "/ntrailer".
        ' (PDFs can use Lf or CrLf as line-endings, so it would be extra work to search for the more useful "/ntrailer/n"
        Dim strSearch As String, strPos As Long
        strSearch = vbLf & "trailer"
        strPos = InStrRev(strTest, strSearch)
        If (strPos > 0) Then
        
            'Starting from the trailer and moving forward, look for an /Encrypt entry
            strSearch = "/Encrypt "
            Dim posEncrypt As Long
            posEncrypt = InStr(strPos, strTest, strSearch)
            
            'A file with an encryption tag is encrypted.
            IsFileEncrypted = (posEncrypt > 0)
            
            'If you wanted to further see if encryption is tied to a user-supplied password,
            ' you could do so following the ideas at this link:
            ' https://stackoverflow.com/questions/35594240/how-to-check-if-pdf-is-password-protected-using-static-tools
            '
            '...but that's overkill for PD at present.
        
        End If
        
    End If
    
End Function

'Render the current page to an arbitrary DC.
' Units are in PIXELS here - not points!
Friend Sub RenderCurrentPageToDC(ByVal dstDC As Long, ByVal dstLeft As Long, ByVal dstTop As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, Optional ByVal pageRotate As PDFium_Orientation = FPDF_Normal, Optional ByVal renderOptions As PDFium_RenderOptions = 0&)
    
    Const FUNC_NAME As String = "RenderCurrentPageToDC"
    If (m_hPage = 0) Then
        InternalError FUNC_NAME, "no page loaded"
        Exit Sub
    End If
    
    CallCDeclW FPDF_RenderPage, vbEmpty, dstDC, m_hPage, dstLeft, dstTop, dstWidth, dstHeight, pageRotate, renderOptions
    
End Sub

'Render the current page to a pdDIB object.
' Units are in PIXELS here - not points!
Friend Sub RenderCurrentPageToPDDib(ByRef dstDIB As pdDIB, ByVal dstLeft As Long, ByVal dstTop As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, Optional ByVal pageRotate As PDFium_Orientation = FPDF_Normal, Optional ByVal renderOptions As PDFium_RenderOptions = 0&)
    
    Const FUNC_NAME As String = "RenderCurrentPageToPDDib"
    If (m_hPage = 0) Then
        InternalError FUNC_NAME, "no page loaded"
        Exit Sub
    End If
    
    'PDFium can render to a few different pixel formats; we only need BGRA for PD
    Const FPDFBitmap_BGRA = 4
    
    'Wrap a PDFium bitmap around the pdDIB object
    'FPDF_BITMAP FPDFBitmap_CreateEx(int width, int height, int format, void* first_scan, int stride);
    Dim hBitmap As Long
    hBitmap = CallCDeclW(FPDFBitmap_CreateEx, vbLong, dstDIB.GetDIBWidth, dstDIB.GetDIBHeight, FPDFBitmap_BGRA, dstDIB.GetDIBPointer, dstDIB.GetDIBStride)
    
    'Render directly into the DIB
    CallCDeclW FPDF_RenderPageBitmap, vbEmpty, hBitmap, m_hPage, dstLeft, dstTop, dstWidth, dstHeight, pageRotate, renderOptions
    
    'CallCDeclW FPDF_RenderPage, vbEmpty, dstDC, m_hPage, dstLeft, dstTop, dstWidth, dstHeight, pageRotate, renderOptions
    
    'Free the temporary pdfium bitmap object we created
    CallCDeclW FPDFBitmap_Destroy, vbEmpty, hBitmap
    
End Sub

Private Sub Class_Initialize()
    
    'Initialize all module-level arrays
    ReDim m_vType(0 To MAX_PARAM_COUNT - 1) As Integer
    ReDim m_vPtr(0 To MAX_PARAM_COUNT - 1) As Long
        
    'Copy the central list of DLL proc addresses
    Plugin_PDF.CopyPDFiumProcAddresses m_ProcAddresses
    
End Sub

Private Sub Class_Terminate()
    
    'Close the current page, if any
    If (m_hPage <> 0) Then CloseCurrentPage
    
    'Free any open PDF handles before exiting
    ClosePDF
    
End Sub

Private Function GetLastPDFiumError() As FPDF_ErrorCodes
    
    'If the required 3rd-party library isn't available, bail
    If (Not Plugin_PDF.IsPDFiumAvailable()) Then Exit Function
    
    'From the header:
    ' "If the previous SDK call succeeded, the return value of this
    '  function is not defined. This function only works in conjunction
    '  with APIs that mention FPDF_GetLastError() in their documentation."
    'FPDF_EXPORT unsigned long FPDF_CALLCONV FPDF_GetLastError();
    Dim tmpReturn As Long
    tmpReturn = CallCDeclW(FPDF_GetLastError, vbLong)
    GetLastPDFiumError = tmpReturn
    
End Function

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
        PDDebug.LogAction "pdPDF." & funcName & "() reported an error: " & errDescription
    Else
        Debug.Print "pdPDF." & funcName & "() reported an error: " & errDescription
    End If
End Sub
