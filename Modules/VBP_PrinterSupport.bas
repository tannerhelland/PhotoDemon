Attribute VB_Name = "Printing"
'***************************************************************************
'Printer support functions
'Copyright 2003-2015 by Tanner Helland
'Created: 4/April/03
'Last updated: 09/August/14
'Last update: perform necessary cleanup for printer temp files in Vista+
'
'This module includes code based off an article written by Cassandra Roads of Professional Logics Corporation (PLC).
' You can download the original, unmodified version of Cassandra's code from this link (good as of 12 Nov 2014):
' http://www.tek-tips.com/faqs.cfm?fid=3603
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Per its name, DeviceCapabilities is used to retrieve printer capabilities
Private Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal lpsDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, lpOutput As Any, ByVal lpDevMode As Long) As Long

'Constants that define printer attributes we want to request from the printer (supported paper sizes, etc)
Private Const DC_PAPERS As Long = 2
Private Const DC_PAPERSIZE As Long = 3
Private Const DC_PAPERNAMES As Long = 16

'Printer orientation constants
Private Const DMORIENT_PORTRAIT = 1
Private Const DMORIENT_LANDSCAPE = 2

'Internal Windows Private Constants that define paper sizes.  As paper size text may be localized, it can't be relied upon for
' matching standard image sizes.  1/10th mm differences in reported sizes for standard media also makes it unreliable.
' These Private Constants can be used to absolutely match up a given paper size with a standard size.
Private Const DMPAPER_LETTER = 1
Private Const DMPAPER_LEGAL = 5
Private Const DMPAPER_10X11 = 45
Private Const DMPAPER_10X14 = 16
Private Const DMPAPER_11X17 = 17
Private Const DMPAPER_15X11 = 46
Private Const DMPAPER_9X11 = 44
Private Const DMPAPER_A_PLUS = 57
Private Const DMPAPER_A2 = 66
Private Const DMPAPER_A3 = 8
Private Const DMPAPER_A3_EXTRA = 63
Private Const DMPAPER_A3_EXTRA_TRANSVERSE = 68
Private Const DMPAPER_A3_TRANSVERSE = 67
Private Const DMPAPER_A4 = 9
Private Const DMPAPER_A4_EXTRA = 53
Private Const DMPAPER_A4_PLUS = 60
Private Const DMPAPER_A4_TRANSVERSE = 55
Private Const DMPAPER_A4SMALL = 10
Private Const DMPAPER_A5 = 11
Private Const DMPAPER_A5_EXTRA = 64
Private Const DMPAPER_A5_TRANSVERSE = 61
Private Const DMPAPER_B_PLUS = 58
Private Const DMPAPER_B4 = 12
Private Const DMPAPER_B5 = 13
Private Const DMPAPER_B5_EXTRA = 65
Private Const DMPAPER_B5_TRANSVERSE = 62
Private Const DMPAPER_CSHEET = 24
Private Const DMPAPER_DSHEET = 25
Private Const DMPAPER_ENV_10 = 20
Private Const DMPAPER_ENV_11 = 21
Private Const DMPAPER_ENV_12 = 22
Private Const DMPAPER_ENV_14 = 23
Private Const DMPAPER_ENV_9 = 19
Private Const DMPAPER_ENV_B4 = 33
Private Const DMPAPER_ENV_B5 = 34
Private Const DMPAPER_ENV_B6 = 35
Private Const DMPAPER_ENV_C3 = 29
Private Const DMPAPER_ENV_C4 = 30
Private Const DMPAPER_ENV_C5 = 28
Private Const DMPAPER_ENV_C6 = 31
Private Const DMPAPER_ENV_C65 = 32
Private Const DMPAPER_ENV_DL = 27
Private Const DMPAPER_ENV_INVITE = 47
Private Const DMPAPER_ENV_ITALY = 36
Private Const DMPAPER_ENV_MONARCH = 37
Private Const DMPAPER_ENV_PERSONAL = 38
Private Const DMPAPER_ESHEET = 26
Private Const DMPAPER_EXECUTIVE = 7
Private Const DMPAPER_FANFOLD_LGL_GERMAN = 41
Private Const DMPAPER_FANFOLD_STD_GERMAN = 40
Private Const DMPAPER_FANFOLD_US = 39
Private Const DMPAPER_FIRST = 1
Private Const DMPAPER_FOLIO = 14
Private Const DMPAPER_ISO_B4 = 42
Private Const DMPAPER_JAPANESE_POSTCARD = 43
Private Const DMPAPER_LAST = 41
Private Const DMPAPER_LEDGER = 4
Private Const DMPAPER_LEGAL_EXTRA = 51
Private Const DMPAPER_LETTER_EXTRA = 50
Private Const DMPAPER_LETTER_EXTRA_TRANSVERSE = 56
Private Const DMPAPER_LETTER_PLUS = 59
Private Const DMPAPER_LETTER_TRANSVERSE = 54
Private Const DMPAPER_LETTERSMALL = 2
Private Const DMPAPER_NOTE = 18
Private Const DMPAPER_QUARTO = 15
Private Const DMPAPER_STATEMENT = 6
Private Const DMPAPER_TABLOID = 3
Private Const DMPAPER_TABLOID_EXTRA = 52
Private Const DMPAPER_USER = 256

'If the user has attempted to print during this session, this value will be set to TRUE, and the corresponding temp file
' will be marked.  When PD closes down, if we can access the file, assume printing is complete and delete it.
Private m_userPrintedThisSession As Boolean
Private m_temporaryPrintPath As String

'Call this function at shutdown time to perform any printer-related cleanup
Public Sub performPrinterCleanup()

    If m_userPrintedThisSession Then
        
        Dim cFile As pdFSO
        Set cFile = New pdFSO
        
        If cFile.FileExist(m_temporaryPrintPath) Then cFile.KillFile m_temporaryPrintPath
        
    End If

End Sub

'This simple function can be used to route printing through the default "Windows Photo Printer" dialog.
Public Sub printViaWindowsPhotoPrinter()

    Message "Preparing image for printing..."
    
    'Create a temporary copy of the currently active image, composited against a white background
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    pdImages(g_CurrentImage).getCompositedImage tmpDIB, False
    If tmpDIB.getDIBColorDepth <> 24 Then tmpDIB.convertTo24bpp
    
    'Windows itself handles the heavy lifting for printing.  We just write a temp file that contains the image data.
    Dim tmpFilename As String
    tmpFilename = g_UserPreferences.GetTempPath & "PhotoDemon_print.png"
    
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "Preparing to print: " & tmpFilename
    #End If
    
    'Write the temporary DIB out to a temporary PNG file, then free it
    Saving.QuickSaveDIBAsPNG tmpFilename, tmpDIB
    Set tmpDIB = Nothing
    
    'Store the print state, so we can perform clean-up as necessary at shutdown time
    m_userPrintedThisSession = True
    m_temporaryPrintPath = tmpFilename
    
    Message "Image successfully sent to Windows Photo Printer."
    
    'Once that is complete, use ShellExecute to launch the default Windows Photo Print dialog.  (Note that this
    ' DOES NOT work on XP.)
    Dim actionName As String
    actionName = "print"
    
    ShellExecute getModalOwner().hWnd, StrPtr(actionName), StrPtr(tmpFilename), 0&, 0&, SW_SHOWNORMAL
    
End Sub

'Use the API to retrieve all supported paper sizes for the current printer
Public Function getPaperSizes(ByVal printerIndex As Long, ByRef paperSizeNames() As String, ByRef paperIDs() As Integer, ByRef exactPaperSizes() As POINTAPI) As Boolean

    'We're going to use the printer name and port frequently, so cache their names in advance
    Dim pName As String, pPort As String
    pName = Printers(printerIndex).DeviceName
    pPort = Printers(printerIndex).Port

    'Start by retrieving the paper size count; we need this to prep all our arrays
    Dim numOfPaperSizes As Long
    numOfPaperSizes = DeviceCapabilities(pName, pPort, DC_PAPERNAMES, ByVal vbNullString, 0)
    
    'Prep the various size-related arrays
    ReDim paperSizeNames(0 To numOfPaperSizes - 1) As String
    ReDim paperIDs(0 To numOfPaperSizes - 1) As Integer
    ReDim exactPaperSizes(0 To numOfPaperSizes - 1) As POINTAPI
    
    'Paper size names are returned as one giant-ass string.  Each individual name occupies 64 characters, and each
    ' is null-terminated (unless it consumes all 64 characters, in which case we have to terminate it manually).
    Dim giantPaperNameList As String
    giantPaperNameList = String(numOfPaperSizes * 64, 0)
    
    DeviceCapabilities pName, pPort, DC_PAPERNAMES, ByVal giantPaperNameList, 0
    
    'Now we have to manually parse the returned string into the array
    Dim i As Long
    Dim tmpString As String
    
    For i = 0 To numOfPaperSizes - 1
        tmpString = Mid$(giantPaperNameList, (i * 64) + 1, 64)
        tmpString = TrimNull(tmpString)
        paperSizeNames(i) = tmpString
    Next i
    
    'Next comes the matching list of paper size IDs.  See the matching list of dmPaperSize constants at:
    ' http://msdn.microsoft.com/en-us/library/windows/desktop/dd183565%28v=vs.85%29.aspx
    DeviceCapabilities pName, pPort, DC_PAPERS, paperIDs(0), 0
    
    'Next comes the list of paper widths and heights.  These are mm-accurate measurements of each paper size,
    ' which is hugely helpful for rendering our print preview accurately.
    DeviceCapabilities pName, pPort, DC_PAPERSIZE, exactPaperSizes(0), 0
        
    getPaperSizes = True

End Function
        

