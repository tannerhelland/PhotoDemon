VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCommonDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'I'm pretty sure I got this common dialog wrapper from http://vbaccelerator.com
'This is an older version of this project, but it still works like a champ.  Many
'thanks to Steve McMahon and his sweet site, http://vbaccelerator.com

Option Explicit


Public Enum EErrorCommonDialog
    eeBaseCommonDialog = 13450  ' CommonDialog
End Enum

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalCompact Lib "kernel32" (ByVal dwMinFree As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Sub CopyMemoryStr Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, ByVal lpvSource As String, ByVal cbCopy As Long)

Private Const MAX_PATH = 260
Private Const MAX_FILE = 260

Private Type OPENFILENAME
    lStructSize As Long          ' Filled with UDT size
    hWndOwner As Long            ' Tied to Owner
    hInstance As Long            ' Ignored (used only by templates)
    lpstrFilter As String        ' Tied to Filter
    lpstrCustomFilter As String  ' Ignored (exercise for reader)
    nMaxCustFilter As Long       ' Ignored (exercise for reader)
    nFilterIndex As Long         ' Tied to FilterIndex
    lpstrFile As String          ' Tied to FileName
    nMaxFile As Long             ' Handled internally
    lpstrFileTitle As String     ' Tied to FileTitle
    nMaxFileTitle As Long        ' Handled internally
    lpstrInitialDir As String    ' Tied to InitDir
    lpstrTitle As String         ' Tied to DlgTitle
    flags As Long                ' Tied to Flags
    nFileOffset As Integer       ' Ignored (exercise for reader)
    nFileExtension As Integer    ' Ignored (exercise for reader)
    lpstrDefExt As String        ' Tied to DefaultExt
    lCustData As Long            ' Ignored (needed for hooks)
    lpfnHook As Long             ' Ignored (good luck with hooks)
    lpTemplateName As Long       ' Ignored (good luck with templates)
End Type

Private Declare Function GetOpenFileName Lib "COMDLG32" _
    Alias "GetOpenFileNameA" (file As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "COMDLG32" _
    Alias "GetSaveFileNameA" (file As OPENFILENAME) As Long
Private Declare Function GetFileTitle Lib "COMDLG32" _
    Alias "GetFileTitleA" (ByVal szFile As String, _
    ByVal szTitle As String, ByVal cbBuf As Long) As Long

Public Enum EOpenFile
    OFN_READONLY = &H1
    OFN_OVERWRITEPROMPT = &H2
    OFN_HIDEREADONLY = &H4
    OFN_NOCHANGEDIR = &H8
    OFN_SHOWHELP = &H10
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_NOVALIDATE = &H100
    OFN_ALLOWMULTISELECT = &H200
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_PATHMUSTEXIST = &H800
    OFN_FILEMUSTEXIST = &H1000
    OFN_CREATEPROMPT = &H2000
    OFN_SHAREAWARE = &H4000
    OFN_NOREADONLYRETURN = &H8000&
    OFN_NOTESTFILECREATE = &H10000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOLONGNAMES = &H40000
    OFN_EXPLORER = &H80000
    OFN_NODEREFERENCELINKS = &H100000
    OFN_LONGNAMES = &H200000
End Enum

Private Type TCHOOSECOLOR
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long
End Type

Private Declare Function ChooseColor Lib "COMDLG32.DLL" _
    Alias "ChooseColorA" (Color As TCHOOSECOLOR) As Long

Public Enum EChooseColor
    CC_RGBInit = &H1
    CC_FullOpen = &H2
    CC_PreventFullOpen = &H4
    CC_ColorShowHelp = &H8
' Win95 only
    CC_SolidColor = &H80
    CC_AnyColor = &H100
' End Win95 only
    CC_ENABLEHOOK = &H10
    CC_ENABLETEMPLATE = &H20
    CC_EnableTemplateHandle = &H40
End Enum
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Type TCHOOSEFONT
    lStructSize As Long         ' Filled with UDT size
    hWndOwner As Long           ' Caller's window handle
    hdc As Long                 ' Printer DC/IC or NULL
    lpLogFont As Long           ' Pointer to LOGFONT
    iPointSize As Long          ' 10 * size in points of font
    flags As Long               ' Type flags
    rgbColors As Long           ' Returned text color
    lCustData As Long           ' Data passed to hook function
    lpfnHook As Long            ' Pointer to hook function
    lpTemplateName As Long      ' Custom template name
    hInstance As Long           ' Instance handle for template
    lpszStyle As String         ' Return style field
    nFontType As Integer        ' Font type bits
    iAlign As Integer           ' Filler
    nSizeMin As Long            ' Minimum point size allowed
    nSizeMax As Long            ' Maximum point size allowed
End Type
Private Declare Function ChooseFont Lib "COMDLG32" _
    Alias "ChooseFontA" (chfont As TCHOOSEFONT) As Long

Private Const LF_FACESIZE = 32
Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Public Enum EChooseFont
    CF_ScreenFonts = &H1
    CF_PrinterFonts = &H2
    CF_BOTH = &H3
    CF_FontShowHelp = &H4
    CF_UseStyle = &H80
    CF_EFFECTS = &H100
    CF_AnsiOnly = &H400
    CF_NoVectorFonts = &H800
    CF_NoOemFonts = CF_NoVectorFonts
    CF_NoSimulations = &H1000
    CF_LimitSize = &H2000
    CF_FixedPitchOnly = &H4000
    CF_WYSIWYG = &H8000&  ' Must also have ScreenFonts And PrinterFonts
    CF_ForceFontExist = &H10000
    CF_ScalableOnly = &H20000
    CF_TTOnly = &H40000
    CF_NoFaceSel = &H80000
    CF_NoStyleSel = &H100000
    CF_NoSizeSel = &H200000
    ' Win95 only
    CF_SelectScript = &H400000
    CF_NoScriptSel = &H800000
    CF_NoVertFonts = &H1000000

    CF_InitToLogFontStruct = &H40
    CF_Apply = &H200
    CF_EnableHook = &H8
    CF_EnableTemplate = &H10
    CF_EnableTemplateHandle = &H20
    CF_FontNotSupported = &H238
End Enum

' These are extra nFontType bits that are added to what is returned to the
' EnumFonts callback routine

Public Enum EFontType
    Simulated_FontType = &H8000&
    Printer_FontType = &H4000
    Screen_FontType = &H2000
    Bold_FontType = &H100
    Italic_FontType = &H200
    Regular_FontType = &H400
End Enum

Private Type TPRINTDLG
    lStructSize As Long
    hWndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hdc As Long
    flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As Long
    lpSetupTemplateName As Long
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type

'  DEVMODE collation selections
Private Const DMCOLLATE_FALSE = 0
Private Const DMCOLLATE_TRUE = 1

Private Declare Function PrintDlg Lib "COMDLG32.DLL" _
    Alias "PrintDlgA" (prtdlg As TPRINTDLG) As Integer

Public Enum EPrintDialog
    PD_ALLPAGES = &H0
    PD_SELECTION = &H1
    PD_PAGENUMS = &H2
    PD_NOSELECTION = &H4
    PD_NOPAGENUMS = &H8
    PD_COLLATE = &H10
    PD_PRINTTOFILE = &H20
    PD_PRINTSETUP = &H40
    PD_NOWARNING = &H80
    PD_RETURNDC = &H100
    PD_RETURNIC = &H200
    PD_RETURNDEFAULT = &H400
    PD_SHOWHELP = &H800
    PD_ENABLEPRINTHOOK = &H1000
    PD_ENABLESETUPHOOK = &H2000
    PD_ENABLEPRINTTEMPLATE = &H4000
    PD_ENABLESETUPTEMPLATE = &H8000&
    PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
    PD_ENABLESETUPTEMPLATEHANDLE = &H20000
    PD_USEDEVMODECOPIES = &H40000
    PD_USEDEVMODECOPIESANDCOLLATE = &H40000
    PD_DISABLEPRINTTOFILE = &H80000
    PD_HIDEPRINTTOFILE = &H100000
    PD_NONETWORKBUTTON = &H200000
End Enum

Private Type DEVNAMES
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
End Type

Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Type DevMode
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

' New Win95 Page Setup dialogs are up to you
Private Type POINTL
    x As Long
    Y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Private Type TPAGESETUPDLG
    lStructSize                 As Long
    hWndOwner                   As Long
    hDevMode                    As Long
    hDevNames                   As Long
    flags                       As Long
    ptPaperSize                 As POINTL
    rtMinMargin                 As RECT
    rtMargin                    As RECT
    hInstance                   As Long
    lCustData                   As Long
    lpfnPageSetupHook           As Long
    lpfnPagePaintHook           As Long
    lpPageSetupTemplateName     As Long
    hPageSetupTemplate          As Long
End Type

' EPaperSize constants same as vbPRPS constants
Public Enum EPaperSize
    epsLetter = 1          ' Letter, 8 1/2 x 11 in.
    epsLetterSmall         ' Letter Small, 8 1/2 x 11 in.
    epsTabloid             ' Tabloid, 11 x 17 in.
    epsLedger              ' Ledger, 17 x 11 in.
    epsLegal               ' Legal, 8 1/2 x 14 in.
    epsStatement           ' Statement, 5 1/2 x 8 1/2 in.
    epsExecutive           ' Executive, 7 1/2 x 10 1/2 in.
    epsA3                  ' A3, 297 x 420 mm
    epsA4                  ' A4, 210 x 297 mm
    epsA4Small             ' A4 Small, 210 x 297 mm
    epsA5                  ' A5, 148 x 210 mm
    epsB4                  ' B4, 250 x 354 mm
    epsB5                  ' B5, 182 x 257 mm
    epsFolio               ' Folio, 8 1/2 x 13 in.
    epsQuarto              ' Quarto, 215 x 275 mm
    eps10x14               ' 10 x 14 in.
    eps11x17               ' 11 x 17 in.
    epsNote                ' Note, 8 1/2 x 11 in.
    epsEnv9                ' Envelope #9, 3 7/8 x 8 7/8 in.
    epsEnv10               ' Envelope #10, 4 1/8 x 9 1/2 in.
    epsEnv11               ' Envelope #11, 4 1/2 x 10 3/8 in.
    epsEnv12               ' Envelope #12, 4 1/2 x 11 in.
    epsEnv14               ' Envelope #14, 5 x 11 1/2 in.
    epsCSheet              ' C size sheet
    epsDSheet              ' D size sheet
    epsESheet              ' E size sheet
    epsEnvDL               ' Envelope DL, 110 x 220 mm
    epsEnvC3               ' Envelope C3, 324 x 458 mm
    epsEnvC4               ' Envelope C4, 229 x 324 mm
    epsEnvC5               ' Envelope C5, 162 x 229 mm
    epsEnvC6               ' Envelope C6, 114 x 162 mm
    epsEnvC65              ' Envelope C65, 114 x 229 mm
    epsEnvB4               ' Envelope B4, 250 x 353 mm
    epsEnvB5               ' Envelope B5, 176 x 250 mm
    epsEnvB6               ' Envelope B6, 176 x 125 mm
    epsEnvItaly            ' Envelope, 110 x 230 mm
    epsenvmonarch          ' Envelope Monarch, 3 7/8 x 7 1/2 in.
    epsEnvPersonal         ' Envelope, 3 5/8 x 6 1/2 in.
    epsFanfoldUS           ' U.S. Standard Fanfold, 14 7/8 x 11 in.
    epsFanfoldStdGerman    ' German Standard Fanfold, 8 1/2 x 12 in.
    epsFanfoldLglGerman    ' German Legal Fanfold, 8 1/2 x 13 in.
    epsUser = 256          ' User-defined
End Enum

' EPrintQuality constants same as vbPRPQ constants
Public Enum EPrintQuality
    epqDraft = -1
    epqLow = -2
    epqMedium = -3
    epqHigh = -4
End Enum

Public Enum EOrientation
    eoPortrait = 1
    eoLandscape
End Enum

Private Declare Function PageSetupDlg Lib "COMDLG32" _
    Alias "PageSetupDlgA" (lppage As TPAGESETUPDLG) As Boolean

Public Enum EPageSetup
    PSD_Defaultminmargins = &H0 ' Default (printer's)
    PSD_InWinIniIntlMeasure = &H0
    PSD_MINMARGINS = &H1
    PSD_MARGINS = &H2
    PSD_INTHOUSANDTHSOFINCHES = &H4
    PSD_INHUNDREDTHSOFMILLIMETERS = &H8
    PSD_DISABLEMARGINS = &H10
    PSD_DISABLEPRINTER = &H20
    PSD_NoWarning = &H80
    PSD_DISABLEORIENTATION = &H100
    PSD_ReturnDefault = &H400
    PSD_DISABLEPAPER = &H200
    PSD_ShowHelp = &H800
    PSD_EnablePageSetupHook = &H2000
    PSD_EnablePageSetupTemplate = &H8000&
    PSD_EnablePageSetupTemplateHandle = &H20000
    PSD_EnablePagePaintHook = &H40000
    PSD_DisablePagePainting = &H80000
End Enum

Public Enum EPageSetupUnits
    epsuInches
    epsuMillimeters
End Enum

' Common dialog errors

Private Declare Function CommDlgExtendedError Lib "COMDLG32" () As Long

Public Enum EDialogError
    CDERR_DIALOGFAILURE = &HFFFF

    CDERR_GENERALCODES = &H0
    CDERR_STRUCTSIZE = &H1
    CDERR_INITIALIZATION = &H2
    CDERR_NOTEMPLATE = &H3
    CDERR_NOHINSTANCE = &H4
    CDERR_LOADSTRFAILURE = &H5
    CDERR_FINDRESFAILURE = &H6
    CDERR_LOADRESFAILURE = &H7
    CDERR_LOCKRESFAILURE = &H8
    CDERR_MEMALLOCFAILURE = &H9
    CDERR_MEMLOCKFAILURE = &HA
    CDERR_NOHOOK = &HB
    CDERR_REGISTERMSGFAIL = &HC

    PDERR_PRINTERCODES = &H1000
    PDERR_SETUPFAILURE = &H1001
    PDERR_PARSEFAILURE = &H1002
    PDERR_RETDEFFAILURE = &H1003
    PDERR_LOADDRVFAILURE = &H1004
    PDERR_GETDEVMODEFAIL = &H1005
    PDERR_INITFAILURE = &H1006
    PDERR_NODEVICES = &H1007
    PDERR_NODEFAULTPRN = &H1008
    PDERR_DNDMMISMATCH = &H1009
    PDERR_CREATEICFAILURE = &H100A
    PDERR_PRINTERNOTFOUND = &H100B
    PDERR_DEFAULTDIFFERENT = &H100C

    CFERR_CHOOSEFONTCODES = &H2000
    CFERR_NOFONTS = &H2003
    CFERR_MAXLESSTHANMIN = &H2003

    FNERR_FILENAMECODES = &H3000
    FNERR_SUBCLASSFAILURE = &H3001
    FNERR_INVALIDFILENAME = &H3002
    FNERR_BUFFERTOOSMALL = &H3003

    CCERR_CHOOSECOLORCODES = &H5000
End Enum

' Array of custom colors lasts for life of app
Private alCustom(0 To 15) As Long, fNotFirst As Boolean

Public Enum EPrintRange
    eprAll
    eprPageNumbers
    eprSelection
End Enum
Private m_lApiReturn As Long
Private m_lExtendedError As Long
Private m_dvmode As DevMode

Public Property Get APIReturn() As Long
    'return object's APIReturn property
    APIReturn = m_lApiReturn
End Property
Public Property Get ExtendedError() As Long
    'return object's ExtendedError property
    ExtendedError = m_lExtendedError
End Property

#If fComponent Then
Private Sub Class_Initialize()
    InitColors
End Sub
#End If

Function VBGetOpenFileName(Filename As String, _
                           Optional FileTitle As String, _
                           Optional FileMustExist As Boolean = True, _
                           Optional MultiSelect As Boolean = False, _
                           Optional ReadOnly As Boolean = False, _
                           Optional HideReadOnly As Boolean = False, _
                           Optional Filter As String = "All (*.*)| *.*", _
                           Optional FilterIndex As Long = 1, _
                           Optional InitDir As String, _
                           Optional DlgTitle As String, _
                           Optional DefaultExt As String, _
                           Optional Owner As Long = -1, _
                           Optional flags As Long = 0) As Boolean

    Dim opfile As OPENFILENAME, s As String
    
    m_lApiReturn = 0
    m_lExtendedError = 0

With opfile
    .lStructSize = Len(opfile)
    
    ' Add in specific flags and strip out non-VB flags
    
    .flags = (-FileMustExist * OFN_FILEMUSTEXIST) Or _
            (-MultiSelect * OFN_ALLOWMULTISELECT) Or _
             (-ReadOnly * OFN_READONLY) Or _
             (-HideReadOnly * OFN_HIDEREADONLY) Or _
             (flags And CLng(Not (OFN_ENABLEHOOK Or _
                                  OFN_ENABLETEMPLATE)))
    ' Owner can take handle of owning window
    If Owner <> -1 Then .hWndOwner = Owner
    ' InitDir can take initial directory string
    .lpstrInitialDir = InitDir
    ' DefaultExt can take default extension
    .lpstrDefExt = DefaultExt
    ' DlgTitle can take dialog box title
    .lpstrTitle = DlgTitle
    
    ' To make Windows-style filter, replace | and : with nulls
    Dim ch As String, i As Integer
    For i = 1 To Len(Filter)
        ch = Mid$(Filter, i, 1)
        If ch = "|" Or ch = ":" Then
            s = s & vbNullChar
        Else
            s = s & ch
        End If
    Next
    ' Put double null at end
    s = s & vbNullChar & vbNullChar
    .lpstrFilter = s
    .nFilterIndex = FilterIndex

    ' Pad file and file title buffers to maximum path
    s = Filename & String$(MAX_PATH - Len(Filename), 0)
    .lpstrFile = s
    .nMaxFile = MAX_PATH
    s = FileTitle & String$(MAX_FILE - Len(FileTitle), 0)
    .lpstrFileTitle = s
    .nMaxFileTitle = MAX_FILE
    ' All other fields set to zero
    
    m_lApiReturn = GetOpenFileName(opfile)
    Select Case m_lApiReturn
    Case 1
        ' Success
        VBGetOpenFileName = True
        Filename = StrZToStr(.lpstrFile)
        FileTitle = StrZToStr(.lpstrFileTitle)
        flags = .flags
        ' Return the filter index
        FilterIndex = .nFilterIndex
        ' Look up the filter the user selected and return that
        Filter = FilterLookup(.lpstrFilter, FilterIndex)
        If (.flags And OFN_READONLY) Then ReadOnly = True
    Case 0
        ' Cancelled
        VBGetOpenFileName = False
        Filename = ""
        FileTitle = ""
        flags = 0
        FilterIndex = -1
        Filter = ""
    Case Else
        ' Extended error
        m_lExtendedError = CommDlgExtendedError()
        VBGetOpenFileName = False
        Filename = ""
        FileTitle = ""
        flags = 0
        FilterIndex = -1
        Filter = ""
    End Select
End With
End Function
Private Function StrZToStr(s As String) As String
    StrZToStr = Left$(s, lstrlen(s))
End Function

Function VBGetSaveFileName(Filename As String, _
                           Optional FileTitle As String, _
                           Optional OverWritePrompt As Boolean = True, _
                           Optional Filter As String = "All (*.*)| *.*", _
                           Optional FilterIndex As Long = 1, _
                           Optional InitDir As String, _
                           Optional DlgTitle As String, _
                           Optional DefaultExt As String, _
                           Optional Owner As Long = -1, _
                           Optional flags As Long) As Boolean
            
    Dim opfile As OPENFILENAME, s As String

    m_lApiReturn = 0
    m_lExtendedError = 0

With opfile
    .lStructSize = Len(opfile)
    
    ' Add in specific flags and strip out non-VB flags
    .flags = (-OverWritePrompt * OFN_OVERWRITEPROMPT) Or _
             OFN_HIDEREADONLY Or _
             (flags And CLng(Not (OFN_ENABLEHOOK Or _
                                  OFN_ENABLETEMPLATE)))
    ' Owner can take handle of owning window
    If Owner <> -1 Then .hWndOwner = Owner
    ' InitDir can take initial directory string
    .lpstrInitialDir = InitDir
    ' DefaultExt can take default extension
    .lpstrDefExt = DefaultExt
    ' DlgTitle can take dialog box title
    .lpstrTitle = DlgTitle
    
    ' Make new filter with bars (|) replacing nulls and double null at end
    Dim ch As String, i As Integer
    For i = 1 To Len(Filter)
        ch = Mid$(Filter, i, 1)
        If ch = "|" Or ch = ":" Then
            s = s & vbNullChar
        Else
            s = s & ch
        End If
    Next
    ' Put double null at end
    s = s & vbNullChar & vbNullChar
    .lpstrFilter = s
    .nFilterIndex = FilterIndex

    ' Pad file and file title buffers to maximum path
    s = Filename & String$(MAX_PATH - Len(Filename), 0)
    .lpstrFile = s
    .nMaxFile = MAX_PATH
    s = FileTitle & String$(MAX_FILE - Len(FileTitle), 0)
    .lpstrFileTitle = s
    .nMaxFileTitle = MAX_FILE
    ' All other fields zero
    
    m_lApiReturn = GetSaveFileName(opfile)
    Select Case m_lApiReturn
    Case 1
        VBGetSaveFileName = True
        Filename = StrZToStr(.lpstrFile)
        FileTitle = StrZToStr(.lpstrFileTitle)
        flags = .flags
        ' Return the filter index
        FilterIndex = .nFilterIndex
        ' Look up the filter the user selected and return that
        Filter = FilterLookup(.lpstrFilter, FilterIndex)
    Case 0
        ' Cancelled:
        VBGetSaveFileName = False
        Filename = ""
        FileTitle = ""
        flags = 0
        FilterIndex = 0
        Filter = ""
    Case Else
        ' Extended error:
        VBGetSaveFileName = False
        m_lExtendedError = CommDlgExtendedError()
        Filename = ""
        FileTitle = ""
        flags = 0
        FilterIndex = 0
        Filter = ""
    End Select
End With
End Function

Private Function FilterLookup(ByVal sFilters As String, ByVal iCur As Long) As String
    Dim iStart As Long, iEnd As Long, s As String
    iStart = 1
    If Len(sFilters) = 0 Then Exit Function
    Do
        ' Cut out both parts marked by null character
        iEnd = InStr(iStart, sFilters, vbNullChar)
        If iEnd = 0 Then Exit Function
        iEnd = InStr(iEnd + 1, sFilters, vbNullChar)
        If iEnd Then
            s = Mid$(sFilters, iStart, iEnd - iStart)
        Else
            s = Mid$(sFilters, iStart)
        End If
        iStart = iEnd + 1
        If iCur = 1 Then
            FilterLookup = s
            Exit Function
        End If
        iCur = iCur - 1
    Loop While iCur
End Function

Function VBGetFileTitle(sFile As String) As String
    Dim sFileTitle As String, cFileTitle As Integer

    cFileTitle = MAX_PATH
    sFileTitle = String$(MAX_PATH, 0)
    cFileTitle = GetFileTitle(sFile, sFileTitle, MAX_PATH)
    If cFileTitle Then
        VBGetFileTitle = ""
    Else
        VBGetFileTitle = Left$(sFileTitle, InStr(sFileTitle, vbNullChar) - 1)
    End If

End Function

' ChooseColor wrapper
Function VBChooseColor(Color As Long, _
                       Optional AnyColor As Boolean = True, _
                       Optional FullOpen As Boolean = False, _
                       Optional DisableFullOpen As Boolean = False, _
                       Optional Owner As Long = -1, _
                       Optional flags As Long) As Boolean

    Dim chclr As TCHOOSECOLOR
    chclr.lStructSize = Len(chclr)
    
    ' Color must get reference variable to receive result
    ' Flags can get reference variable or constant with bit flags
    ' Owner can take handle of owning window
    If Owner <> -1 Then chclr.hWndOwner = Owner

    ' Assign color (default uninitialized value of zero is good default)
    chclr.rgbResult = Color

    ' Mask out unwanted bits
    Dim afMask As Long
    afMask = CLng(Not (CC_ENABLEHOOK Or _
                       CC_ENABLETEMPLATE))
    ' Pass in flags
    chclr.flags = afMask And (CC_RGBInit Or _
                  IIf(AnyColor, CC_AnyColor, CC_SolidColor) Or _
                  (-FullOpen * CC_FullOpen) Or _
                  (-DisableFullOpen * CC_PreventFullOpen))

    ' If first time, initialize to white
    If fNotFirst = False Then InitColors

    chclr.lpCustColors = VarPtr(alCustom(0))
    ' All other fields zero
    
    m_lApiReturn = ChooseColor(chclr)
    Select Case m_lApiReturn
    Case 1
        ' Success
        VBChooseColor = True
        Color = chclr.rgbResult
    Case 0
        ' Cancelled
        VBChooseColor = False
        Color = -1
    Case Else
        ' Extended error
        m_lExtendedError = CommDlgExtendedError()
        VBChooseColor = False
        Color = -1
    End Select

End Function

Private Sub InitColors()
    Dim i As Integer
    ' Initialize with first 16 system interface colors
    For i = 0 To 15
        alCustom(i) = GetSysColor(i)
    Next
    fNotFirst = True
End Sub

' Property to read or modify custom colors (use to save colors in registry)
Public Property Get CustomColor(i As Integer) As Long
    ' If first time, initialize to white
    If fNotFirst = False Then InitColors
    If i >= 0 And i <= 15 Then
        CustomColor = alCustom(i)
    Else
        CustomColor = -1
    End If
End Property

Public Property Let CustomColor(i As Integer, iValue As Long)
    ' If first time, initialize to system colors
    If fNotFirst = False Then InitColors
    If i >= 0 And i <= 15 Then
        alCustom(i) = iValue
    End If
End Property

Private Property Get DevMode() As DevMode
    DevMode = m_dvmode
End Property

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.EXEName & ".CommonDialog"
        Err.Raise COMError(e), sSource, sText
    Else
        ' Raise standard Visual Basic error
        sSource = App.EXEName & ".VBError"
        Err.Raise e, sSource
    End If
End Sub
#End If


Private Sub StrToBytes(ab() As Byte, s As String)
    If IsArrayEmpty(ab) Then
        ' Assign to empty array
        ab = StrConv(s, vbFromUnicode)
    Else
        Dim cab As Long
        ' Copy to existing array, padding or truncating if necessary
        cab = UBound(ab) - LBound(ab) + 1
        If Len(s) < cab Then s = s & String$(cab - Len(s), 0)
        'If UnicodeTypeLib Then
        '    Dim st As String
        '    st = StrConv(s, vbFromUnicode)
        '    CopyMemoryStr ab(LBound(ab)), st, cab
        'Else
            CopyMemoryStr ab(LBound(ab)), s, cab
        'End If
    End If
End Sub


Private Function BytesToStr(ab() As Byte) As String
    BytesToStr = StrConv(ab, vbUnicode)
End Function

Private Function COMError(e As Long) As Long
    COMError = e Or vbObjectError
End Function
'
Private Function IsArrayEmpty(va As Variant) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = va(LBound(va))
    IsArrayEmpty = (Err <> 0)
End Function





