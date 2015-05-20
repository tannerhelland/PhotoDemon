Attribute VB_Name = "Uniscribe_API_Decs"
'***************************************************************************
'Uniscribe API Types
'Copyright 2015-2015 by Tanner Helland
'Created: 14/May/15
'Last updated: 19/May/15
'Last update: refactor the module to only keep types here; functions were split into the new pdUniscribe class.
'
'Relevant MSDN page for all things Uniscribe:
' https://msdn.microsoft.com/en-us/library/windows/desktop/dd374091%28v=vs.85%29.aspx
'
'Many thanks to Michael Kaplan for his endless work in demystifying Windows text handling.  Of particular value to
' this module is his personal blog - http://www.siao2.com/ - which I referenced liberally during the assembly of this
' class.  His blog includes an article with some Uniscribe-related VB declarations from his (now out of print) book
' on VB internationalization, but I am deliberately avoiding linking it as many of the declarations are missing or
' simply incorrect.  I *strongly* recommend referring to MSDN directly if you plan on working with Uniscribe.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'-----------------------------
'   Windows API enumerations
Public Enum GCPCLASS
     GCPCLASS_LATIN = 1
     GCPCLASS_ARABIC = 2
     GCPCLASS_HEBREW = 2
     GCPCLASS_NEUTRAL = 3
     GCPCLASS_LOCALNUMBER = 4
     GCPCLASS_LATINNUMBER = 5
     GCPCLASS_LATINNUMERICTERMINATOR = 6
     GCPCLASS_LATINNUMERICSEPARATOR = 7
     GCPCLASS_NUMERICSEPARATOR = 8
     GCPCLASS_POSTBOUNDRTL = &H10
     GCPCLASS_PREBOUNDLTR = &H40
     GCPCLASS_PREBOUNDRTL = &H80
     GCPCLASS_POSTBOUNDLTR = &H20
     GCPGLYPH_LINKAFTER = &H4000
     GCPGLYPH_LINKBEFORE = &H8000
End Enum

Public Enum HRESULT
    S_FALSE = &H1
    S_OK = &H0
    E_INVALIDARG = &H80070057
    E_OUTOFMEMORY = &H8007000E
    E_PENDING = &H8000000A
    USP_E_SCRIPT_NOT_IN_FONT = &H80040200
End Enum

#If False Then
    Private Const S_FALSE = &H1, S_OK = &H0, E_INVALIDARG = &H80070057, E_OUTOFMEMORY = &H8007000E, E_PENDING = &H8000000A, USP_E_SCRIPT_NOT_IN_FONT = &H80040200
#End If

'-----------------------------
' Uniscribe ENUMs
Public Enum SCRIPT
    SCRIPT_UNDEFINED = 0
End Enum

Public Enum SCRIPT_JUSTIFY
    SCRIPT_JUSTIFY_NONE = 0
    SCRIPT_JUSTIFY_ARABIC_BLANK = 1
    SCRIPT_JUSTIFY_CHARACTER = 2
    SCRIPT_JUSTIFY_RESERVED1 = 3
    SCRIPT_JUSTIFY_BLANK = 4
    SCRIPT_JUSTIFY_RESERVED2 = 5
    SCRIPT_JUSTIFY_RESERVED3 = 6
    SCRIPT_JUSTIFY_ARABIC_NORMAL = 7
    SCRIPT_JUSTIFY_ARABIC_KASHIDA = 8
    SCRIPT_JUSTIFY_ARABIC_ALEF = 9
    SCRIPT_JUSTIFY_ARABIC_HA = 10
    SCRIPT_JUSTIFY_ARABIC_RA = 11
    SCRIPT_JUSTIFY_ARABIC_BA = 12
    SCRIPT_JUSTIFY_ARABIC_BARA = 13
    SCRIPT_JUSTIFY_ARABIC_SEEN = 14
    SCRIPT_JUSTIFY_RESERVED4 = 15
End Enum

Public Enum SSA_FLAGS
    SSA_PASSWORD = &H1            ' Input string contains a single character to be duplicated iLength times
    SSA_TAB = &H2                 ' Expand tabs
    SSA_CLIP = &H4                ' Clip string at iReqWidth
    SSA_FIT = &H8                 ' Justify string to iReqWidth
    SSA_DZWG = &H10               ' Provide representation glyphs for control characters
    SSA_FALLBACK = &H20           ' Use fallback fonts
    SSA_BREAK = &H40              ' Return break flags (character and word stops)
    SSA_GLYPHS = &H80             ' Generate glyphs, positions and attributes
    SSA_RTL = &H100               ' Base embedding level 1
    SSA_GCP = &H200               ' Return missing glyphs and LogCLust with GetCharacterPlacement conventions
    SSA_HOTKEY = &H400            ' Replace '&’ with underline on subsequent codepoint
    SSA_METAFILE = &H800          ' Write items with ExtTextOutW Unicode calls, not glyphs
    SSA_LINK = &H1000             ' Apply FE font linking/association to non-complex text
    SSA_HIDEHOTKEY = &H2000       ' Remove first '&’ from displayed string
    SSA_HOTKEYONLY = &H2400       ' Display underline only.
   
    ' Internal flags
    SSA_PIDX = &H10000000         ' Internal
    SSA_LAYOUTRTL = &H20000000    ' Internal - Used when DC is mirrored
    SSA_DONTGLYPH = &H40000000    ' Internal - Used only by GDI during metafiling - Use ExtTextOutA for positioning
End Enum

Public Enum SCRIPT_IS_COMPLEX_FLAGS
    SIC_COMPLEX = 1      ' Treat complex script letters as complex
    SIC_ASCIIDIGIT = 2   ' Treat digits U+0030 through U+0039 as copmplex
    SIC_NEUTRAL = 4      ' Treat neutrals as complex
End Enum

Public Enum SCRIPT_DIGITSUBSTITUTE_FLAGS
    SCRIPT_DIGITSUBSTITUTE_CONTEXT = 0       ' Substitute to match preceeding letters
    SCRIPT_DIGITSUBSTITUTE_NONE = 1          ' No substitution
    SCRIPT_DIGITSUBSTITUTE_NATIONAL = 2      ' Substitute with official national digits
    SCRIPT_DIGITSUBSTITUTE_TRADITIONAL = 3   ' Substitute with traditional digits of the locale
End Enum

Public Enum SCRIPT_GET_CMAP_FLAGS
    SGCM_RTL = &H1&             ' Return mirrored glyph for mirrorable Unicode codepoints
End Enum

'-----------------------------
'   Uniscribe Types

'-----------------------------
' Windows API types
Public Type ABC
    abcA As Long
    abcB As Long
    abcC As Long
End Type

Public Type RECT_SIZE
    cx As Long
    cy As Long
End Type

' This is the C-friendly version of SCRIPT_DIGITSUBSTITUTE_VB
' which will be packed properly
Public Type SCRIPT_DIGITSUBSTITUTE
    NationalDigitLanguage As Integer
    TraditionalDigitLanguage As Integer
    DigitSubstitute As Byte
    dwReserved As Long
End Type

' This is the C-friendly version of SCRIPT_CONTROL_VB
' which will be packed properly
Public Type SCRIPT_CONTROL
    uDefaultLanguage As Integer
    fBitFields As Byte
    fReserved As Integer
End Type

' This is the C-friendly version of SCRIPT_STATE_VB
' which will be packed properly
Public Type SCRIPT_STATE
    fBitFields1 As Byte
    fBitFields2 As Byte
End Type

' This is the C-friendly version of SCRIPT_VISATTR_VB
' which will be packed properly
Public Type SCRIPT_VISATTR
    uJustification As SCRIPT_JUSTIFY
    fBitFields1 As Byte
    fBitFields2 As Byte
End Type

' This is the C-friendly version of SCRIPT_ANALYSIS_VB
' which will be packed properly
Public Type SCRIPT_ANALYSIS
    fBitFields1 As Byte
    fBitFields2 As Byte
    s As SCRIPT_STATE
End Type

' This is the C-friendly version of SCRIPT_LOGATTR_VB
' which will be packed properly
Public Type SCRIPT_LOGATTR
    fBitFields As Byte
End Type

Public Type SCRIPT_CACHE
    p As Long
End Type

Public Type SCRIPT_FONTPROPERTIES
    cBytes As Long
    wgBlank As Integer
    wgDefault As Integer
    wgInvalid As Integer
    wgKashida As Integer
    iKashidaWidth As Long
End Type

' UNDONE: This structure may not work well
' for using SCRIPT_PROPERTIES because it may
' not be aligned properly. Why oh why did they
' have to use bitfields?
Public Type SCRIPT_PROPERTIES
    langID As Integer
    fBitFields(1 To 3) As Byte
End Type

Public Type SCRIPT_ITEM
    iCharPos As Long
    analysis As SCRIPT_ANALYSIS
End Type

Public Type GOFFSET
    du As Long
    dv As Long
End Type

Public Type SCRIPT_TABDEF
    cTabStops As Long
    iScale As Long
    pTabStops() As Long
    iTabOrigin As Long
End Type

' We do not use this struct since we have to pass it ByVal
' some times and ByRef other times. All it is a pointer to a
' BLOB of data in memory, anyway, so we will use a Long
Public Type SCRIPT_STRING_ANALYSIS
    p As Long
End Type

'-----------------------------
'   VB friendly versions of Uniscribe Types

' You will have to use SCRIPT_CONTROL to call the
' API to make sure the structure is packed properly
Public Type SCRIPT_CONTROL_VB
    uDefaultLanguage As Long    '  :16
    fContextDigits As Byte  ' As Long   :1
    fInvertPreBoundDir As Byte  ' As Long   :1
    fInvertPostBoundDir As Byte ' As Long   :1
    fLinkStringBefore As Byte   ' As Long   :1
    fLinkStringAfter As Byte    ' As Long   :1
    fNeutralOverride As Byte    ' As Long   :1
    fNumericOverride As Byte    ' As Long   :1
    fLegacyBidiClass As Byte    ' As Long   :1
    fReserved As Byte   ' As Long   :8
End Type

' You will have to use SCRIPT_STATE to call the
' API to make sure the structure is packed properly
Public Type SCRIPT_STATE_VB
    uBidiLevel As Integer   ':5
    fOverrideDirection As Integer   ':1
    fInhibitSymSwap As Integer  ':1
    fCharShape As Integer   ':1
    fDigitSubstitute As Integer ':1
    fInhibitLigate As Integer   ':1
    fDisplayZWG As Integer  ':1
    fArabicNumContext As Integer    ':1
    fGcpClusters As Integer ':1
    fReserved As Integer    ':1
    fEngineReserved As Integer  ':2
End Type

' You will have to use SCRIPT_VISATTR to call the
' API to make sure the structure is packed properly
Public Type SCRIPT_VISATTR_VB
       uJustification As SCRIPT_JUSTIFY ':4
       fClusterStart As Integer ':1
       fDiacritic As Integer    ':1
       fZeroWidth As Integer    ':1
       fReserved As Integer ':1
       fShapeReserved As Integer    ':8
End Type

' You will have to use SCRIPT_ANALYSIS to call the
' API to make sure the structure is packed properly
Public Type SCRIPT_ANALYSIS_VB
    eScript As Integer  ':10
    fRTL As Integer ':1
    fLayoutRTL As Integer   ':1
    fLinkBefore As Integer  ':1
    fLinkAfter As Integer   ':1
    fLogicalOrder As Integer    ':1
    fNoGlyphIndex As Integer    ':1
    s As SCRIPT_STATE
End Type

' You will have to use SCRIPT_LOGATTR to call the
' API to make sure the structure is packed properly
Public Type SCRIPT_LOGATTR_VB
    fSoftBreak As Byte  ':1
    fWhiteSpace As Byte ':1
    fCharStop As Byte   ':1
    fWordStop As Byte   ':1
    fInvalid As Byte    ':1
    fReserved As Byte   ':3
End Type

' You will have to use SCRIPT_PROPERTIES to call the
' API to make sure the structure is packed properly
Public Type SCRIPT_PROPERTIES_VB
    langID As Long  ':16
    fNumeric As Long    ':1
    fComplex As Long    ':1
    fNeedsWordBreaking As Long  ':1
    fNeedsCaretInfo As Long ':1
    bCharSet As Long    ':8
    fControl As Long    ':1
    fPrivateUseArea  As Long  ':1
    fNeedsCharacterJustify As Long  ':1
    fInvalidGlyph As Long   ':1
    fInvalidLogAttr As Long ':1
    fCDM As Long    ':1
   
    ' Added in later versions of UNISCRIBE (usp10.h)
    fAmbiguousCharSet As Long   ':1
    fClusterSizeVaries As Long  ':1
    fRejectInvalid As Long  ':1
End Type

'Glyph data used by PhotoDemon.  An array of this custom struct is filled when the caller requests a copy of our internal Uniscribe data.
' A fair amount of work is required to pull data out of the various incredibly complicated Uniscribe returns, so don't request copies of
' this data any more than is absolutely necessary.
Public Type pdGlyphUniscribe
    glyphIndex As Long
    AdvanceWidth As Long
    isZeroWidth As Boolean
    GlyphOffset As GOFFSET
    glyphPath As pdGraphicsPath     'GDI+ GraphicsPath wrapper containing the fully translated font outline.  This value is not filled by
                                    ' this module; only pdGlyphCollection is capable of generating it.
End Type
