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

'ENUMs
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

#If False Then
    Private Const GCPCLASS_LATIN = 1, GCPCLASS_ARABIC = 2, GCPCLASS_HEBREW = 2, GCPCLASS_NEUTRAL = 3, GCPCLASS_LOCALNUMBER = 4
    Private Const GCPCLASS_LATINNUMBER = 5, GCPCLASS_LATINNUMERICTERMINATOR = 6, GCPCLASS_LATINNUMERICSEPARATOR = 7, GCPCLASS_NUMERICSEPARATOR = 8
    Private Const GCPCLASS_POSTBOUNDRTL = &H10, GCPCLASS_PREBOUNDLTR = &H40, GCPCLASS_PREBOUNDRTL = &H80, GCPCLASS_POSTBOUNDLTR = &H20
    Private Const GCPGLYPH_LINKAFTER = &H4000, GCPGLYPH_LINKBEFORE = &H8000
#End If

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

Public Enum SCRIPT
    SCRIPT_UNDEFINED = 0
End Enum

#If False Then
    Private Const SCRIPT_UNDEFINED = 0
#End If

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

#If False Then
    Private Const SCRIPT_JUSTIFY_NONE = 0, SCRIPT_JUSTIFY_ARABIC_BLANK = 1, SCRIPT_JUSTIFY_CHARACTER = 2, SCRIPT_JUSTIFY_RESERVED1 = 3
    Private Const SCRIPT_JUSTIFY_BLANK = 4, SCRIPT_JUSTIFY_RESERVED2 = 5, SCRIPT_JUSTIFY_RESERVED3 = 6, SCRIPT_JUSTIFY_ARABIC_NORMAL = 7
    Private Const SCRIPT_JUSTIFY_ARABIC_KASHIDA = 8, SCRIPT_JUSTIFY_ARABIC_ALEF = 9, SCRIPT_JUSTIFY_ARABIC_HA = 10, SCRIPT_JUSTIFY_ARABIC_RA = 11
    Private Const SCRIPT_JUSTIFY_ARABIC_BA = 12, SCRIPT_JUSTIFY_ARABIC_BARA = 13, SCRIPT_JUSTIFY_ARABIC_SEEN = 14, SCRIPT_JUSTIFY_RESERVED4 = 15
#End If

Public Enum SCRIPT_IS_COMPLEX_FLAGS
    SIC_COMPLEX = 1      ' Treat complex script letters as complex
    SIC_ASCIIDIGIT = 2   ' Treat digits U+0030 through U+0039 as copmplex
    SIC_NEUTRAL = 4      ' Treat neutrals as complex
End Enum

#If False Then
    Private Const SIC_COMPLEX = 1, SIC_ASCIIDIGIT = 2, SIC_NEUTRAL = 4
#End If

Public Enum SCRIPT_DIGITSUBSTITUTE_FLAGS
    SCRIPT_DIGITSUBSTITUTE_CONTEXT = 0       ' Substitute to match preceeding letters
    SCRIPT_DIGITSUBSTITUTE_NONE = 1          ' No substitution
    SCRIPT_DIGITSUBSTITUTE_NATIONAL = 2      ' Substitute with official national digits
    SCRIPT_DIGITSUBSTITUTE_TRADITIONAL = 3   ' Substitute with traditional digits of the locale
End Enum

#If False Then
    Private Const SCRIPT_DIGITSUBSTITUTE_CONTEXT = 0, SCRIPT_DIGITSUBSTITUTE_NONE = 1, SCRIPT_DIGITSUBSTITUTE_NATIONAL = 2, SCRIPT_DIGITSUBSTITUTE_TRADITIONAL = 3
#End If

Public Enum SCRIPT_GET_CMAP_FLAGS
    SGCM_RTL = &H1&             ' Return mirrored glyph for mirrorable Unicode codepoints
End Enum

#If False Then
    Private Const SGCM_RTL = &H1&             ' Return mirrored glyph for mirrorable Unicode codepoints
#End If

'TYPES
Public Type ABC
    abcA As Long
    abcB As Long
    abcC As Long
End Type

Public Type RECT_SIZE
    cx As Long
    cy As Long
End Type

'For various script types, note the mixed use of Integer and Byte types.  Many of these types use variable-length bitfields
' to report boolean data, and listing those bitfields as bytes, explicitly, prevents us from having to deal with messy handling
' of sign bits and endianness ordering.
Public Type SCRIPT_DIGITSUBSTITUTE
    NationalDigitLanguage As Integer
    TraditionalDigitLanguage As Integer
    DigitSubstitute As Byte
    dwReserved As Long
End Type

Public Type SCRIPT_CONTROL
    uDefaultLanguage As Integer
    fBitFields1 As Byte
    fBitFields2 As Byte
End Type

Public Type SCRIPT_STATE
    fBitFields1 As Byte
    fBitFields2 As Byte
End Type

Public Type SCRIPT_VISATTR
    uJustification As SCRIPT_JUSTIFY
    fBitFields1 As Byte
    fBitFields2 As Byte
End Type

Public Type SCRIPT_ANALYSIS
    fBitFields1 As Byte
    fBitFields2 As Byte
    s As SCRIPT_STATE
End Type

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

'WORD OF WARNING: I haven't made use of this struct yet, but it may not work correctly due to VB's automatic alignment of structs along
' word boundaries.  YOU MUST TEST THIS if using it in an array, with advice to add dummy bytes as padding if alignment issues follow.
Public Type SCRIPT_PROPERTIES
    langID As Integer
    fBitFields1 As Byte
    fBitFields2 As Byte
    fBitFields3 As Byte
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

'Glyph data used by PhotoDemon.  An array of this custom struct is filled when the caller requests a copy of pdUniscribe's internal data.
' A fair amount of work is required to pull data out of the various incredibly complicated Uniscribe structs, so don't request copies of
' this data more than is absolutely necessary.
Public Type pdGlyphUniscribe
    glyphIndex As Long
    AdvanceWidth As Long
    isZeroWidth As Boolean
    GlyphOffset As GOFFSET
    glyphPath As pdGraphicsPath     'GDI+ GraphicsPath wrapper containing the fully translated font outline.  Note that this value is not
                                    ' filled by PD's Uniscribe interface; pdGlyphCollection actually handles that step.
End Type
