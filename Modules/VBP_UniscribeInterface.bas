Attribute VB_Name = "Uniscribe_Interface"
'***************************************************************************
'Uniscribe API Interface
'Copyright 2015-2015 by Tanner Helland
'Created: 14/May/15
'Last updated: 14/May/15
'Last update: start initial build
'
'Relevant MSDN page: https://msdn.microsoft.com/en-us/library/windows/desktop/dd374091%28v=vs.85%29.aspx
'
'Many thanks to Michael Kaplan for his endless work in demystifying Windows text handling.  Of particular value to
' this module is this link from his personal blog: http://www.siao2.com/2006/06/12/628714.aspx
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

'-----------------------------
'   Uniscribe APIs
Private Declare Function ScriptApplyDigitSubstitution Lib "usp10" (psds As SCRIPT_DIGITSUBSTITUTE, psc As SCRIPT_CONTROL, pss As SCRIPT_STATE) As Long
Private Declare Function ScriptApplyLogicalWidth Lib "usp10" (piDx() As Long, ByVal cChars As Long, ByVal cGlyphs As Long, pwLogClust() As Integer, psva As SCRIPT_VISATTR, piAdvance() As Long, psa As SCRIPT_ANALYSIS, pABC As ABC, piJustify As Long) As Long
Private Declare Function ScriptBreak Lib "usp10" (pwcChars As Long, ByVal cChars As Long, psa As SCRIPT_ANALYSIS, psla As SCRIPT_LOGATTR) As Long
Private Declare Function ScriptCPtoX Lib "usp10" (ByVal iCP As Long, ByVal fTrailing As Long, ByVal cChars As Long, ByVal cGlyphs As Long, pwLogClust As Integer, psva As SCRIPT_VISATTR, piAdvance As Long, psa As SCRIPT_ANALYSIS, piX As Long) As Long
Private Declare Function ScriptCacheGetHeight Lib "usp10" (ByVal srcDC As Long, psc As SCRIPT_CACHE, tmHeight As Long) As Long
Private Declare Function ScriptFreeCache Lib "usp10" (psc As SCRIPT_CACHE) As Long
Private Declare Function ScriptGetCMap Lib "usp10" (ByVal srcDC As Long, psc As SCRIPT_CACHE, ByVal pwcInChars As Long, ByVal cChars As Long, ByVal dwFlags As SCRIPT_GET_CMAP_FLAGS, pwOutGlyphs() As Integer) As Long
Private Declare Function ScriptGetFontProperties Lib "usp10" (ByVal srcDC As Long, psc As SCRIPT_CACHE, sfp As SCRIPT_FONTPROPERTIES) As Long
Private Declare Function ScriptGetGlyphABCWidth Lib "usp10" (ByVal srcDC As Long, psc As SCRIPT_CACHE, ByVal wGlyph As Integer, pABC As ABC) As Long
Private Declare Function ScriptGetLogicalWidths Lib "usp10" (psa As SCRIPT_ANALYSIS, ByVal cChars As Long, ByVal cGlyphs As Long, piGlyphWidth() As Long, pwLogClust() As Integer, psva As SCRIPT_VISATTR, piDx As Long) As Long
Private Declare Function ScriptGetProperties Lib "usp10" (ppSp As SCRIPT_PROPERTIES, piNumScripts As Long) As Long
Private Declare Function ScriptIsComplex Lib "usp10" (ByVal pwcInChars As Long, ByVal cInChars As Long, ByVal dwFlags As SCRIPT_IS_COMPLEX_FLAGS) As Long
Private Declare Function ScriptItemize Lib "usp10" (ByVal pwcInChars As Long, ByVal cInChars As Long, ByVal cMaxItems As Long, ByVal ptrToScriptControl As Long, ByVal ptrToScriptState As Long, ByVal ptrToPItems As Long, ByRef pcItems As Long) As Long
Private Declare Function ScriptJustify Lib "usp10" (psva As SCRIPT_VISATTR, piAdvance() As Long, ByVal cGlyphs As Long, ByVal iDx As Long, ByVal iMinKashida As Long, piJustify() As Long) As Long
Private Declare Function ScriptLayout Lib "usp10" (ByVal cRuns As Long, ByVal ptrToPBLevel As Long, ByVal ptrToPIVisualToLogical As Long, ByVal ptrToPILogicalToVisual As Long) As Long
Private Declare Function ScriptPlace Lib "usp10" (ByVal srcDC As Long, ByRef psc As SCRIPT_CACHE, ByVal ptrToIntpwGlyphs As Long, ByVal cGlyphs As Long, ByVal ptrToSVpsva As Long, ByRef psa As SCRIPT_ANALYSIS, ByVal ptrToLngpiAdvance As Long, ByVal ptrTopGoffset As Long, ByRef pABC As ABC) As Long
Private Declare Function ScriptRecordDigitSubstitution Lib "usp10" (ByVal Locale As Long, psds As SCRIPT_DIGITSUBSTITUTE) As Long
Private Declare Function ScriptShape Lib "usp10" (ByVal srcDC As Long, ByVal ptrToSCpsc As Long, ByVal pwcChars As Long, ByVal cChars As Long, ByVal cMaxGlyphs As Long, ByVal ptrToSApas As Long, ByVal ptrToIntpwOutGlyphs As Long, ByVal ptrToIntpwLogClust As Long, ByVal ptrToSVpsva As Long, ByRef pcGlyphs As Long) As Long
Private Declare Function ScriptTextOut Lib "usp10" (ByVal srcDC As Long, psc As SCRIPT_CACHE, ByVal x As Long, ByVal y As Long, ByVal fuOptions As Long, lprc As RECTL, psa As SCRIPT_ANALYSIS, ByVal pwcReserved As Long, ByVal iReserved As Long, pwGlyphs() As Integer, ByVal cGlyphs As Long, piAdvance() As Long, piJustify As Any, pGoffset As GOFFSET) As Long
Private Declare Function ScriptXtoCP Lib "usp10" (ByVal iX As Long, ByVal cChars As Long, ByVal cGlyphs As Long, pwLogClust() As Integer, psva As SCRIPT_VISATTR, piAdvance() As Long, psa As SCRIPT_ANALYSIS, piCP As Long, piTrailing As Long) As Long

'-----------------------------
'   Uniscribe Script* APIs
Private Declare Function ScriptStringAnalyse Lib "usp10" (ByVal srcDC As Long, ByVal pString As Long, ByVal cString As Long, ByVal cGlyphs As Long, ByVal iCharset As Long, ByVal dwFlags As SSA_FLAGS, ByVal iReqWidth As Long, ByRef psControl As Any, ByRef psState As Any, ByRef piDx As Long, ByRef pTabdef As Any, ByRef pbInClass As GCPCLASS, ByRef pssa As Long) As Long
Private Declare Function ScriptStringCPtoX Lib "usp10" (ByVal ssa As Long, ByVal iCP As Long, ByVal fTrailing As Long, pX As Long) As Long
Private Declare Function ScriptStringFree Lib "usp10" (ByRef pssa As Long) As Long
Private Declare Function ScriptStringGetLogicalWidths Lib "usp10" (ByVal ssa As Long, piDx() As Long) As Long
Private Declare Function ScriptStringGetOrder Lib "usp10" (ByVal ssa As Long, puOrder As Long) As Long
Private Declare Function ScriptStringOut Lib "usp10" (ByVal ssa As Long, ByVal iX As Long, ByVal iY As Long, ByVal uOptions As Long, prc As RECTL, ByVal iMinSel As Long, ByVal iMaxSel As Long, ByVal fDisabled As Long) As Long
Private Declare Function ScriptString_pcOutChars Lib "usp10" (ByVal ssa As Long) As Long
Private Declare Function ScriptString_pLogAttr Lib "usp10" (ByVal ssa As Long) As Long
Private Declare Function ScriptString_pSize Lib "usp10" (ByVal ssa As Long) As Long
Private Declare Function ScriptStringValidate Lib "usp10" (ByVal ssa As Long) As Long
Private Declare Function ScriptStringXtoCP Lib "usp10" (ByVal ssa As Long, ByVal iX As Long, picH As Long, piTrailing As Long) As Long

'Internal storage stuff written by Tanner (specifically for PD)

'Glyph data used by PhotoDemon.  An array of this custom struct is filled when the caller requests a copy of our internal Uniscribe data.
' A fair amount of work is required to pull data out of the various incredibly complicated Uniscribe returns, so don't request copies of
' this data any more than is absolutely necessary.
Public Type pdGlyphUniscribe
    GlyphIndex As Long
    AdvanceWidth As Long
    isZeroWidth As Boolean
    GlyphOffset As GOFFSET
End Type

'Current string associated with our cache; if this doesn't change, we can skip most Uniscribe processing.
Private m_CurrentString As String

'SCRIPT_ITEM cache generated by Step 1
Private m_ScriptItemCacheOK As Boolean
Private m_ScriptItemsCache() As SCRIPT_ITEM

'Visual order of the runs in m_ScriptItemsCache()
Private m_VisualToLogicalOrder() As Long

'Logical order of the runs in m_ScriptItemsCache()
Private m_LogicalToVisualOrder() As Long

'Opaque SCRIPT_CACHE handle.  This is allocated by the system on the initial call to ScriptShape.  We must free this handle
' when we are done with it.
Private m_ScriptCache As SCRIPT_CACHE

'Glyph cache.  This is generated by step 3, and is crucial for rendering as it contains the actual glyph indices from the current font!
Private m_GlyphCache() As Integer
Private m_NumOfGlyphs As Long

'Logical cluster cache.  This is generated by step 3, and per MSDN, "the value of each element is the offset from the first glyph in the
' run to the first glyph in the cluster containing the corresponding character."  Basically, this provides a mapping from character to
' glyph.  As such, one logical cluster entry exists for each character (not glyph!)
Private m_LogicalClusterCache() As Integer

'Glyph visual attributes cache.  This is generated by step 3, and one visual attribute entry is present for each glyph (not char!)
Private m_VisualAttributesCache() As SCRIPT_VISATTR

'Advance width cache.  This is generated by step 4, and one advance width is present for each glyph (not char!).  These values are the
' offsets, in pixels, from one glyph to the next.
Private m_AdvanceWidthCache() As Long

'Glyph offset cache.  This is generated by step 4, and one offset is present for each glyph (not char!).  This struct is only filled
' for combining glyphs - for example, for an "A" with an accent "`" over it, some fonts may only provide a plain "A" glyph, and a plain
' "`" accent glyph.  This offset tells us where to move the "`" accent glyph over the A, and the same "`" accent over a different glyph
' may use a different offset (e.g. for a lower-case "a").
Private m_GlyphOffsetCache() As GOFFSET




'———————
'   Wrappers around several Uniscribe functions that allow slightly
'   more friendly VB interaction
'
' ScriptStringFreeC
' ScriptString_pcOutCharsC
' ScriptString_pSizeC
' ScriptString_pLogAttrC
' ScriptStringAnalyseC
' ScriptStringCPtoXC
' ScriptStringXtoCPC
'
' ScriptIsComplex
'———————
Public Function ScriptStringFreeC(ssa As Long) As Long
    
    If ssa <> 0 Then
        ScriptStringFreeC = ScriptStringFree(ssa)
        ssa = 0&
    End If
    
End Function

Public Function ScriptString_pcOutCharsC(ssa As Long) As Long
    
    Dim pcch As Long
    pcch = ScriptString_pcOutChars(ssa)
    If pcch <> 0 Then
        CopyMemory ScriptString_pcOutCharsC, ByVal pcch, Len(pcch)
    End If
    
End Function

Public Function ScriptString_pSizeC(ssa As Long) As RECT_SIZE
    
    Dim psiz As Long
    psiz = ScriptString_pSize(ssa)
    If psiz <> 0 Then
        CopyMemory ScriptString_pSizeC, ByVal psiz, Len(ScriptString_pSizeC)
    End If
    
End Function

Public Sub ScriptString_pLogAttrC(ssa As Long, cch As Long, rgsla() As SCRIPT_LOGATTR_VB)

    Dim prgtsla As Long
    Dim rgtsla() As SCRIPT_LOGATTR
    Dim itsla As Long
    Dim byt As Byte
   
    ' Call Uniscribe to get the LogAttr info
    prgtsla = ScriptString_pLogAttr(ssa)
   
    If prgtsla <> 0 Then
        ' Success! Lets put the pointer into a struct and prepare some memory
        ReDim rgtsla(0 To cch - 1)
        CopyMemory rgtsla(0), ByVal prgtsla, CLng(Len(rgtsla(0)) * cch)
        ReDim rgsla(0 To cch - 1)
       
        ' Convert the unfriendly C type into a friendly VB type that can be used elsewhere
        For itsla = 0 To cch - 1
            byt = rgtsla(itsla).fBitFields
            With rgsla(itsla)
                .fSoftBreak = RightShift((byt And &H1), 0)
                .fWhiteSpace = RightShift((byt And &H2), 1)
                .fCharStop = RightShift((byt And &H4), 2)
                .fWordStop = RightShift((byt And &H8), 3)
                .fInvalid = RightShift((byt And &H10), 4)
                .fReserved = RightShift((byt And &HE0), 5) ' &HE0 = (2 ^ 5 Or 2 ^ 6 Or 2 ^ 7)
            End With
        Next itsla
        Erase rgtsla
    End If
    
End Sub

Public Function ScriptStringAnalyseC(srcDC As Long, stAnalyse As String, cch As Long, ByVal dwFlags As SSA_FLAGS, iReqWidth As Long, Optional vSCV As Variant, Optional vSSV As Variant, Optional vST As Variant) As Long

    Dim ssa As Long
    Dim sc As SCRIPT_CONTROL
    Dim ss As SCRIPT_STATE
    Dim st As SCRIPT_TABDEF
    
    If Not IsMissing(vSCV) Then
        sc.uDefaultLanguage = vSCV.uDefaultLanguage
        sc.fBitFields = _
                            LeftShift(vSCV.fContextDigits, 0) Or _
                            LeftShift(vSCV.fInvertPreBoundDir, 1) Or _
                            LeftShift(vSCV.fInvertPostBoundDir, 2) Or _
                            LeftShift(vSCV.fLinkStringBefore, 3) Or _
                            LeftShift(vSCV.fLinkStringAfter, 4) Or _
                            LeftShift(vSCV.fNeutralOverride, 5) Or _
                            LeftShift(vSCV.fNumericOverride, 6) Or _
                            LeftShift(vSCV.fLegacyBidiClass, 7)
    End If
   
    If Not IsMissing(vSSV) Then
        ss.fBitFields1 = _
                            LeftShift(vSSV.uBidiLevel, 4) Or _
                            LeftShift(vSSV.fOverrideDirection, 5) Or _
                            LeftShift(vSSV.fInhibitSymSwap, 6) Or _
                            LeftShift(vSSV.fCharShape, 7)
        ss.fBitFields2 = _
                            LeftShift(vSSV.fDigitSubstitute, 0) Or _
                            LeftShift(vSSV.fInhibitLigate, 1) Or _
                            LeftShift(vSSV.fDisplayZWG, 2) Or _
                            LeftShift(vSSV.fArabicNumContext, 3) Or _
                            LeftShift(vSSV.fGcpClusters, 4)
    End If
   
    If Not IsMissing(vST) And ((dwFlags And SSA_TAB) = SSA_TAB) Then
        st.cTabStops = vST.cTabStops
        st.iScale = vST.iScale
        st.pTabStops = vST.pTabStops
        st.iTabOrigin = vST.iTabOrigin
    End If
   
    If ScriptStringAnalyse(srcDC, StrPtr(stAnalyse), cch, 0, -1, dwFlags, iReqWidth, sc, ss, ByVal 0&, st, ByVal 0&, ssa) = S_OK Then
        ScriptStringAnalyseC = ssa
    End If
    
End Function

Public Function ScriptStringCPtoXC(ssa As Long, iCP As Long, fTrailing As Long) As Long
    
    Dim pX As Long
    If ScriptStringCPtoX(ssa, iCP, fTrailing, pX) = S_OK Then
        ScriptStringCPtoXC = pX
    End If
    
End Function

Public Function ScriptStringXtoCPC(ssa As Long, ByVal iX As Long, piTrailing As Long) As Long
    
    Dim picH As Long
    If ScriptStringXtoCP(ssa, iX, picH, piTrailing) = S_OK Then
        ScriptStringXtoCPC = picH
    End If
    
End Function

Public Function ScriptIsComplexC(stIn As String, Optional Flags As SCRIPT_IS_COMPLEX_FLAGS) As Boolean
    
    Dim hr As Long
   
    hr = ScriptIsComplex(StrPtr(stIn), Len(stIn), Flags)
    If hr = S_OK Then
        ScriptIsComplexC = True
    ElseIf hr = S_FALSE Then
        ScriptIsComplexC = False
    Else
        Err.Raise hr
    End If
    
End Function

Public Function ScriptRecordDigitSubstitutionC(Locale As Long) As SCRIPT_DIGITSUBSTITUTE
    
    Dim psds As SCRIPT_DIGITSUBSTITUTE

    If ScriptRecordDigitSubstitution(Locale, psds) = S_OK Then
        ScriptRecordDigitSubstitutionC = psds
    End If
    
End Function

'———————
' IchNext / IchPrev
'
'   Takes a SCRIPT_STRING_ANALYSIS and a character position and
'   returns the next or previous character position or word position, taking
'   Uniscribe "clusters" into account
'———————
Public Function IchNext(ssa As Long, ByVal ich As Long, fWord As Boolean) As Long

    Dim cch As Long
    Dim rgsla() As SCRIPT_LOGATTR_VB
    
    cch = ScriptString_pcOutCharsC(ssa)
    Call ScriptString_pLogAttrC(ssa, cch, rgsla())
    
    Do Until ich >= cch - 1
        ich = ich + 1
        If (rgsla(ich).fCharStop And Not fWord) Then Exit Do    ' We are at the end of a character
        If (rgsla(ich).fWordStop And fWord) Then Exit Do    ' We are at the end of a word
    Loop
    
    If ich > cch - 1 Then ich = cch ' Take care of the boundary cases
    IchNext = ich
    
End Function

Public Function IchPrev(ssa As Long, ByVal ich As Long, fWord As Boolean) As Long

    Dim cch As Long
    Dim rgsla() As SCRIPT_LOGATTR_VB
    
    If ich > 0 Then ' Make sure we are at the beginning of the string already
        cch = ScriptString_pcOutCharsC(ssa)
        Call ScriptString_pLogAttrC(ssa, cch, rgsla())
        Do Until ich <= 0
            If (rgsla(ich).fCharStop And Not fWord) Then Exit Do    ' We are at the end of a character
            If (rgsla(ich).fWordStop And fWord) Then Exit Do    ' We are at the end of a word
            ich = ich - 1
        Loop
    End If
    
    If ich < 0 Then ich = 0 ' Take care of the boundary cases
    IchPrev = ich
    
End Function

'———————
' IchBreakSpot
'
'   Find the appropriate place to break for this line. Here
'   is the algorithm used:
'
'   1) If all text will fit or no line breaking is specified, then output the whole string
'   2) If #1 is not true, find the first hard break within the text that could fit on the line
'   3) If #2 could not be found, then look for the last softbreak or whitespace within the text that could fit on the line.
'   4) If #3 is a whitespace, then break AFTER the character
'   5) If #3 is a soft break, than break before the character
'———————
Public Function IchBreakSpot(st As String, rgsla() As SCRIPT_LOGATTR_VB, cch As Long, Optional fNoLineBreaks As Boolean = False) As Long

    Dim ich As Long
   
    ' First check for a hard break
    ich = InStr(1, st, vbCrLf, vbBinaryCompare) - 1
    
    If ich >= 0 And ich <= cch - 1 Then
        ' Use the hard break that was found
        IchBreakSpot = ich
        
    ElseIf Len(st) > cch Then
        ' There are more characters then there is space to output, on this line
        ' at least. So walk the string backwards, looking for a break character.
        For ich = cch - 1 To 0 Step -1
            With rgsla(ich)
                ' Check to see if its a soft break char or a white space char
                If .fWhiteSpace Or .fSoftBreak Then
                    If .fWhiteSpace Then
                        ' White space means break AFTER this character
                        IchBreakSpot = ich
                    ElseIf ich > 0 Then
                        ' Its a softbreak. If we have the characters to spare,
                        ' subtract one because we should be breaking BEFORE
                        ' the character, not AFTER.
                        IchBreakSpot = ich - 1
                    Else
                        ' There are not enough chars to go after. This probably should
                        ' never happen, but we may as well make sure.
                        IchBreakSpot = 0
                    End If
                    Exit For
                End If
            End With
        Next ich
    End If
   
    ' Assume cch is where it's at if it has never been set
    If IchBreakSpot = 0 Then IchBreakSpot = cch
    
End Function

'———————
' UniscribeExtTextOutW
'
'   The Uniscribe-aware version of ExtTextOutW
'———————
Public Function UniscribeExtTextOutW(ByRef srcDC As Long, wOptions As Long, lpRect As RECTL, ByVal st As String, Optional x1 As Long = 0, Optional x2 As Long = 0) As Long

    On Error Resume Next
    Dim ssa As Long
    Dim xWidth As Long
    Dim cch As Long
    Dim ichBreak As Long
    Dim siz As RECT_SIZE
    Dim rgsla() As SCRIPT_LOGATTR_VB
    Dim rct As RECTL
   
    ' deep copy the rect since may be modifying it
    rct.Left = lpRect.Left
    rct.Right = lpRect.Right
    rct.Top = lpRect.Top
    rct.Bottom = lpRect.Bottom
   
    xWidth = rct.Right - rct.Left
   
    ' Keep going till all of the string is done
    Do Until Len(st) = 0
        ssa = ScriptStringAnalyseC(srcDC, st, Len(st), SSA_GLYPHS Or SSA_FALLBACK Or SSA_CLIP Or SSA_BREAK, xWidth)
        If ssa <> 0 Then
            cch = ScriptString_pcOutCharsC(ssa)
            Call ScriptString_pLogAttrC(ssa, cch, rgsla())
       
            ' Get the appropriate break point for this line (see comments in
            ' IchBreakSpot for a better understanding of "appropriate"
            ' CONSIDER: MULTILINE: To support multiple lines, the fNoLineBreaks flag
            ' below would have to be set to False. The rest of the function depends on it!
            ichBreak = IchBreakSpot(st, rgsla(), cch, True)

            ' Free up the analysis, we need to do it again with the new break info
            Call ScriptStringFreeC(ssa)
       
            ' reanalyze the string
            ssa = ScriptStringAnalyseC(srcDC, st, ichBreak, SSA_GLYPHS Or SSA_FALLBACK Or SSA_CLIP Or SSA_BREAK, xWidth)
            If ssa <> 0 Then
                siz = ScriptString_pSizeC(ssa)
                cch = ScriptString_pcOutCharsC(ssa)
               
                ' Output the string, now that we have done all the preparation
                Call ScriptStringOut(ssa, rct.Left, rct.Top, wOptions, rct, x1, x2, 0&)
               
                ' Remove the portion of the string that has been output and adjust the rect
                ' for the next line
                st = Mid$(st, cch + 1)
                rct.Top = rct.Top + siz.cy
            End If
            ' Free up the analysis, we need to (so we can do the next one)!
            Call ScriptStringFreeC(ssa)
        End If
    Loop
End Function

'———————-
' LeftShift
'
'   Since VB does not have a left shift operator
'   LeftShift(8,2) is equivalent to 8 << 2
'———————-
Private Function LeftShift(ByVal lNum As Long, ByVal LShift As Long) As Long
    LeftShift = lNum * (2 ^ LShift)
End Function

'———————-
' RightShift
'
'   Since VB does not have a right shift operator
'   RightShift(8,2) is equivalent to 8 >> 2
'———————-
Private Function RightShift(ByVal lNum As Long, ByVal LShift As Long) As Long
    RightShift = lNum \ (2 ^ LShift)
End Function

'START FUNCTIONS WRITTEN BY TANNER

'Set an arbitrary 1-bit position [0-31] in a Long-type value to either 1 or 0.  Position 0 is the LEAST-SIGNIFICANT BIT, and Position 31
' is the SIGN BIT for a standard VB Long.
'
' Inputs:
'  1) position of the flag, which must be in the range [0, 31]
'  2) value of the flag, TRUE for 1, FALSE for 0
'  3) The Long-type value where you want the flag placed
Private Sub setBitFlag(ByVal flagPosition As Long, ByVal flagValue As Boolean, ByRef copyOfDstLong As Long)

    If (flagPosition >= 0) And (flagPosition <= 31) Then
    
        'Create a bitmask, using flagPosition to determine bit offset
        Dim longMask As Long
        If flagPosition < 31 Then longMask = 2 ^ flagPosition Else longMask = &H80000000
        
        'Write a 1 flag
        If flagValue Then
        
            'Blend the flag with the source Long using OR, which will force the target bit to 1 while preserving the rest of the Long's contents
            copyOfDstLong = copyOfDstLong Or longMask
        
        'Write a 0 flag
        Else
        
            'Blend the INVERSE flag with the source Long, which will force the target bit to 0 while preserving the rest of the Long's contents
            copyOfDstLong = copyOfDstLong And Not longMask
        
        End If
        
    Else
        Debug.Print "C'mon - you know Longs only have 32 bits.  Flag positions must be on the range [0, 31]!"
    End If

End Sub

'Retrieve an arbitrary 1-bit position [0-31] in a Long-type value.  Position 0 is the LEAST-SIGNIFICANT BIT, and Position 31
' is the SIGN BIT for a standard VB Long.
'
' Inputs:
'  1) position of the flag, which must be in the range [0, 31]
'  2) The Long from which you want the bit retrieved
Private Function getBitFlag(ByVal flagPosition As Long, ByVal srcLong As Long) As Boolean

    If (flagPosition >= 0) And (flagPosition <= 31) Then
    
        'Create a bitmask, using flagPosition to determine bit offset
        Dim longMask As Long
        If flagPosition < 31 Then longMask = 2 ^ flagPosition Else longMask = &H80000000
        
        getBitFlag = ((srcLong And longMask) <> 0)
        
    Else
        Debug.Print "C'mon - you know Longs only have 32 bits.  Flag positions must be on the range [0, 31]!"
    End If

End Function

'Free any/all previously allocated cache structures
Public Function freeUniscribeCaches()
    ScriptFreeCache m_ScriptCache
End Function

'Step1_ScriptItemize is the first step in processing a Uniscribe string.  It generates the crucial SCRIPT_ITEM array used for pretty much
' every subsequent Uniscribe interaction.  This array is cached internally, in m_ScriptItemsCache().
'
'Returns TRUE if successful; FALSE otherwise.
Public Function Step1_ScriptItemize(ByRef srcString As String) As Boolean
    
    m_ScriptItemCacheOK = False
    
    'Uniscribe does not handle 0-length strings, so make sure at least one character is present
    If Len(srcString) < 1 Then srcString = " "
    
    'Make a deep copy of the source string
    m_CurrentString = srcString
    
    'Values we need to manually set:
    ' pwcInChars [In]: Pointer to a Unicode string to itemize.
    ' cInChars [In]: Number of characters in pwcInChars to itemize.
    ' cMaxItems [In]: Maximum number of SCRIPT_ITEM structures defining items to process.
    ' psControl [in, optional]: Pointer to a SCRIPT_CONTROL structure indicating the type of itemization to perform.
    '                           Alternatively, the application can set this parameter to NULL if no SCRIPT_CONTROL properties are needed.
    ' psState [in, optional]: Pointer to a SCRIPT_STATE structure indicating the initial bidirectional algorithm state.
    '                         Alternatively, the application can set this parameter to NULL if the script state is not needed.
    ' pItems [out]: Pointer to a buffer in which the function retrieves SCRIPT_ITEM structures representing the items that have been processed.
    '               The buffer should be (cMaxItems + 1) * sizeof(SCRIPT_ITEM) bytes in length. It is invalid to call this function with a
    '               buffer to hold less than two SCRIPT_ITEM structures. The function always adds a terminal item to the item analysis array
    '               so that the length of the item with zero-based index "i" is always available as:
    '               pItems[i+1].iCharPos - pItems[i].iCharPos;
    ' pcItems [out]: Pointer to the number of SCRIPT_ITEM structures processed.
    
    'Determine the maximum number of SCRIPT_ITEM structures to process.  This is an arbitrary value; MSDN says "The function returns E_OUTOFMEMORY
    ' if the value of cMaxItems is insufficient. As in all error cases, no items are fully processed and no part of the output array contains
    ' defined values. If the function returns E_OUTOFMEMORY, the application can call it again with a larger pItems buffer."
    '
    'Because PD doesn't work with particularly large strings (e.g. we're not MS Office), we err on the side of safety and use
    ' a very large buffer.  Note that the buffer itself must be one larger than the number of SCRIPT_ITEM values passed, and it
    ' can never be less than two.
    Dim numScriptItems As Long
    numScriptItems = Len(srcString) * 2
    
    ReDim m_ScriptItemsCache(0 To numScriptItems) As SCRIPT_ITEM
    
    'SCRIPT_CONTROL and SCRIPT_STATE primarily deal with extremely technical details of localization and glyph runs.  They are not relevant
    ' to our usage in PD.  For more details, see:
    ' SCRIPT_CONTROL: https://msdn.microsoft.com/en-us/library/windows/desktop/dd368800%28v=vs.85%29.aspx
    ' SCRIPT_STATE: https://msdn.microsoft.com/en-us/library/windows/desktop/dd374043%28v=vs.85%29.aspx
    '
    '(That said, if complaints ever arise in the future, we can revisit using tips from http://www.catch22.net/tuts/uniscribe-mysteries)
    Dim dummyScriptControl As SCRIPT_CONTROL
    Dim dummyScriptState As SCRIPT_STATE
    
    Dim numItemsFilled As Long
    
    Dim unsReturn As HRESULT
    unsReturn = ScriptItemize(StrPtr(srcString), Len(srcString), numScriptItems, VarPtr(dummyScriptControl), VarPtr(dummyScriptState), VarPtr(m_ScriptItemsCache(0)), numItemsFilled)
    
    'Account for potential out of memory errors
    Do While unsReturn = E_OUTOFMEMORY
        
        'Double the allotted cache size and try again
        numScriptItems = numScriptItems * 2
        ReDim m_ScriptItemsCache(0 To numScriptItems - 1) As SCRIPT_ITEM
        unsReturn = ScriptItemize(StrPtr(srcString), Len(srcString), numScriptItems, ByVal 0&, ByVal 0&, VarPtr(m_ScriptItemsCache(0)), numItemsFilled)
        
    Loop
    
    'If ScriptItemize still failed, it was not due to memory errors
    If unsReturn = S_OK Then
        
        Debug.Print "Uniscribe itemized the string correctly.  " & numItemsFilled & " SCRIPT_ITEM structs were filled."
        
        'Trim our cache to its relevant size, and note that we MUST leave a spare entry at the end!
        ReDim Preserve m_ScriptItemsCache(0 To numItemsFilled) As SCRIPT_ITEM
        Step1_ScriptItemize = True
        
    Else
        Debug.Print "WARNING!  ScriptItemize failed with code " & unsReturn & ".  Please investigate."
        Step1_ScriptItemize = False
    End If
    
    'Cache the success/failure value, so subsequent calls don't error out
    m_ScriptItemCacheOK = Step1_ScriptItemize

End Function

'Step2_ScriptLayout is the second step in processing a Uniscribe string.  (Actually, there could be another step before this, where we
' manually break the results of Step1 into more fine-grained runs, accounting for differences in font.  PD doesn't support this right now
' so I don't provide a wrapper for that.)  Step 2 takes the "runs" generated by Step 1, and converts them from logical input order to
' visual output order.  This is crucial when intermixing LTR and RTL text, as the order in which characters are entered may be totally
' different from the order in which they are displayed.
'
'Note that this function takes no inputs; it relies entirely on the output of Step 1 for its behavior.
'
'Returns TRUE if successful; FALSE otherwise.  (A failure at step 1 that was not dealt with by the caller will cause this function to
' return FALSE, as it relies on the output of Step 1 to generate a correct layout.)
Public Function Step2_ScriptLayout() As Boolean
    
    'If a SCRIPT_ITEM cache does not exist, exit now
    If Not m_ScriptItemCacheOK Then
        Debug.Print "WARNING!  Step 1 failed, so Step 2 cannot proceed!"
        Step2_ScriptLayout = False
        Exit Function
    End If
    
    'Values we need to manually set:
    ' cRuns [In]: Number of runs to process.
    ' pbLevel [In]: Pointer to an array, of length indicated by cRuns, containing run embedding levels. Embedding levels for all runs on
    '               the line must be included, ordered logically.
    ' piVisualToLogical [out, optional]: Pointer to an array, of length indicated by cRuns, in which this function retrieves the run
    ' embedding levels reordered to visual order. The first array element represents the run to display at the far left, and subsequent
    ' entries should be displayed progressing from left to right. The function sets this parameter to NULL if there is no output.
    ' piLogicalToVisual [out, optional]: Pointer to an array, of length indicated by cRuns, in which this function retrieves the visual
    ' run positions. The first array element is the relative visual position where the first logical run should be displayed, the leftmost
    ' display position being 0. The function sets this parameter to NULL if there is no output.
    
    'The number of runs to process is simply the length of the SCRIPT_ITEM cache from step 1
    Dim numOfRuns As Long
    numOfRuns = UBound(m_ScriptItemsCache) + 1
    
    'Run embedding levels were automatically calculated by Step 1.  These values are buried deep within the item cache, unfortunately,
    ' so we need to retrieve them and store them in their own array.
    '
    'Note also that this structure must be DWORD-aligned, to avoid access issues
    Dim elBound As Long
    elBound = ((numOfRuns * 4) + 3) \ 4
    
    Dim embeddingLevels() As Byte
    ReDim embeddingLevels(0 To elBound - 1) As Byte
    
    Dim tmpPBLevel As Long, extractedPBLevel As Long, tmpFlag As Boolean
    
    Dim i As Long
    For i = 0 To numOfRuns - 1
        
        'Retrieving the embedded bidi level is ugly, to put it kindly.  The top 5 bits of the embedded SCRIPT_STATE value for each run
        ' contain LTR vs RTL information.  There is no good way to retrieve this data in VB, so we are lazy and just do it manually.
        '
        'Start by copying the byte into a Long, which makes retrieval easier.
        tmpPBLevel = m_ScriptItemsCache(i).analysis.s.fBitFields1
        
        'Technically the bitfield contains 5 bits for potential values on the range [0, 31].  But there are really only values on the
        ' range [0, 3], basically every combination of LTR and RTL text embedded within each other.  As such, we only need to retrieve
        ' the bottom 2 of 5 bytes.
        tmpFlag = getBitFlag(3, tmpPBLevel)
        If tmpFlag Then extractedPBLevel = 1 Else extractedPBLevel = 0
        
        tmpFlag = getBitFlag(4, tmpPBLevel)
        If tmpFlag Then extractedPBLevel = extractedPBLevel + 2
        
        'Store the calculated value in our embeddingLevels() result array
        embeddingLevels(i) = extractedPBLevel
        
    Next i
    
    'With our bidi levels set, we can now determine the visual order of runs
    
    'Prep conversion arrays
    ReDim m_VisualToLogicalOrder(0 To numOfRuns - 1) As Long
    ReDim m_LogicalToVisualOrder(0 To numOfRuns - 1) As Long
    
    Dim ptrToVLO As Long, ptrToLVO As Long
    ptrToVLO = VarPtr(m_VisualToLogicalOrder(0))
    ptrToLVO = VarPtr(m_LogicalToVisualOrder(0))
    
    'Retrieve the final layout order of the runs
    Dim unsReturn As HRESULT
    unsReturn = ScriptLayout(numOfRuns, VarPtr(embeddingLevels(0)), ptrToVLO, ptrToLVO)
    
    'unsReturn = ScriptLayout(numOfRuns, VarPtr(embeddingLevels(0)), ByVal 0&, ByVal 0&)
    
    If unsReturn = S_OK Then
        
        Debug.Print "Uniscribe laid out the string correctly."
        Step2_ScriptLayout = True
        
    Else
        Debug.Print "WARNING!  ScriptLayout failed with code " & unsReturn & ".  Please investigate."
        Step2_ScriptLayout = False
    End If
    
End Function

'Step3_ScriptShape is the third step in processing a Uniscribe string (and arguably the most important!).  ScriptShape is the far more
' powerful Uniscribe version of GDI's GetCharacterPlacement function.  Basically, it operates on a single run, and does all the messy
' behind-the-scene work to convert character values into glyph indices of the current font.  As such, it is the first step to require
' a DC, and you must (obviously) have the relevant font selected into the DC *prior* to calling this!
'
'NOTE: in the future, I would very much like to have two codepaths here: one that uses the old ScriptShape function, and a new one that
' uses ScriptShapeOpenType.  This would allow us to take full advantage of OpenType's many awesome features - but right now, I mostly
' just want to get the damn thing working.
'
'At present, this only operates on the first run returned by ScriptItemize.  Some messy work is required to deal with multiple runs;
' I'll tackle that after I have single runs working.
'
'Returns TRUE if successful; FALSE otherwise.  (A failure at step 1 or 2 that was not dealt with by the caller will cause this function to
' return FALSE, as it relies on the output of Steps 1 and 2 to generate correct shapes.)
Public Function Step3_ScriptShape(ByRef srcDC As Long) As Boolean
    
    'If a SCRIPT_ITEM cache does not exist, exit now
    If Not m_ScriptItemCacheOK Then
        Debug.Print "WARNING!  Step 1 or 2 failed, so Step 3 cannot proceed!"
        Step3_ScriptShape = False
        Exit Function
    End If
    
    'Values we must generate prior to calling the API:
    ' hDC [In]: Handle to the device context. For more information, see Caching.
    ' psc [in, out]: Pointer to a SCRIPT_CACHE structure identifying the script cache.
    ' pwcChars [In]: Pointer to an array of Unicode characters defining the run.
    ' cChars [In]: Number of characters in the Unicode run.
    ' cMaxGlyphs [In]: Maximum number of glyphs to generate, and the length of pwOutGlyphs. A reasonable value is (1.5 * cChars + 16),
    '                  but this value might be insufficient in some circumstances.
    ' psa [in, out]: Pointer to the SCRIPT_ANALYSIS structure for the run, containing the results from an earlier call to ScriptItemize.
    ' pwOutGlyphs [out]: Pointer to a buffer in which this function retrieves an array of glyphs with size as indicated by cMaxGlyphs.
    ' pwLogClust [out]: Pointer to a buffer in which this function retrieves an array of logical cluster information. Each array element
    '                   corresponds to a character in the array of Unicode characters; therefore this array has the number of elements
    '                   indicated by cChars. The value of each element is the offset from the first glyph in the run to the first glyph
    '                   in the cluster containing the corresponding character. Note that, when the fRTL member is set to TRUE in the
    '                   SCRIPT_ANALYSIS structure, the elements decrease as the array is read.
    ' psva [out]: Pointer to a buffer in which this function retrieves an array of SCRIPT_VISATTR structures containing visual attribute
    '             information. Since each glyph has only one visual attribute, this array has the number of elements indicated by cMaxGlyphs.
    ' pcGlyphs [out]: Pointer to the location in which this function retrieves the number of glyphs indicated in pwOutGlyphs.

    'Determine an intial size for the glyph cache.  We use the MSDN recommended approach for calculations.
    Dim cMaxGlyphs As Long
    cMaxGlyphs = 1.5 * CDbl(Len(m_CurrentString)) + 16
    
    'Prep the logical cluster cache, which has the same length as cChars
    ReDim m_LogicalClusterCache(0 To Len(m_CurrentString) - 1) As Integer
    
    Dim unsReturn As HRESULT
    
    Do
        
        'Prep all glyph-specific caches to be the same size as cMaxGlyphs
        ReDim m_GlyphCache(0 To cMaxGlyphs - 1) As Integer
        ReDim m_VisualAttributesCache(0 To cMaxGlyphs - 1) As SCRIPT_VISATTR
    
        unsReturn = ScriptShape(srcDC, VarPtr(m_ScriptCache), StrPtr(m_CurrentString), Len(m_CurrentString), cMaxGlyphs, VarPtr(m_ScriptItemsCache(0).analysis), VarPtr(m_GlyphCache(0)), VarPtr(m_LogicalClusterCache(0)), VarPtr(m_VisualAttributesCache(0)), m_NumOfGlyphs)
        
        'Because cMaxGlyphs initialization is a guessing game, ScriptShape will return an out of memory code if the buffer proves insufficient.
        ' If this happens, we must increase the size of the buffer and try again.
        If unsReturn = E_OUTOFMEMORY Then
            cMaxGlyphs = cMaxGlyphs * 2
        End If
        
    Loop While unsReturn = E_OUTOFMEMORY
    
    'm_NumOfGlyphs now contains the number of glyphs generated by ScriptShape.  Trim the glyph-specific caches to match.
    If m_NumOfGlyphs > 1 Then
        ReDim Preserve m_GlyphCache(0 To m_NumOfGlyphs - 1) As Integer
        ReDim Preserve m_VisualAttributesCache(0 To m_NumOfGlyphs - 1) As SCRIPT_VISATTR
    'Else
    '    ReDim Preserve m_GlyphCache(0 To 1) As Integer
    '    ReDim Preserve m_VisualAttributesCache(0 To 1) As SCRIPT_VISATTR
    End If
    
    'Return success/failure
    If unsReturn = S_OK Then
        
        Debug.Print "Uniscribe shaped the string correctly."
        Step3_ScriptShape = True
        
    Else
        Debug.Print "WARNING!  ScriptShape failed with code " & unsReturn & ".  Please investigate."
        Step3_ScriptShape = False
    End If
    
End Function

'Step4_ScriptPlace is the fourth step in processing a Uniscribe string.  It could technically be called Step 3.5, as it operates
' directly on the results of ScriptShape; in fact, I may merge the two functions in the future.
'
'While ScriptShape converts characters to glyphs, ScriptPlace calculates positioning for all glyphs.  Once again, the choice of
' font is crucial, so this step also requires a DC (nd you must have the relevant font selected into the DC *prior* to calling this!)
'
'NOTE: in the future, I would very much like to have two codepaths here: one that uses the old ScriptPlace function, and a new one that
' uses ScriptPlaceOpenType.  This would allow us to take full advantage of OpenType's many awesome features - but right now, I mostly
' just want to get the damn thing working.
'
'At present, this only operates on the first run returned by ScriptItemize.  Some messy work is required to deal with multiple runs;
' I'll tackle that after I have single runs working.
'
'Returns TRUE if successful; FALSE otherwise.  (A failure at step 1 or 2 that was not dealt with by the caller will cause this function to
' return FALSE, as it relies on the output of Steps 1 and 2 to generate correct shapes.)
Public Function Step4_ScriptPlace(ByRef srcDC As Long) As Boolean

    'If a SCRIPT_ITEM cache does not exist, exit now
    If Not m_ScriptItemCacheOK Then
        Debug.Print "WARNING!  Step 1 or 2 failed, so Step 4 cannot proceed!"
        Step4_ScriptPlace = False
        Exit Function
    End If
    
    'Values we must generate prior to calling the API:
    ' hDC [In]: Handle to the device context. For more information, see Caching.
    ' psc [in, out]: Pointer to a SCRIPT_CACHE structure identifying the script cache.
    ' pwGlyphs [In]: Pointer to a glyph buffer obtained from an earlier call to the ScriptShape function.
    ' cGlyphs [In]: Count of glyphs in the glyph buffer.
    ' psva [In]: Pointer to an array of SCRIPT_VISATTR structures indicating visual attributes.
    ' psa [in, out]: Pointer to a SCRIPT_ANALYSIS structure. On input, this structure is obtained from a previous call to ScriptItemize.
    '                On output, this structure contains values retrieved by ScriptPlace.
    ' piAdvance [out]: Pointer to an array in which this function retrieves advance width information.
    ' pGoffset [out]: Optional. Pointer to an array of GOFFSET structures in which this function retrieves the x and y offsets of
    '                 combining glyphs. This array must be of length indicated by cGlyphs.
    ' pABC [out]: Pointer to an ABC structure in which this function retrieves the ABC width for the *entire run*.
    
    'Prep the glyph offset cache
    ReDim m_GlyphOffsetCache(0 To m_NumOfGlyphs - 1) As GOFFSET
    
    'Prep the advance width cache
    ReDim m_AdvanceWidthCache(0 To m_NumOfGlyphs - 1) As Long
    
    'Prep a single ABC struct, to store the ABC width for this entire run.
    Dim tmpABC As ABC
    
    Dim unsReturn As HRESULT
    unsReturn = ScriptPlace(srcDC, m_ScriptCache, VarPtr(m_GlyphCache(0)), m_NumOfGlyphs, VarPtr(m_VisualAttributesCache(0)), m_ScriptItemsCache(0).analysis, VarPtr(m_AdvanceWidthCache(0)), VarPtr(m_GlyphOffsetCache(0)), tmpABC)
    
    'Return success/failure
    If unsReturn = S_OK Then
        
        Debug.Print "Uniscribe placed the string correctly."
        Step4_ScriptPlace = True
        
    Else
        Debug.Print "WARNING!  ScriptPlace failed with code " & unsReturn & ".  Please investigate."
        Step4_ScriptPlace = False
    End If

End Function

'Retrieve a copy of the currently calculated glyph cache, in a custom PD format.
'
' Returns a value >= 0 indicating the number of glyphs in the returned struct.  0 indicated failure or an empty string.
Public Function getCopyOfGlyphCache(ByRef dstGlyphArray() As pdGlyphUniscribe) As Long
    
    ReDim dstGlyphArray(0 To m_NumOfGlyphs - 1) As pdGlyphUniscribe
    
    Dim i As Long
    For i = 0 To m_NumOfGlyphs - 1
    
        With dstGlyphArray(i)
            .GlyphIndex = m_GlyphCache(i)
            .AdvanceWidth = m_AdvanceWidthCache(i)
            .GlyphOffset = m_GlyphOffsetCache(i)
            .isZeroWidth = False
        End With
    
    Next i
    
    getCopyOfGlyphCache = m_NumOfGlyphs
    
End Function

Public Sub printDebugInfo()
    
    Dim tmpGlyphCache() As pdGlyphUniscribe
    Dim numOfGlyphs As Long
    numOfGlyphs = getCopyOfGlyphCache(tmpGlyphCache)
    
    Debug.Print "-- Glyph data returned by Uniscribe --"
    
    If numOfGlyphs > 0 Then
    
        Dim i As Long
        For i = 0 To numOfGlyphs - 1
            With tmpGlyphCache(i)
                Debug.Print i & ": " & .GlyphIndex & " (" & .AdvanceWidth & ")  (" & .GlyphOffset.du & ", " & .GlyphOffset.dv & ")"
            End With
        Next i
        
    End If
    
    Debug.Print "-- End of glyph data  --"
    
End Sub
