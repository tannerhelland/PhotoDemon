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

'——————————-
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

'——————————-
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

'——————————-
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

'——————————-
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
    langid As Integer
    fBitFields(1 To 3) As Byte
End Type

Public Type SCRIPT_ITEM
    iCharPos As Long
    a As SCRIPT_ANALYSIS
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

'——————————-
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
    langid As Long  ':16
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

'——————————-
'   Uniscribe APIs
Private Declare Function ScriptApplyDigitSubstitution Lib "usp10" (psds As SCRIPT_DIGITSUBSTITUTE, psc As SCRIPT_CONTROL, pss As SCRIPT_STATE) As Long
Private Declare Function ScriptApplyLogicalWidth Lib "usp10" (piDx() As Long, ByVal cChars As Long, ByVal cGlyphs As Long, pwLogClust() As Integer, psva As SCRIPT_VISATTR, piAdvance() As Long, pSA As SCRIPT_ANALYSIS, pABC As ABC, piJustify As Long) As Long
Private Declare Function ScriptBreak Lib "usp10" (pwcChars As Long, ByVal cChars As Long, pSA As SCRIPT_ANALYSIS, psla As SCRIPT_LOGATTR) As Long
Private Declare Function ScriptCPtoX Lib "usp10" (ByVal iCP As Long, ByVal fTrailing As Long, ByVal cChars As Long, ByVal cGlyphs As Long, pwLogClust As Integer, psva As SCRIPT_VISATTR, piAdvance As Long, pSA As SCRIPT_ANALYSIS, piX As Long) As Long
Private Declare Function ScriptCacheGetHeight Lib "usp10" (ByVal srcDC As Long, psc As SCRIPT_CACHE, tmHeight As Long) As Long
Private Declare Function ScriptFreeCache Lib "usp10" (psc As SCRIPT_CACHE) As Long
Private Declare Function ScriptGetCMap Lib "usp10" (ByVal srcDC As Long, psc As SCRIPT_CACHE, ByVal pwcInChars As Long, ByVal cChars As Long, ByVal dwFlags As SCRIPT_GET_CMAP_FLAGS, pwOutGlyphs() As Integer) As Long
Private Declare Function ScriptGetFontProperties Lib "usp10" (ByVal srcDC As Long, psc As SCRIPT_CACHE, sfp As SCRIPT_FONTPROPERTIES) As Long
Private Declare Function ScriptGetGlyphABCWidth Lib "usp10" (ByVal srcDC As Long, psc As SCRIPT_CACHE, ByVal wGlyph As Integer, pABC As ABC) As Long
Private Declare Function ScriptGetLogicalWidths Lib "usp10" (pSA As SCRIPT_ANALYSIS, ByVal cChars As Long, ByVal cGlyphs As Long, piGlyphWidth() As Long, pwLogClust() As Integer, psva As SCRIPT_VISATTR, piDx As Long) As Long
Private Declare Function ScriptGetProperties Lib "usp10" (ppSp As SCRIPT_PROPERTIES, piNumScripts As Long) As Long
Private Declare Function ScriptIsComplex Lib "usp10" (ByVal pwcInChars As Long, ByVal cInChars As Long, ByVal dwFlags As SCRIPT_IS_COMPLEX_FLAGS) As Long
Private Declare Function ScriptItemize Lib "usp10" (ByVal pwcInChars As Long, ByVal cInChars As Long, ByVal cMaxItems As Long, psControl As SCRIPT_CONTROL, psState As SCRIPT_STATE, pItems() As SCRIPT_ITEM, pcItems As Long) As Long
Private Declare Function ScriptJustify Lib "usp10" (psva As SCRIPT_VISATTR, piAdvance() As Long, ByVal cGlyphs As Long, ByVal iDx As Long, ByVal iMinKashida As Long, piJustify() As Long) As Long
Private Declare Function ScriptLayout Lib "usp10" (ByVal cRuns As Long, pbLevel() As Byte, piVisualToLogical() As Long, piLogicalToVisual() As Long) As Long
Private Declare Function ScriptPlace Lib "usp10" (ByVal srcDC As Long, psc As SCRIPT_CACHE, pwGlyphs() As Integer, ByVal cGlyphs As Long, psva As SCRIPT_VISATTR, pSA As SCRIPT_ANALYSIS, piAdvance() As Long, pGoffset As GOFFSET, pABC As ABC) As Long
Private Declare Function ScriptRecordDigitSubstitution Lib "usp10" (ByVal Locale As Long, psds As SCRIPT_DIGITSUBSTITUTE) As Long
Private Declare Function ScriptShape Lib "usp10" (ByVal srcDC As Long, psc As SCRIPT_CACHE, ByVal pwcChars As Long, ByVal cChars As Long, ByVal cMaxGlyphs As Long, pas As SCRIPT_ANALYSIS, pwOutGlyphs() As Integer, pwLogClust() As Integer, psva As SCRIPT_VISATTR, pcGlyphs As Long) As Long
Private Declare Function ScriptTextOut Lib "usp10" (ByVal srcDC As Long, psc As SCRIPT_CACHE, ByVal x As Long, ByVal y As Long, ByVal fuOptions As Long, lprc As RECTL, pSA As SCRIPT_ANALYSIS, ByVal pwcReserved As Long, ByVal iReserved As Long, pwGlyphs() As Integer, ByVal cGlyphs As Long, piAdvance() As Long, piJustify As Any, pGoffset As GOFFSET) As Long
Private Declare Function ScriptXtoCP Lib "usp10" (ByVal iX As Long, ByVal cChars As Long, ByVal cGlyphs As Long, pwLogClust() As Integer, psva As SCRIPT_VISATTR, piAdvance() As Long, pSA As SCRIPT_ANALYSIS, piCP As Long, piTrailing As Long) As Long

'——————————-
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
Private Function LeftShift(ByVal lNum As Long, ByVal lShift As Long) As Long
    LeftShift = lNum * (2 ^ lShift)
End Function

'———————-
' RightShift
'
'   Since VB does not have a right shift operator
'   RightShift(8,2) is equivalent to 8 >> 2
'———————-
Private Function RightShift(ByVal lNum As Long, ByVal lShift As Long) As Long
    RightShift = lNum \ (2 ^ lShift)
End Function



