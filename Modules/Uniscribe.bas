Attribute VB_Name = "Uniscribe"
'***************************************************************************
'Uniscribe API Types
'Copyright 2015-2026 by Tanner Helland
'Created: 14/May/15
'Last updated: 01/June/16
'Last update: construct a new dynamic cache when retrieving script properties; this greatly improves performance for
'             the font dropdowns, as they need script info to generate idealized font previews
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
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Some consts are unused at present, but they are tedious to translate into VB6 format,
' so I've left them in case they prove useful in the future.
'Public Enum GCPCLASS
'    GCPCLASS_LATIN = 1
'    GCPCLASS_ARABIC = 2
'    GCPCLASS_HEBREW = 2
'    GCPCLASS_NEUTRAL = 3
'    GCPCLASS_LOCALNUMBER = 4
'    GCPCLASS_LATINNUMBER = 5
'    GCPCLASS_LATINNUMERICTERMINATOR = 6
'    GCPCLASS_LATINNUMERICSEPARATOR = 7
'    GCPCLASS_NUMERICSEPARATOR = 8
'    GCPCLASS_POSTBOUNDRTL = &H10
'    GCPCLASS_PREBOUNDLTR = &H40
'    GCPCLASS_PREBOUNDRTL = &H80
'    GCPCLASS_POSTBOUNDLTR = &H20
'    GCPGLYPH_LINKAFTER = &H4000
'    GCPGLYPH_LINKBEFORE = &H8000&
'End Enum
'
'#If False Then
'    Private Const GCPCLASS_LATIN = 1, GCPCLASS_ARABIC = 2, GCPCLASS_HEBREW = 2, GCPCLASS_NEUTRAL = 3, GCPCLASS_LOCALNUMBER = 4
'    Private Const GCPCLASS_LATINNUMBER = 5, GCPCLASS_LATINNUMERICTERMINATOR = 6, GCPCLASS_LATINNUMERICSEPARATOR = 7, GCPCLASS_NUMERICSEPARATOR = 8
'    Private Const GCPCLASS_POSTBOUNDRTL = &H10, GCPCLASS_PREBOUNDLTR = &H40, GCPCLASS_PREBOUNDRTL = &H80, GCPCLASS_POSTBOUNDLTR = &H20
'    Private Const GCPGLYPH_LINKAFTER = &H4000, GCPGLYPH_LINKBEFORE = &H8000&
'#End If

Public Enum hResult
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

'Public Enum SCRIPT_JUSTIFY
'    SCRIPT_JUSTIFY_NONE = 0
'    SCRIPT_JUSTIFY_ARABIC_BLANK = 1
'    SCRIPT_JUSTIFY_CHARACTER = 2
'    SCRIPT_JUSTIFY_RESERVED1 = 3
'    SCRIPT_JUSTIFY_BLANK = 4
'    SCRIPT_JUSTIFY_RESERVED2 = 5
'    SCRIPT_JUSTIFY_RESERVED3 = 6
'    SCRIPT_JUSTIFY_ARABIC_NORMAL = 7
'    SCRIPT_JUSTIFY_ARABIC_KASHIDA = 8
'    SCRIPT_JUSTIFY_ARABIC_ALEF = 9
'    SCRIPT_JUSTIFY_ARABIC_HA = 10
'    SCRIPT_JUSTIFY_ARABIC_RA = 11
'    SCRIPT_JUSTIFY_ARABIC_BA = 12
'    SCRIPT_JUSTIFY_ARABIC_BARA = 13
'    SCRIPT_JUSTIFY_ARABIC_SEEN = 14
'    SCRIPT_JUSTIFY_RESERVED4 = 15
'End Enum
'
'#If False Then
'    Private Const SCRIPT_JUSTIFY_NONE = 0, SCRIPT_JUSTIFY_ARABIC_BLANK = 1, SCRIPT_JUSTIFY_CHARACTER = 2, SCRIPT_JUSTIFY_RESERVED1 = 3
'    Private Const SCRIPT_JUSTIFY_BLANK = 4, SCRIPT_JUSTIFY_RESERVED2 = 5, SCRIPT_JUSTIFY_RESERVED3 = 6, SCRIPT_JUSTIFY_ARABIC_NORMAL = 7
'    Private Const SCRIPT_JUSTIFY_ARABIC_KASHIDA = 8, SCRIPT_JUSTIFY_ARABIC_ALEF = 9, SCRIPT_JUSTIFY_ARABIC_HA = 10, SCRIPT_JUSTIFY_ARABIC_RA = 11
'    Private Const SCRIPT_JUSTIFY_ARABIC_BA = 12, SCRIPT_JUSTIFY_ARABIC_BARA = 13, SCRIPT_JUSTIFY_ARABIC_SEEN = 14, SCRIPT_JUSTIFY_RESERVED4 = 15
'#End If

'Public Enum SCRIPT_IS_COMPLEX_FLAGS
'    SIC_COMPLEX = 1      ' Treat complex script letters as complex
'    SIC_ASCIIDIGIT = 2   ' Treat digits U+0030 through U+0039 as complex
'    SIC_NEUTRAL = 4      ' Treat neutrals as complex
'End Enum
'
'#If False Then
'    Private Const SIC_COMPLEX = 1, SIC_ASCIIDIGIT = 2, SIC_NEUTRAL = 4
'#End If

'Public Enum SCRIPT_DIGITSUBSTITUTE_FLAGS
'    SCRIPT_DIGITSUBSTITUTE_CONTEXT = 0       ' Substitute to match preceeding letters
'    SCRIPT_DIGITSUBSTITUTE_NONE = 1          ' No substitution
'    SCRIPT_DIGITSUBSTITUTE_NATIONAL = 2      ' Substitute with official national digits
'    SCRIPT_DIGITSUBSTITUTE_TRADITIONAL = 3   ' Substitute with traditional digits of the locale
'End Enum
'
'#If False Then
'    Private Const SCRIPT_DIGITSUBSTITUTE_CONTEXT = 0, SCRIPT_DIGITSUBSTITUTE_NONE = 1, SCRIPT_DIGITSUBSTITUTE_NATIONAL = 2, SCRIPT_DIGITSUBSTITUTE_TRADITIONAL = 3
'#End If

'Public Enum SCRIPT_GET_CMAP_FLAGS
'    SGCM_RTL = &H1&             ' Return mirrored glyph for mirrorable Unicode codepoints
'End Enum
'
'#If False Then
'    Private Const SGCM_RTL = &H1&             ' Return mirrored glyph for mirrorable Unicode codepoints
'#End If

'TYPES
Public Type ABC
    abcA As Long
    abcB As Long
    abcC As Long
End Type

'For various script types, note the mixed use of Integer and Byte types.  Many of these types use variable-length bitfields
' to report boolean data, and listing those bitfields as bytes, explicitly, prevents us from having to deal with messy handling
' of sign bits and endianness ordering.
'Public Type SCRIPT_DIGITSUBSTITUTE
'    NationalDigitLanguage As Integer
'    TraditionalDigitLanguage As Integer
'    DigitSubstitute As Byte
'    dwReserved As Long
'End Type

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
    fBitFields1 As Byte
    fBitFieldReserved As Byte
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

'WORD OF WARNING: I haven't made use of this struct yet, but it may not work correctly due to
' VB's automatic alignment of structs along word boundaries.  YOU MUST TEST THIS if using it in
' an array, with advice to add dummy bytes as padding if alignment issues follow.
'Public Type SCRIPT_PROPERTIES
'    langID As Integer
'    fBitFields1 As Byte
'    fBitFields2 As Byte
'    fBitFields3 As Byte
'End Type

Public Type SCRIPT_ITEM
    iCharPos As Long
    analysis As SCRIPT_ANALYSIS
End Type

Public Type GOFFSET
    du As Long
    dv As Long
End Type

'Public Type SCRIPT_TABDEF
'    cTabStops As Long
'    iScale As Long
'    pTabStops() As Long
'    iTabOrigin As Long
'End Type

'Glyph data used by PhotoDemon.  An array of this custom struct is filled when the caller requests a copy of pdUniscribe's internal data.
' A fair amount of work is required to pull data out of the various incredibly complicated Uniscribe structs, so don't request copies of
' this data more than is absolutely necessary.
Public Type PDGlyphUniscribe
    glyphIndex As Long
    glyphOffset As GOFFSET
    
    advanceWidth As Long
    
    glyphPath As pd2DPath     'GDI+ GraphicsPath wrapper containing the fully translated font outline.  Note that this value is not
                                    ' filled by PD's Uniscribe interface; pdGlyphCollection actually handles that step.
    isSoftBreak As Boolean
    isWhiteSpace As Boolean
    isZeroWidth As Boolean
    isHardLineBreak As Boolean
    
    abcWidth As ABC
    
    finalX As Single        'Final (x, y) positioning has nothing to do with Uniscribe.  PD calculates this internally, using all the
    finalY As Single        ' metrics supplied by Uniscribe, and all the metrics supplied by the user.
    lineID As Long          'Which line (0-based) this glyph sits on.  PD marks this during the positioning loop to simplify rendering.
    
    isFirstGlyphOnLine As Boolean   'These two values could be inferred from LineID, but it's faster to simply mark them during
    isLastGlyphOnLine As Boolean    ' processing.  We use these markers to account for ABC overhang on leading or trailing chars.
    
End Type

'When retrieving OpenType tags, it's convenient to reduce the unsigned Longs into a 4-byte struct.
'Private Type pdOpenTypeTag
'    b1 As Byte
'    b2 As Byte
'    b3 As Byte
'    b4 As Byte
'End Type

'Some Uniscribe API declarations are script-independent.  These are nice for retrieving generic font or glyph information, so we declare
' them here (instead of inside the pdUniscribe class).
Private Declare Function ScriptGetFontScriptTags Lib "usp10" (ByVal srcDC As Long, ByVal ptrToScriptCache As Long, ByVal ptrToScriptAnalysis As Long, ByVal cMaxTags As Long, ByVal ptrToScriptTagsArray As Long, ByVal ptrToNumOfTags As Long) As Long
Private Declare Function ScriptFreeCache Lib "usp10" (psc As SCRIPT_CACHE) As Long

'When checking font script tags, we only need to declare an array once, then simply clear it and reuse it.
Private m_ScriptCachesAreReady As Boolean
Private m_ScriptTags() As Long
Private Const OPENTYPE_MAX_NUM_SCRIPT_TAGS As Long = 114

'Retrieving supported scripts is a time- and energy-intensive process.  We can alleviate a lot of the bottlenecks by
' caching supported scripts after they've been retrieved.
Private Const INITIAL_SCRIPT_CACHE_SIZE As Long = 16
Private m_ScriptFontNames() As String
Private m_ScriptCache() As PD_FONT_PROPERTY
Private m_NumOfScripts As Long

'Valid OpenType Script Tag values.  This list is *not* comprehensive!
' A full list is available here: https://www.microsoft.com/typography/otspec/scripttags.htm
Public Enum OPENTYPE_SCRIPT_TAG
    OST_ARABIC = &H62617261
    OST_CJK_IDEOGRAPHIC = &H696E6168
    OST_CYRILLIC = &H6C727963
    OST_HEBREW = &H72626568
    OST_LATIN = &H6E74616C
    OST_THAI = &H69616874
End Enum

'Given a font name, retrieve all scripts explicitly supported.  If no scripts are supported, the destination property struct
' will be marked with SCRIPT_UNKNOWN.
'
'Returns: value >=0 identifying how many scripts are supported.  Also, dstFontProperty will be filled accordingly.
Public Function GetScriptsSupportedByFont(ByVal srcFontName As String, ByRef dstFontProperty As PD_FONT_PROPERTY) As Long
    
    If OS.IsVistaOrLater Then
        
        Dim i As Long
        
        'Before manually retrieving script information, see if the properties for this font have already been retrieved.
        If (m_NumOfScripts > 0) Then
            
            For i = 0 To m_NumOfScripts - 1
                If Strings.StringsEqual(srcFontName, m_ScriptFontNames(i), False) Then
                    dstFontProperty = m_ScriptCache(i)
                    GetScriptsSupportedByFont = dstFontProperty.numSupportedScripts
                    Exit Function
                End If
            Next i
            
        Else
            If (Not VBHacks.IsArrayInitialized(m_ScriptFontNames)) Then
                ReDim m_ScriptFontNames(0 To INITIAL_SCRIPT_CACHE_SIZE - 1) As String
                ReDim m_ScriptCache(0 To INITIAL_SCRIPT_CACHE_SIZE - 1) As PD_FONT_PROPERTY
            End If
        End If
        
        'Create a dummy font handle matching the current name
        Dim tmpFont As Long, tmpDC As Long
        If Fonts.QuickCreateFontAndDC(srcFontName, tmpFont, tmpDC) Then
            
            'As of May 2015, OpenType only supports 114 tags, so a font can't return more values than this!
            ' (We actually size it to 114 + 1, just in case Uniscribe gets picky about having a little extra breathing room.)
            If (Not m_ScriptCachesAreReady) Then
                ReDim m_ScriptTags(0 To OPENTYPE_MAX_NUM_SCRIPT_TAGS) As Long
                m_ScriptCachesAreReady = True
            Else
                FillMemory VarPtr(m_ScriptTags(0)), OPENTYPE_MAX_NUM_SCRIPT_TAGS * 4, 0
            End If
            
            'Unfortunately, the nature of this function means it's impossible to take advantage of Uniscribe's internal
            ' caching mechanisms.  As such, we basically have to submit a bunch of blank cache structs.
            Dim tmpCache As SCRIPT_CACHE
            Dim numTagsReceived As Long
            
            'Retrieve a list of supported scripts
            Dim unsReturn As hResult
            unsReturn = ScriptGetFontScriptTags(tmpDC, VarPtr(tmpCache), 0&, OPENTYPE_MAX_NUM_SCRIPT_TAGS, VarPtr(m_ScriptTags(0)), VarPtr(numTagsReceived))
            
            'Success!  Copy a list of relevant parameters into the destination font property object
            If (unsReturn = S_OK) Then
                
                'Resize the target array as necessary
                If (numTagsReceived > 0) Then
                    
                    'Mark supported scripts as KNOWN
                    dstFontProperty.ScriptsKnown = True
                    
                    'Copy all tags into the destination object
                    dstFontProperty.numSupportedScripts = numTagsReceived
                    
                    'NOTE: if you want, you can also preserve the full list of supported scripts.  PD doesn't do this at present,
                    ' because it greatly increases the copying time required for structs (as the list of potential scripts is
                    ' 100+ items).  To use this capability, you must also go into the Fonts module and uncomment
                    ' the .SupportedScripts array inside the PD_FONT_PROPERTY definition.
                    'ReDim dstFontProperty.SupportedScripts(0 To numTagsReceived - 1) As Long
                    'CopyMemoryStrict VarPtr(dstFontProperty.SupportedScripts(0)), VarPtr(m_ScriptTags(0)), numTagsReceived * 4
                    
                    'Mark a few known tags in advance, as it's helpful to have immediate access to these.
                    For i = 0 To numTagsReceived - 1
                        
                        Select Case m_ScriptTags(i)
                        
                            Case OST_ARABIC
                                dstFontProperty.Supports_Arabic = True
                            
                            Case OST_CJK_IDEOGRAPHIC
                                dstFontProperty.Supports_CJK = True
                                
                            Case OST_CYRILLIC
                                dstFontProperty.Supports_Cyrillic = True
                                
                            Case OST_HEBREW
                                dstFontProperty.Supports_Hebrew = True
                                
                            Case OST_LATIN
                                dstFontProperty.Supports_Latin = True
                            
                            Case OST_THAI
                                dstFontProperty.Supports_Thai = True
                        
                        End Select
                        
                    Next i
                    
                    'Return the number of script tags received for this object
                    GetScriptsSupportedByFont = numTagsReceived
                    
                    'Alternatively, if you're curious, you can dump a list of supported script names to the debug window
                    'If srcFontName = "Cambria Math" Then
                    '
                    '    Dim tmpString As String
                    '    Dim tmpTag As pdOpenTypeTag
                    '    For i = 0 To numTagsReceived - 1
                    '        CopyMemoryStrict VarPtr(tmpTag), VarPtr(m_ScriptTags(i)), 4
                    '        With tmpTag
                    '            tmpString = ChrW$(.byte1) & ChrW$(.byte2) & ChrW$(.byte3) & ChrW$(.byte4)
                    '        End With
                    '        Debug.Print tmpString
                    '    Next i
                    '
                    'End If
                    
                'If no tags were received, mark this script as "undefined".  We generally assume Latin character support only
                ' for such a font.
                Else
                    dstFontProperty.ScriptsKnown = False
                End If
                
            Else
                PDDebug.LogAction "WARNING!  Couldn't retrieve supported script list for " & srcFontName
            End If
            
            'Remember to free our temporary font and DC when we're done with them
            Fonts.QuickDeleteFontAndDC tmpFont, tmpDC
            
            'Also, let Uniscribe know we're done with our copy of their cache
            ScriptFreeCache tmpCache
            
            'Even in failure, we want to cache this script so that we don't have to test it again in the future.
            If (m_NumOfScripts > UBound(m_ScriptFontNames)) Then
                ReDim Preserve m_ScriptFontNames(0 To m_NumOfScripts * 2 - 1) As String
                ReDim Preserve m_ScriptCache(0 To m_NumOfScripts * 2 - 1) As PD_FONT_PROPERTY
            End If
            
            m_ScriptFontNames(m_NumOfScripts) = srcFontName
            m_ScriptCache(m_NumOfScripts) = dstFontProperty
            m_NumOfScripts = m_NumOfScripts + 1
        
        Else
            GetScriptsSupportedByFont = 0
        End If
        
    End If
    
End Function
