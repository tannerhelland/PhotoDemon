Attribute VB_Name = "Font_Management"
'***************************************************************************
'PhotoDemon Font Manager
'Copyright 2013-2015 by Tanner Helland
'Created: 31/May/13
'Last updated: 26/April/15
'Last update: start splitting relevant bits from pdFont into this separate manager module.  pdFont still exists for
'              GDI font rendering purposes.
'
'For many years, PhotoDemon has used the pdFont class for GDI text rendering.  Unfortunately, that class was designed before I
' knew much about GDI font management, and it has some sketchy design considerations that make it a poor fit for PD's text tool.
'
'As part of a broader overhaul to PD's text management, this new Font_Management module has been created.  Its job is to manage a
' font cache for this system, which individual font classes can then query for things like font existence, style, and more.
'
'Obviously, this class relies heavily on WAPI.  Functions are documented to the best of my knowledge and ability.
'
'Dependencies: font section of PD's Public_Enums_And_Types module
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'****************************************************************************************
'Note: these types are used in the callback function for EnumFontFamiliesEx; as such, I have to declare them as public.

Public Const LF_FACESIZEW = 64

Public Type LOGFONTW
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
    lfFaceName(0 To LF_FACESIZEW - 1) As Byte
End Type

Public Type NEWTEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
    ntmFlags As Long
    ntmSizeEM As Long
    ntmCellHeight As Long
    ntmAveWidth As Long
End Type

'END ENUMFONTFAMILESEX ENUMS
'****************************************************************************************

'Retrieve specific metrics on a font (in our case, crucial for aligning button images against the font baseline and ascender)
Private Declare Function GetCharABCWidthsFloat Lib "gdi32" Alias "GetCharABCWidthsFloatW" (ByVal hDC As Long, ByVal firstCharCodePoint As Long, ByVal secondCharCodePoint As Long, ByVal ptrToABCFloatArray As Long) As Long
Public Type ABCFLOAT
    abcfA As Single
    abcfB As Single
    abcfC As Single
End Type

Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsW" (ByVal hDC As Long, ByRef lpMetrics As TEXTMETRIC) As Long
Public Declare Function GetOutlineTextMetrics Lib "gdi32" Alias "GetOutlineTextMetricsW" (ByVal hDC As Long, ByVal cbData As Long, ByVal ptrToOTMStruct As Long) As Long

Public Type TEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Integer
    tmLastChar As Integer
    tmDefaultChar As Integer
    tmBreakChar As Integer
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type

Public Type TEXTMETRIC_PADDED_W
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Integer
    tmLastChar As Integer
    tmDefaultChar As Integer
    tmBreakChar As Integer
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
    tmPaddingByte1 As Byte
    tmPaddingByte2 As Byte
    tmPaddingByte3 As Byte
End Type

Public Type PANOSE_PADDED
    bPaddingByte0 As Byte
    bFamilyType As Byte
    bSerifStyle As Byte
    bWeight As Byte
    bProportion As Byte
    bContrast As Byte
    bStrokeVariation As Byte
    bArmStyle As Byte
    bLetterform As Byte
    bMidline As Byte
    bXHeight As Byte
    bPaddingByte1 As Byte
End Type

Public Type OUTLINETEXTMETRIC
    otmSize As Long
    otmTextMetrics As TEXTMETRIC_PADDED_W
    otmPanoseNumber As PANOSE_PADDED
    otmfsSelection As Long
    otmfsType As Long
    otmsCharSlopeRise As Long
    otmsCharSlopeRun As Long
    otmItalicAngle As Long
    otmEMSquare As Long
    otmAscent As Long
    otmDescent As Long
    otmLineGap As Long
    otmsCapEmHeight As Long
    otmsXHeight As Long
    otmrcFontBox As RECTL
    otmMacAscent As Long
    otmMacDescent As Long
    otmMacLineGap As Long
    otmusMinimumPPEM As Long
    otmptSubscriptSize As POINTAPI
    otmptSubscriptOffset As POINTAPI
    otmptSuperscriptSize As POINTAPI
    otmptSuperscriptOffset As POINTAPI
    otmsStrikeoutSize As Long
    otmsStrikeoutPosition As Long
    otmsUnderscoreSize As Long
    otmsUnderscorePosition As Long
    otmpFamilyName As Long
    otmpFaceName As Long
    otmpStyleName As Long
    otmpFullName As Long
End Type

'Possible charsets for the tmCharSet byte, above
Public Enum FONT_CHARSETS
    CS_ANSI = 0
    CS_DEFAULT = 1
    CS_SYMBOL = 2
    CS_MAC = 77
    CS_SHIFTJIS = 128
    CS_HANGEUL = 129
    CS_JOHAB = 130
    CS_GB2312 = 134
    CS_CHINESEBIG5 = 136
    CS_GREEK = 161
    CS_TURKISH = 162
    CS_HEBREW = 177
    CS_ARABIC = 178
    CS_BALTIC = 186
    CS_RUSSIAN = 204
    CS_THAI = 222
    CS_EASTEUROPE = 238
    CS_OEM = 255
End Enum

#If False Then
    Private Const CS_ANSI = 0, CS_DEFAULT = 1, CS_SYMBOL = 2, CS_MAC = 77, CS_SHIFTJIS = 128, CS_HANGEUL = 129, CS_JOHAB = 130
    Private Const CS_GB2312 = 134, CS_CHINESEBIG5 = 136, CS_GREEK = 161, CS_TURKISH = 162, CS_HEBREW = 177, CS_ARABIC = 178
    Private Const CS_BALTIC = 186, CS_RUSSIAN = 204, CS_THAI = 222, CS_EASTEUROPE = 238, CS_OEM = 255
#End If

'Font enumeration types
Private Const LF_FACESIZEA = 32
Private Const DEFAULT_CHARSET = 1

'NOTE: several crucial types for this class are listed in the Public_Enums_And_Types module.

'ntmFlags field flags
Private Const NTM_REGULAR = &H40&
Private Const NTM_BOLD = &H20&
Private Const NTM_ITALIC = &H1&

'tmPitchAndFamily flags
Private Const TMPF_FIXED_PITCH = &H1
Private Const TMPF_VECTOR = &H2
Private Const TMPF_DEVICE = &H8
Private Const TMPF_TRUETYPE = &H4

Private Const ELF_VERSION = 0
Private Const ELF_CULTURE_LATIN = 0

'EnumFonts Masks
Private Const RASTER_FONTTYPE = &H1
Private Const DEVICE_FONTTYPE = &H2
Private Const TRUETYPE_FONTTYPE = &H4

Private Declare Function EnumFontFamiliesEx Lib "gdi32" Alias "EnumFontFamiliesExW" (ByVal hDC As Long, ByRef lpLogFontW As LOGFONTW, ByVal lpEnumFontFamExProc As Long, ByRef lParam As Any, ByVal dwFlags As Long) As Long

'GDI font weight (boldness)
Private Const FW_DONTCARE As Long = 0
Private Const FW_THIN As Long = 100
Private Const FW_EXTRALIGHT As Long = 200
Private Const FW_ULTRALIGHT As Long = 200
Private Const FW_LIGHT As Long = 300
Private Const FW_NORMAL As Long = 400
Private Const FW_REGULAR As Long = 400
Private Const FW_MEDIUM As Long = 500
Private Const FW_SEMIBOLD As Long = 600
Private Const FW_DEMIBOLD As Long = 600
Private Const FW_BOLD As Long = 700
Private Const FW_EXTRABOLD As Long = 800
Private Const FW_ULTRABOLD As Long = 800
Private Const FW_HEAVY As Long = 900
Private Const FW_BLACK As Long = 900

'GDI font quality
Private Const DEFAULT_QUALITY As Long = 0
Private Const DRAFT_QUALITY As Long = 1
Private Const PROOF_QUALITY As Long = 2
Private Const NONANTIALIASED_QUALITY As Long = 3
Private Const ANTIALIASED_QUALITY As Long = 4
Private Const CLEARTYPE_QUALITY As Byte = 5

'GDI font creation and management
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectW" (ByRef lpLogFont As LOGFONTW) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'Various non-font-specific WAPI functions helpful for font assembly
Private Const LogPixelsX = 88
Private Const LOGPIXELSY = 90
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

'Some system-specific font settings are cached at initialization time, and unchanged for the life of the program.
' (TODO: watch for relevant window messages on Win 8.1+ that may change these.)
Private m_LogPixelsX As Long, m_LogPixelsY As Long

'Internal font caches.  PD uses these to populate things like font selection dropdowns.
Private m_PDFontCache As pdStringStack
Private Const INITIAL_PDFONTCACHE_SIZE As Long = 256
Private m_LastFontAdded As String

'Some external functions retrieve specific data from a TextMetrics struct.  We cache our own struct so we can use it
' on such function calls.
Private m_TmpTextMetric As TEXTMETRIC

'This function provides some helper wrappers for selecting fonts into a DC.  Rather than track the previously selected object
' (which will only ever be a stock object), we simply re-select a stock font into the DC prior to deleting the temporary font.
Private Const SYSTEM_FONT As Long = 13
Private Declare Function GetStockObject Lib "gdi32" (ByVal fnObject As Long) As Long

'PD's internal font property object and collection.  This is generated by the BuildFontCacheProperties function, below.
Public Type PD_FONT_PROPERTY
    ScriptsKnown As Boolean
    Supports_Arabic As Boolean
    Supports_CJK As Boolean
    Supports_Cyrillic As Boolean
    Supports_Hebrew As Boolean
    Supports_Latin As Boolean
    Supports_Thai As Boolean
    numSupportedScripts As Byte
    SupportedScripts() As Long
End Type

Public g_PDFontProperties() As PD_FONT_PROPERTY

Private m_Unicode As pdUnicode

'PD paints pretty much all of its own text.  Rather than burden each individual control with maintaining their own font object,
' we maintain a cache of the interface font at all requested sizes.  If an object needs to draw interface text, they can query
' us for a matching font object.
Private m_ProgramFontCollection As pdFontCollection

'Want to draw program text onto something?  Call this function to find out what font size is required.
' If you will subsequently use the returned font size for testing, you can set "cacheIfNovel = True" to automatically cache a copy
' of the font at the newly detected font size.
Public Function FindFontSizeSingleLine(ByRef srcString As String, ByVal pxWidth As Long, ByVal initialFontSize As Single, Optional ByVal isBold As Boolean = False, Optional ByVal isItalic As Boolean = False, Optional ByVal isUnderline As Boolean = False, Optional ByVal cacheIfNovel As Boolean = True) As Single
    
    'Inside the designer, we need to make sure the font collection exists
    If Not g_IsProgramRunning Then
        If m_ProgramFontCollection Is Nothing Then InitProgramFontCollection
    End If
    
    'Add this font size+style combination to the collection
    Dim fontIndex As Long
    fontIndex = m_ProgramFontCollection.AddFontToCache(g_InterfaceFont, initialFontSize, isBold, isItalic, isUnderline)
    
    'Retrieve a handle to that font
    Dim tmpFont As pdFont
    Set tmpFont = m_ProgramFontCollection.GetFontObjectByPosition(fontIndex)
    
    'Return a smaller font size, as necessary, to fit the requested pixel width
    FindFontSizeSingleLine = tmpFont.GetMaxFontSizeToFitStringWidth(srcString, pxWidth, initialFontSize)
    
    'If the caller plans to use this new font size for immediate rendering, immediately cache a copy of the font at this new size
    If cacheIfNovel And (FindFontSizeSingleLine <> initialFontSize) Then
        m_ProgramFontCollection.AddFontToCache g_InterfaceFont, FindFontSizeSingleLine, isBold, isItalic, isUnderline
    End If
    
End Function

'Same as FindFontSizeSingleLine(), above, but with support for wordwrap
' If you will subsequently use the returned font size for testing, you can set "cacheIfNovel = True" to automatically cache a copy
' of the font at the newly detected font size.
Public Function FindFontSizeWordWrap(ByRef srcString As String, ByVal pxWidth As Long, ByVal pxHeight As Long, ByVal initialFontSize As Single, Optional ByVal isBold As Boolean = False, Optional ByVal isItalic As Boolean = False, Optional ByVal isUnderline As Boolean = False, Optional ByVal cacheIfNovel As Boolean = True) As Single
    
    'Inside the designer, we need to make sure the font collection exists
    If Not g_IsProgramRunning Then
        If m_ProgramFontCollection Is Nothing Then InitProgramFontCollection
    End If
    
    'Retrieve a handle to a matching pdFont object
    Dim tmpFont As pdFont
    Set tmpFont = Font_Management.GetMatchingUIFont(initialFontSize, isBold, isItalic, isUnderline)
    
    'Return a smaller font size, as necessary, to fit the requested pixel width
    FindFontSizeWordWrap = tmpFont.GetMaxFontSizeToFitWordWrap(srcString, pxWidth, pxHeight, initialFontSize)
    
    'If the caller plans to use this new font size for immediate rendering, immediately cache a copy of the font at this new size
    If cacheIfNovel And (FindFontSizeWordWrap <> initialFontSize) Then
        m_ProgramFontCollection.AddFontToCache g_InterfaceFont, FindFontSizeWordWrap, isBold, isItalic, isUnderline
    End If
    
End Function

'Want direct access to a UI font instance?  Get one here.  Note that only size, bold, italic, and underline are currently matched,
' as PD doesn't use strikethrough fonts anywhere.
Public Function GetMatchingUIFont(ByVal FontSize As Single, Optional ByVal isBold As Boolean = False, Optional ByVal isItalic As Boolean = False, Optional ByVal isUnderline As Boolean = False) As pdFont
    
    'Inside the designer, we need to make sure the font collection exists
    If Not g_IsProgramRunning Then
        If m_ProgramFontCollection Is Nothing Then InitProgramFontCollection
    End If
    
    'Add this font size+style combination to the collection, as necessary
    Dim fontIndex As Long
    fontIndex = m_ProgramFontCollection.AddFontToCache(g_InterfaceFont, FontSize, isBold, isItalic, isUnderline)
    
    'Return the handle of the newly created (and/or previously cached) pdFont object
    Set GetMatchingUIFont = m_ProgramFontCollection.GetFontObjectByPosition(fontIndex)
    
End Function

'If functions want their own copy of all available fonts on this PC, call this function
Public Function GetCopyOfSystemFontList(ByRef dstStringStack As pdStringStack) As Boolean
    If dstStringStack Is Nothing Then Set dstStringStack = New pdStringStack
    dstStringStack.cloneStack m_PDFontCache
End Function

'Build a system font cache.  Note that this is an expensive operation, and should never be called more than once.
' RETURNS: 0 if failure, Number of fonts (>= 0) if successful.  (Note that the *total number of all fonts* is returned,
'          not just TrueType ones.)
Public Function BuildFontCaches() As Long
    
    Set m_PDFontCache = New pdStringStack
    Set m_Unicode = New pdUnicode
    
    'Retrieve the current system LOGFONT conversion values
    UpdateLogFontValues
    
    'Next, prep a full font list for the advanced typography tool.
    '(We won't know the full number of available fonts until the Enum function finishes, so prep an extra-large buffer in advance.)
    m_PDFontCache.resetStack INITIAL_PDFONTCACHE_SIZE
    GetAllAvailableFonts
    
    'Because the font cache(s) will potentially be accessed by tons of external functions, it pays to sort them just once,
    ' at initialization time.
    m_PDFontCache.trimStack
    m_PDFontCache.SortAlphabetically True
    
    'TESTING ONLY!  Curious about the list of fonts?  Use this line to write it out to the immediate window
    'm_PDFontCache.DEBUG_dumpResultsToImmediateWindow
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "FYI - number of fonts found on this PC: " & m_PDFontCache.getNumOfStrings
    #End If
    
    'We have one other piece of initialization to do here.  Prep the program UI font cache that outside functions can use for
    ' their own UI painting.
    InitProgramFontCollection
    
End Function

'Converting between normal font sizes and GDI font sizes is convoluted, and it relies on a system-specific LOGPIXELSY value.
' We must cache that value before requesting fonts from the system.
Private Sub UpdateLogFontValues()
    Dim tmpDC As Long
    tmpDC = Drawing.GetMemoryDC()
    m_LogPixelsX = GetDeviceCaps(tmpDC, LogPixelsX)
    m_LogPixelsY = GetDeviceCaps(tmpDC, LOGPIXELSY)
    Drawing.FreeMemoryDC tmpDC
End Sub

'Prep the program font cache.  Individual functions may need to call this inside the designer, because PD's normal run-time
' initialization steps won't have fired.
Private Sub InitProgramFontCollection()
    
    Set m_ProgramFontCollection = New pdFontCollection
    
    'When outside callers request a copy of the system font, they are allowed to request any size+style they want.  Font name,
    ' however, never varies, so tell the font cache to only compare size and style when matching font requests.
    m_ProgramFontCollection.SetCacheMode FCM_SizeAndStyle
    
End Sub

'Retrieve all available fonts on this PC, regardless of font type
Private Function GetAllAvailableFonts() As Boolean

    'Prep a default LOGFONTW instance.  Note that EnumFontFamiliesEx only checks three params:
    ' lfCharSet:  If set to DEFAULT_CHARSET, the function enumerates all uniquely-named fonts in all character sets.
    '             (If there are two fonts with the same name, only one is enumerated.)
    '             If set to a valid character set value, the function enumerates only fonts in the specified character set.
    ' lfFaceName: If set to an empty string, the function enumerates one font in each available typeface name.
    '             If set to a valid typeface name, the function enumerates all fonts with the specified name.
    ' lfPitchAndFamily: Must be set to zero for all language versions of the operating system.
    Dim tmpLogFont As LOGFONTW
    tmpLogFont.lfCharSet = DEFAULT_CHARSET
    
    'Enumerate font families using a temporary DC
    Dim tmpDC As Long
    tmpDC = Drawing.GetMemoryDC()
    EnumFontFamiliesEx tmpDC, tmpLogFont, AddressOf EnumFontFamExProc, ByVal 0, 0
    Drawing.FreeMemoryDC tmpDC
    
    'If at least one font was found, return TRUE
    GetAllAvailableFonts = CBool(m_PDFontCache.getNumOfStrings > 0)

End Function

'Callback function for EnumFontFamiliesEx
Public Function EnumFontFamExProc(ByRef lpElfe As LOGFONTW, ByRef lpNtme As NEWTEXTMETRIC, ByVal srcFontType As Long, ByVal lParam As Long) As Long

    'Start by retrieving the font face name from the LogFontW struct
    Dim thisFontFace As String
    thisFontFace = String$(LF_FACESIZEA, 0)
    CopyMemory ByVal StrPtr(thisFontFace), ByVal VarPtr(lpElfe.lfFaceName(0)), LF_FACESIZEW
    thisFontFace = m_Unicode.TrimNull(thisFontFace)
    
    'If this font face is identical to the previous font face, do not add it
    Dim fontUsable As Boolean
    fontUsable = CBool(StrComp(thisFontFace, m_LastFontAdded, vbBinaryCompare) <> 0)
    
    'We also want to ignore fonts with @ in front of their name, as these are merely duplicates of existing fonts.
    ' (The @ signifies improved support for vertical text, which may someday be useful... but right now I have enough
    '  on my plate without worrying about that.)
    If fontUsable Then
        fontUsable = CBool(StrComp(Left$(thisFontFace, 1), "@", vbBinaryCompare) <> 0)
    End If
    
    'For now, we are also ignoring raster fonts, as they create unwanted complications
    If fontUsable Then
        fontUsable = CBool(CLng(srcFontType And RASTER_FONTTYPE) = 0)
    End If
    
    'If this font is a worthy addition, add it now
    If fontUsable Then
        
        m_PDFontCache.AddString thisFontFace
        
        'Make a copy of the last added font, so we can ignore duplicates
        m_LastFontAdded = thisFontFace
        
        'NOTE: Perhaps it could be helpful to cache the font type, or possibly use it to determine if fonts should be ignored?
        'm_PDFontCache(m_NumOfFonts).FontType = srcFontType
                
    End If
    
    'Return 1 so the enumeration continues
    EnumFontFamExProc = 1
    
End Function

'After the font cache has been successfully assembled, you can use this function to assemble a list of properties for each font.
Public Sub BuildFontCacheProperties()
    
    'Make sure the font cache exists
    If m_PDFontCache.getNumOfStrings > 0 Then
        
        'Sync the font property cache size to the font cache size
        ReDim g_PDFontProperties(0 To m_PDFontCache.getNumOfStrings - 1) As PD_FONT_PROPERTY
        
        Dim i As Long
        
        'Font properties can only be gathered on Vista or later
        If g_IsVistaOrLater Then
        
            'Iterate each font, gathering properties as we go
            'For i = 0 To UBound(g_PDFontProperties)
                
                'I'm temporarily disabling this while I investigate some performance issues on slow PCs
                'Uniscribe_Interface.getScriptsSupportedByFont m_PDFontCache.GetString(i), g_PDFontProperties(i)
                
                'Debug only: list fonts that support CJK forms
                'If g_PDFontProperties(i).Supports_CJK Then Debug.Print "Supports CJK: " & m_PDFontCache.GetString(i)
                
            'Next i
            
        'On XP, all scripts are marked as "unknown"
        Else
            
            'Temporarily disabled for the reasons explained above.
            'For i = 0 To UBound(g_PDFontProperties)
            '    g_PDFontProperties(i).ScriptsKnown = False
            'Next i
            
        End If
        
    End If
    
End Sub

'Given a DC with a font selected into it, return the primary charset for that DC
Public Function GetCharsetOfDC(ByRef srcDC As Long) As FONT_CHARSETS
    GetTextMetrics srcDC, m_TmpTextMetric
    GetCharsetOfDC = m_TmpTextMetric.tmCharSet
End Function

'Given some standard font characteristics (font face, style, etc), fill a corresponding LOGFONTW struct with matching values.
' This is helpful as PD stores characteristics in VB-friendly formats (e.g. booleans for styles), while LOGFONTW uses custom
' descriptors (e.g. font size, which is not calculated in an obvious way).
Public Sub FillLogFontW_Basic(ByRef dstLogFontW As LOGFONTW, ByRef srcFontFace As String, ByVal srcFontBold As Boolean, ByVal srcFontItalic As Boolean, ByVal srcFontUnderline As Boolean, ByVal srcFontStrikeout As Boolean)

    With dstLogFontW
    
        'For Unicode compatibility, the font face must be copied directly, without internal VB translations
        Dim copyLength As Long
        copyLength = Len(srcFontFace) * 2
        If copyLength > LF_FACESIZEW Then copyLength = LF_FACESIZEW
        CopyMemory ByVal VarPtr(.lfFaceName(0)), ByVal StrPtr(srcFontFace), copyLength
        
        'Bold is a unique style, because it must be translated to a corresponding weight measurement
        If srcFontBold Then
            .lfWeight = FW_BOLD
        Else
            .lfWeight = FW_NORMAL
        End If
        
        'Other styles all use the same pattern: multiply the bool by -1 to obtain a matching byte-type style
        .lfItalic = -1 * srcFontItalic
        .lfUnderline = -1 * srcFontUnderline
        .lfStrikeOut = -1 * srcFontStrikeout
        
        'While we're here, set charset to the default value; PD does not deviate from this (at present)
        .lfCharSet = DEFAULT_CHARSET
        
    End With
    
End Sub

'Fill a LOGFONTW struct with a matching PD font size (typically in pixels, but points are also supported)
Public Sub FillLogFontW_Size(ByRef dstLogFontW As LOGFONTW, ByVal FontSize As Single, ByVal fontMeasurementUnit As pdFontUnit)

    With dstLogFontW
        
        Select Case fontMeasurementUnit
        
            'Pixels use a modified version of the standard Windows formula; note that this assumes 96 DPI at present - high DPI
            ' systems still need testing!  TODO!
            Case pdfu_Pixel
                
                'Convert font size to points
                FontSize = FontSize * 0.75      '(72 / 96, technically, where 96 is the current screen DPI)
                
                'Use the standard point-based formula
                .lfHeight = Font_Management.ConvertToGDIFontSize(FontSize)
                
            'Points are converted using a standard Windows formula; see https://msdn.microsoft.com/en-us/library/dd145037%28v=vs.85%29.aspx
            Case pdfu_Point
                .lfHeight = Font_Management.ConvertToGDIFontSize(FontSize)
        
        End Select
        
        'Per convention, font width is set to 0 so the font mapper can select an aspect-ratio preserved width for us
        .lfWidth = 0
        
    End With
    
End Sub

Public Function ConvertToGDIFontSize(ByVal srcFontSize As Single) As Long
    If m_LogPixelsY = 0 Then UpdateLogFontValues
    ConvertToGDIFontSize = -1 * Internal_MulDiv(srcFontSize, m_LogPixelsY, 72#)
End Function

'It really isn't necessary to rely on the system MulDiv values for the sizes used for fonts.  By using CLng, we can mimic
' MulDiv's rounding behavior as well.
Private Function Internal_MulDiv(ByVal srcNumber As Single, ByVal srcNumerator As Single, ByVal srcDenominator As Single) As Long
    Internal_MulDiv = CLng((srcNumber * srcNumerator) / srcDenominator)
End Function

'Once I have a better idea of what I can do with font quality, I'll be switching the fontQuality Enum to something internal to PD.
' But right now, I'm still in the exploratory phase, and trying to figure out whether different font quality settings affect
' the glyph outline returned.  (They should, technically, since hinting affects font shape.)
Public Sub FillLogFontW_Quality(ByRef dstLogFontW As LOGFONTW, ByVal fontQuality As GdiPlusTextRenderingHint)

    Dim gdiFontQuality As Long
    
    'Per http://stackoverflow.com/questions/1203087/why-is-graphics-measurestring-returning-a-higher-than-expected-number?lq=1
    ' this function mirrors the .NET conversion from GDI+ text rendering hints to GDI font quality settings.  Mapping deliberately
    ' ignores some settings (no idea why, but if the .NET stack does it, there's probably a reason)
    Select Case fontQuality
    
        Case TextRenderingHintSystemDefault
            gdiFontQuality = DEFAULT_QUALITY
            
        Case TextRenderingHintSingleBitPerPixel
            gdiFontQuality = DRAFT_QUALITY
        
        Case TextRenderingHintSingleBitPerPixelGridFit
            gdiFontQuality = PROOF_QUALITY
        
        Case TextRenderingHintAntiAlias
            gdiFontQuality = ANTIALIASED_QUALITY
        
        Case TextRenderingHintAntiAliasGridFit
            gdiFontQuality = ANTIALIASED_QUALITY
        
        Case TextRenderingHintClearTypeGridFit
            gdiFontQuality = CLEARTYPE_QUALITY
        
        Case Else
            Debug.Print "Unknown font quality passed; please double-check requests to fillLogFontW_Quality"
    
    End Select
    
    dstLogFontW.lfQuality = gdiFontQuality

End Sub

'Retrieve a text metrics struct for a given DC.  Obviously, the desired font needs to be selected into the DC *prior* to calling this.
Public Function FillTextMetrics(ByRef srcDC As Long, ByRef dstTextMetrics As TEXTMETRIC) As Boolean
    
    Dim gtmReturn As Long
    gtmReturn = GetTextMetrics(srcDC, dstTextMetrics)
    
    FillTextMetrics = CBool(gtmReturn <> 0)
    
End Function

Public Function FillOutlineTextMetrics(ByRef srcDC As Long, ByRef dstOutlineMetrics As OUTLINETEXTMETRIC) As Boolean
    
    'Retrieve the size required by the struct
    Dim structSize As Long
    structSize = GetOutlineTextMetrics(srcDC, 0, ByVal 0&)
    
    'Because GOTM adds four trailing strings to the struct, we need to prep an array large enough to receive the entire structure
    ' PLUS those strings.  After retrieving the full struct + strings, we'll manually extract just the struct portion.
    Dim tmpBytes() As Byte
    ReDim tmpBytes(0 To structSize) As Byte
    
    Dim gtmReturn As Long
    gtmReturn = GetOutlineTextMetrics(srcDC, structSize, VarPtr(tmpBytes(0)))
    
    'If the byte array was filled successfully, parse out the struct now.  (I don't have need for the additional name values
    ' right now, so they're just ignored.)
    If (gtmReturn <> 0) Then
        
        FillOutlineTextMetrics = True
        CopyMemory ByVal VarPtr(dstOutlineMetrics), ByVal VarPtr(tmpBytes(0)), LenB(dstOutlineMetrics)
        
    Else
        FillOutlineTextMetrics = False
    End If
    
End Function

'Given a filled LOGFONTW struct (hopefully filled by the fillLogFontW_* functions above!), attempt to create an actual font object.
' Returns TRUE if successful; FALSE otherwise.
Public Function createGDIFont(ByRef srcLogFont As LOGFONTW, ByRef dstFontHandle As Long) As Boolean
    dstFontHandle = CreateFontIndirect(srcLogFont)
    createGDIFont = CBool(dstFontHandle <> 0)
End Function

'Delete a GDI font; returns TRUE if successful
Public Function DeleteGDIFont(ByVal srcFontHandle As Long) As Boolean
    DeleteGDIFont = CBool(DeleteObject(srcFontHandle) <> 0)
End Function

'Given a GDI font handle and a Unicode code point, return an ABC float for the corresponding glyph.
' By default, the passed font handle MUST NOT BE SELECTED INTO A DC.  However, to make interaction easier with GDI rendering code,
' you can set the optional fontHandleIsReallyDC value to TRUE, and obviously pass in a DC instead of font handle, and this function
' will assume you have already selected the relevant font into the DC for it.
Public Function GetABCWidthOfGlyph(ByVal srcFontHandle As Long, ByVal charCodeInQuestion As Long, ByRef dstABCFloat As ABCFLOAT, Optional ByVal fontHandleIsReallyDC As Boolean = False) As Boolean
    
    Dim gdiReturn As Long
    
    'If the user has selected the font into a DC for us, this function is incredibly simple
    If fontHandleIsReallyDC Then
    
        'Retrieve the character positioning values
        gdiReturn = GetCharABCWidthsFloat(srcFontHandle, charCodeInQuestion, charCodeInQuestion, VarPtr(dstABCFloat))
    
    'If the user only has a bare font handle, we have to handle the DC step ourselves, unfortunately
    Else
        
        'Temporarily select the font into a local DC
        Dim origFont As Long, tmpDC As Long
        tmpDC = Drawing.GetMemoryDC()
        origFont = SelectObject(tmpDC, srcFontHandle)
        
        'Retrieve the character positioning values
        gdiReturn = GetCharABCWidthsFloat(tmpDC, charCodeInQuestion, charCodeInQuestion, VarPtr(dstABCFloat))
        
        'Release the font
        SelectObject tmpDC, origFont
        Drawing.FreeMemoryDC tmpDC
    
    End If
    
    'GetCharABCWidthsFloat() returns a non-zero value if successful
    GetABCWidthOfGlyph = CBool(gdiReturn <> 0)
    
End Function

'Given a font name, quickly generate a GDI font handle with default settings, and shove it into a temporary DC.
' IMPORTANT NOTE: the caller needs to cache the font and DC handle, then pass them to the clean-up function below
Public Function QuickCreateFontAndDC(ByRef srcFontName As String, ByRef dstFont As Long, ByRef dstDC As Long) As Boolean
    
    Dim tmpLogFont As LOGFONTW
    FillLogFontW_Basic tmpLogFont, srcFontName, False, False, False, False
    If createGDIFont(tmpLogFont, dstFont) Then
        
        'Create a temporary DC and select the font into it
        dstDC = Drawing.GetMemoryDC()
        SelectObject dstDC, dstFont
        
        QuickCreateFontAndDC = True
        
    Else
        QuickCreateFontAndDC = False
    End If
    
End Function

Public Sub QuickDeleteFontAndDC(ByRef srcFont As Long, ByRef srcDC As Long)
    
    'Remove the font
    SelectObject srcDC, GetStockObject(SYSTEM_FONT)
    
    'Kill both the font and the DC
    DeleteObject srcFont
    Drawing.FreeMemoryDC srcDC
    
End Sub
