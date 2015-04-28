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
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsW" (ByVal hDC As Long, ByRef lpMetrics As TEXTMETRIC) As Long
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
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type

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

'GDI+ font collection interfaces; these are temporarily in-use while I sort out OpenType handling
Private Declare Function GdipNewInstalledFontCollection Lib "gdiplus" (ByRef dstFontCollectionHandle As Long) As Long
Private Declare Function GdipGetFontCollectionFamilyCount Lib "gdiplus" (ByVal srcFontCollection As Long, ByRef dstNumFound As Long) As Long
Private Declare Function GdipGetFontCollectionFamilyList Lib "gdiplus" (ByVal srcFontCollection As Long, ByVal sizeOfDstBuffer As Long, ByVal ptrToDstFontFamilyArray As Long, ByRef dstNumFound As Long) As Long
Private Declare Function GdipGetFamilyName Lib "gdiplus" (ByVal srcFontFamily As Long, ByVal ptrDstNameBuffer As Long, ByVal languageID As Integer) As Long
Private Const LF_FACESIZE As Long = 32          'Note: this represents 32 *chars*, not bytes!
Private Const LANG_NEUTRAL As Integer = &H0

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

'GDI font creation
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectW" (ByRef lpLogFont As LOGFONTW) As Long

'Various non-font-specific WAPI functions helpful for font assembly
Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

'Some system-specific font settings are cached at initialization time, and unchanged for the life of the program
Private curLogPixelsX As Long, curLogPixelsY As Long

'Internal font cache.  PD uses this to populate things like font selection dropdowns.
Private m_PDFontCache As pdStringStack
Private Const INITIAL_PDFONTCACHE_SIZE As Long = 256
Private m_LastFontAdded As String

'Temporary DIB (more importantly - DC) for testing and/or applying font settings
Private m_TestDIB As pdDIB

'If functions want their own copy of available fonts, call this function
Public Function getCopyOfFontCache(ByRef dstStringStack As pdStringStack) As Boolean
    If dstStringStack Is Nothing Then Set dstStringStack = New pdStringStack
    dstStringStack.cloneStack m_PDFontCache
End Function

'Build a system font cache.  Note that this is an expensive operation, and should never be called more than once.
' RETURNS: 0 if failure, Number of fonts (>= 0) if successful
Public Function buildFontCache(Optional ByVal getTrueTypeOnly As Boolean = False) As Long
    
    'Prep the default font cache
    Set m_PDFontCache = New pdStringStack
    
    'Prep a tiny internal DIB for testing font settings
    Set m_TestDIB = New pdDIB
    m_TestDIB.createBlank 4, 4, 32, 0, 0
    
    'Use the DIB to retrieve system-specific font conversion values
    curLogPixelsX = GetDeviceCaps(m_TestDIB.getDIBDC, LOGPIXELSX)
    curLogPixelsY = GetDeviceCaps(m_TestDIB.getDIBDC, LOGPIXELSY)
    
    'We now branch into two possible directions:
    ' 1) If getTrueTypeOnly is FALSE, we retrieve all fonts on the PC via GDI's EnumFontFamiliesEx
    ' 2) If getTrueTypeOnly is TRUE, we retrieve only TrueType fonts via GDI+'s getFontFamilyCollectionList function.
    If getTrueTypeOnly Then
        
        'GDI+ will return the font count prior to enumeration, so we don't need to prep the string stack in advance
        getAllAvailableTrueTypeFonts
        
    Else
        
        'We won't know the full number of available fonts until the Enum function finishes, so prep an extra-large
        ' buffer in advance.
        m_PDFontCache.resetStack INITIAL_PDFONTCACHE_SIZE
        getAllAvailableFonts
        
    End If
    
    
    'Because the font cache will potentially be accessed by tons of external functions, it pays to sort it just once,
    ' at initialization time.
    m_PDFontCache.trimStack
    m_PDFontCache.SortAlphabetically
    
    'TESTING ONLY!  Curious about the list of fonts?  Use this line to write it out to the immediate window
    'm_PDFontCache.DEBUG_dumpResultsToImmediateWindow
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "FYI - number of fonts found on this PC: " & m_PDFontCache.getNumOfStrings
    #End If
    
End Function

'Retrieve all available fonts on this PC, regardless of font type
Private Function getAllAvailableFonts() As Boolean

    'Prep a default LOGFONTW instance.  Note that EnumFontFamiliesEx only checks three params:
    ' lfCharSet:  If set to DEFAULT_CHARSET, the function enumerates all uniquely-named fonts in all character sets.
    '             (If there are two fonts with the same name, only one is enumerated.)
    '             If set to a valid character set value, the function enumerates only fonts in the specified character set.
    ' lfFaceName: If set to an empty string, the function enumerates one font in each available typeface name.
    '             If set to a valid typeface name, the function enumerates all fonts with the specified name.
    ' lfPitchAndFamily: Must be set to zero for all language versions of the operating system.
    Dim tmpLogFont As LOGFONTW
    tmpLogFont.lfCharSet = DEFAULT_CHARSET
    
    'Enumerate font families using our temporary DIB DC
    EnumFontFamiliesEx m_TestDIB.getDIBDC, tmpLogFont, AddressOf EnumFontFamExProc, ByVal 0, 0
    
    'If at least one font was found, return TRUE
    getAllAvailableFonts = CBool(m_PDFontCache.getNumOfStrings > 0)

End Function

'Callback function for EnumFontFamiliesEx
Public Function EnumFontFamExProc(ByRef lpElfe As LOGFONTW, ByRef lpNtme As NEWTEXTMETRIC, ByVal srcFontType As Long, ByVal lParam As Long) As Long

    'Start by retrieving the font face name from the LogFontW struct
    Dim thisFontFace As String
    thisFontFace = String$(LF_FACESIZEA, 0)
    CopyMemory ByVal StrPtr(thisFontFace), ByVal VarPtr(lpElfe.lfFaceName(0)), LF_FACESIZEW
    thisFontFace = TrimNull(thisFontFace)
    
    'If this font face is identical to the previous font face, do not add it
    Dim fontUsable As Boolean
    fontUsable = CBool(StrComp(thisFontFace, m_LastFontAdded, vbBinaryCompare) <> 0)
    
    'We also want to ignore fonts with @ in front of their name, as these are merely duplicates of existing fonts.
    ' (The @ signifies improved support for vertical text, which may someday be useful... but right now I have enough
    '  on my plate without worrying about that.)
    If fontUsable Then
        fontUsable = CBool(StrComp(Left$(thisFontFace, 1), "@", vbBinaryCompare) <> 0)
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

'Helper function for returning a string stack of currently installed, GDI+ compatible (e.g. TrueType) fonts
Private Function getAllAvailableTrueTypeFonts() As Boolean
    
    'Create a new GDI+ font collection object
    Dim fontCollection As Long
    If GdipNewInstalledFontCollection(fontCollection) = 0 Then
    
        'Get the family count
        Dim fontCount As Long
        If GdipGetFontCollectionFamilyCount(fontCollection, fontCount) = 0 Then
        
            'Prep a Long-type array, which will receive the list of fonts installed on this machine
            Dim fontList() As Long
            If fontCount > 0 Then ReDim fontList(0 To fontCount - 1) As Long Else ReDim fontList(0) As Long
        
            'I don't know if it's possible for GDI+ to return a different amount of fonts than it originally reported,
            ' but since it takes a parameter for numFound, let's use it
            Dim fontsFound As Long
            If GdipGetFontCollectionFamilyList(fontCollection, fontCount, VarPtr(fontList(0)), fontsFound) = 0 Then
            
                'Populate our string stack with the names of this collection; also, since we know the approximate size of
                ' the stack in advance, we can accurately prep the stack's buffer.
                m_PDFontCache.resetStack fontCount
                
                'Retrieve all font names
                Dim i As Long, thisFontName As String
                For i = 0 To fontsFound - 1
                    
                    'Retrieve the name for this entry
                    thisFontName = String$(LF_FACESIZE, 0)
                    If GdipGetFamilyName(fontList(i), StrPtr(thisFontName), LANG_NEUTRAL) = 0 Then
                        m_PDFontCache.AddString TrimNull(thisFontName)
                    End If
                    
                Next i
                
                'Return success
                getAllAvailableTrueTypeFonts = True
            
            Else
                Debug.Print "WARNING! GDI+ refused to return a font collection list."
                getAllAvailableTrueTypeFonts = False
            End If
        
        Else
            Debug.Print "WARNING! GDI+ refused to return a font collection count."
            getAllAvailableTrueTypeFonts = False
        End If
    
    Else
        Debug.Print "WARNING! GDI+ refused to return a font collection object."
        getAllAvailableTrueTypeFonts = False
    End If
    
End Function

'This function is identical to PD's publicly declared "TrimNull" function in File_And_Path_Handling.  It is included here to reduce
' external dependencies for this class.
Private Function TrimNull(ByRef origString As String) As String

    'See if the incoming string contains null chars
    Dim nullPosition As Long
    nullPosition = InStr(origString, ChrW$(0))
    
    'If it does, trim accordingly
    If nullPosition Then
       TrimNull = Left$(origString, nullPosition - 1)
    Else
       TrimNull = origString
    End If
  
End Function

'Given some standard font characteristics (font face, style, etc), fill a corresponding LOGFONTW struct with matching values.
' This is helpful as PD stores characteristics in VB-friendly formats (e.g. booleans for styles), while LOGFONTW uses custom
' descriptors (e.g. font size, which is not calculated in an obvious way).
Public Sub fillLogFontW_Basic(ByRef dstLogFontW As LOGFONTW, ByRef srcFontFace As String, ByVal srcFontBold As Boolean, ByVal srcFontItalic As Boolean, ByVal srcFontUnderline As Boolean, ByVal srcFontStrikeout As Boolean)

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
Public Sub fillLogFontW_Size(ByRef dstLogFontW As LOGFONTW, ByVal FontSize As Single, ByVal fontMeasurementUnit As pdFontUnit)

    With dstLogFontW
        
        Select Case fontMeasurementUnit
        
            'Pixels use a modified version of the standard Windows formula; note that this assumes 96 DPI at present - high DPI
            ' systems still need testing!  TODO!
            Case pdfu_Pixel
                
                'Convert font size to points
                FontSize = FontSize * 0.75      '(72 / 96, technically, where 96 is the current screen DPI)
                
                'Use the standard point-based formula
                .lfHeight = -1 * internal_MulDiv(FontSize, curLogPixelsY, 72)
                
            'Points are converted using a standard Windows formula; see https://msdn.microsoft.com/en-us/library/dd145037%28v=vs.85%29.aspx
            Case pdfu_Point
                .lfHeight = -1 * internal_MulDiv(FontSize, curLogPixelsY, 72)
        
        End Select
        
        'Per convention, font width is set to 0 so the font mapper can select an aspect-ratio preserved width for us
        .lfWidth = 0
        
    End With
    
End Sub

'It really isn't necessary to rely on the system MulDiv values for the sizes used for fonts.  By using CLng, we can mimic
' MulDiv's rounding behavior as well.
Private Function internal_MulDiv(ByVal srcNumber As Single, ByVal srcNumerator As Single, ByVal srcDenominator As Single) As Long
    internal_MulDiv = CLng((srcNumber * srcNumerator) / srcDenominator)
End Function

'Once I have a better idea of what I can do with font quality, I'll be switching the fontQuality Enum to something internal to PD.
' But right now, I'm still in the exploratory phase, and trying to figure out whether different font quality settings affect
' the glyph outline returned.  (They should, technically, since hinting affects font shape.)
Public Sub fillLogFontW_Quality(ByRef dstLogFontW As LOGFONTW, ByVal fontQuality As GdiPlusTextRenderingHint)

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
Public Function fillTextMetrics(ByRef srcDC As Long, ByRef dstTextMetrics As TEXTMETRIC) As Boolean
    
    Dim gtmReturn As Long
    gtmReturn = GetTextMetrics(srcDC, dstTextMetrics)
    
    fillTextMetrics = CBool(gtmReturn <> 0)
    
End Function

'Given a filled LOGFONTW struct (hopefully filled by the fillLogFontW_* functions above!), attempt to create an actual font object.
' Returns TRUE if successful; FALSE otherwise.
Public Function createGDIFont(ByRef srcLogFont As LOGFONTW, ByRef dstFontHandle As Long) As Boolean
    dstFontHandle = CreateFontIndirect(srcLogFont)
    createGDIFont = CBool(dstFontHandle <> 0)
End Function
