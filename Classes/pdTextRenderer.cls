VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdTextRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon GDI+ Text Manager and Renderer
'Copyright 2015-2017 by Tanner Helland
'Created: 22/April/15
'Last updated: 19/May/15
'Last update: various changes to improve compatibility with the new Uniscribe backend (see pdGlyphCollection)
'
'This class handles the messy business of font management for PD's text and typography layers.  It intermixes
' GDI, GDI+, and Uniscribe as necessary to provide a comprehensive text rendering solution.
'
'I originally planned to split this class in two, as a GDI-specific backend and a separate GDI+-specific backend (with
' Uniscribe stuff handled separately).  But there are complications with this.  GDI and GDI+ have different strengths
' and weaknesses (GDI has no 32-bpp rendering support, while GDI+ doesn't support OpenType), so to really provide a
' decent font experience without native libraries, one must utilize parts of all the libraries.
'
'Hence this class.  Here's the basic setup:
'
'For the WAPI text rendering engine...
'
' - With TrueType fonts, GDI+ is used.  This saves us a lot of grief, and lets us make use of some nice, typographic-specific
' features (like variable antialiasing contrast).
'
' - With OpenType fonts, GDI is used out of necessity.  Individual glyphs are retrieved, parsed, and converted into a usable
' format (e.g. OpenType's quadratic B-splines are converted into GDI+'s native cubic Bezier curve format).  The final,
' resultant shape(s) are then passed to GDI+ as a generic path object, which we can render with any of GDI+'s native
' features (antialiasing, fancy gradients, etc).
'
'This unpleasant hybrid approach is part of the shit we deal with as Windows developers, I guess.
'
'For the PHOTODEMON text rendering engine...
'
' - A comprehensive Uniscribe interface is used for text shaping, layout, and placement.  This allows us to make use of just
'   about every possible OpenType font feature, which is amazing!
'
' - Glyph paths are retrieved via GDI, and manually parsed into GDI+ path objects.  The details are handled by the
'   pdGlyphCollection class.
'
' - Actual rendering duties are handled by GDI+.
'
'Dependencies:
'  - Fonts module, which handles some of the messier bits of GDI interaction
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'To simplify the process of setting/getting text settings for a given layer, this Enum is used to pass text properties
Public Enum PD_TEXT_PROPERTY
    ptp_Text = 0
    ptp_FontColor = 1
    ptp_FontFace = 2
    ptp_FontSize = 3
    ptp_FontSizeUnit = 4
    ptp_FontBold = 5
    ptp_FontItalic = 6
    ptp_FontUnderline = 7
    ptp_FontStrikeout = 8
    ptp_HorizontalAlignment = 9
    ptp_VerticalAlignment = 10
    ptp_TextAntialiasing = 11
    ptp_TextContrast = 12
    ptp_RenderingEngine = 13
    ptp_TextHinting = 14
    ptp_WordWrap = 15
    ptp_FillActive = 16
    ptp_FillBrush = 17
    ptp_OutlineActive = 18
    ptp_OutlinePen = 19
    ptp_BackgroundActive = 20
    ptp_BackgroundBrush = 21
    ptp_BackBorderActive = 22
    ptp_BackBorderPen = 23
    ptp_LineSpacing = 24
    ptp_MarginLeft = 25
    ptp_MarginTop = 26
    ptp_MarginRight = 27
    ptp_MarginBottom = 28
    ptp_CharRemap = 29
    ptp_CharSpacing = 30
    ptp_CharOrientation = 31
    ptp_CharJitterX = 32
    ptp_CharJitterY = 33
    ptp_CharInflation = 34
    ptp_CharMirror = 35
End Enum

#If False Then
    Const ptp_Text = 0, ptp_FontColor = 1, ptp_FontFace = 2, ptp_FontSize = 3, ptp_FontSizeUnit = 4, ptp_FontBold = 5, ptp_FontItalic = 6
    Const ptp_FontUnderline = 7, ptp_FontStrikeout = 8, ptp_HorizontalAlignment = 9, ptp_VerticalAlignment = 10, ptp_TextAntialiasing = 11
    Const ptp_TextContrast = 12, ptp_RenderingEngine = 13, ptp_TextHinting = 14, ptp_WordWrap = 15, ptp_FillActive = 16, ptp_FillBrush = 17
    Const ptp_OutlineActive = 18, ptp_OutlinePen = 19, ptp_BackgroundActive = 20, ptp_BackgroundBrush = 21, ptp_BackBorderActive = 22
    Const ptp_BackBorderPen = 23, ptp_LineSpacing = 24, ptp_MarginLeft = 25, ptp_MarginTop = 26, ptp_MarginRight = 27, ptp_MarginBottom = 28
    Const ptp_CharRemap = 29, ptp_CharSpacing = 30, ptp_CharOrientation = 31, ptp_CharJitterX = 32, ptp_CharJitterY = 33, ptp_CharInflation = 34
    Const ptp_CharMirror = 35
#End If

'At present, only WAPI font backends are available.  This may change in the future (Cairo+Pango would be particularly nice to implement).
Private Enum fontBackend
    fb_UNKNOWN = 0
    fb_GDIPLUS = 1
    fb_GDI = 2
End Enum

#If False Then
    Private Const fb_UNKNOWN = 0, fb_GDIPLUS = 1, fb_GDI = 2
#End If

'Because I want switches between GDI and GDI+ to be transparent, this class will dynamically track which engine is required
' for rendering.  If a setting changes (and renders the font cache invalid), this value will be reset to UNKNOWN.  On the next
' font object creation attempt, the required backend will be automatically determined, based on the current class settings.
Private m_CurrentBackend As fontBackend

'This class provides two basic rendering engines: WAPI and MANUAL.  WAPI limits you to innate features of GDI and GDI+ (BOOOORING).
' MANUAL offloads all rendering to PD's manual glyph renderer, and enables... well, pretty much anything I can think to implement.
Public Enum TEXT_RENDERING_ENGINE
    tre_WAPI = 0
    tre_PHOTODEMON = 1
End Enum

#If False Then
    Private Const tre_WAPI = 0, tre_PHOTODEMON = 1
#End If

Private m_RenderingEngine As TEXT_RENDERING_ENGINE

'This class's custom glyph renderer supports various wordwrap modes
Public Enum TEXT_WORDWRAP_MODE
    tre_None = 0
    tre_ManualOnly = 1
    tre_AutoBreakCharacter = 2
    tre_AutoBreakWords = 3
End Enum

#If False Then
    Private Const tre_None = 0, tre_ManualOnly = 1, tre_AutoBreakCharacter = 2, tre_AutoBreakWords = 3
#End If

'GDI+ font functions are prone to failure, for reasons both known (e.g. no OpenType support) and unknown (e.g. GDI+ font handling is
' a pile of garbage).  Here are a few common failure states.
Private Const gdipFontFamilyNotFound As Long = 14
Private Const gdipFontStyleNotFound As Long = 15
Private Const gdipNotTrueTypeFont As Long = 16

'GDI+ supports multiple measurement modes, and these may actually be relevant for fonts.  Because GDI+ provides some units
' that we don't want to expose to users (e.g. "World", "Display"), PD uses its own font size enumeration, which we translate
' to corresponding GDI+ units as necessary.
Public Enum pdFontUnit
    pdfu_Pixel = 0
    pdfu_Point = 1
End Enum

#If False Then
    Private Const pdfu_Pixel = 0, pdfu_Point = 1
#End If

Private Enum GdiPlusMeasurementUnit
    gpUnitWorld = 0       '(non-physical unit; do not use)
    gpUnitDisplay = 1     '(varies by display, as you can imagine; do not use)
    gpUnitPixel = 2
    gpUnitPoint = 3       '1/72 in
    gpUnitInch = 4        '1 in
    gpUnitDocument = 5    '1/300 in
    gpUnitMillimeter = 6  '1 mm
End Enum

#If False Then
    Private Const gpUnitWorld = 0, gpUnitDisplay = 1, gpUnitPixel = 2, gpUnitPoint = 3, gpUnitInch = 4, gpUnitDocument = 5, gpUnitMillimeter = 6
#End If

'GDI+ maps a number of enums to this default "QualityMode" enum.  These values tend to be unpredictable, so I recommend sticking with
' explicit mapping options whenever possible.
Private Enum GdiPlusQualityMode
    QualityModeInvalid = -1
    QualityModeDefault = 0
    QualityModeLow = 1
    QualityModeHigh = 2
End Enum

#If False Then
    Private Const QualityModeInvalid = -1, QualityModeDefault = 0, QualityModeLow = 1, QualityModeHigh = 2
#End If

'SmoothingMode controls generic AA operations for pens
Private Enum GdiPlusSmoothingMode
    SmoothingModeInvalid = QualityModeInvalid
    SmoothingModeDefault = QualityModeDefault
    SmoothingModeHighSpeed = QualityModeLow
    SmoothingModeHighQuality = QualityModeHigh
    GP_SM_None
    GP_SM_Antialias
End Enum

#If False Then
    Private Const SmoothingModeInvalid = QualityModeInvalid, SmoothingModeDefault = QualityModeDefault, SmoothingModeHighSpeed = QualityModeLow, SmoothingModeHighQuality = QualityModeHigh
    Private Const GP_SM_None = 0, GP_SM_Antialias = 0
#End If

'GDI+ supports multiple text antialiasing modes, but half of them don't work on 32-bpp rendering targets (UGGGGGH).  Rather than expose a confusing list
' of operators that may or may not be relevant, PD uses its own AA enumeration, which we translate to corresponding GDI+ units as necessary.
Public Enum pdTextAA
    pdta_None = 0
    pdta_Standard = 1
    pdta_Crisp = 2
End Enum

#If False Then
    Private Const pdta_None = 0, pdta_Standard = 1, pdta_Crisp = 2
#End If

'Text hinting and AA are merged into a single setting in GDI+
Public Enum GdiPlusTextRenderingHint
    TextRenderingHintSystemDefault = 0              'System default; unused because it is unpredictable
    TextRenderingHintSingleBitPerPixelGridFit = 1   'Hinting, No AA
    TextRenderingHintSingleBitPerPixel = 2          'No Hinting, No AA
    TextRenderingHintAntiAliasGridFit = 3           'Hinting, grayscale AA (forces font sizes to integer values)
    TextRenderingHintAntiAlias = 4                  'No Hinting, grayscale AA (supports any size, at a fuzziness trade-off)
    TextRenderingHintClearTypeGridFit = 5           'Hinting, LCD-specific AA.  There is no "non-hinted ClearType" variant.
End Enum

#If False Then
    Private Const TextRenderingHintSystemDefault = 0, TextRenderingHintSingleBitPerPixelGridFit = 1, TextRenderingHintSingleBitPerPixel = 2
    Private Const TextRenderingHintAntiAliasGridFit = 3, TextRenderingHintAntiAlias = 4, TextRenderingHintClearTypeGridFit = 5
#End If

'String format settings
Private Enum StringFormatFlags
    StringFormatFlagsDirectionRightToLeft = &H1
    StringFormatFlagsDirectionVertical = &H2
    StringFormatFlagsFitBlackBox = &H4
    StringFormatFlagsDisplayFormatControl = &H20
    StringFormatFlagsNoFontFallback = &H400
    StringFormatFlagsMeasureTrailingSpaces = &H800
    StringFormatFlagsNoWrap = &H1000
    StringFormatFlagsLineLimit = &H2000
    StringFormatFlagsNoClip = &H4000
End Enum

#If False Then
    Private Const StringFormatFlagsDirectionRightToLeft = &H1, StringFormatFlagsDirectionVertical = &H2, StringFormatFlagsFitBlackBox = &H4, StringFormatFlagsDisplayFormatControl = &H20
    Private Const StringFormatFlagsNoFontFallback = &H400, StringFormatFlagsMeasureTrailingSpaces = &H800, StringFormatFlagsNoWrap = &H1000, StringFormatFlagsLineLimit = &H2000
    Private Const StringFormatFlagsNoClip = &H4000
#End If

'Font style settings
Private Enum FontStyle
    FontStyleRegular = 0
    FontStyleBold = 1
    FontStyleItalic = 2
    FontStyleBoldItalic = 3
    FontStyleUnderline = 4
    FontStyleStrikeout = 8
End Enum

#If False Then
    Const FontStyleRegular = 0, FontStyleBold = 1, FontStyleItalic = 2, FontStyleBoldItalic = 3, FontStyleUnderline = 4, FontStyleStrikeout = 8
#End If

'GDI+ string format settings.  Note that "near" and "far" monikers are used, to better support RTL languages.
Public Enum GdiPlusStringAlignment
    StringAlignmentNear = 0
    StringAlignmentCenter = 1
    StringAlignmentFar = 2
End Enum

#If False Then
    Private Const StringAlignmentNear = 0, StringAlignmentCenter = 1, StringAlignmentFar = 2
#End If

'GDI+ font family functions
Private Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal ptrToSrcFontName As Long, ByVal srcFontCollection As Long, ByRef dstFontFamily As Long) As Long
Private Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal srcFontFamily As Long) As Long
Private Declare Function GdipIsStyleAvailable Lib "gdiplus" (ByVal srcFontFamily As Long, ByVal srcStyleToTest As FontStyle, ByRef dstIsStyleAvailable As Long) As Long

'GDI+ font functions
Private Declare Function GdipCreateFont Lib "gdiplus" (ByVal srcFontFamily As Long, ByVal srcFontSize As Single, ByVal srcFontStyle As FontStyle, ByVal srcMeasurementUnit As GdiPlusMeasurementUnit, ByRef dstCreatedFont As Long) As Long
Private Declare Function GdipDeleteFont Lib "gdiplus" (ByVal srcFont As Long) As Long

'GDI+ string format functions
Private Declare Function GdipCreateStringFormat Lib "gdiplus" (ByVal formatAttributes As StringFormatFlags, ByVal srcLanguage As Long, ByRef dstStringFormat As Long) As GP_Result
Private Declare Function GdipStringFormatGetGenericTypographic Lib "gdiplus" (ByRef dstStringFormat As Long) As Long
Private Declare Function GdipDeleteStringFormat Lib "gdiplus" (ByVal srcStringFormat As Long) As Long
Private Declare Function GdipSetStringFormatFlags Lib "gdiplus" (ByVal dstStringFormat As Long, ByVal newFlags As Long) As Long
Private Declare Function GdipGetStringFormatFlags Lib "gdiplus" (ByVal srcStringFormat As Long, ByRef dstFlags As Long) As Long
Private Declare Function GdipGetStringFormatAlign Lib "gdiplus" (ByVal srcStringFormat As Long, ByRef dstAlignment As GdiPlusStringAlignment) As GP_Result
Private Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal dstStringFormat As Long, ByVal newAlignment As GdiPlusStringAlignment) As Long
Private Declare Function GdipSetStringFormatLineAlign Lib "gdiplus" (ByVal dstStringFormat As Long, ByVal newLineAlignment As GdiPlusStringAlignment) As Long
Private Declare Function GdipMeasureString Lib "gdiplus" (ByVal srcGraphics As Long, ByVal srcStringPtr As Long, ByVal numCharsInString As Long, ByVal srcFontObject As Long, ByRef srcLayoutRect As RECTF, ByVal srcStringFormat As Long, ByRef dstBoundingBox As RECTF, ByRef optNumCharactersFitted As Long, ByRef optNumLinesFitted As Long) As Long

'GDI+ graphics container font functions
Private Declare Function GdipSetTextRenderingHint Lib "gdiplus" (ByVal dstGraphics As Long, ByVal newRenderHintMode As GdiPlusTextRenderingHint) As Long
Private Declare Function GdipSetTextContrast Lib "gdiplus" (ByVal dstGraphics As Long, ByVal textContrast As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal dstGraphics As Long, ByVal newSmoothingMode As GdiPlusSmoothingMode) As Long

'GDI+ Render functions
Private Declare Function GdipDrawString Lib "gdiplus" (ByVal dstGraphics As Long, ByVal srcStringPtr As Long, ByVal strLength As Long, ByVal gdipFontHandle As Long, ByRef layoutRect As RECTF, ByVal gdipStringFormat As Long, ByVal gdipBrush As Long) As Long
Private Declare Function GdipDrawPath Lib "gdiplus" (ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal srcPath As Long) As Long
Private Declare Function GdipDrawRectangle Lib "gdiplus" (ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal srcX As Single, ByVal srcW As Single, ByVal srcWidth As Single, ByVal srcHeight As Single) As Long
Private Declare Function GdipFillPath Lib "gdiplus" (ByVal dstGraphics As Long, ByVal srcBrush As Long, ByVal srcPath As Long) As Long
Private Declare Function GdipFillRectangleI Lib "gdiplus" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long

'GDI generic object management
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'GDI DC settings
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextAlign Lib "gdi32" (ByVal hDC As Long, ByVal wFlags As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long

'GDI font rendering functions
Private Declare Function ExtTextOutNoRect Lib "gdi32" Alias "ExtTextOutW" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal wOptions As Long, ByVal fakeRectPointer As Long, ByVal lpStrPointer As Long, ByVal nCount As Long, ByRef lpDx As Long) As Long
Private Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutW" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal wOptions As Long, ByRef lpRect As RECT, ByVal lpStrPointer As Long, ByVal nCount As Long, ByRef lpDx As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32W" (ByVal hDC As Long, ByVal lpStrPointer As Long, ByVal cbString As Long, ByRef lpSize As POINTAPI) As Long

'GDI text alignment flags
Private Const TA_LEFT = 0
Private Const TA_RIGHT = 2
Private Const TA_CENTER = 6

Private Const TA_TOP = 0
Private Const TA_BOTTOM = 8
Private Const TA_BASELINE = 24

Private Const TA_UPDATECP = 1
Private Const TA_NOUPDATECP = 0

'Back color modes (not useful here except during debug mode)
Private Const FONT_TRANSPARENT = &H1
Private Const FONT_OPAQUE = &H2

'Formatting constants for DrawText
Private Const DT_TOP = &H0
Private Const DT_LEFT = &H0
Private Const DT_CENTER = &H1
Private Const DT_RIGHT = &H2
Private Const DT_VCENTER = &H4
Private Const DT_BOTTOM = &H8
Private Const DT_WORDBREAK = &H10
Private Const DT_SINGLELINE = &H20
Private Const DT_EXPANDTABS = &H40
Private Const DT_TABSTOP = &H80
Private Const DT_NOCLIP = &H100
Private Const DT_EXTERNALLEADING = &H200
Private Const DT_CALCRECT = &H400
Private Const DT_NOPREFIX = &H800
Private Const DT_INTERNAL = &H1000
Private Const DT_EDITCONTROL = &H2000
Private Const DT_PATH_ELLIPSIS = &H4000
Private Const DT_END_ELLIPSIS = &H8000
Private Const DT_MODIFYSTRING = &H10000
Private Const DT_RTLREADING = &H20000
Private Const DT_WORD_ELLIPSIS = &H40000

'The current text-to-be-rendered
Private m_CurrentText As String

'This class internally maintains various font properties.  These properties must be converted into specific GDI+ parameters via
' various means, but they are stored in human-friendly format to simplify serializing a class instance to an XML string.
Private m_FontFace As String
Private m_FontColor As Long
Private m_FontSize As Single
Private m_FontSizeUnit As pdFontUnit
Private m_FontBold As Boolean
Private m_FontItalic As Boolean
Private m_FontUnderline As Boolean
Private m_FontStrikeout As Boolean

'Not all fonts support all subtypes (bold, italic, strikethrough).  When a font is loaded, this class will test for
' available subtypes automatically.
Private m_BoldSupported As Boolean
Private m_ItalicSupported As Boolean
Private m_UnderlineSupported As Boolean
Private m_StrikeoutSupported As Boolean

'Some string settings are not stored in the font itself, but in a GDI+ object called "StringFormat".  These are similar
' to per-DC settings in GDI (e.g. string alignment).  Because it is cheap to modify a StringFormat object, PD automatically
' syncs these internal values to a GDI+ StringFormat handle whenever they are changed.  This spares us from needing to do
' it during rendering stages.
Private m_HorizontalAlignment As GdiPlusStringAlignment
Private m_VerticalAlignment As GdiPlusStringAlignment

'Even *more* string settings are not stored in the font itself, or in a StringFormat object, but in the target
' GDI+ Graphics container.  These must be assigned to the graphics container prior to painting text, so there's not really
' an easy way to cache these settings.  (We could keep a temporary graphics container on hand, but we still have to clone
' it prior to rendering, so I doubt there's any gain there.)  As such, these values are not relevant until text is
' actually painted onto a target object.
Private m_TextAntialiasing As pdTextAA
Private m_TextContrast As Long

'The per-glyph PHOTODEMON renderer supports a bunch of unique options, including tons of appearance-related ones
Private m_TextHinting As Boolean
Private m_WordWrap As TEXT_WORDWRAP_MODE
Private m_FillActive As Boolean
Private m_FillBrush As String
Private m_OutlineActive As Boolean
Private m_OutlinePen As String
Private m_BackgroundActive As Boolean
Private m_BackgroundBrush As String
Private m_BackBorderActive As Boolean
Private m_BackBorderPen As String

'...as well as customized spacing and layout options
Private m_LineSpacing As Single
Private m_MarginLeft As Single
Private m_MarginTop As Single
Private m_MarginRight As Single
Private m_MarginBottom As Single

'...and a bunch of character-level modifications
Private m_CharRemap As Long
Private m_CharSpacing As Double
Private m_CharOrientation As Double
Private m_CharJitterX As Double
Private m_CharJitterY As Double
Private m_CharInflation As Double
Private m_CharMirror As Long

'For performance reasons, this class caches various font objects and handles from both GDI and GDI+.  This spares us from
' having to recreate expensive font data during rendering steps.
Private m_GDIPlusFont As Long
Private m_GDIPlusFontFamily As Long
Private m_GDIPlusStringFormat As Long
Private m_GDIFont As Long

'If we have to fallback to GDI, we can't actually use GDI for rendering.  Instead, we use it to generate glyph data, which we
' then retrieve, translate into a usable format, and rendering ourselves.  (Yes, it's a major PITA.)
Private m_GlyphCollection As pdGlyphCollection

'If a font object has been created, and a setting has been changed (e.g. font name, size, etc), we must recreate the font.
' All relevant property changes will set this value to FALSE to signify a re-cache is required.
Private m_FontCacheClean As Boolean

'If we have to fall back to plain GDI for rendering, we use a temporary 24-bpp DIB to cache the results.  This is then manually
' transferred to the 32-bpp target.
Private m_TmpGdiDIB As pdDIB

'Get/set the current rendering string.  PD uses a lot of tricks to improve rendering performance, including ignoring redraw requests if none of the
' class's internals have changed, so the current text must be assigned via these functions prior to actual rendering.
Public Function GetCurrentText() As String
    GetCurrentText = m_CurrentText
End Function

Public Sub SetCurrentText(ByRef newText As String)
    m_CurrentText = newText
End Sub

'Get functions for various font styles, and whether said styles are supported by the current font.
' (If no font name has been set, the function returns *will not be valid*)
Public Function GetFontColor() As Long
    GetFontColor = m_FontColor
End Function

Public Sub SetFontColor(ByVal newFontColor As Long)
    m_FontColor = newFontColor
End Sub

Public Function GetFontBold() As Boolean
    GetFontBold = m_FontBold
End Function

Public Function IsFontBoldSupported() As Boolean
    IsFontBoldSupported = m_BoldSupported
End Function

Public Function GetFontItalic() As Boolean
    GetFontItalic = m_FontItalic
End Function

Public Function IsFontItalicSupported() As Boolean
    IsFontItalicSupported = m_ItalicSupported
End Function

Public Function GetFontUnderline() As Boolean
    GetFontUnderline = m_FontUnderline
End Function

Public Function IsFontUnderlineSupported() As Boolean
    IsFontUnderlineSupported = m_UnderlineSupported
End Function

Public Function GetFontStrikeout() As Boolean
    GetFontStrikeout = m_FontStrikeout
End Function

Public Function IsFontStrikeoutSupported() As Boolean
    IsFontStrikeoutSupported = m_StrikeoutSupported
End Function

'Set functions for various font styles.  Note that these functions reset the current font cache, so please use them
' only if necessary.
Public Sub SetFontBold(ByVal newValue As Boolean)
    If newValue <> m_FontBold Then
        m_FontBold = newValue
        m_FontCacheClean = False
    End If
End Sub

Public Sub SetFontItalic(ByVal newValue As Boolean)
    If newValue <> m_FontItalic Then
        m_FontItalic = newValue
        m_FontCacheClean = False
    End If
End Sub

Public Sub SetFontUnderline(ByVal newValue As Boolean)
    If newValue <> m_FontUnderline Then
        m_FontUnderline = newValue
        m_FontCacheClean = False
    End If
End Sub

Public Sub SetFontStrikeout(ByVal newValue As Boolean)
    If newValue <> m_FontStrikeout Then
        m_FontStrikeout = newValue
        m_FontCacheClean = False
    End If
End Sub

'Get/set functions for font size and unit.  By default, this class uses pixels.  Other units may be supported
' in the future.
Public Function GetFontSize() As Single
    GetFontSize = m_FontSize
End Function

Public Sub SetFontSize(ByVal newValue As Single)
    If newValue <> m_FontSize Then
        m_FontSize = newValue
        m_FontCacheClean = False
    End If
End Sub

Friend Function GetFontSizeUnit() As pdFontUnit
    GetFontSizeUnit = m_FontSizeUnit
End Function

Friend Sub SetFontSizeUnit(ByVal newUnit As pdFontUnit)
    If newUnit <> m_FontSizeUnit Then
        m_FontSizeUnit = newUnit
        m_FontCacheClean = False
    End If
End Sub

'Get/set the text rendering engine.  This class currently supports two engines: WAPI (which will use GDI or GDI+, as appropriate), and PHOTODEMON,
' which uses GDI and Uniscribe to retrieve individual glyphs, but then does all positioning and rendering manually.
'
'Note that this setting does not initiate a font refresh, as PD generates a GDI font copy by default, regardless of rendering engine or font backend.
Public Function GetRenderingEngine() As TEXT_RENDERING_ENGINE
    GetRenderingEngine = m_RenderingEngine
End Function

Public Sub SetRenderingEngine(ByVal newEngine As TEXT_RENDERING_ENGINE)
    m_RenderingEngine = newEngine
End Sub

'TODO: XML get/set functions for getting/setting all parameters at once
Public Function GetAllFontSettingsAsXML() As String

    'We don't need a full-featured XML creator (e.g. with encoding decs and the like), so we can use the smaller, lighter pdParamXML class
    Dim cSerialize As pdParamXML
    Set cSerialize = New pdParamXML
    
    With cSerialize
        .Reset 1#
        .AddParam "TextLayerText", m_CurrentText
        .AddParam "TextLayerColor", m_FontColor
        .AddParam "TextLayerFontFace", m_FontFace
        .AddParam "TextLayerFontSize", m_FontSize
        .AddParam "TextLayerFontSizeUnit", m_FontSizeUnit
        .AddParam "TextLayerFontBold", m_FontBold
        .AddParam "TextLayerFontItalic", m_FontItalic
        .AddParam "TextLayerFontUnderline", m_FontUnderline
        .AddParam "TextLayerFontStrikeout", m_FontStrikeout
        .AddParam "TextLayerHorizontalAlignment", m_HorizontalAlignment
        .AddParam "TextLayerVerticalAlignment", m_VerticalAlignment
        .AddParam "TextLayerAntialiasing", m_TextAntialiasing
        .AddParam "TextLayerContrast", m_TextContrast
        .AddParam "TextRenderingEngine", m_RenderingEngine
        .AddParam "TextHinting", m_TextHinting
        .AddParam "TextWordWrap", m_WordWrap
        .AddParam "TextFillActive", m_FillActive
        .AddParam "TextFillBrush", m_FillBrush
        .AddParam "TextOutlineActive", m_OutlineActive
        .AddParam "TextOutlinePen", m_OutlinePen
        .AddParam "TextBackgroundActive", m_BackgroundActive
        .AddParam "TextBackgroundBrush", m_BackgroundBrush
        .AddParam "TextBackBorderActive", m_BackBorderActive
        .AddParam "TextBackBorderPen", m_BackBorderPen
        .AddParam "TextLineSpacing", m_LineSpacing
        .AddParam "TextMarginLeft", m_MarginLeft
        .AddParam "TextMarginTop", m_MarginTop
        .AddParam "TextMarginRight", m_MarginRight
        .AddParam "TextMarginBottom", m_MarginBottom
        .AddParam "TextCharRemap", m_CharRemap
        .AddParam "TextCharSpacing", m_CharSpacing
        .AddParam "TextCharOrientation", m_CharOrientation
        .AddParam "TextCharJitterX", m_CharJitterX
        .AddParam "TextCharJitterY", m_CharJitterY
        .AddParam "TextCharInflation", m_CharInflation
        .AddParam "TextCharMirror", m_CharMirror
    End With
    
    'Return the completed XML string
    GetAllFontSettingsAsXML = cSerialize.GetParamString()

End Function

'Set all text and font parameters at once, using an XML string created by a previous call to getAllFontSettingsAsXML.
' (Note that this function manually marks the font cache as dirty, so you should only use it when loading a file for the first time.)
Public Function SetAllFontSettingsFromXML(ByRef srcXMLString As String) As Boolean
    
    'We don't need a full-featured XML creator (e.g. with encoding decs and the like), so we can use the smaller, lighter pdParamXML class
    Dim cSerialize As pdParamXML
    Set cSerialize = New pdParamXML
    
    With cSerialize
        .SetParamString srcXMLString
        
        SetGenericTextProperty ptp_Text, .GetString("TextLayerText", "")
        SetGenericTextProperty ptp_FontColor, .GetLong("TextLayerColor", 0)
        SetGenericTextProperty ptp_FontFace, .GetString("TextLayerFontFace", g_InterfaceFont)
        SetGenericTextProperty ptp_FontSize, .GetDouble("TextLayerFontSize", 16#)
        SetGenericTextProperty ptp_FontSizeUnit, .GetLong("TextLayerFontSizeUnit", pdfu_Pixel)
        SetGenericTextProperty ptp_FontBold, .GetBool("TextLayerFontBold", False)
        SetGenericTextProperty ptp_FontItalic, .GetBool("TextLayerFontItalic", False)
        SetGenericTextProperty ptp_FontUnderline, .GetBool("TextLayerFontUnderline", False)
        SetGenericTextProperty ptp_FontStrikeout, .GetBool("TextLayerFontStrikeout", False)
        SetGenericTextProperty ptp_HorizontalAlignment, .GetLong("TextLayerHorizontalAlignment", StringAlignmentNear)
        SetGenericTextProperty ptp_VerticalAlignment, .GetLong("TextLayerVerticalAlignment", StringAlignmentNear)
        SetGenericTextProperty ptp_TextAntialiasing, .GetLong("TextLayerAntialiasing", pdta_Standard)
        SetGenericTextProperty ptp_TextContrast, .GetLong("TextLayerContrast", 5)
        SetGenericTextProperty ptp_RenderingEngine, .GetLong("TextRenderingEngine", tre_WAPI)
        SetGenericTextProperty ptp_TextHinting, .GetBool("TextHinting", False)
        SetGenericTextProperty ptp_WordWrap, .GetLong("TextWordWrap", tre_AutoBreakWords)
        SetGenericTextProperty ptp_FillActive, .GetBool("TextFillActive", True)
        SetGenericTextProperty ptp_FillBrush, .GetString("TextFillBrush", "")
        SetGenericTextProperty ptp_OutlineActive, .GetBool("TextOutlineActive", False)
        SetGenericTextProperty ptp_OutlinePen, .GetString("TextOutlinePen", "")
        SetGenericTextProperty ptp_BackgroundActive, .GetBool("TextBackgroundActive", False)
        SetGenericTextProperty ptp_BackgroundBrush, .GetString("TextBackgroundBrush", "")
        SetGenericTextProperty ptp_BackBorderActive, .GetBool("TextBackBorderActive", False)
        SetGenericTextProperty ptp_BackBorderPen, .GetString("TextBackBorderPen", "")
        SetGenericTextProperty ptp_LineSpacing, .GetDouble("TextLineSpacing", 0)
        SetGenericTextProperty ptp_MarginLeft, .GetDouble("TextMarginLeft", 0)
        SetGenericTextProperty ptp_MarginTop, .GetDouble("TextMarginTop", 0)
        SetGenericTextProperty ptp_MarginRight, .GetDouble("TextMarginRight", 0)
        SetGenericTextProperty ptp_MarginBottom, .GetDouble("TextMarginBottom", 0)
        SetGenericTextProperty ptp_CharRemap, .GetLong("TextCharRemap", 0)
        SetGenericTextProperty ptp_CharSpacing, .GetDouble("TextCharSpacing", 0)
        SetGenericTextProperty ptp_CharOrientation, .GetDouble("TextCharOrientation", 0)
        SetGenericTextProperty ptp_CharJitterX, .GetDouble("TextCharJitterX", 0)
        SetGenericTextProperty ptp_CharJitterY, .GetDouble("TextCharJitterY", 0)
        SetGenericTextProperty ptp_CharInflation, .GetDouble("TextCharInflation", 0)
        SetGenericTextProperty ptp_CharMirror, .GetLong("TextCharMirror", 0)
    End With
    
    'This function does not currently provide a fail state; as long as the load request comes from PD herself, failure should be impossible.
    SetAllFontSettingsFromXML = True
    
End Function

'Rather than using dedicated get functions or properties, you can retrieve font and text values via this universal function.
Friend Function GetGenericTextProperty(ByVal desiredProperty As PD_TEXT_PROPERTY) As Variant

    Select Case desiredProperty
        
        Case ptp_Text
            GetGenericTextProperty = m_CurrentText
        
        Case ptp_FontColor
            GetGenericTextProperty = m_FontColor
        
        Case ptp_FontFace
            GetGenericTextProperty = m_FontFace
        
        Case ptp_FontSize
            GetGenericTextProperty = m_FontSize
        
        Case ptp_FontSizeUnit
            GetGenericTextProperty = m_FontSizeUnit
        
        Case ptp_FontBold
            GetGenericTextProperty = m_FontBold
        
        Case ptp_FontItalic
            GetGenericTextProperty = m_FontItalic
        
        Case ptp_FontUnderline
            GetGenericTextProperty = m_FontUnderline
        
        Case ptp_FontStrikeout
            GetGenericTextProperty = m_FontStrikeout
        
        Case ptp_HorizontalAlignment
            GetGenericTextProperty = m_HorizontalAlignment
        
        Case ptp_VerticalAlignment
            GetGenericTextProperty = m_VerticalAlignment
        
        Case ptp_TextAntialiasing
            GetGenericTextProperty = m_TextAntialiasing
        
        Case ptp_TextContrast
            GetGenericTextProperty = m_TextContrast
            
        Case ptp_RenderingEngine
            GetGenericTextProperty = m_RenderingEngine
            
        Case ptp_TextHinting
            GetGenericTextProperty = m_TextHinting
            
        Case ptp_WordWrap
            GetGenericTextProperty = m_WordWrap
            
        Case ptp_FillActive
            GetGenericTextProperty = m_FillActive
        
        Case ptp_FillBrush
            GetGenericTextProperty = m_FillBrush
        
        Case ptp_OutlineActive
            GetGenericTextProperty = m_OutlineActive
        
        Case ptp_OutlinePen
            GetGenericTextProperty = m_OutlinePen
        
        Case ptp_BackgroundActive
            GetGenericTextProperty = m_BackgroundActive
        
        Case ptp_BackgroundBrush
            GetGenericTextProperty = m_BackgroundBrush
        
        Case ptp_BackBorderActive
            GetGenericTextProperty = m_BackBorderActive
            
        Case ptp_BackBorderPen
            GetGenericTextProperty = m_BackBorderPen
        
        Case ptp_LineSpacing
            GetGenericTextProperty = m_LineSpacing
            
        Case ptp_MarginLeft
            GetGenericTextProperty = m_MarginLeft
            
        Case ptp_MarginTop
            GetGenericTextProperty = m_MarginTop
            
        Case ptp_MarginRight
            GetGenericTextProperty = m_MarginRight
            
        Case ptp_MarginBottom
            GetGenericTextProperty = m_MarginBottom
            
        Case ptp_CharRemap
            GetGenericTextProperty = m_CharRemap
            
        Case ptp_CharSpacing
            GetGenericTextProperty = m_CharSpacing
            
        Case ptp_CharOrientation
            GetGenericTextProperty = m_CharOrientation
            
        Case ptp_CharJitterX
            GetGenericTextProperty = m_CharJitterX
            
        Case ptp_CharJitterY
            GetGenericTextProperty = m_CharJitterY
            
        Case ptp_CharInflation
            GetGenericTextProperty = m_CharInflation
            
        Case ptp_CharMirror
            GetGenericTextProperty = m_CharMirror
    
    End Select

End Function


'Rather than using dedicated set functions or properties, you can modify values via this universal function.
' Returns: TRUE if a value was successfully changed; FALSE if the requested value change was meaningless.  The caller can use this return
'           to know whether the associated text needs to be redrawn.
Friend Function SetGenericTextProperty(ByVal desiredProperty As PD_TEXT_PROPERTY, ByVal newValue As Variant) As Boolean

    Select Case desiredProperty
        
        Case ptp_Text
            If Strings.StringsNotEqual(m_CurrentText, CStr(newValue), False) Then
                m_CurrentText = newValue
                SetGenericTextProperty = True
            End If
        
        Case ptp_FontColor
            If m_FontColor <> CLng(newValue) Then
                m_FontColor = CLng(newValue)
                SetGenericTextProperty = True
            End If
        
        'TODO: consider raising some kind of error or message if the listed font is not available.  This is a possibility for PDI images
        ' traded between PCs.
        Case ptp_FontFace
            If Strings.StringsNotEqual(m_FontFace, CStr(newValue), True) Then
                m_FontFace = newValue
                m_FontCacheClean = False
                m_CurrentBackend = fb_UNKNOWN
                SetGenericTextProperty = True
            End If
        
        Case ptp_FontSize
            If m_FontSize <> CSng(newValue) Then
                m_FontSize = CSng(newValue)
                m_FontCacheClean = False
                SetGenericTextProperty = True
            End If
        
        Case ptp_FontSizeUnit
            If m_FontSizeUnit <> CLng(newValue) Then
                m_FontSizeUnit = CLng(newValue)
                m_FontCacheClean = False
                SetGenericTextProperty = True
            End If
        
        Case ptp_FontBold
            If m_FontBold <> CBool(newValue) Then
                m_FontBold = CBool(newValue)
                m_FontCacheClean = False
                SetGenericTextProperty = True
            End If
        
        Case ptp_FontItalic
            If m_FontItalic <> CBool(newValue) Then
                m_FontItalic = CBool(newValue)
                m_FontCacheClean = False
                SetGenericTextProperty = True
            End If
        
        Case ptp_FontUnderline
            If m_FontUnderline <> CBool(newValue) Then
                m_FontUnderline = CBool(newValue)
                m_FontCacheClean = False
                SetGenericTextProperty = True
            End If
        
        Case ptp_FontStrikeout
            If m_FontStrikeout <> CBool(newValue) Then
                m_FontStrikeout = CBool(newValue)
                m_FontCacheClean = False
                SetGenericTextProperty = True
            End If
        
        'When using GDI+ as our text layout engine, we set alignment changes immediately, as it's inexpensive to do so
        Case ptp_HorizontalAlignment
            If m_HorizontalAlignment <> CLng(newValue) Then
                m_HorizontalAlignment = CLng(newValue)
                SetGenericTextProperty = True
                If (m_GDIPlusStringFormat <> 0) Then GdipSetStringFormatAlign m_GDIPlusStringFormat, m_HorizontalAlignment
            End If
        
        Case ptp_VerticalAlignment
            If m_VerticalAlignment <> CLng(newValue) Then
                m_VerticalAlignment = CLng(newValue)
                SetGenericTextProperty = True
                If (m_GDIPlusStringFormat <> 0) Then GdipSetStringFormatLineAlign m_GDIPlusStringFormat, m_VerticalAlignment
            End If
        
        Case ptp_TextAntialiasing
            If m_TextAntialiasing <> CLng(newValue) Then
                m_TextAntialiasing = CLng(newValue)
                SetGenericTextProperty = True
                
                'If we're using GDI for rendering, changing antialiasing requires us to recreate the font
                If m_CurrentBackend = fb_GDI Then m_FontCacheClean = False
                
            End If
        
        'Text contrast is a weird one: the values vary from 0 to 12, and 4 is the default.  In PD, I have modified it to use a 0-10 scale with
        ' 3 as the default, and we convert it to a 12-based measurement if working with GDI+ specifically.  Note that this *ONLY WORKS IN
        ' BASIC RENDERING MODE.  GDI doesn't support this measurement natively, so we have to rig it using a manual system.  PD's manual glyph
        ' rendering methods ignore this setting entirely.
        Case ptp_TextContrast
            If m_TextContrast <> CLng(newValue) Then
                m_TextContrast = CLng(newValue)
                
                If (m_TextContrast < 0) Then
                    m_TextContrast = 0
                ElseIf (m_TextContrast > 10) Then
                    m_TextContrast = 10
                End If
                
                SetGenericTextProperty = True
            End If
            
        Case ptp_RenderingEngine
            If m_RenderingEngine <> CLng(newValue) Then
                m_RenderingEngine = CLng(newValue)
                SetGenericTextProperty = True
            End If
            
        Case ptp_TextHinting
            If m_TextHinting <> CBool(newValue) Then
                m_TextHinting = CBool(newValue)
                m_FontCacheClean = False
                SetGenericTextProperty = True
            End If
            
        Case ptp_WordWrap
            If m_WordWrap <> CLng(newValue) Then
                m_WordWrap = CLng(newValue)
                SetGenericTextProperty = True
            End If
            
        Case ptp_FillActive
            If m_FillActive <> CBool(newValue) Then
                m_FillActive = CBool(newValue)
                SetGenericTextProperty = True
            End If
        
        Case ptp_FillBrush
            If m_FillBrush <> newValue Then
                m_FillBrush = newValue
                SetGenericTextProperty = True
            End If
        
        Case ptp_OutlineActive
            If m_OutlineActive <> CBool(newValue) Then
                m_OutlineActive = CBool(newValue)
                SetGenericTextProperty = True
            End If
        
        Case ptp_OutlinePen
            If m_OutlinePen <> newValue Then
                m_OutlinePen = newValue
                SetGenericTextProperty = True
            End If
        
        Case ptp_BackgroundActive
            If m_BackgroundActive <> CBool(newValue) Then
                m_BackgroundActive = CBool(newValue)
                SetGenericTextProperty = True
            End If
        
        Case ptp_BackgroundBrush
            If m_BackgroundBrush <> newValue Then
                m_BackgroundBrush = newValue
                SetGenericTextProperty = True
            End If
        
        Case ptp_BackBorderActive
            If m_BackBorderActive <> CBool(newValue) Then
                m_BackBorderActive = CBool(newValue)
                SetGenericTextProperty = True
            End If
        
        Case ptp_BackBorderPen
            If m_BackBorderPen <> newValue Then
                m_BackBorderPen = newValue
                SetGenericTextProperty = True
            End If
        
        Case ptp_LineSpacing
            If m_LineSpacing <> CSng(newValue) Then
                m_LineSpacing = CSng(newValue)
                SetGenericTextProperty = True
            End If
        
        Case ptp_MarginLeft
            If m_MarginLeft <> CSng(newValue) Then
                m_MarginLeft = CSng(newValue)
                SetGenericTextProperty = True
            End If
        
        Case ptp_MarginTop
            If m_MarginTop <> CSng(newValue) Then
                m_MarginTop = CSng(newValue)
                SetGenericTextProperty = True
            End If
        
        Case ptp_MarginRight
            If m_MarginRight <> CSng(newValue) Then
                m_MarginRight = CSng(newValue)
                SetGenericTextProperty = True
            End If
        
        Case ptp_MarginBottom
            If m_MarginBottom <> CSng(newValue) Then
                m_MarginBottom = CSng(newValue)
                SetGenericTextProperty = True
            End If
            
        Case ptp_CharRemap
            If m_CharRemap <> CLng(newValue) Then
                m_CharRemap = CLng(newValue)
                SetGenericTextProperty = True
            End If
            
        Case ptp_CharSpacing
            If m_CharSpacing <> CDbl(newValue) Then
                m_CharSpacing = CDbl(newValue)
                SetGenericTextProperty = True
            End If
            
        Case ptp_CharOrientation
            If m_CharOrientation <> CDbl(newValue) Then
                m_CharOrientation = CDbl(newValue)
                SetGenericTextProperty = True
            End If
            
        Case ptp_CharJitterX
            If m_CharJitterX <> CDbl(newValue) Then
                m_CharJitterX = CDbl(newValue)
                SetGenericTextProperty = True
            End If
            
        Case ptp_CharJitterY
            If m_CharJitterY <> CDbl(newValue) Then
                m_CharJitterY = CDbl(newValue)
                SetGenericTextProperty = True
            End If
            
        Case ptp_CharInflation
            If m_CharInflation <> CDbl(newValue) Then
                m_CharInflation = CDbl(newValue)
                SetGenericTextProperty = True
            End If
            
        Case ptp_CharMirror
            If m_CharMirror <> CLng(newValue) Then
                m_CharMirror = CLng(newValue)
                SetGenericTextProperty = True
            End If
        
        Case Else
            Debug.Print "WARNING!  Unknown text property requested from pdTextRenderer.setGenericTextProperty()"
        
    End Select
    
End Function

'Get/Set font face.  These are more cumbersome than other font settings.
Public Function GetFontFace() As String
    GetFontFace = m_FontFace
End Function

'Set a given font face, as specified by a font name.
Public Sub SetFontFace(ByVal fontName As String)
    
    'If this font face has already been set, ignore this request
    If Strings.StringsNotEqual(fontName, m_FontFace, False) Then
        m_FontFace = fontName
        m_FontCacheClean = False
        m_CurrentBackend = fb_UNKNOWN
    End If
    
End Sub

'Convert a pdFontUnit enum to a GDI+ format measurement unit.  Note that not all GDI+ units are reachable this way, by design.
Private Function ConvertPDFontUnitToGDIPlusFontUnit(ByRef srcPDFontUnit As pdFontUnit) As GdiPlusMeasurementUnit
    
    Select Case srcPDFontUnit
    
        Case pdfu_Pixel
            ConvertPDFontUnitToGDIPlusFontUnit = gpUnitPixel
        
        Case pdfu_Point
            ConvertPDFontUnitToGDIPlusFontUnit = gpUnitPoint
            
    End Select
    
End Function

'When all font parameters have been set, we can actually create a font!
Public Function CreateFontObject() As Boolean
    
    'If the font cache is clean, this request is redundant; ignore it
    If m_FontCacheClean Then
        Debug.Print "Font cache is clean; ignoring CreateFontObject request"
        CreateFontObject = True
        Exit Function
    End If
    
    'Reset the current text engine marker, as we won't know which engine is required until we try to create the font.
    m_CurrentBackend = fb_UNKNOWN
    
    'If we have already created any font objects (in any engine), free them now
    If m_GDIPlusFont <> 0 Then
        GdipDeleteFont m_GDIPlusFont
        m_GDIPlusFont = 0
    End If
    
    If m_GDIPlusFontFamily <> 0 Then
        GdipDeleteFontFamily m_GDIPlusFontFamily
        m_GDIPlusFontFamily = 0
    End If
    
    If m_GDIFont <> 0 Then
        DeleteObject m_GDIFont
        m_GDIFont = 0
    End If
    
    'First, try to create a GDI+ font family.  This saves us a lot of grief.
    ' (Note that this step is skipped if manual rendering mode is activated; in that case, we will only use GDI fonts)
    Dim fontCreationSuccess As Boolean
    fontCreationSuccess = False
    
    If m_RenderingEngine = tre_WAPI Then fontCreationSuccess = CreateGDIPlusFont()
    
    'GDI+ reported success, and we obtained a valid GDI+ font handle.
    If fontCreationSuccess And (m_GDIPlusFont <> 0) Then
        
        'Note that GDI+ works fine for this particular combination of settings, and return success.
        Debug.Print "GDI+ is now set as the active font handling backend."
        m_CurrentBackend = fb_GDIPLUS
        m_FontCacheClean = True
        CreateFontObject = True
        
        'For text metric purposes, we create a duplicate font in GDI format.  This is a hassle, but it's necessary to retrieve
        ' detailed metrics on individual glyphs, as GDI+ doesn't expose very detailed font information.
        CreateGDIFont
        
    'If GDI+ failed, fall back to GDI and try again.
    Else
        
        fontCreationSuccess = CreateGDIFont()
        
        'If GDI is successful...
        If fontCreationSuccess And (m_GDIFont <> 0) Then
        
            'Note that GDI is required for this combination of settings
            Debug.Print "GDI is now set as the active font handling backend."
            m_CurrentBackend = fb_GDI
            m_FontCacheClean = True
            CreateFontObject = True
            
        Else
        
            'For some reason, this combination of settings doesn't work at all.  I hope this never happens, but you never know.
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "WARNING!  pdTextRenderer's CreateFontObject failed with both GDI+ and GDI; sorry!"
            #End If
            
            m_CurrentBackend = fb_UNKNOWN
            m_FontCacheClean = False
            CreateFontObject = False
            
        End If
        
    End If
        
End Function

'Attempt to create a GDI+ font matching the current class settings.  Note that GDI+ creates two relevant objects: font family, and font itself
Private Function CreateGDIPlusFont() As Boolean
    
    Debug.Print "Attempting to create GDI+ font object matching the current pdTextRenderer settings..."
    
    'If a GDI+ font family already exists, free it
    If m_GDIPlusFontFamily <> 0 Then GdipDeleteFontFamily m_GDIPlusFontFamily
    
    'Attempt to retrieve a GDI+ font family matching this name
    Dim gdipReturn As Long
    gdipReturn = GdipCreateFontFamilyFromName(StrPtr(m_FontFace), 0&, m_GDIPlusFontFamily)
    
    'If GDI+ reported success and a non-zero handle was obtained, carry on!
    If (gdipReturn = 0) And (m_GDIPlusFontFamily <> 0) Then
        
        Debug.Print "It's a miracle!  GDI+ font family created successfully."
        
        'As a convenience for future efforts, cache the font styles supported by this font family
        Dim testResult As Long
        GdipIsStyleAvailable m_GDIPlusFontFamily, FontStyleBold, testResult
        m_BoldSupported = CBool(testResult <> 0)
        
        GdipIsStyleAvailable m_GDIPlusFontFamily, FontStyleItalic, testResult
        m_ItalicSupported = CBool(testResult <> 0)
        
        GdipIsStyleAvailable m_GDIPlusFontFamily, FontStyleUnderline, testResult
        m_UnderlineSupported = CBool(testResult <> 0)
        
        GdipIsStyleAvailable m_GDIPlusFontFamily, FontStyleStrikeout, testResult
        m_StrikeoutSupported = CBool(testResult <> 0)
        
        'Next, we will use the family handle to create an actual font; this includes any font styles (bold, italic, etc) and the
        ' current font size.
        gdipReturn = GdipCreateFont(m_GDIPlusFontFamily, m_FontSize, GetFontStylesAsGdipLong, ConvertPDFontUnitToGDIPlusFontUnit(m_FontSizeUnit), m_GDIPlusFont)
        
        'Check for known errors; GDI+ cannot synthesize some font sizes, for example
        If (gdipReturn = 0) And (m_GDIPlusFont <> 0) Then
        
            'Font creation successful.  Set the font cache flag, so we know not to create the font again.
            Debug.Print "WOW, two miracles in a row!  GDI+ font object created successfully."
            
            'Return success
            CreateGDIPlusFont = True
            
        Else
        
            'Font creation failure.  Not much we can do now.
            Debug.Print "WARNING!  GDI+ failed to create a font object based on this font family."
            ReportGDIPlusFailure gdipReturn
            
            'Delete the font family we created as well, then exit
            If m_GDIPlusFontFamily <> 0 Then GdipDeleteFontFamily m_GDIPlusFontFamily
            CreateGDIPlusFont = False
            
        End If
        
    'GDI+ font selection can fail for a variety of reasons; report the failure reason and exit immediately
    Else
        
        Debug.Print "GDI+ font family could not be created; font is probably OpenType."
        ReportGDIPlusFailure gdipReturn
        CreateGDIPlusFont = False
        
    End If
    
End Function

'Convert the current arrangement of font styles into a Long-type value compatible with GDI+'s style declarations
Private Function GetFontStylesAsGdipLong() As Long
    
    GetFontStylesAsGdipLong = 0
    
    If m_FontBold Then GetFontStylesAsGdipLong = (GetFontStylesAsGdipLong Or FontStyleBold)
    If m_FontItalic Then GetFontStylesAsGdipLong = (GetFontStylesAsGdipLong Or FontStyleItalic)
    If m_FontUnderline Then GetFontStylesAsGdipLong = (GetFontStylesAsGdipLong Or FontStyleUnderline)
    If m_FontStrikeout Then GetFontStylesAsGdipLong = (GetFontStylesAsGdipLong Or FontStyleStrikeout)
    
End Function

'Convert the current PD-specific text rendering hint into a Long-type value compatible with GDI+'s text declarations
Private Function ConvertTextAAToGdipTextHint() As GdiPlusTextRenderingHint
    
    Select Case m_TextAntialiasing
    
        Case pdta_None
            ConvertTextAAToGdipTextHint = TextRenderingHintSingleBitPerPixel 'TextRenderingHintSingleBitPerPixelGridFit
        
        Case pdta_Standard
            ConvertTextAAToGdipTextHint = TextRenderingHintAntiAlias
        
        Case pdta_Crisp
            ConvertTextAAToGdipTextHint = TextRenderingHintAntiAliasGridFit
    
    End Select
    
End Function

'If a GDI+ function fails, pass the return value to this function so the debugger can attempt to output a helpful message.
Private Sub ReportGDIPlusFailure(ByVal failCode As Long)

    'For performance reasons, this sub is disabled in production builds.  GDI+ failures are expected and frequent, so my
    ' main interest in them is academic.
    #If DEBUGMODE = 1 Then
        
        Select Case failCode
        
            Case gdipFontFamilyNotFound
                pdDebug.LogAction "(GDI+ error code analysis: font family wasn't located)"
                
            Case gdipFontStyleNotFound
                pdDebug.LogAction "(GDI+ error code analysis: requested font style wasn't found)"
                
            Case gdipNotTrueTypeFont
                pdDebug.LogAction "(GDI+ error code analysis: requested font is not a TrueType font)"
                
            Case Else
                pdDebug.LogAction "(GDI+ failed for an unknown reason: #" & failCode & ")"
        
        End Select
        
    #End If
    
End Sub

'Attempt to create a GDI font matching the current class settings
Private Function CreateGDIFont() As Boolean
    
    Debug.Print "Attempting to create GDI font object matching the current pdTextRenderer settings..."
    
    'If a GDI font handle already exists, free it.  Note that the font must have been unselected from a DC; this is typically handled
    ' by the renderer function itself.
    If m_GDIFont <> 0 Then DeleteObject m_GDIFont
    
    'The font management module makes this task a lot easier.  Start by creating a LOGFONTW container.
    Dim tmpLogFont As LOGFONTW
    
    'Use the font management module to populate the struct
    Fonts.FillLogFontW_Basic tmpLogFont, m_FontFace, m_FontBold, m_FontItalic, m_FontUnderline, m_FontStrikeout
    Fonts.FillLogFontW_Size tmpLogFont, m_FontSize, m_FontSizeUnit
    Fonts.FillLogFontW_Quality tmpLogFont, ConvertTextAAToGdipTextHint
    
    'Attempt to create a matching GDI font
    If Fonts.CreateGDIFont(tmpLogFont, m_GDIFont) Then
        
        'Success!  Note that GDI is an acceptable engine for this font, and return success
        Debug.Print "GDI font creation worked!  Proceeding with rendering."
        CreateGDIFont = True
        
    Else
        
        'Failure!  PD doesn't have a fallback beyond GDI, so this whole request is doomed.
        ' (This shouldn't technically be possible, but if it is, I'd rather laugh than cry.)
        Debug.Print "GDI font creation failed miserably!  Everything is falling apart!  AAAAHHHH!"
        CreateGDIFont = False
    
    End If
    
End Function

'Use this function to render arbitrary text to an arbitrary DIB.  Performance should be decent, but if you're going to be making
' a ton of successive calls, it's probably worthwhile to cache a GDI+ graphics instance that points to the target.  (This class does
' not currently offer a function for that, FYI.)
'
'Note also that the crColor parameter should not include alpha; alpha in PD is handled by layers, not the text renderer.
Public Function RenderTextToDIB(ByRef dstDIB As pdDIB, ByVal x1 As Single, ByVal y1 As Single, ByVal textWidth As Single, ByVal textHeight As Single) As Boolean
    
    'As a convenience to the user, create the font as necessary
    If (Not m_FontCacheClean) Then
        Debug.Print "Font settings have changed; attempting to recreate font now..."
        CreateFontObject
    End If
    
    'If the font cache *still* isn't clean, something went horribly wrong
    If (Not m_FontCacheClean) Then
        Debug.Print "PD failed to create the requested font.  Abandoning render request."
        RenderTextToDIB = False
        Exit Function
    Else
        Debug.Print "Font object looks good.  Proceeding with rendering."
    End If
    
    'PD currently supports two rendering modes: WAPI (which uses GDI or GDI+ to render the text, and is thus limited by the settings
    ' those two engines provide), or an internal PHOTODEMON engine, which manually extracts font glyphs and renders them ourselves.
    If (m_RenderingEngine = tre_WAPI) Then
        
        'These backends have no reason to render if text isn't present
        If (Len(m_CurrentText) <> 0) Then
        
                'We should now know whether GDI or GDI+ is required for this combination of settings.
            Select Case m_CurrentBackend
                
                'GDI+ is sufficient
                Case fb_GDIPLUS
                    RenderTextToDIB = RenderTextToDIB_GDIPlus(dstDIB, m_CurrentText, m_FontColor, x1, y1, textWidth, textHeight)
                    
                    'GDI+ may fail when actually drawing (UGH).  If it does, silently fall back to GDI.
                    If (Not RenderTextToDIB) Then
                        Debug.Print "GDI+ text render failed; falling back to GDI renderer..."
                        RenderTextToDIB = RenderTextToDIB_GDI(dstDIB, m_CurrentText, m_FontColor, x1, y1, textWidth, textHeight)
                    End If
                    
                'GDI is required
                Case fb_GDI
                    RenderTextToDIB = RenderTextToDIB_GDI(dstDIB, m_CurrentText, m_FontColor, x1, y1, textWidth, textHeight)
                    
                'GDI and GDI+ both failed; at present, we have no fallback for this case
                Case fb_UNKNOWN
                    Debug.Print "GDI and GDI+ both failed; something is probably wrong with this font.  Consider blacklisting it...?"
                    RenderTextToDIB = False
                    
            End Select
            
            'Debug.Print "FYI - renderTextToDIB returned " & CStr(renderTextToDIB)
            
        End If
    
    ElseIf (m_RenderingEngine = tre_PHOTODEMON) Then
    
        'PD's internal font engine works by manually translating GDI glyphs into GDI+ paths, then rendering the paths manually.
        ' Because glyph geometry is very energy intensive, we use a separate class that caches glyphs as they are created.
        ' This saves us from having to translate glyphs more than once, which is a non-trivial boost on text with one or more
        ' repeat characters.
        
        'Start by creating a pdGlyphCollection instance, if one doesn't already exist.
        If (m_GlyphCollection Is Nothing) Then Set m_GlyphCollection = New pdGlyphCollection
        
        'Notify the glyph collection of our GDI font handle.  It obviously needs this in order to generate glyphs.
        m_GlyphCollection.NotifyOfGDIFontChoice m_GDIFont, m_FontUnderline, m_FontStrikeout
        
        'Next, pass the glyph collection our target string.  The glyph collection will generate a list of required glyphs and
        ' automatically retrieve any glyph geometry it hasn't already calculated.  (For example, if the last request it received
        ' was "cat", but this request is "cats", only the 's' needs to be newly retrieved.)
        '
        'Also - because this step may modify the string (via remapping settings), we use a temporary copy of the current string.
        Dim tmpString As String
        tmpString = m_CurrentText
        m_GlyphCollection.BuildGlyphCollection tmpString, m_TextHinting, m_CharRemap
        
        'After all glyphs have been created, it's *very important* that we regain control over our GDI font, as we need it for
        ' the actual rendering step!
        m_GlyphCollection.RequestGDIFontRelease
        
        'With the glyph collection successfully assembled, proceed with rendering
        RenderTextToDIB = RenderTextToDIB_Glyphs(dstDIB, tmpString, x1, y1, textWidth, textHeight)
    
    Else
        Debug.Print "WARNING!  Unknown rendering engine set in pdTextRenderer."
        RenderTextToDIB = False
    End If
    
End Function

'Render text using GDI+
Private Function RenderTextToDIB_GDIPlus(ByRef dstDIB As pdDIB, ByRef srcString As String, ByVal crColor As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal textWidth As Single, ByVal textHeight As Single) As Boolean

    'TESTING ONLY!  Fill with white prior to rendering
    'GDI_Plus.GDIPlusFillDIBRect dstDIB, 0, 0, dstDIB.getDIBWidth, dstDIB.getDIBHeight, vbWhite
    
    'Start by acquiring a GDI+ handle to the destination DIB
    Dim dstGraphics As Long
    dstGraphics = GDI_Plus.GetGDIPlusGraphicsFromDC(dstDIB.GetDIBDC, False)
    
    'Next, set up all container-specific text settings
    If (GdipSetTextRenderingHint(dstGraphics, ConvertTextAAToGdipTextHint) <> 0) Then
        Debug.Print "Failed to set text rendering hint " & m_TextAntialiasing
    End If
    
    'PD stores text contrast on a 0-10 scale, but GDI+ inexplicably uses a 0-12 scale.
    If (GdipSetTextContrast(dstGraphics, CLng(CDbl(m_TextContrast) * 1.2)) <> 0) Then
        Debug.Print "Failed to set text contrast " & m_TextContrast
    End If
    
    'New GDI+ definitions should look like this:
    If Not GDI_Plus.SetGDIPlusGraphicsProperty(dstGraphics, P2_SurfaceBlendUsingSRGBGamma, GP_CQ_AssumeLinear) Then Debug.Print "Failed to set compositing quality."
    If Not GDI_Plus.SetGDIPlusGraphicsProperty(dstGraphics, P2_SurfacePixelOffset, GP_POM_Half) Then Debug.Print "Failed to set pixel offset mode."
    
    'Create a solid fill brush.  (In the future, we'll add more impressive options!)
    Dim gdipBrush As Long
    gdipBrush = GDI_Plus.GetGDIPlusSolidBrushHandle(crColor, 255)
    
    'Convert the input rect to a RECTF
    Dim boundingRect As RECTF
    With boundingRect
        .Left = x1
        .Top = y1
        .Width = textWidth
        .Height = textHeight
    End With
    
    'Hypothetically, we could attempt to retrieve a bounding box for the string.  This would be relevant when compensating for overhang
    ' on the left side, for synthesized fonts like Times New Roman - Italic.  Unfortunately, GDI+ is stupid and it doesn't report
    ' left-bound overhang, so our internal function GetStringBoundingBox() isn't helpful at all.
    
    'We can, however, use GDI to calculate it instead.  (TODO: cache this value, and only re-check it if the string and/or font changes.)
    Dim firstCharABCWidth As ABCFLOAT
    If Fonts.GetABCWidthOfGlyph(m_GDIFont, AscW(Left$(srcString, 1)), firstCharABCWidth) Then
        
        'If the offset value is negative, we want to increase our rendering offset proportionally.  (Note that we add an
        ' extra pixel to account for possible antialiasing along the border.)
        If (firstCharABCWidth.abcfA < 0) Then
            
            Dim newLeft As Long, newWidth As Long
            newLeft = boundingRect.Left + Abs(firstCharABCWidth.abcfA) + 1
            boundingRect.Left = newLeft
            
            'Because fonts with a significant left overhang are likely to have a large right overhang as well, mirror the
            ' overhang to the right boundary as well.  (We do this because GDI+ word wrap is unpredictable, and we can't retrieve
            ' where it decides to wordwrap - so we have to make assumptions like this.)
            newWidth = boundingRect.Width - (Abs(firstCharABCWidth.abcfC) + 1)
            If (newWidth > 0) Then boundingRect.Width = newWidth
            
        End If
        
    End If
        
    'Finally, use GDI+ to render the actual string
    Dim gdipReturn As GP_Result
    gdipReturn = GdipDrawString(dstGraphics, StrPtr(srcString), Len(srcString), m_GDIPlusFont, boundingRect, m_GDIPlusStringFormat, gdipBrush)
    
    If (gdipReturn = GP_OK) Then
        RenderTextToDIB_GDIPlus = True
    Else
        Debug.Print "WARNING!  GdipDrawString failed with error code " & gdipReturn
        Debug.Print "WARNING!  (Extra debug info: hFont - " & m_GDIPlusFont & ", hStringFormat - " & m_GDIPlusStringFormat & ")"
        RenderTextToDIB_GDIPlus = False
    End If
    
    'Release the temporary GDI+ objects we created
    GDI_Plus.ReleaseGDIPlusBrush gdipBrush
    GDI_Plus.ReleaseGDIPlusGraphics dstGraphics

End Function

'Use GdipMeasureString to return the bounding box of a given string and assembled GDI+ objects.  This is helpful for knowing how the string
' fits into the desired target rect.
Private Function GetStringBoundingBox(ByRef srcGraphics As Long, ByRef srcString As String, ByRef srcLayoutRect As RECTF, ByRef dstBoundRect As RECTF) As Boolean

    If (m_GDIPlusFont <> 0) Then
        
        Dim gdipReturn As Long
        gdipReturn = GdipMeasureString(srcGraphics, StrPtr(srcString), Len(srcString), m_GDIPlusFont, srcLayoutRect, m_GDIPlusStringFormat, dstBoundRect, ByVal 0&, ByVal 0&)
        GetStringBoundingBox = CBool(gdipReturn = 0)
        
    Else
        GetStringBoundingBox = 0
    End If

End Function

'Render text using GDI+, and a glyph collection properly assembled inside m_GlyphCollection.
Private Function RenderTextToDIB_Glyphs(ByRef dstDIB As pdDIB, ByRef srcString As String, ByVal x1 As Single, ByVal y1 As Single, ByVal textWidth As Single, ByVal textHeight As Single) As Boolean

    'TESTING ONLY!  Fill with white prior to rendering
    'GDI_Plus.GDIPlusFillDIBRect dstDIB, 0, 0, dstDIB.getDIBWidth, dstDIB.getDIBHeight, vbWhite
    
    'Start by acquiring a GDI+ handle to the destination DIB
    Dim dstGraphics As Long
    dstGraphics = GDI_Plus.GetGDIPlusGraphicsFromDC(dstDIB.GetDIBDC, False)
    
    'Next, convert the input rect to a RECTF
    Dim boundingRect As RECTF
    With boundingRect
        .Left = x1
        .Top = y1
        .Width = textWidth
        .Height = textHeight
    End With
    
    'Next, set up all container-specific path settings
    Dim targetSmoothingMode As GdiPlusSmoothingMode, targetOffsetMode As GP_PixelOffsetMode
    
    'Compositing quality does not change; GDI+ is unreliable in its handling of compositing quality (which assumes gamma correction),
    ' and it can cause premultiplied alpha to be calculated incorrectly, leading to random hot pixels.
    GDI_Plus.SetGDIPlusGraphicsProperty dstGraphics, P2_SurfaceBlendUsingSRGBGamma, GP_CQ_AssumeLinear
    
    'Antialiasing settings currently affect both smoothing mode and offset mode.  Offset mode is particularly relevant when working with
    ' tiny text, as it greatly improves clarity to use half mode.  (Also, note that antialiasing affects *all drawing operations on the target*,
    ' by design, so you cannot currently set different AA options for text vs background elements.)
    Select Case m_TextAntialiasing
        
        Case pdta_None
            targetSmoothingMode = SmoothingModeHighSpeed
            targetOffsetMode = GP_POM_None
        
        Case pdta_Standard
            targetSmoothingMode = GP_SM_Antialias
            targetOffsetMode = GP_POM_None
        
        Case pdta_Crisp
            targetSmoothingMode = GP_SM_Antialias
            targetOffsetMode = GP_POM_Half
        
    End Select
    
    'Note that we don't actually set container settings yet, because we first want to draw some generic elements (like the background fill)
    ' at maximum speed.  Once those are finished, we'll set target smoothing and offset modes.
        
    'If fill mode is active, create a relevant brush.  Note that not all brushes are currently implemented!
    Dim gdipFillBrush As Long, textFillBrush As pd2DBrush, backgroundFillBrush As pd2DBrush
    
    If m_FillActive Then
        Set textFillBrush = New pd2DBrush
        textFillBrush.SetBrushPropertiesFromXML m_FillBrush
        gdipFillBrush = textFillBrush.GetHandle
    Else
        gdipFillBrush = 0
    End If
        
    'If outline mode is active, create a relevant pen.
    Dim gdipOutlinePen As Long, penTextOutline As pd2DPen, penBackgroundOutline As pd2DPen
    Set penTextOutline = New pd2DPen
    Set penBackgroundOutline = New pd2DPen
    
    If m_OutlineActive Then
        penTextOutline.SetPenPropertiesFromXML m_OutlinePen
        If penTextOutline.CreatePen() Then gdipOutlinePen = penTextOutline.GetHandle
    Else
        gdipOutlinePen = 0
    End If
    
    'If background mode is active, create a background-specific brush
    Dim gdipBackgroundBrush As Long
    
    If m_BackgroundActive Then
        Set backgroundFillBrush = New pd2DBrush
        backgroundFillBrush.SetBrushPropertiesFromXML m_BackgroundBrush
        backgroundFillBrush.SetBoundaryRect boundingRect
        gdipBackgroundBrush = backgroundFillBrush.GetHandle
    Else
        gdipBackgroundBrush = 0
    End If
        
    'If a background outline is active, create a pen for that, too
    Dim gdipBackBorderPen As Long
    
    If m_BackBorderActive Then
        penBackgroundOutline.SetPenPropertiesFromXML m_BackBorderPen
        If penBackgroundOutline.CreatePen() Then gdipBackBorderPen = penBackgroundOutline.GetHandle
    Else
        gdipBackBorderPen = 0
    End If
    
    'Next, modify the bounding rect according to a few different factors:
    
    '1) If outline mode is active, increase padding by 1/2 the outline size.  This ensures that the outline is not cropped.
    If m_OutlineActive Then
        
        With boundingRect
            .Left = .Left + penTextOutline.GetPenProperty(P2_PenWidth)
            .Top = .Top + penTextOutline.GetPenProperty(P2_PenWidth)
            .Width = .Width - penTextOutline.GetPenProperty(P2_PenWidth) * 2
            .Height = .Height - penTextOutline.GetPenProperty(P2_PenWidth) * 2
        End With
    
    End If
    
    '2) The caller can specify custom margin adjustments, either positive or negative.  Apply those now.
    With boundingRect
        .Left = .Left + m_MarginLeft
        .Top = .Top + m_MarginTop
        .Width = .Width - (m_MarginLeft + m_MarginRight)
        .Height = .Height - (m_MarginTop + m_MarginBottom)
    End With
    
    '3) If a background border is active, we further modify the bounding rect by the size of the border
    Dim borderWidth As Single, halfBorderWidth As Single
    
    If m_BackBorderActive Then
        
        borderWidth = penBackgroundOutline.GetPenProperty(P2_PenWidth)
        halfBorderWidth = borderWidth / 2
        
        With boundingRect
            .Left = .Left + halfBorderWidth
            .Top = .Top + halfBorderWidth
            .Width = .Width - borderWidth
            .Height = .Height - borderWidth
        End With
        
    End If
    
    '4) After all boundary rect modifications are made, we need to perform a failsafe check to make sure width and height aren't negative.
    With boundingRect
        
        If (.Width < 0) Then
            .Left = .Left + .Width
            .Width = 0
        End If
        
        If (.Height < 0) Then
            .Top = .Top + .Height
            .Height = 0
        End If
        
    End With
    
    'Retrieve a full, composite path from the glyph collector.  It handles messy business like character positioning,
    ' and simply hands us a fully composited path, ready for rendering.
    Dim finalTextPath As pd2DPath
    m_GlyphCollection.NotifyCustomLayoutSettings m_LineSpacing, m_CharSpacing, m_CharOrientation, m_CharJitterX, m_CharJitterY, m_CharInflation, m_CharMirror
    
    If m_GlyphCollection.AssembleCompositePath(finalTextPath, boundingRect, m_HorizontalAlignment, m_VerticalAlignment, m_WordWrap) Then
        
        Dim exactTextBounds As RECTF
        exactTextBounds = finalTextPath.GetPathBoundariesF
        
        Dim gdipReturn As Long
                
        'If background mode is active, fill it now
        If (gdipBackgroundBrush <> 0) Then
            
            'Prior to filling, set high-speed offset mode.
            GDI_Plus.SetGDIPlusGraphicsProperty dstGraphics, P2_SurfacePixelOffset, GP_POM_None
            
            'If an outline is being drawn, we don't need antialiasing, which also improves performance
            If (gdipBackBorderPen <> 0) Then
                GdipSetSmoothingMode dstGraphics, SmoothingModeHighSpeed
            Else
                GdipSetSmoothingMode dstGraphics, GP_SM_Antialias
            End If
            
            gdipReturn = GdipFillRectangleI(dstGraphics, gdipBackgroundBrush, x1, y1, textWidth, textHeight)
            If (gdipReturn <> 0) Then Debug.Print "WARNING!  GdipFillPath failed; return code was " & gdipReturn
            
        End If
        
        'Now we can set our desired smoothing and offset modes, which are constant for all remaining draw operations
        GdipSetSmoothingMode dstGraphics, targetSmoothingMode
        GDI_Plus.SetGDIPlusGraphicsProperty dstGraphics, P2_SurfacePixelOffset, targetOffsetMode
                
        'If background border is active, stroke it now
        If (gdipBackBorderPen <> 0) Then
            gdipReturn = GdipDrawRectangle(dstGraphics, gdipBackBorderPen, x1 + halfBorderWidth, y1 + halfBorderWidth, textWidth - borderWidth, textHeight - borderWidth)
            If gdipReturn <> 0 Then Debug.Print "WARNING!  GdipDrawRectangle failed; return code was " & gdipReturn
        End If
                
        'If fill mode is active, draw the text now
        If (gdipFillBrush <> 0) Then
        
            'Update the fill brush against the final rect established by the path renderer.
            ' (This additional step is necessary to ensure that gradients span the full width/height of the text.)
            textFillBrush.SetBrushPropertiesFromXML m_FillBrush
            boundingRect = finalTextPath.GetPathBoundariesF
            textFillBrush.SetBoundaryRect boundingRect
            gdipFillBrush = textFillBrush.GetHandle
            
            gdipReturn = GdipFillPath(dstGraphics, gdipFillBrush, finalTextPath.GetHandle)
            If gdipReturn <> 0 Then Debug.Print "WARNING!  GdipFillPath failed; return code was " & gdipReturn
        End If
        
        'If outline mode is active, stroke the text now
        If (gdipOutlinePen <> 0) Then
            gdipReturn = GdipDrawPath(dstGraphics, gdipOutlinePen, finalTextPath.GetHandle)
            If (gdipReturn <> 0) Then Debug.Print "WARNING!  GdipDrawPath failed; return code was " & gdipReturn
        End If
        
    Else
        Debug.Print "WARNING!  m_GlyphCollection.assembleCompositePath returned FALSE.  Please investigate."
    End If
    
    'Release the temporary GDI+ objects we created.  (Note that GDI+ objects silently created via pd2D objects are released automatically.)
    GDI_Plus.ReleaseGDIPlusGraphics dstGraphics
    
    'Return success/failure contingent on the return of the GDI+ rendering call
    RenderTextToDIB_Glyphs = CBool(gdipReturn = 0)
    
End Function

'Render text using GDI using a dummy implementation, only.  This only exists to test things like GDI font creation; I am actively working on
' superceding it with a much more advanced implementation.
Private Function RenderTextToDIB_GDI(ByRef dstDIB As pdDIB, ByRef srcString As String, ByVal crColor As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal textWidth As Single, ByVal textHeight As Single) As Boolean

    'Because GDI doesn't support 32-bpp rendering targets, we must do all rendering to a 24-bpp surface, then copy it over manually
    If (m_TmpGdiDIB Is Nothing) Then
        Set m_TmpGdiDIB = New pdDIB
        m_TmpGdiDIB.CreateBlank 4, 4, 24, 0
    End If
    
    'Start by setting up all container-specific text settings; note that these are only used in our temporary DIB!
    Dim oldFont As Long
    oldFont = SelectObject(m_TmpGdiDIB.GetDIBDC, m_GDIFont)
    
    'Set other font parameters.  Note that this function only draws white text onto a black background.  Color is applied later.
    SetTextColor m_TmpGdiDIB.GetDIBDC, RGB(255, 255, 255)
    SetTextAlign m_TmpGdiDIB.GetDIBDC, 0
    
    'Enable transparent font rendering
    SetBkMode m_TmpGdiDIB.GetDIBDC, FONT_TRANSPARENT
    
    'You'd think drawing opaque text to a 32-bpp container would work, but nope, this just means that GDI functions make the
    ' ENTIRE TEXT AREA transparent. THANKS, MICROSOFT!
    'SetBkMode dstDIB.getDIBDC, FONT_OPAQUE
    'SetBkColor dstDIB.getDIBDC, vbWhite
    
    'We now need to figure out the size of our drawn text.  Wordwrap is supported if the text extends past the horizontal boundary
    ' of the target DIB.
    
    'Start by retrieving default size values
    Dim targetWidth As Long, targetHeight As Long
    Dim txtSize As POINTAPI
    GetTextExtentPoint32 m_TmpGdiDIB.GetDIBDC, StrPtr(srcString), Len(srcString), txtSize
    
    Dim tmpRect As RECT
    tmpRect.Left = 0
    
    'A left-side overhang is possible on some fonts, particularly fonts where italics have to be synthesized.
    ' There's no automatic way to handle this, so we do it manually (as usual).
    Dim firstCharABCWidth As ABCFLOAT
    If Fonts.GetABCWidthOfGlyph(m_TmpGdiDIB.GetDIBDC, AscW(Left$(srcString, 1)), firstCharABCWidth, True) Then
        
        'If the offset value is negative, we want to increase our rendering offset proportionally.  (Note that we add an
        ' extra pixel to account for possible antialiasing along the border.)
        If firstCharABCWidth.abcfA < 0 Then
        
            Dim newLeft As Long
            newLeft = tmpRect.Left + Abs(firstCharABCWidth.abcfA) + 1
            
        End If
        
    End If
        
    
    'If the width of this string (in the current font) is less than the width of the target DIB, we will use the string's
    ' precise render dimensions for our target DIB.
    If txtSize.x + tmpRect.Left < dstDIB.GetDIBWidth Then
        
        'GDI adds 1/6 em width to account for glyph overhang.  Italic mode doubles the default overhang (and we apply that correction
        ' by default, to cover fonts where italics is the only available mode).
        targetWidth = txtSize.x + CLng((CDbl(txtSize.y) / 6 + 1) * 2) + tmpRect.Left
        If targetWidth > dstDIB.GetDIBWidth Then targetWidth = dstDIB.GetDIBWidth
        
        targetHeight = txtSize.y + 1
    
    'The string extends beyond the width of the DIB. Wordwrap is required.  Calculate height now.
    Else
    
        targetWidth = dstDIB.GetDIBWidth - 1
        
        tmpRect.Top = 0
        tmpRect.Bottom = 0
        tmpRect.Right = targetWidth
        
        DrawText m_TmpGdiDIB.GetDIBDC, StrPtr(srcString), Len(srcString), tmpRect, GetDrawTextAlignmentFlags() Or DT_CALCRECT Or DT_WORDBREAK Or DT_EXTERNALLEADING Or DT_NOPREFIX
        
        'Take the height of the rendered string, or the height of the target DIB, whichever is smaller
        If tmpRect.Bottom > dstDIB.GetDIBHeight Then
            targetHeight = dstDIB.GetDIBHeight
        Else
            targetHeight = tmpRect.Bottom + 1
        End If
    
    End If
    
    'Recreate the target DIB, as necessary
    If (m_TmpGdiDIB.GetDIBWidth <> targetWidth) Or (m_TmpGdiDIB.GetDIBHeight <> targetHeight) Then
    
        'Remove the font from the temporary DC, because we are about to create it.
        SelectObject m_TmpGdiDIB.GetDIBDC, oldFont
    
        'Create the temporary DIB.  We know this will be smaller or equal to the size of the target DIB, greatly simplifying the eventual transfer process.
        m_TmpGdiDIB.CreateBlank targetWidth, targetHeight, 24
        m_TmpGdiDIB.ResetDIB
        
        'Re-select the font into the DIB
        oldFont = SelectObject(m_TmpGdiDIB.GetDIBDC, m_GDIFont)
        
        'Reset other font parameters.
        SetTextColor m_TmpGdiDIB.GetDIBDC, RGB(255, 255, 255)
        SetTextAlign m_TmpGdiDIB.GetDIBDC, 0
        SetBkMode m_TmpGdiDIB.GetDIBDC, FONT_TRANSPARENT
        
    Else
        m_TmpGdiDIB.ResetDIB
    End If
    
    'Populate our clipping rect with the final width and height values we generated
    tmpRect.Right = targetWidth
    tmpRect.Bottom = targetHeight
    
    'If overhang is present on the first glyph, it's likely to be present on other glyphs as well.  Mirror the left overhang (if any)
    ' across the right border.  (We do this because GDI word wrap is unpredictable, and we can't retrieve where it decides to
    ' wordwrap - so we have to make assumptions like this.)
    If firstCharABCWidth.abcfA < 0 Then
        Dim newRight As Long
        newRight = tmpRect.Right - (Abs(firstCharABCWidth.abcfC) + 1)
        If newRight > tmpRect.Left Then tmpRect.Right = newRight
    End If
    
    'Use DrawText for the actual rendering
    Dim retDrawText As Long
    retDrawText = DrawText(m_TmpGdiDIB.GetDIBDC, StrPtr(srcString), Len(srcString), tmpRect, DT_WORDBREAK Or DT_NOCLIP Or GetDrawTextAlignmentFlags())
    
    'We can actually determine a return value now
    RenderTextToDIB_GDI = CBool(retDrawText <> 0)
    
    'Remove the font, so we can use again it later!
    SelectObject m_TmpGdiDIB.GetDIBDC, oldFont
    
    'With the temporary DIB successfully renderered, our new job is to transfer its contents to the destination DIB.
    ' Unfortunately, we have to do this manually, using per-pixel code.
    
    'Start by clearing the destination DIB
    dstDIB.ResetDIB
    dstDIB.SetInitialAlphaPremultiplicationState True
        
    'Next, we need to prepare lookup tables based on the translation of gray pixels (since we rendered white text to a black background)
    ' to the target color, with alpha accounted for.  Note that alpha is easy - the grayscale value of the temporary DIB represents
    ' alpha - so we just need to calculate premultiplied colors.
    
    'Start by extracting individual RGB components from the text color.
    Dim r As Long, g As Long, b As Long, a As Long
    r = Colors.ExtractRed(crColor)
    g = Colors.ExtractGreen(crColor)
    b = Colors.ExtractBlue(crColor)
    
    'Build RGB lookup tables for premultiplied alpha.  Note that we must also calculate an alpha lookup table, so we can manually implement
    ' GDI+-specific features like "text clarity".
    Dim rLookup() As Byte, gLookup() As Byte, bLookup() As Byte, aLookup() As Byte
    ReDim rLookup(0 To 255) As Byte
    ReDim gLookup(0 To 255) As Byte
    ReDim bLookup(0 To 255) As Byte
    ReDim aLookup(0 To 255) As Byte
    
    Dim x As Long, y As Long
    Dim preMultiplicationFactor As Double, textContrastFactor As Double, alphaCalculation As Double
    
    '"Standard" and "Crisp" AA modes need to mimic the Text Contrast setting GDI+ provides.  We use a custom calculation
    ' that provides similar results.
    
    'Convert m_TextContrast from the range [0, 10] to [0.5, 3.0]
    textContrastFactor = (m_TextContrast / 4) + 0.5
    
    For x = 0 To 255
        
        a = x
        
        'Start by calculating alpha.  This varies according to the current antialiasing mode and clarity.
        Select Case m_TextAntialiasing
        
            Case pdta_None
                
                'Split alpha to strictly 0 or 255 values.  (Also, ignore text contrast completely.)
                If a < 127 Then
                    a = 0
                Else
                    a = 255
                End If
            
            Case pdta_Standard
                
                'Modify incoming alpha according to our text contrast measurement
                alphaCalculation = a / 255
                alphaCalculation = alphaCalculation ^ textContrastFactor
                
                'Convert back to the [0, 255] range
                a = alphaCalculation * 255
                If a < 0 Then
                    a = 0
                ElseIf a > 255 Then
                    a = 255
                End If
            
            Case pdta_Crisp
            
                'Crisp is almost identical to regular alpha, except we convert alpha to a ^ 2 curve in advance
                alphaCalculation = a / 255
                alphaCalculation = alphaCalculation * alphaCalculation
                alphaCalculation = alphaCalculation ^ textContrastFactor
                
                'Convert back to the [0, 255] range
                a = alphaCalculation * 255
                If a < 0 Then
                    a = 0
                ElseIf a > 255 Then
                    a = 255
                End If
        
        End Select
        
        preMultiplicationFactor = a / 255
        
        rLookup(x) = r * preMultiplicationFactor
        gLookup(x) = g * preMultiplicationFactor
        bLookup(x) = b * preMultiplicationFactor
        aLookup(x) = a
        
    Next x
    
    'Lookup tables are now ready.  Time to transfer the bits!
        
    'Create a local array and point it at the pixel data of the target image
    Dim dstImageData() As Byte
    Dim dstSA As SAFEARRAY2D
    PrepSafeArray dstSA, dstDIB
    CopyMemory ByVal VarPtrArray(dstImageData()), VarPtr(dstSA), 4
    
    'Create a second local array and point it as the pixel data of our temporary image
    Dim srcImageData() As Byte
    Dim srcSA As SAFEARRAY2D
    PrepSafeArray srcSA, m_TmpGdiDIB
    CopyMemory ByVal VarPtrArray(srcImageData()), VarPtr(srcSA), 4
    
    'Local loop variables can be more efficiently cached by VB's compiler, so we transfer all relevant loop data here
    Dim finalX As Long, finalY As Long
    finalX = m_TmpGdiDIB.GetDIBWidth - 1
    finalY = m_TmpGdiDIB.GetDIBHeight - 1
    
    'As a failsafe, make sure finalX and Y don't extend beyond our target DIB, either
    If finalX >= dstDIB.GetDIBWidth Then finalX = dstDIB.GetDIBWidth - 1
    If finalY >= dstDIB.GetDIBHeight Then finalY = dstDIB.GetDIBHeight - 1
    
    'Finally, we also need to calculate x/y offset values.  For performance reasons, this class renders text to a temporary DIB
    ' at the exact size of the text's bounding rect.  Because we must manually convert these results to 32-bpp, then transfer them
    ' to the destination image, the alignment of the temporary DIB vs the destination DIB varies according to the underlying
    ' text alignment.
    
    'Calculate offsets now
    Dim xOffset As Long, yOffset As Long
    
    Select Case m_HorizontalAlignment
    
        Case StringAlignmentNear
            xOffset = 0
        
        Case StringAlignmentCenter
            xOffset = (dstDIB.GetDIBWidth - m_TmpGdiDIB.GetDIBWidth) \ 2
        
        Case StringAlignmentFar
            xOffset = dstDIB.GetDIBWidth - m_TmpGdiDIB.GetDIBWidth
    
    End Select
    
    Select Case m_VerticalAlignment
    
        Case StringAlignmentNear
            yOffset = 0
        
        Case StringAlignmentCenter
            yOffset = (dstDIB.GetDIBHeight - m_TmpGdiDIB.GetDIBHeight) \ 2
        
        Case StringAlignmentFar
            yOffset = dstDIB.GetDIBHeight - m_TmpGdiDIB.GetDIBHeight
    
    End Select
    
    'Add failsafe offset checks
    If xOffset < 0 Then xOffset = 0
    If yOffset < 0 Then yOffset = 0
    
    Dim xStrideSrc As Long, xStrideDst As Long
    
    'Start transferring pixels
    For x = 0 To finalX
        xStrideSrc = x * 3
        xStrideDst = (x + xOffset) * 4
    For y = 0 To finalY
        
        'Because we know the source DIB is grayscale, we can skip processing of all pixels with value = 0.
        g = srcImageData(xStrideSrc, y)
        
        If g > 0 Then
        
            'We don't need to retrieve red and green, as they are identical to blue.
            
            'Populate the destination pixel at this point with our lookup table values.
            dstImageData(xStrideDst, y + yOffset) = bLookup(g)
            dstImageData(xStrideDst + 1, y + yOffset) = gLookup(g)
            dstImageData(xStrideDst + 2, y + yOffset) = rLookup(g)
            dstImageData(xStrideDst + 3, y + yOffset) = aLookup(g)
        
        End If
        
    Next y
    Next x
    
    'Deallocate both arrays and exit
    CopyMemory ByVal VarPtrArray(srcImageData), 0&, 4
    CopyMemory ByVal VarPtrArray(dstImageData), 0&, 4
    
End Function

'Convert the current alignment flags into DrawText-compatible alignment values
Private Function GetDrawTextAlignmentFlags() As Long

    GetDrawTextAlignmentFlags = 0
    
    'Start with horizontal flags
    Select Case m_HorizontalAlignment
        
        Case StringAlignmentNear
            GetDrawTextAlignmentFlags = GetDrawTextAlignmentFlags Or DT_LEFT
        
        Case StringAlignmentCenter
            GetDrawTextAlignmentFlags = GetDrawTextAlignmentFlags Or DT_CENTER
        
        Case StringAlignmentFar
            GetDrawTextAlignmentFlags = GetDrawTextAlignmentFlags Or DT_RIGHT
        
    End Select
    
    'Add vertical flags
    Select Case m_VerticalAlignment
        
        Case StringAlignmentNear
            GetDrawTextAlignmentFlags = GetDrawTextAlignmentFlags Or DT_TOP
        
        Case StringAlignmentCenter
            GetDrawTextAlignmentFlags = GetDrawTextAlignmentFlags Or DT_VCENTER
        
        Case StringAlignmentFar
            GetDrawTextAlignmentFlags = GetDrawTextAlignmentFlags Or DT_BOTTOM
        
    End Select

End Function

'If this class is no longer required but the caller doesn't want to fully release it, they can call this sub to have us release all
' caches and system handles (hFonts, GDI+ fonts, etc).  Note that the class needs to be FULLY reinitialized after this function is called.
Public Sub ReleaseAsManyResourcesAsPossible()
    
    'Release any GDI+ objects that are easily re-created
    If (m_GDIPlusFont <> 0) Then GdipDeleteFont m_GDIPlusFont
    If (m_GDIPlusFontFamily <> 0) Then GdipDeleteFontFamily m_GDIPlusFontFamily
    
    'Release any GDI objects that are easily re-created
    If (m_GDIFont <> 0) Then DeleteObject m_GDIFont
    
    'Free our glyph collection, if one exists
    Set m_GlyphCollection = Nothing
    
End Sub

Private Sub Class_Initialize()
    
    'TEMPORARY!  Force the default, basic rendering engine while we get the basic text tool sorted out.
    m_RenderingEngine = tre_WAPI
        
    'Because GDI+ StringFormat creation is cheap, create a default typographic GDI+ StringFormat object now.
    ' This saves us having to recreate the object later.
    '
    'Note that a default typographic StringFormat has the following options set:
    ' FormatFlags: 24580 (NoClip, FitBlackBox and LineLimit.)
    ' Alignment: Near
    ' LineAlignment: Near
    ' Trimming: None
    ' HotkeyPrefix: None
    ' DigitSubstitutionMethod: User
    ' DigitSubstitutionLanguage: 0
    'GdipStringFormatGetGenericTypographic m_GDIPlusStringFormat
    
    'NOTE!  In Feb 2017, testing revealed that GDI+ reuses the object returned by GdipStringFormatGetGenericTypographic.
    ' This is some kind of shared object, and when you have multiple text layers in a single image, the layers all share
    ' (and thus repeatedly overwrite!) this same object.
    '
    'The only way around this is to manually create a typographic string format from scratch.  Uuuuugh GDI+ headaches!
    Const LANG_NEUTRAL As Long = 0&
    GdipCreateStringFormat StringFormatFlagsNoClip Or StringFormatFlagsFitBlackBox Or StringFormatFlagsLineLimit, LANG_NEUTRAL, m_GDIPlusStringFormat
    
    'Make a few adjustments to allow partially visible lines to still appear
    Dim tmpCopyFlags As Long
    GdipGetStringFormatFlags m_GDIPlusStringFormat, tmpCopyFlags
    tmpCopyFlags = tmpCopyFlags And (Not StringFormatFlagsLineLimit)
    GdipSetStringFormatFlags m_GDIPlusStringFormat, tmpCopyFlags
    
    'Mark the current font cache as dirty
    m_FontCacheClean = False
    
    'Reset the font backend marker (this is set on a per-font basis, based on a fairly sophisticated combination of settings)
    m_CurrentBackend = fb_UNKNOWN
        
    'By default, this class uses pixels
    m_FontSizeUnit = pdfu_Pixel
    
    'Set some default font properties
    m_FontFace = g_InterfaceFont
    m_FontSize = 16#                '16px = 12 points at 96 DPI
    
    m_HorizontalAlignment = StringAlignmentNear
    m_VerticalAlignment = StringAlignmentNear
    
    m_WordWrap = tre_AutoBreakWords
    m_FillActive = True
    
    
    'Set default graphics container text settings
    m_TextContrast = 4
    m_TextAntialiasing = pdta_Standard

End Sub

Private Sub Class_Terminate()
    
    'Release any GDI+ objects we created
    If (m_GDIPlusFont <> 0) Then GdipDeleteFont m_GDIPlusFont
    If (m_GDIPlusFontFamily <> 0) Then GdipDeleteFontFamily m_GDIPlusFontFamily
    If (m_GDIPlusStringFormat <> 0) Then GdipDeleteStringFormat m_GDIPlusStringFormat
    
    'Release any GDI objects we created
    If (m_GDIFont <> 0) Then DeleteObject m_GDIFont
    
End Sub
