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
'PhotoDemon Text Layer Renderer (using some hellish combination of GDI, GDI+, Uniscribe, and custom code)
'Copyright 2015-2026 by Tanner Helland
'Created: 22/April/15
'Last updated: 07/April/25
'Last update: add support for creating GDI+ fonts from user-specified (non-system) font files
'
'This class handles the messy business of font management for PD's basic and advanced text layers.
' It intermixes GDI, GDI+, and Uniscribe as necessary to provide a comprehensive text rendering solution.
'
'I originally planned to split this class in two, as a GDI-specific backend and a separate GDI+-specific backend
' (with Uniscribe stuff handled separately).  But there are complications with this.  GDI and GDI+ have different
' strengths and weaknesses (GDI has no 32-bpp rendering support, while GDI+ doesn't support OpenType), so to really
' provide a decent font rendering experience on Windows, it's unavoiadable to drag multiple libraries into the mix.
' (This is a recurring theme with literally all things on Windows - there's never one "best choice" library
' for anything.)
'
'Here's the basic setup for this class:
'
'PhotoDemon's basic text tool attempts to use built-in Windows renderers, with the following caveats...
'
' - On TrueType fonts, GDI+ is used.  This saves us a lot of grief, and lets us make use of some nice,
' typographic-specific features (like variable antialiasing contrast).
'
' - With OpenType fonts, GDI is used out of necessity.  Individual glyphs are retrieved, parsed,
' and converted into a usable format (e.g. OpenType's quadratic B-splines are converted into GDI+'s native
' cubic Bezier curve format).  The final, resultant shape(s) are then passed to GDI+ as a generic path,
' which we can then render with native GDI+ features (antialiasing, fancy gradients, etc).
'
'This unpleasant hybrid approach is part of the shit you deal with as a Windows developer.
'
'For the PHOTODEMON text rendering engine...
'
' - A comprehensive Uniscribe interface is used for text shaping, layout, and placement.  This allows me to use
'   pretty much every possible OpenType font feature, which is amazing!
'
' - Glyph paths are retrieved via GDI, and manually parsed into GDI+ path objects.  The details are handled by the
'   pdGlyphCollection class.
'
' - Rendering the final path(s) is done by GDI+.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'When debugging, performance reports can be helpful
Private Const REPORT_TEXT_RENDER_TIMING As Boolean = False

'This class switches between GDI and GDI+ as necessary.  It sets this flag automatically based on current class settings
' and success/failure of various WAPI calls.
Private Enum PD_FontBackend
    fb_UNKNOWN = 0
    fb_GDIPLUS = 1
    fb_GDI = 2
End Enum

#If False Then
    Private Const fb_UNKNOWN = 0, fb_GDIPLUS = 1, fb_GDI = 2
#End If

'If a text setting changes (and renders the current font cache invalid), this backend value gets reset to UNKNOWN.
' On the next font object creation attempt, the required backend is automatically determined based on both class settings
' and whether we could initialize the engine we want.  (Note that text will *always* be rendered successfully,
' so you don't need to concern yourself with how PD renders the text - it'll figure out a solution regardless of
' GDI/GDI+ weirdness.)
Private m_CurrentBackend As PD_FontBackend

'This class provides two different text rendering engines: WAPI and MANUAL.  WAPI is used for basic text layers, and it is
' thus limited by the innate feature support of GDI and GDI+ (BOOOORING).  The "PhotoDemon" backend is a homebrew text
' renderer that can do pretty much anything we want.  It's extremely powerful, but note that it still dips into WAPI for
' various features (e.g. Uniscribe is used for glyph layout, and then we apply our own features atop those results).
Public Enum PD_TextEngine
    te_WAPI = 0
    te_PhotoDemon = 1
End Enum

#If False Then
    Private Const te_WAPI = 0, te_PhotoDemon = 1
#End If

Private m_RenderingEngine As PD_TextEngine

'GDI+ font functions are prone to failure, for reasons both known (e.g. no OpenType support) and unknown
' (e.g. GDI+ font handling is a pile of garbage).  Here are a few common failure states.
Private Const GDIP_FONT_FAMILY_NOT_FOUND As Long = 14
Private Const GDIP_FONT_STYLE_NOT_FOUND As Long = 15
Private Const GDIP_FONT_NOT_TRUETYPE As Long = 16

'GDI+ supports multiple measurement modes, and these may actually be relevant for fonts.  Because GDI+ provides
' units that we don't want to expose to users (e.g. "World", "Display"), PD uses its own font size enumeration,
' which we translate to corresponding GDI+ units as necessary.
Public Enum PD_FontUnit
    fu_Pixel = 0
    fu_Point = 1
End Enum

#If False Then
    Private Const fu_Pixel = 0, fu_Point = 1
#End If

'GDI+ supports multiple text antialiasing modes, but half of them don't work on 32-bpp rendering targets (UGGGGGH).
' Rather than expose a confusing list of operators that may or may not be relevant, PD uses its own AA enumeration,
' which we translate to corresponding GDI+ units as necessary.
Public Enum PD_TextAntialiasing
    pdta_None = 0
    pdta_Standard = 1
    pdta_Crisp = 2
    pdta_Smooth = 3
End Enum

#If False Then
    Private Const pdta_None = 0, pdta_Standard = 1, pdta_Crisp = 2, pdta_Smooth = 3
#End If

'Text hinting and AA are merged into a single setting in GDI+
Public Enum GP_TextRenderingHint
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
Private Enum GP_StringFormatFlags
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
Private Enum GP_FontStyle
    FontStyleRegular = 0
    FontStyleBold = 1
    FontStyleItalic = 2
    FontStyleBoldItalic = 3
    FontStyleUnderline = 4
    FontStyleStrikeout = 8
End Enum

#If False Then
    Private Const FontStyleRegular = 0, FontStyleBold = 1, FontStyleItalic = 2, FontStyleBoldItalic = 3, FontStyleUnderline = 4, FontStyleStrikeout = 8
#End If

'GDI+ font family functions
Private Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal ptrToSrcFontName As Long, ByVal srcFontCollection As Long, ByRef dstFontFamily As Long) As GP_Result
Private Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal srcFontFamily As Long) As GP_Result
Private Declare Function GdipIsStyleAvailable Lib "gdiplus" (ByVal srcFontFamily As Long, ByVal srcStyleToTest As GP_FontStyle, ByRef dstIsStyleAvailable As Long) As GP_Result

'GDI+ font functions
Private Declare Function GdipCreateFont Lib "gdiplus" (ByVal srcFontFamily As Long, ByVal srcFontSize As Single, ByVal srcFontStyle As GP_FontStyle, ByVal srcMeasurementUnit As GP_Unit, ByRef dstCreatedFont As Long) As GP_Result
Private Declare Function GdipDeleteFont Lib "gdiplus" (ByVal srcFont As Long) As GP_Result
Private Declare Function GdipGetFontCollectionFamilyCount Lib "gdiplus" (ByVal hFontCollection As Long, ByRef dstNumFound As Long) As GP_Result
Private Declare Function GdipGetFontCollectionFamilyList Lib "gdiplus" (ByVal hFontCollection As Long, ByVal numSought As Long, ByVal ptrDstArrayFamilies As Long, ByRef dstNumActuallyFound As Long) As GP_Result
Private Declare Function GdipGetFamilyName Lib "gdiplus" (ByVal pFontFamily As Long, ByVal pDstName As Long, ByVal srcLangID As Integer) As GP_Result

'GDI+ string format functions
Private Declare Function GdipCreateStringFormat Lib "gdiplus" (ByVal formatAttributes As GP_StringFormatFlags, ByVal srcLanguage As Long, ByRef dstStringFormat As Long) As GP_Result
Private Declare Function GdipDeleteStringFormat Lib "gdiplus" (ByVal srcStringFormat As Long) As GP_Result
Private Declare Function GdipSetStringFormatFlags Lib "gdiplus" (ByVal dstStringFormat As Long, ByVal newFlags As Long) As GP_Result
Private Declare Function GdipGetStringFormatFlags Lib "gdiplus" (ByVal srcStringFormat As Long, ByRef dstFlags As Long) As GP_Result
Private Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal dstStringFormat As Long, ByVal newAlignment As GP_StringAlignment) As GP_Result
Private Declare Function GdipSetStringFormatLineAlign Lib "gdiplus" (ByVal dstStringFormat As Long, ByVal newLineAlignment As GP_StringAlignment) As GP_Result

'GDI+ graphics container font functions
Private Declare Function GdipSetTextRenderingHint Lib "gdiplus" (ByVal dstGraphics As Long, ByVal newRenderHintMode As GP_TextRenderingHint) As GP_Result
Private Declare Function GdipSetTextContrast Lib "gdiplus" (ByVal dstGraphics As Long, ByVal textContrast As Long) As GP_Result
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal dstGraphics As Long, ByVal newSmoothingMode As GP_SmoothingMode) As GP_Result

'GDI+ Render functions
Private Declare Function GdipDrawString Lib "gdiplus" (ByVal dstGraphics As Long, ByVal srcStringPtr As Long, ByVal strLength As Long, ByVal gdipFontHandle As Long, ByRef layoutRect As RectF, ByVal gdipStringFormat As Long, ByVal gdipBrush As Long) As GP_Result
Private Declare Function GdipDrawPath Lib "gdiplus" (ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal srcPath As Long) As GP_Result
Private Declare Function GdipDrawRectangle Lib "gdiplus" (ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal srcX As Single, ByVal srcW As Single, ByVal srcWidth As Single, ByVal srcHeight As Single) As GP_Result
Private Declare Function GdipFillPath Lib "gdiplus" (ByVal dstGraphics As Long, ByVal srcBrush As Long, ByVal srcPath As Long) As GP_Result
Private Declare Function GdipFillRectangleI Lib "gdiplus" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As GP_Result

'GDI generic object management
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'GDI DC settings
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextAlign Lib "gdi32" (ByVal hDC As Long, ByVal wFlags As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long

'GDI font rendering functions
Private Declare Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32W" (ByVal hDC As Long, ByVal lpStrPointer As Long, ByVal cbString As Long, ByRef lpSize As PointAPI) As Long

'Back color modes (PD only uses transparent fonts)
Private Const FONT_TRANSPARENT As Long = &H1

'Formatting constants for DrawText
Private Enum Win32_DrawTextConstants
    DT_TOP = &H0
    DT_LEFT = &H0
    DT_CENTER = &H1
    DT_RIGHT = &H2
    DT_VCENTER = &H4
    DT_BOTTOM = &H8
    DT_WORDBREAK = &H10
    DT_SINGLELINE = &H20
    DT_EXPANDTABS = &H40
    DT_TABSTOP = &H80
    DT_NOCLIP = &H100
    DT_EXTERNALLEADING = &H200
    DT_CALCRECT = &H400
    DT_NOPREFIX = &H800
    DT_INTERNAL = &H1000
    DT_EDITCONTROL = &H2000
    DT_PATH_ELLIPSIS = &H4000
    DT_END_ELLIPSIS = &H8000&
    DT_MODIFYSTRING = &H10000
    DT_RTLREADING = &H20000
    DT_WORD_ELLIPSIS = &H40000
End Enum

#If False Then
    Private Const DT_TOP = &H0, DT_LEFT = &H0, DT_CENTER = &H1, DT_RIGHT = &H2, DT_VCENTER = &H4, DT_BOTTOM = &H8, DT_WORDBREAK = &H10, DT_SINGLELINE = &H20, DT_EXPANDTABS = &H40, DT_TABSTOP = &H80, DT_NOCLIP = &H100, DT_EXTERNALLEADING = &H200, DT_CALCRECT = &H400, DT_NOPREFIX = &H800, DT_INTERNAL = &H1000, DT_EDITCONTROL = &H2000, DT_PATH_ELLIPSIS = &H4000, DT_END_ELLIPSIS = &H8000&, DT_MODIFYSTRING = &H10000, DT_RTLREADING = &H20000, DT_WORD_ELLIPSIS = &H40000
#End If

'The current text-to-be-rendered
Private m_CurrentText As String

'This class internally maintains various font properties.  These properties must be converted into specific GDI+ parameters via
' various means, but they are stored in human-friendly format to simplify serializing a class instance to an XML string.
Private m_FontFace As String
Private m_FontColor As Long
Private m_FontSize As Single
Private m_FontSizeUnit As PD_FontUnit
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
Private m_HorizontalAlignment As GP_StringAlignment
Private m_VerticalAlignment As GP_StringAlignment
Private m_AlignLastLine As GP_StringAlignment

'Even *more* string settings are not stored in the font itself, or in a StringFormat object, but in the target
' GDI+ Graphics container.  These must be assigned to the graphics container prior to painting text, so there's not really
' an easy way to cache these settings.  (We could keep a temporary graphics container on hand, but we still have to clone
' it prior to rendering, so I doubt there's any gain there.)  As such, these values are not relevant until text is
' actually painted onto a target object.
Private m_TextAntialiasing As PD_TextAntialiasing
Private m_TextContrast As Long

'The per-glyph PHOTODEMON renderer supports a bunch of unique options, including tons of appearance-related ones
Private m_TextHinting As Boolean
Private m_WordWrap As PD_TextWordwrap
Private m_StretchToFit As PD_TextStretchToFit
Private m_FillActive As Boolean
Private m_FillBrush As String
Private m_OutlineActive As Boolean
Private m_OutlinePen As String
Private m_BackgroundActive As Boolean
Private m_BackgroundBrush As String
Private m_BackBorderActive As Boolean
Private m_BackBorderPen As String
Private m_OutlineAboveFill As Boolean

'...as well as customized spacing and layout options
Private m_LineSpacing As Single
Private m_MarginLeft As Single
Private m_MarginTop As Single
Private m_MarginRight As Single
Private m_MarginBottom As Single

'...and a bunch of character-level modifications
Private m_CharRemap As PD_StringRemap
Private m_CharSpacing As Double
Private m_CharOrientation As Double
Private m_CharJitterX As Double
Private m_CharJitterY As Double
Private m_CharInflation As Double
Private m_CharMirror As PD_CharacterMirror

'For performance reasons, this class caches various font objects and handles from both GDI and GDI+.  This spares us from
' having to recreate expensive font data during rendering steps.
Private m_GDIPlusFont As Long
Private m_GDIPlusFontFamily As Long, m_GDIPlusFontFamilyIsPrivate As Boolean
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
Private m_tmpDIB As pdDIB

'Matching font names to font families is cumbersome, and we don't want to do it more than we have to.
' After a successful match, the result is cached here.
Private m_UserFontMatches As pdVariantHash

'Get/set the current rendering string.  PD uses a lot of tricks to improve rendering performance, including ignoring redraw requests if none of the
' class's internals have changed, so the current text must be assigned via these functions prior to actual rendering.
Friend Function GetCurrentText() As String
    GetCurrentText = m_CurrentText
End Function

Friend Sub SetCurrentText(ByRef newText As String)
    m_CurrentText = newText
End Sub

'Get functions for various font styles, and whether said styles are supported by the current font.
' (If no font name has been set, the function returns *will not be valid*)
Friend Function GetFontColor() As Long
    GetFontColor = m_FontColor
End Function

Friend Sub SetFontColor(ByVal newFontColor As Long)
    m_FontColor = newFontColor
End Sub

Friend Function GetFontBold() As Boolean
    GetFontBold = m_FontBold
End Function

Friend Function IsFontBoldSupported() As Boolean
    IsFontBoldSupported = m_BoldSupported
End Function

Friend Function GetFontItalic() As Boolean
    GetFontItalic = m_FontItalic
End Function

Friend Function IsFontItalicSupported() As Boolean
    IsFontItalicSupported = m_ItalicSupported
End Function

Friend Function GetFontUnderline() As Boolean
    GetFontUnderline = m_FontUnderline
End Function

Friend Function IsFontUnderlineSupported() As Boolean
    IsFontUnderlineSupported = m_UnderlineSupported
End Function

Friend Function GetFontStrikeout() As Boolean
    GetFontStrikeout = m_FontStrikeout
End Function

Friend Function IsFontStrikeoutSupported() As Boolean
    IsFontStrikeoutSupported = m_StrikeoutSupported
End Function

'Set functions for various font styles.
' Note that these functions reset the current font cache, so all of them contain checks for "If (newSetting != oldSetting)".
' This lets us skip cache rebuilds unless absolutely necessary.
Friend Sub SetFontBold(ByVal newValue As Boolean)
    If (newValue <> m_FontBold) Then
        m_FontBold = newValue
        m_FontCacheClean = False
    End If
End Sub

Friend Sub SetFontItalic(ByVal newValue As Boolean)
    If (newValue <> m_FontItalic) Then
        m_FontItalic = newValue
        m_FontCacheClean = False
    End If
End Sub

Friend Sub SetFontUnderline(ByVal newValue As Boolean)
    If newValue <> m_FontUnderline Then
        m_FontUnderline = newValue
        m_FontCacheClean = False
    End If
End Sub

Friend Sub SetFontStrikeout(ByVal newValue As Boolean)
    If (newValue <> m_FontStrikeout) Then
        m_FontStrikeout = newValue
        m_FontCacheClean = False
    End If
End Sub

'Get/set functions for font size and unit.  By default, this class uses pixels.  Other units may be supported
' in the future.
Friend Function GetFontSize() As Single
    GetFontSize = m_FontSize
End Function

Friend Sub SetFontSize(ByVal newValue As Single)
    If (newValue <> m_FontSize) Then
        m_FontSize = newValue
        m_FontCacheClean = False
    End If
End Sub

Friend Function GetFontSizeUnit() As PD_FontUnit
    GetFontSizeUnit = m_FontSizeUnit
End Function

Friend Sub SetFontSizeUnit(ByVal newUnit As PD_FontUnit)
    If (newUnit <> m_FontSizeUnit) Then
        m_FontSizeUnit = newUnit
        m_FontCacheClean = False
    End If
End Sub

Private Function GetFontAntialiasingStringFromEnum(ByVal srcValue As PD_TextAntialiasing) As String
    If (srcValue = pdta_None) Then
        GetFontAntialiasingStringFromEnum = "none"
    ElseIf (srcValue = pdta_Standard) Then
        GetFontAntialiasingStringFromEnum = "standard"
    ElseIf (srcValue = pdta_Crisp) Then
        GetFontAntialiasingStringFromEnum = "crisp"
    ElseIf (srcValue = pdta_Crisp) Then
        GetFontAntialiasingStringFromEnum = "smooth"
    Else
        InternalError "GetFontAntialiasingStringFromEnum", "bad value: " & srcValue
        GetFontAntialiasingStringFromEnum = "standard"
    End If
End Function

Private Function GetFontAntialiasingEnumFromString(ByRef srcValue As String) As PD_TextAntialiasing
    If (srcValue = "none") Then
        GetFontAntialiasingEnumFromString = pdta_None
    ElseIf (srcValue = "standard") Then
        GetFontAntialiasingEnumFromString = pdta_Standard
    ElseIf (srcValue = "crisp") Then
        GetFontAntialiasingEnumFromString = pdta_Crisp
    ElseIf (srcValue = "smooth") Then
        GetFontAntialiasingEnumFromString = pdta_Smooth
    Else
        InternalError "GetFontAntialiasingEnumFromString", "bad value: " & srcValue
        GetFontAntialiasingEnumFromString = pdta_Standard
    End If
End Function

'When serializing text properties to XML, use string representations instead of hard-coded enums
Private Function GetFontSizeStringFromUnit(ByVal srcUnit As PD_FontUnit) As String
    If (srcUnit = fu_Pixel) Then
        GetFontSizeStringFromUnit = "px"
    ElseIf (srcUnit = fu_Point) Then
        GetFontSizeStringFromUnit = "pt"
    Else
        InternalError "GetFontSizeStringFromUnit", "bad unit: " & srcUnit
        GetFontSizeStringFromUnit = "px"
    End If
End Function

Private Function GetFontSizeUnitFromString(ByRef srcString As String) As PD_FontUnit
    If (srcString = "px") Then
        GetFontSizeUnitFromString = fu_Pixel
    ElseIf (srcString = "pt") Then
        GetFontSizeUnitFromString = fu_Point
    Else
        InternalError "GetFontSizeUnitFromString", "bad string: " & srcString
        GetFontSizeUnitFromString = fu_Pixel
    End If
End Function

Private Function GetRendererStringFromEnum(ByVal srcValue As PD_TextEngine) As String
    If (srcValue = te_WAPI) Then
        GetRendererStringFromEnum = "wapi"
    ElseIf (srcValue = te_PhotoDemon) Then
        GetRendererStringFromEnum = "pd"
    Else
        InternalError "GetRendererStringFromEnum", "bad value: " & srcValue
        GetRendererStringFromEnum = "wapi"
    End If
End Function

Private Function GetRendererEnumFromString(ByRef srcValue As String) As PD_TextEngine
    If (srcValue = "wapi") Then
        GetRendererEnumFromString = te_WAPI
    ElseIf (srcValue = "pd") Then
        GetRendererEnumFromString = te_PhotoDemon
    Else
        InternalError "GetRendererEnumFromString", "bad value: " & srcValue
        GetRendererEnumFromString = te_WAPI
    End If
End Function

Private Function GetStretchToFitEnumFromString(ByRef srcValue As String) As PD_TextStretchToFit
    If (srcValue = "none") Then
        GetStretchToFitEnumFromString = stf_None
    ElseIf (srcValue = "box") Then
        GetStretchToFitEnumFromString = stf_Box
    ElseIf (srcValue = "slab") Then
        GetStretchToFitEnumFromString = stf_Slab
    End If
End Function

Private Function GetStretchToFitStringFromEnum(ByVal srcValue As PD_TextStretchToFit) As String
    If (srcValue = stf_None) Then
        GetStretchToFitStringFromEnum = "none"
    ElseIf (srcValue = stf_Box) Then
        GetStretchToFitStringFromEnum = "box"
    ElseIf (srcValue = stf_Slab) Then
        GetStretchToFitStringFromEnum = "slab"
    End If
End Function

Private Function GetTextMirrorStringFromEnum(ByVal srcValue As PD_CharacterMirror) As String
    If (srcValue = cm_None) Then
        GetTextMirrorStringFromEnum = "none"
    ElseIf (srcValue = cm_Horizontal) Then
        GetTextMirrorStringFromEnum = "horz"
    ElseIf (srcValue = cm_Vertical) Then
        GetTextMirrorStringFromEnum = "vert"
    ElseIf (srcValue = cm_Both) Then
        GetTextMirrorStringFromEnum = "both"
    Else
        InternalError "GetTextMirrorStringFromEnum", "bad value: " & srcValue
        GetTextMirrorStringFromEnum = "none"
    End If
End Function

Private Function GetTextMirrorEnumFromString(ByRef srcValue As String) As PD_CharacterMirror
    If (srcValue = "none") Then
        GetTextMirrorEnumFromString = cm_None
    ElseIf (srcValue = "horz") Then
        GetTextMirrorEnumFromString = cm_Horizontal
    ElseIf (srcValue = "vert") Then
        GetTextMirrorEnumFromString = cm_Vertical
    ElseIf (srcValue = "both") Then
        GetTextMirrorEnumFromString = cm_Both
    Else
        InternalError "GetTextMirrorEnumFromString", "bad value: " & srcValue
        GetTextMirrorEnumFromString = cm_None
    End If
End Function

Private Function GetTextRemapStringFromEnum(ByVal srcValue As PD_StringRemap) As String
    If (srcValue = sr_None) Then
        GetTextRemapStringFromEnum = "none"
    ElseIf (srcValue = sr_LowerCase) Then
        GetTextRemapStringFromEnum = "lower"
    ElseIf (srcValue = sr_UpperCase) Then
        GetTextRemapStringFromEnum = "upper"
    ElseIf (srcValue = sr_Titlecase) Then
        GetTextRemapStringFromEnum = "title"
    ElseIf (srcValue = sr_Hiragana) Then
        GetTextRemapStringFromEnum = "hiragana"
    ElseIf (srcValue = sr_Katakana) Then
        GetTextRemapStringFromEnum = "katakana"
    ElseIf (srcValue = sr_ChineseSimple) Then
        GetTextRemapStringFromEnum = "chinese-simple"
    ElseIf (srcValue = sr_ChineseTraditional) Then
        GetTextRemapStringFromEnum = "chinese-traditional"
    Else
        InternalError "GetTextRemapStringFromEnum", "bad value: " & srcValue
        GetTextRemapStringFromEnum = "none"
    End If
End Function

Private Function GetTextRemapEnumFromString(ByRef srcValue As String) As PD_StringRemap
    If (srcValue = "none") Then
        GetTextRemapEnumFromString = sr_None
    ElseIf (srcValue = "lower") Then
        GetTextRemapEnumFromString = sr_LowerCase
    ElseIf (srcValue = "upper") Then
        GetTextRemapEnumFromString = sr_UpperCase
    ElseIf (srcValue = "title") Then
        GetTextRemapEnumFromString = sr_Titlecase
    ElseIf (srcValue = "hiragana") Then
        GetTextRemapEnumFromString = sr_Hiragana
    ElseIf (srcValue = "katakana") Then
        GetTextRemapEnumFromString = sr_Katakana
    ElseIf (srcValue = "chinese-simple") Then
        GetTextRemapEnumFromString = sr_ChineseSimple
    ElseIf (srcValue = "chinese-traditional") Then
        GetTextRemapEnumFromString = sr_ChineseTraditional
    Else
        InternalError "GetTextRemapEnumFromString", "bad value: " & srcValue
        GetTextRemapEnumFromString = sr_None
    End If
End Function

Private Function GetWordwrapStringFromEnum(ByVal srcValue As PD_TextWordwrap) As String
    If (srcValue = tww_None) Then
        GetWordwrapStringFromEnum = "none"
    ElseIf (srcValue = tww_Manual) Then
        GetWordwrapStringFromEnum = "manual"
    ElseIf (srcValue = tww_AutoCharacter) Then
        GetWordwrapStringFromEnum = "char"
    ElseIf (srcValue = tww_AutoWord) Then
        GetWordwrapStringFromEnum = "word"
    Else
        InternalError "GetWordwrapStringFromEnum", "bad value: " & srcValue
        GetWordwrapStringFromEnum = "none"
    End If
End Function

Private Function GetWordwrapEnumFromString(ByRef srcValue As String) As PD_TextWordwrap
    If (srcValue = "none") Then
        GetWordwrapEnumFromString = tww_None
    ElseIf (srcValue = "manual") Then
        GetWordwrapEnumFromString = tww_Manual
    ElseIf (srcValue = "char") Then
        GetWordwrapEnumFromString = tww_AutoCharacter
    ElseIf (srcValue = "word") Then
        GetWordwrapEnumFromString = tww_AutoWord
    Else
        InternalError "GetWordwrapEnumFromString", "bad value: " & srcValue
        GetWordwrapEnumFromString = tww_None
    End If
End Function

'XML get/set functions for getting/setting all parameters at once.
Friend Function GetAllFontSettingsAsXML() As String

    'We don't need a full-featured XML creator (e.g. with encoding decs and the like), so we can use the smaller,
    ' lighter pdSerialize class
    Dim cSerialize As pdSerialize
    Set cSerialize = New pdSerialize
    
    'Add a basic header and version info
    cSerialize.AddXMLString "<photodemon-text-object>"
    cSerialize.AddParam "text-version", 1&, True, True
    
    'Next comes the huge list of possible text properties
    With cSerialize
        .AddParam "obj-text", m_CurrentText, True, False
        .AddParam "obj-font-color", m_FontColor, True, True
        .AddParam "obj-font-face", m_FontFace, True, False
        .AddParam "obj-font-size", m_FontSize, True, True
        .AddParam "obj-font-size-unit", GetFontSizeStringFromUnit(m_FontSizeUnit), True, True
        .AddParam "obj-font-bold", m_FontBold, True, True
        .AddParam "obj-font-italic", m_FontItalic, True, True
        .AddParam "obj-font-underling", m_FontUnderline, True, True
        .AddParam "obj-font-strikeout", m_FontStrikeout, True, True
        .AddParam "obj-align-horizontal", GetAlignmentStringFromUnit(m_HorizontalAlignment), True, True
        .AddParam "obj-align-vertical", GetAlignmentStringFromUnit(m_VerticalAlignment), True, True
        .AddParam "obj-antialiasing-type", GetFontAntialiasingStringFromEnum(m_TextAntialiasing), True, True
        .AddParam "obj-antialiasing-contrast", m_TextContrast, True, True
        .AddParam "obj-text-renderer", GetRendererStringFromEnum(m_RenderingEngine), True, True
        .AddParam "obj-hinting", m_TextHinting, True, True
        .AddParam "obj-wordwrap-mode", GetWordwrapStringFromEnum(m_WordWrap), True, True
        .AddParam "obj-stretch-to-fit", GetStretchToFitStringFromEnum(m_StretchToFit), True, True
        .AddParam "obj-fill-active", m_FillActive, True, True
        .AddParam "obj-fill-brush", m_FillBrush, True, False
        .AddParam "obj-stroke-active", m_OutlineActive, True, True
        .AddParam "obj-stroke-pen", m_OutlinePen, True, False
        .AddParam "obj-background-fill-active", m_BackgroundActive, True, True
        .AddParam "obj-background-fill-brush", m_BackgroundBrush, True, False
        .AddParam "obj-background-stroke-active", m_BackBorderActive, True, True
        .AddParam "obj-background-stroke-pen", m_BackBorderPen, True, False
        .AddParam "obj-line-spacing", m_LineSpacing, True, True
        .AddParam "obj-margin-left", m_MarginLeft, True, True
        .AddParam "obj-margin-top", m_MarginTop, True, True
        .AddParam "obj-margin-right", m_MarginRight, True, True
        .AddParam "obj-margin-bottom", m_MarginBottom, True, True
        .AddParam "obj-character-remap", GetTextRemapStringFromEnum(m_CharRemap), True, True
        .AddParam "obj-character-spacing", m_CharSpacing, True, True
        .AddParam "obj-character-rotation", m_CharOrientation, True, True
        .AddParam "obj-character-jitter-x", m_CharJitterX, True, True
        .AddParam "obj-character-jitter-y", m_CharJitterY, True, True
        .AddParam "obj-character-inflation", m_CharInflation, True, True
        .AddParam "obj-character-mirror", GetTextMirrorStringFromEnum(m_CharMirror), True, True
        .AddParam "obj-align-last-line", GetAlignmentStringFromUnit(m_AlignLastLine), True, True
        .AddParam "obj-outline-above-fill", m_OutlineAboveFill, True, True
    End With
    
    'Close out the text object
    cSerialize.AddXMLString "</photodemon-text-object>"
    
    'Return the completed XML string
    GetAllFontSettingsAsXML = cSerialize.GetParamString()

End Function

'Set all text and font parameters at once, using an XML string created by a previous call
' to GetAllFontSettingsAsXML(). (Note that this function manually marks the font cache
' as dirty, so you should only use it when loading a serialized object for the first time.)
Friend Function SetAllFontSettingsFromXML(ByRef srcXMLString As String) As Boolean
    
    'We don't need a full-featured XML creator (e.g. with encoding decs and the like), so we can use the smaller,
    ' lighter pdSerialize class
    Dim cSerialize As pdSerialize
    Set cSerialize = New pdSerialize
    cSerialize.SetParamString srcXMLString
    
    'If basic validation fails, revert to the legacy font serialization engine
    Dim useLegacyLoader As Boolean
    useLegacyLoader = (Not cSerialize.DoesParamExist("photodemon-text-object"))
    If (Not useLegacyLoader) Then useLegacyLoader = (cSerialize.GetLong("text-version", 0, True) < 1)
    
    If useLegacyLoader Then
        Set cSerialize = Nothing
        SetAllFontSettingsFromXML = SetAllFontSettingsFromXML_Legacy(srcXMLString)
    Else
        
        With cSerialize
            SetGenericTextProperty ptp_Text, .GetString("obj-text", vbNullString)
            SetGenericTextProperty ptp_FontColor, .GetLong("obj-font-color", vbBlack)
            SetGenericTextProperty ptp_FontFace, .GetString("obj-font-face", Fonts.GetUIFontName())
            SetGenericTextProperty ptp_FontSize, .GetDouble("obj-font-size", 16#)
            SetGenericTextProperty ptp_FontSizeUnit, GetFontSizeUnitFromString(.GetString("obj-font-size-unit", GetFontSizeStringFromUnit(fu_Pixel), True))
            SetGenericTextProperty ptp_FontBold, .GetBool("obj-font-bold", False)
            SetGenericTextProperty ptp_FontItalic, .GetBool("obj-font-italic", False)
            SetGenericTextProperty ptp_FontUnderline, .GetBool("obj-font-underline", False)
            SetGenericTextProperty ptp_FontStrikeout, .GetBool("obj-font-strikeout", False)
            SetGenericTextProperty ptp_HorizontalAlignment, GetAlignmentUnitFromString(.GetString("obj-align-horizontal", GetAlignmentStringFromUnit(StringAlignmentNear), True))
            SetGenericTextProperty ptp_VerticalAlignment, GetAlignmentUnitFromString(.GetString("obj-align-vertical", GetAlignmentStringFromUnit(StringAlignmentNear), True))
            SetGenericTextProperty ptp_TextAntialiasing, GetFontAntialiasingEnumFromString(.GetString("obj-antialiasing-type", GetFontAntialiasingStringFromEnum(pdta_Standard), True))
            SetGenericTextProperty ptp_TextContrast, .GetLong("obj-antialiasing-contrast", 5)
            SetGenericTextProperty ptp_RenderingEngine, GetRendererEnumFromString(.GetString("obj-text-renderer", GetRendererStringFromEnum(te_WAPI), True))
            SetGenericTextProperty ptp_TextHinting, .GetBool("obj-hinting", False)
            SetGenericTextProperty ptp_WordWrap, GetWordwrapEnumFromString(.GetString("obj-wordwrap-mode", GetWordwrapStringFromEnum(tww_AutoWord), True))
            SetGenericTextProperty ptp_StretchToFit, GetStretchToFitEnumFromString(.GetString("obj-stretch-to-fit", False, True))
            SetGenericTextProperty ptp_FillActive, .GetBool("obj-fill-active", True)
            SetGenericTextProperty ptp_FillBrush, .GetString("obj-fill-brush", vbNullString)
            SetGenericTextProperty ptp_OutlineActive, .GetBool("obj-stroke-active", False)
            SetGenericTextProperty ptp_OutlinePen, .GetString("obj-stroke-pen", vbNullString)
            SetGenericTextProperty ptp_BackgroundActive, .GetBool("obj-background-fill-active", False)
            SetGenericTextProperty ptp_BackgroundBrush, .GetString("obj-background-fill-brush", vbNullString)
            SetGenericTextProperty ptp_BackBorderActive, .GetBool("obj-background-stroke-active", False)
            SetGenericTextProperty ptp_BackBorderPen, .GetString("obj-background-stroke-pen", vbNullString)
            SetGenericTextProperty ptp_LineSpacing, .GetDouble("obj-line-spacing", 0#)
            SetGenericTextProperty ptp_MarginLeft, .GetDouble("obj-margin-left", 0#)
            SetGenericTextProperty ptp_MarginTop, .GetDouble("obj-margin-top", 0#)
            SetGenericTextProperty ptp_MarginRight, .GetDouble("obj-margin-right", 0#)
            SetGenericTextProperty ptp_MarginBottom, .GetDouble("obj-margin-bottom", 0#)
            SetGenericTextProperty ptp_CharRemap, GetTextRemapEnumFromString(.GetString("obj-character-remap", GetTextRemapStringFromEnum(sr_None), True))
            SetGenericTextProperty ptp_CharSpacing, .GetDouble("obj-character-spacing", 0#)
            SetGenericTextProperty ptp_CharOrientation, .GetDouble("obj-character-rotation", 0#)
            SetGenericTextProperty ptp_CharJitterX, .GetDouble("obj-character-jitter-x", 0#)
            SetGenericTextProperty ptp_CharJitterY, .GetDouble("obj-character-jitter-y", 0#)
            SetGenericTextProperty ptp_CharInflation, .GetDouble("obj-character-inflation", 0#)
            SetGenericTextProperty ptp_CharMirror, GetTextMirrorEnumFromString(.GetString("obj-character-mirror", GetTextMirrorStringFromEnum(cm_None), True))
            SetGenericTextProperty ptp_AlignLastLine, GetAlignmentUnitFromString(.GetString("obj-align-last-line", GetAlignmentStringFromUnit(StringAlignmentNear), True))
            SetGenericTextProperty ptp_OutlineAboveFill, .GetBool("obj-outline-above-fill", True, True)
        End With
        
        'This function does not currently provide a fail state; as long as the load request comes from PD herself, failure should be impossible.
        SetAllFontSettingsFromXML = True
        
    End If
        
End Function

Private Function GetAlignmentStringFromUnit(ByVal srcValue As GP_StringAlignment) As String
    If (srcValue = StringAlignmentNear) Then
        GetAlignmentStringFromUnit = "near"
    ElseIf (srcValue = StringAlignmentCenter) Then
        GetAlignmentStringFromUnit = "center"
    ElseIf (srcValue = StringAlignmentFar) Then
        GetAlignmentStringFromUnit = "far"
    ElseIf (srcValue = StringAlignmentJustify) Then
        GetAlignmentStringFromUnit = "justify"
    Else
        InternalError "GetAlignmentStringFromUnit", "bad value: " & srcValue
        GetAlignmentStringFromUnit = "near"
    End If
End Function

Private Function GetAlignmentUnitFromString(ByRef srcValue As String) As GP_StringAlignment
    If (srcValue = "near") Then
        GetAlignmentUnitFromString = StringAlignmentNear
    ElseIf (srcValue = "center") Then
        GetAlignmentUnitFromString = StringAlignmentCenter
    ElseIf (srcValue = "far") Then
        GetAlignmentUnitFromString = StringAlignmentFar
    ElseIf (srcValue = "justify") Then
        GetAlignmentUnitFromString = StringAlignmentJustify
    Else
        InternalError "GetAlignmentUnitFromString", "bad value: " & srcValue
        GetAlignmentUnitFromString = StringAlignmentNear
    End If
End Function

'Rather than using dedicated get functions or properties, you can retrieve font and text values via this universal function.
Friend Function GetGenericTextProperty(ByVal desiredProperty As PD_TextProperty) As Variant

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
            If (m_RenderingEngine <> te_PhotoDemon) And (m_HorizontalAlignment >= StringAlignmentJustify) Then GetGenericTextProperty = StringAlignmentNear
        
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
        
        Case ptp_StretchToFit
            GetGenericTextProperty = m_StretchToFit
            
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
            
        Case ptp_AlignLastLine
            GetGenericTextProperty = m_AlignLastLine
            
        Case ptp_OutlineAboveFill
            GetGenericTextProperty = m_OutlineAboveFill
    
    End Select

End Function

'Rather than using dedicated set functions or properties, you can modify values via this universal function.
' Returns: TRUE if a value was successfully changed; FALSE if the requested value change was meaningless.
'           The caller can use this return to know whether the associated text needs to be redrawn.
Friend Function SetGenericTextProperty(ByVal desiredProperty As PD_TextProperty, ByVal newValue As Variant) As Boolean

    Select Case desiredProperty
        
        Case ptp_Text
            If Strings.StringsNotEqual(m_CurrentText, CStr(newValue), False) Then
                m_CurrentText = newValue
                SetGenericTextProperty = True
            End If
        
        Case ptp_FontColor
            If (m_FontColor <> CLng(newValue)) Then
                m_FontColor = CLng(newValue)
                SetGenericTextProperty = True
            End If
        
        'TODO: consider raising some kind of error or message if the listed font is not available?
        ' This is a possibility for PDI images traded between PCs.
        Case ptp_FontFace
            If Strings.StringsNotEqual(m_FontFace, CStr(newValue), True) Then
                m_FontFace = newValue
                m_FontCacheClean = False
                m_CurrentBackend = fb_UNKNOWN
                SetGenericTextProperty = True
            End If
        
        Case ptp_FontSize
            If (m_FontSize <> CSng(newValue)) Then
                m_FontSize = CSng(newValue)
                m_FontCacheClean = False
                SetGenericTextProperty = True
            End If
        
        Case ptp_FontSizeUnit
            If (m_FontSizeUnit <> CLng(newValue)) Then
                m_FontSizeUnit = CLng(newValue)
                m_FontCacheClean = False
                SetGenericTextProperty = True
            End If
        
        Case ptp_FontBold
            If (m_FontBold <> CBool(newValue)) Then
                m_FontBold = CBool(newValue)
                m_FontCacheClean = False
                SetGenericTextProperty = True
            End If
        
        Case ptp_FontItalic
            If (m_FontItalic <> CBool(newValue)) Then
                m_FontItalic = CBool(newValue)
                m_FontCacheClean = False
                SetGenericTextProperty = True
            End If
        
        Case ptp_FontUnderline
            If (m_FontUnderline <> CBool(newValue)) Then
                m_FontUnderline = CBool(newValue)
                m_FontCacheClean = False
                SetGenericTextProperty = True
            End If
        
        Case ptp_FontStrikeout
            If (m_FontStrikeout <> CBool(newValue)) Then
                m_FontStrikeout = CBool(newValue)
                m_FontCacheClean = False
                SetGenericTextProperty = True
            End If
        
        'When using GDI+ as our text layout engine, we set alignment changes immediately, as it's inexpensive to do so
        Case ptp_HorizontalAlignment
            If (m_HorizontalAlignment <> CLng(newValue)) Then
                m_HorizontalAlignment = CLng(newValue)
                SetGenericTextProperty = True
                If (m_GDIPlusStringFormat <> 0) And (m_HorizontalAlignment < StringAlignmentJustify) Then GdipSetStringFormatAlign m_GDIPlusStringFormat, m_HorizontalAlignment
            End If
        
        Case ptp_VerticalAlignment
            If (m_VerticalAlignment <> CLng(newValue)) Then
                m_VerticalAlignment = CLng(newValue)
                SetGenericTextProperty = True
                If (m_GDIPlusStringFormat <> 0) Then GdipSetStringFormatLineAlign m_GDIPlusStringFormat, m_VerticalAlignment
            End If
        
        Case ptp_TextAntialiasing
            If (m_TextAntialiasing <> CLng(newValue)) Then
                m_TextAntialiasing = CLng(newValue)
                SetGenericTextProperty = True
                
                'If we're using GDI for rendering, changing antialiasing requires us to recreate the font
                If (m_CurrentBackend = fb_GDI) Then m_FontCacheClean = False
                
            End If
        
        'Text contrast is a weird one: the values vary from 0 to 12, and 4 is the default.
        ' In PD, I have modified it to use a 0-10 scale with 3 as the default, and we convert it to a
        ' 12-based measurement if working with GDI+ specifically.
        '
        'Note that this *ONLY WORKS IN BASIC RENDERING MODE.*  GDI doesn't support this measurement natively,
        ' so we have to rig it using a manual system.
        '
        '(PD's manual glyph rendering methods ignore this setting entirely.)
        Case ptp_TextContrast
            If (m_TextContrast <> CLng(newValue)) Then
                m_TextContrast = CLng(newValue)
                
                If (m_TextContrast < 0) Then
                    m_TextContrast = 0
                ElseIf (m_TextContrast > 10) Then
                    m_TextContrast = 10
                End If
                
                SetGenericTextProperty = True
            End If
            
        Case ptp_RenderingEngine
            If (m_RenderingEngine <> CLng(newValue)) Then
                m_RenderingEngine = CLng(newValue)
                SetGenericTextProperty = True
            End If
            
        Case ptp_TextHinting
            If (m_TextHinting <> CBool(newValue)) Then
                m_TextHinting = CBool(newValue)
                m_FontCacheClean = False
                SetGenericTextProperty = True
            End If
            
        Case ptp_WordWrap
            If (m_WordWrap <> CLng(newValue)) Then
                m_WordWrap = CLng(newValue)
                SetGenericTextProperty = True
            End If
        
        Case ptp_StretchToFit
            If (m_StretchToFit <> CLng(newValue)) Then
                m_StretchToFit = CLng(newValue)
                SetGenericTextProperty = True
            End If
            
        Case ptp_FillActive
            If (m_FillActive <> CBool(newValue)) Then
                m_FillActive = CBool(newValue)
                SetGenericTextProperty = True
            End If
        
        Case ptp_FillBrush
            If (m_FillBrush <> newValue) Then
                m_FillBrush = newValue
                SetGenericTextProperty = True
            End If
        
        Case ptp_OutlineActive
            If (m_OutlineActive <> CBool(newValue)) Then
                m_OutlineActive = CBool(newValue)
                SetGenericTextProperty = True
            End If
        
        Case ptp_OutlinePen
            If (m_OutlinePen <> newValue) Then
                m_OutlinePen = newValue
                SetGenericTextProperty = True
            End If
        
        Case ptp_BackgroundActive
            If (m_BackgroundActive <> CBool(newValue)) Then
                m_BackgroundActive = CBool(newValue)
                SetGenericTextProperty = True
            End If
        
        Case ptp_BackgroundBrush
            If (m_BackgroundBrush <> newValue) Then
                m_BackgroundBrush = newValue
                SetGenericTextProperty = True
            End If
        
        Case ptp_BackBorderActive
            If (m_BackBorderActive <> CBool(newValue)) Then
                m_BackBorderActive = CBool(newValue)
                SetGenericTextProperty = True
            End If
        
        Case ptp_BackBorderPen
            If (m_BackBorderPen <> newValue) Then
                m_BackBorderPen = newValue
                SetGenericTextProperty = True
            End If
        
        Case ptp_LineSpacing
            If (m_LineSpacing <> CSng(newValue)) Then
                m_LineSpacing = CSng(newValue)
                SetGenericTextProperty = True
            End If
        
        Case ptp_MarginLeft
            If (m_MarginLeft <> CSng(newValue)) Then
                m_MarginLeft = CSng(newValue)
                SetGenericTextProperty = True
            End If
        
        Case ptp_MarginTop
            If (m_MarginTop <> CSng(newValue)) Then
                m_MarginTop = CSng(newValue)
                SetGenericTextProperty = True
            End If
        
        Case ptp_MarginRight
            If (m_MarginRight <> CSng(newValue)) Then
                m_MarginRight = CSng(newValue)
                SetGenericTextProperty = True
            End If
        
        Case ptp_MarginBottom
            If (m_MarginBottom <> CSng(newValue)) Then
                m_MarginBottom = CSng(newValue)
                SetGenericTextProperty = True
            End If
            
        Case ptp_CharRemap
            If (m_CharRemap <> CLng(newValue)) Then
                m_CharRemap = CLng(newValue)
                SetGenericTextProperty = True
            End If
            
        Case ptp_CharSpacing
            If (m_CharSpacing <> CDbl(newValue)) Then
                m_CharSpacing = CDbl(newValue)
                SetGenericTextProperty = True
            End If
            
        Case ptp_CharOrientation
            If (m_CharOrientation <> CDbl(newValue)) Then
                m_CharOrientation = CDbl(newValue)
                SetGenericTextProperty = True
            End If
            
        Case ptp_CharJitterX
            If (m_CharJitterX <> CDbl(newValue)) Then
                m_CharJitterX = CDbl(newValue)
                SetGenericTextProperty = True
            End If
            
        Case ptp_CharJitterY
            If (m_CharJitterY <> CDbl(newValue)) Then
                m_CharJitterY = CDbl(newValue)
                SetGenericTextProperty = True
            End If
            
        Case ptp_CharInflation
            If (m_CharInflation <> CDbl(newValue)) Then
                m_CharInflation = CDbl(newValue)
                SetGenericTextProperty = True
            End If
            
        Case ptp_CharMirror
            If (m_CharMirror <> CLng(newValue)) Then
                m_CharMirror = CLng(newValue)
                SetGenericTextProperty = True
            End If
        
        Case ptp_AlignLastLine
            If (m_AlignLastLine <> CLng(newValue)) Then
                m_AlignLastLine = CLng(newValue)
                SetGenericTextProperty = True
            End If
            
        Case ptp_OutlineAboveFill
            If (m_OutlineAboveFill <> CBool(newValue)) Then
                m_OutlineAboveFill = CBool(newValue)
                SetGenericTextProperty = True
            End If
        
        Case Else
            PDDebug.LogAction "WARNING!  Unknown text property requested from pdTextRenderer.setGenericTextProperty()"
        
    End Select
    
End Function

'Get/Set font face.  These are more cumbersome than other font settings.
Friend Function GetFontFace() As String
    GetFontFace = m_FontFace
End Function

'Set a given font face, as specified by a font name.
Friend Sub SetFontFace(ByVal fontName As String)
    
    'If this font face has already been set, ignore this request
    If Strings.StringsNotEqual(fontName, m_FontFace, False) Then
        m_FontFace = fontName
        m_FontCacheClean = False
        m_CurrentBackend = fb_UNKNOWN
    End If
    
End Sub

'Convert a PD_FontUnit enum to a GDI+ format measurement unit.  Note that not all GDI+ units are reachable this way, by design.
Private Function ConvertPDFontUnitToGdipFontUnit(ByRef srcFontUnit As PD_FontUnit) As GP_Unit
    Select Case srcFontUnit
        Case fu_Pixel
            ConvertPDFontUnitToGdipFontUnit = GP_U_Pixel
        Case fu_Point
            ConvertPDFontUnitToGdipFontUnit = GP_U_Point
    End Select
End Function

'When all font parameters have been set, we can actually create a font!
Friend Function CreateFontObject() As Boolean
    
    'If the font cache is clean, this request is redundant; ignore it
    If m_FontCacheClean Then
        CreateFontObject = True
        Exit Function
    End If
    
    'Reset the current text engine marker, as we won't know which engine is required until we try to create the font.
    m_CurrentBackend = fb_UNKNOWN
    
    'If we have already created any font objects (in any engine), free them now
    If (m_GDIPlusFont <> 0) Then
        GdipDeleteFont m_GDIPlusFont
        m_GDIPlusFont = 0
    End If
    
    'Private font families are families created user-added font files.  They are maintained separately
    ' and freed via special APIs, so *do not* release them via GdipDeleteFontFamily.
    If (m_GDIPlusFontFamily <> 0) Then
        If (Not m_GDIPlusFontFamilyIsPrivate) Then GdipDeleteFontFamily m_GDIPlusFontFamily
        m_GDIPlusFontFamily = 0
    End If
    
    If (m_GDIFont <> 0) Then
        DeleteObject m_GDIFont
        m_GDIFont = 0
    End If
    
    'First, try to create a GDI+ font family.  This saves us a lot of grief.
    ' (Note that this step is skipped if manual rendering mode is activated; in that case, we will only use GDI fonts)
    Dim fontCreationSuccess As Boolean
    fontCreationSuccess = False
    
    If m_RenderingEngine = te_WAPI Then fontCreationSuccess = CreateGDIPlusFont()
    
    'GDI+ reported success, and we obtained a valid GDI+ font handle.
    If fontCreationSuccess And (m_GDIPlusFont <> 0) Then
        
        'Note that GDI+ works fine for this particular combination of settings, and return success.
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
            m_CurrentBackend = fb_GDI
            m_FontCacheClean = True
            CreateFontObject = True
            
        Else
        
            'For some reason, this combination of settings doesn't work at all.  I hope this never happens, but you never know.
            PDDebug.LogAction "WARNING!  pdTextRenderer's CreateFontObject failed with both GDI+ and GDI; sorry!"
            m_CurrentBackend = fb_UNKNOWN
            m_FontCacheClean = False
            CreateFontObject = False
            
        End If
        
    End If
        
End Function

'Attempt to create a GDI+ font matching the current class settings.
' Note that GDI+ creates two relevant objects: font family, and font itself.
Private Function CreateGDIPlusFont() As Boolean
    
    'If a GDI+ font family already exists, free it.
    ' (Note that private families - e.g. font families created from user-added familes - are *not* deleted here,
    '  by design; they are persistent for the session and released at shutdown time.)
    If (m_GDIPlusFontFamily <> 0) Then
        If (Not m_GDIPlusFontFamilyIsPrivate) Then GdipDeleteFontFamily m_GDIPlusFontFamily
        m_GDIPlusFontFamily = 0
    End If
    m_GDIPlusFontFamilyIsPrivate = False
    
    'Attempt to retrieve a GDI+ font family matching this name
    Dim gdipReturn As GP_Result
    gdipReturn = GdipCreateFontFamilyFromName(StrPtr(m_FontFace), 0&, m_GDIPlusFontFamily)
    
    'If this failed, the font may be a user font family
    If (gdipReturn <> GP_OK) And (Fonts.UserFonts_GetNumAdded() > 0) Then
        
        'The user added one or more fonts this session.
        
        'Get a handle to the GDI+ private collection for this session
        Dim hGdipUserFonts As Long
        hGdipUserFonts = GDI_Plus.GDIPlus_GetUserFontCollection()
        If (hGdipUserFonts <> 0) Then
            
            'See if we can simply pull from previously cached successes
            Dim idxBest As Long, idxReturn As Variant
            Dim cachedFontFamily As Long
            
            If (m_UserFontMatches Is Nothing) Then Set m_UserFontMatches = New pdVariantHash
            If m_UserFontMatches.GetItemByKey(m_FontFace, idxReturn) Then
                cachedFontFamily = idxReturn
                
            'This font face has never been created before.  Do so now!
            Else
                
                'Retrieve a list of user-added font families.
                Dim lstFontFamilies() As Long, numFontFamilies As Long
                If (GdipGetFontCollectionFamilyCount(hGdipUserFonts, numFontFamilies) = GP_OK) Then
                If (numFontFamilies > 0) Then
                    
                    ReDim lstFontFamilies(0 To numFontFamilies - 1) As Long
                    If (GdipGetFontCollectionFamilyList(hGdipUserFonts, numFontFamilies, VarPtr(lstFontFamilies(0)), numFontFamilies) <> GP_OK) Then
                        numFontFamilies = 0
                    End If
                    
                '/End numFontFamilies > 0
                End If
                '/End GdipGetFontCollectionFamilyCount
                End If
                
                'Ensure we have some user font families to query
                If (numFontFamilies > 0) Then
                    
                    'Find the best match for the requested font name
                    Const LF_FACESIZEW As Long = 64
                    
                    'Use Levenshtein distance for matching
                    Dim minDistance As Long, curDistance As Long
                    idxBest = 0
                    minDistance = 100
                    
                    Dim i As Long
                    For i = 0 To numFontFamilies - 1
                        
                        Dim srcName As String
                        srcName = String$(LF_FACESIZEW \ 2, 0)
                        
                        Const LANG_NEUTRAL As Integer = 0
                        
                        If (GdipGetFamilyName(lstFontFamilies(i), StrPtr(srcName), LANG_NEUTRAL) = GP_OK) Then
                            srcName = Strings.TrimNull(srcName)
                            curDistance = Strings.StringDistance(m_FontFace, srcName, True)
                            If (curDistance < minDistance) Then
                                minDistance = curDistance
                                idxBest = i
                            End If
                        End If
                    
                    Next i
                    
                    'Store this result in the hash table
                    m_UserFontMatches.AddItem m_FontFace, lstFontFamilies(idxBest)
                    cachedFontFamily = lstFontFamilies(idxBest)
                    
                '/numFontFamilies > 0
                End If
            
            '/End /not fontFoundAlready
            End If
            
            'Ensure the font can be created using the best-matched name
            If (cachedFontFamily <> 0) Then
                
                gdipReturn = GdipCreateFont(cachedFontFamily, m_FontSize, GetFontStylesAsGdipLong, ConvertPDFontUnitToGdipFontUnit(m_FontSizeUnit), m_GDIPlusFont)
                If (gdipReturn = GP_OK) Then
                    If (m_GDIPlusFont <> 0) Then GdipDeleteFont m_GDIPlusFont
                    m_GDIPlusFont = 0
                    m_GDIPlusFontFamilyIsPrivate = True
                    m_GDIPlusFontFamily = cachedFontFamily
                End If
                
            End If
                
        '/ (hGdipUserFonts <> 0)
        End If
    
    '/ (gdipReturn <> GP_OK) And (Fonts.UserFonts_GetNumAdded() > 0)
    End If
    
    'If GDI+ reported success and a non-zero handle was obtained, carry on!
    If (gdipReturn = GP_OK) And (m_GDIPlusFontFamily <> 0) Then
        
        'As a convenience for future efforts, cache the font styles supported by this font family
        Dim testResult As Long
        GdipIsStyleAvailable m_GDIPlusFontFamily, FontStyleBold, testResult
        m_BoldSupported = (testResult <> 0)
        
        GdipIsStyleAvailable m_GDIPlusFontFamily, FontStyleItalic, testResult
        m_ItalicSupported = (testResult <> 0)
        
        GdipIsStyleAvailable m_GDIPlusFontFamily, FontStyleUnderline, testResult
        m_UnderlineSupported = (testResult <> 0)
        
        GdipIsStyleAvailable m_GDIPlusFontFamily, FontStyleStrikeout, testResult
        m_StrikeoutSupported = (testResult <> 0)
        
        'Next, we will use the family handle to create an actual font; this includes any font styles (bold, italic, etc)
        ' and the current font size.
        gdipReturn = GdipCreateFont(m_GDIPlusFontFamily, m_FontSize, GetFontStylesAsGdipLong, ConvertPDFontUnitToGdipFontUnit(m_FontSizeUnit), m_GDIPlusFont)
        
        'Check for known errors; GDI+ cannot synthesize some font sizes, for example
        If (gdipReturn = GP_OK) And (m_GDIPlusFont <> 0) Then
            CreateGDIPlusFont = True
        Else
        
            ReportGDIPlusFailure gdipReturn
            
            'Private font families are families created user-added font files.  They are maintained separately
            ' and freed via special APIs, so *do not* release them via GdipDeleteFontFamily.
            If (m_GDIPlusFontFamily <> 0) Then
                If (Not m_GDIPlusFontFamilyIsPrivate) Then GdipDeleteFontFamily m_GDIPlusFontFamily
                m_GDIPlusFontFamily = 0
            End If
            
            CreateGDIPlusFont = False
            
        End If
        
    'GDI+ font selection can fail for a variety of reasons; report the failure reason and exit immediately
    Else
        PDDebug.LogAction "Warning: GDI+ font family could not be created; font is probably OpenType."
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
Private Function ConvertTextAAToGdipTextHint() As GP_TextRenderingHint
    
    Select Case m_TextAntialiasing
    
        Case pdta_None
            ConvertTextAAToGdipTextHint = TextRenderingHintSingleBitPerPixel
        
        Case pdta_Standard
            ConvertTextAAToGdipTextHint = TextRenderingHintAntiAlias
        
        Case pdta_Crisp
            ConvertTextAAToGdipTextHint = TextRenderingHintAntiAliasGridFit
    
    End Select
    
End Function

'If a GDI+ function fails, pass the return value to this function so the debugger can attempt to output a helpful message.
Private Sub ReportGDIPlusFailure(ByVal failCode As Long)

    'FYI, GDI+ failures are expected and frequent, so my interest in them is primarily academic at present.
    Select Case failCode
    
        Case GDIP_FONT_FAMILY_NOT_FOUND
            PDDebug.LogAction "(GDI+ error code analysis: font family wasn't located)"
            
        Case GDIP_FONT_STYLE_NOT_FOUND
            PDDebug.LogAction "(GDI+ error code analysis: requested font style wasn't found)"
            
        Case GDIP_FONT_NOT_TRUETYPE
            PDDebug.LogAction "(GDI+ error code analysis: requested font is not a TrueType font)"
            
        Case Else
            PDDebug.LogAction "(GDI+ failed for an unknown reason: #" & failCode & ")"
    
    End Select
    
End Sub

'Attempt to create a GDI font matching the current class settings
Private Function CreateGDIFont() As Boolean
    
    'If a GDI font handle already exists, free it.  Note that the font must have been unselected from a DC;
    ' this is typically handled by the renderer function itself.
    If (m_GDIFont <> 0) Then DeleteObject m_GDIFont
    
    'The font management module makes this task a lot easier.  Start by creating a LOGFONTW container.
    Dim tmpLogFont As LOGFONTW
    
    'Use the font management module to populate the struct
    Fonts.FillLogFontW_Basic tmpLogFont, m_FontFace, m_FontBold, m_FontItalic, m_FontUnderline, m_FontStrikeout
    Fonts.FillLogFontW_Size tmpLogFont, m_FontSize, m_FontSizeUnit
    Fonts.FillLogFontW_Quality tmpLogFont, ConvertTextAAToGdipTextHint
    
    'Attempt to create a matching GDI font
    CreateGDIFont = Fonts.CreateGDIFont(tmpLogFont, m_GDIFont)
    If (Not CreateGDIFont) Then PDDebug.LogAction "GDI font creation failed!"
    
End Function

'Use this function to render arbitrary text to an arbitrary DIB.  All class properties must be initialized
' *prior* to this point (correct behavior is not otherwise guaranteed).
Friend Function RenderTextToDIB(ByRef dstDIB As pdDIB, ByVal x1 As Single, ByVal y1 As Single, ByVal textWidth As Single, ByVal textHeight As Single) As Boolean
    
    'As a convenience to the user, create the font as necessary
    If (Not m_FontCacheClean) Then CreateFontObject
    
    'If the font cache *still* isn't clean, something went horribly wrong
    If (Not m_FontCacheClean) Then
        RenderTextToDIB = False
        Exit Function
    End If
    
    'Check for zero-length text
    If (LenB(m_CurrentText) = 0) Then
        RenderTextToDIB = True
        Exit Function
    End If
    
    'PD currently supports two rendering modes: WAPI (which uses GDI or GDI+ to render the text, and is thus limited by the settings
    ' those two engines provide), or an internal PHOTODEMON engine, which manually extracts font glyphs and renders them ourselves.
    If (m_RenderingEngine = te_WAPI) Then
        
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
                PDDebug.LogAction "GDI and GDI+ both failed; something is probably wrong with this font.  Consider blacklisting it...?"
                RenderTextToDIB = False
                
        End Select
            
        'Debug.Print "FYI - renderTextToDIB returned " & CStr(renderTextToDIB)
        
    ElseIf (m_RenderingEngine = te_PhotoDemon) Then
        
        Dim startTime As Currency
        If REPORT_TEXT_RENDER_TIMING Then VBHacks.GetHighResTime startTime
        
        'PD's internal font engine works by manually translating GDI glyphs into GDI+ paths, then rendering the paths manually.
        ' Because glyph geometry is very energy intensive, we use a separate class that caches glyphs as they are created.
        ' This saves us from having to translate glyphs more than once, which is a non-trivial boost on text with one or more
        ' repeat characters.
        
        'Start by creating a pdGlyphCollection instance, if one doesn't already exist.
        If (m_GlyphCollection Is Nothing) Then Set m_GlyphCollection = New pdGlyphCollection
        
        'Notify the glyph collection of our GDI font handle.  It obviously needs this in order to generate glyphs.
        m_GlyphCollection.NotifyOfGDIFontChoice m_GDIFont, m_FontUnderline, m_FontStrikeout
        
        If REPORT_TEXT_RENDER_TIMING Then
            PDDebug.LogAction "pdTextRenderer - init glyph collection: " & VBHacks.GetTimeDiffNowAsString(startTime)
            VBHacks.GetHighResTime startTime
        End If
        
        'Next, pass the glyph collection our target string.  The glyph collection will generate a list of required glyphs and
        ' automatically retrieve any glyph geometry it hasn't already calculated.  (For example, if the last request it received
        ' was "cat", but this request is "cats", only the 's' needs to be newly retrieved.)
        '
        'Also - because this step may modify the string (via remapping settings), we use a temporary copy of the current string.
        Dim tmpString As String
        tmpString = m_CurrentText
        m_GlyphCollection.BuildGlyphCollection tmpString, m_TextHinting, m_CharRemap
        
        If REPORT_TEXT_RENDER_TIMING Then
            PDDebug.LogAction "pdTextRenderer - build glyphs: " & VBHacks.GetTimeDiffNowAsString(startTime)
            VBHacks.GetHighResTime startTime
        End If
        
        'After all glyphs have been created, it's *very important* that we regain control over our GDI font, as we need it for
        ' the actual rendering step!
        m_GlyphCollection.RequestGDIFontRelease
        
        'With the glyph collection successfully assembled, proceed with rendering
        RenderTextToDIB = RenderTextToDIB_Glyphs(dstDIB, tmpString, x1, y1, textWidth, textHeight)
        
        If REPORT_TEXT_RENDER_TIMING Then
            PDDebug.LogAction "pdTextRenderer - render glyphs: " & VBHacks.GetTimeDiffNowAsString(startTime)
            VBHacks.GetHighResTime startTime
        End If
        
    'Failsafe
    Else
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
    If (GdipSetTextRenderingHint(dstGraphics, ConvertTextAAToGdipTextHint) <> GP_OK) Then PDDebug.LogAction "Failed to set text rendering hint " & m_TextAntialiasing
    
    'PD stores text contrast on a 0-10 scale, but GDI+ inexplicably uses a 0-12 scale.
    If (GdipSetTextContrast(dstGraphics, CLng(CDbl(m_TextContrast) * 1.2)) <> GP_OK) Then PDDebug.LogAction "Failed to set text contrast " & m_TextContrast
    
    'Half-pixel offset works better with GDI+ hinting strategy
    If Not GDI_Plus.SetGDIPlusGraphicsPixelOffset(dstGraphics, GP_POM_Half) Then PDDebug.LogAction "Failed to set pixel offset mode."
    
    'GDI+ may assume gamma-corrected blending which doesn't always play nicely
    ' with alpha premultiplication (because it is, by design, non-linear)
    GDI_Plus.SetGDIPlusGraphicsBlendUsingSRGBGamma dstGraphics, GP_CQ_AssumeLinear
    
    'Create a solid fill brush.  (In the future, we'll add more impressive options!)
    Dim gdipBrush As Long
    gdipBrush = GDI_Plus.GetGDIPlusSolidBrushHandle(crColor, 255)
    
    'Convert the input rect to a RECTF
    Dim boundingRect As RectF
    With boundingRect
        .Left = x1
        .Top = y1
        .Width = textWidth
        .Height = textHeight
    End With
    
    'Hypothetically, we could attempt to retrieve a bounding box for the string.  This would be relevant when
    ' compensating for overhang on the left side, for synthesized fonts like Times New Roman - Italic.
    ' Unfortunately, GDI+ is stupid and it doesn't report left-bound overhang, so our internal function
    ' GetStringBoundingBox() isn't helpful here.
    
    'We can, however, use GDI to calculate it instead.
    Dim firstCharABCWidth As ABCFLOAT
    If Fonts.GetABCWidthOfGlyph(m_GDIFont, AscW(Left$(srcString, 1)), firstCharABCWidth) Then
        
        'If the offset value is negative, we want to increase our rendering offset proportionally.  (Note that we add an
        ' extra pixel to account for possible antialiasing along the border.)
        If (firstCharABCWidth.abcfA < 0) Then
            
            Dim newLeft As Long, newWidth As Long
            newLeft = boundingRect.Left + Abs(firstCharABCWidth.abcfA) + 1
            boundingRect.Left = newLeft
            
            'Because fonts with a significant left overhang are likely to have a large right overhang as well,
            ' mirror the overhang to the right boundary as well.  (We do this because GDI+ word wrap is unpredictable,
            ' and we can't retrieve where it decides to wordwrap - so we have to make assumptions.)
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
        PDDebug.LogAction "WARNING!  GdipDrawString failed with error code " & gdipReturn
        PDDebug.LogAction "WARNING!  (Extra debug info: hFont - " & m_GDIPlusFont & ", hStringFormat - " & m_GDIPlusStringFormat & ")"
        RenderTextToDIB_GDIPlus = False
    End If
    
    'Release the temporary GDI+ objects we created
    GDI_Plus.ReleaseGDIPlusBrush gdipBrush
    GDI_Plus.ReleaseGDIPlusGraphics dstGraphics

End Function

'Render text using GDI+, and a glyph collection properly assembled inside m_GlyphCollection.
Private Function RenderTextToDIB_Glyphs(ByRef dstDIB As pdDIB, ByRef srcString As String, ByVal x1 As Single, ByVal y1 As Single, ByVal textWidth As Single, ByVal textHeight As Single) As Boolean

    'TESTING ONLY!  Fill with white prior to rendering
    'GDI_Plus.GDIPlusFillDIBRect dstDIB, 0, 0, dstDIB.getDIBWidth, dstDIB.getDIBHeight, vbWhite
    
    Dim startTime As Currency
    If REPORT_TEXT_RENDER_TIMING Then VBHacks.GetHighResTime startTime
    
    'Start by acquiring a GDI+ handle to the destination DIB
    Dim dstGraphics As Long
    dstGraphics = GDI_Plus.GetGDIPlusGraphicsFromDC(dstDIB.GetDIBDC, False)
    
    'Next, convert the input rect to a RECTF
    Dim boundingRect As RectF
    With boundingRect
        .Left = x1
        .Top = y1
        .Width = textWidth
        .Height = textHeight
    End With
    
    'Next, set up all container-specific path settings
    Dim targetSmoothingMode As GP_SmoothingMode, targetOffsetMode As GP_PixelOffsetMode
    
    'Compositing quality does not change; GDI+ is unreliable in its handling of compositing quality
    ' (which assumes gamma correction), and it can cause premultiplied alpha to be calculated
    ' incorrectly, leading to random hot pixels.
    GDI_Plus.SetGDIPlusGraphicsBlendUsingSRGBGamma dstGraphics, GP_CQ_AssumeLinear
    
    'Antialiasing settings currently affect both smoothing mode and offset mode.
    ' Offset mode is particularly relevant when working with tiny text, as it greatly improves clarity (when TRUE).
    ' (Also, note that antialiasing affects *all drawing operations on the target*, by design, so you cannot
    ' currently set different AA options for text vs background elements.)
    Select Case m_TextAntialiasing
        
        Case pdta_None
            targetSmoothingMode = GP_SM_None
            targetOffsetMode = GP_POM_None
        
        Case pdta_Standard
            targetSmoothingMode = GP_SM_Antialias
            targetOffsetMode = GP_POM_None
        
        Case pdta_Crisp, pdta_Smooth
            targetSmoothingMode = GP_SM_Antialias
            targetOffsetMode = GP_POM_Half
        
    End Select
    
    'Note that we don't actually set container settings yet, because we first want to draw some generic elements
    ' (like the background fill) at max speed.  Once those are finished, we'll set target smoothing and offset modes.
        
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
            .Left = .Left + penTextOutline.GetPenWidth()
            .Top = .Top + penTextOutline.GetPenWidth()
            .Width = .Width - penTextOutline.GetPenWidth() * 2
            .Height = .Height - penTextOutline.GetPenWidth() * 2
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
        
        borderWidth = penBackgroundOutline.GetPenWidth()
        halfBorderWidth = borderWidth / 2
        
        With boundingRect
            .Left = .Left + halfBorderWidth
            .Top = .Top + halfBorderWidth
            .Width = .Width - borderWidth
            .Height = .Height - borderWidth
        End With
        
    End If
    
    '4) After all boundary rect modifications are made, we need to perform a failsafe check
    ' to ensure width and height aren't negative.
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
    
    If REPORT_TEXT_RENDER_TIMING Then
        PDDebug.LogAction "RenderTextToDIB_Glyphs - init: " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    If m_GlyphCollection.AssembleCompositePath(finalTextPath, boundingRect, m_HorizontalAlignment, m_VerticalAlignment, m_WordWrap, m_StretchToFit, m_AlignLastLine) Then
        
        If REPORT_TEXT_RENDER_TIMING Then
            PDDebug.LogAction "RenderTextToDIB_Glyphs - assembly: " & VBHacks.GetTimeDiffNowAsString(startTime)
            VBHacks.GetHighResTime startTime
        End If
    
        Dim exactTextBounds As RectF
        exactTextBounds = finalTextPath.GetPathBoundariesF
        
        Dim gdipReturn As GP_Result
                
        'If background mode is active, fill it now
        If (gdipBackgroundBrush <> 0) Then
            
            'Prior to filling, set high-speed offset mode.
            GDI_Plus.SetGDIPlusGraphicsPixelOffset dstGraphics, GP_POM_None
            
            'If an outline is being drawn, we don't need antialiasing, which also improves performance
            If (gdipBackBorderPen <> 0) Then
                GdipSetSmoothingMode dstGraphics, GP_SM_HighSpeed
            Else
                GdipSetSmoothingMode dstGraphics, GP_SM_Antialias
            End If
            
            gdipReturn = GdipFillRectangleI(dstGraphics, gdipBackgroundBrush, x1, y1, textWidth, textHeight)
            If (gdipReturn <> GP_OK) Then PDDebug.LogAction "WARNING!  GdipFillPath failed; return code was " & gdipReturn
            
        End If
        
        'Now we can set our desired smoothing and offset modes, which are constant for all remaining draw operations
        GdipSetSmoothingMode dstGraphics, targetSmoothingMode
        GDI_Plus.SetGDIPlusGraphicsPixelOffset dstGraphics, targetOffsetMode
                
        'If background border is active, stroke it now
        If (gdipBackBorderPen <> 0) Then
            gdipReturn = GdipDrawRectangle(dstGraphics, gdipBackBorderPen, x1 + halfBorderWidth, y1 + halfBorderWidth, textWidth - borderWidth, textHeight - borderWidth)
            If (gdipReturn <> GP_OK) Then PDDebug.LogAction "WARNING!  GdipDrawRectangle failed; return code was " & gdipReturn
        End If
        
        If REPORT_TEXT_RENDER_TIMING Then
            PDDebug.LogAction "RenderTextToDIB_Glyphs - bkgd: " & VBHacks.GetTimeDiffNowAsString(startTime)
            VBHacks.GetHighResTime startTime
        End If
    
        'Before attempting to render the text itself, make sure a boundary rect exists.
        ' (Bad/corrupted font files may not produce usable glyph data.)
        If (boundingRect.Width > 0!) And (boundingRect.Height > 0!) Then
            
            'If fill mode is active, prep a pd2D fill brush
            If (gdipFillBrush <> 0) Then
            
                'Update the fill brush against the final rect established by the path renderer.
                ' (This additional step is necessary to ensure that gradients span the full width/height of the text.)
                textFillBrush.SetBrushPropertiesFromXML m_FillBrush
                boundingRect = finalTextPath.GetPathBoundariesF
                textFillBrush.SetBoundaryRect boundingRect
                gdipFillBrush = textFillBrush.GetHandle
            
            End If
            
            'We now vary rendering order according to the "outline on top" property.  (By default, PD renders a text outline,
            ' if any, *above* the fill - but the user can toggle this behavior.)
            'If outline mode is active *and* the user wants the outline on top, stroke the text now
            If (gdipOutlinePen <> 0) And (Not m_OutlineAboveFill) Then
                gdipReturn = GdipDrawPath(dstGraphics, gdipOutlinePen, finalTextPath.GetHandle)
                If (gdipReturn <> GP_OK) Then PDDebug.LogAction "WARNING!  GdipDrawPath failed; return code was " & gdipReturn
                If REPORT_TEXT_RENDER_TIMING Then
                    PDDebug.LogAction "RenderTextToDIB_Glyphs - stroke: " & VBHacks.GetTimeDiffNowAsString(startTime)
                    VBHacks.GetHighResTime startTime
                End If
            End If
            
            'Fill (if any)
            If (gdipFillBrush <> 0) Then
                
                gdipReturn = GdipFillPath(dstGraphics, gdipFillBrush, finalTextPath.GetHandle)
                If (gdipReturn <> GP_OK) Then PDDebug.LogAction "WARNING!  GdipFillPath failed; return code was " & gdipReturn
                
                'If the user wants "smooth" antialiasing, we can achieve this by performing a fast scaledown+scaleup op on the
                ' filled surface.  This softens antialiased boundaries slightly, with a minimal penalty to perf.
                If (m_TextAntialiasing = pdta_Smooth) Then
                    
                    Dim tmpDIB As pdDIB
                    Set tmpDIB = New pdDIB
                    
                    Const SMOOTH_SCALE_FACTOR As Single = 0.95!
                    
                    Dim tmpWidth As Long, tmpHeight As Long
                    tmpWidth = Int(dstDIB.GetDIBWidth * SMOOTH_SCALE_FACTOR + 0.5)
                    If (tmpWidth >= dstDIB.GetDIBWidth) Then tmpWidth = dstDIB.GetDIBWidth - 1
                    If (tmpWidth <= 0) Then tmpWidth = dstDIB.GetDIBWidth
                    
                    tmpHeight = Int(dstDIB.GetDIBHeight * SMOOTH_SCALE_FACTOR + 0.5)
                    If (tmpHeight >= dstDIB.GetDIBHeight) Then tmpHeight = dstDIB.GetDIBHeight - 1
                    If (tmpHeight <= 0) Then tmpHeight = dstDIB.GetDIBHeight
                    
                    tmpDIB.CreateBlank tmpWidth, tmpHeight, 32, 0, 0
                    tmpDIB.SetInitialAlphaPremultiplicationState True
                    
                    GDI_Plus.GDIPlus_StretchBlt tmpDIB, 0, 0, tmpWidth, tmpHeight, dstDIB, 0, 0, dstDIB.GetDIBWidth, dstDIB.GetDIBHeight, interpolationType:=GP_IM_HighQualityBicubic, isZoomedIn:=True, dstCopyIsOkay:=True
                    dstDIB.ResetDIB 0
                    GDI_Plus.GDIPlus_StretchBlt dstDIB, 0, 0, dstDIB.GetDIBWidth, dstDIB.GetDIBHeight, tmpDIB, 0, 0, tmpWidth, tmpHeight, interpolationType:=GP_IM_HighQualityBicubic, isZoomedIn:=True, dstCopyIsOkay:=True
                    Set tmpDIB = Nothing
                    
                End If
                
                If REPORT_TEXT_RENDER_TIMING Then
                    PDDebug.LogAction "RenderTextToDIB_Glyphs - fill: " & VBHacks.GetTimeDiffNowAsString(startTime)
                    VBHacks.GetHighResTime startTime
                End If
                
            End If
            
            'If outline mode is active *and* the user wants the outline on top, stroke the text *now*
            If (gdipOutlinePen <> 0) And m_OutlineAboveFill Then
                gdipReturn = GdipDrawPath(dstGraphics, gdipOutlinePen, finalTextPath.GetHandle)
                If (gdipReturn <> GP_OK) Then PDDebug.LogAction "WARNING!  GdipDrawPath failed; return code was " & gdipReturn
                If REPORT_TEXT_RENDER_TIMING Then
                    PDDebug.LogAction "RenderTextToDIB_Glyphs - stroke: " & VBHacks.GetTimeDiffNowAsString(startTime)
                    VBHacks.GetHighResTime startTime
                End If
            End If
            
        End If
        
    Else
        PDDebug.LogAction "WARNING!  m_GlyphCollection.AssembleCompositePath returned FALSE.  Please investigate."
    End If
    
    'Release the temporary GDI+ objects we created.  (Note that GDI+ objects silently created via pd2D objects are released automatically.)
    GDI_Plus.ReleaseGDIPlusGraphics dstGraphics
    
    'Return success/failure contingent on the return of the GDI+ rendering call
    RenderTextToDIB_Glyphs = (gdipReturn = GP_OK)
    
End Function

'Render text using GDI.  This only exists as a fallback if GDI+ rendering fails (which is far more common pre-Windows-10).
' This function is unlikely to be called on Win 10+ due to improvements in GDI+ font handling.
Private Function RenderTextToDIB_GDI(ByRef dstDIB As pdDIB, ByRef srcString As String, ByVal crColor As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal textWidth As Single, ByVal textHeight As Single) As Boolean

    'Because GDI doesn't support 32-bpp rendering targets, we must do all rendering to a 24-bpp surface, then copy it over manually
    If (m_tmpDIB Is Nothing) Then
        Set m_tmpDIB = New pdDIB
        m_tmpDIB.CreateBlank 4, 4, 24, 0
    End If
    
    'Start by setting up all container-specific text settings; note that these are only used in our temporary DIB!
    Dim oldFont As Long
    oldFont = SelectObject(m_tmpDIB.GetDIBDC, m_GDIFont)
    
    'Set other font parameters.  Note that this function only draws white text onto a black background.  Color is applied later.
    SetTextColor m_tmpDIB.GetDIBDC, RGB(255, 255, 255)
    SetTextAlign m_tmpDIB.GetDIBDC, 0
    
    'Enable transparent font rendering
    SetBkMode m_tmpDIB.GetDIBDC, FONT_TRANSPARENT
    
    'You'd think drawing opaque text to a 32-bpp container would work, but nope, this just means that
    ' GDI functions make the ENTIRE TEXT AREA transparent. THANKS, MICROSOFT!
    'SetBkMode dstDIB.getDIBDC, FONT_OPAQUE
    'SetBkColor dstDIB.getDIBDC, vbWhite
    
    'We now need to figure out the size of our drawn text.  Wordwrap is supported if the text extends
    ' past the horizontal boundary of the target DIB.
    
    'Start by retrieving default size values
    Dim targetWidth As Long, targetHeight As Long
    Dim txtSize As PointAPI
    GetTextExtentPoint32 m_tmpDIB.GetDIBDC, StrPtr(srcString), Len(srcString), txtSize
    
    Dim tmpRect As RECT
    tmpRect.Left = 0
    
    'A left-side overhang is possible on some fonts, particularly fonts where italics have to be synthesized.
    ' There's no automatic way to handle this, so we do it manually (as usual).
    Dim firstCharABCWidth As ABCFLOAT
    If Fonts.GetABCWidthOfGlyph(m_tmpDIB.GetDIBDC, AscW(Left$(srcString, 1)), firstCharABCWidth, True) Then
        
        'If the offset value is negative, we want to increase our rendering offset proportionally.
        ' (Note that we add an extra pixel to account for possible antialiasing along the border.)
        If (firstCharABCWidth.abcfA < 0) Then
            Dim newLeft As Long
            newLeft = tmpRect.Left + Abs(firstCharABCWidth.abcfA) + 1
        End If
        
    End If
    
    'If the width of this string (in the current font) is less than the width of the target DIB,
    ' we will use the string's precise render dimensions for our target DIB.
    If (txtSize.x + tmpRect.Left < dstDIB.GetDIBWidth) Then
        
        'GDI adds 1/6 em width to account for glyph overhang.  Italic mode doubles the default overhang
        ' (and we apply that correction by default, to cover fonts where italics is the only available mode).
        targetWidth = txtSize.x + CLng((CDbl(txtSize.y) / 6 + 1) * 2) + tmpRect.Left
        If targetWidth > dstDIB.GetDIBWidth Then targetWidth = dstDIB.GetDIBWidth
        
        targetHeight = txtSize.y + 1
    
    'The string extends beyond the width of the DIB. Wordwrap is required.  Calculate height now.
    Else
    
        targetWidth = dstDIB.GetDIBWidth - 1
        
        tmpRect.Top = 0
        tmpRect.Bottom = 0
        tmpRect.Right = targetWidth
        
        DrawText m_tmpDIB.GetDIBDC, StrPtr(srcString), Len(srcString), tmpRect, GetDrawTextAlignmentFlags() Or DT_CALCRECT Or DT_WORDBREAK Or DT_EXTERNALLEADING Or DT_NOPREFIX
        
        'Take the height of the rendered string, or the height of the target DIB, whichever is smaller
        If (tmpRect.Bottom > dstDIB.GetDIBHeight) Then
            targetHeight = dstDIB.GetDIBHeight
        Else
            targetHeight = tmpRect.Bottom + 1
        End If
    
    End If
    
    'Recreate the target DIB, as necessary
    If (m_tmpDIB.GetDIBWidth <> targetWidth) Or (m_tmpDIB.GetDIBHeight <> targetHeight) Then
    
        'Remove the font from the temporary DC, because we are about to create it.
        SelectObject m_tmpDIB.GetDIBDC, oldFont
    
        'Create the temporary DIB.  We know this will be smaller or equal to the size of the target DIB,
        ' greatly simplifying the eventual transfer process.
        m_tmpDIB.CreateBlank targetWidth, targetHeight, 24
        m_tmpDIB.ResetDIB
        
        'Re-select the font into the DIB
        oldFont = SelectObject(m_tmpDIB.GetDIBDC, m_GDIFont)
        
        'Reset other font parameters.
        SetTextColor m_tmpDIB.GetDIBDC, RGB(255, 255, 255)
        SetTextAlign m_tmpDIB.GetDIBDC, 0
        SetBkMode m_tmpDIB.GetDIBDC, FONT_TRANSPARENT
        
    Else
        m_tmpDIB.ResetDIB
    End If
    
    'Populate our clipping rect with the final width and height values we generated
    tmpRect.Right = targetWidth
    tmpRect.Bottom = targetHeight
    
    'If overhang is present on the first glyph, it's likely to be present on other glyphs as well.
    ' Mirror the left overhang (if any) across the right border.  (We do this because GDI word wrap
    ' is unpredictable, and we can't retrieve where it decides to wordwrap - so we can only make assumptions.)
    If (firstCharABCWidth.abcfA < 0) Then
        Dim newRight As Long
        newRight = tmpRect.Right - (Abs(firstCharABCWidth.abcfC) + 1)
        If (newRight > tmpRect.Left) Then tmpRect.Right = newRight
    End If
    
    'Use DrawText for the actual rendering
    Dim retDrawText As Long
    retDrawText = DrawText(m_tmpDIB.GetDIBDC, StrPtr(srcString), Len(srcString), tmpRect, DT_WORDBREAK Or DT_NOCLIP Or GetDrawTextAlignmentFlags())
    
    'We can actually determine a return value now
    RenderTextToDIB_GDI = (retDrawText <> 0)
    
    'Remove the font, so we can use again it later!
    SelectObject m_tmpDIB.GetDIBDC, oldFont
    
    'With the temporary DIB successfully renderered, our new job is to transfer its contents to the destination DIB.
    ' Unfortunately, we have to do this manually, using per-pixel code.
    
    'Start by clearing the destination DIB
    dstDIB.ResetDIB
    dstDIB.SetInitialAlphaPremultiplicationState True
        
    'Next, we need to prepare lookup tables based on the translation of gray pixels (since we rendered white text
    ' to a black background) to the target color, with alpha accounted for.  Note that alpha is easy - the grayscale
    ' value of the temporary DIB represents alpha - so we just need to calculate premultiplied colors.
    
    'Start by extracting individual RGB components from the text color.
    Dim r As Long, g As Long, b As Long, a As Long
    r = Colors.ExtractRed(crColor)
    g = Colors.ExtractGreen(crColor)
    b = Colors.ExtractBlue(crColor)
    
    'Build RGB lookup tables for premultiplied alpha.
    ' (Note that we must also calculate an alpha lookup table, so we can manually implement
    '  GDI+-specific features like "text clarity".)
    Dim clrLookup(0 To 255) As RGBQuad
    
    Dim x As Long, y As Long
    Dim preMultiplicationFactor As Double, textContrastFactor As Double, alphaCalculation As Double
    
    '"standard" and "crisp" AA modes need to mimic the Text Contrast setting GDI+ provides.
    ' We use a custom calculation that provides similar results.
    
    'Convert m_TextContrast from the range [0, 10] to [0.5, 3.0]
    textContrastFactor = (m_TextContrast / 4) + 0.5
    
    For x = 0 To 255
        
        a = x
        
        'Start by calculating alpha.  This varies according to the current antialiasing mode and clarity.
        Select Case m_TextAntialiasing
        
            Case pdta_None
                
                'Split alpha to strictly 0 or 255 values.  (Also, ignore text contrast completely.)
                If (a < 127) Then
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
                If (a < 0) Then
                    a = 0
                ElseIf (a > 255) Then
                    a = 255
                End If
            
            Case pdta_Crisp
            
                'Crisp is almost identical to regular alpha, except we convert alpha to a ^ 2 curve in advance
                alphaCalculation = a / 255
                alphaCalculation = alphaCalculation * alphaCalculation
                alphaCalculation = alphaCalculation ^ textContrastFactor
                
                'Convert back to the [0, 255] range
                a = alphaCalculation * 255
                If (a < 0) Then
                    a = 0
                ElseIf (a > 255) Then
                    a = 255
                End If
        
        End Select
        
        preMultiplicationFactor = a / 255#
        
        With clrLookup(x)
            .Red = Int(r * preMultiplicationFactor)
            .Green = Int(g * preMultiplicationFactor)
            .Blue = Int(b * preMultiplicationFactor)
            .Alpha = a
        End With
        
    Next x
    
    'Lookup tables are now ready.  Time to transfer the bits!
        
    'Create a local array and point it at the pixel data of the target image
    Dim dstImageData() As RGBQuad, dstSA As SafeArray1D
    
    'Create a second local array and point it as the pixel data of our temporary image
    Dim srcImageData() As Byte, srcSA As SafeArray2D
    m_tmpDIB.WrapArrayAroundDIB srcImageData, srcSA
    
    Dim finalX As Long, finalY As Long
    finalX = m_tmpDIB.GetDIBWidth - 1
    finalY = m_tmpDIB.GetDIBHeight - 1
    
    'As a failsafe, make sure finalX and Y don't extend beyond our target DIB, either
    If (finalX >= dstDIB.GetDIBWidth) Then finalX = dstDIB.GetDIBWidth - 1
    If (finalY >= dstDIB.GetDIBHeight) Then finalY = dstDIB.GetDIBHeight - 1
    
    'Finally, we also need to calculate x/y offset values.  For performance reasons, this class renders text
    ' to a temporary DIB at the exact size of the text's bounding rect.  Because we must manually convert these
    ' results to 32-bpp, then transfer them to the destination image, the alignment of the temporary DIB vs the
    ' destination DIB varies according to the underlying text alignment.
    
    'Calculate offsets now
    Dim xOffset As Long, yOffset As Long
    
    Select Case m_HorizontalAlignment
        
        'Justified text does *not* work with this renderer
        Case StringAlignmentNear, StringAlignmentJustify
            xOffset = 0
        
        Case StringAlignmentCenter
            xOffset = (dstDIB.GetDIBWidth - m_tmpDIB.GetDIBWidth) \ 2
        
        Case StringAlignmentFar
            xOffset = dstDIB.GetDIBWidth - m_tmpDIB.GetDIBWidth
    
    End Select
    
    Select Case m_VerticalAlignment
    
        Case StringAlignmentNear
            yOffset = 0
        
        Case StringAlignmentCenter
            yOffset = (dstDIB.GetDIBHeight - m_tmpDIB.GetDIBHeight) \ 2
        
        Case StringAlignmentFar
            yOffset = dstDIB.GetDIBHeight - m_tmpDIB.GetDIBHeight
    
    End Select
    
    'Add failsafe offset checks
    If (xOffset < 0) Then xOffset = 0
    If (yOffset < 0) Then yOffset = 0
    
    'Start transferring pixels
    For y = 0 To finalY
        dstDIB.WrapRGBQuadArrayAroundScanline dstImageData, dstSA, y + yOffset
    For x = 0 To finalX
        
        'Because we know the source DIB is grayscale, we can skip processing of all pixels with value = 0.
        g = srcImageData(x * 3, y)
        If (g <> 0) Then dstImageData(x + xOffset) = clrLookup(g)
        
    Next x
    Next y
    
    'Deallocate both arrays and exit
    dstDIB.UnwrapRGBQuadArrayFromDIB dstImageData
    m_tmpDIB.UnwrapArrayFromDIB srcImageData
    
End Function

'Convert the current alignment flags into DrawText-compatible alignment values
Private Function GetDrawTextAlignmentFlags() As Win32_DrawTextConstants

    GetDrawTextAlignmentFlags = 0
    
    'Start with horizontal flags
    Select Case m_HorizontalAlignment
        
        Case StringAlignmentNear
            GetDrawTextAlignmentFlags = GetDrawTextAlignmentFlags Or DT_LEFT
        
        Case StringAlignmentCenter
            GetDrawTextAlignmentFlags = GetDrawTextAlignmentFlags Or DT_CENTER
        
        Case StringAlignmentFar
            GetDrawTextAlignmentFlags = GetDrawTextAlignmentFlags Or DT_RIGHT
            
        'PhotoDemon can render justified text, but Win32 cannot.
        Case Else
            GetDrawTextAlignmentFlags = GetDrawTextAlignmentFlags Or DT_LEFT
        
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

'If this class is no longer required but the caller doesn't want to fully release it, they can call this sub
' to have us release all caches and system handles (hFonts, GDI+ fonts, etc).  Note that the class needs to be
' FULLY reinitialized after this function is called.
Friend Sub ReleaseAsManyResourcesAsPossible()
    
    'Release any GDI+ objects that are easily re-created
    If (m_GDIPlusFont <> 0) Then
        GdipDeleteFont m_GDIPlusFont
        m_GDIPlusFont = 0
    End If
    
    'Private font families are families created user-added font files.  They are maintained separately
    ' and freed via special APIs, so *do not* release them via GdipDeleteFontFamily.
    If (m_GDIPlusFontFamily <> 0) Then
        If (Not m_GDIPlusFontFamilyIsPrivate) Then GdipDeleteFontFamily m_GDIPlusFontFamily
        m_GDIPlusFontFamily = 0
    End If
    
    'Release any GDI objects that are easily re-created
    If (m_GDIFont <> 0) Then
        DeleteObject m_GDIFont
        m_GDIFont = 0
    End If
    
    'Free our glyph collection, if one exists
    Set m_GlyphCollection = Nothing
    
End Sub

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String)
    If UserPrefs.GenerateDebugLogs Then
        PDDebug.LogAction "pdTextRenderer." & funcName & "() reported an error: " & errDescription
    Else
        Debug.Print "pdTextRenderer." & funcName & "() reported an error: " & errDescription
    End If
End Sub

Private Sub Class_Initialize()
    
    m_RenderingEngine = te_WAPI
        
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
    m_FontSizeUnit = fu_Pixel
    
    'Set some default font properties
    m_FontFace = Fonts.GetUIFontName()
    m_FontSize = 16# '16px = 12 points at 96 DPI
    
    m_HorizontalAlignment = StringAlignmentNear
    m_VerticalAlignment = StringAlignmentNear
    
    m_WordWrap = tww_AutoWord
    m_FillActive = True
    m_StretchToFit = stf_None
    m_OutlineAboveFill = True
    m_TextHinting = True
    
    'Set default graphics container text settings
    m_TextContrast = 4
    m_TextAntialiasing = pdta_Standard

End Sub

Private Sub Class_Terminate()
    
    'Release any GDI+ objects we created
    If (m_GDIPlusFont <> 0) Then GdipDeleteFont m_GDIPlusFont
    
    'Private font families are families created user-added font files.  They are maintained separately
    ' and freed via special APIs, so *do not* release them via GdipDeleteFontFamily.
    If (m_GDIPlusFontFamily <> 0) And (Not m_GDIPlusFontFamilyIsPrivate) Then GdipDeleteFontFamily m_GDIPlusFontFamily
    If (m_GDIPlusStringFormat <> 0) Then GdipDeleteStringFormat m_GDIPlusStringFormat
    
    'Release any GDI objects we created
    If (m_GDIFont <> 0) Then DeleteObject m_GDIFont
    
End Sub




'**********************************************************************************************
'Legacy support functions follow.  Do NOT modify these functions except to fix security issues.

'Set all text and font parameters at once, using an XML string created by a previous call to GetAllFontSettingsAsXML.
' (Note that this function manually marks the font cache as dirty, so you should only use it when loading a file for the first time.)
Private Function SetAllFontSettingsFromXML_Legacy(ByRef srcXMLString As String) As Boolean
    
    'We don't need a full-featured XML creator (e.g. with encoding decs and the like), so we can use the smaller, lighter pdSerialize class
    Dim cSerialize As pdSerialize
    Set cSerialize = New pdSerialize
    
    With cSerialize
        .SetParamString srcXMLString
        
        SetGenericTextProperty ptp_Text, .GetString("TextLayerText", vbNullString)
        SetGenericTextProperty ptp_FontColor, .GetLong("TextLayerColor", 0)
        SetGenericTextProperty ptp_FontFace, .GetString("TextLayerFontFace", Fonts.GetUIFontName())
        SetGenericTextProperty ptp_FontSize, .GetDouble("TextLayerFontSize", 16#)
        SetGenericTextProperty ptp_FontSizeUnit, .GetLong("TextLayerFontSizeUnit", fu_Pixel)
        SetGenericTextProperty ptp_FontBold, .GetBool("TextLayerFontBold", False)
        SetGenericTextProperty ptp_FontItalic, .GetBool("TextLayerFontItalic", False)
        SetGenericTextProperty ptp_FontUnderline, .GetBool("TextLayerFontUnderline", False)
        SetGenericTextProperty ptp_FontStrikeout, .GetBool("TextLayerFontStrikeout", False)
        SetGenericTextProperty ptp_HorizontalAlignment, .GetLong("TextLayerHorizontalAlignment", StringAlignmentNear)
        SetGenericTextProperty ptp_VerticalAlignment, .GetLong("TextLayerVerticalAlignment", StringAlignmentNear)
        SetGenericTextProperty ptp_TextAntialiasing, .GetLong("TextLayerAntialiasing", pdta_Standard)
        SetGenericTextProperty ptp_TextContrast, .GetLong("TextLayerContrast", 5)
        SetGenericTextProperty ptp_RenderingEngine, .GetLong("TextRenderingEngine", te_WAPI)
        SetGenericTextProperty ptp_TextHinting, .GetBool("TextHinting", False)
        SetGenericTextProperty ptp_WordWrap, .GetLong("TextWordWrap", tww_AutoWord)
        SetGenericTextProperty ptp_StretchToFit, stf_None
        SetGenericTextProperty ptp_FillActive, .GetBool("TextFillActive", True)
        SetGenericTextProperty ptp_FillBrush, .GetString("TextFillBrush", vbNullString)
        SetGenericTextProperty ptp_OutlineActive, .GetBool("TextOutlineActive", False)
        SetGenericTextProperty ptp_OutlinePen, .GetString("TextOutlinePen", vbNullString)
        SetGenericTextProperty ptp_BackgroundActive, .GetBool("TextBackgroundActive", False)
        SetGenericTextProperty ptp_BackgroundBrush, .GetString("TextBackgroundBrush", vbNullString)
        SetGenericTextProperty ptp_BackBorderActive, .GetBool("TextBackBorderActive", False)
        SetGenericTextProperty ptp_BackBorderPen, .GetString("TextBackBorderPen", vbNullString)
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
        SetGenericTextProperty ptp_AlignLastLine, StringAlignmentNear
        SetGenericTextProperty ptp_OutlineAboveFill, True
    End With
    
    'This function does not currently provide a fail state; as long as the load request comes from PD herself, failure should be impossible.
    SetAllFontSettingsFromXML_Legacy = True
    
End Function
