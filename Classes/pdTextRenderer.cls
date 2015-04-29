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
'Copyright 2015-2015 by Tanner Helland
'Created: 22/April/15
'Last updated: 26/April/15
'Last update: migrate the rendering-specific bits out of this class, and into pdTextBackend_WAPI
'
'This class handles the messy business of font management for PD's new text layers.  It switches between GDI and GDI+ as
' necessary to provide a comprehensive text rendering solution.
'
'I had originally planned to split this class in two, as a GDI-specific backend and a separate GDI+-specific backend.
' But there are complications with this.  GDI and GDI+ have different strengths and weaknesses (GDI has no 32-bpp rendering
' support, while GDI+ doesn't support OpenType), so to really provide a decent font experience without native libraries,
' one must utilize both libraries.
'
'Hence this class.  Here's the basic setup:
'
'For TrueType fonts, GDI+ is used.  This saves us a lot of grief, and lets us make use of some nice, typographic-specific
' features (like variable antialiasing contrast).
'
'For OpenType fonts, GDI is used out of necessity.  Individual glyphs are retrieved, parsed, and converted into a usable
' format (e.g. OpenType's quadratic B-splines are converted into GDI+'s native cubic Bézier curve format).  The final,
' resultant shape(s) are then passed to GDI+ as a generic path object, which we can render with any of GDI+'s native
' features (antialiasing, fancy gradients, etc).
'
'This unpleasant hybrid approach is part of the shit we deal with as Windows developers, I guess.
'
'Dependencies:
'  - Font_Management module, which handles some of the messier bits of GDI interaction
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'At present, only a WAPI font rendering backend is available.  This may change in the future.
Private Enum fontRenderingBackend
    frb_WAPI = 0
End Enum

#If False Then
    Private Const frb_WAPI = 0
#End If

Private m_CurrentFontBackend As fontRenderingBackend

'Because I want switches between GDI and GDI+ to be transparent, this class will dynamically track which engine is required
' for rendering.  If a setting changes (and renders the font cache invalid), this value will be reset to UNKNOWN.  On the next
' font object creation attempt, the required backend will be noted.
Private Enum textEngineRequired
    te_UNKNOWN = 0
    te_GDIPLUS = 1
    te_GDI = 2
End Enum

#If False Then
    Private Const te_UNKNOWN = 0, te_GDIPLUS = 1, te_GDI = 2
#End If

Private m_curTextEngine As textEngineRequired

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

'Compositing quality constants; there are really only three settings: HighSpeed, AssumeLinear, and GammaCorrected
' (in that order, going from fast performance to slow performance)
Private Enum GdiPlusCompositingQuality
    CompositingQualityInvalid = QualityModeInvalid
    CompositingQualityDefault = QualityModeDefault
    CompositingQualityHighSpeed = QualityModeLow
    CompositingQualityHighQuality = QualityModeHigh
    CompositingQualityGammaCorrected
    CompositingQualityAssumeLinear
End Enum

#If False Then
    Private Const CompositingQualityInvalid = QualityModeInvalid, CompositingQualityDefault = QualityModeDefault, CompositingQualityHighSpeed = QualityModeLow
    Private Const CompositingQualityHighQuality = QualityModeHigh, CompositingQualityGammaCorrected = 0, CompositingQualityAssumeLinear = 0
#End If

'SmoothingMode controls generic AA operations for pens
Private Enum GdiPlusSmoothingMode
    SmoothingModeInvalid = QualityModeInvalid
    SmoothingModeDefault = QualityModeDefault
    SmoothingModeHighSpeed = QualityModeLow
    SmoothingModeHighQuality = QualityModeHigh
    SmoothingModeNone
    SmoothingModeAntiAlias
End Enum

#If False Then
    Private Const SmoothingModeInvalid = QualityModeInvalid, SmoothingModeDefault = QualityModeDefault, SmoothingModeHighSpeed = QualityModeLow, SmoothingModeHighQuality = QualityModeHigh
    Private Const SmoothingModeNone = 0, SmoothingModeAntiAlias = 0
#End If

'Pixel offset mode controls the "theoretical" positioning of pixel commands.  Half mode treats pixel coordinates as occurring "between" physical pixels,
' which can improve AA output
Private Enum GdiPlusPixelOffsetMode
    PixelOffsetModeInvalid = QualityModeInvalid
    PixelOffsetModeDefault = QualityModeDefault
    PixelOffsetModeHighSpeed = QualityModeLow
    PixelOffsetModeHighQuality = QualityModeHigh
    PixelOffsetModeNone
    PixelOffsetModeHalf
End Enum

#If False Then
    Private Const PixelOffsetModeInvalid = QualityModeInvalid, PixelOffsetModeDefault = QualityModeDefault, PixelOffsetModeHighSpeed = QualityModeLow, PixelOffsetModeHighQuality = QualityModeHigh
    Private Const PixelOffsetModeNone = 0, PixelOffsetModeHalf = 0
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
Private Declare Function GdipStringFormatGetGenericTypographic Lib "gdiplus" (ByRef dstStringFormat As Long) As Long
Private Declare Function GdipDeleteStringFormat Lib "gdiplus" (ByVal srcStringFormat As Long) As Long
Private Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal dstStringFormat As Long, ByVal newAlignment As GdiPlusStringAlignment) As Long
Private Declare Function GdipSetStringFormatLineAlign Lib "gdiplus" (ByVal dstStringFormat As Long, ByVal newLineAlignment As GdiPlusStringAlignment) As Long

'GDI+ graphics container font functions
Private Declare Function GdipSetTextRenderingHint Lib "gdiplus" (ByVal dstGraphics As Long, ByVal newRenderHintMode As GdiPlusTextRenderingHint) As Long
Private Declare Function GdipSetTextContrast Lib "gdiplus" (ByVal dstGraphics As Long, ByVal textContrast As Long) As Long
Private Declare Function GdipSetCompositingQuality Lib "gdiplus" (ByVal dstGraphics As Long, ByVal newCompositingQuality As GdiPlusCompositingQuality) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal dstGraphics As Long, ByVal newSmoothingMode As GdiPlusSmoothingMode) As Long
Private Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal dstGraphics As Long, ByVal newPixelOffsetMode As GdiPlusPixelOffsetMode) As Long

'GDI+ Render functions
Private Declare Function GdipDrawString Lib "gdiplus" (ByVal dstGraphics As Long, ByVal ptrToString As Long, ByVal strLength As Long, ByVal gdipFontHandle As Long, ByRef layoutRect As RECTF, ByVal gdipStringFormat As Long, ByVal gdipBrush As Long) As Long
Private Declare Function GdipDrawPath Lib "gdiplus" (ByVal dstGraphics As Long, ByVal srcPen As Long, ByVal srcPath As Long) As Long
Private Declare Function GdipFillPath Lib "gdiplus" (ByVal dstGraphics As Long, ByVal srcBrush As Long, ByVal srcPath As Long) As Long

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

'This class internally maintains various font properties.  These properties must be converted into specific GDI+ parameters via
' various means, but they are stored in human-friendly format to simplify serializing a class instance to an XML string.
Private m_FontFace As String
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

'TODO: expose a PD-specific hinting + AA solution, rather than GDI+'s very confusing options.  (I'm waiting until I've
' better tested the various options before making this change.)
Private m_TextRenderingHint As GdiPlusTextRenderingHint
Private m_TextContrast As Long

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

'Get functions for various font styles, and whether said styles are supported by the current font.
' (If no font name has been set, the function returns *will not be valid*)
Public Function getFontBold() As Boolean
    getFontBold = m_FontBold
End Function

Public Function isFontBoldSupported() As Boolean
    isFontBoldSupported = m_BoldSupported
End Function

Public Function getFontItalic() As Boolean
    getFontItalic = m_FontItalic
End Function

Public Function isFontItalicSupported() As Boolean
    isFontItalicSupported = m_ItalicSupported
End Function

Public Function getFontUnderline() As Boolean
    getFontUnderline = m_FontUnderline
End Function

Public Function isFontUnderlineSupported() As Boolean
    isFontUnderlineSupported = m_UnderlineSupported
End Function

Public Function getFontStrikeout() As Boolean
    getFontStrikeout = m_FontStrikeout
End Function

Public Function isFontStrikeoutSupported() As Boolean
    isFontStrikeoutSupported = m_StrikeoutSupported
End Function

'Set functions for various font styles.  Note that these functions reset the current font cache, so please use them
' only if necessary.
Public Sub setFontBold(ByVal newValue As Boolean)
    If newValue <> m_FontBold Then
        m_FontBold = newValue
        m_FontCacheClean = False
    End If
End Sub

Public Sub setFontItalic(ByVal newValue As Boolean)
    If newValue <> m_FontItalic Then
        m_FontItalic = newValue
        m_FontCacheClean = False
    End If
End Sub

Public Sub setFontUnderline(ByVal newValue As Boolean)
    If newValue <> m_FontUnderline Then
        m_FontUnderline = newValue
        m_FontCacheClean = False
    End If
End Sub

Public Sub setFontStrikeout(ByVal newValue As Boolean)
    If newValue <> m_FontStrikeout Then
        m_FontStrikeout = newValue
        m_FontCacheClean = False
    End If
End Sub

'Get/set functions for font size and unit.  By default, this class uses pixels.  Other units may be supported
' in the future.
Public Function getFontSize() As Single
    getFontSize = m_FontSize
End Function

Public Sub setFontSize(ByVal newValue As Single)
    If newValue <> m_FontSize Then
        m_FontSize = newValue
        m_FontCacheClean = False
    End If
End Sub

Friend Function getFontSizeUnit() As pdFontUnit
    getFontSizeUnit = m_FontSizeUnit
End Function

Friend Sub setFontSizeUnit(ByVal newUnit As pdFontUnit)
    If newUnit <> m_FontSizeUnit Then
        m_FontSizeUnit = newUnit
        m_FontCacheClean = False
    End If
End Sub

'Get/set string alignment.  Note that unlike other settings, these functions have two significant features:
' 1) These Set functions immediately update a corresponding GDI+ StringFormat object.
' 2) These Set functions do not affect the current font cache
Friend Function getHorizontalAlignment() As GdiPlusStringAlignment
    getHorizontalAlignment = m_HorizontalAlignment
End Function

Friend Sub setHorizontalAlignment(ByVal newAlignment As GdiPlusStringAlignment)
    
    If m_HorizontalAlignment <> newAlignment Then
        m_HorizontalAlignment = newAlignment
        
        'If using GDI+, set this format immediately, as it's inexpensive to do so
        If m_CurrentFontBackend = frb_WAPI Then
            GdipSetStringFormatAlign m_GDIPlusStringFormat, newAlignment
        End If
        
    End If
    
End Sub

Friend Function getVerticalAlignment() As GdiPlusStringAlignment
    getVerticalAlignment = m_VerticalAlignment
End Function

Friend Sub setVerticalAlignment(ByVal newAlignment As GdiPlusStringAlignment)
    
    If m_VerticalAlignment <> newAlignment Then
        m_VerticalAlignment = newAlignment
        
        'If using GDI+, set this format immediately, as it's inexpensive to do so
        If m_CurrentFontBackend = frb_WAPI Then
            GdipSetStringFormatLineAlign m_GDIPlusStringFormat, newAlignment
        End If
        
    End If
    
End Sub

'Get/set container-specific font settings.  These are stored in normal variables, but they cannot be passed to GDI+
' until rendering time, as they are container-specific.  As such, they do not affect the current font cache.
Friend Function getTextRenderingHint() As GdiPlusTextRenderingHint
    getTextRenderingHint = m_TextRenderingHint
End Function

Friend Sub setTextRenderingHint(ByVal newRenderingHint As GdiPlusTextRenderingHint)
    m_TextRenderingHint = newRenderingHint
End Sub

'Text contrast is a weird one: the values vary from 0 to 12, and 4 is the default.
Public Function getTextContrast() As Long
    getTextContrast = m_TextContrast
End Function

Public Sub setTextContrast(ByVal newContrast As Long)
    
    'Force contrast into valid bounds
    If newContrast < 0 Then
        newContrast = 0
    ElseIf newContrast > 12 Then
        newContrast = 12
    End If
    
    m_TextContrast = newContrast
    
End Sub

'TODO: XML get/set functions for getting/setting all parameters at once
Public Function getAllFontSettingsAsXML(Optional ByVal ignoreFluffEntries As Boolean = False) As String

    'pdXML handles string assembly for us
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    
    xmlEngine.prepareNewXML "pdTextSettings"
    
    If Not ignoreFluffEntries Then
        xmlEngine.writeBlankLine
        xmlEngine.writeComment "This XML string contains initialization settings for the pdTextRenderer class."
        xmlEngine.writeBlankLine
    End If
    
    'Properties are added in turn.
    xmlEngine.writeTag "FontFace", m_FontFace
    xmlEngine.writeTag "FontSize", m_FontSize
    xmlEngine.writeTag "FontSizeUnit", m_FontSizeUnit
    xmlEngine.writeTag "FontBold", m_FontBold
    xmlEngine.writeTag "FontItalic", m_FontItalic
    xmlEngine.writeTag "FontUnderline", m_FontUnderline
    xmlEngine.writeTag "FontStrikeout", m_FontStrikeout
    xmlEngine.writeTag "FontHorizontalAlignment", m_HorizontalAlignment
    xmlEngine.writeTag "FontVerticalAlignment", m_VerticalAlignment
    xmlEngine.writeTag "FontRenderingHint", m_TextRenderingHint
    xmlEngine.writeTag "FontContrast", m_TextContrast
    
    'Return the completed XML string
    getAllFontSettingsAsXML = xmlEngine.returnCurrentXMLString

End Function

Public Function setAllFontSettingsFromXML(ByRef srcXMLString As String) As Boolean
    
    'Note that this function does not manually mark the font cache as dirty.  Individual setXYZ functions will do that as necessary.
    
    'Prep an XML parser
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    
    'Load the source XML string
    If xmlEngine.loadXMLFromString(srcXMLString) Then
    
        'Validate the XML subtype
        If xmlEngine.isPDDataType("pdTextSettings") Then
        
            'Retrieve all relevant settings in turn.  If a setting is not supported (due to an older version of PD writing this
            ' XML string), supply a default value.
            
            'TODO: raise some kind of error or message if the listed font is not available.  This is a possibility for PDI images
            ' traded between users.
            setFontFace xmlEngine.getUniqueTag_String("FontFace", g_InterfaceFont)
            setFontSize xmlEngine.getUniqueTag_Double("FontSize", 16)
            setFontSizeUnit xmlEngine.getUniqueTag_Long("FontSizeUnit", pdfu_Pixel)
            setFontBold xmlEngine.getUniqueTag_Boolean("FontBold", False)
            setFontItalic xmlEngine.getUniqueTag_Boolean("FontItalic", False)
            setFontUnderline xmlEngine.getUniqueTag_Boolean("FontUnderline", False)
            setFontStrikeout xmlEngine.getUniqueTag_Boolean("FontStrikeout", False)
            setHorizontalAlignment xmlEngine.getUniqueTag_Long("FontHorizontalAlignment", StringAlignmentNear)
            setVerticalAlignment xmlEngine.getUniqueTag_Long("FontVerticalAlignment", StringAlignmentNear)
            setTextRenderingHint xmlEngine.getUniqueTag_Long("FontRenderingHint", TextRenderingHintAntiAliasGridFit)
            setTextContrast xmlEngine.getUniqueTag_Long("FontContrast", 4)
            
            'Return success!
            setAllFontSettingsFromXML = True
        
        Else
            setAllFontSettingsFromXML = False
        End If
        
    Else
        setAllFontSettingsFromXML = False
    End If
    
End Function

'Get/Set font face.  These are more cumbersome than other font settings.
Public Function getFontFace() As String
    getFontFace = m_FontFace
End Function

'Set a given font face, as specified by a font name.
Public Sub setFontFace(ByVal fontName As String)
    
    'If this font face has already been set, ignore this request
    If (StrComp(fontName, m_FontFace, vbBinaryCompare) <> 0) Then
        m_FontFace = fontName
        m_FontCacheClean = False
    End If
    
End Sub

'Convert a pdFontUnit enum to a GDI+ format measurement unit.  Note that not all GDI+ units are reachable this way, by design.
Private Function convertPDFontUnitToGDIPlusFontUnit(ByRef srcPDFontUnit As pdFontUnit) As GdiPlusMeasurementUnit
    
    Select Case srcPDFontUnit
    
        Case pdfu_Pixel
            convertPDFontUnitToGDIPlusFontUnit = gpUnitPixel
        
        Case pdfu_Point
            convertPDFontUnitToGDIPlusFontUnit = gpUnitPoint
            
    End Select
    
End Function

'When all font parameters have been set, we can actually create a font!
Public Function createFontObject() As Boolean
    
    'If the font cache is clean, this request is redundant; ignore it
    If m_FontCacheClean Then
        createFontObject = True
        Exit Function
    End If
    
    'Reset the current text engine marker, as we won't know which engine is required until we try to create the font.
    m_curTextEngine = te_UNKNOWN
    
    'If we have already created any font objects (in either engine), free them now
    If m_GDIPlusFont <> 0 Then GdipDeleteFont m_GDIPlusFont
    If m_GDIPlusFontFamily <> 0 Then GdipDeleteFontFamily m_GDIPlusFontFamily
    
    'First, try to create a GDI+ font family.  This saves us a lot of grief.
    Dim fontCreationSuccess As Boolean
    
    'As of 28 April 2015, I am deactivating the GDI+ codepath.  PD's manual glyph rendering code is stable enough for me to
    ' rely on it exclusively for the remainder of this testing cycle, I think.  (That said, I'm leaving this here in case
    ' I ever need to revisit in the future.)
    'fontCreationSuccess = createGDIPlusFont()
    
    'GDI+ reported success, and we obtained a valid GDI+ font handle, carry on
    If fontCreationSuccess And (m_GDIPlusFont <> 0) Then
        
        'Note that GDI+ works fine for this particular combination of settings, and return success
        m_curTextEngine = te_GDIPLUS
        m_FontCacheClean = True
        createFontObject = True
        
    'If GDI+ failed, fall back to GDI and try again.
    Else
        
        fontCreationSuccess = createGDIFont()
        
        'If GDI is successful...
        If fontCreationSuccess And (m_GDIFont <> 0) Then
        
            'Note that GDI is required for this combination of settings
            m_curTextEngine = te_GDI
            m_FontCacheClean = True
            createFontObject = True
            
        Else
        
            'For some reason, this combination of settings doesn't work at all.  I hope this never happens, but you never know.
            #If DEBUGMODE = 1 Then
                pdDebug.LogAction "WARNING!  pdTextRenderer's createFontObject failed with both GDI+ and GDI; sorry!"
            #End If
            
            m_curTextEngine = te_UNKNOWN
            m_FontCacheClean = False
            createFontObject = False
            
        End If
        
    End If
        
End Function

'Attempt to create a GDI+ font matching the current class settings.  Note that GDI+ creates two relevant objects: font family, and font itself
Private Function createGDIPlusFont() As Boolean
    
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
        gdipReturn = GdipCreateFont(m_GDIPlusFontFamily, m_FontSize, getFontStylesAsGdipLong, convertPDFontUnitToGDIPlusFontUnit(m_FontSizeUnit), m_GDIPlusFont)
        
        'Check for known errors; GDI+ cannot synthesize some font sizes, for example
        If (gdipReturn = 0) And (m_GDIPlusFont <> 0) Then
        
            'Font creation successful.  Set the font cache flag, so we know not to create the font again.
            Debug.Print "WOW, two miracles in a row!  GDI+ font object created successfully."
            
            'Return success
            createGDIPlusFont = True
            
        Else
        
            'Font creation failure.  Not much we can do now.
            Debug.Print "WARNING!  GDI+ failed to create a font object based on this font family."
            reportGDIPlusFailure gdipReturn
            
            'Delete the font family we created as well, then exit
            If m_GDIPlusFontFamily <> 0 Then GdipDeleteFontFamily m_GDIPlusFontFamily
            createGDIPlusFont = False
            
        End If
        
    'GDI+ font selection can fail for a variety of reasons; report the failure reason and exit immediately
    Else
        
        Debug.Print "GDI+ font family could not be created; font is probably OpenType."
        reportGDIPlusFailure gdipReturn
        createGDIPlusFont = False
        
    End If
    
End Function

'Convert the current arrangement of font styles into a Long-type value compatible with GDI+'s style declarations
Private Function getFontStylesAsGdipLong() As Long
    
    getFontStylesAsGdipLong = 0
    
    If m_FontBold Then getFontStylesAsGdipLong = (getFontStylesAsGdipLong Or FontStyleBold)
    If m_FontItalic Then getFontStylesAsGdipLong = (getFontStylesAsGdipLong Or FontStyleItalic)
    If m_FontUnderline Then getFontStylesAsGdipLong = (getFontStylesAsGdipLong Or FontStyleUnderline)
    If m_FontStrikeout Then getFontStylesAsGdipLong = (getFontStylesAsGdipLong Or FontStyleStrikeout)
    
End Function

'If a GDI+ function fails, pass the return value to this function so the debugger can attempt to output a helpful message.
Private Sub reportGDIPlusFailure(ByVal failCode As Long)

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
Private Function createGDIFont() As Boolean
    
    'Reset the current font cache state, as we're about to wipe any existing font handles
    m_FontCacheClean = False
    
    Debug.Print "Attempting to create GDI font object matching the current pdTextRenderer settings..."
    
    'If a GDI font handle already exists, free it.  Note that the font must have been unselected from a DC; this is typically handled
    ' by the renderer function itself.
    If m_GDIFont <> 0 Then DeleteObject m_GDIFont
    
    'The font management module makes this task a lot easier.  Start by creating a LOGFONTW container.
    Dim tmpLogFont As LOGFONTW
    
    'Use the font management module to populate the struct
    Font_Management.fillLogFontW_Basic tmpLogFont, m_FontFace, m_FontBold, m_FontItalic, m_FontUnderline, m_FontStrikeout
    Font_Management.fillLogFontW_Size tmpLogFont, m_FontSize, m_FontSizeUnit
    Font_Management.fillLogFontW_Quality tmpLogFont, m_TextRenderingHint
    
    'Attempt to create a matching GDI font
    If Font_Management.createGDIFont(tmpLogFont, m_GDIFont) Then
        
        'Success!  Note that GDI is an acceptable engine for this font, and return success
        Debug.Print "GDI font fallback worked!  Proceeding with rendering."
        m_curTextEngine = te_GDI
        m_FontCacheClean = True
        createGDIFont = True
        
    Else
        
        'Failure!  PD doesn't have a fallback beyond GDI, so this whole request is doomed.
        ' (This shouldn't technically be possible, but if it is, I'd rather laugh than cry.)
        Debug.Print "GDI font fallback failed miserably!  Everything is falling apart!  AAAAHHHH!"
        m_curTextEngine = te_UNKNOWN
        m_FontCacheClean = False
        createGDIFont = False
    
    End If
    
End Function

'Use this function to render arbitrary text to an arbitrary DIB.  Performance should be decent, but if you're going to be making
' a ton of successive calls, it's probably worthwhile to cache a GDI+ graphics instance that points to the target.  (This class does
' not currently offer a function for that, FYI.)
'
'Note also that the crColor parameter should not include alpha; alpha in PD is handled by layers, not the text renderer.
Public Function renderTextToDIB(ByRef dstDIB As pdDIB, ByRef srcString As String, ByVal crColor As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal textWidth As Single, ByVal textHeight As Single) As Boolean
    
    'As a convenience to the user, create the font as necessary
    If (Not m_FontCacheClean) Then createFontObject
    
    'We should now know whether GDI or GDI+ is required for this combination of settings.
    Select Case m_curTextEngine
        
        'GDI+ is sufficient
        Case te_GDIPLUS
            renderTextToDIB = renderTextToDIB_GDIPlus(dstDIB, srcString, crColor, x1, y1, textWidth, textHeight)
        
        'GDI is required
        Case te_GDI
            
            'Because GDI doesn't support 32-bpp rendering targets, we must manually translate GDI glyphs into GDI+ paths, then render
            ' the paths manually.  Because this is hugely energy intensive, our separate class attempts to cache glyphs as they are
            ' created.  This saves us from having to translate glyphs more than once, which is a non-trivial boost on long chunks
            ' of text.
            
            'Start by creating a pdGlyphCollection instance, if one doesn't already exist.
            If (m_GlyphCollection Is Nothing) Then Set m_GlyphCollection = New pdGlyphCollection
            
            'Notify the glyph collection of our GDI font handle.  It obviously needs this in order to generate glyphs.
            m_GlyphCollection.notifyOfGDIFontChoice m_GDIFont
            
            'TODO: retrieve kerning pairs in advance...?
            
            'Next, pass it our string.  It will generate a list of required glyphs and automatically retrieve any glyphs it hasn't
            ' already calculated.  (For example, if the last request it received was "cat", but this request is "cats", only the 's'
            ' needs to be newly retrieved.)
            m_GlyphCollection.buildGlyphCollection srcString, True
            
            'After all glyphs have been created, it's *very important* that we regain control over our GDI font
            m_GlyphCollection.requestGDIFontRelease
            
            'With the glyph collection successfully assembled, proceed with rendering
            renderTextToDIB = renderTextToDIB_Glyphs(dstDIB, srcString, crColor, x1, y1, textWidth, textHeight)
            
            'While I work on the above code, here's a dummy rendering mode for testing
            'renderTextToDIB = renderTextToDIB_GDI_Dummy(dstDIB, srcString, crColor, x1, y1, textWidth, textHeight)
        
        
        'GDI and GDI+ both failed; at present, we have no fallback for this case
        Case te_UNKNOWN
            Debug.Print "GDI and GDI+ both failed; something is probably wrong with this font.  Consider blacklisting it...?"
            
    End Select
    
End Function

'Render text using GDI+
Private Function renderTextToDIB_GDIPlus(ByRef dstDIB As pdDIB, ByRef srcString As String, ByVal crColor As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal textWidth As Single, ByVal textHeight As Single) As Boolean

    'TESTING ONLY!  Fill with white prior to rendering
    'GDI_Plus.GDIPlusFillDIBRect dstDIB, 0, 0, dstDIB.getDIBWidth, dstDIB.getDIBHeight, vbWhite
    
    'Start by acquiring a GDI+ handle to the destination DIB
    Dim dstGraphics As Long
    dstGraphics = GDI_Plus.getGDIPlusGraphicsFromDC(dstDIB.getDIBDC, False)
    
    'Next, set up all container-specific text settings
    If GdipSetTextRenderingHint(dstGraphics, m_TextRenderingHint) <> 0 Then
        Debug.Print "Failed to set text rendering hint " & m_TextRenderingHint
    End If
    
    If GdipSetTextContrast(dstGraphics, m_TextContrast) <> 0 Then
        Debug.Print "Failed to set text contrast " & m_TextContrast
    End If
    
    If GdipSetCompositingQuality(dstGraphics, CompositingQualityAssumeLinear) <> 0 Then
        Debug.Print "Failed to set compositing quality."
    End If
    
    If GdipSetPixelOffsetMode(dstGraphics, PixelOffsetModeHalf) <> 0 Then
        Debug.Print "Failed to set pixel offset mode."
    End If
    
    'Create a solid fill brush.  (In the future, we'll add more impressive options!)
    Dim gdipBrush As Long
    gdipBrush = GDI_Plus.getGDIPlusBrushHandle(crColor, 255)
    
    'Convert the input rect to a RECTF
    Dim boundingRect As RECTF
    With boundingRect
        .Left = x1
        .Top = y1
        .Width = textWidth
        .Height = textHeight
    End With
    
    'Finally, use GDI+ to render the actual string
    Dim gdipReturn As Long
    gdipReturn = GdipDrawString(dstGraphics, StrPtr(srcString), Len(srcString), m_GDIPlusFont, boundingRect, m_GDIPlusStringFormat, gdipBrush)
    
    If gdipReturn = 0 Then
        renderTextToDIB_GDIPlus = True
    Else
        Debug.Print "WARNING!  GdipDrawString failed with error code " & gdipReturn
        Debug.Print "WARNING!  (Extra debug info: hFont - " & m_GDIPlusFont & ", hStringFormat - " & m_GDIPlusStringFormat & ")"
        renderTextToDIB_GDIPlus = False
    End If
    
    'Release the temporary GDI+ objects we created
    GDI_Plus.releaseGDIPlusBrush gdipBrush
    GDI_Plus.releaseGDIPlusGraphics dstGraphics

End Function

'Render text using GDI+, and a glyph collection properly assembled inside m_GlyphCollection.
Private Function renderTextToDIB_Glyphs(ByRef dstDIB As pdDIB, ByRef srcString As String, ByVal crColor As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal textWidth As Single, ByVal textHeight As Single) As Boolean

    'TESTING ONLY!  Fill with white prior to rendering
    'GDI_Plus.GDIPlusFillDIBRect dstDIB, 0, 0, dstDIB.getDIBWidth, dstDIB.getDIBHeight, vbWhite
    
    'Start by acquiring a GDI+ handle to the destination DIB
    Dim dstGraphics As Long
    dstGraphics = GDI_Plus.getGDIPlusGraphicsFromDC(dstDIB.getDIBDC, False)
    
    'Next, set up all container-specific path settings
    If GdipSetCompositingQuality(dstGraphics, CompositingQualityAssumeLinear) <> 0 Then
        Debug.Print "Failed to set compositing quality."
    End If
    
    If GdipSetPixelOffsetMode(dstGraphics, PixelOffsetModeHalf) <> 0 Then
        Debug.Print "Failed to set pixel offset mode."
    End If
    
    If GdipSetSmoothingMode(dstGraphics, SmoothingModeAntiAlias) <> 0 Then
        Debug.Print "Failed to set smoothing mode."
    End If
    
    'Create a solid fill brush.  (In the future, we'll add more impressive options!)
    Dim gdipBrush As Long
    gdipBrush = GDI_Plus.getGDIPlusBrushHandle(crColor, 255)
    
    'Convert the input rect to a RECTF
    Dim boundingRect As RECTF
    With boundingRect
        .Left = x1
        .Top = y1
        .Width = textWidth
        .Height = textHeight
    End With
    
    'Retrieve a full, composite path from the glyph collector.  It handles messy business like character positioning,
    ' and simply hands us a fully composited path, ready for rendering.
    Dim finalTextPath As pdGraphicsPath
    If m_GlyphCollection.assembleCompositePath(srcString, finalTextPath, boundingRect) Then
        
        'Draw the path
        Dim gdipReturn As Long
        gdipReturn = GdipFillPath(dstGraphics, gdipBrush, finalTextPath.getCopyOfPathHandle)
        
        If gdipReturn <> 0 Then Debug.Print "WARNING!  GdipFillPath failed; return code was " & gdipReturn
        
    Else
        Debug.Print "WARNING!  m_GlyphCollection.assembleCompositePath returned FALSE.  Please investigate."
    End If
    
    'TODO: handle returns
    'Dim gdipReturn As Long
   '
   ' If gdipReturn = 0 Then
   '     renderTextToDIB_GDIPlus = True
   ' Else
   '     Debug.Print "WARNING!  GdipDrawString failed with error code " & gdipReturn
   '     Debug.Print "WARNING!  (Extra debug info: hFont - " & m_GDIPlusFont & ", hStringFormat - " & m_GDIPlusStringFormat & ")"
   '     renderTextToDIB_GDIPlus = False
   ' End If
    
    'Release the temporary GDI+ objects we created
    GDI_Plus.releaseGDIPlusBrush gdipBrush
    GDI_Plus.releaseGDIPlusGraphics dstGraphics

End Function

'Render text using GDI using a dummy implementation, only.  This only exists to test things like GDI font creation; I am actively working on
' superceding it with a much more advanced implementation.
Private Function renderTextToDIB_GDI_Dummy(ByRef dstDIB As pdDIB, ByRef srcString As String, ByVal crColor As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal textWidth As Single, ByVal textHeight As Single) As Boolean

    'TESTING ONLY!  Since GDI font rendering forcibly sets alpha to 0 on any pixels it touches, we go ahead and fill
    ' the layer with white prior to rendering.
    GDI_Plus.GDIPlusFillDIBRect dstDIB, x1, y1, textWidth, textHeight, vbWhite
    
    'Start by setting up all container-specific text settings
    Dim oldFont As Long
    oldFont = SelectObject(dstDIB.getDIBDC, m_GDIFont)
    
    'Set other font parameters
    SetTextColor dstDIB.getDIBDC, crColor
    SetTextAlign dstDIB.getDIBDC, 0        'Setting actual text alignment is TODO!
    
    'Enable transparent font rendering; if nothing else, this lets us see a "cutout" where GDI wants to render the font
    SetBkMode dstDIB.getDIBDC, FONT_TRANSPARENT
    
    'You'd think drawing opaque text would work, but nope, this just means that GDI functions make the ENTIRE TEXT AREA transparent.
    ' THANKS, MICROSOFT!
    'SetBkMode dstDIB.getDIBDC, FONT_OPAQUE
    'SetBkColor dstDIB.getDIBDC, vbWhite
    
    'GDI rendering requires an integer rect; populate one now
    Dim tmpRect As RECT
    tmpRect.Left = x1
    tmpRect.Top = y1
    tmpRect.Right = x1 + textWidth
    tmpRect.Bottom = y1 + textHeight
    
    'Use DrawText for the actual rendering
    Dim retDrawText As Long
    retDrawText = DrawText(dstDIB.getDIBDC, StrPtr(srcString), Len(srcString), tmpRect, DT_WORDBREAK Or DT_LEFT Or DT_TOP)
    
    'Restore prior font settings (except text color and alignment; we don't care about those because PD will override those automatically)
    SelectObject dstDIB.getDIBDC, oldFont
    
    'This is just debug code, so I don't care about return values
    Debug.Print "FYI, renderTextToDIB_GDI_Dummy reported its outcome as " & UCase(CStr(CBool(retDrawText <> 0)))
    renderTextToDIB_GDI_Dummy = True

End Function

Private Sub Class_Initialize()
    
    'Only one backend is available at present
    m_CurrentFontBackend = frb_WAPI
    
    'Create a relevant font rendering backend
    If m_CurrentFontBackend = frb_WAPI Then
        
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
        GdipStringFormatGetGenericTypographic m_GDIPlusStringFormat
        
    End If
    
    'Mark the current font cache as dirty
    m_FontCacheClean = False
    
    'Reset the text engine marker (we don't know this until we attempt to create a given font)
    m_curTextEngine = te_UNKNOWN
        
    'By default, this class uses pixels
    m_FontSizeUnit = pdfu_Pixel
    
    'Set some default font properties
    m_FontFace = g_InterfaceFont
    m_FontSize = 16#                '16px = 12 points at 96 DPI
    
    m_HorizontalAlignment = StringAlignmentNear
    m_VerticalAlignment = StringAlignmentNear
    
    'Set default graphics container text settings
    m_TextContrast = 4
    m_TextRenderingHint = TextRenderingHintAntiAliasGridFit

End Sub

Private Sub Class_Terminate()
    
    'Perform any backend-specific clean-up
    If m_CurrentFontBackend = frb_WAPI Then
    
        'Release any GDI+ objects we created
        If m_GDIPlusFont <> 0 Then GdipDeleteFont m_GDIPlusFont
        If m_GDIPlusFontFamily <> 0 Then GdipDeleteFontFamily m_GDIPlusFontFamily
        If m_GDIPlusStringFormat <> 0 Then GdipDeleteStringFormat m_GDIPlusStringFormat
        
        'Release any GDI objects we created
        If m_GDIFont <> 0 Then DeleteObject m_GDIFont
        
    End If
    
End Sub
