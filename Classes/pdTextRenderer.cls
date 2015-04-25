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
'Last updated: 22/April/15
'Last update: initial build
'
'While PhotoDemon already has a capable GDI font renderer in pdFont, text rendering onto actual images is a whole other
' can of worms.
'
'The biggest problem with GDI rendering is its total lack of support for 32-bpp rendering targets.  This is a deal-breaker
' for text layers, so I have no choice but to lean on GDI+ (short of writing some incredibly ugly functions for rasterizing
' GDI font render results, and extrapolating font data - ugh no).
'
'As such, this text renderer class was originally designed against GDI+ text capabilities.  As a failsafe against future
' improvements, I have included some framework code for variable backends.  Pango intrigues me as another rendering option,
' so I may look to it in the future if I can't wrangle GDI+ into supporting all the features I'd like.  (As it is, I'll
' be manually implementing some features, like line and character spacing, but that's okay.)
'
'Much like pdFont, this class requires you to *manually create a font object before using it*.  Font creation is fairly
' cumbersome in GDI+, and this class will take care of some initializations in the background.  But because things like
' font size must be declared *at font creation time*, it is very likely that a single session will require many font
' creation/destruction cycles.  To that end, after any/all font settings are adjusted, you MUST CREATE THE FONT prior to
' actually rendering with this class.
'
'Dependencies:
' - pdStringStack (used to report supported fonts as detected by GDI+)
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

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

'Font collection interfaces
Private Declare Function GdipNewInstalledFontCollection Lib "gdiplus" (ByRef dstFontCollectionHandle As Long) As Long
Private Declare Function GdipGetFontCollectionFamilyCount Lib "gdiplus" (ByVal srcFontCollection As Long, ByRef dstNumFound As Long) As Long
Private Declare Function GdipGetFontCollectionFamilyList Lib "gdiplus" (ByVal srcFontCollection As Long, ByVal sizeOfDstBuffer As Long, ByVal ptrToDstFontFamilyArray As Long, ByRef dstNumFound As Long) As Long
Private Declare Function GdipGetFamilyName Lib "gdiplus" (ByVal srcFontFamily As Long, ByVal ptrDstNameBuffer As Long, ByVal languageID As Integer) As Long
Private Const LF_FACESIZE As Long = 32          'Note: this represents 32 *chars*, not bytes!
Private Const LANG_NEUTRAL As Integer = &H0

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

'Font family functions
Private Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal ptrToSrcFontName As Long, ByVal srcFontCollection As Long, ByRef dstFontFamily As Long) As Long
Private Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal srcFontFamily As Long) As Long
Private Declare Function GdipIsStyleAvailable Lib "gdiplus" (ByVal srcFontFamily As Long, ByVal srcStyleToTest As FontStyle, ByRef dstIsStyleAvailable As Long) As Long

'Font functions
Private Declare Function GdipCreateFont Lib "gdiplus" (ByVal srcFontFamily As Long, ByVal srcFontSize As Single, ByVal srcFontStyle As FontStyle, ByVal srcMeasurementUnit As GdiPlusMeasurementUnit, ByRef dstCreatedFont As Long) As Long
Private Declare Function GdipDeleteFont Lib "gdiplus" (ByVal srcFont As Long) As Long

'String format functions
Private Declare Function GdipStringFormatGetGenericTypographic Lib "gdiplus" (ByRef dstStringFormat As Long) As Long
Private Declare Function GdipDeleteStringFormat Lib "gdiplus" (ByVal srcStringFormat As Long) As Long
Private Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal dstStringFormat As Long, ByVal newAlignment As GdiPlusStringAlignment) As Long
Private Declare Function GdipSetStringFormatLineAlign Lib "gdiplus" (ByVal dstStringFormat As Long, ByVal newLineAlignment As GdiPlusStringAlignment) As Long

'Graphics container font functions
Private Declare Function GdipSetTextRenderingHint Lib "gdiplus" (ByVal dstGraphics As Long, ByVal newRenderHintMode As GdiPlusTextRenderingHint) As Long
Private Declare Function GdipSetTextContrast Lib "gdiplus" (ByVal dstGraphics As Long, ByVal textContrast As Long) As Long
Private Declare Function GdipSetCompositingQuality Lib "gdiplus" (ByVal dstGraphics As Long, ByVal newCompositingQuality As GdiPlusCompositingQuality) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal dstGraphics As Long, ByVal newSmoothingMode As GdiPlusSmoothingMode) As Long
Private Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal dstGraphics As Long, ByVal newPixelOffsetMode As GdiPlusPixelOffsetMode) As Long

'Render functions
Private Declare Function GdipDrawString Lib "gdiplus" (ByVal dstGraphics As Long, ByVal ptrToString As Long, ByVal strLength As Long, ByVal gdipFontHandle As Long, ByRef layoutRect As RECTF, ByVal gdipStringFormat As Long, ByVal gdipBrush As Long) As Long

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

'For performance reasons, this class caches various GDI+ font objects and handles.  This spares us from having to recreate
' expensive font data during rendering steps.
Private m_GDIPlusFont As Long
Private m_GDIPlusFontFamily As Long
Private m_GDIPlusStringFormat As Long

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
        GdipSetStringFormatAlign m_GDIPlusStringFormat, newAlignment
    End If
    
End Sub

Friend Function getVerticalAlignment() As GdiPlusStringAlignment
    getVerticalAlignment = m_VerticalAlignment
End Function

Friend Sub setVerticalAlignment(ByVal newAlignment As GdiPlusStringAlignment)
    
    If m_VerticalAlignment <> newAlignment Then
        m_VerticalAlignment = newAlignment
        GdipSetStringFormatLineAlign m_GDIPlusStringFormat, newAlignment
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

'Set a given font face, as specified by a font name (hopefully captured via getListOfInstalledFonts, below)
' Returns: TRUE if font exists; FALSE otherwise
Public Function setFontFace(ByVal fontName As String) As Boolean
    
    'If this font face has already been set, and a font has been successfully created, ignore this request
    If (StrComp(fontName, m_FontFace, vbBinaryCompare) = 0) And (m_GDIPlusFontFamily <> 0) Then
        setFontFace = True
        Exit Function
    End If
    
    'If a font family already exists, delete it
    If m_GDIPlusFontFamily <> 0 Then GdipDeleteFontFamily m_GDIPlusFontFamily
    
    'Attempt to retrieve the font family matching this name
    Dim gdipReturn As Long
    gdipReturn = GdipCreateFontFamilyFromName(StrPtr(fontName), 0&, m_GDIPlusFontFamily)
    
    'This GDI+ function supports several unique fail conditions; test for them
    If (gdipReturn = 0) And (m_GDIPlusFontFamily <> 0) Then
        
        'The font appears to have been created successfully!
        m_FontFace = fontName
        
        'Prior to exiting, cache the font styles supported by this font family
        Dim testResult As Long
        GdipIsStyleAvailable m_GDIPlusFontFamily, FontStyleBold, testResult
        m_BoldSupported = CBool(testResult <> 0)
        
        GdipIsStyleAvailable m_GDIPlusFontFamily, FontStyleItalic, testResult
        m_ItalicSupported = CBool(testResult <> 0)
        
        GdipIsStyleAvailable m_GDIPlusFontFamily, FontStyleUnderline, testResult
        m_UnderlineSupported = CBool(testResult <> 0)
        
        GdipIsStyleAvailable m_GDIPlusFontFamily, FontStyleStrikeout, testResult
        m_StrikeoutSupported = CBool(testResult <> 0)
        
        'Note that the current font cache, if one exists, is no longer valid
        m_FontCacheClean = False
        
        'Return success
        setFontFace = True
        
    Else
        
        Const FontFamilyNotFound As Long = 14
        Const FontStyleNotFound As Long = 15
        Const NotTrueTypeFont As Long = 16
        
        #If DEBUG_MODE = 1 Then
        
            Select Case gdipReturn
            
                Case FontFamilyNotFound
                    pdDebug.LogAction "pdTextRenderer.setFontFace failed: font family wasn't found"
                    
                Case FontStyleNotFound
                    pdDebug.LogAction "pdTextRenderer.setFontFace failed: requested font style wasn't found"
                    
                Case NotTrueTypeFont
                    pdDebug.LogAction "pdTextRenderer.setFontFace failed: requested font is not a TrueType font"
                    
                Case Else
                    pdDebug.LogAction "pdTextRenderer.setFontFace failed for unknown reasons"
            
            End Select
            
        #End If
        
        setFontFace = False
        
    End If
    
End Function

'When all font parameters have been set, we can actually create a font!
Public Function createFontObject() As Boolean
    
    'If the font cache is clean, this request is redundant; ignore it
    If m_FontCacheClean Then
        createFontObject = True
        Exit Function
    End If
    
    Debug.Print "Creating GDI+ font object..."
    
    'If a font already exists, delete it now
    If m_GDIPlusFont <> 0 Then GdipDeleteFont m_GDIPlusFont
    
    'If a font family does not exist, attempt to create one now
    If m_GDIPlusFontFamily = 0 Then setFontFace m_FontFace
    
    'If a font family *still* does not exist, the font name is invalid.
    If m_GDIPlusFontFamily = 0 Then
        Debug.Print "WARNING! Invalid font name specified!"
        createFontObject = False
    End If
    
    'Attempt to create the font
    Dim gdipReturn As Long
    gdipReturn = GdipCreateFont(m_GDIPlusFontFamily, m_FontSize, getFontStylesAsGdipLong(), convertPDFontUnitToGDIPlusFontUnit(m_FontSizeUnit), m_GDIPlusFont)
    
    'Check for known errors; GDI+ does not support OpenType fonts, for example
    If (gdipReturn = 0) And (m_GDIPlusFont <> 0) Then
    
        'Font creation successful.  Set the font cache flag, so we know not to create the font again.
        m_FontCacheClean = True
        createFontObject = True
        
    Else
    
        'Font creation failure.  Not much we can do now.
        m_FontCacheClean = False
        createFontObject = False
        
    End If
    
End Function

'Convert a pdFontUnit enum to a GDI+ format measurement unit.  Note that not all GDI+ units are reachable this way, by design.
Private Function convertPDFontUnitToGDIPlusFontUnit(ByRef srcPDFontUnit As pdFontUnit) As GdiPlusMeasurementUnit
    
    Select Case srcPDFontUnit
    
        Case pdfu_Pixel
            convertPDFontUnitToGDIPlusFontUnit = gpUnitPixel
        
        Case pdfu_Point
            convertPDFontUnitToGDIPlusFontUnit = gpUnitPoint
            
    End Select
    
End Function

'Convert the current arrangement of font styles into a Long-type value compatible with GDI+'s style declarations
Private Function getFontStylesAsGdipLong() As Long
    
    getFontStylesAsGdipLong = 0
    
    If m_FontBold Then getFontStylesAsGdipLong = (getFontStylesAsGdipLong Or FontStyleBold)
    If m_FontItalic Then getFontStylesAsGdipLong = (getFontStylesAsGdipLong Or FontStyleItalic)
    If m_FontUnderline Then getFontStylesAsGdipLong = (getFontStylesAsGdipLong Or FontStyleUnderline)
    If m_FontStrikeout Then getFontStylesAsGdipLong = (getFontStylesAsGdipLong Or FontStyleStrikeout)
    
End Function

'Use this function to render arbitrary text to an arbitrary DIB.  Performance should be decent, but if you're going to be making
' a ton of successive calls, it's probably worthwhile to cache a GDI+ graphics instance that points to the target.  (This class does
' not currently offer a function for that, FYI.)
'
'Note also that the crColor parameter should not include alpha; alpha in PD is handled by layers, not the text renderer.
Public Function renderTextToDIB(ByRef dstDIB As pdDIB, ByRef srcString As String, ByVal crColor As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal textWidth As Single, ByVal textHeight As Single) As Boolean
    
    'As a convenience to the user, create the font as necessary
    If (Not m_FontCacheClean) Then createFontObject
    
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
    
    If gdipReturn <> 0 Then
        Debug.Print "WARNING!  GdipDrawString failed with error code " & gdipReturn
        Debug.Print "WARNING!  (Extra debug info: hFont - " & m_GDIPlusFont & ", hStringFormat - " & m_GDIPlusStringFormat & ")"
    End If
    
    'Release the temporary GDI+ objects we created
    GDI_Plus.releaseGDIPlusBrush gdipBrush
    GDI_Plus.releaseGDIPlusGraphics dstGraphics

End Function

'Helper function for returning a string stack of currently installed, GDI+ compatible (e.g. TrueType) fonts
Public Function getListOfInstalledFonts(ByRef dstStringStack As pdStringStack) As Boolean
    
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
                If dstStringStack Is Nothing Then Set dstStringStack = New pdStringStack
                dstStringStack.resetStack fontCount
                
                'Retrieve all font names
                Dim i As Long, thisFontName As String
                For i = 0 To fontsFound - 1
                    
                    'Retrieve the name for this entry
                    thisFontName = String$(LF_FACESIZE, 0)
                    If GdipGetFamilyName(fontList(i), StrPtr(thisFontName), LANG_NEUTRAL) = 0 Then
                        dstStringStack.AddString TrimNull(thisFontName)
                    End If
                    
                Next i
                
                'Return success
                getListOfInstalledFonts = True
            
            Else
                Debug.Print "WARNING! GDI+ refused to return a font collection list."
                getListOfInstalledFonts = False
            End If
        
        Else
            Debug.Print "WARNING! GDI+ refused to return a font collection count."
            getListOfInstalledFonts = False
        End If
    
    Else
        Debug.Print "WARNING! GDI+ refused to return a font collection object."
        getListOfInstalledFonts = False
    End If
    
End Function

'This function is identical to PD's publicly declared "TrimNull" function in File_And_Path_Handling.  It is included here to reduce
' external dependencies for this class.
Private Function TrimNull(ByVal origString As String) As String

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

Private Sub Class_Initialize()

    'Mark the current font cache as dirty
    m_FontCacheClean = False
    
    'By default, this class uses pixels
    m_FontSizeUnit = pdfu_Pixel
    
    'Set some default font properties
    m_FontFace = g_InterfaceFont
    m_FontSize = 16#                '16px = 12 points at 96 DPI
    
    'Because StringFormat creation is cheap, create a default typographic GDI+ StringFormat object now.
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
    m_HorizontalAlignment = StringAlignmentNear
    m_VerticalAlignment = StringAlignmentNear
    
    'Set default graphics container text settings
    m_TextContrast = 4
    m_TextRenderingHint = TextRenderingHintAntiAliasGridFit

End Sub

Private Sub Class_Terminate()
    
    'Release any GDI+ objects we created
    If m_GDIPlusFont <> 0 Then GdipDeleteFont m_GDIPlusFont
    If m_GDIPlusFontFamily <> 0 Then GdipDeleteFontFamily m_GDIPlusFontFamily
    If m_GDIPlusStringFormat <> 0 Then GdipDeleteStringFormat m_GDIPlusStringFormat
    
End Sub
