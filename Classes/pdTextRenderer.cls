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

'GDI+ supports multiple measurement modes, and these may actually be relevant for fonts
Public Enum GdiPlusMeasurementUnit
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

'Font family functions
Private Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal ptrToSrcFontName As Long, ByVal srcFontCollection As Long, ByRef dstFontFamily As Long) As Long
Private Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal srcFontFamily As Long) As Long
Private Declare Function GdipIsStyleAvailable Lib "gdiplus" (ByVal srcFontFamily As Long, ByVal srcStyleToTest As FontStyle, ByRef dstIsStyleAvailable As Long) As Long

'Font functions
Private Declare Function GdipCreateFont Lib "gdiplus" (ByVal srcFontFamily As Long, ByVal srcFontSize As Single, ByVal srcFontStyle As FontStyle, ByVal srcMeasurementUnit As GdiPlusMeasurementUnit, ByRef dstCreatedFont As Long) As Long
Private Declare Function GdipDeleteFont Lib "gdiplus" (ByVal srcFont As Long) As Long

'This class internally maintains various font properties.  These properties must be converted into specific GDI+ parameters via
' various means, but they are stored in human-friendly format to simplify serializing a class instance to an XML string.
Private m_FontFace As String
Private m_FontSize As Single
Private m_FontSizeUnit As GdiPlusMeasurementUnit
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

'For performance reasons, this class caches various GDI+ font objects and handles.  This spares us from having to recreate
' expensive font data during rendering steps.
Private m_GDIPlusFont As Long
Private m_GDIPlusFontFamily As Long

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

Friend Function getFontSizeUnit() As GdiPlusMeasurementUnit
    getFontSizeUnit = m_FontSizeUnit
End Function

Friend Sub setFontSizeUnit(ByVal newUnit As GdiPlusMeasurementUnit)
    If newUnit <> m_FontSizeUnit Then
        m_FontSizeUnit = newUnit
        m_FontCacheClean = False
    End If
End Sub

'TODO: XML get/set functions for getting/setting all parameters at once
Public Function getAllFontSettingsAsXML() As String

End Function

Public Sub setAllFontSettingsFromXML(ByRef srcXMLString As String)
    m_FontCacheClean = False
End Sub

'Get/Set font face.  These are more cumbersome than other font settings.
Public Function getFontFace() As String
    getFontFace = m_FontFace
End Function

'Set a given font face, as specified by a font name (hopefully captured via getListOfInstalledFonts, below)
' Returns: TRUE if font exists; FALSE otherwise
Public Function setFontFace(ByVal fontName As String) As Boolean
    
    'If a font family already exists, delete it
    If m_GDIPlusFontFamily <> 0 Then GdipDeleteFontFamily m_GDIPlusFontFamily
    
    'Attempt to retrieve the font family matching this name
    Dim gdipReturn As Long
    gdipReturn = GdipCreateFontFamilyFromName(StrPtr(fontName), 0&, m_GDIPlusFontFamily)
    
    'This GDI+ function supports several unique fail conditions; test for them
    If (gdipReturn = 0) And (m_GDIPlusFontFamily <> 0) Then
        
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
    
    'If a font already exists, delete it now
    If m_GDIPlusFont <> 0 Then GdipDeleteFont m_GDIPlusFont
    
    'Make sure a font face was specified, as it's the most crucial input parameter
    If Len(Trim$(m_FontFace)) = 0 Then
        Debug.Print "WARNING! No font name specified!"
        createFontObject = False
    End If
    
    'Attempt to create the font
    Dim gdipReturn As Long
    'gdipreturn = gdipcreatefont(m_gdipfontfamily,m_fontsize,
    
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
    m_FontSizeUnit = gpUnitPixel
    
    'Set some default font properties
    m_FontFace = g_InterfaceFont
    m_FontSize = 16#

End Sub

Private Sub Class_Terminate()
    
    'Release any GDI+ objects we created
    If m_GDIPlusFont <> 0 Then GdipDeleteFont m_GDIPlusFont
    If m_GDIPlusFontFamily <> 0 Then GdipDeleteFontFamily m_GDIPlusFontFamily
    
End Sub
