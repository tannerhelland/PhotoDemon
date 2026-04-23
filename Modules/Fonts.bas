Attribute VB_Name = "Fonts"
'***************************************************************************
'PhotoDemon Font Manager
'Copyright 2013-2026 by Tanner Helland
'Created: 31/May/13
'Last updated: 21/April/26
'Last update: add management for "recently used" fonts, and automatically relay these to font request functions
'
'For many years, PhotoDemon has used the pdFont class for GDI text rendering.  Unfortunately, that class
' was designed before I knew much about GDI font management, and it has some sketchy design considerations
' that make it a poor fit for PD's text tool.
'
'As part of a broader overhaul to PD's text management, this new Fonts module has been created.  Its job
' is to manage a font cache for this system, which individual font classes can then query for things like
' font existence, style, and more.
'
'Obviously, this class relies heavily on WAPI.  Functions are documented to the best of my knowledge and ability.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'****************************************************************************************
'Note: these types are used in the callback function for EnumFontFamiliesEx; as such, I have to declare them as public.

Private Const LF_FACESIZEW As Long = 64, LF_FACESIZEA As Long = 32
Private Const DEFAULT_CHARSET As Long = 1

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
    otmrcFontBox As RectL
    otmMacAscent As Long
    otmMacDescent As Long
    otmMacLineGap As Long
    otmusMinimumPPEM As Long
    otmptSubscriptSize As PointAPI
    otmptSubscriptOffset As PointAPI
    otmptSuperscriptSize As PointAPI
    otmptSuperscriptOffset As PointAPI
    otmsStrikeoutSize As Long
    otmsStrikeoutPosition As Long
    otmsUnderscoreSize As Long
    otmsUnderscorePosition As Long
    otmpFamilyName As Long
    otmpFaceName As Long
    otmpStyleName As Long
    otmpFullName As Long
End Type

Private Declare Function EnumFontFamiliesEx Lib "gdi32" Alias "EnumFontFamiliesExW" (ByVal hDC As Long, ByRef lpLogFontW As LOGFONTW, ByVal lpEnumFontFamExProc As Long, ByRef lParam As Any, ByVal dwFlags As Long) As Long

'GDI font weight (boldness)
Public Enum GDI_FontWeight
    fw_DontCare = 0
    fw_Thin = 100
    fw_Extralight = 200
    fw_Ultralight = 200
    fw_Light = 300
    fw_Normal = 400
    fw_Regular = 400
    fw_Medium = 500
    fw_SemiBold = 600
    fw_DemiBold = 600
    fw_Bold = 700
    fw_ExtraBold = 800
    fw_UltraBold = 800
    fw_Heavy = 900
    fw_Black = 900
End Enum

#If False Then
    Private Const fw_DontCare = 0, fw_Thin = 100, fw_Extralight = 200, fw_Ultralight = 200, fw_Light = 300, fw_Normal = 400, fw_Regular = 400, fw_Medium = 500, fw_SemiBold = 600, fw_DemiBold = 600, fw_Bold = 700, fw_ExtraBold = 800, fw_UltraBold = 800, fw_Heavy = 900, fw_Black = 900
#End If

Public Enum GDI_FontQuality
    fq_Default = 0
    fq_Draft = 1
    fq_Proof = 2
    fq_NonAntialiased = 3
    fq_Antialiased = 4
    fq_ClearType = 5
End Enum

#If False Then
    Private Const fq_Default = 0, fq_Draft = 1, fq_Proof = 2, fq_NonAntialiased = 3, fq_Antialiased = 4, fq_ClearType = 5
#End If

Public Enum GDI_TextAlign

    ta_NOUPDATECP = 0
    ta_UPDATECP = 1

    ta_LEFT = 0
    ta_RIGHT = 2
    ta_CENTER = 6

    ta_TOP = 0
    ta_BOTTOM = 8
    ta_BASELINE = 24
    ta_RTLREADING = 256
    ta_MASK = (ta_BASELINE + ta_CENTER + ta_UPDATECP + ta_RTLREADING)

    'Vertical text layouts can use these slightly altered enums for easier reading
    vta_BASELINE = ta_BASELINE
    vta_LEFT = ta_BOTTOM
    vta_RIGHT = ta_TOP
    vta_CENTER = ta_CENTER
    vta_BOTTOM = ta_RIGHT
    vta_TOP = ta_LEFT
    
End Enum

#If False Then
    Private Const ta_NOUPDATECP = 0, ta_UPDATECP = 1, ta_LEFT = 0, ta_RIGHT = 2, ta_CENTER = 6, ta_TOP = 0, ta_BOTTOM = 8, ta_BASELINE = 24, ta_RTLREADING = 256, ta_MASK = (ta_BASELINE + ta_CENTER + ta_UPDATECP + ta_RTLREADING)
    Private Const vta_BASELINE = ta_BASELINE, vta_LEFT = ta_BOTTOM, vta_RIGHT = ta_TOP, vta_CENTER = ta_CENTER, vta_BOTTOM = ta_RIGHT, vta_TOP = ta_LEFT
#End If

'GDI font creation and management
Private Declare Function AddFontResourceExW Lib "gdi32" (ByVal lFontName As Long, ByVal lFontCharacteristics As Long, ByVal lReserved As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectW" (ByRef lpLogFont As LOGFONTW) As Long
Private Declare Function RemoveFontResourceExW Lib "gdi32" (ByVal lFontName As Long, ByVal lFontCharacteristics As Long, ByVal lReserved As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Const FR_PRIVATE As Long = &H10 'Used by AddFontResourceEx

'Various non-font-specific WAPI functions helpful for font assembly
Private Const logPixelsX As Long = 88
Private Const LOGPIXELSY As Long = 90
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long

'Some system-specific font settings are cached at initialization time, and unchanged for the life of the program.
' (TODO: watch for relevant window messages on Win 8.1+ that may change these.)
Private m_LogPixelsX As Long, m_LogPixelsY As Long

'The name of this system's UI font is set here; all PD controls will render using this font face.
Private m_InterfaceFontName As String

'Internal font caches.  PD uses these to populate things like font selection dropdowns.
Private m_PDFontCache As pdStringStack
Private Const INITIAL_PDFONTCACHE_SIZE As Long = 64
Private m_LastFontAdded As String

'Fonts recently used by the user.  When the user chooses a font for something (like a text layer), it's up to the tool
' to relay that choice here so we we can keep our list up-to-date.  When we are notified of changes to the recent font list,
' we'll send out a notification that font selectors can listen for (and respond to).
Private m_RecentFonts As pdStringStack
Private Const RECENT_FONTS_FILENAME As String = "recent_fonts.xml"
Private m_MaxRecentFonts As Long

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
    numSupportedScripts As Integer
    
    'At present, this item is unused.  See the Uniscribe.GetScriptsSupportedByFont() function for more details.
    'SupportedScripts() As Long
    
End Type

'PD paints pretty much all of its own text.  Rather than burden each individual control with maintaining their own font object,
' we maintain a cache of the interface font at all requested sizes.  If an object needs to draw interface text, they can query
' us for a matching font object.
Private m_ProgramFontCollection As pdFontCollection

'To improve compile-time performance, we cache a dummy font object.  This object is ignored at run-time,
' but during compile-time, we return it for all GetMatchingUIFont() calls instead of using a more
' sophisticated caching system (as the cache gets thrashed by VB instantiating and destroying compile-time
' objects willy-nilly).
Private m_DummyFont As pdFont

'As of 2025.4, users can add arbitrary fonts from custom folders.  We must free these fonts at termination.
Private m_UserAddedFonts As pdStringStack

Public Sub DetermineUIFont()
    
    Dim tmpFontCheck As pdFont
    Set tmpFontCheck = New pdFont
    
    'Users can override PD's default font by adding a "UIFont" entry to the "Interface"
    ' segment of PD's settings file.  By default, this entry is *not* added because I
    ' don't want people changing this willy-nilly (because I can't guarantee that all
    ' UI elements will reflow correctly).
    Dim userFontEnabled As Boolean
    userFontEnabled = (LenB(UserPrefs.GetUIFontName()) > 0)
    If userFontEnabled Then userFontEnabled = tmpFontCheck.DoesFontExist(UserPrefs.GetUIFontName())
    If userFontEnabled Then
        m_InterfaceFontName = UserPrefs.GetUIFontName()
    
    'By default, PD uses "Segoe UI" if present; "Tahoma" otherwise
    Else
        If tmpFontCheck.DoesFontExist("Segoe UI") Then m_InterfaceFontName = "Segoe UI" Else m_InterfaceFontName = "Tahoma"
    End If
    
    Set tmpFontCheck = Nothing
    
End Sub

Public Function GetUIFontName() As String
    GetUIFontName = m_InterfaceFontName
End Function

'Want to draw program text onto something?  Call this function to find out what font size is required.
' If you will subsequently use the returned font size for testing, you can set "cacheIfNovel = True" to automatically cache a copy
' of the font at the newly detected font size.
Public Function FindFontSizeSingleLine(ByRef srcString As String, ByVal pxWidth As Long, ByVal initialFontSize As Single, Optional ByVal isBold As Boolean = False, Optional ByVal isItalic As Boolean = False, Optional ByVal isUnderline As Boolean = False, Optional ByVal cacheIfNovel As Boolean = True) As Single
    
    'Inside the designer, we need to make sure the font collection exists
    If (m_ProgramFontCollection Is Nothing) Then InitProgramFontCollection
    
    'Add this font size+style combination to the collection
    Dim fontIndex As Long
    fontIndex = m_ProgramFontCollection.AddFontToCache(Fonts.GetUIFontName(), initialFontSize, isBold, isItalic, isUnderline)
    
    'Retrieve a handle to that font
    Dim tmpFont As pdFont
    Set tmpFont = m_ProgramFontCollection.GetFontObjectByPosition(fontIndex)
    
    'Return a smaller font size, as necessary, to fit the requested pixel width
    FindFontSizeSingleLine = tmpFont.GetMaxFontSizeToFitStringWidth(srcString, pxWidth, initialFontSize)
    
    'If the caller plans to use this new font size for immediate rendering, immediately cache a copy of the font at this new size
    If cacheIfNovel And (FindFontSizeSingleLine <> initialFontSize) Then
        m_ProgramFontCollection.AddFontToCache Fonts.GetUIFontName(), FindFontSizeSingleLine, isBold, isItalic, isUnderline
    End If
    
End Function

'Same as FindFontSizeSingleLine(), above, but with support for wordwrap
' If you will subsequently use the returned font size for testing, you can set "cacheIfNovel = True" to automatically cache a copy
' of the font at the newly detected font size.
Public Function FindFontSizeWordWrap(ByRef srcString As String, ByVal pxWidth As Long, ByVal pxHeight As Long, ByVal initialFontSize As Single, Optional ByVal isBold As Boolean = False, Optional ByVal isItalic As Boolean = False, Optional ByVal isUnderline As Boolean = False, Optional ByVal cacheIfNovel As Boolean = True) As Single
    
    'Inside the designer, we need to make sure the font collection exists
    If (m_ProgramFontCollection Is Nothing) Then InitProgramFontCollection
    
    'Retrieve a handle to a matching pdFont object
    Dim tmpFont As pdFont
    Set tmpFont = Fonts.GetMatchingUIFont(initialFontSize, isBold, isItalic, isUnderline)
    
    'Return a smaller font size, as necessary, to fit the requested pixel width
    FindFontSizeWordWrap = tmpFont.GetMaxFontSizeToFitWordWrap(srcString, pxWidth, pxHeight, initialFontSize)
    
    'If the caller plans to use this new font size for immediate rendering, immediately cache a copy of the font at this new size
    If cacheIfNovel And (FindFontSizeWordWrap <> initialFontSize) Then
        m_ProgramFontCollection.AddFontToCache Fonts.GetUIFontName(), FindFontSizeWordWrap, isBold, isItalic, isUnderline
    End If
    
End Function

'Want direct access to a UI font instance?  Get one here.  Note that only size, bold, italic, and underline are currently matched,
' as PD doesn't use strikethrough fonts anywhere.
Public Function GetMatchingUIFont(ByVal srcFontSize As Single, Optional ByVal isBold As Boolean = False, Optional ByVal isItalic As Boolean = False, Optional ByVal isUnderline As Boolean = False) As pdFont
    
    'Inside the designer, we need to make sure the font collection exists
    If (m_ProgramFontCollection Is Nothing) Then InitProgramFontCollection
    
    'During compile-time, we don't need access to all of PD's font features.  Just return a dummy UI font
    ' unless the program is actually running
    If PDMain.IsProgramRunning Then

        'Add this font size+style combination to the collection, as necessary
        Dim fontIndex As Long
        fontIndex = m_ProgramFontCollection.AddFontToCache(m_InterfaceFontName, srcFontSize, isBold, isItalic, isUnderline)

        'Return the handle of the newly created (and/or previously cached) pdFont object
        Set GetMatchingUIFont = m_ProgramFontCollection.GetFontObjectByPosition(fontIndex)

    Else
        If (m_DummyFont Is Nothing) Then
            VBHacks.EnableHighResolutionTimers
            Set m_DummyFont = New pdFont
            m_DummyFont.SetFontPropsAllAtOnce m_InterfaceFontName, srcFontSize, False, False, False
            m_DummyFont.CreateFontObject
        Else
            If (srcFontSize <> m_DummyFont.GetFontSize) Then
                m_DummyFont.DeleteCurrentFont
                m_DummyFont.SetFontSize srcFontSize
                m_DummyFont.CreateFontObject
            End If
        End If
        Set GetMatchingUIFont = m_DummyFont
    End If
    
End Function

'If functions want their own copy of all available fonts on this PC, call this function
Public Function GetCopyOfSystemFontList(ByRef dstFontsSystem As pdStringStack, ByRef dstFontsRecent As pdStringStack) As Boolean
    If (dstFontsSystem Is Nothing) Then Set dstFontsSystem = New pdStringStack
    dstFontsSystem.CloneStack m_PDFontCache
    If (dstFontsRecent Is Nothing) Then Set dstFontsRecent = New pdStringStack
    dstFontsRecent.CloneStack m_RecentFonts
End Function

'If the caller just wants to know the size of a default string, it's better to use this function.  That spares them from having to
' create a redundant font object just to measure text.
Public Function GetDefaultStringHeight(ByVal FontSize As Single, Optional ByVal isBold As Boolean = False, Optional ByVal isItalic As Boolean = False, Optional ByVal isUnderline As Boolean = False) As Single
    Dim tmpFont As pdFont
    Set tmpFont = Fonts.GetMatchingUIFont(FontSize, isBold, isItalic, isUnderline)
    GetDefaultStringHeight = tmpFont.GetHeightOfString("FfAaBbCctbpqjy1234567890")
    Set tmpFont = Nothing
End Function

'If the caller just wants to measure string width, it's better to use this function.  That spares them from having to
' create a redundant font object just to measure text.
Public Function GetDefaultStringWidth(ByRef srcString As String, ByVal FontSize As Single, Optional ByVal isBold As Boolean = False, Optional ByVal isItalic As Boolean = False, Optional ByVal isUnderline As Boolean = False) As Single
    Dim tmpFont As pdFont
    Set tmpFont = Fonts.GetMatchingUIFont(FontSize, isBold, isItalic, isUnderline)
    GetDefaultStringWidth = tmpFont.GetWidthOfString(srcString)
    Set tmpFont = Nothing
End Function

'When the max number of recent fonts is changed (via Tools > Options > Fonts), notify us via this function.
' We'll update the "recently used" font list and send out notifications to font UI elements so they can update.
Public Sub NotifyRecentFontsMaxCount(ByVal newFontMax As Long)
    
    'Update our list of recent fonts to reflect the new maximum
    If (m_MaxRecentFonts <> newFontMax) Then
        
        m_MaxRecentFonts = newFontMax
        
        'Update the recent font list (as needed), and notify all loaded font dropdowns of the change.
        If (Not m_RecentFonts Is Nothing) Then
            If (m_MaxRecentFonts < m_RecentFonts.GetNumOfStrings()) Then
                m_RecentFonts.KeepTopNStringsOnly m_MaxRecentFonts
                If (Not g_WindowManager Is Nothing) Then UserControls.PostPDMessage WM_PD_FONTSUPDATED
            End If
        End If
        
    End If
    
End Sub

'When the user uses a font for a tool (e.g. changing the font for a text layer), notify us via this function.
' We'll update the "recently used" font list and send out notifications to font UI elements so they can update.
Public Sub NotifyFontUsed(ByRef srcFontName As String)
    
    'Update our list of recent fonts
    If (m_RecentFonts Is Nothing) Then Set m_RecentFonts = New pdStringStack
    m_RecentFonts.AddString_CheckDuplicatesFirst srcFontName
    
    'Ensure the list does not grow past the user's current max number of allowed recent fonts
    If (m_MaxRecentFonts > 0) Then m_RecentFonts.KeepTopNStringsOnly m_MaxRecentFonts
    
    'Notify any loaded font dropdowns of the change, so they can update themselves accordingly.
    If (Not g_WindowManager Is Nothing) Then UserControls.PostPDMessage WM_PD_FONTSUPDATED
    
End Sub

'Build a system font cache.  Note that this is an expensive operation, and should never be called more than once.
' RETURNS: 0 if failure, Number of fonts (>= 0) if successful.  (Note that the *total number of all fonts* is returned,
'          not just TrueType ones.)
Public Function BuildFontCaches() As Long
    
    Set m_PDFontCache = New pdStringStack
    Set m_RecentFonts = New pdStringStack
    
    'Retrieve the current system LOGFONT conversion values
    UpdateLogFontValues
    
    'Next, prep a full font list for the advanced text tool.
    '(We won't know the full number of available fonts until the Enum function finishes,
    ' so prep an extra-large buffer in advance.)
    m_PDFontCache.ResetStack INITIAL_PDFONTCACHE_SIZE
    GetAllAvailableFonts
    
    'Because the font cache(s) will potentially be accessed by tons of external functions,
    ' it pays to sort them just once, at initialization time.
    m_PDFontCache.TrimStack
    m_PDFontCache.SortAlphabetically True
    
    'TESTING ONLY!  Curious about the list of fonts?  Use this line to write it out to the immediate window
    'm_PDFontCache.DEBUG_dumpResultsToImmediateWindow
    PDDebug.LogAction "FYI - number of fonts found on this PC: " & m_PDFontCache.GetNumOfStrings
    
    'Next, we need to retrieve the user's previously saved "recent font list", if any
    Dim recentFontsFile As String
    recentFontsFile = UserPrefs.GetFontPath() & RECENT_FONTS_FILENAME
    
    If Files.FileExists(recentFontsFile) Then
        
        'Load the recent fonts file into a serializer class (it will help us retrieve settings from XML)
        Dim tmpFileContents As String
        
        Dim cSerialize As pdSerialize
        Set cSerialize = New pdSerialize
        If Files.FileLoadAsString(recentFontsFile, tmpFileContents, True) Then
            
            cSerialize.SetParamString tmpFileContents
            
            'Next, extract the list of recent fonts
            Dim tmpFontList As pdStringStack
            Set tmpFontList = New pdStringStack
            If tmpFontList.RecreateStackFromSerializedString(cSerialize.GetString("recent-fonts", vbNullString, True)) Then
                
                'Validate each font in the file; if one doesn't exist on this PC, don't load it
                If (tmpFontList.GetNumOfStrings > 0) Then
                    
                    Dim i As Long, srcFontName As String
                    For i = 0 To tmpFontList.GetNumOfStrings() - 1
                        srcFontName = tmpFontList.GetString(i)
                        If (m_PDFontCache.ContainsString(srcFontName, True) >= 0) Then
                            m_RecentFonts.AddString srcFontName
                        Else
                            'Font doesn't exist on this PC; don't load it!
                        End If
                    Next i
                    
                End If
                
            Else
                PDDebug.LogAction "WARNING: recent font list is corrupt!"
            End If
        Else
            PDDebug.LogAction "WARNING: couldn't load recent font file " & recentFontsFile
        End If
        
    'If the recent fonts file *doesn't* exist, that's fine - we don't need to do any additional initialization here.
    End If
    
    'Retrieve the max number of recent fonts we should maintain in the various font dropdowns across the program.
    m_MaxRecentFonts = UserPrefs.GetPref_Long("Interface", "recent-font-max", 10)
    m_RecentFonts.KeepTopNStringsOnly m_MaxRecentFonts
    
    'We have one other piece of initialization to do here.  Prep the program UI font cache that outside functions
    ' can use for their own UI painting.  This cache *only* uses the current app font, but in different sizes
    ' and styles.
    InitProgramFontCollection
    
End Function

'Converting between normal font sizes and GDI font sizes is convoluted, and it relies on a system-specific LOGPIXELSY value.
' We must cache that value before requesting fonts from the system.
Private Sub UpdateLogFontValues()
    Dim tmpDC As Long
    tmpDC = GDI.GetMemoryDC()
    m_LogPixelsX = GetDeviceCaps(tmpDC, logPixelsX)
    m_LogPixelsY = GetDeviceCaps(tmpDC, LOGPIXELSY)
    GDI.FreeMemoryDC tmpDC
End Sub

'Prep the program font cache.  Individual functions may need to call this inside the designer,
' because PD's normal run-time initialization steps won't have fired.
Private Sub InitProgramFontCollection()
    
    Set m_ProgramFontCollection = New pdFontCollection
    If (m_UserAddedFonts Is Nothing) Then Set m_UserAddedFonts = New pdStringStack
    
    'When outside callers request a copy of the system font, they are allowed to request any size+style they want.
    ' Font name, however, never varies, so tell the font cache to only compare size and style when matching font requests.
    m_ProgramFontCollection.SetCacheMode FCM_SizeAndStyle
    
End Sub

'Retrieve all available fonts on this PC, regardless of font type.
' As of 2025.4, this also retrieves fonts from custom folders added by the user.
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
    
    'Before enumerating all available fonts, load any available user fonts for this session.
    ' (Adding these *here* ensures they show up in subsequent EnumFontFamilesEx() calls.)
    InitializeUserFonts
    
    'Enumerate font families using a temporary DC
    Dim tmpDC As Long
    tmpDC = GDI.GetMemoryDC()
    EnumFontFamiliesEx tmpDC, tmpLogFont, AddressOf EnumFontFamExProc, ByVal 0&, 0&
    GDI.FreeMemoryDC tmpDC
    
    'If at least one font was found, return TRUE
    GetAllAvailableFonts = (m_PDFontCache.GetNumOfStrings > 0)

End Function

Private Sub InitializeUserFonts()
    
    Dim lstFontFolders As pdStringStack
    Set lstFontFolders = New pdStringStack
    
    'Always add the default PD font folder
    lstFontFolders.AddString UserPrefs.GetFontPath()
    
    'Load any/all custom font folders saved in previous sessions
    Const FONT_PRESETS_FILE As String = "font_folders.txt"
    Dim srcFile As String
    srcFile = UserPrefs.GetPresetPath & FONT_PRESETS_FILE
    If Files.FileExists(srcFile) Then
        
        'Load preset file
        Dim srcList As String
        If Files.FileLoadAsString(srcFile, srcList, True) Then
            
            'Iterate lines in the file
            Dim cStack As pdStringStack: Set cStack = New pdStringStack
            If cStack.CreateFromMultilineString(srcList, vbCrLf) Then
                
                Dim i As Long
                For i = 0 To cStack.GetNumOfStrings - 1
                    srcFile = cStack.GetString(i)
                    If (LenB(srcFile) > 0) Then
                        If Files.PathExists(srcFile, False) Then lstFontFolders.AddString srcFile
                    End If
                Next i
                
            End If
            
        End If
        
    End If
    
    'Iterate all font folders, iterate files *within* those folders, and add each in turn.
    Dim srcFontFolder As String
    Do While lstFontFolders.PopString(srcFontFolder)
        
        'Only retrieve actual font files, not text or license or other files
        Dim lstFontFiles As pdStringStack
        Set lstFontFiles = New pdStringStack
        If Files.RetrieveAllFiles(srcFontFolder, lstFontFiles, True, False, "ttf|ttc|otf") Then
            
            If (m_UserAddedFonts Is Nothing) Then Set m_UserAddedFonts = New pdStringStack
            
            'Iterate all discovered fonts
            Dim srcFontFile As String
            Do While lstFontFiles.PopString(srcFontFile)
                
                Dim addedSuccessfully As Boolean
                addedSuccessfully = (AddFontResourceExW(StrPtr(srcFontFile), FR_PRIVATE, 0&) <> 0&)
                
                'On successful additions to GDI, note the added file, then add the same font
                ' to GDI+ (which ensures availability to GDI+ font calls).
                If addedSuccessfully Then
                    m_UserAddedFonts.AddString srcFontFile
                    GDIPlus_AddRuntimeFont srcFontFile
                End If
                
            Loop
            
        '/End If: 1+ files retrieved
        End If
        
    Loop
    
End Sub

Public Function UserFonts_GetNumAdded() As Long
    If (Not m_UserAddedFonts Is Nothing) Then UserFonts_GetNumAdded = m_UserAddedFonts.GetNumOfStrings()
End Function

'At shutdown, release any/all loaded fonts
Public Sub ReleaseUserFonts()
    
    'Start by unloading GDI+ fonts
    GDI_Plus.GDIPlus_ReleaseRuntimeFonts
    
    'Next comes GDI fonts
    If (Not m_UserAddedFonts Is Nothing) Then
        
        Dim srcFontFile As String
        Do While m_UserAddedFonts.PopString(srcFontFile)
            RemoveFontResourceExW StrPtr(srcFontFile), FR_PRIVATE, 0&
        Loop
        
    End If
    
    'Finally, save the user's recently used fonts (in font dropdowns), if any
    If (Not m_RecentFonts Is Nothing) Then
        If (m_RecentFonts.GetNumOfStrings > 0) Then
             
            'Save the current list of recent fonts to a normal PD XML object
            Dim cSerialize As pdSerialize
            Set cSerialize = New pdSerialize
            cSerialize.Reset
            cSerialize.AddParam "recent-fonts", m_RecentFonts.SerializeStackToSingleString()
            
            'Use the /Data/Fonts folder for storage
            Dim fontFolder As String
            fontFolder = UserPrefs.GetFontPath()
            
            'Dump the serialized settings to file
            If (Not Files.FileSaveAsText(cSerialize.GetParamString(), fontFolder & RECENT_FONTS_FILENAME, True, True)) Then
                PDDebug.LogAction "WARNING: couldn't save recent font list!  Font dropdowns may not store last font correctly."
            End If
            
        End If
    End If
    
End Sub

'Callback function for EnumFontFamiliesEx
Public Function EnumFontFamExProc(ByRef lpElfe As LOGFONTW, ByRef lpNtme As NEWTEXTMETRIC, ByVal srcFontType As Long, ByVal lParam As Long) As Long

    'Start by retrieving the font face name from the LogFontW struct
    Dim thisFontFace As String
    thisFontFace = String$(LF_FACESIZEA, 0)
    CopyMemoryStrict StrPtr(thisFontFace), VarPtr(lpElfe.lfFaceName(0)), LF_FACESIZEW
    thisFontFace = Strings.TrimNull(thisFontFace)
    
    'Perform some basic checks to see if this font is usable
    Dim fontUsable As Boolean
    fontUsable = (LenB(thisFontFace) > 0)
    
    'If this font face is identical to the previous font face, do not add it
    If fontUsable Then fontUsable = Strings.StringsNotEqual(thisFontFace, m_LastFontAdded, False)
    
    'We also want to ignore fonts with @ in front of their name, as these are merely duplicates of existing fonts.
    ' (The @ signifies improved support for vertical text, which may someday be useful... but right now I have enough
    '  on my plate without worrying about that.)
    If fontUsable Then fontUsable = Strings.StringsNotEqual(Left$(thisFontFace, 1), "@", False)
    
    'For now, we are also ignoring raster fonts, as they create unwanted complications
    Const RASTER_FONTTYPE As Long = &H1
    If fontUsable Then fontUsable = ((srcFontType And RASTER_FONTTYPE) = 0)
    
    'If this font is a worthy addition, add it now
    If fontUsable Then
        
        m_PDFontCache.AddString thisFontFace
        
        'Make a copy of the last added font name, so we can ignore duplicates
        m_LastFontAdded = thisFontFace
        
        'NOTE: Perhaps it could be helpful to cache the font type, or possibly use it to determine if fonts should be ignored?
        ' (At present, this is ignored in favor of a more extensive, Uniscribe-based analysis that determines actual
        '  full-fledged Unicode range support.)
        'm_PDFontCache(m_NumOfFonts).FontType = srcFontType
        
    End If
    
    'Return 1 so the enumeration continues
    EnumFontFamExProc = 1
    
End Function

'Given some standard font characteristics (font face, style, etc), fill a corresponding LOGFONTW struct with matching values.
' This is helpful as PD stores characteristics in VB-friendly formats (e.g. booleans for styles), while LOGFONTW uses custom
' descriptors (e.g. font size, which is not calculated in an obvious way).
Public Sub FillLogFontW_Basic(ByRef dstLogFontW As LOGFONTW, ByRef srcFontFace As String, ByVal srcFontBold As Boolean, ByVal srcFontItalic As Boolean, ByVal srcFontUnderline As Boolean, ByVal srcFontStrikeout As Boolean)

    With dstLogFontW
    
        'For Unicode compatibility, the font face must be copied directly, without internal VB translations
        Dim copyLength As Long
        copyLength = LenB(srcFontFace)
        If (copyLength > LF_FACESIZEW) Then copyLength = LF_FACESIZEW
        CopyMemoryStrict VarPtr(.lfFaceName(0)), StrPtr(srcFontFace), copyLength
        
        'Bold is a unique style, because it must be translated to a corresponding weight measurement
        If srcFontBold Then .lfWeight = fw_Bold Else .lfWeight = fw_Normal
        
        'Other styles all use the same pattern: multiply the bool by -1 to obtain a matching byte-type style
        .lfItalic = -1 * srcFontItalic
        .lfUnderline = -1 * srcFontUnderline
        .lfStrikeOut = -1 * srcFontStrikeout
        
        'While we're here, set charset to the default value; PD does not deviate from this (at present)
        .lfCharSet = DEFAULT_CHARSET
        
    End With
    
End Sub

'Fill a LOGFONTW struct with a matching PD font size (typically in pixels, but points are also supported)
Public Sub FillLogFontW_Size(ByRef dstLogFontW As LOGFONTW, ByVal FontSize As Single, ByVal fontMeasurementUnit As PD_FontUnit)

    With dstLogFontW
        
        Select Case fontMeasurementUnit
        
            'Pixels use a modified version of the standard Windows formula; note that this assumes
            ' 96 DPI at present - high DPI systems still need testing!  TODO!
            Case fu_Pixel
                
                'Convert font size to points
                FontSize = FontSize * 0.75      '(72 / 96, technically, where 96 is the current screen DPI)
                
                'Use the standard point-based formula
                .lfHeight = Fonts.ConvertToGDIFontSize(FontSize)
                
            'Points are converted using a standard Windows formula; see https://msdn.microsoft.com/en-us/library/dd145037%28v=vs.85%29.aspx
            Case fu_Point
                .lfHeight = Fonts.ConvertToGDIFontSize(FontSize)
        
        End Select
        
        'Per convention, font width is set to 0 so the font mapper can select an aspect-ratio preserved width for us
        .lfWidth = 0
        
    End With
    
End Sub

Public Function ConvertToGDIFontSize(ByVal srcFontSize As Single) As Long
    If (m_LogPixelsY = 0#) Then UpdateLogFontValues
    ConvertToGDIFontSize = -1# * Internal_MulDiv(srcFontSize, m_LogPixelsY, 72!)
End Function

'It really isn't necessary to rely on the system MulDiv values for the sizes used for fonts.
Private Function Internal_MulDiv(ByVal srcNumber As Single, ByVal srcNumerator As Single, ByVal srcDenominator As Single) As Long
    Internal_MulDiv = Int((srcNumber * srcNumerator) / srcDenominator)
End Function

'Once I have a better idea of what I can do with font quality, I'll be switching the fontQuality Enum to something internal to PD.
' But right now, I'm still in the exploratory phase, and trying to figure out whether different font quality settings affect
' the glyph outline returned.  (They should, technically, since hinting affects font shape.)
Public Sub FillLogFontW_Quality(ByRef dstLogFontW As LOGFONTW, ByVal fontQuality As GP_TextRenderingHint)

    Dim gdiFontQuality As GDI_FontQuality
    
    'Per http://stackoverflow.com/questions/1203087/why-is-graphics-measurestring-returning-a-higher-than-expected-number?lq=1
    ' this function mirrors the .NET conversion from GDI+ text rendering hints to GDI font quality settings.  Mapping deliberately
    ' ignores some settings (no idea why, but if the .NET stack does it, there's probably a reason)
    Select Case fontQuality
    
        Case TextRenderingHintSystemDefault
            gdiFontQuality = fq_Default
            
        Case TextRenderingHintSingleBitPerPixel
            gdiFontQuality = fq_Draft
        
        Case TextRenderingHintSingleBitPerPixelGridFit
            gdiFontQuality = fq_Proof
        
        Case TextRenderingHintAntiAlias
            gdiFontQuality = fq_Antialiased
        
        Case TextRenderingHintAntiAliasGridFit
            gdiFontQuality = fq_Antialiased
        
        Case TextRenderingHintClearTypeGridFit
            gdiFontQuality = fq_ClearType
        
        Case Else
            Debug.Print "Unknown font quality passed; please double-check requests to fillLogFontW_Quality"
    
    End Select
    
    dstLogFontW.lfQuality = gdiFontQuality

End Sub

'Retrieve a text metrics struct for a given DC.  Obviously, the desired font needs to be selected into the DC *prior* to calling this.
Public Function FillTextMetrics(ByRef srcDC As Long, ByRef dstTextMetrics As TEXTMETRIC) As Boolean
    Dim gtmReturn As Long
    gtmReturn = GetTextMetrics(srcDC, dstTextMetrics)
    FillTextMetrics = (gtmReturn <> 0)
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
    FillOutlineTextMetrics = (gtmReturn <> 0)
    If FillOutlineTextMetrics Then CopyMemoryStrict VarPtr(dstOutlineMetrics), VarPtr(tmpBytes(0)), LenB(dstOutlineMetrics)
    
End Function

'Given a filled LOGFONTW struct (hopefully filled by the fillLogFontW_* functions above!), attempt to create an actual font object.
' Returns TRUE if successful; FALSE otherwise.
Public Function CreateGDIFont(ByRef srcLogFont As LOGFONTW, ByRef dstFontHandle As Long) As Boolean
    dstFontHandle = CreateFontIndirect(srcLogFont)
    CreateGDIFont = (dstFontHandle <> 0)
    If CreateGDIFont Then PDDebug.UpdateResourceTracker PDRT_hFont, 1
End Function

'Delete a GDI font; returns TRUE if successful
Public Function DeleteGDIFont(ByVal srcFontHandle As Long) As Boolean
    DeleteGDIFont = (DeleteObject(srcFontHandle) <> 0)
    If DeleteGDIFont Then PDDebug.UpdateResourceTracker PDRT_hFont, -1
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
        tmpDC = GDI.GetMemoryDC()
        origFont = SelectObject(tmpDC, srcFontHandle)
        
        'Retrieve the character positioning values
        gdiReturn = GetCharABCWidthsFloat(tmpDC, charCodeInQuestion, charCodeInQuestion, VarPtr(dstABCFloat))
        
        'Release the font
        SelectObject tmpDC, origFont
        GDI.FreeMemoryDC tmpDC
    
    End If
    
    'GetCharABCWidthsFloat() returns a non-zero value if successful
    GetABCWidthOfGlyph = (gdiReturn <> 0)
    
End Function

'Given a font name, quickly generate a GDI font handle with default settings, and shove it into a temporary DC.
' IMPORTANT NOTE: the caller needs to cache the font and DC handle, then pass them to the clean-up function below
Public Function QuickCreateFontAndDC(ByRef srcFontName As String, ByRef dstFont As Long, ByRef dstDC As Long) As Boolean
    
    Dim tmpLogFont As LOGFONTW
    FillLogFontW_Basic tmpLogFont, srcFontName, False, False, False, False
    QuickCreateFontAndDC = CreateGDIFont(tmpLogFont, dstFont)
    
    'Create a temporary DC and select the font into it
    If QuickCreateFontAndDC Then
        dstDC = GDI.GetMemoryDC()
        SelectObject dstDC, dstFont
    End If
    
End Function

Public Sub QuickDeleteFontAndDC(ByRef srcFont As Long, ByRef srcDC As Long)
    
    'Remove the font
    SelectObject srcDC, GetStockObject(SYSTEM_FONT)
    
    'Kill both the font and the DC
    If (DeleteObject(srcFont) <> 0) Then PDDebug.UpdateResourceTracker PDRT_hFont, -1
    GDI.FreeMemoryDC srcDC
    
End Sub
