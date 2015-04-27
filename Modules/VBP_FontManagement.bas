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
'Note: these types are used in the callback function for EnumFontFamiliesEx; as such, I have declared them as public,
'       despite them not really being used anywhere but inside this module.

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

'END FONT-SPECIFIC DECLARATIONS
'****************************************************************************************

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


