VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Palette Container
'Copyright 2018-2018 by Tanner Helland
'Created: 16/January/18
'Last updated: 23/January/18
'Last update: add support for Adobe color swatches (.aco format)
'
'This class currently has some non-obvious overlap with the Palettes module.  In the future, it would be nice to
' compartmentalize more palette-specific functionality within this class, then pass around class instances instead
' of bare RGBQuad arrays.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Declare Sub CopyMemory_Strict Lib "kernel32" Alias "RtlMoveMemory" (ByVal dstPointer As Long, ByVal srcPointer As Long, ByVal numOfBytes As Long)

'Adobe swatch format requires custom color space enums
Private Enum AdobeSwatchColorSpace
    
    ASCS_RGB = 0
    ASCS_HSB = 1
    ASCS_CMYK = 2
    ASCS_Lab = 7
    ASCS_Gray = 8
    
    'The following color spaces are not publicly documented.  We cannot currently decode them.
    ASCS_Pantone = 3
    ASCS_Focoltone = 4
    ASCS_Trumatch = 5
    ASCS_Toyo88 = 6
    ASCS_HKS = 10

End Enum

#If False Then
    Private Const ASCS_RGB = 0, ASCS_HSB = 1, ASCS_CMYK = 2, ASCS_Lab = 7, ASCS_Gray = 8, ASCS_Pantone = 3, ASCS_Focoltone = 4, ASCS_Trumatch = 5, ASCS_Toyo88 = 6, ASCS_HKS = 10
#End If

Private m_PaletteFilename As String
Private m_PaletteName As String
Private m_PaletteColorCount As Long
Private m_PaletteColors() As RGBQuad

'A specially designed QuickSort algorithm is used to sort palettes.  We define our own internal sort criteria
' using this custom type, which allows us to implement different types of palette sorts "on the fly" by simply
' plugging-in different sort criterias.
Private Type PaletteSort
    pSortCriteria As Single
    pOrigIndex As Long
End Type

Friend Function GetPaletteColorCount() As Long
    GetPaletteColorCount = m_PaletteColorCount
End Function

Friend Function GetPaletteFilename() As String
    GetPaletteFilename = m_PaletteFilename
End Function

Friend Function GetPaletteName() As String
    GetPaletteName = m_PaletteName
End Function

Friend Sub SetPaletteName(ByRef newPaletteName As String)
    m_PaletteName = newPaletteName
End Sub

Friend Function CopyPaletteToArray(ByRef dstPalette() As RGBQuad) As Boolean

    CopyPaletteToArray = (m_PaletteColorCount > 0)
    
    If CopyPaletteToArray Then
        
        On Error Resume Next
        If (LBound(dstPalette) <> 0) Or (UBound(dstPalette) <> m_PaletteColorCount - 1) Then ReDim dstPalette(0 To m_PaletteColorCount - 1) As RGBQuad
        On Error GoTo 0
        
        CopyMemory_Strict VarPtr(dstPalette(0)), VarPtr(m_PaletteColors(0)), m_PaletteColorCount * 4
    
    End If

End Function

'Given a path to a supported palette file, return TRUE if the file can be successfully parsed for palette data; FALSE otherwise
Friend Function LoadPaletteFromFile(ByRef srcFile As String) As Boolean

    'Branch according to format; on failure, attempt the next mechanism.  Note that some formats
    ' do enforce set file extensions, and we attempt to honor that here.
    
    'Look for Adobe's basic color table file (*much* simpler than swatches)
    If Strings.StringsEqual(Files.FileGetExtension(srcFile), "act", True) Then LoadPaletteFromFile = LoadPaletteAdobeColorTable(srcFile)
    
    'Look for Adobe color swatches next
    If (Not LoadPaletteFromFile) Then LoadPaletteFromFile = LoadPaletteAdobeSwatch(srcFile)
    
    'Look for GIMP palettes next
    If (Not LoadPaletteFromFile) Then LoadPaletteFromFile = LoadPaletteGIMP(srcFile)
    
    'Look for PaintShop Pro palettes next
    If (Not LoadPaletteFromFile) Then LoadPaletteFromFile = LoadPalettePaintShopPro(srcFile)
    
    'Look for Paint.NET palettes next
    If (Not LoadPaletteFromFile) And (Strings.StringsEqual(Files.FileGetExtension(srcFile), "txt", True)) Then LoadPaletteFromFile = LoadPalettePaintDotNet(srcFile)
    
End Function

'Given a valid path to a Photoshop-format .act file (Adobe Color Table), return an array of RGBQuad entries
Private Function LoadPaletteAdobeSwatch(ByRef srcFile As String, Optional ByVal optimizeForPerformance As Boolean = False) As Boolean
    
    On Error GoTo BadSwatchData
    
    'This function references the Adobe spec frequently, particularly for things like magic numbers.  Here is the
    ' link I used (good as of Jan '18): http://www.adobe.com/devnet-apps/photoshop/fileformatashtml/#50577411_31265
    
    'ACO files contain no validation entries in their header, so we use the file extension as initial validation.
    ' (We also check file size to ensure that the file is large enough to contain at least one color entry.
    ' The "magic number" file size is calculated using data from the spec.)
    If (Strings.StringsEqual(Files.FileGetExtension(srcFile), "aco", True) And (Files.FileLenW(srcFile) >= 14)) Then
        
        'Retrieve the file as a pdStream object (which makes traversal *much* easier).
        Dim srcData As pdStream
        If Files.FileLoadAsPDStream(srcFile, srcData) Then
            
            'Owing to their Mac heritage, Adobe files store everything as big-endian.  As such, we'll be using
            ' non-standard data retrieval calls from our pdStream object.
            
            'Note also that per the spec, swatch files are just long lists of 16-bit words.  This makes
            ' traversal straightforward.
            
            'Start by ensuring that the initial version number is good.  It needs to be 1(Apparently version "0" and "1" are both
            ' valid descriptors for version 1 files - go figure?)
            Dim srcInt As Integer
            srcInt = srcData.ReadInt_BE()
            If ((srcInt = 1) Or (srcInt = 0)) Then
                
                'Read the number of colors stored in the file
                Dim numOfColors As Long
                numOfColors = srcData.ReadIntUnsigned_BE()
                If (numOfColors > 0) Then
                    
                    Debug.Print "Adobe swatch contains " & CStr(numOfColors) & " colors."
                    
                    'Prep our internal palette structs to match the number of retrieved colors.  Note that we currently
                    ' limit the retrieved color count to 256 (although Adobe swatch counts can be 16-bits!).
                    Dim limitedMax As Long
                    limitedMax = numOfColors
                    
                    If (limitedMax > 256) Then limitedMax = 256
                    ReDim m_PaletteColors(0 To limitedMax - 1) As RGBQuad
                    m_PaletteColorCount = 0
                    
                    Dim i As Long
                    For i = 0 To numOfColors - 1
                        
                        If (m_PaletteColorCount < 256) Then
                        
                            'We now want to retrieve each color in turn.  Color swatches use a standardized 10-byte format,
                            ' comprised of five unsigned 16-bit entries:
                            ' 1) Color space (an internal Adobe enum)
                            ' 2) w, x, y, z values, which must be post-processed differently depending on (1)
                            
                            'Start by retrieving the color space
                            Dim curSpace As AdobeSwatchColorSpace
                            curSpace = srcData.ReadInt_BE()
                            
                            'A dummy int can be used to advance the file pointer
                            Dim dummyInt As Integer
                            
                            Select Case curSpace
                            
                                Case ASCS_RGB
                                    
                                    'RGB constants are stored as 16-bit unsigned ints, plus an empty two bytes at the end.
                                    ' (PD downconverts the 16-bit ints to 8-bit.)
                                    Dim r As Long, g As Long, b As Long
                                    r = srcData.ReadIntUnsigned_BE()
                                    g = srcData.ReadIntUnsigned_BE()
                                    b = srcData.ReadIntUnsigned_BE()
                                    dummyInt = srcData.ReadInt()
                                    
                                    'Downconvert the 16-bit entries to 8-bit
                                    r = (r \ 256)
                                    g = (g \ 256)
                                    b = (b \ 256)
                                    
                                    'Failsafe checks
                                    If (r > 255) Then r = 255
                                    If (g > 255) Then g = 255
                                    If (b > 255) Then b = 2555
                                    
                                    With m_PaletteColors(m_PaletteColorCount)
                                        .Red = r
                                        .Green = g
                                        .Blue = b
                                        .Alpha = 255
                                    End With
                                    
                                    m_PaletteColorCount = m_PaletteColorCount + 1
                                    
                                Case ASCS_CMYK
                                    Dim c As Long, m As Long, y As Long, k As Long
                                    c = srcData.ReadIntUnsigned_BE()
                                    m = srcData.ReadIntUnsigned_BE()
                                    y = srcData.ReadIntUnsigned_BE()
                                    k = srcData.ReadIntUnsigned_BE()
                                    
                                    'Downconvert the 16-bit entries to 8-bit and invert
                                    c = 255 - (c \ 256)
                                    m = 255 - (m \ 256)
                                    y = 255 - (y \ 256)
                                    k = 255 - (k \ 256)
                                    
                                    'Convert these to RGB literals using an entirely different spec as our guide
                                    ' (https://web.archive.org/web/20090909133852/http://partners.adobe.com/public/developer/en/ps/psrefman.pdf)
                                    r = 255 - PDMath.Min2Int(255, c + k)
                                    g = 255 - PDMath.Min2Int(255, m + k)
                                    b = 255 - PDMath.Min2Int(255, y + k)
                                    
                                    'Failsafe checks
                                    If (r > 255) Then r = 255
                                    If (g > 255) Then g = 255
                                    If (b > 255) Then b = 2555
                                    
                                    With m_PaletteColors(m_PaletteColorCount)
                                        .Red = r
                                        .Green = g
                                        .Blue = b
                                        .Alpha = 255
                                    End With
                                    
                                    m_PaletteColorCount = m_PaletteColorCount + 1
                                
                                'Grayscale is inexplicably stored on the range [0, 10000]
                                Case ASCS_Gray
                                    g = srcData.ReadIntUnsigned_BE()
                                    dummyInt = srcData.ReadInt()
                                    dummyInt = srcData.ReadInt()
                                    dummyInt = srcData.ReadInt()
                                
                                    'Translate gray into a usable RGB space
                                    g = CLng(CStr(g) / 39.0625)
                                    
                                    'Failsafe checks
                                    If (g > 255) Then g = 255
                                    
                                    With m_PaletteColors(m_PaletteColorCount)
                                        .Red = g
                                        .Green = g
                                        .Blue = b
                                        .Alpha = 255
                                    End With
                                    
                                    m_PaletteColorCount = m_PaletteColorCount + 1
                                    
                                'These two spaces could be implemented without too much trouble, but I don't currently have
                                ' swatches available to test their correctness:
                                'Case ASCS_HSB
                                'Case ASCS_Lab
                                                            
                                'All other color spaces are undocumented and unsupported
                                Case Else
                                    InternalProblem "Adobe Swatch color space is unsupported (" & CStr(curSpace) & ")"
                                    
                                    'Advance the pointer to the next color
                                    srcData.SetPosition 8, FILE_CURRENT
                            
                            End Select
                            
                        'End (palette count < 256)
                        End If
                        
                    Next i
                    
                    'TODO: look for a v2 swatch following the v1 entry.  All data will be identical, except colors
                    ' can also provide name descriptors.
                    
                    'Make sure at least one valid color was found
                    If (m_PaletteColorCount > 0) Then
                    
                        'As a failsafe, remove duplicate entries
                        FindAndRemoveDuplicates
                        
                        'If we haven't errored out, consider this a successful parse
                        m_PaletteName = Files.FileGetName(srcFile, True)
                        m_PaletteFilename = srcFile
                        LoadPaletteAdobeSwatch = True
                        
                    End If
                
                Else
                    InternalProblem "ACO file reported empty palette (" & CStr(numOfColors) & ")"
                End If
                
            Else
                InternalProblem "ACO file reported bad version number (" & CStr(srcInt) & ")"
            End If
            
        End If
        
    End If
    
    Exit Function
    
BadSwatchData:
    InternalProblem "LoadPaletteAdobeSwatch error # " & CStr(Err.Number) & ": " & Err.Description
    m_PaletteName = vbNullString
    m_PaletteColorCount = 0
    m_PaletteFilename = vbNullString
    LoadPaletteAdobeSwatch = False
    
End Function

'PaintShop Pro palettes use multiple extensions, but all are just text files with a line-delimited list of
' RGB triples.
Private Function LoadPalettePaintShopPro(ByRef srcFile As String, Optional ByVal optimizeForPerformance As Boolean = False) As Boolean

    On Error GoTo BadPaintShopProPalette
    
    'PaintShop Pro palettes have a very basic header, and they use the standardized file extensions
    ' ".pal" (old format) and ".psppalette" (new format).  To my knowledge, the file contents themselves
    ' are identical for both formats.
    If (Strings.StringsEqual(Files.FileGetExtension(srcFile), "pal", True) Or Strings.StringsEqual(Files.FileGetExtension(srcFile), "psppalette", True)) Then
        
        'Dump the full file into a string
        Dim rawFileString As String
        LoadPalettePaintShopPro = Files.FileLoadAsString(srcFile, rawFileString, True)
        
        If LoadPalettePaintShopPro Then
    
            'Prep our internal palette structs.  Note that PaintShop Pro palettes have a maximum size
            ' of 96 colors (per https://www.getpaint.net/doc/latest/WorkingWithPalettes.html),
            ' but we will automatically resize the collection to remove duplicate entries after
            ' all colors are loaded.
            ReDim m_PaletteColors(0 To 255) As RGBQuad
            m_PaletteColorCount = 0
            
            'To simplify processing, split the string by lines.
            Dim fileLines As pdStringStack
            Set fileLines = New pdStringStack
            fileLines.CreateFromMultilineString rawFileString
            
            'Parse each line in turn.  (Normally we would pop lines like a stack, but the user may want
            ' to preserve the original palette order - in which case we need to also process the text
            ' lines in their original order.)
            Const SPACE_CHAR As String = " "
            Dim curLine As String, i As Long
            For i = 0 To fileLines.GetNumOfStrings - 1
                
                curLine = Trim$(LCase$(fileLines.GetString(i)))
                
                If (m_PaletteColorCount < 256) Then
                    
                    'Start by looking for at least two spaces in the trimmed string (indicating at least three unique entries)
                    If (InStr(1, curLine, SPACE_CHAR, vbBinaryCompare) <> InStrRev(curLine, SPACE_CHAR, -1, vbBinaryCompare)) Then
                    
                        'This string contains two spaces.  Extract the first string-delimited entry.
                        Dim targetColor As String, targetColorInt As Long
                        targetColor = Left$(curLine, InStr(1, curLine, SPACE_CHAR, vbBinaryCompare) - 1)
                        
                        'Attempt to convert this to a number; if it fails, that's okay; this is some kind of invalid line
                        ' and we can ignore further parsing.
                        On Error GoTo BadLineColor
                        m_PaletteColors(m_PaletteColorCount).Red = CByte(targetColor)
                        On Error GoTo 0
                        
                        'Trim the color we've parsed out of the string, then repeat the above steps
                        curLine = Trim$(Right$(curLine, Len(curLine) - InStr(1, curLine, SPACE_CHAR, vbBinaryCompare)))
                        targetColor = Left$(curLine, InStr(1, curLine, SPACE_CHAR, vbBinaryCompare) - 1)
                        On Error GoTo BadLineColor
                        m_PaletteColors(m_PaletteColorCount).Green = CByte(targetColor)
                        On Error GoTo 0
                        
                        '...and one last time, for the blue component.  Note that the resulting string may not
                        ' have a trailing space, so we forcibly add one to simplify the parser.
                        curLine = Trim$(Right$(curLine, Len(curLine) - InStr(1, curLine, SPACE_CHAR, vbBinaryCompare))) & SPACE_CHAR
                        targetColor = Left$(curLine, InStr(1, curLine, SPACE_CHAR, vbBinaryCompare) - 1)
                        On Error GoTo BadLineColor
                        m_PaletteColors(m_PaletteColorCount).Blue = CByte(targetColor)
                        On Error GoTo 0
                        
                        'PSP palettes do not support alpha channels.  Forcibly set a value of 255.
                        m_PaletteColors(m_PaletteColorCount).Alpha = 255
                        
                        'If we made it all the way here, this line was successfully parsed for color data.
                        ' Advance the color count tracker and resume the line parser.
                        m_PaletteColorCount = m_PaletteColorCount + 1
BadLineColor:
                    End If
                    
                End If
            
            Next i
            
            'If we haven't errored out, consider this a valid parse if we found any valid colors
            LoadPalettePaintShopPro = (m_PaletteColorCount > 0)
            
            'If the load was successful, assign a fake palette name using the filename
            If LoadPalettePaintShopPro Then
                m_PaletteName = Files.FileGetName(srcFile, True)
                m_PaletteFilename = srcFile
            End If
            
            'On valid loads, check and remove any duplicate palette entries.  (As an ASCII format,
            ' PSP palettes make no guarantees, and we don't want to waste time palette-matching
            ' duplicate entries.)
            FindAndRemoveDuplicates
            
            'If the user plans to use this palette for color-matching (and really, what else are they gonna
            ' do with it?) we can pre-sort the palette to improve matching performance.
            If optimizeForPerformance Then SortByLuminance
            
        Else
            LoadPalettePaintShopPro = False
        End If
        
    End If
    
    Exit Function
    
BadPaintShopProPalette:
    m_PaletteName = vbNullString
    m_PaletteColorCount = 0
    m_PaletteFilename = vbNullString
    LoadPalettePaintShopPro = False

End Function

'Paint.NET palettes are just text files with a list of hex chars.  Unlike other palette formats,
' they *do* support alpha values.
Private Function LoadPalettePaintDotNet(ByRef srcFile As String, Optional ByVal optimizeForPerformance As Boolean = False) As Boolean

    On Error GoTo BadPaintDotNetPalette
    
    'Paint.NET palettes don't have a header for guaranteed byte length.  Instead, we simply look for valid
    ' ARGB hex descriptors inside a plain text file.
    If Strings.StringsEqual(Files.FileGetExtension(srcFile), "txt", True) Then
        
        'Dump the full file into a string
        Dim rawFileString As String
        LoadPalettePaintDotNet = Files.FileLoadAsString(srcFile, rawFileString, True)
        
        If LoadPalettePaintDotNet Then
    
            'Prep our internal palette structs.  Note that Paint.NET palettes are technically limited
            ' to 16 or 256 colors (https://www.codeproject.com/articles/31964/palettes-you-ve-gotta-love-them)
            ' but we follow the modern convention and basically ignore that number completely in favor
            ' of just parsing as many colors as we find in the file.
            ReDim m_PaletteColors(0 To 255) As RGBQuad
            m_PaletteColorCount = 0
            
            'To simplify processing, split the string by lines.
            Dim fileLines As pdStringStack
            Set fileLines = New pdStringStack
            fileLines.CreateFromMultilineString rawFileString
            
            'Before continuing, let's parse a few basic header lines (which are standard across all
            ' PSP palette formats).
            
            'First up is the JASC identifier
            Dim curLine As String
            curLine = Trim$(fileLines.GetString(0))
            
            Dim okayToParseColors As Boolean
            okayToParseColors = False
            
            If Strings.StringsEqual(curLine, "JASC-PAL", True) Then
                
                okayToParseColors = True
                
                'Next, check the version.  To my knowledge, this never varies from 0100
                curLine = Trim$(fileLines.GetString(1))
                #If DEBUGMODE = 1 Then
                    If Strings.StringsNotEqual(curLine, "0100", False) Then pdDebug.LogAction "WARNING!  JASC palette contains unexpected version number: " & curLine
                #End If
                
                'Finally, retrieve line 3.  This line should be some number on the range [1, 256].
                curLine = Trim$(fileLines.GetString(2))
                If (CLng(curLine) < 1 Or CLng(curLine) > 256) Then
                    #If DEBUGMODE = 1 Then
                        pdDebug.LogAction "WARNING!  JASC palette contains invalid color count: " & curLine
                    #End If
                    okayToParseColors = False
                End If
                
            End If
            
            'If we're allowed to continue parsing the file, parse each line in turn.
            ' (Normally we would pop lines like a stack, but the user may want to preserve the
            ' original palette order - in which case we need to also process the text lines in
            ' their original file order.)
            Dim i As Long
            For i = 3 To fileLines.GetNumOfStrings - 1
                
                curLine = Trim$(LCase$(fileLines.GetString(i)))
                
                'Comment lines start with a ";" - these can be completely ignored
                If Strings.StringsEqual(Left$(curLine, 1), ";", False) Then
                    'Do nothing
                
                'Any other line should be treated as an ARGB hex entry.  There is no prepended
                ' #, 0x, or other indicator.
                ElseIf (Len(curLine) = 8) Then
                    
                    On Error GoTo BadLineColor
                    
                    'Fill the RGBA entries directly
                    With m_PaletteColors(m_PaletteColorCount)
                        .Alpha = CByte("&h" & Left$(curLine, 2))
                        .Red = CByte("&h" & Mid$(curLine, 3, 2))
                        .Green = CByte("&h" & Mid$(curLine, 5, 2))
                        .Blue = CByte("&h" & Right$(curLine, 2))
                    End With
                    
                    On Error GoTo 0
                    
                    m_PaletteColorCount = m_PaletteColorCount + 1
                    
BadLineColor:
                
                End If
            
            Next i
            
            'If we haven't errored out, consider this a valid parse if we found any valid colors
            LoadPalettePaintDotNet = (m_PaletteColorCount > 0)
            
            'If the load was successful, assign a fake palette name using the filename
            If LoadPalettePaintDotNet Then
                m_PaletteName = Files.FileGetName(srcFile, True)
                m_PaletteFilename = srcFile
            End If
            
            'On valid loads, check and remove any duplicate palette entries.  (As an ASCII format,
            ' Paint.NET makes no guarantees, but we don't want to waste time palette-matching
            ' duplicate entries.)
            FindAndRemoveDuplicates
            
            'If the user plans to use this palette for color-matching (and really, what else are they gonna
            ' do with it?) we can pre-sort the palette to improve matching performance.
            If optimizeForPerformance Then SortByLuminance
            
        Else
            LoadPalettePaintDotNet = False
        End If
        
    End If
    
    Exit Function
    
BadPaintDotNetPalette:
    m_PaletteName = vbNullString
    m_PaletteColorCount = 0
    m_PaletteFilename = vbNullString
    LoadPalettePaintDotNet = False

End Function

'Given a valid path to a Photoshop-format .act file (Adobe Color Table), return an array of RGBQuad entries
Private Function LoadPaletteAdobeColorTable(ByRef srcFile As String, Optional ByVal optimizeForPerformance As Boolean = False) As Boolean
    
    On Error GoTo BadACTData
    
    'Because ACT files don't have a header, we must test byte length.  They are only allowed to be
    ' 768 or 772 bytes, per the spec at http://www.adobe.com/devnet-apps/photoshop/fileformatashtml/#50577411_pgfId-1070626
    If (Files.FileLenW(srcFile) = 768) Or (Files.FileLenW(srcFile) = 772) Then
    
        'For now, also test file extension (as a failsafe)
        If Strings.StringsEqual(Files.FileGetExtension(srcFile), "act", True) Then
        
            'File size and extension match.  Start retrieving colors!
            Dim srcBytes() As Byte
            If Files.FileLoadAsByteArray(srcFile, srcBytes) Then
                
                'Prep our internal palette structs
                ReDim m_PaletteColors(0 To 255) As RGBQuad
                m_PaletteColorCount = 0
                
                Dim i As Long
                For i = 0 To 255
                    With m_PaletteColors(i)
                        .Red = srcBytes(i * 3)
                        .Green = srcBytes(i * 3 + 1)
                        .Blue = srcBytes(i * 3 + 2)
                        .Alpha = 255
                    End With
                Next i
                
                'ACT files always contain 256 colors
                m_PaletteColorCount = 256
                
                'Duplicates must always be removed, as ACT files don't guarantee a descriptor for
                ' "how many relevant colors exist"
                FindAndRemoveDuplicates
                
                'If we haven't errored out, consider this a successful parse
                m_PaletteName = Files.FileGetName(srcFile, True)
                m_PaletteFilename = srcFile
                LoadPaletteAdobeColorTable = True
                
            End If
        
        End If
        
    End If
    
    Exit Function
    
BadACTData:
    m_PaletteName = vbNullString
    m_PaletteColorCount = 0
    m_PaletteFilename = vbNullString
    LoadPaletteAdobeColorTable = False
    
End Function

'Given a valid path to a GIMP-format .gpl file, return an array of RGBQuad entries
Private Function LoadPaletteGIMP(ByRef srcFile As String, Optional ByVal removeDuplicateColors As Boolean = True, Optional ByVal optimizeForPerformance As Boolean = False) As Boolean
    
    On Error GoTo InvalidPalette
    
    'GIMP palettes contain extra information that PD doesn't currently have a way to use.  We only parse the
    ' information we require, so you may need to look elsewhere for a full parsing.
    
    'Dump the full file into a string
    Dim rawFileString As String
    LoadPaletteGIMP = Files.FileLoadAsString(srcFile, rawFileString, True)
    
    If LoadPaletteGIMP Then
    
        'GIMP palette files always start with the text "GIMP Palette"
        If Strings.StringsEqual(Left$(rawFileString, 12), "GIMP Palette", True) Then
        
            'This appears to be a valid GIMP palette file.  Hypothetically, line order should be fixed,
            ' but we parse the file as if line order is *not* fixed.  Let me know if you encounter a file
            ' where this approach is invalid.
            
            'Prep our internal palette structs
            ReDim m_PaletteColors(0 To 255) As RGBQuad
            m_PaletteColorCount = 0
            
            'To simplify processing, split the string by lines.
            Dim fileLines As pdStringStack
            Set fileLines = New pdStringStack
            fileLines.CreateFromMultilineString rawFileString
            
            Const SPACE_CHAR As String = " "
            
            'Parse each line in turn.  (Normally we would pop lines like a stack, but the user may want
            ' to preserve the original palette order - in which case we need to also process the text
            ' lines in their original order.)
            Dim curLine As String, i As Long
            For i = 0 To fileLines.GetNumOfStrings - 1
                
                curLine = fileLines.GetString(i)
                
                'AFAIK, there is no formal GIMP spec for palette files.  As such, they can come in a variety
                ' of shapes and sizes.  One thing we want to standardize (to simplify parsing) is replacing
                ' tab chars with space chars; VB's lack of a generic "whitespace" identifier makes this choice
                ' the least of several evils.
                If (InStr(1, curLine, vbTab, vbBinaryCompare) <> 0) Then curLine = Replace$(curLine, vbTab, SPACE_CHAR, , , vbBinaryCompare)
                
                'Empty lines can be ignored
                If (LenB(Trim$(curLine)) = 0) Then
                    'Do nothing
                    
                'Comment lines start with a #; these can be completely ignored
                ElseIf Strings.StringsEqual(Left$(curLine, 1), "#", False) Then
                    'Do nothing
                
                'The palette name is stored on a line prefaced by "Name: "
                ElseIf Strings.StringsEqual(Left$(curLine, 5), "Name:", True) Then
                    m_PaletteName = Trim$(Right$(curLine, Len(curLine) - 5))
                
                'Color descriptor lines contain three numbers, separated by one or more spaces (as the columns
                ' are forcibly aligned).  Here are two examples of valid color lines:
                
                '232   0  50
                ' 26 130  38 ColorNameHere (occurs 6454)
                
                'Because of these variations in formatting, we have to search for colors in a somewhat complicated way.
                ElseIf (m_PaletteColorCount < 256) Then
                    
                    ' Start by looking for at least two spaces in the trimmed string (indicating at least three unique entries)
                    curLine = Trim$(curLine)
                    If (InStr(1, curLine, SPACE_CHAR, vbBinaryCompare) <> InStrRev(curLine, SPACE_CHAR, -1, vbBinaryCompare)) Then
                    
                        'This string contains two spaces.  Extract the first string-delimited entry.
                        Dim targetColor As String, targetColorInt As Long
                        targetColor = Left$(curLine, InStr(1, curLine, SPACE_CHAR, vbBinaryCompare) - 1)
                        
                        'Attempt to convert this to a number; if it fails, that's okay; this is some kind of invalid line
                        ' and we can ignore further parsing.
                        On Error GoTo BadLineColor
                        m_PaletteColors(m_PaletteColorCount).Red = CByte(targetColor)
                        On Error GoTo 0
                        
                        'Trim the color we've parsed out of the string, then repeat the above steps
                        curLine = Trim$(Right$(curLine, Len(curLine) - InStr(1, curLine, SPACE_CHAR, vbBinaryCompare)))
                        targetColor = Left$(curLine, InStr(1, curLine, SPACE_CHAR, vbBinaryCompare) - 1)
                        On Error GoTo BadLineColor
                        m_PaletteColors(m_PaletteColorCount).Green = CByte(targetColor)
                        On Error GoTo 0
                        
                        '...and one last time, for the blue component.  Note that the resulting string may not
                        ' have a trailing space, so we forcibly add one to simplify the parser.
                        curLine = Trim$(Right$(curLine, Len(curLine) - InStr(1, curLine, SPACE_CHAR, vbBinaryCompare))) & SPACE_CHAR
                        targetColor = Left$(curLine, InStr(1, curLine, SPACE_CHAR, vbBinaryCompare) - 1)
                        On Error GoTo BadLineColor
                        m_PaletteColors(m_PaletteColorCount).Blue = CByte(targetColor)
                        On Error GoTo 0
                        
                        'GIMP palettes do not support alpha channels.  Forcibly set a value of 255.
                        m_PaletteColors(m_PaletteColorCount).Alpha = 255
                        
                        'If we made it all the way here, this line was successfully parsed for color data.
                        ' Advance the color count tracker and resume the line parser.
                        m_PaletteColorCount = m_PaletteColorCount + 1
                        
BadLineColor:
                    
                    End If
                    
                End If
            
            Next i
            
            'If we haven't errored out, consider this a valid parse
            LoadPaletteGIMP = True
            
            'On valid loads, check and remove any duplicate palette entries.  (As an ASCII format, GIMP makes
            ' no guarantees about this, but we don't want to waste time palette-matching duplicate entries.)
            If removeDuplicateColors Then FindAndRemoveDuplicates
            
            'If the user plans to use this palette for color-matching (and really, what else are they gonna
            ' do with it?) we can pre-sort the palette to improve matching performance.
            If optimizeForPerformance Then SortByLuminance
            
        Else
            LoadPaletteGIMP = False
        End If
    
    End If
    
    If LoadPaletteGIMP Then m_PaletteFilename = srcFile
    
    Exit Function

InvalidPalette:
    
    m_PaletteName = vbNullString
    m_PaletteColorCount = 0
    m_PaletteFilename = vbNullString
    LoadPaletteGIMP = False

End Function

'Many palette files are text-based, and nothing prevents problematic situations like duplicate colors.
' (In fact, some shitty formats *cough* Paint.NET *cough* require a *fiex* number of colors, which practically
' guarantees duplicate entries for small palettes.)  After loading a palette from file, I recommend calling this
' function to remove any duplicate palette entries.  It is *not* well-optimized (potentially O(n^2)) but seeing
' as PD only supports 8-bit palettes, I have not yet found it worth the trouble to optimize further.
Private Sub FindAndRemoveDuplicates()

    'Only palettes with multiple entries need to be searched for duplicates
    If (m_PaletteColorCount <= 1) Then Exit Sub
    
    Dim needsRemoval() As Byte
    ReDim needsRemoval(0 To 255) As Byte
    
    Dim i As Long, j As Long, targetValue As RGBQuad, duplicatesFound As Boolean
    For i = 0 To m_PaletteColorCount - 2
        
        'Array accesses are slow in VB; use a local value instead
        targetValue = m_PaletteColors(i)
        
        For j = i + 1 To m_PaletteColorCount - 1
            With m_PaletteColors(j)
                If (.Red = targetValue.Red) And (.Green = targetValue.Green) And (.Blue = targetValue.Blue) And (.Alpha = targetValue.Alpha) Then
                    needsRemoval(j) = 1
                    duplicatesFound = True
                End If
            End With
        Next j
        
    Next i
    
    'Remove all flagged entries
    If duplicatesFound Then
    
        Dim numRemoved As Long
        
        For i = 1 To m_PaletteColorCount - 1
            If (needsRemoval(i) = 0) Then
                m_PaletteColors(i - numRemoved) = m_PaletteColors(i)
            Else
                numRemoved = numRemoved + 1
            End If
        Next i
        
        m_PaletteColorCount = m_PaletteColorCount - numRemoved
        ReDim Preserve m_PaletteColors(0 To m_PaletteColorCount - 1) As RGBQuad
        
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "Palette contained duplicate entries; color count reduced from " & CStr(m_PaletteColorCount + numRemoved) & " to " & CStr(m_PaletteColorCount) & " colors."
        #End If
        
    End If
    
End Sub

'When constructing palette matching trees, it is often useful to sort a palette by luminance
' (e.g. along the diagonal of the RGB cube).  This makes it easier to produce a balanced nearest-neighbor
' structure for color-matching.
Friend Sub SortByLuminance()

    'We don't actually use true "luminance" for the sort; instead, we treat RGB coordinates as literal
    ' positions within a 3D RGB cube, and we use their positions to sort along the cube's diagonal.
    ' (Note that alpha is *not* considered during the sort, by design.)
    Dim pSort() As PaletteSort
    ReDim pSort(0 To m_PaletteColorCount - 1) As PaletteSort
    
    Dim r As Long, g As Long, b As Long
    
    Dim i As Long
    For i = 0 To m_PaletteColorCount - 1
        
        With m_PaletteColors(i)
            r = .Red
            g = .Green
            b = .Blue
        End With
        
        pSort(i).pOrigIndex = i
        pSort(i).pSortCriteria = r * r + g * g + b * b
        
    Next i
    
    'We can now use the sort criteria to perform a normal quicksort
    SortPalette pSort
    
    'The pSort() array has now been sorted according to its pSortCriteria contents.  We want to rebuild
    ' our palette array in a matching order.
    Dim newPalette() As RGBQuad
    ReDim newPalette(0 To m_PaletteColorCount - 1) As RGBQuad
    For i = 0 To m_PaletteColorCount - 1
        newPalette(i) = m_PaletteColors(pSort(i).pOrigIndex)
    Next i
    
    'Release the old array in favor of the newly sorted one
    m_PaletteColors = newPalette
    
End Sub

'Use QuickSort to sort a palette.  The srcPaletteSort must be assembled by the caller, with the .pSortCriteria
' filled with a Single that represents "color order".  (Not defining this strictly allows for many different types
' of palette sorts, based on the caller's needs.)
Private Sub SortPalette(ByRef srcPaletteSort() As PaletteSort)
    SortInner srcPaletteSort, 0, UBound(srcPaletteSort)
End Sub

'Basic QuickSort function.  Recursive calls will sort the palette on the range [lowVal, highVal].  The first
' call must be on the range [0, UBound(srcPaletteSort)].
Private Sub SortInner(ByRef srcPaletteSort() As PaletteSort, ByVal lowVal As Long, ByVal highVal As Long)
    
    'Ignore the search request if the bounds are mismatched
    If (lowVal < highVal) Then
        
        'Sort some sub-portion of the list, and use the returned pivot to repeat the sort process
        Dim j As Long
        j = SortPartition(srcPaletteSort, lowVal, highVal)
        SortInner srcPaletteSort, lowVal, j - 1
        SortInner srcPaletteSort, j + 1, highVal
    End If
    
End Sub

'Basic QuickSort partition function.  All values in the range [lowVal, highVal] are sorted against a pivot value, j.
' The final pivot position is returned, and our caller can use that to request two new sorts on either side of the pivot.
Private Function SortPartition(ByRef srcPaletteSort() As PaletteSort, ByVal lowVal As Long, ByVal highVal As Long) As Long
    
    Dim i As Long, j As Long
    i = lowVal
    j = highVal + 1
    
    Dim v As Single
    v = srcPaletteSort(lowVal).pSortCriteria
    
    Dim tmpSort As PaletteSort
    
    Do
        
        'Compare the pivot against points beneath it
        Do
            i = i + 1
            If (i = highVal) Then Exit Do
        Loop While (srcPaletteSort(i).pSortCriteria < v)
        
        'Compare the pivot against points above it
        Do
            j = j - 1
            
            'A failsafe exit check here would be redundant, since we already check this state above
            'If (j = lowVal) Then Exit Do
        Loop While (v < srcPaletteSort(j).pSortCriteria)
        
        'If the pivot has arrived at its final location, exit
        If (i >= j) Then Exit Do
        
        'Swap the values at indexes i and j
        tmpSort = srcPaletteSort(j)
        srcPaletteSort(j) = srcPaletteSort(i)
        srcPaletteSort(i) = tmpSort
        
    Loop
    
    'Move the pivot value into its final location
    tmpSort = srcPaletteSort(j)
    srcPaletteSort(j) = srcPaletteSort(lowVal)
    srcPaletteSort(lowVal) = tmpSort
    
    'Return the pivot's final position
    SortPartition = j
    
End Function

Private Sub InternalProblem(ByRef errMsg As String)
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "WARNING!  pdPalette returned an error: " & errMsg
    #Else
        Debug.Print "WARNING!  pdPalette returned an error: " & errMsg
    #End If
End Sub
