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
'PhotoDemon Palette Container and Parser
'Copyright 2018-2026 by Tanner Helland
'Created: 16/January/18
'Last updated: 07/February/18
'Last update: add support for writing Adobe Swatch Exchange files (.ase)
'
'(This class currently has some non-obvious overlap with the Palettes module.  In the future, it would be nice to
' compartmentalize more palette-specific functionality within this class, then pass around class instances instead
' of bare RGBQuad arrays.)
'
'The pdPalette class was built to simplify management of various palette filetypes, including Photoshop, GIMP,
' PaintShop Pro, Paint.NET, and others.  It is capable of parsing palettes from each of these programs, and storing
' their relevant (and non-relevant!) data inside VB-friendly objects.
'
'Because some palette formats (like Adobe's Swatch Exchange - ASE) support the notion of multiple palettes within
' one file, this class delegates some palette responsibilities to a pdPaletteChild class.  One pdPalette instance
' may require multiple pdPaletteChild instances.  That said, memory usage is deliberately kept to a minimum,
' and pdPaletteChild provides a "remove duplicate colors" feature to improve performance even further.
'
'If you encounter other palette formats and you'd like PhotoDemon to support them, please let me know via GitHub.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Adobe swatch format requires custom color space enums
Private Enum AdobeSwatchColorSpace
    
    ASCS_RGB = 0
    ASCS_HSB = 1    'Support TODO; I need .aco files in this format for testing!
    ASCS_CMYK = 2
    ASCS_Lab = 7    'Support TODO; I need .aco files in this format for testing!
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

'If a loaded palette file contains multiple groups, this pdPalette instance will also contain multiple child classes.
' One instance (0) is always guaranteed to exist, as it's created when the class is instantiated.
Private m_NumOfGroups As Long
Private m_Palettes() As pdPaletteChild

Friend Function ChildPalette(Optional ByVal childIndex As Long = 0) As pdPaletteChild
    Set ChildPalette = m_Palettes(childIndex)
End Function

Friend Function GetPaletteColorCount(Optional ByVal palGroupIndex As Long = 0) As Long
    If (m_NumOfGroups > 0) Then GetPaletteColorCount = m_Palettes(palGroupIndex).GetNumOfColors() Else GetPaletteColorCount = 0
End Function

Friend Function GetPaletteFilename(Optional ByVal palGroupIndex As Long = 0) As String
    If (m_NumOfGroups > 0) Then GetPaletteFilename = m_Palettes(palGroupIndex).GetPaletteFilename Else GetPaletteFilename = vbNullString
End Function

Friend Function GetPaletteGroupCount() As Long
    GetPaletteGroupCount = m_NumOfGroups
End Function

Friend Function GetPaletteName(Optional ByVal palGroupIndex As Long = 0) As String
    If (m_NumOfGroups > 0) Then GetPaletteName = m_Palettes(palGroupIndex).GetPaletteName Else GetPaletteName = vbNullString
End Function

Friend Sub SetPaletteName(ByVal newPaletteName As String, Optional ByVal palGroupIndex As Long = 0)
    If (m_NumOfGroups > 0) Then m_Palettes(palGroupIndex).SetPaletteName newPaletteName
End Sub

Friend Function CopyPaletteToArray(ByRef dstPalette() As RGBQuad, Optional ByVal palGroupIndex As Long = 0) As Boolean
    
    'Because this function involves unsafe memcpy behavior, we try to perform comprehensive validation
    Dim okayToCopy As Boolean
    okayToCopy = (palGroupIndex >= 0) And (palGroupIndex <= UBound(m_Palettes)) And (m_NumOfGroups > 0)
    If okayToCopy Then okayToCopy = (Not m_Palettes(palGroupIndex) Is Nothing)
    If okayToCopy Then okayToCopy = (m_Palettes(palGroupIndex).GetNumOfColors > 0)
    If okayToCopy Then CopyPaletteToArray = m_Palettes(palGroupIndex).CopyRGBQuadsToArray(dstPalette)
    
End Function

Friend Function CreateFromPaletteArray(ByRef srcPalette() As RGBQuad, ByVal numColors As Long, Optional ByVal resetBeforeLoading As Boolean = True) As Boolean
    
    CreateFromPaletteArray = False
    
    If (numColors > 0) Then
        
        If resetBeforeLoading Then
            Me.Reset
        Else
            If (m_NumOfGroups <= UBound(m_Palettes)) Then ReDim Preserve m_Palettes(0 To m_NumOfGroups + 1) As pdPaletteChild
            If (m_Palettes(m_NumOfGroups) Is Nothing) Then Set m_Palettes(m_NumOfGroups) = New pdPaletteChild
        End If
        
        Dim i As Long
        For i = 0 To numColors - 1
            m_Palettes(m_NumOfGroups).AddColor srcPalette(i)
        Next i
        
        m_Palettes(m_NumOfGroups).SetPaletteFilename vbNullString
        m_Palettes(m_NumOfGroups).SetPaletteName "internal"
        m_NumOfGroups = m_NumOfGroups + 1
        
        CreateFromPaletteArray = True
        
    Else
        InternalProblem "WARNING!  pdPalette.CreateFromPaletteArray() was psased an invalid number of colors (" & numColors & ")"
    End If

End Function

'Want to merge two separate pdPalette objects?  This will take care of it for you, regardless of the palette group count in
' either object.
Friend Function AppendExistingPalette(ByRef srcPalette As pdPalette) As Boolean
    
    AppendExistingPalette = False
    
    On Error GoTo AppendFailure
    
    Dim okToCopy As Boolean
    okToCopy = (Not srcPalette Is Nothing)
    If okToCopy Then okToCopy = (srcPalette.GetPaletteGroupCount() > 0)
    
    If okToCopy Then
        
        'Loop through all source palettes, and add them one-at-a-time to this one
        Dim i As Long
        For i = 0 To srcPalette.GetPaletteGroupCount() - 1
            If (m_NumOfGroups >= UBound(m_Palettes)) Then ReDim Preserve m_Palettes(0 To m_NumOfGroups + 1) As pdPaletteChild
            Set m_Palettes(m_NumOfGroups) = srcPalette.ChildPalette(i)
            m_NumOfGroups = m_NumOfGroups + 1
        Next i
        
        AppendExistingPalette = True
        
    Else
        InternalProblem "WARNING!  pdPalette.AppendExistingPalette() was passed a null source palette."
    End If
    
    Exit Function

AppendFailure:

    AppendExistingPalette = False
    InternalProblem "WARNING!  pdPalette.AppendExistingPalette() experienced error #" & Err.Number & ": " & Err.Description

End Function

Friend Sub SetNewPaletteCount(Optional ByVal newPaletteSize As Long = 256, Optional ByVal palGroupIndex As Long = 0)
    Dim okayToProceed As Boolean
    okayToProceed = (palGroupIndex >= 0) And (palGroupIndex <= UBound(m_Palettes)) And (m_NumOfGroups > 0)
    If okayToProceed Then okayToProceed = (Not m_Palettes(palGroupIndex) Is Nothing)
    If okayToProceed Then m_Palettes(palGroupIndex).SetNumOfColors newPaletteSize
End Sub

'Given a path to a supported palette file, return TRUE if the file can be successfully parsed for palette data; FALSE otherwise
Friend Function LoadPaletteFromFile(ByRef srcFile As String, Optional ByVal removeDuplicateColors As Boolean = True, Optional ByVal retainOriginalColorOrder As Boolean = True, Optional ByVal resetBeforeLoading As Boolean = True) As Boolean
    
    'Ensure we have an empty palette available for loading
    If resetBeforeLoading Or (m_NumOfGroups = 0) Then
        Me.Reset
    Else
        If (m_NumOfGroups >= UBound(m_Palettes)) Then ReDim Preserve m_Palettes(0 To m_NumOfGroups + 1) As pdPaletteChild
        If (m_Palettes(m_NumOfGroups) Is Nothing) Then Set m_Palettes(m_NumOfGroups) = New pdPaletteChild Else m_Palettes(m_NumOfGroups).Reset
    End If
    
    'Track our current group count; we'll reset to this value if a load operation fails
    Dim startIndex As Long
    startIndex = m_NumOfGroups
    
    'Fail on some obvious problem cases
    If (LenB(srcFile) = 0) Then Exit Function
    If (Not Files.FileExists(srcFile)) Then Exit Function
    
    'All Load_XYZ functions operate on a pdStream object.  Load the target file into a pdStream, and then pass *that*
    ' to subsequent load functions.
    Dim cStream As pdStream
    Set cStream = New pdStream
    If cStream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFile) Then
        
        'Although we parse files using pdStream, note that we still pass the source filename (and path) to all
        ' load functions.  Many palette formats are primitive, and they may require validation against things like
        ' a precise file size - and this is easier to do via filename than stream.
        
        'Branch according to format; on failure, attempt the next mechanism.
        ' (Note that we manually reset our palette collection between attempts, in case a load routine failed
        '  mid-way through a load attempt.)
        
        'Look for PD's internal palette format
        If (Not LoadPaletteFromFile) Then
            m_NumOfGroups = startIndex
            m_Palettes(m_NumOfGroups).Reset
            LoadPaletteFromFile = LoadPalettePhotoDemon(srcFile, cStream, removeDuplicateColors, retainOriginalColorOrder)
            If LoadPaletteFromFile Then cStream.StopStream True
        End If
        
        'Look for Adobe's basic color table format
        If (Not LoadPaletteFromFile) Then
            m_NumOfGroups = startIndex
            m_Palettes(m_NumOfGroups).Reset
            LoadPaletteFromFile = LoadPaletteAdobeColorTable(srcFile, cStream, removeDuplicateColors, retainOriginalColorOrder)
            If LoadPaletteFromFile Then cStream.StopStream True
        End If
        
        'Look for Adobe color swatches next
        If (Not LoadPaletteFromFile) Then
            m_NumOfGroups = startIndex
            m_Palettes(m_NumOfGroups).Reset
            LoadPaletteFromFile = LoadPaletteAdobeSwatch(srcFile, cStream, removeDuplicateColors, retainOriginalColorOrder)
            If LoadPaletteFromFile Then cStream.StopStream True
        End If
        
        'Look for the new-ish Adobe swatch exchange fromat next
        If (Not LoadPaletteFromFile) Then
            m_NumOfGroups = startIndex
            m_Palettes(m_NumOfGroups).Reset
            LoadPaletteFromFile = LoadPaletteAdobeSwatchExchange(srcFile, cStream, removeDuplicateColors, retainOriginalColorOrder)
            If LoadPaletteFromFile Then cStream.StopStream True
        End If
        
        'Look for GIMP palettes next.  Note that GIMP palettes are supposed to have the .gpl extension,
        ' but for whatever reason, some users stick to .txt.  As such, this function will attempt to parse
        ' *anything* it's handed.  It may throw up debug messages about string conversion errors if the
        ' source file is binary, but these issues are handled safely.
        If (Not LoadPaletteFromFile) Then
            m_NumOfGroups = startIndex
            m_Palettes(m_NumOfGroups).Reset
            LoadPaletteFromFile = LoadPaletteGIMP(srcFile, cStream, removeDuplicateColors, retainOriginalColorOrder)
            If LoadPaletteFromFile Then cStream.StopStream True
        End If
        
        'Look for PaintShop Pro palettes next
        If (Not LoadPaletteFromFile) Then
            m_NumOfGroups = startIndex
            m_Palettes(m_NumOfGroups).Reset
            LoadPaletteFromFile = LoadPalettePaintShopPro(srcFile, cStream, removeDuplicateColors, retainOriginalColorOrder)
            If LoadPaletteFromFile Then cStream.StopStream True
        End If
        
        'Look for Paint.NET palettes next
        If (Not LoadPaletteFromFile) And (Strings.StringsEqual(Files.FileGetExtension(srcFile), "txt", True)) Then
            m_NumOfGroups = startIndex
            m_Palettes(m_NumOfGroups).Reset
            LoadPaletteFromFile = LoadPalettePaintDotNet(srcFile, cStream, removeDuplicateColors, retainOriginalColorOrder)
            If LoadPaletteFromFile Then cStream.StopStream True
        End If
        
    End If
    
    'If all palette parsers failed, reset our internals
    If (Not LoadPaletteFromFile) Then Me.Reset
    
End Function

'Given a valid path to a Photoshop-format .act file (Adobe Color Table), return an array of RGBQuad entries
Private Function LoadPaletteAdobeColorTable(ByRef srcFile As String, ByRef srcStream As pdStream, Optional ByVal removeDuplicateColors As Boolean = True, Optional ByVal retainOriginalColorOrder As Boolean = True) As Boolean
    
    On Error GoTo BadACTData
    
    If (srcStream Is Nothing) Then Exit Function
    
    'Reset the stream pointer and perform some basic validation checks
    srcStream.SetPosition 0, FILE_BEGIN
    Dim passedValidation As Boolean: passedValidation = True
    
    'If this palette came from file, test file extension (as a failsafe)
    If (LenB(srcFile) <> 0) Then passedValidation = Strings.StringsEqual(Files.FileGetExtension(srcFile), "act", True)
    
    'Because ACT files don't have a header, we must test byte length.  They are only allowed to be
    ' 768 or 772 bytes, per the spec at http://www.adobe.com/devnet-apps/photoshop/fileformatashtml/#50577411_pgfId-1070626
    If passedValidation Then passedValidation = (srcStream.GetStreamSize() = 768) Or (srcStream.GetStreamSize() = 772)
        
    'File size and/or extension match closely enough to attempt a load.  Start retrieving colors!
    If passedValidation Then
        
        Dim srcBytes() As Byte
        If (srcStream.ReadBytes(srcBytes, srcStream.GetStreamSize(), True) = srcStream.GetStreamSize()) Then
        
            Dim i As Long, tmpQuad As RGBQuad
            For i = 0 To 255
                With tmpQuad
                    .Red = srcBytes(i * 3)
                    .Green = srcBytes(i * 3 + 1)
                    .Blue = srcBytes(i * 3 + 2)
                    .Alpha = 255
                End With
                m_Palettes(m_NumOfGroups).AddColor tmpQuad
            Next i
            
            'In a 772-byte file, the last four bytes supposedly use two bytes for the number of colors to use,
            ' and another two bytes for a transparent index.  In my testing, these last four bytes are frequently
            ' all set to zero, which makes the "number of colors to use" indicator useless. As such, we don't
            ' currently use bytes; instead, we manually remove duplicate entries from the final palette, which is
            ' a "good enough" way to figure out how many intentional colors are in the palette.
            
            'If original color order is not required, we can sort the palette prior to removing
            ' duplicate palette entries; this makes the process much faster.
            If removeDuplicateColors Then
                If retainOriginalColorOrder Then
                    m_Palettes(m_NumOfGroups).FindAndRemoveDuplicates
                Else
                    m_Palettes(m_NumOfGroups).SortFixedOrder
                    m_Palettes(m_NumOfGroups).FindAndRemoveDuplicates_Fast
                End If
            End If
            
            'If we haven't errored out, consider this a successful parse
            m_Palettes(m_NumOfGroups).SetPaletteName Files.FileGetName(srcFile, True)
            m_Palettes(m_NumOfGroups).SetPaletteFilename srcFile
            m_NumOfGroups = m_NumOfGroups + 1
            LoadPaletteAdobeColorTable = True
            
        End If
        
    End If
    
    Exit Function
    
BadACTData:
    LoadPaletteAdobeColorTable = False
    
End Function

'Given a valid path to a Photoshop-format .aco file (Adobe Photoshop Swatch), return an array of RGBQuad entries
Private Function LoadPaletteAdobeSwatch(ByRef srcFile As String, ByRef srcStream As pdStream, Optional ByVal removeDuplicateColors As Boolean = True, Optional ByVal retainOriginalColorOrder As Boolean = True) As Boolean
    
    On Error GoTo BadSwatchData
    
    If (srcStream Is Nothing) Then Exit Function
    
    'Reset the stream pointer and perform some basic validation checks
    srcStream.SetPosition 0, FILE_BEGIN
    Dim passedValidation As Boolean: passedValidation = True
    
    'ACO files contain no validation entries in their header, so we use the file extension as initial validation.
    ' (We also check file size to ensure that the file is large enough to contain at least one color entry.
    ' The "magic number" file size is calculated using data from the spec.)
    If (LenB(srcFile) <> 0) Then passedValidation = (Strings.StringsEqual(Files.FileGetExtension(srcFile), "aco", True) And (Files.FileLenW(srcFile) >= 14))
    
    'This function references the Adobe spec frequently, particularly for things like magic numbers.  Here is the
    ' link I used (good as of Jan '18): http://www.adobe.com/devnet-apps/photoshop/fileformatashtml/#50577411_31265
    If passedValidation Then
        
        'Owing to their Mac heritage, Adobe files store everything as big-endian.  As such, we'll be using
        ' _BE suffixed retrieval functions from the pdStream object.
        
        'Note also that per the spec, swatch files are just long lists of 16-bit words.  This makes
        ' traversal straightforward.
        
        'Start by ensuring that the initial version number is good.  It needs to be 1 or 2; I've seen 3rd-parties
        ' state that "0" is also an option, but the official Adobe spec doesn't allow this, and I haven't encountered
        ' any palettes "in the wild" that do this, so we require "1" or "2" for now.
        Dim curSwatchVersion As Integer
        curSwatchVersion = srcStream.ReadInt_BE()
        If (curSwatchVersion <= 0) Or (curSwatchVersion > 2) Then InternalProblem "ACO file reported bad version number (" & CStr(curSwatchVersion) & ")"
        
        Dim dualSwatchFile As Boolean
            
        'Now comes something kinda strange.  The spec says that Version 2 swatches should always include a full
        ' Version 1 swatch, followed by an identical Version 2 swatch (plus the features Version 2 allows).
        ' This means that files *should* always start with a version 1 swatch, but as usual, there are swatch
        ' writers that ignore this and just dump a version 2 swatch into the file.  To cover both cases,
        ' we loop our check.
        Do While (curSwatchVersion = 1) Or (curSwatchVersion = 2)
            
            'Read the number of colors stored in the file
            Dim numOfColors As Long
            numOfColors = srcStream.ReadIntUnsigned_BE()
            If (numOfColors > 0) Then
                
                'Because this may be a combo swatch (with a v2 swatch following a v1 version), we always want
                ' to manually reset the current palette object before loading colors to it.  Otherwise, we risk
                ' ending up with two of every color!
                If dualSwatchFile Then Set m_Palettes(m_NumOfGroups) = New pdPaletteChild
                
                Dim i As Long, j As Long
                For i = 0 To numOfColors - 1
                    
                    'We now want to retrieve each color in turn.  Color swatches use a standardized 10-byte format,
                    ' comprised of five unsigned 16-bit entries:
                    ' 1) Color space (an internal Adobe enum)
                    ' 2) w, x, y, z values, which must be post-processed differently depending on (1)
                        
                    'Start by retrieving the color space
                    Dim curSpace As AdobeSwatchColorSpace
                    curSpace = srcStream.ReadInt_BE()
                    
                    'A dummy int can be used to advance the file pointer
                    Dim dummyInt As Integer, tmpQuad As RGBQuad, colorGood As Boolean
                    
                    colorGood = False
                    
                    Select Case curSpace
                    
                        Case ASCS_RGB
                            
                            'RGB constants are stored as 16-bit unsigned ints, plus an empty two bytes at the end.
                            ' (PD downconverts the 16-bit ints to 8-bit.)
                            Dim r As Long, g As Long, b As Long
                            r = srcStream.ReadIntUnsigned_BE()
                            g = srcStream.ReadIntUnsigned_BE()
                            b = srcStream.ReadIntUnsigned_BE()
                            dummyInt = srcStream.ReadInt()
                            
                            'Downconvert the 16-bit entries to 8-bit
                            r = (r \ 257)
                            g = (g \ 257)
                            b = (b \ 257)
                            
                            'Failsafe checks
                            If (r > 255) Then r = 255
                            If (g > 255) Then g = 255
                            If (b > 255) Then b = 255
                            
                            With tmpQuad
                                .Red = r
                                .Green = g
                                .Blue = b
                                .Alpha = 255
                            End With
                            
                            colorGood = True
                            
                        Case ASCS_CMYK
                            Dim c As Single, m As Single, y As Single, k As Single
                            c = srcStream.ReadIntUnsigned_BE()
                            m = srcStream.ReadIntUnsigned_BE()
                            y = srcStream.ReadIntUnsigned_BE()
                            k = srcStream.ReadIntUnsigned_BE()
                            
                            'Downconvert the 16-bit entries to 8-bit and invert
                            c = 255! - (c / 257!)
                            m = 255! - (m / 257!)
                            y = 255! - (y / 257!)
                            k = 255! - (k / 257!)
                            
                            'Convert these to RGB literals using an entirely different spec as our guide
                            ' (https://web.archive.org/web/20090909133852/http://partners.adobe.com/public/developer/en/ps/psrefman.pdf)
                            r = 255 - PDMath.Min2Int(255, Int(c + k + 0.5!))
                            g = 255 - PDMath.Min2Int(255, Int(m + k + 0.5!))
                            b = 255 - PDMath.Min2Int(255, Int(y + k + 0.5!))
                            
                            'Failsafe checks
                            If (r > 255) Then r = 255 Else If (r < 0) Then r = 0
                            If (g > 255) Then g = 255 Else If (g < 0) Then g = 0
                            If (b > 255) Then b = 255 Else If (b < 0) Then b = 0
                            
                            With tmpQuad
                                .Red = r
                                .Green = g
                                .Blue = b
                                .Alpha = 255
                            End With
                            
                            colorGood = True
                        
                        'Grayscale is inexplicably stored on the range [0, 10000]
                        Case ASCS_Gray
                        
                            g = srcStream.ReadIntUnsigned_BE()
                            dummyInt = srcStream.ReadInt()
                            dummyInt = srcStream.ReadInt()
                            dummyInt = srcStream.ReadInt()
                        
                            'Translate gray into a usable RGB space
                            g = CLng(CStr(g) / 39.0625)
                            
                            'Failsafe checks
                            If (g > 255) Then g = 255
                            
                            With tmpQuad
                                .Red = g
                                .Green = g
                                .Blue = g
                                .Alpha = 255
                            End With
                            
                            colorGood = True
                            
                        'These two spaces could be implemented without too much trouble, but I don't currently have
                        ' swatches available to test their correctness:
                        'Case ASCS_HSB
                        'Case ASCS_Lab
                        
                        'All other color spaces are undocumented and unsupported
                        Case Else
                            InternalProblem "Adobe Swatch color space is unsupported (" & CStr(curSpace) & ")"
                            
                            'Advance the pointer to the next color
                            srcStream.SetPosition 8, FILE_CURRENT
                    
                    End Select
                        
                    'If this is a v2 swatch, the color definition is followed by a color name.  Names are actually
                    ' very similar to BSTRs: prefaced by a 4-byte length field, representing the number of characters
                    ' in the string (not bytes), followed by the string of Unicode values, two bytes per character.
                    ' (Again, these are big-endian however, so we need to manually reverse all bytes as we go.)
                    If (curSwatchVersion = 2) Then
                        
                        'Retrieve the color size
                        Dim colorLen As Long, srcColorName As String
                        colorLen = srcStream.ReadLong_BE()
                        
                        'If the length is non-zero, resize the string to match
                        If (colorLen > 0) Then
                        
                            srcColorName = String$(colorLen, 0)
                            
                            'Retrieve each character in turn, then add it to the string
                            Dim tmpChar As Long
                            For j = 1 To colorLen
                                tmpChar = srcStream.ReadIntUnsigned_BE()
                                Mid$(srcColorName, j, 1) = ChrW$(tmpChar)
                            Next j
                            
                            srcColorName = Strings.TrimNull(srcColorName)
                            
                        Else
                            srcColorName = vbNullString
                        End If
                        
                        'If we were able to retrieve RGBA components, add both this color and its name
                        ' to the running collection.
                        If colorGood Then m_Palettes(m_NumOfGroups).AddColor tmpQuad, srcColorName
                        
                    'On v1 swatches, color names are not supported; just add the RGBA components
                    Else
                        If colorGood Then m_Palettes(m_NumOfGroups).AddColor tmpQuad
                    End If
                    
                Next i
                
                'Make sure at least one valid color was found
                If (m_Palettes(m_NumOfGroups).GetNumOfColors() > 0) Then
                    
                    'If original color order is not required, we can sort the palette prior to removing
                    ' duplicate palette entries; this makes the process much faster.
                    If removeDuplicateColors Then
                        If retainOriginalColorOrder Then
                            m_Palettes(m_NumOfGroups).FindAndRemoveDuplicates
                        Else
                            m_Palettes(m_NumOfGroups).SortFixedOrder
                            m_Palettes(m_NumOfGroups).FindAndRemoveDuplicates_Fast
                        End If
                    End If
                    
                    'If we haven't errored out, consider this a successful parse
                    m_Palettes(m_NumOfGroups).SetPaletteName Files.FileGetName(srcFile, True)
                    m_Palettes(m_NumOfGroups).SetPaletteFilename srcFile
                    m_NumOfGroups = m_NumOfGroups + 1
                    LoadPaletteAdobeSwatch = True
                    
                End If
            
            Else
                InternalProblem "ACO file reported empty palette (" & CStr(numOfColors) & ")"
            End If
            
            'If we haven't reached the end of the file, this is a hybrid swatch (where a v2 swatch follows the
            ' existing v1 version).  If the version 2 swatch validates, let's erase our existing swatch data and
            ' use the v2 entries instead.  (They *should* be identical, but better safe than sorry.)
            If (srcStream.GetPosition < srcStream.GetStreamSize) Then
                curSwatchVersion = srcStream.ReadInt_BE()
                dualSwatchFile = True
                m_NumOfGroups = m_NumOfGroups - 1
            Else
                curSwatchVersion = 0
                dualSwatchFile = False
            End If
            
        'If a v2 swatch follows the v1 swatch, retrieve it next
        Loop
        
    End If
    
    Exit Function
    
BadSwatchData:
    InternalProblem "LoadPaletteAdobeSwatch error # " & CStr(Err.Number) & ": " & Err.Description
    LoadPaletteAdobeSwatch = False
    
End Function

'Given a valid path to a modern ASE file (Adobe Swatch Exchange), extract and process all swatches inside the file.
Private Function LoadPaletteAdobeSwatchExchange(ByRef srcFile As String, ByRef srcStream As pdStream, Optional ByVal removeDuplicateColors As Boolean = True, Optional ByVal retainOriginalColorOrder As Boolean = True) As Boolean
    
    On Error GoTo BadSwatchData
    
    If (srcStream Is Nothing) Then Exit Function
    
    'As the swatch exchange loader may load multiple palettes (ASE files can contain embedded "groups"), we need to make
    ' a note of our initial group index; this will serve as our base for loading the file
    Dim baseIndex As Long
    baseIndex = m_NumOfGroups
    
    'Reset the stream pointer and perform some basic validation checks
    srcStream.SetPosition 0, FILE_BEGIN
    Dim passedValidation As Boolean: passedValidation = True
    
    'ASE files contain a proper validation entry in their header, so we'll start by checking the file extension,
    ' then proceed with a more detailed validation.  (Note that we also double-check file length, to ensure that a
    ' full header exists in the file.)
    If (LenB(srcFile) <> 0) Then passedValidation = Strings.StringsEqual(Files.FileGetExtension(srcFile), "ase", True)
    If passedValidation Then passedValidation = (srcStream.GetStreamSize() >= 12)
    
    'Adobe has not publicly documented the ASE format, which means we have to rely on reverse-engineering to decode
    ' the format.  Several links proved very useful in this regard; thank you to each of them:
    ' http://www.selapa.net/swatches/colors/fileformats.php#adobe_ase
    ' https://www.cyotek.com/blog/reading-adobe-swatch-exchange-ase-files-using-csharp
    If passedValidation Then
        
        'Owing to their Mac heritage, Adobe files store everything as big-endian.  As such, we'll be using
        ' many _BE suffixed retrieval functions from the pdStream object.
        
        'Start by checking the first four chars for the required "ASEF" descriptor.
        Dim headerValidated As Boolean
        headerValidated = Strings.StringsEqual(srcStream.ReadString_ASCII(4), "ASEF", False)
        
        'If the header validated correctly, we also need to verify the file's version number.  Version is
        ' defined as major/minor, two 16-bit integers.  (The only supported version is 1.0.)
        If headerValidated Then headerValidated = (srcStream.ReadInt_BE() = 1)
        If headerValidated Then headerValidated = (srcStream.ReadInt_BE() = 0)
        
        'If all validation checks have passed, we have one final one to perform before proceeding.
        ' The next entry is a 4-byte integer defining the number of blocks in the file.  Blocks come in
        ' three types: group start, group end, and color entry.
        Dim numOfBlocks As Long
        If headerValidated Then
            numOfBlocks = srcStream.ReadLong_BE()
            headerValidated = (numOfBlocks > 0)
        End If
        
        'If the file contains a non-zero block count, let's parse it for color data!
        If headerValidated Then
            
            'One frustration with the ASE format is that the number of colors stored in the swatch is not
            ' reported in advance.  The only way to know this is to actually parse each block in turn.
            
            'Iterate through all blocks.  We may skip certain block types as they are not currently useful
            ' (or even decodable, in the case of 3rd-party extensions).
            Dim i As Long, j As Long
            For i = 1 To numOfBlocks
                
                'Start by retrieving block type and length.  IMPORTANTLY!  Note that block length does
                ' *not* include the six bytes required for the block type and length descriptors - it only
                ' defines the block length *following* those six bytes.
                Dim blockType As Long, blockLength As Long
                blockType = srcStream.ReadIntUnsigned_BE()
                blockLength = srcStream.ReadLong_BE()
                
                'Group end blocks have zero-length, by design.  We process them implicitly (whenever a
                ' new group is encountered or the file ends), so we can ignore them here.
                If (blockLength > 0) Then
                
                    'Because we're going to be poking around inside this block, let's make a note of the
                    ' current stream pointer position.  We can use this (plus the block length) to properly
                    ' align the stream pointer after processing this block, regardless of how much block data
                    ' we actually retrieve/use.
                    Dim blockStreamPointer As Long
                    blockStreamPointer = srcStream.GetPosition()
                    
                    Dim curGroupName As String, curColorName As String, nameLength As Long, tmpChar As Long
                    Dim numGroupsFound As Long
                    
                    'Next, split processing according to block type.
                    Select Case blockType
                        
                        'Color descriptor
                        Case 1
                        
                            'Color descriptors start with this color's name.  For additional details, see the
                            ' "Group start" case below, as it contains full details on name parsing.
                            nameLength = srcStream.ReadIntUnsigned_BE()
                            If (nameLength > 0) Then
                            
                                curColorName = String$(nameLength, 0)
                                
                                For j = 1 To nameLength - 1
                                    tmpChar = srcStream.ReadIntUnsigned_BE()
                                    Mid$(curColorName, j, 1) = ChrW$(tmpChar)
                                Next j
                                
                                curColorName = Trim$(Strings.TrimNull(curColorName))
                                
                                'Like most Adobe formats, other software may write bad names.  Remove some obvious problem chars.
                                If InStr(1, curColorName, vbLf, vbBinaryCompare) Then curColorName = Replace$(curColorName, vbLf, vbCrLf, , , vbBinaryCompare)
                                If InStr(1, curColorName, vbCrLf, vbBinaryCompare) Then curColorName = Replace$(curColorName, vbCrLf, vbNullString, , , vbBinaryCompare)
                                
                                'Manually advance the pointer by one 16-bit word to account for the null
                                ' string terminator we didn't retrieve.
                                tmpChar = srcStream.ReadInt()
                                
                            End If
                            
                            'Next is a 4-byte ASCII string defining the current color space.  It may contain
                            ' several different values.
                            Dim colorSpaceName As String
                            colorSpaceName = UCase$(Trim$(srcStream.ReadString_ASCII(4)))
                            
                            Dim tmpQuad As RGBQuad
                            Dim r As Single, g As Single, b As Single
                            Dim c As Single, m As Single, y As Single, k As Single
                            
                            Select Case colorSpaceName
                                
                                'To handle CMYK correctly, we'd actually want to load up a proper CMYK ICC profile
                                ' and manually translate all colors via that profile.  Because Adobe's EULA only
                                ' allows its CMYK profiles to ship as standalone files (never bundled with software),
                                ' I'd need to use a public domain one, like the one helpfully provided by Graeme Gill
                                ' of Argyll CMS fame:
                                ' https://sourceforge.net/p/lcms/mailman/message/32755884/
                                ' www.argyllcms.com/cmyk.icm
                                
                                'This is the CMYK profile Krita ships, and it's likely to perform much better than
                                ' the "dummy" CMYK transform we currently use.  That said, shipping full profiles is
                                ' still on my to-do list, so until that time, this is marked as TODO.
                                Case "CMYK"
                                
                                    'CMYK values are stored as four single-precision floats.
                                    c = srcStream.ReadFloat_BE()
                                    m = srcStream.ReadFloat_BE()
                                    y = srcStream.ReadFloat_BE()
                                    k = srcStream.ReadFloat_BE()
                                    
                                    'Convert these to RGB literals using an entirely different spec as our guide
                                    ' (https://web.archive.org/web/20090909133852/http://partners.adobe.com/public/developer/en/ps/psrefman.pdf)
                                    r = 1! - PDMath.Min2Float_Single(1!, c + k)
                                    g = 1! - PDMath.Min2Float_Single(1!, m + k)
                                    b = 1! - PDMath.Min2Float_Single(1!, y + k)
                                    
                                    'Failsafe checks for HDR spaces
                                    If (r > 1!) Then r = 1! Else If (r < 0!) Then r = 0!
                                    If (g > 1!) Then g = 1! Else If (g < 0!) Then g = 0!
                                    If (b > 1!) Then b = 1! Else If (b < 0!) Then b = 0!
                                    
                                    'Rounding must be used to arrive at identical values as palettes exported via
                                    ' Photoshop to formats that use integer RGB values.
                                    With tmpQuad
                                        .Red = Int(r * 255! + 0.5!)
                                        .Green = Int(g * 255! + 0.5!)
                                        .Blue = Int(b * 255! + 0.5!)
                                        .Alpha = 255
                                    End With
                                    
                                    m_Palettes(m_NumOfGroups).AddColor tmpQuad, curColorName
                                    
                                Case "RGB"
                                
                                    'To make our life harder, RGB values are stored as three single-precision floats.
                                    ' Unfortunately, the floats are stored as big-endian values (ugh).  So don't
                                    ' forget to convert them!
                                    r = srcStream.ReadFloat_BE()
                                    g = srcStream.ReadFloat_BE()
                                    b = srcStream.ReadFloat_BE()
                                    
                                    'Failsafe checks for HDR spaces
                                    If (r > 1!) Then r = 1! Else If (r < 0!) Then r = 0!
                                    If (g > 1!) Then g = 1! Else If (g < 0!) Then g = 0!
                                    If (b > 1!) Then b = 1! Else If (b < 0!) Then b = 0!
                                    
                                    'Rounding must be used to arrive at identical values as palettes exported via
                                    ' Photoshop to formats that use integer RGB values.
                                    With tmpQuad
                                        .Red = Int(r * 255! + 0.5!)
                                        .Green = Int(g * 255! + 0.5!)
                                        .Blue = Int(b * 255! + 0.5!)
                                        .Alpha = 255
                                    End With
                                    
                                    m_Palettes(m_NumOfGroups).AddColor tmpQuad, curColorName
                                    
                                'I need LAB swatches for testing!
                                Case "LAB"
                                
                                Case "Gray"
                                
                                    'Gray is identical to RGB, but (obviously) only one value is stored.
                                    g = srcStream.ReadFloat_BE()
                                    If (g > 1!) Then g = 1! Else If (g < 0!) Then g = 0!
                                    
                                    With tmpQuad
                                        .Red = Int(g * 255! + 0.5!)
                                        .Green = .Red
                                        .Blue = .Red
                                        .Alpha = 255
                                    End With
                                    
                                    m_Palettes(m_NumOfGroups).AddColor tmpQuad, curColorName
                            
                            End Select
                            
                            '/End color block parsing
                            
                        'Group start.  This block is *optional* and not provided by all swatches!
                        Case &HC001&
                            
                            'If this is *not* the first group we've found in this file, we need to manually
                            ' create a new palette instance here.
                            If (numGroupsFound > 0) Then
                                m_NumOfGroups = m_NumOfGroups + 1
                                ReDim Preserve m_Palettes(0 To m_NumOfGroups) As pdPaletteChild
                                Set m_Palettes(m_NumOfGroups) = New pdPaletteChild
                            End If
                            
                            'Start by retrieving the length of the name, *in WCHARS*, not bytes.  Note also that
                            ' the length *includes* a null terminator, which we don't need to process.
                            nameLength = srcStream.ReadIntUnsigned_BE()
                            
                            'Next, retrieve the string, and manually remove 1 from its length (as VB will add
                            ' its own null-pointer at the end of the BSTR).  Note also that we can't rely on
                            ' convenient WAPI functions here, because the string chars are big-endian.
                            If (nameLength > 0) Then
                            
                                curGroupName = String$(nameLength, 0)
                                
                                'Retrieve each character in turn, then add it to the string
                                For j = 1 To nameLength - 1
                                    tmpChar = srcStream.ReadIntUnsigned_BE()
                                    Mid$(curGroupName, j, 1) = ChrW$(tmpChar)
                                Next j
                                
                            End If
                            
                            'Note that we don't need to synchronize the stream pointer here, because we forcibly
                            ' reset it after processing this block.
                            
                            'Assign this name to the current group.
                            m_Palettes(m_NumOfGroups).SetPaletteName Trim$(Strings.TrimNull(curGroupName))
                            
                            'Note that we've found a group in this file
                            numGroupsFound = numGroupsFound + 1
                        
                        'Group end; this has already been handled, above (because group end blocks are zero-length)
                        Case &HC002&
                            
                            'Make sure that the currently-loaded group contains at least one valid color.  If it doesn't,
                            ' we want to reset this group and reuse it.
                            If (m_Palettes(m_NumOfGroups).GetNumOfColors <= 0) Then
                                m_Palettes(m_NumOfGroups).Reset
                                m_NumOfGroups = m_NumOfGroups - 1
                            End If
                        
                        'Some other block type.  This is not strictly prohibited by the spec, and apparently some
                        ' ASE writers use custom blocks to store metadata.  We have no way to understand these,
                        ' so just ignore 'em.
                        Case Else
                    
                    End Select
                    
                    'Because different block types are processed differently, and some may not manually read to
                    ' the end of a block, forcibly reset the stream pointer using known-good values.
                    srcStream.SetPosition blockStreamPointer, FILE_BEGIN
                    srcStream.SetPosition blockLength, FILE_CURRENT
                
                End If
                
            Next i
            
            'Make sure at least one valid palette was loaded
            If (numGroupsFound > baseIndex) Then
                
                'Unlike other palette files, we now need to loop through all child palettes (as there may be
                ' multiple ones) and assign proper values to each!
                m_NumOfGroups = m_NumOfGroups + 1
                
                For i = baseIndex To m_NumOfGroups - 1
                    
                    If (Not m_Palettes(i) Is Nothing) Then
                    
                        m_Palettes(i).SetPaletteFilename srcFile
                        
                        'If original color order is not required, we can sort the palette prior to removing
                        ' duplicate palette entries; this makes the process much faster.
                        If removeDuplicateColors Then
                            If retainOriginalColorOrder Then
                                m_Palettes(i).FindAndRemoveDuplicates
                            Else
                                m_Palettes(i).SortFixedOrder
                                m_Palettes(i).FindAndRemoveDuplicates_Fast
                            End If
                        End If
                        
                    End If
                    
                Next i
                    
                'If this swatch file contained only one group, and that group didn't have a name,
                ' substitute the filename as the palette name.
                If ((m_NumOfGroups - baseIndex) = 1) Then
                    If (LenB(m_Palettes(m_NumOfGroups - 1).GetPaletteName) = 0) Then m_Palettes(m_NumOfGroups - 1).SetPaletteName Files.FileGetName(srcFile, True)
                End If
                
                'If we haven't errored out, consider this a successful parse
                LoadPaletteAdobeSwatchExchange = True
                
            Else
                InternalProblem "LoadPaletteAdobeSwatchExchange() reports no valid colors found.  This palette may use an unsupported color space."
                LoadPaletteAdobeSwatchExchange = False
            End If
            
        'Header did not validate
        Else
            InternalProblem "LoadPaletteAdobeSwatchExchange() reports that this file failed basic validation."
            LoadPaletteAdobeSwatchExchange = False
        End If
        
    'Bad file extension; for security reasons, do not attempt to load.
    End If
    
    Exit Function
    
BadSwatchData:
    InternalProblem "LoadPaletteAdobeSwatchExchange error # " & CStr(Err.Number) & ": " & Err.Description
    LoadPaletteAdobeSwatchExchange = False
    
End Function

'Given a valid path to a GIMP-format .gpl file, return an array of RGBQuad entries
Private Function LoadPaletteGIMP(ByRef srcFile As String, ByRef srcStream As pdStream, Optional ByVal removeDuplicateColors As Boolean = True, Optional ByVal retainOriginalColorOrder As Boolean = True) As Boolean
    
    On Error GoTo InvalidPalette
    
    If (srcStream Is Nothing) Then Exit Function
    
    'Because GIMP palettes are just text files, it's easier to parse them as strings
    srcStream.SetPosition 0, FILE_BEGIN
    
    Dim rawFileString As String
    rawFileString = srcStream.ReadString_UnknownEncoding(srcStream.GetStreamSize())
    
    If (LenB(rawFileString) <> 0) Then
    
        'GIMP palette files always start with the text "GIMP Palette"
        If Strings.StringsEqual(Left$(rawFileString, 12), "GIMP Palette", True) Then
        
            'This appears to be a valid GIMP palette file.  Hypothetically, line order should be fixed,
            ' but we parse the file as if line order is *not* fixed.  Let me know if you encounter a file
            ' where this approach is invalid.
            
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
                ' tab chars with space chars; VB's lack of a generic "whitespace" identifier makes this
                ' approach the least of several evils.
                If (InStr(1, curLine, vbTab, vbBinaryCompare) <> 0) Then curLine = Replace$(curLine, vbTab, SPACE_CHAR, , , vbBinaryCompare)
                
                'Empty lines can be ignored
                If (LenB(Trim$(curLine)) = 0) Then
                    'Do nothing
                    
                'Comment lines start with a #; these can be completely ignored
                ElseIf Strings.StringsEqual(Left$(curLine, 1), "#", False) Then
                    'Do nothing
                
                'The palette name is stored on a line prefaced by "Name: "
                ElseIf Strings.StringsEqual(Left$(curLine, 5), "Name:", True) Then
                    m_Palettes(m_NumOfGroups).SetPaletteName Trim$(Right$(curLine, Len(curLine) - 5))
                
                'Color descriptor lines contain three numbers, separated by one or more spaces (as the columns
                ' are forcibly aligned).  Here are two examples of valid color lines:
                
                '232   0  50
                ' 26 130  38 ColorNameHere (occurs 6454)
                
                'Because of these variations in formatting, we have to search for colors in a somewhat complicated way.
                Else
                    
                    ' Start by looking for at least two spaces in the trimmed string (indicating at least three unique entries)
                    curLine = Trim$(curLine)
                    If (InStr(1, curLine, SPACE_CHAR, vbBinaryCompare) <> InStrRev(curLine, SPACE_CHAR, -1, vbBinaryCompare)) Then
                    
                        'This string contains at least two spaces.  Extract the first string-delimited entry.
                        Dim targetColor As String, tmpQuad As RGBQuad
                        targetColor = Left$(curLine, InStr(1, curLine, SPACE_CHAR, vbBinaryCompare) - 1)
                        
                        'Attempt to convert this to a number; if it fails, that's okay; this is some kind of invalid line
                        ' and we can ignore further parsing.
                        On Error GoTo BadLineColor
                        tmpQuad.Red = CByte(targetColor)
                        On Error GoTo 0
                        
                        'Trim the color we've parsed out of the string, then repeat the above steps
                        curLine = Trim$(Right$(curLine, Len(curLine) - InStr(1, curLine, SPACE_CHAR, vbBinaryCompare)))
                        targetColor = Left$(curLine, InStr(1, curLine, SPACE_CHAR, vbBinaryCompare) - 1)
                        On Error GoTo BadLineColor
                        tmpQuad.Green = CByte(targetColor)
                        On Error GoTo 0
                        
                        '...and one last time, for the blue component.  Note that the resulting string may not
                        ' have a trailing space, so we forcibly add one to simplify the parser.
                        curLine = Trim$(Right$(curLine, Len(curLine) - InStr(1, curLine, SPACE_CHAR, vbBinaryCompare))) & SPACE_CHAR
                        targetColor = Left$(curLine, InStr(1, curLine, SPACE_CHAR, vbBinaryCompare) - 1)
                        On Error GoTo BadLineColor
                        tmpQuad.Blue = CByte(targetColor)
                        On Error GoTo 0
                        
                        'GIMP palettes do not support alpha channels.  Forcibly set a value of 255.
                        tmpQuad.Alpha = 255
                        
                        'If any text remains in this row, we can treat it as this color's name.  Color names are not
                        ' guaranteed to exist inside GIMP palettes (and they may not even be part of the spec), but many
                        ' of the sample palettes I downloaded from the Internet came with names, so we support 'em here.
                        Dim srcColorName As String
                        srcColorName = vbNullString
                        If (InStr(1, Trim$(curLine), SPACE_CHAR, vbBinaryCompare) <> 0) Then
                            curLine = Trim$(Right$(curLine, Len(curLine) - InStr(1, curLine, SPACE_CHAR, vbBinaryCompare))) & SPACE_CHAR
                            If (LenB(curLine) <> 0) Then srcColorName = curLine
                        End If
                        
                        'If we made it all the way here, this line was successfully parsed for color data.
                        ' Advance the color count tracker and resume the line parser.
                        m_Palettes(m_NumOfGroups).AddColor tmpQuad, srcColorName
                        
BadLineColor:
                    End If
                    
                End If
            
            Next i
            
            'If we haven't errored out, consider this a valid parse
            LoadPaletteGIMP = True
            
            'If original color order is not required, we can sort the palette prior to removing
            ' duplicate palette entries; this makes the process much faster.
            If removeDuplicateColors Then
                If retainOriginalColorOrder Then
                    m_Palettes(m_NumOfGroups).FindAndRemoveDuplicates
                Else
                    m_Palettes(m_NumOfGroups).SortFixedOrder
                    m_Palettes(m_NumOfGroups).FindAndRemoveDuplicates_Fast
                End If
            End If
            
        Else
            LoadPaletteGIMP = False
        End If
    
    End If
    
    If LoadPaletteGIMP Then
        m_Palettes(m_NumOfGroups).SetPaletteFilename srcFile
        m_NumOfGroups = m_NumOfGroups + 1
    End If
    
    Exit Function

InvalidPalette:
    LoadPaletteGIMP = False

End Function

'Paint.NET palettes are just text files with a list of hex chars.  Unlike other palette formats,
' they *do* support alpha values.
Private Function LoadPalettePaintDotNet(ByRef srcFile As String, ByRef srcStream As pdStream, Optional ByVal removeDuplicateColors As Boolean = True, Optional ByVal retainOriginalColorOrder As Boolean = True) As Boolean

    On Error GoTo BadPaintDotNetPalette
    
    If (srcStream Is Nothing) Then Exit Function
    
    'Paint.NET palettes don't provide an easy validation method, so if this palette came from a file, test its
    ' file extension before continuing.
    Dim failsafeCheck As Boolean: failsafeCheck = True
    If (LenB(srcFile) <> 0) Then failsafeCheck = Strings.StringsEqual(Files.FileGetExtension(srcFile), "txt", True)
    
    If failsafeCheck Then
        
        'Because Paint.NET palettes are just text files, it's easier to parse them as strings
        srcStream.SetPosition 0, FILE_BEGIN
        
        Dim rawFileString As String
        rawFileString = srcStream.ReadString_UnknownEncoding(srcStream.GetStreamSize())
        
        If (LenB(rawFileString) <> 0) Then
            
            'FYI: Paint.NET palettes are technically limited to 96 colors...
            ' (https://www.getpaint.net/doc/latest/WorkingWithPalettes.html)
            ' ...but we ignore this and simply grab as many colors as we find in the file.
            
            'To simplify processing, split the string by lines.
            Dim fileLines As pdStringStack
            Set fileLines = New pdStringStack
            fileLines.CreateFromMultilineString rawFileString
            
            ' (Normally we would pop lines like a stack, but the user may want to preserve the
            ' original palette order - in which case we need to also process the text lines in
            ' their original file order.)
            Dim i As Long, curLine As String
            For i = 0 To fileLines.GetNumOfStrings - 1
                
                curLine = Trim$(LCase$(fileLines.GetString(i)))
                
                'Comment lines start with a ";" - these can be completely ignored
                If Strings.StringsEqual(Left$(curLine, 1), ";", False) Then
                    'Do nothing
                
                'Any other line should be treated as an ARGB hex entry.  There is no prepended
                ' #, 0x, or other indicator.
                ElseIf (Len(curLine) = 8) Then
                    
                    Dim tmpQuad As RGBQuad
                    
                    On Error GoTo BadLineColor
                    
                    'Fill the RGBA entries directly
                    With tmpQuad
                        .Alpha = CByte("&h" & Left$(curLine, 2))
                        .Red = CByte("&h" & Mid$(curLine, 3, 2))
                        .Green = CByte("&h" & Mid$(curLine, 5, 2))
                        .Blue = CByte("&h" & Right$(curLine, 2))
                    End With
                    
                    On Error GoTo 0
                    
                    m_Palettes(m_NumOfGroups).AddColor tmpQuad
                    
BadLineColor:
                End If
            
            Next i
            
            'If we haven't errored out, consider this a valid parse if we found any valid colors
            LoadPalettePaintDotNet = (m_Palettes(m_NumOfGroups).GetNumOfColors() > 0)
            
            'If the load was successful, assign a fake palette name using the filename
            If LoadPalettePaintDotNet Then
            
                m_Palettes(m_NumOfGroups).SetPaletteName Files.FileGetName(srcFile, True)
                m_Palettes(m_NumOfGroups).SetPaletteFilename srcFile
                
                'If original color order is not required, we can sort the palette prior to removing
                ' duplicate palette entries; this makes the process much faster.
                If removeDuplicateColors Then
                    If retainOriginalColorOrder Then
                        m_Palettes(m_NumOfGroups).FindAndRemoveDuplicates
                    Else
                        m_Palettes(m_NumOfGroups).SortFixedOrder
                        m_Palettes(m_NumOfGroups).FindAndRemoveDuplicates_Fast
                    End If
                End If
            
                m_NumOfGroups = m_NumOfGroups + 1
                
            End If
            
        Else
            LoadPalettePaintDotNet = False
        End If
        
    End If
    
    Exit Function
    
BadPaintDotNetPalette:
    LoadPalettePaintDotNet = False

End Function
 
'PaintShop Pro palettes use multiple extensions, but all are just text files with a line-delimited list of
' RGB triples.
Private Function LoadPalettePaintShopPro(ByRef srcFile As String, ByRef srcStream As pdStream, Optional ByVal removeDuplicateColors As Boolean = True, Optional ByVal retainOriginalColorOrder As Boolean = True) As Boolean

    On Error GoTo BadPaintShopProPalette
    
    If (srcStream Is Nothing) Then Exit Function
    
    'PaintShop Pro palettes have a very basic header, and they use the standardized file extensions ".pal" (old format)
    ' and ".psppalette" (new format).  To my knowledge, the file contents themselves are identical for both formats -
    ' but as a failsafe, let's at least make sure that file extension is okay.  (Obviously, this check is only relevant
    ' if this palette is coming directly from a file.)
    Dim failsafeCheck As Boolean: failsafeCheck = True
    If (LenB(srcFile) <> 0) Then failsafeCheck = (Strings.StringsEqual(Files.FileGetExtension(srcFile), "pal", True) Or Strings.StringsEqual(Files.FileGetExtension(srcFile), "psppalette", True))
    
    If failsafeCheck Then
        
        'Because PSP palettes are just text files, it's easier to parse them as strings
        srcStream.SetPosition 0, FILE_BEGIN
        
        Dim rawFileString As String
        rawFileString = srcStream.ReadString_UnknownEncoding(srcStream.GetStreamSize())
        
        If (LenB(rawFileString) <> 0) Then
            
            'FYI: PaintShop Pro palettes are technically limited to 16 or 256 colors...
            ' (https://www.codeproject.com/articles/31964/palettes-you-ve-gotta-love-them)
            ' ...but we follow the modern convention of ignoring that number and parsing as many valid colors
            ' as we find in the file.
            
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
                If Strings.StringsNotEqual(curLine, "0100", False) Then PDDebug.LogAction "WARNING!  JASC palette contains unexpected version number: " & curLine
                
                'Finally, retrieve line 3.  This line should be some number on the range [1, 256].
                curLine = Trim$(fileLines.GetString(2))
                If (CLng(curLine) < 1 Or CLng(curLine) > 256) Then
                    PDDebug.LogAction "WARNING!  JASC palette contains invalid color count: " & curLine
                    okayToParseColors = False
                End If
                
            End If
            
            'Parse each line in turn.  (Normally we would pop lines like a stack, but the user may want
            ' to preserve the original palette order - in which case we need to also process the text
            ' lines in their original order.)
            Const SPACE_CHAR As String = " "
            Dim i As Long
            For i = 3 To fileLines.GetNumOfStrings - 1
                
                curLine = Trim$(LCase$(fileLines.GetString(i)))
                
                'Start by looking for at least two spaces in the trimmed string (indicating at least three unique entries)
                If (InStr(1, curLine, SPACE_CHAR, vbBinaryCompare) <> InStrRev(curLine, SPACE_CHAR, -1, vbBinaryCompare)) Then
                
                    'This string contains two spaces.  Extract the first string-delimited entry.
                    Dim targetColor As String, tmpQuad As RGBQuad
                    targetColor = Left$(curLine, InStr(1, curLine, SPACE_CHAR, vbBinaryCompare) - 1)
                    
                    'Attempt to convert this to a number; if it fails, that's okay; this is some kind of invalid line
                    ' and we can ignore further parsing.
                    On Error GoTo BadLineColor
                    tmpQuad.Red = CByte(targetColor)
                    On Error GoTo 0
                    
                    'Trim the color we've parsed out of the string, then repeat the above steps
                    curLine = Trim$(Right$(curLine, Len(curLine) - InStr(1, curLine, SPACE_CHAR, vbBinaryCompare)))
                    targetColor = Left$(curLine, InStr(1, curLine, SPACE_CHAR, vbBinaryCompare) - 1)
                    On Error GoTo BadLineColor
                    tmpQuad.Green = CByte(targetColor)
                    On Error GoTo 0
                    
                    '...and one last time, for the blue component.  Note that the resulting string may not
                    ' have a trailing space, so we forcibly add one to simplify the parser.
                    curLine = Trim$(Right$(curLine, Len(curLine) - InStr(1, curLine, SPACE_CHAR, vbBinaryCompare))) & SPACE_CHAR
                    targetColor = Left$(curLine, InStr(1, curLine, SPACE_CHAR, vbBinaryCompare) - 1)
                    On Error GoTo BadLineColor
                    tmpQuad.Blue = CByte(targetColor)
                    On Error GoTo 0
                    
                    'PSP palettes do not support alpha channels.  Forcibly set a value of 255.
                    tmpQuad.Alpha = 255
                    
                    'If we made it all the way here, this line was successfully parsed for color data.
                    ' Add the color to our running collection resume the line parser.
                    m_Palettes(m_NumOfGroups).AddColor tmpQuad
                    
BadLineColor:
                End If
                    
            Next i
            
            'If we haven't errored out, consider this a valid parse if we found any valid colors
            LoadPalettePaintShopPro = (m_Palettes(m_NumOfGroups).GetNumOfColors() > 0)
            
            'If the load was successful, assign a fake palette name using the filename
            If LoadPalettePaintShopPro Then
            
                m_Palettes(m_NumOfGroups).SetPaletteName Files.FileGetName(srcFile, True)
                m_Palettes(m_NumOfGroups).SetPaletteFilename srcFile
                
                'If original color order is not required, we can sort the palette prior to removing
                ' duplicate palette entries; this makes the process much faster.
                If removeDuplicateColors Then
                    If retainOriginalColorOrder Then
                        m_Palettes(m_NumOfGroups).FindAndRemoveDuplicates
                    Else
                        m_Palettes(m_NumOfGroups).SortFixedOrder
                        m_Palettes(m_NumOfGroups).FindAndRemoveDuplicates_Fast
                    End If
                End If
                
                'Finally, increment our group count
                m_NumOfGroups = m_NumOfGroups + 1
            
            End If
            
        Else
            LoadPalettePaintShopPro = False
        End If
        
    End If
    
    Exit Function
    
BadPaintShopProPalette:
   LoadPalettePaintShopPro = False

End Function

'PhotoDemon palettes are simple text files.  Unlike other palette formats, they provide full support for:
' - Unlimited color counts
' - RGBA values
' - Color names
' ...while still being human-readable text files.
'
'This format is roughly modeled after Paint.NET palette files; see the SavePalettePhotoDemon function
' for a complete description of the minor changes we've made.
Private Function LoadPalettePhotoDemon(ByRef srcFile As String, ByRef srcStream As pdStream, Optional ByVal removeDuplicateColors As Boolean = False, Optional ByVal retainOriginalColorOrder As Boolean = True) As Boolean

    On Error GoTo BadPDPalette
    
    If (srcStream Is Nothing) Then Exit Function
    
    'While we'll also validate the first line of the palette file, we may as well validate the file extension as well.
    ' (PD itself will never write palettes with a different extension, and we don't want to encourage other software
    '  doing that either.)
    Dim failsafeCheck As Boolean: failsafeCheck = True
    If (LenB(srcFile) <> 0) Then failsafeCheck = Strings.StringsEqual(Files.FileGetExtension(srcFile), "pdpalette", True)
    
    If failsafeCheck Then
        
        'Because PD palettes are just text files, it's easiest to parse them as strings
        srcStream.SetPosition 0, FILE_BEGIN
        
        Dim rawFileString As String
        rawFileString = srcStream.ReadString_UTF8(srcStream.GetStreamSize())
        
        If (LenB(rawFileString) <> 0) Then
            
            'To simplify processing, convert the string into a "stack" of "lines".
            Dim fileLines As pdStringStack
            Set fileLines = New pdStringStack
            fileLines.CreateFromMultilineString rawFileString
            
            'Validate the first line of the stack
            If Strings.StringsEqual(Trim$(fileLines.GetString(0)), "# PhotoDemon palette", True) Then
                
                '(Normally we would pop lines like a stack, but the user may want to preserve the
                ' original palette order - in which case we need to also process the text lines in
                ' their original file order.)
                Dim i As Long, curLine As String
                For i = 0 To fileLines.GetNumOfStrings - 1
                    
                    'Retrieve the current line and trim any leading or trailing whitespace
                    curLine = Trim$(fileLines.GetString(i))
                    
                    'Ignore null-length lines
                    If (LenB(curLine) <> 0) Then
                        
                        'Comment lines start with a "#" - these can be completely ignored
                        If Strings.StringsEqual(Left$(curLine, 1), "#", False) Then
                            'Do nothing
                        
                        'Any other line should be treated as an RGBA hex entry, with an optional trailing color name.
                        ' (There is never a prepended #, 0x, or other hex indicator.)
                        ElseIf (Len(curLine) >= 6) Then
                            
                            Dim tmpQuad As RGBQuad, srcColorName As String
                            
                            'VB6 error handling is a nightmare, alas...
                            On Error GoTo BadLineColor
                            
                            'Fill the RGB entries directly
                            With tmpQuad
                                .Red = CByte("&h" & Mid$(curLine, 1, 2))
                                .Green = CByte("&h" & Mid$(curLine, 3, 2))
                                .Blue = CByte("&h" & Mid$(curLine, 5, 2))
                                
                                'Alpha is optional, but PD will always write it
                                If (Len(curLine) >= 8) Then
                                    
                                    'Ensure this isn't just an RGB line with an appended name
                                    If (LCase$(Mid$(curLine, 7, 2)) <> " ") Then .Alpha = CByte("&h" & Mid$(curLine, 7, 2))
                                    
                                End If
                            End With
                            
                            'Color names are always optional.  Look for a space in the current string, and if one exists,
                            ' take all characters after it (trimming blank padding on either side).
                            If (InStr(1, curLine, " ", vbBinaryCompare) <> 0) Then
                                srcColorName = Trim$(Right$(curLine, Len(curLine) - InStr(1, curLine, " ", vbBinaryCompare)))
                            Else
                                srcColorName = vbNullString
                            End If
                            
                            'Resume normal VB6 error handling, then add this color to the collection
                            On Error GoTo 0
                            m_Palettes(m_NumOfGroups).AddColor tmpQuad, srcColorName
                            
BadLineColor:
                        End If
                        
                    End If
                
                Next i
                
            'If the first line doesn't validate, this is not a valid PD palette file; do nothing
            Else
            
            End If
                
            'If we haven't errored out, consider this a valid parse if we found any valid colors
            LoadPalettePhotoDemon = (m_Palettes(m_NumOfGroups).GetNumOfColors() > 0)
            
            'If the load was successful, assign a fake palette name using the filename
            If LoadPalettePhotoDemon Then
            
                m_Palettes(m_NumOfGroups).SetPaletteName Files.FileGetName(srcFile, True)
                m_Palettes(m_NumOfGroups).SetPaletteFilename srcFile
                
                'If original color order is not required, we can sort the palette prior to removing
                ' duplicate palette entries; this makes the process much faster.
                If removeDuplicateColors Then
                    If retainOriginalColorOrder Then
                        m_Palettes(m_NumOfGroups).FindAndRemoveDuplicates
                    Else
                        m_Palettes(m_NumOfGroups).SortFixedOrder
                        m_Palettes(m_NumOfGroups).FindAndRemoveDuplicates_Fast
                    End If
                End If
            
                m_NumOfGroups = m_NumOfGroups + 1
                
            End If
            
        Else
            InternalProblem "LoadPalettePhotoDemon failed to read the source stream as a UTF-8 buffer"
            LoadPalettePhotoDemon = False
        End If
        
    End If
    
    Exit Function
    
BadPDPalette:
    LoadPalettePhotoDemon = False

End Function

'Given a path to a .act file (Adobe Color Table), write the designated palette group to the file.
' - An existing .act file with the same name, if one exists, will be overwritten.
Friend Function SavePaletteAdobeColorTable(ByRef dstFile As String, Optional ByVal useGroupIndex As Long = 0) As Boolean
    
    On Error GoTo BadACTData
    
    'ACT files don't have a header.  Instead, they use a fixed file size of 768 or 772 bytes,
    ' per the spec at http://www.adobe.com/devnet-apps/photoshop/fileformatashtml/#50577411_pgfId-1070626
    Dim dstBytes() As Byte
    ReDim dstBytes(0 To 767) As Byte
        
    'Dump all colors - as RGB triples - into the temp array
    Dim numColorsSafe As Long
    numColorsSafe = m_Palettes(useGroupIndex).GetNumOfColors
    If (numColorsSafe > 256) Then numColorsSafe = 256
    
    Dim i As Long
    For i = 0 To numColorsSafe - 1
        With m_Palettes(useGroupIndex).GetPaletteColor(i)
            dstBytes(i * 3) = .Red
            dstBytes(i * 3 + 1) = .Green
            dstBytes(i * 3 + 2) = .Blue
        End With
    Next i
    
    'If this palette contains less than 256 colors, write the final palette entry to file multiple times;
    ' this prevents black (0, 0, 0) from being "in the palette" by default due to empty trailing bytes.
    If (numColorsSafe < 256) Then
        For i = numColorsSafe To 255
            With m_Palettes(useGroupIndex).GetPaletteColor(numColorsSafe - 1)
                dstBytes(i * 3) = .Red
                dstBytes(i * 3 + 1) = .Green
                dstBytes(i * 3 + 2) = .Blue
            End With
        Next i
    End If
    
    'And that's all there is to it!  Dump the finished palette to file.
    SavePaletteAdobeColorTable = Files.FileCreateFromByteArray(dstBytes, dstFile, True)
    
    Exit Function
    
BadACTData:
    SavePaletteAdobeColorTable = False
    
End Function

'Given a path to an .aco file (Adobe Photoshop Swatch extension), write the designated palette group to the file.
' - An existing .aco file with the same name, if one exists, will be overwritten.
' - Swatch version can be specified.  By default (-1), PhotoDemon will choose the best format for you.  If the
'    source palette does not contain color names, a v1 swatch will be written.  If the source palette *does* contain
'    color names, PD will write a v2 version.  You can override this by specifying the version you want.
Friend Function SavePaletteAdobeSwatch(ByRef dstFile As String, Optional ByVal swatchVersion As Long = -1, Optional ByVal useGroupIndex As Long = 0) As Boolean
    
    On Error GoTo BadSwatchData
    
    If Files.FileExists(dstFile) Then Files.FileDelete dstFile
    
    'This function references the Adobe spec frequently, particularly for things like magic numbers.  Here is the
    ' link I used (good as of Jan '18): http://www.adobe.com/devnet-apps/photoshop/fileformatashtml/#50577411_31265
    
    'Owing to their Mac heritage, Adobe files store everything as big-endian.  As such, we'll use a pdStream object
    ' with its _BE suffixed functions to greatly simplify the write process.
    Dim dstData As pdStream
    Set dstData = New pdStream
    If dstData.StartStream(PD_SM_FileBacked, PD_SA_ReadWrite, dstFile, , , OptimizeSequentialAccess) Then
    
        'Per the Adobe spec, swatch files are just long lists of 16-bit words.  This makes writing them straightforward.
        
        'Start by writing the version number we're using.  Version 2 allows us to also write color names.
        ' (If the user passed -1, choose a version for them.  Palettes without color names will be preferentially written
        '  in the v1 format.)
        If (swatchVersion < 1) Then
            If m_Palettes(useGroupIndex).DoesPaletteUseColorNames Then swatchVersion = 2 Else swatchVersion = 1
        ElseIf (swatchVersion > 2) Then
            swatchVersion = 2
        End If
        
        Dim swatchIterate As Long
        For swatchIterate = 1 To swatchVersion
        
            dstData.WriteInt_BE swatchIterate
            
            'Regardless of swatch version, we always write a v1 swatch first.  v2 swatches are simply v1 swatches with
            ' a special trailer.
            
            'Next, we need to write the number of colors in the palette
            Dim numOfColors As Long
            numOfColors = m_Palettes(useGroupIndex).GetNumOfColors()
            dstData.WriteIntU_BE numOfColors
            
            'PD always writes colors as RGB
            Dim targetColorSpace As AdobeSwatchColorSpace
            targetColorSpace = ASCS_RGB
            
            'Loop through all colors and write them as we go!
            Dim tmpPalEntry As PDPaletteEntry
            Dim r As Long, g As Long, b As Long
            
            Dim i As Long
            For i = 0 To numOfColors - 1
                
                tmpPalEntry = m_Palettes(useGroupIndex).GetPaletteEntry(i)
                
                'Write the color space identifier
                dstData.WriteInt_BE targetColorSpace
                
                'Write RGB values, but convert them to unsigned Ints on the range [0, 65535] first
                With tmpPalEntry.ColorValue
                    r = .Red
                    g = .Green
                    b = .Blue
                End With
                
                dstData.WriteIntU_BE r * 257
                dstData.WriteIntU_BE g * 257
                dstData.WriteIntU_BE b * 257
                dstData.WriteIntU_BE 0
                
                'If this is a v2 swatch, we also want to write this color's name out to file.
                If (swatchIterate = 2) Then
                    
                    'Retrieve the color name, if any
                    Dim srcColorName As String
                    srcColorName = tmpPalEntry.ColorName
                    
                    'Write the length of this color's name
                    Dim colorLen As Long
                    colorLen = Len(srcColorName)
                    dstData.WriteLong_BE colorLen
                    
                    'If the length is non-zero, write each string char in turn, converting to big-endian as we go
                    If (colorLen > 0) Then
                        
                        Dim j As Long
                        For j = 1 To colorLen
                            dstData.WriteIntU_BE AscW(Mid$(srcColorName, j, 1))
                        Next j
                        
                    End If
                    
                End If
                
            Next i
            
        Next swatchIterate
        
        'That's all!
        dstData.StopStream
        
        SavePaletteAdobeSwatch = True
        
    End If
    
    Exit Function
    
BadSwatchData:
    InternalProblem "SavePaletteAdobeSwatch error # " & CStr(Err.Number) & ": " & Err.Description
    SavePaletteAdobeSwatch = False
    
End Function

'Given a path to an .ase file (Adobe Swatch Exchange), write all existing palette groups to the file.
' (ASE files are the only format that supports multiple palettes within a single file.)
' - An existing .ase file with the same name, if one exists, will be overwritten.
Friend Function SavePaletteAdobeSwatchExchange(ByRef dstFile As String) As Boolean
    
    On Error GoTo BadSwatchData
    
    If Files.FileExists(dstFile) Then Files.FileDelete dstFile
    
    'Adobe has not publicly documented the ASE format, which means we have to rely on reverse-engineering to decode
    ' the format.  Several links proved very useful in this regard; thank you to each of them:
    ' http://www.selapa.net/swatches/colors/fileformats.php#adobe_ase
    ' https://www.cyotek.com/blog/reading-adobe-swatch-exchange-ase-files-using-csharp
    
    'As usual, pdStream saves the day by making it very easy to stream arbitrary bytes out to file
    Dim dstData As pdStream
    Set dstData = New pdStream
    SavePaletteAdobeSwatchExchange = dstData.StartStream(PD_SM_MemoryBacked, PD_SA_ReadWrite)
    
    If SavePaletteAdobeSwatchExchange Then
        
        'Owing to their Mac heritage, Adobe files store everything as big-endian.  As such, we'll be using
        ' many _BE suffixed write functions with the pdStream object.
                
        'Start by writing the required four-char "ASEF" descriptor.
        dstData.WriteString_ASCII "ASEF"
                
        'Next comes file version number.  Version is defined as major/minor, two 16-bit integers.
        ' (The only currently supported version is 1.0.)
        dstData.WriteInt_BE 1
        dstData.WriteInt_BE 0
        
        ' The next entry is a 4-byte integer defining the number of blocks in the file.  Blocks come in
        ' three types: group start, group end, and color entry.  We can figure out this number by counting
        ' each unique group as two blocks (start/end) plus one block for each color in each palette.
        Dim numOfBlocks As Long
        Dim i As Long, j As Long
        For i = 0 To Me.GetPaletteGroupCount() - 1
            If (Me.GetPaletteGroupCount > 0) Or (LenB(m_Palettes(i).GetPaletteName) <> 0) Then numOfBlocks = numOfBlocks + 2
            numOfBlocks = numOfBlocks + m_Palettes(i).GetNumOfColors()
        Next i
        
        dstData.WriteLong_BE numOfBlocks
        
        'Time to write each palette!
        For i = 0 To Me.GetPaletteGroupCount() - 1
        
            'Each block uses a standardized header:
            ' - 2 bytes to define the block "type"
            ' - 4 bytes to define the block "length", which does *not* include these 6 bytes in the header.
            ' The important thing to note when calculating block "length" is that it includes the length of
            ' any text stored in the block - such as group and/or color names.
            
            'We'll only divide palettes into groups if this palettte has a name, or if there are multiple palette
            ' groups in this object.
            Dim wroteAGroup As Boolean
            wroteAGroup = (LenB(m_Palettes(i).GetPaletteName) <> 0) Or (Me.GetPaletteGroupCount() > 1)
            
            If wroteAGroup Then
            
                'Start by writing a "group start" block for this palette.  (This magic number comes from Adobe.)
                dstData.WriteInt_BE &HC001
                
                'This block only contains a group name, and its length is calculated as:
                ' - LenB(m_palettes(i).GetPaletteName)
                ' - If the above is non-zero, we need 2-bytes to define string length, and two extra bytes for a
                '   trailing null WCHAR.
                If (LenB(m_Palettes(i).GetPaletteName) = 0) Then
                    dstData.WriteLong_BE 0
                Else
                    
                    dstData.WriteLong_BE LenB(m_Palettes(i).GetPaletteName) + 4
                    
                    'Time to write a Unicode string for this palette.  Start by writing a 2-byte string length
                    ' (the length *IN WCHARS*, confusingly), and include an extra byte for the trailing null char.
                    dstData.WriteInt_BE Len(m_Palettes(i).GetPaletteName) + 1
                    
                    'Next, write the string, and explicitly request a trailing null
                    dstData.WriteString_UnicodeBE m_Palettes(i).GetPaletteName, True
                    
                End If
                
            End If
            
            'Next, we need to write out all colors in this palette.
            For j = 0 To m_Palettes(i).GetNumOfColors - 1
            
                'Write a color block indicator (magic number comes from Adobe)
                dstData.WriteInt_BE &H1
                
                'Color block length is contingent on the presence of color names.  By default, RGB colors
                ' always require 18 bytes: 4 for the color descriptor ("RGB "), 12 for 3 RGB values, stored as
                ' 32-bit floats, then 2 bytes for the color "type" (global, spot, normal).  Those 18 bytes
                ' are further increased by 2 bytes for name length, plus LenB(colorname) + 2 if a name exists.
                Dim nameLengthInBytes As Long, colorBlockLength As Long, tmpPaletteEntry As PDPaletteEntry
                tmpPaletteEntry = m_Palettes(i).GetPaletteEntry(j)
                
                nameLengthInBytes = 2
                If (LenB(tmpPaletteEntry.ColorName) <> 0) Then nameLengthInBytes = nameLengthInBytes + LenB(tmpPaletteEntry.ColorName) + 2
                colorBlockLength = 18 + nameLengthInBytes
                
                'Write our calculated block length
                dstData.WriteLong_BE colorBlockLength
                
                'Write name length and the name itself, if any.  (Confusingly, note that the name length is
                ' in *WCHARS*, not BYTES.
                If (nameLengthInBytes > 2) Then
                    dstData.WriteInt_BE Len(tmpPaletteEntry.ColorName) + 1
                    dstData.WriteString_UnicodeBE tmpPaletteEntry.ColorName, True
                Else
                    dstData.WriteInt_BE 0
                End If
                
                'Write the color space identifier
                dstData.WriteString_ASCII "RGB "
                
                'Write individual RGB values as 32-bit floats on the range [0, 1]
                dstData.WriteFloat_BE CSng(tmpPaletteEntry.ColorValue.Red) / 255#
                dstData.WriteFloat_BE CSng(tmpPaletteEntry.ColorValue.Green) / 255#
                dstData.WriteFloat_BE CSng(tmpPaletteEntry.ColorValue.Blue) / 255#
                
                'Write two empty bytes as the color "type"
                dstData.WriteInt_BE 0
                
            Next j
            
            'If we wrote a "group start" block, we need to mirror it with a "group end" block.
            If wroteAGroup Then
                
                'Again, use Adobe's magic number
                dstData.WriteInt_BE &HC002
                
                'Close-group blocks are always 0-length
                dstData.WriteLong_BE 0
                
            End If
            
        Next i
        
        'Write the finished palette to file!
        SavePaletteAdobeSwatchExchange = Files.FileCreateFromPtr(dstData.Peek_PointerOnly(0), dstData.GetStreamSize(), dstFile)
        
        'Free the underlying stream object
        dstData.StopStream
        
    End If
    
    Exit Function
    
BadSwatchData:
    InternalProblem "SavePaletteAdobeSwatchExchange error # " & CStr(Err.Number) & ": " & Err.Description
    SavePaletteAdobeSwatchExchange = False
    
End Function

'Given a path to a .gpl file (GIMP Palette), write the designated palette group to the file.
' - An existing .gpl file with the same name, if one exists, will be overwritten.
Friend Function SavePaletteGIMP(ByRef dstFile As String, Optional ByVal useGroupIndex As Long = 0, Optional ByVal addPDCredit As Boolean = True) As Boolean
    
    On Error GoTo InvalidPalette
    
    If Files.FileExists(dstFile) Then Files.FileDelete dstFile
    
    'GIMP palettes are just text files, and pdString makes it easy to efficiently assemble an in-memory string
    Dim dstText As pdString
    Set dstText = New pdString
    
    'Start by writing a few standard lines, kind of like a header
    With dstText
        .AppendLine "GIMP Palette"
        If (LenB(m_Palettes(useGroupIndex).GetPaletteName()) <> 0) Then .AppendLine "Name: " & m_Palettes(useGroupIndex).GetPaletteName()
        .AppendLine "#"
        If addPDCredit Then
            .AppendLine "# " & g_Language.TranslateMessage("palette exported by PhotoDemon %1", "(https://photodemon.org)")
            .AppendLine "#"
        End If
    End With
    
    Const SPACE_CHAR As String = " "
            
    'Next come all the colors.  Color names, if they exist, will be written automatically.
    Dim i As Long, tmpPaletteEntry As PDPaletteEntry, colorLine As String
    For i = 0 To m_Palettes(useGroupIndex).GetNumOfColors() - 1
        
        tmpPaletteEntry = m_Palettes(useGroupIndex).GetPaletteEntry(i)
        
        With tmpPaletteEntry
            colorLine = GIMPColorHelper(.ColorValue.Red) & SPACE_CHAR & GIMPColorHelper(.ColorValue.Green) & SPACE_CHAR & GIMPColorHelper(.ColorValue.Blue)
            
            'I realize that use of a tab char here is odd, but the palette files that ship with GIMP files
            ' all use tab delimiters (instead of spaces) before the color name.  I don't know why.
            If (LenB(.ColorName) <> 0) Then colorLine = colorLine & vbTab & .ColorName
        End With
        
        dstText.AppendLine colorLine
        
    Next i
    
    'That's all there is to it!  Write the finished text out to file.
    Dim finalText As String
    finalText = dstText.ToString()
    SavePaletteGIMP = Files.FileSaveAsText(finalText, dstFile, True, False)
    
    Exit Function

InvalidPalette:
    SavePaletteGIMP = False

End Function

'GIMP palette files like to format colors as strings with guaranteed length 3; spaces are prepended as necessary.
Private Function GIMPColorHelper(ByVal srcColor As Long) As String
    GIMPColorHelper = CStr(srcColor)
    If (Len(GIMPColorHelper) < 3) Then GIMPColorHelper = String$(3 - Len(GIMPColorHelper), " ") & GIMPColorHelper
End Function

'Given a path to a .txt file (Paint.NET palette), write the designated palette group to the file.
' - An existing .txt file with the same name, if one exists, will be overwritten.
Friend Function SavePalettePaintDotNet(ByRef dstFile As String, Optional ByVal useGroupIndex As Long = 0, Optional ByVal addPDCredit As Boolean = True, Optional ByVal forceAlphaTo255 As Boolean = True) As Boolean

    On Error GoTo BadPaintDotNetPalette
    
    If Files.FileExists(dstFile) Then Files.FileDelete dstFile
    
    'Paint.NET palettes are just text files, and pdString makes it easy to efficiently assemble an in-memory string
    Dim dstText As pdString
    Set dstText = New pdString
    
    'Start by writing a few standard lines, kind of like a header
    With dstText
        .AppendLine "; paint.net palette file"
        If addPDCredit Then .AppendLine ";  " & g_Language.TranslateMessage("palette exported by PhotoDemon %1", "(https://photodemon.org)")
    End With
    
    'Next come all the colors.  Paint.NET does not support color names, so only RGB values will be written.
    Dim i As Long, tmpPaletteEntry As PDPaletteEntry, colorLine As String
    
    'Note that Paint.NET palettes also top out at 96 colors; < 96 colors are okay (as the remaining entries
    ' will be filled with white by Paint.NET, but > 96 is disallowed.  If you want to enforce this,
    ' uncomment the boundary check below.
    Dim numColorsBound As Long
    numColorsBound = m_Palettes(useGroupIndex).GetNumOfColors()
    'If (numColorsBound > 96) Then numColorsBound = 96
    
    For i = 0 To numColorsBound - 1
        
        tmpPaletteEntry = m_Palettes(useGroupIndex).GetPaletteEntry(i)
        colorLine = String$(8, 0)
        
        'Per the "spec" (https://www.getpaint.net/doc/latest/WorkingWithPalettes.html),
        ' Paint.NET palette files are text files where all colors are defined as AARRGGBB hex entries.
        ' In the example file they provide, the hex chars are uppercase - so we mimic that here.
        If forceAlphaTo255 Then
            Mid$(colorLine, 1, 2) = "FF"
        Else
            Mid$(colorLine, 1, 2) = UCase$(Colors.GetTwoCharHexFromByte(tmpPaletteEntry.ColorValue.Alpha))
        End If
        
        Mid$(colorLine, 3, 2) = UCase$(Colors.GetTwoCharHexFromByte(tmpPaletteEntry.ColorValue.Red))
        Mid$(colorLine, 5, 2) = UCase$(Colors.GetTwoCharHexFromByte(tmpPaletteEntry.ColorValue.Green))
        Mid$(colorLine, 7, 2) = UCase$(Colors.GetTwoCharHexFromByte(tmpPaletteEntry.ColorValue.Blue))
        
        dstText.AppendLine colorLine
        
    Next i
    
    'That's all there is to it!  Write the finished text out to file.
    Dim finalText As String
    finalText = dstText.ToString()
    SavePalettePaintDotNet = Files.FileSaveAsText(finalText, dstFile, True, False)
    
    Exit Function
    
BadPaintDotNetPalette:
    SavePalettePaintDotNet = False

End Function

'Given a path to a .pal file (JASC palette), write the designated palette group to the file.
' - An existing .pal file with the same name, if one exists, will be overwritten.
Friend Function SavePalettePaintShopPro(ByRef dstFile As String, Optional ByVal useGroupIndex As Long = 0) As Boolean

    On Error GoTo BadPaintShopProPalette
    If Files.FileExists(dstFile) Then Files.FileDelete dstFile
    
    'PSP palettes are just text files, and pdString makes it easy to efficiently assemble an in-memory string
    Dim dstText As pdString
    Set dstText = New pdString
    
    'Note that the PSP spec supposedly restricts palettes to 16 or 256 colors, but there are many palettes
    ' "in the wild" that don't abide these conditions - so we just write out the size of the palette as-is,
    ' with a hard limit of 256 colors in case the source palette is very large.
    Dim numColorsBound As Long
    numColorsBound = m_Palettes(useGroupIndex).GetNumOfColors()
    If (numColorsBound > 256) Then numColorsBound = 256
    
    'Start by writing the basic JASC header
    With dstText
        .AppendLine "JASC-PAL"
        .AppendLine "0100"
        .AppendLine CStr(numColorsBound)
    End With
    
    'Next come all the colors.  PSP palettes don't support color names, so only RGB values will be written.
    Const SPACE_CHAR As String = " "
    Dim i As Long, tmpPaletteEntry As PDPaletteEntry, colorLine As String
    For i = 0 To numColorsBound - 1
        
        tmpPaletteEntry = m_Palettes(useGroupIndex).GetPaletteEntry(i)
        With tmpPaletteEntry.ColorValue
            colorLine = CStr(.Red) & SPACE_CHAR & CStr(.Green) & SPACE_CHAR & CStr(.Blue)
        End With
        dstText.AppendLine colorLine
        
    Next i
    
    'That's all there is to it!  Write the finished text out to file.
    Dim finalText As String
    finalText = dstText.ToString()
    SavePalettePaintShopPro = Files.FileSaveAsText(finalText, dstFile, True, False)
    
    Exit Function
    
BadPaintShopProPalette:
   SavePalettePaintShopPro = False

End Function

'I don't like to invent my own formats, but palettes are one of those weird ones where every existing format
' has a bunch of irritating limitations and/or inconsistencies.  As such, for internal tools especially it's
' useful to have a predictable palette format with guaranteed behavior on things like alpha bytes.
'
'The format of a PD palette is similar to Paint.NET, with colors being simple hex strings, but we also
' add support for unlimited color counts (even > 256 because that's no problem internally, and it makes it
' easier to apply things like retro palettes for systems with 4096 color counts), color names, and other
' palette features.
'
'Note that the palette files are ultimately UTF-8 encoded (no BOM) with Windows line endings (CRLF).
Friend Function SavePalettePhotoDemon(ByRef dstFile As String, Optional ByVal useGroupIndex As Long = 0, Optional ByVal addPDCredit As Boolean = True, Optional ByVal forceAlphaTo255 As Boolean = False) As Boolean

    On Error GoTo BadPDPalette
    
    If Files.FileExists(dstFile) Then Files.FileDelete dstFile
    
    'PhotoDemon palettes are just text files, and pdString makes it easy to efficiently assemble an in-memory string
    Dim dstText As pdString
    Set dstText = New pdString
    
    'Start by writing a standard header
    With dstText
        
        'This opening line *IS CURRENTLY REQUIRED*; it makes it much easier for us to correctly identify
        ' a palette file, without relying on complicated heuristics
        .AppendLine "# PhotoDemon palette"
        
        'Any #-prefixed line following the first is considered a comment line.
        ' You can stick whatever you want in them and PD will ignore it
        If addPDCredit Then .AppendLine "# palette exported by PhotoDemon " & Updates.GetPhotoDemonVersion() & ", https://photodemon.org"
        
    End With
    
    'Next come all the colors.  RGBA values and color names (if any) are fully supported
    Dim i As Long, tmpPaletteEntry As PDPaletteEntry, colorLine As String
    
    Dim numColorsBound As Long
    numColorsBound = m_Palettes(useGroupIndex).GetNumOfColors()
    
    For i = 0 To numColorsBound - 1
        
        tmpPaletteEntry = m_Palettes(useGroupIndex).GetPaletteEntry(i)
        colorLine = String$(8, 0)
        
        'Colors are written in RGBA format.  A few notes:
        ' - alpha can be omitted (full opacity will be assumed), but PD always writes it
        ' - case of hex chars doesn't matter, but PD always writes lower-case
        ' - shortened strings are not allowed - the hex string must be 6 chars (RGB) or 8 chars (RGBA), no exceptions
        ' - if a color name exists, add a single space after the RGB/A string, followed by the color name.
        '   (No restrictions are currently placed on color names; the text is ultimately written as UTF-8 so
        '    valid Unicode of any size is theoretically okay!)
        Mid$(colorLine, 1, 2) = LCase$(Colors.GetTwoCharHexFromByte(tmpPaletteEntry.ColorValue.Red))
        Mid$(colorLine, 3, 2) = LCase$(Colors.GetTwoCharHexFromByte(tmpPaletteEntry.ColorValue.Green))
        Mid$(colorLine, 5, 2) = LCase$(Colors.GetTwoCharHexFromByte(tmpPaletteEntry.ColorValue.Blue))
        If forceAlphaTo255 Then
            Mid$(colorLine, 7, 2) = "ff"
        Else
            Mid$(colorLine, 7, 2) = LCase$(Colors.GetTwoCharHexFromByte(tmpPaletteEntry.ColorValue.Alpha))
        End If
        
        'If the color has a name, append it after a space char
        If (LenB(m_Palettes(useGroupIndex).GetPaletteEntry(i).ColorName) <> 0) Then colorLine = colorLine & " " & m_Palettes(useGroupIndex).GetPaletteEntry(i).ColorName
        
        dstText.AppendLine colorLine
        
    Next i
    
    'That's all there is to it!  Write the finished string out to file.
    Dim finalText As String
    finalText = dstText.ToString()
    SavePalettePhotoDemon = Files.FileSaveAsText(finalText, dstFile, True, False)
    
    Exit Function
    
BadPDPalette:
    SavePalettePhotoDemon = False

End Function

Friend Sub Reset()
    m_NumOfGroups = 0
    ReDim m_Palettes(0) As pdPaletteChild
    Set m_Palettes(0) = New pdPaletteChild
End Sub

Private Sub InternalProblem(ByRef errMsg As String)
    PDDebug.LogAction "WARNING!  pdPalette returned an error: " & errMsg
End Sub

Private Sub Class_Initialize()
    Me.Reset
End Sub
