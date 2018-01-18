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
'Last updated: 18/January/18
'Last update: forcibly remove duplicate entries from imported palettes
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

    'Branch according to format; at present, only .gpl files are supported
    LoadPaletteFromFile = LoadPaletteFromGIMPFile(srcFile)
    
End Function

'Given a valid path to a GIMP-format .gpl file, return an array of RGBQuad entries
Private Function LoadPaletteFromGIMPFile(ByRef srcFile As String, Optional ByVal removeDuplicateColors As Boolean = True, Optional ByVal optimizeForPerformance As Boolean = False) As Boolean
    
    On Error GoTo InvalidPalette
    
    'GIMP palettes contain extra information that PD doesn't currently have a way to use.  We only parse the
    ' information we require, so you may need to look elsewhere for a full parsing.
    
    'Dump the full file into a string
    Dim rawFileString As String
    LoadPaletteFromGIMPFile = Files.FileLoadAsString(srcFile, rawFileString, True)
    
    If LoadPaletteFromGIMPFile Then
    
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
            
            'Parse each line in turn
            Dim curLine As String
            Do While fileLines.PopString(curLine)
                
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
                        targetColorInt = CLng(targetColor)
                        On Error GoTo 0
                        
                        'If we're still here, the color was parsed correctly.  Store it as the red component of the
                        ' current palette entry.
                        m_PaletteColors(m_PaletteColorCount).Red = targetColorInt
                        
                        'Trim the color we've parsed out of the string, then repeat the above steps
                        curLine = Trim$(Right$(curLine, Len(curLine) - InStr(1, curLine, SPACE_CHAR, vbBinaryCompare)))
                        targetColor = Left$(curLine, InStr(1, curLine, SPACE_CHAR, vbBinaryCompare) - 1)
                        On Error GoTo BadLineColor
                        targetColorInt = CLng(targetColor)
                        On Error GoTo 0
                        m_PaletteColors(m_PaletteColorCount).Green = targetColorInt
                        
                        '...and one last time, for the blue component.  Note that the resulting string may not
                        ' have a trailing space, so we forcibly add one to simplify the parser.
                        curLine = Trim$(Right$(curLine, Len(curLine) - InStr(1, curLine, SPACE_CHAR, vbBinaryCompare))) & SPACE_CHAR
                        targetColor = Left$(curLine, InStr(1, curLine, SPACE_CHAR, vbBinaryCompare) - 1)
                        On Error GoTo BadLineColor
                        targetColorInt = CLng(targetColor)
                        On Error GoTo 0
                        m_PaletteColors(m_PaletteColorCount).Blue = targetColorInt
                        
                        'GIMP palettes do not support alpha channels.  Forcibly set a value of 255.
                        m_PaletteColors(m_PaletteColorCount).Alpha = 255
                        
                        'If we made it all the way here, this line was successfully parsed for color data.
                        ' Advance the color count tracker and resume the line parser.
                        m_PaletteColorCount = m_PaletteColorCount + 1
                        
BadLineColor:
                    
                    End If
                    
                End If
            
            Loop
            
            'If we haven't errored out, consider this a valid parse
            LoadPaletteFromGIMPFile = True
            
            'On valid loads, check and remove any duplicate palette entries.  (As an ASCII format, GIMP makes
            ' no guarantees about this, but we don't want to waste time palette-matching duplicate entries.)
            If removeDuplicateColors Then FindAndRemoveDuplicates
            
            'If the user plans to use this palette for color-matching (and really, what else are they gonna
            ' do with it?) we can pre-sort the palette to improve matching performance.
            If optimizeForPerformance Then SortByLuminance
            
        End If
    
    End If
    
    If LoadPaletteFromGIMPFile Then m_PaletteFilename = srcFile
    
    Exit Function

InvalidPalette:
    
    m_PaletteName = vbNullString
    m_PaletteColorCount = 0
    m_PaletteFilename = vbNullString
    LoadPaletteFromGIMPFile = False

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
