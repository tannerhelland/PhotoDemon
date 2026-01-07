VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdPaletteChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Palette Child Container (see pdPalette for additional details)
'Copyright 2018-2026 by Tanner Helland
'Created: 25/January/18
'Last updated: 01/Feburary/18
'Last update: remove 256 color limit on palettes; size is now unrestricted
'
'Please the pdPalette class for additional details on how this child class works.  Individual palette files
' (Adobe Swatch Exchange files, in particular) can actually contain multiple independent palette groups.
' This is helpful for exchanging swatches, per the name, but it creates some complications for PD where
' I originally assumed that all palettes would be independent entities.
'
'To work around this, a single pdPalette instance now manages one (by default) or more (as necessary)
' pdPaletteChild instances, depending on needs.  This allows us to properly support ASE files, and it
' may prove useful in the future if we decide to ship a palette collection with PD (as all built-in
' palettes can be managed by a single pdPalette instance).
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'If you want to restrict palettes to a maximum count (e.g. 256 colors), set this value to 256
' or any other limit.  PD has its own internal palette matchers that support color indices > 8-bits,
' so we don't need this feature, but some GDI functions and/or image formats won't work with large
' color counts and this can be a helpful way to restrict color count.
Private Const MAX_PALETTE_SIZE As Long = LONG_MAX

'This initial palette size is always allocated by default on a new class instance.  Make it smaller
' to conserve memory; make it larger to improve performance when adding new colors to the palette.
Private Const INIT_PALETTE_SIZE As Long = 8

'Current number of colors in the palette.  This value is independent of the size of the palette array.
Private m_NumOfColors As Long
Private m_Palette() As PDPaletteEntry

'Palette name and filename.  These may be the same; only certain palettes support the concept of an
' embedded "palette name".
Private m_PaletteName As String
Private m_PaletteFilename As String

'A specially designed QuickSort algorithm is used to sort palettes.  We define our own internal sort criteria
' using this custom type, which allows us to implement different types of palette sorts "on the fly" by simply
' plugging-in different sort criterias.
Private Type PaletteSort
    pSortCriteria As Single
    pOrigIndex As Long
End Type

Friend Function CopyRGBQuadsToArray(ByRef dstArray() As RGBQuad) As Boolean
    
    CopyRGBQuadsToArray = (m_NumOfColors > 0)
    
    If CopyRGBQuadsToArray Then
        ReDim dstArray(0 To m_NumOfColors - 1) As RGBQuad
        Dim i As Long
        For i = 0 To m_NumOfColors - 1
            dstArray(i) = m_Palette(i).ColorValue
        Next i
    End If
    
End Function

'Create a new palette instance from a source RGB quad collection.  The source array must be 0-based.
Friend Sub CreateFromRGBQuads(ByRef srcArray() As RGBQuad, Optional ByVal paletteName As String = vbNullString, Optional ByVal paletteFilename As String = vbNullString)
    
    If (UBound(srcArray) >= LBound(srcArray)) Then
        
        ReDim m_Palette(0 To UBound(srcArray)) As PDPaletteEntry
        m_NumOfColors = UBound(srcArray) + 1
        
        Dim i As Long
        For i = 0 To m_NumOfColors - 1
            m_Palette(i).ColorValue = srcArray(i)
        Next i
        
        m_PaletteName = paletteName
        m_PaletteFilename = paletteFilename
        
    End If
    
End Sub

'If at least one color has a name, this function will return TRUE.
Friend Function DoesPaletteUseColorNames() As Boolean
    
    If (Me.GetNumOfColors > 0) Then
        
        Dim i As Long
        For i = 0 To Me.GetNumOfColors - 1
            DoesPaletteUseColorNames = (LenB(m_Palette(i).ColorName) <> 0)
            If DoesPaletteUseColorNames Then Exit Function
        Next i
        
    Else
        DoesPaletteUseColorNames = False
    End If
    
End Function

Friend Function GetNumOfColors() As Long
    GetNumOfColors = m_NumOfColors
End Function

Friend Sub SetNumOfColors(ByVal newNum As Long)
    If (newNum > UBound(m_Palette) + 1) Then ReDim Preserve m_Palette(0 To newNum - 1) As PDPaletteEntry
    m_NumOfColors = newNum
End Sub

Friend Function GetPaletteColor(ByVal palIndex As Long) As RGBQuad
    GetPaletteColor = m_Palette(palIndex).ColorValue
End Function

Friend Function GetPaletteColorAsLong(ByVal palIndex As Long) As Long
    GetPaletteColorAsLong = RGB(m_Palette(palIndex).ColorValue.Red, m_Palette(palIndex).ColorValue.Green, m_Palette(palIndex).ColorValue.Blue)
End Function

Friend Function GetPaletteEntry(ByVal palIndex As Long) As PDPaletteEntry
    GetPaletteEntry = m_Palette(palIndex)
End Function

Friend Function GetPaletteFilename() As String
    GetPaletteFilename = m_PaletteFilename
End Function

Friend Function GetPaletteName() As String
    GetPaletteName = m_PaletteName
End Function

Friend Sub SetPaletteFilename(ByRef newFilename As String)
    m_PaletteFilename = newFilename
End Sub

Friend Sub SetPaletteName(ByRef newName As String)
    m_PaletteName = newName
End Sub

'Add a new color to this palette.  Returns the index of the added color, which can be helpful for knowing
' how many colors currently exist in the palette as-you-go.
Friend Function AddColor(ByRef srcQuad As RGBQuad, Optional ByRef srcColorName As String = vbNullString) As Long
    
    'We currently enforce a fixed 256 color palette size limit.  I will look at removing this in the future,
    ' pending additional testing on functions that rely on 8-bit palette indices.
    If (m_NumOfColors >= MAX_PALETTE_SIZE) Then Exit Function
    
    If (m_NumOfColors > UBound(m_Palette)) Then ReDim Preserve m_Palette(0 To m_NumOfColors * 2 - 1) As PDPaletteEntry
    
    With m_Palette(m_NumOfColors)
        .ColorValue = srcQuad
        .ColorName = srcColorName
    End With
    
    m_NumOfColors = m_NumOfColors + 1
    
End Function

Friend Sub Reset()
    m_NumOfColors = 0
    ReDim m_Palette(0 To INIT_PALETTE_SIZE - 1) As PDPaletteEntry
    m_PaletteName = vbNullString
    m_PaletteFilename = vbNullString
End Sub

Private Sub Class_Initialize()
    Me.Reset
End Sub

'Many palette files are text-based, and nothing prevents problematic situations like duplicate colors.
' (In fact, some shitty formats *cough* Paint.NET *cough* require a *fiex* number of colors, which practically
' guarantees duplicate entries for small palettes.)  After loading a palette from file, I recommend calling this
' function to remove any duplicate palette entries.  It is *not* well-optimized (potentially O(n^2)) but seeing
' as PD only supports 8-bit palettes, I have not yet found it worth the trouble to optimize further.
Friend Sub FindAndRemoveDuplicates()

    'Only palettes with multiple entries need to be searched for duplicates
    If (m_NumOfColors <= 1) Then Exit Sub
    
    Dim needsRemoval() As Byte
    ReDim needsRemoval(0 To UBound(m_Palette)) As Byte
    
    Dim i As Long, j As Long, targetValue As RGBQuad, duplicatesFound As Boolean
    For i = 0 To m_NumOfColors - 2
        
        'Array accesses are slow in VB; use a local value instead
        targetValue = m_Palette(i).ColorValue
        
        For j = i + 1 To m_NumOfColors - 1
            With m_Palette(j)
                If (.ColorValue.Red = targetValue.Red) And (.ColorValue.Green = targetValue.Green) And (.ColorValue.Blue = targetValue.Blue) And (.ColorValue.Alpha = targetValue.Alpha) Then
                    If Strings.StringsEqual(.ColorName, m_Palette(i).ColorName) Then
                        needsRemoval(j) = 1
                        duplicatesFound = True
                    End If
                End If
            End With
        Next j
        
    Next i
    
    'Remove all flagged entries
    If duplicatesFound Then
    
        Dim numRemoved As Long
        
        For i = 1 To m_NumOfColors - 1
            If (needsRemoval(i) = 0) Then
                m_Palette(i - numRemoved) = m_Palette(i)
            Else
                numRemoved = numRemoved + 1
            End If
        Next i
        
        m_NumOfColors = m_NumOfColors - numRemoved
        ReDim Preserve m_Palette(0 To m_NumOfColors - 1) As PDPaletteEntry
        PDDebug.LogAction "Palette contained duplicate entries; color count reduced from " & CStr(m_NumOfColors + numRemoved) & " to " & CStr(m_NumOfColors) & " colors."
        
    End If
    
End Sub

'As a faster option, if you first sort the palette by RGB order, removing duplicates is a piece of cake as
' we can simply check neighboring entries for matches.  Thanks to our nice QuickSort implementation, this
' is a much faster O(nLog(n)) operation (compared to O(n^2) for the regular remove duplicates function).
Friend Sub FindAndRemoveDuplicates_Fast()

    'Only palettes with multiple entries need to be searched for duplicates
    If (m_NumOfColors <= 1) Then Exit Sub
    
    Dim needsRemoval() As Byte
    ReDim needsRemoval(0 To UBound(m_Palette)) As Byte
    
    Dim i As Long, targetValue As RGBQuad, duplicatesFound As Boolean
    For i = 0 To UBound(m_Palette) - 1
        
        'Because the palette is sorted, we only need to compare neighboring entries
        targetValue = m_Palette(i).ColorValue
        
        With m_Palette(i + 1)
            If (.ColorValue.Red = targetValue.Red) And (.ColorValue.Green = targetValue.Green) And (.ColorValue.Blue = targetValue.Blue) And (.ColorValue.Alpha = targetValue.Alpha) Then
                If Strings.StringsEqual(.ColorName, m_Palette(i).ColorName) Then
                    needsRemoval(i + 1) = 1
                    duplicatesFound = True
                End If
            End If
        End With
        
    Next i
    
    'Remove all flagged entries
    If duplicatesFound Then
    
        Dim numRemoved As Long
        
        For i = 1 To m_NumOfColors - 1
            If (needsRemoval(i) = 0) Then
                m_Palette(i - numRemoved) = m_Palette(i)
            Else
                numRemoved = numRemoved + 1
            End If
        Next i
        
        m_NumOfColors = m_NumOfColors - numRemoved
        ReDim Preserve m_Palette(0 To m_NumOfColors - 1) As PDPaletteEntry
        PDDebug.LogAction "Palette contained duplicate entries; color count reduced from " & CStr(m_NumOfColors + numRemoved) & " to " & CStr(m_NumOfColors) & " colors."
        
    End If
    
End Sub

'Some palette sort methods are better when order is randomized.  (Note that we use a fixed key
' for debugging and profiling purposes.)
Friend Sub RandomizeOrder()
    
    If (m_NumOfColors <= 1) Then Exit Sub
    
    Randomize -1
    Randomize 1234567
    
    Dim rndLimit As Long
    rndLimit = (m_NumOfColors - 1)
    
    Dim i As Long, firstIndex As Long, secondIndex As Long, tmpEntry As PDPaletteEntry
    For i = 0 To m_NumOfColors * 2 - 1
        firstIndex = Rnd * rndLimit
        secondIndex = Rnd * rndLimit
        tmpEntry = m_Palette(firstIndex)
        m_Palette(firstIndex) = m_Palette(secondIndex)
        m_Palette(secondIndex) = tmpEntry
    Next i

End Sub

'When constructing KD-trees specifically, sorting along a single channel (likely red?) can lead to a more
' balanced tree than a luminance sort, as it's more likely to spread the palette contents over the entire
' RGB cube, instead of concentrating them along a lone diagonal.
Friend Sub SortByChannel(Optional ByVal channelIndex As Long = 0)

    Dim pSort() As PaletteSort
    ReDim pSort(0 To m_NumOfColors - 1) As PaletteSort
    
    Dim r As Long, g As Long, b As Long, a As Long
    
    Dim multTable() As Single
    ReDim multTable(0 To 3) As Single
    multTable(0) = 2 ^ 0
    multTable(1) = 2 ^ 8
    multTable(2) = 2 ^ 16
    multTable(3) = 2 ^ 24
        
    Dim i As Long
    For i = 0 To m_NumOfColors - 1
        
        With m_Palette(i).ColorValue
            r = .Red
            g = .Green
            b = .Blue
            a = .Alpha
        End With
        
        pSort(i).pOrigIndex = i
        
        If (channelIndex = 0) Then
            pSort(i).pSortCriteria = CSng(r) * multTable(3) + CSng(g) * multTable(2) + CSng(b) * multTable(1) + CSng(a) * multTable(0)
        ElseIf (channelIndex = 1) Then
            pSort(i).pSortCriteria = CSng(g) * multTable(3) + CSng(b) * multTable(2) + CSng(r) * multTable(1) + CSng(a) * multTable(0)
        ElseIf (channelIndex = 2) Then
            pSort(i).pSortCriteria = CSng(b) * multTable(3) + CSng(r) * multTable(2) + CSng(g) * multTable(1) + CSng(a) * multTable(0)
        Else
            pSort(i).pSortCriteria = CSng(a) * multTable(3) + CSng(g) * multTable(2) + CSng(r) * multTable(1) + CSng(b) * multTable(0)
        End If
        
    Next i
    
    FinishSort pSort
    
End Sub

'When removing duplicates from a large palette, it will go much faster if you first sort the array
' into its literal RGB order (e.g. from lowest combined long value to highest).  Then you can remove
' duplicate entries by simply checking neighbors.
Friend Sub SortFixedOrder()

    'We don't actually use true "luminance" for the sort; instead, we treat RGB coordinates as literal
    ' positions within a 3D RGB cube, and we use their positions to sort along the cube's diagonal.
    ' (Note that alpha is *not* considered during the sort, by design.)
    Dim pSort() As PaletteSort
    ReDim pSort(0 To m_NumOfColors - 1) As PaletteSort
    
    Dim r As Long, g As Long, b As Long
    
    Dim i As Long
    For i = 0 To m_NumOfColors - 1
        
        With m_Palette(i).ColorValue
            r = .Red
            g = .Green
            b = .Blue
        End With
        
        pSort(i).pOrigIndex = i
        pSort(i).pSortCriteria = RGB(r, g, b)
        
    Next i
    
    FinishSort pSort
    
End Sub

'When constructing palette matching trees, it is often useful to sort a palette by luminance
' (e.g. along the diagonal of the RGB cube).  This makes it easier to produce a balanced nearest-neighbor
' structure for color-matching.
Friend Sub SortByLuminance()

    'We don't actually use true "luminance" for the sort; instead, we treat RGB coordinates as literal
    ' positions within a 3D RGB cube, and we use their positions to sort along the cube's diagonal.
    ' (Note that alpha is *not* considered during the sort, by design.)
    Dim pSort() As PaletteSort
    ReDim pSort(0 To m_NumOfColors - 1) As PaletteSort
    
    Dim r As Long, g As Long, b As Long
    
    Dim i As Long
    For i = 0 To m_NumOfColors - 1
        
        With m_Palette(i).ColorValue
            r = .Red
            g = .Green
            b = .Blue
        End With
        
        pSort(i).pOrigIndex = i
        pSort(i).pSortCriteria = r * r + g * g + b * b
        
    Next i
    
    FinishSort pSort
    
End Sub

'After assembling a PaletteSort array (which contains sort keys for palette entries), call this function
' assemble a new palette based on that sort criteria.
Private Sub FinishSort(ByRef srcPaletteSort() As PaletteSort)
    
    If (m_NumOfColors <= 1) Then Exit Sub
    
    'We can now use the sort criteria to perform a normal quicksort
    SortPalette srcPaletteSort
    
    'The pSort() array has now been sorted according to its pSortCriteria contents.  We want to rebuild
    ' our palette array in a matching order.
    Dim newPalette() As PDPaletteEntry
    ReDim newPalette(0 To m_NumOfColors - 1) As PDPaletteEntry
    
    Dim i As Long
    For i = 0 To m_NumOfColors - 1
        newPalette(i) = m_Palette(srcPaletteSort(i).pOrigIndex)
    Next i
    
    'Release the old array in favor of the newly sorted one
    m_Palette = newPalette
    
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
