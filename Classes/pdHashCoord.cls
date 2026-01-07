VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdHashCoord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'Fast Histogram class (focused on RGB/A support), using hash-table + separate chaining (linked-list) for collisions
'Copyright 2022-2026 by Tanner Helland
'Created: 07/March/22
'Last updated: 14/June/22
'Last update: split off a new coordinate-driven variation from the original pdHistogramHash class
'
'This class is derived from the pdHistogramHash class.  Look there for full details.  The main change
' here is that this class stores colors alongside coordinates; PD uses this during 3D LUT generation to
' map best-fit colors from their original values to their final (modified) values.  Original colors are
' matched, and the saved coordinates are used to find the "new" color at that same position.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Table entry.  Color data is stored alongside coordinate data, which allows you to look up a color and
' retrieve a corresponding coordinate for it.
Private Type CoordEntry
    tblColor As Long
    tblCoord As PointLong
    idxNext As Long
End Type

'Initial table size is effectively arbitrary.  65k entries at 16-bytes-per-entry is about
' ~2 MB of space, which is a good compromise between tiny images (which won't need that
' much space) and extremely color-dense images (which will only need to resize the table
' a handful of times).
Private Const INIT_TABLE_SIZE As Long = 2 ^ 16

'Hash table and mask that lets us map into the table.  Mask must always be of the form 2^n-1.
' (A new mask must be generated whenever the table is resized.)
Private HASH_TABLE_MASK As Long
Private m_hashTable() As CoordEntry

'Hash collisons are resolved by placement into an overflow table, which expands linearly.
' The overflow table size is currently set to always match the size of the hash table;
' this simplifies resize operations, and lends itself to good table coverage.
Private m_overflowTable() As CoordEntry

'Current index into the next available position in the overflow table.
Private m_idxOverflow As Long

'Returns TRUE if the added color is unique, FALSE otherwise
Friend Function AddColor(ByVal srcColor As Long, ByRef srcCoord As PointLong) As Boolean
    
    AddColor = True
    
    'Generate a (very cheap) hash for the incoming color.  This formula deliberately weights
    ' alpha less than RGB values (2 ^ 9, 2 ^ 25) and typically results in ~80% hash table
    ' coverage before a resize is required.
    Dim idxTable As Long
    idxTable = srcColor Xor (srcColor \ 512) Xor (srcColor \ 33554432)
    idxTable = idxTable And HASH_TABLE_MASK
    
    'Probe the initial hash table
    If (m_hashTable(idxTable).tblCoord.x = LONG_MAX) Then
    
        'Empty table position; initialize against this color
        With m_hashTable(idxTable)
            .tblColor = srcColor
            .tblCoord = srcCoord
        End With
    
    'Table position occupied
    Else
        
        'Look for a matching color
        If (m_hashTable(idxTable).tblColor = srcColor) Then
        
            'Match!  No need to do anything.
            ' (You could overwrite the coordinate, I suppose, but the important thing to us is that
            ' we simply have *a* coordinate matching *this* color - which coordinate is irrelevant.)
            AddColor = False
            
        'Color mismatch
        Else
            
            'If this table position is occupied *and* the color doesn't match,
            ' we need to move into the overflow table.
            
            'See if a linked list has already been initialized for this table entry.
            If (m_hashTable(idxTable).idxNext = 0) Then
                
                'Place this color as a new entry in the overflow table.
                m_hashTable(idxTable).idxNext = m_idxOverflow
                m_overflowTable(m_idxOverflow).tblColor = srcColor
                m_overflowTable(m_idxOverflow).tblCoord = srcCoord
                m_idxOverflow = m_idxOverflow + 1
                If (m_idxOverflow > UBound(m_overflowTable)) Then IncreaseTableSize
                
            Else
            
                'Continue probing entries until we find a match or an empty place in the overflow table
                idxTable = m_hashTable(idxTable).idxNext
                
                Do
                    
                    If (m_overflowTable(idxTable).tblColor = srcColor) Then
                        
                        'Colors match!  No need to update this entry.
                        AddColor = False
                        Exit Function
                        
                    Else
                        
                        'If this is the end of the linked list, add this entry to the table
                        If (m_overflowTable(idxTable).idxNext = 0) Then
                            m_overflowTable(idxTable).idxNext = m_idxOverflow
                            m_overflowTable(m_idxOverflow).tblColor = srcColor
                            m_overflowTable(m_idxOverflow).tblCoord = srcCoord
                            m_idxOverflow = m_idxOverflow + 1
                            If (m_idxOverflow > UBound(m_overflowTable)) Then IncreaseTableSize
                            Exit Function
                        
                        'More colors to probe; reassign the table index, and let the loop continue naturally
                        Else
                            idxTable = m_overflowTable(idxTable).idxNext
                        End If
                    
                    End If
                    
                Loop
                
            End If
            
        End If
        
    End If
        
End Function

'Given a particular color, retrieve the matching coordinate, if any.
' Returns TRUE if color was found; FALSE otherwise.
Friend Function GetCoordForColor(ByVal srcColor As Long, ByRef dstCoord As PointLong) As Boolean
    
    GetCoordForColor = False
    
    'Obviously, we must use the same hash as the AddColor function, above.
    Dim idxTable As Long
    idxTable = srcColor Xor (srcColor \ 512) Xor (srcColor \ 33554432)
    idxTable = idxTable And HASH_TABLE_MASK
    
    'Probe the initial hash table
    If (m_hashTable(idxTable).tblColor = srcColor) Then
        
        'Match!  Return the stored coordinate and exit immediately
        GetCoordForColor = True
        dstCoord = m_hashTable(idxTable).tblCoord
        Exit Function
    
    'Color mismatch
    Else
        
        'Move into the overflow table if we can.
        If (m_hashTable(idxTable).idxNext = 0) Then
            
            'No overflow index exists for this entry, thus the color doesn't exist in the table.
            Exit Function
        
        '1+ entries exist in the overflow table
        Else
        
            'Continue probing entries until we find a match or an empty place in the overflow table
            idxTable = m_hashTable(idxTable).idxNext
            
            Do
                
                If (m_overflowTable(idxTable).tblColor = srcColor) Then
                    
                    'Colors match!  Return the stored coord.
                    GetCoordForColor = True
                    dstCoord = m_overflowTable(idxTable).tblCoord
                    Exit Function
                    
                Else
                    
                    'If this is the end of the linked list, the color doesn't exist
                    If (m_overflowTable(idxTable).idxNext = 0) Then
                        Exit Function
                    
                    'More colors to probe; reassign the table index, and let the loop continue naturally
                    Else
                        idxTable = m_overflowTable(idxTable).idxNext
                    End If
                
                End If
                
            Loop
            
        End If
    
    End If
    
End Function

'Count total number of unique entries in the table (corresponds to RGBA color count)
Friend Function GetNumUniqueEntries() As Long
    
    'This step is extremely simple because we don't actually need to traverse anything linked-list-style.
    
    'Instead, simply probe the hash table and count all non-zero counts.
    Dim i As Long
    For i = 0 To UBound(m_hashTable)
        If (m_hashTable(i).tblCoord.x <> LONG_MAX) Then GetNumUniqueEntries = GetNumUniqueEntries + 1
    Next i
    
    'Next, we *don't need to iterate* the overflow table.
    
    'Because it's a linear table, we are guaranteed that each entry in the table is unique!
    ' Just add the overflow count to the hash table count.
    GetNumUniqueEntries = GetNumUniqueEntries + (m_idxOverflow - 1)

End Function

'Retrieve the list of collected colors (*no* coordinates) in a convenient RGBQuad array.
' ALSO - this is important - return the number of colors.  The returned array dimension may not precisely match
' the final color count.  This is a perf optimization that allows us to allocate each array just once.
Friend Function GetUniqueColors(ByRef dstQuadArray() As RGBQuad) As Long
    
    'Set the target array to a guaranteed "safe" size.
    GetUniqueColors = Me.GetNumUniqueEntries()
    ReDim dstQuadArray(0 To GetUniqueColors - 1) As RGBQuad
    
    'Wrap a fake wrapper around the RGBQuad array; this lets us use direct assignment from Long -> RGBQuad
    Dim fakeLongArray() As Long, tmpSA As SafeArray1D
    VBHacks.WrapArrayAroundPtr_Long fakeLongArray, tmpSA, VarPtr(dstQuadArray(0)), GetUniqueColors * 4
    
    'Iterate the list, copying relevant items into the destination array
    Dim i As Long, idxDst As Long
    For i = 0 To UBound(m_hashTable)
        If (m_hashTable(i).tblCoord.x <> LONG_MAX) Then
            fakeLongArray(idxDst) = m_hashTable(i).tblColor
            idxDst = idxDst + 1
        End If
    Next i
    
    For i = 0 To m_idxOverflow - 2
        fakeLongArray(idxDst) = m_overflowTable(i).tblColor
        idxDst = idxDst + 1
    Next i
    
    'Free the unsafe array wrapper
    VBHacks.UnwrapArrayFromPtr_Long fakeLongArray
    
End Function

'Retrieve the list of collected colors (and coordinates) in convenient RGBQuad (color) and Long (count) arrays.
' ALSO - this is important - return the number of colors in the arrays (same color count for both, obviously).
' This is important because the returned array dimensions may not precisely match the final color count.
' This is a perf optimization that allows us to allocate each array just once.
Friend Function GetUniqueColorsAndCoords(ByRef dstQuadArray() As RGBQuad, ByRef dstCoords() As PointLong) As Long
    
    'Set each target array to a guaranteed "safe" size.
    GetUniqueColorsAndCoords = Me.GetNumUniqueEntries()
    
    ReDim dstQuadArray(0 To GetUniqueColorsAndCoords - 1) As RGBQuad
    ReDim dstCoords(0 To GetUniqueColorsAndCoords - 1) As PointLong
    
    'Wrap a fake wrapper around the RGBQuad array; this lets us use direct assignment from Long -> RGBQuad
    Dim fakeLongArray() As Long, tmpSA As SafeArray1D
    VBHacks.WrapArrayAroundPtr_Long fakeLongArray, tmpSA, VarPtr(dstQuadArray(0)), GetUniqueColorsAndCoords * 4
    
    'Iterate the list, copying relevant items into each destination array
    Dim i As Long, idxDst As Long
    For i = 0 To UBound(m_hashTable)
        If (m_hashTable(i).tblCoord.x <> LONG_MAX) Then
            fakeLongArray(idxDst) = m_hashTable(i).tblColor
            dstCoords(idxDst) = m_hashTable(i).tblCoord
            idxDst = idxDst + 1
        End If
    Next i
    
    For i = 0 To m_idxOverflow - 2
        fakeLongArray(idxDst) = m_overflowTable(i).tblColor
        dstCoords(idxDst) = m_overflowTable(i).tblCoord
        idxDst = idxDst + 1
    Next i
    
    'Free the unsafe array wrapper
    VBHacks.UnwrapArrayFromPtr_Long fakeLongArray
    
End Function

'This function imposes a large performance penalty.  *Please* call it sparingly!
Private Sub IncreaseTableSize()
    
    'If we're here, it means we've run out of space in the overflow table.
    ' (In the current implementation, the hash and overflow tables are always identically sized.
    '  If the overflow table overflows, we double the size of *both* the hash table and the
    '  overflow table, then re-add all existing elements.)
    
    'Start by backing up the existing hash tables into temporary arrays
    Dim tmpHash() As CoordEntry, tmpOverflow() As CoordEntry
    ReDim tmpHash(0 To UBound(m_hashTable)) As CoordEntry
    ReDim tmpOverflow(0 To UBound(m_overflowTable)) As CoordEntry
    CopyMemoryStrict VarPtr(tmpHash(0)), VarPtr(m_hashTable(0)), (UBound(m_hashTable) + 1) * LenB(tmpHash(0))
    CopyMemoryStrict VarPtr(tmpOverflow(0)), VarPtr(m_overflowTable(0)), (UBound(m_overflowTable) + 1) * LenB(tmpOverflow(0))
    
    'Calculate new table sizes, then increase the main hash and overflow tables to match
    Dim newTableSize As Long
    newTableSize = (HASH_TABLE_MASK + 1) * 2
    HASH_TABLE_MASK = newTableSize - 1
    
    ReDim m_hashTable(0 To newTableSize - 1) As CoordEntry
    ReDim m_overflowTable(0 To newTableSize - 1) As CoordEntry
    m_idxOverflow = 1
    
    'Whenever the table is resized, we need to flag new entries as "unused".  (Note that we only do this
    ' to the base hash table; the overflow table is traversed differently, according to each linked list's
    ' "next item" value, which must be non-zero by definition if an entry is valid.)
    Dim i As Long
    For i = 0 To UBound(m_hashTable)
        m_hashTable(i).tblCoord.x = LONG_MAX
    Next i
    
    'Re-add all items to the new, larger hash table
    For i = 0 To UBound(tmpHash)
        If (tmpHash(i).tblCoord.x <> LONG_MAX) Then Me.AddColor tmpHash(i).tblColor, tmpHash(i).tblCoord
    Next i
    
    'By definition, the overflow table was full prior to this resize, so we don't need to check for
    ' non-zero count values. (Note also that we start at position 1 because 0 is a reserved value
    ' indicating "no linked entry".)
    For i = 1 To UBound(tmpOverflow)
        Me.AddColor tmpOverflow(i).tblColor, tmpOverflow(i).tblCoord
    Next i
    
End Sub

Private Sub Class_Initialize()
    
    'Create the initial table(s) and bit-mask
    ReDim m_hashTable(0 To INIT_TABLE_SIZE - 1) As CoordEntry
    ReDim m_overflowTable(0 To INIT_TABLE_SIZE - 1) As CoordEntry
    HASH_TABLE_MASK = INIT_TABLE_SIZE - 1
    
    'Whenever the table is resized, we need to flag new entries as "unused".  (Note that we only do this
    ' to the base hash table; the overflow table is traversed differently, according to each linked list's
    ' "next item" value, which must be non-zero by definition if an entry is valid.)
    Dim i As Long
    For i = 0 To UBound(m_hashTable)
        m_hashTable(i).tblCoord.x = LONG_MAX
    Next i
    
    '0 is used to denote "no children", so ensure the overflow index starts at 1
    m_idxOverflow = 1
    
End Sub
