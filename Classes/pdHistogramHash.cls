VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdHistogramHash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'Fast Histogram class (focused on RGB/A support), using hash-table + separate chaining (linked-list) for collisions
'Copyright 2022-2026 by Tanner Helland
'Created: 07/March/22
'Last updated: 22/March/22
'Last update: fix bug on properly resetting overflow index after table resize
'
'A number of RGB/A-based image processing functions can be accelerated using histograms.
' For RGB images, a naive 16-million-color array can be used to track unique colors, but this
' naive approach breaks down for RGBA colors (full 4GB possibility space, argh).
'
'Similarly, if you need to do something like generate a list of unique histogram entries
' (i.e. when generating an optimized palette), the naive approach is extremely slow to iterate,
' especially when sparsely populated.
'
'This class exists as a specialized hash table to both accelerate *and* minimize memory usage of
' histogram-driven tasks.  It is very fast to populate and very fast to retrieve a full list of
' unique entries (including their counts, a critical component of the whole "histogram" thing).
'
'Collisions are resolved using linked list indices into a dedicated overflow array.  This works well
' for cache locality but is relatively naive in implementation - a perf-friendly trade-off for
' most VB6 implementations, generally speaking.
'
'This class will dynamically resize the hash table whenever the overflow table fills. By default,
' the overflow table is set to the same size as the hash table, which means that when the overflow
' fills and a resize is triggered, both the hash table *and* the overflow table will double in size.
' Powers-of-two are used for table size.  This is not ideal from a hashing standpoint (primes hash
' better) but it's a great fit for VB because we can use simple AND masks for table assignment,
' improving performance further.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Histogram entry.  Colors are always type LONG (even if RGB, not RGBA).
Private Type MergedEntry
    hstColor As Long    'RGB/A color
    hstCount As Long    'Number of occurrences
    idxNext As Long     'Index to the next entry in the list (always in the overflow table)
End Type

'Initial table size is effectively arbitrary.  65k entries at 12-bytes-per-entry is about
' ~1.5 MB of space, which is a good compromise between tiny images (which won't need that
' much space) and extremely color-dense images (which will only need to resize the table
' a handful of times).
Private Const INIT_TABLE_SIZE As Long = 2 ^ 16

'Hash table and mask that lets us map into the table.  Mask must always be of the form 2^n-1.
' (A new mask must be generated whenever the table is resized.)
Private HASH_TABLE_MASK As Long
Private m_hashTable() As MergedEntry

'Hash collisons are resolved by placement into an overflow table, which expands linearly.
' The overflow table size is currently set to always match the size of the hash table;
' this simplifies resize operations, and lends itself to good table coverage.
Private m_overflowTable() As MergedEntry

'Current index into the next available position in the overflow table.
Private m_idxOverflow As Long

Friend Sub AddColor(ByVal srcColor As Long, Optional ByVal srcColorCount As Long = 1)
    
    'Generate a (very cheap) hash for the incoming color.  This formula deliberately weights
    ' alpha less than RGB values (2 ^ 9, 2 ^ 25) and typically results in ~80% hash table
    ' coverage before a resize is required.
    Dim idxTable As Long
    idxTable = srcColor Xor (srcColor \ 512) Xor (srcColor \ 33554432)
    idxTable = idxTable And HASH_TABLE_MASK
    
    'Probe the initial hash table
    If (m_hashTable(idxTable).hstCount = 0) Then
    
        'Empty table position; initialize against this color
        With m_hashTable(idxTable)
            .hstColor = srcColor
            .hstCount = srcColorCount
        End With
    
    'Table position occupied
    Else
        
        'Look for a matching color
        If (m_hashTable(idxTable).hstColor = srcColor) Then
        
            'Match!  Increment count and exit
            m_hashTable(idxTable).hstCount = m_hashTable(idxTable).hstCount + srcColorCount
            
        'Color mismatch
        Else
            
            'If this table position is occupied *and* the color doesn't match,
            ' we need to move into the overflow table.
            
            'See if a linked list has already been initialized for this table entry.
            If (m_hashTable(idxTable).idxNext = 0) Then
                
                'Place this color as a new entry in the overflow table.
                m_hashTable(idxTable).idxNext = m_idxOverflow
                m_overflowTable(m_idxOverflow).hstColor = srcColor
                m_overflowTable(m_idxOverflow).hstCount = srcColorCount
                m_idxOverflow = m_idxOverflow + 1
                If (m_idxOverflow > UBound(m_overflowTable)) Then IncreaseTableSize
                
            Else
            
                'Continue probing entries until we find a match or an empty place in the overflow table
                idxTable = m_hashTable(idxTable).idxNext
                
                Do
                    
                    If (m_overflowTable(idxTable).hstColor = srcColor) Then
                        
                        'Colors match!  Increment count and exit
                        m_overflowTable(idxTable).hstCount = m_overflowTable(idxTable).hstCount + srcColorCount
                        Exit Sub
                        
                    Else
                        
                        'If this is the end of the linked list, add this entry to the table
                        If (m_overflowTable(idxTable).idxNext = 0) Then
                            m_overflowTable(idxTable).idxNext = m_idxOverflow
                            m_overflowTable(m_idxOverflow).hstColor = srcColor
                            m_overflowTable(m_idxOverflow).hstCount = srcColorCount
                            m_idxOverflow = m_idxOverflow + 1
                            If (m_idxOverflow > UBound(m_overflowTable)) Then IncreaseTableSize
                            Exit Sub
                        
                        'More colors to probe; reassign the table index, and let the loop continue naturally
                        Else
                            idxTable = m_overflowTable(idxTable).idxNext
                        End If
                    
                    End If
                    
                Loop
                
            End If
            
        End If
        
    End If
        
End Sub

'Count total number of unique entries in the table (corresponds to RGBA color count)
Friend Function GetNumUniqueEntries() As Long
    
    'This step is extremely simple because we don't actually need to traverse anything linked-list-style.
    
    'Instead, simply probe the hash table and count all non-zero counts.
    Dim i As Long
    For i = 0 To UBound(m_hashTable)
        If (m_hashTable(i).hstCount > 0) Then GetNumUniqueEntries = GetNumUniqueEntries + 1
    Next i
    
    'Next, we *don't need to iterate* the overflow table.
    
    'Because it's a linear table, we are guaranteed that each entry in the table is unique!
    ' Just add the overflow count to the hash table count.
    GetNumUniqueEntries = GetNumUniqueEntries + (m_idxOverflow - 1)

End Function

'Retrieve the list of collected colors (and counts) in convenient RGBQuad (color) and Long (count) arrays.
' ALSO - this is important - return the number of colors in the arrays (same color count for both, obviously).
' This is important because the returned array dimensions may not precisely match the final color count.
' This is a perf optimization that allows us to allocate each array just once.
Friend Function GetUniqueColors(ByRef dstQuadArray() As RGBQuad, dstCounts() As Long) As Long
    
    'Set each target array to a guaranteed "safe" size.
    GetUniqueColors = Me.GetNumUniqueEntries()
    
    ReDim dstQuadArray(0 To GetUniqueColors - 1) As RGBQuad
    ReDim dstCounts(0 To GetUniqueColors - 1) As Long
    
    'Wrap a fake wrapper around the RGBQuad array; this lets us use direct assignment from Long -> RGBQuad
    Dim fakeLongArray() As Long, tmpSA As SafeArray1D
    VBHacks.WrapArrayAroundPtr_Long fakeLongArray, tmpSA, VarPtr(dstQuadArray(0)), GetUniqueColors * 4
    
    'Iterate the list, copying relevant items into each destination array
    Dim i As Long, idxDst As Long
    For i = 0 To UBound(m_hashTable)
        If (m_hashTable(i).hstCount > 0) Then
            fakeLongArray(idxDst) = m_hashTable(i).hstColor
            dstCounts(idxDst) = m_hashTable(i).hstCount
            idxDst = idxDst + 1
        End If
    Next i
    
    For i = 0 To m_idxOverflow - 2
        fakeLongArray(idxDst) = m_overflowTable(i).hstColor
        dstCounts(idxDst) = m_overflowTable(i).hstCount
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
    Dim tmpHash() As MergedEntry, tmpOverflow() As MergedEntry
    ReDim tmpHash(0 To UBound(m_hashTable)) As MergedEntry
    ReDim tmpOverflow(0 To UBound(m_overflowTable)) As MergedEntry
    CopyMemoryStrict VarPtr(tmpHash(0)), VarPtr(m_hashTable(0)), (UBound(m_hashTable) + 1) * 12
    CopyMemoryStrict VarPtr(tmpOverflow(0)), VarPtr(m_overflowTable(0)), (UBound(m_overflowTable) + 1) * 12
    
    'Calculate new table sizes, then increase the main hash and overflow tables to match
    Dim newTableSize As Long
    newTableSize = (HASH_TABLE_MASK + 1) * 2
    HASH_TABLE_MASK = newTableSize - 1
    
    ReDim m_hashTable(0 To newTableSize - 1) As MergedEntry
    ReDim m_overflowTable(0 To newTableSize - 1) As MergedEntry
    m_idxOverflow = 1
    
    'Re-add all items to the new, larger hash table
    Dim i As Long
    For i = 0 To UBound(tmpHash)
        If (tmpHash(i).hstCount > 0) Then Me.AddColor tmpHash(i).hstColor, tmpHash(i).hstCount
    Next i
    
    'By definition, the overflow table was full prior to this resize, so we don't need to check for
    ' non-zero count values. (Note also that we start at position 1 because 0 is a reserved value
    ' indicating "no linked entry".)
    For i = 1 To UBound(tmpOverflow)
        Me.AddColor tmpOverflow(i).hstColor, tmpOverflow(i).hstCount
    Next i
    
End Sub

Private Sub Class_Initialize()
    
    'Create the initial table(s) and bit-mask
    ReDim m_hashTable(0 To INIT_TABLE_SIZE - 1) As MergedEntry
    ReDim m_overflowTable(0 To INIT_TABLE_SIZE - 1) As MergedEntry
    HASH_TABLE_MASK = INIT_TABLE_SIZE - 1
    
    '0 is used to denote "no children", so ensure the overflow index starts at 1
    m_idxOverflow = 1
    
End Sub
