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
'Copyright 2022-2022 by Tanner Helland
'Created: 07/March/22
'Last updated: 08/March/22
'Last update: wrap up initial build
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
' VB6 implementations, generally speaking.
'
'Note also that this class does *not* currently dynamically resize the hash table.  It uses a
' fixed size based on average photo color counts.  Dynamic resizing could be helpful on extreme
' outliers and I may tackle this someday.
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

'Initial hash table, fixed size.  64k is effectively arbitrary; it's a decent starting size for
' average RGBA images.
Private Const HASH_TABLE_MASK As Long = &HFFFF&
Private m_hashTable(0 To 65535) As MergedEntry

'Hash collisons are resolved by placement into an overflow table, which expands linearly.
' Initial size is arbitrary; resizes are expensive so we try to minimize resize attempts,
' but note that each entry is 12-bytes long so we don't want to make it *too* large.
Private Const INIT_OVERFLOW_SIZE As Long = 65536
Private m_OverflowTable() As MergedEntry

'Current index into the next available position in the overflow table.
Private m_idxOverflow As Long

Friend Sub AddColor(ByVal srcColor As Long)
    
    'Generate a (very cheap) hash for the incoming color.  This formula deliberately weights
    ' alpha less than RGB values (2 ^ 9, 2 ^ 25) and performs very well on small images
    Dim idxTable As Long
    idxTable = srcColor Xor (srcColor \ 512) Xor (srcColor \ 33554432)
    idxTable = idxTable And HASH_TABLE_MASK
    
    'Probe the initial hash table
    If (m_hashTable(idxTable).hstCount = 0) Then
    
        'Empty table position; initialize against this color
        With m_hashTable(idxTable)
            .hstColor = srcColor
            .hstCount = 1
        End With
    
    'Table position occupied
    Else
        
        'Look for a matching color
        If (m_hashTable(idxTable).hstColor = srcColor) Then
        
            'Match!  Increment count and exit
            m_hashTable(idxTable).hstCount = m_hashTable(idxTable).hstCount + 1
            
        'Color mismatch
        Else
            
            'If this table position is occupied *and* the color doesn't match,
            ' we need to move into the overflow table.
            
            'See if a linked list has already been initialized for this table entry.
            If (m_hashTable(idxTable).idxNext = 0) Then
                
                'Place this color as a new entry in the overflow table.
                m_hashTable(idxTable).idxNext = m_idxOverflow
                If (m_idxOverflow > UBound(m_OverflowTable)) Then ReDim Preserve m_OverflowTable(0 To m_idxOverflow * 2 - 1) As MergedEntry
                m_OverflowTable(m_idxOverflow).hstColor = srcColor
                m_OverflowTable(m_idxOverflow).hstCount = 1
                m_idxOverflow = m_idxOverflow + 1
                
            Else
            
                'Continue probing entries until we find a match or an empty place in the overflow table
                idxTable = m_hashTable(idxTable).idxNext
                
                Do
                    
                    If (m_OverflowTable(idxTable).hstColor = srcColor) Then
                        
                        'Colors match!  Increment count and exit
                        m_OverflowTable(idxTable).hstCount = m_OverflowTable(idxTable).hstCount + 1
                        Exit Sub
                        
                    Else
                        
                        'If this is the end of the linked list, add this entry to the table
                        If (m_OverflowTable(idxTable).idxNext = 0) Then
                            m_OverflowTable(idxTable).idxNext = m_idxOverflow
                            If (m_idxOverflow > UBound(m_OverflowTable)) Then ReDim Preserve m_OverflowTable(0 To m_idxOverflow * 2 - 1) As MergedEntry
                            m_OverflowTable(m_idxOverflow).hstColor = srcColor
                            m_OverflowTable(m_idxOverflow).hstCount = 1
                            m_idxOverflow = m_idxOverflow + 1
                            Exit Sub
                        
                        'More colors to probe; reassign the table index, and let the loop continue naturally
                        Else
                            idxTable = m_OverflowTable(idxTable).idxNext
                        End If
                    
                    End If
                    
                Loop
                
                'NOP - the loop is guaranteed to terminate internally, so this line should never trigger
                Exit Sub
                
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
        fakeLongArray(idxDst) = m_OverflowTable(i).hstColor
        dstCounts(idxDst) = m_OverflowTable(i).hstCount
        idxDst = idxDst + 1
    Next i
    
    'Free the unsafe array wrapper
    VBHacks.UnwrapArrayFromPtr_Long fakeLongArray
    
End Function

Private Sub Class_Initialize()
    
    ReDim m_OverflowTable(0 To INIT_OVERFLOW_SIZE - 1) As MergedEntry
    
    '0 is used to denote "no children", so ensure the overflow index starts at 1
    m_idxOverflow = 1
    
End Sub
