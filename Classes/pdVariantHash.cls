VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdVariantHash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'Fast Variant-based collection, using hash-table + separate chaining (linked-list) for collisions
'Copyright 2024-2026 by Tanner Helland
'Created: 09/September/24
'Last updated: 09/September/24
'Last update: split off from string-based class; I'm keeping that for localized text performance,
'             but this class will be more flexible thanks to varied options for storage types
'
'This class is a direct copy of the pdStringHash class, with strings swapped out for variants.
' This provides additional options for data storage at a moderate hit to performance.
'
'PhotoDemon currently uses this class in the customizable hotkeys UI to map action IDs across
' menu and hotkey arrays.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit
Option Compare Binary

'Item entry in the table.  Key and value are stored together, alongside a link to the next item in this chain.
Private Type MergedEntry
    sKey As String
    sItem As Variant
    idxNext As Long     'Index to the next entry in the list (always in the overflow table); 0 if unlinked
End Type

'Initial table size is effectively arbitrary.  In a given session, PD is guaranteed to query at least
' ~600 translations (as measured on a cold start against v9.0 nightly builds), and probably unlikely
' to query more than ~1200 translations (if every on-canvas tool is accessed), so a table size of 512
' works well while ensuring that even in a worst-case scenario (every translation in the program being
' touched in a single session) we only need a few table resizes.
'
'For reference, PD currently ships ~2600 phrases, but remember that this hash table size applies to
' *both* the base hash table itself AND a same-sized overflow table.  Base hash table occupancy rates
' are ~80+% before the average resize event, so the overflow table must exceed the INIT_TABLE_SIZE
' before both table sizes double - which means that a base size of 512 entries guarantees no table
' resizes on a cold start (1024 slots with ~90% coverage across both), and a worst-case scenario of
' two table resizes if a user somehow triggers every translation in the project.
'
'(Finally, all this math doesn't really matter because table resizes are extremely fast, but we try
' to minimize memory allocations as a general rule!)
Private Const INIT_TABLE_SIZE As Long = 2 ^ 9

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

'Add a Key+Value pair to the table.
'
'Returns TRUE if added successfully; FALSE otherwise.  (FALSE is not currently returned under any circumstances.)
Friend Function AddItem(ByRef srcKey As String, ByVal srcItem As Variant) As Boolean
    
    'Assume successful addition
    AddItem = True
    
    'Always start by retrieving the hash index
    Dim idxTable As Long
    idxTable = GetKeyHash(srcKey)
    
    'Probe the initial hash table entry and use it if possible
    If (LenB(m_hashTable(idxTable).sKey) = 0) Then
        m_hashTable(idxTable).sKey = srcKey
        m_hashTable(idxTable).sItem = srcItem
        
    'Table position occupied.  Search the overflow table.
    Else
        
        'Look for a matching key, and do nothing if a match is found
        If (m_hashTable(idxTable).sKey = srcKey) Then
            
            'Key already exists in the table.  Overwrite it with the new value, then exit.
            m_hashTable(idxTable).sItem = srcItem
            
        'Key mismatch
        Else
            
            'If this table position is occupied *and* the string doesn't match,
            ' we need to move into the overflow table.
            
            'See if a linked list has already been initialized for this table entry.
            If (m_hashTable(idxTable).idxNext = 0) Then
                
                'Place this string as a new entry in the overflow table.
                m_hashTable(idxTable).idxNext = m_idxOverflow
                m_overflowTable(m_idxOverflow).sKey = srcKey
                m_overflowTable(m_idxOverflow).sItem = srcItem
                m_idxOverflow = m_idxOverflow + 1
                If (m_idxOverflow > UBound(m_overflowTable)) Then IncreaseTableSize
                
            Else
            
                'Continue probing entries until we find a match or an empty place in the overflow table
                idxTable = m_hashTable(idxTable).idxNext
                
                Do
                    
                    'Key already exists in the table.  Overwrite it with the new value, then exit.
                    If (m_overflowTable(idxTable).sKey = srcKey) Then
                        m_overflowTable(idxTable).sItem = srcItem
                        Exit Function
                        
                    Else
                        
                        'This is the end of the linked list; add this entry to the table
                        If (m_overflowTable(idxTable).idxNext = 0) Then
                            m_overflowTable(idxTable).idxNext = m_idxOverflow
                            m_overflowTable(m_idxOverflow).sKey = srcKey
                            m_overflowTable(m_idxOverflow).sItem = srcItem
                            m_idxOverflow = m_idxOverflow + 1
                            If (m_idxOverflow > UBound(m_overflowTable)) Then IncreaseTableSize
                            Exit Function
                        
                        'More strings to probe; reassign the table index, and let the loop continue naturally
                        Else
                            idxTable = m_overflowTable(idxTable).idxNext
                        End If
                    
                    End If
                    
                Loop
                
            End If
            
        End If
        
    End If
        
End Function

'Return the stored item for a given key.  Returns TRUE if key exists, FALSE otherwise.
Friend Function GetItemByKey(ByRef srcKey As String, ByRef dstItem As Variant) As Boolean
    
    'Assume successful retrieval
    GetItemByKey = True
    
    'Always start by retrieving the hash index
    Dim idxTable As Long
    idxTable = GetKeyHash(srcKey)
    
    'Probe the initial hash table entry and use it if possible
    If (LenB(m_hashTable(idxTable).sKey) = 0) Then
        
        'Key doesn't exist
        GetItemByKey = False
        
    'Table position occupied.  Compare keys, and if that fails, search the overflow table.
    Else
        
        If (m_hashTable(idxTable).sKey = srcKey) Then
            
            'Match!
            dstItem = m_hashTable(idxTable).sItem
            GetItemByKey = True
            
        'Key mismatch
        Else
            
            'This table position is occupied *and* the string doesn't match,
            ' so we need to move into the overflow table.
            
            'See if a linked list has already been initialized for this table entry.
            If (m_hashTable(idxTable).idxNext = 0) Then
                
                'No linked list for this entry.  Return failure.
                GetItemByKey = False
                
            Else
            
                'Continue probing entries until we find a match or the end of this list
                idxTable = m_hashTable(idxTable).idxNext
                
                Do
                    
                    'Match!
                    If (m_overflowTable(idxTable).sKey = srcKey) Then
                        dstItem = m_overflowTable(idxTable).sItem
                        GetItemByKey = True
                        Exit Function
                    
                    'Mismatch
                    Else
                        
                        'If we reach the end of the linked list with no match, return failure.
                        If (m_overflowTable(idxTable).idxNext = 0) Then
                            GetItemByKey = False
                            Exit Function
                        
                        'More strings to probe; reassign the table index, and let the loop continue naturally
                        Else
                            idxTable = m_overflowTable(idxTable).idxNext
                        End If
                    
                    End If
                    
                Loop
                
            End If
            
        End If
        
    End If
    
End Function

Friend Function GetNumOfItems() As Long
    
    GetNumOfItems = 0
    
    Dim i As Long
    For i = 0 To UBound(m_hashTable)
        If (LenB(m_hashTable(i).sKey) <> 0) Then GetNumOfItems = GetNumOfItems + 1
    Next i
    
    If (m_idxOverflow > 1) Then GetNumOfItems = GetNumOfItems + (m_idxOverflow - 1)
    
End Function

'Erase the table and restore its default configuration
Friend Sub Reset()

    'Create the initial table(s) and bit-mask
    ReDim m_hashTable(0 To INIT_TABLE_SIZE - 1) As MergedEntry
    ReDim m_overflowTable(0 To INIT_TABLE_SIZE - 1) As MergedEntry
    HASH_TABLE_MASK = INIT_TABLE_SIZE - 1
    
    '0 is used to denote "no children", so ensure the overflow index starts at 1
    m_idxOverflow = 1
    
End Sub

'Hash function.  Could be inlined for a nice perf boost, as relevant.  Note that this
' function ALWAYS returns an index into the hash table (i.e. it is pre-masked for you).
Private Function GetKeyHash(ByRef srcKey As String) As Long
    
    'Note that zero-length keys have undefined behavior, by design.  (PD does not use these.)
    ' You could handle these yourself, but you'd need to add an extra tracking member to the underlying
    ' table struct to distinguish between null and unused entries.
    
    'Hash the string using libdeflate, and use the length of the incoming string as the
    ' initial crc key.  (This slightly improves entropy for our purposes.)
    GetKeyHash = Plugin_libdeflate.GetCrc32(StrPtr(srcKey), LenB(srcKey), LenB(srcKey), False) And HASH_TABLE_MASK
    
End Function

'This function imposes a large performance penalty.  *Please* call it sparingly!
Private Sub IncreaseTableSize()
    
    'TODO: switch to *SafeArray swapping here, to improve performance
    
    'If we're here, it means we've run out of space in the overflow table.
    ' (In the current implementation, the hash and overflow tables are always identically sized.
    '  If the overflow table overflows, we double the size of *both* the hash table and the
    '  overflow table, then re-add all existing elements.)
    
    'Start by backing up the existing hash tables into temporary arrays
    Dim tmpHash() As MergedEntry, tmpOverflow() As MergedEntry
    ReDim tmpHash(0 To UBound(m_hashTable)) As MergedEntry
    ReDim tmpOverflow(0 To UBound(m_overflowTable)) As MergedEntry
    
    'Unfortunately, because we're using object types, we can't do a simple copy-memory because
    ' it will cause strings to deallocate.  Instead, copy the old-fashioned way (ugh).  Note that
    ' we *could* do StrPtr swaps here for better perf, but I am looking at using UTF-8 storage
    ' which would use 50% less memory... so until a firm decision is made, use this simple copy.
    Dim i As Long
    For i = 0 To UBound(tmpHash)
        tmpHash(i) = m_hashTable(i)
        tmpOverflow(i) = m_overflowTable(i)
    Next i
    
    'Calculate new table sizes, then increase the main hash and overflow tables to match
    Dim newTableSize As Long
    newTableSize = (HASH_TABLE_MASK + 1) * 2
    HASH_TABLE_MASK = newTableSize - 1
    
    ReDim m_hashTable(0 To newTableSize - 1) As MergedEntry
    ReDim m_overflowTable(0 To newTableSize - 1) As MergedEntry
    m_idxOverflow = 1
    
    'Re-add all items to the new, larger hash table
    For i = 0 To UBound(tmpHash)
        If (LenB(tmpHash(i).sKey) <> 0) Then Me.AddItem tmpHash(i).sKey, tmpHash(i).sItem
    Next i
    
    'By definition, the overflow table was full prior to this resize, so we don't need to check length
    ' on entries. (Note also that we start at position 1 because 0 is a reserved value indicating
    ' "no linked entry".)
    For i = 1 To UBound(tmpOverflow)
        Me.AddItem tmpOverflow(i).sKey, tmpOverflow(i).sItem
    Next i
    
End Sub

Private Sub Class_Initialize()
    Me.Reset
End Sub

'If you're curious about occupancy rates, this function will describe it to you
Friend Sub PrintClassDebugInfo()
    Dim numItemsInTable As Long
    numItemsInTable = Me.GetNumOfItems()
    PDDebug.LogAction "Unique entries in this cache: " & numItemsInTable
    PDDebug.LogAction " (" & Format$((numItemsInTable - (m_idxOverflow - 1)) / (UBound(m_hashTable) + 1), "0.0%") & " table occupancy, final table size was " & (UBound(m_hashTable) + 1) & "x2)"
End Sub
