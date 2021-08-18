VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdStringStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon "String Stack" Class
'Copyright 2014-2021 by Tanner Helland
'Created: 05/February/15
'Last updated: 20/July/21
'Last update: add a fast "RemoveString" function; this isn't actually needed in PD, but in the separate
'             update-file-generator (which needs to manually flag some files as unnecessary)
'
'Per its name, this class provides a simple interface to a stack comprised of strings.  PD often has need to deal
' with large string collections (iterating folders, image metadata, etc), and rather than manually settings up
' collections for each instance, I've decided to simply use this small class.
'
'Note that it's not *technically* a stack, by design, as it's sometimes helpful to retrieve data from the middle
' of the stack (rather than enforcing a strict push/pop access system).  But I like the name "string stack" so I
' went with it. ;)
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private m_Strings() As String
Private m_NumOfStrings As Long
Private Const INIT_STACK_SIZE = 16

'Because recursion is painfully slow in VB, our QuickSort implementation uses a stack instead.
Private Type QSStack
    sLB As Long
    sUB As Long
End Type

Private Const INIT_QUICKSORT_STACK_SIZE As Long = 256
Private m_qsStack() As QSStack
Private m_qsStackPtr As Long

'Add a string to the stack.  Return value is the index of the added location (which can be used to infer the number of strings
' in the stack, obviously).
Friend Function AddString(ByRef srcString As String) As Long

    'Resize the stack as necessary
    If (m_NumOfStrings > UBound(m_Strings)) Then ReDim Preserve m_Strings(0 To m_NumOfStrings * 2 - 1) As String
    
    'Add the string
    m_Strings(m_NumOfStrings) = srcString
    
    AddString = m_NumOfStrings
    m_NumOfStrings = m_NumOfStrings + 1
        
End Function

'Append a full stack to this one
Friend Sub AppendStack(ByRef srcStack As pdStringStack)
    If (Not srcStack Is Nothing) Then
        If (srcStack.GetNumOfStrings > 0) Then
            Dim i As Long
            For i = 0 To srcStack.GetNumOfStrings() - 1
                Me.AddString srcStack.GetString(i)
            Next i
        End If
    End If
End Sub

'Check the stack for an existing string.  Returns an index >= 0 if the string exists in the stack;
' returns a negative number if the string does *not* occur.
Friend Function ContainsString(ByRef srcString As String, Optional ByVal ignoreCase As Boolean = False) As Long

    ContainsString = -1
    If (m_NumOfStrings > 0) Then
    
        Dim i As Long
        For i = 0 To m_NumOfStrings - 1
            If Strings.StringsEqual(srcString, m_Strings(i), ignoreCase) Then
                ContainsString = i
                Exit For
            End If
        Next i
        
    End If

End Function

'Pop the top string off the stack.  Returns TRUE if pop is successful, FALSE if stack is empty.  Caller is responsible for
' allocating their own destination string, which this function simply fills.
'
'The function was designed to make popping the entire stack convenient (e.g. Do While strStack.PopString(tmpString)...)
'
'Note that this function DOES NOT shrink the string array to match.  This is by design.  If you want to resize the string array
' after a pop, manually call TrimStack().  (But seriously - don't do this unless you really need to, as the performance
' implications are severe.)
Friend Function PopString(ByRef dstString As String) As Boolean
    
    If (m_NumOfStrings > 0) Then
        m_NumOfStrings = m_NumOfStrings - 1
        dstString = m_Strings(m_NumOfStrings)
        PopString = True
    Else
        PopString = False
    End If
    
End Function

'Return the size of the stack
Friend Function GetNumOfStrings() As Long
    GetNumOfStrings = m_NumOfStrings
End Function

'Set a new stack size.  By design, this only works if the new size is <= the current stack size
' (otherwise we'd need to allocate new memory and what would we even fill it with - null strings?).
Friend Sub SetNumOfStrings(ByVal newNumStrings As Long)
    If (newNumStrings <= m_NumOfStrings) Then m_NumOfStrings = newNumStrings
End Sub

'Trim the stack to its exact size.
' IMPORTANT NOTE!  Don't do this any more than you have to, as it's not performance-friendly.
Friend Sub TrimStack()
    ReDim Preserve m_Strings(0 To m_NumOfStrings - 1) As String
End Sub

Friend Sub FreeString(ByVal strIndex As Long)
    If (strIndex >= 0) And (strIndex < m_NumOfStrings) Then m_Strings(strIndex) = vbNullString
End Sub

'Retrieve a string from the stack
Friend Function GetString(ByVal strIndex As Long) As String

    If (strIndex >= 0) And (strIndex < m_NumOfStrings) Then
        GetString = m_Strings(strIndex)
    Else
        Debug.Print "WARNING!  Someone asked pdStringStack for a string outside stack bounds.  Fix this!"
    End If

End Function

'Retrieve a string pointer from the stack; helpful for API interactions
Friend Function GetStringPointer(ByVal strIndex As Long) As Long

    If (strIndex >= 0) And (strIndex < m_NumOfStrings) Then
        GetStringPointer = StrPtr(m_Strings(strIndex))
    Else
        Debug.Print "WARNING!  Someone asked pdStringStack for a string outside stack bounds.  Fix this!"
    End If

End Function

'This function may seem like a ridiculous addition, but it's actually very helpful in PD.  pdStringStack is used by pdFSO
' when retrieving all subfolders inside some base folder.  When performing something like a tree copy, I like to pre-sort
' the subfolder list by length.  This greatly simplifies the code required to create the new folder tree prior to performing
' the copy; creating the folders in advance greatly accelerates the copy operation, as we don't have to perform "do my
' folders exist?" checks on every damn file.
Friend Sub SortStackByLength(Optional ByVal sortAscending As Boolean = True)

    If (m_NumOfStrings > 1) Then
    
        'Given PD's standard use-case (subfolder trees, as mentioned above), the existing stack order is typically pretty close
        ' to sorted.  This saves us from needing an elaborate search algorithm; instead, a simple in-place bubble sort
        ' is predictable and good enough.
        Dim i As Long, j As Long, loopBound As Long
        loopBound = m_NumOfStrings - 1
        
        'Loop through all entries in the stack, sorting them as we go
        For i = 0 To loopBound
            For j = 0 To loopBound
                
                'Compare two entries, and if the longer one precedes the shorter one, swap them
                If sortAscending Then
                    If LenB(m_Strings(i)) < LenB(m_Strings(j)) Then SwapIndices i, j
                
                'An opposite check is used for descending order.
                Else
                    If LenB(m_Strings(i)) > LenB(m_Strings(j)) Then SwapIndices i, j
                End If
                
            Next j
        Next i
        
    End If

End Sub

'Helper for string sort functions.  To avoid new string allocations, this function simply swaps BSTR pointers.
Friend Sub SwapIndices(ByVal strIndex1 As Long, ByVal strIndex2 As Long)
    Dim tmpStrPtr As Long
    tmpStrPtr = StrPtr(m_Strings(strIndex1))
    GetMem4 VarPtr(m_Strings(strIndex2)), ByVal VarPtr(m_Strings(strIndex1))
    GetMem4 VarPtr(tmpStrPtr), ByVal VarPtr(m_Strings(strIndex2))
End Sub

'Instead of swapping two entries, this function will move a string to a new index, then shift all remaining strings
' in the list to match.  Obviously there are performance issues with this, so use it sparingly!
Friend Sub MoveStringToNewPosition(ByVal srcStringIndex As Long, ByVal dstStringIndex As Long)
    
    If (srcStringIndex >= 0) And (srcStringIndex < m_NumOfStrings) And (dstStringIndex >= 0) And (dstStringIndex < m_NumOfStrings) And (srcStringIndex <> dstStringIndex) Then
    
        Dim i As Long, tmpString As String
        tmpString = m_Strings(srcStringIndex)
        
        If (srcStringIndex < dstStringIndex) Then
            For i = srcStringIndex To dstStringIndex - 1
                m_Strings(i) = m_Strings(i + 1)
            Next i
        Else
            For i = srcStringIndex To dstStringIndex + 1 Step -1
                m_Strings(i) = m_Strings(i - 1)
            Next i
        End If
        
        m_Strings(dstStringIndex) = tmpString
        
    End If

End Sub

'Per the name, sort the current stack alphabetically.  In PD, a sorted stack typically needs to have any
' duplicate entries removed (like e.g. font lists, which Windows may report in strange ways), so this
' function also handles that duty, as requested.
Friend Sub SortAlphabetically(Optional ByVal removeDuplicates As Boolean = False)
    
    If (m_NumOfStrings <= 1) Then Exit Sub
    
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    'Sort the stack in-place
    QuickSortStringStack
    
    'If the user wants duplicates removed, do so now.  (This check is fast and easy, because duplicates are
    ' sitting next to each other in the sorted list.)
    If removeDuplicates Then
        
        Dim itemsRemoved As Long
        itemsRemoved = 0
        
        Dim itemCount As Long
        itemCount = m_NumOfStrings - 1
        
        Dim i As Long, j As Long
        i = 0
        
        Do While (i < (itemCount - itemsRemoved))
            
            'If this string and the string above it match, shift everything above it downward
            If Strings.StringsEqual(m_Strings(i), m_Strings(i + 1), True) Then
                
                If (itemCount - itemsRemoved) < UBound(m_Strings) Then
                    For j = i To itemCount - itemsRemoved
                        m_Strings(j) = m_Strings(j + 1)
                    Next j
                End If
                
                itemsRemoved = itemsRemoved + 1
                
            Else
                i = i + 1
            End If
            
        Loop
        
        'If one or more items were removed, mark the new array size accordingly
        If (itemsRemoved > 0) Then m_NumOfStrings = m_NumOfStrings - itemsRemoved
        
    End If
    
    'Want timing reports?  Here you go:
    'pdDebug.LogAction "String collection sorted in " & VBHacks.GetTimeDiffNowAsString(startTime)
    
End Sub

Private Sub QuickSortStringStack(Optional ByVal useLogicalSortOrder As Boolean = False)
    
    If (m_NumOfStrings > 1) Then
    
        'Prep our internal stack
        ReDim m_qsStack(0 To INIT_QUICKSORT_STACK_SIZE - 1) As QSStack
        m_qsStackPtr = 0
        m_qsStack(0).sLB = 0
        m_qsStack(0).sUB = m_NumOfStrings - 1
        
        If useLogicalSortOrder Then
            NaiveQuickSortExtended_Logical
        Else
            NaiveQuickSortExtended
        End If
        
        'Free the stack before exiting
        Erase m_qsStack
        
    End If
    
End Sub

'Semi-standard QuickSort implementation, with VB-specific enhancements provided by georgekar, and further
' enhancements by myself to further improve performance.  VB6 has any number of quirks that require special
' workarounds and optimizations, and because this function is primarily used for sorting API returns
' (like font names), its use-cases can vary wildly.  This modified implementation should will perform well
' under a variety of circumstances, including already-sorted, reverse-sorted, even-odd, and other typically
' rare scenarios, which is good as Windows rarely makes guarantees with sort order when returning strings
' (again, like installed font names).
'
'georgekar's original, unmodified implementation can be found here:
' http://www.vbforums.com/showthread.php?781043-VB6-Dual-Pivot-QuickSort
Private Sub NaiveQuickSortExtended()
    
    Dim lowVal As Long, highVal As Long
    Dim i As Long, j As Long
    Dim v As String, vPtr As Long
    
    Do
        
        'Load the next set of boundaries, and reset all pivots
        lowVal = m_qsStack(m_qsStackPtr).sLB
        highVal = m_qsStack(m_qsStackPtr).sUB
        
        'Check for single-entry ranges
        If (highVal - lowVal = 1) Then
            i = lowVal
            If (Strings.StrCompSortPtr(StrPtr(m_Strings(i)), StrPtr(m_Strings(highVal))) > 0) Then SwapIndices i, highVal 'Tmp = m_Strings(i): m_Strings(i) = m_Strings(highVal): m_Strings(highVal) = Tmp
            GoTo NextSortItem
        Else
            
            'Bisect this range
            i = (lowVal + highVal) \ 2
            
            'Migrate all equal entries into place
            If (Strings.StrCompSortPtr(StrPtr(m_Strings(i)), StrPtr(m_Strings(lowVal))) = 0) Then
                
                j = highVal - 1
                i = lowVal
                
                Do
                    i = i + 1
                    If (i > j) Then
                        If (Strings.StrCompSortPtr(StrPtr(m_Strings(highVal)), StrPtr(m_Strings(lowVal))) < 0) Then SwapIndices lowVal, highVal
                        GoTo NextSortItem
                    End If
                Loop Until (Strings.StrCompSortPtr(StrPtr(m_Strings(i)), StrPtr(m_Strings(lowVal))) <> 0)
                
                v = m_Strings(i)
                If (i > lowVal) Then If (Strings.StrCompSortPtr(StrPtr(m_Strings(lowVal)), StrPtr(m_Strings(i))) > 0) Then SwapIndices lowVal, i
            
            'Move the pointer until we arrive at an unsorted pivot
            Else
                v = m_Strings(i)
                vPtr = StrPtr(m_Strings(i))
                i = lowVal
                Do While (Strings.StrCompSortPtr(StrPtr(m_Strings(i)), vPtr) < 0): i = i + 1: Loop
            End If
        
        'End special case handling
        End If
        
        'Resume standard QuickSort behavior
        j = highVal
        
        Do
            'Advance from the right
            Do While (Strings.StrCompSortPtr(StrPtr(m_Strings(j)), StrPtr(v)) > 0): j = j - 1: Loop
            
            'Swap as necessary
            If (i <= j) Then
                SwapIndices i, j
                i = i + 1
                j = j - 1
            End If
            
            If (i > j) Then Exit Do
            
            'Advance from the left
            Do While (Strings.StrCompSortPtr(StrPtr(m_Strings(i)), StrPtr(v)) < 0): i = i + 1: Loop
            
        Loop
        
        'Conditionally add new entries to the processing stack
        If (lowVal < j) Then
            m_qsStack(m_qsStackPtr).sLB = lowVal
            m_qsStack(m_qsStackPtr).sUB = j
            m_qsStackPtr = m_qsStackPtr + 1
        End If
        
        If (i < highVal) Then
            m_qsStack(m_qsStackPtr).sLB = i
            m_qsStack(m_qsStackPtr).sUB = highVal
            m_qsStackPtr = m_qsStackPtr + 1
        End If
        
'Yep, VB6 requires us to use GOTO and line labels.  There is no "Continue For" equivalent.
NextSortItem:
        
        'Decrement the stack pointer
        m_qsStackPtr = m_qsStackPtr - 1
        
    Loop While (m_qsStackPtr >= 0)
    
End Sub

'Per the name, sort the current stack the same way Windows Explorer does.  As users can create
' or import their own lists of filenames (e.g. in the Batch Process window), we also perform
' a failsafe check for duplicate entries.
Friend Sub SortLogically(Optional ByVal removeDuplicates As Boolean = True)
    
    If (m_NumOfStrings <= 1) Then Exit Sub
    
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    'Sort the stack in-place
    QuickSortStringStack True
    
    'If the user wants duplicates removed, do so now.  (This check is fast and easy, because duplicates are
    ' sitting next to each other in the sorted list.)
    If removeDuplicates Then
        
        Dim itemsRemoved As Long
        itemsRemoved = 0
        
        Dim itemCount As Long
        itemCount = m_NumOfStrings - 1
        
        Dim i As Long, j As Long
        i = 0
        
        Do While (i < (itemCount - itemsRemoved))
            
            'If this string and the string above it match, shift everything above it downward
            If (Strings.StrCompSortPtr_Filenames(StrPtr(m_Strings(i)), StrPtr(m_Strings(i + 1))) = 0) Then
                
                If (itemCount - itemsRemoved) < UBound(m_Strings) Then
                    For j = i To itemCount - itemsRemoved
                        m_Strings(j) = m_Strings(j + 1)
                    Next j
                End If
                
                itemsRemoved = itemsRemoved + 1
                
            Else
                i = i + 1
            End If
            
        Loop
        
        'If one or more items were removed, mark the new array size accordingly
        If (itemsRemoved > 0) Then m_NumOfStrings = m_NumOfStrings - itemsRemoved
        
    End If
    
    'Want timing reports?  Here you go:
    'pdDebug.LogAction "String collection sorted in " & VBHacks.GetTimeDiffNowAsString(startTime)
    
End Sub

'Semi-standard QuickSort implementation, using StrCmpLogicalW (the sort function used by
' Windows Explorer), with VB-specific enhancements provided by georgekar, and further
' enhancements by myself to further improve performance.
'
'georgekar's original, unmodified implementation can be found here:
' http://www.vbforums.com/showthread.php?781043-VB6-Dual-Pivot-QuickSort
Private Sub NaiveQuickSortExtended_Logical()
    
    Dim lowVal As Long, highVal As Long
    Dim i As Long, j As Long
    Dim v As String, vPtr As Long
    
    Do
        
        'Load the next set of boundaries, and reset all pivots
        lowVal = m_qsStack(m_qsStackPtr).sLB
        highVal = m_qsStack(m_qsStackPtr).sUB
        
        'Check for single-entry ranges
        If (highVal - lowVal = 1) Then
            i = lowVal
            If (Strings.StrCompSortPtr_Filenames(StrPtr(m_Strings(i)), StrPtr(m_Strings(highVal))) > 0) Then SwapIndices i, highVal  'Tmp = m_Strings(i): m_Strings(i) = m_Strings(highVal): m_Strings(highVal) = Tmp
            GoTo NextSortItem
        Else
            
            'Bisect this range
            i = (lowVal + highVal) \ 2
            
            'Migrate all equal entries into place
            If (Strings.StrCompSortPtr_Filenames(StrPtr(m_Strings(i)), StrPtr(m_Strings(lowVal))) = 0) Then
                
                j = highVal - 1
                i = lowVal
                
                Do
                    i = i + 1
                    If (i > j) Then
                        If (Strings.StrCompSortPtr_Filenames(StrPtr(m_Strings(highVal)), StrPtr(m_Strings(lowVal))) < 0) Then SwapIndices lowVal, highVal
                        GoTo NextSortItem
                    End If
                Loop Until (Strings.StrCompSortPtr_Filenames(StrPtr(m_Strings(i)), StrPtr(m_Strings(lowVal))) <> 0)
                
                v = m_Strings(i)
                If (i > lowVal) Then If (Strings.StrCompSortPtr_Filenames(StrPtr(m_Strings(lowVal)), StrPtr(m_Strings(i))) > 0) Then SwapIndices lowVal, i
            
            'Move the pointer until we arrive at an unsorted pivot
            Else
                v = m_Strings(i)
                vPtr = StrPtr(m_Strings(i))
                i = lowVal
                Do While (Strings.StrCompSortPtr_Filenames(StrPtr(m_Strings(i)), vPtr) < 0): i = i + 1: Loop
            End If
        
        'End special case handling
        End If
        
        'Resume standard QuickSort behavior
        j = highVal
        
        Do
            'Advance from the right
            Do While (Strings.StrCompSortPtr_Filenames(StrPtr(m_Strings(j)), StrPtr(v)) > 0): j = j - 1: Loop
            
            'Swap as necessary
            If (i <= j) Then
                SwapIndices i, j
                i = i + 1
                j = j - 1
            End If
            
            If (i > j) Then Exit Do
            
            'Advance from the left
            Do While (Strings.StrCompSortPtr_Filenames(StrPtr(m_Strings(i)), StrPtr(v)) < 0): i = i + 1: Loop
            
        Loop
        
        'Conditionally add new entries to the processing stack
        If (lowVal < j) Then
            m_qsStack(m_qsStackPtr).sLB = lowVal
            m_qsStack(m_qsStackPtr).sUB = j
            m_qsStackPtr = m_qsStackPtr + 1
        End If
        
        If (i < highVal) Then
            m_qsStack(m_qsStackPtr).sLB = i
            m_qsStack(m_qsStackPtr).sUB = highVal
            m_qsStackPtr = m_qsStackPtr + 1
        End If
        
'Yep, VB6 requires us to use GOTO and line labels.  There is no "Continue For" equivalent.
NextSortItem:
        
        'Decrement the stack pointer
        m_qsStackPtr = m_qsStackPtr - 1
        
    Loop While (m_qsStackPtr >= 0)
    
End Sub

'Clone another string stack
Friend Sub CloneStack(ByRef stackToClone As pdStringStack)
    
    'Initialize this stack to the size of the target
    Me.ResetStack stackToClone.GetNumOfStrings()
    
    'Copy all strings
    Dim i As Long
    For i = 0 To stackToClone.GetNumOfStrings() - 1
        Me.AddString stackToClone.GetString(i)
    Next i
    
End Sub

'Return our list of strings as a bare string array
Friend Sub GetCopyOfStringArray(ByRef dstStringArray() As String)
    ReDim dstStringArray(0 To m_NumOfStrings - 1) As String
    Dim i As Long
    For i = 0 To m_NumOfStrings - 1
        dstStringArray(i) = m_Strings(i)
    Next i
End Sub

'Fill this stack with the contents of a bare string array.  Do not pass an uninitialized array.
Friend Sub CreateFromStringArray(ByRef srcStringArray() As String)
    Dim i As Long
    For i = LBound(srcStringArray) To UBound(srcStringArray)
        Me.AddString srcStringArray(i)
    Next i
End Sub

'Fill this stack by parsing a string.  VBA.Split() is currently used; obviously a faster solution could be implemented.
Friend Sub CreateFromMultilineString(ByRef srcString As String, Optional ByVal lineDelimiter As String = vbCrLf)
    Dim tmpArray() As String
    tmpArray = Split(srcString, lineDelimiter, , vbBinaryCompare)
    Me.CreateFromStringArray tmpArray
End Sub

'Remove a string from the current stack.
' IMPORTANT NOTE: use this function sparingly, because even though it's optimized using
' "pointer magic" (or the closest thing to it VB has) to migrate strings in the stack
' into position over the removed one, it still has to iterate all strings past the target one
Friend Sub RemoveStringByIndex(ByVal stringIndex As Long)
    
    If (stringIndex >= 0) And (stringIndex < m_NumOfStrings) Then
    
        'Free the target string
        m_Strings(stringIndex) = vbNullString
        
        'Use pointer tricks to shift all neighboring strings "down" in the stack
        Dim idxCurString As Long
        idxCurString = stringIndex + 1
        
        Do While (idxCurString < m_NumOfStrings)
            GetMem4_Ptr VarPtr(m_Strings(idxCurString)), VarPtr(m_Strings(idxCurString - 1))
            idxCurString = idxCurString + 1
        Loop
        
        'Nullify the final string pointer in the table to prevent a leak (or double-free)
        If (stringIndex < m_NumOfStrings - 1) Then PutMem4 VarPtr(m_Strings(m_NumOfStrings - 1)), 0&
        
        'Decrement string counter
        m_NumOfStrings = m_NumOfStrings - 1
        
    End If
    
End Sub

'Convenience wrapper to RemoveStringByIndex(), above.
' Returns TRUE if string is found and removed successfully; FALSE means the
' string doesn't exist in the current stack.
Friend Function RemoveStringByText(ByRef srcString As String, Optional ByVal ignoreCase As Boolean = False) As Boolean
    Dim idxTarget As Long
    idxTarget = Me.ContainsString(srcString, ignoreCase)
    RemoveStringByText = (idxTarget >= 0)
    If RemoveStringByText Then Me.RemoveStringByIndex idxTarget
End Function

'Clear the current stack.  An optional stack size can be passed; if it is not passed, it will default to INIT_STACK_SIZE
Friend Sub ResetStack(Optional ByVal newStackSize As Long = INIT_STACK_SIZE)
    
    On Error GoTo FailsafeReset
    
    'Failsafe bounds check
    If (newStackSize <= 0) Then newStackSize = INIT_STACK_SIZE
    
    'Reset the array (but only if necessary!)
    If (m_NumOfStrings = 0) Then
        ReDim m_Strings(0 To newStackSize - 1) As String
    Else
        If (UBound(m_Strings) <> newStackSize - 1) Then ReDim m_Strings(0 To newStackSize - 1) As String
    End If
    
    m_NumOfStrings = 0
    
    Exit Sub
    
FailsafeReset:
    If (newStackSize <= 0) Then newStackSize = INIT_STACK_SIZE
    ReDim m_Strings(0 To newStackSize - 1) As String
    
End Sub

Friend Function SerializeStackToSingleString() As String

    If (m_NumOfStrings > 0) Then
        
        'The first entry in the serialized string is always the string count
        Dim finalString As pdString
        Set finalString = New pdString
        finalString.Append CStr(m_NumOfStrings)
        
        Dim tstPipeString As String: tstPipeString = "|"
        Dim rplPipeString As String: rplPipeString = "&pipe;"
        
        Dim i As Long
        For i = 0 To m_NumOfStrings - 1
            finalString.Append tstPipeString
            If (InStr(1, m_Strings(i), tstPipeString, vbBinaryCompare) <> 0) Then
                finalString.Append Replace$(m_Strings(i), tstPipeString, rplPipeString, , , vbBinaryCompare)
            Else
                finalString.Append m_Strings(i)
            End If
        Next i
        
        SerializeStackToSingleString = finalString.ToString()
        
    End If
    
End Function

Friend Function RecreateStackFromSerializedString(ByRef srcString As String) As Boolean
    
    On Error GoTo RecreateFailure
    
    Me.ResetStack
    
    Dim tstPipeString As String: tstPipeString = "|"
    Dim rplPipeString As String: rplPipeString = "&pipe;"
    
    If (InStr(1, srcString, tstPipeString, vbBinaryCompare) <> 0) Then
    
        Dim stringArray() As String
        stringArray = Split(srcString, tstPipeString, , vbBinaryCompare)
        
        m_NumOfStrings = CLng(stringArray(LBound(stringArray)))
        
        Dim i As Long
        For i = 0 To m_NumOfStrings - 1
            
            If (InStr(1, stringArray(i + 1), rplPipeString, vbBinaryCompare) <> 0) Then
                Me.AddString Replace$(stringArray(i + 1), rplPipeString, tstPipeString, , , vbBinaryCompare)
            Else
                Me.AddString stringArray(i + 1)
            End If
        Next i
        
    End If
    
    RecreateStackFromSerializedString = True
    Exit Function
    
RecreateFailure:
    
    Debug.Print "WARNING!  Failed to un-serialize string: " & srcString
    RecreateStackFromSerializedString = False
    Exit Function

End Function

Private Sub Class_Initialize()
    
    'Always start with an initialized array
    Me.ResetStack
        
End Sub

Private Sub Class_Terminate()
    Me.ResetStack
End Sub

'DEBUG ONLY!  I sometimes find it helpful to investigate the contents of the stack.  This function makes it trivial to do so.
' I also append "--" to the start and end of the string, to help me see if extra whitespace chars are present.
Friend Sub DEBUG_DumpResultsToImmediateWindow()
    
    If (m_NumOfStrings > 0) Then
        Dim i As Long
        For i = 0 To m_NumOfStrings - 1
            Debug.Print i & ": -- " & m_Strings(i) & " -- "
        Next i
    Else
        Debug.Print " -- String stack is empty -- "
    End If
    
End Sub

'DEBUG ONLY!  I sometimes find it helpful to investigate the contents of the stack.
' This function makes it trivial to dump the stack to a standalone file.
Friend Function DEBUG_DumpResultsToFile(ByRef dstFilename As String) As Boolean
    
    Dim finalString As pdString
    Set finalString = New pdString
    
    If (m_NumOfStrings > 0) Then
        Dim i As Long
        For i = 0 To m_NumOfStrings - 1
            finalString.AppendLine m_Strings(i)
        Next i
    Else
        finalString.Append " -- String stack is empty -- "
    End If
    
    DEBUG_DumpResultsToFile = Files.FileSaveAsText(finalString.ToString(), dstFilename, True, True)
    
End Function

