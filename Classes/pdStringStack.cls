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
'Copyright 2014-2017 by Tanner Helland
'Created: 05/February/15
'Last updated: 30/August/15
'Last update: improve performance of stack resetting by only allocating new memory if absolutely necessary
'
'Per its name, this class provides a simple interface to a stack comprised of strings.  PD often has need to deal
' with large string collections (iterating folders, image metadata, etc), and rather than manually settings up
' collections for each instance, I've decided to simply use this small class.
'
'Note that it's not *technically* a stack, by design, as it's sometimes helpful to retrieve data from the middle
' of the stack (rather than enforcing a strict push/pop access system).  But I like the name "string stack" so I
' went with it. ;)
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private m_Strings() As String
Private m_NumOfStrings As Long
Private Const INIT_STACK_SIZE = 16

Private Declare Sub CopyMemory_Custom Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDst As Long, ByRef lpSrc As Long, ByVal byteLength As Long)
Private Declare Function GetMem4 Lib "msvbvm60" (ByVal pDWordSrc As Long, ByVal pDWordDst As Long) As Long
Private Declare Function PutMem4 Lib "msvbvm60" (ByVal ptrDst As Long, ByVal newValue As Long) As Long

'Because recursion is painfully slow in VB, our QuickSort implementation uses a stack instead.
Private Type QSStack
    sLB As Long
    sUB As Long
End Type

Private Const INIT_QUICKSORT_STACK_SIZE As Long = 256
Private m_qsStack() As QSStack
Private m_qsStackSize As Long, m_qsStackPtr As Long

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

'Trim the stack to its exact size.  IMPORTANT NOTE!  Don't do this any more than you have to, as it's not performance-friendly.
Friend Sub TrimStack()
    ReDim Preserve m_Strings(0 To m_NumOfStrings - 1) As String
End Sub

'Retrieve a string from the stack, with optional support for locale invariant conversions (when the caller expects the string
' to represent a number of some sort)
Friend Function GetString(ByVal strIndex As Long, Optional ByVal assumeLocaleInvariantNumber As Boolean = False) As String

    If (strIndex >= 0) And (strIndex < m_NumOfStrings) Then
        
        'Some callers may use this function to return a numeric value as a String, e.g. prior to creating a param string.
        ' They can use the assumeLocaleInvariantNumber parameter to notify us of this, and we will translate the key
        ' at this point to a safe, locale-invariant string representation.
        If assumeLocaleInvariantNumber Then
        
            'If the string representation of this key can be coerced into a numeric value, use a (rather ugly) series
            ' of transforms to ensure that the string representation of the number *never* varies by locale.  This is
            ' important as the original string may be locale-specific (especially if it originated from a text box),
            ' but we only want to use locale-invariant versions internally.
            Dim testString As String
            testString = m_Strings(strIndex)
            
            If IsNumberLocaleUnaware(testString) Then
                GetString = Trim$(Str(Val(testString)))
            Else
                GetString = testString
            End If
        
        Else
            GetString = m_Strings(strIndex)
        End If
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

'Locale-unaware check for strings that can successfully be converted to numbers.  Thank you to
' http://stackoverflow.com/questions/18368680/vb6-isnumeric-behaviour-in-windows-8-windows-2012
' for the code.  (Note that the original function listed there is buggy!  I had to add some
' fixes for exponent strings, which the original code did not handle correctly.)
Private Function IsNumberLocaleUnaware(ByRef Expression As String) As Boolean
    
    Dim Negative As Boolean, Number As Boolean, Period As Boolean, Positive As Boolean, Exponent As Boolean
    Dim x As Long
    For x = 1& To Len(Expression)
        Select Case Mid$(Expression, x, 1&)
        Case "0" To "9"
            Number = True
        Case "-"
            If Period Or Number Or Negative Or Positive Then Exit Function
            Negative = True
        Case "."
            If Period Or Exponent Then Exit Function
            Period = True
        Case "E", "e"
            If Not Number Then Exit Function
            If Exponent Then Exit Function
            Exponent = True
            Number = False
            Negative = False
            Period = False
        Case "+"
            If Not Exponent Then Exit Function
            If Number Or Negative Or Positive Then Exit Function
            Positive = True
        Case " ", vbTab, vbVerticalTab, vbCr, vbLf, vbFormFeed
            If Period Or Number Or Exponent Or Negative Then Exit Function
        Case Else
            Exit Function
        End Select
    Next x
        
    IsNumberLocaleUnaware = Number
    
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
                    If Len(m_Strings(i)) < Len(m_Strings(j)) Then SwapIndices i, j
                
                'An opposite check is used for descending order.
                Else
                    If Len(m_Strings(i)) > Len(m_Strings(j)) Then SwapIndices i, j
                End If
                
            Next j
        Next i
        
    End If

End Sub

'Helper for string sort functions.  To avoid new string allocations, this function simply swaps BSTR pointers.
Friend Sub SwapIndices(ByVal strIndex1 As Long, ByVal strIndex2 As Long)
    Dim tmpStrPtr As Long
    tmpStrPtr = StrPtr(m_Strings(strIndex1))
    GetMem4 VarPtr(m_Strings(strIndex2)), VarPtr(m_Strings(strIndex1))
    GetMem4 VarPtr(tmpStrPtr), VarPtr(m_Strings(strIndex2))
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
    QuickSortStringStack3Way
    
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
    'Debug.Print "String stack sorted in " & VBHacks.GetTimeDiffNowAsString(startTime)
    
End Sub

'Bentley-McIlroy 3-way partitioning QuickSort implementation developed with help from this PDF:
' http://www.cs.princeton.edu/~rs/talks/QuicksortIsOptimal.pdf
'
'Because recursion can quickly become slow in VB, we use an internal stack to track remaining range searches.
Private Sub QuickSortStringStack3Way()
    
    If (m_NumOfStrings > 1) Then
    
        'Prep our internal stack
        ReDim m_qsStack(0 To INIT_QUICKSORT_STACK_SIZE - 1) As QSStack
        m_qsStackPtr = 0
        m_qsStack(0).sLB = 0
        m_qsStack(0).sUB = m_NumOfStrings - 1
        
        QSInnerStack3Way
        
        'Free the stack before exiting
        Erase m_qsStack
        
    End If
    
End Sub

Private Sub QSInnerStack3Way()
    
    Dim lowVal As Long, highVal As Long
    Dim i As Long, j As Long, k As Long
    Dim p As Long, q As Long
    Dim v As String, vPtr As Long
    
    Do
        
        'Load the next set of boundaries, and reset all pivots
        lowVal = m_qsStack(m_qsStackPtr).sLB
        highVal = m_qsStack(m_qsStackPtr).sUB
        
        i = lowVal - 1
        j = highVal
        p = lowVal - 1
        q = highVal
        v = m_Strings(highVal)
        vPtr = StrPtr(v)
        
        'Start the "infinite" loop comprising the bulk of the sort
        Do
            
            'Search from the left
            Do
                i = i + 1
                If (i >= highVal) Then Exit Do
            Loop While Strings.StrCompSortPtr(StrPtr(m_Strings(i)), vPtr) < 0
            
            'Search from the right
            Do
                j = j - 1
                If (j <= lowVal) Then Exit Do
            Loop While Strings.StrCompSortPtr(vPtr, StrPtr(m_Strings(j))) < 0
            
            'If indices cross, this segment is fully sorted; exit!
            If (i >= j) Then Exit Do
            
            'If we're still here, these two indices are out-of-order.  Swap 'em.
            SwapIndices i, j
            
            'Handle equality cases specially.  (This improves performance on lists with many
            ' duplicate elements.)
            If (Strings.StrCompSortPtr(StrPtr(m_Strings(i)), vPtr) = 0) Then
                p = p + 1
                SwapIndices p, i
            End If
            
            If (Strings.StrCompSortPtr(vPtr, StrPtr(m_Strings(j))) = 0) Then
                q = q - 1
                SwapIndices j, q
            End If
            
        Loop
        
        'Swap all equal entries to the center of the partition, so we don't have to process them again
        SwapIndices i, highVal
        j = i - 1
        i = i + 1
        
        For k = lowVal To p - 1
            SwapIndices k, j
            j = j - 1
        Next k
        
        For k = highVal - 1 To q + 1 Step -1
            SwapIndices i, k
            i = i + 1
        Next k
        
        'Push (up to) two new entries on the processing stack
        If (m_qsStackPtr >= UBound(m_qsStack)) Then ReDim Preserve m_qsStack(0 To m_qsStackPtr * 2 - 1) As QSStack
        
        'We now have a trade-off to make:
        ' 1) We can perform two if/then checks to see if this segment needs to be subdivided further.
        '    If either of these tests succeeds, we can skip adding a new segment to the stack.
        ' 2) We can perform zero if/then checks and forcibly add new segments to the stack.
        '    Then, in the outer Do/Loop, we can perform an "If (lowVal < highVal)" check, and immediately
        '    decrement the stack if the test fails.
        '
        'There are trade-offs to both mechanisms, but given VB's array handling performance, I think
        ' it's preferable to use branches instead of risking extra array accesses and ReDims.
        If (j > lowVal) Then
            m_qsStack(m_qsStackPtr).sLB = lowVal
            m_qsStack(m_qsStackPtr).sUB = j
            m_qsStackPtr = m_qsStackPtr + 1
        End If
        If (highVal > i) Then
            m_qsStack(m_qsStackPtr).sLB = i
            m_qsStack(m_qsStackPtr).sUB = highVal
            m_qsStackPtr = m_qsStackPtr + 1
        End If
        
        'Decrement the stack pointer
        m_qsStackPtr = m_qsStackPtr - 1
        
    Loop While (m_qsStackPtr >= 0)
    
End Sub

'Clone another string stack
Friend Sub CloneStack(ByRef stackToClone As pdStringStack)
    
    'Initialize this stack to the size of the target
    Me.ResetStack stackToClone.GetNumOfStrings
    
    'Copy all strings
    Dim i As Long
    For i = 0 To stackToClone.GetNumOfStrings - 1
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

    If m_NumOfStrings > 0 Then
        
        'The first entry in the serialized string is always the string count
        Dim finalString As String
        finalString = CStr(m_NumOfStrings)
        
        Dim tstPipeString As String: tstPipeString = "|"
        Dim rplPipeString As String: rplPipeString = "&pipe;"
        
        Dim i As Long
        For i = 0 To m_NumOfStrings - 1
            If (InStr(1, m_Strings(i), tstPipeString, vbBinaryCompare) <> 0) Then
                finalString = finalString & tstPipeString & Replace(m_Strings(i), tstPipeString, rplPipeString, , , vbBinaryCompare)
            Else
                finalString = finalString & tstPipeString & m_Strings(i)
            End If
        Next i
        
    End If
    
End Function

Friend Function RecreateStackFromSerializedString(ByRef srcString As String) As Boolean
    
    On Error GoTo RecreateFailure
    
    Me.ResetStack
    
    Dim tstPipeString As String: tstPipeString = "|"
    Dim rplPipeString As String: rplPipeString = "&pipe;"
    
    If InStr(1, srcString, tstPipeString, vbBinaryCompare) <> 0 Then
    
        Dim stringArray() As String
        stringArray = Split(srcString, tstPipeString, , vbBinaryCompare)
        
        m_NumOfStrings = CLng(stringArray(LBound(stringArray)))
        
        Dim i As Long
        For i = 0 To m_NumOfStrings - 1
            
            If (InStr(1, stringArray(i + 1), rplPipeString, vbBinaryCompare) <> 0) Then
                Me.AddString Replace(stringArray(i + 1), rplPipeString, tstPipeString, , , vbBinaryCompare)
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
Friend Sub DEBUG_dumpResultsToImmediateWindow()
    
    If (m_NumOfStrings > 0) Then
        Dim i As Long
        For i = 0 To m_NumOfStrings - 1
            Debug.Print i & ": -- " & m_Strings(i) & " -- "
        Next i
    Else
        Debug.Print " -- String stack is empty -- "
    End If
End Sub
