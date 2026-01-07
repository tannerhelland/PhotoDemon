VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Generic (int) Stack
'Copyright 2015-2026 by Tanner Helland
'Created: 05/February/15
'Last updated: 08/March/24
'Last update: various helper functions to help with page selection during PDF import
'
'Per its name, this class provides a simple interface to a stack comprised of ints.
'
'Note that it's not *technically* a stack, by design, as it provides helper functions for retrieving data
' from the middle of the stack (rather than enforcing a strict push/pop access system).
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private m_Ints() As Long
Private m_NumOfInts As Long
Private Const INIT_STACK_SIZE As Long = 16

'Add a value to the stack.  Return value is the index of the added location (which can be used to infer the
' number of ints in the stack, obviously).
Friend Function AddInt(ByVal srcInt As Long) As Long

    'Resize the stack as necessary
    If (m_NumOfInts > UBound(m_Ints)) Then ReDim Preserve m_Ints(0 To m_NumOfInts * 2 - 1) As Long
    
    'Add the string
    m_Ints(m_NumOfInts) = srcInt
    
    AddInt = m_NumOfInts
    m_NumOfInts = m_NumOfInts + 1
        
End Function

'Does a given integer appear in the stack?  Uses a naive loop for detection, so consider perf as necessary.
' RETURNS: index >= 0 if int exists, -1 if int does *not* exist.  Returns first instance if multiples exist.
Friend Function DoesIntExist(ByVal srcInt As Long) As Long
    
    DoesIntExist = -1
    If (m_NumOfInts <= 0) Then Exit Function
    
    Dim i As Long
    For i = 0 To m_NumOfInts - 1
        If (m_Ints(i) = srcInt) Then
            DoesIntExist = i
            Exit Function
        End If
    Next i
    
End Function

'Pop the top int off the stack.  Returns TRUE if pop is successful, FALSE if stack is empty.
'
'The function was designed to make popping the entire stack convenient (e.g. Do While strStack.PopString(tmpString)...)
'
'Note that this function DOES NOT shrink the storage array to match.  This is by design.  If you want to resize the
' storage array after a pop, manually call TrimStack().  (But seriously - don't do this unless you really need to,
' as the performance implications are severe.)
Friend Function PopInt(ByRef dstInt As Long) As Boolean
    
    If (m_NumOfInts > 0) Then
        m_NumOfInts = m_NumOfInts - 1
        dstInt = m_Ints(m_NumOfInts)
        PopInt = True
    Else
        PopInt = False
    End If
    
End Function

'Return the size of the stack
Friend Function GetNumOfInts() As Long
    GetNumOfInts = m_NumOfInts
End Function

'Trim the stack to its exact size.  IMPORTANT NOTE!  Don't do this any more than you have to, as it's not performance-friendly.
Friend Sub TrimStack()
    If (m_NumOfInts > 0) Then ReDim Preserve m_Ints(0 To m_NumOfInts - 1) As Long
End Sub

'Retrieve a value from the stack at any arbitrary position
Friend Function GetInt(ByVal intIndex As Long) As Long
    If (intIndex >= 0) And (intIndex < m_NumOfInts) Then
        GetInt = m_Ints(intIndex)
    Else
        Debug.Print "WARNING!  Someone asked pdStack for a value outside stack bounds.  Fix this!"
    End If
End Function

'Clone another stack.  This could be optimized via memcpy, but it's only used sparingly at present,
' so I've gone with a "safer" implementation.
Friend Sub CloneStack(ByRef stackToClone As pdStack)
    
    'Initialize this stack to the size of the target
    Me.ResetStack stackToClone.GetNumOfInts
    
    'Copy all values
    Dim i As Long
    For i = 0 To stackToClone.GetNumOfInts - 1
        Me.AddInt stackToClone.GetInt(i)
    Next i
    
End Sub

'Return our list of strings as a bare int array
Friend Sub GetCopyOfIntArray(ByRef dstArray() As Long)
    ReDim dstArray(0 To m_NumOfInts - 1) As Long
    Dim i As Long
    For i = 0 To m_NumOfInts - 1
        dstArray(i) = m_Ints(i)
    Next i
End Sub

'Fill this stack with the contents of a bare int array.  Do not pass an uninitialized array.
Friend Sub CreateFromArray(ByRef srcArray() As Long)
    Dim i As Long
    For i = LBound(srcArray) To UBound(srcArray)
        Me.AddInt srcArray(i)
    Next i
End Sub

'Clear the current stack.  An optional stack size can be passed; if it is not passed, it will default to INIT_STACK_SIZE
Friend Sub ResetStack(Optional ByVal newStackSize As Long = INIT_STACK_SIZE)
    
    On Error GoTo FailsafeReset
    
    'Failsafe bounds check
    If (newStackSize <= 0) Then newStackSize = INIT_STACK_SIZE
    
    'Reset the array (but only if necessary!)
    If (m_NumOfInts = 0) Then
        ReDim m_Ints(0 To newStackSize - 1) As Long
    Else
        If (UBound(m_Ints) <> newStackSize - 1) Then ReDim m_Ints(0 To newStackSize - 1) As Long
    End If
    
    m_NumOfInts = 0
    
    Exit Sub
    
FailsafeReset:
    If (newStackSize <= 0) Then newStackSize = INIT_STACK_SIZE
    ReDim m_Ints(0 To newStackSize - 1) As Long
    
End Sub

'Reverse the order of items in the stack.  PD uses this when importing PDFs as the user can choose to
' swap page order (putting the first page at the *top* of the stack instead of the traditional bottom).
Friend Sub ReverseStack()
    
    If (m_NumOfInts > 0) Then
        
        Dim ubStack As Long
        ubStack = m_NumOfInts - 1
        
        Dim i As Long
        For i = 0 To (m_NumOfInts \ 2) - 1
            SwapIndices i, ubStack - i
        Next i
        
    End If
    
End Sub

'Sort stack entries from least to most (if sortAscending is TRUE)
Friend Sub SortStackByValue(Optional ByVal sortAscending As Boolean = True)

    If (m_NumOfInts > 1) Then
    
        'Given PD's standard use-case for this function (PDF import, at present), the existing stack order
        ' is typically already sorted.  This saves us from needing an elaborate search algorithm;
        ' instead, a simple in-place bubble sort is predictable and good enough.
        Dim i As Long, j As Long, loopBound As Long
        loopBound = m_NumOfInts - 1
        
        For i = 0 To loopBound
            For j = i To loopBound
                
                'Compare two entries, and if the larger one precedes the smaller one, swap them
                If sortAscending Then
                    If (m_Ints(j) < m_Ints(i)) Then SwapIndices i, j
                    
                'An opposite check is used for descending order.
                Else
                    If (m_Ints(j) > m_Ints(i)) Then SwapIndices i, j
                End If
                
            Next j
        Next i
        
    End If

End Sub

'Helper for sort functions.
Private Sub SwapIndices(ByVal idx1 As Long, ByVal idx2 As Long)
    Dim tmpLong As Long
    tmpLong = m_Ints(idx2)
    m_Ints(idx2) = m_Ints(idx1)
    m_Ints(idx1) = tmpLong
End Sub

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
    If (m_NumOfInts > 0) Then
        Dim i As Long
        For i = 0 To m_NumOfInts - 1
            Debug.Print i & ": -- " & m_Ints(i) & " -- "
        Next i
    Else
        Debug.Print " -- String stack is empty -- "
    End If
End Sub
