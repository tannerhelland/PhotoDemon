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
'Copyright 2014-2015 by Tanner Helland
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
'Special thanks to the following individuals:
'   - Ellis Dee, for specialized sort function "Snake Sort" which is well-suited to partially sorted arrays.  Original source here:
'       http://www.vbforums.com/showthread.php?473677-VB6-Sorting-algorithms-%28sort-array-sorting-arrays%29&p=2933240#post2933240
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private m_Strings() As String
Private m_NumOfStrings As Long
Private Const INIT_STACK_SIZE = 16

'Add a string to the stack.  Return value is the index of the added location (which can be used to infer the number of strings
' in the stack, obviously).
Public Function AddString(ByRef srcString As String) As Long

    'Resize the stack as necessary
    If m_NumOfStrings > UBound(m_Strings) Then ReDim Preserve m_Strings(0 To m_NumOfStrings * 2 - 1) As String
    
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
' after a pop, manually call trimStack().  (But seriously - don't do this unless you really need to, as the performance
' implications are severe.)
Public Function PopString(ByRef dstString As String) As Boolean
    
    If m_NumOfStrings > 0 Then
        m_NumOfStrings = m_NumOfStrings - 1
        dstString = m_Strings(m_NumOfStrings)
        PopString = True
    Else
        PopString = False
    End If
    
End Function

'Return the size of the stack
Public Function getNumOfStrings() As Long
    getNumOfStrings = m_NumOfStrings
End Function

'Trim the stack to its exact size.  IMPORTANT NOTE!  Don't do this any more than you have to, as it's not performance-friendly.
Public Sub trimStack()
    ReDim Preserve m_Strings(0 To m_NumOfStrings - 1) As String
End Sub

'Retrieve a string from the stack, with optional support for locale invariant conversions (when the caller expects the string
' to represent a number of some sort)
Public Function GetString(ByVal strIndex As Long, Optional ByVal assumeLocaleInvariantNumber As Boolean = False) As String

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
Public Function GetStringPointer(ByVal strIndex As Long) As Long

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
    
    Dim Negative As Boolean
    Dim Number As Boolean
    Dim Period As Boolean
    Dim Positive As Boolean
    Dim Exponent As Boolean
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

'This function may seem like a ridiculous addition, but it's actually very helpful in PD.  pdStringStack is used by pdFSO when retrieving all subfolders
' inside some base folder.  When performing something like a tree copy, I like to pre-sort the subfolder list by length.  This greatly simplifies the
' code required to create the new folder tree prior to performing the copy; creating the folders in advance greatly accelerates the copy operation, as we
' don't have to perform "do my folders exist?" checks on every damn file.
Public Sub sortStackByLength(Optional ByVal sortAscending As Boolean = True)

    If m_NumOfStrings > 1 Then
    
        'Given PD's standard use-case (subfolder trees, as mentioned above), the existing stack order is typically pretty close to sorted.  This saves us
        ' from needing an elaborate search algorithm; instead, a simple in-place insertion sort performs very well.
        Dim i As Long, j As Long, loopBound As Long
        loopBound = m_NumOfStrings - 1
        
        'Loop through all entries in the stack, sorting them as we go
        For i = 0 To loopBound
            For j = 0 To loopBound
                
                'Compare two entries, and if the longer one precedes the shorter one, swap them
                If sortAscending Then
                    If Len(m_Strings(i)) < Len(m_Strings(j)) Then swapStrings i, j
                
                'An opposite check is used for descending order.
                Else
                    If Len(m_Strings(i)) > Len(m_Strings(j)) Then swapStrings i, j
                End If
                
            Next j
        Next i
        
    End If

End Sub

'Helper for string sorting, above.  This is not optimized at all, but it's sufficient for PD's current usage...
Private Sub swapStrings(ByVal strIndex1 As Long, ByVal strIndex2 As Long)
    Dim tmpString As String
    tmpString = m_Strings(strIndex1)
    m_Strings(strIndex1) = m_Strings(strIndex2)
    m_Strings(strIndex2) = tmpString
End Sub

'Modified sort function based off "Snake Sort" by Ellis Dee.  Original source here:
' http://www.vbforums.com/showthread.php?473677-VB6-Sorting-algorithms-%28sort-array-sorting-arrays%29&p=2933240#post2933240
Public Sub SortAlphabetically(Optional ByVal removeDuplicates As Boolean = False)
    
    Dim i As Long
    Dim iMin As Long
    Dim iMax As Long
    Dim lngIndex() As Long
    Dim lngLevel As Long
    Dim lngOldLevel As Long
    Dim lngNewLevel As Long
    Dim varMirror As Variant
    Dim lngDirection As Long
    Dim blnMirror As Boolean
    Dim varSwap As String
    
    iMin = LBound(m_Strings)
    iMax = UBound(m_Strings)
    
    ReDim lngIndex((iMax - iMin + 3) \ 2)
    lngIndex(0) = iMin
    i = iMin
    
    'Initial loop: locate cutoffs for each ordered section
    Do Until i >= iMax
    
        Select Case lngDirection
            Case 1
                Do Until i = iMax
                    If StrComp(m_Strings(i), (i + 1)) = 1 Then Exit Do
                    i = i + 1
                Loop
            Case -1
                Do Until i = iMax
                    If StrComp(m_Strings(i), m_Strings(i + 1)) = -1 Then Exit Do
                    i = i + 1
                Loop
            Case Else
                Do Until i = iMax
                    If StrComp(m_Strings(i), m_Strings(i + 1)) <> 0 Then Exit Do
                    i = i + 1
                Loop
                If i = iMax Then lngDirection = 1
        End Select
        
        If lngDirection = 0 Then
            If StrComp(m_Strings(i), m_Strings(i + 1)) = 1 Then
                lngDirection = -1
            Else
                lngDirection = 1
            End If
        Else
            lngLevel = lngLevel + 1
            lngIndex(lngLevel) = i * lngDirection
            lngDirection = 0
        End If
        
        i = i + 1
        
    Loop
    
    If Abs(lngIndex(lngLevel)) < iMax Then
        If lngDirection = 0 Then lngDirection = 1
        lngLevel = lngLevel + 1
        lngIndex(lngLevel) = i * lngDirection
    End If
    
    'If the list is already sorted, exit
    If lngLevel <= 1 Then
    
        'If sorted descending, reverse before exiting
        If lngIndex(lngLevel) < 0 Then
            For i = 0 To (iMax - iMin) \ 2
                varSwap = m_Strings(iMin + i)
                m_Strings(iMin + i) = m_Strings(iMax - i)
                m_Strings(iMax - i) = varSwap
            Next
        End If
        
        Exit Sub
        
    End If
    
    'Main loop - merge section pairs together until only one section left
    ReDim varMirror(iMin To iMax)
    
    Do Until lngLevel = 1
    
        lngOldLevel = lngLevel
        
        For lngLevel = 1 To lngLevel - 1 Step 2
        
            If blnMirror Then
                SnakeSortMerge varMirror, lngIndex(lngLevel - 1), lngIndex(lngLevel), lngIndex(lngLevel + 1), m_Strings
            Else
                SnakeSortMerge m_Strings, lngIndex(lngLevel - 1), lngIndex(lngLevel), lngIndex(lngLevel + 1), varMirror
            End If
            
            lngNewLevel = lngNewLevel + 1
            lngIndex(lngNewLevel) = Abs(lngIndex(lngLevel + 1))
            
        Next
        
        If lngOldLevel Mod 2 = 1 Then
        
            If blnMirror Then
                For i = lngIndex(lngNewLevel) + 1 To iMax
                    m_Strings(i) = varMirror(i)
                Next
            Else
                For i = lngIndex(lngNewLevel) + 1 To iMax
                    varMirror(i) = m_Strings(i)
                Next
            End If
            
            lngNewLevel = lngNewLevel + 1
            lngIndex(lngNewLevel) = lngIndex(lngOldLevel)
            
        End If
        
        lngLevel = lngNewLevel
        lngNewLevel = 0
        blnMirror = Not blnMirror
        
    Loop
    
    'Copy back to main array if necessary
    If blnMirror Then
        For i = iMin To iMax
            m_Strings(i) = varMirror(i)
        Next
    End If
    
    'If the user wants duplicates removed, do so now
    If removeDuplicates Then
        
        Dim itemsRemoved As Long
        itemsRemoved = 0
        
        Dim itemCount As Long
        itemCount = m_NumOfStrings - 1
        
        Dim j As Long
        
        For i = 0 To itemCount
            
            'If we've reached the point where the array ends due to shifting, exit now
            If i = itemCount - itemsRemoved Then Exit For
            
            'If this string and the string above it match, shift everything above it downward
            If StrComp(m_Strings(i), m_Strings(i + 1), vbBinaryCompare) = 0 Then
                
                #If DEBUGMODE = 1 Then
                    Debug.Print "Duplicate string found: " & m_Strings(i)
                #End If
                
                For j = i To UBound(m_Strings) - 1
                    m_Strings(j) = m_Strings(j + 1)
                Next j
                
                itemsRemoved = itemsRemoved + 1
                
            End If
            
        Next i
        
        'If one or more items were removed, mark the new array size accordingly
        If itemsRemoved > 0 Then m_NumOfStrings = m_NumOfStrings - itemsRemoved
        
    End If
    
End Sub
 
'Helper function to Snake Sort, above.  Modified version of a function originally written by Ellis Dee.  Original source here:
' http://www.vbforums.com/showthread.php?473677-VB6-Sorting-algorithms-%28sort-array-sorting-arrays%29&p=2933240#post2933240
Private Sub SnakeSortMerge(pvarSource As Variant, plngLeft As Long, plngMid As Long, plngRight As Long, pvarDest As Variant)
    
    Dim l As Long
    Dim lMin As Long
    Dim lMax As Long
    Dim LStep As Long
    Dim r As Long
    Dim RMin As Long
    Dim RMax As Long
    Dim RStep As Long
    Dim O As Long
    
    If plngLeft <> 0 Then O = Abs(plngLeft) + 1
    
    If plngMid > 0 Then
        lMin = O
        lMax = Abs(plngMid)
        LStep = 1
    Else
        lMin = Abs(plngMid)
        lMax = O
        LStep = -1
    End If
    
    If plngRight > 0 Then
        RMin = Abs(plngMid) + 1
        RMax = Abs(plngRight)
        RStep = 1
    Else
        RMin = Abs(plngRight)
        RMax = Abs(plngMid) + 1
        RStep = -1
    End If
    
    l = lMin
    r = RMin
    
    Do
    
        If StrComp(pvarSource(l), pvarSource(r)) <= 0 Then
            
            pvarDest(O) = pvarSource(l)
            
            If l = lMax Then
                For r = r To RMax Step RStep
                    O = O + 1
                    pvarDest(O) = pvarSource(r)
                Next
                Exit Do
            
            End If
            
            l = l + LStep
            
        Else
        
            pvarDest(O) = pvarSource(r)
            
            If r = RMax Then
                For l = l To lMax Step LStep
                    O = O + 1
                    pvarDest(O) = pvarSource(l)
                Next
                Exit Do
            End If
            
            r = r + RStep
            
        End If
        
        O = O + 1
        
    Loop
    
End Sub

'Clone another string stack
Public Sub cloneStack(ByRef stackToClone As pdStringStack)
    
    'Initialize this stack to the size of the target
    Me.resetStack stackToClone.getNumOfStrings
    
    'Copy all strings
    Dim i As Long
    For i = 0 To stackToClone.getNumOfStrings - 1
        Me.AddString stackToClone.GetString(i)
    Next i
    
End Sub

'Clear the current stack.  An optional stack size can be passed; if it is not passed, it will default to INIT_STACK_SIZE
Public Sub resetStack(Optional ByVal newStackSize As Long = INIT_STACK_SIZE)
    
    On Error GoTo FailsafeReset
    
    'Failsafe bounds check
    If newStackSize <= 0 Then newStackSize = INIT_STACK_SIZE
    
    'Reset the array (but only if necessary!)
    If m_NumOfStrings = 0 Then
        ReDim m_Strings(0 To newStackSize - 1) As String
    Else
        If UBound(m_Strings) <> newStackSize - 1 Then ReDim m_Strings(0 To newStackSize - 1) As String
    End If
    
    m_NumOfStrings = 0
    
    Exit Sub
    
FailsafeReset:
    If newStackSize <= 0 Then newStackSize = INIT_STACK_SIZE
    ReDim m_Strings(0 To newStackSize - 1) As String
    
End Sub

Private Sub Class_Initialize()
    
    'Always start with an initialized array
    resetStack
        
End Sub

Private Sub Class_Terminate()
    resetStack
End Sub

'DEBUG ONLY!  I sometimes find it helpful to investigate the contents of the stack.  This function makes it trivial to do so.
' I also append "--" to the start and end of the string, to help me see if extra whitespace chars are present.
Public Sub DEBUG_dumpResultsToImmediateWindow()
    
    If m_NumOfStrings > 0 Then
        Dim i As Long
        For i = 0 To m_NumOfStrings - 1
            Debug.Print i & ": -- " & m_Strings(i) & " -- "
        Next i
    Else
        Debug.Print " -- String stack is empty -- "
    End If
End Sub

