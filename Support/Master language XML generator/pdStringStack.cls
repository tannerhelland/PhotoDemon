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
'Last updated: 05/February/15
'Last update: initial build
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

'Clear the current stack.
Public Sub resetStack()
    ReDim m_Strings(0 To INIT_STACK_SIZE - 1) As String
    m_NumOfStrings = 0
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

