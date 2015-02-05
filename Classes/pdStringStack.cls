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
' allocating their own string.
'
'The function was designed to make popping the entire stack convenient (e.g. Do While strStack.PopString(tmpString)...)
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
Public Function getString(ByVal strIndex As Long, Optional ByVal assumeLocaleInvariantNumber As Boolean = False) As String

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
            
            If Text_Support.IsNumberLocaleUnaware(testString) Then
                getString = Trim$(Str(Val(testString)))
            Else
                getString = testString
            End If
        
        Else
            getString = m_Strings(strIndex)
        End If
    Else
        Debug.Print "WARNING!  Someone asked pdStringStack for a string outside stack bounds.  Fix this!"
    End If

End Function

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

'DEBUG ONLY!
Public Sub DEBUG_dumpResultsToImmediateWindow()
    
    If m_NumOfStrings > 0 Then
        Dim i As Long
        For i = 0 To m_NumOfStrings - 1
            Debug.Print " -- " & m_Strings(i) & " -- "
        Next i
    Else
        Debug.Print " -- String stack is empty -- "
    End If
End Sub

