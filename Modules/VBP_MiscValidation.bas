Attribute VB_Name = "Misc_Validation"
'***************************************************************************
'Miscellaneous Functions Related to Validating User Input
'Copyright ©2000-2013 by Tanner Helland
'Created: 6/12/01
'Last updated: 03/October/12
'Last update: First build
'***************************************************************************

Option Explicit

'Validate a given text box entry.
Public Sub textValidate(ByRef srcTextBox As TextBox, Optional ByVal negAllowed As Boolean = False, Optional ByVal floatAllowed As Boolean = False)

    'Convert the input number to a string
    Dim numString As String
    numString = srcTextBox.Text
    
    'Remove any incidental white space before processing
    numString = Trim(numString)
    
    'Create a string of valid numerical characters, based on the input specifications
    Dim validChars As String
    validChars = "0123456789"
    If negAllowed Then validChars = validChars & "-"
    If floatAllowed Then validChars = validChars & "."
    
    'Make note of the cursor position so we can restore it after removing invalid text
    Dim cursorPos As Long
    cursorPos = srcTextBox.SelStart
    
    'Loop through the text box contents and remove any invalid characters
    Dim i As Long
    Dim invLoc As Long
    
    For i = 1 To Len(numString)
        
        'Compare a single character from the text box against our list of valid characters
        invLoc = InStr(validChars, Mid$(numString, i, 1))
        
        'If this character was NOT found in the list of valid characters, remove it from the string
        If invLoc = 0 Then
        
            numString = Left$(numString, i - 1) & Right$(numString, Len(numString) - i)
            
            'Modify the position of the cursor to match (so the text box maintains the same cursor position)
            If i >= (cursorPos - 1) Then cursorPos = cursorPos - 1
            
            'Move the loop variable back by 1 so the next character is properly checked
            i = i - 1
            
        End If
            
    Next i
        
    'Place the newly validated string back in the text box
    srcTextBox.Text = numString
    srcTextBox.Refresh
    srcTextBox.SelStart = cursorPos

End Sub

'Check a Long-type value to see if it falls within a given range
Public Function RangeValid(ByVal checkVal As Long, ByVal cMin As Long, ByVal cMax As Long) As Boolean
    If (checkVal >= cMin) And (checkVal <= cMax) Then
        RangeValid = True
    Else
        pdMsgBox "%1 is not a valid entry." & vbCrLf & "Please enter a value between %2 and %3.", vbExclamation + vbOKOnly + vbApplicationModal, "Invalid entry", checkVal, cMin, cMax
        RangeValid = False
    End If
End Function

'Check a Variant-type value to see if it's numeric
Public Function NumberValid(ByVal checkVal) As Boolean
    If Not IsNumeric(checkVal) Then
        pdMsgBox "%1 is not a valid entry." & vbCrLf & "Please enter a numeric value.", vbExclamation + vbOKOnly + vbApplicationModal, "Invalid entry", checkVal
        NumberValid = False
    Else
        NumberValid = True
    End If
End Function

'A pleasant combination of RangeValid and NumberValid
Public Function EntryValid(ByVal checkVal As Variant, ByVal cMin As Long, ByVal cMax As Long, Optional ByVal displayNumError As Boolean = True, Optional ByVal displayRangeError As Boolean = True) As Boolean
    If Not IsNumeric(checkVal) Then
        If displayNumError = True Then pdMsgBox "%1 is not a valid entry." & vbCrLf & "Please enter a numeric value.", vbExclamation + vbOKOnly + vbApplicationModal, "Invalid entry", checkVal
        EntryValid = False
    Else
        If (checkVal >= cMin) And (checkVal <= cMax) Then
            EntryValid = True
        Else
            If displayRangeError = True Then pdMsgBox "%1 is not a valid entry." & vbCrLf & "Please enter a value between %2 and %3.", vbExclamation + vbOKOnly + vbApplicationModal, "Invalid entry", checkVal, cMin, cMax
            EntryValid = False
        End If
    End If
End Function
