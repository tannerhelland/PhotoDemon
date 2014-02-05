Attribute VB_Name = "Text_Validation"
'***************************************************************************
'Miscellaneous Functions Related to Validating User Input
'Copyright ©2000-2014 by Tanner Helland
'Created: 6/12/01
'Last updated: 05/February/14
'Last update: New CDblCustom() function to replace CDbl.  It is designed for use with my internal XML writing routines,
'              and will check for either comma or period as a decimal separator.  This allows the user to change locale
'              without all their presets breaking.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
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
Public Function RangeValid(ByVal checkVal As Variant, ByVal cMin As Double, ByVal cMax As Double) As Boolean
    If (checkVal >= cMin) And (checkVal <= cMax) Then
        RangeValid = True
    Else
        pdMsgBox "%1 is not a valid entry." & vbCrLf & "Please enter a value between %2 and %3.", vbExclamation + vbOKOnly + vbApplicationModal, "Invalid entry", checkVal, cMin, cMax
        RangeValid = False
    End If
End Function

'Check a Variant-type value to see if it's numeric
Public Function NumberValid(ByVal checkVal As Variant) As Boolean
    If Not IsNumeric(checkVal) Then
        pdMsgBox "%1 is not a valid entry." & vbCrLf & "Please enter a numeric value.", vbExclamation + vbOKOnly + vbApplicationModal, "Invalid entry", checkVal
        NumberValid = False
    Else
        NumberValid = True
    End If
End Function

'A pleasant combination of RangeValid and NumberValid
Public Function EntryValid(ByVal checkVal As Variant, ByVal cMin As Double, ByVal cMax As Double, Optional ByVal displayNumError As Boolean = True, Optional ByVal displayRangeError As Boolean = True) As Boolean
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

'A custom CDbl function that accepts both commas and decimals as a separator
Public Function CDblCustom(ByVal srcString As String) As Double

    'Replace commas with periods
    If InStr(1, srcString, ",") > 0 Then srcString = Replace(srcString, ",", ".")
    
    'We can now use Val() to convert to Double
    If IsNumeric(srcString) Then
        CDblCustom = Val(srcString)
    Else
        CDblCustom = 0
    End If

End Function
