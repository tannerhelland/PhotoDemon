Attribute VB_Name = "Text_Support"
'***************************************************************************
'Miscellaneous functions related to specialized text handling
'Copyright 2000-2015 by Tanner Helland
'Created: 6/12/01
'Last updated: 07/May/14
'Last update: Fix bugs with IsNumberLocaleUnaware() so that very large and very small exponents are handled correctly.
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
        PDMsgBox "%1 is not a valid entry." & vbCrLf & "Please enter a value between %2 and %3.", vbExclamation + vbOKOnly + vbApplicationModal, "Invalid entry", checkVal, cMin, cMax
        RangeValid = False
    End If
End Function

'Check a Variant-type value to see if it's numeric
Public Function NumberValid(ByVal checkVal As Variant) As Boolean
    If Not IsNumeric(checkVal) Then
        PDMsgBox "%1 is not a valid entry." & vbCrLf & "Please enter a numeric value.", vbExclamation + vbOKOnly + vbApplicationModal, "Invalid entry", checkVal
        NumberValid = False
    Else
        NumberValid = True
    End If
End Function

'A pleasant combination of RangeValid and NumberValid
Public Function EntryValid(ByVal checkVal As Variant, ByVal cMin As Double, ByVal cMax As Double, Optional ByVal displayNumError As Boolean = True, Optional ByVal displayRangeError As Boolean = True) As Boolean
    If Not IsNumeric(checkVal) Then
        If displayNumError = True Then PDMsgBox "%1 is not a valid entry." & vbCrLf & "Please enter a numeric value.", vbExclamation + vbOKOnly + vbApplicationModal, "Invalid entry", checkVal
        EntryValid = False
    Else
        If (checkVal >= cMin) And (checkVal <= cMax) Then
            EntryValid = True
        Else
            If displayRangeError = True Then PDMsgBox "%1 is not a valid entry." & vbCrLf & "Please enter a value between %2 and %3.", vbExclamation + vbOKOnly + vbApplicationModal, "Invalid entry", checkVal, cMin, cMax
            EntryValid = False
        End If
    End If
End Function

'A custom CDbl function that accepts both commas and decimals as a separator
Public Function CDblCustom(ByVal srcString As String) As Double

    'Replace commas with periods
    If InStr(1, srcString, ",") > 0 Then srcString = Replace(srcString, ",", ".")
    
    'We can now use Val() to convert to Double
    If IsNumberLocaleUnaware(srcString) Then
        CDblCustom = Val(srcString)
    Else
        CDblCustom = 0
    End If

End Function

'Locale-unaware check for strings that can successfully be converted to numbers.  Thank you to
' http://stackoverflow.com/questions/18368680/vb6-isnumeric-behaviour-in-windows-8-windows-2012
' for the code.  (Note that the original function listed there is buggy!  I had to add some
' fixes for exponent strings, which the original code did not handle correctly.)
Public Function IsNumberLocaleUnaware(ByRef Expression As String) As Boolean
    
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

'For a given string, see if it has a trailing number value in parentheses (e.g. "Image (2)").  If it does have a
' trailing number, return the string with the number incremented by one.  If there is no trailing number, apply one.
Public Function incrementTrailingNumber(ByVal srcString As String) As String

    'Start by figuring out if the string is already in the format: "text (#)"
    srcString = Trim(srcString)
    
    Dim numToAppend As Long
    
    'Check the trailing character.  If it is a closing parentheses ")", we need to analyze more
    If Right(srcString, 1) = ")" Then
    
        Dim i As Long
        For i = Len(srcString) - 2 To 1 Step -1
            
            'If this char isn't a number, see if it's an initial parentheses: "("
            If Not (IsNumeric(Mid(srcString, i, 1))) Then
                
                'If it is a parentheses, then this string already has a "(#)" appended to it.  Figure out what
                ' the number inside the parentheses is, and strip that entire block from the string.
                If Mid(srcString, i, 1) = "(" Then
                
                    numToAppend = CLng(Val(Mid(srcString, i + 1, Len(srcString) - i - 1)))
                    srcString = Left(srcString, i - 2)
                    Exit For
                
                'If this character is non-numeric and NOT an initial parentheses, this string does not already have a
                ' number appended (in the expected format). Treat it like any other string and append " (2)" to it
                Else
                    numToAppend = 2
                    Exit For
                End If
                
            End If
        
        'If this character IS a number, keep scanning.
        Next i
    
    'If the string is not already in the format "text (#)", append a " (2)" to it
    Else
        numToAppend = 2
    End If
    
    incrementTrailingNumber = srcString & " (" & Trim$(CStr(numToAppend)) & ")"

End Function

'PhotoDemon's software processor requires that all parameters be passed as a string, with individual parameters separated by
' the pipe "|" character.  This function can be used to automatically assemble any number of parameters into such a string.
Public Function buildParams(ParamArray allParams() As Variant) As String

    buildParams = ""

    If UBound(allParams) >= LBound(allParams) Then
    
        Dim tmpString As String
        
        Dim i As Long
        For i = LBound(allParams) To UBound(allParams)
        
            If IsNumeric(allParams(i)) Then
                tmpString = Trim$(Str(allParams(i)))
            Else
                tmpString = Trim$(allParams(i))
            End If
        
            If Len(tmpString) <> 0 Then
                
                'Add the string (properly escaped) to the param string
                buildParams = buildParams & escapeParamCharacters(tmpString)
                
            Else
                buildParams = buildParams & " "
            End If
            
            If i < UBound(allParams) Then buildParams = buildParams & "|"
            
        Next i
    
    End If

End Function

'Given a parameter to be added to a param string, apply any necessary escaping
Public Function escapeParamCharacters(ByVal srcString As String) As String

    escapeParamCharacters = srcString
                
    'The most crucial character to escape is the pipe "|", as PD uses it to separate individual params
    If InStr(1, escapeParamCharacters, "|", vbBinaryCompare) > 0 Then
        
        'In lieu of a better escape system, use the HTML system
        escapeParamCharacters = Replace$(escapeParamCharacters, "|", "&#124;")
        
    End If
    
End Function

'Given a parameter that is ready to be removed from a passed string and reported to a calling function, replace
' any escaped characters with their correct equivalents.
Public Function unEscapeParamCharacters(ByVal srcString As String) As String

    unEscapeParamCharacters = srcString
    
    'At present, the only character PD forcibly escapes is the pipe "|"
    If InStr(1, unEscapeParamCharacters, "&#124;", vbBinaryCompare) > 0 Then
        
        'In lieu of a better escape system, use the HTML system
        unEscapeParamCharacters = Replace$(unEscapeParamCharacters, "&#124;", "|")
        
    End If
    
End Function

'As of PD 7.0, XML strings are universally used for parameter parsing.  The old pipe-delimited system is currently being
' replaced in favor of this lovely little helper function.
Public Function buildParamList(ParamArray allParams() As Variant) As String
    
    'pdParamXML handles all the messy work for us
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    
    On Error GoTo buildParamListFailure
    
    If UBound(allParams) >= LBound(allParams) Then
    
        Dim tmpName As String, tmpValue As Variant
        
        Dim i As Long
        For i = LBound(allParams) To UBound(allParams) Step 2
            
            'Parameters must be passed in a strict name/value order.  An odd number of parameters will cause crashes.
            tmpName = allParams(i)
            
            If (i + 1) <= UBound(allParams) Then
                tmpValue = allParams(i + 1)
            Else
                Err.Raise 9
            End If
            
            'Add this key/value pair to the current running param string
            cParams.addParam tmpName, tmpValue
            
        Next i
    
    End If
    
    buildParamList = cParams.getParamString
    
    Exit Function
    
buildParamListFailure:
        
    #If DEBUGMODE = 1 Then
        pdDebug.LogAction "WARNING!  buildParamList failed to create a parameter string!"
    #End If
    
    buildParamList = ""
    
End Function
