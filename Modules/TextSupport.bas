Attribute VB_Name = "TextSupport"
'***************************************************************************
'Miscellaneous functions related to specialized text handling
'Copyright 2000-2026 by Tanner Helland
'Created: 6/12/01
'Last updated: 08/March/24
'Last update: new page range parser (added as part of PDF import)
'
'PhotoDemon interacts with a *lot* of text input.  This module contains various bits of text support code,
' typically based around tasks like "validate a user's text input" or "convert arbitrary text input
' into a usable numeric value".
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Check a Long-type value to see if it falls within a given range
Public Function RangeValid(ByVal checkVal As Variant, ByVal cMin As Double, ByVal cMax As Double) As Boolean
    If (checkVal >= cMin) And (checkVal <= cMax) Then
        RangeValid = True
    Else
        PDMsgBox "%1 is not a valid entry." & vbCrLf & "Please enter a value between %2 and %3.", vbExclamation Or vbOKOnly, "Invalid entry", checkVal, cMin, cMax
        RangeValid = False
    End If
End Function

'Check a Variant-type value to see if it's numeric
Public Function NumberValid(ByVal checkVal As Variant) As Boolean
    If (Not IsNumeric(checkVal)) Then
        PDMsgBox "%1 is not a valid entry." & vbCrLf & "Please enter a numeric value.", vbExclamation Or vbOKOnly, "Invalid entry", checkVal
        NumberValid = False
    Else
        NumberValid = True
    End If
End Function

'A pleasant combination of RangeValid and NumberValid
Public Function EntryValid(ByVal checkVal As Variant, ByVal cMin As Double, ByVal cMax As Double, Optional ByVal displayNumError As Boolean = True, Optional ByVal displayRangeError As Boolean = True) As Boolean
    If Not IsNumeric(checkVal) Then
        If displayNumError Then PDMsgBox "%1 is not a valid entry." & vbCrLf & "Please enter a numeric value.", vbExclamation Or vbOKOnly, "Invalid entry", checkVal
        EntryValid = False
    Else
        If (checkVal >= cMin) And (checkVal <= cMax) Then
            EntryValid = True
        Else
            If displayRangeError Then PDMsgBox "%1 is not a valid entry." & vbCrLf & "Please enter a value between %2 and %3.", vbExclamation Or vbOKOnly, "Invalid entry", checkVal, cMin, cMax
            EntryValid = False
        End If
    End If
End Function

'PD uses this (cheap, possibly ill-conceived) custom CDbl() function to coerce arbitrary floating-point
' text into a proper numeric type, regardless of locale settings.  The function *will* fail if thousands
' separators are present - the text *must* be limited to a single separator of standard type (".", ",",
' and the Arabic decimal separator 0x066b are currently allowed).
'
'This function is used to work around one of the more annoying aspects of portable software - the possibility
' of underlying system/user locale data changing arbitrarily, and possibly changing in ways that inconvenience
' the user (e.g. a U.S. traveler trying to use a portable app while on vacation in the E.U.).  The hope is
' that it allows users to enter floating-point values however they want, without worrying about system
' settings they may/may not have control over.
Public Function CDblCustom(ByVal srcString As String) As Double
    
    'Start by normalizing the incoming string.  This will convert any non-standard Unicode chars
    ' (e.g. weird extended-range numeric representations) into their standard 0-9 equivalent.
    srcString = Strings.StringNormalize(srcString)
    
    'Coerce arbitrary decimal separators into the standard, locale-invariant "."
    If (InStr(1, srcString, ",", vbBinaryCompare) <> 0) Then srcString = Replace$(srcString, ",", ".", , , vbBinaryCompare)
    If (InStr(1, srcString, ChrW$(&H66B&), vbBinaryCompare) <> 0) Then srcString = Replace$(srcString, ChrW$(&H66B&), ".", , , vbBinaryCompare)
    
    'Perform a final check to make sure the string looks like a valid, locale-invariant number;
    ' if it does, use VB's built-in Val() to convert to Double.
    If TextSupport.IsNumberLocaleUnaware(srcString) Then
        CDblCustom = Val(srcString)
    Else
        CDblCustom = 0#
    End If

End Function

'Convert a page-range string (e.g. "1, 3, 5-10") to a stack of integers (e.g. "1, 3, 5, 6, 7, 8, 9, 10").
' Returns TRUE if the page range was successfully parsed and at least one valid page exists;
' FALSE if there were errors.  (You can use this return for basic input validation.)
'
'Because the user will supply pages on the range [1, n] but internally we often reference on the range [1, n-1]
' you can pass -1 to the "modifyAllValuesByThis" parameter to subtract one from all pages found in the string
Public Function ConvertPageRangeToStack(ByRef srcRange As String, ByRef dstStack As pdStack, Optional ByVal modifyAllValuesByThis As Long = 0) As Boolean
    
    If (dstStack Is Nothing) Then Set dstStack = New pdStack
    
    'Assume failure
    ConvertPageRangeToStack = False
    
    'I want this function to be a little more resilient and user-friendly than various pre-built page range
    ' parsers found online.
    
    'For example, users should be allowed to use spaces as delimiters (e.g. "1 3 5-10") and the parser
    ' should handle that gracefully.
    
    'To simplify parsing, let's turn all commas into spaces (and standardize other whitespace chars)
    ' and then tokenize against whitespace, ignoring everything that isn't a number or a dash/hyphen
    Dim processedRange As String
    processedRange = srcRange
    If (InStr(1, processedRange, vbTab, vbBinaryCompare) <> 0) Then processedRange = Replace$(processedRange, vbTab, " ", 1, -1, vbBinaryCompare)
    If (InStr(1, processedRange, vbCr, vbBinaryCompare) <> 0) Then processedRange = Replace$(processedRange, vbCr, " ", 1, -1, vbBinaryCompare)
    If (InStr(1, processedRange, vbLf, vbBinaryCompare) <> 0) Then processedRange = Replace$(processedRange, vbLf, " ", 1, -1, vbBinaryCompare)
    If (InStr(1, processedRange, ",", vbBinaryCompare) <> 0) Then processedRange = Replace$(processedRange, ",", " ", 1, -1, vbBinaryCompare)
    If (InStr(1, processedRange, ";", vbBinaryCompare) <> 0) Then processedRange = Replace$(processedRange, ";", " ", 1, -1, vbBinaryCompare)
    
    'Internationalized list separators follow
    'https://en.wikipedia.org/wiki/Comma#Languages_other_than_Western_European
    If (InStr(1, processedRange, ChrW$(&H3001&), vbBinaryCompare) <> 0) Then processedRange = Replace$(processedRange, ChrW$(&H3001&), " ", 1, -1, vbBinaryCompare)
    
    'Ensure hyphens are separated by whitespace too
    Const HYPHEN_CHAR As String = "-"
    If (InStr(1, processedRange, HYPHEN_CHAR, vbBinaryCompare) <> 0) Then processedRange = Replace$(processedRange, HYPHEN_CHAR, " - ", 1, -1, vbBinaryCompare)
    
    'Condense any/all double-plus space occurrences down to single spaces
    Const DOUBLE_SPACE As String = "  ", SINGLE_SPACE As String = " "
    Do While (InStr(1, processedRange, DOUBLE_SPACE, vbBinaryCompare) <> 0)
        processedRange = Replace$(processedRange, DOUBLE_SPACE, SINGLE_SPACE, 1, -1, vbBinaryCompare)
    Loop
    
    'With whitespace standardized, we can now split on spaces and have a nicely tokenized list of (hopefully)
    ' just numbers and just hyphens.  This makes it trivial to process.
    Dim listOfTokens() As String
    listOfTokens = Split(processedRange, SINGLE_SPACE, -1, vbBinaryCompare)
    
    If (UBound(listOfTokens) >= LBound(listOfTokens)) Then
        
        Dim i As Long, j As Long, lastNumber As Long, nextNumber As Long
        i = LBound(listOfTokens)
        
        Do While (i <= UBound(listOfTokens))
            
            'Add numbers directly to the destination stack (and remember this as the "last number encountered")
            If IsNumeric(listOfTokens(i)) Then
                lastNumber = CLng(listOfTokens(i))
                dstStack.AddInt lastNumber + modifyAllValuesByThis
            
            'Hyphens denote ranges
            ElseIf (listOfTokens(i) = HYPHEN_CHAR) Then
                
                'Ensure that another number follows this one.
                If (i < UBound(listOfTokens)) Then
                    
                    'Ensure the next entry is numeric
                    If IsNumeric(listOfTokens(i + 1)) Then
                        
                        nextNumber = CLng(listOfTokens(i + 1))
                        
                        'Ensure the next number is *larger* than the current one.
                        If (nextNumber > lastNumber) Then
                            
                            'Add all numbers between the last number (exclusive) and this one to the stack
                            For j = (lastNumber + 1) To nextNumber
                                dstStack.AddInt j + modifyAllValuesByThis
                            Next j
                            
                            'Increment the token pointer *to* the trailing number; the i + 1 at the end of this
                            ' loop will increment it again, to ensure it points past this hyphen group
                            i = i + 1
                            
                        '/number after hyphen is <= the number before the hyphen; failure!
                        Else
                            Exit Function
                        End If
                        
                    '/next token isn't numeric; failure!
                    Else
                        Exit Function
                    End If
                    
                '/no following token; failure!
                Else
                    Exit Function
                End If
            
            'Any other entries are invalid; failure!
            Else
                Exit Function
            End If
            
            'Advance to the next token
            i = i + 1
            
        Loop
        
        'Return success if at least one valid page was found
        ConvertPageRangeToStack = (dstStack.GetNumOfInts > 0)
    
    'No valid tokens found; abandon ship
    Else
        ConvertPageRangeToStack = False
    End If
    
End Function

'Because VB6's built-in Format$() function uses locale-specific decimal signs, and it exhibits stupid
' behavior with things like a trailing decimal point, this shorthand function can be used to produce
' nicely formatted strings for floating-point values in a locale-independent manner.  Note that it
' does *not* handle thousands separators, by design.
Public Function FormatInvariant(ByVal srcValue As Variant, ByVal newFormat As String) As String
    FormatInvariant = Format$(srcValue, newFormat)
    If (InStr(1, FormatInvariant, ",") <> 0) Then FormatInvariant = Replace$(FormatInvariant, ",", ".")
    If (Right$(FormatInvariant, 1) = ".") Then FormatInvariant = Left$(FormatInvariant, Len(FormatInvariant) - 1)
End Function

'Locale-unaware check for strings that can successfully be converted to numbers.  Thank you to
' http://stackoverflow.com/questions/18368680/vb6-isnumeric-behaviour-in-windows-8-windows-2012
' for the code.  (Note that the original function listed there is buggy!  I had to add fixes for
' exponent strings because the original code did not handle them correctly.)
Public Function IsNumberLocaleUnaware(ByRef srcExpression As String) As Boolean
    
    Dim numIsNegative As Boolean, numIsPositive As Boolean
    Dim txtIsNumber As Boolean, txtIsPeriod As Boolean, txtIsExponent As Boolean
    
    Dim x As Long
    For x = 1 To Len(srcExpression)
    
        Select Case Mid$(srcExpression, x, 1)
            
            Case "0" To "9"
                txtIsNumber = True
            Case "-"
                If txtIsPeriod Or txtIsNumber Or numIsNegative Or numIsPositive Then Exit Function
                numIsNegative = True
            Case "."
                If (txtIsPeriod Or txtIsExponent) Then Exit Function
                txtIsPeriod = True
            Case "E", "e"
                If (Not txtIsNumber) Then Exit Function
                If txtIsExponent Then Exit Function
                txtIsExponent = True
                txtIsNumber = False
                numIsNegative = False
                txtIsPeriod = False
            Case "+"
                If (Not txtIsExponent) Then Exit Function
                If (txtIsNumber Or numIsNegative Or numIsPositive) Then Exit Function
                numIsPositive = True
            Case " ", vbTab, vbVerticalTab, vbCr, vbLf, vbFormFeed
                If (txtIsPeriod Or txtIsNumber Or txtIsExponent Or numIsNegative) Then Exit Function
            Case Else
                Exit Function
        
        End Select
        
    Next x
    
    IsNumberLocaleUnaware = txtIsNumber
    
End Function

'For a given string, see if it has a trailing number value in parentheses (e.g. "Image (2)").  If it does have a
' trailing number, return the string with the number incremented by one.  If there is no trailing number, apply one.
Public Function IncrementTrailingNumber(ByVal srcString As String) As String

    'Start by figuring out if the string is already in the format: "text (#)"
    srcString = Trim$(srcString)
    
    Dim numToAppend As Long
    
    'Check the trailing character.  If it is a closing parentheses ")", we need to analyze more
    If Strings.StringsEqual(Right$(srcString, 1), ")", False) Then
    
        Dim i As Long
        For i = Len(srcString) - 2 To 1 Step -1
            
            'If this char isn't a number, see if it's an initial parentheses: "("
            If Not (IsNumeric(Mid$(srcString, i, 1))) Then
                
                'If it is a parentheses, then this string already has a "(#)" appended to it.  Figure out what
                ' the number inside the parentheses is, and strip that entire block from the string.
                If Strings.StringsEqual(Mid$(srcString, i, 1), "(", False) Then
                
                    numToAppend = CLng(Mid$(srcString, i + 1, Len(srcString) - i - 1)) + 1
                    srcString = Left$(srcString, i - 1)
                    Exit For
                
                'If this character is non-numeric and NOT an initial parentheses, this string does not already have a
                ' number appended (in the expected format). Treat it like any other string and append " (2)" to it
                Else
                    numToAppend = 2
                    srcString = srcString & " "
                    Exit For
                End If
                
            End If
        
        'If this character IS a number, keep scanning.
        Next i
    
    'If the string is not already in the format "text (#)", append a " (2)" to it
    Else
        numToAppend = 2
        srcString = srcString & " "
    End If
    
    IncrementTrailingNumber = srcString & "(" & CStr(numToAppend) & ")"

End Function

'As of PD 7.0, XML strings are universally used for parameter parsing.
Public Function BuildParamList(ParamArray allParams() As Variant) As String
    
    'pdSerialize handles all the messy work for us
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    On Error GoTo BuildParamListFailure
    
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
            cParams.AddParam tmpName, tmpValue
            
        Next i
    
    End If
    
    BuildParamList = cParams.GetParamString
    
    Exit Function
    
BuildParamListFailure:
    
    PDDebug.LogAction "WARNING!  buildParamList failed to create a parameter string!"
    BuildParamList = vbNullString
    
End Function

'Given two strings - a test candidate string, and a string comprised only of valid characters - return TRUE if the
' test string is comprised only of characters from the valid character list.
Public Function ValidateCharacters(ByVal srcText As String, ByVal listOfValidChars As String, Optional ByVal compareCaseInsensitive As Boolean = True) As Boolean
    
    ValidateCharacters = True
    
    'For case-insensitive comparisons, lcase both strings in advance
    If compareCaseInsensitive Then
        srcText = LCase$(srcText)
        listOfValidChars = LCase$(listOfValidChars)
    End If
    
    'I'm not sure if there's a better way to do this, but basically, we need to individually check each character
    ' in the string against the valid char list.  If a character is NOT located in the valid char list, return FALSE,
    ' and if the whole string checks out, return TRUE.
    Dim i As Long
    For i = 1 To Len(srcText)
        
        'If this invalid character exists in the target string, return FALSE
        If (InStr(1, listOfValidChars, Mid$(srcText, i, 1), vbBinaryCompare) = 0) Then
            ValidateCharacters = False
            Exit For
        End If
        
    Next i
    
End Function

'Return TRUE if a test string is comprised only of valid hex chars (0-9, a-f).
Public Function ValidateHexChars(ByRef srcText As String) As Boolean
    
    ValidateHexChars = True
    
    'I'm not sure if there's a better way to do this, but basically, we need to individually
    ' check each character in the string against the valid hex char list.  If a character is
    ' NOT located in the valid char list, return FALSE, and if the whole string checks out,
    ' return TRUE.
    Dim i As Long, chrTest As Long
    For i = 1 To Len(srcText)
        
        chrTest = AscW(Mid$(srcText, i, 1))
        
        'If this invalid character exists in the target string, return FALSE
        If (chrTest < 48) Or (chrTest > 103) Or ((chrTest > 57) And (chrTest < 65)) Or ((chrTest > 70) And (chrTest < 97)) Then
            ValidateHexChars = False
            Exit For
        End If
        
    Next i
    
End Function
