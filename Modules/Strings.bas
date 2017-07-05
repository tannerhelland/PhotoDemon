Attribute VB_Name = "Strings"
'***************************************************************************
'Additional string support functions
'Copyright 2017-2017 by Tanner Helland
'Created: 13/June/17
'Last updated: 13/June/17
'Last update: initial build; string functions are currently spread across a number of different objects,
'             and I'd like to perform certain actions without needing to instantiate a class.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Locale identifiers; these need to be specified for certain string functions
Public Enum PD_LocaleIdentifier
    pdli_Invariant = &H7F
    pdli_SystemDefault = &H800
    pdli_UserDefault = &H400
End Enum

#If False Then
    Private Const pdli_Invariant = &H7F, pdli_SystemDefault = &H800, pdli_UserDefault = &H400
#End If

Private Enum StrCmpFlags
    LINGUISTIC_IGNORECASE = &H10        'Ignore case, as linguistically appropriate.
    LINGUISTIC_IGNOREDIACRITIC = &H20   'Ignore nonspacing characters, as linguistically appropriate.
    NORM_IGNORECASE = &H1               'Ignore case. For many scripts (notably Latin scripts), NORM_IGNORECASE coincides with LINGUISTIC_IGNORECASE.
    NORM_IGNORENONSPACE = &H2           'Ignore nonspacing characters. For many scripts (notably Latin scripts), NORM_IGNORENONSPACE coincides with LINGUISTIC_IGNOREDIACRITIC.
    NORM_IGNORESYMBOLS = &H4            'Ignore symbols and punctuation
    NORM_IGNOREWIDTH = &H20000          'Chinese/Japanese: ignore the difference between half-width and full-width characters
    NORM_IGNOREKANATYPE = &H10000       'Do not differentiate between hiragana and katakana characters
    NORM_LINGUISTIC_CASING = &H8000000  'Use language rules (not filesystem rules)
    SORT_DIGITSASNUMBERS = &H8          'Win 7+ only
    SORT_STRINGSORT = &H1000            'Treat punctuation as symbols
End Enum

#If False Then
    Private Const LINGUISTIC_IGNORECASE = &H10, LINGUISTIC_IGNOREDIACRITIC = &H20, NORM_IGNORECASE = &H1, NORM_IGNORENONSPACE = &H2, NORM_IGNORESYMBOLS = &H4, NORM_IGNOREWIDTH = &H20000, NORM_IGNOREKANATYPE = &H10000, NORM_LINGUISTIC_CASING = &H8000000, SORT_DIGITSASNUMBERS = &H8, SORT_STRINGSORT = &H1000
#End If

'While not technically Uniscribe-specific, this class wraps some other Unicode bits as a convenience
Public Enum PD_STRING_REMAP
    PDSR_NONE = 0
    PDSR_LOWERCASE = 1
    PDSR_UPPERCASE = 2
    PDSR_HIRAGANA = 3
    PDSR_KATAKANA = 4
    PDSR_SIMPLE_CHINESE = 5
    PDSR_TRADITIONAL_CHINESE = 6
    PDSR_TITLECASE_WIN7 = 7
End Enum

Private Declare Function CompareStringW Lib "kernel32" (ByVal lcID As PD_LocaleIdentifier, ByVal cmpFlags As StrCmpFlags, ByVal ptrToStr1 As Long, ByVal str1Len As Long, ByVal ptrToStr2 As Long, ByVal str2Len As Long) As Long
Private Declare Function CompareStringOrdinal Lib "kernel32" (ByVal ptrToStr1 As Long, ByVal str1Len As Long, ByVal ptrToStr2 As Long, ByVal str2Len As Long, ByVal bIgnoreCase As Long) As Long

'High-performance string equality function.  Returns TRUE/FALSE for equality, with support for case-insensitivity.
Public Function StringsEqual(ByVal firstString As String, ByVal secondString As String, Optional ByVal ignoreCase As Boolean = False) As Boolean
    
    'Cheat and compare length first
    If (Len(firstString) <> Len(secondString)) Then
        StringsEqual = False
    Else
        If ignoreCase Then
            If g_IsVistaOrLater Then
                StringsEqual = (CompareStringOrdinal(StrPtr(firstString), Len(firstString), StrPtr(secondString), Len(secondString), 1&) = 2&)
            Else
                StringsEqual = (CompareStringW(pdli_SystemDefault, NORM_IGNORECASE, StrPtr(firstString), -1&, StrPtr(secondString), -1&) = 2&)
            End If
        Else
            StringsEqual = VBHacks.MemCmp(StrPtr(firstString), StrPtr(secondString), Len(firstString) * 2)
        End If
    End If
    
End Function

'Convenience not-wrapper to StringsEqual, above
Public Function StringsNotEqual(ByVal firstString As String, ByVal secondString As String, Optional ByVal ignoreCase As Boolean = False) As Boolean
    StringsNotEqual = Not StringsEqual(firstString, secondString, ignoreCase)
End Function

'When passing file and path strings among API calls, they often have to be pre-initialized to some arbitrary buffer length
' (typically MAX_PATH).  When finished, the string needs to be resized to remove any null chars.  Use this function to do so.
Public Function TrimNull(ByRef origString As String) As String

    'Find a null char, if any
    Dim nullPosition As Long
    nullPosition = InStr(origString, Chr$(0))
    
    If (nullPosition > 0) Then
       TrimNull = Left$(origString, nullPosition - 1)
    Else
       TrimNull = origString
    End If
  
End Function

