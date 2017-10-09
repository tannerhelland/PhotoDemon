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

Private Const CP_UTF8 As Long = 65001   'Fixed constant for UTF-8 "codepage" transformations
Private Const CRYPT_STRING_BASE64 As Long = 1&
Private Const CRYPT_STRING_NOCR As Long = &H80000000
Private Const CRYPT_STRING_NOCRLF As Long = &H40000000
Private Const LOCALE_SYSTEM_DEFAULT As Long = &H800&

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

'(Both LCMapString variants use the same constants)
Private Enum REMAP_STRING_API
    LCMAP_LOWERCASE = &H100&
    LCMAP_UPPERCASE = &H200&
    LCMAP_TITLECASE = &H300&      'Windows 7 only!

    LCMAP_HIRAGANA = &H100000
    LCMAP_KATAKANA = &H200000

    LCMAP_LINGUISTIC_CASING = &H1000000     'Per MSDN, "Use linguistic rules for casing, instead of file system rules (default)."
                                            '           This flag is valid with LCMAP_LOWERCASE or LCMAP_UPPERCASE only."

    LCMAP_SIMPLIFIED_CHINESE = &H2000000
    LCMAP_TRADITIONAL_CHINESE = &H4000000
End Enum

Private Declare Function CryptBinaryToString Lib "crypt32" Alias "CryptBinaryToStringW" (ByVal ptrBinaryData As Long, ByVal numBytesToConvert As Long, ByVal dwFlags As Long, ByVal ptrToDstString As Long, ByRef sizeOfStringBuffer As Long) As Long
Private Declare Function CryptStringToBinary Lib "crypt32" Alias "CryptStringToBinaryW" (ByVal pszString As Long, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, ByRef pcbBinary As Long, ByRef pdwSkip As Long, ByRef pdwFlags As Long) As Long

Private Declare Sub CopyMemoryStrict Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDst As Long, ByVal lpSrc As Long, ByVal byteLength As Long)
Private Declare Function CompareStringW Lib "kernel32" (ByVal lcID As PD_LocaleIdentifier, ByVal cmpFlags As StrCmpFlags, ByVal ptrToStr1 As Long, ByVal str1Len As Long, ByVal ptrToStr2 As Long, ByVal str2Len As Long) As Long
Private Declare Function CompareStringOrdinal Lib "kernel32" (ByVal ptrToStr1 As Long, ByVal str1Len As Long, ByVal ptrToStr2 As Long, ByVal str2Len As Long, ByVal bIgnoreCase As Long) As Long
Private Declare Function LCMapStringW Lib "kernel32" (ByVal localeID As Long, ByVal dwMapFlags As REMAP_STRING_API, ByVal lpSrcStringPtr As Long, ByVal lenSrcString As Long, ByVal lpDstStringPtr As Long, ByVal lenDstString As Long) As Long
Private Declare Function LCMapStringEx Lib "kernel32" (ByVal lpLocaleNameStringPt As Long, ByVal dwMapFlags As REMAP_STRING_API, ByVal lpSrcStringPtr As Long, ByVal lenSrcString As Long, ByVal lpDstStringPtr As Long, ByVal lenDstString As Long, ByVal lpVersionInformationPtr As Long, ByVal lpReserved As Long, ByVal sortHandle As Long) As Long 'Vista+ only!  (Note the lack of a trailing W in the function name.)
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal dstCodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function SysAllocStringByteLen Lib "oleaut32" (ByVal srcPtr As Long, ByVal strLength As Long) As String
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal dstCodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long

'Apply basic heuristics to the first (n) bytes of a potentially UTF-8 source, and return a "best-guess" on whether the bytes
' represent valid UTF-8 data.
'
'This is based off a similar function by Dana Seaman, who noted an original source of http://www.geocities.co.jp/SilkRoad/4511/vb/utf8.htm
' I have modified the function to ignore invalid 5- and 6- byte extensions, and to shorten the validation range as the original 2048 count
' seems excessive.  (For a 24-byte sequence, the risk of a false positive is less than 1 in 1,000,000;
' see http://stackoverflow.com/questions/4520184/how-to-detect-the-character-encoding-of-a-text-file?lq=1.  False negative results have
' a higher probability, but ~100 characters should be enough to determine this, especially given the typical use-cases in PD.)
'
'For additional details on UTF-8 heuristics, see:
'  https://github.com/neitanod/forceutf8/blob/master/src/ForceUTF8/Encoding.php
'  http://www-archive.mozilla.org/projects/intl/UniversalCharsetDetection.html (very detailed)
Private Function AreBytesUTF8(ByRef textBytes() As Byte, Optional ByVal verifyLength As Long = 128) As Boolean
    
    AreBytesUTF8 = False
    
    If (verifyLength > 0) Then
    
        Dim utf8Size As Long, lIsUtf8 As Long, i As Long
        
        'If the requested verification length exceeds the size of the array, just search the entire array
        If (verifyLength > UBound(textBytes)) Then verifyLength = UBound(textBytes)
        
        'Scan through the byte array, looking for patterns specific to UTF-8
        Dim pos As Long: pos = 0
        Do While (pos < verifyLength)
        
            'If this is a standard ANSI value, it doesn't tell us anything useful - advance to the next byte
            If (textBytes(pos) <= &H7F) Then
                pos = pos + 1
            
            'If this value is a continuation byte (128-191), invalid byte (192-193), or Latin-1 identifier (194), we know
            ' the text is *not* UTF-8.  Exit now.
            ElseIf (textBytes(pos) < &HC0) Then
                AreBytesUTF8 = False
                Exit Function
            
            'Other byte values are potential multibyte UTF-8 markers.  We will advance the pointer by a matching amount, and scan
            ' intermediary bytes to make sure they do not contain invalid markers.
            ElseIf (textBytes(pos) <= &HF4) Then
                
                'These special-range UTF-8 markers are used to represent multi-byte encodings.  Detect how many bytes are included
                ' in this character
                If ((textBytes(pos) And &HC0) = &HC0) Then
                    utf8Size = 1
                ElseIf ((textBytes(pos) And &HE0) = &HE0) Then
                    utf8Size = 2
                ElseIf ((textBytes(pos) And &HF0) = &HF0) Then
                    utf8Size = 3
                End If
                
                'If the position exceeds the length we are supposed to verify, exit now and rely on previous detection
                ' passes to return a yes/no result.
                If ((pos + utf8Size) >= verifyLength) Then Exit Do
                
                'Scan the intermediary bytes of this character to ensure that no invalid markers are contained.
                For i = (pos + 1) To (pos + utf8Size)
                    
                    'Valid UTF-8 continuation bytes must not exceed &H80
                    If ((textBytes(i) And &HC0) <> &H80) Then
                        
                        'This is an invalid marker; exit immediately
                        AreBytesUTF8 = False
                        Exit Function
                        
                    End If
                    
                Next i
                
                'If we made it all the way here, all bytes in this multibyte set are valid.  Note that we've found at least one
                ' valid UTF-8 multibyte encoding, and carry on with the next character
                lIsUtf8 = lIsUtf8 + 1
                pos = pos + utf8Size + 1
            
            'Byte values above 0xF4 are always invalid (http://en.wikipedia.org/wiki/UTF-8).  Exit immediately and report failure.
            Else
                AreBytesUTF8 = False
                Exit Function
            End If
            
        Loop
        
        'If we found at least one valid, multibyte UTF-8 sequence, return TRUE.  If we did not encounter such a sequence, then all
        ' characters fall within the ASCII range.  This is "indeterminate", and returning TRUE or FALSE is really a matter of preference.
        ' Default to whatever return you think is most likely.  (In PD's case, we assume UTF-8, as files are likely coming from internal
        ' files.)
        If (lIsUtf8 > 0) Then
            AreBytesUTF8 = True
        
        'Indeterminate case
        Else
            AreBytesUTF8 = True
        End If
        
    'If no validation length is passed, any heuristics are pointless - exit immediately.
    End If
    
End Function

'Convert a byte array into a base-64 encoded string, using standard Windows libraries.
' Returns TRUE if successful; FALSE otherwise.
Public Function BytesToBase64(ByRef srcArray() As Byte, ByRef dstBase64 As String) As Boolean
    
    BytesToBase64 = False
    
    'Retrieve the necessary output buffer size.
    Dim bufferSize As Long
    If (CryptBinaryToString(VarPtr(srcArray(LBound(srcArray))), UBound(srcArray) - LBound(srcArray) + 1, CRYPT_STRING_BASE64 Or CRYPT_STRING_NOCRLF, 0&, bufferSize) <> 0) Then
        dstBase64 = String$(bufferSize - 1, 0)
        BytesToBase64 = (CryptBinaryToString(VarPtr(srcArray(LBound(srcArray))), UBound(srcArray) - LBound(srcArray) + 1, CRYPT_STRING_BASE64 Or CRYPT_STRING_NOCRLF, StrPtr(dstBase64), bufferSize) <> 0)
    End If
    
End Function

Public Function BytesToBase64Ex(ByVal ptrToSrcData As Long, ByVal lenOfSrcDataInBytes As Long, ByRef dstBase64 As String) As Boolean
    
    BytesToBase64Ex = False
    
    'Retrieve the necessary output buffer size.
    Dim bufferSize As Long
    If (CryptBinaryToString(ptrToSrcData, lenOfSrcDataInBytes, CRYPT_STRING_BASE64 Or CRYPT_STRING_NOCRLF, 0&, bufferSize) <> 0) Then
        dstBase64 = String$(bufferSize - 1, 0)
        BytesToBase64Ex = (CryptBinaryToString(ptrToSrcData, lenOfSrcDataInBytes, CRYPT_STRING_BASE64 Or CRYPT_STRING_NOCRLF, StrPtr(dstBase64), bufferSize) <> 0)
    End If
    
End Function

'Convert a base-64 encoded string into a byte array, using standard Windows libraries.
' Returns TRUE if successful; FALSE otherwise.
'
'Thanks to vbForums user dilettante for the original version of this code (retrieved here: http://www.vbforums.com/showthread.php?514815-JPEG-Base-64&p=3186994&viewfull=1#post3186994)
Public Function BytesFromBase64(ByRef dstArray() As Byte, ByRef srcBase64 As String) As Boolean
    
    BytesFromBase64 = False
    
    'Retrieve the necessary output buffer size.
    Dim lngOutLen As Long, dwActualUsed As Long
    If (CryptStringToBinary(StrPtr(srcBase64), Len(srcBase64), CRYPT_STRING_BASE64, ByVal 0&, lngOutLen, 0&, dwActualUsed) <> 0) Then
        ReDim dstArray(lngOutLen - 1) As Byte
        BytesFromBase64 = (CryptStringToBinary(StrPtr(srcBase64), Len(srcBase64), CRYPT_STRING_BASE64, VarPtr(dstArray(0)), lngOutLen, 0&, dwActualUsed) <> 0)
    End If
    
End Function

'WARNING!  This function allows you to directly decode a Base64 string to an arbitrary destination pointer.
' Because this function cannot resize the destination memory space (obviously), you *must* use some external
' knowledge to size your destination buffer appropriately.  This function will return the number of bytes used,
' but crashes and/or security compromises are possible if you do not validate your destination size correctly.
Public Function BytesFromBase64Ex(ByVal dstPtr As Long, ByRef dstBufferSize As Long, ByRef srcBase64 As String) As Boolean
    BytesFromBase64Ex = False
    Dim dwActualUsed As Long
    BytesFromBase64Ex = (CryptStringToBinary(StrPtr(srcBase64), Len(srcBase64), CRYPT_STRING_BASE64, dstPtr, dstBufferSize, 0&, dwActualUsed) <> 0)
End Function

'Given an arbitrary pointer to a null-terminated CHAR or WCHAR run, measure the resulting string and copy the results
' into a VB string.
'
'For security reasons, if an upper limit of the string's length is known in advance (e.g. MAX_PATH), pass that limit
' via the optional maxLength parameter to avoid a buffer overrun.  This function has a hard-coded limit of 65k chars,
' a limit you can easily lift but which makes sense for PD.  If a string exceeds the limit (whether passed or
' hard-coded), *a string will still be created and returned*, but it will be clamped to the max length.
'
'If the string length is known in advance, and WCHARS are being used, please use the faster (and more secure)
' StringFromUTF16_FixedLen() function, below.
Public Function StringFromCharPtr(ByVal srcPointer As Long, Optional ByVal srcStringIsUnicode As Boolean = True, Optional ByVal maxLength As Long = -1) As String
    
    'Check string length
    Dim strLength As Long
    If srcStringIsUnicode Then strLength = lstrlenW(srcPointer) Else strLength = lstrlenA(srcPointer)
    
    'Make sure the length/pointer isn't null
    If (strLength <= 0) Then
        StringFromCharPtr = vbNullString
    Else
        
        'Make sure the string's length is valid.
        Dim maxAllowedLength As Long
        If (maxLength = -1) Then maxAllowedLength = 65535 Else maxAllowedLength = maxLength
        If (strLength > maxAllowedLength) Then strLength = maxAllowedLength
        
        'Create the target string and copy the bytes over
        If srcStringIsUnicode Then
            StringFromCharPtr = String$(strLength, 0)
            CopyMemoryStrict StrPtr(StringFromCharPtr), srcPointer, strLength * 2
        Else
            StringFromCharPtr = SysAllocStringByteLen(srcPointer, strLength)
        End If
    
    End If
    
End Function

'Given an array of arbitrary bytes, perform a series of heuristics to perform a "best-guess" conversion to VB's internal DBCS string format.
'
'Currently supported formats include big- and little-endian UTF-16, UTF-8, DBCS, and ANSI variants.  Note that ANSI variants are *always*
' converted using the current codepage, as codepage heuristics are complicated and unwieldy.
'
'For best results, pass text directly from a file into this function, as BOMs are very helpful when determining format.
'
'This function can optionally normalize line endings, but note that this is time-consuming.
'
'Finally, if you know the incoming string format in advance, it will be faster to perform your own format-specific conversion,
' because heuristics (particularly UTF-8 without BOM) can be time-consuming.
'
'RETURNS: TRUE if successful; FALSE otherwise.  Note that TRUE may not guarantee a correct string, especially if the incoming data
' is garbage, or if the "true" text format is an esoteric ANSI codepage.
Public Function StringFromMysteryBytes(ByRef srcBytes() As Byte, ByRef dstString As String, Optional ByVal forceWindowsLineEndings As Boolean = True) As Boolean
    
    On Error GoTo StringConversionFailed
    
    'There are a number of different ways to convert an arbitrary byte array to a string; this temporary string will be used to
    ' translate data between byte arrays and VB strings as necessary.
    Dim tmpString As String
    
    'Start running some string encoding heuristics.  BOMs are checked first, as they're easiest to handle.  Note that no attempts
    ' are currently made to detect UTF-32, due to its extreme rarity.  (That said, heursitics for it are simple;
    ' see http://stackoverflow.com/questions/4520184/how-to-detect-the-character-encoding-of-a-text-file/4522251#4522251)
    
    'First, check for UTF-8 BOM (0xEFBBBF).  This isn't common in the wild (as UTF-8 doesn't require a BOM), but we forcibly write
    ' it on all internal PD files because it lets us skip heuristics.
    If (srcBytes(0) = &HEF) And (srcBytes(1) = &HBB) And (srcBytes(2) = &HBF) Then
    
        'A helper function converts the UTF-8 bytes for us; all we need to do is remove the BOM
        dstString = Mid$(Strings.StringFromUTF8(srcBytes), 2)
        
    'Next, check for BOM 0xFFFE, which indicates little-endian UTF-16 (e.g. VB's internal format)
    ElseIf (srcBytes(0) = 255) And (srcBytes(1) = 254) Then
        
        'Cast the byte array straight into a string, then remove the BOM.
        tmpString = srcBytes
        dstString = Right$(tmpString, Len(tmpString) - 2)
        
    'Next, check for big-endian UTF-16 (0xFEFF)
    ElseIf (srcBytes(0) = 254) And (srcBytes(1) = 255) Then
      
        'Swizzle the incoming array
        Dim tmpSwap As Byte, i As Long
        
        For i = 0 To UBound(srcBytes) Step 2
            tmpSwap = srcBytes(i)
            srcBytes(i) = srcBytes(i + 1)
            srcBytes(i + 1) = tmpSwap
        Next i
        
        'Cast the newly ordered byte array straight into a string, then remove the BOM
        tmpString = srcBytes
        dstString = Right$(tmpString, Len(tmpString) - 2)
        
    'All BOM checks failed.  Time to run heuristics, and try to "guess" at the correct format.
    Else
        
        'Check for UTF-8 data without a BOM.  These heuristics are near-perfect for avoiding false-positives, but there will always
        ' be a (very) low risk of false-negatives.  The default character search length can be extended to reduce false-negative risk.
        If Strings.AreBytesUTF8(srcBytes) Then
            dstString = Strings.StringFromUTF8(srcBytes)
            
        'If the bytes do not appear to be UTF-8, we could theoretically run one final ANSI check.  US-ANSI data falls into the
        ' [0, 127] range, exclusively, so it's easy to identify.  If, however, the file contains bytes outside this range,
        ' we're SOL, because extended bytes vary according to the original creation locale (which we do not know).  In that case,
        ' we can't really do anything but use the current user locale and hope for the best, so rather than differentiate between
        ' these cases, I just do a forcible conversion using the current codepage.
        Else
            dstString = StrConv(srcBytes, vbUnicode)
            Debug.Print "FYI: Strings.StringFromMysteryBytes received a string with unclear encoding.  Current user's codepage was assumed."
        End If
        
    End If
    
    'If the caller is concerned about inconsistent line-endings, we can forcibly convert everything to vbCrLf.
    ' This harms performance (as we need to cover both the CR-only case (OSX) and LF-only case (Linux/Unix)),
    ' but it ensures that any combination of linefeed characters are properly normalized against vbCrLf.
    If forceWindowsLineEndings Then
        
        'To ensure "perfect" line-ending fixes, we would need to scan the file completely and search for orphaned vbLf or vbCr
        ' occurrences (e.g. this would catch standalone lines with variable endings).  However, the likelihood of this occurring
        ' on a PD-specific file is basically 0%, so to improve performance, we simply check for files where there are no vbCrLf
        ' pairs, but there *are* standalone vbLf and/or vbCr chars.
        
        'First, see if the file consists of something other than pure vbCrLf
        Dim needToNormalize As Boolean
        needToNormalize = (InStr(1, dstString, vbCrLf, vbBinaryCompare) = 0) And ((InStr(1, dstString, vbCr, vbBinaryCompare) <> 0) Or (InStr(1, dstString, vbLf, vbBinaryCompare) <> 0))
        
        If needToNormalize Then
            
            'Force all existing vbCrLf instances to vbLf
            If (InStr(1, dstString, vbCrLf, vbBinaryCompare) <> 0) Then dstString = Replace$(dstString, vbCrLf, vbLf, , , vbBinaryCompare)
            
            'Force all existing vbCr instances to vbLf
            If (InStr(1, dstString, vbCr, vbBinaryCompare) <> 0) Then dstString = Replace$(dstString, vbCr, vbLf, , , vbBinaryCompare)
            
            'With everything normalized against vbLf, convert all vbLf instances to vbCrLf
            If (InStr(1, dstString, vbLf, vbBinaryCompare) <> 0) Then dstString = Replace$(dstString, vbLf, vbCrLf, , , vbBinaryCompare)
            
        End If
    
    End If
    
    StringFromMysteryBytes = True
    
    Exit Function
    
StringConversionFailed:

    InternalError "Strings.StringFromMysteryBytes() failed; string conversion abandoned.", Err.Number
    StringFromMysteryBytes = False

End Function

'Given a byte array containing UTF-8 data, return the data as a VB string.  A custom length can also be specified;
' if it's missing, the full input array will be used.
Public Function StringFromUTF8(ByRef Utf8() As Byte, Optional ByVal customDataLength As Long = -1) As String
    
    'Use MultiByteToWideChar() to calculate the required size of the final string (e.g. UTF-8 expanded to VB's default wide character set).
    Dim lenWideString As Long
    If (customDataLength < 0) Then customDataLength = UBound(Utf8) + 1
    lenWideString = MultiByteToWideChar(CP_UTF8, 0, VarPtr(Utf8(0)), customDataLength, 0, 0)
    
    'If the returned length is 0, MultiByteToWideChar failed.  This typically only happens if totally invalid characters are found.
    If (lenWideString = 0) Then
        InternalError "Strings.StringFromUTF8() failed because MultiByteToWideChar did not return a valid buffer length (#" & Err.LastDllError & ")."
        StringFromUTF8 = vbNullString
        
    'The returned length is non-zero.  Prep a buffer, then retrieve the bytes.
    Else
    
        'Prep a temporary string buffer
        StringFromUTF8 = String$(lenWideString, 0)
        
        'Use the API to perform the actual conversion
        lenWideString = MultiByteToWideChar(CP_UTF8, 0, VarPtr(Utf8(0)), customDataLength, StrPtr(StringFromUTF8), lenWideString)
        
        'Make sure the conversion was successful.  (There is generally no reason for it to succeed when calculating a buffer length, only to
        ' fail here, but better safe than sorry.)
        If (lenWideString = 0) Then
            InternalError "Strings.StringFromUTF8() failed because MultiByteToWideChar could not perform the conversion, despite returning a valid buffer length (#" & Err.LastDllError & ")."
            StringFromUTF8 = vbNullString
        End If
        
    End If
    
End Function

'Given an arbitrary pointer (often to a VB array, but it doesn't matter) and a length IN BYTES, copy that chunk
' of bytes to a VB string.  The bytes must already be in Unicode format (UCS-2 or UTF-16).
Public Function StringFromUTF16_FixedLen(ByVal srcPointer As Long, ByVal lengthInBytes As Long, Optional ByVal trimNullChars As Boolean = True) As String
    StringFromUTF16_FixedLen = String$(lengthInBytes \ 2, 0)
    CopyMemoryStrict StrPtr(StringFromUTF16_FixedLen), srcPointer, lengthInBytes
    If trimNullChars Then StringFromUTF16_FixedLen = Strings.TrimNull(StringFromUTF16_FixedLen)
End Function

'Apply some kind of remap conversion ("change case" in Latin languages) using WAPI.
' IMPORTANT: some LCMAP constants *are only available under Windows 7*, so be aware of which requests fail on earlier OSes.
Public Function StringRemap(ByRef srcString As String, ByVal remapType As PD_STRING_REMAP) As String
    
    'If the remap type is 0, do nothing
    If (remapType = PDSR_NONE) Then
        StringRemap = srcString
    Else
    
        'Convert the incoming remap type to an API equivalent
        Dim apiFlags As REMAP_STRING_API
        
        Select Case remapType
        
            Case PDSR_LOWERCASE
                apiFlags = LCMAP_LINGUISTIC_CASING Or LCMAP_LOWERCASE
            
            Case PDSR_UPPERCASE
                apiFlags = LCMAP_LINGUISTIC_CASING Or LCMAP_UPPERCASE
                
            Case PDSR_HIRAGANA
                apiFlags = LCMAP_HIRAGANA
                
            Case PDSR_KATAKANA
                apiFlags = LCMAP_KATAKANA
                
            Case PDSR_SIMPLE_CHINESE
                apiFlags = LCMAP_SIMPLIFIED_CHINESE
                
            Case PDSR_TRADITIONAL_CHINESE
                apiFlags = LCMAP_TRADITIONAL_CHINESE
                
            Case PDSR_TITLECASE_WIN7
                apiFlags = LCMAP_TITLECASE
                
                'If the remap type is "titlecase" and we're on Vista or earlier, do nothing
                If (Not OS.IsWin7OrLater) Then
                    StringRemap = srcString
                    Exit Function
                End If
        
        End Select
        
        'For Latin languages, the length of the new string shouldn't change, but with CJK languages, there are no guarantees.  As a failsafe,
        ' double the length of the temporary destination buffer.
        Dim dstString As String
        dstString = String$(Len(srcString) * 2, 0)
        
        'Use the Vista+ variant preferentially, as it has received additional updates versus the backward-compatible function.
        Dim apiSuccess As Boolean
        
        If OS.IsVistaOrLater Then
            apiSuccess = (LCMapStringEx(0&, apiFlags, StrPtr(srcString), Len(srcString), StrPtr(dstString), Len(dstString), 0&, 0&, 0&) <> 0)
            If (Not apiSuccess) Then InternalError "LCMapStringEx() failed on /" & srcString & "/ and PD remap type " & remapType & "."
        Else
            apiSuccess = (LCMapStringW(LOCALE_SYSTEM_DEFAULT, apiFlags, StrPtr(srcString), Len(srcString), StrPtr(dstString), Len(dstString)) <> 0)
            If (Not apiSuccess) Then InternalError "LCMapStringW() failed on /" & srcString & "/ and PD remap type " & remapType & "."
        End If
        
        'Because we use a huge destination buffer (as a failsafe), trailing null chars are inevitable.  Trim them before returning.
        If apiSuccess Then StringRemap = Strings.TrimNull(dstString) Else StringRemap = srcString
        
    End If
    
End Function

'High-performance string equality function.  Returns TRUE/FALSE for equality, with support for case-insensitivity.
Public Function StringsEqual(ByRef firstString As String, ByRef secondString As String, Optional ByVal ignoreCase As Boolean = False) As Boolean
    
    'Cheat and compare length first
    If (Len(firstString) <> Len(secondString)) Then
        StringsEqual = False
    Else
        If ignoreCase Then
            If OS.IsVistaOrLater Then
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
Public Function StringsNotEqual(ByRef firstString As String, ByRef secondString As String, Optional ByVal ignoreCase As Boolean = False) As Boolean
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

'Given a VB string, fill a byte array with matching UTF-8 data.  Returns TRUE if successful; FALSE otherwise
Public Function UTF8FromString(ByRef srcString As String, ByRef dstUtf8() As Byte, Optional ByRef lenUTF8 As Long) As Boolean
    UTF8FromString = Strings.UTF8FromStrPtr(StrPtr(srcString), Len(srcString), dstUtf8, lenUTF8)
End Function

'Given a pointer to a VB string, fill a byte array with matching UTF-8 data.  Returns TRUE if successful; FALSE otherwise
Public Function UTF8FromStrPtr(ByVal srcPtr As Long, ByVal srcLenInChars As Long, ByRef dstUtf8() As Byte, Optional ByRef lenUTF8 As Long, Optional ByVal baseArrIndexToWrite As Long = 0) As Boolean
    
    UTF8FromStrPtr = False
    
    'Use WideCharToMultiByte() to calculate the required size of the final UTF-8 array.
    lenUTF8 = WideCharToMultiByte(CP_UTF8, 0, srcPtr, srcLenInChars, 0, 0, 0, 0)
    
    'If the returned length is 0, WideCharToMultiByte failed.  This typically only happens if totally invalid character combinations are found.
    If (lenUTF8 = 0) Then
        InternalError "Strings.UTF8FromStrPtr() failed because WideCharToMultiByte did not return a valid buffer length (#" & Err.LastDllError & ")."
        
    'The returned length is non-zero.  Prep a buffer, then process the bytes.
    Else
        
        'Prep a temporary byte buffer.  In some places in PD, we'll reuse the same buffer for multiple string copies,
        ' so to improve performance, only resize the destination array as necessary.
        If VBHacks.IsArrayInitialized(dstUtf8) Then
            If ((UBound(dstUtf8) - LBound(dstUtf8) + 1 + baseArrIndexToWrite) < lenUTF8) Then ReDim dstUtf8(0 To lenUTF8 - 1 + baseArrIndexToWrite) As Byte
        Else
            ReDim dstUtf8(0 To lenUTF8 - 1 + baseArrIndexToWrite) As Byte
        End If
        
        'Use the API to perform the actual conversion
        lenUTF8 = WideCharToMultiByte(CP_UTF8, 0, srcPtr, srcLenInChars, VarPtr(dstUtf8(baseArrIndexToWrite)), lenUTF8, 0, 0)
        
        'Make sure the conversion was successful.  (There is generally no reason for it to succeed when calculating a buffer length, only to
        ' fail here, but better safe than sorry.)
        UTF8FromStrPtr = (lenUTF8 <> 0)
        If (Not UTF8FromStrPtr) Then InternalError "Strings.UTF8FromStrPtr() failed because WideCharToMultiByte could not perform the conversion, despite returning a valid buffer length (#" & Err.LastDllError & ")."
        
    End If
    
End Function

'Internal string-related errors are passed here.  PD writes these to a debug log, but only in debug builds; you can choose to
' handle errors differently.
Private Sub InternalError(ByVal errComment As String, Optional ByVal errNumber As Long = 0)
    #If DEBUGMODE = 1 Then
        If (errNumber <> 0) Then
            pdDebug.LogAction "WARNING!  VB error in Strings module (#" & Err.Number & "): " & Err.Description & " || " & errComment
        Else
            pdDebug.LogAction "WARNING!  Strings module internal error: " & errComment
        End If
    #End If
End Sub

