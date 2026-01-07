VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon String Builder class
'Copyright 2017-2026 by Tanner Helland
'Created: 06/October/17
'Last updated: 06/October/17
'Last update: initial build
'
'I was going to delay construction of a string builder until post-7.0, but I ran into some issues with the
' pdGradient class and serializing huge gradient node collections, so it made sense to write a little
' string builder sooner rather than later.
'
'VB6's problematic string performance is well-studied, and I'm sure there are more comprehensive solutions
' to this problem out there, but like most other classes in this project, it's nice to have a PD-specific
' solution that does exactly what we need (no more, no less).
'
'Instead of wrapping an array, note that we directly buffer an underlying BSTR.  This simplifies some
' operations that are otherwise tedious (e.g. using built-in VB string operations, checking buffer length,
' ensuring null-termination).  This also allows us to use internal VB string functions, where desired,
' although this class may still preferentially choose various string APIs for performance reasons
' (e.g. to avoid searching large stretches of null chars at the end of the buffer).
'
'Functions are (roughly) modeled after .NET's StringBuilder class:
' https://msdn.microsoft.com/en-us/library/system.text.stringbuilder(v=vs.110).aspx
' ...but with the usual modifications to imitate things like function overloading.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'*************************************************************************

Option Explicit

'shlwapi provides convenient C++ string wrappers, but note that they always rely on null-terminated strings,
' and they always return pointers (not 0-based offsets within the string).  We have to translate these
' manually to get VB-like results.
Private Declare Function StrStrW Lib "shlwapi" (ByVal pszFirst As Long, ByVal pszSrch As Long) As Long
Private Declare Function StrStrIW Lib "shlwapi" (ByVal pszFirst As Long, ByVal pszSrch As Long) As Long

'At present, we use a mathematically pleasing phi growth strategy.
' (https://stackoverflow.com/questions/1100311/what-is-the-ideal-growth-rate-for-a-dynamically-allocated-array)
Private Const GROWTH_STRATEGY As Double = ((1# + 2.23606797749979) * 0.5)

'The actual string data is stored here.  While this is obviously a standard VB string, note that it may be
' padded by a (potentially large) series of null chars - so Len() can be used to track *buffer size*, but not
' *string length*; for string length, use the m_Length variable.
Private m_Buffer As String
Private m_Length As Long

Private Sub Class_Initialize()

    'Technically you could forcibly initialize the buffer here, but because PD uses tiny strings in a
    ' number of places (e.g. individual string instances that are only appended to once), we don't actually
    ' initialize the buffer until absolutely necessary.

End Sub

Private Sub Class_Terminate()

    'Because we use a standard VB string as our buffer, no special unload commands are required.

End Sub

Friend Sub Append(ByRef srcString As String)
    
    'Make sure we've got room to append.
    ' NOTE: if performance is crucial, in-line the If statement from EnsureCapacity, like so:
    ' If (m_Length + Len(srcString)) > Len(m_Buffer) Then EnsureCapacity m_Length + Len(srcString)
    EnsureCapacity m_Length + Len(srcString)
    
    'Copy the string into place, then update our internal length tracker
    CopyMemoryStrict StrPtr(m_Buffer) + m_Length * 2, StrPtr(srcString), Len(srcString) * 2
    m_Length = m_Length + Len(srcString)
    
End Sub

'Other helper append functions; these all follow the same pattern, so refer to the (simpler) base Append
' function for details.
Friend Sub AppendFromPtr(ByVal srcPtr As Long, ByVal srcLenInChars As Long)
    EnsureCapacity m_Length + srcLenInChars
    CopyMemoryStrict StrPtr(m_Buffer) + m_Length * 2, srcPtr, srcLenInChars * 2
    m_Length = m_Length + srcLenInChars
End Sub

'Append a full line of text, and terminate with the standard windows carriage return + linefeed combo
Friend Sub AppendLine(ByRef srcString As String)
    EnsureCapacity m_Length + Len(srcString) + 2
    CopyMemoryStrict StrPtr(m_Buffer) + m_Length * 2, StrPtr(srcString), Len(srcString) * 2
    m_Length = m_Length + Len(srcString) + 2
    
    'Manually write a Windows-style CrLf feed after the original string
    PutMem2 StrPtr(m_Buffer) + m_Length * 2 - 4, 13
    PutMem2 StrPtr(m_Buffer) + m_Length * 2 - 2, 10
End Sub

'Append a full line of text, but terminate using a Unix-style linefeed-only char
Friend Sub AppendLineLf(ByRef srcString As String)
    EnsureCapacity m_Length + Len(srcString) + 1
    CopyMemoryStrict StrPtr(m_Buffer) + m_Length * 2, StrPtr(srcString), Len(srcString) * 2
    m_Length = m_Length + Len(srcString) + 1
    
    'Manually write only a linefeed char after the original string
    PutMem2 StrPtr(m_Buffer) + m_Length * 2 - 2, 10
End Sub

'CrLf linebreak
Friend Sub AppendLineBreak()
    m_Length = m_Length + 2
    EnsureCapacity m_Length
    PutMem2 StrPtr(m_Buffer) + m_Length * 2 - 4, 13
    PutMem2 StrPtr(m_Buffer) + m_Length * 2 - 2, 10
End Sub

'Lf only
Friend Sub AppendLineBreakLf()
    m_Length = m_Length + 1
    EnsureCapacity m_Length
    PutMem2 StrPtr(m_Buffer) + m_Length * 2 - 2, 10
End Sub

Friend Sub AppendPDString(ByRef srcPDString As pdString)
    EnsureCapacity m_Length + srcPDString.GetLength
    CopyMemoryStrict StrPtr(m_Buffer) + m_Length * 2, srcPDString.GetStrPtr(), srcPDString.GetLength() * 2
    m_Length = m_Length + srcPDString.GetLength()
End Sub

'Ensure the buffer as a whole is at least (n) chars long.
Friend Sub EnsureCapacity(ByVal requiredCapacityInChars As Long)
    
    If (Len(m_Buffer) < requiredCapacityInChars) Then
    
        'Time to enlarge the buffer!  Figure out what our new size would be using the current growth strategy.
        ' (See the top of the class for details on how this multiplier was selected.)
        Dim newSize As Long
        newSize = Int(Len(m_Buffer) * GROWTH_STRATEGY + 0.5)
        
        'If the caller wants more size than the default allocator would give us, just use their request.
        If (requiredCapacityInChars > newSize) Then newSize = requiredCapacityInChars
        
        'Create a new string, with null-padding filling the end of the newly allocated buffer
        m_Buffer = m_Buffer & String$(newSize - Len(m_Buffer), 0)
        
    End If
    
End Sub

'Find a specific character (single-byte only).  Use StrStr for Unicode-aware or case-insensitive comparisons.
' Returns 0 if char is not found.
Friend Function FindChar(ByRef srcChar As String, Optional ByVal startPos As Long = 1) As Long
    FindChar = InStr(startPos, m_Buffer, srcChar, vbBinaryCompare)
    If (FindChar > m_Length) Then FindChar = 0
End Function

'1-based current buffer length
Friend Function GetCapacity() As Long
    GetCapacity = Len(m_Buffer)
End Function

'Length of useful string data, *not* necessarily the size of the buffer.  (By extension, it will always be
' equal-to or shorter than the current buffer size.)
Friend Function GetLength() As Long
    GetLength = m_Length
End Function

'Places an entire line of text inside GetLine.  Delimiter can be manually specified.  Returns FALSE
' if no line is found (treat as EOF).
'Note that the delimiter itself *IS NOT* returned as part of the string, by design; to account for
' the next search position (while doing a line-by-line retrieval, for example), you can calculate
' the start of the next line as "startPos + Len(returnedLine) + Len(lineDelimiter) + 1"
Friend Function GetLine(ByRef dstLine As String, Optional ByVal startPos As Long = 1, Optional ByVal lineDelimiter As String = vbCrLf) As Boolean
    If (startPos < m_Length) Then
        Dim delPos As Long
        delPos = InStr(startPos, m_Buffer, lineDelimiter, vbBinaryCompare)
        If (delPos > m_Length) Then delPos = 0  'String may be null-padded beyond its BSTR length
        GetLine = (delPos > 0)
        If GetLine Then dstLine = Mid$(m_Buffer, startPos, delPos - startPos) Else dstLine = vbNullString
    End If
End Function

'Use this value responsibly!  (Also, note that this function may return 0; our internal buffer is not guaranteed
' to be initialized, especially if none of the "Append" functions have ever been called.)
Friend Function GetStrPtr() As Long
    GetStrPtr = StrPtr(m_Buffer)
End Function

'We deliberately wrap certain VB functions, as they'll inadvertently copy null bytes at the end of our buffer.
Friend Function MidW(ByVal startPos As Long, Optional ByVal midLength As Long = 0) As String
    If (midLength = 0) Then midLength = (m_Length - startPos) + 1
    MidW = Mid$(m_Buffer, startPos, midLength)
End Function

'Want to erase the entire buffer with some new string?  This function provides a shortcut way to do it
' (vs clearing the buffer and manually appending a string).
Friend Sub Reset(Optional ByRef newString As String = vbNullString, Optional ByVal ensureTrailingNulls As Boolean = True)
    
    'Reuse our existing allocation if we can (taking care to zero-out trailing bytes, as a precaution)
    m_Length = Len(newString)
    If (LenB(m_Buffer) > 0) Then VBHacks.ZeroMemory StrPtr(m_Buffer), Len(m_Buffer) * 2
    EnsureCapacity m_Length
    
    'Copy the new string directly into place
    If (LenB(newString) > 0) Then CopyMemoryStrict StrPtr(m_Buffer), StrPtr(newString), LenB(newString)
    
End Sub

'Modify the current "end of data" pointer.  *ZERO VALIDATION IS PERFORMED ON THE PASSED LENGTH PARAMETER*,
' by design - so use this function at your own risk.
Friend Sub SetLength(ByVal newLength As Long)
    m_Length = newLength
End Sub

'InStr replacement.  To avoid wasting time searching null-padding at the end of the string buffer,
' we lean on C++ string operators provided by shlwapi.  (They automatically stop searching when a
' null-char is reached; if this behavior is *not* desired, look below for an InStr-based implementation.)
Friend Function StrStr(ByRef strToLookFor As String, Optional ByVal searchStartPos As Long = 1) As Long
    
    If ((searchStartPos - 1) <= Len(m_Buffer)) Then
        
        'Use PD's faster Boyer-Moore engine when comparing longer strings
        If (LenB(strToLookFor) > 8) Then
            StrStr = Strings.StrStrBM(m_Buffer, strToLookFor, searchStartPos, True, m_Length)
        Else
            StrStr = StrStrW(StrPtr(m_Buffer) + (searchStartPos - 1) * 2, StrPtr(strToLookFor))
            If (StrStr <> 0) Then StrStr = (StrStr - StrPtr(m_Buffer)) \ 2 + 1
        End If
        
    End If
    
    'FYI, an equivalent VB-only implementation would be as simple as:
    'StrStr = InStr(searchStartPos, m_Buffer, strToLookFor, vbBinaryCompare)
    
End Function

'Case-insensitive InStr replacement.
Friend Function StrStrI(ByRef strToLookFor As String, Optional ByVal searchStartPos As Long = 1) As Long
    If ((searchStartPos - 1) <= Len(m_Buffer)) Then
        StrStrI = StrStrIW(StrPtr(m_Buffer) + (searchStartPos - 1) * 2, StrPtr(strToLookFor))
        If (StrStrI <> 0) Then StrStrI = (StrStrI - StrPtr(m_Buffer)) \ 2 + 1
    End If
End Function

'Return a copy of our internal buffer, auto-trimmed to fit its current length.
Friend Function ToString() As String
    If (LenB(m_Buffer) > 0) And (m_Length > 0) Then ToString = Left$(m_Buffer, m_Length) Else ToString = vbNullString
End Function

'Because PD uses a ton of XML data, this class exposes some XML-specific operations.  These use a number of
' internal shortcuts specific to XML, and they can easily be chopped out if you don't need 'em.  Just delete
' all functions and subs starting with "XML" (which I've tried to keep in a block).
Friend Sub XMLAppend(ByRef appendTagName As String, ByRef appendTagValue As String)
    
    'Failsafe check for valid tag names
    If (LenB(appendTagName) = 0) Then Exit Sub
    
    'We need 5 extra chars: "<></>" in addition to the length of the strings we're passed.  (Note also
    ' that we need 2x the length of the tag name - 1x for the opening tag, 1x for the closing tag.)
    Dim lenRequired As Long
    lenRequired = Len(appendTagName) * 2 + Len(appendTagValue) + 5
    EnsureCapacity m_Length + lenRequired
    
    'Manually write all tags in order.  This is kinda tedious, but it's faster than calling "Append"
    ' for each char, as that would force multiple validations of the buffer size.
    Dim targetPtr As Long
    targetPtr = StrPtr(m_Buffer) + m_Length * 2
    
    'Opening tag (AscW("<") = 60) (AscW(">") = 62)
    PutMem2 targetPtr, 60
    targetPtr = targetPtr + 2
    
    CopyMemoryStrict targetPtr, StrPtr(appendTagName), Len(appendTagName) * 2
    targetPtr = targetPtr + Len(appendTagName) * 2
    
    PutMem2 targetPtr, 62
    targetPtr = targetPtr + 2
    
    'Tag value
    If (LenB(appendTagValue) > 0) Then
        CopyMemoryStrict targetPtr, StrPtr(appendTagValue), Len(appendTagValue) * 2
        targetPtr = targetPtr + Len(appendTagValue) * 2
    End If
    
    'Closing tag (AscW("<") = 60) (AscW(">") = 62) (AscW("/") = 47)
    PutMem2 targetPtr, 60
    targetPtr = targetPtr + 2
    
    PutMem2 targetPtr, 47
    targetPtr = targetPtr + 2
    
    CopyMemoryStrict targetPtr, StrPtr(appendTagName), Len(appendTagName) * 2
    targetPtr = targetPtr + Len(appendTagName) * 2
    
    PutMem2 targetPtr, 62
    targetPtr = targetPtr + 2
    
    'Perform a final update of the current length
    m_Length = m_Length + lenRequired

End Sub

'Update an existing value (e.g. given a tag name, replace the text inside that tag with new text).
' Returns: TRUE if the update is successful; FALSE if the update fails.  An update will only fail if the tag name
'          cannot be found inside the current XML string.
Friend Function XMLUpdateTagValue(ByRef tagName As String, ByRef newTagValue As String, Optional ByVal searchStartPos As Long = 1) As Boolean
    
    XMLUpdateTagValue = False
    
    'Find the location of the specified tag.
    Dim tagPosStart As Long
    tagPosStart = Me.StrStr("<" & tagName & ">", searchStartPos)
    
    'Only proceed if the tag was found
    If (tagPosStart > 0) Then
        
        'Advance the tag position to where the tag value lies
        tagPosStart = tagPosStart + Len(tagName) + 2
        
        'Find where the *closing* tag starts
        Dim tagPosEnd As Long
        tagPosEnd = Me.StrStr("</" & tagName & ">", tagPosStart)
        
        If (tagPosEnd > 0) Then
            
            XMLUpdateTagValue = True
            
            'If the new tag value is *longer* than the previous one, we may need to extend the buffer first.
            ' (Similarly, if the new tag and old tag are identical in length, we don't have to adjust the
            '  buffer at all, making this op very fast.)
            Dim curTagLength As Long
            curTagLength = (tagPosEnd - tagPosStart)
            
            'If the new tag value is identical in length to the previous tag value, we don't have to shuffle
            ' anything around.  Check for this rare but awesome case.
            If (curTagLength = Len(newTagValue)) Then
            
                'Perform an in-place write
                CopyMemoryStrict StrPtr(m_Buffer) + tagPosStart * 2 - 2, StrPtr(newTagValue), Len(newTagValue) * 2
            
            'The lengths differ, meaning we need a temporary buffer.
            Else
            
                'If the new tag is longer, ensure we have enough room to fit it
                If (curTagLength < Len(newTagValue)) Then Me.EnsureCapacity m_Length + (Len(newTagValue) - curTagLength) + 1
                
                'Before proceeding, copy the trailing text into a temporary buffer.
                Dim tmpBuffer As String
                tmpBuffer = Mid$(m_Buffer, tagPosEnd, m_Length - tagPosEnd + 1)
                
                'Add the new tag value
                CopyMemoryStrict StrPtr(m_Buffer) + tagPosStart * 2 - 2, StrPtr(newTagValue), Len(newTagValue) * 2
                
                'Copy the trailing text into place *after* the buffer
                tagPosStart = tagPosStart * 2 - 2 + Len(newTagValue) * 2
                CopyMemoryStrict StrPtr(m_Buffer) + tagPosStart, StrPtr(tmpBuffer), Len(tmpBuffer) * 2
                
                'Update our length accordingly
                m_Length = m_Length + Len(newTagValue) - curTagLength
                
                'If the new length is shorter than the old length, null-pad the trailing bytes accordingly
                If (Len(newTagValue) < curTagLength) Then FillMemory StrPtr(m_Buffer) + m_Length * 2, (curTagLength - Len(newTagValue)) * 2, 0
                
            End If
            
        End If
    
    End If
        
End Function
