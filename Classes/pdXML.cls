VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon XML Interface (reading, writing, parsing, etc)
'Copyright ©2012-2013 by Tanner Helland
'Created: 30/July/13
'Last updated: 01/August/13
'Last update: initial build
'
'In 2013, PD became increasingly reliant on XML-format files.  The translation engine was the first to require XML
' interoperability (by design), followed a few months later by the metadata engine.  After the success of these
' two projects, a decision was made to switch all custom PhotoDemon filetypes to XML format.  This should provide
' excellent interoperability with 3rd-party projects, as well as provide a measure of future-proofing, since new
' features can be easily added without breaking old files (by simply adding new tags to file entries - tags that
' will simply be ignored by old copies of the software).
'
'Rather than write unique XML parsers for each custom filetype, this universal class was created.  It is meant to
' serve as a broad-purpose XML file interface, with strong support for reading, writing, and parsing XML files.
'
'The primary focus of this class is supporting the barebones XML features required for various PhotoDemon filetypes.
' IT IS NOT MEANT TO BE A FULL-FEATURED TO-SPEC XML PARSER, but it could certainly be extended to support additional
' XML features as needed.
'
'Many thanks to Frank Donckers, who helped prototype the original translation XML engine (which I referenced frequently
' while designing this class).
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'All PhotoDemon-compatible files must have their data wrapped in the following top-level tag
Private Const ROOT_TAG As String = "<pdData>"
Private Const ROOT_TAG_CLOSE As String = "</pdData>"

'The contents of the assigned XML file, stored in memory as one (potentially very long) string.
Private xmlContents As String

'If this class is being used to write out a new XML file, this function can be called to initialize the blank file.
Public Sub prepareBlankXML()

    xmlContents = "<?xml version=""1.0"" encoding=""windows-1252""?>" & vbCrLf & vbCrLf
    xmlContents = xmlContents & ROOT_TAG & vbCrLf & vbCrLf & ROOT_TAG_CLOSE & vbCrLf

End Sub

'Write a blank line into the XML file.  This has no practical purpose, but I'm a big fan of pretty XML output, so I like to use
' blank lines to separate tag families.
Public Function writeBlankLine() As Boolean

    'Find the </pdData> tag that signifies the end of PD-compatible XML data
    Dim tagLocation As Long
    tagLocation = InStr(1, xmlContents, ROOT_TAG_CLOSE, vbBinaryCompare)
    
    If tagLocation > 0 Then
    
        'Split the XML file into two halves: the half before the root tag, and the half after
        Dim topHalf As String, bottomHalf As String
        splitStringIn2 xmlContents, tagLocation - 1, topHalf, bottomHalf
        
        'Reassemble the primary string with a blank line inserted
        xmlContents = topHalf & vbCrLf & bottomHalf
        
        writeBlankLine = True
    Else
        writeBlankLine = False
    End If

End Function

'Write a comment into the XML file.  This has no practical purpose, but it can be helpful for end-users to understand the file's contents.
Public Function writeComment(ByVal commentText As String) As Boolean

    'Find the </pdData> tag that signifies the end of PD-compatible XML data
    Dim tagLocation As Long
    tagLocation = InStr(1, xmlContents, ROOT_TAG_CLOSE, vbBinaryCompare)
    
    If tagLocation > 0 Then
    
        'Split the XML file into two halves: the half before the root tag, and the half after
        Dim topHalf As String, bottomHalf As String
        splitStringIn2 xmlContents, tagLocation - 1, topHalf, bottomHalf
        
        'Reassemble the primary string with a blank line inserted
        xmlContents = topHalf & "<!-- " & commentText & " -->" & vbCrLf & bottomHalf
        
        writeComment = True
    Else
        writeComment = False
    End If

End Function

'Write a new XML tag to the master XML string.  By default, new tags are written to the end of the file, but the writeAtStart
' param can be set to TRUE to write tags at the top.
Public Function writeTag(ByVal tagName As String, ByVal tagValue As String, Optional ByVal writeAtStart As Boolean = False) As Boolean

    'Find the </pdData> tag that signifies the end of PD-compatible XML data
    Dim tagLocation As Long
    If writeAtStart Then
        tagLocation = InStr(1, xmlContents, ROOT_TAG, vbBinaryCompare)
    Else
        tagLocation = InStr(1, xmlContents, ROOT_TAG_CLOSE, vbBinaryCompare)
    End If

    If tagLocation > 0 Then
    
        'Split the XML file into two halves: the half before the root tag, and the half after
        Dim topHalf As String, bottomHalf As String
        If writeAtStart Then
            splitStringIn2 xmlContents, tagLocation + Len(ROOT_TAG), topHalf, bottomHalf
        Else
            splitStringIn2 xmlContents, tagLocation - 1, topHalf, bottomHalf
        End If
        
        'Build a string with the tag name and value we were passed
        Dim newTagLine As String
        newTagLine = "<" & tagName & ">" & tagValue & "</" & tagName & ">" & vbCrLf
        
        'Reassemble the primary string
        xmlContents = topHalf & newTagLine & bottomHalf
        
        writeTag = True
    Else
        writeTag = False
    End If

End Function

'Given a string and a position, split it into two strings at that position
Private Function splitStringIn2(ByRef srcString As String, ByVal splitPosition As Long, ByRef dstFirstHalf As String, ByRef dstSecondHalf As String)
    dstFirstHalf = Left(srcString, splitPosition)
    dstSecondHalf = Right(srcString, Len(srcString) - splitPosition)
End Function

'Once a valid XML file has been loaded, we need to see if it contains valid XML data for the current operation.  The client can
' do this by scanning for any number of tags it expects to find in the XML file.  If all are found, return TRUE.
Public Function validateLoadedXMLData(ParamArray expectedTags() As Variant) As Boolean

    'Start by looking for the <pdData> tags that surround all PhotoDemon-specific XML files
    If (InStr(1, xmlContents, ROOT_TAG, vbBinaryCompare) = 0) Or (InStr(1, xmlContents, ROOT_TAG_CLOSE, vbBinaryCompare) = 0) Then
        validateLoadedXMLData = False
        Exit Function
    End If

    'Search the xmlContents string for each tag in the validation request
    If Not IsMissing(expectedTags) Then
    
        Dim i As Long
        For i = LBound(expectedTags) To UBound(expectedTags)
            If InStr(1, xmlContents, expectedTags(i), vbTextCompare) = 0 Then
                validateLoadedXMLData = False
                Exit Function
            End If
        Next i
    
    End If
    
    validateLoadedXMLData = True

End Function

'Load an XML file into memory.  This function will also do some basic validation to ensure the requested file is actually XML.
' Returns: TRUE if file found, loaded, and validated successfully.  FALSE otherwise.
Public Function loadXMLFile(ByVal xmlPath As String) As Boolean
    
    If FileExist(xmlPath) Then
    
        'Load the file's contents into a string
        xmlContents = getFileAsString(xmlPath)
        
        'Check for an XML header
        If verifyXMLHeader(xmlContents) Then
            loadXMLFile = True
        Else
            loadXMLFile = False
        End If
    
    Else
        loadXMLFile = False
    End If
    
End Function

'Given an XML file (or sometimes, just the first 1024 bytes of an XML file), check to see if it has a valid XML header.
Private Function verifyXMLHeader(ByRef fileContents As String) As Boolean
        
    'Check for "<?xml" in the file.  We don't care about encoding, etc - just check "<?xml" to keep things quick.
    If InStr(1, fileContents, "<?xml", vbTextCompare) > 0 Then
        verifyXMLHeader = True
    Else
        verifyXMLHeader = False
    End If

End Function

'Retrieve an entire file and return it as a string.
Private Function getFileAsString(ByVal fName As String) As String
    
    Dim tmpFileContents As String
    
    'Ensure that the file exists before attempting to load it
    If FileExist(fName) Then
        
        Dim fileNum As Integer
        fileNum = FreeFile
    
        Open fName For Binary As #fileNum
            tmpFileContents = Space$(LOF(fileNum))
            Get #fileNum, , tmpFileContents
        Close #fileNum
    
        'Remove all tabs from the source file (which may have been added in by an XML editor, but are not relevant to our own XML processing)
        If InStr(1, tmpFileContents, vbTab) <> 0 Then tmpFileContents = Replace(tmpFileContents, vbTab, "")
        
        getFileAsString = tmpFileContents
    
    Else
        getFileAsString = ""
    End If
    
End Function

'Given an XML string, apply basic indentation
Private Sub applyIndentation(ByRef srcString As String)

    Dim numOfTabs As Long
    numOfTabs = 0
    
    'Start by splitting up the XML array into individual lines
    Dim xmlArray() As String
    xmlArray = Split(xmlContents, vbCrLf)
    
    'Next, loop through each line, and apply TAB characters to the start of each line as necessary
    Dim curTag As String, tagPosition As Long
    Dim i As Long, j As Long
    For i = 0 To UBound(xmlArray)
    
        'Trim any existing white space from this line
        xmlArray(i) = Trim$(xmlArray(i))
    
        'See if this line contains any tags
        tagPosition = InStr(1, xmlArray(i), "<", vbBinaryCompare)
        If tagPosition > 0 Then
        
            'This line contains a tag.  Retrieve the tag's name.
            curTag = Mid$(xmlArray(i), tagPosition + 1, InStr(1, xmlArray(i), ">", vbBinaryCompare) - tagPosition - 1)
            
            'Check for a closing tag, which would mean we need to place the current line one tab-stop to the left
            If InStr(1, curTag, "/", vbBinaryCompare) > 0 Then numOfTabs = numOfTabs - 1
            
        End If
    
        'Apply any accumulated tabs to the start of this line
        If numOfTabs > 0 Then
            For j = 1 To numOfTabs
                xmlArray(i) = vbTab & xmlArray(i)
            Next j
        End If
        
        'Increment or decrement the current tab count based on the presence of an opening tag but no closing tag
        If Not (InStr(1, curTag, "/", vbBinaryCompare) > 0) And (InStr(1, xmlArray(i), "<", vbBinaryCompare) > 0) Then
            If Not (InStr(1, xmlArray(i), "</", vbBinaryCompare) > 0) And Not (InStr(1, xmlArray(i), "<!--", vbBinaryCompare) > 0) Then numOfTabs = numOfTabs + 1
        End If
        
    Next i
    
    'Once all tabs have been inserted, reassemble the original string
    srcString = Join$(xmlArray, vbCrLf)

End Sub

'Write the current XML contents out to file.  By default, the output will have tabs added to it to make the output "pretty".
' This behavior can be avoided by setting the suppressIndentation param to TRUE.
Public Function writeXMLToFile(ByVal dstFile As String, Optional ByVal suppressIndentation As Boolean = False)

    'Make the XML contents pretty by providing some basic indentation
    Dim fileContents As String
    fileContents = xmlContents
    
    If Not suppressIndentation Then applyIndentation fileContents
    
    'If the file already exists, remove it
    If FileExist(dstFile) Then Kill dstFile
    
    'Open the specified file
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open dstFile For Output As #fileNum
        Print #fileNum, fileContents
    Close #fileNum
    
    writeXMLToFile = True

End Function

'The next block of functions returns a unique tag value in the specified format.  "Unique" tags are those that only exist once in
' a file, so their location does not matter, as they can only appear once.
Public Function getUniqueTag_String(ByVal tagName As String, Optional ByVal defaultReturn As String = "") As String
    
    Dim tmpString As String
    tmpString = getTextBetweenTags(tagName)
    
    If Len(tagName) > 0 Then
        getUniqueTag_String = CStr(tmpString)
    Else
        getUniqueTag_String = defaultReturn
    End If
    
End Function

Public Function getUniqueTag_Long(ByVal tagName As String, Optional ByVal defaultReturn As Long = 0) As String
    
    Dim tmpString As String
    tmpString = getTextBetweenTags(tagName)
    
    If Len(tagName) > 0 Then
        getUniqueTag_Long = CLng(tmpString)
    Else
        getUniqueTag_Long = defaultReturn
    End If
    
End Function

'Given a tag name, return the text between the opening and closing occurrences of that tag.  This function will always return the first
' occurence of the specified tag, starting at the specified search position (1 by default).  If the tag is not found, a blank string will
' be returned.
' Optionally, a Long-type variable can be supplied as whereTagFound if the calling function wants to know where the tag was located.
Private Function getTextBetweenTags(ByRef tagName As String, Optional ByVal searchLocation As Long = 1, Optional ByRef whereTagFound As Long = -1) As String

    Dim tagStart As Long, tagEnd As Long
    tagStart = InStr(searchLocation, xmlContents, "<" & tagName & ">", vbTextCompare)

    'If the tag was found in the file, we also need to find the closing tag.
    If tagStart > 0 Then
    
        tagEnd = InStr(tagStart, xmlContents, "</" & tagName & ">", vbTextCompare)
        
        'If the closing tag exists, return everything between that and the opening tag
        If tagEnd > tagStart Then
            
            'Increment the tag start location by the length of the tag plus two (+1 for each bracket: <>)
            tagStart = tagStart + Len(tagName) + 2
            
            'If the user passed a long, they want to know where this tag was found - return the location just after the
            ' location where the closing tag was located.
            If whereTagFound <> -1 Then whereTagFound = tagEnd + Len(tagName) + 2
            getTextBetweenTags = Mid(xmlContents, tagStart, tagEnd - tagStart)
            
        Else
            getTextBetweenTags = "ERROR: requested tag wasn't properly closed!"
        End If
        
    Else
        getTextBetweenTags = ""
    End If

End Function
