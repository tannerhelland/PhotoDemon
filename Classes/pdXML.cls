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
'Last updated: 03/September/13
'Last update: added getLocationOfParentTag() function, which the translation engine needs for finding the <phrase>
'             tag that encloses an original/translation text pair
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
'Many thanks to Frank Donckers, who helped prototype the original translation XML engine (which heavily influenced
' the design and code of this class).
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'All PhotoDemon-compatible files must have their data wrapped in the following top-level tag
Private Const ROOT_TAG As String = "<pdData>"
Private Const ROOT_TAG_CLOSE As String = "</pdData>"
Private Const PD_DATA_ID As String = "<pdDataType>"
Private Const PD_DATA_ID_CLOSE As String = "</pdDataType>"

'The contents of the assigned XML file, stored in memory as one (potentially very long) string.
Private xmlContents As String

'If this class is being used to write out a new XML file, this function can be called to initialize the blank file.
Public Sub prepareNewXML(ByVal pdDataType As String)

    xmlContents = "<?xml version=""1.0"" encoding=""windows-1252""?>" & vbCrLf & vbCrLf
    xmlContents = xmlContents & ROOT_TAG & vbCrLf & vbCrLf & PD_DATA_ID & pdDataType & PD_DATA_ID_CLOSE & vbCrLf & ROOT_TAG_CLOSE & vbCrLf

End Sub

'PhotoDemon-specific XML files are required to encode a data type (filter, macro, etc).  This function can be used to quickly
' retrieve that type, allowing the calling function to determine if a proper filetype has been loaded for their operation.
Public Function isPDDataType(ByVal expectedType As String) As Boolean
    
    Dim reportedType As String
    reportedType = getTextBetweenTags("pdDataType")
    
    If StrComp(reportedType, expectedType, vbTextCompare) = 0 Then
        isPDDataType = True
    Else
        isPDDataType = False
    End If
    
End Function

'Write a blank line into the XML file.  This has no practical purpose, but I'm a big fan of pretty XML output, so I like to use
' blank lines to separate tag families.
Public Function writeBlankLine() As Boolean

    'Find the </pdData> tag that signifies the end of PD-compatible XML data
    Dim tagLocation As Long
    tagLocation = InStrRev(xmlContents, ROOT_TAG_CLOSE, , vbBinaryCompare)
    
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
    tagLocation = InStrRev(xmlContents, ROOT_TAG_CLOSE, , vbBinaryCompare)
    
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
Public Function writeTag(ByVal tagName As String, ByVal tagContents As String, Optional ByVal doNotCloseTag As Boolean = False, Optional ByVal writeAtStart As Boolean = False) As Boolean

    'Find the </pdData> tag that signifies the end of PD-compatible XML data
    Dim tagLocation As Long
    If writeAtStart Then
        tagLocation = InStr(1, xmlContents, ROOT_TAG, vbBinaryCompare)
    Else
        tagLocation = InStrRev(xmlContents, ROOT_TAG_CLOSE, , vbBinaryCompare)
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
        newTagLine = "<" & tagName & ">" & tagContents
        
        If Not doNotCloseTag Then newTagLine = newTagLine & "</" & tagName & ">" & vbCrLf Else newTagLine = newTagLine & vbCrLf
        
        'Reassemble the primary string
        xmlContents = topHalf & newTagLine & bottomHalf
        
        writeTag = True
    Else
        writeTag = False
    End If

End Function

'Simple - does a given tag exist?  Both simple and complex tags will be checked.
Public Function doesTagExist(ByVal tagName As String, Optional ByVal attributeName As String = "", Optional ByVal attributeValue As String = "") As Boolean

    'If an attribute is provided, finding the tag is a bit messier
    If Len(attributeName) > 0 Then
    
        If InStr(1, xmlContents, "<" & tagName & " " & attributeName & "=""" & attributeValue & """>", vbTextCompare) Then
            doesTagExist = True
        Else
            doesTagExist = False
        End If
    
    'If no attribute is provided, finding the tag is simple
    Else
    
        If InStr(1, xmlContents, "<" & tagName & ">", vbTextCompare) Then
            doesTagExist = True
        Else
            doesTagExist = False
        End If
        
    End If

End Function

'Update an already existant tag located within a specific subsection of the XML file.  If the tag is not found, it will be added
' at the end of the section.
Public Function updateTag(ByVal tagName As String, ByVal tagContents As String, Optional ByVal sectionName As String = "", Optional ByVal sectionAttribute As String = "", Optional ByVal sectionAttributeValue As String = "", Optional ByVal createIfMissing As Boolean = True) As Boolean

    'Create a start and end tag to search for, which will vary contingent on the presence of a section request
    Dim startTag As String, closeTag As String
    If Len(sectionName) > 0 Then
        If Len(sectionAttribute) > 0 Then
            startTag = "<" & sectionName & " " & sectionAttribute & "=""" & sectionAttributeValue & """>"
        Else
            startTag = "<" & sectionName & ">"
        End If
        closeTag = "</" & sectionName & ">"
    Else
        startTag = ROOT_TAG
        closeTag = ROOT_TAG_CLOSE
    End If
    
    Dim sectionLocation As Long, sectionStartLocation As Long
    sectionStartLocation = 0
    
    'If a section is specified, add the tag at the end of that section.  Otherwise, add it at the end of the XML file.
    If Len(sectionAttribute) > 0 Then
    
        'Finding the proper section close tag for sections with attributes is a bit trickier.  Start by finding the
        ' start location of the requested section+attribute, then find the close tag that follows that.
        sectionStartLocation = InStr(1, xmlContents, startTag, vbTextCompare)
        
        If sectionStartLocation > 0 Then
            sectionLocation = InStr(sectionStartLocation, xmlContents, closeTag, vbTextCompare)
        Else
            
            'If the section tag was not found, we should really create the section for the user.  But as that's not a
            ' a possibility at present, simply fail and exit.  (In the future, if new sections are added to the
            ' user prefs file, I will need to implement this.)
            updateTag = False
            Exit Function
        End If
    
    Else
        sectionLocation = InStrRev(xmlContents, closeTag, , vbTextCompare)
    End If
    
    'We can only update the tag if its section was found.
    If sectionLocation > 0 Then
    
        'See if the tag already exists
        Dim tagLocation As Long
        tagLocation = InStrRev(xmlContents, "<" & tagName & ">", sectionLocation, vbTextCompare)
        
        Dim topHalf As String, bottomHalf As String
        
        'If the tag already exists, just update its value.  Otherwise, write out the tag as new at the end of the requested section.
        If (tagLocation > 0) And (tagLocation > sectionStartLocation) Then
        
            'Split the XML file into two halves: the half before the relevant tag, and the half after
            Dim tagCloseLocation As Long
            tagCloseLocation = InStr(tagLocation, xmlContents, "</" & tagName & ">", vbTextCompare)
            splitStringIn2 xmlContents, tagCloseLocation - 1, topHalf, bottomHalf
            
            'The "topHalf" string now includes everything before the closing tag.  We want to split it again, leaving out the
            ' previous value so we can replace it with our own.
            
            'Reassemble the string with the new value
            xmlContents = Left$(topHalf, tagLocation + Len(tagName) + 1) & tagContents & bottomHalf
            
            updateTag = True
        
        'The tag does not exist, so we need to add it to the end of the requested section
        Else
        
            If createIfMissing Then
            
                'Split the XML file into two halves: the half before the closing tag, and the half after
                splitStringIn2 xmlContents, sectionLocation - 1, topHalf, bottomHalf
                
                'Build a string with the tag name and value we were passed
                Dim newTagLine As String
                newTagLine = "<" & tagName & ">" & tagContents & "</" & tagName & ">" & vbCrLf
                
                'Reassemble the primary string
                xmlContents = topHalf & newTagLine & bottomHalf
                
                updateTag = True
                
            Else
                updateTag = False
            End If
            
        End If
        
    Else
        updateTag = False
    End If

End Function

'Write a new XML tag to the master XML string, including a single attribute and value.  By default, new tags are written to the end of
' the file, but the writeAtStart param can be set to TRUE to write tags at the top.
' If you don't want the tag automatically closed, set the doNotCloseTag parameter to TRUE.
Public Function writeTagWithAttribute(ByVal tagName As String, ByVal tagAttribute As String, ByVal attributeValue As String, ByVal tagContents As String, Optional ByVal doNotCloseTag As Boolean = False, Optional ByVal writeAtStart As Boolean = False) As Boolean

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
        newTagLine = "<" & tagName & " " & tagAttribute & "=""" & attributeValue & """>" & tagContents
        
        If Not doNotCloseTag Then newTagLine = newTagLine & "</" & tagName & ">" & vbCrLf Else newTagLine = newTagLine & vbCrLf
        
        'Reassemble the primary string
        xmlContents = topHalf & newTagLine & bottomHalf
        
        writeTagWithAttribute = True
    Else
        writeTagWithAttribute = False
    End If

End Function

'Close a tag that has been previously left open
Public Function closeTag(ByVal tagName As String) As Boolean
    
    'Find the </pdData> tag that signifies the end of PD-compatible XML data
    Dim tagLocation As Long
    tagLocation = InStrRev(xmlContents, ROOT_TAG_CLOSE, , vbBinaryCompare)
    
    If tagLocation > 0 Then
    
        'Split the XML file into two halves: the half before the root tag, and the half after
        Dim topHalf As String, bottomHalf As String
        splitStringIn2 xmlContents, tagLocation - 1, topHalf, bottomHalf
        
        'Reassemble the primary string with the closing tag inserted
        xmlContents = topHalf & "</" & tagName & ">" & vbCrLf & bottomHalf
        
        closeTag = True
    Else
        closeTag = False
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
    
    'Next, make sure the file specifies some type of PhotoDemon data
    If InStr(1, xmlContents, PD_DATA_ID, vbBinaryCompare) = 0 Then
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
    
    'Finally, remove any lines following the trailing </pdData> tag
    i = UBound(xmlArray)
    Do While (InStr(1, xmlArray(i), ROOT_TAG_CLOSE, vbBinaryCompare) = 0) And (i > 0)
        i = i - 1
    Loop
    
    'The i variable is now pointing at the line number of the closing tag.  ReDim the array to remove anything past this point.
    If i > 0 Then ReDim Preserve xmlArray(0 To i + 1) As String
    
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
    
    'If the file already exists, remove it.  (I have also added a DoEvents here after noticing random errors in this
    ' sub - it's important to wait for the file to be deleted, so that the write attempt below does not fail.)
    If FileExist(dstFile) Then
        Kill dstFile
        DoEvents
    End If
    
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
Public Function getUniqueTag_String(ByVal tagName As String, Optional ByVal defaultReturn As String = "", Optional ByVal searchLocation As Long = 1, Optional ByVal xmlSection As String = "", Optional ByVal xmlSectionAttribute As String = "", Optional ByVal xmlSectionAttributeValue As String = "") As String
    
    'If a section was provided, start our unique tag search there.  At present, we don't care if our search extends past
    ' that section, but only because we know it will never happen!
    If Len(xmlSection) > 0 Then
    
        If Len(xmlSectionAttribute) > 0 Then
            searchLocation = InStr(1, xmlContents, "<" & xmlSection & " " & xmlSectionAttribute & "=""" & xmlSectionAttributeValue & """>", vbTextCompare)
        Else
            searchLocation = InStr(1, xmlContents, "<" & xmlSection & ">", vbTextCompare)
        End If
        If searchLocation = 0 Then
            getUniqueTag_String = defaultReturn
            Exit Function
        End If
    End If
    
    Dim tmpString As String
    tmpString = getTextBetweenTags(tagName, searchLocation)
    
    If Len(tagName) > 0 Then
        getUniqueTag_String = CStr(tmpString)
    Else
        getUniqueTag_String = defaultReturn
    End If
    
End Function

Public Function getUniqueTag_Long(ByVal tagName As String, Optional ByVal defaultReturn As Long = 0, Optional ByVal searchLocation As Long = 1) As String
    
    Dim tmpString As String
    tmpString = getTextBetweenTags(tagName, searchLocation)
    
    If Len(tagName) > 0 Then
        getUniqueTag_Long = CLng(tmpString)
    Else
        getUniqueTag_Long = defaultReturn
    End If
    
End Function

Public Function getUniqueTag_Double(ByVal tagName As String, Optional ByVal defaultReturn As Double = 0, Optional ByVal searchLocation As Long = 1) As Double
    
    Dim tmpString As String
    tmpString = getTextBetweenTags(tagName, searchLocation)
    
    If Len(tagName) > 0 Then
        getUniqueTag_Double = CDbl(tmpString)
    Else
        getUniqueTag_Double = defaultReturn
    End If
    
End Function

'Return a location pointer immediately following the location of a given tag (assumed to be unique)
Public Function getLocationOfTag(ByVal tagName As String) As Long
    getLocationOfTag = InStr(1, xmlContents, "<" & tagName & ">", vbTextCompare)
End Function

'Return a location pointer immediately following the location of a given tag+attribute combo
Public Function getLocationOfTagPlusAttribute(ByVal tagName As String, ByVal tagAttribute As String, ByVal tagAttributeValue As String) As Long

    Dim searchLocation As Long
    searchLocation = InStr(1, xmlContents, "<" & tagName & " ", vbTextCompare)
    
    'Run a loop, finding matching tag entries as we go.
    Do While searchLocation > 0
    
        'Search location is now pointing at the location of the next tagName occurrence in the XML file.  From that location, look for
        ' a matching attribute tag.  (It's assumed that one exists...)
        Dim attributeLocation As Long
        attributeLocation = InStr(searchLocation, xmlContents, tagAttribute, vbTextCompare)
        
        'From the attribute location, we know the value has to appear immediately following tagAttribute=", so look for it there
        If StrComp(tagAttributeValue, Mid$(xmlContents, attributeLocation + Len(attributeLocation) + 3, Len(tagAttributeValue))) = 0 Then
            
            'A match was found!  Return this tag location and exit.
            getLocationOfTagPlusAttribute = attributeLocation + 5
            Exit Function
            
        End If
        
        'If we're here, a matching attribute was not found.  Find the next matching tag occurrence and continue.
        searchLocation = InStr(attributeLocation, xmlContents, "<" & tagName & " ", vbTextCompare)
        
    Loop
    
    'If we made it all the way here, we were unable to find a matching tag/attribute combination.
    getLocationOfTagPlusAttribute = 0
    
End Function

'Given a tag name and contents, return a pointer to the location of a specified "parent" tag.  (This is used by the
' translation engine to find the "phrase" tag enclosing an original/translation tag pair.)
Public Function getLocationOfParentTag(ByVal parentTagName As String, ByVal tagName As String, ByVal tagContents As String) As Long

    'First, find the requested tag's location
    Dim searchLocation As Long
    searchLocation = InStr(1, xmlContents, "<" & tagName & ">" & tagContents & "</" & tagName & ">", vbTextCompare)
    
    'Assuming that tag was found, look for the nearest parent tag occurrence
    If searchLocation > 0 Then
        getLocationOfParentTag = InStrRev(xmlContents, "<" & parentTagName & ">", searchLocation, vbTextCompare)
    Else
        getLocationOfParentTag = 0
    End If
    
End Function

'Given a tag name, return the text between the opening and closing occurrences of that tag.  This function will always return the first
' occurence of the specified tag, starting at the specified search position (1 by default).  If the tag is not found, a blank string will
' be returned.
' Optionally, a Long-type variable can be supplied as whereTagFound if the calling function wants to know where the tag was located.
Private Function getTextBetweenTags(ByRef tagName As String, Optional ByVal searchLocation As Long = 1, Optional ByRef whereTagFound As Long = -1) As String

    Dim tagStart As Long, tagEnd As Long
    tagStart = InStr(searchLocation, xmlContents, "<" & tagName & ">", vbTextCompare)

    'If the tag was not found, it's possible that the tag has an attribute inside it, making the previous check fail.  Try again now.
    If tagStart = 0 Then
        tagStart = InStr(searchLocation, xmlContents, "<" & tagName & " ", vbTextCompare)
    End If

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
            getTextBetweenTags = Mid$(xmlContents, tagStart, tagEnd - tagStart)
            
        Else
            getTextBetweenTags = "ERROR: requested tag wasn't properly closed!"
        End If
        
    Else
        getTextBetweenTags = ""
    End If

End Function

'Given a string, replace any characters that are not allowed with underscores; this is used to generate dynamic tag names
Public Function getXMLSafeTagName(ByVal srcText As String) As String

    Dim goodString As String
    
    'Remove any incidental white space before processing
    goodString = Trim(srcText)
    
    'Create a string of valid numerical characters, based on the XML spec at http://www.w3.org/TR/1998/REC-xml-19980210.html#sec-common-syn
    Dim validChars As String
    validChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
    validChars = validChars & "0123456789"
    validChars = validChars & ".-_:"
    
    'Loop through the text box contents and remove any invalid characters
    Dim i As Long
    
    For i = 1 To Len(goodString)
        
        'Compare a single character from the text against our list of valid characters. If a character is NOT found
        ' in the list of valid characters, replace it with an underscore.
        If InStr(validChars, Mid$(goodString, i, 1)) = 0 Then
            goodString = Left$(goodString, i - 1) & "_" & Right$(goodString, Len(goodString) - i)
        End If
            
    Next i
    
    getXMLSafeTagName = goodString

End Function

'Given a tag name and attribute type, find all the matching attribute values in the file.  The calling function can then use
' these to pull specific tags from a given tag/attribute section.
Public Function findAllAttributeValues(ByRef sArray() As String, ByVal tagName As String, ByVal attributeName As String) As Boolean

    ReDim sArray(0) As String
    Dim tmpString As String, tmpStringArray() As String
    
    Dim tagsFound As Long
    tagsFound = 0
    
    'Find the first occurrence of the string in the file (if any)
    Dim searchLocation As Long, endLocation As Long
    searchLocation = 1
    searchLocation = InStr(searchLocation, xmlContents, "<" & tagName & " " & attributeName & "=""", vbTextCompare)
    
    Do While searchLocation > 0
    
        'Make room in the target array for the new string
        tagsFound = tagsFound + 1
        ReDim Preserve sArray(0 To tagsFound - 1) As String
        
        'This is a somewhat sloppy way to extract the attribute, but oh well - find the end of this tag line.
        endLocation = InStr(searchLocation, xmlContents, ">", vbBinaryCompare)
        
        'Strip out just this tag, using the start and end locations we've found
        tmpString = Mid$(xmlContents, searchLocation, endLocation - searchLocation)
        
        'Now, parse the string by quotation mark (").  The middle entry contains the attribute ID we want.
        tmpStringArray = Split(tmpString, """")
        sArray(tagsFound - 1) = tmpStringArray(1)
        
        'Find the next occurrence of the requested string
        searchLocation = InStr(searchLocation + 2, xmlContents, "<" & tagName & " " & attributeName & "=""", vbTextCompare)
    
    Loop
    
    'If at least one matching tag was found, return true
    If tagsFound > 0 Then
        findAllAttributeValues = True
    Else
        findAllAttributeValues = False
    End If

End Function
