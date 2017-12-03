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
'Copyright 2013-2015 by Tanner Helland
'Created: 30/July/13
'Last updated: 18/August/14
'Last update: make the text comparison mode user-settable; some PD-centric XML files adhere to strict formatting,
'              allowing us to use binary comparison mode for much better performance.
'
'In 2013, PD became increasingly reliant on XML-format files.  The translation engine was the first to require XML
' interoperability (by design), followed a few months later by the metadata engine.  After the success of these
' two projects, a decision was made to switch all custom PhotoDemon filetypes to XML format.  This should provide
' excellent interoperability with 3rd-party projects, as well as provide a measure of future-proofing, since new
' features can be easily added without breaking old files (by simply adding new tags to file entries - tags that
' will simply be ignored by old copies of the software).
'
'Rather than write unique XML parsers for each custom filetype, this universal class was created.  It is meant to
' serve as a broad-purpose XML file interface, with strong support for reading, writing, parsing, and providing
' in-place updates to PhotoDemon-specific XML data.
'
'To that end, the primary focus of this class is supporting the (relatively barebones) XML features required for
' various PhotoDemon filetypes.  IT IS NOT MEANT TO BE A FULL-FEATURED TO-SPEC XML PARSER, but it could certainly
' be extended to support additional XML features as needed.
'
'Many thanks to Frank Donckers, who helped prototype the original translation XML engine (which heavily influenced
' the design of this class).
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'All PhotoDemon-compatible files must have their data wrapped in the following top-level tag
Private Const ROOT_TAG As String = "<pdData>"
Private Const ROOT_TAG_CLOSE As String = "</pdData>"
Private Const PD_DATA_ID As String = "<pdDataType>"
Private Const PD_DATA_ID_CLOSE As String = "</pdDataType>"

'The contents of the assigned XML file, stored in memory as one (potentially very long) string.
Private m_xmlContents As String

'Text comparison mode.  By default, this is set to m_TextCompareMode.  Outside functions can modify this via setTextCompareMode().
Private m_TextCompareMode As VbCompareMethod

'File operations are made easier by using the pdFSO class, which wraps a bunch of Unicode-friendly file APIs
Private m_File As pdFSO

Public Sub setTextCompareMode(ByVal newCompareMode As VbCompareMethod)
    m_TextCompareMode = newCompareMode
End Sub

'If this class is being used to write out a new XML file, this function can be called to initialize the blank file.
Public Sub prepareNewXML(ByVal pdDataType As String)

    m_xmlContents = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & vbCrLf
    m_xmlContents = m_xmlContents & ROOT_TAG & vbCrLf & vbCrLf & PD_DATA_ID & pdDataType & PD_DATA_ID_CLOSE & vbCrLf & ROOT_TAG_CLOSE & vbCrLf

End Sub

'PhotoDemon-specific XML files are required to encode a data type (filter, macro, etc).  This function can be used to quickly
' retrieve that type, allowing the calling function to determine if a proper filetype has been loaded for their operation.
Public Function isPDDataType(ByVal expectedType As String) As Boolean
    
    Dim reportedType As String
    reportedType = getTextBetweenTags("pdDataType")
    
    If StrComp(reportedType, expectedType, m_TextCompareMode) = 0 Then
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
    tagLocation = InStrRev(m_xmlContents, ROOT_TAG_CLOSE, , vbBinaryCompare)
    
    If tagLocation > 0 Then
    
        'Split the XML file into two halves: the half before the root tag, and the half after
        Dim topHalf As String, bottomHalf As String
        splitStringIn2 m_xmlContents, tagLocation - 1, topHalf, bottomHalf
        
        'Reassemble the primary string with a blank line inserted
        m_xmlContents = topHalf & vbCrLf & bottomHalf
        
        writeBlankLine = True
    Else
        writeBlankLine = False
    End If

End Function

'Write a comment into the XML file.  This has no practical purpose, but it can be helpful for end-users to understand the file's contents.
Public Function writeComment(ByVal commentText As String) As Boolean

    'Find the </pdData> tag that signifies the end of PD-compatible XML data
    Dim tagLocation As Long
    tagLocation = InStrRev(m_xmlContents, ROOT_TAG_CLOSE, , vbBinaryCompare)
    
    If tagLocation > 0 Then
    
        'Split the XML file into two halves: the half before the root tag, and the half after
        Dim topHalf As String, bottomHalf As String
        splitStringIn2 m_xmlContents, tagLocation - 1, topHalf, bottomHalf
        
        'Reassemble the primary string with a blank line inserted
        m_xmlContents = topHalf & "<!-- " & commentText & " -->" & vbCrLf & bottomHalf
        
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
        tagLocation = InStr(1, m_xmlContents, ROOT_TAG, vbBinaryCompare)
    Else
        tagLocation = InStrRev(m_xmlContents, ROOT_TAG_CLOSE, , vbBinaryCompare)
    End If

    If tagLocation > 0 Then
    
        'Split the XML file into two halves: the half before the root tag, and the half after
        Dim topHalf As String, bottomHalf As String
        If writeAtStart Then
            splitStringIn2 m_xmlContents, tagLocation + Len(ROOT_TAG), topHalf, bottomHalf
        Else
            splitStringIn2 m_xmlContents, tagLocation - 1, topHalf, bottomHalf
        End If
        
        'Build a string with the tag name and value we were passed
        Dim newTagLine As String
        newTagLine = "<" & tagName & ">" & tagContents
        
        If Not doNotCloseTag Then newTagLine = newTagLine & "</" & tagName & ">" & vbCrLf Else newTagLine = newTagLine & vbCrLf
        
        'Reassemble the primary string
        m_xmlContents = topHalf & newTagLine & bottomHalf
        
        writeTag = True
    Else
        writeTag = False
    End If

End Function

'Add a new section to the XML file.  To keep things simple, this is always done at the END of the file.
Public Function writeNewSection(ByVal sectionName As String, Optional ByVal sectionAttribute As String = "", Optional ByVal sectionAttributeValue As String = "") As Boolean
    
    'First, make sure the section does not already exist
    If InStr(1, m_xmlContents, "<" & sectionName & ">", m_TextCompareMode) > 0 Then
        writeNewSection = False
        Exit Function
    End If
    
    'Find the </pdData> tag that signifies the end of PD-compatible XML data
    Dim tagLocation As Long
    tagLocation = InStrRev(m_xmlContents, ROOT_TAG_CLOSE, , vbBinaryCompare)
    
    If tagLocation > 0 Then
    
        'Split the XML file into two halves: the half before the root tag, and the half after
        Dim topHalf As String, bottomHalf As String
        splitStringIn2 m_xmlContents, tagLocation - 1, topHalf, bottomHalf
        
        'Build a string with the tag name and value we were passed
        Dim newTagSection As String
        newTagSection = "<" & sectionName
        
        If Len(sectionAttribute) <> 0 Then
            newTagSection = newTagSection & " " & sectionAttribute & "=""" & sectionAttributeValue & """>"
        Else
            newTagSection = newTagSection & ">"
        End If
        
        newTagSection = newTagSection & vbCrLf & "</" & sectionName & ">" & vbCrLf & vbCrLf
        
        'Reassemble the primary XML string
        m_xmlContents = topHalf & newTagSection & bottomHalf
        
        writeNewSection = True
    Else
        writeNewSection = False
    End If
    
End Function

'Simple - does a given tag exist?  Both simple and complex tags will be checked.
Public Function doesTagExist(ByVal tagName As String, Optional ByVal attributeName As String = "", Optional ByVal attributeValue As String = "") As Boolean

    'If an attribute is provided, finding the tag is a bit messier
    If Len(attributeName) <> 0 Then
    
        If InStr(1, m_xmlContents, "<" & tagName & " " & attributeName & "=""" & attributeValue & """>", m_TextCompareMode) Then
            doesTagExist = True
        Else
            doesTagExist = False
        End If
    
    'If no attribute is provided, finding the tag is simple
    Else
    
        If InStr(1, m_xmlContents, "<" & tagName & ">", m_TextCompareMode) Then
            doesTagExist = True
        Else
            doesTagExist = False
        End If
        
    End If

End Function

'This is a bit different from the updateTag function below (which exists primarily to deal with INI-style entries).  This function asks
' for a location in the master string.  The first occurrence of the tag AFTER that location will be updated wiht the new value.  The
' current value of the tag does not need to be known.
'
'The intended purpose of this function is adding/updating translations to a PD language file.  In that case, getLocationOfParentTag is
' used to find the <phrase> wrapper, and that location is then passed to this function so the immediately following <translation> tag
' can be updated.
'
'NOTE: Unlike the updateTag function, if the requested tagName cannot be found at this location, this function will not add a new tag
' to the file.  It will simply fail and exit.  (updateTag will add the requested tag to the end of the specified section.)
Public Function updateTagAtLocation(ByVal tagName As String, ByVal newTagContents As String, Optional ByVal startLocation As Long = 1) As Boolean

    Dim tagLocation As Long
    tagLocation = InStr(startLocation, m_xmlContents, "<" & tagName & ">", m_TextCompareMode)
    
    Dim topHalf As String, bottomHalf As String
    
    'If the tag was located successfully, update it with its new contents
    If tagLocation > 0 Then
        
        'Split the XML file into two halves: the half before the relevant tag, and the half after
        Dim tagCloseLocation As Long
        tagCloseLocation = InStr(tagLocation, m_xmlContents, "</" & tagName & ">", m_TextCompareMode)
        splitStringIn2 m_xmlContents, tagCloseLocation - 1, topHalf, bottomHalf
        
        'The "topHalf" string now includes everything before the closing tag.  Chop it off at the end of the start tag,
        ' add the new contents, then add the bottom half of the original XML string.
        m_xmlContents = Left$(topHalf, tagLocation + Len(tagName) + 1) & newTagContents & bottomHalf
        
        updateTagAtLocation = True
    Else
        updateTagAtLocation = False
    End If

End Function

'Update an already existant tag located within a specific subsection of the XML file.  If the tag is not found, it will be added
' at the end of the section.
Public Function updateTag(ByVal tagName As String, ByVal tagContents As String, Optional ByVal sectionName As String = "", Optional ByVal sectionAttribute As String = "", Optional ByVal sectionAttributeValue As String = "", Optional ByVal createIfMissing As Boolean = True) As Boolean

    'Create a start and end tag to search for, which will vary contingent on the presence of a section request
    Dim startTag As String, closeTag As String
    If Len(sectionName) <> 0 Then
        If Len(sectionAttribute) <> 0 Then
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
    If Len(sectionAttribute) <> 0 Then
    
        'Finding the proper section close tag for sections with attributes is a bit trickier.  Start by finding the
        ' start location of the requested section+attribute, then find the close tag that follows that.
        sectionStartLocation = InStr(1, m_xmlContents, startTag, m_TextCompareMode)
        
        If sectionStartLocation > 0 Then
            sectionLocation = InStr(sectionStartLocation, m_xmlContents, closeTag, m_TextCompareMode)
        Else
            
            'If the section tag was not found, and createIfMissing is TRUE, create the section for the user.
            If createIfMissing Then
                writeNewSection sectionName, sectionAttribute, sectionAttributeValue
                
                'Find the start location again
                sectionStartLocation = InStr(1, m_xmlContents, startTag, m_TextCompareMode)
                sectionLocation = InStr(sectionStartLocation, m_xmlContents, closeTag, m_TextCompareMode)
                
            Else
                updateTag = False
                Exit Function
            End If
            
        End If
    
    Else
        sectionLocation = InStrRev(m_xmlContents, closeTag, , m_TextCompareMode)
        
        'If the section wasn't found, add it
        If sectionLocation = 0 Then
            writeNewSection sectionName
            sectionLocation = InStrRev(m_xmlContents, closeTag, , m_TextCompareMode)
        End If
        
    End If
    
    'We can only update the tag if its section was found.
    If sectionLocation > 0 Then
    
        'See if the tag already exists
        Dim tagLocation As Long
        tagLocation = InStrRev(m_xmlContents, "<" & tagName & ">", sectionLocation, m_TextCompareMode)
        
        Dim topHalf As String, bottomHalf As String
        
        'If the tag already exists, just update its value.  Otherwise, write out the tag as new at the end of the requested section.
        If (tagLocation > 0) And (tagLocation > sectionStartLocation) Then
        
            'Split the XML file into two halves: the half before the relevant tag, and the half after
            Dim tagCloseLocation As Long
            tagCloseLocation = InStr(tagLocation, m_xmlContents, "</" & tagName & ">", m_TextCompareMode)
            splitStringIn2 m_xmlContents, tagCloseLocation - 1, topHalf, bottomHalf
            
            'The "topHalf" string now includes everything before the closing tag.  Chop it off at the end of the start tag (e.g. after
            ' the closing bracket), add the new contents, then add the bottom half of the original XML string.
            m_xmlContents = Left$(topHalf, tagLocation + Len(tagName) + 1) & tagContents & bottomHalf
            
            updateTag = True
        
        'The tag does not exist, so we need to add it to the end of the requested section
        Else
        
            If createIfMissing Then
            
                'Split the XML file into two halves: the half before the closing tag, and the half after
                splitStringIn2 m_xmlContents, sectionLocation - 1, topHalf, bottomHalf
                
                'Build a string with the tag name and value we were passed
                Dim newTagLine As String
                newTagLine = "<" & tagName & ">" & tagContents & "</" & tagName & ">" & vbCrLf
                
                'Reassemble the primary string
                m_xmlContents = topHalf & newTagLine & bottomHalf
                
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
        tagLocation = InStr(1, m_xmlContents, ROOT_TAG, vbBinaryCompare)
    Else
        tagLocation = InStrRev(m_xmlContents, ROOT_TAG_CLOSE, , vbBinaryCompare)
    End If

    If tagLocation > 0 Then
    
        'Split the XML file into two halves: the half before the root tag, and the half after
        Dim topHalf As String, bottomHalf As String
        If writeAtStart Then
            splitStringIn2 m_xmlContents, tagLocation + Len(ROOT_TAG), topHalf, bottomHalf
        Else
            splitStringIn2 m_xmlContents, tagLocation - 1, topHalf, bottomHalf
        End If
        
        'Build a string with the tag name and value we were passed
        Dim newTagLine As String
        newTagLine = "<" & tagName & " " & tagAttribute & "=""" & attributeValue & """>" & tagContents
        
        If Not doNotCloseTag Then newTagLine = newTagLine & "</" & tagName & ">" & vbCrLf Else newTagLine = newTagLine & vbCrLf
        
        'Reassemble the primary string
        m_xmlContents = topHalf & newTagLine & bottomHalf
        
        writeTagWithAttribute = True
    Else
        writeTagWithAttribute = False
    End If

End Function

'Close a tag that has been previously left open
Public Function closeTag(ByVal tagName As String) As Boolean
    
    'Find the </pdData> tag that signifies the end of PD-compatible XML data
    Dim tagLocation As Long
    tagLocation = InStrRev(m_xmlContents, ROOT_TAG_CLOSE, , vbBinaryCompare)
    
    If tagLocation > 0 Then
    
        'Split the XML file into two halves: the half before the root tag, and the half after
        Dim topHalf As String, bottomHalf As String
        splitStringIn2 m_xmlContents, tagLocation - 1, topHalf, bottomHalf
        
        'Reassemble the primary string with the closing tag inserted
        m_xmlContents = topHalf & "</" & tagName & ">" & vbCrLf & bottomHalf
        
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
    If (InStr(1, m_xmlContents, ROOT_TAG, vbBinaryCompare) = 0) Or (InStrRev(m_xmlContents, ROOT_TAG_CLOSE, -1, vbBinaryCompare) = 0) Then
        validateLoadedXMLData = False
        Exit Function
    End If
    
    'Next, make sure the file specifies some type of PhotoDemon data
    If InStr(1, m_xmlContents, PD_DATA_ID, vbBinaryCompare) = 0 Then
        validateLoadedXMLData = False
        Exit Function
    End If

    'Search the m_xmlContents string for each tag in the validation request
    If UBound(expectedTags) >= LBound(expectedTags) Then
    
        Dim i As Long
        For i = LBound(expectedTags) To UBound(expectedTags)
            If InStr(1, m_xmlContents, expectedTags(i), m_TextCompareMode) = 0 Then
                validateLoadedXMLData = False
                Exit Function
            End If
        Next i
    
    End If
    
    validateLoadedXMLData = True

End Function

'Load an XML file from a string.  This function will also do some basic validation to ensure the requested string actually contains XML.
' Returns: TRUE if string is successfully validated and loaded.  FALSE otherwise.
Public Function loadXMLFromString(ByRef xmlString As String) As Boolean
    
    'Make an internal copy of the string
    m_xmlContents = xmlString
    
    'Validate the XML header
    If verifyXMLHeader(m_xmlContents) Then
        loadXMLFromString = True
    Else
        loadXMLFromString = False
    End If
    
End Function

'Load an XML file into memory.  This function will also do some basic validation to ensure the requested file is actually XML.
' Returns: TRUE if file found, loaded, and validated successfully.  FALSE otherwise.
Public Function loadXMLFile(ByVal xmlPath As String) As Boolean
    
    If m_File.FileExist(xmlPath) Then
    
        'Load the file's contents into a string
        If m_File.LoadTextFileAsString(xmlPath, m_xmlContents) Then
            
            'Remove all tabs from the source file.  (These pad the XML data unnecessarily, so we erase them when loading XML, and restore
            ' them when writing XML.)
            If InStr(1, m_xmlContents, vbTab, vbBinaryCompare) <> 0 Then m_xmlContents = Replace$(m_xmlContents, vbTab, "", , , vbBinaryCompare)
            
            'Check for an XML header
            If verifyXMLHeader(m_xmlContents) Then
                loadXMLFile = True
            Else
                loadXMLFile = False
            End If
            
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
    If InStr(1, fileContents, "<?xml", m_TextCompareMode) > 0 Then
        verifyXMLHeader = True
    Else
        verifyXMLHeader = False
    End If

End Function

'Retrieve an entire file and return it as a string.
Private Function getFileAsString(ByVal fName As String) As String
    
    Dim tmpFileContents As String
    
    'Ensure that the file exists before attempting to load it
    If m_File.FileExist(fName) Then
        
        Dim fileNum As Integer
        fileNum = FreeFile
        
        Open fName For Binary As #fileNum
            tmpFileContents = Space$(LOF(fileNum))
            Get #fileNum, , tmpFileContents
        Close #fileNum
        
        'Once Unicode support is added, investigate alternate means of file loading, perhaps like the code below??
        'Open fName For Input As #fileNum
        '    tmpFileContents = StrConv(InputB(LOF(fileNum), fileNum), vbUnicode)
        'Close #fileNum
    
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
    xmlArray = Split(m_xmlContents, vbCrLf)
    
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
            curTag = Mid$(xmlArray(i), tagPosition + 1, InStr(tagPosition, xmlArray(i), ">", vbBinaryCompare) - tagPosition - 1)
            
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

'Return the current XML contents as one enormous string.  By default, the output will have tabs added to it to make the output "pretty".
' This behavior can be avoided by setting the suppressIndentation param to TRUE.
Public Function returnCurrentXMLString(Optional ByVal suppressIndentation As Boolean = False) As String

    'Make the XML contents pretty by providing some basic indentation
    Dim tmpXMLCopy As String
    tmpXMLCopy = m_xmlContents
    
    If Not suppressIndentation Then applyIndentation tmpXMLCopy
    
    returnCurrentXMLString = tmpXMLCopy

End Function

'Write the current XML contents out to file.  By default, the output will have tabs added to it to make the output "pretty".
' This behavior can be avoided by setting the suppressIndentation param to TRUE.
Public Function writeXMLToFile(ByVal dstFile As String, Optional ByVal suppressIndentation As Boolean = False)
    
    On Error GoTo CouldNotWriteXML
    
    'Make the XML contents pretty by providing some basic indentation
    Dim fileContents As String
    fileContents = m_xmlContents
    
    If Not suppressIndentation Then applyIndentation fileContents
    
    'If the file contains an old-style windows-1252 encoding declaration, replace it with UTF-8
    If InStr(1, fileContents, "windows-1252", vbBinaryCompare) Then fileContents = Replace$(fileContents, "windows-1252", "UTF-8")
    
    'Allow pdFSO to handle the file write for us
    writeXMLToFile = m_File.SaveStringToTextFile(fileContents, dstFile)
    
    Exit Function
    
CouldNotWriteXML:

    Debug.Print "WARNING! A request to pdXML.writeXMLToFile failed for unknown reasons.  Please investigate."
    writeXMLToFile = False

End Function

'The next block of functions returns a unique tag value in the specified format.  "Unique" tags are those that only exist once in
' a file, so their location does not matter, as they can only appear once.
Public Function getUniqueTag_String(ByRef tagName As String, Optional ByVal defaultReturn As String = "", Optional ByVal searchLocation As Long = 1, Optional ByVal xmlSection As String = "", Optional ByVal xmlSectionAttribute As String = "", Optional ByVal xmlSectionAttributeValue As String = "") As String
    
    'If a section was provided, start our unique tag search there.  At present, we don't care if our search extends past
    ' that section, but only because we know it will never happen!
    If Len(xmlSection) <> 0 Then
    
        If Len(xmlSectionAttribute) <> 0 Then
            searchLocation = InStr(1, m_xmlContents, "<" & xmlSection & " " & xmlSectionAttribute & "=""" & xmlSectionAttributeValue & """>", m_TextCompareMode)
        Else
            searchLocation = InStr(1, m_xmlContents, "<" & xmlSection & ">", m_TextCompareMode)
        End If
        If searchLocation = 0 Then
            getUniqueTag_String = defaultReturn
            Exit Function
        End If
    End If
    
    Dim tmpString As String
    tmpString = getTextBetweenTags(tagName, searchLocation)
    
    If Len(tmpString) <> 0 Then
        getUniqueTag_String = Trim$(tmpString)
    Else
        getUniqueTag_String = defaultReturn
    End If
    
End Function

Public Function getUniqueTag_Long(ByVal tagName As String, Optional ByVal defaultReturn As Long = 0, Optional ByVal searchLocation As Long = 1) As Long
    
    Dim tmpString As String
    tmpString = getTextBetweenTags(tagName, searchLocation)
    
    If Len(tmpString) <> 0 Then
        getUniqueTag_Long = CLng(tmpString)
    Else
        getUniqueTag_Long = defaultReturn
    End If
    
End Function

Public Function getUniqueTag_Boolean(ByVal tagName As String, Optional ByVal defaultReturn As Long = False, Optional ByVal searchLocation As Long = 1) As Boolean
    
    Dim tmpString As String
    tmpString = getTextBetweenTags(tagName, searchLocation)
    
    If Len(tmpString) <> 0 Then
        getUniqueTag_Boolean = CBool(tmpString)
    Else
        getUniqueTag_Boolean = defaultReturn
    End If
    
End Function

Public Function getUniqueTag_Double(ByVal tagName As String, Optional ByVal defaultReturn As Double = 0, Optional ByVal searchLocation As Long = 1) As Double
    
    Dim tmpString As String
    tmpString = getTextBetweenTags(tagName, searchLocation)
    
    If Len(tmpString) <> 0 Then
        getUniqueTag_Double = CDblCustom(tmpString)
    Else
        getUniqueTag_Double = defaultReturn
    End If
    
End Function

'Return a location pointer immediately following the location of a given tag (assumed to be unique)
Public Function getLocationOfTag(ByVal tagName As String) As Long
    getLocationOfTag = InStr(1, m_xmlContents, "<" & tagName & ">", m_TextCompareMode)
End Function

'Return a location pointer immediately following the location of a given tag+attribute combo
Public Function getLocationOfTagPlusAttribute(ByVal tagName As String, ByVal tagAttribute As String, ByVal tagAttributeValue As String) As Long

    Dim searchLocation As Long
    searchLocation = InStr(1, m_xmlContents, "<" & tagName & " ", m_TextCompareMode)
    
    'Run a loop, finding matching tag entries as we go.
    Do While searchLocation > 0
    
        'Search location is now pointing at the location of the next tagName occurrence in the XML file.  From that location, look for
        ' a matching attribute tag.  (It's assumed that one exists...)
        Dim attributeLocation As Long
        attributeLocation = InStr(searchLocation, m_xmlContents, tagAttribute, m_TextCompareMode)
        
        'From the attribute location, we know the value has to appear immediately following tagAttribute=", so look for it there
        If StrComp(tagAttributeValue, Mid$(m_xmlContents, attributeLocation + Len(tagAttribute) + 2, Len(tagAttributeValue))) = 0 Then
            
            'A match was found!  Return this tag location and exit.
            getLocationOfTagPlusAttribute = attributeLocation + 5
            Exit Function
            
        End If
        
        'If we're here, a matching attribute was not found.  Find the next matching tag occurrence and continue.
        searchLocation = InStr(attributeLocation, m_xmlContents, "<" & tagName & " ", m_TextCompareMode)
        
    Loop
    
    'If we made it all the way here, we were unable to find a matching tag/attribute combination.
    getLocationOfTagPlusAttribute = 0
    
End Function

'Given a tag name and contents, return a pointer to the location of a specified "parent" tag.  (This is used by the
' translation engine to find the "phrase" tag enclosing an original/translation tag pair.)
Public Function getLocationOfParentTag(ByRef parentTagName As String, ByRef tagName As String, ByRef tagContents As String) As Long

    'First, find the requested tag's location
    Dim searchLocation As Long
    searchLocation = InStr(1, m_xmlContents, "<" & tagName & ">" & tagContents & "</" & tagName & ">", m_TextCompareMode)
    
    'Assuming that tag was found, look for the nearest parent tag occurrence
    If searchLocation > 0 Then
        getLocationOfParentTag = InStrRev(m_xmlContents, "<" & parentTagName & ">", searchLocation, m_TextCompareMode)
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
    tagStart = InStr(searchLocation, m_xmlContents, "<" & tagName & ">", m_TextCompareMode)

    'If the tag was not found, it's possible that the tag has an attribute inside it, making the previous check fail.  Try again now.
    If tagStart = 0 Then
        tagStart = InStr(searchLocation, m_xmlContents, "<" & tagName & " ", m_TextCompareMode)
    End If

    'If the tag was found in the file, we also need to find the closing tag.
    If tagStart > 0 Then
    
        tagEnd = InStr(tagStart, m_xmlContents, "</" & tagName & ">", m_TextCompareMode)
        
        'If the closing tag exists, return everything between that and the opening tag
        If tagEnd > tagStart Then
            
            'Increment the tag start location by the length of the tag plus two (+1 for each bracket: <>)
            tagStart = tagStart + Len(tagName) + 2
            
            'If the user passed a long, they want to know where this tag was found - return the location just after the
            ' location where the closing tag was located.
            If whereTagFound <> -1 Then whereTagFound = tagEnd + Len(tagName) + 2
            getTextBetweenTags = Mid$(m_xmlContents, tagStart, tagEnd - tagStart)
            
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
    searchLocation = InStr(searchLocation, m_xmlContents, "<" & tagName & " " & attributeName & "=""", m_TextCompareMode)
    
    Do While searchLocation > 0
    
        'Make room in the target array for the new string
        tagsFound = tagsFound + 1
        ReDim Preserve sArray(0 To tagsFound - 1) As String
        
        'This is a somewhat sloppy way to extract the attribute, but oh well - find the end of this tag line.
        endLocation = InStr(searchLocation, m_xmlContents, ">", vbBinaryCompare)
        
        'Strip out just this tag, using the start and end locations we've found
        tmpString = Mid$(m_xmlContents, searchLocation, endLocation - searchLocation)
        
        'Now, parse the string by quotation mark (").  The middle entry contains the attribute ID we want.
        tmpStringArray = Split(tmpString, """")
        sArray(tagsFound - 1) = tmpStringArray(1)
        
        'Find the next occurrence of the requested string
        searchLocation = InStr(searchLocation + 2, m_xmlContents, "<" & tagName & " " & attributeName & "=""", m_TextCompareMode)
    
    Loop
    
    'If at least one matching tag was found, return true
    If tagsFound > 0 Then
        findAllAttributeValues = True
    Else
        findAllAttributeValues = False
    End If

End Function

'Given a tag name, find all the occurrences of that tag in the file.  Those occurrence locations will be placed in a
' Long-type array, which the calling function can use to retrieve individual values at its leisure.
' NOTE: per my personal requirements for this function, it assumes simple tags only (e.g. <tagname>).  This improves
' performance.  This function could easily be modified to also find tags with attributes.
Public Function findAllTagLocations(ByRef locationArray() As Long, ByVal tagName As String, Optional ByVal useDoEvents As Boolean = False) As Boolean

    Dim numOfTags As Long, tagArrayBound As Long
    numOfTags = 0
    tagArrayBound = 100
    ReDim locationArray(0 To tagArrayBound) As Long
    
    Dim searchTag As String
    searchTag = "<" & tagName & ">"
    
    Dim searchLocation As Long
    searchLocation = 1
    
    'If the tag doesn't occur even once, fail the function and exit.
    If InStr(searchLocation, m_xmlContents, searchTag, m_TextCompareMode) > 0 Then
    
        'Find one initial tag
        searchLocation = InStr(searchLocation, m_xmlContents, searchTag, m_TextCompareMode)
    
        'Continue searching the parent string until no more occurrences are found
        Do
            
            'Store the tag's location in the results array
            locationArray(numOfTags) = searchLocation
            
            'Increment tag count.  If it exceeds the bounds of the array, resize the array to match.
            numOfTags = numOfTags + 1
            
            If numOfTags > tagArrayBound Then
                tagArrayBound = tagArrayBound * 2
                ReDim Preserve locationArray(0 To tagArrayBound) As Long
            End If
            
            'Find the next occurrence of the target tag
            searchLocation = InStr(searchLocation + 1, m_xmlContents, searchTag, m_TextCompareMode)
        
            If useDoEvents Then
                If (numOfTags And 7) = 0 Then DoEvents
            End If
        
        Loop While searchLocation > 0
    
        'Now that all tags have been located, resize the array to match the number of tags found, then exit
        ReDim Preserve locationArray(0 To numOfTags - 1) As Long
        findAllTagLocations = True
        
    Else
    
        findAllTagLocations = False
    
    End If
    
End Function

'A custom CDbl function that accepts both commas and decimals as a separator; this is important when moving floating-point data,
' represented as strings, between locales.
Private Function CDblCustom(ByVal srcString As String) As Double

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
Private Function IsNumberLocaleUnaware(ByRef Expression As String) As Boolean
    
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

Private Sub Class_Initialize()

    'This XML class uses text comparison mode by default.  Note that this is fraught with internationalization issues,
    ' so be aware of the considerations before relying on output in locales other than en-US.
    m_TextCompareMode = vbTextCompare
    
    'Prepare a module-level pdFSO instance, which will make file operations much easier
    Set m_File = New pdFSO

End Sub
