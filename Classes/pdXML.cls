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
'Copyright 2013-2026 by Tanner Helland
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
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'All PhotoDemon-compatible files must have their data wrapped in the following top-level tag
Private Const ROOT_TAG As String = "<pdData>"
Private Const ROOT_TAG_CLOSE As String = "</pdData>"
Private Const PD_DATA_ID As String = "<pdDataType>"
Private Const PD_DATA_ID_CLOSE As String = "</pdDataType>"

'To prevent creating tons of temporary string objects, we cache some XML tags at class creation time.
Private Const XML_LT As String = "<"
Private Const XML_GT As String = ">"

'The contents of the assigned XML file, stored in memory as one (potentially very long) string.
Private m_xmlContents As String

'Text comparison mode.  By default, this is set to m_TextCompareMode.  Outside functions can modify this via setTextCompareMode().
Private m_TextCompareMode As VbCompareMethod

Friend Sub SetTextCompareMode(ByVal newCompareMode As VbCompareMethod)
    m_TextCompareMode = newCompareMode
End Sub

'If this class is being used to write out a new XML file, this function can be called to initialize the blank file.
Friend Sub PrepareNewXML(ByRef pdDataType As String)
    m_xmlContents = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & vbCrLf
    m_xmlContents = m_xmlContents & ROOT_TAG & vbCrLf & vbCrLf & PD_DATA_ID & pdDataType & PD_DATA_ID_CLOSE & vbCrLf & ROOT_TAG_CLOSE & vbCrLf
End Sub

'PhotoDemon-specific XML files are required to encode a data type (filter, macro, etc).  This function can be used to quickly
' retrieve that type, allowing the calling function to determine if a proper filetype has been loaded for their operation.
Friend Function IsPDDataType(ByRef expectedType As String) As Boolean
    IsPDDataType = Strings.StringsEqual(GetTextBetweenTags("pdDataType"), expectedType, (m_TextCompareMode = vbTextCompare))
End Function

'Write a blank line into the XML file.  This has no practical purpose, but I like pretty XML output, so PD occasionally uses
' blank lines to separate tag families.
Friend Function WriteBlankLine() As Boolean

    'Find the </pdData> tag that signifies the end of PD-compatible XML data
    Dim tagLocation As Long
    tagLocation = InStrRev(m_xmlContents, ROOT_TAG_CLOSE, , vbBinaryCompare)
    
    If tagLocation > 0 Then
    
        'Split the XML file into two halves: the half before the root tag, and the half after
        Dim topHalf As String, bottomHalf As String
        SplitStringIn2 m_xmlContents, tagLocation - 1, topHalf, bottomHalf
        
        'Reassemble the primary string with a blank line inserted
        m_xmlContents = topHalf & vbCrLf & bottomHalf
        
        WriteBlankLine = True
    Else
        WriteBlankLine = False
    End If

End Function

'Write an existing string into the XML file.  The assumption is that this is valid XML the user has already prepared by some other means.
Friend Function WriteGenericText(ByRef srcString As String) As Boolean

    'Find the </pdData> tag that signifies the end of PD-compatible XML data
    Dim tagLocation As Long
    tagLocation = InStrRev(m_xmlContents, ROOT_TAG_CLOSE, , vbBinaryCompare)
    
    If tagLocation > 0 Then
    
        'Split the XML file into two halves: the half before the root tag, and the half after
        Dim topHalf As String, bottomHalf As String
        SplitStringIn2 m_xmlContents, tagLocation - 1, topHalf, bottomHalf
        
        'Reassemble the primary string with the user's text inserted
        m_xmlContents = topHalf & srcString & bottomHalf
        
        WriteGenericText = True
    Else
        WriteGenericText = False
    End If

End Function

'Write a comment into the XML file.  This has no practical purpose, but it can be helpful for end-users to understand the file's contents.
Friend Function WriteComment(ByRef commentText As String) As Boolean

    'Find the </pdData> tag that signifies the end of PD-compatible XML data
    Dim tagLocation As Long
    tagLocation = InStrRev(m_xmlContents, ROOT_TAG_CLOSE, , vbBinaryCompare)
    
    If (tagLocation > 0) Then
    
        'Split the XML file into two halves: the half before the root tag, and the half after
        Dim topHalf As String, bottomHalf As String
        SplitStringIn2 m_xmlContents, tagLocation - 1, topHalf, bottomHalf
        
        'Reassemble the primary string with a blank line inserted
        m_xmlContents = topHalf & "<!-- " & commentText & " -->" & vbCrLf & bottomHalf
        
        WriteComment = True
    Else
        WriteComment = False
    End If

End Function

'Write a new XML tag to the XML string.  By default, new tags are written to the end of the file,
' but the writeAtStart param can be set to TRUE to write tags at the top.
Friend Function WriteTag(ByVal tagName As String, ByVal tagContents As String, Optional ByVal doNotCloseTag As Boolean = False, Optional ByVal writeAtStart As Boolean = False) As Boolean

    'Find the </pdData> tag that signifies the end of PD-compatible XML data
    Dim tagLocation As Long
    If writeAtStart Then
        tagLocation = InStr(1, m_xmlContents, ROOT_TAG, vbBinaryCompare)
    Else
        tagLocation = InStrRev(m_xmlContents, ROOT_TAG_CLOSE, , vbBinaryCompare)
    End If

    If (tagLocation > 0) Then
    
        'Split the XML file into two halves: the half before the root tag, and the half after
        Dim topHalf As String, bottomHalf As String
        If writeAtStart Then
            SplitStringIn2 m_xmlContents, tagLocation + Len(ROOT_TAG), topHalf, bottomHalf
        Else
            SplitStringIn2 m_xmlContents, tagLocation - 1, topHalf, bottomHalf
        End If
        
        'Build a string with the tag name and value we were passed
        Dim newTagLine As String
        newTagLine = XML_LT & tagName & XML_GT & tagContents
        
        If (Not doNotCloseTag) Then newTagLine = newTagLine & "</" & tagName & XML_GT & vbCrLf Else newTagLine = newTagLine & vbCrLf
        
        'Reassemble the primary string
        m_xmlContents = topHalf & newTagLine & bottomHalf
        
        WriteTag = True
    Else
        WriteTag = False
    End If

End Function

'Add a new section to the XML file.  To keep things simple, this is always done at the END of the file.
Friend Function WriteNewSection(ByVal sectionName As String, Optional ByVal sectionAttribute As String = vbNullString, Optional ByVal sectionAttributeValue As String = vbNullString) As Boolean
    
    'First, make sure the section does not already exist
    If (InStr(1, m_xmlContents, XML_LT & sectionName & XML_GT, m_TextCompareMode) > 0) Then
        WriteNewSection = False
        Exit Function
    End If
    
    'Find the </pdData> tag that signifies the end of PD-compatible XML data
    Dim tagLocation As Long
    tagLocation = InStrRev(m_xmlContents, ROOT_TAG_CLOSE, , vbBinaryCompare)
    
    If (tagLocation > 0) Then
    
        'Split the XML file into two halves: the half before the root tag, and the half after
        Dim topHalf As String, bottomHalf As String
        SplitStringIn2 m_xmlContents, tagLocation - 1, topHalf, bottomHalf
        
        'Build a string with the tag name and value we were passed
        Dim newTagSection As String
        newTagSection = XML_LT & sectionName
        
        If (LenB(sectionAttribute) <> 0) Then
            newTagSection = newTagSection & " " & sectionAttribute & "=""" & sectionAttributeValue & """>"
        Else
            newTagSection = newTagSection & XML_GT
        End If
        
        newTagSection = newTagSection & vbCrLf & "</" & sectionName & XML_GT & vbCrLf & vbCrLf
        
        'Reassemble the primary XML string
        m_xmlContents = topHalf & newTagSection & bottomHalf
        
        WriteNewSection = True
    Else
        WriteNewSection = False
    End If
    
End Function

'Simple - does a given tag exist?  Both simple and complex tags will be checked.
Friend Function DoesTagExist(ByRef tagName As String, Optional ByRef attributeName As String = vbNullString, Optional ByRef attributeValue As String = vbNullString, Optional ByRef foundPosition As Long = 0) As Boolean
    
    'If an attribute is provided, finding the tag is a bit messier
    If (LenB(attributeName) <> 0) Then
        foundPosition = InStr(1, m_xmlContents, XML_LT & tagName & " " & attributeName & "=""" & attributeValue & """>", m_TextCompareMode)
        
    'If no attribute is provided, finding the tag is simple
    Else
        foundPosition = InStr(1, m_xmlContents, XML_LT & tagName & XML_GT, m_TextCompareMode)
    End If
    
    DoesTagExist = (foundPosition <> 0)

End Function

'This is a bit different from the updateTag function below (which exists primarily to deal with INI-style entries).
' This function asks for a location in the XML string.  The first occurrence of the tag AFTER that location will
' be updated with the new value.  The current value of the tag does not need to be known.
'
'The intended purpose of this function is adding/updating translations to a PD language file.  In that case,
' GetLocationOfParentTag() is used to find the <phrase> wrapper, and that location is then passed to this function
' so the immediately following <translation> tag can be updated.
'
'NOTE: Unlike the updateTag function, if the requested tagName cannot be found at this location,
' this function will not add a new tag to the file.  It will simply fail and exit.
' (UpdateTag() will add the requested tag to the end of the specified section.)
Friend Function UpdateTagAtLocation(ByVal tagName As String, ByVal newTagContents As String, Optional ByVal startLocation As Long = 1) As Boolean

    Dim tagLocation As Long
    tagLocation = InStr(startLocation, m_xmlContents, XML_LT & tagName & XML_GT, m_TextCompareMode)
    
    Dim topHalf As String, bottomHalf As String
    
    'If the tag was located successfully, update it with its new contents
    If tagLocation > 0 Then
        
        'Split the XML file into two halves: the half before the relevant tag, and the half after
        Dim tagCloseLocation As Long
        tagCloseLocation = InStr(tagLocation, m_xmlContents, "</" & tagName & XML_GT, m_TextCompareMode)
        SplitStringIn2 m_xmlContents, tagCloseLocation - 1, topHalf, bottomHalf
        
        'The "topHalf" string now includes everything before the closing tag.  Chop it off at the end of the start tag,
        ' add the new contents, then add the bottom half of the original XML string.
        m_xmlContents = Left$(topHalf, tagLocation + Len(tagName) + 1) & newTagContents & bottomHalf
        
        UpdateTagAtLocation = True
    Else
        UpdateTagAtLocation = False
    End If

End Function

'Update an already existant tag located within a specific subsection of the XML file.  If the tag is not found, it will be added
' at the end of the section.
Friend Function UpdateTag(ByVal tagName As String, ByVal tagContents As String, Optional ByVal sectionName As String = vbNullString, Optional ByVal sectionAttribute As String = vbNullString, Optional ByVal sectionAttributeValue As String = vbNullString, Optional ByVal createIfMissing As Boolean = True) As Boolean

    'Create a start and end tag to search for, which will vary contingent on the presence of a section request
    Dim startTag As String, closingTag As String
    If (LenB(sectionName) <> 0) Then
        If (LenB(sectionAttribute) <> 0) Then
            startTag = XML_LT & sectionName & " " & sectionAttribute & "=""" & sectionAttributeValue & """>"
        Else
            startTag = XML_LT & sectionName & XML_GT
        End If
        closingTag = "</" & sectionName & XML_GT
    Else
        startTag = ROOT_TAG
        closingTag = ROOT_TAG_CLOSE
    End If
    
    Dim sectionLocation As Long, sectionStartLocation As Long
    sectionStartLocation = 0
    
    'If a section is specified, add the tag at the end of that section.  Otherwise, add it at the end of the XML file.
    If (LenB(sectionAttribute) <> 0) Then
    
        'Finding the proper section close tag for sections with attributes is a bit trickier.  Start by finding the
        ' start location of the requested section+attribute, then find the close tag that follows that.
        sectionStartLocation = InStr(1, m_xmlContents, startTag, m_TextCompareMode)
        
        If (sectionStartLocation > 0) Then
            sectionLocation = InStr(sectionStartLocation, m_xmlContents, closingTag, m_TextCompareMode)
        Else
            
            'If the section tag was not found, and createIfMissing is TRUE, create the section for the user.
            If createIfMissing Then
                WriteNewSection sectionName, sectionAttribute, sectionAttributeValue
                
                'Find the start location again
                sectionStartLocation = InStr(1, m_xmlContents, startTag, m_TextCompareMode)
                sectionLocation = InStr(sectionStartLocation, m_xmlContents, closingTag, m_TextCompareMode)
                
            Else
                UpdateTag = False
                Exit Function
            End If
            
        End If
    
    Else
    
        If (LenB(sectionName) <> 0) Then
            
            sectionLocation = InStrRev(m_xmlContents, closingTag, , m_TextCompareMode)
            
            'If the section wasn't found, add it
            If (sectionLocation = 0) Then
                WriteNewSection sectionName
                sectionLocation = InStrRev(m_xmlContents, closingTag, , m_TextCompareMode)
            End If
            
        Else
            sectionLocation = InStrRev(m_xmlContents, closingTag, , m_TextCompareMode)
        End If
        
    End If
    
    'We can only update the tag if its section was found.
    If (sectionLocation > 0) Then
    
        'See if the tag already exists
        Dim tagLocation As Long
        tagLocation = InStrRev(m_xmlContents, XML_LT & tagName & XML_GT, sectionLocation, m_TextCompareMode)
        
        Dim topHalf As String, bottomHalf As String
        
        'If the tag already exists, just update its value.  Otherwise, write out the tag as new at the end of the requested section.
        If (tagLocation > 0) And (tagLocation > sectionStartLocation) Then
        
            'Split the XML file into two halves: the half before the relevant tag, and the half after
            Dim tagCloseLocation As Long
            tagCloseLocation = InStr(tagLocation, m_xmlContents, "</" & tagName & XML_GT, m_TextCompareMode)
            SplitStringIn2 m_xmlContents, tagCloseLocation - 1, topHalf, bottomHalf
            
            'The "topHalf" string now includes everything before the closing tag.  Chop it off at the end of the start tag (e.g. after
            ' the closing bracket), add the new contents, then add the bottom half of the original XML string.
            m_xmlContents = Left$(topHalf, tagLocation + Len(tagName) + 1) & tagContents & bottomHalf
            
            UpdateTag = True
        
        'The tag does not exist, so we need to add it to the end of the requested section
        Else
        
            If createIfMissing Then
            
                'Split the XML file into two halves: the half before the closing tag, and the half after
                SplitStringIn2 m_xmlContents, sectionLocation - 1, topHalf, bottomHalf
                
                'Build a string with the tag name and value we were passed
                Dim newTagLine As String
                newTagLine = XML_LT & tagName & XML_GT & tagContents & "</" & tagName & XML_GT & vbCrLf
                
                'Reassemble the primary string
                m_xmlContents = topHalf & newTagLine & bottomHalf
                
                UpdateTag = True
                
            Else
                UpdateTag = False
            End If
            
        End If
        
    Else
        UpdateTag = False
    End If

End Function

'Write a new XML tag to the XML string, including a single attribute and value.
' By default, new tags are written to the end of the file, but the writeAtStart param can be set to TRUE
' to write tags at the top.
' If you don't want the tag automatically closed, set the doNotCloseTag parameter to TRUE.
Friend Function WriteTagWithAttribute(ByVal tagName As String, ByVal tagAttribute As String, ByVal attributeValue As String, ByVal tagContents As String, Optional ByVal doNotCloseTag As Boolean = False, Optional ByVal writeAtStart As Boolean = False) As Boolean

    'Find the </pdData> tag that signifies the end of PD-compatible XML data
    Dim tagLocation As Long
    If writeAtStart Then
        tagLocation = InStr(1, m_xmlContents, ROOT_TAG, vbBinaryCompare)
    Else
        tagLocation = InStrRev(m_xmlContents, ROOT_TAG_CLOSE, , vbBinaryCompare)
    End If

    If (tagLocation > 0) Then
    
        'Split the XML file into two halves: the half before the root tag, and the half after
        Dim topHalf As String, bottomHalf As String
        If writeAtStart Then
            SplitStringIn2 m_xmlContents, tagLocation + Len(ROOT_TAG), topHalf, bottomHalf
        Else
            SplitStringIn2 m_xmlContents, tagLocation - 1, topHalf, bottomHalf
        End If
        
        'Build a string with the tag name and value we were passed
        Dim newTagLine As String
        newTagLine = XML_LT & tagName & " " & tagAttribute & "=""" & attributeValue & """>" & tagContents
        
        If (Not doNotCloseTag) Then newTagLine = newTagLine & "</" & tagName & XML_GT & vbCrLf Else newTagLine = newTagLine & vbCrLf
        
        'Reassemble the primary string
        m_xmlContents = topHalf & newTagLine & bottomHalf
        
        WriteTagWithAttribute = True
    Else
        WriteTagWithAttribute = False
    End If

End Function

'Close a tag that has been previously left open
Friend Function CloseTag(ByVal tagName As String) As Boolean
    
    'Find the </pdData> tag that signifies the end of PD-compatible XML data
    Dim tagLocation As Long
    tagLocation = InStrRev(m_xmlContents, ROOT_TAG_CLOSE, , vbBinaryCompare)
    
    If (tagLocation > 0) Then
    
        'Split the XML file into two halves: the half before the root tag, and the half after
        Dim topHalf As String, bottomHalf As String
        SplitStringIn2 m_xmlContents, tagLocation - 1, topHalf, bottomHalf
        
        'Reassemble the primary string with the closing tag inserted
        m_xmlContents = topHalf & "</" & tagName & XML_GT & vbCrLf & bottomHalf
        
        CloseTag = True
    Else
        CloseTag = False
    End If
    
End Function

'Given a string and a position, split it into two strings at that position
Private Sub SplitStringIn2(ByRef srcString As String, ByVal splitPosition As Long, ByRef dstFirstHalf As String, ByRef dstSecondHalf As String)
    dstFirstHalf = Left$(srcString, splitPosition)
    dstSecondHalf = Right$(srcString, Len(srcString) - splitPosition)
End Sub

'Count left (<) and right (>) angle brackets and return TRUE if both counts match.
' Note that this doesn't do anything more than count occurrences, so it may return FALSE in situations
' where tag contents contain a bracket.  (PD is lax about allowing this - translation files, for example,
' allow this behavior to make life easier for translators.)
'
'Because of this, do *not* use this function unless you are *certain* that the XML file can't/won't
' contain brackets within tag contents.  Otherwise, you may reject perfectly valid settings files.
'
'(Currently, PD only uses this for the primary settings file.)
Friend Function BasicXMLValidation() As Boolean
    
    Dim numLeftBrackets As Long, numRightBrackets As Long
    numLeftBrackets = Strings.CountCharOccurrences(m_xmlContents, "<")
    numRightBrackets = Strings.CountCharOccurrences(m_xmlContents, ">")
    
    BasicXMLValidation = (numLeftBrackets = numRightBrackets)
    
End Function

'Once a valid XML file has been loaded, we need to see if it contains valid XML data for the current operation.
' The client can do this by scanning for any number of tags it expects to find in the XML file.
' If all are found, return TRUE.
Friend Function ValidateLoadedXMLData(ParamArray expectedTags() As Variant) As Boolean

    'Start by looking for the <pdData> tags that surround all PhotoDemon-specific XML files
    If (InStr(1, m_xmlContents, ROOT_TAG, vbBinaryCompare) = 0) Or (InStrRev(m_xmlContents, ROOT_TAG_CLOSE, -1, vbBinaryCompare) = 0) Then
        ValidateLoadedXMLData = False
        Exit Function
    End If
    
    'Next, make sure the file specifies some type of PhotoDemon data
    If InStr(1, m_xmlContents, PD_DATA_ID, vbBinaryCompare) = 0 Then
        ValidateLoadedXMLData = False
        Exit Function
    End If

    'Search the m_xmlContents string for each tag in the validation request
    If (UBound(expectedTags) >= LBound(expectedTags)) Then
    
        Dim i As Long
        For i = LBound(expectedTags) To UBound(expectedTags)
            If InStr(1, m_xmlContents, expectedTags(i), m_TextCompareMode) = 0 Then
                ValidateLoadedXMLData = False
                Exit Function
            End If
        Next i
    
    End If
    
    ValidateLoadedXMLData = True

End Function

'Load an XML file from a string.  This function will also do some basic validation to ensure the requested string actually contains XML.
' Returns: TRUE if string is successfully validated and loaded.  FALSE otherwise.
Friend Function LoadXMLFromString(ByRef xmlString As String) As Boolean
    m_xmlContents = xmlString
    LoadXMLFromString = VerifyXMLHeader(m_xmlContents)
End Function

'Load an XML file into memory.  This function will also do some basic validation to ensure the requested file is actually XML.
' Returns: TRUE if file found, loaded, and validated successfully.  FALSE otherwise.
Friend Function LoadXMLFile(ByRef xmlPath As String) As Boolean
    
    If Files.FileLoadAsString(xmlPath, m_xmlContents) Then
            
        'Remove all tabs from the source file.  (These pad the XML data unnecessarily, so we erase them when loading XML, and restore
        ' them when writing XML.)
        If (InStr(1, m_xmlContents, vbTab, vbBinaryCompare) <> 0) Then m_xmlContents = Replace$(m_xmlContents, vbTab, vbNullString, , , vbBinaryCompare)
        
        'Check for an XML header
        LoadXMLFile = VerifyXMLHeader(m_xmlContents)
    
    Else
        LoadXMLFile = False
    End If
    
End Function

'Given an XML file (or sometimes, just the first 1024 bytes of an XML file), check to see if it has a valid XML header.
Private Function VerifyXMLHeader(ByRef fileContents As String) As Boolean
        
    'Check for "<?xml" in the file.  We don't care about encoding, etc - just check "<?xml" to keep things quick.
    VerifyXMLHeader = (InStr(1, fileContents, "<?xml", m_TextCompareMode) <> 0)

End Function

'Given an XML string, apply basic indentation
Private Sub ApplyIndentation(ByRef dstString As String)

    Dim numOfTabs As Long
    numOfTabs = 0
    
    'Start by splitting up the XML array into individual lines
    Dim xmlArray() As String
    xmlArray = Split(m_xmlContents, vbCrLf, -1, vbBinaryCompare)
    
    'Next, loop through each line, and apply TAB characters to the start of each line as necessary
    Dim curTag As String, tagPosition As Long
    Dim i As Long
    For i = 0 To UBound(xmlArray)
    
        'Trim any existing white space from this line
        xmlArray(i) = Trim$(xmlArray(i))
    
        'See if this line contains any tags
        tagPosition = InStr(1, xmlArray(i), XML_LT, vbBinaryCompare)
        If (tagPosition > 0) Then
        
            'This line contains a tag.  Retrieve the tag's name.
            curTag = Mid$(xmlArray(i), tagPosition + 1, InStr(tagPosition, xmlArray(i), XML_GT, vbBinaryCompare) - tagPosition - 1)
            
            'Check for a closing tag, which would mean we need to place the current line one tab-stop to the left
            If InStr(1, curTag, "/", vbBinaryCompare) > 0 Then numOfTabs = numOfTabs - 1
            
        End If
    
        'Apply any accumulated tabs to the start of this line
        If (numOfTabs > 0) Then xmlArray(i) = String$(numOfTabs, vbTab) & xmlArray(i)
        
        'Increment or decrement the current tab count based on the presence of an opening tag but no closing tag
        If (InStr(1, curTag, "/", vbBinaryCompare) = 0) And (InStr(1, xmlArray(i), XML_LT, vbBinaryCompare) > 0) Then
            If (InStr(1, xmlArray(i), "</", vbBinaryCompare) = 0) And (InStr(1, xmlArray(i), "<!--", vbBinaryCompare) = 0) And (InStr(1, xmlArray(i), "<?", vbBinaryCompare) = 0) Then numOfTabs = numOfTabs + 1
        End If
        
    Next i
    
    'Finally, remove any lines following the trailing </pdData> tag
    i = UBound(xmlArray)
    Do While (InStr(1, xmlArray(i), ROOT_TAG_CLOSE, vbBinaryCompare) = 0) And (i > 0)
        i = i - 1
    Loop
    
    'The i variable is now pointing at the line number of the closing tag.  ReDim the array to remove anything past this point.
    If (i > 0) Then ReDim Preserve xmlArray(0 To i + 1) As String
    
    'Once all tabs have been inserted, reassemble the original string
    dstString = Join$(xmlArray, vbCrLf)

End Sub

'Return the current XML contents as one enormous string.  By default, the output will have tabs added to it to make the output "pretty".
' This behavior can be avoided by setting the suppressIndentation param to TRUE.
Friend Function ReturnCurrentXMLString(Optional ByVal suppressIndentation As Boolean = False) As String
    If suppressIndentation Then ReturnCurrentXMLString = m_xmlContents Else ApplyIndentation ReturnCurrentXMLString
End Function

'Return a pointer to the current XML string.  All the usual caveats of pointers in VB apply, obviously.
Friend Function GetXMLStringPtr() As Long
    GetXMLStringPtr = StrPtr(m_xmlContents)
End Function

'Return the length of the current XML string.  It's assumed this function will be used in conjunction with
' GetXMLStringPtr(), above.
Friend Function GetXMLStringLen_Chars() As Long
    GetXMLStringLen_Chars = Len(m_xmlContents)
End Function

'Write the current XML contents out to file.  By default, the output will have tabs added to it to make the output "pretty".
' This behavior can be avoided by setting the suppressIndentation param to TRUE.
Friend Function WriteXMLToFile(ByRef dstFile As String, Optional ByVal suppressIndentation As Boolean = False) As Boolean
    
    On Error GoTo CouldNotWriteXML
    
    If (LenB(m_xmlContents) <> 0) Then
    
        'Make the XML contents pretty by providing some basic indentation
        Dim fileContents As String
        If suppressIndentation Then fileContents = m_xmlContents Else ApplyIndentation fileContents
        
        'If the file contains an old-style windows-1252 encoding declaration, replace it with UTF-8
        If (InStr(1, fileContents, "windows-1252", vbBinaryCompare) <> 0) Then fileContents = Replace$(fileContents, "windows-1252", "UTF-8")
        
        'Allow pdFSO to handle the file write for us
        WriteXMLToFile = Files.FileSaveAsText(fileContents, dstFile)
        
    Else
        WriteXMLToFile = False
    End If
    
    Exit Function
    
CouldNotWriteXML:
    PDDebug.LogAction "WARNING! A request to pdXML.WriteXMLToFile failed for unknown reasons (#" & Err.Number & ", " & Err.Description & ") on file (" & dstFile & ")."
    WriteXMLToFile = False

End Function

'The next block of functions returns a unique tag value in the specified format.  "Unique" tags are those that only exist once in
' a file, so their location does not matter, as they can only appear once.
Friend Function GetUniqueTag_String(ByRef tagName As String, Optional ByVal defaultReturn As String = vbNullString, Optional ByVal searchLocation As Long = 1, Optional ByVal xmlSection As String = vbNullString, Optional ByVal xmlSectionAttribute As String = vbNullString, Optional ByVal xmlSectionAttributeValue As String = vbNullString) As String
    
    'If a section was provided, start our unique tag search there.  At present, we don't care if our search extends past
    ' that section, but only because we know it will never happen!
    If (LenB(xmlSection) <> 0) Then
    
        If (LenB(xmlSectionAttribute) <> 0) Then
            searchLocation = InStr(1, m_xmlContents, XML_LT & xmlSection & " " & xmlSectionAttribute & "=""" & xmlSectionAttributeValue & """>", m_TextCompareMode)
        Else
            searchLocation = InStr(1, m_xmlContents, XML_LT & xmlSection & XML_GT, m_TextCompareMode)
        End If
        
        If (searchLocation = 0) Then
            GetUniqueTag_String = defaultReturn
            Exit Function
        End If
        
    End If
    
    GetUniqueTag_String = GetTextBetweenTags(tagName, searchLocation)
    
    If (LenB(GetUniqueTag_String) <> 0) Then
        GetUniqueTag_String = Trim$(GetUniqueTag_String)
    Else
        GetUniqueTag_String = defaultReturn
    End If
    
End Function

Friend Function GetUniqueTag_StringEx(ByRef tagName As String, ByRef tagAttribute As String, ByRef tagAttributeValue As String, Optional ByVal defaultReturn As String = vbNullString, Optional ByVal searchLocation As Long = 1) As String
    
    Dim tagStart As Long, tagEnd As Long
    If GetTagCharacterRange(tagStart, tagEnd, tagName, tagAttribute, tagAttributeValue, searchLocation) Then
        GetUniqueTag_StringEx = Mid$(m_xmlContents, tagStart, tagEnd - tagStart)
    Else
        GetUniqueTag_StringEx = defaultReturn
    End If
        
End Function

Friend Function GetUniqueTag_Long(ByRef tagName As String, Optional ByVal defaultReturn As Long = 0, Optional ByVal searchLocation As Long = 1) As Long
    
    Dim tmpString As String
    tmpString = GetTextBetweenTags(tagName, searchLocation)
    
    If (LenB(tmpString) <> 0) Then
        GetUniqueTag_Long = CLng(tmpString)
    Else
        GetUniqueTag_Long = defaultReturn
    End If
    
End Function

Friend Function GetUniqueTag_Boolean(ByRef tagName As String, Optional ByVal defaultReturn As Boolean = False, Optional ByVal searchLocation As Long = 1) As Boolean
    
    Dim tmpString As String
    tmpString = GetTextBetweenTags(tagName, searchLocation)
    
    If (LenB(tmpString) <> 0) Then
        GetUniqueTag_Boolean = CBool(tmpString)
    Else
        GetUniqueTag_Boolean = defaultReturn
    End If
    
End Function

Friend Function GetUniqueTag_Double(ByRef tagName As String, Optional ByVal defaultReturn As Double = 0, Optional ByVal searchLocation As Long = 1) As Double
    
    Dim tmpString As String
    tmpString = GetTextBetweenTags(tagName, searchLocation)
    
    If (LenB(tmpString) <> 0) Then
        GetUniqueTag_Double = CDblCustom(tmpString)
    Else
        GetUniqueTag_Double = defaultReturn
    End If
    
End Function

Friend Function GetUniqueTag_Currency(ByRef tagName As String, Optional ByVal defaultReturn As Currency = 0@, Optional ByVal searchLocation As Long = 1) As Currency
    
    Dim tmpString As String
    tmpString = GetTextBetweenTags(tagName, searchLocation)
    
    If (LenB(tmpString) <> 0) Then
        GetUniqueTag_Currency = CCur(tmpString)
    Else
        GetUniqueTag_Currency = defaultReturn
    End If
    
End Function

'The next block of functions returns non-unique tag values in a specified format.  "Non-unique" tags are those that may exist
' in multiple sections within a file.  A value will only be returned if the requested tag appears within the specified XML section.
Friend Function GetNonUniqueTag_String(ByRef tagName As String, ByRef xmlSection As String, Optional ByRef xmlSectionAttribute As String = vbNullString, Optional ByRef xmlSectionAttributeValue As String = vbNullString, Optional ByRef defaultReturn As String = vbNullString, Optional ByVal defaultStartPosition As Long = 1) As String
        
    Dim searchStartLocation As Long, searchEndLocation As Long
    
    'If a section was provided, start our unique tag search there.  At present, we don't care if our search extends past
    ' that section, but only because we know it will never happen!
    If (LenB(xmlSection) <> 0) Then
    
        If (LenB(xmlSectionAttribute) <> 0) Then
            searchStartLocation = InStr(defaultStartPosition, m_xmlContents, XML_LT & xmlSection & " " & xmlSectionAttribute & "=""" & xmlSectionAttributeValue & """>", m_TextCompareMode)
            If (searchStartLocation <> 0) Then searchEndLocation = InStr(searchStartLocation, m_xmlContents, "</" & xmlSection & XML_GT, m_TextCompareMode)
        Else
            searchStartLocation = InStr(defaultStartPosition, m_xmlContents, XML_LT & xmlSection & XML_GT, m_TextCompareMode)
            If (searchStartLocation <> 0) Then searchEndLocation = InStr(searchStartLocation, m_xmlContents, "</" & xmlSection & XML_GT, m_TextCompareMode)
        End If
        
        If (searchStartLocation = 0) Then
            GetNonUniqueTag_String = defaultReturn
            Exit Function
        End If
        
        If (searchEndLocation = 0) Then searchEndLocation = LONG_MAX
    Else
        searchStartLocation = defaultStartPosition
        searchEndLocation = LONG_MAX
    End If
    
    Dim tmpString As String, foundLocation As Long
    tmpString = GetTextBetweenTags(tagName, searchStartLocation, foundLocation)
    
    If (LenB(tmpString) <> 0) And (foundLocation < searchEndLocation) Then
        GetNonUniqueTag_String = Trim$(tmpString)
    Else
        GetNonUniqueTag_String = defaultReturn
    End If
    
End Function

'Given a precise character position, return the value of the tag at that position.
Friend Function GetTagValueAtPreciseLocation(ByVal startPosition As Long) As String

    'Find the next XML_GT after the requested position occurrence.
    Dim startPos As Long
    startPos = InStr(startPosition, m_xmlContents, XML_GT, vbBinaryCompare)
    
    'Find the next XML_LT after startPos, above.
    Dim endPos As Long
    endPos = InStr(startPos + 1, m_xmlContents, XML_LT, vbBinaryCompare)
    
    'Return the value between the two positions
    If (endPos > startPos) Then
        GetTagValueAtPreciseLocation = Mid$(m_xmlContents, startPos + 1, (endPos - startPos) - 1)
    Else
        GetTagValueAtPreciseLocation = vbNullString
    End If

End Function

'Return a location pointer immediately following the location of a given tag (assumed to be unique)
Friend Function GetLocationOfTag(ByVal tagName As String, Optional ByVal startLocation As Long = 1) As Long
    GetLocationOfTag = InStr(startLocation, m_xmlContents, XML_LT & tagName & XML_GT, m_TextCompareMode)
End Function

'Return a location pointer immediately following the location of a given tag+attribute combo
Friend Function GetLocationOfTagPlusAttribute(ByVal tagName As String, ByVal tagAttribute As String, ByVal tagAttributeValue As String, Optional ByVal startPosition As Long = 1) As Long

    Dim searchLocation As Long
    searchLocation = InStr(startPosition, m_xmlContents, XML_LT & tagName & " ", m_TextCompareMode)
    
    'Run a loop, finding matching tag entries as we go.
    Do While searchLocation > 0
    
        'Search location is now pointing at the location of the next tagName occurrence in the XML file.  From that location, look for
        ' a matching attribute tag.  (It's assumed that one exists...)
        Dim attributeLocation As Long
        attributeLocation = InStr(searchLocation, m_xmlContents, tagAttribute, m_TextCompareMode)
        
        'From the attribute location, we know the value has to appear immediately following tagAttribute=", so look for it there
        If StrComp(tagAttributeValue, Mid$(m_xmlContents, attributeLocation + Len(tagAttribute) + 2, Len(tagAttributeValue))) = 0 Then
            
            'A match was found!  Return this tag location and exit.
            GetLocationOfTagPlusAttribute = attributeLocation + 5
            Exit Function
            
        End If
        
        'If we're here, a matching attribute was not found.  Find the next matching tag occurrence and continue.
        searchLocation = InStr(attributeLocation, m_xmlContents, XML_LT & tagName & " ", m_TextCompareMode)
        
    Loop
    
    'If we made it all the way here, we were unable to find a matching tag/attribute combination.
    GetLocationOfTagPlusAttribute = 0
    
End Function

'Return the valid range (character positions) of a given XML block.  Start and End position Long values must be provided by the caller.
' Unlike other functions, this returns a TRUE if successful; FALSE otherwise.
Friend Function GetTagCharacterRange(ByRef startPosition As Long, ByRef endPosition As Long, ByRef tagName As String, Optional ByRef tagAttribute As String = vbNullString, Optional ByRef tagAttributeValue As String = vbNullString, Optional ByVal searchStartPosition As Long = 1) As Boolean
    
    Dim tagFound As Boolean
    tagFound = False
    
    Dim searchLocation As Long
    searchLocation = InStr(searchStartPosition, m_xmlContents, XML_LT & tagName & " ", m_TextCompareMode)
        
    'Run a loop, finding matching tag entries as we go.
    Do While (searchLocation > 0)
    
        'Search location is now pointing at the location of the next tagName occurrence in the XML file.
        
        'If a matching attribute wasn't specified, exit now.
        If (LenB(tagAttribute) = 0) Then
            tagFound = True
            Exit Do
        
        'If a matching attribute *was* specified, try to match it now.
        Else
            
            'Within this tag, see if the attribute matches the one supplied by the user.  (It's assumed than an attribute tag
            ' exists, and is distinct from other occurrences in the file.)
            Dim attributeLocation As Long
            attributeLocation = InStr(searchLocation, m_xmlContents, tagAttribute, m_TextCompareMode)
            
            'From the attribute location, we know the value has to appear immediately following tagAttribute=", so look for it there
            If (StrComp(tagAttributeValue, Mid$(m_xmlContents, attributeLocation + Len(tagAttribute) + 2, Len(tagAttributeValue)), m_TextCompareMode) = 0) Then
                
                'A match was found!  Mark this tag location and exit.
                tagFound = True
                
                'The + 4 is used because there are four extra characters, two before the attribute value and two after (=",">)
                searchLocation = attributeLocation + Len(tagAttribute) + Len(tagAttributeValue) + 4
                Exit Do
                
            End If
            
        End If
        
        'If we reach this line, a matching attribute was not found.  Find the next matching tag occurrence and continue.
        searchLocation = InStr(attributeLocation, m_xmlContents, XML_LT & tagName & " ", m_TextCompareMode)
        
    Loop
    
    'If the tag was located successfully, searchLocation will be pointing at it.
    If tagFound Then
    
        'We now have a start position.  Calculate a matching end position.
        endPosition = InStr(searchLocation, m_xmlContents, "</" & tagName & XML_GT, m_TextCompareMode)
        
        GetTagCharacterRange = (endPosition > 0)
        If GetTagCharacterRange Then startPosition = searchLocation
        
    'The requested tag could not be located.  Return FALSE.
    Else
        GetTagCharacterRange = False
    End If
    
End Function

'Given a tag name and contents, return a pointer to the location of a specified "parent" tag.  (This is used by the
' translation engine to find the "phrase" tag enclosing an original/translation tag pair.)
Friend Function GetLocationOfParentTag(ByRef parentTagName As String, ByRef tagName As String, ByRef tagContents As String, Optional ByVal useWideChars As Boolean = True) As Long

    'First, find the requested tag's location
    Dim searchLocation As Long
    If (m_TextCompareMode = vbBinaryCompare) Then
        searchLocation = Strings.StrStrBM(m_xmlContents, XML_LT & tagName & XML_GT & tagContents & "</" & tagName & XML_GT, , useWideChars)
    Else
        searchLocation = InStr(1, m_xmlContents, XML_LT & tagName & XML_GT & tagContents & "</" & tagName & XML_GT, m_TextCompareMode)
    End If
    
    'Assuming that tag was found, look for the nearest parent tag occurrence
    If (searchLocation > 0) Then
        GetLocationOfParentTag = InStrRev(m_xmlContents, XML_LT & parentTagName & XML_GT, searchLocation, m_TextCompareMode)
    Else
        GetLocationOfParentTag = 0
    End If
    
End Function

'Given a tag name, return the text between the opening and closing occurrences of that tag.  This function will always return the first
' occurence of the specified tag, starting at the specified search position (1 by default).  If the tag is not found, a blank string will
' be returned.
' Optionally, a Long-type variable can be supplied as whereTagFound if the calling function wants to know where the tag was located.
Private Function GetTextBetweenTags(ByRef tagName As String, Optional ByVal searchLocation As Long = 1, Optional ByRef whereTagFound As Long = -1) As String

    Dim tagStart As Long, tagEnd As Long
    tagStart = InStr(searchLocation, m_xmlContents, XML_LT & tagName & XML_GT, m_TextCompareMode)

    'If the tag was not found, it's possible that the tag has an attribute inside it, making the previous check fail.  Try again now.
    If (tagStart = 0) Then
        tagStart = InStr(searchLocation, m_xmlContents, XML_LT & tagName & " ", m_TextCompareMode)
    End If

    'If the tag was found in the file, we also need to find the closing tag.
    If (tagStart > 0) Then
    
        tagEnd = InStr(tagStart, m_xmlContents, "</" & tagName & XML_GT, m_TextCompareMode)
        
        'If the closing tag exists, return everything between that and the opening tag
        If (tagEnd > tagStart) Then
            
            'Increment the tag start location by the length of the tag plus two (+1 for each bracket: <>)
            tagStart = tagStart + Len(tagName) + 2
            
            'If the user passed a long, they want to know where this tag was found - return the location just after the
            ' location where the closing tag was located.
            whereTagFound = tagEnd + Len(tagName) + 2
            GetTextBetweenTags = Mid$(m_xmlContents, tagStart, tagEnd - tagStart)
            
        Else
            GetTextBetweenTags = "ERROR: requested tag wasn't properly closed!"
        End If
        
    Else
        GetTextBetweenTags = vbNullString
    End If

End Function

'Given a string, replace any characters that are not allowed with underscores; this is used to generate dynamic tag names
Friend Function GetXMLSafeTagName(ByRef srcText As String) As String

    Dim goodString As String
    
    'Remove any incidental white space before processing
    goodString = Trim$(srcText)
    
    'Create a string of valid numerical characters, based on the XML spec at http://www.w3.org/TR/1998/REC-xml-19980210.html#sec-common-syn
    Const validChars As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.-_:"
    
    'Loop through the text box contents and remove any invalid characters
    Dim i As Long
    For i = 1 To Len(goodString)
        
        'Compare a single character from the text against our list of valid characters. If a character is NOT found
        ' in the list of valid characters, replace it with an underscore.
        If (InStr(1, validChars, Mid$(goodString, i, 1), vbBinaryCompare) = 0) Then Mid$(goodString, i, 1) = "_"
            
    Next i
    
    GetXMLSafeTagName = goodString

End Function

'Given a tag name and attribute type, find all the matching attribute values in the file.  The calling function can then use
' these to pull specific tags from a given tag/attribute section.
Friend Function FindAllAttributeValues(ByRef sArray() As String, ByVal tagName As String, ByVal attributeName As String) As Boolean

    ReDim sArray(0) As String
    Dim tmpString As String, tmpStringArray() As String
    
    Dim tagsFound As Long
    tagsFound = 0
    
    'Find the first occurrence of the string in the file (if any)
    Dim searchLocation As Long, endLocation As Long
    searchLocation = InStr(1, m_xmlContents, XML_LT & tagName & " " & attributeName & "=""", m_TextCompareMode)
    
    Do While (searchLocation > 0)
    
        'Make room in the target array for the new string
        tagsFound = tagsFound + 1
        ReDim Preserve sArray(0 To tagsFound - 1) As String
        
        'This is a somewhat sloppy way to extract the attribute, but oh well - find the end of this tag line.
        endLocation = InStr(searchLocation, m_xmlContents, XML_GT, vbBinaryCompare)
        
        'Strip out just this tag, using the start and end locations we've found
        tmpString = Mid$(m_xmlContents, searchLocation, endLocation - searchLocation)
        
        'Now, parse the string by quotation mark (").  The middle entry contains the attribute ID we want.
        tmpStringArray = Split(tmpString, """")
        sArray(tagsFound - 1) = tmpStringArray(1)
        
        'Find the next occurrence of the requested string
        searchLocation = InStr(searchLocation + 2, m_xmlContents, XML_LT & tagName & " " & attributeName & "=""", m_TextCompareMode)
    
    Loop
    
    'If at least one matching tag was found, return true
    FindAllAttributeValues = (tagsFound > 0)

End Function

'Given a tag name, find all the occurrences of that tag in the file.  Those occurrence locations will be placed in a
' Long-type array, which the calling function can use to retrieve individual values at its leisure.
' NOTE: per my personal requirements for this function, it assumes simple tags only (e.g. <tagname>).  This improves
' performance.  This function could easily be modified to also find tags with attributes.
Friend Function FindAllTagLocations(ByRef locationArray() As Long, ByRef tagName As String) As Boolean
    
    FindAllTagLocations = False
    
    Dim numOfTags As Long
    numOfTags = 0
    
    Const INIT_OCCURRENCES_COUNT As Long = 4096
    Dim tagArrayBound As Long
    tagArrayBound = INIT_OCCURRENCES_COUNT - 1
    
    ReDim locationArray(0 To tagArrayBound) As Long
    
    Dim searchTag As String
    searchTag = XML_LT & tagName & XML_GT
    
    Dim searchLocation As Long
    searchLocation = 1
    
    'If the tag doesn't occur even once, fail the function and exit.
    If (InStr(searchLocation, m_xmlContents, searchTag, m_TextCompareMode) > 0) Then
    
        'Find one initial tag
        searchLocation = InStr(searchLocation, m_xmlContents, searchTag, m_TextCompareMode)
    
        'Continue searching the parent string until no more occurrences are found
        Do
            
            'Store the tag's location in the results array
            locationArray(numOfTags) = searchLocation
            
            'Increment tag count.  If it exceeds the bounds of the array, resize the array to match.
            numOfTags = numOfTags + 1
            If (numOfTags > tagArrayBound) Then
                tagArrayBound = tagArrayBound * 2 + 1
                ReDim Preserve locationArray(0 To tagArrayBound) As Long
            End If
            
            'Find the next occurrence of the target tag
            searchLocation = InStr(searchLocation + 1, m_xmlContents, searchTag, m_TextCompareMode)
            
        Loop While (searchLocation > 0)
        
        'Now that all tags have been located, resize the array to match the number of tags found, then exit
        FindAllTagLocations = (numOfTags > 0)
        If FindAllTagLocations Then ReDim Preserve locationArray(0 To numOfTags - 1) As Long
        
    End If
    
End Function

Private Sub Class_Initialize()

    'This XML class uses text comparison mode by default.  Note that this is fraught with internationalization issues,
    ' so be aware of the considerations before relying on output in locales other than en-US.
    m_TextCompareMode = vbTextCompare
    
End Sub
