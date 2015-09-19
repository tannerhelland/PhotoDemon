VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdMetadata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Image Metadata Handler
'Copyright 2013-2015 by Tanner Helland
'Created: 23/May/13
'Last updated: 05/January/15
'Last update: perform special checks for certain PNG chunks (sRGB, cHRM, gAMA).  PD now processes the data in these
'              chunks, if present, which allows us to provide identical rendering to Firefox and Chrome.
'
'As of version 6.0, PhotoDemon now provides support for loading and saving image metadata.  What is metadata, you ask?
' See http://en.wikipedia.org/wiki/Metadata#Photographs for more details.
'
'Because metadata handling is complicated, I have created this dedicated metadata handler class.  Each pdImage object
' has one (1) metadata class called "imgMetadata".  That class is publicly accessible.  If metadata is available at
' load-time, the imgMetadata object will automatically (and asynchronously) be filled with any metadata found in the
' source file.
'
'At present, metadata handling is 100% tied to the ExifTool plugin (see the "Plugin_ExifTool_Interface" module for more
' detail).  Because metadata is insanely complex to both read and write - if you want it done correctly - I have no
' plans to write a fallback for users without the plugin.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'The number of metadata items currently stored
Private maxMetadataCount As Long

'Array that contains the metadata itself
Private iMetadata() As mdItem

'When extracting metadata, we request a full XML dump from ExifTool.  This XML file is stored - in its entirety - in memory,
' as it is the only way to precisely restore the XML data in its *exact* form when save-time comes.
Private originalMetadataXML As String

'If GPS data is found, mark its location
Private hasGPSCoordinates As Boolean
Private gpsIndex As Long

'Current metadata export setting for this image
Private metadataExportPreference As Long

'Return the original metadata XML string.  Note that aside from some minimal processing (e.g. removing ExifTool's {ready} marker),
' this will be the unmodified XML string originally passed to this class, e.g. without any parsing or validation of the data!
Public Function getOriginalXMLMetadataString() As String
    getOriginalXMLMetadataString = originalMetadataXML
End Function

'Retrieve resolution data from the metadata collection, if it exists.  Returns TRUE if resolution metadata entrist exist; FALSE
' if they do not.  Note that this function always returns resolution in inches.  Pixels per CM measurements will be automatically
' converted to PPI.
Public Function getResolution(ByRef xResolution As Double, ByRef yResolution As Double) As Boolean

    On Error GoTo mdResolutionError

    Dim xResTagIndex As Long, yResTagIndex As Long

    If doesTagExist("XResolution", xResTagIndex) And doesTagExist("YResolution", yResTagIndex) Then
    
        'Check for the presence of a resolution type.  If the resolution is in CM, we need to convert it to inches.
        ' Similarly, if no unit is specified, assume inches.
        Dim unitIsCentimeters As Boolean
        unitIsCentimeters = False
        
        Dim resUnitIndex As Long
        If doesTagExist("ResolutionUnit", resUnitIndex) Then
        
            'Note that 3 is a magic number used by EXIF to indicate CM.  See http://www.sno.phy.queensu.ca/~phil/exiftool/TagNames/EXIF.html
            If CLng(iMetadata(resUnitIndex).ActualValue) = 3 Then unitIsCentimeters = True
            
        End If
        
        'Retrieve X and Y resolution values
        xResolution = CDblCustom(iMetadata(xResTagIndex).Value)
        yResolution = CDblCustom(iMetadata(yResTagIndex).Value)
        
        'If the units are in centimeters, convert them to inches
        If unitIsCentimeters Then
            xResolution = getInchesFromCM(xResolution)
            yResolution = getInchesFromCM(yResolution)
        End If
        
        'Return TRUE!
        getResolution = True
    
    'Resolution tags don't exist.  Return FALSE.
    Else
    
        getResolution = False
    
    End If
    
    Exit Function
    
'If an error occurs, return FALSE
mdResolutionError:

    getResolution = False

End Function

'Does a given tag exist in the metadata collection?  Returns TRUE if yes, FALSE, if no.
' If the tag does exist, the optional byref tagIndex will be set to the index.  It will be set to -1 if the tag is not found.
Private Function doesTagExist(ByVal tagName As String, Optional ByRef tagIndex As Long = -1, Optional ByVal searchMode As VbCompareMethod = vbTextCompare) As Boolean

    doesTagExist = False
    
    'If this image has no metadata, exit now
    If maxMetadataCount <= 0 Then Exit Function

    'Scan the metadata collection, looking for the given tag name
    Dim i As Long
    For i = 0 To maxMetadataCount - 1
    
        'Tag name found!  Mark the index and exit.
        If StrComp(iMetadata(i).Name, tagName, searchMode) = 0 Then
        
            doesTagExist = True
            tagIndex = i
            Exit For
        
        End If
        
    Next i

End Function

'Does a given tag exist in the metadata collection?  Returns TRUE if yes, FALSE, if no.
' If the tag does exist, the optional byref tagIndex will be set to the index.  It will be set to -1 if the tag is not found.
Private Function doesTagExistFullName(ByVal tagNameAndGroup As String, Optional ByRef tagIndex As Long = -1, Optional ByVal searchMode As VbCompareMethod = vbTextCompare) As Boolean

    doesTagExistFullName = False
    
    'If this image has no metadata, exit now
    If maxMetadataCount <= 0 Then Exit Function
    
    'Scan the metadata collection, looking for the given tag name
    Dim i As Long
    For i = 0 To maxMetadataCount - 1
    
        'Tag name found!  Mark the index and exit.
        If StrComp(iMetadata(i).FullGroupAndName, tagNameAndGroup, searchMode) = 0 Then
        
            doesTagExistFullName = True
            tagIndex = i
            Exit For
        
        End If
        
    Next i

End Function

'Return the actual value of a given tag.  If the tag is not found, a null string is returned.
' It is assumed that the caller has used doesTagExistFullName() before calling this function, so they know whether a null string is
' the actual return value of the given metadata entry, or a fail state.
Private Function getTagValue(ByVal tagNameAndGroup As String, Optional ByVal searchMode As VbCompareMethod = vbTextCompare) As String

    Dim tagIndex As Long
    
    If doesTagExistFullName(tagNameAndGroup, tagIndex, searchMode) Then
        getTagValue = getMetadataEntry(tagIndex).Value
    Else
        getTagValue = ""
    End If

End Function

'Metadata export preferences are typically handled according a global user setting (settable via Tools -> Options).  When a
' new image is loaded, its metadata handling preference will be set to that.  However, some dialogs (e.g. JPEG) allow the user
' to override metadata handling for a given image.  In that case, this preference may deviate from the global setting.
Public Sub setMetadataExportPreference(ByVal newExportPreference As Long)
    metadataExportPreference = newExportPreference
End Sub

'Given an index and a source metadata object, fill that object with the contents of the matching entry in the iMetadata() array
Friend Function getMetadataEntry(ByVal objIndex As Long) As mdItem

    'Make sure the passed index is valid
    If (objIndex < maxMetadataCount) And (objIndex >= 0) Then getMetadataEntry = iMetadata(objIndex)

End Function

'Return the total metadata count for this object.  Note that binary and list items ARE included in this count.
Public Function getMetadataCount() As Long
    getMetadataCount = maxMetadataCount
End Function

'PD now supports a few different metadata storage methods.  This will return TRUE if XML metadata is present.
Public Function hasXMLMetadata() As Boolean
    
    'If metadata has been loaded in XML format, maxMetadataCount will be > 0.
    If maxMetadataCount = 0 Then
        hasXMLMetadata = False
    Else
        hasXMLMetadata = True
    End If
    
End Function

'It can be desirable to know if we have checked for GPS data and not found any, vs we have NOT looked for GPS data and we have no
' idea if the image contains any.
Public Function haveAttemptedToFindGPSData() As Boolean
    If maxMetadataCount = 0 Then
        haveAttemptedToFindGPSData = False
    Else
        haveAttemptedToFindGPSData = True
    End If
End Function

'Some metadata options require the presence of GPS data in the photo.  If a GPS entry was found at metadata load-time, return TRUE.
Public Function hasGPSMetadata() As Boolean
    hasGPSMetadata = hasGPSCoordinates
End Function

'If GPS metadata exists, this function can be used to retrieve the exact latitude and longitude values
Public Function fillLatitudeLongitude(ByRef latString As String, ByRef lonString As String) As Boolean
    
    If hasGPSCoordinates Then
    
        Dim tmpSplit() As String
        tmpSplit = Split(iMetadata(gpsIndex).ActualValue, " ")
        latString = tmpSplit(0)
        lonString = tmpSplit(1)
    
        fillLatitudeLongitude = True
    
    Else
        fillLatitudeLongitude = False
    End If
    
End Function

'This function will return a new XML file string with all "marked for removal" entries removed
Private Function retrieveModifiedXMLString() As String
    
    'If this image has no metadata, exit now
    If maxMetadataCount <= 0 Then
        retrieveModifiedXMLString = ""
        Exit Function
    End If
    
    Dim newXMLString As String
    newXMLString = originalMetadataXML
    
    Dim tagStart As Long, tagEnd As Long
    
    'Note that external functions are responsible for marking which entries they want removed.  This function presumes that all
    ' modifications have already been marked, so all we need to do is fulfill the actual removal step.
    Dim i As Long
    For i = 0 To maxMetadataCount - 1
    
        If iMetadata(i).markedForRemoval Then
        
            'Find this tag's start position
            tagStart = InStr(1, newXMLString, "<" & iMetadata(i).FullGroupAndName)
            
            'Next, find this tag's end position
            tagEnd = InStr(tagStart, newXMLString, "</" & iMetadata(i).FullGroupAndName) + Len(iMetadata(i).FullGroupAndName) + 3
            
            'Use those positions to fully remove the tag from the master XML string
            newXMLString = Left$(newXMLString, tagStart - 1) & Right$(newXMLString, Len(newXMLString) - tagEnd)
        
        End If
    
    Next i
        
    retrieveModifiedXMLString = newXMLString
    
End Function

'Given a valid XML-formatted metadata string from ExifTool, parse it into usable PD metadata format.
' (The parent image's ID format is also passed to this function, because a custom "ready" tag with the image's ID must be manually
'  removed from the XML string.)
Public Function loadAllMetadata(ByVal mdString As String, ByVal parentImageID As Long) As Boolean

    'If the ExifTool plugin can't be found, abandon all metadata loading attempts
    If Not g_ExifToolEnabled Then
        Message "ExifTool plugin missing.  Metadata import abandoned."
        loadAllMetadata = False
        Exit Function
    End If
    
    maxMetadataCount = 0
    hasGPSCoordinates = False: gpsIndex = 0
    
    'If no metadata was found, we can safely exit now
    If Len(mdString) = 0 Then
        loadAllMetadata = False
        Exit Function
    End If
    
    Dim i As Long
    
    'DEBUGGING ONLY: sometimes it is helpful to dump the full metadata text to the clipboard.  Uncomment as necessary.
    
    'This rather ungainly set of code is required to copy Unicode text to the clipboard
    'Dim tmpUnicode As pdUnicode
    'Set tmpUnicode = New pdUnicode
    'Dim tmpBytes() As Byte
    'ReDim tmpBytes(0 To Len(mdString) * 2) As Byte
    'CopyMemory ByVal VarPtr(tmpBytes(0)), ByVal StrPtr(mdString), UBound(tmpBytes)
    'Dim tmpClip As cCustomClipboard
    'Set tmpClip = New cCustomClipboard
    'tmpClip.ClearClipboard
    'tmpClip.ClipboardOpen FormMain.hWnd
    'tmpClip.SetBinaryData 13, tmpBytes      '13 is CF_UNICODE
    'tmpClip.ClipboardClose
    
    'VB clipboard text copy is much easier (but note that a forcible ANSI conversion will be applied)
    'Clipboard.Clear
    'Clipboard.SetText mdString
    
    'Due to recent changes to the way PD handles asynchronous metadata parsing, the returned string will contain text similar to
    ' {ready123}, where 123 is the ID number of this image.  Because we don't actually want to store that text, use the supplied
    ' parentImageID to find and remove the text in question.
    If InStrRev(mdString, "{ready", , vbBinaryCompare) <> 0 Then mdString = Replace$(mdString, "{ready" & parentImageID & "}", "")
    
    'We now have a (potentially very large) string of XML-formatted metadata.  Before doing anything else, make a backup copy of
    ' that string.  We need this backup when it comes time to write the metadata back out to file.
    originalMetadataXML = mdString
    
    'Our next job is to parse our very large XML string into individual components within the iMetadata() array. We accomplish this by
    ' stepping through the XML and looking for tag/value pairs as we go.  Because the XML is precisely structured, we can rely on a few
    ' key markers to do this successfully.
    
    'Start by chopping up the string into an array, delimited into individual lines
    Dim xmlLines() As String
    xmlLines = Split(mdString, vbCrLf)
    
    'Each line that contains a tag name will also contain the text "rdf:parseType='Resource'".  Because we don't know what the tag
    ' or group name will be, we look for that text instead.
    Dim curLine As Long
    curLine = 0
    
    Dim ignoreThisTag As Boolean
    
    Dim tmpString As String
    Dim tmpSplitOuter() As String, tmpSplitInner() As String
    
    Dim charPos As Long
    
    'If character sets are unknown, we don't really have a choice but to perform a blind Base-64 to text conversion.
    ' pdUnicode can help us do that.
    Dim uniHelper As pdUnicode
    Set uniHelper = New pdUnicode
    Dim tmpBytes() As Byte
    
    ReDim iMetadata(0) As mdItem
    
    Do
    
        'If our definitive text is found, parse this line to retrieve this tag's Group and Name values
        If InStr(1, UCase$(xmlLines(curLine)), "RDF:PARSETYPE='RESOURCE'", vbBinaryCompare) > 0 Then
            
            ignoreThisTag = False
        
            'Divide the tag into two segments: the prefix (which contains Group and Name) and the rdf parse-type, which we don't care about.
            tmpSplitOuter = Split(xmlLines(curLine), "rdf:parseType='Resource")
            
            'Store the full group and name text, because we will use that to relocate this tag (if necessary) at write-time
            tmpString = Trim$(tmpSplitOuter(0))
            iMetadata(maxMetadataCount).FullGroupAndName = Right$(tmpString, Len(tmpString) - 1)
            
            'Now, divide the prefix into two segments: Group and Name
            tmpSplitInner = Split(tmpSplitOuter(0), ":")
            
            'It is now trivial to retrieve the Group and Name values
            iMetadata(maxMetadataCount).Group = Trim$(Right$(tmpSplitInner(0), Len(tmpSplitInner(0)) - 2))
            iMetadata(maxMetadataCount).Name = Trim$(tmpSplitInner(1))
                        
            'A description line always follows a group/name line.  Retrieve that value now.
            curLine = curLine + 1
            xmlLines(curLine) = Trim$(xmlLines(curLine))
            iMetadata(maxMetadataCount).Description = Trim$(Mid$(xmlLines(curLine), 10, InStr(1, xmlLines(curLine), "</et") - 10))
            
            'A human-friendly value always follows a description line, but parsing it is more difficult than the first two entries.
            ' Why?  The variety of Value properties that can exist.  My testing shows three main types of values:
            
            ' 1) Simple value.  This is the most common type of entry, and it looks like:
            '    <et:prt>value</et:prt>
            '    These values are easy to parse.
            
            ' 2) Complex (binary) value.  Used for things like thumbnails, and not generally relevant to display to the user.  Looks like:
            '    <et:prt rdf:datatype='http://www.w3.org/2001/XMLSchema#base64Binary'>
            '       zgACAAkABgAxAABhY3NwTVNGVAAAAABJRUMgc1JHQgAAAAAAAAAAAAAAAAAA
            '       ...many more lines of gibberish...
            '    </et:prt>
            '    Not difficult to parse, but FYI there is really no reason to store or attempt to convert the specific binary values.
            '    One exception to this is user comments, which can include newline characters and thus fool the auto-detector - note
            '    that we perform a special check for this case.
            
            ' 3) List-type values.  Used primarily by PhotoShop.  Looks like:
            '<et:prt>
            '  <rdf:Bag>
            '    <rdf:li>saved</rdf:li>
            '    <rdf:li>derived</rdf:li>
            '    <rdf:li>saved</rdf:li>
            '    ...many more lines of this
            '  </rdf:Bag>
            '</et:prt>
            '    More problematic than the first two types, as a LOT of rdf lines can exist, and there is possible merit to extracting their values.
            '    At present, these are condensed into a single string delimited by semicolons.  Might revisit this approach in the future.
            
            'Using this knowledge of various data types, check to see if the value is binary or list-type.  If it is,
            ' mark it and proceed to ignore the tag's full contents.
            curLine = curLine + 1
            
            'Simple data will exist on a single line; use that knowledge to separate such values out.
            If (InStr(1, xmlLines(curLine), "<et:prt>") > 0) And (InStr(1, xmlLines(curLine), "</et:prt>") > 0) Then
                iMetadata(maxMetadataCount).isValueBinary = False
                iMetadata(maxMetadataCount).isValueList = False
            Else
                'The current value is either binary, or a list.  Check for the presence of an <rdf:Bag> tag to distinguish between the two.
                If (InStr(1, xmlLines(curLine + 1), "<rdf:Bag>") > 0) Then
                    iMetadata(maxMetadataCount).isValueBinary = False
                    iMetadata(maxMetadataCount).isValueList = True
                    iMetadata(maxMetadataCount).Value = "(list of values too long to display)"
                Else
                    
                    'If a datatype is specified, this is true binary data.
                    If InStr(1, xmlLines(curLine), "rdf:datatype=") > 0 Then
                    
                        iMetadata(maxMetadataCount).isValueBinary = True
                        iMetadata(maxMetadataCount).isValueList = False
                        iMetadata(maxMetadataCount).isValueMultiLine = False
                        
                        'Check to see if the data is Base-64 encoded.  If it is, we'll retrieve the Base64 string for future reference.
                        ' (This is most relevant for IPTC data, as many metadata writers will default to a local codepage but not actually
                        '  embed the codepage in the file.  This makes retrieval pretty much impossible, but we may be able to expose an
                        '  option to the user in the future, so they can attempt to apply a manual codepage conversion.)
                        If InStr(1, xmlLines(curLine), "base64Binary") > 0 Then
                            iMetadata(maxMetadataCount).isValueBase64 = True
                            iMetadata(maxMetadataCount).Value = ""
                        Else
                            iMetadata(maxMetadataCount).isValueBase64 = False
                            iMetadata(maxMetadataCount).Value = "(binary data)"
                        End If
                    
                    'If a datatype is not specified, this is most likely a non-binary value with (non-spec!) line-breaks.  This invalid
                    ' output confuses the parser, but we will attempt to retrieve it via special means later in the function.
                    Else
                    
                        iMetadata(maxMetadataCount).isValueBinary = True
                        iMetadata(maxMetadataCount).isValueList = False
                        iMetadata(maxMetadataCount).isValueMultiLine = True
                        
                    End If
                        
                End If
            End If
            
            If iMetadata(maxMetadataCount).isValueBinary Or iMetadata(maxMetadataCount).isValueList Then
            
                'SPECIAL CHECK FOR COMMENTS
                'Exif specifically disallows linebreak characters, but metadata is prone to these kinds of exceptions.  If the previous parse
                ' step detected a potential multiline comment (defined by a line with an open tag, but no rdf type declaration), or if ExifTool
                ' deemed it necessary to provide the data in Base-64 because the character set wasn't obvious, we will attempt to retrieve the
                ' data into a string.  In the case of multiline text, we'll manually replace linefeed characters with spaces to bring the
                ' metadata into spec.
                If iMetadata(maxMetadataCount).isValueMultiLine Or iMetadata(maxMetadataCount).isValueBase64 Then
                
                    'First, Retrieve its value using a custom parse strategy.
                    xmlLines(curLine) = Trim$(xmlLines(curLine))
                    
                    charPos = InStrRev(xmlLines(curLine), ">", , vbBinaryCompare)
                    If charPos > 0 Then
                        iMetadata(maxMetadataCount).Value = Trim$(Right$(xmlLines(curLine), Len(xmlLines(curLine)) - charPos))
                    End If
                    
                    Do
                        curLine = curLine + 1
                        
                        xmlLines(curLine) = Trim$(xmlLines(curLine))
                        
                        'Ignore empty lines of any input type
                        If Len(xmlLines(curLine)) > 0 Then
                            
                            'Check for the end of this XML section; if present, a custom parse strategy is required
                            If InStr(1, xmlLines(curLine), "</et:prt>") > 0 Then
                            
                                'See if there are any useful characters on this line; if there are, add them to the comment before exiting
                                If Len(xmlLines(curLine)) > 9 Then
                                    
                                    'Base-64 is loaded untouched
                                    If iMetadata(maxMetadataCount).isValueBase64 Then
                                        iMetadata(maxMetadataCount).Value = iMetadata(maxMetadataCount).Value & Left$(xmlLines(curLine), Len(xmlLines(curLine)) - 9)
                                    
                                    'Multiline text is converted to single-line text
                                    Else
                                        iMetadata(maxMetadataCount).Value = iMetadata(maxMetadataCount).Value & " " & Left$(xmlLines(curLine), Len(xmlLines(curLine)) - 9)
                                    End If
                                    
                                End If
                            
                            'If there is no closing XML tag (yet), grab the entire line without modification
                            Else
                                
                                'Base-64 is loaded untouched
                                If iMetadata(maxMetadataCount).isValueBase64 Then
                                    iMetadata(maxMetadataCount).Value = iMetadata(maxMetadataCount).Value & xmlLines(curLine)
                                
                                'Multiline text is converted to single-line text
                                Else
                                    iMetadata(maxMetadataCount).Value = iMetadata(maxMetadataCount).Value & " " & Left$(xmlLines(curLine), Len(xmlLines(curLine)) - 9)
                                End If
                                
                            End If
                        
                        'End for blank strings
                        End If
                        
                    Loop While InStr(1, xmlLines(curLine), "</et:prt>") = 0
                    
                    'Remove dead space from the parsed comment string
                    If iMetadata(maxMetadataCount).isValueBase64 Then
                        
                        'Store the Base64 value for future usage
                        iMetadata(maxMetadataCount).Base64Value = Trim$(iMetadata(maxMetadataCount).Value)
                        
                        'Generate the value by a heuristic conversion; in the future, we can try something smarter, but for now,
                        ' this at least gives us a baseline value.
                        uniHelper.Base64Decode tmpBytes, iMetadata(maxMetadataCount).Base64Value
                        uniHelper.ConvertUnknownBytesToString tmpBytes, iMetadata(maxMetadataCount).Value, False
                        
                    Else
                        iMetadata(maxMetadataCount).Value = Trim$(iMetadata(maxMetadataCount).Value)
                    End If
                    
                Else
            
                    'This is generic binary and/or list data, which tends to be enormous.  ExifTool is kind enough to always start
                    ' these tag types on a new line - so move down a line, then start searching for the end of the chunk.
                    Do
                        curLine = curLine + 1
                    Loop While InStr(1, xmlLines(curLine), "</et:prt>") = 0
                    
                End If
                
            Else
                xmlLines(curLine) = Trim$(xmlLines(curLine))
                iMetadata(maxMetadataCount).Value = Trim$(Mid$(xmlLines(curLine), 9, InStr(1, xmlLines(curLine), "</et") - 9))
            End If
            
            curLine = curLine + 1
            
            'For some tags, the human-friendly value will be the actual value of the tag.  For others, the tag will have some kind
            ' of obscure value that ExifTool has kindly translated for us.  We still want to retrieve the actual, technical value
            ' if available - so check for it now.
            
            'Note that actual value tags can also come in binary and list varieties, so they are analyzed with an algorithm identical
            ' to the one directly above.  The only difference is an initial check, because a separate "actual value" tag may not exist.
            
            If InStr(1, xmlLines(curLine), "<et:val") > 0 Then
            
                'Simple data will exist on a single line; use that knowledge to separate such values out.
                If (InStr(1, xmlLines(curLine), "<et:val>") > 0) And (InStr(1, xmlLines(curLine), "</et:val>") > 0) Then
                    iMetadata(maxMetadataCount).isActualValueBinary = False
                    iMetadata(maxMetadataCount).isActualValueList = False
                Else
                    'The current value is either binary, or a list.  Check for the presence of an <rdf:Bag> tag to distinguish between the two.
                    If (InStr(1, xmlLines(curLine + 1), "<rdf:Bag>") > 0) Then
                        iMetadata(maxMetadataCount).isActualValueBinary = False
                        iMetadata(maxMetadataCount).isActualValueList = True
                        iMetadata(maxMetadataCount).ActualValue = "(list of values too long to display)"
                    Else
                        iMetadata(maxMetadataCount).isActualValueBinary = True
                        iMetadata(maxMetadataCount).isActualValueList = False
                        iMetadata(maxMetadataCount).ActualValue = "(binary data)"
                    End If
                End If
                                
                'Binary and list-type values can be ignored.  Only retrieve the data if it is neither of these types.
                If iMetadata(maxMetadataCount).isActualValueBinary Or iMetadata(maxMetadataCount).isActualValueList Then
                
                    'Binary and list data can be enormous.  ExifTool is kind enough to always start it on a new line for us -
                    ' so move down a line, then start searching for the end of the binary or list chunk.
                    Do
                        curLine = curLine + 1
                    Loop While InStr(1, xmlLines(curLine), "</et:val>") = 0
                Else
                    xmlLines(curLine) = Trim$(xmlLines(curLine))
                    iMetadata(maxMetadataCount).ActualValue = Trim$(Mid$(xmlLines(curLine), 9, InStr(1, xmlLines(curLine), "</et") - 9))
                End If
                
                curLine = curLine + 1
                
            End If
            
            'Now that we have fully parsed the tag, we need to run a few additional checks.
            
            
            'SPECIAL CHECK FOR MULTILINE COMMENTS
            
            'If the current tag is a multiline comment, it will have been erroneously marked as containing binary data.
            ' Reset it now.
            If (StrComp(iMetadata(maxMetadataCount).Name, "Comment") = 0) And iMetadata(maxMetadataCount).isValueBinary Then
            
                'Remove the binary identifier from this tag
                iMetadata(maxMetadataCount).ActualValue = iMetadata(maxMetadataCount).Value
                iMetadata(maxMetadataCount).isActualValueBinary = False
                iMetadata(maxMetadataCount).isValueBinary = False
            
            End If
            
            
            'SPECIAL CHECK FOR UNEXTRACTED BINARY DATA
            
            'I've gone back and forth on whether to extract binary data or not.  There are pros and cons to extracting it, and
            ' among the big cons are that it adds a significant amount of processing time to RAW images.
             
            'Either way, ExifTool sometimes makes its own decision about what binary data to extract.  When it does this, it
            ' will give a comment similar to "(binary data 12345 bytes, use -b option to extract)".  Check for entries like
            ' this, and remove the "use -b option" instructions.
            If (InStr(1, iMetadata(maxMetadataCount).Value, ", use -b option to extract") > 0) Then
            
                'Remove the binary identifier from this tag
                iMetadata(maxMetadataCount).Value = Replace(iMetadata(maxMetadataCount).Value, ", use -b option to extract", "")
            
            End If
            
            If (InStr(1, iMetadata(maxMetadataCount).ActualValue, ", use -b option to extract") > 0) Then
            
                'Remove the binary identifier from this tag
                iMetadata(maxMetadataCount).ActualValue = Replace(iMetadata(maxMetadataCount).ActualValue, ", use -b option to extract", "")
            
            End If
            
            
            'ANALYZE GROUP
            
            'First, analyze the current Group.  Some groups can be ignored - check for those first.
            tmpString = iMetadata(maxMetadataCount).Group
            If StrComp(tmpString, "ExifTool") = 0 Then ignoreThisTag = True
            
            'Next, ExifTool provides VERY detailed group information.  I prefer to simplify its groups into something more manageable.
            
            'First, there's no need for a distinction between "System" and "File".  Group these under a "General" heading.
            If StrComp(tmpString, "System") = 0 Then
                iMetadata(maxMetadataCount).SubGroup = g_Language.TranslateMessage("System")
                iMetadata(maxMetadataCount).Group = g_Language.TranslateMessage("General")
            End If
            If StrComp(tmpString, "File") = 0 Then
                iMetadata(maxMetadataCount).SubGroup = g_Language.TranslateMessage("File")
                iMetadata(maxMetadataCount).Group = g_Language.TranslateMessage("General")
            End If
            
            'Next, ExifTool breaks down EXIF tags into their individual IFD blocks.  Simplify all those under a single "Exif" header.
            If InStr(1, UCase$(tmpString), "IFD1", vbBinaryCompare) > 0 Then
                iMetadata(maxMetadataCount).SubGroup = iMetadata(maxMetadataCount).Group
                iMetadata(maxMetadataCount).Group = "Exif (thumbnail)"
            ElseIf InStr(1, UCase$(tmpString), "IFD", vbBinaryCompare) > 0 Then
                iMetadata(maxMetadataCount).SubGroup = iMetadata(maxMetadataCount).Group
                iMetadata(maxMetadataCount).Group = "Exif"
            End If
            
            'Next, ExifTool separates color profiles into header, profile, and measurement subcategories.  Condense these into
            ' a single "ICC Profile".
            If InStr(1, tmpString, "ICC", vbBinaryCompare) > 0 Then
                iMetadata(maxMetadataCount).SubGroup = iMetadata(maxMetadataCount).Group
                iMetadata(maxMetadataCount).Group = g_Language.TranslateMessage("ICC Profile")
            End If
            
            'Next, XMP tags often contain an additional identifier, such as "XMP-photoshop".  Condense these into a single XMP group.
            If InStr(1, tmpString, "XMP", vbBinaryCompare) > 0 Then
                iMetadata(maxMetadataCount).SubGroup = iMetadata(maxMetadataCount).Group
                iMetadata(maxMetadataCount).Group = "XMP"
            End If
            
            'Next, one of ExifTool's neatest features is its "Composite" tags - basically, tags that don't actually exist in the file,
            ' but which can be inferred from other information.  Let's add an asterisk to these, to denote that the tags are only
            ' inferred, not physically present.
            If StrComp(tmpString, "Composite") = 0 Then iMetadata(maxMetadataCount).Group = g_Language.TranslateMessage("Composite") & "*"
            
            'ANALYZE NAME
            tmpString = iMetadata(maxMetadataCount).Name
            
            'Check for GPS coordinates.  If found, we can enable some fun bonus metadata options.
            If StrComp(tmpString, "GPSPosition") = 0 Then
                hasGPSCoordinates = True
                gpsIndex = maxMetadataCount
            End If
            
            'ANALYZE VALUE
            
            'Next come .Value checks.  Certain values may be invalid - if so, ignore this tag.
            tmpString = iMetadata(maxMetadataCount).Value
            
            'REMOVE JFIF DUPLICATES
            
            'There is no technical prohibition against duplicate metadata tags - and in fact, some software may choose
            ' to deliberately add duplicate tags in multiple groups (IPTC, EXIF, XMP, etc) to maximize compatibility
            ' with other software packages.
            
            'PhotoDemon allows the presence of multiple tags everywhere except in JFIF headers.  Embedded thumbnails may have
            ' their own JFIF header, which includes information like x/y resolution that is totally irrelevant.  Thus we check
            ' for duplicate JFIF tags, and remove them as necessary.
            If (InStr(1, UCase$(iMetadata(maxMetadataCount).Group), "JFIF", vbBinaryCompare) > 0) And (Not ignoreThisTag) Then
            
                'Check the metadata list up to this point.  If a matching group and technical tag name are found,
                ' move this tag into a separate "thumbnail" category.
                For i = 0 To maxMetadataCount - 1
                    If StrComp(iMetadata(i).FullGroupAndName, iMetadata(maxMetadataCount).FullGroupAndName, vbTextCompare) = 0 Then
                    
                        'Regroup this tag into a new group: JFIF (thumbnail)
                        iMetadata(maxMetadataCount).Group = "JFIF (thumbnail)"
                        Exit For
                    
                    End If
                Next i
            
            End If
            
            'Binary data serves no purpose in the user interface part of our metadata handler.  Remove any binary tags.
            'If InStr(1, UCase$(tmpString), "BINARY DATA", vbBinaryCompare) > 0 Then ignoreThisTag = True
            
            'Increment the max metadata count, and redim our metadata array as necessary
            If Not ignoreThisTag Then
                maxMetadataCount = maxMetadataCount + 1
                ReDim Preserve iMetadata(0 To maxMetadataCount) As mdItem
            End If
        
        'If this line doesn't contain a tag, move to the next one
        Else
                
            'Move to the next line
            curLine = curLine + 1
                
        End If
        
    Loop While curLine < UBound(xmlLines)
    
    
    'If present, metadata has now been loaded and successfully parsed.
    
    'PD supports handling of some esoteric metadata elements, which can affect the RGB/A data of the image.  Look for any such tags now,
    ' and if found, apply any changes immediately.
    If maxMetadataCount > 0 Then
    
    'Also, apply some failsafe checks for pdImage boundaries.  I would prefer to include all this on one line, but VB doesn't short-circuit,
    ' so I have no choice but to list them like this.
    If (parentImageID <= UBound(pdImages)) Then
    If Not pdImages(parentImageID) Is Nothing Then
    
        Dim tagIndex As Long
        
        'First up are some PNG chunks that require special handling.  The PNG spec (http://www.libpng.org/pub/png/spec/1.2/png-1.2-pdg.html)
        ' provides instructions on which color modification chunks should be given preference, and we mirror that behavior here.  PD's order
        ' of operations for PNG chunk processing is: ICC, sRGB, cHRM, gAMA.  ICC handling will happen separately, away from the load process,
        ' and if an ICC profile was found, this whole section can be skipped.
        
        If (pdImages(parentImageID).originalFileFormat = FIF_PNG) And (Not pdImages(parentImageID).getActiveDIB.ICCProfile.hasICCData) Then
        
            Dim srcGamma As Double
            srcGamma = 0
            
            Dim cHRMHandled As Boolean, gAMAHandled As Boolean, sRGBHandled As Boolean
            cHRMHandled = False: gAMAHandled = False: sRGBHandled = False
            
            'First, we check for an sRGB chunk.  Technically, we should use the transform requested by the chunk value, but I don't have
            ' source code for all possible transform modes, so we're just going to do the default one (absolute colorimetric).
            If doesTagExistFullName("PNG:SRGBRendering", tagIndex, vbBinaryCompare) Then
            
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "PNG file contains sRGB chunk. Applying color conversion now..."
                #End If
                
                Color_Management.convertRGBUsingCustomEndpoints pdImages(parentImageID).getActiveDIB, 0.64, 0.33, 0.3, 0.6, 0.15, 0.06, 0.3127, 0.329, 0, False, True
                
                sRGBHandled = True
            
            End If
            
            'Next, we're going to look for the "PNG:WhitePointX/Y" pair of tags.  These are found in the cHRM chunk of the
            ' PNG file, and they specify a custom white point to be used as part of color management.  These days, cHRM has largely
            ' been superceded by use of ICC profiles and/or sRGB identifiers, but cHRM is much more compact and may still be used by
            ' websites obsessed with file size.  (The chunk also provides a nice way to test our custom gamut code, which is why I
            ' implemented this feature in the first place!)
            If (Not sRGBHandled) And doesTagExistFullName("PNG:WhitePointX", tagIndex, vbBinaryCompare) And doesTagExistFullName("PNG:WhitePointY", tagIndex, vbBinaryCompare) Then
            
                'Those two tags were found.  Use of an sRGB tag or ICC profile supercedes the use of specific cHRM values, so check for
                ' those next.  If either are found, ignore further processing.
                If (Not doesTagExistFullName("PNG:ProfileName", tagIndex, vbBinaryCompare)) And (Not doesTagExistFullName("PNG:SRGBRendering", tagIndex, vbBinaryCompare)) Then
                
                    'White point values alone are not enough to properly handle color conversion.  We also need RGB tags.
                    If doesTagExistFullName("PNG:RedX", tagIndex, vbBinaryCompare) And doesTagExistFullName("PNG:RedY", tagIndex, vbBinaryCompare) And doesTagExistFullName("PNG:GreenX", tagIndex, vbBinaryCompare) And doesTagExistFullName("PNG:GreenY", tagIndex, vbBinaryCompare) And doesTagExistFullName("PNG:BlueX", tagIndex, vbBinaryCompare) And doesTagExistFullName("PNG:BlueY", tagIndex, vbBinaryCompare) Then
                    
                        'All tags exist.  Retrieve their values now.
                        Dim WhiteX As Double, WhiteY As Double, RedX As Double, RedY As Double, GreenX As Double, GreenY As Double, BlueX As Double, BlueY As Double
                        WhiteX = CDblCustom(getTagValue("PNG:WhitePointX", vbBinaryCompare))
                        WhiteY = CDblCustom(getTagValue("PNG:WhitePointY", vbBinaryCompare))
                        RedX = CDblCustom(getTagValue("PNG:RedX", vbBinaryCompare))
                        RedY = CDblCustom(getTagValue("PNG:RedY", vbBinaryCompare))
                        GreenX = CDblCustom(getTagValue("PNG:GreenX", vbBinaryCompare))
                        GreenY = CDblCustom(getTagValue("PNG:GreenY", vbBinaryCompare))
                        BlueX = CDblCustom(getTagValue("PNG:BlueX", vbBinaryCompare))
                        BlueY = CDblCustom(getTagValue("PNG:BlueY", vbBinaryCompare))
                        
                        'Check for matching inputs, including zeroes.  These indicate a writer error.
                        If ((WhiteX = 0) And (WhiteY = 0)) Or ((RedX = 0) And (RedY = 0)) Or ((GreenX = 0) And (GreenY = 0)) Or ((BlueX = 0) And (BlueY = 0)) Then
                            
                            #If DEBUGMODE = 1 Then
                                pdDebug.LogAction "FYI: custom cHRM data appears to be invalid; ignoring custom conversion.  (List of values follows, in ExifTool order.)"
                                pdDebug.LogAction WhiteX & ", " & WhiteY & ", " & RedX & ", " & RedY & ", " & GreenX & ", " & GreenY & ", " & BlueX & ", " & BlueY
                            #End If
                            
                        'End points appear to be valid!  Continue with the transform.
                        Else
                        
                            'Per the spec, files using the cHRM chunk are also supposed to provide a gamma value.  Without it, we have to
                            ' guess at the right gamma value, which defeats the whole purpose of providing color correction data.
                            ' Look for gamma now.
                            If doesTagExistFullName("PNG:Gamma", tagIndex, vbBinaryCompare) Then srcGamma = CDblCustom(getTagValue("PNG:Gamma", vbBinaryCompare))
                            
                            'If gamma does not exist in the file, or it does exist but it is encoded incorrectly, we'll still proceed with
                            ' the conversion, but we have no choice but to assume a default gamma value.
                            If srcGamma = 0 Then
                            
                                #If DEBUGMODE = 1 Then
                                    pdDebug.LogAction "WARNING! PNG file has cHRM but not gAMA. Default sRGB gamma transform will be used."
                                #End If
                            
                            End If
                            
                            #If DEBUGMODE = 1 Then
                                pdDebug.LogAction "PNG file with custom cHRM data found. Applying color conversion now..."
                            #End If
                            
                            'We're now going to do something weird.  If the reported cHRM endpoints match sRGB endpoints, we're going to pass 0 to the
                            ' gamut conversion function.  This enables the sRGB gamma function, which is a rather complicated two-part ramp (instead of
                            ' a standard gamma function).  It also produces results identical to Firefox for libpng's test image library.
                            '
                            'If the cHRM chunk is any other color space, we'll use the supplied gamma value, per the PNG spec
                            If (RedX = 0.64) And (RedY = 0.33) And (GreenX = 0.3) And (GreenY = 0.6) And (BlueX = 0.15) And (BlueY = 0.06) And (WhiteX = 0.3127) And (WhiteY = 0.329) Then
                                Color_Management.convertRGBUsingCustomEndpoints pdImages(parentImageID).getActiveDIB, RedX, RedY, GreenX, GreenY, BlueX, BlueY, WhiteX, WhiteY, 0, False, True
                            Else
                                Color_Management.convertRGBUsingCustomEndpoints pdImages(parentImageID).getActiveDIB, RedX, RedY, GreenX, GreenY, BlueX, BlueY, WhiteX, WhiteY, srcGamma, False, True
                            End If
                            
                            'FOR TESTING PURPOSES: this sample line of code uses hard-coded sRGB values
                            'Color_Management.convertRGBUsingCustomEndpoints pdImages(parentImageID).getActiveDIB, 0.64, 0.33, 0.3, 0.6, 0.15, 0.06, 0.3127, 0.329, 0, False, True
                            
                            'FOR TESTING PURPOSES: this sample line of code uses had-coded Adobe RGB values
                            'Color_Management.convertRGBUsingCustomEndpoints pdImages(parentImageID).getActiveDIB, 0.64, 0.33, 0.21, 0.71, 0.15, 0.06, 0.3127, 0.329, 2.2, False, True
                            
                            'Note that cHRM has been handled; this prevents subsequent gAMA checks from applying a second round of
                            ' corrections to the image.
                            cHRMHandled = True
                            
                            #If DEBUGMODE = 1 Then
                                pdDebug.LogAction "Custom cHRM data analysis complete."
                            #End If
                            
                            'Refresh the screen to reflect the new colors
                            If parentImageID = g_CurrentImage Then
                                Viewport_Engine.Stage1_InitializeBuffer pdImages(parentImageID), FormMain.mainCanvas(0)
                                toolbar_ImageTabs.notifyUpdatedImage parentImageID
                            End If
                            
                        End If
                    
                    Else
                        #If DEBUGMODE = 1 Then
                            pdDebug.LogAction "FYI: custom cHRM data has been ignored because not all endpoint values exist!"
                        #End If
                    End If
                
                Else
                    If sRGBHandled Then
                        #If DEBUGMODE = 1 Then
                            pdDebug.LogAction "FYI: custom cHRM data has been ignored in favor of sRGB data."
                        #End If
                    End If
                End If
            
            'End cHRM chunk handling for PNG files
            End If
            
            'The next tag we're going to look for is the "PNG:Gamma" value.  FreeImage provides an option for handling this value,
            ' but they do not handle it correctly (by "correctly" I mean, "like web browsers, for better or worse").  Handling gamma
            ' here also allows us to provide gamma correction when GDI+ is used.  Note that this behavior should probably be tied to
            ' a preference, because gamma handling may screw up some images that were encoded incorrectly.
            If (Not cHRMHandled) And (Not sRGBHandled) And doesTagExistFullName("PNG:Gamma", tagIndex, vbBinaryCompare) Then
            
                srcGamma = CDblCustom(getTagValue("PNG:Gamma", vbBinaryCompare))
                
                'If gamma does not exist in the file, or it does exist but it is encoded incorrectly, we obviously won't make use of it.
                If srcGamma <= 0 Then
                
                    #If DEBUGMODE = 1 Then
                        pdDebug.LogAction "WARNING! PNG file has gAMA, but the reported value is invalid.  No gamma correction applied."
                    #End If
                
                'If gamma does exist, and it is encoded correctly (as far as we can tell), apply it now
                Else
                
                    #If DEBUGMODE = 1 Then
                        pdDebug.LogAction "PNG file with custom gAMA data found. Applying gamma correction now..."
                    #End If
                    
                    Filters_Layers.GammaCorrectDIB pdImages(parentImageID).getActiveDIB, 2.2 / srcGamma, True
                    
                    #If DEBUGMODE = 1 Then
                        pdDebug.LogAction "Custom cHRM data analysis complete."
                    #End If
                    
                    'Note that gAMA has been handled; this prevents subsequent checks from possibly applying a second round of
                    ' corrections to the image.
                    gAMAHandled = True
                    
                    'Refresh the screen to reflect the new colors
                    If parentImageID = g_CurrentImage Then
                        Viewport_Engine.Stage1_InitializeBuffer pdImages(parentImageID), FormMain.mainCanvas(0)
                        toolbar_ImageTabs.notifyUpdatedImage parentImageID
                    End If
                    
                End If
            
            'END gAMA chunk handling
            End If
            
        'END special PNG format handling
        End If

        
    'End special metadata analysis and usage
    End If
    End If
    End If
    
    
    'If no metadata was found, return a FAIL state.  (ExifTool should always return *some* metadata, e.g. file date/time and generic
    ' stuff like that).  Otherwise, if at least one piece of valid metadata was found, return SUCCESS.
    If maxMetadataCount = 0 Then
        loadAllMetadata = False
    Else
        loadAllMetadata = True
    End If
    
End Function

'Given a path to a saved image (yes, that's right - make certain it is already saved!) and a Long-type value specifying what metadata to save,
' embed all relevant metadata.  (Note: this needs to be done immediately following a save, and is generally handled by the
' PhotoDemon_SaveImage function.)
Public Function writeAllMetadata(ByVal outputPath As String, ByRef srcPDImage As pdImage) As Boolean
    
    'Metadata only needs to be written out if it exists (obviously)
    If hasXMLMetadata Then
        
        'XML writing is handled by PD's master XML engine
        Dim xmlEngine As pdXML
        Set xmlEngine = New pdXML
        
        Dim newXMLString As String, tmpXMLPath As String
        Dim fileNum As Integer
        
        Dim cSHA2 As CSHA256
        Set cSHA2 = New CSHA256
        
        'The user has three choices for metadata handling, settable from the Tools -> Preferences -> Saving panel.
        ' (Note: some export dialogs (e.g. JPEG) also allow the user to override the global metadata preference for a given image.)
        Select Case metadataExportPreference
            
            'Save all relevant metadata.  Previously, we provided an option for "save all metadata, regardless of relevance",
            ' but that has since been removed (for a variety of reasons).  If someone previously specified that option, we
            ' treat it the same as "save all relevant metadata".
            
            'Save as much relevant metadata as possible (default action)
            Case 0, 1
            
                Message "Embedding metadata..."
                
                'In the future, we will do some preprocessing to the XML string to remove invalid entries, etc.  But for now,
                ' simply dump the full original XML contents to a temp file.
                newXMLString = originalMetadataXML
                
                ' Since we can assume the user has rights to the output directory, write the temporary XML file there
                tmpXMLPath = g_UserPreferences.GetTempPath
                
                'We'll also append a partial hash to the temporary XML filename to avoid overwriting any existing XML files in the folder
                tmpXMLPath = tmpXMLPath & "metadata_" & Left$(cSHA2.SHA256(outputPath), 16) & ".xml"
                
                'Write the XML file
                xmlEngine.loadXMLFromString newXMLString
                xmlEngine.writeXMLToFile tmpXMLPath, True
                                
                'The next part is easy - supply the XML file to ExifTool, and request of copying of metadata from the XML to the image file.
                If writeMetadata(tmpXMLPath, outputPath, srcPDImage) Then
                    Message "Metadata embedded successfully."
                    
                Else
                    Message "Metadata may not have been embedded (but saved image still written successfully)."
                End If
                
            'Save metadata, but remove items with potential privacy concerns (GPS coordinates, serial numbers, etc)
            Case 2
            
                Message "Removing metadata with potential privacy concerns..."
            
                'First thing we need to do is mark all tags to be removed.  Note that GPS tags will be automatically removed by
                ' a parameter we pass to ExifTool, but any other privacy tags must be removed manually.
                Dim i As Long
                Dim removeThisTag As Boolean
                Dim sMetadataName As String
                For i = 0 To maxMetadataCount - 1
                
                    removeThisTag = False
                
                    'Compare this tag against a manually curated database of potentially private tags.  If a match is found,
                    ' mark this tag for removal.

                    sMetadataName = UCase$(iMetadata(i).Name)
        
                    If InStr(1, sMetadataName, "FIRMWARE", vbBinaryCompare) > 0 Then
                        removeThisTag = True
                    ElseIf InStr(1, sMetadataName, "SERIAL", vbBinaryCompare) > 0 Then
                        removeThisTag = True
                    ElseIf InStr(1, sMetadataName, "NAME", vbBinaryCompare) > 0 Then
                        removeThisTag = True
                    ElseIf InStr(1, sMetadataName, "BABY", vbBinaryCompare) > 0 Then
                        removeThisTag = True
                    ElseIf InStr(1, sMetadataName, "SOFTWARE", vbBinaryCompare) > 0 Then
                        removeThisTag = True
                    ElseIf InStr(1, sMetadataName, "ABOUT", vbBinaryCompare) > 0 Then
                        removeThisTag = True
                    ElseIf InStr(1, sMetadataName, "CREATOR", vbBinaryCompare) > 0 Then
                        removeThisTag = True
                    ElseIf InStr(1, sMetadataName, "COUNTRY", vbBinaryCompare) > 0 Then
                        removeThisTag = True
                    ElseIf InStr(1, sMetadataName, "CITY", vbBinaryCompare) > 0 Then
                        removeThisTag = True
                    ElseIf InStr(1, sMetadataName, "LOCATION", vbBinaryCompare) > 0 Then
                        removeThisTag = True
                    ElseIf InStr(1, sMetadataName, "BY-LINE", vbBinaryCompare) > 0 Then
                        removeThisTag = True
                    ElseIf InStr(1, sMetadataName, "TOOLKIT", vbBinaryCompare) > 0 Then
                        removeThisTag = True
                    ElseIf InStr(1, sMetadataName, "HISTORY", vbBinaryCompare) > 0 Then
                        removeThisTag = True
                    ElseIf InStr(1, UCase$(iMetadata(i).Value), "XMP.IID", vbBinaryCompare) > 0 Then
                        removeThisTag = True
                    ElseIf InStr(1, UCase$(iMetadata(i).Value), "XMP.DID", vbBinaryCompare) > 0 Then
                        removeThisTag = True
                        
                    End If
                    
                    iMetadata(i).markedForRemoval = removeThisTag
                    
                Next i
                
                'Now that all problematic tags have been identified, we need to generate a modified XML string without those entries.
                newXMLString = retrieveModifiedXMLString
                
                Message "Embedding metadata..."
                
                ' Since we can assume the user has rights to the output directory, write the temporary XML file there
                tmpXMLPath = getDirectory(outputPath)
                
                'We'll also append a partial hash to the temporary XML filename to avoid overwriting any existing XML files in the folder
                tmpXMLPath = tmpXMLPath & "metadata_" & Left$(cSHA2.SHA256(outputPath), 16) & ".xml"
                
                'Write the XML file
                xmlEngine.loadXMLFromString newXMLString
                xmlEngine.writeXMLToFile tmpXMLPath, True
                
                'The next part is easy - supply the XML file to ExifTool, and request of copying of metadata from the XML to the image file.
                If writeMetadata(tmpXMLPath, outputPath, srcPDImage, True) Then
                    Message "Metadata embedded successfully."
                Else
                    Message "Metadata may not have been embedded (but saved image still written successfully)."
                End If
                            
            'Do not save metadata.  This is by far the easiest method.  :)
            Case 3
        
        End Select
        
    End If
    
    writeAllMetadata = True
    
End Function

Private Sub Class_Initialize()
    maxMetadataCount = 0
    ReDim iMetadata(0) As mdItem
End Sub
