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
'Copyright 2013-2026 by Tanner Helland
'Created: 23/May/13
'Last updated: 09/June/16
'Last update: improve parsing behavior against incorrectly formed multiline print values that are non-binary in nature
'
'As of version 6.0, PhotoDemon now provides support for loading and saving image metadata.  What is metadata, you ask?
' See https://en.wikipedia.org/wiki/Metadata#Photographs for more details.
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
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'The number of metadata items currently stored
Private m_NumOfTags As Long

'Array that contains the metadata itself
Private m_Tags() As PDMetadataItem

'When extracting metadata, we request a full XML dump from ExifTool.  This XML file is stored - in its entirety - in memory,
' as it is the only way to precisely restore the XML data in its *exact* form when save-time comes.
Private m_OriginalXML As String

'We cache whether or not GPS coords are present; this simplifies UI synching of geotagging controls
Private m_HasGPSCoordinates As Boolean

Private Const LEFT_TAG_TEST As String = " <"   'The preceding space is crucial; do not remove it!
Private Const RIGHT_TAG_TEST As String = ">"

'Return the original metadata XML string.  Note that aside from some minimal processing (e.g. removing ExifTool's {ready} marker),
' this will be the unmodified XML string originally passed to this class, e.g. without any parsing or validation of the data!
Friend Function GetOriginalXMLMetadataString() As String
    GetOriginalXMLMetadataString = m_OriginalXML
End Function

Friend Sub GetOriginalXMLMetadataStrPtrAndLen(ByRef dstPtr As Long, ByRef dstLen As Long)
    dstPtr = StrPtr(m_OriginalXML)
    dstLen = Len(m_OriginalXML)
End Sub

'Return a current copy of the metadata collection, serialized into a single massive string.  This string contains all user edits,
' and generally speaking, it should be used in parallel to the original XML string, above.  (Both are needed to fully recreate
' metadata inside a new image.)
Friend Function GetSerializedXMLData() As String
    
    If (m_NumOfTags > 0) Then
        
        'We're actually going to use two serializers - one for the full metadata collection,
        ' and a second one for individual tags (which have *huge* lists of properties).
        Dim mainSerializer As pdSerialize
        Set mainSerializer = New pdSerialize
        mainSerializer.Reset 2#
        mainSerializer.AddParam "tag-count", m_NumOfTags, True, True
        
        Dim tagSerializer As pdSerialize
        Set tagSerializer = New pdSerialize
        
        Dim i As Long
        For i = 0 To m_NumOfTags - 1
            
            'Serialize each tag in turn, then add the finished result to the main serializer
            tagSerializer.Reset 2#
            
            'It's important to note that we need to be careful here with flagging items as "XML-safe".
            ' (This is an optional shortcut parameter to .AddParam() that tells it the passed value
            '  does *not* need to be manually delimited.)  .AddParam() is smart enough to know that
            ' numeric values are XML-safe, but strings may or may not be depending on the presence of
            ' invalid XML chars ("<", ">", "&").  Flagging entries as "XML-safe" improves performance,
            ' but it will break subsequent parsing if used incorrectly - so be wary!
            With m_Tags(i)
                tagSerializer.AddParam "group-and-name", .TagGroupAndName, True
                tagSerializer.AddParam "group", .TagGroup, True
                tagSerializer.AddParam "group-readable", .TagGroupFriendly, True
                tagSerializer.AddParam "name", .tagName, True
                tagSerializer.AddParam "name-readable", .TagNameFriendly, True
                tagSerializer.AddParam "table", .TagTable, True
                tagSerializer.AddParam "id", .TagID, True
                tagSerializer.AddParam "value-readable", .TagValueFriendly, True
                tagSerializer.AddParam "value", .TagValue, True
                tagSerializer.AddParam "has-index", .HasIndex, True
                tagSerializer.AddParam "is-list", .IsTagList, True
                tagSerializer.AddParam "is-binary", .IsTagBinary, True
                tagSerializer.AddParam "binary-was-extracted", .WasBinaryExtracted, True
                tagSerializer.AddParam "base64", .TagBase64Value, True, True
                tagSerializer.AddParam "internal-use", .InternalUseOnly, True
                tagSerializer.AddParam "internal-index", .TagIndexInternal, True
                tagSerializer.AddParam "remove", .TagMarkedForRemoval, True
                'Note that there is another internal "marked for removal" flag here - we do *not* save it, however.
                ' (See comments in parameter description in the ExifTool module for details.)
                tagSerializer.AddParam "modified-now", .UserModifiedThisSession, True
                tagSerializer.AddParam "modified-any", .UserModifiedAllSessions, True
                tagSerializer.AddParam "user-value", .UserValueNew, True
                tagSerializer.AddParam "user-id", .UserIDNew, True
                tagSerializer.AddParam "hit-database", .DB_TagHitDatabase, True
                tagSerializer.AddParam "is-writable", .DB_IsWritable, True
                tagSerializer.AddParam "type-count", .DB_TypeCount, True
                tagSerializer.AddParam "data-type", .DB_DataType, True
                tagSerializer.AddParam "data-type-strict", .DB_DataTypeStrict, True
                tagSerializer.AddParam "is-avoid", .DBF_IsAvoid, True
                tagSerializer.AddParam "is-bag", .DBF_IsBag, True
                tagSerializer.AddParam "is-binary", .DBF_IsBinary, True
                tagSerializer.AddParam "is-flattened", .DBF_IsFlattened, True
                tagSerializer.AddParam "is-list", .DBF_IsList, True
                tagSerializer.AddParam "is-mandatory", .DBF_IsMandatory, True
                tagSerializer.AddParam "is-permanent", .DBF_IsPermanent, True
                tagSerializer.AddParam "is-protected", .DBF_IsProtected, True
                tagSerializer.AddParam "is-sequence", .DBF_IsSequence, True
                tagSerializer.AddParam "is-unknown", .DBF_IsUnknown, True
                tagSerializer.AddParam "is-unsafe", .DBF_IsUnsafe, True
                tagSerializer.AddParam "description", .DB_Description, True
                tagSerializer.AddParam "hardcoded-list", .DB_HardcodedList, True
                tagSerializer.AddParam "hardcoded-list-size", .DB_HardcodedListSize, True
                If .DB_HardcodedList And (.DB_HardcodedListSize > 0) Then
                    If (Not .DB_StackIDs Is Nothing) Then tagSerializer.AddParam "stack-ids", .DB_StackIDs.SerializeStackToSingleString(), True
                    If (Not .DB_StackValues Is Nothing) Then tagSerializer.AddParam "stack-values", .DB_StackValues.SerializeStackToSingleString(), True
                End If
                'During debugging, it may be helpful to cache the pure data returned from ExifTool; uncomment to allow this
                'tagSerializer.AddParam "debug-data", .TagDebugData, True
            End With
            
            'Add the finished XML string to the main serializer
            mainSerializer.AddParam "tag-" & CStr(i), tagSerializer.GetParamString(), True, False
            
        Next i
        
    End If
    
    GetSerializedXMLData = mainSerializer.GetParamString()
    
End Function

Friend Sub RecreateFromSerializedXMLData(ByRef srcString As String)
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString srcString
    
    Dim i As Long
    
    'Check for legacy encoding
    If (cParams.GetParamVersion() >= 2#) Then
        
        'Make sure there's a non-zero amount of tags to decode
        m_NumOfTags = cParams.GetLong("tag-count", 0, True)
        If (m_NumOfTags < 0) Then m_NumOfTags = 0
        ReDim m_Tags(0 To m_NumOfTags) As PDMetadataItem
        
        If (m_NumOfTags > 0) Then
            
            Dim tagSerializer As pdSerialize
            Set tagSerializer = New pdSerialize
            
            For i = 0 To m_NumOfTags - 1
                
                'Serialize each tag in turn, then add the finished result to the main serializer
                tagSerializer.Reset
                tagSerializer.SetParamString cParams.GetString("tag-" & CStr(i), vbNullString, True)
                
                'It's important to note that we need to be careful here with flagging items as "XML-safe".
                ' (This is an optional shortcut parameter to .AddParam() that tells it the passed value
                '  does *not* need to be manually delimited.)  .AddParam() is smart enough to know that
                ' numeric values are XML-safe, but strings may or may not be depending on the presence of
                ' invalid XML chars ("<", ">", "&").  Flagging entries as "XML-safe" improves performance,
                ' but it will break subsequent parsing if used incorrectly - so be wary!
                With m_Tags(i)
                    .TagGroupAndName = tagSerializer.GetString("group-and-name", , True)
                    .TagGroup = tagSerializer.GetString("group", , True)
                    .TagGroupFriendly = tagSerializer.GetString("group-readable", , True)
                    .tagName = tagSerializer.GetString("name", , True)
                    .TagNameFriendly = tagSerializer.GetString("name-readable", , True)
                    .TagTable = tagSerializer.GetString("table", , True)
                    .TagID = tagSerializer.GetString("id", , True)
                    .TagValueFriendly = tagSerializer.GetString("value-readable", , True)
                    .TagValue = tagSerializer.GetString("value", , True)
                    .HasIndex = tagSerializer.GetBool("has-index", , True)
                    .IsTagList = tagSerializer.GetBool("is-list", , True)
                    .IsTagBinary = tagSerializer.GetBool("is-binary", , True)
                    .WasBinaryExtracted = tagSerializer.GetBool("binary-was-extracted", , True)
                    .TagBase64Value = tagSerializer.GetString("base64", , True)
                    .InternalUseOnly = tagSerializer.GetBool("internal-use", , True)
                    .TagIndexInternal = tagSerializer.GetLong("internal-index", , True)
                    .TagMarkedForRemoval = tagSerializer.GetBool("remove", , True)
                    .UserModifiedThisSession = tagSerializer.GetBool("modified-now", , True)
                    .UserModifiedAllSessions = tagSerializer.GetBool("modified-any", , True)
                    .UserValueNew = tagSerializer.GetString("user-value", , True)
                    .UserIDNew = tagSerializer.GetString("user-id", , True)
                    .DB_TagHitDatabase = tagSerializer.GetBool("hit-database")
                    .DB_IsWritable = tagSerializer.GetBool("is-writable", , True)
                    .DB_TypeCount = tagSerializer.GetLong("type-count", , True)
                    .DB_DataType = tagSerializer.GetString("data-type", , True)
                    .DB_DataTypeStrict = tagSerializer.GetLong("data-type-strict", , True)
                    .DBF_IsAvoid = tagSerializer.GetBool("is-avoid", , True)
                    .DBF_IsBag = tagSerializer.GetBool("is-bag", , True)
                    .DBF_IsBinary = tagSerializer.GetBool("is-binary", , True)
                    .DBF_IsFlattened = tagSerializer.GetBool("is-flattened", , True)
                    .DBF_IsList = tagSerializer.GetBool("is-list", , True)
                    .DBF_IsMandatory = tagSerializer.GetBool("is-mandatory", , True)
                    .DBF_IsPermanent = tagSerializer.GetBool("is-permanent", , True)
                    .DBF_IsProtected = tagSerializer.GetBool("is-protected", , True)
                    .DBF_IsSequence = tagSerializer.GetBool("is-sequence", , True)
                    .DBF_IsUnknown = tagSerializer.GetBool("is-unknown", , True)
                    .DBF_IsUnsafe = tagSerializer.GetBool("is-unsafe", , True)
                    .DB_Description = tagSerializer.GetString("description", , True)
                    .DB_HardcodedList = tagSerializer.GetBool("hardcoded-list", , True)
                    .DB_HardcodedListSize = tagSerializer.GetLong("hardcoded-list-size", , True)
                    If .DB_HardcodedList And (.DB_HardcodedListSize > 0) Then
                        Set .DB_StackIDs = New pdStringStack
                        Set .DB_StackValues = New pdStringStack
                        .DB_StackIDs.RecreateStackFromSerializedString tagSerializer.GetString("stack-ids", , True)
                        .DB_StackValues.RecreateStackFromSerializedString tagSerializer.GetString("stack-values", , True)
                    End If
                    
                    'During debugging, it may be helpful to cache the pure data returned from ExifTool; uncomment to allow this
                    '.TagDebugData = tagSerializer.GetString("debug-data", , True)
                    
                End With
                
            Next i
            
        End If
    
    'Legacy block
    Else
    
        If (cParams.GetLong("PDMetadataTagCount", 0) > 0) Then
        
            m_NumOfTags = cParams.GetLong("PDMetadataTagCount", 0)
            ReDim m_Tags(0 To m_NumOfTags) As PDMetadataItem
            
            For i = 0 To m_NumOfTags - 1
                ExifTool.RecoverSerializedTag_Legacy cParams.GetString("PDMD_TAG_NO_" & CStr(i)), m_Tags(i)
            Next i
            
        End If
        
    End If
    
End Sub

'Retrieve resolution data from the metadata collection, if it exists.  Returns TRUE if resolution metadata entrist exist; FALSE
' if they do not.  Note that this function always returns resolution in inches.  Pixels per CM measurements will be automatically
' converted to PPI.
Friend Function GetResolution(ByRef xResolution As Double, ByRef yResolution As Double) As Boolean

    On Error GoTo mdResolutionError

    Dim xResTagIndex As Long, yResTagIndex As Long
    Dim unitIsCentimeters As Boolean, resUnitIndex As Long
    
    'Resolution can appear in quite a few different places.  To try and get the user the value they expect,
    ' we have to parse those places in a particular order.
    Dim resFound As Boolean
    
    'First, look for a Photoshop-specific resolution tag.  I don't like doing this, because it reinforces the
    ' notion that Adobe gets to do whatever the hell it wants and others will just "play along", but because
    ' Photoshop-exported JPEGs can often have 3+ different resolution values in 3+ different locations, we
    ' default to the internal Photoshop resolution value, if one exists.  (See this GitHub issue for additional
    ' details/discussion: https://github.com/tannerhelland/PhotoDemon/issues/260 )
    If DoesTagExistFullName("Photoshop:XResolution", xResTagIndex) And DoesTagExistFullName("Photoshop:YResolution", yResTagIndex) Then
        
        'Photoshop resolution tags exist.  To my knowledge, PS resolution values are always stored in inches.
        ' (There is a "displayed units" tag, but even if this is set to cm, the stored resolution values
        '  should still be in inches.)
        resFound = True
        xResolution = TextSupport.CDblCustom(m_Tags(xResTagIndex).TagValue)
        yResolution = TextSupport.CDblCustom(m_Tags(yResTagIndex).TagValue)
        
    End If
    
    'If Photoshop-specific tags weren't found, query EXIF-specific resolution values next
    If (Not resFound) Then
    
        If DoesTagExistFullName("IFD0:XResolution", xResTagIndex) And DoesTagExistFullName("IFD0:YResolution", yResTagIndex) Then
        
            'EXIF resolution tags exist.
            resFound = True
            xResolution = TextSupport.CDblCustom(m_Tags(xResTagIndex).TagValue)
            yResolution = TextSupport.CDblCustom(m_Tags(yResTagIndex).TagValue)
            
            'EXIF resolution data can be in CM or INCHES.  Conver them as necessary.
            unitIsCentimeters = False
            If DoesTagExistFullName("IFD0:ResolutionUnit", resUnitIndex) Then
                unitIsCentimeters = Strings.StringsNotEqual(m_Tags(resUnitIndex).TagValueFriendly, "inches", True) And Strings.StringsNotEqual(m_Tags(resUnitIndex).TagValueFriendly, "in", True)
            End If
            
            If unitIsCentimeters Then
                xResolution = Units.GetInchesFromCM(xResolution)
                yResolution = Units.GetInchesFromCM(yResolution)
            End If
            
        End If
    
    End If
    
    'If neither Photoshop nor EXIF tags were found, look for generic x/y resolution data last.  We don't care
    ' what group these come from - it'll vary by format (e.g. JFIF for JPEGs, pHYs for PNGs, etc)
    If (Not resFound) Then
    
        If DoesTagExist("XResolution", xResTagIndex) And DoesTagExist("YResolution", yResTagIndex) Then
        
            'Check for the presence of a resolution type.  (If no unit is specified, assume inches.)
            unitIsCentimeters = False
            If DoesTagExist("ResolutionUnit", resUnitIndex) Then
                unitIsCentimeters = Strings.StringsNotEqual(m_Tags(resUnitIndex).TagValueFriendly, "inches", True) And Strings.StringsNotEqual(m_Tags(resUnitIndex).TagValueFriendly, "in", True)
            End If
            
            'Retrieve X and Y resolution values
            xResolution = TextSupport.CDblCustom(m_Tags(xResTagIndex).TagValue)
            yResolution = TextSupport.CDblCustom(m_Tags(yResTagIndex).TagValue)
            
            'If the units are in centimeters, convert them to inches
            If unitIsCentimeters Then
                xResolution = GetInchesFromCM(xResolution)
                yResolution = GetInchesFromCM(yResolution)
            End If
            
            resFound = True
            
        'Resolution tags don't exist.
        End If
        
    End If
    
    'If any valid resolution data was found, return it.  (Source group doesn't matter.)
    GetResolution = resFound
    
    Exit Function
    
mdResolutionError:
    GetResolution = False

End Function

'Does a given tag exist in the metadata collection?  Returns TRUE if yes, FALSE, if no.
' If the tag does exist, the optional byref tagIndex will be set to the index.  It will be set to -1 if the tag is not found.
Private Function DoesTagExist(ByVal tagName As String, Optional ByRef tagIndex As Long = -1, Optional ByVal searchMode As VbCompareMethod = vbTextCompare) As Boolean

    DoesTagExist = False
    
    'If this image has no metadata, exit now
    If (m_NumOfTags <= 0) Then Exit Function

    'Scan the metadata collection, looking for the given tag name
    Dim i As Long
    For i = 0 To m_NumOfTags - 1
        If Strings.StringsEqual(m_Tags(i).tagName, tagName, (searchMode = vbTextCompare)) Then
            DoesTagExist = True
            tagIndex = i
            Exit For
        End If
    Next i

End Function

'Does a given tag exist in the metadata collection?  Returns TRUE if yes, FALSE, if no.
' If the tag does exist, the optional byref tagIndex will be set to the index.  It will be set to -1 if the tag is not found.
Friend Function DoesTagExistFullName(ByVal tagNameAndGroup As String, Optional ByRef tagIndex As Long = -1, Optional ByVal searchMode As VbCompareMethod = vbTextCompare) As Boolean

    DoesTagExistFullName = False
    
    'If this image has no metadata, exit now
    If (m_NumOfTags <= 0) Then Exit Function
    
    'Scan the metadata collection, looking for the given tag name
    Dim i As Long
    For i = 0 To m_NumOfTags - 1
    
        'Tag name found!  Mark the index and exit.
        If Strings.StringsEqual(m_Tags(i).TagGroupAndName, tagNameAndGroup, (searchMode = vbTextCompare)) Then
            DoesTagExistFullName = True
            tagIndex = i
            Exit For
        End If
        
    Next i

End Function

'Return the human-friendly value of a given tag.  If the tag is not found, a null string is returned.
'
'(It's assumed that the caller has called DoesTagExistFullName() before calling this function, so they know whether a null string is
' the actual return value of the given metadata entry, or a fail state.)
Friend Function GetTagValue(ByVal tagNameAndGroup As String, Optional ByVal searchMode As VbCompareMethod = vbTextCompare, Optional ByVal returnPrintFriendlyValue As Boolean = False) As String

    Dim tagIndex As Long
    
    If DoesTagExistFullName(tagNameAndGroup, tagIndex, searchMode) Then
        If returnPrintFriendlyValue Then
            GetTagValue = GetMetadataEntry(tagIndex).TagValueFriendly
        Else
            GetTagValue = GetMetadataEntry(tagIndex).TagValue
        End If
    Else
        GetTagValue = vbNullString
    End If

End Function

'Given an index and a source metadata object, fill that object with the contents of the matching entry in the m_Tags() array
Friend Function GetMetadataEntry(ByVal mtdIndex As Long) As PDMetadataItem
    If (mtdIndex < m_NumOfTags) And (mtdIndex >= 0) Then GetMetadataEntry = m_Tags(mtdIndex)
End Function

Friend Function SetMetadataEntryByIndex(ByVal mtdIndex As Long, ByRef newMetadata As PDMetadataItem) As Boolean
    If (mtdIndex < m_NumOfTags) And (mtdIndex >= 0) Then
        m_Tags(mtdIndex) = newMetadata
        SetMetadataEntryByIndex = True
    End If
End Function

'Return the total metadata count for this object.  Note that binary and list items ARE included in this count.
Friend Function GetMetadataCount() As Long
    GetMetadataCount = m_NumOfTags
End Function

Friend Function HasMetadata() As Boolean
    HasMetadata = (m_NumOfTags <> 0)
End Function

'It can be desirable to know if we have checked for GPS data and not found any, vs we have NOT looked for GPS data and we have no
' idea if the image contains any.
Friend Function HaveAttemptedToFindGPSData() As Boolean
    HaveAttemptedToFindGPSData = (m_NumOfTags > 0)
End Function

'Some metadata options require the presence of GPS data in the photo.  If a GPS entry was found at metadata load-time, return TRUE.
Friend Function HasGPSMetadata() As Boolean
    HasGPSMetadata = m_HasGPSCoordinates
End Function

'If GPS metadata exists, this function can be used to retrieve the exact latitude and longitude values
Friend Function FillLatitudeLongitude(ByRef latString As String, ByRef lonString As String) As Boolean
    
    If m_HasGPSCoordinates Then
        
        Dim GPSTagIndex As Long
        If DoesTagExist("GPSPosition", GPSTagIndex, vbBinaryCompare) Then
        
            Dim tmpSplit() As String
            tmpSplit = Split(m_Tags(GPSTagIndex).TagValue, " ")
            latString = tmpSplit(0)
            lonString = tmpSplit(1)
            
        End If
    
        FillLatitudeLongitude = True
    
    Else
        FillLatitudeLongitude = False
    End If
    
End Function

'This function will return a new XML file string with all "marked for removal" entries removed
Private Function RetrieveModifiedXMLString() As String
    
    'If this image has no metadata, exit now
    If (m_NumOfTags <= 0) Then
        RetrieveModifiedXMLString = vbNullString
    Else
        
        Dim newXMLString As String
        newXMLString = m_OriginalXML
        
        Dim tagStart As Long, tagEnd As Long
        
        'Before searching for individual tags, we need to manually remove the "rdf:about" entry
        ' ExifTool now places in the file.  (This points to the original file where the metadata
        ' came from - we don't want to write that to a temp file if we don't have to.)
        tagStart = InStr(1, newXMLString, "rdf:about=", vbBinaryCompare)
        If (tagStart > 0) Then
        
            'This is not the most robust solution, but as of v11.89, ExifTool always follows this
            ' rdf:about line with a newline character.  Search for that and use it as our anchor for removal.
            tagEnd = InStr(tagStart, newXMLString, vbCrLf, vbBinaryCompare)
            If (tagEnd > tagStart) Then newXMLString = Left$(newXMLString, tagStart - 1) & Right$(newXMLString, Len(newXMLString) - tagEnd)
        
        End If
        
        'Note that external functions are responsible for marking which entries they want removed.  This function presumes that all
        ' modifications have already been marked, so all we need to do is fulfill the actual removal step.
        Dim i As Long
        For i = 0 To m_NumOfTags - 1
        
            If (m_Tags(i).TagMarkedForRemoval Or m_Tags(i).TagMarkedForRemoval_Export) Then
                
                'Find this tag's start position
                tagStart = InStr(1, newXMLString, "<" & m_Tags(i).TagGroupAndName & ">")
                If (tagStart > 0) Then
                
                    'Remove the leading newline, if any
                    If (InStrRev(newXMLString, vbCrLf, tagStart, vbBinaryCompare) = tagStart - 3) Then tagStart = tagStart - 2
                    
                    'Next, find this tag's end position
                    tagEnd = InStr(tagStart, newXMLString, "</" & m_Tags(i).TagGroupAndName & ">") + Len(m_Tags(i).TagGroupAndName) + 3
                    If (tagEnd > tagStart) Then
                        
                        'If the tag is followed by a newline, remove that too
                        If (InStr(tagEnd, newXMLString, vbCrLf, vbBinaryCompare) = tagEnd + 1) Then tagEnd = tagEnd + 2
                        
                        'Use those positions to fully remove the tag from the parent XML string
                        newXMLString = Left$(newXMLString, tagStart - 1) & Right$(newXMLString, Len(newXMLString) - tagEnd)
                        
                    End If
                
                End If
            
            End If
        
        Next i
            
        RetrieveModifiedXMLString = newXMLString
        
    End If
        
End Function

'Given a valid XML-formatted metadata string from ExifTool, parse it into usable PD metadata format.
' (The parent image's ID format is also passed to this function, because a custom "ready" tag with the image's ID must be manually
'  removed from the XML string.)
Friend Function LoadAllMetadata(ByRef mdString As String, ByVal parentImageID As Long, Optional ByVal sourceIsUndoFile As Boolean = False) As Boolean

    'If the ExifTool plugin can't be found, abandon all metadata loading attempts
    If (Not PluginManager.IsPluginCurrentlyEnabled(CCP_ExifTool)) Then
        Message "ExifTool plugin missing.  Metadata import abandoned."
        LoadAllMetadata = False
        Exit Function
    End If
    
    'If no metadata was returned, we can also abandon this function
    If (LenB(mdString) = 0) Then
        LoadAllMetadata = False
        Exit Function
    End If
    
    'We now know that 1) ExifTool exists, and 2) it returned *something* for this image.
    
    'Because PD intercepts metadata asynchronously, the returned string will contain trailing text similar to {ready123},
    ' where 123 is the ID number of this image.  Because we don't actually want to store that text, use the supplied
    ' parentImageID to find and remove the trailing synchronizer bit.
    If (InStrRev(mdString, "{ready", , vbBinaryCompare) <> 0) Then mdString = Replace$(mdString, "{ready" & parentImageID & "}", vbNullString)
    
    'We now have a (potentially very large) string of XML-formatted metadata.  Before doing anything else, make a backup copy of
    ' that string.  We need this backup when it comes time to write the metadata back out to file.
    
    '(It's a bit confusing, but PD works by passing ExifTool the full original metadata XML, followed by a list of specialized
    ' instructions on what to remove, what to add, and what to change.  This is the most efficient way to write metadata using
    ' ExifTool's rather complex interaction structure.)
    m_OriginalXML = mdString
    
    m_NumOfTags = 0
    ReDim m_Tags(0 To 15) As PDMetadataItem
    
    'ExifTool outputs very clean XML output, with each tag sitting on its own line, and tag hierarchy indicated by preceding
    ' spaces (one per hierarchy level).  As such, it is elegantly parsed by a line-by-line method.  Split it into lines now.
    Dim xmlLines() As String
    xmlLines = Split(mdString, vbCrLf)
    
    If VBHacks.IsArrayInitialized(xmlLines) Then
        
        'With the data successfully split, we now want to advance the parser to the first valid tag entry.
        Dim curLine As Long: curLine = 0
        If PARSE_AdvanceToFirstMeaningfulLine(xmlLines, curLine) Then
            
            'curLine now contains the line index of the first valid tag entry.  From here, all tags will be listed back-to-back,
            ' so we can simply keep iterating through the list until we hit an invalid entry.
            Do While PARSE_LikelyTagLine(xmlLines, curLine)
            Loop
            
        Else
            Debug.Print "Metadata received successfully, but no valid tag entries were found; abandoning further processing."
        End If
        
    Else
        Debug.Print "Raw XML metadata received, but splitting by line-ending seems to have failed??"
    End If
    
    'm_Tags() now contains all successfully parsed metadata entries in the file.  We're going to do a little extra processing to
    ' the tag list, to try and remove any potentially confusing or unhelpful entries.
    
    m_HasGPSCoordinates = False
    
    If (m_NumOfTags > 0) Then
    
        Dim i As Long
        For i = 0 To m_NumOfTags - 1
        
            'First, we're going to perform a special check for unextracted binary data.
            
            'ExifTool exposes a "-b" flag that instructs it to extract binary data for further examination.  I've gone back and forth
            ' on whether to use this.  There are pros and cons to extracting binary data, and among the big cons are its tendency to
            ' add a significant amount of processing time to RAW images.
            
            'At present, I have *not* enabled it, but ExifTool sometimes makes its own decision about what binary data is worth extracting.
            ' When it does this, it will post a comment similar to "(binary data 12345 bytes, use -b option to extract)".  We want to find
            ' any entries like this and remove the developer-centric "use -b option..." instructions.
            If (InStr(1, m_Tags(i).TagValue, ", use -b option to extract") <> 0) Then
                m_Tags(i).TagValue = Replace$(m_Tags(i).TagValue, ", use -b option to extract", vbNullString, , , vbBinaryCompare)
            End If
            
            If (InStr(1, m_Tags(i).TagValueFriendly, ", use -b option to extract") <> 0) Then
                m_Tags(i).TagValueFriendly = Replace$(m_Tags(i).TagValueFriendly, ", use -b option to extract", vbNullString, , , vbBinaryCompare)
            End If
            
            'Next, look for GPS coordinate tags.  If found, we can enable some fun bonus metadata options (like pulling up Google Maps
            ' with the photo's location already helpfully tagged).
            If Strings.StringsEqual(m_Tags(i).tagName, "GPSPosition", vbBinaryCompare) Then
                m_HasGPSCoordinates = True
                Interface.SyncInterfaceToCurrentImage
            End If
            
            'Next, we want to mark any tags that are not relevant to the user.  We will still preserve these tags internally,
            ' but they will be forcibly hidden on the editor screen.
            If Strings.StringsEqual(m_Tags(i).TagGroup, "ExifTool") Then m_Tags(i).InternalUseOnly = True
            
            'Next, we want to improve the grouping of tags.  ExifTool's grouping is somewhat erratic; for example, all software
            ' treats ICC profiles as single, contiguous units, but ExifTool tends to group ICC properties into a bunch of
            ' highly specific subcategories.  I prefer to use the same grouping as other software, to avoid confusing users.
            CalculateIdealGroupName m_Tags(i), i, parentImageID
            
        Next i
    
    End If
    
    'We now have all tags sorted into appropriate categories, and we've extracted as much data as we can out of the XML source
    ' provided by ExifTool.
    
    'The last task we need to perform is looking for obscure metadata tags that might affect the image's on-screen appearance.
    ' (The need for this is generally rare, FYI - mostly weird PNG chunks and multipage TIFFs that have per-page details.)
    If (m_NumOfTags <> 0) And (Not sourceIsUndoFile) Then
        If ApplyMetadataPostProcessing(parentImageID) Then Interface.SyncInterfaceToCurrentImage
    End If
    
    LoadAllMetadata = (m_NumOfTags <> 0)
    
End Function

Private Function CalculateIdealGroupName(ByRef dstMetadata As PDMetadataItem, ByVal tagIndex As Long, ByVal parentImageID As Long) As Boolean
    
    Dim i As Long
    
    Dim GroupName As String, groupNameCheck As String
    GroupName = dstMetadata.TagGroup
    groupNameCheck = UCase$(dstMetadata.TagGroup)
    
    dstMetadata.TagGroupFriendly = GroupName
    
    'First, there's no need for a distinction between "System" and "File".  Group these under a "General" heading.
    ' (Pre-7.0's release I am experimenting with leaving these as-is.)
    'If StrComp(groupNameCheck, "SYSTEM", vbBinaryCompare) = 0 Then dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("General")
    'If StrComp(groupNameCheck, "FILE", vbBinaryCompare) = 0 Then dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("General")
    
    'Next, ExifTool provides "Composite" tags - basically, tags that do not exist in the file, but whose value can be inferred
    ' from other pieces of information.  As "Composite" doesn't make much sense, we use "Inferred" instead.
    If Strings.StringsEqual(groupNameCheck, "COMPOSITE", False) Then dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("Inferred")
    
    'Next, ExifTool breaks down EXIF tags into their individual IFD blocks.  Simplify all of these under a single "Exif" header.
    ' (For TIFFs, which may have multiple pages, we handle this differently.)
    If PDImages.IsImageNonNull(parentImageID) Then
        If (PDImages.GetImageByID(parentImageID).GetOriginalFileFormat <> PDIF_TIFF) Then
            If InStr(1, groupNameCheck, "IFD1", vbBinaryCompare) > 0 Then
                dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("Exif (thumbnail)")
            ElseIf InStr(1, groupNameCheck, "IFD", vbBinaryCompare) > 0 Then
                dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("Exif")
            End If
        End If
    End If

    'Next, ExifTool separates color profiles into header, profile, and measurement subcategories.  Condense these into
    ' a single "ICC Profile" group.
    If InStr(1, groupNameCheck, "ICC", vbBinaryCompare) > 0 Then dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("ICC profile")
    
    'Next, XMP tags often contain an additional identifier, such as "XMP-photoshop".  We'll display names for a few of the most
    ' common XMP groups, while everything else gets condensed under an "XMP (misc)" heading.
    If InStr(1, groupNameCheck, "XMP", vbBinaryCompare) > 0 Then CalculateIdealXMPGroupName dstMetadata
    
    'Next comes a tricky one.  There is no technical prohibition against duplicate metadata tags - and in fact, some software may
    ' choose to deliberately add duplicate tags in multiple groups (IPTC, EXIF, XMP, etc) to maximize compatibility.
    
    'PhotoDemon allows the presence of multiple tags everywhere except in JFIF headers.  JPEG thumbnails embedded inside a JPEG
    ' parent image will have their own JFIF header, which includes information like x/y resolution that is totally irrelevant
    ' (and confusing, as it makes the image look like it has multiple headers!).  Here, we want to check for duplicate JFIF tags,
    ' and place them into their own category as necessary.
    If (InStr(1, groupNameCheck, "JFIF", vbBinaryCompare) > 0) Then

        'Check the metadata list up to this point.  If a matching group and technical tag name are found,
        ' move this tag into a separate "thumbnail" category.  (Note that we want to avoid matching this tag against itself,
        ' so we only search up to its *current* position.)
        If (tagIndex > 0) Then
            For i = 0 To tagIndex - 1
                If Strings.StringsEqual(dstMetadata.TagGroupAndName, m_Tags(i).TagGroupAndName, False) Then
                    dstMetadata.TagGroupFriendly = "JFIF (thumbnail)"
                    Exit For
                End If
            Next i
        End If
    
    End If
    
    CalculateIdealGroupName = (LenB(dstMetadata.TagGroupFriendly) <> 0)
    
End Function

Private Function CalculateIdealXMPGroupName(ByRef dstMetadata As PDMetadataItem) As Boolean
    
    Dim GroupName As String
    GroupName = LCase$(dstMetadata.TagGroup)
    
    If InStr(1, GroupName, "xmp", vbBinaryCompare) > 0 Then
        
        'There are hundreds of potential XMP namespaces.  Rather than try to convert all of them into human-readable versions,
        ' PD just sorts through the most common ones.  Other coders are welcome to expand this list, using the reference at
        ' http://www.sno.phy.queensu.ca/~phil/exiftool/TagNames/XMP.html
        If Strings.StringsEqual(GroupName, "xmp-aas") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Apple)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-acdsee") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (ACD Systems)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-album") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Adobe Album)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-apple-fi") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Apple)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-aux") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Adobe)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-cc") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Creative Commons)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-cell") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Sony Ericcson)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-creatoratom") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Adobe After Effects)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-crs") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Photoshop Camera Raw)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-dc") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Dublin Core)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-dex") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Description Explorer)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-dicom") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (DICOM)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-digikam") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (DigiKam)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-exif") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (EXIF)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-exifex") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (EXIF 2.3)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-expressionmedia") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Microsoft Expression Media)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-extensis") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Extensis Portfolio)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-fpv") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Fast Picture Viewer)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-getty") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Getty Images)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-gpano") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Google Photosphere)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-ics") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (IDimager)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-iptccore") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (IPTC Core)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-iptcext") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (IPTC Extension)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-lr") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Adobe Lightroom)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-mediapro") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (iView MediaPro)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-microsoft") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Microsoft Photo 1.0)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-mp") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Microsoft Photo 1.2)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-mp1") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Microsoft Photo 1.1)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-mwg-coll") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (MWG)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-mwg-kw") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (MWG)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-mwg-rs") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (MWG)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-pdf") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Adobe PDF)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-pdfx") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Adobe PDF extended)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-photomech") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (PhotoMechanic)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-photoshop") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Adobe Photoshop)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-pixellive") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (PixelLive)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-plus") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Plus License Data)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-pmi") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (PRISM for images)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-prism") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (PRISM)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-prl") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (PRISM Rights)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-prm") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (PRISM Recipe)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-pur") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (PRISM Usage Rights)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-rdf") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (RDF)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-swf") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Adobe SWF)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-tiff") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (EXIF-TIFF)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-x") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Toolkit)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-xmp") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (General)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-xmpbj") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Basic Job)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-xmpdm") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Dynamic Media)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-xmpmm") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Media Management)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-xmpnote") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Note)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-xmpplus") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Picture Licensing)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-xmprights") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Rights Management)")
        ElseIf Strings.StringsEqual(GroupName, "xmp-xmptpg") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Paged Text)")
        Else
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Misc)")
        End If
        
        CalculateIdealXMPGroupName = True
        
    End If
    
End Function

Private Function ApplyMetadataPostProcessing(ByVal parentImageID As Long) As Boolean
    
    Dim ppMayBeNeeded As Boolean: ppMayBeNeeded = False
    
    ppMayBeNeeded = (m_NumOfTags <> 0)
    If ppMayBeNeeded Then ppMayBeNeeded = PDImages.IsImageNonNull(parentImageID)
    
    Dim relevantTagIndex As Long
        
    'At present, the only post-processing supplied by this function is limited to:
    ' 1) DPI information (as a failsafe, because FreeImage doesn't have a way to report that
    '    "this image lacks DPI information"; instead, it always reports 72 DPI.  ExifTool
    '    will give us a more precise report, and it will can resolve conflicting entries.)
    ' 2) Precise palette color count of 8-bpp non-PNG images (because FreeImage always
    '    assumes palettes have 256 colors, which is both stupid and obnoxious).
    ' 3) Multipage TIFFs that may have per-page information stored (like page names)
    ' 4) ICC Profiles that various image parsers are unable to retrieve (but ExifTool can).
    '    This is limited to esoteric ICC+format combinations like GIF, which can *technically*
    '    be created (by non-traditional means) but which lack official support.
    
    
    'Because of this, we only need to proceed if the image meets some very specific criteria.
    If ppMayBeNeeded Then
        
        'Knock DPI out first
        Dim xResolution As Double, yResolution As Double
        If Me.GetResolution(xResolution, yResolution) Then PDImages.GetImageByID(parentImageID).SetDPI xResolution, yResolution
        
        'Check for 8-bpp, non-PNG images.  (PD's internal PNG decoder handles palettes correctly, but images coming from
        ' FreeImage or GDI+ may not.)
        If (Not ppMayBeNeeded) Then ppMayBeNeeded = (PDImages.GetImageByID(parentImageID).HasOriginalPalette()) And (PDImages.GetImageByID(parentImageID).GetOriginalFileFormat <> PDIF_PNG)
        
        'Check for multipage TIFFs that were loaded using the "one layer per page" method
        If (Not ppMayBeNeeded) Then ppMayBeNeeded = (PDImages.GetImageByID(parentImageID).GetOriginalFileFormat = PDIF_TIFF) And (PDImages.GetImageByID(parentImageID).ImgStorage.GetEntry_Boolean("MultipageImportActive", False))
        
        'If this file contains an ICC profile, but our various load engines were unable to retrieve it
        ' (as is possible with some esoteric formats), we may be able to use ExifTool to manually retrieve
        ' the full profile data.
        If (Not ppMayBeNeeded) Then
            If (LenB(PDImages.GetImageByID(parentImageID).GetColorProfile_Original) = 0) Then ppMayBeNeeded = DoesTagExistFullName("ICC_Profile:ICC_Profile", relevantTagIndex, vbBinaryCompare) And ColorManagement.UseEmbeddedICCProfiles()
        End If
        
    End If
    
    If ppMayBeNeeded Then
        
        'TIFF-specific tags:
        If (PDImages.GetImageByID(parentImageID).GetOriginalFileFormat = PDIF_TIFF) Then
        
            'Look for page name tags.  If found, assign them to their matching image layers.
            Dim i As Long
            For i = 0 To PDImages.GetImageByID(parentImageID).GetNumOfLayers - 1
                If DoesTagExistFullName("IFD" & i & ":PageName", relevantTagIndex, vbBinaryCompare) Then
                    PDImages.GetImageByID(parentImageID).GetLayerByIndex(i).SetLayerName GetTagValue("IFD" & i & ":PageName", vbBinaryCompare)
                    ApplyMetadataPostProcessing = True
                End If
            Next i
        
        'END TIFF TAGS
        End If
        
        'NON-FORMAT-SPECIFIC tags:
        If (LenB(PDImages.GetImageByID(parentImageID).GetColorProfile_Original) = 0) And DoesTagExistFullName("ICC_Profile:ICC_Profile", relevantTagIndex, vbBinaryCompare) And (Not PDImages.GetImageByID(parentImageID).ImgStorage.DoesKeyExist("Tone-mapping")) And ColorManagement.UseEmbeddedICCProfiles() Then
            
            'ExifTool pulled an ICC profile from this image file, but our original image load function
            ' did not.  This may be an esoteric format or one with broken formatting.  Let's see if we
            ' can still apply the ICC profile post-load.  (This may not possible if the profile is tied
            ' to something like CMYK, and we have already converted it to RGB - but we can still try.)
            
            'Ask ExifTool to extract the ICC profile to a standalone temp file
            If ExifTool.ExtractICCMetadataToFile(PDImages.GetImageByID(parentImageID)) Then
                
                PDDebug.LogAction "Pausing load process while ExifTool extracts ICC profile to standalone file..."
                
                'Next, we need to pause while ExifTool does its thing.  To minimize re-entrance issues, we only want to
                ' yield to WM_TIMER messages (which is how we asynchronously receive data from ExifTool - timer+pipe).
                Do While ExifTool.IsVerificationModeActive()
                    VBHacks.DoEventsTimersOnly
                Loop
                
                'Retrieve the temp file name ExifTool used
                Dim srcFilename As String
                srcFilename = ExifTool.GetExtractedICCProfilePath()
                
                'Pull the ICC profile into a byte array
                Dim iccBytes() As Byte
                If Files.FileLoadAsByteArray(srcFilename, iccBytes) Then
                    
                    'Load the ICC bytes into a pdICCProfile object
                    Dim tmpProfile As pdICCProfile
                    Set tmpProfile = New pdICCProfile
                    If tmpProfile.LoadICCFromPtr(UBound(iccBytes) + 1, VarPtr(iccBytes(0))) Then
                        
                        PDDebug.LogAction "ICC profile retrieved successfully.  Applying now..."
                        If ColorManagement.ConvertDIBToSRGB(PDImages.GetImageByID(parentImageID).GetActiveDIB, tmpProfile) Then
                        
                            'The conversion worked!  Add this profile to our central cache, and flag the parent image accordingly
                            Dim profHash As String
                            profHash = ColorManagement.AddProfileToCache(tmpProfile)
                            PDImages.GetImageByID(parentImageID).SetColorProfile_Original profHash
                            PDImages.GetImageByID(parentImageID).GetActiveDIB.SetColorManagementState cms_ProfileConverted
                            PDImages.GetImageByID(parentImageID).GetActiveDIB.SetColorProfileHash ColorManagement.GetSRGBProfileHash()
                            
                            'If the image is active, redraw the display to reflect the new colors
                            If (parentImageID = PDImages.GetActiveImageID()) Then
                                Viewport.Stage1_InitializeBuffer PDImages.GetImageByID(parentImageID), FormMain.MainCanvas(0)
                                Interface.NotifyImageChanged parentImageID
                            End If
                        
                        End If
                        
                        ApplyMetadataPostProcessing = True
                    
                    End If
                    
                End If
                
                'Free the temp file we used
                Files.FileDeleteIfExists srcFilename
                
            End If
                
        End If
        
        If ApplyMetadataPostProcessing Then PDDebug.LogAction "NOTE!  Metadata-based post-processing was applied to this image."
        
    'END all preprocessing checks
    End If
    
End Function

Private Function PARSE_AdvanceToFirstMeaningfulLine(ByRef xmlLines() As String, ByRef linePointer As Long) As Boolean
    
    linePointer = 0
    
    Dim i As Long
    For i = LBound(xmlLines) To UBound(xmlLines)
        If PARSE_IsLinePotentialTag(xmlLines(i)) Then
            linePointer = i
            Exit For
        End If
    Next i
    
    PARSE_AdvanceToFirstMeaningfulLine = (linePointer <> 0)

End Function

Private Function PARSE_IsLinePotentialTag(ByRef srcXMLLine As String) As Boolean
    If Strings.StringsEqual(Left$(srcXMLLine, 2), LEFT_TAG_TEST, False) Then
        If Strings.StringsEqual(Right$(srcXMLLine, 1), RIGHT_TAG_TEST, False) Then
            PARSE_IsLinePotentialTag = True
        End If
    End If
End Function

Private Function PARSE_LikelyTagLine(ByRef xmlLines() As String, ByRef linePointer As Long) As Boolean
    
    If linePointer < UBound(xmlLines) Then
        
        'Make sure this line looks like a valid tag definition
        If PARSE_IsLinePotentialTag(xmlLines(linePointer)) Then
        
            'Make sure we have room in our metadata collection for the new tag
            If m_NumOfTags > UBound(m_Tags) Then ReDim Preserve m_Tags(0 To m_NumOfTags * 2 - 1) As PDMetadataItem
            
            '...while sLine and eLine are the line indices where this tag starts and ends
            Dim sLine As Long, eLine As Long
            Dim tagFinished As Boolean
            
            'Tag entries come in many shapes and sizes, but they all follow a predictable pattern.  Here is a typical sample:
            ' <IFD0:ResolutionUnit>
            '  <rdf:Description et:id='296' et:table='Exif::Main'>
            '   <et:desc>Resolution Unit</et:desc>
            '   <et:prt>inches</et:prt>
            '   <et:val>2</et:val>
            '  </rdf:Description>
            ' </IFD0:ResolutionUnit>
            
            '*Not all lines are required*, except for:
            ' 1) The initial rdf:Description.  The contents of this tag can vary, but it will always list a table.
            ' 2) The <et:desc> tag.  This is a human-readable version of the tag name.
            ' 3) The <et:prt> tag.  This is a human-readable ("print") version of the tag value.  For some tags (e.g. dates or strings),
            '     this will be the *only* value provided.  For known tags with hard-coded constants (like the example above),
            '     an <et:val> tag will also be provided, to tell us what the actual, physical value of the tag is.
            
            'Other rare descriptors can exist, but we'll get to those in turn.
            
            'Start by stripping out the full group+tag name combination.  We need this value to find where the tag entry ends.
            Dim curTagName As String, endTagName As String, curTagValue As String
            If PARSE_GetTagName(xmlLines(linePointer), curTagName) Then
                
                tagFinished = False
                sLine = linePointer
                
                'Use the tag name to find the line where this tag ends
                endTagName = "</" & curTagName & ">"
                Do While (InStr(1, xmlLines(linePointer), endTagName, vbBinaryCompare) = 0)
                    linePointer = linePointer + 1
                    If linePointer > UBound(xmlLines) Then
                        PARSE_LikelyTagLine = False
                        Exit Function
                    End If
                Loop
                
                eLine = linePointer
                
                'Make sure the tag comprises at least *6* lines (inclusive).  This is the amount assumed by our parser, and a
                ' deviation indicates an unexpected change in ExifTool's output formatting.
                If (eLine - sLine) >= 5 Then
                
                    'With the tag's start and end line locations now known, we can commence with parsing all the lines in-between.
                    
                    'Start by storing the group and tag names
                    m_Tags(m_NumOfTags).TagGroupAndName = curTagName
                    PARSE_GetTagAndGroupName curTagName, m_Tags(m_NumOfTags)
                    
                    'Next, parse the description line.  This will supply the tag table and low-level numeric ID (if any).
                    ' (Also, any duplicate tags will be tagged by their "et:index" value.)
                    PARSE_RDFDescription xmlLines(sLine + 1), m_Tags(m_NumOfTags)
                    
                    'After the RDF description comes the Exiftool-specific description of the tag name.  This may
                    ' include spaces and special chars (unlike the formal name).
                    If PARSE_WholeLine(xmlLines(sLine + 2), curTagValue, "et:desc") Then m_Tags(m_NumOfTags).TagNameFriendly = ExifTool.PARSE_UnescapeXML(curTagValue)
                    If PARSE_WholeLine(xmlLines(sLine + 3), curTagValue, "et:prt") Then
                        m_Tags(m_NumOfTags).TagValueFriendly = ExifTool.PARSE_UnescapeXML(curTagValue)
                    
                    'If the print tag is not contained on a single line, it may be a binary tag.  These only appear if the -b
                    ' option is sent to ExifTool, and they result in XML like this:
                    ' <et:prt rdf:datatype='http://www.w3.org/2001/XMLSchema#base64Binary'>
                    '/9j/4gxYSUNDX1BST0ZJTEUAAQEAAAxITGlubwIQAABtbnRyUkdCIFhZWiAH
                    '...many more Base64 lines...
                    '</et:prt>
                    
                    'PD can parse these via a specialized Binary tag parser.
                    Else
                        tagFinished = PARSE_UnprocessedBinaryValue(xmlLines, sLine, eLine, m_Tags(m_NumOfTags))
                    End If
                    
                    'For many tags, this will be the full extent of the tag's contents.  Let's look for one more semi-common tag
                    ' before exiting - a formal value tag, which is the non-human-friendly version of the tag's contents.
                    If PARSE_WholeLine(xmlLines(sLine + 4), curTagValue, "et:val") Then
                        m_Tags(m_NumOfTags).TagValue = curTagValue
                        tagFinished = PARSE_EndOfTagDetails(xmlLines(sLine + 5))
                    
                    'Check for a closing rdf description; if it exists, we can ignore any further parsing.
                    Else
                        m_Tags(m_NumOfTags).TagValue = m_Tags(m_NumOfTags).TagValueFriendly
                        
                        'In 99% of cases, this is just a simple tag that ends at (sLine + 4).  Rare, fully processed binary tags,
                        ' (as demonstrated in the previous If/Then block) will also end here, but the finished line of their tag
                        ' will have already been calculated by a previous step, and tagFinished will be TRUE.
                        If (Not tagFinished) Then tagFinished = PARSE_EndOfTagDetails(xmlLines(sLine + 4))
                        
                    End If
                    
                    'If a closing rdf description has already been located, this tag has been successfully parsed
                    If (Not tagFinished) Then
                        
                        'This tag contains even more details.  It is possibly a binary-type tag (like a thumbnail) or a list-type tag
                        ' (common in images coming from PhotoShop, as PS tracks thing like image history).  It could also be a
                        ' comment or other string-based tag that contains illegal linebreaks.
                        
                        'Let's start by checking for list-type values (e.g. an rdf:Bag comprised of multiple list nodes).
                        ' rdf:Bag entries are well-formed, so they are easily detected and parsed.
                        tagFinished = PARSE_RDFBagCheck(xmlLines, sLine, eLine, m_Tags(m_NumOfTags))
                        
                        'If this is not a list-type tag, check for a binary tag.
                        If (Not tagFinished) Then
                            tagFinished = PARSE_BinaryValueCheck(xmlLines, sLine, eLine, m_Tags(m_NumOfTags))
                        End If
                        
                        'If this is not a list tag or a binary tag, it must be an invalid tag of some kind - possibly one that uses
                        ' a linebreak character.  (Linebreak characters are specifically disallowed by most format specs, but because
                        ' we're nice, we'll try and salvage them.)
                        If (Not tagFinished) Then
                            tagFinished = PARSE_InvalidMultilineTag(xmlLines, sLine, eLine, m_Tags(m_NumOfTags))
                            If (Not tagFinished) Then Debug.Print "WARNING: weird tag encountered: " & m_Tags(m_NumOfTags).TagGroupAndName
                        End If
                        
                    End If
                    
                    'With all parsing finished, we can finally increment the tag count and proceed to the next tag
                    m_NumOfTags = m_NumOfTags + 1
                    
                Else
                    Debug.Print "Ending line of this tag is too close to starting line - abandoning parse!"
                End If
                
                'Increment the line pointer one final time, so that it points to the line *after* this one
                linePointer = eLine + 1
                PARSE_LikelyTagLine = True
                
            End If
            
        End If
        
    End If

End Function

'Unprocessed binary values are only available if the -b option is sent to ExifTool at extraction time.  Otherwise, binary tags
' are generally ignored, unless their length is extremely small (a few bytes)
Private Function PARSE_UnprocessedBinaryValue(ByRef xmlLines() As String, ByRef tagStartLine As Long, ByRef tagEndLine As Long, ByRef dstMetadata As PDMetadataItem) As Boolean
    
    dstMetadata.IsTagBinary = False
    
    'Here is the layout of a base-64 encoded binary tag (this is taken from an actual, real-world image):
'    <Composite:ThumbnailImage>
'     <rdf:Description et:id='ThumbnailImage' et:table='Composite'>
'      <et:desc>Thumbnail Image</et:desc>
'      <et:prt rdf:datatype='http://www.w3.org/2001/XMLSchema#base64Binary'>
'/9j/4gxYSUNDX1BST0ZJTEUAAQEAAAxITGlubwIQAABtbnRyUkdCIFhZWiAH
'zgACAAkABgAxAABhY3NwTVNGVAAAAABJRUMgc1JHQgAAAAAAAAAAAAAAAAAA
'...many more Base64 lines...
'2Q==
'</et:prt>
'     </rdf:Description>
'    </Composite:ThumbnailImage>
    
    Dim i As Long
    Dim finalValueLine As Long, fullTagValue As String
    
    'ExifTool's output is consistent (including the datatype identifier), so we can use line separators to quickly check for such tags.
    If Strings.StringsEqual(Trim$(xmlLines(tagStartLine + 3)), "<et:prt rdf:datatype='http://www.w3.org/2001/XMLSchema#base64Binary'>", False) Then
        
        'This is a binary tag.  *STARTING WITH THE NEXT LINE*, everything that follows is part of the tag, until the closing
        ' </et:val> tag is reached.  (Note that we will need to strip a trailing CrLf from the value.)
        dstMetadata.IsTagBinary = True
        
        'Looping backward from the closing tag line (which has already been discovered!), look for the closing </et:prt> tag.
        For i = (tagEndLine - 1) To tagStartLine Step -1
            If Strings.StringsEqual(Trim$(xmlLines(i)), "</et:prt>", False) Then Exit For
        Next i
        
        If (i > tagStartLine) Then
            
            finalValueLine = i - 1
            
            For i = (tagStartLine + 4) To finalValueLine
                fullTagValue = fullTagValue & xmlLines(i)
            Next i
            
            'Strip any vbCrLfs in the data and store the raw Base64 data.  (Without additional semantics, it's hard to know what format
            ' the data actually represents.)
            'If (InStr(1, fullTagValue, vbCrLf, vbBinaryCompare) <> 0) Then fullTagValue = Replace$(fullTagValue, vbCrLf, vbNullString, , , vbBinaryCompare)
            dstMetadata.TagBase64Value = fullTagValue
            dstMetadata.TagValueFriendly = fullTagValue
            dstMetadata.WasBinaryExtracted = True
            
        End If
        
        PARSE_UnprocessedBinaryValue = dstMetadata.IsTagBinary
        
    'If this line does *not* include a base-64 tag, it is a normal printable tag -- just a multiline one.
    Else
        
        'Looping backward from the closing tag line (which has already been discovered!), look for the closing </et:prt> tag.
        For i = (tagEndLine - 1) To tagStartLine Step -1
            If Strings.StringsEqual(Trim$(xmlLines(i)), "</et:prt>", False) Then Exit For
        Next i
        
        If (i >= tagStartLine) Then
            
            finalValueLine = i - 1
            
            If (Len(xmlLines(tagStartLine + 3)) > 8) Then
                fullTagValue = xmlLines(tagStartLine + 3)
                fullTagValue = Right$(fullTagValue, Len(fullTagValue) - (InStr(1, fullTagValue, "<et:prt>", vbBinaryCompare) + 7))
            End If
            
            For i = (tagStartLine + 4) To finalValueLine
                fullTagValue = fullTagValue & xmlLines(i)
            Next i
            
            'Store the result as a normal tag.
            dstMetadata.TagValueFriendly = ExifTool.PARSE_UnescapeXML(fullTagValue)
            dstMetadata.TagValue = dstMetadata.TagValueFriendly
            PARSE_UnprocessedBinaryValue = True
            
        End If
        
    End If
    
End Function

Private Function PARSE_BinaryValueCheck(ByRef xmlLines() As String, ByRef tagStartLine As Long, ByRef tagEndLine As Long, ByRef dstMetadata As PDMetadataItem) As Boolean
    
    dstMetadata.IsTagBinary = False
    
    'Here is the layout of a base-64 encoded binary tag (this is taken from an actual, real-world image):
'    <IPTC:CodedCharacterSet>
'        <rdf:Description et:id='90' et:table='IPTC::EnvelopeRecord'>
'            <et:desc>Coded Character Set</et:desc>
'            <et:prt>UTF8</et:prt>
'            <et:val rdf:datatype='http://www.w3.org/2001/XMLSchema#base64Binary'>
'GyVH
'</et:val>
'        </rdf:Description>
'    </IPTC:CodedCharacterSet>
    
    'ExifTool's output is consistent (including the datatype identifier), so we can use line separators to quickly check for such tags.
    If Strings.StringsEqual(Trim$(xmlLines(tagStartLine + 4)), "<et:val rdf:datatype='http://www.w3.org/2001/XMLSchema#base64Binary'>", False) Then
        
        'This is a binary tag.  *STARTING WITH THE NEXT LINE*, everything that follows is part of the tag, until the closing
        ' </et:val> tag is reached.  (Note that we will need to strip a trailing CrLf from the value.)
        dstMetadata.IsTagBinary = True
        
        'Looping backward from the closing tag line (which has already been discovered!), look for the closing </et:val> tag.
        Dim i As Long
        For i = (tagEndLine - 1) To tagStartLine Step -1
            If Strings.StringsEqual(Trim$(xmlLines(i)), "</et:val>", False) Then Exit For
        Next i
        
        If (i > tagStartLine) Then
            
            Dim finalValueLine As Long:  finalValueLine = i
            Dim fullTagValue As String
            
            For i = (tagStartLine + 5) To finalValueLine
                fullTagValue = fullTagValue & xmlLines(i)
            Next i
            
            'Strip any vbCrLfs in the data and store the raw Base64 data.  (Without additional semantics, it's hard to know what format
            ' the data actually represents.)
            If (InStr(1, fullTagValue, vbCrLf, vbBinaryCompare) <> 0) Then fullTagValue = Replace$(fullTagValue, vbCrLf, vbNullString, , , vbBinaryCompare)
            dstMetadata.TagBase64Value = fullTagValue
            dstMetadata.WasBinaryExtracted = True
            
            'ExifTool will sometimes store characters of an unknown codepage as Base64 data.  Attempt a heuristic-based conversion to a
            ' raw string and store the result; this is the best we can do at present.
            Dim tmpBytes() As Byte
            Strings.BytesFromBase64 tmpBytes, fullTagValue
            Strings.StringFromMysteryBytes tmpBytes, dstMetadata.TagValue, False
            
        End If
        
    End If
    
    PARSE_BinaryValueCheck = dstMetadata.IsTagBinary
    
End Function

Private Function PARSE_RDFBagCheck(ByRef xmlLines() As String, ByRef tagStartLine As Long, ByRef tagEndLine As Long, ByRef dstMetadata As PDMetadataItem) As Boolean
    
    dstMetadata.IsTagList = False
    
    'Here is the layout of a basic XMP list (this is taken from an actual, real-world image):
    '<XMP-xmpMM:HistorySoftwareAgent>
    '  <rdf:Description et:id='HistorySoftwareAgent' et:table='XMP::xmpMM'>
    '    <et:desc>History Software Agent</et:desc>
    '    <et:prt>Adobe Photoshop CS6 (Windows);Adobe Photoshop CS6 (Windows);Adobe Photoshop CS6 (Windows);Adobe Photoshop CS6 (Windows);Adobe Photoshop CS6 (Windows);Adobe Photoshop CS6 (Windows);Adobe Photoshop CS6 (Windows);Adobe Photoshop CS6 (Windows)</et:prt>
    '    <et:val>
    '    <rdf:Bag>
    '        <rdf:li>Adobe Photoshop CS6 (Windows)</rdf:li>
    '        <rdf:li>Adobe Photoshop CS6 (Windows)</rdf:li>
    '        <rdf:li>Adobe Photoshop CS6 (Windows)</rdf:li>
    '        <rdf:li>Adobe Photoshop CS6 (Windows)</rdf:li>
    '        <rdf:li>Adobe Photoshop CS6 (Windows)</rdf:li>
    '        <rdf:li>Adobe Photoshop CS6 (Windows)</rdf:li>
    '        <rdf:li>Adobe Photoshop CS6 (Windows)</rdf:li>
    '        <rdf:li>Adobe Photoshop CS6 (Windows)</rdf:li>
    '    </rdf:Bag>
    '    </et:val>
    '    </rdf:Description>
    '</XMP-xmpMM:HistorySoftwareAgent>
    
    'ExifTool's output is consistent, so we can use line separators to quickly check for such tags.
    
    'Start by looking for a bare <et:val> tag on the (startLine + 4) line
    If Strings.StringsEqual(Trim$(xmlLines(tagStartLine + 4)), "<et:val>", False) Then
        
        'Check the next line for a bare <rdf:Bag> tag
        If Strings.StringsEqual(Trim$(xmlLines(tagStartLine + 5)), "<rdf:Bag>", False) Then
            
            'This is a list.  We could manually parse the list, but for now, just use the semicolon-delimited version from the
            ' <et:prt> tag.
            dstMetadata.IsTagList = True
            
        End If
    End If
    
    PARSE_RDFBagCheck = dstMetadata.IsTagList
    
End Function

Private Function PARSE_InvalidMultilineTag(ByRef xmlLines() As String, ByRef tagStartLine As Long, ByRef tagEndLine As Long, ByRef dstMetadata As PDMetadataItem) As Boolean
    
    PARSE_InvalidMultilineTag = False
    dstMetadata.IsTagBinary = False
    dstMetadata.IsTagList = False
    
    'Here is the layout of an invalid, multiline XMP tag (this is taken from an actual, real-world image):
    ' <XMP-dc:Description>
    '  <rdf:Description et:id='description' et:table='XMP::dc'>
    '   <et:desc>Description</et:desc>
    '   <et:prt>???? ????
    '
    '??? ???? ? ????? ?????
    '
    '
    '
    '
    '
    ' *** Local Caption *** ??????</et:prt>
    '  </rdf:Description>
    ' </XMP-dc:Description>
    
    'Unlike other non-standard metadata checks, there is no easy way to check for tags like this, so they simply get a "failsafe"
    ' check after checks for all valid types have failed.
    If (InStr(1, xmlLines(tagStartLine + 3), "<et:prt>", vbBinaryCompare) <> 0) And (InStr(1, xmlLines(tagStartLine + 3), "</et:prt>", vbBinaryCompare) = 0) Then
        
        'This appears to be an invalid multiline tag.  Search subsequent lines for the trailing </et:prt> tag, and note that valid
        ' tag content may appear on the same line where the closing tag is found.  (This is different from other multiline types.)
        
        'Looping backward from the closing tag line (which has already been discovered!), look for the closing </et:prt> tag.
        Dim i As Long
        For i = (tagEndLine - 1) To tagStartLine Step -1
            If (InStr(1, xmlLines(i), "</et:prt>", vbBinaryCompare) <> 0) Then Exit For
        Next i
        
        If (i > tagStartLine) Then
            
            'Assemble the full string into one contiguous length
            Dim finalValueLine As Long:  finalValueLine = i
            Dim fullTagValue As String
            
            For i = (tagStartLine + 3) To finalValueLine
                fullTagValue = fullTagValue & xmlLines(i)
            Next i
            
            'Strip any vbCrLfs in the data, then remove the opening/closing tags
            If (InStr(1, fullTagValue, vbCrLf, vbBinaryCompare) <> 0) Then fullTagValue = Replace$(fullTagValue, vbCrLf, vbNullString, , , vbBinaryCompare)
            fullTagValue = Replace$(fullTagValue, "<et:prt>", vbNullString, , , vbBinaryCompare)
            fullTagValue = Replace$(fullTagValue, "</et:prt>", vbNullString, , , vbBinaryCompare)
            
            'Store the result as a normal tag.
            dstMetadata.TagValueFriendly = ExifTool.PARSE_UnescapeXML(fullTagValue)
            dstMetadata.TagValue = dstMetadata.TagValueFriendly
            PARSE_InvalidMultilineTag = True
            
        End If
        
    End If
    
End Function

Private Function PARSE_EndOfTagDetails(ByRef srcLine As String) As Boolean
    PARSE_EndOfTagDetails = (InStr(1, srcLine, "</rdf:Description>", vbBinaryCompare) <> 0)
End Function

Private Function PARSE_WholeLine(ByRef srcLine As String, ByRef dstString As String, Optional ByVal tagValidate As String = vbNullString) As Boolean

    If (LenB(tagValidate) <> 0) Then
        Dim failedValidation As Boolean
        If (InStr(1, srcLine, "<" & tagValidate & ">", vbBinaryCompare) = 0) Then failedValidation = True
        If (InStr(1, srcLine, "</" & tagValidate & ">", vbBinaryCompare) = 0) Then failedValidation = True
        If failedValidation Then Exit Function
    End If
    
    Dim sPos As Long, ePos As Long
    sPos = InStr(1, srcLine, ">", vbBinaryCompare)
    If (sPos > 0) Then
        ePos = InStrRev(srcLine, "<", , vbBinaryCompare) - 1
        If (ePos > 0) Then
            PARSE_WholeLine = True
            dstString = Mid$(srcLine, sPos + 1, ePos - sPos)
        End If
    End If

End Function

Private Function PARSE_RDFDescription(ByRef srcLine As String, ByRef dstMetadata As PDMetadataItem) As Boolean

    Dim etID As String, etTable As String, etIndex As String
    etID = PARSE_GetNamedAttributeValue(srcLine, "et:id")
    
    'Make sure an ID exists. If it doesn't, something is wrong with the XML output
    If (LenB(etID) > 0) Then
    
        dstMetadata.TagID = etID
        
        etTable = PARSE_GetNamedAttributeValue(srcLine, "et:table")
        If (LenB(etTable) > 0) Then dstMetadata.TagTable = etTable
            
        etIndex = PARSE_GetNamedAttributeValue(srcLine, "et:index")
        If (LenB(etIndex) > 0) Then
            dstMetadata.HasIndex = True
            dstMetadata.TagIndexInternal = CLng(etIndex)
        End If
        
        PARSE_RDFDescription = True
        
    Else
        Debug.Print "WARNING!  RDF tag ID does not exist.  Please investigate!"
    End If

End Function

Private Function PARSE_GetNamedAttributeValue(ByRef srcLine As String, ByRef srcAttribute As String) As String

    Dim posAttribute As Long
    posAttribute = InStr(1, srcLine, srcAttribute & "=", vbBinaryCompare)
    
    If (posAttribute > 0) Then
    
        Dim dFirst As Long, dLast As Long
        dFirst = InStr(posAttribute + 1, srcLine, "'", vbBinaryCompare)
        dLast = InStr(dFirst + 1, srcLine, "'", vbBinaryCompare)
        
        If (dLast > dFirst) Then
            PARSE_GetNamedAttributeValue = Mid$(srcLine, dFirst + 1, (dLast - dFirst) - 1)
        End If
    
    End If

End Function

Private Function PARSE_GetTagAndGroupName(ByRef fullName As String, ByRef dstMetadata As PDMetadataItem) As Boolean
    
    Dim locColon As Long
    locColon = InStr(1, fullName, ":", vbBinaryCompare)
    
    If (locColon <> 0) Then
        dstMetadata.TagGroup = Trim$(Left$(fullName, locColon - 1))
        dstMetadata.tagName = Trim$(Right$(fullName, Len(fullName) - locColon))
        PARSE_GetTagAndGroupName = True
    End If
    
End Function

Private Function PARSE_GetTagName(ByRef srcLine As String, ByRef dstLine As String) As Boolean
    Dim sPos As Long, ePos As Long
    sPos = InStr(1, srcLine, "<", vbBinaryCompare)
    If (sPos > 0) Then
        ePos = InStrRev(srcLine, ">", , vbBinaryCompare) - 1
        If (ePos > 0) Then
            PARSE_GetTagName = True
            dstLine = Mid$(srcLine, sPos + 1, ePos - sPos)
        End If
    End If
End Function

'Want to strip metadata from an image?  Call this to forcibly clear out all metadata information from this object,
' and return it to a pristine, childlike state.
Friend Sub Reset()
    m_NumOfTags = 0
    ReDim m_Tags(0) As PDMetadataItem
    m_OriginalXML = vbNullString
    m_HasGPSCoordinates = False
End Sub

'Given a path to a saved image (yes, that's right - make certain it is already saved!) and a reference
' to the image in question, embed as much or as little metadata as the image's settings specify.
' (Note: this absolutely needs to be done *after* the destination file has been saved, because ExifTool
' requires an on-disk source for writing.  In PD, this function is generally called only by PhotoDemon_SaveImage().)
Friend Function WriteAllMetadata(ByRef outputPath As String, ByRef srcImage As pdImage) As Boolean
    
    'XML writing is handled by PD's central XML engine
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    
    Dim newXMLString As String, tmpXMLPath As String
    
    'Retrieve the XML setting, if any, from the target image object, and write metadata accordingly.
    Dim metadataParams As String
    metadataParams = srcImage.ImgStorage.GetEntry_String("MetadataSettings", vbNullString)
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString metadataParams
    
    'First, see if metadata export is even allowed.
    If cParams.GetBool("MetadataExportAllowed", True) Then
        
        Dim i As Long
        
        'See if the caller wants their metadata anonymized prior to export.
        Dim forciblyAnonymize As Boolean
        forciblyAnonymize = cParams.GetBool("MetadataAnonymize", True)
        
        If (forciblyAnonymize And Me.HasMetadata) Then
        
            Message "Removing metadata with potential privacy concerns..."
        
            'NOTE: GPS tags will be automatically removed by a special parameter we pass to ExifTool,
            ' but any other privacy-problematic tags must be manually flagged.
            For i = 0 To m_NumOfTags - 1
                If (Not m_Tags(i).TagMarkedForRemoval) Then
                    m_Tags(i).TagMarkedForRemoval_Export = ExifTool.DoesTagHavePrivacyConcerns(m_Tags(i))
                End If
            Next i
            
        End If
        
        'Next, manually remove certain metadata entries that are explicitly noted as "not writable";
        ' these include things like format or ExifTool version numbers, which ExifTool will raise
        ' warnings about (as those values need to be implicitly handled by ExifTool itself,
        ' rather than being forcibly set to their previous values).
        For i = 0 To m_NumOfTags - 1
            If (Not m_Tags(i).TagMarkedForRemoval) And (Not m_Tags(i).TagMarkedForRemoval_Export) Then
                m_Tags(i).TagMarkedForRemoval_Export = ExifTool.ShouldTagNeverBeWritten(m_Tags(i))
            End If
        Next i
        
        'Now that all problematic tags have been identified, we need to generate a modified XML string
        ' *without* those entries.
        ' (NOTE: removal of these tags will *still be requested* when we interact with ExifTool,
        '  but by preemptively removing said entries, we add an extra layer of protection against parsing
        '  errors and other glitches.)
        newXMLString = RetrieveModifiedXMLString()
        
        'Note that at this point, the image may not have any metadata.  That's okay - we just skip the part
        ' where we copy over metadata from an existing image.
        Message "Embedding metadata..."
        
        'If, however, the image *does* have metadata, we need to write it out to a temporary XML file.
        If (LenB(newXMLString) <> 0) Then
            
            'Since we can assume the user has rights to the output directory, write the temporary XML file there
            tmpXMLPath = UserPrefs.GetTempPath()
            
            'We'll also append a partial hash to the temporary XML filename to avoid overwriting any existing
            ' XML files in the folder.
            Dim cHash As pdCrypto: Set cHash = New pdCrypto
            tmpXMLPath = tmpXMLPath & "metadata_" & cHash.QuickHashString(CStr(Timer) & outputPath) & ".xml"
            
            'Write the XML file
            Files.FileSaveAsText newXMLString, tmpXMLPath, True, False
            
        Else
            tmpXMLPath = vbNullString
        End If
        
        'The next part is easy - supply the XML file to ExifTool, and ask it to copy metadata from the XML file
        ' we just wrote to the image file.
        WriteAllMetadata = ExifTool.WriteMetadata(tmpXMLPath, outputPath, srcImage, forciblyAnonymize, metadataParams)
        
        If WriteAllMetadata Then
            PDDebug.LogAction "ExifTool initiated: " & tmpXMLPath
        
        'On the off chance that something went wrong, kill the temporary XML file we created.
        Else
            PDDebug.LogAction "ExifTool embed request failed; temp file is " & tmpXMLPath
            Files.FileDeleteIfExists tmpXMLPath
        End If
        
    'If metadata export is *not* allowed, we don't have to do a thing!
    Else
        WriteAllMetadata = True
    End If
    
End Function

Private Sub Class_Initialize()
    Me.Reset
End Sub
