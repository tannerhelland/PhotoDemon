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
'Copyright 2013-2016 by Tanner Helland
'Created: 23/May/13
'Last updated: 09/June/16
'Last update: improve parsing behavior against incorrectly formed multiline print values that are non-binary in nature
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
Public Function GetOriginalXMLMetadataString() As String
    GetOriginalXMLMetadataString = m_OriginalXML
End Function

'Return a current copy of the metadata collection, serialized into a single massive string.  This string contains all user edits,
' and generally speaking, it should be used in parallel to the original XML string, above.  (Both are needed to fully recreate
' metadata inside a new image.)
Public Function GetSerializedXMLData() As String
    
    If m_NumOfTags > 0 Then
        
        Dim cParams As pdParamXML
        Set cParams = New pdParamXML
        cParams.Reset 1#
        
        cParams.AddParam "PDMetadataTagCount", m_NumOfTags
        
        Dim i As Long
        For i = 0 To m_NumOfTags - 1
            cParams.AddParam "PDMD_TAG_NO_" & CStr(i), ExifTool.SerializeTagToString(m_Tags(i))
        Next i
        
    End If
    
    GetSerializedXMLData = cParams.GetParamString
    
End Function

Public Sub RecreateFromSerializedXMLData(ByRef srcString As String)
    
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString srcString
    
    If cParams.GetLong("PDMetadataTagCount", 0) > 0 Then
    
        m_NumOfTags = cParams.GetLong("PDMetadataTagCount", 0)
        ReDim m_Tags(0 To m_NumOfTags) As PDMetadataItem
        
        Dim i As Long
        For i = 0 To m_NumOfTags - 1
            ExifTool.RecoverTagFromSerializedString cParams.GetString("PDMD_TAG_NO_" & CStr(i)), m_Tags(i)
        Next i
        
    End If
    
End Sub

'Retrieve resolution data from the metadata collection, if it exists.  Returns TRUE if resolution metadata entrist exist; FALSE
' if they do not.  Note that this function always returns resolution in inches.  Pixels per CM measurements will be automatically
' converted to PPI.
Public Function GetResolution(ByRef xResolution As Double, ByRef yResolution As Double) As Boolean

    On Error GoTo mdResolutionError

    Dim xResTagIndex As Long, yResTagIndex As Long

    If DoesTagExist("XResolution", xResTagIndex) And DoesTagExist("YResolution", yResTagIndex) Then
    
        'Check for the presence of a resolution type.  If the resolution is in CM, we need to convert it to inches.
        ' Similarly, if no unit is specified, assume inches.
        Dim unitIsCentimeters As Boolean
        unitIsCentimeters = False
        
        Dim resUnitIndex As Long
        If DoesTagExist("ResolutionUnit", resUnitIndex) Then
        
            'Note that 3 is a magic number used by EXIF to indicate CM.  See http://www.sno.phy.queensu.ca/~phil/exiftool/TagNames/EXIF.html
            If CLng(m_Tags(resUnitIndex).TagValue) = 3 Then unitIsCentimeters = True
            
        End If
        
        'Retrieve X and Y resolution values
        xResolution = CDblCustom(m_Tags(xResTagIndex).TagValue)
        yResolution = CDblCustom(m_Tags(yResTagIndex).TagValue)
        
        'If the units are in centimeters, convert them to inches
        If unitIsCentimeters Then
            xResolution = GetInchesFromCM(xResolution)
            yResolution = GetInchesFromCM(yResolution)
        End If
        
        GetResolution = True
    
    'Resolution tags don't exist.
    Else
        GetResolution = False
    End If
    
    Exit Function
    
mdResolutionError:
    GetResolution = False

End Function

'Does a given tag exist in the metadata collection?  Returns TRUE if yes, FALSE, if no.
' If the tag does exist, the optional byref tagIndex will be set to the index.  It will be set to -1 if the tag is not found.
Private Function DoesTagExist(ByVal TagName As String, Optional ByRef tagIndex As Long = -1, Optional ByVal searchMode As VbCompareMethod = vbTextCompare) As Boolean

    DoesTagExist = False
    
    'If this image has no metadata, exit now
    If m_NumOfTags <= 0 Then Exit Function

    'Scan the metadata collection, looking for the given tag name
    Dim i As Long
    For i = 0 To m_NumOfTags - 1
        If StrComp(m_Tags(i).TagName, TagName, searchMode) = 0 Then
            DoesTagExist = True
            tagIndex = i
            Exit For
        End If
    Next i

End Function

'Does a given tag exist in the metadata collection?  Returns TRUE if yes, FALSE, if no.
' If the tag does exist, the optional byref tagIndex will be set to the index.  It will be set to -1 if the tag is not found.
Public Function DoesTagExistFullName(ByVal tagNameAndGroup As String, Optional ByRef tagIndex As Long = -1, Optional ByVal searchMode As VbCompareMethod = vbTextCompare) As Boolean

    DoesTagExistFullName = False
    
    'If this image has no metadata, exit now
    If (m_NumOfTags <= 0) Then Exit Function
    
    'Scan the metadata collection, looking for the given tag name
    Dim i As Long
    For i = 0 To m_NumOfTags - 1
    
        'Tag name found!  Mark the index and exit.
        If StrComp(m_Tags(i).TagGroupAndName, tagNameAndGroup, searchMode) = 0 Then
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
Public Function GetTagValue(ByVal tagNameAndGroup As String, Optional ByVal searchMode As VbCompareMethod = vbTextCompare, Optional ByVal returnPrintFriendlyValue As Boolean = False) As String

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
Public Function GetMetadataCount() As Long
    GetMetadataCount = m_NumOfTags
End Function

Public Function HasMetadata() As Boolean
    HasMetadata = CBool(m_NumOfTags <> 0)
End Function

'It can be desirable to know if we have checked for GPS data and not found any, vs we have NOT looked for GPS data and we have no
' idea if the image contains any.
Public Function HaveAttemptedToFindGPSData() As Boolean
    If (m_NumOfTags = 0) Then
        HaveAttemptedToFindGPSData = False
    Else
        HaveAttemptedToFindGPSData = True
    End If
End Function

'Some metadata options require the presence of GPS data in the photo.  If a GPS entry was found at metadata load-time, return TRUE.
Public Function HasGPSMetadata() As Boolean
    HasGPSMetadata = m_HasGPSCoordinates
End Function

'If GPS metadata exists, this function can be used to retrieve the exact latitude and longitude values
Public Function FillLatitudeLongitude(ByRef latString As String, ByRef lonString As String) As Boolean
    
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
    If m_NumOfTags <= 0 Then
        RetrieveModifiedXMLString = ""
        Exit Function
    End If
    
    Dim newXMLString As String
    newXMLString = m_OriginalXML
    
    Dim tagStart As Long, tagEnd As Long
    
    'Note that external functions are responsible for marking which entries they want removed.  This function presumes that all
    ' modifications have already been marked, so all we need to do is fulfill the actual removal step.
    Dim i As Long
    For i = 0 To m_NumOfTags - 1
    
        If m_Tags(i).TagMarkedForRemoval Then
        
            'Find this tag's start position
            tagStart = InStr(1, newXMLString, "<" & m_Tags(i).TagGroupAndName)
            
            'Next, find this tag's end position
            tagEnd = InStr(tagStart, newXMLString, "</" & m_Tags(i).TagGroupAndName) + Len(m_Tags(i).TagGroupAndName) + 3
            
            'Use those positions to fully remove the tag from the master XML string
            newXMLString = Left$(newXMLString, tagStart - 1) & Right$(newXMLString, Len(newXMLString) - tagEnd)
        
        End If
    
    Next i
        
    RetrieveModifiedXMLString = newXMLString
    
End Function

'Given a valid XML-formatted metadata string from ExifTool, parse it into usable PD metadata format.
' (The parent image's ID format is also passed to this function, because a custom "ready" tag with the image's ID must be manually
'  removed from the XML string.)
Public Function LoadAllMetadata(ByVal mdString As String, ByVal parentImageID As Long) As Boolean

    'If the ExifTool plugin can't be found, abandon all metadata loading attempts
    If (Not g_ExifToolEnabled) Then
        Message "ExifTool plugin missing.  Metadata import abandoned."
        LoadAllMetadata = False
        Exit Function
    End If
    
    'If no metadata was returned, we can also abandon this function
    If Len(mdString) = 0 Then
        LoadAllMetadata = False
        Exit Function
    End If
    
    'We now know that 1) ExifTool exists, and 2) it returned *something* for this image.
    
    'If you want to examine the XML output yourself (e.g. while debugging), you can dump the full XML contents to the clipboard
    ' using the following two lines of code.  (Note that a forcible ANSI conversion will be applied.)
    'Clipboard.Clear
    'Clipboard.SetText mdString
    
    'Because PD intercepts metadata asynchronously, the returned string will contain trailing text similar to {ready123},
    ' where 123 is the ID number of this image.  Because we don't actually want to store that text, use the supplied
    ' parentImageID to find and remove the trailing synchronizer bit.
    If InStrRev(mdString, "{ready", , vbBinaryCompare) <> 0 Then mdString = Replace$(mdString, "{ready" & parentImageID & "}", "")
    
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
    
    If VB_Hacks.IsArrayInitialized(xmlLines) Then
        
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
                m_Tags(i).TagValue = Replace$(m_Tags(i).TagValue, ", use -b option to extract", "", , , vbBinaryCompare)
            End If
            
            If (InStr(1, m_Tags(i).TagValueFriendly, ", use -b option to extract") <> 0) Then
                m_Tags(i).TagValueFriendly = Replace$(m_Tags(i).TagValueFriendly, ", use -b option to extract", "", , , vbBinaryCompare)
            End If
            
            'Next, look for GPS coordinate tags.  If found, we can enable some fun bonus metadata options (like pulling up Google Maps
            ' with the photo's location already helpfully tagged).
            If (StrComp(m_Tags(i).TagName, "GPSPosition", vbBinaryCompare) = 0) Then
                m_HasGPSCoordinates = True
                Interface.SyncInterfaceToCurrentImage
            End If
            
            'Next, we want to mark any tags that are not relevant to the user.  We will still preserve these tags internally,
            ' but they will be forcibly hidden on the editor screen.
            If StrComp(m_Tags(i).TagGroup, "ExifTool") = 0 Then m_Tags(i).InternalUseOnly = True
            
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
    If (m_NumOfTags <> 0) Then
        If ApplyMetadataPostProcessing(parentImageID) Then Interface.SyncInterfaceToCurrentImage
    End If
    
    LoadAllMetadata = CBool(m_NumOfTags <> 0)
    
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
    If StrComp(groupNameCheck, "COMPOSITE", vbBinaryCompare) = 0 Then dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("Inferred")
    
    'Next, ExifTool breaks down EXIF tags into their individual IFD blocks.  Simplify all of these under a single "Exif" header.
    ' (For TIFFs, which may have multiple pages, we handle this differently.)
    If (Not (pdImages(parentImageID) Is Nothing)) Then
        If (pdImages(parentImageID).originalFileFormat <> PDIF_TIFF) Then
            If InStr(1, groupNameCheck, "IFD1", vbBinaryCompare) > 0 Then
                dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("Exif (thumbnail)")
            ElseIf InStr(1, groupNameCheck, "IFD", vbBinaryCompare) > 0 Then
                dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("Exif")
            End If
        End If
    End If

    'Next, ExifTool separates color profiles into header, profile, and measurement subcategories.  Condense these into
    ' a single "ICC Profile" group.
    If InStr(1, groupNameCheck, "ICC", vbBinaryCompare) > 0 Then dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("ICC Profile")
    
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
                If StrComp(dstMetadata.TagGroupAndName, m_Tags(i).TagGroupAndName, vbBinaryCompare) = 0 Then
                    dstMetadata.TagGroupFriendly = "JFIF (thumbnail)"
                    Exit For
                End If
            Next i
        End If
    
    End If
    
    CalculateIdealGroupName = CBool(Len(dstMetadata.TagGroupFriendly) <> 0)
    
End Function

Private Function CalculateIdealXMPGroupName(ByRef dstMetadata As PDMetadataItem) As Boolean
    
    Dim GroupName As String
    GroupName = LCase$(dstMetadata.TagGroup)
    
    If InStr(1, GroupName, "xmp", vbBinaryCompare) > 0 Then
        
        'There are hundreds of potential XMP namespaces.  Rather than try to convert all of them into human-readable versions,
        ' PD just sorts through the most common ones.  Other coders are welcome to expand this list, using the reference at
        ' http://www.sno.phy.queensu.ca/~phil/exiftool/TagNames/XMP.html
        If StringsEqual(GroupName, "xmp-aas") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Apple)")
        ElseIf StringsEqual(GroupName, "xmp-acdsee") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (ACD Systems)")
        ElseIf StringsEqual(GroupName, "xmp-album") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Adobe Album)")
        ElseIf StringsEqual(GroupName, "xmp-apple-fi") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Apple)")
        ElseIf StringsEqual(GroupName, "xmp-aux") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Adobe)")
        ElseIf StringsEqual(GroupName, "xmp-cc") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Creative Commons)")
        ElseIf StringsEqual(GroupName, "xmp-cell") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Sony Ericcson)")
        ElseIf StringsEqual(GroupName, "xmp-creatoratom") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Adobe After Effects)")
        ElseIf StringsEqual(GroupName, "xmp-crs") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Photoshop Camera Raw)")
        ElseIf StringsEqual(GroupName, "xmp-dc") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Dublin Core)")
        ElseIf StringsEqual(GroupName, "xmp-dex") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Description Explorer)")
        ElseIf StringsEqual(GroupName, "xmp-dicom") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (DICOM)")
        ElseIf StringsEqual(GroupName, "xmp-digikam") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (DigiKam)")
        ElseIf StringsEqual(GroupName, "xmp-exif") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (EXIF)")
        ElseIf StringsEqual(GroupName, "xmp-exifex") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (EXIF 2.3)")
        ElseIf StringsEqual(GroupName, "xmp-expressionmedia") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Microsoft Expression Media)")
        ElseIf StringsEqual(GroupName, "xmp-extensis") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Extensis Portfolio)")
        ElseIf StringsEqual(GroupName, "xmp-fpv") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Fast Picture Viewer)")
        ElseIf StringsEqual(GroupName, "xmp-getty") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Getty Images)")
        ElseIf StringsEqual(GroupName, "xmp-gpano") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Google Photosphere)")
        ElseIf StringsEqual(GroupName, "xmp-ics") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (IDimager)")
        ElseIf StringsEqual(GroupName, "xmp-iptccore") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (IPTC Core)")
        ElseIf StringsEqual(GroupName, "xmp-iptcext") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (IPTC Extension)")
        ElseIf StringsEqual(GroupName, "xmp-lr") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Adobe Lightroom)")
        ElseIf StringsEqual(GroupName, "xmp-mediapro") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (iView MediaPro)")
        ElseIf StringsEqual(GroupName, "xmp-microsoft") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Microsoft Photo 1.0)")
        ElseIf StringsEqual(GroupName, "xmp-mp") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Microsoft Photo 1.2)")
        ElseIf StringsEqual(GroupName, "xmp-mp1") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Microsoft Photo 1.1)")
        ElseIf StringsEqual(GroupName, "xmp-mwg-coll") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (MWG)")
        ElseIf StringsEqual(GroupName, "xmp-mwg-kw") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (MWG)")
        ElseIf StringsEqual(GroupName, "xmp-mwg-rs") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (MWG)")
        ElseIf StringsEqual(GroupName, "xmp-pdf") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Adobe PDF)")
        ElseIf StringsEqual(GroupName, "xmp-pdfx") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Adobe PDF extended)")
        ElseIf StringsEqual(GroupName, "xmp-photomech") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (PhotoMechanic)")
        ElseIf StringsEqual(GroupName, "xmp-photoshop") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Adobe Photoshop)")
        ElseIf StringsEqual(GroupName, "xmp-pixellive") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (PixelLive)")
        ElseIf StringsEqual(GroupName, "xmp-plus") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Plus License Data)")
        ElseIf StringsEqual(GroupName, "xmp-pmi") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (PRISM for images)")
        ElseIf StringsEqual(GroupName, "xmp-prism") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (PRISM)")
        ElseIf StringsEqual(GroupName, "xmp-prl") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (PRISM Rights)")
        ElseIf StringsEqual(GroupName, "xmp-prm") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (PRISM Recipe)")
        ElseIf StringsEqual(GroupName, "xmp-pur") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (PRISM Usage Rights)")
        ElseIf StringsEqual(GroupName, "xmp-rdf") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (RDF)")
        ElseIf StringsEqual(GroupName, "xmp-swf") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Adobe SWF)")
        ElseIf StringsEqual(GroupName, "xmp-tiff") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (EXIF-TIFF)")
        ElseIf StringsEqual(GroupName, "xmp-x") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Toolkit)")
        ElseIf StringsEqual(GroupName, "xmp-xmp") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (General)")
        ElseIf StringsEqual(GroupName, "xmp-xmpbj") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Basic Job)")
        ElseIf StringsEqual(GroupName, "xmp-xmpdm") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Dynamic Media)")
        ElseIf StringsEqual(GroupName, "xmp-xmpmm") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Media Management)")
        ElseIf StringsEqual(GroupName, "xmp-xmpnote") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Note)")
        ElseIf StringsEqual(GroupName, "xmp-xmpplus") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Picture Licensing)")
        ElseIf StringsEqual(GroupName, "xmp-xmprights") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Rights Management)")
        ElseIf StringsEqual(GroupName, "xmp-xmptpg") Then
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Paged Text)")
        Else
            dstMetadata.TagGroupFriendly = g_Language.TranslateMessage("XMP (Misc)")
        End If
        
        CalculateIdealXMPGroupName = True
        
    End If
    
End Function

Private Function StringsEqual(ByRef str1 As String, ByRef str2 As String) As Boolean
    StringsEqual = CBool(StrComp(str1, str2, vbBinaryCompare) = 0)
End Function

Private Function ApplyMetadataPostProcessing(ByVal parentImageID As Long) As Boolean
    
    Dim ppMayBeNeeded As Boolean: ppMayBeNeeded = False
    
    ppMayBeNeeded = CBool(m_NumOfTags <> 0)
    If ppMayBeNeeded Then ppMayBeNeeded = CBool(parentImageID <= UBound(pdImages))
    If ppMayBeNeeded Then ppMayBeNeeded = Not (pdImages(parentImageID) Is Nothing)
    
    'At present, the only post-processing supplied by this function is limited to:
    ' 1) Rare PNG chunks that the color-management system is unable to handle on its own
    ' 2) Multipage TIFFs that may have per-page information stored (like lpage names)
    
    'Because of this, we only need to proceed if the image meets some very specific criteria.
    If ppMayBeNeeded Then
        
        'Check for PNGs that do not possess an ICC profile
        ppMayBeNeeded = CBool(pdImages(parentImageID).originalFileFormat = PDIF_PNG) And (Not pdImages(parentImageID).GetActiveDIB.ICCProfile.HasICCData)
        
        'Check for multipage TIFFs that were loaded using the "one layer per page" method
        If (Not ppMayBeNeeded) Then ppMayBeNeeded = CBool(pdImages(parentImageID).originalFileFormat = PDIF_TIFF) And (pdImages(parentImageID).imgStorage.GetEntry_Boolean("MultipageImportActive", False))
        
    End If
    
    If ppMayBeNeeded Then
    
        Dim relevantTagIndex As Long
        
        'PNG-specific tags:
        If (pdImages(parentImageID).originalFileFormat = PDIF_PNG) Then
        
            'First up are some PNG chunks that require special handling.  The PNG spec (http://www.libpng.org/pub/png/spec/1.2/png-1.2-pdg.html)
            ' provides instructions on which color modification chunks should be given preference, and we mirror that behavior here.  PD's order
            ' of operations for PNG chunk processing is: ICC, sRGB, cHRM, gAMA.
            
            'Note that ICC handling supercedes all other chunks, so an ICC profile was found, this whole section *must* be skipped.
            ' (We apply this check at the top of the function.)
            
            Dim srcGamma As Double: srcGamma = 0
            
            Dim cHRMHandled As Boolean, gAMAHandled As Boolean, sRGBHandled As Boolean
            cHRMHandled = False: gAMAHandled = False: sRGBHandled = False
                
            'First, check for an sRGB chunk.  Technically, we should use the transform requested by the chunk value, but I don't have
            ' source code for all possible transform modes, so we're just going to do the default one (absolute colorimetric).
            If DoesTagExistFullName("PNG:SRGBRendering", relevantTagIndex, vbBinaryCompare) Then
            
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "PNG file contains sRGB chunk. Applying color conversion now..."
                #End If
                
                ColorManagement.ConvertRGBUsingCustomEndpoints pdImages(parentImageID).GetActiveDIB, 0.64, 0.33, 0.3, 0.6, 0.15, 0.06, 0.3127, 0.329, 0, False, True
                
                sRGBHandled = True
            
            End If
                
            'Next, we're going to look for the "PNG:WhitePointX/Y" pair of tags.  These are found in the cHRM chunk of the
            ' PNG file, and they specify a custom white point to be used as part of color management.  These days, cHRM has largely
            ' been superceded by use of ICC profiles and/or sRGB identifiers, but cHRM is much more compact and may still be used by
            ' websites obsessed with file size.  (The chunk also provides a nice way to test our custom gamut code, which is why I
            ' implemented this feature in the first place!)
            If (Not sRGBHandled) And DoesTagExistFullName("PNG:WhitePointX", relevantTagIndex, vbBinaryCompare) And DoesTagExistFullName("PNG:WhitePointY", relevantTagIndex, vbBinaryCompare) Then
            
                'Both tags were found.  Use of an sRGB tag or ICC profile supercedes the use of specific cHRM values, so check for
                ' those next.  If either are found, ignore further processing.
                If (Not DoesTagExistFullName("PNG:ProfileName", relevantTagIndex, vbBinaryCompare)) And (Not DoesTagExistFullName("PNG:SRGBRendering", relevantTagIndex, vbBinaryCompare)) Then
                
                    'White point values alone are not enough to properly handle color conversion.  We also need RGB tags.
                    If DoesTagExistFullName("PNG:RedX", relevantTagIndex, vbBinaryCompare) And DoesTagExistFullName("PNG:RedY", relevantTagIndex, vbBinaryCompare) And DoesTagExistFullName("PNG:GreenX", relevantTagIndex, vbBinaryCompare) And DoesTagExistFullName("PNG:GreenY", relevantTagIndex, vbBinaryCompare) And DoesTagExistFullName("PNG:BlueX", relevantTagIndex, vbBinaryCompare) And DoesTagExistFullName("PNG:BlueY", relevantTagIndex, vbBinaryCompare) Then
                    
                        'All required tags exist.  Retrieve their values now.
                        Dim WhiteX As Double, WhiteY As Double, RedX As Double, RedY As Double, GreenX As Double, GreenY As Double, BlueX As Double, BlueY As Double
                        WhiteX = CDblCustom(GetTagValue("PNG:WhitePointX", vbBinaryCompare))
                        WhiteY = CDblCustom(GetTagValue("PNG:WhitePointY", vbBinaryCompare))
                        RedX = CDblCustom(GetTagValue("PNG:RedX", vbBinaryCompare))
                        RedY = CDblCustom(GetTagValue("PNG:RedY", vbBinaryCompare))
                        GreenX = CDblCustom(GetTagValue("PNG:GreenX", vbBinaryCompare))
                        GreenY = CDblCustom(GetTagValue("PNG:GreenY", vbBinaryCompare))
                        BlueX = CDblCustom(GetTagValue("PNG:BlueX", vbBinaryCompare))
                        BlueY = CDblCustom(GetTagValue("PNG:BlueY", vbBinaryCompare))
                        
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
                            If DoesTagExistFullName("PNG:Gamma", relevantTagIndex, vbBinaryCompare) Then srcGamma = CDblCustom(GetTagValue("PNG:Gamma", vbBinaryCompare))
                            
                            'If gamma does not exist in the file, or it does exist but it is encoded incorrectly, we'll still proceed with
                            ' the conversion, but we have no choice but to assume a default gamma value.
                            If (srcGamma = 0) Then
                            
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
                                ColorManagement.ConvertRGBUsingCustomEndpoints pdImages(parentImageID).GetActiveDIB, RedX, RedY, GreenX, GreenY, BlueX, BlueY, WhiteX, WhiteY, 0, False, True
                            Else
                                ColorManagement.ConvertRGBUsingCustomEndpoints pdImages(parentImageID).GetActiveDIB, RedX, RedY, GreenX, GreenY, BlueX, BlueY, WhiteX, WhiteY, srcGamma, False, True
                            End If
                            
                            'FOR TESTING PURPOSES: this sample line of code uses hard-coded sRGB values
                            'ColorManagement.convertRGBUsingCustomEndpoints pdImages(parentImageID).getActiveDIB, 0.64, 0.33, 0.3, 0.6, 0.15, 0.06, 0.3127, 0.329, 0, False, True
                            
                            'FOR TESTING PURPOSES: this sample line of code uses hard-coded Adobe RGB values
                            'ColorManagement.convertRGBUsingCustomEndpoints pdImages(parentImageID).getActiveDIB, 0.64, 0.33, 0.21, 0.71, 0.15, 0.06, 0.3127, 0.329, 2.2, False, True
                            
                            'Note that cHRM has been handled; this prevents subsequent gAMA checks from applying a second round of
                            ' corrections to the image.
                            cHRMHandled = True
                            
                            #If DEBUGMODE = 1 Then
                                pdDebug.LogAction "Custom cHRM data analysis complete."
                            #End If
                            
                            'Refresh the screen to reflect the new colors.  (This is necessary because metadata is parsed asynchronously,
                            ' so the image may already be visible on the screen!)
                            If (parentImageID = g_CurrentImage) Then
                                Viewport_Engine.Stage1_InitializeBuffer pdImages(parentImageID), FormMain.mainCanvas(0)
                                Interface.NotifyImageChanged parentImageID
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
            If (Not cHRMHandled) And (Not sRGBHandled) And DoesTagExistFullName("PNG:Gamma", relevantTagIndex, vbBinaryCompare) Then
            
                srcGamma = CDblCustom(GetTagValue("PNG:Gamma", vbBinaryCompare))
                
                'If gamma does not exist in the file, or it does exist but it is encoded incorrectly, we obviously won't make use of it.
                If (srcGamma <= 0) Then
                
                    #If DEBUGMODE = 1 Then
                        pdDebug.LogAction "WARNING! PNG file has gAMA, but the reported value is invalid.  No gamma correction applied."
                    #End If
                
                'If gamma does exist, and it is encoded correctly (as far as we can tell), apply it now
                Else
                
                    #If DEBUGMODE = 1 Then
                        pdDebug.LogAction "PNG file with custom gAMA data found. Applying gamma correction now..."
                    #End If
                    
                    Filters_Layers.GammaCorrectDIB pdImages(parentImageID).GetActiveDIB, 2.2 / srcGamma, True
                    
                    #If DEBUGMODE = 1 Then
                        pdDebug.LogAction "Custom cHRM data analysis complete."
                    #End If
                    
                    'Note that gAMA has been handled; this prevents subsequent checks from possibly applying a second round of
                    ' corrections to the image.
                    gAMAHandled = True
                    
                    'Refresh the screen to reflect the new colors
                    If (parentImageID = g_CurrentImage) Then
                        Viewport_Engine.Stage1_InitializeBuffer pdImages(parentImageID), FormMain.mainCanvas(0)
                        Interface.NotifyImageChanged parentImageID
                    End If
                    
                End If
            
            'END gAMA chunk handling
            End If
                
        'END PNG TAGS
        End If
        
        'TIFF-specific tags:
        If (pdImages(parentImageID).originalFileFormat = PDIF_TIFF) Then
        
            'Look for page name tags.  If found, assign them to their matching image layers.
            Dim i As Long
            For i = 0 To pdImages(parentImageID).GetNumOfLayers - 1
                If DoesTagExistFullName("IFD" & i & ":PageName", relevantTagIndex, vbBinaryCompare) Then
                    pdImages(parentImageID).GetLayerByIndex(i).SetLayerName GetTagValue("IFD" & i & ":PageName", vbBinaryCompare)
                End If
            Next i
        
        'END TIFF TAGS
        End If
        
        ApplyMetadataPostProcessing = True
        
        #If DEBUGMODE = 1 Then
            If ApplyMetadataPostProcessing Then pdDebug.LogAction "NOTE!  Metadata-based post-processing was applied to this image."
        #End If
        
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
    
    PARSE_AdvanceToFirstMeaningfulLine = CBool(linePointer <> 0)

End Function

Private Function PARSE_IsLinePotentialTag(ByRef srcXMLLine As String) As Boolean
    If StrComp(Left$(srcXMLLine, 2), LEFT_TAG_TEST, vbBinaryCompare) = 0 Then
        If StrComp(Right$(srcXMLLine, 1), RIGHT_TAG_TEST, vbBinaryCompare) = 0 Then
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
            
            'We're going to be doing lots of string parsing in this function.  sPos and ePos are used to mark the start and end
            ' of a parse operation...
            Dim sPos As Long, ePos As Long
            
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
                        ' a linebreak character.
                        ' TODO: handle this tag specially??
                        If (Not tagFinished) Then
                            Debug.Print "WARNING: weird tag encountered: " & m_Tags(m_NumOfTags).TagGroupAndName
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
    If StrComp(Trim$(xmlLines(tagStartLine + 3)), "<et:prt rdf:datatype='http://www.w3.org/2001/XMLSchema#base64Binary'>", vbBinaryCompare) = 0 Then
        
        'This is a binary tag.  *STARTING WITH THE NEXT LINE*, everything that follows is part of the tag, until the closing
        ' </et:val> tag is reached.  (Note that we will need to strip a trailing CrLf from the value.)
        dstMetadata.IsTagBinary = True
        
        'Looping backward from the closing tag line (which has already been discovered!), look for the closing </et:prt> tag.
        For i = (tagEndLine - 1) To tagStartLine Step -1
            If StrComp(Trim$(xmlLines(i)), "</et:prt>", vbBinaryCompare) = 0 Then Exit For
        Next i
        
        If (i > tagStartLine) Then
            
            finalValueLine = i - 1
            
            For i = (tagStartLine + 4) To finalValueLine
                fullTagValue = fullTagValue & xmlLines(i)
            Next i
            
            'Strip any vbCrLfs in the data and store the raw Base64 data.  (Without additional semantics, it's hard to know what format
            ' the data actually represents.)
            'If (InStr(1, fullTagValue, vbCrLf, vbBinaryCompare) <> 0) Then fullTagValue = Replace$(fullTagValue, vbCrLf, "", , , vbBinaryCompare)
            dstMetadata.TagBase64Value = fullTagValue
            dstMetadata.TagValueFriendly = fullTagValue
            dstMetadata.WasBinaryExtracted = True
            
        End If
        
        PARSE_UnprocessedBinaryValue = dstMetadata.IsTagBinary
        
    'If this line does *not* include a base-64 tag, it is a normal printable tag -- just a multiline one.
    Else
        
        'Looping backward from the closing tag line (which has already been discovered!), look for the closing </et:prt> tag.
        For i = (tagEndLine - 1) To tagStartLine Step -1
            If StrComp(Trim$(xmlLines(i)), "</et:prt>", vbBinaryCompare) = 0 Then Exit For
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
    If StrComp(Trim$(xmlLines(tagStartLine + 4)), "<et:val rdf:datatype='http://www.w3.org/2001/XMLSchema#base64Binary'>", vbBinaryCompare) = 0 Then
        
        'This is a binary tag.  *STARTING WITH THE NEXT LINE*, everything that follows is part of the tag, until the closing
        ' </et:val> tag is reached.  (Note that we will need to strip a trailing CrLf from the value.)
        dstMetadata.IsTagBinary = True
        
        'Looping backward from the closing tag line (which has already been discovered!), look for the closing </et:val> tag.
        Dim i As Long
        For i = (tagEndLine - 1) To tagStartLine Step -1
            If StrComp(Trim$(xmlLines(i)), "</et:val>", vbBinaryCompare) = 0 Then Exit For
        Next i
        
        If (i > tagStartLine) Then
            
            Dim finalValueLine As Long:  finalValueLine = i
            Dim fullTagValue As String
            
            For i = (tagStartLine + 5) To finalValueLine
                fullTagValue = fullTagValue & xmlLines(i)
            Next i
            
            'Strip any vbCrLfs in the data and store the raw Base64 data.  (Without additional semantics, it's hard to know what format
            ' the data actually represents.)
            If (InStr(1, fullTagValue, vbCrLf, vbBinaryCompare) <> 0) Then fullTagValue = Replace$(fullTagValue, vbCrLf, "", , , vbBinaryCompare)
            dstMetadata.TagBase64Value = fullTagValue
            dstMetadata.WasBinaryExtracted = True
            
            'ExifTool will sometimes store characters of an unknown codepage as Base64 data.  Attempt a heuristic-based conversion to a
            ' raw string and store the result; this is the best we can do at present.
            Dim cUnicode As pdUnicode
            Set cUnicode = New pdUnicode
            
            Dim tmpBytes() As Byte
            cUnicode.Base64Decode tmpBytes, fullTagValue
            cUnicode.ConvertUnknownBytesToString tmpBytes, dstMetadata.TagValue, False
            
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
    If StrComp(Trim$(xmlLines(tagStartLine + 4)), "<et:val>", vbBinaryCompare) = 0 Then
        
        'Check the next line for a bare <rdf:Bag> tag
        If StrComp(Trim$(xmlLines(tagStartLine + 5)), "<rdf:Bag>", vbBinaryCompare) = 0 Then
            
            'This is a list.  We could manually parse the list, but for now, just use the semicolon-delimited version from the
            ' <et:prt> tag.
            dstMetadata.IsTagList = True
            
        End If
    End If
    
    PARSE_RDFBagCheck = dstMetadata.IsTagList
    
End Function

Private Function PARSE_EndOfTagDetails(ByRef srcLine As String) As Boolean
    PARSE_EndOfTagDetails = CBool(InStr(1, srcLine, "</rdf:Description>", vbBinaryCompare) <> 0)
End Function

Private Function PARSE_WholeLine(ByRef srcLine As String, ByRef dstString As String, Optional ByVal tagValidate As String = vbNullString) As Boolean

    If Len(tagValidate) <> 0 Then
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
    If (Len(etID) > 0) Then
    
        dstMetadata.TagID = etID
        
        etTable = PARSE_GetNamedAttributeValue(srcLine, "et:table")
        If Len(etTable) > 0 Then dstMetadata.TagTable = etTable
            
        etIndex = PARSE_GetNamedAttributeValue(srcLine, "et:index")
        If Len(etIndex) > 0 Then
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
        dstMetadata.TagName = Trim$(Right$(fullName, Len(fullName) - locColon))
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

'Given a path to a saved image (yes, that's right - make certain it is already saved!) and a reference to the image in question,
' embed as much or as little metadata as the image's settings specify.
' (Note: this absolutely needs to be done *after* the destination file has been saved, because ExifTool requires an on-disk
'  source for writing.  In PD, this function is generally called only by the PhotoDemon_SaveImage function.)
Public Function WriteAllMetadata(ByVal outputPath As String, ByRef srcImage As pdImage) As Boolean
    
    'XML writing is handled by PD's master XML engine
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    
    Dim newXMLString As String, tmpXMLPath As String
    Dim fileNum As Integer
    
    Dim cSHA2 As CSHA256
    Set cSHA2 = New CSHA256
    
    'Retrieve the XML setting, if any, from the target image object, and write metadata accordingly.
    Dim metadataParams As String
    metadataParams = srcImage.imgStorage.GetEntry_String("MetadataSettings", vbNullString)
    
    Dim cParams As pdParamXML
    Set cParams = New pdParamXML
    cParams.SetParamString metadataParams
    
    'First, see if metadata export is even allowed
    If cParams.GetBool("MetadataExportAllowed", True) Then
    
        'See if the caller wants their metadata anonymized prior to export
        Dim forciblyAnonymize As Boolean
        forciblyAnonymize = cParams.GetBool("MetadataAnonymize", True)
        
        If forciblyAnonymize And Me.HasMetadata Then
        
            Message "Removing metadata with potential privacy concerns..."
        
            'NOTE: GPS tags will be automatically removed by a special parameter we pass to ExifTool, but any other
            ' privacy-problematic tags must be manually flagged.
            Dim i As Long
            For i = 0 To m_NumOfTags - 1
                If (Not m_Tags(i).TagMarkedForRemoval) Then m_Tags(i).TagMarkedForRemoval = ExifTool.DoesTagHavePrivacyConcerns(m_Tags(i))
            Next i
            
        End If
            
        'Now that all problematic tags have been identified, we need to generate a modified XML string without those entries.
        ' (NOTE: removal of these tags will *still be requested* when we interact with ExifTool, but by preemptively removing
        '        said entries, we add an extra layer of protection against parsing errors and other glitches.)
        newXMLString = RetrieveModifiedXMLString
        
        Message "Embedding metadata..."
        
        'Note that at this point, the image may not have any metadata.  That's okay - we just skip the part where
        ' we copy over metadata from an existing image.
        
        'If, however, the image *does* have metadata, we need to write it out to a temporary XML file.
        If (Len(newXMLString) <> 0) Then
            
            'Since we can assume the user has rights to the output directory, write the temporary XML file there
            tmpXMLPath = GetDirectory(outputPath)
                    
            'We'll also append a partial hash to the temporary XML filename to avoid overwriting any existing XML files in the folder
            tmpXMLPath = tmpXMLPath & "metadata_" & Left$(cSHA2.SHA256(CStr(Timer) & outputPath), 16) & ".xml"
            
            'Write the XML file
            xmlEngine.loadXMLFromString newXMLString
            xmlEngine.writeXMLToFile tmpXMLPath, True
        
        Else
            tmpXMLPath = vbNullString
        End If
        
        'The next part is easy - supply the XML file to ExifTool, and request of copying of metadata from the XML to the image file.
        WriteAllMetadata = WriteMetadata(tmpXMLPath, outputPath, srcImage, forciblyAnonymize, metadataParams)
        
        If WriteAllMetadata Then
            Message "Metadata embedded successfully."
        Else
            Message "Metadata may not have been embedded (but saved image still written successfully)."
        End If
        
    'If metadata export is *not* allowed, we don't have to do a thing!
    Else
        WriteAllMetadata = True
    End If
    
End Function

Private Sub Class_Initialize()
    m_NumOfTags = 0
    ReDim m_Tags(0) As PDMetadataItem
End Sub
