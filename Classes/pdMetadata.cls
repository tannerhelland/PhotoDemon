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
'Copyright ©2012-2013 by Tanner Helland
'Created: 23/May/13
'Last updated: 23/May/13
'Last update: initial build
'
'As of version 5.6, PhotoDemon now provides support for loading and saving image metadata.  What is metadata, you ask?
' See http://en.wikipedia.org/wiki/Metadata#Photographs for more details.
'
'Because metadata handling is complicated, I have wrapped all related functions into this class.  Each pdImage object
' has one (1) metadata class called "imgMetadata".  That class is publicly accessible.  If metadata is available at
' load-time, the imgMetadata object will automatically be filled with any metadata found in the source file.
'
'At present, metadata handling is 100% tied FreeImage.  I have no plans for a GDI+ fallback (at this time).  The primary
' reason for this is that to my understanding, GDI+ does not allow for writing metadata out to an image without first
' loading that image into memory.  For JPEG images, this requires decompression and then recompression, which reduces
' image quality.  (Similarly, I don't believe GDI+ provides a lossless way to handle JPEG loading/saving - but I'd love
' to be proven wrong.)  For more information on how FreeImage handles metadata, see the FreeImage documentation at:
' http://freeimage.sourceforge.net/download.html
'
'Finally, note that FreeImage supports 12 types of metadata, which is why the range [0, 11] is frequently used throughout
' this class.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'Number of tags available for each metadata model
Private numTags(0 To 11) As Long

'The highest metadata count of the 12 sub-types.  We need this to know how to dimension our 2D metadata array.
Private maxMetadataCount As Long

'Metadata itself, in FreeImage-compatible format.  Note that the second array dimension will be set according to
' maxMetadataCount above.  Chances are that only one type of metadata will actually contain that many values; in
' fact, most metadata subtypes will probably contain 0 values.  So DO NOT assume all metadata types have
' maxMetadataCount values available.
Private fiMetadata() As FREE_IMAGE_TAG

'If no metadata was found during the original FreeImage load process, this will return FALSE.  If any metadata of any
' compatible type was found, this will return TRUE.
Public Function hasMetadata() As Boolean
    If maxMetadataCount = 0 Then hasMetadata = False Else hasMetadata = True
End Function

'DO NOTE USE THIS FUNCTION IN PRODUCTION!  I include it here simply for ensuring that metadata was loaded correctly.
Public Sub testMetadata()

    Dim msgString As String

    'For each metadata object, display a message box with the relevant metadata value
    Dim i As Long, j As Long
    For i = 0 To 11
    
        If numTags(i) > 0 Then
            For j = 0 To numTags(i) - 1
                msgString = "Metadata model: " & getFriendlySubtypeName(i) & vbCrLf
                msgString = msgString & "Tag count: " & (j + 1) & " of " & numTags(i) & vbCrLf & vbCrLf
                msgString = msgString & "ID: " & fiMetadata(i, j).Id & vbCrLf
                msgString = msgString & "Description: " & fiMetadata(i, j).Description & vbCrLf
                msgString = msgString & "Value: " & CStr(fiMetadata(i, j).StringValue) & vbCrLf
                MsgBox msgString
            Next j
        End If
    
    Next i

End Sub

'Given a metadata sub-type, return a friendly (and translated!) string describing that sub-type
Private Function getFriendlySubtypeName(ByVal mdSubType As Long) As String

    Dim tmpString As String, prefixString As String

    Select Case mdSubType
        
        'Generic comments.  May be EXIF, may not - FreeImage handles these transparently.
        Case 0
            tmpString = g_Language.TranslateMessage("Comments")
        
        'Generic EXIF data (specifically, metadata common to both TIFF and EXIF files)
        Case 1
            tmpString = g_Language.TranslateMessage("Basic EXIF metadata")
        
        'EXIF-specific metadata
        Case 2
            tmpString = g_Language.TranslateMessage("Advanced EXIF metadata")
        
        'GPS metadata (EXIF-type)
        Case 3
            tmpString = g_Language.TranslateMessage("GPS metadata")
        
        'Maker note metadata (EXIF-type - there is no way to interpret this data, as it is not standardized)
        Case 4
            tmpString = g_Language.TranslateMessage("Custom camera metadata (""maker note"")")
        
        'Interoperability metadata (EXIF-type)
        Case 5
            tmpString = g_Language.TranslateMessage("Interoperability metadata")
        
        'IPTC metadata
        Case 6
            tmpString = g_Language.TranslateMessage("IPTC metadata (an old model, no longer recommended for use)")
        
        'XMP metadata
        Case 7
            tmpString = g_Language.TranslateMessage("XMP metadata (XML format developed by Adobe)")
        
        'GeoTiff metadata
        Case 8
            tmpString = g_Language.TranslateMessage("GeoTIFF metadata")
        
        'Animation metadata (animated GIFs only)
        Case 9
            tmpString = g_Language.TranslateMessage("Animation metadata (GIF only)")
        
        'Custom metadata (anything FreeImage can't convert to other metadata types)
        Case 10
            tmpString = g_Language.TranslateMessage("Custom metadata")
        
        'Raw EXIF metadata (required for writing EXIF data to file - unlikely to ever be found in loaded metadata, as
        ' FreeImage will attempt to convert all EXIF data to a more useful sub-type)
        Case 11
            tmpString = g_Language.TranslateMessage("Raw, uncategorized EXIF metadata")
    
    End Select
    
    prefixString = g_Language.TranslateMessage("Type")
    prefixString = prefixString & " " & mdSubType & ":"
    
    getFriendlySubtypeName = prefixString & " " & tmpString

End Function

'Given a handle to a FreeImage-type image, attempt to load all metadata.
' (Portions of this sub are based on code in the "FreeImage_GetAllMetadataTags" function by Carsten Klein)
Public Sub loadAllMetadata(ByVal fiImageHandle As Long)

    'Start by determining the maximum sub-type metadata count.  We need this value so we know how to dim our metadata array.
    maxMetadataCount = FreeImage_GetMaxMetadataCount(fiImageHandle)
    
    'If no metadata was found, exit now
    If maxMetadataCount = 0 Then Exit Sub
    
    'Redim the metadata array to fit the largest metadata sub-type.  This is not the most efficient way to store metadata,
    ' but short of building a custom linked-list datatype, it's the easiest way to handle it in VB.
    ReDim fiMetadata(0 To 11, 0 To maxMetadataCount) As FREE_IMAGE_TAG
    
    'Proceed to load each metadata type individually.
    Dim fiMetadataType As Long, i As Long
    Dim h_Metadata As Long, tmpTag As Long
    
    For fiMetadataType = 0 To 11
        
        'See if metadata exists for this sub-type.  If it does, proceed to load the metadata.
        numTags(fiMetadataType) = FreeImage_GetMetadataCount(fiMetadataType, fiImageHandle)
        If numTags(fiMetadataType) > 0 Then
        
            'Get a pointer to the first metadata object of this type.  If a valid pointer is returned, proceed to
            ' enumerate all subsequent tags.
            h_Metadata = FreeImage_FindFirstMetadata(fiMetadataType, fiImageHandle, tmpTag)
            If h_Metadata <> 0 Then
                
                i = 0
                Do
                    fiMetadata(fiMetadataType, i) = pGetTagFromTagPtr(fiMetadataType, tmpTag)
                    i = i + 1
                Loop While (FreeImage_FindNextMetadataInt(h_Metadata, tmpTag) <> 0)
                FreeImage_FindCloseMetadata h_Metadata
            
            End If
            
        End If
        
    Next fiMetadataType
    
    'Our fiMetadata() array is now filled with all compatible metadata from the source image.

End Sub

'Given a handle to a FreeImage image object, find the maximum sub-type metadata count
Private Function FreeImage_GetMaxMetadataCount(ByVal fiImageHandle As Long) As Long

    Dim i As Long
    Dim maxMetadata As Long, tmpMetadata As Long
    maxMetadata = 0
    
    For i = 0 To 11
        tmpMetadata = FreeImage_GetMetadataCount(i, fiImageHandle)
        If tmpMetadata > maxMetadata Then maxMetadata = tmpMetadata
    Next i
    
    FreeImage_GetMaxMetadataCount = maxMetadata

End Function

'My (vain?) attempt to write an EXIF to XMP translator
Public Function writeDataAsXMPString() As String

    Dim XMPText As String
    
    'Start by writing out the standard XMP packet header
    XMPText = "<?xpacket begin=""?"" id=""W5M0MpCehiHzreSzNTczkc9d""?>" & vbCrLf
    XMPText = XMPText & "<x:xmpmeta xmlns:x=""adobe:ns:meta/"" x:xmptk=""Adobe XMP Core 5.4-c002 1.000000, 0000/00/00-00:00:00        "">" & vbCrLf
    XMPText = XMPText & "<rdf:RDF xmlns:rdf=""http://www.w3.org/1999/02/22-rdf-syntax-ns#"">" & vbCrLf
    XMPText = XMPText & "      <rdf:Description rdf:about=""""" & vbCrLf
    XMPText = XMPText & "            xmlns:xmp=""http://ns.adobe.com/xap/1.0/"">" & vbCrLf
    XMPText = XMPText & "         <xmp:CreatorTool>PhotoDemon</xmp:CreatorTool>" & vbCrLf
    XMPText = XMPText & "      </rdf:Description>" & vbCrLf
    
    Dim curTitle As String
    
    'FreeImage doesn't write out "description" tags, and I have no plans to manually add them.  So our XMP packet data is going to be a little ugly.  Oh well.
    Dim fiMetadataType As Long, j As Long
    For fiMetadataType = 0 To 10
    
        'Only process tags that exist (obviously!)
        If numTags(fiMetadataType) > 0 Then
            For j = 0 To numTags(fiMetadataType) - 1
                
                curTitle = getXMPfromEXIF(fiMetadata(fiMetadataType, j).Key)
                If curTitle <> "" Then
                    XMPText = XMPText & "<" & curTitle & ">" & fiMetadata(fiMetadataType, j).StringValue & "</" & curTitle & ">" & vbCrLf
                End If
                
            Next j
        End If
    
    Next fiMetadataType
    
    'Terminate the XMP packet
    XMPText = XMPText & "</rdf:RDF>"
    XMPText = XMPText & "</x:xmpmeta>"
    XMPText = XMPText & "<?xpacket end=""w""?>"

    writeDataAsXMPString = XMPText

End Function

'Given a handle to a FreeImage object, write out all currently stored metadata.  (This only needs to be done immediately prior to a save.)
Public Function writeAllMetadata(ByVal fiImageHandle As Long) As Boolean
    
    'New strategy: convert as much EXIF data to XMP format as possible.
    Dim XMPPacket As String
    XMPPacket = writeDataAsXMPString
    
    Dim tmpTag As FREE_IMAGE_TAG
    tmpTag = FreeImage_CreateTagEx(FIMD_XMP, "XMLPacket", FIDT_ASCII, XMPPacket)

    Dim totalSuccess As Boolean
    totalSuccess = True
    
    If tmpTag.TagPtr <> 0 Then
                    
                    'Attempt to set this tag to the FreeImage handle we were passed
                    If FreeImage_SetMetadataEx(fiImageHandle, tmpTag, "XMLPacket", FIMD_XMP, True) Then
                    'If FreeImage_SetMetadataEx(fiImageHandle, tmpTag, fiMetadata(fiMetadataType, j).Key, fiMetadataType, True) Then
                        
                        'Note that normally we would need to delete the tag object just created; however, the RefreshTag
                        ' parameter used above handles that for us.
                        
                    Else
                        totalSuccess = False
                        Message "Failed to set metadata tag to FreeImage object."
                    End If
                    
                Else
                    totalSuccess = False
                    Message "Failed to create metadata tag."
                End If
                
                Exit Function
    
    
    Dim fiMetadataType As Long, j As Long
        
    For fiMetadataType = 0 To 10
    
        'Only process tags that exist (obviously!)
        If numTags(fiMetadataType) > 0 Then
            For j = 0 To numTags(fiMetadataType) - 1
                
                'Create a new FreeImage tag with data identical to this one
                Select Case fiMetadata(fiMetadataType, j).Type
                    
                    Case FIDT_SRATIONAL, FIDT_RATIONAL
                        'tmpTag = FreeImage_CreateTagExRational(FIMD_EXIF_RAW, fiMetadata(fiMetadataType, j).Key, fiMetadata(fiMetadataType, j).Type, fiMetadata(fiMetadataType, j).RationalValue(), fiMetadata(fiMetadataType, j).Count, fiMetadata(fiMetadataType, j).Id)
                        'tmpTag = FreeImage_CreateTagExRational(fiMetadataType, fiMetadata(fiMetadataType, j).Key, fiMetadata(fiMetadataType, j).Type, fiMetadata(fiMetadataType, j).RationalValue(), fiMetadata(fiMetadataType, j).Count, fiMetadata(fiMetadataType, j).Id)
                        tmpTag.TagPtr = 0
                    
                    Case Else
                        'tmpTag = FreeImage_CreateTagEx(fiMetadataType, fiMetadata(fiMetadataType, j).Key, fiMetadata(fiMetadataType, j).Type, fiMetadata(fiMetadataType, j).Value, fiMetadata(fiMetadataType, j).Count, fiMetadata(fiMetadataType, j).Id)
                        tmpTag.TagPtr = 0
                        
                End Select
                
                If tmpTag.TagPtr <> 0 Then
                    
                    'Attempt to set this tag to the FreeImage handle we were passed
                    If FreeImage_SetMetadataEx(fiImageHandle, tmpTag, tmpTag.Key, FIMD_XMP, True) Then
                    'If FreeImage_SetMetadataEx(fiImageHandle, tmpTag, fiMetadata(fiMetadataType, j).Key, fiMetadataType, True) Then
                        
                        'Note that normally we would need to delete the tag object just created; however, the RefreshTag
                        ' parameter used above handles that for us.
                        
                    Else
                        totalSuccess = False
                        Message "Failed to set metadata tag to FreeImage object."
                    End If
                    
                Else
                    totalSuccess = False
                    Message "Failed to create metadata tag."
                End If
                
            Next j
        End If
    
    Next fiMetadataType
    
    'Now, for the hell of it, we are going to attempt to retrieve the full raw metadata buffer.  I don't know if this will exist, but it's worth a try...
    loadAllMetadata fiImageHandle
    testMetadata
    'MsgBox fiMetadata(11, 0).StringValue
    
    writeAllMetadata = totalSuccess
    
End Function

'Given the name of an EXIF property, return the corresponding XMP title.  This can be used to convert EXIF properties to XMP format.
Private Function getXMPfromEXIF(ByVal exifName As String) As String

    Select Case exifName
    
        Case "Make"
            getXMPfromEXIF = "tiff:Make"
            
        Case "Model"
            getXMPfromEXIF = "tiff:Model"
            
        Case Else
            getXMPfromEXIF = ""
    
    End Select
    
End Function

Private Sub Class_Initialize()
    maxMetadataCount = 0
End Sub
