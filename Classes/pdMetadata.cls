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
'Last updated: 25/May/13
'Last update: converted storage model to use ExifTool output instead of FreeImage
'
'As of version 5.6, PhotoDemon now provides support for loading and saving image metadata.  What is metadata, you ask?
' See http://en.wikipedia.org/wiki/Metadata#Photographs for more details.
'
'Because metadata handling is complicated, I have wrapped all related functions into this class.  Each pdImage object
' has one (1) metadata class called "imgMetadata".  That class is publicly accessible.  If metadata is available at
' load-time, the imgMetadata object will automatically be filled with any metadata found in the source file.
'
'At present, metadata handling is 100% tied to the ExifTool plugin (see the "Plugin_ExifTool_Interface" module for more
' detail).  Because metadata is insanely complex to both read and write - if you want it done correctly - I have no
' plans to write a fallback for users without the plugin.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'The number of metadata items currently stored
Private maxMetadataCount As Long

Private Type mdItem
    Group As String
    Name As String
    Description As String
    Value As String
    TechnicalValue As String
End Type

'Array that contains the metadata itself
Private iMetadata() As mdItem

'If no metadata was found during the original FreeImage load process, this will return FALSE.  If any metadata of any
' compatible type was found, this will return TRUE.
Public Function hasMetadata() As Boolean
    If maxMetadataCount = 0 Then hasMetadata = False Else hasMetadata = True
End Function

'DO NOTE USE THIS FUNCTION IN PRODUCTION!  I include it here simply for ensuring the metadata was loaded correctly.
Public Sub testMetadata()

    Dim msgString As String

    If maxMetadataCount = 0 Then Exit Sub

    'For each metadata object, display a message box with the relevant metadata value
    Dim i As Long
    For i = 0 To maxMetadataCount - 1
    
        msgString = "Tag " & (i + 1) & " of " & maxMetadataCount & vbCrLf & vbCrLf
        msgString = msgString & "Group: " & iMetadata(i).Group & vbCrLf
        msgString = msgString & "Name: " & iMetadata(i).Name & vbCrLf
        msgString = msgString & "Description: " & iMetadata(i).Description & vbCrLf
        msgString = msgString & "Value: " & iMetadata(i).Value & vbCrLf
        msgString = msgString & "Technical value: " & iMetadata(i).TechnicalValue & vbCrLf
        MsgBox msgString
        
    Next i

End Sub

'Given an image path, attempt to load all metadata from that file.
' (The image's format is also passed to this function, which may be helpful in the future for enabling certain options on a per-format basis.)
Public Sub loadAllMetadata(ByVal srcFile As String, ByVal srcFormat As Long)

    'If the ExifTool plugin can't be found, abandon all metadata loading attempts
    If Not g_ExifToolEnabled Then
        Message "ExifTool plugin missing.  Metadata import abandoned."
        Exit Sub
    End If

    maxMetadataCount = 0
        
    'ExifTool will return all metadata in one enormous string, which we will have to parse manually.
    Dim mdString As String
    mdString = getMetadata(srcFile, srcFormat)

    'If no metadata was found, we can safely exit now
    If mdString = "" Then Exit Sub
    
    'We now have a (very large) string of XML-formatted metadata.  Our next job is to parse that data into the iMetadata() array.
    
    'Our next goal is to step through the XML, looking for tag/value pairs as we go.  Because the XML is precisely structured, we
    ' can rely on a few key markers to do this successfully.
    
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
    
    ReDim iMetadata(0) As mdItem
    
    Do
    
        'If our definitive text is found, parse this line to retrieve this tag's Group and Name values
        If InStr(1, xmlLines(curLine), "rdf:parseType='Resource'", vbBinaryCompare) > 0 Then
            
            ignoreThisTag = False
        
            'Divide the tag into two segments: the prefix (which contains Group and Name) and the rdf parse-type, which we don't care about.
            tmpSplitOuter = Split(xmlLines(curLine), "rdf:parseType='Resource")
            
            'Now, divide the prefix into two segments: Group and Name
            tmpSplitInner = Split(tmpSplitOuter(0), ":")
            
            'It is now trivial to retrieve the Group and Name values
            iMetadata(maxMetadataCount).Group = Right$(tmpSplitInner(0), Len(tmpSplitInner(0)) - 2)
            iMetadata(maxMetadataCount).Name = tmpSplitInner(1)
                        
            'A description line always follows a group/name line.  Retrieve that value now.
            curLine = curLine + 1
            xmlLines(curLine) = Trim$(xmlLines(curLine))
            iMetadata(maxMetadataCount).Description = Mid$(xmlLines(curLine), 10, InStr(1, xmlLines(curLine), "</et") - 10)
            
            'A human-friendly value always follows a description line.  Retrieve that value now.
            curLine = curLine + 1
            xmlLines(curLine) = Trim$(xmlLines(curLine))
            iMetadata(maxMetadataCount).Value = Mid$(xmlLines(curLine), 9, InStr(1, xmlLines(curLine), "</et") - 9)
            
            curLine = curLine + 1
            
            'For some tags, the human-friendly value will be the actual value of the tag.  For others, the tag will have some kind
            ' of obscure value that ExifTool has kindly translated for us.  We still want to retrieve the technical value if
            ' available - so do that now.
            If InStr(1, xmlLines(curLine), "<et:val") > 0 Then
                xmlLines(curLine) = Trim$(xmlLines(curLine))
                iMetadata(maxMetadataCount).TechnicalValue = Mid$(xmlLines(curLine), 9, InStr(1, xmlLines(curLine), "</et") - 9)
                curLine = curLine + 1
            End If
            
            'Now that we have fully parsed the tag, we need to run some additional checks on it.
            
            'ANALYZE GROUP
            
            'First, analyze the current Group.  Some groups can be ignored - check for those first.
            tmpString = iMetadata(maxMetadataCount).Group
            If StrComp(tmpString, "ExifTool") = 0 Then ignoreThisTag = True
            
            'Next, ExifTool provides VERY detailed group information.  I prefer to simplify its groups into something more manageable.
            
            'First, there's no need for a distinction between "System" and "File".  Group these under a "General" heading.
            If StrComp(tmpString, "System") = 0 Then iMetadata(maxMetadataCount).Group = "General"
            If StrComp(tmpString, "File") = 0 Then iMetadata(maxMetadataCount).Group = "General"
            
            'Next, ExifTool breaks down EXIF tags into their individual IDF blocks.  Simplify all those under a single "Exif" header.
            If InStr(1, tmpString, "IFD", vbTextCompare) > 0 Then iMetadata(maxMetadataCount).Group = "Exif"
            
            'Next, one of ExifTool's neatest features is its "Composite" tags - basically, tags that don't actually exist in the file,
            ' but which can be inferred from other information.  Let's rename these as "Intelligent", because it sounds better.  :)
            If StrComp(tmpString, "Composite") = 0 Then iMetadata(maxMetadataCount).Group = "Intelligent"
            
            'ANALYZE VALUE
            
            'Next come .Value checks.  Certain values may be invalid - if so, ignore this tag.
            tmpString = iMetadata(maxMetadataCount).Value
            
            'Binary data serves no purpose in the user interface part of our metadata handler.  Remove any binary tags.
            If InStr(1, tmpString, "Binary data", vbTextCompare) > 0 Then ignoreThisTag = True
            
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

    'If no metadata was found, exit now
    If maxMetadataCount = 0 Then Exit Sub

    'DEBUGGING ONLY: I found it helpful to dump the EXIF data to the clipboard while debugging, so I'm leaving this here.
    'Clipboard.Clear
    'Clipboard.SetText mdString
    
End Sub


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

Private Sub Class_Initialize()
    maxMetadataCount = 0
End Sub
