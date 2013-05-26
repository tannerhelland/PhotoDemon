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
    isBinary As Boolean
End Type

'Array that contains the metadata itself
Private iMetadata() As mdItem

'When extracting metadata, we request a full XML dump from ExifTool.  This XML file is stored - in its entirety - in memory,
' as it is the only way to precisely restore the XML data in its *exact* form when save-time comes.
Private originalMetadataXML As String

'If GPS data is found, mark its location
Private hasGPSCoordinates As Boolean
Private gpsIndex As Long

'If no metadata was found during the original FreeImage load process, this will return FALSE.  If any metadata of any
' compatible type was found, this will return TRUE.
Public Function hasMetadata() As Boolean
    If maxMetadataCount = 0 Then hasMetadata = False Else hasMetadata = True
End Function

'Some metadata options require the presence of GPS data in the photo.  If a GPS entry was found at metadata load-time, return TRUE.
Public Function hasGPSMetadata() As Boolean
    hasGPSMetadata = hasGPSCoordinates
End Function

'If GPS metadata exists, this function can be used to retrieve the exact latitude and longitude values
Public Function fillLatitudeLongitude(ByRef latString As String, ByRef lonString As String) As Boolean
    
    If hasGPSCoordinates Then
    
        Dim tmpSplit() As String
        tmpSplit = Split(iMetadata(gpsIndex).TechnicalValue, " ")
        latString = tmpSplit(0)
        lonString = tmpSplit(1)
    
        fillLatitudeLongitude = True
    
    Else
        fillLatitudeLongitude = False
    End If
    
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
    hasGPSCoordinates = False: gpsIndex = 0
    
    'ExifTool will return all metadata in one enormous string, which we will have to parse manually.
    Dim mdString As String
    mdString = getMetadata(srcFile, srcFormat)
    
    'If no metadata was found, we can safely exit now
    If mdString = "" Then Exit Sub
    
    'We now have a (very large) string of XML-formatted metadata.  Before doing anything else, make a backup copy of that string.
    ' We need this backup when it comes time to write the metadata back out to file.
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
    
    Dim tmpLocation As Long
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
            iMetadata(maxMetadataCount).Group = Trim$(Right$(tmpSplitInner(0), Len(tmpSplitInner(0)) - 2))
            iMetadata(maxMetadataCount).Name = Trim$(tmpSplitInner(1))
                        
            'A description line always follows a group/name line.  Retrieve that value now.
            curLine = curLine + 1
            xmlLines(curLine) = Trim$(xmlLines(curLine))
            iMetadata(maxMetadataCount).Description = Trim$(Mid$(xmlLines(curLine), 10, InStr(1, xmlLines(curLine), "</et") - 10))
            
            'A human-friendly value *almost* always follows a description line.  Check to see if this is a binary data line;
            ' if it is, retrieve it an mark is as binary type.  Then proceed to retrieve the data (binary or otherwise).
            curLine = curLine + 1
            xmlLines(curLine) = Trim$(xmlLines(curLine))
            
            'Non-binary data will have no rdf identifier in the tag; use that to separate binary from non-binary
            If InStr(1, xmlLines(curLine), "<et:prt>") > 0 Then iMetadata(maxMetadataCount).isBinary = False Else iMetadata(maxMetadataCount).isBinary = True
            
            If iMetadata(maxMetadataCount).isBinary Then
                'Binary data can be enormous.  ExifTool is kind enough to always start it on a new line for us - so move down a line,
                ' then run a specialized check for the end of the binary chunk.
                Do
                    curLine = curLine + 1
                Loop While StrComp(xmlLines(curLine), "</et:prt>") <> 0
            Else
                iMetadata(maxMetadataCount).Value = Trim$(Mid$(xmlLines(curLine), 9, InStr(1, xmlLines(curLine), "</et") - 9))
            End If
            
            curLine = curLine + 1
            
            'For some tags, the human-friendly value will be the actual value of the tag.  For others, the tag will have some kind
            ' of obscure value that ExifTool has kindly translated for us.  We still want to retrieve the technical value if
            ' available - so do that now.
            
            'Unfortunately, like the base Value tag, the *actual* value tag can also come in a binary variant.  Check for that now.
            ' If a fully closed tag is found, we know this is of the (far more common) non-binary variety
            If InStr(1, xmlLines(curLine), "<et:val>") > 0 Then
                
                xmlLines(curLine) = Trim$(xmlLines(curLine))
                iMetadata(maxMetadataCount).TechnicalValue = Trim$(Mid$(xmlLines(curLine), 9, InStr(1, xmlLines(curLine), "</et") - 9))
                curLine = curLine + 1
                
            'If the tag exists but is not closed, we know it is in binary format
            ElseIf InStr(1, xmlLines(curLine), "<et:val") > 0 Then
            
                 'Binary data can be enormous.  ExifTool is kind enough to always start it on a new line for us - so move down a line,
                ' then run a specialized check for the end of the binary chunk.
                Do
                    curLine = curLine + 1
                Loop While StrComp(xmlLines(curLine), "</et:val>") <> 0
                
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


'Given a path to a saved image (yes, that's right - make certain it is already saved!) and a path to the original file, to a FreeImage object, write out all currently stored metadata.  (This only needs to be done immediately prior to a save.)
Public Function writeAllMetadata(ByVal outputPath As String, ByVal originalFile As String) As Boolean
    
End Function

Private Sub Class_Initialize()
    maxMetadataCount = 0
End Sub
