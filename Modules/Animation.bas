Attribute VB_Name = "Animation"
'***************************************************************************
'Animation Functions
'Copyright 2019-2020 by Tanner Helland
'Created: 20/August/19
'Last updated: 20/August/19
'Last update: migrate animation code from animated GIF engine to here, since we're going to reuse bits of
'             it for animated PNGs.
'
'PhotoDemon was never meant to be an animation editor, but repeat user requests for animated GIF handling
' led to a rudimentary set of import/export/playback features.
'
'This module collects a few useful tools for dealing with animated images.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Public Function GetFrameTimeFromLayerName(ByRef srcLayerName As String, Optional ByVal defaultTimeIfMissing As Long = 100) As Long
    
    'Default to 100 ms, per convention
    GetFrameTimeFromLayerName = defaultTimeIfMissing
    
    'Look for a trailing parenthesis
    Dim strStart As Long, strEnd As Long
    strEnd = InStrRev(srcLayerName, ")", -1, vbBinaryCompare)
    If (strEnd > 0) Then
        
        'Find the nearest leading parenthesis
        strStart = InStrRev(srcLayerName, "(", strEnd, vbBinaryCompare)
        If (strStart > 0) And (strStart < strEnd - 1) Then
        
            'Retrieve the text between said parentheses
            Dim tmpString As String
            tmpString = Mid$(srcLayerName, strStart + 1, (strEnd - strStart) - 1)
            
            'Finally, strip any non-numeric characters from the text.  (Frame times are typically displayed
            ' as "100ms", and we don't want the "ms" bit.)
            Dim ascLow As Long, ascHigh As Long
            ascLow = AscW("0")
            ascHigh = AscW("9")
            
            Dim finalString As pdString
            Set finalString = New pdString
            
            Dim i As Long, singleChar As String
            For i = 1 To Len(tmpString)
                singleChar = Mid$(tmpString, i, 1)
                If (AscW(singleChar) >= ascLow) And (AscW(singleChar) <= ascHigh) Then finalString.Append singleChar
            Next i
            
            On Error GoTo BadNumber
            GetFrameTimeFromLayerName = CLng(finalString.ToString())
            
            'Enforce a minimum frametime of 0 ms, and leave it to decoders to deal with that case as necessary
            If (GetFrameTimeFromLayerName < 0) Then GetFrameTimeFromLayerName = 0
            
BadNumber:
        
        End If
        
    End If
    
End Function

