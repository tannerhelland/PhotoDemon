Attribute VB_Name = "Text_Processing"
'***************************************************************************
'Text Operations Module
'©2000-2012 Tanner Helland
'Created: 05/July/12
'Last updated: 05/July/12
'Last update: Initial build.  The "Miscellaneous" module was getting overrun
'             with text-related code, so it was time to create a module for
'             just those.
'
'Handles various string operations, mostly related to parsing and generating
' valid filenames and paths.
'
'***************************************************************************

Option Explicit

'Make sure the right backslash of a path is existant
Public Function FixPath(ByVal tempString As String) As String
    If Right(tempString, 1) <> "\" Then
        FixPath = tempString & "\"
    Else
        FixPath = tempString
    End If
End Function

'Pull the directory out of a filename
Public Sub StripDirectory(ByRef sString As String)
    For x = Len(sString) - 1 To 1 Step -1
        If (Mid(sString, x, 1) = "/") Or (Mid(sString, x, 1) = "\") Then
            sString = Left(sString, x)
            Exit Sub
        End If
    Next x
End Sub

'Pull the filename ONLY (no directory) off a path
Public Sub StripFilename(ByRef sString As String)
    For x = Len(sString) - 1 To 1 Step -1
        If (Mid(sString, x, 1) = "/") Or (Mid(sString, x, 1) = "\") Then
            sString = Right(sString, Len(sString) - x)
            Exit Sub
        End If
    Next x
End Sub

'Pull the filename & directory out WITHOUT any extension (but with the ".")
Public Sub StripOffExtension(ByRef sString As String)
    For x = Len(sString) - 1 To 1 Step -1
        If (Mid(sString, x, 1) = ".") Then
            sString = Left(sString, x - 1)
            Exit Sub
        End If
    Next x
End Sub

'Function to strip the extension from a filename (taken long ago from the Internet; thank you to whoever wrote it!)
Public Function GetExtension(FileName As String) As String
    Dim pathLoc As Long, extLoc As Long
    Dim i As Long, j As Long

    For i = Len(FileName) To 1 Step -1
        If Mid(FileName, i, 1) = "." Then
            extLoc = i
            For j = Len(FileName) To 1 Step -1
                If Mid(FileName, j, 1) = "\" Then
                    pathLoc = j
                    Exit For
                End If
            Next j
            Exit For
        End If
    Next i
    
    If pathLoc > extLoc Then
        GetExtension = ""
    Else
        If extLoc = 0 Then GetExtension = ""
        GetExtension = Mid(FileName, extLoc + 1, Len(FileName) - extLoc)
    End If
            
End Function

'Take a string and replace any invalid characters with "_"
Public Sub makeValidWindowsFilename(ByRef FileName As String)

    Dim strInvalidChars As String
    strInvalidChars = "/*?""<>|"
    
    Dim invLoc As Long
    
    For x = 1 To Len(strInvalidChars)
        invLoc = InStr(FileName, Mid$(strInvalidChars, x, 1))
        If invLoc <> 0 Then
            FileName = Left(FileName, invLoc - 1) & "_" & Right(FileName, Len(FileName) - invLoc)
        End If
    Next x

End Sub

'Remove the accelerator (e.g. "Ctrl+0") from the tail end of a string
Public Sub StripAcceleratorFromCaption(ByRef sString As String)
    For x = Len(sString) - 1 To 1 Step -1
        If (Mid(sString, x, 1) = vbTab) Then
            sString = Left(sString, x - 1)
            Exit Sub
        End If
    Next x
End Sub

'This lovely function comes from "penagate"; it was downloaded from http://www.vbforums.com/showthread.php?t=342995 on 08 June '12
Public Function GetDomainName(ByVal Address As String) As String
        
    Dim strOutput As String, strTemp As String
    Dim lngLoopCount As Long
    Dim lngBCount As Long, lngCharCount As Long
    
    strOutput$ = Replace(Address, "\", "/")
        
    lngCharCount = Len(strOutput)
    
    If (InStrB(1, strOutput, "/")) Then
        
        Do Until ((strTemp = "/") Or (lngLoopCount = lngCharCount))
            lngLoopCount = lngLoopCount + 1
            strTemp = Mid$(strOutput, lngBCount + 1, 1)
            lngBCount = lngBCount + 1
        Loop
        
    End If
        
    strOutput = Right$(strOutput, Len(strOutput) - lngBCount)
    lngBCount = 0
    strTemp = "/"
    
    If (InStrB(1, strOutput, "/")) Then
        
        Do Until strTemp <> "/"
            strTemp = Mid$(strOutput, lngBCount + 1, 1)
            If strTemp = "/" Then lngBCount = lngBCount + 1
        Loop
    
    End If
        
    strOutput = Right$(strOutput, Len(strOutput) - lngBCount)
    strOutput = Left$(strOutput, InStr(1, strOutput, "/", vbTextCompare) - 1)
    GetDomainName = strOutput

End Function

