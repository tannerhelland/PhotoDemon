Attribute VB_Name = "Text_Processing"
'***************************************************************************
'Text Operations Module
'Copyright ©2012-2013 by Tanner Helland
'Created: 05/July/12
'Last updated: 09/June/13
'Last update: rewrote file extension function to be more intelligent about extension length
'
'Handles various string operations, mostly related to parsing and generating valid filenames, extensions, and paths.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
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

'Given a full file path (path + name + extension), remove everything but the directory structure
Public Sub StripDirectory(ByRef sString As String)
    
    Dim x As Long
    
    For x = Len(sString) - 1 To 1 Step -1
        If (Mid(sString, x, 1) = "/") Or (Mid(sString, x, 1) = "\") Then
            sString = Left(sString, x)
            Exit Sub
        End If
    Next x
    
End Sub

'Given a full file path (path + name + extension), return everything but the directory structure
Public Function getDirectory(ByRef sString As String) As String
    
    Dim x As Long
    
    For x = Len(sString) - 1 To 1 Step -1
        If (Mid(sString, x, 1) = "/") Or (Mid(sString, x, 1) = "\") Then
            getDirectory = Left(sString, x)
            Exit Function
        End If
    Next x
    
End Function


'Pull the filename ONLY (no directory) off a path
Public Sub StripFilename(ByRef sString As String)
    
    Dim x As Long
    
    For x = Len(sString) - 1 To 1 Step -1
        If (Mid(sString, x, 1) = "/") Or (Mid(sString, x, 1) = "\") Then
            sString = Right(sString, Len(sString) - x)
            Exit Sub
        End If
    Next x
    
End Sub

'Return the filename chunk of a path
Public Function getFilename(ByVal sString As String) As String

    Dim i As Long
    
    For i = Len(sString) - 1 To 1 Step -1
        If (Mid(sString, i, 1) = "/") Or (Mid(sString, i, 1) = "\") Then
            getFilename = Right(sString, Len(sString) - i)
            Exit Function
        End If
    Next i
    
End Function

'Pull the filename & directory out WITHOUT any extension (but with the ".")
Public Sub StripOffExtension(ByRef sString As String)

    Dim x As Long

    For x = Len(sString) - 1 To 1 Step -1
        If (Mid(sString, x, 1) = ".") Then
            sString = Left(sString, x - 1)
            Exit Sub
        End If
    Next x
    
End Sub

'Function to strip the extension from a filename
Public Function GetExtension(sFile As String) As String
    
    Dim i As Long
    For i = Len(sFile) To 1 Step -1
    
        'If we find a path before we find an extension, return a blank string
        If (Mid(sFile, i, 1) = "\") Or (Mid(sFile, i, 1) = "/") Then
            GetExtension = ""
            Exit Function
        End If
        
        If Mid(sFile, i, 1) = "." Then
            GetExtension = Right$(sFile, Len(sFile) - i)
            Exit Function
        End If
    Next i
    
    'If we reach this point, no extension was found
    GetExtension = ""
            
End Function

'Take a string and replace any invalid characters with "_"
Public Sub makeValidWindowsFilename(ByRef FileName As String)

    Dim strInvalidChars As String
    strInvalidChars = "/*?""<>|"
    
    Dim invLoc As Long
    
    Dim x As Long
    For x = 1 To Len(strInvalidChars)
        invLoc = InStr(FileName, Mid$(strInvalidChars, x, 1))
        If invLoc <> 0 Then
            FileName = Left(FileName, invLoc - 1) & "_" & Right(FileName, Len(FileName) - invLoc)
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

