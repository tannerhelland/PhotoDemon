Attribute VB_Name = "Miscellaneous"
'***************************************************************************
'Miscellaneous Operations Handler
'©2000-2012 Tanner Helland
'Created: 6/12/01
'Last updated: 10/June/12
'Last update: Get rid of SetFormTopMost. It's obnoxious as hell when other programs force forms on top,
'             so I definitely don't want to emulate the behavior.
'
'Handles messaging, value checking, RGB Extraction, checking if files exist,
'variable truncation, and the old array transfer routines.
'
'***************************************************************************

Option Explicit

'Straight from MSDN (with some help from PSC) - generate a "browse for folder" dialog
Public Function BrowseForFolder(ByVal srcHwnd As Long) As String
    Dim objShell   As Shell
    Dim objFolder  As Folder
    Dim returnString As String
        
    Set objShell = New Shell
        Set objFolder = objShell.BrowseForFolder(srcHwnd, "Please select a folder:", 0)
            If (Not objFolder Is Nothing) Then
                returnString = objFolder.Items.Item.Path
            Else
                returnString = ""
            End If
        Set objFolder = Nothing
    Set objShell = Nothing
    BrowseForFolder = returnString
End Function

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

'These two routines make it easier to interact with the progress bar; note that they are disabled while a batch
' conversion is running - this is because the batch conversion tool appropriates the scroll bar for itself
Public Sub SetProgBarMax(ByVal val As Long)
    If MacroStatus <> MacroBATCH Then cProgBar.Max = val
End Sub

Public Sub SetProgBarVal(ByVal val As Long)
    If MacroStatus <> MacroBATCH Then
        cProgBar.Value = val
        cProgBar.Draw
    End If
End Sub

'Display the current mouse position in the main form's status bar
Public Sub SetBitmapCoordinates(ByVal X1 As Long, ByVal Y1 As Long)
    Dim ZoomVal As Single
    ZoomVal = Zoom.ZoomArray(FormMain.CmbZoom.ListIndex)
    X1 = FormMain.ActiveForm.HScroll.Value + Int(X1 / ZoomVal)
    Y1 = FormMain.ActiveForm.VScroll.Value + Int(Y1 / ZoomVal)
    FormMain.lblCoordinates.Caption = "(" & X1 & "," & Y1 & ")"
    DoEvents
End Sub

'Display the specified size in the main form's status bar
Public Sub DisplaySize(ByVal iWidth As Long, ByVal iHeight As Long)
    FormMain.lblImgSize.Caption = "Size: " & iWidth & "x" & iHeight
    DoEvents
End Sub

'This popular function is used to display a message in the main form's status bar
Public Sub Message(ByVal MString As String)
    If MacroStatus = MacroSTART Then MString = MString & " {-Recording-}"
    If MacroStatus <> MacroBATCH Then
        If FormMain.Visible = True Then
            cProgBar.Text = MString
            cProgBar.Draw
        End If
    End If
    
    Debug.Print MString
    
    'If we're logging program messages, open up a log file and dump the message in there
    If LogProgramMessages = True Then
        Dim fileNum As Integer
        fileNum = FreeFile
        Open ProgramPath & PROGRAMNAME & "_DebugMessages.log" For Append As #fileNum
            Print #fileNum, MString
            If MString = "Finished." Then Print #fileNum, vbCrLf
        Close #fileNum
    End If
    
End Sub

'A pleasant combination of RangeValid and NumberValid
Public Function EntryValid(ByVal check As Variant, ByVal Min As Long, ByVal Max As Long, Optional ByVal displayNumError As Boolean = True, Optional ByVal displayRangeError As Boolean = True) As Boolean
    If Not IsNumeric(check) Then
        If displayNumError = True Then MsgBox check & " is not a valid entry." & vbCrLf & "Please enter a numeric value.", vbCritical + vbOKOnly + vbApplicationModal, PROGRAMNAME
        EntryValid = False
    Else
        If (check >= Min) And (check <= Max) Then
            EntryValid = True
        Else
            If displayRangeError = True Then MsgBox check & " is not a valid entry." & vbCrLf & "Value must be between " & Min & " and " & Max & ".", vbCritical + vbOKOnly + vbApplicationModal, PROGRAMNAME
            EntryValid = False
        End If
    End If
End Function

Public Function RangeValid(ByVal check As Long, ByVal Min As Long, ByVal Max As Long) As Boolean
    If (check >= Min) And (check <= Max) Then
        RangeValid = True
    Else
        MsgBox check & " is not a valid entry.  Value must be between " & Min & " and " & Max & ".", vbCritical + vbOKOnly + vbApplicationModal, PROGRAMNAME
        RangeValid = False
    End If
End Function

Public Function NumberValid(ByVal check) As Boolean
    If Not IsNumeric(check) Then
        MsgBox check & " is not a valid entry.  Please enter a numeric value.", vbCritical + vbOKOnly + vbApplicationModal, PROGRAMNAME
        NumberValid = False
    Else
        NumberValid = True
    End If
End Function

Public Function ExtractR(ByVal CurrentColor As Long) As Integer
    ExtractR = CurrentColor Mod 256
End Function

Public Function ExtractG(ByVal CurrentColor As Long) As Integer
    ExtractG = (CurrentColor \ 256) And 255
End Function

Public Function ExtractB(ByVal CurrentColor As Long) As Integer
    ExtractB = (CurrentColor \ 65536) And 255
End Function

Public Sub FinishUp()
    SetProgBarVal 0
    Message "Finished."
    FormMain.ActiveForm.BackBuffer.Picture = FormMain.ActiveForm.BackBuffer.Image
    ScrollViewport FormMain.ActiveForm
End Sub

'Convert to absolute byte values (Integer-type)
Public Sub ByteMe(ByRef TempVar As Integer)
    If TempVar > 255 Then TempVar = 255
    If TempVar < 0 Then TempVar = 0
End Sub
'Convert to absolute byte values (Long-type)
Public Sub ByteMeL(ByRef TempVar As Long)
    If TempVar > 255 Then TempVar = 255
    If TempVar < 0 Then TempVar = 0
End Sub

'Returns a boolean as to whether or not fName exists
Public Function FileExist(fName As String) As Boolean
    On Error Resume Next
    Dim Temp As Long
    Temp = GetAttr(fName)
    FileExist = Not CBool(Err)
End Function

'Returns a boolean as to whether or not dName exists
Public Function DirectoryExist(dName As String) As Boolean
    On Error Resume Next
    Dim Temp As Long
    Temp = GetAttr(dName) And vbDirectory
    DirectoryExist = Not CBool(Err)
End Function

'Blend byte1 w/ byte2 based on Percent (integer)
Public Function MixColors(ByVal Color1 As Byte, ByVal Color2 As Byte, ByVal Percent1 As Integer) As Integer
    MixColors = (((100 - Percent1) * Color1) + (Percent1 * Color2)) * 0.01
End Function

'Pass this a text box and it will select all text currently in the text box
Public Function AutoSelectText(ByRef tBox As TextBox)
    tBox.SelStart = 0
    tBox.SelLength = Len(tBox.Text)
End Function

'This lovely function comes from "penagate"; it was downloaded from http://www.vbforums.com/showthread.php?t=342995 on 08 June '12
Public Function GetDomainName(ByVal Address As String) As String
        Dim strOutput       As String
        Dim strTemp         As String
        Dim lngLoopCount    As Long
        Dim lngBCount       As Long
        Dim lngCharCount    As Long
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
