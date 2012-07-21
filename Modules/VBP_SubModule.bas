Attribute VB_Name = "Miscellaneous"
'***************************************************************************
'Miscellaneous Operations Handler
'Copyright ©2000-2012 by Tanner Helland
'Created: 6/12/01
'Last updated: 05/July/12
'Last update: Moved all string functions to a dedicated module.
'
'Handles messaging, value checking, RGB Extraction, checking if files exist,
' variable truncation, and old array transfer routines.
'
'***************************************************************************

Option Explicit

'Straight from MSDN - generate a "browse for folder" dialog
Public Function BrowseForFolder(ByVal srcHwnd As Long) As String
    
    Dim objShell As Shell
    Dim objFolder As Folder
    Dim returnString As String
        
    Set objShell = New Shell
    Set objFolder = objShell.BrowseForFolder(srcHwnd, "Please select a folder:", 0)
            
    If (Not objFolder Is Nothing) Then returnString = objFolder.Items.Item.Path Else returnString = ""
    
    Set objFolder = Nothing
    Set objShell = Nothing
    
    BrowseForFolder = returnString
    
End Function

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
    
    'If we're logging program messages, open up a log file and dump the message there
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
