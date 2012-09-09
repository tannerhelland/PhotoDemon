Attribute VB_Name = "Miscellaneous"
'***************************************************************************
'Miscellaneous Operations Handler
'Copyright ©2000-2012 by Tanner Helland
'Created: 6/12/01
'Last updated: 12/August/12
'Last update: Built several subs dedicated to assigning the system's hand cursor to clickable controls.
'
'Handles messaging, value checking, RGB Extraction, checking if files exist,
' variable truncation, and old array transfer routines.
'
'***************************************************************************

Option Explicit

'Used to set the cursor for an object to the system's hand cursor
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal HWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Const IDC_HAND  As Long = 32649
Private Const GCL_HCURSOR = (-12)

'This variable will hold the value of the loaded hand cursor.  We need to delete it (via DestroyCursor) when the program exits.
Dim hc_Handle As Long


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

'These three routines make it easier to interact with the progress bar; note that two are disabled while a batch
' conversion is running - this is because the batch conversion tool appropriates the scroll bar for itself
Public Sub SetProgBarMax(ByVal val As Long)
    If MacroStatus <> MacroBATCH Then cProgBar.Max = val
End Sub

Public Function getProgBarMax() As Long
    getProgBarMax = cProgBar.Max
End Function

Public Sub SetProgBarVal(ByVal val As Long)
    If MacroStatus <> MacroBATCH Then
        cProgBar.Value = val
        cProgBar.Draw
    End If
End Sub

'Display the current mouse position in the main form's status bar
Public Sub SetBitmapCoordinates(ByVal x1 As Long, ByVal y1 As Long)
    Dim ZoomVal As Single
    ZoomVal = Zoom.ZoomArray(FormMain.CmbZoom.ListIndex)
    x1 = FormMain.ActiveForm.HScroll.Value + Int(x1 / ZoomVal)
    y1 = FormMain.ActiveForm.VScroll.Value + Int(y1 / ZoomVal)
    FormMain.lblCoordinates.Caption = "(" & x1 & "," & y1 & ")"
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
        If displayNumError = True Then MsgBox check & " is not a valid entry." & vbCrLf & "Please enter a numeric value.", vbCritical + vbOKOnly + vbApplicationModal, "Invalid entry"
        EntryValid = False
    Else
        If (check >= Min) And (check <= Max) Then
            EntryValid = True
        Else
            If displayRangeError = True Then MsgBox check & " is not a valid entry." & vbCrLf & "Please enter a value between " & Min & " and " & Max & ".", vbCritical + vbOKOnly + vbApplicationModal, "Invalid entry"
            EntryValid = False
        End If
    End If
End Function

Public Function RangeValid(ByVal check As Long, ByVal Min As Long, ByVal Max As Long) As Boolean
    If (check >= Min) And (check <= Max) Then
        RangeValid = True
    Else
        MsgBox check & " is not a valid entry.  Please enter a value between " & Min & " and " & Max & ".", vbCritical + vbOKOnly + vbApplicationModal, PROGRAMNAME
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

'Extract the red, green, or blue value from an RGB() Long
Public Function ExtractR(ByVal CurrentColor As Long) As Integer
    ExtractR = CurrentColor Mod 256
End Function

Public Function ExtractG(ByVal CurrentColor As Long) As Integer
    ExtractG = (CurrentColor \ 256) And 255
End Function

Public Function ExtractB(ByVal CurrentColor As Long) As Integer
    ExtractB = (CurrentColor \ 65536) And 255
End Function

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

'Returns a boolean as to whether or not a given file exists
Public Function FileExist(ByRef fName As String) As Boolean
    On Error Resume Next
    Dim Temp As Long
    Temp = GetAttr(fName)
    FileExist = Not CBool(Err)
End Function

'Returns a boolean as to whether or not a given directory exists
Public Function DirectoryExist(ByRef dName As String) As Boolean
    On Error Resume Next
    Dim Temp As Long
    Temp = GetAttr(dName) And vbDirectory
    DirectoryExist = Not CBool(Err)
End Function

'Blend byte1 w/ byte2 based on mixRatio.  mixRatio is expected to be a value between 0 and 1.
Public Function BlendColors(ByVal Color1 As Byte, ByVal Color2 As Byte, ByRef mixRatio As Single) As Byte
    BlendColors = ((1 - mixRatio) * Color1) + (mixRatio * Color2)
End Function

'Pass this a text box and it will select all text currently in the text box
Public Function AutoSelectText(ByRef tBox As TextBox)
    tBox.SetFocus
    tBox.SelStart = 0
    tBox.SelLength = Len(tBox.Text)
End Function

'Load the hand cursor into memory
Public Sub initHandCursor()
    hc_Handle = LoadCursor(0, IDC_HAND)
End Sub

'Remove the hand cursor from memory
Public Sub destroyHandCursor()
    DestroyCursor hc_Handle
End Sub

'Set all command buttons, scroll bars, option buttons, check boxes, list boxes, combo boxes, and file/directory/drive boxes to use the system hand cursor
Public Sub setHandCursorForAll(ByRef tForm As Form)

    Dim eControl As Control
    
    For Each eControl In tForm.Controls
        If ((TypeOf eControl Is CommandButton) Or (TypeOf eControl Is HScrollBar) Or (TypeOf eControl Is VScrollBar) Or (TypeOf eControl Is OptionButton) Or (TypeOf eControl Is CheckBox) Or (TypeOf eControl Is ListBox) Or (TypeOf eControl Is ComboBox) Or (TypeOf eControl Is FileListBox) Or (TypeOf eControl Is DirListBox) Or (TypeOf eControl Is DriveListBox)) And (Not TypeOf eControl Is PictureBox) Then
            eControl.MouseIcon = LoadPicture("")
            eControl.MousePointer = 0
            setHandCursor eControl
        End If
    Next
    
End Sub

'Set a single object ot use a particular hand cursor
Public Sub setHandCursor(ByRef tControl As Control)
    SetClassLong tControl.HWnd, GCL_HCURSOR, hc_Handle
End Sub
