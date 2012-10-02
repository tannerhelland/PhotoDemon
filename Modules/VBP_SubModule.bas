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
Private Const IDC_HAND As Long = 32649
Private Const GCL_HCURSOR = (-12)

'Used to convert a system color (such as "button face") to a literal RGB value
Private Declare Function TranslateColor Lib "OLEPRO32.DLL" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, col As Long) As Long
    
'This variable will hold the value of the loaded hand cursor.  We need to delete it (via DestroyCursor) when the program exits.
Dim hc_Handle As Long

'Given an OLE color, return an RGB
Public Function GetRealColor(ByVal Color As OLE_COLOR) As Long
    TranslateColor Color, 0, GetRealColor
End Function

'Validate a given number.
Public Sub textValidate(ByRef srcTextBox As TextBox, Optional ByVal negAllowed As Boolean = False, Optional ByVal floatAllowed As Boolean = False)

    'Convert the input number to a string
    Dim numString As String
    numString = srcTextBox.Text
    
    'Remove any incidental white space before processing
    numString = Trim(numString)
    
    'Create a string of valid numerical characters, based on the input specifications
    Dim validChars As String
    validChars = "0123456789"
    If negAllowed Then validChars = validChars & "-"
    If floatAllowed Then validChars = validChars & "."
    
    'Make note of the cursor position so we can restore it after removing invalid text
    Dim cursorPos As Long
    cursorPos = srcTextBox.SelStart
    
    'Loop through the text box contents and remove any invalid characters
    Dim i As Long, j As Long
    Dim invLoc As Long
    
    For i = 1 To Len(numString)
        
        'Compare a single character from the text box against our list of valid characters
        invLoc = InStr(validChars, Mid$(numString, i, 1))
        
        'If this character was NOT found in the list of valid characters, remove it from the string
        If invLoc = 0 Then
        
            numString = Left$(numString, i - 1) & Right$(numString, Len(numString) - i)
            
            'Modify the position of the cursor to match (so the text box maintains the same cursor position)
            If i >= (cursorPos - 1) Then cursorPos = cursorPos - 1
            
            'Move the loop variable back by 1 so the next character is properly checked
            i = i - 1
            
        End If
            
    Next i
        
    'Place the newly validated string back in the text box
    srcTextBox.Text = numString
    srcTextBox.Refresh
    srcTextBox.SelStart = cursorPos

End Sub

'Populate a text box with a given integer value.  This is done constantly across the program, so I use a sub to handle it, as
' there may be additional validations that need to be performed, and it's nice to be able to adjust those from a single location.
Public Sub copyToTextBoxI(ByRef dstTextBox As TextBox, ByVal srcValue As Long)

    'Remember the current cursor position
    Dim cursorPos As Long
    cursorPos = dstTextBox.SelStart

    'Overwrite the current text box value with the new value
    dstTextBox = CStr(srcValue)
    dstTextBox.Refresh
    
    'Restore the cursor to its original position
    If cursorPos >= Len(dstTextBox) Then cursorPos = Len(dstTextBox)
    dstTextBox.SelStart = cursorPos

End Sub

'Populate a text box with a given floating-point value.  This is done constantly across the program, so I use a sub to handle it, as
' there may be additional validations that need to be performed, and it's nice to be able to adjust those from a single location.
Public Sub copyToTextBoxF(ByVal srcValue As Double, ByRef dstTextBox As TextBox)

    'Remember the current cursor position
    Dim cursorPos As Long
    cursorPos = dstTextBox.SelStart

    'PhotoDemon never allows more than two significant digits for floating-point text boxes
    dstTextBox = Format(CStr(srcValue), "#0.00")
    dstTextBox.Refresh
    
    'Restore the cursor to its original position
    If cursorPos >= Len(dstTextBox) Then cursorPos = Len(dstTextBox)
    dstTextBox.SelStart = cursorPos

End Sub


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
Public Sub SetProgBarMax(ByVal pbVal As Long)
    If MacroStatus <> MacroBATCH Then cProgBar.Max = pbVal
End Sub

Public Function getProgBarMax() As Long
    getProgBarMax = cProgBar.Max
End Function

Public Sub SetProgBarVal(ByVal pbVal As Long)
    If MacroStatus <> MacroBATCH Then
        cProgBar.Value = pbVal
        cProgBar.Draw
    End If
End Sub

'Let a form know whether the mouse pointer is over its image or just the viewport
Public Function isMouseOverImage(ByVal x1 As Long, ByVal y1 As Long, ByRef srcForm As Form) As Boolean

    If (x1 >= pdImages(srcForm.Tag).targetLeft) And (x1 <= pdImages(srcForm.Tag).targetLeft + pdImages(srcForm.Tag).targetWidth) Then
        If (y1 >= pdImages(srcForm.Tag).targetTop) And (y1 <= pdImages(srcForm.Tag).targetTop + pdImages(srcForm.Tag).targetHeight) Then
            isMouseOverImage = True
            Exit Function
        Else
            isMouseOverImage = False
        End If
        isMouseOverImage = False
    End If

End Function

'Calculate and display the current mouse position.
' INPUTS: x and y coordinates of the mouse cursor, current form, and optionally two long-type variables to receive the relative
'          coordinates (e.g. location on the image) of the current mouse position.
Public Sub displayImageCoordinates(ByVal x1 As Single, ByVal y1 As Single, ByRef srcForm As Form, Optional ByRef copyX As Single, Optional ByRef copyY As Single)

    If isMouseOverImage(x1, y1, srcForm) Then
            
        'Grab the current zoom value
        Static ZoomVal As Single
        ZoomVal = Zoom.ZoomArray(pdImages(srcForm.Tag).CurrentZoomValue)
            
        'Calculate x and y positions, while taking into account zoom and scroll values
        x1 = srcForm.HScroll.Value + Int((x1 - pdImages(srcForm.Tag).targetLeft) / ZoomVal)
        y1 = srcForm.VScroll.Value + Int((y1 - pdImages(srcForm.Tag).targetTop) / ZoomVal)
            
        'When zoomed very far out, the values might be calculated incorrectly.  Force them to the image dimensions if necessary.
        If x1 < 0 Then x1 = 0
        If y1 < 0 Then y1 = 0
        If x1 > pdImages(srcForm.Tag).Width Then x1 = pdImages(srcForm.Tag).Width
        If y1 > pdImages(srcForm.Tag).Height Then y1 = pdImages(srcForm.Tag).Height
        
        'If the user has requested copies of these coordinates, assign them now
        If copyX Then copyX = x1
        If copyY Then copyY = y1
        
        FormMain.lblCoordinates.Caption = "(" & x1 & "," & y1 & ")"
        FormMain.lblCoordinates.Refresh
        'DoEvents
        
    End If
    
End Sub

'If an x or y location is NOT in the image, find the nearest coordinate that IS in the image
Public Sub findNearestImageCoordinates(ByRef x1 As Single, ByRef y1 As Single, ByRef srcForm As Form)

    'Grab the current zoom value
    Static ZoomVal As Single
    ZoomVal = Zoom.ZoomArray(pdImages(srcForm.Tag).CurrentZoomValue)

    'Calculate x and y positions, while taking into account zoom and scroll values
    x1 = srcForm.HScroll.Value + Int((x1 - pdImages(srcForm.Tag).targetLeft) / ZoomVal)
    y1 = srcForm.VScroll.Value + Int((y1 - pdImages(srcForm.Tag).targetTop) / ZoomVal)
    
    'Force any invalid values to their nearest matching point in the image
    If x1 < 0 Then x1 = 0
    If y1 < 0 Then y1 = 0
    If x1 > pdImages(srcForm.Tag).Width Then x1 = pdImages(srcForm.Tag).Width
    If y1 > pdImages(srcForm.Tag).Height Then y1 = pdImages(srcForm.Tag).Height

End Sub

'Display the specified size in the main form's status bar
Public Sub DisplaySize(ByVal iWidth As Long, ByVal iHeight As Long)
    FormMain.lblImgSize.Caption = "size: " & iWidth & "x" & iHeight
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

'Because VB6 apps tend to look pretty lame on modern version of Windows, we do a bit of beautification to every form when
' it's loaded.  This routine is nice because every form calls it at least once, so we can make centralized changes without
' having to rewrite code in every individual form.
Public Sub makeFormPretty(ByRef tForm As Form)

    'STEP 1: give all clickable controls a hand icon instead of the default pointer
    ' (Note: this code will set all command buttons, scroll bars, option buttons, check boxes, list boxes, combo boxes, and file/directory/drive boxes to use the system hand cursor)
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
