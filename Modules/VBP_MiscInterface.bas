Attribute VB_Name = "Misc_Interface"
'***************************************************************************
'Miscellaneous Functions Related to the User Interface
'Copyright ©2000-2012 by Tanner Helland
'Created: 6/12/01
'Last updated: 03/October/12
'Last update: First build
'***************************************************************************


Option Explicit

'Used to set the cursor for an object to the system's hand cursor
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal HWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long

Public Const IDC_APPSTARTING = 32650&
Public Const IDC_HAND = 32649&
Public Const IDC_ARROW = 32512&
Public Const IDC_CROSS = 32515&
Public Const IDC_IBEAM = 32513&
Public Const IDC_ICON = 32641&
Public Const IDC_NO = 32648&
Public Const IDC_SIZEALL = 32646&
Public Const IDC_SIZENESW = 32643&
Public Const IDC_SIZENS = 32645&
Public Const IDC_SIZENWSE = 32642&
Public Const IDC_SIZEWE = 32644&
Public Const IDC_UPARROW = 32516&
Public Const IDC_WAIT = 32514&

Private Const GCL_HCURSOR = (-12)

'These variables will hold the values of all custom-loaded cursors.
' They need to be deleted (via DestroyCursor) when the program exits; this is handled by unloadAllCursors.
Dim hc_Handle_Arrow As Long
Dim hc_Handle_Cross As Long
Dim hc_Handle_Hand As Long
Dim hc_Handle_SizeAll As Long
Dim hc_Handle_SizeNESW As Long
Dim hc_Handle_SizeNS As Long
Dim hc_Handle_SizeNWSE As Long
Dim hc_Handle_SizeWE As Long

'Because VB6 apps tend to look pretty lame on modern version of Windows, we do a bit of beautification to every form when
' it's loaded.  This routine is nice because every form calls it at least once, so we can make centralized changes without
' having to rewrite code in every individual form.
Public Sub makeFormPretty(ByRef tForm As Form)

    Dim eControl As Control
    
    For Each eControl In tForm.Controls
        
        'STEP 1: give all clickable controls a hand icon instead of the default pointer
        ' (Note: this code will set all command buttons, scroll bars, option buttons, check boxes,
        ' list boxes, combo boxes, and file/directory/drive boxes to use the system hand cursor)
        If ((TypeOf eControl Is CommandButton) Or (TypeOf eControl Is HScrollBar) Or (TypeOf eControl Is VScrollBar) Or (TypeOf eControl Is OptionButton) Or (TypeOf eControl Is CheckBox) Or (TypeOf eControl Is ListBox) Or (TypeOf eControl Is ComboBox) Or (TypeOf eControl Is FileListBox) Or (TypeOf eControl Is DirListBox) Or (TypeOf eControl Is DriveListBox)) And (Not TypeOf eControl Is PictureBox) Then
            eControl.MouseIcon = LoadPicture("")
            eControl.MousePointer = 0
            setHandCursor eControl
        End If
        
        'STEP 2: reset the .TabStop property of scroll bars.  This removes the obnoxious "flickering effect" in both
        ' the IDE, and for users of the crappy old "classic theme" under modern versions of Windows.
        If (TypeOf eControl Is HScrollBar) Or (TypeOf eControl Is VScrollBar) Then
            eControl.TabStop = False
        End If
        
    Next
    
End Sub

'Perform any drawing routines related to the main form
Public Sub RedrawMainForm()

    'Draw a subtle gradient on the left-hand pane
    FormMain.picLeftPane.Refresh
    DrawGradient FormMain.picLeftPane, RGB(240, 240, 240), RGB(201, 211, 226), True
    
    'Redraw the progress bar
    FormMain.picProgBar.Refresh
    cProgBar.Draw
    
End Sub

'Load all system cursors into memory
Public Sub InitAllCursors()

    hc_Handle_Arrow = LoadCursor(0, IDC_ARROW)
    hc_Handle_Cross = LoadCursor(0, IDC_CROSS)
    hc_Handle_Hand = LoadCursor(0, IDC_HAND)
    hc_Handle_SizeAll = LoadCursor(0, IDC_SIZEALL)
    hc_Handle_SizeNESW = LoadCursor(0, IDC_SIZENESW)
    hc_Handle_SizeNS = LoadCursor(0, IDC_SIZENS)
    hc_Handle_SizeNWSE = LoadCursor(0, IDC_SIZENWSE)
    hc_Handle_SizeWE = LoadCursor(0, IDC_SIZEWE)

End Sub

'Remove the hand cursor from memory
Public Sub unloadAllCursors()
    DestroyCursor hc_Handle_Hand
    DestroyCursor hc_Handle_Arrow
    DestroyCursor hc_Handle_Cross
    DestroyCursor hc_Handle_SizeAll
    DestroyCursor hc_Handle_SizeNESW
    DestroyCursor hc_Handle_SizeNS
    DestroyCursor hc_Handle_SizeNWSE
    DestroyCursor hc_Handle_SizeWE
End Sub

'Set a single object to use the hand cursor
Public Sub setHandCursor(ByRef tControl As Control)
    SetClassLong tControl.HWnd, GCL_HCURSOR, hc_Handle_Hand
End Sub

'Set a single form to use the arrow cursor
Public Sub setArrowCursor(ByRef tControl As Form)
    SetClassLong tControl.HWnd, GCL_HCURSOR, hc_Handle_Arrow
End Sub

'Set a single form to use the cross cursor
Public Sub setCrossCursor(ByRef tControl As Form)
    SetClassLong tControl.HWnd, GCL_HCURSOR, hc_Handle_Cross
End Sub
    
'Set a single form to use the Size All cursor
Public Sub setSizeAllCursor(ByRef tControl As Form)
    SetClassLong tControl.HWnd, GCL_HCURSOR, hc_Handle_SizeAll
End Sub

'Set a single form to use the Size NESW cursor
Public Sub setSizeNESWCursor(ByRef tControl As Form)
    SetClassLong tControl.HWnd, GCL_HCURSOR, hc_Handle_SizeNESW
End Sub

'Set a single form to use the Size NS cursor
Public Sub setSizeNSCursor(ByRef tControl As Form)
    SetClassLong tControl.HWnd, GCL_HCURSOR, hc_Handle_SizeNS
End Sub

'Set a single form to use the Size NWSE cursor
Public Sub setSizeNWSECursor(ByRef tControl As Form)
    SetClassLong tControl.HWnd, GCL_HCURSOR, hc_Handle_SizeNWSE
End Sub

'Set a single form to use the Size WE cursor
Public Sub setSizeWECursor(ByRef tControl As Form)
    SetClassLong tControl.HWnd, GCL_HCURSOR, hc_Handle_SizeWE
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
    
    If IsProgramCompiled = False Then Debug.Print MString
    
    'If we're logging program messages, open up a log file and dump the message there
    If LogProgramMessages = True Then
        Dim fileNum As Integer
        fileNum = FreeFile
        Open DataPath & PROGRAMNAME & "_DebugMessages.log" For Append As #fileNum
            Print #fileNum, MString
            If MString = "Finished." Then Print #fileNum, vbCrLf
        Close #fileNum
    End If
    
End Sub

'Pass AutoSelectText a text box and it will select all text currently in the text box
Public Function AutoSelectText(ByRef tBox As TextBox)
    tBox.SetFocus
    tBox.SelStart = 0
    tBox.SelLength = Len(tBox.Text)
End Function

