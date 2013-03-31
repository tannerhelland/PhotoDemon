Attribute VB_Name = "Interface"
'***************************************************************************
'Miscellaneous Functions Related to the User Interface
'Copyright ©2000-2013 by Tanner Helland
'Created: 6/12/01
'Last updated: 03/October/12
'Last update: First build
'***************************************************************************


Option Explicit

'Experimental subclassing to fix background color problems
' Many thanks to pro VB programmer LaVolpe for this workaround for themed controls not respecting their owner's backcolor properly.
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetProp Lib "user32.dll" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32.dll" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function DefWindowProc Lib "user32.dll" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_PRINTCLIENT As Long = &H318
Private Const WM_PAINT As Long = &HF&
Private Const GWL_WNDPROC As Long = -4
Private Const WM_DESTROY As Long = &H2

'Used to set the cursor for an object to the system's hand cursor
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
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
Private hc_Handle_Arrow As Long
Private hc_Handle_Cross As Long
Public hc_Handle_Hand As Long       'The hand cursor handle is used by the jcButton control as well, so it is declared publicly.
Private hc_Handle_SizeAll As Long
Private hc_Handle_SizeNESW As Long
Private hc_Handle_SizeNS As Long
Private hc_Handle_SizeNWSE As Long
Private hc_Handle_SizeWE As Long

'Used to measure the expected length of a string
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, ByRef lpSize As POINTAPI) As Long

'These constants are used to toggle visibility of display elements.
Public Const VISIBILITY_TOGGLE As Long = 0
Public Const VISIBILITY_FORCEDISPLAY As Long = 1
Public Const VISIBILITY_FORCEHIDE As Long = 2

'Because VB6 apps tend to look pretty lame on modern version of Windows, we do a bit of beautification to every form when
' it's loaded.  This routine is nice because every form calls it at least once, so we can make centralized changes without
' having to rewrite code in every individual form.
Public Sub makeFormPretty(ByRef tForm As Form)

    'Before doing anything else, make sure the form's default cursor is set to an arrow
    tForm.MouseIcon = LoadPicture("")
    tForm.MousePointer = 0

    'FORM STEP 1: Enumerate through every control on the form.  We will be making changes on-the-fly on a per-control basis.
    Dim eControl As Control
    
    For Each eControl In tForm.Controls
        
        'STEP 1: give all clickable controls a hand icon instead of the default pointer.
        ' (Note: this code will set all command buttons, scroll bars, option buttons, check boxes,
        ' list boxes, combo boxes, and file/directory/drive boxes to use the system hand cursor)
        If ((TypeOf eControl Is CommandButton) Or (TypeOf eControl Is HScrollBar) Or (TypeOf eControl Is VScrollBar) Or (TypeOf eControl Is OptionButton) Or (TypeOf eControl Is CheckBox) Or (TypeOf eControl Is ListBox) Or (TypeOf eControl Is ComboBox) Or (TypeOf eControl Is FileListBox) Or (TypeOf eControl Is DirListBox) Or (TypeOf eControl Is DriveListBox)) And (Not TypeOf eControl Is PictureBox) Then
            setHandCursor eControl
        End If
        
        'STEP 2: if the current system is Vista or later, and the user has requested modern typefaces via Edit -> Preferences,
        ' redraw all control fonts using Segoe UI.
        If g_IsVistaOrLater And ((TypeOf eControl Is TextBox) Or (TypeOf eControl Is CommandButton) Or (TypeOf eControl Is OptionButton) Or (TypeOf eControl Is CheckBox) Or (TypeOf eControl Is ListBox) Or (TypeOf eControl Is ComboBox) Or (TypeOf eControl Is FileListBox) Or (TypeOf eControl Is DirListBox) Or (TypeOf eControl Is DriveListBox) Or (TypeOf eControl Is Label)) And (Not TypeOf eControl Is PictureBox) Then
            If g_UseFancyFonts Then
                eControl.FontName = "Segoe UI"
            Else
                eControl.FontName = "Tahoma"
            End If
        End If
        
        If g_IsVistaOrLater And ((TypeOf eControl Is jcbutton) Or (TypeOf eControl Is smartOptionButton) Or (TypeOf eControl Is smartCheckBox)) Then
            If g_UseFancyFonts Then
                eControl.Font.Name = "Segoe UI"
            Else
                eControl.Font.Name = "Tahoma"
            End If
        End If
                        
        'STEP 3: remove TabStop from each picture box.  They should never receive focus, but I often forget to change this
        ' at design-time.
        If (TypeOf eControl Is PictureBox) Then eControl.TabStop = False
                        
        'STEP 4: correct tab stops so that the OK button is always 0, and Cancel is always 1
        If (TypeOf eControl Is CommandButton) Then
            If (eControl.Caption = "&OK") Then
                If eControl.TabIndex <> 0 Then eControl.TabIndex = 0
            End If
        End If
        If (TypeOf eControl Is CommandButton) Then
            If (eControl.Caption = "&Cancel") Then eControl.TabIndex = 1
        End If
                
    Next
    
    'FORM STEP 2: subclass this form and force controls to render transparent borders properly.
    If g_IsProgramCompiled Then SubclassFrame tForm.hWnd, False
    
    'FORM STEP 3: translate the form (and all controls on it)
    If g_Language.translationActive And tForm.Enabled Then
        g_Language.activateShortcut tForm.Name
        g_Language.applyTranslations tForm
        g_Language.deactivateShortcut
    End If
    
    'Refresh all non-MDI forms after making the changes above
    If tForm.Name <> "FormMain" Then
        tForm.Refresh
    Else
        'The main from is a bit different - if it has been translated or changed, it needs menu icons reassigned.
        If FormMain.Visible Then ApplyAllMenuIcons
    End If
        
End Sub

'This sub is used to render control backgrounds as transparent
Public Sub SubclassFrame(FramehWnd As Long, ReleaseSubclass As Boolean)
    Dim prevProc As Long

    prevProc = GetProp(FramehWnd, "scPproc")
    If ReleaseSubclass Then
        If prevProc Then
            SetWindowLong FramehWnd, GWL_WNDPROC, prevProc
            SetProp FramehWnd, "scPproc", 0&
        End If
    ElseIf prevProc = 0& Then
        SetProp FramehWnd, "scPproc", GetWindowLong(FramehWnd, GWL_WNDPROC)
        SetWindowLong FramehWnd, GWL_WNDPROC, AddressOf WndProc_Frame
    End If
End Sub

Private Function WndProc_Frame(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Dim prevProc As Long
    
    prevProc = GetProp(hWnd, "scPproc")
    If prevProc = 0& Then
        WndProc_Frame = DefWindowProc(hWnd, uMsg, wParam, lParam)
    ElseIf uMsg = WM_PRINTCLIENT Then
        SendMessage hWnd, WM_PAINT, wParam, ByVal 0&
    Else
        If uMsg = WM_DESTROY Then SubclassFrame hWnd, True
        WndProc_Frame = CallWindowProc(prevProc, hWnd, uMsg, wParam, lParam)
    End If
End Function

'The next two subs can be used to show or hide the left and right toolbar panes.  An input parameter can be specified to force behavior.
' INPUT VALUES:
' 0 (or none) - toggle the visibility to the opposite state (const VISIBILITY_TOGGLE)
' 1 - make the pane visible                                 (const VISIBILITY_FORCEDISPLAY)
' 2 - hide the pane                                         (const VISIBILITY_FORCEHIDE)
Public Sub ChangeLeftPane(Optional ByVal howToToggle As Long = 0)

    Select Case howToToggle
    
        Case VISIBILITY_TOGGLE
        
            'Write the new value to the INI
            g_UserPreferences.SetPreference_Boolean "General Preferences", "HideLeftPanel", Not g_UserPreferences.GetPreference_Boolean("General Preferences", "HideLeftPanel", False)

            'Toggle the text and picture box accordingly
            If g_UserPreferences.GetPreference_Boolean("General Preferences", "HideLeftPanel", False) Then
                FormMain.MnuLeftPanel.Caption = g_Language.TranslateMessage("Show left panel (file tools)")
                FormMain.picLeftPane.Visible = False
            Else
                FormMain.MnuLeftPanel.Caption = g_Language.TranslateMessage("Hide left panel (file tools)")
                FormMain.picLeftPane.Visible = True
            End If
    
            'Ask the menu icon handler to redraw the menu image with the new icon
            ResetMenuIcons
        
        Case VISIBILITY_FORCEDISPLAY
            FormMain.MnuLeftPanel.Caption = g_Language.TranslateMessage("Hide left panel (file tools)")
            FormMain.picLeftPane.Visible = True
            
        Case VISIBILITY_FORCEHIDE
            FormMain.MnuLeftPanel.Caption = g_Language.TranslateMessage("Show left panel (file tools)")
            FormMain.picLeftPane.Visible = False
            
    End Select

End Sub

Public Sub ChangeRightPane(Optional ByVal howToToggle As Long)

    Select Case howToToggle
    
        Case VISIBILITY_TOGGLE
        
            'Write the new value to the INI
            g_UserPreferences.SetPreference_Boolean "General Preferences", "HideRightPanel", Not g_UserPreferences.GetPreference_Boolean("General Preferences", "HideRightPanel", False)

            'Toggle the text and picture box accordingly
            If g_UserPreferences.GetPreference_Boolean("General Preferences", "HideRightPanel", False) Then
                FormMain.MnuRightPanel.Caption = g_Language.TranslateMessage("Show right panel (image tools)")
                FormMain.picRightPane.Visible = False
            Else
                FormMain.MnuRightPanel.Caption = g_Language.TranslateMessage("Hide right panel (image tools)")
                FormMain.picRightPane.Visible = True
            End If
    
            'Ask the menu icon handler to redraw the menu image with the new icon
            ResetMenuIcons
        
        Case VISIBILITY_FORCEDISPLAY
            FormMain.MnuRightPanel.Caption = g_Language.TranslateMessage("Hide right panel (image tools)")
            FormMain.picRightPane.Visible = True
            
        Case VISIBILITY_FORCEHIDE
            FormMain.MnuRightPanel.Caption = g_Language.TranslateMessage("Show right panel (image tools)")
            FormMain.picRightPane.Visible = False
            
    End Select

End Sub

'When a themed form is unloaded, it may be desirable to release certain changes made to it - or in our case, unsubclass it.
' This function should be called when any themed form is unloaded.
Public Sub ReleaseFormTheming(ByRef tForm As Form)
    If g_IsProgramCompiled Then SubclassFrame tForm.hWnd, True
    Set tForm = Nothing
End Sub

'Perform any drawing routines related to the main form
Public Sub RedrawMainForm()

    'Draw a subtle gradient on either pane if visible
    If FormMain.picLeftPane.Visible Then
        FormMain.picLeftPane.Refresh
        DrawGradient FormMain.picLeftPane, RGB(240, 240, 240), RGB(201, 211, 226), True
    End If
    
    If FormMain.picRightPane.Visible Then
        FormMain.picRightPane.Refresh
        DrawGradient FormMain.picRightPane, RGB(201, 211, 226), RGB(240, 240, 240), True
    End If
    
    'Redraw the progress bar
    FormMain.picProgBar.Refresh
    g_ProgBar.Draw
    
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
    tControl.MouseIcon = LoadPicture("")
    tControl.MousePointer = 0
    SetClassLong tControl.hWnd, GCL_HCURSOR, hc_Handle_Hand
End Sub

Public Sub setHandCursorToHwnd(ByVal dstHwnd As Long)
    SetClassLong dstHwnd, GCL_HCURSOR, hc_Handle_Hand
End Sub

'Set a single object to use the arrow cursor
Public Sub setArrowCursorToObject(ByRef tControl As Control)
    tControl.MouseIcon = LoadPicture("")
    tControl.MousePointer = 0
    SetClassLong tControl.hWnd, GCL_HCURSOR, hc_Handle_Arrow
End Sub

Public Sub setArrowCursorToHwnd(ByVal dstHwnd As Long)
    SetClassLong dstHwnd, GCL_HCURSOR, hc_Handle_Arrow
End Sub

'Set a single form to use the arrow cursor
Public Sub setArrowCursor(ByRef tControl As Form)
    SetClassLong tControl.hWnd, GCL_HCURSOR, hc_Handle_Arrow
End Sub

'Set a single form to use the cross cursor
Public Sub setCrossCursor(ByRef tControl As Form)
    SetClassLong tControl.hWnd, GCL_HCURSOR, hc_Handle_Cross
End Sub
    
'Set a single form to use the Size All cursor
Public Sub setSizeAllCursor(ByRef tControl As Form)
    SetClassLong tControl.hWnd, GCL_HCURSOR, hc_Handle_SizeAll
End Sub

'Set a single form to use the Size NESW cursor
Public Sub setSizeNESWCursor(ByRef tControl As Form)
    SetClassLong tControl.hWnd, GCL_HCURSOR, hc_Handle_SizeNESW
End Sub

'Set a single form to use the Size NS cursor
Public Sub setSizeNSCursor(ByRef tControl As Form)
    SetClassLong tControl.hWnd, GCL_HCURSOR, hc_Handle_SizeNS
End Sub

'Set a single form to use the Size NWSE cursor
Public Sub setSizeNWSECursor(ByRef tControl As Form)
    SetClassLong tControl.hWnd, GCL_HCURSOR, hc_Handle_SizeNWSE
End Sub

'Set a single form to use the Size WE cursor
Public Sub setSizeWECursor(ByRef tControl As Form)
    SetClassLong tControl.hWnd, GCL_HCURSOR, hc_Handle_SizeWE
End Sub

'Display the specified size in the main form's status bar
Public Sub DisplaySize(ByVal iWidth As Long, ByVal iHeight As Long)
    FormMain.lblImgSize.Caption = g_Language.TranslateMessage("size") & ": " & iWidth & "x" & iHeight
    FormMain.lblImgSize.Refresh
End Sub

'This wrapper is used in place of the standard MsgBox function.  At present it's just a wrapper around MsgBox, but
' in the future I may replace the dialog function with something custom.
Public Function pdMsgBox(ByVal pMessage As String, ByVal pButtons As VbMsgBoxStyle, ByVal pTitle As String, ParamArray ExtraText() As Variant) As VbMsgBoxResult

    Dim newMessage As String, newTitle As String
    newMessage = pMessage
    newTitle = pTitle

    'All messages are translatable, but we don't want to translate them if the translation object isn't ready yet
    If (Not (g_Language Is Nothing)) Then
        If g_Language.readyToTranslate Then
            If g_Language.translationActive Then
                newMessage = g_Language.TranslateMessage(pMessage)
                newTitle = g_Language.TranslateMessage(pTitle)
            End If
        End If
    End If
    
    'Once the message is translated, we can add back in any optional parameters
    If Not IsMissing(ExtraText) Then
    
        Dim i As Long
        For i = LBound(ExtraText) To UBound(ExtraText)
            newMessage = Replace$(newMessage, "%" & i + 1, CStr(ExtraText(i)))
        Next i
    
    End If

    pdMsgBox = MsgBox(newMessage, pButtons, newTitle)

End Function

'This popular function is used to display a message in the main form's status bar.
' INPUTS:
' 1) the message to be displayed (mString), with any run-time dependent values
Public Sub Message(ByVal mString As String, ParamArray ExtraText() As Variant)

    Dim newString As String
    newString = mString

    'All messages are translatable, but we don't want to translate them if the translation object isn't ready yet
    If (Not (g_Language Is Nothing)) Then
        If g_Language.readyToTranslate Then
            If g_Language.translationActive Then newString = g_Language.TranslateMessage(mString)
        End If
    End If
    
    'Once the message is translated, we can add back in any optional parameters
    If Not IsMissing(ExtraText) Then
    
        Dim i As Long
        For i = LBound(ExtraText) To UBound(ExtraText)
            newString = Replace$(newString, "%" & i + 1, CStr(ExtraText(i)))
        Next i
    
    End If

    If MacroStatus = MacroSTART Then newString = newString & " {-" & g_Language.TranslateMessage("Recording") & "-}"
    
    If MacroStatus <> MacroBATCH Then
        If FormMain.Visible Then
            g_ProgBar.Text = newString
            g_ProgBar.Draw
        End If
    End If
    
    If Not g_IsProgramCompiled Then Debug.Print newString
    
    'If we're logging program messages, open up a log file and dump the message there
    If g_LogProgramMessages = True Then
        Dim fileNum As Integer
        fileNum = FreeFile
        Open g_UserPreferences.getDataPath & PROGRAMNAME & "_DebugMessages.log" For Append As #fileNum
            Print #fileNum, mString
            If mString = "Finished." Then Print #fileNum, vbCrLf
        Close #fileNum
    End If
    
End Sub

'Pass AutoSelectText a text box and it will select all text currently in the text box
Public Function AutoSelectText(ByRef tBox As TextBox)
    If Not tBox.Visible Then Exit Function
    If Not tBox.Enabled Then Exit Function
    tBox.SetFocus
    tBox.SelStart = 0
    tBox.SelLength = Len(tBox.Text)
End Function

'When the mouse is moved outside the primary image, clear the image coordinates display
Public Sub ClearImageCoordinatesDisplay()
    FormMain.lblCoordinates.Caption = ""
    FormMain.lblCoordinates.Refresh
End Sub

'Populate the passed combo box with options related to distort filter edge-handle options.  Also, select the specified method by default.
Public Sub popDistortEdgeBox(ByRef cmbEdges As ComboBox, Optional ByVal defaultEdgeMethod As EDGE_OPERATOR)

    cmbEdges.Clear
    cmbEdges.AddItem " clamp them to the nearest available pixel"
    cmbEdges.AddItem " reflect them across the nearest edge"
    cmbEdges.AddItem " wrap them around the image"
    cmbEdges.AddItem " erase them"
    cmbEdges.ListIndex = defaultEdgeMethod
    
End Sub

'The labels on the main form's left-hand toolbox may extend past the edge of the toolbar when using a language other than English.
' This function can be used to iterate through the labels, and set them to their largest possible size without extending past
' the edge of the toolbar.
Public Sub FixToolboxText()

    Dim i As Long
    
    'Size 12 font is ideal, if possible
    Dim currentFontSize As Single
    currentFontSize = 12
    
    'We need an object with a DC to use for API calls related to font measuring.  The left-hand toolbox works fine for this.
    FormMain.picLeftPane.Font = FormMain.lblLeftToolBox(0).Font
    FormMain.picLeftPane.FontSize = currentFontSize
        
    'Next, iterate through the controls to find the largest (widest) one
    Dim maxWidth As Long, maxIndex As Long
    maxWidth = 0

    Dim txtWidth As Long
    
    For i = 0 To FormMain.lblLeftToolBox.Count - 1
        txtWidth = getPixelWidthOfString(FormMain.lblLeftToolBox(i).Caption, FormMain.picLeftPane.hDC)
        If txtWidth > maxWidth Then
            maxWidth = txtWidth
            maxIndex = i
        End If
    Next i
    
    'See if the largest (widest) caption extends past the edge of the toolbox
    Dim leftPadding As Long
    leftPadding = FormMain.lblLeftToolBox(0).Left
    
    If maxWidth + leftPadding > FormMain.picLeftPane.ScaleWidth Then
                
        'Iterate through increasingly smaller font sizes, until one is reached that fits inside the toolbox
        Do
        
            currentFontSize = currentFontSize - 1
            FormMain.picLeftPane.FontSize = currentFontSize
            txtWidth = getPixelWidthOfString(FormMain.lblLeftToolBox(maxIndex).Caption, FormMain.picLeftPane.hDC)
        
        Loop While txtWidth + leftPadding > FormMain.picLeftPane.ScaleWidth
        
    End If
    
    'currentFontSize now contains the most appropriate font size for the current language.  Apply it to all toolbox labels.
    For i = 0 To FormMain.lblLeftToolBox.Count - 1
        FormMain.lblLeftToolBox(i).FontSize = currentFontSize
        FormMain.lblLeftToolBox(i).Refresh
    Next i
    
    'Finally, we need to bottom-align all labels manually.
    ' (NOTE: this is an inelegant and ugly way to do this, but I've yet to devise a better solution for bottom-alignment using VB labels)
    Dim newHeight As Long
    newHeight = getPixelHeightOfString(FormMain.lblLeftToolBox(0).Caption, FormMain.picLeftPane.hDC)
    
    FormMain.lblLeftToolBox(0).Top = 25 - newHeight
    FormMain.lblLeftToolBox(1).Top = 98 - newHeight
    FormMain.lblLeftToolBox(2).Top = 186 - newHeight
    FormMain.lblLeftToolBox(3).Top = 282 - newHeight
    
End Sub

'Return the width (and below, height) of a string, in pixels, according to the font assigned to fontContainerDC
Public Function getPixelWidthOfString(ByVal srcString As String, ByVal fontContainerDC As Long) As Long
    Dim txtSize As POINTAPI
    GetTextExtentPoint32 fontContainerDC, srcString, Len(srcString), txtSize
    getPixelWidthOfString = txtSize.X
End Function

Public Function getPixelHeightOfString(ByVal srcString As String, ByVal fontContainerDC As Long) As Long
    Dim txtSize As POINTAPI
    GetTextExtentPoint32 fontContainerDC, srcString, Len(srcString), txtSize
    getPixelHeightOfString = txtSize.Y
End Function
