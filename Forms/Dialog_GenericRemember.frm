VERSION 5.00
Begin VB.Form dialog_GenericMemory 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " "
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7185
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   311
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   479
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdPictureBox picIcon 
      Height          =   735
      Left            =   240
      Top             =   360
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1296
   End
   Begin PhotoDemon.pdButton cmdAnswer 
      Height          =   735
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   6780
      _ExtentX        =   11959
      _ExtentY        =   1085
      Caption         =   "Yes"
   End
   Begin PhotoDemon.pdCheckBox chkRemember 
      Height          =   330
      Left            =   240
      TabIndex        =   0
      Top             =   4200
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   582
      Caption         =   " "
   End
   Begin PhotoDemon.pdLabel lblExplanation 
      Height          =   1290
      Left            =   1200
      Top             =   150
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   2275
      Caption         =   ""
      ForeColor       =   2105376
      Layout          =   1
   End
   Begin PhotoDemon.pdButton cmdAnswer 
      Height          =   735
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   6780
      _ExtentX        =   11959
      _ExtentY        =   1296
      Caption         =   "No"
   End
   Begin PhotoDemon.pdButton cmdAnswer 
      Height          =   735
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   3240
      Width           =   6780
      _ExtentX        =   11959
      _ExtentY        =   1296
      Caption         =   "Cancel"
   End
End
Attribute VB_Name = "dialog_GenericMemory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Generic Yes/No/Cancel Dialog with automatic "Remember My Choice" handling
'Copyright 2012-2026 by Tanner Helland
'Created: 01/December/12
'Last updated: 04/May/15
'Last update: merge code from other dialogs into a single, universal version
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'The user's selected option (Yes, No, Cancel)
Private userAnswer As VbMsgBoxResult

'Whether the user wants us to default to this action in the future (True, False)
Private rememberMyChoice As Boolean

'We want to temporarily suspend an hourglass cursor if necessary
Private restoreCursor As Boolean

'Icon used in the dialog, if any
Private m_iconDIB As pdDIB

'Returns yes/no/cancel
Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = userAnswer
End Property

'Returns TRUE if the user checked "remember my answer"
Public Property Get GetRememberAnswerState() As Boolean
    GetRememberAnswerState = rememberMyChoice
End Property

'The ShowDialog routine presents the user with the form.
Public Sub ShowDialog(ByVal questionText As String, ByVal yesButtonText As String, ByVal noButtonText As String, ByVal cancelButtonText As String, ByVal rememberCheckBoxText As String, ByVal dialogTitleText As String, Optional ByVal sysIcon As SystemIconConstants = 0, Optional ByVal defaultAnswer As VbMsgBoxResult = vbCancel, Optional ByVal defaultRemember As Boolean = False, Optional ByVal resNameYesImg As String = "generic_ok", Optional ByVal resNameNoImg As String = "generic_cancel", Optional ByVal resNameCancelImg As String = vbNullString)

    'On the off chance that this dialog is raised during long-running processing, make a note of the cursor prior to displaying the dialog.
    ' We will restore the hourglass cursor (as necessary) when the dialog exits.
    restoreCursor = (Screen.MousePointer = vbHourglass)
    If restoreCursor Then Screen.MousePointer = vbNormal
    
    'Fit the question text
    Dim tmpFont As pdFont
    Set tmpFont = Fonts.GetMatchingUIFont(lblExplanation.FontSize)
    
    Dim txtWidth As Long, txtHeight As Long
    txtWidth = tmpFont.GetWidthOfString(questionText)
    If (txtWidth > lblExplanation.GetWidth) Then
        txtHeight = tmpFont.GetHeightOfWordwrapString(questionText, lblExplanation.GetWidth)
        lblExplanation.Alignment = vbLeftJustify
    Else
        txtHeight = tmpFont.GetHeightOfString(questionText)
        lblExplanation.Alignment = vbCenter
    End If
    
    'If the text is not long enough to require auto-fitting, reposition the label accordingly
    If (txtHeight < lblExplanation.GetHeight) Then
        lblExplanation.SetPositionAndSize lblExplanation.GetLeft, (cmdAnswer(0).GetTop - (txtHeight + 2)) \ 2, lblExplanation.GetWidth, txtHeight + 2
    End If
    
    'Automatically draw the requested icon using the system icon set
    If (sysIcon <> 0) Then
    
        Dim picIconSize As Long
        picIconSize = Interface.FixDPI(56)
        
        picIcon.SetWidth picIconSize
        picIcon.SetHeight picIconSize
        picIcon.SetTop (cmdAnswer(0).GetTop - picIconSize) \ 2
        
        Dim newLeft As Long
        newLeft = picIcon.GetLeft + picIconSize + Interface.FixDPI(16)
        lblExplanation.SetPositionAndSize newLeft, lblExplanation.GetTop, cmdAnswer(0).GetLeft + cmdAnswer(0).GetWidth - newLeft, lblExplanation.GetHeight
        
        Dim iconSuccess As Boolean
        
        If (sysIcon = IDI_EXCLAMATION) Then
            iconSuccess = IconsAndCursors.LoadResourceToDIB("generic_warning", m_iconDIB, picIconSize, picIconSize, 0)
        ElseIf (sysIcon = IDI_QUESTION) Then
            iconSuccess = IconsAndCursors.LoadResourceToDIB("generic_question", m_iconDIB, picIconSize, picIconSize, 0)
        Else
            iconSuccess = IconsAndCursors.LoadResourceToDIB("generic_info", m_iconDIB, picIconSize, picIconSize, 0)
        End If
        
        If (Not iconSuccess) Then Set m_iconDIB = Nothing
        picIcon.Visible = iconSuccess
        
    Else
        picIcon.Visible = False
        lblExplanation.SetPositionAndSize cmdAnswer(0).Left, lblExplanation.GetTop, cmdAnswer(0).GetWidth, lblExplanation.GetHeight
    End If
    
    'Based on the size of the text explanation, and the presence of an icon, resize the form to match
    If (sysIcon = 0) Then
        
        Dim heightDiff As Long, newTopPos As Long
        newTopPos = Interface.FixDPI(32)
        heightDiff = cmdAnswer(0).GetTop - (lblExplanation.GetHeight + (newTopPos * 2))
        
        lblExplanation.SetTop newTopPos
        cmdAnswer(0).SetTop cmdAnswer(0).GetTop - heightDiff
        cmdAnswer(1).SetTop cmdAnswer(1).GetTop - heightDiff
        cmdAnswer(2).SetTop cmdAnswer(2).GetTop - heightDiff
        chkRemember.SetTop chkRemember.GetTop - heightDiff
        
        If (Not g_WindowManager Is Nothing) Then
            Dim winRect As RectL
            g_WindowManager.GetWindowRect_API_Universal Me.hWnd, VarPtr(winRect)
            g_WindowManager.SetSizeByHWnd Me.hWnd, winRect.Right - winRect.Left, (winRect.Bottom - winRect.Top) - heightDiff, True
        End If
    
    End If
    
    'Set the default answer.  (When the form is displayed, this will be used to assign focus to the corresponding button.)
    userAnswer = defaultAnswer
    
    'Apply captions
    lblExplanation.Caption = questionText
    cmdAnswer(0).Caption = yesButtonText
    cmdAnswer(1).Caption = noButtonText
    cmdAnswer(2).Caption = cancelButtonText
    chkRemember.Caption = rememberCheckBoxText
    If (Not g_WindowManager Is Nothing) Then g_WindowManager.SetWindowCaptionW Me.hWnd, dialogTitleText
    
    'The caller can specify whether "remember my choice" is checked by default
    chkRemember.Value = defaultRemember
    
    'Prep button icons at load-time
    Dim buttonIconSize As Long
    buttonIconSize = Interface.FixDPI(32)
    If (LenB(resNameYesImg) <> 0) Then cmdAnswer(0).AssignImage resNameYesImg, , buttonIconSize, buttonIconSize
    If (LenB(resNameNoImg) <> 0) Then cmdAnswer(1).AssignImage resNameNoImg, , buttonIconSize, buttonIconSize
    If (LenB(resNameCancelImg) <> 0) Then cmdAnswer(2).AssignImage resNameCancelImg, , buttonIconSize, buttonIconSize
    
    'Apply visual themes and translations
    Interface.ApplyThemeAndTranslations Me

    'Display the form
    ShowPDDialog vbModal, Me, True

End Sub

'Update the dialog's return value based on the pressed command button
Private Sub cmdAnswer_Click(Index As Integer)
    
    'Note the user's answer (yes/no/cancel)
    Select Case Index
    
        Case 0
            userAnswer = vbYes
            
        Case 1
            userAnswer = vbNo
            
        Case 2
            userAnswer = vbCancel
            
    End Select
    
    'Note the user's preference for remembering this decision
    rememberMyChoice = chkRemember.Value
    
    'Notify the central interface manager of this result
    Interface.NotifyShowDialogResult userAnswer, True
    
    'If a non-standard cursor was in use prior to displaying the dialog, restore it now
    If restoreCursor Then Screen.MousePointer = vbHourglass
    
    'Hiding the form allows the showPDDialog function to continue
    Me.Hide
    
End Sub

Private Sub Form_Activate()
    
    'Set focus to the default answer specified by the caller
    If (Not g_WindowManager Is Nothing) Then
        Select Case userAnswer
            Case vbYes
                g_WindowManager.SetFocusAPI cmdAnswer(0).hWnd
            Case vbNo
                g_WindowManager.SetFocusAPI cmdAnswer(1).hWnd
            Case vbCancel
                g_WindowManager.SetFocusAPI cmdAnswer(2).hWnd
        End Select
    End If
    
    'With the proper button set, we must reset the "userAnswer" variable to vbCancel, in case the user closes the dialog by
    ' some mechanism other than clicking a button (e.g. the corner x).
    userAnswer = vbCancel

End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub picIcon_DrawMe(ByVal targetDC As Long, ByVal ctlWidth As Long, ByVal ctlHeight As Long)
    GDI.FillRectToDC targetDC, 0, 0, ctlWidth, ctlHeight, g_Themer.GetGenericUIColor(UI_Background)
    If (Not m_iconDIB Is Nothing) Then m_iconDIB.AlphaBlendToDC targetDC
End Sub
