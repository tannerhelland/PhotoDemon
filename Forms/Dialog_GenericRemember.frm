VERSION 5.00
Begin VB.Form dialog_GenericMemory 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " "
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7185
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   311
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   479
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdButton cmdAnswer 
      Height          =   735
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   6780
      _extentx        =   11959
      _extenty        =   1085
      caption         =   "&Yes"
   End
   Begin PhotoDemon.smartCheckBox chkRemember 
      Height          =   330
      Left            =   240
      TabIndex        =   0
      Top             =   4200
      Width           =   6735
      _extentx        =   11880
      _extenty        =   582
      caption         =   " "
   End
   Begin PhotoDemon.pdLabel lblExplanation 
      Height          =   1290
      Left            =   960
      Top             =   150
      Width           =   6015
      _extentx        =   10610
      _extenty        =   2275
      caption         =   ""
      forecolor       =   2105376
      layout          =   1
   End
   Begin PhotoDemon.pdButton cmdAnswer 
      Height          =   735
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   6780
      _extentx        =   11959
      _extenty        =   1296
      caption         =   "&No"
   End
   Begin PhotoDemon.pdButton cmdAnswer 
      Height          =   735
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   3240
      Width           =   6780
      _extentx        =   11959
      _extenty        =   1296
      caption         =   "&Cancel"
   End
End
Attribute VB_Name = "dialog_GenericMemory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Generic Yes/No/Cancel Dialog with automatic "Remember My Choice" handling
'Copyright 2012-2015 by Tanner Helland
'Created: 01/December/12
'Last updated: 04/May/15
'Last update: merge code from other dialogs into a single, universal version
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'The user's selected option (Yes, No, Cancel)
Private userAnswer As VbMsgBoxResult

'Whether the user wants us to default to this action in the future (True, False)
Private rememberMyChoice As Boolean

'We want to temporarily suspend an hourglass cursor if necessary
Private restoreCursor As Boolean

'Returns yes/no/cancel
Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = userAnswer
End Property

'Returns TRUE if the user checked "remember my answer"
Public Property Get getRememberAnswerState() As Boolean
    getRememberAnswerState = rememberMyChoice
End Property


'The ShowDialog routine presents the user with the form.  FormID MUST BE SET in advance of calling this.
Public Sub showDialog(ByVal questionText As String, ByVal yesButtonText As String, ByVal noButtonText As String, ByVal cancelButtonText As String, ByVal rememberCheckBoxText As String, ByVal dialogTitleText As String, Optional ByVal icon As SystemIconConstants = IDI_QUESTION, Optional ByVal defaultAnswer As VbMsgBoxResult = vbCancel, Optional ByVal defaultRemember As Boolean = False)

    'On the off chance that this dialog is raised during long-running processing, make a note of the cursor prior to displaying the dialog.
    ' We will restore the hourglass cursor (as necessary) when the dialog exits.
    If Screen.MousePointer = vbHourglass Then
        restoreCursor = True
        Screen.MousePointer = vbNormal
    Else
        restoreCursor = False
    End If

    'Automatically draw the requested icon using the system icon set
    Dim iconY As Long
    iconY = fixDPI(18)
    If g_UseFancyFonts Then iconY = iconY + fixDPI(2)
    DrawSystemIcon icon, Me.hDC, fixDPI(22), iconY
    
    'Set the default answer.  (When the form is displayed, this will be used to assign focus to the corresponding button.)
    userAnswer = defaultAnswer
    
    'Apply captions
    lblExplanation.Caption = questionText
    cmdAnswer(0).Caption = yesButtonText
    cmdAnswer(1).Caption = noButtonText
    cmdAnswer(2).Caption = cancelButtonText
    chkRemember.Caption = rememberCheckBoxText
    Me.Caption = dialogTitleText
    
    'The caller can specify whether "remember my choice" is checked by default
    If defaultRemember Then
        chkRemember.Value = vbChecked
    Else
        chkRemember.Value = vbUnchecked
    End If

    'Apply visual themes and translations
    makeFormPretty Me

    'Display the form
    showPDDialog vbModal, Me, True

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
    rememberMyChoice = CBool(chkRemember.Value)
    
    'If a non-standard cursor was in use prior to displaying the dialog, restore it now
    If restoreCursor Then Screen.MousePointer = vbHourglass
    
    'Hiding the form allows the showPDDialog function to continue
    Me.Hide
    
End Sub

Private Sub Form_Activate()
    
    'Set focus to the default answer specified by the caller
    Select Case userAnswer
    
        Case vbYes
            cmdAnswer(0).SetFocus
        
        Case vbNo
            cmdAnswer(1).SetFocus
        
        Case vbCancel
            cmdAnswer(2).SetFocus
    
    End Select
    
    'With the proper button set, we must reset the "userAnswer" variable to vbCancel, in case the user closes the dialog by
    ' some mechanism other than clicking a button (e.g. the corner x).
    userAnswer = vbCancel

End Sub

Private Sub Form_Load()
    
    'Prep button icons at load-time
    cmdAnswer(0).AssignImage "LRGACCEPT"
    cmdAnswer(1).AssignImage "LRGCANCEL"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

