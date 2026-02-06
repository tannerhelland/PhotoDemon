VERSION 5.00
Begin VB.Form dialog_InputBox 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " "
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9525
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
   ScaleHeight     =   128
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin PhotoDemon.pdCommandBarMini cmdBarMini 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   1170
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   1323
      DontAutoUnloadParent=   -1  'True
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   480
      Left            =   120
      Top             =   120
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   847
      Caption         =   ""
      FontSize        =   11
   End
   Begin PhotoDemon.pdTextBox txtInput 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   660
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   661
      FontSize        =   11
   End
End
Attribute VB_Name = "dialog_InputBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Unicode Input Box
'Copyright 2025-2026 by Tanner Helland
'Created: 28/December/25
'Last updated: 06/January/26
'Last update: wrap up initial build
'
'I avoided adding an input box to the project for many years, but in 2025 a user requested support for
' importing password-protected PDF documents.  This dialog was the result.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'The user button click from the dialog (OK/Cancel)
Private m_userAnswer As VbMsgBoxResult

'The user-entered text into the primary edit box
Private m_userText As String

Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = m_userAnswer
End Property

'User text is only relevant when DialogResult is OK; ignore otherwise
Public Property Get UserEnteredText() As String
    UserEnteredText = m_userText
End Property

'The ShowDialog routine presents the user with this form.
Public Sub ShowDialog(ByRef dialogTitle As String, ByRef dialogText As String, Optional ByRef defaultUserText As String = vbNullString)

    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    m_userAnswer = vbCancel
    
    'Make sure that a proper cursor is set
    Screen.MousePointer = 0
    
    'Apply the user's text (must already be localized!)
    If (Not g_WindowManager Is Nothing) Then g_WindowManager.SetWindowCaptionW Me.hWnd, dialogTitle
    Me.lblTitle.Caption = dialogText
    Me.txtInput.Text = defaultUserText
    If (LenB(defaultUserText) > 0) Then Me.txtInput.SelectAll
    
    'Theme the dialog
    Interface.ApplyThemeAndTranslations Me, True, False
    
    'Display the dialog
    Me.Show vbModal, FormMain
    
End Sub

Private Sub cmdBarMini_CancelClick()
    m_userAnswer = vbCancel
    m_userText = vbNullString
End Sub

Private Sub cmdBarMini_OKClick()
    m_userAnswer = vbOK
    m_userText = Me.txtInput.Text
End Sub

Private Sub Form_Activate()
    
    'Set focus to the text entry box
    Me.txtInput.SetFocusToEditBox
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Interface.ReleaseFormTheming Me
End Sub

'When the user presses "Enter" in the text box, capture and use to trigger the OK button
Private Sub txtInput_KeyPress(ByVal Shift As ShiftConstants, ByVal vKey As Long, preventFurtherHandling As Boolean)
    If (vKey = VK_RETURN) Then
        preventFurtherHandling = False
        cmdBarMini.ClickOKForMe
    ElseIf (vKey = VK_ESCAPE) Then
        preventFurtherHandling = False
        cmdBarMini.ClickCancelForMe
    End If
End Sub
