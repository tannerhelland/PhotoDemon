VERSION 5.00
Begin VB.Form dialog_IDEWarning 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9045
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
   ScaleHeight     =   410
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   603
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdPictureBox picWarning 
      Height          =   615
      Left            =   330
      Top             =   240
      Width           =   615
      _ExtentX        =   873
      _ExtentY        =   1085
   End
   Begin PhotoDemon.pdButton cmdOK 
      Height          =   750
      Left            =   1125
      TabIndex        =   0
      Top             =   4560
      Width           =   6690
      _ExtentX        =   11800
      _ExtentY        =   1323
      Caption         =   "I understand the risks of running PhotoDemon in the IDE"
   End
   Begin PhotoDemon.pdCheckBox chkRepeat 
      Height          =   300
      Left            =   1140
      TabIndex        =   1
      Top             =   5520
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   582
      Caption         =   "Do not display this warning again"
   End
   Begin PhotoDemon.pdLabel lblWarning 
      Height          =   885
      Index           =   3
      Left            =   360
      Top             =   3210
      Width           =   8175
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   ""
      ForeColor       =   4210752
      Layout          =   1
   End
   Begin PhotoDemon.pdLabel lblWarning 
      Height          =   765
      Index           =   2
      Left            =   360
      Top             =   2400
      Width           =   8385
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   ""
      ForeColor       =   4210752
      Layout          =   1
   End
   Begin PhotoDemon.pdLabel lblWarning 
      Height          =   1245
      Index           =   1
      Left            =   360
      Top             =   1080
      Width           =   8385
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   ""
      ForeColor       =   4210752
      Layout          =   1
   End
   Begin PhotoDemon.pdLabel lblWarning 
      Height          =   525
      Index           =   0
      Left            =   1005
      Top             =   390
      Width           =   7695
      _ExtentX        =   0
      _ExtentY        =   0
      Alignment       =   2
      Caption         =   ""
      ForeColor       =   2105376
      Layout          =   1
   End
End
Attribute VB_Name = "dialog_IDEWarning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'IDE Warning Dialog
'Copyright 2011-2026 by Tanner Helland
'Created: 15/December/12
'Last updated: 07/March/20
'Last update: remove translation wrappers for text; it's fine for these warnings to appear
'             in English, only - and it reduces burdens for volunteer translators!
'
'Generally speaking, it's not a great idea to run PhotoDemon in the IDE.  This dialog
' is used to warn the user of the associated risks with doing so.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Theme-specific icons are fully supported
Private m_warningDIB As pdDIB

'The user input from the dialog
Private userAnswer As VbMsgBoxResult

Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = userAnswer
End Property

'The ShowDialog routine presents the user with the form.
Public Sub ShowDialog()
    
    'Prep a warning icon
    Dim warningIconSize As Long
    warningIconSize = Interface.FixDPI(32)
    
    If Not IconsAndCursors.LoadResourceToDIB("generic_warning", m_warningDIB, warningIconSize, warningIconSize, 0) Then
        Set m_warningDIB = Nothing
        picWarning.Visible = False
    End If
    
    'This hack prevents these strings from being picked up by PD's translation generator.
    ' It does not inspect individual strings for translatable text.
    Dim strDialog(0 To 4) As String
    strDialog(0) = "You are running PhotoDemon inside the Visual Basic IDE.  This is not recommended."
    strDialog(1) = "Please compile PhotoDemon before using it.  Many features that rely on subclassing are disabled in the IDE, but some - such as custom command buttons - cannot be disabled without severely impacting the program's functionality.  As such, you may experience IDE instability and crashes, especially if you close the program using the IDE's Stop button."
    strDialog(2) = "Additionally, like all other photo editors, PhotoDemon relies heavily on multidimensional arrays. Array performance is severely degraded in the IDE, so some functions may perform very slowly."
    strDialog(3) = "If you insist on running PhotoDemon in the IDE, please do not submit bugs regarding IDE crashes or freezes.  PhotoDemon's developers can only address issues and bugs that affect the compiled .exe."
    strDialog(4) = "Visual Basic IDE Detected"
    lblWarning(0).Caption = strDialog(0)
    lblWarning(1).Caption = strDialog(1)
    lblWarning(2).Caption = strDialog(2)
    lblWarning(3).Caption = strDialog(3)
    
    'It's okay to use a non-Unicode-safe assignment here, because this warning is intentionally non-translated
    Me.Caption = strDialog(4)
    
    Dim buttonIconSize As Long
    buttonIconSize = Interface.FixDPI(32)
    cmdOK.AssignImage "generic_ok", , buttonIconSize, buttonIconSize
    
    'Provide a default answer of "first image only" (in the event that the user clicks the "x" button in the top-right)
    userAnswer = vbOK

    'Apply any custom styles to the form
    ApplyThemeAndTranslations Me

    'Display the form
    ShowPDDialog vbModal, Me, True

End Sub

Private Sub CmdOK_Click()
    If chkRepeat.Value Then UserPrefs.SetPref_Boolean "Core", "Display IDE Warning", False
    Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub picWarning_DrawMe(ByVal targetDC As Long, ByVal ctlWidth As Long, ByVal ctlHeight As Long)
    GDI.FillRectToDC targetDC, 0, 0, ctlWidth, ctlHeight, g_Themer.GetGenericUIColor(UI_Background)
    If (Not m_warningDIB Is Nothing) Then m_warningDIB.AlphaBlendToDC targetDC, , (ctlWidth - m_warningDIB.GetDIBWidth) \ 2, (ctlHeight - m_warningDIB.GetDIBHeight) \ 2
End Sub
