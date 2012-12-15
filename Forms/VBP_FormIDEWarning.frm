VERSION 5.00
Begin VB.Form dialog_IDEWarning 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Visual Basic IDE Detected"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9045
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
   ScaleHeight     =   440
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   603
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkRepeat 
      Appearance      =   0  'Flat
      Caption         =   "  Do not display this warning again"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   390
      Left            =   2790
      TabIndex        =   2
      Top             =   5970
      Value           =   1  'Checked
      Width           =   3855
   End
   Begin PhotoDemon.jcbutton cmdOK 
      Default         =   -1  'True
      Height          =   735
      Left            =   1125
      TabIndex        =   0
      Top             =   4560
      Width           =   6690
      _ExtentX        =   11800
      _ExtentY        =   1296
      ButtonStyle     =   13
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "I understand the risks of running PhotoDemon in the IDE"
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormIDEWarning.frx":0000
      PictureAlign    =   0
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipType     =   1
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   $"VBP_FormIDEWarning.frx":1052
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   885
      Index           =   3
      Left            =   360
      TabIndex        =   5
      Top             =   3240
      Width           =   8175
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   $"VBP_FormIDEWarning.frx":111A
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   765
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Top             =   2400
      Width           =   8385
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   $"VBP_FormIDEWarning.frx":11DD
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1245
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   8385
      WordWrap        =   -1  'True
   End
   Begin VB.Line lineSeparator 
      BorderColor     =   &H8000000D&
      X1              =   8
      X2              =   592
      Y1              =   376
      Y2              =   376
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   "You are running PhotoDemon inside the Visual Basic IDE.  This is not recommended."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00202020&
      Height          =   525
      Index           =   0
      Left            =   1005
      TabIndex        =   1
      Top             =   390
      Width           =   7695
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "dialog_IDEWarning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'IDE Warning Dialog
'Copyright ©2011-2012 by Tanner Helland
'Created: 15/December/12
'Last updated: 15/December/12
'Last update: initial build; this replaces the generic message box previously used
'
'Generally speaking, it's not a great idea to run PhotoDemon in the IDE.  This dialog
' is used to warn the user of the associated risks with doing so.
'
'***************************************************************************

Option Explicit

'The user input from the dialog
Private userAnswer As VbMsgBoxResult

Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = userAnswer
End Property

'The ShowDialog routine presents the user with the form.  FormID MUST BE SET in advance of calling this.
Public Sub ShowDialog()

    'Automatically draw a warning icon using the system icon set
    Dim iconY As Long
    iconY = 18
    If useFancyFonts Then iconY = iconY + 2
    DrawSystemIcon IDI_EXCLAMATION, Me.hDC, 22, iconY
    
    'Provide a default answer of "first image only" (in the event that the user clicks the "x" button in the top-right)
    userAnswer = vbOK

    'Apply any custom styles to the form
    makeFormPretty Me

    'Display the form
    Me.Show vbModal, FormMain

End Sub

'OK button
Private Sub cmdOK_Click()

    If CBool(chkRepeat.Value) Then userPreferences.SetPreference_Boolean "General Preferences", "DisplayIDEWarning", False
    Me.Hide

End Sub

