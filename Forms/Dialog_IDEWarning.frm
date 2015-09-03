VERSION 5.00
Begin VB.Form dialog_IDEWarning 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Visual Basic IDE Detected"
   ClientHeight    =   6150
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
   ScaleHeight     =   410
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   603
   ShowInTaskbar   =   0   'False
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
   Begin PhotoDemon.smartCheckBox chkRepeat 
      Height          =   300
      Left            =   1140
      TabIndex        =   1
      Top             =   5520
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   582
      Caption         =   "Do not display this warning again"
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   "Warning"
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
      Caption         =   "Warning"
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
      Caption         =   "Warning"
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
      TabIndex        =   2
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
'Copyright 2011-2015 by Tanner Helland
'Created: 15/December/12
'Last updated: 15/December/12
'Last update: initial build; this replaces the generic message box previously used
'
'Generally speaking, it's not a great idea to run PhotoDemon in the IDE.  This dialog
' is used to warn the user of the associated risks with doing so.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'The user input from the dialog
Private userAnswer As VbMsgBoxResult

Public Property Get DialogResult() As VbMsgBoxResult
    DialogResult = userAnswer
End Property

'The ShowDialog routine presents the user with the form.  FormID MUST BE SET in advance of calling this.
Public Sub showDialog()

    'Automatically draw a warning icon using the system icon set
    Dim iconY As Long
    iconY = fixDPI(18)
    If g_UseFancyFonts Then iconY = iconY + fixDPI(2)
    DrawSystemIcon IDI_EXCLAMATION, Me.hDC, fixDPI(22), iconY
    
    lblWarning(1).Caption = g_Language.TranslateMessage("Please compile PhotoDemon before using it.  Many features that rely on subclassing are disabled in the IDE, but some - such as custom command buttons - cannot be disabled without severely impacting the program's functionality.  As such, you may experience IDE instability and crashes, especially if you close the program using the IDE's Stop button.")
    lblWarning(2).Caption = g_Language.TranslateMessage("Additionally, like all other photo editors, PhotoDemon relies heavily on multidimensional arrays. Array performance is severely degraded in the IDE, so some functions may perform very slowly.")
    lblWarning(3).Caption = g_Language.TranslateMessage("If you insist on running PhotoDemon in the IDE, please do not submit bugs regarding IDE crashes or freezes.  PhotoDemon's developers can only address issues and bugs that affect the compiled .exe.")
    
    cmdOK.AssignImage "LRGACCEPT"
    
    'Provide a default answer of "first image only" (in the event that the user clicks the "x" button in the top-right)
    userAnswer = vbOK

    'Apply any custom styles to the form
    makeFormPretty Me

    'Display the form
    showPDDialog vbModal, Me, True

End Sub

Private Sub CmdOK_Click()
    If CBool(chkRepeat.Value) Then g_UserPreferences.SetPref_Boolean "Core", "Display IDE Warning", False
    Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub
