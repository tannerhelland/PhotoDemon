VERSION 5.00
Begin VB.Form FormLanguageEditor 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Language Editor"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15360
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
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Default         =   -1  'True
      Height          =   615
      Left            =   11880
      TabIndex        =   4
      Top             =   8310
      Width           =   1725
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   615
      Left            =   13860
      TabIndex        =   3
      Top             =   8310
      Width           =   1365
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "&Previous"
      Enabled         =   0   'False
      Height          =   615
      Left            =   10080
      TabIndex        =   2
      Top             =   8310
      Width           =   1725
   End
   Begin VB.Label lblBackground 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   -240
      TabIndex        =   5
      Top             =   8235
      Width           =   17415
   End
   Begin VB.Label lblExplanation 
      BackStyle       =   0  'Transparent
      Caption         =   "(text populated at run-time)"
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
      Height          =   7320
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   780
      Width           =   3135
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      Index           =   0
      X1              =   224
      X2              =   224
      Y1              =   48
      Y2              =   544
   End
   Begin VB.Label lblWizardTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Step 1: select the language file to edit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "FormLanguageEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

Private Sub Form_Load()
    
    MsgBox "This tool is currently under heavy development.  It may not work as expected.", vbInformation + vbOKOnly + vbApplicationModal, "Development warning"
    
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

