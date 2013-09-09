VERSION 5.00
Begin VB.Form FormWait 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Applying changes..."
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9015
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   106
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   601
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.Timer tmrProgBar 
      Interval        =   50
      Left            =   120
      Top             =   1560
   End
   Begin VB.PictureBox picProgBar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   585
      TabIndex        =   1
      Top             =   840
      Width           =   8775
   End
   Begin VB.Label lblWaitTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "please wait"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00900000&
      Height          =   405
      Left            =   240
      TabIndex        =   0
      Top             =   240
      UseMnemonic     =   0   'False
      Width           =   8490
   End
End
Attribute VB_Name = "FormWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'System progress bar control
Private sysProgBar As cProgressBarOfficial

Private Sub Form_Load()

    Set sysProgBar = New cProgressBarOfficial
    sysProgBar.CreateProgressBar picProgBar.hWnd, 0, 0, picProgBar.ScaleWidth, picProgBar.ScaleHeight, True, True, True, True
    sysProgBar.Max = 100
    sysProgBar.Min = 0
    sysProgBar.Value = 0
    sysProgBar.Marquee = True
    sysProgBar.Value = 0
    
    'Turn on the progress bar timer, which is used to move the marquee progress bar
    'tmrProgBar.Enabled = True
    
End Sub

Private Sub tmrProgBar_Timer()

    sysProgBar.Value = sysProgBar.Value + 1
    If sysProgBar.Value = sysProgBar.Max Then sysProgBar.Value = sysProgBar.Min
    
    sysProgBar.Refresh
    
End Sub
