VERSION 5.00
Begin VB.Form FormWait 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Please wait a moment..."
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9015
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
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
   ScaleHeight     =   169
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   601
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin PhotoDemon.pdProgressBar pbMarquee 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   8775
      _ExtentX        =   0
      _ExtentY        =   0
   End
   Begin PhotoDemon.pdLabel lblWaitTitle 
      Height          =   405
      Left            =   240
      Top             =   240
      Width           =   8490
      _ExtentX        =   0
      _ExtentY        =   0
      Alignment       =   2
      Caption         =   "Please wait"
      FontBold        =   -1  'True
      FontSize        =   12
      ForeColor       =   9437184
   End
   Begin PhotoDemon.pdLabel lblWaitDescription 
      Height          =   960
      Left            =   240
      Top             =   1560
      Visible         =   0   'False
      Width           =   8490
      _ExtentX        =   14975
      _ExtentY        =   1905
      Alignment       =   2
      Caption         =   ""
      ForeColor       =   9437184
      Layout          =   1
   End
End
Attribute VB_Name = "FormWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_ModalUnloadCheck As pdTimer
Attribute m_ModalUnloadCheck.VB_VarHelpID = -1

Private Sub Form_Load()

    Interface.ApplyThemeAndTranslations Me
    pbMarquee.MarqueeMode = True
    
    Set m_ModalUnloadCheck = New pdTimer
    m_ModalUnloadCheck.Interval = 16
    m_ModalUnloadCheck.StartTimer
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If (Not m_ModalUnloadCheck Is Nothing) Then
        m_ModalUnloadCheck.StopTimer
        Set m_ModalUnloadCheck = Nothing
    End If
    
    Interface.ReleaseFormTheming Me
    
End Sub

Private Sub m_ModalUnloadCheck_Timer()
    
    DoEvents
    
    'If the dialog is raised modally, an asynchronous method must be used to unload the window.  Set this global flag
    ' to unload the window asynchronously.
    If g_UnloadWaitWindow Then
        g_UnloadWaitWindow = False
        Unload Me
    End If
    
End Sub
