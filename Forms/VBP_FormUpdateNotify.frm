VERSION 5.00
Begin VB.Form FormUpdateNotify 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Update ready"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9195
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
   ScaleHeight     =   186
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   613
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin PhotoDemon.smartCheckBox chkNotify 
      Height          =   330
      Left            =   120
      TabIndex        =   2
      Top             =   2370
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   159
      Caption         =   "in the future, do not notify me of updates"
      Value           =   0
   End
   Begin PhotoDemon.pdHyperlink lblReleaseAnnouncement 
      Height          =   270
      Left            =   840
      Top             =   930
      Width           =   8130
      _ExtentX        =   14340
      _ExtentY        =   503
      Alignment       =   2
      Caption         =   "(text populated at run-time)"
      FontSize        =   11
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Keep working"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   4680
      TabIndex        =   1
      Top             =   1500
      Width           =   4455
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Restart PhotoDemon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1500
      Width           =   4455
   End
   Begin PhotoDemon.pdLabel lblUpdate 
      Height          =   735
      Left            =   960
      Top             =   120
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   1296
      FontSize        =   11
      Layout          =   1
   End
End
Attribute VB_Name = "FormUpdateNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Tooltip As pdToolTip

Private Sub cmdUpdate_Click(Index As Integer)
    
    'Regardless of the user's choice, we always update their notification preference
    g_UserPreferences.SetPref_Boolean "Updates", "Update Notifications", Not CBool(chkNotify.Value)
    
    Select Case Index
    
        'Restart now
        Case 0
        
            'Set a program-wide restart flag, which PD will use post-patch to initiate a restart.
            g_UserWantsRestart = True
            
            'Hide this dialog
            Me.Visible = False
            
            'Initiate shutdown
            Process "Exit program", True
            
        'Restart later
        Case 1
        
            'If the user wants to keep working, we don't have to do anything special.
            ' (PhotoDemon will automatically apply the remaining patches at shut-down time.)
            Unload Me
    
    End Select
    
End Sub

Private Sub Form_Load()
    
    'Load the "notify of updates" preference
    If g_UserPreferences.GetPref_Boolean("Updates", "Update Notifications", True) Then
        chkNotify.Value = vbUnchecked
    Else
        chkNotify.Value = vbChecked
    End If
    
    'Theme the dialog
    makeFormPretty Me
    
    'Set the release announcement URL
    Dim raURL As String
    raURL = Software_Updater.getReleaseAnnouncementURL
    If Len(raURL) <> 0 Then
        lblReleaseAnnouncement.Caption = g_Language.TranslateMessage("Learn more about the new features in version %1", Software_Updater.getUpdateVersion)
        lblReleaseAnnouncement.Visible = True
        lblReleaseAnnouncement.URL = raURL
    Else
        lblReleaseAnnouncement.Caption = ""
        lblReleaseAnnouncement.Visible = False
    End If
    
    'Disable the restart option inside the IDE
    If Not g_IsProgramCompiled Then
        cmdUpdate(0).Caption = g_Language.TranslateMessage("(Sorry, but automatic restarts don't work inside the IDE.)")
        cmdUpdate(0).Enabled = False
    End If
    
    'Automatically draw a relevant icon using the system icon set
    DrawSystemIcon IDI_ASTERISK, Me.hDC, fixDPI(16), fixDPI(12)
    
    'Display the update message.  (pdLabel automatically handles translations, as necessary.)
    lblUpdate.Caption = "A new version of PhotoDemon is available.  Restart the program to complete the update process."
    
    'Add a few tooltips
    Set m_Tooltip = New pdToolTip
    
    m_Tooltip.setTooltip cmdUpdate(0).hWnd, Me.hWnd, "Restart now to access to the latest version of the program.", "Apply update now"
    m_Tooltip.setTooltip cmdUpdate(1).hWnd, Me.hWnd, "If you're in the middle of something, feel free to keep working.  The update process will automatically complete whenever you next use the program.", "Apply update later"
    
    'Position the form at the bottom-right corner of the main program window.
    Me.Move (FormMain.Left + FormMain.Width) - (Me.Width + 90), (FormMain.Top + FormMain.Height) - (Me.Height + 90)
    
End Sub
