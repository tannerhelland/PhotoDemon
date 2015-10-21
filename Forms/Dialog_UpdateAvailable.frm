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
   Begin PhotoDemon.pdButton cmdUpdate 
      Height          =   750
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   1323
      Caption         =   "Restart PhotoDemon"
   End
   Begin PhotoDemon.smartCheckBox chkNotify 
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   2370
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   582
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
   Begin PhotoDemon.pdButton cmdUpdate 
      Height          =   750
      Index           =   1
      Left            =   4680
      TabIndex        =   2
      Top             =   1440
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   1323
      Caption         =   "Keep working"
   End
End
Attribute VB_Name = "FormUpdateNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Update Notification form
'Copyright 2014-2015 by Tanner Helland
'Created: 03/March/14
'Last updated: 06/September/15
'Last update: convert buttons to pdButton
'
'This dialog's a simple one: when an update is available, it will notify the user and give them the choice to
' immediately restart+apply, or continue working.  Not much to it!
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************


Option Explicit

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
        
    'Set the release announcement URL
    Dim raURL As String
    raURL = Software_Updater.getReleaseAnnouncementURL
    If Len(raURL) <> 0 Then
        lblReleaseAnnouncement.Caption = g_Language.TranslateMessage("Learn more about the new features in %1", Software_Updater.getUpdateVersion_Friendly)
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
    DrawSystemIcon IDI_ASTERISK, Me.hDC, FixDPI(16), FixDPI(12)
    
    'Display the update message.  (pdLabel automatically handles translations, as necessary.)
    lblUpdate.Caption = "A new version of PhotoDemon is available.  Restart the program to complete the update process."
    
    'Add a few tooltips
    cmdUpdate(0).AssignTooltip "Restart now to access to the latest version of the program.", "Apply update now"
    cmdUpdate(1).AssignTooltip "If you're in the middle of something, feel free to keep working.  The update process will automatically complete whenever you next use the program.", "Apply update later"
    
    'Theme the dialog
    MakeFormPretty Me
    
    'Position the form at the bottom-right corner of the main program window.
    Me.Move (FormMain.Left + FormMain.Width) - (Me.Width + 90), (FormMain.Top + FormMain.Height) - (Me.Height + 90)
    
End Sub
