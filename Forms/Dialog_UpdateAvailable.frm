VERSION 5.00
Begin VB.Form FormUpdateNotify 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Update ready"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9195
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
   ScaleHeight     =   186
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   613
   ShowInTaskbar   =   0   'False
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
   Begin PhotoDemon.pdCheckBox chkNotify 
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   2370
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   582
      Caption         =   "in the future, do not notify me of updates"
      Value           =   0   'False
   End
   Begin PhotoDemon.pdHyperlink lblReleaseAnnouncement 
      Height          =   270
      Left            =   840
      TabIndex        =   3
      Top             =   930
      Width           =   8130
      _ExtentX        =   14340
      _ExtentY        =   503
      Alignment       =   2
      Caption         =   ""
      FontSize        =   11
   End
   Begin PhotoDemon.pdLabel lblUpdate 
      Height          =   735
      Left            =   960
      Top             =   120
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   1296
      Caption         =   ""
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
   Begin PhotoDemon.pdPictureBox picWarning 
      Height          =   615
      Left            =   60
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1085
   End
End
Attribute VB_Name = "FormUpdateNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Update Notification form
'Copyright 2014-2026 by Tanner Helland
'Created: 03/March/14
'Last updated: 06/September/15
'Last update: convert buttons to pdButton
'
'This dialog's a simple one: when an update is available, it will notify the user and give them the choice to
' immediately restart+apply, or continue working.  Not much to it!
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Theme-specific icons are fully supported
Private m_icoDIB As pdDIB

Private Sub cmdUpdate_Click(Index As Integer)
    
    'Regardless of the user's choice, we always update their notification preference
    UserPrefs.SetPref_Boolean "Updates", "Update Notifications", Not chkNotify.Value
    
    Select Case Index
    
        'Restart now
        Case 0
        
            'Set a program-wide restart flag, which PD will use post-patch to initiate a restart.
            Updates.SetRestartAfterUpdate True
            
            'Hide this dialog
            Me.Visible = False
            
            'Initiate shutdown
            Process "Exit program", True
            
        'Restart later
        Case 1
            
            'The update will apply at shutdown time, and the user does *not* want us to restart
            Updates.SetRestartAfterUpdate False
            
            'If the user wants to keep working, we don't have to do anything special.
            ' (PhotoDemon will automatically apply the remaining patches at shut-down time.)
            Unload Me
    
    End Select
    
End Sub

Private Sub Form_Load()
    
    'Load the "notify of updates" preference
    chkNotify.Value = Not UserPrefs.GetPref_Boolean("Updates", "Update Notifications", True)
    
    'Set the release announcement URL
    Dim raURL As String
    raURL = Updates.GetReleaseAnnouncementURL
    If (LenB(raURL) <> 0) Then
        lblReleaseAnnouncement.Caption = g_Language.TranslateMessage("Learn more about the new features in %1", Updates.GetUpdateVersion_Friendly)
        lblReleaseAnnouncement.Visible = True
        lblReleaseAnnouncement.URL = raURL
    Else
        lblReleaseAnnouncement.Caption = vbNullString
        lblReleaseAnnouncement.Visible = False
    End If
    
    'Disable the restart option inside the IDE
    If (Not OS.IsProgramCompiled) Then
        cmdUpdate(0).Caption = g_Language.TranslateMessage("(Sorry, but automatic restarts don't work inside the IDE.)")
        cmdUpdate(0).Enabled = False
    End If
    
    'Prep an information icon
    Dim icoSize As Long
    icoSize = Interface.FixDPI(32)
    
    If Not IconsAndCursors.LoadResourceToDIB("generic_info", m_icoDIB, icoSize, icoSize, 0) Then
        Set m_icoDIB = Nothing
        picWarning.Visible = False
    End If
    
    'Display the update message.  (pdLabel automatically handles translations, as necessary.)
    lblUpdate.Caption = "A new version of PhotoDemon is available.  Restart the program to complete the update process."
    
    'Add a few tooltips
    cmdUpdate(0).AssignTooltip "Restart now to access to the latest version of the program.", "Apply update now"
    cmdUpdate(1).AssignTooltip "If you're in the middle of something, feel free to keep working.  The update process will automatically complete whenever you next use the program.", "Apply update later"
    
    'Theme the dialog
    ApplyThemeAndTranslations Me
    
    'Position the form at the bottom-right corner of the main program window.
    Me.Move (FormMain.Left + FormMain.Width) - (Me.Width + 90), (FormMain.Top + FormMain.Height) - (Me.Height + 90)
    
End Sub

Private Sub picWarning_DrawMe(ByVal targetDC As Long, ByVal ctlWidth As Long, ByVal ctlHeight As Long)
    GDI.FillRectToDC targetDC, 0, 0, ctlWidth, ctlHeight, g_Themer.GetGenericUIColor(UI_Background)
    If (Not m_icoDIB Is Nothing) Then m_icoDIB.AlphaBlendToDC targetDC, , (ctlWidth - m_icoDIB.GetDIBWidth) \ 2, (ctlHeight - m_icoDIB.GetDIBHeight) \ 2
End Sub
