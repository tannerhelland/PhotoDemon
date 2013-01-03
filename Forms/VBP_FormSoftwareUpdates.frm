VERSION 5.00
Begin VB.Form FormSoftwareUpdate 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Update Notifier"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5535
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
   ScaleHeight     =   401
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   369
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOKNo 
      Caption         =   "&OK"
      Height          =   495
      Left            =   3960
      TabIndex        =   15
      Top             =   5280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox picNo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   0
      MousePointer    =   1  'Arrow
      ScaleHeight     =   345
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   369
      TabIndex        =   11
      Top             =   6240
      Visible         =   0   'False
      Width           =   5535
      Begin VB.TextBox txtNoExplanation 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00400000&
         Height          =   1335
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   1440
         Width           =   5055
      End
      Begin VB.Label lblDirectPDDownload 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.tannerhelland.com/photodemon"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   120
         MouseIcon       =   "VBP_FormSoftwareUpdates.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   3360
         Width           =   5355
      End
      Begin VB.Label lblDownloadTitle 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PhotoDemon website:"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1995
         TabIndex        =   13
         Top             =   3000
         Width           =   1605
      End
   End
   Begin VB.CommandButton cmdNoDownloadNoReminder 
      Caption         =   "Not now, not ever.  Do not prompt me about future updates."
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   5280
      Width           =   5055
   End
   Begin VB.CommandButton cmdNoDownload 
      Caption         =   "Not right now, but please remind me again in the future."
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   4680
      Width           =   5055
   End
   Begin VB.CommandButton cmdYesDownload 
      Caption         =   "&Yes!  Take me to the download page."
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   3960
      Width           =   5055
   End
   Begin VB.Label lblQuestion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Do you want to download the new version?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   555
      Left            =   120
      TabIndex        =   10
      Top             =   3405
      Width           =   5340
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000002&
      X1              =   360
      X2              =   8
      Y1              =   208
      Y2              =   208
   End
   Begin VB.Label lblUpdateAnnouncement 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "http://update-announcement-goes-here"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      MouseIcon       =   "VBP_FormSoftwareUpdates.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   2520
      Width           =   5295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000002&
      X1              =   360
      X2              =   8
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line0 
      BorderColor     =   &H80000002&
      X1              =   360
      X2              =   8
      Y1              =   40
      Y2              =   40
   End
   Begin VB.Label lblAnnouncementExplanation 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "To learn about the latest version before downloading it, please visit:"
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   5220
   End
   Begin VB.Label lblNewestAppVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "X.X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   3120
      TabIndex        =   4
      Top             =   1320
      Width           =   2205
   End
   Begin VB.Label lblCurrentAppVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "X.X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   3120
      TabIndex        =   3
      Top             =   960
      Width           =   1950
   End
   Begin VB.Label lblNewVersion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Newest version:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   -480
      TabIndex        =   2
      Top             =   1320
      Width           =   3360
   End
   Begin VB.Label lblCurVersion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Your version:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   -480
      TabIndex        =   1
      Top             =   960
      Width           =   3360
   End
   Begin VB.Label lblUpdateIntro 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A new version of PhotoDemon is available!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   165
      Width           =   5220
   End
End
Attribute VB_Name = "FormSoftwareUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Automatic Software Updater (note: it doesn't do the actual updating, it just CHECKS for updates!)
'Copyright ©2000-2013 by Tanner Helland
'Created: 19/August/12
'Last updated: 19/August/12
'Last update: initial build
'
'Interface for notifying the user that a new version of PhotoDemon is available for download.  This code is simply the
' notification part; the actual update checking is handled within the SoftwareUpdater module.
'
'Note that this code interfaces with the .INI file so the user can opt to not check for updates and never be
' notified again. (FYI - this option can be enabled/disabled from the 'Edit' -> 'Program Preferences' menu.)
'
'***************************************************************************

Option Explicit

'Do not download the update, but prompt the user again in the future
Private Sub cmdNoDownload_Click()
    
    userPreferences.SetPreference_Boolean "General Preferences", "CheckForUpdates", True
    
    Message "Automatic update canceled."
    
    cmdYesDownload.Visible = False
    cmdNoDownload.Visible = False
    cmdNoDownloadNoReminder.Visible = False
    picNo.Left = 0
    picNo.Top = 0
    DoEvents
    txtNoExplanation.Text = "The next time you launch " & PROGRAMNAME & ", it will remind you about this software update." & vbCrLf & vbCrLf & "Note: you can always manually download the latest version of " & PROGRAMNAME & " by visiting the " & PROGRAMNAME & " website."
    picNo.Visible = True
    cmdOKNo.Visible = True
    cmdOKNo.SetFocus
    
End Sub

'Do not download the update, and do not prompt the user again
Private Sub cmdNoDownloadNoReminder_Click()
    
    userPreferences.SetPreference_Boolean "General Preferences", "CheckForUpdates", False
    
    Message "Automatic update canceled."
    
    cmdYesDownload.Visible = False
    cmdNoDownload.Visible = False
    cmdNoDownloadNoReminder.Visible = False
    picNo.Left = 0
    picNo.Top = 0
    DoEvents
    txtNoExplanation.Text = PROGRAMNAME & " will no longer prompt you about software updates.  (If you change your mind in the future, this setting can be reversed from the 'Edit' -> 'Program Preferences' menu.)" & vbCrLf & vbCrLf & "Note: you can always manually download the latest version of " & PROGRAMNAME & " by visiting the " & PROGRAMNAME & " website."
    picNo.Visible = True
    cmdOKNo.Visible = True
    cmdOKNo.SetFocus
    
End Sub

'OK button on the picture box loaded when the user selects one of the "no, don't download" options
Private Sub cmdOKNo_Click()
    Unload Me
End Sub

'Yes, the user wants us to download the new version.  Launch a link to the update page and close this form.
Private Sub cmdYesDownload_Click()
    OpenURL "http://www.tannerhelland.com/photodemon/#download"
    Message "Thanks for downloading the latest PhotoDemon update.  Hope you enjoy it!"
    Unload Me
End Sub

'LOAD form
Private Sub Form_Load()
    
    'Update the form labels with the version information
    lblCurrentAppVersion = App.Major & "." & App.Minor & "." & App.Revision
    lblNewestAppVersion = updateMajor & "." & updateMinor & "." & updateBuild
    
    'If available, display a link to the update notice of this software version
    If updateAnnouncement <> "" Then
        lblUpdateAnnouncement = "What's new in PhotoDemon " & updateMajor & "." & updateMinor & "?"
    Else
        lblAnnouncementExplanation.Visible = False
        lblUpdateAnnouncement.Visible = False
    End If
    
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
End Sub

'Open a browser window with the PhotoDemon download page
Private Sub lblDirectPDDownload_Click()
    OpenURL "http://www.tannerhelland.com/photodemon/#download"
End Sub

'When the user clicks the "more information" link, open a browser window with the blog article pulled from updates.txt
Private Sub lblUpdateAnnouncement_Click()
    OpenURL updateAnnouncement
End Sub
