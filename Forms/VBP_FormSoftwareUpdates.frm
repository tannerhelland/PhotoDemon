VERSION 5.00
Begin VB.Form FormSoftwareUpdate 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Update Notifier"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10710
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
   ScaleHeight     =   354
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   714
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOKNo 
      Caption         =   "&OK"
      Height          =   615
      Left            =   9120
      TabIndex        =   15
      Top             =   4560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox picNo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   0
      MousePointer    =   1  'Arrow
      ScaleHeight     =   353
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   721
      TabIndex        =   11
      Top             =   5520
      Visible         =   0   'False
      Width           =   10815
      Begin VB.TextBox txtNoExplanation 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
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
         Height          =   2415
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   360
         Width           =   10335
      End
      Begin VB.Label lblDirectPDDownload 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.photodemon.org"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   600
         MouseIcon       =   "VBP_FormSoftwareUpdates.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   3480
         Width           =   3090
      End
      Begin VB.Label lblDownloadTitle 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PhotoDemon website:"
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
         Height          =   285
         Left            =   240
         TabIndex        =   13
         Top             =   3120
         Width           =   2355
      End
   End
   Begin VB.CommandButton cmdNoDownloadNoReminder 
      Caption         =   "Not now, not ever.  Do not prompt me about future updates."
      Height          =   735
      Left            =   5400
      TabIndex        =   9
      Top             =   4440
      Width           =   5055
   End
   Begin VB.CommandButton cmdNoDownload 
      Caption         =   "Not right now, but please remind me again in the future."
      Height          =   735
      Left            =   5400
      TabIndex        =   8
      Top             =   3600
      Width           =   5055
   End
   Begin VB.CommandButton cmdYesDownload 
      Caption         =   "&Yes!  Take me to the download page."
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      TabIndex        =   7
      Top             =   3600
      Width           =   5055
   End
   Begin VB.Label lblQuestion 
      AutoSize        =   -1  'True
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
      Height          =   285
      Left            =   240
      TabIndex        =   10
      Top             =   3000
      Width           =   4620
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      X1              =   704
      X2              =   8
      Y1              =   184
      Y2              =   184
   End
   Begin VB.Label lblUpdateAnnouncement 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "http://update-announcement-goes-here"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
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
      Top             =   2190
      Width           =   10335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      X1              =   704
      X2              =   8
      Y1              =   56
      Y2              =   56
   End
   Begin VB.Label lblAnnouncementExplanation 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "To learn about the latest version before downloading, please visit:"
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
      Height          =   360
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   10290
   End
   Begin VB.Label lblNewestAppVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X.X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   7680
      TabIndex        =   4
      Top             =   1200
      Width           =   405
   End
   Begin VB.Label lblCurrentAppVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X.X"
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
      Height          =   285
      Left            =   3600
      TabIndex        =   3
      Top             =   1200
      Width           =   345
   End
   Begin VB.Label lblNewVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Newest version:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   5475
      TabIndex        =   2
      Top             =   1200
      Width           =   1965
   End
   Begin VB.Label lblCurVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Your version:"
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
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   1200
      Width           =   1440
   End
   Begin VB.Label lblUpdateIntro 
      AutoSize        =   -1  'True
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
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4545
   End
End
Attribute VB_Name = "FormSoftwareUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Automatic Software Updater (note: it doesn't do the actual updating, it just CHECKS for updates!)
'Copyright ©2012-2013 by Tanner Helland
'Created: 19/August/12
'Last updated: 19/August/12
'Last update: initial build
'
'Interface for notifying the user that a new version of PhotoDemon is available for download.  This code is simply the
' notification part; the actual update checking is handled within the SoftwareUpdater module.
'
'Note that this code interfaces with the user preferences file so the user can opt to not check for updates and never be
' notified again. (FYI - this option can be enabled/disabled from the 'Tools' -> 'Options' menu.)
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Dim cImgCtl As clsControlImage

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'Do not download the update, but prompt the user again in the future
Private Sub cmdNoDownload_Click()
    
    g_UserPreferences.SetPref_Boolean "Updates", "Check For Updates", True
    
    Message "Automatic update canceled."
    
    cmdYesDownload.Visible = False
    cmdNoDownload.Visible = False
    cmdNoDownloadNoReminder.Visible = False
    picNo.Left = 0
    picNo.Top = 0
    'DoEvents
    txtNoExplanation.Text = g_Language.TranslateMessage("The next time you launch the program, it will remind you about this software update." & vbCrLf & vbCrLf & "Note: you can always manually download the latest version by visiting photodemon.org.")
    picNo.Visible = True
    cmdOKNo.Visible = True
    cmdOKNo.SetFocus
    
End Sub

'Do not download the update, and do not prompt the user again
Private Sub cmdNoDownloadNoReminder_Click()
    
    g_UserPreferences.SetPref_Boolean "Updates", "Check For Updates", False
    
    Message "Automatic update canceled."
    
    cmdYesDownload.Visible = False
    cmdNoDownload.Visible = False
    cmdNoDownloadNoReminder.Visible = False
    picNo.Left = 0
    picNo.Top = 0
    'DoEvents
    txtNoExplanation.Text = "You will no longer be prompted about software updates.  (If you change your mind in the future, this setting can be reversed from the 'Tools' -> 'Options' menu.)" & vbCrLf & vbCrLf & "Note: you can always manually download the latest version by visiting photodemon.org"
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
    OpenURL "http://www.photodemon.org/download"
    Message "Thanks for downloading the latest PhotoDemon update.  Hope you enjoy it!"
    Unload Me
End Sub

'LOAD form
Private Sub Form_Load()
    
    'Extract relevant icons from the resource file, and render them onto the buttons at run-time.
    ' (NOTE: because the icons require manifest theming, they will not appear in the IDE.)
    Set cImgCtl = New clsControlImage
    With cImgCtl
        .LoadImageFromStream cmdYesDownload.hWnd, LoadResData("LRGUPDATE", "CUSTOM"), 32, 32
        .LoadImageFromStream cmdNoDownload.hWnd, LoadResData("LRGDELAY", "CUSTOM"), 32, 32
        .LoadImageFromStream cmdNoDownloadNoReminder.hWnd, LoadResData("LRGCANCEL", "CUSTOM"), 32, 32
        
        .SetMargins cmdYesDownload.hWnd, 10
        .Align(cmdYesDownload.hWnd) = Icon_Left
        
        .SetMargins cmdNoDownload.hWnd, 10
        .Align(cmdNoDownload.hWnd) = Icon_Left
        
        .SetMargins cmdNoDownloadNoReminder.hWnd, 10
        .Align(cmdNoDownloadNoReminder.hWnd) = Icon_Left
        
    End With
    
    'Update the form labels with the version information
    lblCurrentAppVersion = App.Major & "." & App.Minor & "." & App.Revision
    lblNewestAppVersion = updateMajor & "." & updateMinor & "." & updateBuild
    
    'If available, display a link to the update notice of this software version
    If updateAnnouncement <> "" Then
        lblUpdateAnnouncement = g_Language.TranslateMessage("What's new in PhotoDemon") & " " & updateMajor & "." & updateMinor & "?"
    Else
        lblAnnouncementExplanation.Visible = False
        lblUpdateAnnouncement.Visible = False
    End If
    
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Open a browser window with the PhotoDemon download page
Private Sub lblDirectPDDownload_Click()
    OpenURL "http://www.photodemon.org/download"
End Sub

'When the user clicks the "more information" link, open a browser window with the blog article pulled from updates.txt
Private Sub lblUpdateAnnouncement_Click()
    OpenURL updateAnnouncement
End Sub
