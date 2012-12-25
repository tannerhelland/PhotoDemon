VERSION 5.00
Begin VB.Form FormPluginManager 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " PhotoDemon Plugin Manager"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9675
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
   ScaleHeight     =   7080
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstPlugins 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   4620
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label lblInterfaceTitle 
      AutoSize        =   -1  'True
      Caption         =   "FreeImage 3.15.4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C07031&
      Height          =   285
      Index           =   0
      Left            =   3120
      MouseIcon       =   "VBP_FormPluginManager.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   840
      Width           =   1890
   End
   Begin VB.Label lblInterfaceTitle 
      AutoSize        =   -1  'True
      Caption         =   "EZTwain 1.19"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C07031&
      Height          =   285
      Index           =   2
      Left            =   3120
      MouseIcon       =   "VBP_FormPluginManager.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   3600
      Width           =   1470
   End
   Begin VB.Label lblInterfaceTitle 
      AutoSize        =   -1  'True
      Caption         =   "zLib 1.2.5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C07031&
      Height          =   285
      Index           =   1
      Left            =   3120
      MouseIcon       =   "VBP_FormPluginManager.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2160
      Width           =   1050
   End
   Begin VB.Label lblInterfaceTitle 
      AutoSize        =   -1  'True
      Caption         =   "pngnq-s9 2.0.1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C07031&
      Height          =   285
      Index           =   3
      Left            =   3120
      MouseIcon       =   "VBP_FormPluginManager.frx":03F6
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   5040
      Width           =   1635
   End
   Begin VB.Label lblPluginStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GOOD"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000B909&
      Height          =   285
      Left            =   5520
      TabIndex        =   1
      Top             =   240
      Width           =   690
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "current plugin status:"
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
      Left            =   3120
      TabIndex        =   0
      Top             =   240
      Width           =   2265
   End
End
Attribute VB_Name = "FormPluginManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Plugin Manager
'Copyright ©2011-2012 by Tanner Helland
'Created: 21/December/12
'Last updated: 22/December/12
'Last update: finished initial build
'
'Dialog for presenting the user data related to the currently installed plugins.
'
'I seriously considered merging this form with the main Preferences (now Options) dialog, but there
' are simply too many settings present.  Rather than clutter up the main Preferences dialog with
' plugin-related settings, I have moved those all here.
'
'In the future, I suppose this could be merged with the plugin updater to form one happy plugin
' handler, but for now it makes sense to make them both available (and to keep them separate).
'
'***************************************************************************

Option Explicit

Private Sub Form_Load()
    
    'Populate the left-hand list box with all relevant plugins
    lstPlugins.Clear
    lstPlugins.AddItem "Overview", 0
    lstPlugins.AddItem "FreeImage", 1
    lstPlugins.AddItem "zLib", 2
    lstPlugins.AddItem "EZTwain", 3
    lstPlugins.AddItem "pngnq-s9", 4
    
    lstPlugins.ListIndex = 0
    
End Sub
