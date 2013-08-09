VERSION 5.00
Begin VB.Form FormPluginManager 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " PhotoDemon Plugin Manager"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10815
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
   ScaleHeight     =   411
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   721
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   7800
      TabIndex        =   0
      Top             =   5550
      Width           =   1365
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   9270
      TabIndex        =   1
      Top             =   5550
      Width           =   1365
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset all plugin options"
      Height          =   495
      Left            =   120
      TabIndex        =   71
      ToolTipText     =   "Use this to reset all plugin-related options to their default state.  This action cannot be undone."
      Top             =   5520
      Width           =   2775
   End
   Begin VB.ListBox lstPlugins 
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
      Height          =   4920
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   2775
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4935
      Index           =   0
      Left            =   3000
      ScaleHeight     =   329
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   513
      TabIndex        =   3
      Top             =   240
      Width           =   7695
      Begin VB.Label lblDisable 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Disable ExifTool"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C07031&
         Height          =   240
         Index           =   4
         Left            =   6015
         MouseIcon       =   "VBP_FormPluginManager.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   88
         Top             =   4125
         Width           =   1350
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "installed, enabled, and up to date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   270
         Index           =   4
         Left            =   1260
         TabIndex        =   87
         Top             =   4440
         Width           =   3255
      End
      Begin VB.Label lblInterfaceSubheader 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "status:"
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
         Height          =   270
         Index           =   4
         Left            =   480
         TabIndex        =   86
         Top             =   4440
         Width           =   675
      End
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ExifTool"
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
         Index           =   4
         Left            =   240
         TabIndex        =   85
         Top             =   4080
         Width           =   870
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
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   15
         Width           =   2265
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
         Left            =   2460
         TabIndex        =   20
         Top             =   15
         Width           =   690
      End
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "pngnq-s9"
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
         Index           =   3
         Left            =   240
         TabIndex        =   19
         Top             =   3240
         Width           =   1005
      End
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "zLib"
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
         Index           =   1
         Left            =   240
         TabIndex        =   18
         Top             =   1560
         Width           =   420
      End
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EZTwain"
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
         Index           =   2
         Left            =   240
         TabIndex        =   17
         Top             =   2400
         Width           =   915
      End
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FreeImage"
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
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label lblInterfaceSubheader 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "status:"
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
         Height          =   270
         Index           =   0
         Left            =   480
         TabIndex        =   15
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "installed, enabled, and up to date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   270
         Index           =   0
         Left            =   1260
         TabIndex        =   14
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Label lblInterfaceSubheader 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "status:"
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
         Height          =   270
         Index           =   1
         Left            =   480
         TabIndex        =   13
         Top             =   1920
         Width           =   675
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "installed, enabled, and up to date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   270
         Index           =   1
         Left            =   1260
         TabIndex        =   12
         Top             =   1920
         Width           =   3255
      End
      Begin VB.Label lblInterfaceSubheader 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "status:"
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
         Height          =   270
         Index           =   2
         Left            =   480
         TabIndex        =   11
         Top             =   2760
         Width           =   675
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "installed, enabled, and up to date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   270
         Index           =   2
         Left            =   1260
         TabIndex        =   10
         Top             =   2760
         Width           =   3255
      End
      Begin VB.Label lblInterfaceSubheader 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "status:"
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
         Height          =   270
         Index           =   3
         Left            =   480
         TabIndex        =   9
         Top             =   3600
         Width           =   675
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "installed, enabled, and up to date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   270
         Index           =   3
         Left            =   1260
         TabIndex        =   8
         Top             =   3600
         Width           =   3255
      End
      Begin VB.Label lblDisable 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Disable FreeImage"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C07031&
         Height          =   240
         Index           =   0
         Left            =   5760
         MouseIcon       =   "VBP_FormPluginManager.frx":0152
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   765
         Width           =   1605
      End
      Begin VB.Label lblDisable 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Disable zLib"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C07031&
         Height          =   240
         Index           =   1
         Left            =   6360
         MouseIcon       =   "VBP_FormPluginManager.frx":02A4
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   1605
         Width           =   1005
      End
      Begin VB.Label lblDisable 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Disable EZTwain"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C07031&
         Height          =   240
         Index           =   2
         Left            =   5955
         MouseIcon       =   "VBP_FormPluginManager.frx":03F6
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   2445
         Width           =   1410
      End
      Begin VB.Label lblDisable 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Disable pngnq-s9"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C07031&
         Height          =   240
         Index           =   3
         Left            =   5895
         MouseIcon       =   "VBP_FormPluginManager.frx":0548
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   3285
         Width           =   1470
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4935
      Index           =   4
      Left            =   3000
      ScaleHeight     =   329
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   513
      TabIndex        =   52
      Top             =   240
      Width           =   7695
      Begin PhotoDemon.smartCheckBox chkPngnqYUVA 
         Height          =   480
         Left            =   480
         TabIndex        =   74
         Top             =   2970
         Width           =   6195
         _ExtentX        =   10927
         _ExtentY        =   847
         Caption         =   "analyze colors using YUV instead of RGB (slower, but higher quality)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin PhotoDemon.smartCheckBox chkPngnqAlphaExtenuation 
         Height          =   480
         Left            =   480
         TabIndex        =   73
         Top             =   2520
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   847
         Caption         =   "when reducing alpha channels, attempt to preserve values of 0 and 255"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.HScrollBar hsPngnqDither 
         Height          =   255
         Left            =   2280
         Max             =   10
         TabIndex        =   68
         Top             =   4320
         Value           =   5
         Width           =   5295
      End
      Begin VB.HScrollBar hsPngnqSample 
         Height          =   255
         Left            =   2280
         Max             =   -1
         Min             =   -10
         TabIndex        =   64
         Top             =   3600
         Value           =   -3
         Width           =   5295
      End
      Begin VB.Label lblHSDescription 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "full persistence"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   3
         Left            =   6240
         TabIndex        =   70
         Top             =   4665
         Width           =   1080
      End
      Begin VB.Label lblHSDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "none"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   2
         Left            =   2550
         TabIndex        =   69
         Top             =   4665
         Width           =   360
      End
      Begin VB.Label lblPngnqSetting 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "color dithering:"
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
         Height          =   240
         Index           =   1
         Left            =   840
         TabIndex        =   67
         Top             =   4320
         Width           =   1305
      End
      Begin VB.Label lblHSDescription 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "large (slow, high quality)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   1
         Left            =   5520
         TabIndex        =   66
         Top             =   3945
         Width           =   1770
      End
      Begin VB.Label lblHSDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "small (fast, low quality)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   0
         Left            =   2550
         TabIndex        =   65
         Top             =   3945
         Width           =   1665
      End
      Begin VB.Label lblPngnqSetting 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "color sample size:"
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
         Height          =   240
         Index           =   0
         Left            =   570
         TabIndex        =   63
         Top             =   3600
         Width           =   1560
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "pngnq-s9 settings"
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
         Index           =   5
         Left            =   120
         TabIndex        =   62
         Top             =   2160
         Width           =   1890
      End
      Begin VB.Label lblLicenseLink 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "custom license, see PNGNQ-S9-LICENSE file"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C07031&
         Height          =   270
         Index           =   3
         Left            =   2760
         MouseIcon       =   "VBP_FormPluginManager.frx":069A
         MousePointer    =   99  'Custom
         TabIndex        =   61
         Top             =   1560
         Width           =   4290
      End
      Begin VB.Label lblLicense 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "pngnq-s9 license:"
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
         Height          =   270
         Index           =   3
         Left            =   480
         TabIndex        =   60
         Top             =   1560
         Width           =   1680
      End
      Begin VB.Label lblHomepageLink 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://sourceforge.net/projects/pngnqs9/"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C07031&
         Height          =   270
         Index           =   3
         Left            =   2760
         MouseIcon       =   "VBP_FormPluginManager.frx":07EC
         MousePointer    =   99  'Custom
         TabIndex        =   59
         Top             =   1080
         Width           =   4065
      End
      Begin VB.Label lblHomepage 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "pngnq-s9 homepage:"
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
         Height          =   270
         Index           =   3
         Left            =   480
         TabIndex        =   58
         Top             =   1080
         Width           =   2085
      End
      Begin VB.Label lbPluginSubheader 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "version found:"
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
         Height          =   270
         Index           =   13
         Left            =   3960
         TabIndex        =   57
         Top             =   600
         Width           =   1395
      End
      Begin VB.Label lblPluginVersionTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "XX.XX.XX"
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
         Height          =   270
         Index           =   3
         Left            =   2400
         TabIndex        =   56
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lblPluginVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "XX.XX.XX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   270
         Index           =   3
         Left            =   5520
         TabIndex        =   55
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lbPluginSubheader 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "expected version:"
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
         Height          =   270
         Index           =   12
         Left            =   480
         TabIndex        =   54
         Top             =   600
         Width           =   1740
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "pngnq-s9 plugin information"
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
         Index           =   4
         Left            =   120
         TabIndex        =   53
         Top             =   15
         Width           =   3045
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4935
      Index           =   5
      Left            =   3000
      ScaleHeight     =   329
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   513
      TabIndex        =   75
      Top             =   240
      Width           =   7695
      Begin VB.Label lblLicenseLink 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "open-source Perl license"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C07031&
         Height          =   270
         Index           =   4
         Left            =   2640
         MouseIcon       =   "VBP_FormPluginManager.frx":093E
         MousePointer    =   99  'Custom
         TabIndex        =   84
         Top             =   1560
         Width           =   2325
      End
      Begin VB.Label lblLicense 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ExifTool license:"
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
         Height          =   270
         Index           =   4
         Left            =   480
         TabIndex        =   83
         Top             =   1560
         Width           =   1545
      End
      Begin VB.Label lblHomepageLink 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.sno.phy.queensu.ca/~phil/exiftool/"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C07031&
         Height          =   270
         Index           =   4
         Left            =   2640
         MouseIcon       =   "VBP_FormPluginManager.frx":0A90
         MousePointer    =   99  'Custom
         TabIndex        =   82
         Top             =   1080
         Width           =   4575
      End
      Begin VB.Label lblHomepage 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ExifTool homepage:"
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
         Height          =   270
         Index           =   4
         Left            =   480
         TabIndex        =   81
         Top             =   1080
         Width           =   1950
      End
      Begin VB.Label lbPluginSubheader 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "version found:"
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
         Height          =   270
         Index           =   3
         Left            =   3960
         TabIndex        =   80
         Top             =   600
         Width           =   1395
      End
      Begin VB.Label lblPluginVersionTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "XX.XX.XX"
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
         Height          =   270
         Index           =   4
         Left            =   2400
         TabIndex        =   79
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lblPluginVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "XX.XX.XX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   270
         Index           =   4
         Left            =   5520
         TabIndex        =   78
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lbPluginSubheader 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "expected version:"
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
         Height          =   270
         Index           =   2
         Left            =   480
         TabIndex        =   77
         Top             =   600
         Width           =   1740
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ExifTool plugin information"
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
         Index           =   6
         Left            =   120
         TabIndex        =   76
         Top             =   15
         Width           =   2910
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4935
      Index           =   3
      Left            =   3000
      ScaleHeight     =   329
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   513
      TabIndex        =   42
      Top             =   240
      Width           =   7695
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EZTwain plugin information"
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
         Index           =   3
         Left            =   120
         TabIndex        =   51
         Top             =   15
         Width           =   2955
      End
      Begin VB.Label lbPluginSubheader 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "expected version:"
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
         Height          =   270
         Index           =   11
         Left            =   480
         TabIndex        =   50
         Top             =   600
         Width           =   1740
      End
      Begin VB.Label lblPluginVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "XX.XX.XX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   270
         Index           =   2
         Left            =   5520
         TabIndex        =   49
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lblPluginVersionTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "XX.XX.XX"
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
         Height          =   270
         Index           =   2
         Left            =   2400
         TabIndex        =   48
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lbPluginSubheader 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "version found:"
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
         Height          =   270
         Index           =   10
         Left            =   3960
         TabIndex        =   47
         Top             =   600
         Width           =   1395
      End
      Begin VB.Label lblHomepage 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EZTwain homepage:"
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
         Height          =   270
         Index           =   2
         Left            =   480
         TabIndex        =   46
         Top             =   1080
         Width           =   1995
      End
      Begin VB.Label lblHomepageLink 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.eztwain.com/eztwain1.htm"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C07031&
         Height          =   270
         Index           =   2
         Left            =   2640
         MouseIcon       =   "VBP_FormPluginManager.frx":0BE2
         MousePointer    =   99  'Custom
         TabIndex        =   45
         Top             =   1080
         Width           =   3780
      End
      Begin VB.Label lblLicense 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EZTwain license:"
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
         Height          =   270
         Index           =   2
         Left            =   480
         TabIndex        =   44
         Top             =   1560
         Width           =   1590
      End
      Begin VB.Label lblLicenseLink 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "public domain"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C07031&
         Height          =   270
         Index           =   2
         Left            =   2640
         MouseIcon       =   "VBP_FormPluginManager.frx":0D34
         MousePointer    =   99  'Custom
         TabIndex        =   43
         Top             =   1560
         Width           =   1305
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4935
      Index           =   2
      Left            =   3000
      ScaleHeight     =   329
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   513
      TabIndex        =   22
      Top             =   240
      Width           =   7695
      Begin VB.Label lblLicenseLink 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "zLib License"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C07031&
         Height          =   270
         Index           =   1
         Left            =   2280
         MouseIcon       =   "VBP_FormPluginManager.frx":0E86
         MousePointer    =   99  'Custom
         TabIndex        =   31
         Top             =   1560
         Width           =   1140
      End
      Begin VB.Label lblLicense 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "zLib license:"
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
         Height          =   270
         Index           =   1
         Left            =   480
         TabIndex        =   30
         Top             =   1560
         Width           =   1140
      End
      Begin VB.Label lblHomepageLink 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.zlib.net/"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C07031&
         Height          =   270
         Index           =   1
         Left            =   2280
         MouseIcon       =   "VBP_FormPluginManager.frx":0FD8
         MousePointer    =   99  'Custom
         TabIndex        =   29
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lblHomepage 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "zLib homepage:"
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
         Height          =   270
         Index           =   1
         Left            =   480
         TabIndex        =   28
         Top             =   1080
         Width           =   1545
      End
      Begin VB.Label lbPluginSubheader 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "version found:"
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
         Height          =   270
         Index           =   6
         Left            =   3960
         TabIndex        =   27
         Top             =   600
         Width           =   1395
      End
      Begin VB.Label lblPluginVersionTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "XX.XX.XX"
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
         Height          =   270
         Index           =   1
         Left            =   2400
         TabIndex        =   26
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lblPluginVersion 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "XX.XX.XX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   270
         Index           =   1
         Left            =   5520
         TabIndex        =   25
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lbPluginSubheader 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "expected version:"
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
         Height          =   270
         Index           =   7
         Left            =   480
         TabIndex        =   24
         Top             =   600
         Width           =   1740
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "zLib plugin information"
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
         Index           =   2
         Left            =   120
         TabIndex        =   23
         Top             =   15
         Width           =   2460
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4935
      Index           =   1
      Left            =   3000
      ScaleHeight     =   329
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   513
      TabIndex        =   32
      Top             =   240
      Width           =   7695
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FreeImage plugin information"
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
         Index           =   1
         Left            =   120
         TabIndex        =   41
         Top             =   15
         Width           =   3165
      End
      Begin VB.Label lbPluginSubheader 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "expected version:"
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
         Height          =   270
         Index           =   0
         Left            =   480
         TabIndex        =   40
         Top             =   600
         Width           =   1740
      End
      Begin VB.Label lblPluginVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "XX.XX.XX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   270
         Index           =   0
         Left            =   5520
         TabIndex        =   39
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lblPluginVersionTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "XX.XX.XX"
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
         Height          =   270
         Index           =   0
         Left            =   2400
         TabIndex        =   38
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lbPluginSubheader 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "version found:"
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
         Height          =   270
         Index           =   1
         Left            =   3960
         TabIndex        =   37
         Top             =   600
         Width           =   1395
      End
      Begin VB.Label lblHomepage 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FreeImage homepage:"
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
         Height          =   270
         Index           =   0
         Left            =   480
         TabIndex        =   36
         Top             =   1080
         Width           =   2265
      End
      Begin VB.Label lblHomepageLink 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://freeimage.sourceforge.net/"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C07031&
         Height          =   270
         Index           =   0
         Left            =   2880
         MouseIcon       =   "VBP_FormPluginManager.frx":112A
         MousePointer    =   99  'Custom
         TabIndex        =   35
         Top             =   1080
         Width           =   3330
      End
      Begin VB.Label lblLicense 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FreeImage license:"
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
         Height          =   270
         Index           =   0
         Left            =   480
         TabIndex        =   34
         Top             =   1560
         Width           =   1860
      End
      Begin VB.Label lblLicenseLink 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FreeImage Public License (FIPL)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C07031&
         Height          =   270
         Index           =   0
         Left            =   2880
         MouseIcon       =   "VBP_FormPluginManager.frx":127C
         MousePointer    =   99  'Custom
         TabIndex        =   33
         Top             =   1560
         Width           =   3150
      End
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   0
      TabIndex        =   72
      Top             =   5400
      Width           =   10815
   End
End
Attribute VB_Name = "FormPluginManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Plugin Manager
'Copyright ©2012-2013 by Tanner Helland
'Created: 21/December/12
'Last updated: 24/May/13
'Last update: added support for ExifTool
'
'Dialog for presenting the user data related to the currently installed plugins.
'
'I seriously considered merging this form with the main Preferences (now Options) dialog, but there
' are simply too many settings present.  Rather than clutter up the main Preferences dialog with
' plugin-related settings, I have moved those all here.
'
'In the future, I suppose this could be merged with the plugin updater to form one giant plugin
' handler, but for now it makes sense to make both available (and to keep them separate).
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'Green and red hues for use with our GOOD and BAD labels
Private Const GOODCOLOR As Long = 49152 'RGB(0,192,0)
Private Const BADCOLOR As Long = 192    'RGB(192,0,0)

'Much of the version-checking code used in this form was derived from http://allapi.mentalis.org/apilist/GetFileVersionInfo.shtml
' Many thanks to those authors for their work on demystifying some of these more obscure API calls
Private Type VS_FIXEDFILEINFO
   dwSignature As Long
   dwStrucVersionl As Integer     ' e.g. = &h0000 = 0
   dwStrucVersionh As Integer     ' e.g. = &h0042 = .42
   dwFileVersionMSl As Integer    ' e.g. = &h0003 = 3
   dwFileVersionMSh As Integer    ' e.g. = &h0075 = .75
   dwFileVersionLSl As Integer    ' e.g. = &h0000 = 0
   dwFileVersionLSh As Integer    ' e.g. = &h0031 = .31
   dwProductVersionMSl As Integer ' e.g. = &h0003 = 3
   dwProductVersionMSh As Integer ' e.g. = &h0010 = .1
   dwProductVersionLSl As Integer ' e.g. = &h0000 = 0
   dwProductVersionLSh As Integer ' e.g. = &h0031 = .31
   dwFileFlagsMask As Long        ' = &h3F for version "0.42"
   dwFileFlags As Long            ' e.g. VFF_DEBUG Or VFF_PRERELEASE
   dwFileOS As Long               ' e.g. VOS_DOS_WINDOWS16
   dwFileType As Long             ' e.g. VFT_DRIVER
   dwFileSubtype As Long          ' e.g. VFT2_DRV_KEYBOARD
   dwFileDateMS As Long           ' e.g. 0
   dwFileDateLS As Long           ' e.g. 0
End Type
Private Declare Function GetFileVersionInfo Lib "Version" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, ByVal Source As Long, ByVal Length As Long)

'This array will contain the full version strings of our various plugins
Dim vString(0 To 4) As String

'If the user presses "cancel", we need to restore the previous enabled/disabled values
Dim pEnabled(0 To 4) As Boolean

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

Private Sub CollectVersionInfo(ByVal FullFileName As String, ByVal strIndex As Long)
   
   Dim StrucVer As String, FileVer As String, ProdVer As String
   
   Dim rc As Long, lDummy As Long, sBuffer() As Byte
   Dim lBufferLen As Long, lVerPointer As Long, udtVerBuffer As VS_FIXEDFILEINFO
   Dim lVerbufferLen As Long

   '*** Get size ****
   lBufferLen = GetFileVersionInfoSize(FullFileName, lDummy)
   If lBufferLen < 1 Then
      'pdMsgBox "No Version Info available!"
      Exit Sub
   End If

   '**** Store info to udtVerBuffer struct ****
   ReDim sBuffer(lBufferLen)
   rc = GetFileVersionInfo(FullFileName, 0&, lBufferLen, sBuffer(0))
   rc = VerQueryValue(sBuffer(0), "\", lVerPointer, lVerbufferLen)
   MoveMemory udtVerBuffer, lVerPointer, Len(udtVerBuffer)

   '**** Determine Structure Version number - NOT USED ****
   StrucVer = Trim(Format$(udtVerBuffer.dwStrucVersionh)) & "." & Trim(Format$(udtVerBuffer.dwStrucVersionl))

   '**** Determine File Version number ****
   FileVer = Trim(Format$(udtVerBuffer.dwFileVersionMSh)) & "." & Trim(Format$(udtVerBuffer.dwFileVersionMSl)) & "." & Trim(Format$(udtVerBuffer.dwFileVersionLSh)) & "." & Trim(Format$(udtVerBuffer.dwFileVersionLSl))

   '**** Determine Product Version number ****
   ProdVer = Trim(Format$(udtVerBuffer.dwProductVersionMSh)) & "." & Trim(Format$(udtVerBuffer.dwProductVersionMSl)) & "." & Trim(Format$(udtVerBuffer.dwProductVersionLSh)) & "." & Trim(Format$(udtVerBuffer.dwProductVersionLSl))

   vString(strIndex) = ProdVer

End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    
    'Restore the original values for enabled or disabled plugins
    g_ImageFormats.FreeImageEnabled = pEnabled(0)
    g_ZLibEnabled = pEnabled(1)
    g_ScanEnabled = pEnabled(2)
    g_ImageFormats.pngnqEnabled = pEnabled(3)
    g_ExifToolEnabled = pEnabled(4)
    
    Unload Me
    
End Sub

'OK button
Private Sub CmdOK_Click()
    
    Message "Saving plugin options..."
    
    'Hide this form
    Me.Visible = False
    
    'Remember the current container the user is viewing
    g_UserPreferences.SetPref_Long "Plugin Preferences", "LastPluginPreferencesPage", lstPlugins.ListIndex
    
    'Save all plugin-specific settings to the preferences file
    
    'pngnq-s9 settings
        
        'Alpha extenuation
        g_UserPreferences.SetPref_Boolean "Plugin Preferences", "PngnqAlphaExtenuation", CBool(chkPngnqAlphaExtenuation.Value)
        
        'YUV
        g_UserPreferences.SetPref_Boolean "Plugin Preferences", "PngnqYUV", CBool(chkPngnqYUVA.Value)
        
        'Color sample size
        g_UserPreferences.SetPref_Long "Plugin Preferences", "PngnqColorSample", -1 * hsPngnqSample.Value
        
        'Dithering
        g_UserPreferences.SetPref_Long "Plugin Preferences", "PngnqDithering", hsPngnqDither.Value
            
    'Write all enabled/disabled plugin changes to the preferences file
    If g_ImageFormats.FreeImageEnabled Then
        g_UserPreferences.SetPref_Boolean "Plugin Preferences", "ForceFreeImageDisable", False
    Else
        g_UserPreferences.SetPref_Boolean "Plugin Preferences", "ForceFreeImageDisable", True
    End If
            
    'zLib
    If g_ZLibEnabled Then
        g_UserPreferences.SetPref_Boolean "Plugin Preferences", "ForceZLibDisable", False
    Else
        g_UserPreferences.SetPref_Boolean "Plugin Preferences", "ForceZLibDisable", True
    End If
        
    'EZTwain
    If g_ScanEnabled Then
        g_UserPreferences.SetPref_Boolean "Plugin Preferences", "ForceEZTwainDisable", False
    Else
        g_UserPreferences.SetPref_Boolean "Plugin Preferences", "ForceEZTwainDisable", True
    End If
        
    'pngnq-s9
    If g_ImageFormats.pngnqEnabled Then
        g_UserPreferences.SetPref_Boolean "Plugin Preferences", "ForcePngnqDisable", False
    Else
        g_UserPreferences.SetPref_Boolean "Plugin Preferences", "ForcePngnqDisable", True
    End If
    
    'ExifTool
    If g_ExifToolEnabled Then
        g_UserPreferences.SetPref_Boolean "Plugin Preferences", "ForceExifToolDisable", False
    Else
        g_UserPreferences.SetPref_Boolean "Plugin Preferences", "ForceExifToolDisable", True
    End If
    
    'If the user has changed any plugin enable/disable settings, a number of things must be refreshed program-wide
    If (pEnabled(0) <> g_ImageFormats.FreeImageEnabled) Or (pEnabled(1) <> g_ZLibEnabled) Or (pEnabled(2) <> g_ScanEnabled) Or (pEnabled(3) <> g_ImageFormats.pngnqEnabled) Or (pEnabled(4) <> g_ExifToolEnabled) Then
        LoadPlugins
        ApplyAllMenuIcons
        ResetMenuIcons
        g_ImageFormats.generateInputFormats
        g_ImageFormats.generateOutputFormats
    End If
    
    Message "Plugin options saved."
    
    Unload Me
    
End Sub

'RESET all plugin options
Private Sub cmdReset_Click()

    'Set current container to zero
    g_UserPreferences.SetPref_Long "Plugin Preferences", "LastPluginPreferencesPage", 0
    
    'Reset all plugin-specific settings in the preferences file
    
    'pngnq-s9 settings
        
        'Alpha extenuation
        g_UserPreferences.SetPref_Boolean "Plugin Preferences", "PngnqAlphaExtenuation", False
        
        'YUV
        g_UserPreferences.SetPref_Boolean "Plugin Preferences", "PngnqYUV", True
        
        'Color sample size
        g_UserPreferences.SetPref_Long "Plugin Preferences", "PngnqColorSample", 3
        
        'Dithering
        g_UserPreferences.SetPref_Long "Plugin Preferences", "PngnqDithering", 5

    'Enable all plugins if possible
    g_UserPreferences.SetPref_Boolean "Plugin Preferences", "ForceFreeImageDisable", False
    g_UserPreferences.SetPref_Boolean "Plugin Preferences", "ForceZLibDisable", False
    g_UserPreferences.SetPref_Boolean "Plugin Preferences", "ForceEZTwainDisable", False
    g_UserPreferences.SetPref_Boolean "Plugin Preferences", "ForcePngnqDisable", False
    g_UserPreferences.SetPref_Boolean "Plugin Preferences", "ForceExifToolDisable", False
    
    'Reload the plugins (from a system standpoint)
    LoadPlugins
    
    'Reload the dialog
    LoadAllPluginSettings
    
End Sub

'LOAD the form
Private Sub Form_Load()
    
    'Remember which plugins the user has enabled or disabled
    pEnabled(0) = g_ImageFormats.FreeImageEnabled
    pEnabled(1) = g_ZLibEnabled
    pEnabled(2) = g_ScanEnabled
    pEnabled(3) = g_ImageFormats.pngnqEnabled
    pEnabled(4) = g_ExifToolEnabled
    
    'Populate the left-hand list box with all relevant plugins
    lstPlugins.Clear
    lstPlugins.AddItem "Overview", 0
    lstPlugins.AddItem "FreeImage", 1
    lstPlugins.AddItem "zLib", 2
    lstPlugins.AddItem "EZTwain", 3
    lstPlugins.AddItem "pngnq-s9", 4
    lstPlugins.AddItem "ExifTool", 5
    
    lstPlugins.ListIndex = 0
    
    'Load all user-editable settings from the preferences file, and populate all plugin information
    LoadAllPluginSettings
        
    'For some reason, the container picture boxes automatically acquire the pointer of children objects.
    ' Manually force those cursors to arrows to prevent this.
    Dim i As Long
    For i = 0 To picContainer.Count - 1
        setArrowCursorToObject picContainer(i)
    Next i
            
    'Apply visual styles
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'If a translation is active, realign text as necessary
    If g_Language.translationActive Then
        lblPluginStatus.Left = lblTitle(0).Left + lblTitle(0).Width + 8
        
        For i = 0 To lblStatus.Count - 1
            lblStatus(i).Left = lblInterfaceSubheader(i).Left + lblInterfaceSubheader(i).Width + 8
        Next i
        
        For i = 0 To lblHomepage.Count - 1
            lblHomepageLink(i).Left = lblHomepage(i).Left + lblHomepage(i).Width + 8
            lblLicenseLink(i).Left = lblLicense(i).Left + lblLicense(i).Width + 8
        Next i
        
    End If
    
End Sub

'When the dialog is first launched, use this to populate the dialog with any settings the user may have modified
Private Sub LoadAllPluginSettings()

    'Now, check version numbers of each plugin.  This is more complicated than it needs to be, on account of
    ' each plugin having its own unique mechanism for version-checking, but I have wrapped these various functions
    ' inside fairly standard wrapper calls.
    CollectAllVersionNumbers
    
    'We now have a collection of version numbers for our various plugins.  Let's use those to populate our
    ' "good/bad" labels for each plugin.
    UpdatePluginLabels
    
    'Hide all plugin control containers
    Dim i As Long
    For i = 0 To picContainer.Count - 1
        picContainer(i).Visible = False
    Next i
    
    'Enable the last container the user selected
    lstPlugins.ListIndex = g_UserPreferences.GetPref_Long("Plugin Preferences", "LastPluginPreferencesPage", 0)
    picContainer(lstPlugins.ListIndex).Visible = True
    
    'Load all plugin settings from the preferences file
    
    'pngnq-s9 settings
        
        'Alpha extenuation
        If g_UserPreferences.GetPref_Boolean("Plugin Preferences", "PngnqAlphaExtenuation", False) Then chkPngnqAlphaExtenuation.Value = vbChecked Else chkPngnqAlphaExtenuation.Value = vbUnchecked
        
        'YUV
        If g_UserPreferences.GetPref_Boolean("Plugin Preferences", "PngnqYUV", True) Then chkPngnqYUVA.Value = vbChecked Else chkPngnqYUVA.Value = vbUnchecked
        
        'Color sample size
        hsPngnqSample.Value = -1 * g_UserPreferences.GetPref_Long("Plugin Preferences", "PngnqColorSample", 3)
        
        'Dithering
        hsPngnqDither.Value = g_UserPreferences.GetPref_Long("Plugin Preferences", "PngnqDithering", 5)
        
End Sub

'Assuming version numbers have been successfully retrieved, this function can be called to update the
' green/red plugin label display on the main panel.
Private Sub UpdatePluginLabels()
    
    Dim pluginStatus As Boolean
    
    'FreeImage
    pluginStatus = popPluginLabel(0, "FreeImage", "3.15.4", isFreeImageAvailable, g_ImageFormats.FreeImageEnabled)
    
    'zLib
    pluginStatus = pluginStatus And popPluginLabel(1, "zLib", "1.2.8", isZLibAvailable, g_ZLibEnabled)
    
    'EZTwain
    pluginStatus = pluginStatus And popPluginLabel(2, "EZTwain", "1.18.0", isEZTwainAvailable, g_ScanEnabled)
    
    'pngnq-s9
    pluginStatus = pluginStatus And popPluginLabel(3, "pngnq-s9", "2.0.1", isPngnqAvailable, g_ImageFormats.pngnqEnabled)
    
    'ExifTool
    pluginStatus = pluginStatus And popPluginLabel(4, "ExifTool", "9.29", isExifToolAvailable, g_ExifToolEnabled)
    
    If pluginStatus Then
        lblPluginStatus.ForeColor = GOODCOLOR
        lblPluginStatus.Caption = UCase(g_Language.TranslateMessage("GOOD"))
    Else
        lblPluginStatus.ForeColor = BADCOLOR
        lblPluginStatus.Caption = g_Language.TranslateMessage("problems detected")
    End If
        
End Sub

'Retrieve all relevant plugin version numbers and store them in the vString() array
Private Sub CollectAllVersionNumbers()

    'Start by analyzing plugin file metadata for version information.  This works for FreeImage and zLib (but
    ' do it for all four, just in case).
    If isFreeImageAvailable Then CollectVersionInfo g_PluginPath & "freeimage.dll", 0 Else vString(0) = "none"
    If isZLibAvailable Then CollectVersionInfo g_PluginPath & "zlibwapi.dll", 1 Else vString(1) = "none"
    If isEZTwainAvailable Then CollectVersionInfo g_PluginPath & "eztw32.dll", 2 Else vString(2) = "none"
    If isPngnqAvailable Then CollectVersionInfo g_PluginPath & "pngnq-s9.exe", 3 Else vString(3) = "none"
    If isExifToolAvailable Then CollectVersionInfo g_PluginPath & "exiftool.exe", 4 Else vString(4) = "none"
    
    'Special version-checking techniques are required for some plugins.
    
    'The EZTwain DLL provides its own version-checking function
    If isEZTwainAvailable Then vString(2) = getEZTwainVersion Else vString(2) = "none"
    
    'pngnq-s9 doesn't have a version-checking function, and it doesn't write version numbers to the .exe.
    ' So for now, assume the latest version automatically.
    If isPngnqAvailable Then vString(3) = getPngnqVersion Else vString(3) = "none"
    
    'ExifTool can write its version number to stdout.  Capture that now.
    If isExifToolAvailable Then vString(4) = getExifToolVersion() Else vString(4) = "none"
    
    'Remove trailing build numbers from version strings as necessary.  (Note: (4) is left off, as ExifTool
    ' does not report a build number)
    Dim i As Long
    For i = 0 To 3
        If vString(i) <> "none" Then StripOffExtension vString(i)
    Next i

End Sub

'Given a plugin's availability, expected version, and index on this form, populate the relevant labels associated with it.
' This function will return TRUE if the plugin is in good status, FALSE if it isn't (for any reason)
Private Function popPluginLabel(ByVal curPlugin As Long, ByRef pluginName As String, ByRef expectedVersion As String, ByVal isAvailable As Boolean, ByVal isEnabled As Boolean) As Boolean
        
    'Make the individual plugin panels display the expected version
    lblPluginVersionTitle(curPlugin).Caption = expectedVersion
        
    'Is this plugin present on the machine?
    If isAvailable Then
    
        'Make the individual plugin panels display the discovered version
        lblPluginVersion(curPlugin).Caption = vString(curPlugin)
        If StrComp(vString(curPlugin), expectedVersion, vbTextCompare) = 0 Then
            lblPluginVersion(curPlugin).ForeColor = GOODCOLOR
        Else
            lblPluginVersion(curPlugin).ForeColor = BADCOLOR
        End If
        
        'If present, has it been forcibly disabled?
        If isEnabled Then
            lblStatus(curPlugin).Caption = g_Language.TranslateMessage("installed")
            lblDisable(curPlugin).Caption = g_Language.TranslateMessage("disable") & " " & pluginName
            
            'If this plugin is present and enabled, does its version match what we expect?
            If StrComp(vString(curPlugin), expectedVersion, vbTextCompare) = 0 Then
                lblStatus(curPlugin).Caption = lblStatus(curPlugin).Caption & " " & g_Language.TranslateMessage("and up to date")
                lblStatus(curPlugin).ForeColor = GOODCOLOR
                popPluginLabel = True
                
            'Version mismatch
            Else
                lblStatus(curPlugin).Caption = lblStatus(curPlugin).Caption & ", " & g_Language.TranslateMessage("but incorrect version (%1 found, %2 expected)", vString(curPlugin), expectedVersion)
                lblStatus(curPlugin).ForeColor = BADCOLOR
                popPluginLabel = False
            End If
            
        'Plugin is disabled
        Else
            lblStatus(curPlugin).Caption = g_Language.TranslateMessage("installed, but disabled by user")
            lblStatus(curPlugin).ForeColor = BADCOLOR
            lblDisable(curPlugin).Caption = g_Language.TranslateMessage("enable") & " " & pluginName
            popPluginLabel = False
        End If
        
    'Plugin is not present on the machine
    Else
        lblStatus(curPlugin).Caption = g_Language.TranslateMessage("missing")
        lblStatus(curPlugin).ForeColor = BADCOLOR
        lblDisable(curPlugin).Visible = False
        popPluginLabel = False
        lblPluginVersion(curPlugin).Caption = g_Language.TranslateMessage("missing")
        lblPluginVersion(curPlugin).ForeColor = BADCOLOR
    End If
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'The user is now allowed to selectively disable/enable various plugins.  This can be used to test certain program
' parameters, or to force certain behaviors.
Private Sub lblDisable_Click(Index As Integer)

    Select Case Index
    
        'FreeImage
        Case 0
            g_ImageFormats.FreeImageEnabled = Not g_ImageFormats.FreeImageEnabled
            
        'zLib
        Case 1
            g_ZLibEnabled = Not g_ZLibEnabled
            
        'EZTwain
        Case 2
            g_ScanEnabled = Not g_ScanEnabled
            
        'pngnq-s9
        Case 3
            g_ImageFormats.pngnqEnabled = Not g_ImageFormats.pngnqEnabled
        
        'ExifTool
        Case 4
            g_ExifToolEnabled = Not g_ExifToolEnabled
            
    End Select
    
    'Update the various labels to match the new situation
    UpdatePluginLabels

End Sub

Private Sub lblHomepageLink_Click(Index As Integer)

    Select Case Index
        
        'FreeImage
        Case 0
            OpenURL "http://freeimage.sourceforge.net/"
            
        'zLib
        Case 1
            OpenURL "http://www.zlib.net/"
        
        'ezTwain
        Case 2
            OpenURL "http://www.eztwain.com/eztwain1.htm"
        
        'pngnq-s9
        Case 3
            OpenURL "http://sourceforge.net/projects/pngnqs9/"
            
        'ExifTool
        Case 4
            OpenURL "http://www.sno.phy.queensu.ca/~phil/exiftool/"
        
    End Select

End Sub

Private Sub lblLicenseLink_Click(Index As Integer)

    Select Case Index
        
        'FreeImage
        Case 0
            OpenURL "http://freeimage.sourceforge.net/freeimage-license.txt"
            
        'zLib
        Case 1
            OpenURL "http://www.zlib.net/zlib_license.html"
        
        'ezTwain
        Case 2
            OpenURL "http://www.eztwain.com/ezt1faq.htm"
            
        'pngnq-s9
        Case 3
            OpenURL "http://sourceforge.net/projects/pngnqs9/files/"
            
        'ExifTool
        Case 4
            OpenURL "http://www.sno.phy.queensu.ca/~phil/exiftool/#license"
        
    End Select
    
End Sub

'When a new plugin is selected, display only the relevant plugin panel
Private Sub lstPlugins_Click()

    Dim i As Long
    For i = 0 To picContainer.Count - 1
        If i = lstPlugins.ListIndex Then picContainer(i).Visible = True Else picContainer(i).Visible = False
    Next i
    
End Sub
