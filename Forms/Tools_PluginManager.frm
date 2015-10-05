VERSION 5.00
Begin VB.Form FormPluginManager 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " PhotoDemon Plugin Manager"
   ClientHeight    =   7125
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
   Icon            =   "Tools_PluginManager.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   475
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   721
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.commandBarMini cmdBarMini 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   81
      Top             =   6375
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   1323
      BackColor       =   14802140
   End
   Begin PhotoDemon.pdButton cmdReset 
      Height          =   615
      Left            =   120
      TabIndex        =   80
      Top             =   5640
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1085
      Caption         =   "Reset all plugin options"
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
      Height          =   5220
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5895
      Index           =   0
      Left            =   3000
      ScaleHeight     =   393
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   513
      TabIndex        =   1
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
         MouseIcon       =   "Tools_PluginManager.frx":000C
         MousePointer    =   99  'Custom
         TabIndex        =   78
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
         TabIndex        =   77
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
         TabIndex        =   76
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
         TabIndex        =   75
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
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   15
         Width           =   690
      End
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PNGQuant"
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
         TabIndex        =   17
         Top             =   3240
         Width           =   1110
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         MouseIcon       =   "Tools_PluginManager.frx":015E
         MousePointer    =   99  'Custom
         TabIndex        =   5
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
         MouseIcon       =   "Tools_PluginManager.frx":02B0
         MousePointer    =   99  'Custom
         TabIndex        =   4
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
         MouseIcon       =   "Tools_PluginManager.frx":0402
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   2445
         Width           =   1410
      End
      Begin VB.Label lblDisable 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Disable PNGQuant"
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
         Left            =   5835
         MouseIcon       =   "Tools_PluginManager.frx":0554
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   3285
         Width           =   1530
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5895
      Index           =   4
      Left            =   3000
      ScaleHeight     =   393
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   513
      TabIndex        =   50
      Top             =   240
      Width           =   7695
      Begin PhotoDemon.smartCheckBox chkPNGQuantIE6 
         Height          =   330
         Left            =   480
         TabIndex        =   64
         Top             =   2970
         Width           =   7050
         _ExtentX        =   12435
         _ExtentY        =   582
         Caption         =   "improve IE6 compatibility (reduces image quality; use with caution)"
      End
      Begin PhotoDemon.smartCheckBox chkPNGQuantDither 
         Height          =   330
         Left            =   480
         TabIndex        =   63
         Top             =   2520
         Width           =   7050
         _ExtentX        =   12435
         _ExtentY        =   582
         Caption         =   "use dithering to improve output"
      End
      Begin PhotoDemon.sliderTextCombo sltPNGQuantSpeed 
         Height          =   675
         Left            =   480
         TabIndex        =   79
         Top             =   3720
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   1191
         Caption         =   "performance vs image quality"
         FontSizeCaption =   10
         Min             =   1
         Max             =   11
         SliderTrackStyle=   1
         Value           =   3
         NotchPosition   =   2
         NotchValueCustom=   3
      End
      Begin VB.Label lblHSDescription 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "fast, low quality"
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
         Left            =   5280
         TabIndex        =   62
         Top             =   4560
         Width           =   1155
      End
      Begin VB.Label lblHSDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "slow, high quality"
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
         Left            =   960
         TabIndex        =   61
         Top             =   4560
         Width           =   1245
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PNGQuant settings"
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
         TabIndex        =   60
         Top             =   2160
         Width           =   1995
      End
      Begin VB.Label lblLicenseLink 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BSD license"
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
         Left            =   3060
         MouseIcon       =   "Tools_PluginManager.frx":06A6
         MousePointer    =   99  'Custom
         TabIndex        =   59
         Top             =   1560
         Width           =   1110
      End
      Begin VB.Label lblLicense 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PNGQuant license:"
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
         Top             =   1560
         Width           =   1800
      End
      Begin VB.Label lblHomepageLink 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://pngquant.org/"
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
         Left            =   3060
         MouseIcon       =   "Tools_PluginManager.frx":07F8
         MousePointer    =   99  'Custom
         TabIndex        =   57
         Top             =   1080
         Width           =   2040
      End
      Begin VB.Label lblHomepage 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PNGQuant homepage:"
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
         TabIndex        =   56
         Top             =   1080
         Width           =   2205
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
         TabIndex        =   55
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
         TabIndex        =   54
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
         TabIndex        =   53
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
         TabIndex        =   52
         Top             =   600
         Width           =   1740
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PNGQuant plugin information"
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
         TabIndex        =   51
         Top             =   15
         Width           =   3150
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5895
      Index           =   5
      Left            =   3000
      ScaleHeight     =   393
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   513
      TabIndex        =   65
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
         MouseIcon       =   "Tools_PluginManager.frx":094A
         MousePointer    =   99  'Custom
         TabIndex        =   74
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
         TabIndex        =   73
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
         MouseIcon       =   "Tools_PluginManager.frx":0A9C
         MousePointer    =   99  'Custom
         TabIndex        =   72
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
         TabIndex        =   71
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
         TabIndex        =   70
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
         TabIndex        =   69
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
         TabIndex        =   68
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
         TabIndex        =   67
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
         TabIndex        =   66
         Top             =   15
         Width           =   2910
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5895
      Index           =   3
      Left            =   3000
      ScaleHeight     =   393
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   513
      TabIndex        =   40
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
         TabIndex        =   49
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
         TabIndex        =   48
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
         TabIndex        =   47
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
         TabIndex        =   46
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
         TabIndex        =   45
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
         TabIndex        =   44
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
         MouseIcon       =   "Tools_PluginManager.frx":0BEE
         MousePointer    =   99  'Custom
         TabIndex        =   43
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
         TabIndex        =   42
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
         MouseIcon       =   "Tools_PluginManager.frx":0D40
         MousePointer    =   99  'Custom
         TabIndex        =   41
         Top             =   1560
         Width           =   1305
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5895
      Index           =   2
      Left            =   3000
      ScaleHeight     =   393
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   513
      TabIndex        =   20
      Top             =   240
      Width           =   7695
      Begin VB.Label lblLicenseLink 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "zLib license"
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
         MouseIcon       =   "Tools_PluginManager.frx":0E92
         MousePointer    =   99  'Custom
         TabIndex        =   29
         Top             =   1560
         Width           =   1065
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
         TabIndex        =   28
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
         MouseIcon       =   "Tools_PluginManager.frx":0FE4
         MousePointer    =   99  'Custom
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
         Top             =   15
         Width           =   2460
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5895
      Index           =   1
      Left            =   3000
      ScaleHeight     =   393
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   513
      TabIndex        =   30
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   35
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
         TabIndex        =   34
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
         MouseIcon       =   "Tools_PluginManager.frx":1136
         MousePointer    =   99  'Custom
         TabIndex        =   33
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
         TabIndex        =   32
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
         MouseIcon       =   "Tools_PluginManager.frx":1288
         MousePointer    =   99  'Custom
         TabIndex        =   31
         Top             =   1560
         Width           =   3150
      End
   End
End
Attribute VB_Name = "FormPluginManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Plugin Manager
'Copyright 2012-2015 by Tanner Helland
'Created: 21/December/12
'Last updated: 02/July/14
'Last update: replaced all pngnq-s9 interactions with PNGQuant
'
'Dialog for presenting the user data related to the currently installed plugins.
'
'I seriously considered merging this form with the main Preferences (now Options) dialog, but there are simply
' too many settings present.  Rather than clutter up the main Preferences dialog with plugin-related settings,
' I have moved those all here.
'
'In the future, I suppose this could be merged with the plugin updater to form one giant plugin handler, but
' for now it makes sense to make both available (and to keep them separate).
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
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
Private vString(0 To 4) As String

'If the user presses "cancel", we need to restore the previous enabled/disabled values
Private pEnabled(0 To 4) As Boolean

Private Sub CollectVersionInfo(ByVal FullFileName As String, ByVal strIndex As Long)
   
   Dim StrucVer As String, FileVer As String, ProdVer As String
   
   Dim rc As Long, lDummy As Long, sBuffer() As Byte
   Dim lBufferLen As Long, lVerPointer As Long, udtVerBuffer As VS_FIXEDFILEINFO
   Dim lVerbufferLen As Long

   '*** Get size ****
   lBufferLen = GetFileVersionInfoSize(FullFileName, lDummy)
   If lBufferLen < 1 Then
      'Debug.Print "Could not retrieve version information for (" & FullFileName & ")"
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

Private Sub cmdBarMini_CancelClick()
    
    'Restore the original values for enabled or disabled plugins
    g_ImageFormats.FreeImageEnabled = pEnabled(0)
    g_ZLibEnabled = pEnabled(1)
    g_ScanEnabled = pEnabled(2)
    g_ImageFormats.pngQuantEnabled = pEnabled(3)
    g_ExifToolEnabled = pEnabled(4)
    
End Sub

Private Sub cmdBarMini_OKClick()
    
    Message "Saving plugin options..."
    
    'Hide this form
    Me.Visible = False
    
    'Remember the current container the user is viewing
    g_UserPreferences.startBatchPreferenceMode
    g_UserPreferences.SetPref_Long "Plugins", "Last Plugin Preferences Page", lstPlugins.ListIndex
    
    'Save all plugin-specific settings to the preferences file
    
    'PNGQuant settings
        
        'Dithering
        g_UserPreferences.SetPref_Boolean "Plugins", "PNGQuant Dithering", CBool(chkPNGQuantDither.Value)
        
        'IE6 compatibility
        g_UserPreferences.SetPref_Boolean "Plugins", "PNGQuant IE6 Compatibility", CBool(chkPNGQuantIE6.Value)
        
        'Performance vs speed
        g_UserPreferences.SetPref_Long "Plugins", "PNGQuant Performance", sltPNGQuantSpeed.Value
            
            
    'Write all enabled/disabled plugin changes to the preferences file
    If g_ImageFormats.FreeImageEnabled Then
        g_UserPreferences.SetPref_Boolean "Plugins", "Force FreeImage Disable", False
    Else
        g_UserPreferences.SetPref_Boolean "Plugins", "Force FreeImage Disable", True
    End If
            
    'zLib
    If g_ZLibEnabled Then
        g_UserPreferences.SetPref_Boolean "Plugins", "Force ZLib Disable", False
    Else
        g_UserPreferences.SetPref_Boolean "Plugins", "Force ZLib Disable", True
    End If
        
    'EZTwain
    If g_ScanEnabled Then
        g_UserPreferences.SetPref_Boolean "Plugins", "Force EZTwain Disable", False
    Else
        g_UserPreferences.SetPref_Boolean "Plugins", "Force EZTwain Disable", True
    End If
        
    'PNGQuant
    If g_ImageFormats.pngQuantEnabled Then
        g_UserPreferences.SetPref_Boolean "Plugins", "Force PNGQuant Disable", False
    Else
        g_UserPreferences.SetPref_Boolean "Plugins", "Force PNGQuant Disable", True
    End If
    
    'ExifTool
    If g_ExifToolEnabled Then
        g_UserPreferences.SetPref_Boolean "Plugins", "Force ExifTool Disable", False
    Else
        g_UserPreferences.SetPref_Boolean "Plugins", "Force ExifTool Disable", True
    End If
    
    'If the user has changed any plugin enable/disable settings, a number of things must be refreshed program-wide
    If (pEnabled(0) <> g_ImageFormats.FreeImageEnabled) Or (pEnabled(1) <> g_ZLibEnabled) Or (pEnabled(2) <> g_ScanEnabled) Or (pEnabled(3) <> g_ImageFormats.pngQuantEnabled) Or (pEnabled(4) <> g_ExifToolEnabled) Then
        Plugin_Management.LoadAllPlugins
        applyAllMenuIcons
        resetMenuIcons
        g_ImageFormats.generateInputFormats
        g_ImageFormats.generateOutputFormats
    End If
    
    'End batch preference update mode, which will force a write-to-file operation
    g_UserPreferences.endBatchPreferenceMode
    
    Message "Plugin options saved."
    
End Sub

'RESET all plugin options
Private Sub cmdReset_Click()

    'Set current container to zero
    g_UserPreferences.SetPref_Long "Plugins", "Last Plugin Preferences Page", 0
    
    'Reset all plugin-specific settings in the preferences file
    
    'PNGQuant settings
        
        'Dithering
        g_UserPreferences.SetPref_Boolean "Plugins", "PNGQuant Dithering", True
        
        'IE6 compatibility
        g_UserPreferences.SetPref_Boolean "Plugins", "PNGQuant IE6 Compatibility", False
        
        'Performance vs speed
        g_UserPreferences.SetPref_Long "Plugins", "PNGQuant Performance", 3

    'Enable all plugins if possible
    g_UserPreferences.SetPref_Boolean "Plugins", "Force FreeImage Disable", False
    g_UserPreferences.SetPref_Boolean "Plugins", "Force ZLib Disable", False
    g_UserPreferences.SetPref_Boolean "Plugins", "Force EZTwain Disable", False
    g_UserPreferences.SetPref_Boolean "Plugins", "Force PNGQuant Disable", False
    g_UserPreferences.SetPref_Boolean "Plugins", "Force ExifTool Disable", False
    
    'Reload the plugins (from a system standpoint)
    Plugin_Management.LoadAllPlugins
    
    'Reload the dialog
    LoadAllPluginSettings
    
End Sub

'LOAD the form
Private Sub Form_Load()
    
    'Remember which plugins the user has enabled or disabled
    pEnabled(0) = g_ImageFormats.FreeImageEnabled
    pEnabled(1) = g_ZLibEnabled
    pEnabled(2) = g_ScanEnabled
    pEnabled(3) = g_ImageFormats.pngQuantEnabled
    pEnabled(4) = g_ExifToolEnabled
    
    'Populate the left-hand list box with all relevant plugins
    lstPlugins.Clear
    lstPlugins.AddItem "Overview", 0
    lstPlugins.AddItem "FreeImage", 1
    lstPlugins.AddItem "zLib", 2
    lstPlugins.AddItem "EZTwain", 3
    lstPlugins.AddItem "PNGQuant", 4
    lstPlugins.AddItem "ExifTool", 5
    
    lstPlugins.ListIndex = 0
    
    'Load all user-editable settings from the preferences file, and populate all plugin information
    LoadAllPluginSettings
        
    'For some reason, the container picture boxes automatically acquire the pointer of children objects.
    ' Manually force those cursors to arrows to prevent this.
    Dim i As Long
    For i = 0 To picContainer.Count - 1
        setArrowCursor picContainer(i)
    Next i
            
    'Apply translations and visual themes
    MakeFormPretty Me
    
    'If a translation is active, realign text as necessary
    If g_Language.translationActive Then
        lblPluginStatus.Left = lblTitle(0).Left + lblTitle(0).Width + FixDPI(8)
        
        For i = 0 To lblStatus.Count - 1
            lblStatus(i).Left = lblInterfaceSubheader(i).Left + lblInterfaceSubheader(i).Width + FixDPI(8)
        Next i
        
        For i = 0 To lblHomepage.Count - 1
            lblHomepageLink(i).Left = lblHomepage(i).Left + lblHomepage(i).Width + FixDPI(8)
            lblLicenseLink(i).Left = lblLicense(i).Left + lblLicense(i).Width + FixDPI(8)
        Next i
        
    End If
    
End Sub

'When the dialog is first launched, use this to populate the dialog with any settings the user may have modified
Private Sub LoadAllPluginSettings()
    
    'Start batch preference processing mode.
    g_UserPreferences.startBatchPreferenceMode
    
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
    lstPlugins.ListIndex = g_UserPreferences.GetPref_Long("Plugins", "Last Plugin Preferences Page", 0)
    picContainer(lstPlugins.ListIndex).Visible = True
    
    'Load all specialty plugin settings from the preferences file
    
    'PNGQuant settings
        
        'Dithering
        If g_UserPreferences.GetPref_Boolean("Plugins", "PNGQuant Dithering", True) Then chkPNGQuantDither.Value = vbChecked Else chkPNGQuantDither.Value = vbUnchecked
        
        'IE6 compatibility
        If g_UserPreferences.GetPref_Boolean("Plugins", "PNGQuant IE6 Compatibility", False) Then chkPNGQuantIE6.Value = vbChecked Else chkPNGQuantIE6.Value = vbUnchecked
        
        'Performance vs speed
        sltPNGQuantSpeed.Value = g_UserPreferences.GetPref_Long("Plugins", "PNGQuant Performance", 3)
        
    'End batch preference mode
    g_UserPreferences.endBatchPreferenceMode
        
End Sub

'Assuming version numbers have been successfully retrieved, this function can be called to update the
' green/red plugin label display on the main panel.
Private Sub UpdatePluginLabels()
    
    Dim pluginStatus As Boolean
    
    'FreeImage
    pluginStatus = popPluginLabel(0, "FreeImage", EXPECTED_FREEIMAGE_VERSION, isFreeImageAvailable, g_ImageFormats.FreeImageEnabled)
    
    'zLib
    pluginStatus = pluginStatus And popPluginLabel(1, "zLib", EXPECTED_ZLIB_VERSION, isZLibAvailable, g_ZLibEnabled)
    
    'EZTwain
    pluginStatus = pluginStatus And popPluginLabel(2, "EZTwain", EXPECTED_EZTWAIN_VERSION, isEZTwainAvailable, g_ScanEnabled)
    
    'PNGQuant
    pluginStatus = pluginStatus And popPluginLabel(3, "PNGQuant", EXPECTED_PNGQUANT_VERSION, isPngQuantAvailable, g_ImageFormats.pngQuantEnabled)
    
    'ExifTool
    pluginStatus = pluginStatus And popPluginLabel(4, "ExifTool", EXPECTED_EXIFTOOL_VERSION, isExifToolAvailable, g_ExifToolEnabled)
    
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
    If isPngQuantAvailable Then CollectVersionInfo g_PluginPath & "pngquant.exe", 3 Else vString(3) = "none"
    If isExifToolAvailable Then CollectVersionInfo g_PluginPath & "exiftool.exe", 4 Else vString(4) = "none"
    
    'Special version-checking techniques are required for some plugins.
    
    'The EZTwain DLL provides its own version-checking function
    If isEZTwainAvailable Then vString(2) = getEZTwainVersion Else vString(2) = "none"
    
    'PNGQuant can write its version number to stdout.  Capture that now.
    If isPngQuantAvailable Then vString(3) = getPngQuantVersion() Else vString(3) = "none"
    
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
            
        'PNGQuant
        Case 3
            g_ImageFormats.pngQuantEnabled = Not g_ImageFormats.pngQuantEnabled
        
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
        
        'PNGQuant
        Case 3
            OpenURL "http://pngquant.org/"
            
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
            
        'PNGQuant
        Case 3
            OpenURL "https://raw.githubusercontent.com/pornel/pngquant/master/COPYRIGHT"
            
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
