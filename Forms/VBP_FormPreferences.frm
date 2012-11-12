VERSION 5.00
Begin VB.Form FormPreferences 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " PhotoDemon Preferences"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8805
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
   ScaleHeight     =   505
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   587
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset all Preferences"
      Height          =   495
      Left            =   360
      TabIndex        =   44
      Top             =   6840
      Width           =   2085
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   7200
      TabIndex        =   1
      Top             =   6840
      Width           =   1245
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   5880
      TabIndex        =   0
      Top             =   6840
      Width           =   1245
   End
   Begin PhotoDemon.jcbutton cmdCategory 
      Height          =   1140
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   2011
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Interface"
      ForeColor       =   4210752
      ForeColorHover  =   4194304
      Mode            =   2
      Value           =   -1  'True
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormPreferences.frx":0000
      PictureAlign    =   6
      PictureEffectOnDown=   0
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipType     =   1
      TooltipTitle    =   "Interface Preferences"
   End
   Begin PhotoDemon.jcbutton cmdCategory 
      Height          =   1140
      Index           =   3
      Left            =   5280
      TabIndex        =   5
      Top             =   240
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   2011
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Updates"
      ForeColor       =   4210752
      ForeColorHover  =   4194304
      Mode            =   2
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormPreferences.frx":1052
      PictureAlign    =   6
      PictureEffectOnDown=   0
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipType     =   1
      TooltipTitle    =   "Update Preferences"
   End
   Begin PhotoDemon.jcbutton cmdCategory 
      Height          =   1140
      Index           =   1
      Left            =   1920
      TabIndex        =   3
      Top             =   240
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   2011
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Tools"
      ForeColor       =   4210752
      ForeColorHover  =   4194304
      Mode            =   2
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormPreferences.frx":20A4
      PictureAlign    =   6
      PictureEffectOnDown=   0
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipType     =   1
      TooltipTitle    =   "Tool Preferences"
   End
   Begin PhotoDemon.jcbutton cmdCategory 
      Height          =   1140
      Index           =   4
      Left            =   6960
      TabIndex        =   6
      Top             =   240
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   2011
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Advanced"
      ForeColor       =   4210752
      ForeColorHover  =   4194304
      Mode            =   2
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormPreferences.frx":30F6
      PictureAlign    =   6
      PictureEffectOnDown=   0
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipType     =   1
      TooltipTitle    =   "Advanced Settings"
   End
   Begin PhotoDemon.jcbutton cmdCategory 
      Height          =   1140
      Index           =   2
      Left            =   3600
      TabIndex        =   4
      Top             =   240
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   2011
      ButtonStyle     =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Transparency"
      ForeColor       =   4210752
      ForeColorHover  =   4194304
      Mode            =   2
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormPreferences.frx":4148
      PictureAlign    =   6
      PictureEffectOnDown=   0
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipType     =   1
      TooltipTitle    =   "Transparency preferences"
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   4
      Left            =   240
      MousePointer    =   1  'Arrow
      ScaleHeight     =   313
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   23
      Top             =   1800
      Width           =   8295
      Begin VB.CheckBox chkGDIPlusTest 
         Appearance      =   0  'Flat
         Caption         =   "enable GDI+ support"
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
         Height          =   255
         Left            =   240
         TabIndex        =   43
         ToolTipText     =   $"VBP_FormPreferences.frx":519A
         Top             =   3360
         Width           =   3255
      End
      Begin VB.CheckBox chkFreeImageTest 
         Appearance      =   0  'Flat
         Caption         =   "enable FreeImage support"
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
         Height          =   255
         Left            =   240
         TabIndex        =   42
         ToolTipText     =   $"VBP_FormPreferences.frx":528C
         Top             =   2880
         Width           =   3255
      End
      Begin VB.TextBox TxtTempPath 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "automatically generated at run-time"
         ToolTipText     =   "Folder used for temporary files"
         Top             =   1560
         Width           =   5415
      End
      Begin VB.CommandButton CmdTmpPath 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         TabIndex        =   26
         ToolTipText     =   "Click to open a browse-for-folder dialog"
         Top             =   1560
         Width           =   405
      End
      Begin VB.CheckBox ChkLogMessages 
         Appearance      =   0  'Flat
         Caption         =   "log program messages to file (for debugging purposes)"
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
         Height          =   255
         Left            =   240
         TabIndex        =   25
         ToolTipText     =   $"VBP_FormPreferences.frx":537E
         Top             =   600
         Width           =   6975
      End
      Begin VB.Label lblAdvancedWarning 
         BackStyle       =   0  'Transparent
         Caption         =   $"VBP_FormPreferences.frx":5470
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   240
         TabIndex        =   45
         Top             =   3960
         Width           =   7815
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblRuntimeSettings 
         AutoSize        =   -1  'True
         Caption         =   "run-time testing options (NOTE: these are not saved to the INI file)"
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
         TabIndex        =   41
         Top             =   2280
         Width           =   7155
      End
      Begin VB.Label lblTempFolder 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "temporary file folder (used to hold Undo/Redo data):"
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
         Left            =   240
         TabIndex        =   28
         Top             =   1200
         Width           =   4530
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "advanced settings"
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
         TabIndex        =   24
         Top             =   0
         Width           =   1875
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4815
      Index           =   0
      Left            =   240
      MousePointer    =   1  'Arrow
      ScaleHeight     =   321
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   7
      Top             =   1800
      Width           =   8295
      Begin VB.CheckBox chkDropShadow 
         Appearance      =   0  'Flat
         Caption         =   " draw a shadow between the image and the canvas"
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
         Height          =   375
         Left            =   240
         TabIndex        =   40
         ToolTipText     =   " This setting helps images stand out from the canvas behind them"
         Top             =   810
         Width           =   5655
      End
      Begin VB.CheckBox chkFancyFonts 
         Appearance      =   0  'Flat
         Caption         =   " render PhotoDemon text with modern typefaces (only available on Vista or newer)"
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
         Height          =   375
         Left            =   240
         TabIndex        =   35
         ToolTipText     =   "This setting uses ""Segoe UI"" as the PhotoDemon interface font. Leaving it unchecked defaults to ""Tahoma""."
         Top             =   3150
         Width           =   7695
      End
      Begin VB.CheckBox chkConfirmUnsaved 
         Appearance      =   0  'Flat
         Caption         =   " when closing image files, warn me me about unsaved changes"
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
         Height          =   375
         Left            =   240
         TabIndex        =   22
         ToolTipText     =   "Check this if you want to be warned when you try to close an image with unsaved changes"
         Top             =   2190
         Width           =   7215
      End
      Begin VB.ComboBox cmbLargeImages 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1770
         Width           =   4815
      End
      Begin VB.ComboBox cmbCanvas 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   405
         Width           =   4815
      End
      Begin VB.PictureBox picCanvasColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   7680
         MouseIcon       =   "VBP_FormPreferences.frx":5502
         MousePointer    =   99  'Custom
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   37
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Click to change the image window background color"
         Top             =   405
         Width           =   585
      End
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
         Caption         =   "interface text"
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
         TabIndex        =   39
         Top             =   2760
         Width           =   1365
      End
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
         Caption         =   "load / save behavior"
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
         TabIndex        =   38
         Top             =   1380
         Width           =   2145
      End
      Begin VB.Label lblImgOpen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "when loading images, set zoom to: "
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
         Left            =   240
         TabIndex        =   13
         Top             =   1830
         Width           =   3075
      End
      Begin VB.Label lblCanvasFX 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "image canvas background:"
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
         Left            =   240
         TabIndex        =   12
         Top             =   465
         Width           =   2295
      End
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
         Caption         =   "canvas appearance"
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
         TabIndex        =   8
         Top             =   0
         Width           =   1980
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   3
      Left            =   240
      MousePointer    =   1  'Arrow
      ScaleHeight     =   313
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   17
      Top             =   1800
      Width           =   8295
      Begin VB.CheckBox ChkPromptPluginDownload 
         Appearance      =   0  'Flat
         Caption         =   "if core plugins cannot be located, offer to download them"
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
         Height          =   375
         Left            =   240
         TabIndex        =   20
         ToolTipText     =   $"VBP_FormPreferences.frx":5654
         Top             =   1080
         Width           =   6735
      End
      Begin VB.CheckBox chkProgramUpdates 
         Appearance      =   0  'Flat
         Caption         =   "automatically check for software updates every 10 days"
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
         Height          =   495
         Left            =   240
         TabIndex        =   19
         ToolTipText     =   "If this is disabled, you can visit tannerhelland.com/photodemon to manually download the latest version of PhotoDemon"
         Top             =   480
         Width           =   7455
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "update preferences"
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
         TabIndex        =   18
         Top             =   0
         Width           =   2010
      End
      Begin VB.Label lblExplanation 
         BackStyle       =   0  'Transparent
         Caption         =   "(disclaimer populated at run-time)"
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
         Height          =   2775
         Left            =   240
         TabIndex        =   21
         Top             =   1800
         Width           =   7935
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   2
      Left            =   240
      MousePointer    =   1  'Arrow
      ScaleHeight     =   313
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   29
      Top             =   1800
      Width           =   8295
      Begin VB.ComboBox cmbAlphaCheckSize 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   1875
         Width           =   5055
      End
      Begin VB.ComboBox cmbAlphaCheck 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   915
         Width           =   5055
      End
      Begin VB.PictureBox picAlphaOne 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   5520
         MouseIcon       =   "VBP_FormPreferences.frx":56F0
         MousePointer    =   99  'Custom
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   37
         TabIndex        =   32
         TabStop         =   0   'False
         ToolTipText     =   "Click to change the second checkerboard background color for alpha channels"
         Top             =   915
         Width           =   585
      End
      Begin VB.PictureBox picAlphaTwo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   6240
         MouseIcon       =   "VBP_FormPreferences.frx":5842
         MousePointer    =   99  'Custom
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   37
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "Click to change the second checkerboard background color for alpha channels"
         Top             =   915
         Width           =   585
      End
      Begin VB.Label lblAlphaCheckSize 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "transparency checkerboard size:"
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
         Left            =   240
         TabIndex        =   37
         Top             =   1560
         Width           =   2790
      End
      Begin VB.Label lblAlphaCheck 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "transparency checkerboard colors:"
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
         Left            =   240
         TabIndex        =   34
         Top             =   600
         Width           =   2970
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "transparency preferences"
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
         TabIndex        =   30
         Top             =   0
         Width           =   2640
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   1
      Left            =   240
      MousePointer    =   1  'Arrow
      ScaleHeight     =   313
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   14
      Top             =   1800
      Width           =   8295
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "tool preferences"
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
         TabIndex        =   16
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "There are not currently any tool settings."
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
         Left            =   240
         TabIndex        =   15
         Top             =   645
         Width           =   3510
      End
   End
   Begin VB.Line lneVertical 
      BorderColor     =   &H8000000D&
      X1              =   8
      X2              =   576
      Y1              =   106
      Y2              =   106
   End
End
Attribute VB_Name = "FormPreferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Program Preferences Handler
'Copyright ©2000-2012 by Tanner Helland
'Created: 8/November/02
'Last updated: 22/October/12
'Last update: revamped entire interface; settings are now sorted by category.
'
'Module for interfacing with the user's desired program preferences.  Handles
' reading from and copying to the program's ".INI" file.
'
'Note that this form interacts heavily with the INIProcessor module.
'
'***************************************************************************

Option Explicit

'Used to see if the user physically clicked a combo box, or if VB selected it on its own
Dim userInitiatedColorSelection As Boolean
Dim userInitiatedAlphaSelection As Boolean

'For this particular box, update the interface instantly
Private Sub chkFancyFonts_Click()

    useFancyFonts = CBool(chkFancyFonts)
    makeFormPretty Me
    makeFormPretty FormMain

End Sub

'Alpha channel checkerboard selection
Private Sub cmbAlphaCheck_Click()

    'Only respond to user-generated events
    If userInitiatedAlphaSelection = False Then Exit Sub

    'Redraw the sample picture boxes based on the value the user has selected
    AlphaCheckMode = cmbAlphaCheck.ListIndex
    Select Case cmbAlphaCheck.ListIndex
    
        'Case 0 - Highlights
        Case 0
            AlphaCheckOne = RGB(255, 255, 255)
            AlphaCheckTwo = RGB(204, 204, 204)
        
        'Case 1 - Midtones
        Case 1
            AlphaCheckOne = RGB(153, 153, 153)
            AlphaCheckTwo = RGB(102, 102, 102)
        
        'Case 2 - Shadows
        Case 2
            AlphaCheckOne = RGB(51, 51, 51)
            AlphaCheckTwo = RGB(0, 0, 0)
        
        'Case 3 - Custom
        Case 3
            AlphaCheckOne = RGB(255, 204, 246)
            AlphaCheckTwo = RGB(255, 255, 255)
        
    End Select

    'Change the picture boxes to match the current selection
    picAlphaOne.backColor = AlphaCheckOne
    picAlphaTwo.backColor = AlphaCheckTwo

End Sub

'Canvas background selection
Private Sub cmbCanvas_Click()
    
    'Only respond to user-generated events
    If userInitiatedColorSelection = False Then Exit Sub
    
    'Redraw the sample picture box value based on the value the user has selected
    Select Case cmbCanvas.ListIndex
        Case 0
            CanvasBackground = vb3DLight
        Case 1
            CanvasBackground = vb3DShadow
        Case 2
            'Prompt with a color selection box
            Dim retColor As Long
    
            Dim CD1 As cCommonDialog
            Set CD1 = New cCommonDialog
    
            retColor = picCanvasColor.backColor
    
            CD1.VBChooseColor retColor, True, True, False, Me.HWnd
    
            'If a color was selected, change the picture box and associated combo box to match
            If retColor >= 0 Then CanvasBackground = retColor Else CanvasBackground = picCanvasColor.backColor
            
    End Select
    
    DrawSampleCanvasBackground
    
End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    Unload Me
End Sub

'When the category is changed, only display the controls in that category
Private Sub cmdCategory_Click(Index As Integer)
    
    Static catID As Long
    For catID = 0 To cmdCategory.Count - 1
        If catID = Index Then picContainer(catID).Visible = True Else picContainer(catID).Visible = False
    Next catID
    
End Sub

'OK button
Private Sub CmdOK_Click()
    
    'Store whether the user wants to be prompted when closing unsaved images
    ConfirmClosingUnsaved = CBool(chkConfirmUnsaved.Value)
    userPreferences.SetPreference_Boolean "General Preferences", "ConfirmClosingUnsaved", ConfirmClosingUnsaved
    
    If ConfirmClosingUnsaved Then
        FormMain.cmdClose.ToolTip = "Close the current image." & vbCrLf & vbCrLf & "If the current image has not been saved, you will" & vbCrLf & " receive a prompt to save it before it closes."
    Else
        FormMain.cmdClose.ToolTip = "Close the current image." & vbCrLf & vbCrLf & "Because you have turned off save prompts (via Edit -> Preferences)," & vbCrLf & " you WILL NOT receive a prompt to save this image before it closes."
    End If
    
    'Store whether PhotoDemon is allowed to check for updates
    userPreferences.SetPreference_Boolean "General Preferences", "CheckForUpdates", CBool(chkProgramUpdates.Value)
    
    'Store whether PhotoDemon is allowed to offer the automatic download of missing core plugins
    userPreferences.SetPreference_Boolean "General Preferences", "PromptForPluginDownload", CBool(ChkPromptPluginDownload.Value)
    
    'Store whether we'll log system messages or not
    LogProgramMessages = CBool(ChkLogMessages.Value)
    userPreferences.SetPreference_Boolean "General Preferences", "LogProgramMessages", LogProgramMessages
    
    'Store the preference for rendering a drop shadow onto the canvas surrounding an image
    CanvasDropShadow = CBool(chkDropShadow.Value)
    userPreferences.SetPreference_Boolean "General Preferences", "CanvasDropShadow", CanvasDropShadow
    
    If CanvasDropShadow Then canvasShadow.initializeSquareShadow PD_CANVASSHADOWSIZE, PD_CANVASSHADOWSTRENGTH, CanvasBackground
    
    'Store the canvas background preference
    CanvasBackground = CLng(picCanvasColor.backColor)
    userPreferences.SetPreference_Long "General Preferences", "CanvasBackground", CanvasBackground
        
    'Store the alpha checkerboard preference
    userPreferences.SetPreference_Long "General Preferences", "AlphaCheckMode", CLng(cmbAlphaCheck.ListIndex)
    userPreferences.SetPreference_Long "General Preferences", "AlphaCheckOne", CLng(picAlphaOne.backColor)
    userPreferences.SetPreference_Long "General Preferences", "AlphaCheckTwo", CLng(picAlphaTwo.backColor)
    
    'Store the alpha checkerboard size preference
    AlphaCheckSize = cmbAlphaCheckSize.ListIndex
    userPreferences.SetPreference_Long "General Preferences", "AlphaCheckSize", AlphaCheckSize
    
    'Remember whether or not to autozoom large images
    AutosizeLargeImages = cmbLargeImages.ListIndex
    userPreferences.SetPreference_Long "General Preferences", "AutosizeLargeImages", AutosizeLargeImages
    
    'Verify the temporary path
    If LCase(TxtTempPath.Text) <> LCase(userPreferences.getTempPath) Then userPreferences.setTempPath TxtTempPath.Text
    
    'Remember the run-time only settings in the "Advanced" panel
    FreeImageEnabled = CBool(chkFreeImageTest.Value)
    GDIPlusEnabled = CBool(chkGDIPlusTest.Value)
    
    'Store the user's preference regarding interface fonts on modern versions of Windows
    userPreferences.SetPreference_Boolean "General Preferences", "UseFancyFonts", useFancyFonts
    
    'Because some settings affect the way image canvases are rendered, redraw every active canvas
    Dim tForm As Form
    Message "Saving preferences..."
    For Each tForm In VB.Forms
        If tForm.Name = "FormImage" Then PrepareViewport tForm
    Next
    Message "Finished."
        
    Unload Me
    
End Sub

'Regenerate the INI file from scratch.  This can be an effective way to "reset" a PhotoDemon installation.
Private Sub cmdReset_Click()

    'Before resetting, warn the user
    Dim confirmReset As VbMsgBoxResult
    confirmReset = MsgBox("This action will reset all preferences to their default values.  It cannot be undone." & vbCrLf & vbCrLf & "Are you sure you want to continue?", vbApplicationModal + vbExclamation + vbYesNo, "Reset all " & PROGRAMNAME & " preferences")

    'If the user gives final permission, rewrite the INI file from scratch and repopulate this form
    If confirmReset = vbYes Then
        userPreferences.resetPreferences
        LoadAllPreferences
    End If

End Sub

'When the "..." button is clicked, prompt the user with a "browse for folder" dialog
Private Sub CmdTmpPath_Click()
    Dim tString As String
    tString = BrowseForFolder(Me.HWnd)
    If tString <> "" Then TxtTempPath.Text = FixPath(tString)
End Sub

'Load all relevant values from the INI file, and populate their corresponding controls with the user's current settings
Private Sub LoadAllPreferences()
    
    'Start with the canvas background (which also requires populating the canvas background combo box)
    userInitiatedColorSelection = False
    cmbCanvas.Clear
    cmbCanvas.AddItem "System theme: light", 0
    cmbCanvas.AddItem "System theme: dark", 1
    cmbCanvas.AddItem "Custom color (click box to customize)", 2
        
    'Select the proper combo box value based on the CanvasBackground variable
    If CanvasBackground = vb3DLight Then
        'System theme: light
        cmbCanvas.ListIndex = 0
    ElseIf CanvasBackground = vb3DShadow Then
        'System theme: dark
        cmbCanvas.ListIndex = 1
    Else
        'Custom color
        cmbCanvas.ListIndex = 2
    End If
    
    'Draw the current canvas background to the sample picture box
    DrawSampleCanvasBackground
    userInitiatedColorSelection = True
    
    'Next, get the values for alpha-channel checkerboard rendering
    userInitiatedAlphaSelection = False
    cmbAlphaCheck.Clear
    cmbAlphaCheck.AddItem "Highlight checks", 0
    cmbAlphaCheck.AddItem "Midtone checks", 1
    cmbAlphaCheck.AddItem "Shadow checks", 2
    cmbAlphaCheck.AddItem "Custom (click boxes to customize)", 3
    
    cmbAlphaCheck.ListIndex = AlphaCheckMode
    
    picAlphaOne.backColor = AlphaCheckOne
    picAlphaTwo.backColor = AlphaCheckTwo
    
    userInitiatedAlphaSelection = True
    
    'Next, get the current alpha-channel checkerboard size value
    cmbAlphaCheckSize.Clear
    cmbAlphaCheckSize.AddItem "Small (4x4 pixels)", 0
    cmbAlphaCheckSize.AddItem "Medium (8x8 pixels)", 1
    cmbAlphaCheckSize.AddItem "Large (16x16 pixels)", 2
    
    cmbAlphaCheckSize.ListIndex = AlphaCheckSize
    
    'Assign the check box for logging program messages
    If LogProgramMessages Then ChkLogMessages.Value = vbChecked Else ChkLogMessages.Value = vbUnchecked
    
    'Assign the check box for prompting about unsaved images
    If ConfirmClosingUnsaved Then chkConfirmUnsaved.Value = vbChecked Else chkConfirmUnsaved.Value = vbUnchecked
    
    'Assign the check box for rendering a drop shadow around the image
    If CanvasDropShadow Then chkDropShadow.Value = vbChecked Else chkDropShadow.Value = vbUnchecked
    
    'Display the current temporary file path
    TxtTempPath.Text = userPreferences.getTempPath
    
    'We have to pull the "offer to download plugins" value from the INI file, since we don't track
    ' it internally (it's only accessed when PhotoDemon is first loaded)
    If userPreferences.GetPreference_Boolean("General Preferences", "PromptForPluginDownload", True) Then ChkPromptPluginDownload.Value = vbChecked Else ChkPromptPluginDownload.Value = vbUnchecked
    
    'Same for checking for software updates
    If userPreferences.GetPreference_Boolean("General Preferences", "CheckForUpdates", True) Then chkProgramUpdates.Value = vbChecked Else chkProgramUpdates.Value = vbUnchecked
    
    'Populate the "what to do when loading large images" combo box
    cmbLargeImages.Clear
    cmbLargeImages.AddItem "automatically fit the image on-screen", 0
    cmbLargeImages.AddItem "1:1 (100% zoom, or ""actual size"")", 1
    cmbLargeImages.ListIndex = userPreferences.GetPreference_Long("General Preferences", "AutosizeLargeImages", 0)
    
    'Hide the modern typefaces box if the user in on XP.  If the user is on Vista or later, set the box according
    ' to the preference stated in their INI file.
    If Not isVistaOrLater Then
        chkFancyFonts.Caption = " render PhotoDemon text with modern typefaces (only available on Vista or newer)"
        chkFancyFonts.Enabled = False
    Else
        chkFancyFonts.Caption = " render PhotoDemon text with modern typefaces"
        chkFancyFonts.Enabled = True
        If useFancyFonts Then chkFancyFonts.Value = vbChecked Else chkFancyFonts.Value = vbUnchecked
    End If
    
    'Populate and en/disable the run-time only settings in the "Advanced" panel
    If FreeImageEnabled Then chkFreeImageTest.Value = vbChecked Else chkFreeImageTest.Value = vbUnchecked
    If GDIPlusEnabled Then chkGDIPlusTest.Value = vbChecked Else chkGDIPlusTest.Value = vbUnchecked

End Sub

'When the form is loaded, populate the various checkboxes and textboxes with the values from the INI file
Private Sub Form_Load()
    
    Me.Caption = PROGRAMNAME & " Preferences"
    
    'Populate all controls with their corresponding values
    LoadAllPreferences
    
    'Populate the multi-line tooltips for the category command buttons
    'Interface
    cmdCategory(0).ToolTip = "Interface preferences include default setting for canvas backgrounds," & vbCrLf & "transparency checkerboards, and the program's visual theme."
    'Tools
    cmdCategory(1).ToolTip = "Tool preferences currently includes customizable options for the selection tool." & vbCrLf & "In the future, PhotoDemon will gain paint tools, and those settings will appear" & vbCrLf & "here as well."
    'Transparency
    cmdCategory(2).ToolTip = "Transparency preferences control how PhotoDemon displays images" & vbCrLf & "that contain alpha channels (e.g. 32bpp images)."
    'Updates
    cmdCategory(3).ToolTip = "Update preferences control how frequently PhotoDemon checks for" & vbCrLf & "updated versions, and how it handles the download of missing plugins."
    'Advanced
    cmdCategory(4).ToolTip = "Advanced preferences can be safely ignored by regular users." & vbCrLf & "Testers and developers may, however, find these settings useful."
    
    'Populate the network access disclaimer in the "Update" panel
    lblExplanation.Caption = PROGRAMNAME & " provides two non-essential features that require Internet access: checking for software updates, and offering to download core plugins (FreeImage, EZTwain, and ZLib) if they aren't present in the \Data\Plugins subdirectory." _
    & vbCrLf & vbCrLf & "The developers of " & PROGRAMNAME & " take privacy very seriously, so no information - statistical or otherwise - is uploaded by these features. Checking for software updates involves downloading a single ""updates.txt"" file containing the latest PhotoDemon version number. Similarly, downloading missing plugins simply involves downloading one or more of the compressed plugin files from the " & PROGRAMNAME & " server." _
    & vbCrLf & vbCrLf & "If you choose to disable these features, note that you can always visit tannerhelland.com/photodemon to manually download the most recent version of the program."
        
    'Finally, hide the inactive category panels
    Dim i As Long
    For i = 1 To picContainer.Count - 1
        picContainer(i).Visible = False
    Next i
    
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
    'Note: at present, this doesn't seem to be working, and I'm not sure why.  It has something to do with
    ' picture boxes contained within other picture boxes.  Because of this, I've manually set the mouse icon
    ' to an old-school hand cursor (which is all VB will accept).
    'setHandCursor picCanvasColor
    'setHandCursor picAlphaOne
    'setHandCursor picAlphaTwo
    
    'For some reason, the container picture boxes automatically acquire the pointer of children objects.
    ' Manually force those cursors to arrows to prevent this.
    For i = 0 To picContainer.Count - 1
        setArrowCursorToObject picContainer(i)
    Next i
        
End Sub

'Draw a sample of the current background to the PicCanvasColor picture box
Private Sub DrawSampleCanvasBackground()
    
    Me.picCanvasColor.backColor = CanvasBackground
    Me.picCanvasColor.Refresh
    Me.picCanvasColor.Enabled = True
    
End Sub

'Allow the user to change the first checkerboard color for alpha channels
Private Sub picAlphaOne_Click()
    
    Dim retColor As Long
    
    Dim CD1 As cCommonDialog
    Set CD1 = New cCommonDialog
    
    retColor = picAlphaOne.backColor
    
    'Display a Windows color selection box
    CD1.VBChooseColor retColor, True, True, False, Me.HWnd
    
    'If a color was selected, change the picture box and associated combo box to match
    If retColor > 0 Then
    
        AlphaCheckOne = retColor
        picAlphaOne.backColor = retColor
        
        userInitiatedAlphaSelection = False
        cmbAlphaCheck.ListIndex = 3   '3 corresponds to "custom colors"
        userInitiatedAlphaSelection = True
                
    End If
    
End Sub

'Allow the user to change the second checkerboard color for alpha channels
Private Sub picAlphaTwo_Click()
    
    Dim retColor As Long
    
    Dim CD1 As cCommonDialog
    Set CD1 = New cCommonDialog
    
    retColor = picAlphaTwo.backColor
    
    'Display a Windows color selection box
    CD1.VBChooseColor retColor, True, True, False, Me.HWnd
    
    'If a color was selected, change the picture box and associated combo box to match
    If retColor > 0 Then
    
        AlphaCheckTwo = retColor
        picAlphaTwo.backColor = retColor
        
        userInitiatedAlphaSelection = False
        cmbAlphaCheck.ListIndex = 3   '3 corresponds to "custom colors"
        userInitiatedAlphaSelection = True
                
    End If
    
End Sub

'Clicking the sample color box allows the user to pick a new color
Private Sub picCanvasColor_Click()
    
    Dim retColor As Long
    
    Dim CD1 As cCommonDialog
    Set CD1 = New cCommonDialog
    
    retColor = picCanvasColor.backColor
    
    'Display a Windows color selection box
    CD1.VBChooseColor retColor, True, True, False, Me.HWnd
    
    'If a color was selected, change the picture box and associated combo box to match
    If retColor >= 0 Then
    
        CanvasBackground = retColor
        
        userInitiatedColorSelection = False
        If CanvasBackground = vb3DLight Then
            'System theme: light
            cmbCanvas.ListIndex = 0
        ElseIf CanvasBackground = vb3DShadow Then
            'System theme: dark
            cmbCanvas.ListIndex = 1
        Else
            'Custom color
            cmbCanvas.ListIndex = 2
        End If
        userInitiatedColorSelection = True
        
        DrawSampleCanvasBackground
        
    End If
    
End Sub
