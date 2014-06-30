VERSION 5.00
Begin VB.Form FormPreferences 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " PhotoDemon Options"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11505
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
   MousePointer    =   99  'Custom
   ScaleHeight     =   508
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   767
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset all options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2910
      TabIndex        =   85
      Top             =   6990
      Width           =   2580
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
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
      Height          =   495
      Left            =   8520
      TabIndex        =   0
      Top             =   6990
      Width           =   1365
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9990
      TabIndex        =   1
      Top             =   6990
      Width           =   1365
   End
   Begin PhotoDemon.jcbutton cmdCategory 
      Height          =   780
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   1376
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
      BackColor       =   -2147483643
      Caption         =   "Interface"
      ForeColor       =   4210752
      ForeColorHover  =   4194304
      Mode            =   1
      Value           =   -1  'True
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormPreferences.frx":0000
      PictureAlign    =   0
      PictureEffectOnDown=   0
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipTitle    =   "Interface Options"
      ColorScheme     =   3
   End
   Begin PhotoDemon.jcbutton cmdCategory 
      Height          =   780
      Index           =   6
      Left            =   120
      TabIndex        =   5
      Top             =   5160
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   1376
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
      BackColor       =   -2147483643
      Caption         =   "Updates"
      ForeColor       =   4210752
      ForeColorHover  =   4194304
      Mode            =   1
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormPreferences.frx":1052
      PictureAlign    =   0
      PictureEffectOnDown=   0
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipTitle    =   "Update Options"
      ColorScheme     =   3
   End
   Begin PhotoDemon.jcbutton cmdCategory 
      Height          =   780
      Index           =   4
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   1376
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
      BackColor       =   -2147483643
      Caption         =   "Tools"
      ForeColor       =   4210752
      ForeColorHover  =   4194304
      Mode            =   1
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormPreferences.frx":24A4
      PictureAlign    =   0
      PictureEffectOnDown=   0
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipTitle    =   "Tool Options"
      ColorScheme     =   3
   End
   Begin PhotoDemon.jcbutton cmdCategory 
      Height          =   780
      Index           =   7
      Left            =   120
      TabIndex        =   6
      Top             =   6000
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   1376
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
      BackColor       =   -2147483643
      Caption         =   "Advanced"
      ForeColor       =   4210752
      ForeColorHover  =   4194304
      Mode            =   1
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormPreferences.frx":38F6
      PictureAlign    =   0
      PictureEffectOnDown=   0
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipTitle    =   "Advanced Options"
      ColorScheme     =   3
   End
   Begin PhotoDemon.jcbutton cmdCategory 
      Height          =   780
      Index           =   5
      Left            =   120
      TabIndex        =   4
      Top             =   4320
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   1376
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
      BackColor       =   -2147483643
      Caption         =   "Color and Transparency"
      ForeColor       =   4210752
      ForeColorHover  =   4194304
      Mode            =   1
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormPreferences.frx":4D48
      PictureAlign    =   0
      PictureEffectOnDown=   0
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipTitle    =   "Color and Transparency Options"
      ColorScheme     =   3
   End
   Begin PhotoDemon.jcbutton cmdCategory 
      Height          =   780
      Index           =   1
      Left            =   120
      TabIndex        =   28
      Top             =   960
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   1376
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
      BackColor       =   -2147483643
      Caption         =   "Loading"
      ForeColor       =   4210752
      ForeColorHover  =   4194304
      Mode            =   1
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormPreferences.frx":619A
      PictureAlign    =   0
      PictureEffectOnDown=   0
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipTitle    =   "Load (Import) Options"
      ColorScheme     =   3
   End
   Begin PhotoDemon.jcbutton cmdCategory 
      Height          =   780
      Index           =   2
      Left            =   120
      TabIndex        =   44
      Top             =   1800
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   1376
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
      BackColor       =   -2147483643
      Caption         =   "Saving"
      ForeColor       =   4210752
      ForeColorHover  =   4194304
      Mode            =   1
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormPreferences.frx":71EC
      PictureAlign    =   0
      PictureEffectOnDown=   0
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipTitle    =   "Save (Export) Options"
      ColorScheme     =   3
   End
   Begin PhotoDemon.jcbutton cmdCategory 
      Height          =   780
      Index           =   3
      Left            =   120
      TabIndex        =   46
      Top             =   2640
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   1376
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
      BackColor       =   -2147483643
      Caption         =   "File formats"
      ForeColor       =   4210752
      ForeColorHover  =   4194304
      Mode            =   1
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormPreferences.frx":863E
      PictureAlign    =   0
      PictureEffectOnDown=   0
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipTitle    =   "File Format Options"
      ColorScheme     =   3
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6480
      Index           =   4
      Left            =   3000
      MousePointer    =   1  'Arrow
      ScaleHeight     =   432
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   11
      Top             =   240
      Width           =   8295
      Begin PhotoDemon.smartCheckBox chkSelectionClearCrop 
         Height          =   480
         Left            =   240
         TabIndex        =   96
         Top             =   480
         Width           =   6480
         _ExtentX        =   11430
         _ExtentY        =   847
         Caption         =   "automatically clear the active selection after ""Crop to Selection"" is used"
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
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "selections"
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
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   1020
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6480
      Index           =   7
      Left            =   3000
      MousePointer    =   1  'Arrow
      ScaleHeight     =   432
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   16
      Top             =   240
      Width           =   8295
      Begin PhotoDemon.jcbutton cmdCopyReportClipboard 
         Height          =   525
         Left            =   7680
         TabIndex        =   121
         Top             =   4170
         Width           =   525
         _ExtentX        =   873
         _ExtentY        =   873
         ButtonStyle     =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   0
         Caption         =   ""
         PictureNormal   =   "VBP_FormPreferences.frx":9690
         PictureEffectOnOver=   0
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         TooltipBackColor=   0
      End
      Begin VB.TextBox txtHardware 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   120
         Top             =   2520
         Width           =   7335
      End
      Begin PhotoDemon.smartCheckBox chkLogMessages 
         Height          =   480
         Left            =   240
         TabIndex        =   97
         Top             =   360
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   847
         Caption         =   "log all program messages to file "
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
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "automatically generated at run-time"
         ToolTipText     =   "Folder used for temporary files"
         Top             =   1440
         Width           =   7335
      End
      Begin VB.CommandButton cmdTmpPath 
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
         Height          =   450
         Left            =   7680
         TabIndex        =   18
         ToolTipText     =   "Click to open a browse-for-folder dialog"
         Top             =   1395
         Width           =   525
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "hardware acceleration report:"
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
         Left            =   0
         TabIndex        =   119
         Top             =   2040
         Width           =   3120
      End
      Begin VB.Label lblMemoryUsageMax 
         BackStyle       =   0  'Transparent
         Caption         =   "memory usage will be displayed here"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00804040&
         Height          =   480
         Left            =   240
         TabIndex        =   62
         Top             =   5790
         Width           =   7965
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblMemoryUsageCurrent 
         BackStyle       =   0  'Transparent
         Caption         =   "memory usage will be displayed here"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00804040&
         Height          =   480
         Left            =   240
         TabIndex        =   61
         Top             =   5280
         Width           =   7965
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "memory diagnostics"
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
         Left            =   0
         TabIndex        =   60
         Top             =   4920
         Width           =   2130
      End
      Begin VB.Label lblRuntimeSettings 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "temporary file location"
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
         Left            =   0
         TabIndex        =   43
         Top             =   960
         Width           =   2385
      End
      Begin VB.Label lblTempPathWarning 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   240
         TabIndex        =   27
         Top             =   2040
         Width           =   7695
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "debugging"
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
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   1125
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6480
      Index           =   6
      Left            =   3000
      MousePointer    =   1  'Arrow
      ScaleHeight     =   432
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   13
      Top             =   240
      Width           =   8295
      Begin PhotoDemon.smartCheckBox chkPromptPluginDownload 
         Height          =   480
         Left            =   240
         TabIndex        =   87
         Top             =   1080
         Width           =   5280
         _ExtentX        =   9313
         _ExtentY        =   847
         Caption         =   "if core plugins cannot be located, offer to download them"
         Value           =   1
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
      Begin PhotoDemon.smartCheckBox chkProgramUpdates 
         Height          =   480
         Left            =   240
         TabIndex        =   86
         ToolTipText     =   "If this is disabled, you can visit photodemon.org to manually download the latest version of PhotoDemon"
         Top             =   480
         Width           =   5130
         _ExtentX        =   9049
         _ExtentY        =   847
         Caption         =   "automatically check for software updates every 10 days"
         Value           =   1
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
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "update options"
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
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   1575
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
         Height          =   4095
         Left            =   240
         TabIndex        =   15
         Top             =   1800
         Width           =   7935
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6480
      Index           =   0
      Left            =   3000
      MousePointer    =   1  'Arrow
      ScaleHeight     =   432
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   7
      Top             =   240
      Width           =   8295
      Begin PhotoDemon.textUpDown tudRecentFiles 
         Height          =   420
         Left            =   3900
         TabIndex        =   109
         Top             =   4605
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Min             =   1
         Max             =   32
         Value           =   10
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
      Begin PhotoDemon.colorSelector csCanvasColor 
         Height          =   435
         Left            =   6960
         TabIndex        =   104
         Top             =   780
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   767
      End
      Begin PhotoDemon.smartCheckBox chkFancyFonts 
         Height          =   480
         Left            =   240
         TabIndex        =   89
         ToolTipText     =   "This setting uses ""Segoe UI"" as the PhotoDemon interface font. Leaving it unchecked defaults to ""Tahoma""."
         Top             =   5520
         Width           =   7425
         _ExtentX        =   13097
         _ExtentY        =   847
         Caption         =   "render PhotoDemon text with modern typefaces (only available on Vista or newer)"
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
      Begin PhotoDemon.smartCheckBox chkDropShadow 
         Height          =   480
         Left            =   240
         TabIndex        =   88
         ToolTipText     =   "This setting helps images stand out from the canvas behind them"
         Top             =   1230
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   847
         Caption         =   "draw drop shadow between image and canvas"
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
      Begin VB.ComboBox cmbMRUCaption 
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   4080
         Width           =   8055
      End
      Begin VB.ComboBox cmbImageCaption 
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   32
         ToolTipText     =   "Image windows tend to be large, so feel free to display each image's full location in the image window title bars."
         Top             =   2610
         Width           =   8055
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   810
         Width           =   6615
      End
      Begin VB.Label lblRecentFileCount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "maximum number of recent file entries: "
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
         TabIndex        =   108
         Top             =   4680
         Width           =   3480
      End
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "recent files list"
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
         Index           =   13
         Left            =   0
         TabIndex        =   107
         Top             =   3240
         Width           =   1515
      End
      Begin VB.Label lblMRUText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "recently used file shortcuts should be: "
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
         TabIndex        =   36
         Top             =   3720
         Width           =   3315
      End
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "captions"
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
         Left            =   0
         TabIndex        =   34
         Top             =   1800
         Width           =   870
      End
      Begin VB.Label lblImageText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "image window titles should be: "
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
         TabIndex        =   33
         Top             =   2250
         Width           =   2730
      End
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "miscellaneous"
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
         Left            =   0
         TabIndex        =   26
         Top             =   5160
         Width           =   1470
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
         TabIndex        =   10
         Top             =   450
         Width           =   2295
      End
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   1980
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6480
      Index           =   5
      Left            =   3000
      MousePointer    =   1  'Arrow
      ScaleHeight     =   432
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   20
      Top             =   240
      Width           =   8295
      Begin VB.CommandButton cmdColorProfilePath 
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
         Left            =   7380
         TabIndex        =   118
         Top             =   2760
         Width           =   810
      End
      Begin VB.TextBox txtColorProfilePath 
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
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   117
         Text            =   "(none)"
         Top             =   2760
         Width           =   6525
      End
      Begin VB.ComboBox cmbMonitors 
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
         Left            =   780
         Style           =   2  'Dropdown List
         TabIndex        =   115
         Top             =   1950
         Width           =   7440
      End
      Begin PhotoDemon.smartOptionButton optColorManagement 
         Height          =   330
         Index           =   0
         Left            =   240
         TabIndex        =   112
         Top             =   840
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   582
         Caption         =   "use the system color profile"
         Value           =   -1  'True
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
      Begin PhotoDemon.colorSelector csAlphaOne 
         Height          =   435
         Left            =   6240
         TabIndex        =   105
         Top             =   4230
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   767
      End
      Begin PhotoDemon.smartCheckBox chkValidateAlpha 
         Height          =   480
         Left            =   240
         TabIndex        =   90
         Top             =   5760
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   847
         Caption         =   "automatically validate all incoming alpha channels"
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   5250
         Width           =   5895
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   4260
         Width           =   5895
      End
      Begin PhotoDemon.colorSelector csAlphaTwo 
         Height          =   435
         Left            =   7320
         TabIndex        =   106
         Top             =   4230
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   767
      End
      Begin PhotoDemon.smartOptionButton optColorManagement 
         Height          =   330
         Index           =   1
         Left            =   240
         TabIndex        =   113
         Top             =   1200
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   582
         Caption         =   "use one or more custom color profiles"
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
      Begin VB.Label lblColorManagement 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "color profile for selected monitor:"
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
         Index           =   2
         Left            =   780
         TabIndex        =   116
         Top             =   2430
         Width           =   2880
      End
      Begin VB.Label lblColorManagement 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "available monitors:"
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
         Left            =   780
         TabIndex        =   114
         Top             =   1590
         Width           =   1635
      End
      Begin VB.Label lblColorManagement 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "when rendering images to the screen:"
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
         Left            =   240
         TabIndex        =   111
         Top             =   480
         Width           =   3285
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "color management"
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
         Left            =   0
         TabIndex        =   110
         Top             =   0
         Width           =   1980
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
         TabIndex        =   25
         Top             =   4860
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
         TabIndex        =   23
         Top             =   3870
         Width           =   2970
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "transparency management"
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
         Left            =   0
         TabIndex        =   21
         Top             =   3420
         Width           =   2805
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6495
      Index           =   3
      Left            =   3000
      MousePointer    =   1  'Arrow
      ScaleHeight     =   433
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   47
      Top             =   240
      Width           =   8295
      Begin VB.ComboBox cmbFiletype 
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
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   72
         Top             =   960
         Width           =   7395
      End
      Begin VB.PictureBox picFileContainer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3735
         Index           =   1
         Left            =   240
         ScaleHeight     =   249
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   529
         TabIndex        =   76
         Top             =   1680
         Width           =   7935
         Begin PhotoDemon.smartCheckBox chkPNGBackground 
            Height          =   480
            Left            =   360
            TabIndex        =   93
            Top             =   2520
            Width           =   4830
            _ExtentX        =   8520
            _ExtentY        =   847
            Caption         =   "preserve file's original background color, if available"
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
         Begin PhotoDemon.smartCheckBox chkPNGInterlacing 
            Height          =   480
            Left            =   360
            TabIndex        =   92
            Top             =   2040
            Width           =   2430
            _ExtentX        =   4286
            _ExtentY        =   847
            Caption         =   "use interlacing (Adam7)"
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
         Begin VB.HScrollBar hsPNGCompression 
            Height          =   330
            Left            =   360
            Max             =   9
            TabIndex        =   78
            Top             =   1080
            Value           =   9
            Width           =   7095
         End
         Begin VB.Label lblPNGCompression 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "maximum compression"
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
            Left            =   5625
            TabIndex        =   81
            Top             =   1560
            Width           =   1590
         End
         Begin VB.Label lblPNGCompression 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "no compression"
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
            Left            =   600
            TabIndex        =   80
            Top             =   1560
            Width           =   1110
         End
         Begin VB.Label lblFileStuff 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "when saving, compress PNG files at the following level:"
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
            Left            =   360
            TabIndex        =   79
            Top             =   720
            Width           =   4725
         End
         Begin VB.Label lblInterfaceTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PNG (Portable Network Graphic) options"
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
            Index           =   20
            Left            =   120
            TabIndex        =   77
            Top             =   120
            Width           =   4290
         End
      End
      Begin VB.PictureBox picFileContainer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3735
         Index           =   3
         Left            =   240
         ScaleHeight     =   249
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   529
         TabIndex        =   82
         Top             =   1680
         Width           =   7935
         Begin PhotoDemon.smartCheckBox chkTGARLE 
            Height          =   480
            Left            =   360
            TabIndex        =   94
            Top             =   600
            Width           =   4410
            _ExtentX        =   7779
            _ExtentY        =   847
            Caption         =   "use RLE compression when saving TGA images"
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
         Begin VB.Label lblInterfaceTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TGA (Truevision) options"
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
            Index           =   21
            Left            =   120
            TabIndex        =   83
            Top             =   120
            Width           =   2700
         End
      End
      Begin VB.PictureBox picFileContainer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3735
         Index           =   0
         Left            =   240
         ScaleHeight     =   249
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   529
         TabIndex        =   74
         Top             =   1680
         Width           =   7935
         Begin PhotoDemon.smartCheckBox chkBMPRLE 
            Height          =   480
            Left            =   360
            TabIndex        =   95
            Top             =   600
            Width           =   4890
            _ExtentX        =   8625
            _ExtentY        =   847
            Caption         =   "use RLE compression when saving 8bpp BMP images"
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
         Begin VB.Label lblInterfaceTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BMP (Bitmap) options"
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
            Index           =   19
            Left            =   120
            TabIndex        =   75
            Top             =   120
            Width           =   2295
         End
      End
      Begin VB.PictureBox picFileContainer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3735
         Index           =   4
         Left            =   240
         ScaleHeight     =   249
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   529
         TabIndex        =   66
         Top             =   1680
         Width           =   7935
         Begin PhotoDemon.smartCheckBox chkTIFFCMYK 
            Height          =   480
            Left            =   360
            TabIndex        =   91
            Top             =   1560
            Width           =   4230
            _ExtentX        =   7461
            _ExtentY        =   847
            Caption         =   " save TIFFs as separated CMYK (for printing)"
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
         Begin VB.ComboBox cmbTIFFCompression 
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
            TabIndex        =   67
            Top             =   960
            Width           =   7335
         End
         Begin VB.Label lblInterfaceTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TIFF (Tagged Image File Format) options"
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
            Index           =   7
            Left            =   120
            TabIndex        =   70
            Top             =   120
            Width           =   4395
         End
         Begin VB.Label lblFileStuff 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "when saving, compress TIFFs using:"
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
            Left            =   360
            TabIndex        =   68
            Top             =   645
            Width           =   3135
         End
      End
      Begin VB.PictureBox picFileContainer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3855
         Index           =   2
         Left            =   240
         ScaleHeight     =   257
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   529
         TabIndex        =   63
         Top             =   1680
         Width           =   7935
         Begin VB.ComboBox cmbPPMFormat 
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
            Left            =   480
            Style           =   2  'Dropdown List
            TabIndex        =   64
            Top             =   960
            Width           =   7335
         End
         Begin VB.Label lblInterfaceTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PPM (Portable Pixmap) options"
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
            Index           =   12
            Left            =   120
            TabIndex        =   71
            Top             =   120
            Width           =   3285
         End
         Begin VB.Label lblPPMEncoding 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "export PPM files using:"
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
            TabIndex        =   65
            Top             =   600
            Width           =   1950
         End
      End
      Begin VB.Label lblFileFreeImageWarning 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   600
         TabIndex        =   73
         Top             =   5520
         Width           =   7455
      End
      Begin VB.Line lineFiletype 
         BorderColor     =   &H8000000D&
         X1              =   536
         X2              =   16
         Y1              =   103
         Y2              =   103
      End
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "please select a file type:"
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
         Index           =   18
         Left            =   360
         TabIndex        =   69
         Top             =   480
         Width           =   2520
      End
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "file format options"
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
         Index           =   9
         Left            =   0
         TabIndex        =   48
         Top             =   0
         Width           =   1950
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6480
      Index           =   2
      Left            =   3000
      MousePointer    =   1  'Arrow
      ScaleHeight     =   432
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   45
      Top             =   240
      Width           =   8295
      Begin PhotoDemon.smartCheckBox chkConfirmUnsaved 
         Height          =   480
         Left            =   240
         TabIndex        =   103
         ToolTipText     =   "Check this if you want to be warned when you try to close an image with unsaved changes"
         Top             =   360
         Width           =   5445
         _ExtentX        =   9604
         _ExtentY        =   847
         Caption         =   "when closing images, warn me me about unsaved changes"
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
      Begin VB.ComboBox cmbSaveBehavior 
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   99
         Top             =   5925
         Width           =   7980
      End
      Begin VB.ComboBox cmbExportColorDepth 
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   1740
         Width           =   7980
      End
      Begin VB.ComboBox cmbMetadata 
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   4530
         Width           =   7980
      End
      Begin VB.ComboBox cmbDefaultSaveFormat 
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   3135
         Width           =   7980
      End
      Begin VB.Label lblSubheader 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "when saving images that originally contained metadata:"
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
         Index           =   3
         Left            =   240
         TabIndex        =   100
         Top             =   4140
         Width           =   4785
      End
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "metadata (EXIF, GPS, comments, etc.)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00505050&
         Height          =   285
         Index           =   1
         Left            =   0
         TabIndex        =   98
         Top             =   3690
         Width           =   4065
      End
      Begin VB.Label lblSubheader 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "set outgoing color depth:"
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
         Left            =   240
         TabIndex        =   59
         Top             =   1350
         Width           =   2145
      End
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "color depth of saved images"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00505050&
         Height          =   285
         Index           =   17
         Left            =   0
         TabIndex        =   58
         Top             =   930
         Width           =   2985
      End
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "save behavior: overwrite vs make a copy"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00505050&
         Height          =   285
         Index           =   16
         Left            =   0
         TabIndex        =   56
         Top             =   5085
         Width           =   4320
      End
      Begin VB.Label lblSubheader 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "when ""Save"" is used:"
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
         Index           =   2
         Left            =   240
         TabIndex        =   55
         Top             =   5535
         Width           =   1830
      End
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "closing unsaved images"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00505050&
         Height          =   285
         Index           =   11
         Left            =   0
         TabIndex        =   52
         Top             =   0
         Width           =   2505
      End
      Begin VB.Label lblSubheader 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "when using the ""Save As"" command, set the default file format according to:"
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
         Left            =   240
         TabIndex        =   51
         Top             =   2730
         Width           =   6585
      End
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "default file format when saving"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00505050&
         Height          =   285
         Index           =   10
         Left            =   0
         TabIndex        =   50
         Top             =   2310
         Width           =   3285
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6465
      Index           =   1
      Left            =   3000
      MousePointer    =   1  'Arrow
      ScaleHeight     =   431
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   29
      Top             =   240
      Width           =   8295
      Begin VB.ComboBox cmbMultiImage 
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   3090
         Width           =   7920
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   4530
         Width           =   7920
      End
      Begin PhotoDemon.smartCheckBox chkInitialColorDepth 
         Height          =   480
         Left            =   240
         TabIndex        =   101
         Top             =   360
         Width           =   6765
         _ExtentX        =   11933
         _ExtentY        =   847
         Caption         =   "count unique colors in incoming images (to determine optimal color depth)"
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
      Begin PhotoDemon.smartCheckBox chkToneMapping 
         Height          =   480
         Left            =   240
         TabIndex        =   102
         Top             =   1440
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   847
         Caption         =   "automatically apply tone mapping to HDR and RAW images (48, 64, 96, 128bpp)"
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
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "color depth"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00505050&
         Height          =   285
         Index           =   15
         Left            =   60
         TabIndex        =   53
         Top             =   0
         Width           =   1200
      End
      Begin VB.Label lblFreeImageWarning 
         BackStyle       =   0  'Transparent
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
         Height          =   375
         Left            =   120
         TabIndex        =   42
         Top             =   6000
         Width           =   8055
      End
      Begin VB.Label lblMultiImages 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "if an image contains multiple pages:"
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
         TabIndex        =   41
         Top             =   2700
         Width           =   3105
      End
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "multi-page images (animated GIF, icons, TIFF)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00505050&
         Height          =   285
         Index           =   8
         Left            =   60
         TabIndex        =   39
         Top             =   2220
         Width           =   4965
      End
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "high-dynamic range (HDR) images"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00505050&
         Height          =   285
         Index           =   6
         Left            =   60
         TabIndex        =   38
         Top             =   1080
         Width           =   3675
      End
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "zoom"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00505050&
         Height          =   285
         Index           =   5
         Left            =   60
         TabIndex        =   37
         Top             =   3720
         Width           =   585
      End
      Begin VB.Label lblImgOpen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "when an image is first loaded, set its viewport zoom to: "
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
         TabIndex        =   31
         Top             =   4155
         Width           =   4845
      End
   End
   Begin VB.Line lneVertical 
      BorderColor     =   &H8000000D&
      X1              =   184
      X2              =   184
      Y1              =   8
      Y2              =   448
   End
   Begin VB.Label lblBackground 
      Height          =   855
      Left            =   0
      TabIndex        =   84
      Top             =   6840
      Width           =   12135
   End
End
Attribute VB_Name = "FormPreferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Program Preferences Handler
'Copyright ©2002-2014 by Tanner Helland
'Created: 8/November/02
'Last updated: 04/September/13
'Last update: huge code overhaul.  This dialog lacked any sort of code organization, which made it extremely difficult
'             to manage.  I have now reorganized everything (including the preferences XML file itself) by category,
'             including all load/save preference functions.  This should make it much easier to modify this dialog in
'             the future.
'
'Dialog for interfacing with the user's desired program preferences.  Handles reading/writing from/to the persistent
' XML file that actually stores all preferences.
'
'Note that this form interacts heavily with the pdPreferences class.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Used to see if the user physically clicked a combo box, or if VB selected it on its own
Private userInitiatedColorSelection As Boolean
Private userInitiatedAlphaSelection As Boolean

'Some settings are odd - I want them to update in real-time, so the user can see the effects of the change.  But if the user presses
' "cancel", the original settings need to be returned.  Thus, remember these settings, and restore them upon canceling.
Dim originalg_useFancyFonts As Boolean
Dim originalg_CanvasBackground As Long

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'For the font check box, update the interface instantly (so the user can see the option's effects)
Private Sub chkFancyFonts_Click()

    If Me.Visible Then
        g_UseFancyFonts = CBool(chkFancyFonts)
        
        'We must reassign the proper font manually (normally this is only done when the program is loaded)
        If g_UseFancyFonts Then
            g_InterfaceFont = "Segoe UI"
        Else
            g_InterfaceFont = "Tahoma"
        End If
    
        makeFormPretty Me
        FormMain.requestMakeFormPretty
    End If

End Sub

'Alpha channel checkerboard selection; change the color selectors to match
Private Sub cmbAlphaCheck_Click()

    'Only respond to user-generated events
    If userInitiatedAlphaSelection Then

        userInitiatedAlphaSelection = False

        'Redraw the sample picture boxes based on the value the user has selected
        Select Case cmbAlphaCheck.ListIndex
        
            'Case 0 - Highlights
            Case 0
                csAlphaOne.Color = RGB(255, 255, 255)
                csAlphaTwo.Color = RGB(204, 204, 204)
            
            'Case 1 - Midtones
            Case 1
                csAlphaOne.Color = RGB(153, 153, 153)
                csAlphaTwo.Color = RGB(102, 102, 102)
            
            'Case 2 - Shadows
            Case 2
                csAlphaOne.Color = RGB(51, 51, 51)
                csAlphaTwo.Color = RGB(0, 0, 0)
            
            'Case 3 - Custom
            Case 3
                csAlphaOne.Color = RGB(255, 204, 246)
                csAlphaTwo.Color = RGB(255, 255, 255)
            
        End Select
        
        userInitiatedAlphaSelection = True
                
    End If

End Sub

'Canvas background selection; change the color selection box to match
Private Sub cmbCanvas_Click()
    
    If userInitiatedColorSelection Then
    
        'Redraw the sample color box based on the value the user has selected
        Select Case cmbCanvas.ListIndex
            
            Case 0
                userInitiatedColorSelection = False
                g_CanvasBackground = vb3DLight
                csCanvasColor.Color = ConvertSystemColor(g_CanvasBackground)
                userInitiatedColorSelection = True
            
            Case 1
                userInitiatedColorSelection = False
                g_CanvasBackground = vb3DShadow
                csCanvasColor.Color = ConvertSystemColor(g_CanvasBackground)
                userInitiatedColorSelection = True
            
            'Prompt with a color selection box
            Case 2
                csCanvasColor.displayColorSelection
                
        End Select
    
    End If
    
End Sub

'When a new filetype is selected on the File Formats settings page, display the matching options panel
Private Sub cmbFiletype_Click()
    
    Dim ftID As Long
    For ftID = 0 To cmbFiletype.ListCount - 1
        If ftID = cmbFiletype.ListIndex Then picFileContainer(ftID).Visible = True Else picFileContainer(ftID).Visible = False
    Next ftID
    
End Sub

'Whenever the Color and Transparency -> Color Management -> Monitor combo box is changed, load the relevant color profile
' path from the preferences file (if one exists)
Private Sub cmbMonitors_Click()

    'One of the difficulties with tracking multiple monitors is that the user can attach/detach them at will.  They
    ' can also have multiple monitors attached with the same make and model (and retrieving EDIDs is extremely
    ' unpleasant and overwrought).  Per this article (http://www.microsoft.com/msj/0697/monitor/monitor.aspx), PD
    ' uses the HMONITOR handle to store and retrieve monitor-specific settings, because as the article states,
    ' "A physical device has the same HMONITOR value throughout its lifetime, even across changes to display settings,
    ' as long as it remains a part of the desktop."  If the user runs PD without a monitor attached, only to reattach
    ' it later, I have no idea if a new HMONITOR will be assigned or not... and at present, that's frankly not a
    ' huge concern for me.   We'll leave that level of management up to the OS.  For now, just assume that HMONITOR
    ' is a valid way to persistently track individual monitors.

    'Start by retrieving the HMONITOR value for the selected monitor
    Dim hMonitor As Long
    hMonitor = g_cMonitors.Monitors(cmbMonitors.ListIndex + 1).Handle
    
    'Use that to retrieve a stored color profile (if any)
    Dim profilePath As String
    profilePath = g_UserPreferences.GetPref_String("Transparency", "MonitorProfile_" & hMonitor, "(none)")
    
    'If the returned value is "(none)", translate that into the user's language before displaying; otherwise, display
    ' whatever path we retrieved.
    If profilePath = "(none)" Then
        txtColorProfilePath.Text = g_Language.TranslateMessage("(none)")
    Else
        txtColorProfilePath.Text = profilePath
    End If
    
End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    
    'Restore any settings that may have been changed in real-time
    If g_UseFancyFonts <> originalg_useFancyFonts Then
        g_UseFancyFonts = originalg_useFancyFonts
        FormMain.requestMakeFormPretty
    End If
    
    g_CanvasBackground = originalg_CanvasBackground
    
    Unload Me
    
End Sub

'When the preferences category is changed, only display the controls in that category
Private Sub cmdCategory_Click(Index As Integer)
    
    Dim catID As Long
    For catID = 0 To cmdCategory.Count - 1
        If catID = Index Then
            picContainer(catID).Visible = True
            cmdCategory(catID).Value = True
        Else
            picContainer(catID).Visible = False
            cmdCategory(catID).Value = False
        End If
    Next catID
    
End Sub

'Allow the user to select a new color profile for the attached monitor.  Because this text box is re-used for multiple
' settings, save any changes to file immediately, rather than waiting for the user to click OK.
Private Sub cmdColorProfilePath_Click()

    'Disable user input until the dialog closes
    Interface.disableUserInput
    
    Dim sFile As String
    sFile = ""
    
    'Get the last color profile path from the preferences file
    Dim tempPathString As String
    tempPathString = g_UserPreferences.GetPref_String("Paths", "Color Profile", "")
    
    'If no color profile path was found, populate it with the default system color profile path
    If Len(tempPathString) = 0 Then tempPathString = getSystemColorFolder()
    
    'Prepare a common dialog filter list with extensions of known profile types
    Dim cdFilter As String
    cdFilter = g_Language.TranslateMessage("ICC Profiles") & " (.icc, .icm)|*.icc;*.icm"
    cdFilter = cdFilter & "|" & g_Language.TranslateMessage("All files") & "|*.*"
    
    Dim cdTitle As String
    cdTitle = g_Language.TranslateMessage("Please select a color profile")
    
    Dim CC As cCommonDialog
    Set CC = New cCommonDialog
    
    If CC.VBGetOpenFileName(sFile, , True, False, False, True, cdFilter, 0, tempPathString, cdTitle, ".icc", FormPreferences.hWnd, OFN_HIDEREADONLY) Then
        
        'Save this new directory as the default path for future usage
        Dim listPath As String
        listPath = sFile
        StripDirectory listPath
        g_UserPreferences.SetPref_String "Paths", "Color Profile", listPath
        
        'Set the text box to match this color profile, and save the resulting preference out to file.
        txtColorProfilePath = sFile
        
        Dim hMonitor As Long
        hMonitor = g_cMonitors.Monitors(cmbMonitors.ListIndex + 1).Handle
        g_UserPreferences.SetPref_String "Transparency", "MonitorProfile_" & hMonitor, TrimNull(sFile)
        
        'If the "user custom color profiles" option button isn't selected, mark it now
        If Not optColorManagement(1).Value Then optColorManagement(1).Value = True
        
    End If
    
    'Re-enable user input
    Interface.enableUserInput

End Sub

'Copy the hardware acceleration report to the clipboard
Private Sub cmdCopyReportClipboard_Click()
    Clipboard.Clear
    Clipboard.SetText txtHardware
End Sub

'OK button
Private Sub CmdOK_Click()
    
    Message "Saving preferences..."
    
    'First, make note of the active panel, so we can default to that if the user returns to this dialog
    Dim i As Long
    For i = 0 To cmdCategory.Count - 1
        If cmdCategory(i).Value Then g_UserPreferences.SetPref_Long "Core", "Last Preferences Page", i
    Next i
    
    g_UserPreferences.SetPref_Long "Core", "Last File Preferences Page", cmbFiletype.ListIndex
    
    'We may need to access a generic "form" object multiple times, so I declare it at the top of this sub.
    Dim tForm As Form
    
    'Write preferences out to file in category order.  (The preference XML file is order-agnostic, but I try to
    ' maintain the order used in the Preferences dialog itself to make changes easier.)
    
    '***************************************************************************
    
    'BEGIN Interface preferences
    
        'START/END canvas background color
            g_UserPreferences.SetPref_Long "Interface", "Canvas Background", g_CanvasBackground
            
        'START canvas drop shadow
            g_CanvasDropShadow = CBool(chkDropShadow.Value)
            g_UserPreferences.SetPref_Boolean "Interface", "Canvas Drop Shadow", g_CanvasDropShadow
    
            If g_CanvasDropShadow Then g_CanvasShadow.initializeSquareShadow PD_CANVASSHADOWSIZE, PD_CANVASSHADOWSTRENGTH, g_CanvasBackground
        'END canvas drop shadow
    
        'START image window caption length
    
            'Check to see if the new caption length setting matches the old one.  If it does not, rewrite all form captions to match
            If cmbImageCaption.ListIndex <> g_UserPreferences.GetPref_Long("Interface", "Window Caption Length", 0) Then
                For Each tForm In VB.Forms
                    If tForm.Name = "FormImage" Then
                        If cmbImageCaption.ListIndex = 0 Then
                            tForm.Caption = pdImages(tForm.Tag).originalFileNameAndExtension
                        Else
                            If pdImages(tForm.Tag).locationOnDisk <> "" Then tForm.Caption = pdImages(tForm.Tag).locationOnDisk Else tForm.Caption = pdImages(tForm.Tag).originalFileNameAndExtension
                        End If
                    End If
                Next
            End If
            g_UserPreferences.SetPref_Long "Interface", "Window Caption Length", cmbImageCaption.ListIndex
        
        'END image window caption length
        
        'START MRU caption length
        
            'Similarly, check to see if the new MRU caption setting matches the old one.  If it doesn't, reload the MRU.
            If cmbMRUCaption.ListIndex <> g_UserPreferences.GetPref_Long("Interface", "MRU Caption Length", 0) Then
                g_UserPreferences.SetPref_Long "Interface", "MRU Caption Length", cmbMRUCaption.ListIndex
                g_RecentFiles.MRU_SaveToFile
                g_RecentFiles.MRU_LoadFromFile
                resetMenuIcons
            End If
        
        'END MRU caption length
        
        'START maximum MRU count
            Dim newMaxRecentFiles As Long
            
            'Validate the user's supplied recent file limit
            If tudRecentFiles.IsValid Then
                newMaxRecentFiles = tudRecentFiles.Value
            Else
                newMaxRecentFiles = 10
            End If
            
            'If the max number of recent files has changed, update the MRU list to match
            If newMaxRecentFiles <> g_UserPreferences.GetPref_Long("Interface", "Recent Files Limit", 10) Then
                g_UserPreferences.SetPref_Long "Interface", "Recent Files Limit", tudRecentFiles.Value
                g_RecentFiles.MRU_NotifyNewMaxLimit
            End If
            
        'END maximum MRU count
            
        'START/END interface fonts on modern versions of Windows
            g_UserPreferences.SetPref_Boolean "Interface", "Use Fancy Fonts", g_UseFancyFonts
    
    'END Interface preferences
    
    '***************************************************************************
    
    'BEGIN Loading preferences
    
        'START/END verifying incoming color depth
            g_UserPreferences.SetPref_Boolean "Loading", "Verify Initial Color Depth", CBool(chkInitialColorDepth)
    
        'START/END automatically tone-map HDR images
            g_UserPreferences.SetPref_Boolean "Loading", "HDR Tone Mapping", CBool(chkToneMapping)
        
        'START/END multipage image load behavior
            g_UserPreferences.SetPref_Long "Loading", "Multipage Image Prompt", cmbMultiImage.ListIndex
    
        'START initial zoom
            g_AutozoomLargeImages = cmbLargeImages.ListIndex
            g_UserPreferences.SetPref_Long "Loading", "Initial Image Zoom", g_AutozoomLargeImages
        'END initial zoom
    
    
    'END Loading preferences
    
    '***************************************************************************
    
    'BEGIN Saving preferences
    
        'START prompt on unsaved images
            g_ConfirmClosingUnsaved = CBool(chkConfirmUnsaved.Value)
            g_UserPreferences.SetPref_Boolean "Saving", "Confirm Closing Unsaved", g_ConfirmClosingUnsaved
    
            If g_ConfirmClosingUnsaved Then
                toolbar_File.cmdClose.ToolTip = g_Language.TranslateMessage("Close the current image." & vbCrLf & vbCrLf & "If the current image has not been saved, you will receive a prompt to save it before it closes.")
            Else
                toolbar_File.cmdClose.ToolTip = g_Language.TranslateMessage("Close the current image." & vbCrLf & vbCrLf & "Because you have turned off save prompts (via Tools -> Options), you WILL NOT receive a prompt to save this image before it closes.")
            End If
        'END prompt on unsaved images
    
        'START/END outgoing color depth selection
            g_UserPreferences.SetPref_Long "Saving", "Outgoing Color Depth", cmbExportColorDepth.ListIndex
    
        'START/END Save behavior (overwrite or copy)
            g_UserPreferences.SetPref_Long "Saving", "Overwrite Or Copy", cmbSaveBehavior.ListIndex
        
        'START/END "Save As" dialog's suggested file format
            g_UserPreferences.SetPref_Long "Saving", "Suggested Format", cmbDefaultSaveFormat.ListIndex
    
        'START/END metadata export behavior
            g_UserPreferences.SetPref_Long "Saving", "Metadata Export", cmbMetadata.ListIndex + 1
    
    'END Saving preferences
    
    '***************************************************************************
    
    'START File format preferences
    
        'START/END BMP RLE
            g_UserPreferences.SetPref_Boolean "File Formats", "Bitmap RLE", CBool(chkBMPRLE.Value)
        
        'START/END PNG compression
            g_UserPreferences.SetPref_Long "File Formats", "PNG Compression", hsPNGCompression.Value
        
        'START/END PNG interlacing
            g_UserPreferences.SetPref_Boolean "File Formats", "PNG Interlacing", CBool(chkPNGInterlacing.Value)
        
        'START/END PNG background preservation
            g_UserPreferences.SetPref_Boolean "File Formats", "PNG Background Color", CBool(chkPNGBackground.Value)
        
        'START/END PPM encoding
            g_UserPreferences.SetPref_Long "File Formats", "PPM Export Format", cmbPPMFormat.ListIndex
        
        'START/END TGA RLE encoding
            g_UserPreferences.SetPref_Boolean "File Formats", "TGA RLE", CBool(chkTGARLE.Value)
        
        'START/END TIFF compression
            g_UserPreferences.SetPref_Long "File Formats", "TIFF Compression", cmbTIFFCompression.ListIndex
        
        'START/END TIFF CMYK
            g_UserPreferences.SetPref_Boolean "File Formats", "TIFF CMYK", CBool(chkTIFFCMYK.Value)
    
    'END File format preferences
    
    '***************************************************************************
    
    'START Tools preferences
    
        'START/END clear selections after "Crop to Selection"
            g_UserPreferences.SetPref_Boolean "Tools", "Clear Selection After Crop", CBool(chkSelectionClearCrop.Value)
    
    'END Tools preferences
    
    '***************************************************************************
    
    'START Color and Transparency preferences

        'START use system color profile
            g_UserPreferences.SetPref_Boolean "Transparency", "Use System Color Profile", optColorManagement(0)
            g_UseSystemColorProfile = optColorManagement(0)
        'END use system color profile

        'START alpha checkerboard colors
            g_UserPreferences.SetPref_Long "Transparency", "Alpha Check Mode", CLng(cmbAlphaCheck.ListIndex)
            g_UserPreferences.SetPref_Long "Transparency", "Alpha Check One", CLng(csAlphaOne.Color)
            g_UserPreferences.SetPref_Long "Transparency", "Alpha Check Two", CLng(csAlphaTwo.Color)
        'END alpha checkerboard colors
            
        'START alpha checkerboard size
            g_UserPreferences.SetPref_Long "Transparency", "Alpha Check Size", cmbAlphaCheckSize.ListIndex
            
            'Recreate the cached pattern for the alpha background
            Drawing.createAlphaCheckerboardDIB g_CheckerboardPattern
            
        'END alpha checkerboard size
    
        'START/END validate incoming alpha channel data
            g_UserPreferences.SetPref_Boolean "Transparency", "Validate Alpha Channels", CBool(chkValidateAlpha)
    
    'END Color and Transparency preferences
    
    '***************************************************************************
    
    'BEGIN Update preferences
    
        'START/END allowed to check for updates
            g_UserPreferences.SetPref_Boolean "Updates", "Check For Updates", CBool(chkProgramUpdates)
    
        'Store whether PhotoDemon is allowed to offer the automatic download of missing core plugins
            g_UserPreferences.SetPref_Boolean "Updates", "Prompt For Plugin Download", CBool(chkPromptPluginDownload)
    
    'END Update preferences
    
    '***************************************************************************
    
    'BEGIN Advanced preferences
    
        'START log program messages or not
            g_LogProgramMessages = CBool(chkLogMessages)
            g_UserPreferences.SetPref_Boolean "Advanced", "Log Program Messages", g_LogProgramMessages
        'END log program messages or not
    
        'START/END store the temporary path (but only if it's changed)
            If LCase(TxtTempPath) <> LCase(g_UserPreferences.GetTempPath) Then g_UserPreferences.setTempPath TxtTempPath
    
    'END Advanced preferences
    
    '***************************************************************************
    
    'All user preferences have now been written out to file
    
    'Because some preferences affect the program's interface, redraw the active image.
    FormMain.refreshAllCanvases
    FormMain.mainCanvas(0).BackColor = g_CanvasBackground
    
    If g_OpenImageCount > 0 Then
        PrepareViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
    
    Message "Finished."
        
    Unload Me
    
End Sub

'RESET will regenerate the preferences file from scratch.  This can be an effective way to
' "reset" a copy of the program.
Private Sub cmdReset_Click()

    'Before resetting, warn the user
    Dim confirmReset As VbMsgBoxResult
    confirmReset = pdMsgBox("This action will reset all preferences to their default values.  It cannot be undone." & vbCrLf & vbCrLf & "Are you sure you want to continue?", vbApplicationModal + vbExclamation + vbYesNo, "Reset all preferences")

    'If the user gives final permission, rewrite the preferences file from scratch and repopulate this form
    If confirmReset = vbYes Then
        g_UserPreferences.resetPreferences
        LoadAllPreferences
    End If

End Sub

'When the "..." button is clicked, prompt the user with a "browse for folder" dialog
Private Sub CmdTmpPath_Click()
    Dim tString As String
    tString = BrowseForFolder(Me.hWnd)
    If Len(tString) > 0 Then TxtTempPath.Text = FixPath(tString)
End Sub

'Load all relevant values from the preferences file, and populate their corresponding controls with the user's current settings
Private Sub LoadAllPreferences()
    
    'For the sake of order, we will load preferences by category.  (They can be loaded in any order without consequence,
    ' but there are MANY preferences, so maintaining some kind of order is helpful.)
    
    'Note also that many tooltips are manually populated throughout this section.  This is done for translation
    ' purposes; the tooltips themselves are sometimes too long to fit inside a traditional VB control, so the
    ' IDE dumps them to a separate custom .frx resource file where they are difficult to extract. Rather than
    ' mess with that, I manually add the tooltips here so that the automatic translation engine can easily find
    ' all tooltip text.
    
    '***************************************************************************
    
    'START Interface preferences
        
        'START canvas background (which also requires populating the canvas background combo box)
        
            userInitiatedColorSelection = False
        
            cmbCanvas.Clear
            cmbCanvas.AddItem " system theme: light", 0
            cmbCanvas.AddItem " system theme: dark", 1
            cmbCanvas.AddItem " custom color (click box to customize)", 2
                
            'Select the proper combo box value based on the g_CanvasBackground variable
            If g_CanvasBackground = vb3DLight Then
                'System theme: light
                cmbCanvas.ListIndex = 0
            ElseIf g_CanvasBackground = vb3DShadow Then
                'System theme: dark
                cmbCanvas.ListIndex = 1
            Else
                'Custom color
                cmbCanvas.ListIndex = 2
            End If
            
            originalg_CanvasBackground = g_CanvasBackground
            
            'Draw the current canvas background to the sample picture box
            csCanvasColor.Color = g_CanvasBackground
            
            userInitiatedColorSelection = True
            
            'Finally, provide helpful tooltips for the canvas items
            cmbCanvas.ToolTipText = g_Language.TranslateMessage("The image canvas sits ""behind"" the image on the screen.  Dark colors are generally preferable, as they help the image stand out while you work on it.")
            csCanvasColor.ToolTipText = g_Language.TranslateMessage("Click to change the image window background color")
        
        'END canvas background
        
        'START/END drop shadow between image and canvas
            If g_CanvasDropShadow Then chkDropShadow.Value = vbChecked Else chkDropShadow.Value = vbUnchecked
        
        'START image window caption length
            cmbImageCaption.Clear
            cmbImageCaption.AddItem " compact - file name only", 0
            cmbImageCaption.AddItem " descriptive - full location, including folder(s)", 1
            cmbImageCaption.ListIndex = g_UserPreferences.GetPref_Long("Interface", "Window Caption Length", 0)
        'END image window caption length
        
        'START Recent file max count
            lblRecentFileCount.Caption = g_Language.TranslateMessage("maximum number of recent file entries: ")
            tudRecentFiles.Left = lblRecentFileCount.Left + lblRecentFileCount.Width + fixDPI(6)
            tudRecentFiles.Value = g_UserPreferences.GetPref_Long("Interface", "Recent Files Limit", 10)
        'END
        
        'START MRU caption length
            cmbMRUCaption.Clear
            cmbMRUCaption.AddItem " compact - file names only", 0
            cmbMRUCaption.AddItem " descriptive - full locations, including folder(s)", 1
            cmbMRUCaption.ListIndex = g_UserPreferences.GetPref_Long("Interface", "MRU Caption Length", 0)
            cmbMRUCaption.ToolTipText = g_Language.TranslateMessage("The ""Recent Files"" menu width is limited by Windows.  To prevent this menu from overflowing, PhotoDemon can display image names only instead of full image locations.")
        'END MRU caption length
        
        'START modern typefaces
        
            'Hide the modern typefaces box if the user in on XP.  If the user is on Vista or later, set the box according
            ' to the preference stated in their preferences file.
            If Not g_IsVistaOrLater Then
                chkFancyFonts.Caption = g_Language.TranslateMessage("render PhotoDemon text with modern typefaces (only available on Vista or newer)")
                chkFancyFonts.Enabled = False
            Else
                chkFancyFonts.Caption = g_Language.TranslateMessage("render PhotoDemon text with modern typefaces")
                chkFancyFonts.Enabled = True
                If g_UseFancyFonts Then chkFancyFonts.Value = vbChecked Else chkFancyFonts.Value = vbUnchecked
                originalg_useFancyFonts = g_UseFancyFonts
            End If
            
        'END modern typefaces
        
    'END Interface preferences
    
    '***************************************************************************
    
    'START Loading preferences
    
        'START count unique colors at load time
            If g_UserPreferences.GetPref_Boolean("Loading", "Verify Initial Color Depth", True) Then chkInitialColorDepth.Value = vbChecked Else chkInitialColorDepth.Value = vbUnchecked
            chkInitialColorDepth.ToolTipText = g_Language.TranslateMessage("This option allows PhotoDemon to scan incoming images to determine the most appropriate color depth on a case-by-case basis (rather than relying on the source image file's color depth, which may have been chosen arbitrarily).")
        'END count unique colors at load time
        
        'START tone-mapping HDR images at load time
            If g_UserPreferences.GetPref_Boolean("Loading", "HDR Tone Mapping", True) Then chkToneMapping.Value = vbChecked Else chkToneMapping.Value = vbUnchecked
            
            If g_ImageFormats.FreeImageEnabled Then
                chkToneMapping.Enabled = True
            Else
                chkToneMapping.Caption = g_Language.TranslateMessage("feature disabled due to missing plugin")
                chkToneMapping.Enabled = False
            End If
            
            chkToneMapping.ToolTipText = g_Language.TranslateMessage("Tone mapping is used to preserve the tonal range of HDR images.  This setting is very useful for RAW photos and scanned documents, but it adds a significant amount of time to the image load process.")
        'END tone-mapping HDR images at load time
                
        'START multipage images
            cmbMultiImage.Clear
            cmbMultiImage.AddItem " ask me how I want to proceed", 0
            cmbMultiImage.AddItem " load only the first page", 1
            cmbMultiImage.AddItem " load all pages", 2
            cmbMultiImage.ListIndex = g_UserPreferences.GetPref_Long("Loading", "Multipage Image Prompt", 0)
            
            cmbMultiImage.ToolTipText = g_Language.TranslateMessage("Some image formats can hold multiple images in one file.  When these files are encountered, PhotoDemon can ignore the extra images, or it can load them all for you.")
            
            If Not g_ImageFormats.FreeImageEnabled Then
                cmbMultiImage.Clear
                cmbMultiImage.AddItem " feature disabled due to missing plugin", 0
                cmbMultiImage.ListIndex = 0
                cmbMultiImage.Enabled = False
            Else
                cmbMultiImage.Enabled = True
            End If
            
        'END multipage images
        
        'START initial image zoom
            cmbLargeImages.Clear
            cmbLargeImages.AddItem " automatically fit the image on-screen", 0
            cmbLargeImages.AddItem " 1:1 (100% zoom, or ""actual size"")", 1
            cmbLargeImages.ListIndex = g_UserPreferences.GetPref_Long("Loading", "Initial Image Zoom", 0)
            
            cmbLargeImages.ToolTipText = g_Language.TranslateMessage("Any photo larger than 2 megapixels is too big to fit on an average computer monitor.  PhotoDemon can automatically zoom out on large photographs so that the entire image is viewable.")
        'END initial image zoom
    
    'END Loading preferences
    
    '***************************************************************************
    
    'START Saving preferences
    
        'START/END prompt about unsaved images
            If g_ConfirmClosingUnsaved Then chkConfirmUnsaved.Value = vbChecked Else chkConfirmUnsaved.Value = vbUnchecked
    
        'START exported color depth handling
            cmbExportColorDepth.Clear
            cmbExportColorDepth.AddItem " to match the image file's original color depth", 0
            cmbExportColorDepth.AddItem " automatically", 1
            cmbExportColorDepth.AddItem " by asking me what color depth I want to use", 2
            cmbExportColorDepth.ListIndex = g_UserPreferences.GetPref_Long("Saving", "Outgoing Color Depth", 1)
        
            cmbExportColorDepth.ToolTipText = g_Language.TranslateMessage("Some image file types support multiple color depths.  PhotoDemon's developers suggest letting the software choose the best color depth for you, unless you have reason to choose otherwise.")
        'END exported color depth handling
            
        'START suggested save as format
            cmbDefaultSaveFormat.Clear
            cmbDefaultSaveFormat.AddItem " the current file format of the image being saved", 0
            cmbDefaultSaveFormat.AddItem " the last image format I used in the ""Save As"" screen", 1
            cmbDefaultSaveFormat.ListIndex = g_UserPreferences.GetPref_Long("Saving", "Suggested Format", 0)
            
            cmbDefaultSaveFormat.ToolTipText = g_Language.TranslateMessage("Most photo editors use the format of the current image as the default in the ""Save As"" screen.  When working with RAW images that will eventually be saved to JPEG, it is useful to have PhotoDemon remember that - hence the ""last used"" option.")
        'END suggested save as format
        
        'START overwrite vs copy when saving
            cmbSaveBehavior.Clear
            cmbSaveBehavior.AddItem " overwrite the current file (standard behavior)", 0
            cmbSaveBehavior.AddItem " save a new copy, e.g. ""filename (2).jpg"" (safe behavior)", 1
            cmbSaveBehavior.ListIndex = g_UserPreferences.GetPref_Long("Saving", "Overwrite Or Copy", 0)
            
            cmbSaveBehavior.ToolTipText = g_Language.TranslateMessage("In most photo editors, the ""Save"" command saves the image over its original version, erasing that copy forever.  PhotoDemon provides a ""safer"" option, where each save results in a new copy of the file.")
        'END overwrite vs copy when saving
               
        'START metadata export
            cmbMetadata.Clear
            cmbMetadata.AddItem " preserve all relevant metadata", 0
            cmbMetadata.AddItem " preserve all relevant metadata, but remove personal tags (GPS coords, serial #'s, etc)", 1
            cmbMetadata.AddItem " do not preserve metadata", 2
            
            'Previously we provided an option for "preserve all metadata" at position 0.  This option is no longer available
            ' (for a huge variety of reasons).  To compensate for the removal of position 0, we apply some special handling
            ' to this preference.
            Dim tmpPreferenceLong As Long
            tmpPreferenceLong = g_UserPreferences.GetPref_Long("Saving", "Metadata Export", 0)
            If tmpPreferenceLong > 0 Then tmpPreferenceLong = tmpPreferenceLong - 1
            cmbMetadata.ListIndex = tmpPreferenceLong
            
            cmbMetadata.ToolTipText = g_Language.TranslateMessage("Image metadata is extra data placed in an image file by a camera or photo software.  This data can include things like the make and model of the camera, the GPS coordinates where a photo was taken, or many other items.  To view an image's metadata, use the Image -> Metadata menu.")
        'END metadata export
    
    'END Saving preferences
    
    '***************************************************************************
    
    'START File format preferences
    
        'Prepare the file format selection box.  (No preference is associated with this.)
            cmbFiletype.Clear
            cmbFiletype.AddItem "BMP - Bitmap", 0
            cmbFiletype.AddItem "PNG - Portable Network Graphics", 1
            cmbFiletype.AddItem "PPM - Portable Pixmap", 2
            cmbFiletype.AddItem "TGA - Truevision (TARGA)", 3
            cmbFiletype.AddItem "TIFF - Tagged Image File Format", 4
            cmbFiletype.ListIndex = 0
            
            cmbFiletype.ToolTipText = g_Language.TranslateMessage("Some image file types support additional parameters when importing and exporting.  By default, PhotoDemon will manage these for you, but you can specify different parameters if necessary.")
            
        'BMP
        
            'START/END RLE encoding for bitmaps
                If g_UserPreferences.GetPref_Boolean("File Formats", "Bitmap RLE", False) Then chkBMPRLE.Value = vbChecked Else chkBMPRLE.Value = vbUnchecked
                chkBMPRLE.ToolTipText = g_Language.TranslateMessage("Bitmap files only support one type of compression, and they only support it for certain color depths.  PhotoDemon can apply simple RLE compression when saving 8bpp images.")
        
        'PNG
        
            'START/END PNG compression level
                hsPNGCompression.Value = g_UserPreferences.GetPref_Long("File Formats", "PNG Compression", 9)
    
            'START/END interlacing
                If g_UserPreferences.GetPref_Boolean("File Formats", "PNG Interlacing", False) Then chkPNGInterlacing.Value = vbChecked Else chkPNGInterlacing.Value = vbUnchecked
                chkPNGInterlacing.ToolTipText = g_Language.TranslateMessage("PNG interlacing is similar to ""progressive scan"" on JPEGs.  Interlacing slightly increases file size, but an interlaced image can ""fade-in"" while it downloads.")
            
            'START/END background color preservation
                If g_UserPreferences.GetPref_Boolean("File Formats", "PNG Background Color", True) Then chkPNGBackground.Value = vbChecked Else chkPNGBackground.Value = vbUnchecked
                chkPNGBackground.ToolTipText = g_Language.TranslateMessage("PNG files can contain a background color parameter.  This takes up extra space in the file, so feel free to disable it if you don't need background colors.")
        
        'PPM
    
            'START PPM export format
                cmbPPMFormat.Clear
                cmbPPMFormat.AddItem " binary encoding (faster, smaller file size)", 0
                cmbPPMFormat.AddItem " ASCII encoding (human-readable, multi-platform)", 1
                cmbPPMFormat.ListIndex = g_UserPreferences.GetPref_Long("File Formats", "PPM Export Format", 0)
                
                cmbPPMFormat.ToolTipText = g_Language.TranslateMessage("Binary encoding of PPM files is strongly suggested.  (In other words, don't change this setting unless you are certain that ASCII encoding is what you want. :)")
            'END PPM export format
    
        'TGA
    
            'START/END TGA RLE encoding
                If g_UserPreferences.GetPref_Boolean("File Formats", "TGA RLE", False) Then chkTGARLE.Value = vbChecked Else chkTGARLE.Value = vbUnchecked
                chkTGARLE.ToolTipText = g_Language.TranslateMessage("TGA files only support one type of compression.  PhotoDemon can apply simple RLE compression when saving TGA images.")
        
        'TIFF
    
            'START TIFF compression (many options)
                cmbTIFFCompression.Clear
                cmbTIFFCompression.AddItem " default settings - CCITT Group 4 for 1bpp, LZW for all others", 0
                cmbTIFFCompression.AddItem " no compression", 1
                cmbTIFFCompression.AddItem " Macintosh PackBits (RLE)", 2
                cmbTIFFCompression.AddItem " Official DEFLATE ('Adobe-style')", 3
                cmbTIFFCompression.AddItem " PKZIP DEFLATE (also known as zLib DEFLATE)", 4
                cmbTIFFCompression.AddItem " LZW", 5
                cmbTIFFCompression.AddItem " JPEG - 8bpp grayscale or 24bpp color only", 6
                cmbTIFFCompression.AddItem " CCITT Group 3 fax encoding - 1bpp only", 7
                cmbTIFFCompression.AddItem " CCITT Group 4 fax encoding - 1bpp only", 8
                
                cmbTIFFCompression.ListIndex = g_UserPreferences.GetPref_Long("File Formats", "TIFF Compression", 0)
                
                cmbTIFFCompression.ToolTipText = g_Language.TranslateMessage("TIFFs support a variety of compression techniques.  Some of these techniques are limited to specific color depths, so make sure you pick one that matches the images you plan on saving.")
            'END TIFF compression
                
            'START/END TIFF CMYK encoding
                If g_UserPreferences.GetPref_Boolean("File Formats", "TIFF CMYK", False) Then chkTIFFCMYK.Value = vbChecked Else chkTIFFCMYK.Value = vbUnchecked
                chkTIFFCMYK.ToolTipText = g_Language.TranslateMessage("TIFFs support both RGB and CMYK color spaces.  RGB is used by default, but if a TIFF file is going to be used in printed document, CMYK is sometimes required.")
        
    'END File format preferences
    
    '***************************************************************************
    
    'START Tools preferences
    
        'START Clear selections after "Crop to Selection"
            If g_UserPreferences.GetPref_Boolean("Tools", "Clear Selection After Crop", True) Then chkSelectionClearCrop.Value = vbChecked Else chkSelectionClearCrop.Value = vbUnchecked
            chkSelectionClearCrop.ToolTipText = g_Language.TranslateMessage("When the ""Crop to Selection"" command is used, the resulting image will always contain a selection the same size as the full image.  There is generally no need to retain this, so PhotoDemon can automatically clear it for you.")
        'END Clear selections after "Crop to Selection"
        
    'END Tools preferences
    
    '***************************************************************************
    
    'START Color and Transparency preferences
    
        'START color management preferences
            
            'Set the option buttons according to the user's preference
            If g_UserPreferences.GetPref_Boolean("Transparency", "Use System Color Profile", True) Then optColorManagement(0).Value = True Else optColorManagement(1).Value = True
            
            'Load a list of all available monitors
            cmbMonitors.Clear
            
            Dim primaryMonitor As String, secondaryMonitor As String
            primaryMonitor = g_Language.TranslateMessage("Primary monitor") & ": "
            secondaryMonitor = g_Language.TranslateMessage("Secondary monitor") & ": "
            
            Dim primaryIndex As Long
            
            Dim monitorEntry As String
            
            Dim i As Long
            For i = 1 To g_cMonitors.Monitors.Count
                monitorEntry = ""
                
                'Explicitly label the primary monitor
                If g_cMonitors.Monitors(i).isPrimary Then
                    monitorEntry = primaryMonitor
                    primaryIndex = i - 1
                Else
                    monitorEntry = secondaryMonitor
                End If
                
                'Add the monitor's physical size
                monitorEntry = monitorEntry & g_cMonitors.Monitors(i).getMonitorSizeAsString
                
                'Add the monitor's name
                monitorEntry = monitorEntry & " " & g_cMonitors.Monitors(i).getBestMonitorName
                
                'Add the monitor's native resolution
                monitorEntry = monitorEntry & " (" & g_cMonitors.Monitors(i).getMonitorResolutionAsString & ")"
                
                'Add the monitor's description (typically the video card driving the monitor)
                'monitorEntry = monitorEntry & " (" & g_cMonitors.Monitors(i).Description & ")"
                
                'Display this monitor in the list
                cmbMonitors.AddItem monitorEntry, i - 1
                
            Next i
            
            'Display the primary monitor by default; this will also trigger a load of the matching
            ' custom profile, if one exists.
            cmbMonitors.ListIndex = primaryIndex
            
            'Add tooltips to all color-profile-related controls
            optColorManagement(0).ToolTipText = g_Language.TranslateMessage("This setting is the best choice for most users.  If you have no idea what color management is, use this setting.  If you have correctly configured a display profile via the Windows Control Panel, also use this setting.")
            optColorManagement(1).ToolTipText = g_Language.TranslateMessage("To configure custom color profiles on a per-monitor basis, please use this setting.")
            
            cmbMonitors.ToolTipText = g_Language.TranslateMessage("Please specify a color profile for each monitor currently attached to the system.  Note that the text in parentheses is the display adapter driving the named monitor.")
            cmdColorProfilePath.ToolTipText = g_Language.TranslateMessage("Click this button to bring up a ""browse for color profile"" dialog.")
        
        'END color management preferences
    
        'START alpha-channel checkerboard rendering
            userInitiatedAlphaSelection = False
            cmbAlphaCheck.Clear
            cmbAlphaCheck.AddItem " Highlight checks", 0
            cmbAlphaCheck.AddItem " Midtone checks", 1
            cmbAlphaCheck.AddItem " Shadow checks", 2
            cmbAlphaCheck.AddItem " Custom (click boxes to customize)", 3
            
            cmbAlphaCheck.ListIndex = g_UserPreferences.GetPref_Long("Transparency", "Alpha Check Mode", 0)
            
            csAlphaOne.Color = g_UserPreferences.GetPref_Long("Transparency", "Alpha Check One", RGB(255, 255, 255))
            csAlphaTwo.Color = g_UserPreferences.GetPref_Long("Transparency", "Alpha Check Two", RGB(204, 204, 204))
            
            cmbAlphaCheck.ToolTipText = g_Language.TranslateMessage("If an image has transparent areas, a checkerboard is typically displayed ""behind"" the image.  This box lets you change the checkerboard's colors.")
            csAlphaOne.ToolTipText = g_Language.TranslateMessage("Click to change the first checkerboard background color for alpha channels")
            csAlphaTwo.ToolTipText = g_Language.TranslateMessage("Click to change the second checkerboard background color for alpha channels")
            
            userInitiatedAlphaSelection = True
        'END alpha-channel checkerboard rendering
        
        'START alpha-channel checkerboard size
            cmbAlphaCheckSize.Clear
            cmbAlphaCheckSize.AddItem " Small (4x4 pixels)", 0
            cmbAlphaCheckSize.AddItem " Medium (8x8 pixels)", 1
            cmbAlphaCheckSize.AddItem " Large (16x16 pixels)", 2
            
            cmbAlphaCheckSize.ListIndex = g_UserPreferences.GetPref_Long("Transparency", "Alpha Check Size", 1)
            
            cmbAlphaCheckSize.ToolTipText = g_Language.TranslateMessage("If an image has transparent areas, a checkerboard is typically displayed ""behind"" the image.  This box lets you change the checkerboard's size.")
        'END alpha-channel checkerboard size
        
        'START/END validate incoming alpha channels
            If g_UserPreferences.GetPref_Boolean("Transparency", "Validate Alpha Channels", True) Then chkValidateAlpha.Value = vbChecked Else chkValidateAlpha.Value = vbUnchecked
            chkValidateAlpha.ToolTipText = g_Language.TranslateMessage("When checked, this option allows PhotoDemon to automatically remove empty alpha channels from imported images. This improves program performance, reduces RAM usage, and improves file size on exported files.")

    'END Color and Transparency preferences
    
    '***************************************************************************
    
    'START Update preferences
    
        'START/END check for software updates
            If g_UserPreferences.GetPref_Boolean("Updates", "Check For Updates", True) Then chkProgramUpdates.Value = vbChecked Else chkProgramUpdates.Value = vbUnchecked
            
        'START prompt for missing plugin download
            If g_UserPreferences.GetPref_Boolean("Updates", "Prompt For Plugin Download", True) Then chkPromptPluginDownload.Value = vbChecked Else chkPromptPluginDownload.Value = vbUnchecked
            chkPromptPluginDownload.ToolTipText = g_Language.TranslateMessage("PhotoDemon relies on several free, open-source plugins for full functionality. If any of these plugins are missing (for example, if you downloaded PhotoDemon from a 3rd-party site), this option will offer to download the missing plugins for you.")
        'END prompt for missing plugin download
    
        'Populate the network access disclaimer in the "Update" panel
            lblExplanation.Caption = g_Language.TranslateMessage("PhotoDemon provides two non-essential features that require Internet access: checking for software updates, and offering to download core plugins if they aren't present in the \App\PhotoDemon\Plugins subdirectory." & vbCrLf & vbCrLf & "The developers of PhotoDemon take privacy very seriously, so no information - statistical or otherwise - is uploaded by these features. Checking for software updates involves downloading a single ""updates.txt"" file containing the latest software version number. Similarly, downloading missing plugins simply involves downloading one or more compressed plugin files from the PhotoDemon server." & vbCrLf & vbCrLf & "If you choose to disable these features, you can always visit photodemon.org to manually download the most recent version of the program.")
    
    'END Update preferences
    
    '***************************************************************************
    
    'START Advanced preferences
    
        'START log program messages
            If g_LogProgramMessages Then chkLogMessages.Value = vbChecked Else chkLogMessages.Value = vbUnchecked
            chkLogMessages.ToolTipText = g_Language.TranslateMessage("If this is checked, PhotoDemon will create a human-readable .log file that contains the text of every message displayed on the progress bar.  This will increase processing time, so only check this option if you really need debugging data.")
        'END log program messages
            
        'Display the current temporary file path
            TxtTempPath.Text = g_UserPreferences.GetTempPath
    
        'Display what we know about this PC's hardware acceleration capabilities
            txtHardware = getDeviceCapsString()
            
        '...and give the "copy to clipboard" button a tooltip
            cmdCopyReportClipboard.ToolTip = g_Language.TranslateMessage("Copy the report to the system clipboard")
        
        'Display what we know about PD's memory usage
            lblMemoryUsageCurrent.Caption = g_Language.TranslateMessage("current PhotoDemon memory usage:") & " " & Format(Str(GetPhotoDemonMemoryUsage()), "###,###,###,###") & " K"
            lblMemoryUsageMax.Caption = g_Language.TranslateMessage("max PhotoDemon memory usage this session:") & " " & Format(Str(GetPhotoDemonMemoryUsage(True)), "###,###,###,###") & " K"
            If Not g_IsProgramCompiled Then
                lblMemoryUsageCurrent = lblMemoryUsageCurrent.Caption & " (" & g_Language.TranslateMessage("reading not accurate inside IDE") & ")"
                lblMemoryUsageMax = lblMemoryUsageMax.Caption & " (" & g_Language.TranslateMessage("reading not accurate inside IDE") & ")"
            End If
    
    'END Advanced preferences
    
    '***************************************************************************
    
    'All preference controls are now initialized with the matching value stored in the preferences file
    
    
    'Some preferences rely on the presence of the FreeImage plugin.  If the FreeImage plugin is not available,
    ' display a warning about preferences not working as expected.
    If Not g_ImageFormats.FreeImageEnabled Then
        lblFreeImageWarning.Caption = g_Language.TranslateMessage("NOTE: options on this page have been disabled because the FreeImage plugin could not be located.")
        lblFreeImageWarning.Visible = True
        lblFileFreeImageWarning.Caption = g_Language.TranslateMessage("NOTE: Many of these file format options require the FreeImage plugin.  Because you do not have the FreeImage plugin installed, these options may not perform as expected.")
        lblFileFreeImageWarning.Visible = True
    Else
        lblFreeImageWarning.Visible = False
        lblFileFreeImageWarning.Visible = False
    End If
    
End Sub

'When new transparency checkerboard colors are selected, change the corresponding list box to match
Private Sub csAlphaOne_ColorChanged()
    
    If userInitiatedAlphaSelection Then
        userInitiatedAlphaSelection = False
        cmbAlphaCheck.ListIndex = 3         '3 corresponds to "custom colors"
        userInitiatedAlphaSelection = True
    End If
    
End Sub

Private Sub csAlphaTwo_ColorChanged()
    
    If userInitiatedAlphaSelection Then
        userInitiatedAlphaSelection = False
        cmbAlphaCheck.ListIndex = 3         '3 corresponds to "custom colors"
        userInitiatedAlphaSelection = True
    End If
    
End Sub

'When a new canvas background color is selected, update the corresponding list box as necessary
Private Sub csCanvasColor_ColorChanged()
    
    g_CanvasBackground = csCanvasColor.Color
    
    If userInitiatedColorSelection Then
    
        userInitiatedColorSelection = False
        
        'System theme: light
        If g_CanvasBackground = vb3DLight Then
            If cmbCanvas.ListIndex <> 0 Then cmbCanvas.ListIndex = 0
        
        'System theme: dark
        ElseIf g_CanvasBackground = vb3DShadow Then
            If cmbCanvas.ListIndex <> 1 Then cmbCanvas.ListIndex = 1
        
        'Custom color
        Else
            If cmbCanvas.ListIndex <> 2 Then cmbCanvas.ListIndex = 2
        End If
        
        userInitiatedColorSelection = True
        
    End If
    
End Sub

'When the form is loaded, populate the various checkboxes and textboxes with the values from the preferences file
Private Sub Form_Load()

    'Populate all controls with the corresponding values from the preferences file
    LoadAllPreferences
    
    'Populate the multi-line tooltips for the category command buttons (on the left)
        'Interface
        cmdCategory(0).ToolTip = g_Language.TranslateMessage("Interface options include settings for the main PhotoDemon interface, including things like canvas settings, font selection, and positioning.")
        'Loading
        cmdCategory(1).ToolTip = g_Language.TranslateMessage("Load options allow you to customize the way image files enter the application.")
        'Saving
        cmdCategory(2).ToolTip = g_Language.TranslateMessage("Save options allow you to customize the way image files leave the application.")
        'File formats
        cmdCategory(3).ToolTip = g_Language.TranslateMessage("File format options control how PhotoDemon handles certain types of images.")
        'Tools
        cmdCategory(4).ToolTip = g_Language.TranslateMessage("Tool options currently include customizable options for the Selection Tool. In the future, PhotoDemon will gain paint tools, and those settings will appear here as well.")
        'Color and transparency
        cmdCategory(5).ToolTip = g_Language.TranslateMessage("Color and transparency options include settings for color management (ICC profiles), and alpha channel handling.")
        'Updates
        cmdCategory(6).ToolTip = g_Language.TranslateMessage("Update options control how frequently PhotoDemon checks for updated versions, and how it handles the download of missing plugins.")
        'Advanced
        cmdCategory(7).ToolTip = g_Language.TranslateMessage("Advanced options can be safely ignored by regular users. Testers and developers may, however, find these settings useful.")
    
    'Hide all category panels (the proper one will be activated in a moment)
    Dim i As Long
    For i = 0 To picContainer.Count - 1
        picContainer(i).Visible = False
        cmdCategory(i).Value = False
    Next i
    For i = 0 To picFileContainer.Count - 1
        picFileContainer(i).Visible = False
    Next i
    
    'Activate the last preferences panel that the user looked at
    picContainer(g_UserPreferences.GetPref_Long("Core", "Last Preferences Page", 0)).Visible = True
    cmdCategory(g_UserPreferences.GetPref_Long("Core", "Last Preferences Page", 0)).Value = True
    
    'Also, activate the last file format preferences sub-panel that the user looked at
    cmbFiletype.ListIndex = g_UserPreferences.GetPref_Long("Core", "Last File Preferences Page", 1)
    picFileContainer(g_UserPreferences.GetPref_Long("Core", "Last File Preferences Page", 1)).Visible = True
    
    'Translate and decorate the form; note that a custom tooltip object is passed.  makeFormPretty will automatically
    ' populate this object for us, which allows for themed and multiline tooltips.
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'For some reason, the container picture boxes automatically acquire the pointer of children objects.
    ' Manually force those cursors to arrows to prevent this.
    For i = 0 To picContainer.Count - 1
        setArrowCursor picContainer(i)
    Next i
    
    For i = 0 To picFileContainer.Count - 1
        setArrowCursor picFileContainer(i)
    Next i
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'If the selected temp folder doesn't have write access, warn the user
Private Sub TxtTempPath_Change()
    If Not DirectoryExist(TxtTempPath.Text) Then
        lblTempPathWarning.Caption = g_Language.TranslateMessage("WARNING: this folder is invalid (access prohibited).  Please provide a valid folder.  If no new folder is provided, PhotoDemon will use the system's default temp location.")
        lblTempPathWarning.Visible = True
    Else
        lblTempPathWarning.Visible = False
    End If
End Sub

