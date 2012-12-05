VERSION 5.00
Begin VB.Form FormPreferences 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " PhotoDemon Preferences"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11265
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
   ScaleHeight     =   490
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   751
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset all Preferences"
      Height          =   495
      Left            =   2880
      TabIndex        =   38
      ToolTipText     =   "Use this to reset all preferences to their default state.  This action cannot be undone."
      Top             =   6600
      Width           =   2085
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   9720
      TabIndex        =   1
      Top             =   6630
      Width           =   1245
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   8400
      TabIndex        =   0
      Top             =   6630
      Width           =   1245
   End
   Begin PhotoDemon.jcbutton cmdCategory 
      Height          =   780
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2265
      _ExtentX        =   3995
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
      Caption         =   "Interface"
      ForeColor       =   4210752
      ForeColorHover  =   4194304
      Mode            =   2
      Value           =   -1  'True
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormPreferences.frx":0000
      PictureAlign    =   0
      PictureEffectOnDown=   0
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipType     =   1
      TooltipTitle    =   "Interface Preferences"
   End
   Begin PhotoDemon.jcbutton cmdCategory 
      Height          =   780
      Index           =   6
      Left            =   120
      TabIndex        =   5
      Top             =   5160
      Width           =   2265
      _ExtentX        =   3995
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
      Caption         =   "Updates"
      ForeColor       =   4210752
      ForeColorHover  =   4194304
      Mode            =   2
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormPreferences.frx":1052
      PictureAlign    =   0
      PictureEffectOnDown=   0
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipType     =   1
      TooltipTitle    =   "Update Preferences"
   End
   Begin PhotoDemon.jcbutton cmdCategory 
      Height          =   780
      Index           =   4
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   2265
      _ExtentX        =   3995
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
      Caption         =   "Tools"
      ForeColor       =   4210752
      ForeColorHover  =   4194304
      Mode            =   2
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormPreferences.frx":20A4
      PictureAlign    =   0
      PictureEffectOnDown=   0
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipType     =   1
      TooltipTitle    =   "Tool Preferences"
   End
   Begin PhotoDemon.jcbutton cmdCategory 
      Height          =   780
      Index           =   7
      Left            =   120
      TabIndex        =   6
      Top             =   6000
      Width           =   2265
      _ExtentX        =   3995
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
      Caption         =   "Advanced"
      ForeColor       =   4210752
      ForeColorHover  =   4194304
      Mode            =   2
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormPreferences.frx":30F6
      PictureAlign    =   0
      PictureEffectOnDown=   0
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipType     =   1
      TooltipTitle    =   "Advanced Settings"
   End
   Begin PhotoDemon.jcbutton cmdCategory 
      Height          =   780
      Index           =   5
      Left            =   120
      TabIndex        =   4
      Top             =   4320
      Width           =   2265
      _ExtentX        =   3995
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
      Caption         =   "Transparency"
      ForeColor       =   4210752
      ForeColorHover  =   4194304
      Mode            =   2
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormPreferences.frx":4148
      PictureAlign    =   0
      PictureEffectOnDown=   0
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipType     =   1
      TooltipTitle    =   "Transparency preferences"
   End
   Begin PhotoDemon.jcbutton cmdCategory 
      Height          =   780
      Index           =   1
      Left            =   120
      TabIndex        =   42
      Top             =   960
      Width           =   2265
      _ExtentX        =   3995
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
      Caption         =   "Loading"
      ForeColor       =   4210752
      ForeColorHover  =   4194304
      Mode            =   2
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormPreferences.frx":519A
      PictureAlign    =   0
      PictureEffectOnDown=   0
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipType     =   1
      TooltipTitle    =   "Load (Import) Preferences"
   End
   Begin PhotoDemon.jcbutton cmdCategory 
      Height          =   780
      Index           =   2
      Left            =   120
      TabIndex        =   65
      Top             =   1800
      Width           =   2265
      _ExtentX        =   3995
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
      Caption         =   "Saving"
      ForeColor       =   4210752
      ForeColorHover  =   4194304
      Mode            =   2
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormPreferences.frx":61EC
      PictureAlign    =   0
      PictureEffectOnDown=   0
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipType     =   1
      TooltipTitle    =   "Save (Export) Preferences"
   End
   Begin PhotoDemon.jcbutton cmdCategory 
      Height          =   780
      Index           =   3
      Left            =   120
      TabIndex        =   69
      Top             =   2640
      Width           =   2265
      _ExtentX        =   3995
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
      Caption         =   "Performance"
      ForeColor       =   4210752
      ForeColorHover  =   4194304
      Mode            =   2
      HandPointer     =   -1  'True
      PictureNormal   =   "VBP_FormPreferences.frx":723E
      PictureAlign    =   0
      PictureEffectOnDown=   0
      DisabledPictureMode=   1
      CaptionEffects  =   0
      TooltipType     =   1
      TooltipTitle    =   "Performance Preferences"
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6000
      Index           =   2
      Left            =   2760
      MousePointer    =   1  'Arrow
      ScaleHeight     =   400
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   66
      Top             =   345
      Width           =   8295
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
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   74
         ToolTipText     =   $"VBP_FormPreferences.frx":8290
         Top             =   2070
         Width           =   7215
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
         Left            =   840
         TabIndex        =   67
         ToolTipText     =   "Check this if you want to be warned when you try to close an image with unsaved changes"
         Top             =   840
         Width           =   7215
      End
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
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
         Left            =   360
         TabIndex        =   77
         Top             =   480
         Width           =   2505
      End
      Begin VB.Label Label1 
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
         Left            =   840
         TabIndex        =   76
         Top             =   1740
         Width           =   6585
      End
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
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
         Left            =   360
         TabIndex        =   75
         Top             =   1320
         Width           =   3285
      End
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
         Caption         =   "save preferences for specific image types"
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
         TabIndex        =   73
         Top             =   2760
         Width           =   4320
      End
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
         Caption         =   "global save preferences"
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
         TabIndex        =   68
         Top             =   0
         Width           =   2475
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6000
      Index           =   1
      Left            =   2760
      MousePointer    =   1  'Arrow
      ScaleHeight     =   400
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   43
      Top             =   345
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
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   57
         ToolTipText     =   $"VBP_FormPreferences.frx":834A
         Top             =   2340
         Width           =   4095
      End
      Begin VB.CheckBox chkToneMapping 
         Appearance      =   0  'Flat
         Caption         =   "apply tone mapping to imported HDR and RAW images (48, 64, 96, 128bpp)"
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
         Left            =   360
         TabIndex        =   55
         ToolTipText     =   $"VBP_FormPreferences.frx":83F2
         Top             =   405
         Width           =   7455
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
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   44
         ToolTipText     =   $"VBP_FormPreferences.frx":84BC
         Top             =   1350
         Width           =   4815
      End
      Begin VB.Label lblFreeImageWarning 
         Caption         =   $"VBP_FormPreferences.frx":8576
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
         Height          =   735
         Left            =   120
         TabIndex        =   59
         Top             =   5160
         Width           =   8175
      End
      Begin VB.Label lblMultiImages 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "if a file contains multiple images: "
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
         Left            =   360
         TabIndex        =   58
         Top             =   2400
         Width           =   2895
      End
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
         Caption         =   "multi-image files (animated GIF, multipage TIFF)"
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
         Left            =   120
         TabIndex        =   56
         Top             =   1920
         Width           =   5205
      End
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
         Caption         =   "high-dynamic range (HDR) files"
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
         Left            =   120
         TabIndex        =   54
         Top             =   0
         Width           =   3345
      End
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
         Caption         =   "initial viewport"
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
         Left            =   120
         TabIndex        =   53
         Top             =   960
         Width           =   1560
      End
      Begin VB.Label lblImgOpen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "set initial image zoom to: "
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
         Left            =   360
         TabIndex        =   45
         Top             =   1410
         Width           =   2235
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6000
      Index           =   4
      Left            =   2760
      MousePointer    =   1  'Arrow
      ScaleHeight     =   400
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   12
      Top             =   345
      Width           =   8295
      Begin VB.CheckBox chkSelectionClearCrop 
         Appearance      =   0  'Flat
         Caption         =   " automatically clear the active selection after ""Crop to Selection"" is used"
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
         Left            =   360
         TabIndex        =   64
         ToolTipText     =   $"VBP_FormPreferences.frx":868B
         Top             =   480
         Width           =   7455
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
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
         Left            =   120
         TabIndex        =   13
         Top             =   0
         Width           =   1020
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6000
      Index           =   7
      Left            =   2760
      MousePointer    =   1  'Arrow
      ScaleHeight     =   400
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   19
      Top             =   345
      Width           =   8295
      Begin VB.CheckBox chkGDIPlusTest 
         Appearance      =   0  'Flat
         Caption         =   " enable GDI+ support"
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
         TabIndex        =   37
         ToolTipText     =   "Use this to manually disable GDI+ support. This forces PhotoDemon to rely on its FreeImage and internal VB-only routines."
         Top             =   3360
         Width           =   7815
      End
      Begin VB.CheckBox chkFreeImageTest 
         Appearance      =   0  'Flat
         Caption         =   " enable FreeImage support"
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
         TabIndex        =   36
         ToolTipText     =   "Use this to manually disable FreeImage support. This forces PhotoDemon to rely on its GDI+ and internal VB-only routines."
         Top             =   2880
         Width           =   7815
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
         TabIndex        =   23
         Text            =   "automatically generated at run-time"
         ToolTipText     =   "Folder used for temporary files"
         Top             =   4320
         Width           =   6975
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
         Left            =   7320
         TabIndex        =   22
         ToolTipText     =   "Click to open a browse-for-folder dialog"
         Top             =   4320
         Width           =   405
      End
      Begin VB.CheckBox ChkLogMessages 
         Appearance      =   0  'Flat
         Caption         =   " log all program messages to file "
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
         TabIndex        =   21
         ToolTipText     =   $"VBP_FormPreferences.frx":877D
         Top             =   480
         Width           =   6975
      End
      Begin VB.Label lblRuntimeSettings 
         AutoSize        =   -1  'True
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
         Left            =   120
         TabIndex        =   63
         Top             =   3840
         Width           =   2385
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
         Height          =   360
         Left            =   240
         TabIndex        =   62
         Top             =   1890
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
         Height          =   360
         Left            =   240
         TabIndex        =   61
         Top             =   1440
         Width           =   7965
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
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
         Left            =   120
         TabIndex        =   60
         Top             =   960
         Width           =   2130
      End
      Begin VB.Label lblTempPathWarning 
         Caption         =   $"VBP_FormPreferences.frx":886F
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
         TabIndex        =   41
         Top             =   4920
         Width           =   7695
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
         TabIndex        =   35
         Top             =   2400
         Width           =   7155
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
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
         Left            =   120
         TabIndex        =   20
         Top             =   0
         Width           =   1125
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6000
      Index           =   6
      Left            =   2760
      MousePointer    =   1  'Arrow
      ScaleHeight     =   400
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   14
      Top             =   345
      Width           =   8295
      Begin VB.CheckBox ChkPromptPluginDownload 
         Appearance      =   0  'Flat
         Caption         =   " if core plugins cannot be located, offer to download them"
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
         TabIndex        =   17
         ToolTipText     =   $"VBP_FormPreferences.frx":891E
         Top             =   1080
         Width           =   6735
      End
      Begin VB.CheckBox chkProgramUpdates 
         Appearance      =   0  'Flat
         Caption         =   " automatically check for software updates every 10 days"
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   18
         Top             =   1800
         Width           =   7935
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6000
      Index           =   0
      Left            =   2760
      MousePointer    =   1  'Arrow
      ScaleHeight     =   400
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   7
      Top             =   345
      Width           =   8295
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
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   50
         ToolTipText     =   $"VBP_FormPreferences.frx":89BA
         Top             =   2520
         Width           =   4575
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
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   47
         ToolTipText     =   "Image windows tend to be large, so feel free to display each image's full location in the image window title bars."
         Top             =   2040
         Width           =   4575
      End
      Begin VB.CheckBox chkWindowLocation 
         Appearance      =   0  'Flat
         Caption         =   " remember the main window's on-screen location between sessions"
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
         ToolTipText     =   "If selected, this setting will ensure that PhotoDemon starts up in the on-screen location where you last left it."
         Top             =   4680
         Width           =   7815
      End
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
         TabIndex        =   34
         ToolTipText     =   " This setting helps images stand out from the canvas behind them"
         Top             =   960
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
         TabIndex        =   30
         ToolTipText     =   "This setting uses ""Segoe UI"" as the PhotoDemon interface font. Leaving it unchecked defaults to ""Tahoma""."
         Top             =   3600
         Width           =   7695
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
         ToolTipText     =   $"VBP_FormPreferences.frx":8A63
         Top             =   480
         Width           =   4815
      End
      Begin VB.PictureBox picCanvasColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   7680
         MouseIcon       =   "VBP_FormPreferences.frx":8AFC
         MousePointer    =   99  'Custom
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   37
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Click to change the image window background color"
         Top             =   480
         Width           =   585
      End
      Begin VB.Label lblMRUCaption 
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
         TabIndex        =   51
         Top             =   2580
         Width           =   3315
      End
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
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
         Left            =   120
         TabIndex        =   49
         Top             =   1560
         Width           =   870
      End
      Begin VB.Label lblImageCaption 
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
         TabIndex        =   48
         Top             =   2100
         Width           =   2730
      End
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
         Caption         =   "window location"
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
         TabIndex        =   39
         Top             =   4200
         Width           =   1725
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
         TabIndex        =   33
         Top             =   3120
         Width           =   1365
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
         TabIndex        =   11
         Top             =   540
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
      Height          =   6000
      Index           =   5
      Left            =   2760
      MousePointer    =   1  'Arrow
      ScaleHeight     =   400
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   24
      Top             =   345
      Width           =   8295
      Begin VB.CheckBox chkValidateAlpha 
         Appearance      =   0  'Flat
         Caption         =   " automatically validate all incoming alpha channels"
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
         Left            =   360
         TabIndex        =   46
         ToolTipText     =   $"VBP_FormPreferences.frx":8C4E
         Top             =   3240
         Width           =   7695
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
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   31
         ToolTipText     =   $"VBP_FormPreferences.frx":8D20
         Top             =   2010
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
         TabIndex        =   28
         ToolTipText     =   $"VBP_FormPreferences.frx":8DB3
         Top             =   900
         Width           =   5055
      End
      Begin VB.PictureBox picAlphaOne 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   5520
         MouseIcon       =   "VBP_FormPreferences.frx":8E48
         MousePointer    =   99  'Custom
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   37
         TabIndex        =   27
         TabStop         =   0   'False
         ToolTipText     =   "Click to change the second checkerboard background color for alpha channels"
         Top             =   900
         Width           =   585
      End
      Begin VB.PictureBox picAlphaTwo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   6240
         MouseIcon       =   "VBP_FormPreferences.frx":8F9A
         MousePointer    =   99  'Custom
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   37
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Click to change the second checkerboard background color for alpha channels"
         Top             =   900
         Width           =   585
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "validation"
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
         TabIndex        =   52
         Top             =   2760
         Width           =   1020
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
         TabIndex        =   32
         Top             =   1590
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
         TabIndex        =   29
         Top             =   480
         Width           =   2970
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "appearance"
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
         TabIndex        =   25
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6000
      Index           =   3
      Left            =   2760
      MousePointer    =   1  'Arrow
      ScaleHeight     =   400
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   553
      TabIndex        =   70
      Top             =   345
      Width           =   8295
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   " sample checkbox"
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
         Left            =   360
         TabIndex        =   71
         ToolTipText     =   "Check this if you want to be warned when you try to close an image with unsaved changes"
         Top             =   480
         Width           =   7215
      End
      Begin VB.Label lblInterfaceTitle 
         AutoSize        =   -1  'True
         Caption         =   "performance preferences"
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
         Left            =   120
         TabIndex        =   72
         Top             =   0
         Width           =   2625
      End
   End
   Begin VB.Line lneVertical 
      BorderColor     =   &H8000000D&
      X1              =   168
      X2              =   168
      Y1              =   8
      Y2              =   484
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

'Some settings are odd - I want them to update in real-time, so the user can see the effects of the change.  But if the user presses
' "cancel", the original settings need to be returned.  Thus, remember these settings, and restore them upon canceling
Dim originalUseFancyFonts As Boolean
Dim originalAlphaCheckMode As Long
Dim originalAlphaCheckOne As Long
Dim originalAlphaCheckTwo As Long
Dim originalCanvasBackground As Long

'For this particular box, update the interface instantly
Private Sub chkFancyFonts_Click()

    useFancyFonts = CBool(chkFancyFonts)
    makeFormPretty Me
    makeFormPretty FormMain

End Sub

'Alpha channel checkerboard selection
Private Sub cmbAlphaCheck_Click()

    'Only respond to user-generated events
    If userInitiatedAlphaSelection = True Then

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
                
    End If

End Sub

'Canvas background selection
Private Sub cmbCanvas_Click()
    
    'Only respond to user-generated events
    If userInitiatedColorSelection = True Then
    
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
        
                'If a color is selected, change the picture box and associated combo box to match
                If CD1.VBChooseColor(retColor, True, True, False, Me.hWnd) Then
                    CanvasBackground = retColor
                Else
                    CanvasBackground = picCanvasColor.backColor
                End If
        End Select
    
        DrawSampleCanvasBackground
    
    End If
    
End Sub

'CANCEL button
Private Sub CmdCancel_Click()
    
    'Restore any settings that may have been changed in real-time
    If useFancyFonts <> originalUseFancyFonts Then
        useFancyFonts = originalUseFancyFonts
        makeFormPretty FormMain
    End If
    
    AlphaCheckMode = originalAlphaCheckMode
    AlphaCheckOne = originalAlphaCheckOne
    AlphaCheckTwo = originalAlphaCheckTwo
    CanvasBackground = originalCanvasBackground
    
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
    
    'We may need to access a generic "form" object multiple times, so I declare it at the top of this sub.
    Dim tForm As Form
    
    'Store whether the user wants to be prompted when closing unsaved images
    ConfirmClosingUnsaved = CBool(chkConfirmUnsaved.Value)
    userPreferences.SetPreference_Boolean "General Preferences", "ConfirmClosingUnsaved", ConfirmClosingUnsaved
    
    If ConfirmClosingUnsaved Then
        FormMain.cmdClose.ToolTip = "Close the current image." & vbCrLf & vbCrLf & "If the current image has not been saved, you will" & vbCrLf & " receive a prompt to save it before it closes."
    Else
        FormMain.cmdClose.ToolTip = "Close the current image." & vbCrLf & vbCrLf & "Because you have turned off save prompts (via Edit -> Preferences)," & vbCrLf & " you WILL NOT receive a prompt to save this image before it closes."
    End If
    
    'Store the user's preferred behavior for the "Save As" dialog's suggested file format
    userPreferences.SetPreference_Long "General Preferences", "DefaultSaveFormat", cmbDefaultSaveFormat.ListIndex
    
    'Store whether PhotoDemon is allowed to check for updates
    userPreferences.SetPreference_Boolean "General Preferences", "CheckForUpdates", CBool(chkProgramUpdates.Value)
    
    'Store whether PhotoDemon is allowed to offer the automatic download of missing core plugins
    userPreferences.SetPreference_Boolean "General Preferences", "PromptForPluginDownload", CBool(ChkPromptPluginDownload.Value)
    
    'Check to see if the new caption length setting matches the old one.  If it does not, rewrite all form captions to match
    If cmbImageCaption.ListIndex <> userPreferences.GetPreference_Long("General Preferences", "ImageCaptionSize", 0) Then
        For Each tForm In VB.Forms
            If tForm.Name = "FormImage" Then
                If cmbImageCaption.ListIndex = 0 Then
                    tForm.Caption = pdImages(tForm.Tag).OriginalFileNameAndExtension
                Else
                    If pdImages(tForm.Tag).LocationOnDisk <> "" Then tForm.Caption = pdImages(tForm.Tag).LocationOnDisk Else tForm.Caption = pdImages(tForm.Tag).OriginalFileNameAndExtension
                End If
            End If
        Next
    End If
    userPreferences.SetPreference_Long "General Preferences", "ImageCaptionSize", cmbImageCaption.ListIndex
    
    'Similarly, check to see if the new MRU caption setting matches the old one.  If it doesn't, reload the MRU.
    If cmbMRUCaption.ListIndex <> userPreferences.GetPreference_Long("General Preferences", "MRUCaptionSize", 0) Then
        userPreferences.SetPreference_Long "General Preferences", "MRUCaptionSize", cmbMRUCaption.ListIndex
        MRU_SaveToINI
        MRU_LoadFromINI
        ResetMenuIcons
    End If
        
    'Store whether PhotoDemon should validate incoming alpha channel data
    userPreferences.SetPreference_Boolean "General Preferences", "ValidateAlphaChannels", CBool(chkValidateAlpha.Value)
    
    'Store whether HDR images should be tone-mapped at load time
    userPreferences.SetPreference_Boolean "General Preferences", "UseToneMapping", CBool(chkToneMapping.Value)
    
    'Store whether we'll log system messages or not
    LogProgramMessages = CBool(ChkLogMessages.Value)
    userPreferences.SetPreference_Boolean "General Preferences", "LogProgramMessages", LogProgramMessages
    
    'Store the preference for rendering a drop shadow onto the canvas surrounding an image
    CanvasDropShadow = CBool(chkDropShadow.Value)
    userPreferences.SetPreference_Boolean "General Preferences", "CanvasDropShadow", CanvasDropShadow
    
    If CanvasDropShadow Then canvasShadow.initializeSquareShadow PD_CANVASSHADOWSIZE, PD_CANVASSHADOWSTRENGTH, CanvasBackground
    
    'Store the canvas background preference
    userPreferences.SetPreference_Long "General Preferences", "CanvasBackground", CanvasBackground
        
    'Store the alpha checkerboard preference
    userPreferences.SetPreference_Long "General Preferences", "AlphaCheckMode", CLng(cmbAlphaCheck.ListIndex)
    userPreferences.SetPreference_Long "General Preferences", "AlphaCheckOne", CLng(picAlphaOne.backColor)
    userPreferences.SetPreference_Long "General Preferences", "AlphaCheckTwo", CLng(picAlphaTwo.backColor)
    
    'Store the alpha checkerboard size preference
    AlphaCheckSize = cmbAlphaCheckSize.ListIndex
    userPreferences.SetPreference_Long "General Preferences", "AlphaCheckSize", AlphaCheckSize
    
    'Remember how the user wants multipage images to be handled
    userPreferences.SetPreference_Long "General Preferences", "MultipageImagePrompt", cmbMultiImage.ListIndex
    
    'Remember whether or not to autozoom large images
    AutosizeLargeImages = cmbLargeImages.ListIndex
    userPreferences.SetPreference_Long "General Preferences", "AutosizeLargeImages", AutosizeLargeImages
    
    'Verify the temporary path
    If LCase(TxtTempPath.Text) <> LCase(userPreferences.getTempPath) Then userPreferences.setTempPath TxtTempPath.Text
    
    'Remember the run-time only settings in the "Advanced" panel
    If imageFormats.FreeImageEnabled <> CBool(chkFreeImageTest.Value) Then
        imageFormats.FreeImageEnabled = CBool(chkFreeImageTest.Value)
        imageFormats.generateInputFormats
        imageFormats.generateOutputFormats
    End If
    If imageFormats.GDIPlusEnabled <> CBool(chkGDIPlusTest.Value) Then
        imageFormats.GDIPlusEnabled = CBool(chkGDIPlusTest.Value)
        imageFormats.generateInputFormats
        imageFormats.generateOutputFormats
    End If
    
    'Store the user's preference regarding interface fonts on modern versions of Windows
    userPreferences.SetPreference_Boolean "General Preferences", "UseFancyFonts", useFancyFonts
    
    'Store the user's preference for remembering window location
    userPreferences.SetPreference_Boolean "General Preferences", "RememberWindowLocation", CBool(chkWindowLocation.Value)
    
    'Store tool preferences
    
    'Clear selections after "Crop to Selection"
    userPreferences.SetPreference_Boolean "Tool Preferences", "ClearSelectionAfterCrop", CBool(chkSelectionClearCrop.Value)
    
    'Because some settings affect the way image canvases are rendered, redraw every active canvas
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
    tString = BrowseForFolder(Me.hWnd)
    If tString <> "" Then TxtTempPath.Text = FixPath(tString)
End Sub

'Load all relevant values from the INI file, and populate their corresponding controls with the user's current settings
Private Sub LoadAllPreferences()
    
    'Start with the canvas background (which also requires populating the canvas background combo box)
    userInitiatedColorSelection = False
    cmbCanvas.Clear
    cmbCanvas.AddItem "system theme: light", 0
    cmbCanvas.AddItem "system theme: dark", 1
    cmbCanvas.AddItem "custom color (click box to customize)", 2
        
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
    
    originalCanvasBackground = CanvasBackground
    
    'Draw the current canvas background to the sample picture box
    DrawSampleCanvasBackground
    userInitiatedColorSelection = True
    
    'Populate the combo box for default save file format
    cmbDefaultSaveFormat.Clear
    cmbDefaultSaveFormat.AddItem " the current file format of the image being saved", 0
    cmbDefaultSaveFormat.AddItem " the last image format I used in the ""Save As"" screen", 1
    cmbDefaultSaveFormat.ListIndex = userPreferences.GetPreference_Long("General Preferences", "DefaultSaveFormat", 0)
    
    'Populate the combo boxes for caption-related preferences
    cmbImageCaption.Clear
    cmbImageCaption.AddItem "compact - file name only", 0
    cmbImageCaption.AddItem "descriptive - full location, including folder(s)", 1
    cmbImageCaption.ListIndex = userPreferences.GetPreference_Long("General Preferences", "ImageCaptionSize", 0)
    
    cmbMRUCaption.Clear
    cmbMRUCaption.AddItem "compact - file names only", 0
    cmbMRUCaption.AddItem "descriptive - full locations, including folder(s)", 1
    cmbMRUCaption.ListIndex = userPreferences.GetPreference_Long("General Preferences", "MRUCaptionSize", 0)
    
    'Populate the combo box for multipage image handling
    cmbMultiImage.Clear
    cmbMultiImage.AddItem "ask me how I want to proceed", 0
    cmbMultiImage.AddItem "load only the first image", 1
    cmbMultiImage.AddItem "load all the images in the file", 2
    cmbMultiImage.ListIndex = userPreferences.GetPreference_Long("General Preferences", "MultipageImagePrompt", 0)
    
    'Next, get the values for alpha-channel checkerboard rendering
    userInitiatedAlphaSelection = False
    cmbAlphaCheck.Clear
    cmbAlphaCheck.AddItem "Highlight checks", 0
    cmbAlphaCheck.AddItem "Midtone checks", 1
    cmbAlphaCheck.AddItem "Shadow checks", 2
    cmbAlphaCheck.AddItem "Custom (click boxes to customize)", 3
    
    cmbAlphaCheck.ListIndex = AlphaCheckMode
    originalAlphaCheckMode = AlphaCheckMode
    
    picAlphaOne.backColor = AlphaCheckOne
    picAlphaTwo.backColor = AlphaCheckTwo
    originalAlphaCheckOne = AlphaCheckOne
    originalAlphaCheckTwo = AlphaCheckTwo
    
    userInitiatedAlphaSelection = True
    
    'Next, get the current alpha-channel checkerboard size value
    cmbAlphaCheckSize.Clear
    cmbAlphaCheckSize.AddItem "Small (4x4 pixels)", 0
    cmbAlphaCheckSize.AddItem "Medium (8x8 pixels)", 1
    cmbAlphaCheckSize.AddItem "Large (16x16 pixels)", 2
    
    cmbAlphaCheckSize.ListIndex = AlphaCheckSize
    
    'Assign the check box for validating incoming alpha channels on 32bpp images
    If userPreferences.GetPreference_Boolean("General Preferences", "ValidateAlphaChannels", True) Then chkValidateAlpha.Value = vbChecked Else chkValidateAlpha.Value = vbUnchecked
    
    'Assign the check box for using tone mapping on HDR images
    If userPreferences.GetPreference_Boolean("General Preferences", "UseToneMapping", True) Then chkToneMapping.Value = vbChecked Else chkToneMapping.Value = vbUnchecked
    
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
    
    'Same for remember last-used window location
    If userPreferences.GetPreference_Boolean("General Preferences", "RememberWindowLocation", True) Then chkWindowLocation.Value = vbChecked Else chkWindowLocation.Value = vbUnchecked
    
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
        originalUseFancyFonts = useFancyFonts
    End If
        
    'Populate and en/disable the run-time only settings in the "Advanced" panel
    If imageFormats.FreeImageEnabled Then
        chkFreeImageTest.Caption = "enable FreeImage support"
        chkFreeImageTest.Value = vbChecked
    Else
        chkFreeImageTest.Caption = "enable FreeImage support (do not activate this if FreeImage.dll is not available)"
        chkFreeImageTest.Value = vbUnchecked
    End If
    
    If imageFormats.GDIPlusEnabled Then
        chkGDIPlusTest.Caption = "enable GDI+ support"
        chkGDIPlusTest.Value = vbChecked
    Else
        chkGDIPlusTest.Caption = "enable GDI+ support (do not enable this if gdiplus.dll is not available)"
        chkGDIPlusTest.Value = vbUnchecked
    End If
    
    'Next, it's time for tool preferences
    
    'Clear selections after "Crop to Selection"
    If userPreferences.GetPreference_Boolean("Tool Preferences", "ClearSelectionAfterCrop", True) Then chkSelectionClearCrop.Value = vbChecked Else chkSelectionClearCrop.Value = vbUnchecked
    
    'If any preferences rely on FreeImage to operate, en/disable them as necessary
    If imageFormats.FreeImageEnabled = False Then
        chkToneMapping.Value = vbUnchecked
        chkToneMapping.Caption = "feature disabled due to missing plungin"
        chkToneMapping.Enabled = False
        cmbMultiImage.Clear
        cmbMultiImage.AddItem "feature disabled due to missing plugin", 0
        cmbMultiImage.ListIndex = 0
        cmbMultiImage.Enabled = False
        lblFreeImageWarning.Visible = True
    Else
        chkToneMapping.Enabled = True
        cmbMultiImage.Enabled = True
        lblFreeImageWarning.Visible = False
    End If

    'Finally, display some memory usage information
    lblMemoryUsageCurrent.Caption = "current PhotoDemon memory usage: " & Format(CStr(GetPhotoDemonMemoryUsage()), "###,###,###,###") & " K"
    lblMemoryUsageMax.Caption = "max PhotoDemon memory usage this session: " & Format(CStr(GetPhotoDemonMemoryUsage(True)), "###,###,###,###") & " K"
    If Not isProgramCompiled Then
        lblMemoryUsageCurrent = lblMemoryUsageCurrent.Caption & " (reading not accurate inside IDE)"
        lblMemoryUsageMax = lblMemoryUsageMax.Caption & " (reading not accurate inside IDE)"
    End If

End Sub

'When the form is loaded, populate the various checkboxes and textboxes with the values from the INI file
Private Sub Form_Load()
    
    Me.Caption = PROGRAMNAME & " Preferences"
    
    'Populate all controls with their corresponding values
    LoadAllPreferences
    
    'Populate the multi-line tooltips for the category command buttons
    'Interface
    cmdCategory(0).ToolTip = "Interface preferences include default setting for canvas backgrounds," & vbCrLf & "image load/save behavior, and the program's visual theme."
    'Loading
    cmdCategory(1).ToolTip = "Load preferences allow you to customize the way image files enter the application."
    'Saving
    cmdCategory(2).ToolTip = "Save preferences allow you to customize the way image files leave the application."
    'Performance
    cmdCategory(3).ToolTip = "Performance preferences allow you to specify how aggressively PhotoDemon makes use" & vbCrLf & "of the system's available RAM and hard drive space."
    'Tools
    cmdCategory(4).ToolTip = "Tool preferences currently include customizable options for the Selection Tool." & vbCrLf & "In the future, PhotoDemon will gain paint tools, and those settings will appear" & vbCrLf & "here as well."
    'Transparency
    cmdCategory(5).ToolTip = "Transparency preferences control how PhotoDemon displays images" & vbCrLf & "that contain alpha channels (e.g. 32bpp images)."
    'Updates
    cmdCategory(6).ToolTip = "Update preferences control how frequently PhotoDemon checks for" & vbCrLf & "updated versions, and how it handles the download of missing plugins."
    'Advanced
    cmdCategory(7).ToolTip = "Advanced preferences can be safely ignored by regular users." & vbCrLf & "Testers and developers may, however, find these settings useful."
    
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
    
    Me.picCanvasColor.Enabled = True
    Me.picCanvasColor.backColor = ConvertSystemColor(CanvasBackground)
    
End Sub

'Allow the user to change the first checkerboard color for alpha channels
Private Sub picAlphaOne_Click()
    
    Dim retColor As Long
    
    Dim CD1 As cCommonDialog
    Set CD1 = New cCommonDialog
    
    retColor = picAlphaOne.backColor
    
    'Display a Windows color selection box
    CD1.VBChooseColor retColor, True, True, False, Me.hWnd
    
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
    CD1.VBChooseColor retColor, True, True, False, Me.hWnd
    
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
    CD1.VBChooseColor retColor, True, True, False, Me.hWnd
    
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

'Test to see if we can determine folder access...
Private Sub TxtTempPath_Change()
    If Not DirectoryExist(TxtTempPath.Text) Then
        lblTempPathWarning.Visible = True
    Else
        lblTempPathWarning.Visible = False
    End If
End Sub

