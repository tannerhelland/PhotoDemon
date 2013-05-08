VERSION 5.00
Begin VB.MDIForm FormMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H80000010&
   Caption         =   "PhotoDemon by Tanner Helland - www.tannerhelland.com"
   ClientHeight    =   8745
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   15045
   Icon            =   "VBP_FormMain.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScrollBars      =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picRightPane 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8370
      Left            =   12075
      ScaleHeight     =   558
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   198
      TabIndex        =   15
      Top             =   0
      Width           =   2970
      Begin VB.CommandButton cmdTools 
         Caption         =   "Select (Rect)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   465
         Width           =   900
      End
      Begin VB.CommandButton cmdTools 
         Caption         =   "Select (Ellipse)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   1
         Left            =   1110
         TabIndex        =   25
         Top             =   465
         Width           =   900
      End
      Begin VB.PictureBox picTools 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   5535
         Index           =   0
         Left            =   0
         ScaleHeight     =   369
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   198
         TabIndex        =   16
         Top             =   2760
         Visible         =   0   'False
         Width           =   2970
         Begin PhotoDemon.sliderTextCombo sltCornerRounding 
            CausesValidation=   0   'False
            Height          =   495
            Left            =   0
            TabIndex        =   28
            Top             =   3360
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   873
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
         Begin VB.ComboBox cmbSelRender 
            Appearance      =   0  'Flat
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            ItemData        =   "VBP_FormMain.frx":000C
            Left            =   180
            List            =   "VBP_FormMain.frx":000E
            Style           =   2  'Dropdown List
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "Click to change the way selections are rendered"
            Top             =   540
            Width           =   2685
         End
         Begin PhotoDemon.textUpDown tudSelLeft 
            Height          =   405
            Index           =   0
            Left            =   120
            TabIndex        =   18
            Top             =   1440
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   714
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
         Begin PhotoDemon.textUpDown tudSelTop 
            Height          =   405
            Index           =   0
            Left            =   1560
            TabIndex        =   19
            Top             =   1440
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   714
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
         Begin PhotoDemon.textUpDown tudSelWidth 
            Height          =   405
            Index           =   0
            Left            =   120
            TabIndex        =   20
            Top             =   2400
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   714
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
         Begin PhotoDemon.textUpDown tudSelHeight 
            Height          =   405
            Index           =   0
            Left            =   1560
            TabIndex        =   21
            Top             =   2400
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   714
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
         Begin VB.Label lblSelection 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "corner rounding"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00606060&
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   29
            Top             =   3000
            Width           =   1710
         End
         Begin VB.Label lblSelection 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "visual style"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00606060&
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   24
            Top             =   120
            Width           =   1155
         End
         Begin VB.Label lblSelection 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "selection position"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00606060&
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   23
            Top             =   1050
            Width           =   1830
         End
         Begin VB.Label lblSelection 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "selection size"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00606060&
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   22
            Top             =   2010
            Width           =   1380
         End
      End
      Begin VB.Line lineMain 
         BorderColor     =   &H80000002&
         Index           =   1
         X1              =   5
         X2              =   192
         Y1              =   180
         Y2              =   180
      End
      Begin VB.Label lblTools 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "image tools"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   285
         Left            =   120
         TabIndex        =   27
         Top             =   90
         Width           =   1230
      End
   End
   Begin VB.PictureBox picProgBar 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1003
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   8370
      Width           =   15045
   End
   Begin PhotoDemon.vbalHookControl ctlAccelerator 
      Left            =   12000
      Top             =   7560
      _ExtentX        =   1191
      _ExtentY        =   1058
      Enabled         =   0   'False
   End
   Begin VB.PictureBox picLeftPane 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8370
      Left            =   0
      ScaleHeight     =   558
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   70
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   1050
      Begin VB.PictureBox picLogo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2250
         Left            =   360
         Picture         =   "VBP_FormMain.frx":0010
         ScaleHeight     =   150
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   600
         TabIndex        =   10
         Top             =   12000
         Visible         =   0   'False
         Width           =   9000
      End
      Begin VB.ComboBox CmbZoom 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "VBP_FormMain.frx":81F1
         Left            =   60
         List            =   "VBP_FormMain.frx":81F3
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Click to adjust image zoom"
         Top             =   4320
         Width           =   930
      End
      Begin PhotoDemon.jcbutton cmdOpen 
         Height          =   615
         Left            =   60
         TabIndex        =   1
         Top             =   60
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   1085
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
         BackColor       =   15199212
         Caption         =   ""
         HandPointer     =   -1  'True
         PictureNormal   =   "VBP_FormMain.frx":81F5
         DisabledPictureMode=   1
         CaptionEffects  =   0
         TooltipTitle    =   "Open"
      End
      Begin PhotoDemon.jcbutton cmdSave 
         Height          =   615
         Left            =   60
         TabIndex        =   2
         Top             =   1440
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   1085
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
         BackColor       =   15199212
         Caption         =   ""
         HandPointer     =   -1  'True
         PictureNormal   =   "VBP_FormMain.frx":9647
         DisabledPictureMode=   1
         CaptionEffects  =   0
         TooltipTitle    =   "Save"
      End
      Begin PhotoDemon.jcbutton cmdUndo 
         Height          =   615
         Left            =   60
         TabIndex        =   3
         Top             =   2820
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   1085
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
         BackColor       =   15199212
         Caption         =   ""
         HandPointer     =   -1  'True
         PictureNormal   =   "VBP_FormMain.frx":A699
         DisabledPictureMode=   1
         CaptionEffects  =   0
         TooltipTitle    =   "Undo"
         TooltipBackColor=   -2147483643
      End
      Begin PhotoDemon.jcbutton cmdRedo 
         Height          =   615
         Left            =   60
         TabIndex        =   4
         Top             =   3450
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   1085
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
         BackColor       =   15199212
         Caption         =   ""
         HandPointer     =   -1  'True
         PictureNormal   =   "VBP_FormMain.frx":B6EB
         DisabledPictureMode=   1
         CaptionEffects  =   0
         TooltipTitle    =   "Redo"
         TooltipBackColor=   -2147483643
      End
      Begin PhotoDemon.jcbutton cmdClose 
         Height          =   615
         Left            =   60
         TabIndex        =   11
         Top             =   690
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   1085
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
         BackColor       =   15199212
         Caption         =   ""
         HandPointer     =   -1  'True
         PictureNormal   =   "VBP_FormMain.frx":C73D
         DisabledPictureMode=   1
         CaptionEffects  =   0
         TooltipTitle    =   "Close"
      End
      Begin PhotoDemon.jcbutton cmdSaveAs 
         Height          =   615
         Left            =   60
         TabIndex        =   12
         Top             =   2070
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   1085
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
         BackColor       =   15199212
         Caption         =   ""
         HandPointer     =   -1  'True
         PictureNormal   =   "VBP_FormMain.frx":D78F
         DisabledPictureMode=   1
         CaptionEffects  =   0
         TooltipTitle    =   "Save As"
      End
      Begin PhotoDemon.jcbutton cmdZoomIn 
         Height          =   450
         Left            =   525
         TabIndex        =   13
         Top             =   4800
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   794
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
         BackColor       =   15199212
         Caption         =   ""
         HandPointer     =   -1  'True
         PictureNormal   =   "VBP_FormMain.frx":E7E1
         DisabledPictureMode=   1
         CaptionEffects  =   0
         ToolTip         =   "Use this button to increase image zoom."
         TooltipTitle    =   "Zoom In"
      End
      Begin PhotoDemon.jcbutton cmdZoomOut 
         Height          =   450
         Left            =   45
         TabIndex        =   14
         Top             =   4800
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   794
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
         BackColor       =   15199212
         Caption         =   ""
         HandPointer     =   -1  'True
         PictureNormal   =   "VBP_FormMain.frx":EC33
         DisabledPictureMode=   1
         CaptionEffects  =   0
         ToolTip         =   "Use this button to decrease image zoom."
         TooltipTitle    =   "Zoom Out"
      End
      Begin VB.Line lineMain 
         BorderColor     =   &H80000002&
         Index           =   0
         X1              =   2
         X2              =   68
         Y1              =   279
         Y2              =   279
      End
      Begin VB.Label lblRecording 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "macro recording in progress..."
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
         Height          =   1380
         Left            =   30
         TabIndex        =   9
         Top             =   6840
         Visible         =   0   'False
         Width           =   960
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCoordinates 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "(X, Y)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   0
         TabIndex        =   8
         Top             =   6240
         Width           =   990
      End
      Begin VB.Label lblImgSize 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "size:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D1B499&
         Height          =   675
         Left            =   0
         TabIndex        =   7
         Top             =   5460
         Width           =   990
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Menu MnuFile 
      Caption         =   "&File"
      Begin VB.Menu MnuOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu MnuRecent 
         Caption         =   "Open &recent"
         Begin VB.Menu mnuRecDocs 
            Caption         =   "Empty"
            Enabled         =   0   'False
            Index           =   0
         End
         Begin VB.Menu MnuRecentSepBar1 
            Caption         =   "-"
         End
         Begin VB.Menu MnuClearMRU 
            Caption         =   "Clear recent image list"
         End
      End
      Begin VB.Menu MnuAcquire 
         Caption         =   "&Import"
         Begin VB.Menu MnuImportClipboard 
            Caption         =   "From clipboard"
         End
         Begin VB.Menu MnuImportSepBar0 
            Caption         =   "-"
         End
         Begin VB.Menu MnuScanImage 
            Caption         =   "From scanner or camera..."
            Shortcut        =   ^I
         End
         Begin VB.Menu MnuSelectScanner 
            Caption         =   "Select which scanner or camera to use"
         End
         Begin VB.Menu MnuImportSepBar1 
            Caption         =   "-"
         End
         Begin VB.Menu MnuImportFromInternet 
            Caption         =   "Online image..."
         End
         Begin VB.Menu MnuImportSepBar2 
            Caption         =   "-"
         End
         Begin VB.Menu MnuScreenCapture 
            Caption         =   "Screen capture..."
         End
         Begin VB.Menu MnuImportSepBar3 
            Caption         =   "-"
         End
         Begin VB.Menu MnuImportFrx 
            Caption         =   "Visual Basic binary file..."
         End
      End
      Begin VB.Menu MnuFileSepBar 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu MnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu MnuSaveAs 
         Caption         =   "Save &as..."
      End
      Begin VB.Menu MnuFileSepBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuClose 
         Caption         =   "&Close"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu MnuCloseAll 
         Caption         =   "Close all"
      End
      Begin VB.Menu MnuFileSepBar2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuBatchConvert 
         Caption         =   "&Batch process..."
         Shortcut        =   ^B
      End
      Begin VB.Menu MnuFileSepBar3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPrint 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu MnuFileSepBar4 
         Caption         =   "-"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu MnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu MnuUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu MnuRedo 
         Caption         =   "&Redo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu MnuRepeatLast 
         Caption         =   "Repeat &last action"
         Enabled         =   0   'False
         Shortcut        =   ^F
      End
      Begin VB.Menu MnuEditSepBar 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCopy 
         Caption         =   "&Copy to clipboard"
         Shortcut        =   ^C
      End
      Begin VB.Menu MnuPaste 
         Caption         =   "&Paste as new image"
         Shortcut        =   ^V
      End
      Begin VB.Menu MnuEmptyClipboard 
         Caption         =   "&Empty clipboard"
      End
   End
   Begin VB.Menu MnuView 
      Caption         =   "&View"
      Begin VB.Menu MnuFitOnScreen 
         Caption         =   "&Fit image on screen"
      End
      Begin VB.Menu MnuFitWindowToImage 
         Caption         =   "Fit viewport around &image"
      End
      Begin VB.Menu MnuViewSepBar0 
         Caption         =   "-"
      End
      Begin VB.Menu MnuZoomIn 
         Caption         =   "Zoom &in"
      End
      Begin VB.Menu MnuZoomOut 
         Caption         =   "Zoom &out"
      End
      Begin VB.Menu MnuViewSepBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSpecificZoom 
         Caption         =   "16:1 (1600%)"
         Index           =   0
      End
      Begin VB.Menu MnuSpecificZoom 
         Caption         =   "8:1 (800%)"
         Index           =   1
      End
      Begin VB.Menu MnuSpecificZoom 
         Caption         =   "4:1 (400%)"
         Index           =   2
      End
      Begin VB.Menu MnuSpecificZoom 
         Caption         =   "2:1 (200%)"
         Index           =   3
      End
      Begin VB.Menu MnuSpecificZoom 
         Caption         =   "1:1 (actual size, 100%)"
         Index           =   4
      End
      Begin VB.Menu MnuSpecificZoom 
         Caption         =   "1:2 (50%)"
         Index           =   5
      End
      Begin VB.Menu MnuSpecificZoom 
         Caption         =   "1:4 (25%)"
         Index           =   6
      End
      Begin VB.Menu MnuSpecificZoom 
         Caption         =   "1:8 (12.5%)"
         Index           =   7
      End
      Begin VB.Menu MnuSpecificZoom 
         Caption         =   "1:16 (6.25%)"
         Index           =   8
      End
      Begin VB.Menu MnuViewSepBar2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuLeftPanel 
         Caption         =   "Hide left panel (file tools)"
      End
      Begin VB.Menu MnuRightPanel 
         Caption         =   "Hide right panel (image tools)"
      End
   End
   Begin VB.Menu MnuImage 
      Caption         =   "&Image"
      Begin VB.Menu MnuDuplicate 
         Caption         =   "&Duplicate"
         Shortcut        =   ^D
      End
      Begin VB.Menu MnuImageSepBar0 
         Caption         =   "-"
      End
      Begin VB.Menu MnuImageMode 
         Caption         =   "Mode"
         Begin VB.Menu MnuImageMode24bpp 
            Caption         =   "Photo  (RGB  |  24bpp  |  no transparency)"
         End
         Begin VB.Menu MnuImageMode32bpp 
            Caption         =   "Web  (RGBA  |  32bpp  |  transparency)"
         End
      End
      Begin VB.Menu MnuImageSepBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuResample 
         Caption         =   "Resize..."
         Shortcut        =   ^R
      End
      Begin VB.Menu MnuCropSelection 
         Caption         =   "Crop to selection"
      End
      Begin VB.Menu MnuAutocrop 
         Caption         =   "Autocrop Image"
      End
      Begin VB.Menu MnuImageSepBar2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuMirror 
         Caption         =   "Flip horizontal"
      End
      Begin VB.Menu MnuFlip 
         Caption         =   "Flip vertical"
      End
      Begin VB.Menu MnuImageSepBar3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRotateClockwise 
         Caption         =   "Rotate 90° clockwise"
      End
      Begin VB.Menu MnuRotate270Clockwise 
         Caption         =   "Rotate 90° counter-clockwise"
      End
      Begin VB.Menu MnuRotate180 
         Caption         =   "Rotate 180°"
      End
      Begin VB.Menu MnuRotateArbitrary 
         Caption         =   "Arbitrary rotation..."
      End
      Begin VB.Menu MnuImageSepBar4 
         Caption         =   "-"
      End
      Begin VB.Menu MnuIsometric 
         Caption         =   "Convert to isometric view"
      End
      Begin VB.Menu MnuTile 
         Caption         =   "Tile..."
      End
   End
   Begin VB.Menu MnuColorTop 
      Caption         =   "&Color"
      Begin VB.Menu MnuColor 
         Caption         =   "Brightness and contrast..."
         Index           =   0
      End
      Begin VB.Menu MnuColor 
         Caption         =   "Color balance..."
         Index           =   1
      End
      Begin VB.Menu MnuColor 
         Caption         =   "Gamma..."
         Index           =   2
         Shortcut        =   ^G
      End
      Begin VB.Menu MnuColor 
         Caption         =   "Hue and saturation..."
         Index           =   3
         Shortcut        =   ^H
      End
      Begin VB.Menu MnuColor 
         Caption         =   "Levels..."
         Index           =   4
         Shortcut        =   ^L
      End
      Begin VB.Menu MnuColor 
         Caption         =   "Shadows and highlights..."
         Index           =   5
      End
      Begin VB.Menu MnuColor 
         Caption         =   "Temperature..."
         Index           =   6
         Shortcut        =   ^T
      End
      Begin VB.Menu MnuColor 
         Caption         =   "White balance..."
         Index           =   7
         Shortcut        =   ^W
      End
      Begin VB.Menu MnuColor 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu MnuColor 
         Caption         =   "Histogram"
         Index           =   9
         Begin VB.Menu MnuHistogram 
            Caption         =   "Display histogram"
         End
         Begin VB.Menu mnuHistogramSepBar1 
            Caption         =   "-"
         End
         Begin VB.Menu MnuHistogramEqualize 
            Caption         =   "Equalize..."
         End
         Begin VB.Menu MnuHistogramStretch 
            Caption         =   "Stretch"
         End
      End
      Begin VB.Menu MnuColor 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu MnuColor 
         Caption         =   "Color shift"
         Index           =   11
         Begin VB.Menu MnuCShiftR 
            Caption         =   "Shift colors right"
         End
         Begin VB.Menu MnuCShiftL 
            Caption         =   "Shift colors left"
         End
      End
      Begin VB.Menu MnuColor 
         Caption         =   "Rechannel..."
         Index           =   12
      End
      Begin VB.Menu MnuColor 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu MnuColor 
         Caption         =   "Colorize..."
         Index           =   14
      End
      Begin VB.Menu MnuColor 
         Caption         =   "Enhance"
         Index           =   15
         Begin VB.Menu MnuAutoEnhance 
            Caption         =   "Contrast"
            Shortcut        =   +{F1}
         End
         Begin VB.Menu MnuAutoEnhanceHighlights 
            Caption         =   "Highlights"
            Shortcut        =   +{F2}
         End
         Begin VB.Menu MnuAutoEnhanceMidtones 
            Caption         =   "Midtones"
            Shortcut        =   +{F3}
         End
         Begin VB.Menu MnuAutoEnhanceShadows 
            Caption         =   "Shadows"
            Shortcut        =   +{F4}
         End
      End
      Begin VB.Menu MnuColor 
         Caption         =   "Fade"
         Index           =   16
         Begin VB.Menu MnuFadeLow 
            Caption         =   "Low fade"
         End
         Begin VB.Menu MnuFadeMedium 
            Caption         =   "Medium fade"
         End
         Begin VB.Menu MnuFadeHigh 
            Caption         =   "High fade"
         End
         Begin VB.Menu MnuCustomFade 
            Caption         =   "Custom fade..."
         End
         Begin VB.Menu MnuFadeSepBar1 
            Caption         =   "-"
         End
         Begin VB.Menu MnuUnfade 
            Caption         =   "Unfade"
         End
      End
      Begin VB.Menu MnuColor 
         Caption         =   "Grayscale..."
         Index           =   17
      End
      Begin VB.Menu MnuColor 
         Caption         =   "Invert"
         Index           =   18
         Begin VB.Menu MnuNegative 
            Caption         =   "Invert CMYK (film negative)"
         End
         Begin VB.Menu MnuInvertHue 
            Caption         =   "Invert hue"
         End
         Begin VB.Menu mnuInvert 
            Caption         =   "Invert RGB"
         End
         Begin VB.Menu mnuInvertSepBar0 
            Caption         =   "-"
         End
         Begin VB.Menu MnuCompoundInvert 
            Caption         =   "Compound invert"
         End
      End
      Begin VB.Menu MnuColor 
         Caption         =   "Monochrome"
         Index           =   19
         Begin VB.Menu MnuMonochrome 
            Caption         =   "Color to monochrome..."
            Index           =   0
         End
         Begin VB.Menu MnuMonochrome 
            Caption         =   "Monochrome to grayscale..."
            Index           =   1
         End
      End
      Begin VB.Menu MnuColor 
         Caption         =   "Sepia"
         Index           =   20
      End
      Begin VB.Menu MnuColor 
         Caption         =   "-"
         Index           =   21
      End
      Begin VB.Menu MnuColor 
         Caption         =   "Count unique colors"
         Index           =   22
      End
      Begin VB.Menu MnuColor 
         Caption         =   "Reduce color count..."
         Index           =   23
      End
   End
   Begin VB.Menu MnuFilter 
      Caption         =   "Effect&s"
      Begin VB.Menu MnuFadeLastEffect 
         Caption         =   "Fade last effect"
         Enabled         =   0   'False
      End
      Begin VB.Menu MnuFilterSepBar2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuEffectUpper 
         Caption         =   "Artistic"
         Index           =   0
         Begin VB.Menu MnuArtistic 
            Caption         =   "Antique"
            Index           =   0
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Comic book"
            Index           =   1
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Film noir"
            Index           =   2
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Modern art..."
            Index           =   3
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Pencil drawing"
            Index           =   4
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Posterize..."
            Index           =   5
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Relief"
            Index           =   6
         End
      End
      Begin VB.Menu MnuEffectUpper 
         Caption         =   "Blur"
         Index           =   1
         Begin VB.Menu MnuBlurFilter 
            Caption         =   "Soften"
            Index           =   0
         End
         Begin VB.Menu MnuBlurFilter 
            Caption         =   "Soften more"
            Index           =   1
         End
         Begin VB.Menu MnuBlurFilter 
            Caption         =   "Blur"
            Index           =   2
         End
         Begin VB.Menu MnuBlurFilter 
            Caption         =   "Blur more"
            Index           =   3
         End
         Begin VB.Menu MnuBlurFilter 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu MnuBlurFilter 
            Caption         =   "Box blur..."
            Index           =   5
         End
         Begin VB.Menu MnuBlurFilter 
            Caption         =   "Gaussian blur..."
            Index           =   6
         End
         Begin VB.Menu MnuBlurFilter 
            Caption         =   "Grid blur"
            Index           =   7
         End
         Begin VB.Menu MnuBlurFilter 
            Caption         =   "Smart blur..."
            Index           =   8
         End
         Begin VB.Menu MnuBlurFilter 
            Caption         =   "-"
            Index           =   9
         End
         Begin VB.Menu MnuBlurFilter 
            Caption         =   "Pixelate..."
            Index           =   10
         End
      End
      Begin VB.Menu MnuEffectUpper 
         Caption         =   "Distort"
         Index           =   2
         Begin VB.Menu MnuDistortFilter 
            Caption         =   "Apply lens distortion..."
            Index           =   0
         End
         Begin VB.Menu MnuDistortFilter 
            Caption         =   "Correct existing lens distortion..."
            Index           =   1
         End
         Begin VB.Menu MnuDistortFilter 
            Caption         =   "Figured glass (dents)..."
            Index           =   2
         End
         Begin VB.Menu MnuDistortFilter 
            Caption         =   "Kaleiodoscope..."
            Index           =   3
         End
         Begin VB.Menu MnuDistortFilter 
            Caption         =   "Perspective (fixed)..."
            Index           =   4
         End
         Begin VB.Menu MnuDistortFilter 
            Caption         =   "Perspective (free)..."
            Index           =   5
         End
         Begin VB.Menu MnuDistortFilter 
            Caption         =   "Pinch and whirl..."
            Index           =   6
         End
         Begin VB.Menu MnuDistortFilter 
            Caption         =   "Polar conversion..."
            Index           =   7
         End
         Begin VB.Menu MnuDistortFilter 
            Caption         =   "Ripple..."
            Index           =   8
         End
         Begin VB.Menu MnuDistortFilter 
            Caption         =   "Shear..."
            Index           =   9
         End
         Begin VB.Menu MnuDistortFilter 
            Caption         =   "Swirl..."
            Index           =   10
         End
         Begin VB.Menu MnuDistortFilter 
            Caption         =   "Waves..."
            Index           =   11
         End
      End
      Begin VB.Menu MnuEffectUpper 
         Caption         =   "Edge"
         Index           =   3
         Begin VB.Menu MnuEdge 
            Caption         =   "Emboss or engrave..."
            Index           =   0
         End
         Begin VB.Menu MnuEdge 
            Caption         =   "Enhance edges"
            Index           =   1
         End
         Begin VB.Menu MnuEdge 
            Caption         =   "Find edges..."
            Index           =   2
         End
         Begin VB.Menu MnuEdge 
            Caption         =   "Trace contour..."
            Index           =   3
         End
      End
      Begin VB.Menu MnuEffectUpper 
         Caption         =   "Experimental"
         Index           =   4
         Begin VB.Menu MnuAlien 
            Caption         =   "Alien"
         End
         Begin VB.Menu MnuBlackLight 
            Caption         =   "Black light..."
         End
         Begin VB.Menu MnuDream 
            Caption         =   "Dream"
         End
         Begin VB.Menu MnuRadioactive 
            Caption         =   "Radioactive"
         End
         Begin VB.Menu MnuSynthesize 
            Caption         =   "Synthesize"
         End
         Begin VB.Menu MnuHeatmap 
            Caption         =   "Thermograph (heat map)"
         End
         Begin VB.Menu MnuVibrate 
            Caption         =   "Vibrate"
         End
      End
      Begin VB.Menu MnuEffectUpper 
         Caption         =   "Natural"
         Index           =   5
         Begin VB.Menu MnuNatureFilter 
            Caption         =   "Atmosphere"
            Index           =   0
         End
         Begin VB.Menu MnuNatureFilter 
            Caption         =   "Burn"
            Index           =   1
         End
         Begin VB.Menu MnuNatureFilter 
            Caption         =   "Fog"
            Index           =   2
         End
         Begin VB.Menu MnuNatureFilter 
            Caption         =   "Freeze"
            Index           =   3
         End
         Begin VB.Menu MnuNatureFilter 
            Caption         =   "Lava"
            Index           =   4
         End
         Begin VB.Menu MnuNatureFilter 
            Caption         =   "Rainbow"
            Index           =   5
         End
         Begin VB.Menu MnuNatureFilter 
            Caption         =   "Steel"
            Index           =   6
         End
         Begin VB.Menu MnuNatureFilter 
            Caption         =   "Underwater"
            Index           =   7
         End
      End
      Begin VB.Menu MnuEffectUpper 
         Caption         =   "Noise"
         Index           =   6
         Begin VB.Menu MnuNoise 
            Caption         =   "Add film grain..."
            Index           =   0
         End
         Begin VB.Menu MnuNoise 
            Caption         =   "Add RGB noise..."
            Index           =   1
         End
         Begin VB.Menu MnuNoise 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu MnuNoise 
            Caption         =   "Despeckle..."
            Index           =   3
         End
         Begin VB.Menu MnuNoise 
            Caption         =   "Median..."
            Index           =   4
         End
         Begin VB.Menu MnuNoise 
            Caption         =   "Remove orphan pixels"
            Index           =   5
         End
      End
      Begin VB.Menu MnuEffectUpper 
         Caption         =   "Sharpen"
         Index           =   8
         Begin VB.Menu MnuSharpen 
            Caption         =   "Sharpen"
            Index           =   0
         End
         Begin VB.Menu MnuSharpen 
            Caption         =   "Sharpen more"
            Index           =   1
         End
         Begin VB.Menu MnuSharpen 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu MnuSharpen 
            Caption         =   "Unsharp masking..."
            Index           =   3
         End
      End
      Begin VB.Menu MnuEffectUpper 
         Caption         =   "Stylize"
         Index           =   9
         Begin VB.Menu MnuStylize 
            Caption         =   "Diffuse..."
            Index           =   0
         End
         Begin VB.Menu MnuStylize 
            Caption         =   "Dilate..."
            Index           =   1
         End
         Begin VB.Menu MnuStylize 
            Caption         =   "Erode..."
            Index           =   2
         End
         Begin VB.Menu MnuStylize 
            Caption         =   "Solarize..."
            Index           =   3
         End
         Begin VB.Menu MnuStylize 
            Caption         =   "Twins..."
            Index           =   4
         End
         Begin VB.Menu MnuStylize 
            Caption         =   "Vignetting..."
            Index           =   5
         End
      End
      Begin VB.Menu MnuFilterSepBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCustomFilter 
         Caption         =   "Custom filter..."
      End
      Begin VB.Menu MnuTest 
         Caption         =   "Test"
      End
   End
   Begin VB.Menu MnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuTool 
         Caption         =   "Language"
         Index           =   0
         Begin VB.Menu mnuLanguages 
            Caption         =   "English (US)"
            Checked         =   -1  'True
            Index           =   0
         End
      End
      Begin VB.Menu mnuTool 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuTool 
         Caption         =   "&Macros"
         Index           =   2
         Begin VB.Menu MnuPlayMacroRecording 
            Caption         =   "Play saved macro..."
         End
         Begin VB.Menu MnuMacroSepBar1 
            Caption         =   "-"
         End
         Begin VB.Menu MnuStartMacroRecording 
            Caption         =   "&Record new macro"
         End
         Begin VB.Menu MnuStopMacroRecording 
            Caption         =   "Sto&p recording..."
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuTool 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Options"
         Index           =   4
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Plugin manager"
         Index           =   5
         Shortcut        =   ^M
      End
   End
   Begin VB.Menu MnuWindowTop 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu MnuWindow 
         Caption         =   "Next image"
         Index           =   0
      End
      Begin VB.Menu MnuWindow 
         Caption         =   "Previous image"
         Index           =   1
      End
      Begin VB.Menu MnuWindow 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu MnuWindow 
         Caption         =   "&Arrange icons"
         Index           =   3
      End
      Begin VB.Menu MnuWindow 
         Caption         =   "&Cascade"
         Index           =   4
      End
      Begin VB.Menu MnuWindow 
         Caption         =   "Tile &horizontally"
         Index           =   5
      End
      Begin VB.Menu MnuWindow 
         Caption         =   "Tile &vertically"
         Index           =   6
      End
      Begin VB.Menu MnuWindow 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu MnuWindow 
         Caption         =   "&Minimize all windows"
         Index           =   8
      End
      Begin VB.Menu MnuWindow 
         Caption         =   "&Restore all windows"
         Index           =   9
      End
   End
   Begin VB.Menu MnuHelpTop 
      Caption         =   "&Help"
      Begin VB.Menu MnuHelp 
         Caption         =   "Support PhotoDemon with a small donation (thank you!)"
         Index           =   0
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "Check for &updates"
         Index           =   2
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "Submit feedback..."
         Index           =   3
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "Submit bug report..."
         Index           =   4
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "&Visit the PhotoDemon website"
         Index           =   6
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "Download PhotoDemon's source code"
         Index           =   7
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "Read PhotoDemon's license and terms of use"
         Index           =   8
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "&About PhotoDemon"
         Index           =   10
      End
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Please see the included README.txt file for additional information regarding licensing and redistribution.

'PhotoDemon is Copyright ©1999-2013 by Tanner Helland, www.tannerhelland.com

'***************************************************************************
'Main Program MDI Form
'Copyright ©2002-2013 by Tanner Helland
'Created: 15/September/02
'Last updated: 29/April/13
'Last update: rebuilt all selection code to utilize the new text/up-down custom control
'
'This is PhotoDemon's main form.  In actuality, it contains relatively little code.  Its
' primary purpose is sending parameters to other, more interesting sections of the program.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'Used to toggle the command button state of the toolbox buttons
Private Const BM_SETSTATE = &HF3
Private Declare Function SendMessageA Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'Used to render images onto the tool buttons at run-time
' NOTE: TOOLBOX IMAGES WILL NOT APPEAR IN THE IDE.  YOU MUST COMPILE FIRST.
Private cImgCtl As clsControlImage

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'When the selection type is changed, update the corresponding preference and redraw all selections
Private Sub cmbSelRender_Click(Index As Integer)
    
    'Remember the selection type, and write it out to the preferences file as well
    g_selectionRenderPreference = FormMain.cmbSelRender(Index).ListIndex
    g_UserPreferences.SetPreference_Long "Tool Preferences", "LastSelectionType", g_selectionRenderPreference
        
    If NumOfWindows > 0 Then
    
        Dim i As Long
        For i = 0 To NumOfImagesLoaded
            If (Not pdImages(i) Is Nothing) Then
                If pdImages(i).IsActive And pdImages(i).selectionActive Then RenderViewport pdImages(i).containingForm
            End If
        Next i
    
    End If
    
End Sub

'When the zoom combo box is changed, redraw the image using the new zoom value
Private Sub CmbZoom_Click()
    
    'Track the current zoom value
    If NumOfWindows > 0 Then
        pdImages(FormMain.ActiveForm.Tag).CurrentZoomValue = FormMain.CmbZoom.ListIndex
        If FormMain.CmbZoom.ListIndex = 0 Then
            FormMain.cmdZoomIn.Enabled = False
        Else
            If Not FormMain.cmdZoomIn.Enabled Then FormMain.cmdZoomIn.Enabled = True
        End If
        If FormMain.CmbZoom.ListIndex = FormMain.CmbZoom.ListCount - 1 Then
            FormMain.cmdZoomOut.Enabled = False
        Else
            If Not FormMain.cmdZoomOut.Enabled Then FormMain.cmdZoomOut.Enabled = True
        End If
        PrepareViewport FormMain.ActiveForm, "zoom changed by user"
    End If
    
End Sub

Private Sub cmdClose_Click()
    Unload Me.ActiveForm
End Sub

Private Sub cmdOpen_Click()
    Process FileOpen
End Sub

Private Sub cmdRedo_Click()
    Process Redo
End Sub

Private Sub cmdSave_Click()
    Process FileSave
End Sub

Private Sub cmdSaveAs_Click()
    Process FileSaveAs
End Sub

Private Sub cmdTools_Click(Index As Integer)
    g_CurrentTool = Index
    resetToolButtonStates
End Sub

Private Sub cmdTools_LostFocus(Index As Integer)
    g_CurrentTool = Index
    resetToolButtonStates
End Sub

Private Sub cmdTools_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    g_CurrentTool = Index
    resetToolButtonStates
End Sub

'When a new tool button is selected, we need to raise all the others and display the proper options box
Private Sub resetToolButtonStates()
    
    Dim i As Long
    For i = 0 To cmdTools.Count - 1
        If i = g_CurrentTool Then
            SendMessageA cmdTools(i).hWnd, BM_SETSTATE, True, 0
        Else
            SendMessageA cmdTools(i).hWnd, BM_SETSTATE, False, 0
        End If
    Next i
    
    'Next, we need to display the correct tool options panel.  There is no set pattern to this; some tools share
    ' panels, while enabling/disabling certain controls.  Others require completely unique ones.  I've tried to
    ' strike a balance between "as few panels as possible" without going overboard.
    Dim activeToolPanel As Long
    
    Select Case g_CurrentTool
        
        'Rectangular, Elliptical selections
        Case SELECT_RECT, SELECT_CIRC
            activeToolPanel = 0
        
        Case Else
        
    End Select
    
    'If tools share the same panel, they may need to show or hide a few additional controls.  (For example,
    ' "corner rounding", which is needed for rectangular selections but not elliptical ones, despite the two
    ' sharing the same tool panel.)  Do this before showing or hiding the tool panel.
    Select Case g_CurrentTool
    
        'For rectangular selections, show the rounded corners option
        Case SELECT_RECT
            FormMain.lblSelection(3).Visible = True
            FormMain.sltCornerRounding.Visible = True
            
        'For elliptical selections, hide the rounded corners option
        Case SELECT_CIRC
            FormMain.lblSelection(3).Visible = False
            FormMain.sltCornerRounding.Visible = False
    
    End Select
    
    'Display the current tool options panel, while hiding all inactive ones
    For i = 0 To picTools.Count - 1
        If i = activeToolPanel Then
            If Not picTools(i).Visible Then
                picTools(i).Visible = True
                setArrowCursorToObject picTools(i)
            End If
        Else
            If picTools(i).Visible Then picTools(i).Visible = False
        End If
    Next i
    
    newToolSelected
        
End Sub

'When a new tool is selected, we may need to initialize certain values
Private Sub newToolSelected()
    
    Select Case g_CurrentTool
    
        'Rectangular, elliptical selections
        Case SELECT_RECT, SELECT_CIRC
            
            'Load the visual style setting from the INI file
            FormMain.cmbSelRender(0).ListIndex = g_UserPreferences.GetPreference_Long("Tool Preferences", "LastSelectionType", 0)
        
            'If a selection is already active, change its type to match the current selection, then redraw it
            If selectionsAllowed Then
                pdImages(CurrentImage).mainSelection.setSelectionType g_CurrentTool
                If g_CurrentTool = SELECT_RECT Then pdImages(CurrentImage).mainSelection.setRoundedCornerAmount sltCornerRounding.Value
                RenderViewport FormMain.ActiveForm
            End If
            
        Case Else
        
    End Select
    
End Sub

Private Sub cmdUndo_Click()
    Process Undo
End Sub

Private Sub cmdZoomIn_Click()
    FormMain.CmbZoom.ListIndex = FormMain.CmbZoom.ListIndex - 1
End Sub

Private Sub cmdZoomOut_Click()
    FormMain.CmbZoom.ListIndex = FormMain.CmbZoom.ListIndex + 1
End Sub

'THE BEGINNING OF EVERYTHING
' Actually, Sub "Main" in the module "modMain" is loaded first, but all it does is set up native theming.  Once it has done that, FormMain is loaded.
Private Sub MDIForm_Load()

    'Use a global variable to store any command line parameters we may have been passed
    g_CommandLine = Command$
    
    'The bulk of the loading code actually takes place inside the LoadTheprogram subroutine (which can be found in the "Loading" module)
    LoadTheProgram
        
    'Hide the selection tools
    tInit tSelection, False
    
    'Render images to the toolbox command buttons
    Dim i As Long
    
    'Extract relevant icons from the resource file, and render them onto the buttons at run-time.
    ' (NOTE: because the icons require manifest theming, they will not appear in the IDE.)
    Set cImgCtl = New clsControlImage
    If g_IsProgramCompiled Then
        With cImgCtl
            .LoadImageFromStream cmdTools(0).hWnd, LoadResData("T_SELRECT", "CUSTOM"), 16, 16
            .LoadImageFromStream cmdTools(1).hWnd, LoadResData("T_SELCIRCLE", "CUSTOM"), 16, 16
            
            For i = 0 To cmdTools.Count - 1
                cmdTools(i).Caption = ""
                .SetMargins cmdTools(i).hWnd, 0
                .Align(cmdTools(i).hWnd) = Icon_Center
            Next i
        End With
    End If
    
    'Select the last-used tool (which should be saved in the INI file)
    g_CurrentTool = g_UserPreferences.GetPreference_Long("Tool Preferences", "LastActiveTool", 0)
    cmdTools_Click CInt(g_CurrentTool)
        
    'After the program has been successfully loaded, change the focus to the Open Image button
    Me.Visible = True
    If FormMain.Enabled And picLeftPane.Visible Then cmdOpen.SetFocus
        
    'Before continuing with the last few steps of interface initialization, we need to make sure the user is being presented
    ' with an interface they can understand - thus we need to evaluate the current language and make changes as necessary.
    
    'Start by asking the translation engine if it thinks we should display a language dialog.  (The conditions that trigger
    ' this are described in great detail in the pdTranslate class.)
    Dim lDialogReason As Long
    If g_Language.isLanguageDialogNeeded(lDialogReason) Then
    
        'If we are inside this block, the translation engine thinks we should ask the user to pick a language.  The reason
        ' for this is stored in the lDialogReason variable, and the values correspond to the following:
        ' 0) User-initiated dialog (irrelevant in this case; the return should never be 0)
        ' 1) First-time user, and an approximate (but not exact) language match was found.  Ask them to clarify.
        ' 2) First-time user, and no language match found.  Give them a language dialog in English.
        ' 3) Not a first-time user, but the preferred language file couldn't be located.  Ask them to pick a new one.
        
    
    End If
    
    'Start by seeing if we're allowed to check for software updates
    Dim allowedToUpdate As Boolean
    allowedToUpdate = g_UserPreferences.GetPreference_Boolean("General Preferences", "CheckForUpdates", True)
        
    'If updates ARE allowed, see when we last checked.  To be polite, only check once every 10 days.
    If allowedToUpdate Then
    
        Dim lastCheckDate As String
        lastCheckDate = g_UserPreferences.GetPreference_String("General Preferences", "LastUpdateCheck", "")
        
        'If the last update check date was not found, request an update check now
        If lastCheckDate = "" Then
        
            allowedToUpdate = True
        
        'If a last update check date was found, check to see how much time has elapsed since that check
        Else
        
            Dim currentDate As Date
            currentDate = Format$(Now, "Medium Date")
            
            'If 10 days have elapsed, allow an update check
            If CLng(DateDiff("d", CDate(lastCheckDate), currentDate)) >= 10 Then
                allowedToUpdate = True
            
            'If 10 days haven't passed, prevent an update
            Else
                Message "Update check postponed (a check has been performed in the last 10 days)"
                allowedToUpdate = False
            End If
                    
        End If
    
    End If
    
    'If we're STILL allowed to update, do so now (unless this is the first time the user has run the program; in that case, suspend updates)
    If allowedToUpdate And (Not g_IsFirstRun) Then
    
        Message "Checking for software updates (this feature can be disabled from the Tools -> Options menu)..."
    
        Dim updateNeeded As Long
        updateNeeded = CheckForSoftwareUpdate
        
        'CheckForSoftwareUpdate can return one of three values:
        ' 0 - something went wrong (no Internet connection, etc)
        ' 1 - the check was successful, but this version is up-to-date
        ' 2 - the check was successful, and an update is available
        
        Select Case updateNeeded
        
            Case 0
                Message "An error occurred while checking for updates.  Please make sure you have an active Internet connection."
            
            Case 1
                Message "Software is up-to-date."
                
                'Because the software is up-to-date, we can mark this as a successful check in the INI file
                g_UserPreferences.SetPreference_String "General Preferences", "LastUpdateCheck", Format$(Now, "Medium Date")
                
            Case 2
                Message "Software update found!  Launching update notifier..."
                FormSoftwareUpdate.Show vbModal, Me
            
        End Select
            
    End If
    
    'Last but not least, if any core plugin files were marked as "missing," offer to download them
    ' (NOTE: this check is superceded by the update check - since a full program update will include the missing plugins -
    '        so ignore this request if the user was already notified of an update.)
    If (updateNeeded <> 2) And ((Not isZLibAvailable) Or (Not isEZTwainAvailable) Or (Not isFreeImageAvailable) Or (Not isPngnqAvailable)) Then
    
        Message "Some core plugins could not be found. Preparing updater..."
        
        'As a courtesy, if the user has asked us to stop bugging them about downloading plugins, obey their request
        Dim promptToDownload As Boolean
        promptToDownload = g_UserPreferences.GetPreference_Boolean("General Preferences", "PromptForPluginDownload", True)
                
        'Finally, if allowed, we can prompt the user to download the recommended plugin set
        If promptToDownload = True Then
            FormPluginDownloader.Show vbModal, FormMain
            
            'Since plugins may have been downloaded, update the interface to match any new features that may be available.
            LoadPlugins
            ApplyAllMenuIcons
            ResetMenuIcons
            g_ImageFormats.generateInputFormats
            g_ImageFormats.generateOutputFormats
            
        Else
            Message "Ignoring plugin update request per user's INI settings"
        End If
    
    End If
    
    Message "Please load an image.  (The large 'Open Image' button at the top-left should do the trick!)"
    
    'Render the main form with any extra visual styles we've decided to apply
    RedrawMainForm
        
    'Because people may be using this code in the IDE, warn them about the consequences of doing so
    If (Not g_IsProgramCompiled) And (g_UserPreferences.GetPreference_Boolean("General Preferences", "DisplayIDEWarning", True)) Then displayIDEWarning
     
    'Finally, return focus to the main form
    'FormMain.SetFocus
     
End Sub

'Allow the user to drag-and-drop files from Windows Explorer onto the main MDI form
Private Sub MDIForm_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    'Make sure the form is available (e.g. a modal form hasn't stolen focus)
    If Not g_AllowDragAndDrop Then Exit Sub

    'Verify that the object being dragged is some sort of file or file list
    If Data.GetFormat(vbCFFiles) Then
        
        'Copy the filenames into an array
        Dim sFile() As String
        ReDim sFile(0 To Data.Files.Count) As String
        
        Dim oleFilename
        Dim tmpString As String
        
        Dim countFiles As Long
        countFiles = 0
        
        For Each oleFilename In Data.Files
            tmpString = CStr(oleFilename)
            If tmpString <> "" Then
                sFile(countFiles) = tmpString
                countFiles = countFiles + 1
            End If
        Next oleFilename
        
        'Because the OLE drop may include blank strings, verify the size of the array against countFiles
        ReDim Preserve sFile(0 To countFiles - 1) As String
        
        'Pass the list of filenames to PreLoadImage, which will load the images one-at-a-time
        PreLoadImage sFile
        
    End If
    
End Sub

Private Sub MDIForm_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

    'Make sure the form is available (e.g. a modal form hasn't stolen focus)
    If Not g_AllowDragAndDrop Then Exit Sub

    'Check to make sure the type of OLE object is files
    If Data.GetFormat(vbCFFiles) Then
        'Inform the source (Explorer, in this case) that the files will be treated as "copied"
        Effect = vbDropEffectCopy And Effect
    Else
        'If it's not files, don't allow a drop
        Effect = vbDropEffectNone
    End If

End Sub

'If the user is attempting to close the program, run some checks
' Note: in VB6, the order of events for program closing is MDI Parent QueryUnload, MDI children QueryUnload, MDI children Unload, MDI Unload
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'If the histogram form is open, close it
    Unload FormHistogram
    
    'If the user wants us to remember the program's last-used location, store those values to file now
    If g_UserPreferences.GetPreference_Boolean("General Preferences", "RememberWindowLocation", True) Then
    
        g_UserPreferences.SetPreference_Long "General Preferences", "LastWindowState", Me.WindowState
        g_UserPreferences.SetPreference_Long "General Preferences", "LastWindowLeft", Me.Left / Screen.TwipsPerPixelX
        g_UserPreferences.SetPreference_Long "General Preferences", "LastWindowTop", Me.Top / Screen.TwipsPerPixelY
        g_UserPreferences.SetPreference_Long "General Preferences", "LastWindowWidth", Me.Width / Screen.TwipsPerPixelX
        g_UserPreferences.SetPreference_Long "General Preferences", "LastWindowHeight", Me.Height / Screen.TwipsPerPixelY
    
    End If
    
    'Set a public variable to let other functions know that the user has initiated a program-wide shutdown
    g_ProgramShuttingDown = True
    
End Sub

'UNLOAD EVERYTHING
Private Sub MDIForm_Unload(Cancel As Integer)
        
    'By this point, all the child forms should have taken care of their Undo clearing-out.
    ' Just in case, however, prompt a final cleaning.
    ClearALLUndo
    
    'Release GDIPlus (if applicable)
    If g_ImageFormats.GDIPlusEnabled Then releaseGDIPlus
    
    'Destroy all custom-created form icons
    destroyAllIcons
    
    'Release the hand cursor we use for all clickable objects
    unloadAllCursors

    'Save the MRU list to the INI file.  (I've considered doing this as files are loaded, but the
    ' only time that would be an improvement is if the program crashes, and if it does crash, the user
    ' wouldn't want to re-load the problematic image anyway.)
    MRU_SaveToINI
    
    'Save all current tool options to file
    g_UserPreferences.saveToolSettings
    
    ReleaseFormTheming Me
    
End Sub

'All artistic filters are launched here
Private Sub MnuArtistic_Click(Index As Integer)

    Select Case Index
    
        'Antique
        Case 0
            Process Antique
        
        'Comic book
        Case 1
            Process ComicBook
            
        'Film noir
        Case 2
            Process FilmNoir
        
        'Modern art
        Case 3
            Process ModernArt, , , , , , , , , , True
        
        'Pencil drawing
        Case 4
            Process Pencil
                
        'Posterize
        Case 5
            Process Posterize, , , , , , , , , , True
            
        'Relief
        Case 6
            Process Relief
    
    End Select

End Sub

Private Sub MnuAutocrop_Click()
    Process Autocrop
End Sub

Private Sub MnuAutoEnhanceHighlights_Click()
    Process AutoHighlights
End Sub

Private Sub MnuAutoEnhanceMidtones_Click()
    Process AutoMidtones
End Sub

Private Sub MnuAutoEnhanceShadows_Click()
    Process AutoShadows
End Sub

Private Sub MnuBatchConvert_Click()
    g_AllowDragAndDrop = False
    FormBatchWizard.Show vbModal, FormMain
    g_AllowDragAndDrop = True
    'FormBatchConvert.Show 1, FormMain
End Sub

Private Sub MnuBlackLight_Click()
    Process BlackLight, , , , , , , , , , True
End Sub

'All blur filters are handled here
Private Sub MnuBlurFilter_Click(Index As Integer)

    Select Case Index
            
        'Soften
        Case 0
            Process Soften
        
        'Soften more
        Case 1
            Process SoftenMore
        
        'Blur
        Case 2
            Process Blur
        
        'Blur more
        Case 3
            Process BlurMore
        
        'Box blur
        Case 5
            Process BoxBlur, , , , , , , , , , True
            
        'Gaussian blur
        Case 6
            Process GaussianBlur, , , , , , , , , , True
                
        'Grid blur
        Case 7
            Process GridBlur
            
        'Smart Blur
        Case 8
            Process SmartBlur, , , , , , , , , , True
            
        'Pixelate (mosaic)
        Case 10
            Process Mosaic, , , , , , , , , , True
    
    End Select

End Sub

Private Sub MnuClearMRU_Click()
    MRU_ClearList
End Sub

Private Sub MnuClose_Click()
    
    'Note that we are not closing ALL images - just one of them
    g_ClosingAllImages = False
    Unload Me.ActiveForm
    
End Sub

Private Sub MnuCloseAll_Click()

    'Note that the user has opted to close ALL open images
    g_ClosingAllImages = True

    'Go through each image object and close the containing form
    Dim i As Long
    For i = 0 To NumOfImagesLoaded
        If (Not pdImages(i) Is Nothing) Then
            If pdImages(i).IsActive Then Unload pdImages(i).containingForm
        End If
        
        'If the user pressed "cancel", obey their request immediately and stop processing images
        If Not g_ClosingAllImages Then Exit For
        
    Next i

    'Reset the "closing all images" flag
    g_ClosingAllImages = False
    g_DealWithAllUnsavedImages = False
    
End Sub

'All Color sub-menu entries are handled here.
Private Sub MnuColor_Click(Index As Integer)

    Select Case Index
    
        'Brightness/Contrast
        Case 0
            Process BrightnessAndContrast, , , , , , , , , , True
        
        'Color balance
        Case 1
            Process AdjustColorBalance, , , , , , , , , , True
        
        'Gamma correction
        Case 2
            Process GammaCorrection, , , , , , , , , , True
            
        'HSL
        Case 3
            Process AdjustHSL, , , , , , , , , , True
            
        'Levels
        Case 4
            Process ImageLevels, , , , , , , , , , True
        
        'Shadows/Midtones/Highlights
        Case 5
            Process ShadowHighlight, , , , , , , , , , True
        
        'Temperature
        Case 6
            Process AdjustTemperature, , , , , , , , , , True
        
        'White balance
        Case 7
            Process WhiteBalance, , , , , , , , , , True
        
        '<separator>
        Case 8
        
        '<Histogram top-leve>
        Case 9
        
        '<separator>
        Case 10
        
        '<Color shift top-level>
        Case 11
        
        'Rechannel
        Case 12
            Process Rechannel, , , , , , , , , , True
        
        '<separator>
        Case 13
        
        'Colorize
        Case 14
            Process Colorize, , , , , , , , , , True
            
        '<Enhance top-level>
        Case 15
        
        '<Fade top-level>
        Case 16
        
        'Grayscale
        Case 17
            Process GrayScale, , , , , , , , , , True
            
        '<Invert top-level>
        Case 18
        
        '<Monochrome top-level>
        Case 19
            
        'Sepia
        Case 20
            Process Sepia
        
        '<separator>
        Case 21
        
        'Count colors
        Case 22
            Process CountColors
        
        'Reduce color count
        Case 23
            Process ReduceColors, , , , , , , , , , True
        
    End Select

End Sub

Private Sub MnuCompoundInvert_Click()
    Process CompoundInvert, 128
End Sub

Private Sub MnuCropSelection_Click()
    Process CropToSelection
End Sub

Private Sub MnuCShiftL_Click()
    Process ColorShiftLeft, 1
End Sub

Private Sub MnuCShiftR_Click()
    Process ColorShiftRight, 0
End Sub

Private Sub MnuCustomFade_Click()
    Process Fade, , , , , , , , , , True
End Sub

Private Sub MnuCustomFilter_Click()
    Process CustomFilter, , , , , , , , , , True
End Sub

'All distortion filters happen here
Private Sub MnuDistortFilter_Click(Index As Integer)

    Select Case Index
    
        'Apply lens distort
        Case 0
            Process DistortLens, , , , , , , , , , True
        
        'Remove lens distort
        Case 1
            Process DistortLensFix, , , , , , , , , , True
            
        'Etched glass
        Case 2
            Process DistortFiguredGlass, , , , , , , , , , True
        
        'Kaleidoscope
        Case 3
            Process DistortKaleidoscope, , , , , , , , , , True
                
        'Perspective (fixed)
        Case 4
            Process FixedPerspective, , , , , , , , , , True
            
        'Perspective (free)
        Case 5
            Process FreePerspective, , , , , , , , , , True
            
        'Pinch and whirl
        Case 6
            Process DistortPinchAndWhirl, , , , , , , , , , True
        
        'Polar conversion
        Case 7
            Process ConvertPolar, , , , , , , , , , True
        
        'Ripple
        Case 8
            Process DistortRipple, , , , , , , , , , True
        
        'Shear
        Case 9
            Process DistortShear, , , , , , , , , , True
        
        'Swirl
        Case 10
            Process DistortSwirl, , , , , , , , , , True
        
        'Waves
        Case 11
            Process DistortWaves, , , , , , , , , , True
    
    End Select

End Sub

Private Sub MnuDream_Click()
    Process Dream
End Sub

'Duplicate the current image
Private Sub MnuDuplicate_Click()
    
    'This sub can be found in the "Loading" module
    DuplicateCurrentImage
    
End Sub

Private Sub MnuEdge_Click(Index As Integer)

    Select Case Index
        
        'Emboss/engrave
        Case 0
            Process EmbossToColor, , , , , , , , , , True
        
        'Enhance edges
        Case 1
            Process EdgeEnhance
        
        'Find edges
        Case 2
            Process Laplacian, , , , , , , , , , True
        
        'Trace contour
        Case 3
            Process Contour, , , , , , , , , , True
    
    End Select

End Sub

Private Sub MnuFadeHigh_Click()
    Process Fade, 0.75
End Sub

Private Sub MnuFadeLastEffect_Click()
    Process FadeLastEffect
End Sub

Private Sub MnuFadeLow_Click()
    Process Fade, 0.25
End Sub

Private Sub MnuFadeMedium_Click()
    Process Fade, 0.5
End Sub

Private Sub MnuFitOnScreen_Click()
    FitOnScreen
End Sub

Private Sub MnuFitWindowToImage_Click()
    If (FormMain.ActiveForm.WindowState = vbMaximized) Or (FormMain.ActiveForm.WindowState = vbMinimized) Then FormMain.ActiveForm.WindowState = vbNormal
    FitWindowToImage
End Sub

Private Sub MnuHeatmap_Click()
    Process HeatMap
End Sub

'All help menu entries are launched from here
Private Sub MnuHelp_Click(Index As Integer)

    Select Case Index
        
        'Donations are so very, very welcome!
        Case 0
            OpenURL "http://www.tannerhelland.com/donate"
            
        'Check for updates
        Case 2
            Message "Checking for software updates..."
    
            Dim updateNeeded As Long
            updateNeeded = CheckForSoftwareUpdate
    
            'CheckForSoftwareUpdate can return one of three values:
            ' 0 - something went wrong (no Internet connection, etc)
            ' 1 - the check was successful, but this version is up-to-date
            ' 2 - the check was successful, and an update is available
            Select Case updateNeeded
        
                Case 0
                    Message "An error occurred while checking for updates.  Please try again later."
                    
                Case 1
                    Message "This copy of PhotoDemon is the newest available.  (Version %1.%2.%3)", App.Major, App.Minor, App.Revision
                        
                    'Because the software is up-to-date, we can mark this as a successful check in the INI file
                    g_UserPreferences.SetPreference_String "General Preferences", "LastUpdateCheck", Format$(Now, "Medium Date")
                        
                Case 2
                    Message "Software update found!  Launching update notifier..."
                    FormSoftwareUpdate.Show 1, Me
                
            End Select
        
        'Submit feedback
        Case 3
            OpenURL "http://www.tannerhelland.com/photodemon-contact/"
        
        'Submit bug report
        Case 4
            'GitHub requires a login for submitting Issues; check for that first
            Dim msgReturn As VbMsgBoxResult
            
            'If the user has previously been prompted about having a GitHub account, use their previous answer
            If g_UserPreferences.doesValueExist("General Preferences", "HasGitHubAccount") Then
            
                Dim hasGitHub As Boolean
                hasGitHub = g_UserPreferences.GetPreference_Boolean("General Preferences", "HasGitHubAccount", False)
                
                If hasGitHub Then msgReturn = vbYes Else msgReturn = vbNo
            
            'If this is the first time they are submitting feedback, ask them if they have a GitHub account
            Else
            
                msgReturn = pdMsgBox("Thank you for submitting a bug report.  To make sure your bug is addressed as quickly as possible, PhotoDemon needs to know where to send it." & vbCrLf & vbCrLf & "Do you have a GitHub account? (If you have no idea what this means, answer ""No"".)", vbQuestion + vbApplicationModal + vbYesNoCancel, "Thanks for fixing PhotoDemon")
                
                'If their answer was anything but "Cancel", store that answer to file
                If msgReturn = vbYes Then g_UserPreferences.SetPreference_Boolean "General Preferences", "HasGitHubAccount", True
                If msgReturn = vbNo Then g_UserPreferences.SetPreference_Boolean "General Preferences", "HasGitHubAccount", False
                
            End If
            
            'If they have a GitHub account, let them submit the bug there.  Otherwise, send them to the tannerhelland.com contact form
            If msgReturn = vbYes Then
                'Shell a browser window with the GitHub issue report form
                OpenURL "https://github.com/tannerhelland/PhotoDemon/issues/new"
            ElseIf msgReturn = vbNo Then
                'Shell a browser window with the tannerhelland.com PhotoDemon contact form
                OpenURL "http://www.tannerhelland.com/photodemon-contact/"
            End If
            
        'PhotoDemon's homepage
        Case 6
            OpenURL "http://www.tannerhelland.com/photodemon"
            
        'Download source code
        Case 7
            OpenURL "https://github.com/tannerhelland/PhotoDemon"
        
        'Read terms and license agreement
        Case 8
            OpenURL "http://www.tannerhelland.com/photodemon/#license"
            
        'Display About page
        Case 10
            'Before we can display the "About" form, we need to paint the PhotoDemon logo to it.
            Dim logoWidth As Long, logoHeight As Long
            Dim logoAspectRatio As Double
            
            logoWidth = FormMain.picLogo.ScaleWidth
            logoHeight = FormMain.picLogo.ScaleHeight
            logoAspectRatio = CDbl(logoWidth) / CDbl(logoHeight)
            
            FormAbout.Visible = False
            SetStretchBltMode FormAbout.hDC, STRETCHBLT_HALFTONE
            StretchBlt FormAbout.hDC, 0, 0, FormAbout.ScaleWidth, FormAbout.ScaleWidth / logoAspectRatio, FormMain.picLogo.hDC, 0, 0, logoWidth, logoHeight, vbSrcCopy
            FormAbout.Picture = FormAbout.Image
            
            'With the painting done, we can now display the form.
            FormAbout.Show vbModal, FormMain
        
    End Select

End Sub

Private Sub MnuHistogram_Click()
    Process ViewHistogram, , , , , , , , , , True
End Sub

Private Sub MnuHistogramEqualize_Click()
    Process Equalize, , , , , , , , , , True
End Sub

Private Sub MnuHistogramStretch_Click()
    Process StretchHistogram
End Sub

'Convert the current image to 24bpp mode.  (If it is already in 24bpp mode, clicking this has no effect.)
Private Sub MnuImageMode24bpp_Click()

    'Ignore clicks if the current image is in 24bpp mode
    If pdImages(CurrentImage).mainLayer.getLayerColorDepth = 24 Then Exit Sub
    
    Process ChangeImageMode24
    
End Sub

'Convert the current image to 32bpp mode.  (If it is already in 32bpp mode, clicking this has no effect.)
Private Sub MnuImageMode32bpp_Click()

    'Ignore clicks if the current image is in 32bpp mode
    If pdImages(CurrentImage).mainLayer.getLayerColorDepth = 32 Then Exit Sub
    
    Process ChangeImageMode32
    
End Sub

'This is the exact same thing as "Paste as New Image".  It is provided in two locations for convenience.
Private Sub MnuImportClipboard_Click()
    Process cPaste
End Sub

'Attempt to import an image from the Internet
Private Sub MnuImportFromInternet_Click()
    If FormInternetImport.Visible = False Then FormInternetImport.Show 1, FormMain
End Sub

Private Sub MnuImportFrx_Click()
    On Error Resume Next
    If FormImportFrx.Visible = False Then FormImportFrx.Show 1, FormMain
End Sub

Private Sub MnuAlien_Click()
    Process Alien
End Sub

Private Sub MnuInvertHue_Click()
    Process InvertHue
End Sub

Private Sub MnuIsometric_Click()
    Process Isometric
End Sub

'When a language is clicked, immediately activate it
Private Sub mnuLanguages_Click(Index As Integer)

    'Remove the existing translation
    Message "Removing existing translation..."
    g_Language.undoTranslations FormMain
    
    'Apply the new translation
    Message "Applying new translation..."
    g_Language.activateNewLanguage Index
    
    Message "Language changed successfully."

End Sub

'The user can toggle the appearance of the left-hand panel from this menu.  This toggle is also stored in the INI file.
Private Sub MnuLeftPanel_Click()
    
    ChangeLeftPane VISIBILITY_TOGGLE
    
End Sub

Private Sub MnuMonochrome_Click(Index As Integer)
    
    Select Case Index
        
        'Convert color to monochrome
        Case 0
            Process BWMaster, , , , , , , , , , True
        
        'Convert monochrome to grayscale
        Case 1
            Process RemoveBW, , , , , , , , , , True
        
    End Select
    
End Sub

Private Sub MnuNatureFilter_Click(Index As Integer)

    Select Case Index
    
        'Atmosphere
        Case 0
            Process Atmospheric
            
        'Burn
        Case 1
            Process Burn
        
        'Fog
        Case 2
            Process FogEffect
        
        'Freeze
        Case 3
            Process Frozen
        
        'Lava
        Case 4
            Process Lava
                
        'Rainbow
        Case 5
            Process Rainbow
        
        'Steel
        Case 6
            Process Steel
        
        'Water
        Case 7
            Process Water
    
    End Select

End Sub

Private Sub MnuNegative_Click()
    Process Negative
End Sub

Private Sub MnuCopy_Click()
    Process cCopy
End Sub

Private Sub MnuEmptyClipboard_Click()
    Process cEmpty
End Sub

Private Sub MnuExit_Click()
    Unload FormMain
End Sub

Private Sub MnuFlip_Click()
    Process Flip
End Sub

Private Sub MnuInvert_Click()
    Process Invert
End Sub

Private Sub MnuMirror_Click()
    Process Mirror
End Sub

'All noise filters are handled here
Private Sub MnuNoise_Click(Index As Integer)

    Select Case Index
    
        'Film grain
        Case 0
            Process FilmGrain, , , , , , , , , , True
        
        'RGB Noise
        Case 1
            Process Noise, , , , , , , , , , True
            
        'Separator
        Case 2
            
        'Despeckle
        Case 3
            Process CustomDespeckle, , , , , , , , , , True
        
        'Median
        Case 4
            Process Median, , , , , , , , , , True
            
        'Remove orphan pixels
        Case 5
            Process Despeckle
            
    End Select
        
End Sub

Private Sub MnuOpen_Click()
    Process FileOpen
End Sub

Private Sub MnuPaste_Click()
    Process cPaste
End Sub

Private Sub MnuPlayMacroRecording_Click()
    Process MacroPlayRecording
End Sub

Private Sub MnuPrint_Click()
    If FormPrint.Visible = False Then FormPrint.Show 1, FormMain
End Sub

Private Sub MnuRadioactive_Click()
    Process Radioactive
End Sub

'This is triggered whenever a user clicks on one of the "Most Recent Files" entries
Public Sub mnuRecDocs_Click(Index As Integer)
    
    'Load the MRU path that correlates to this index.  (If one is not found, a null string is returned)
    Dim tmpString As String
    tmpString = getSpecificMRU(Index)
    
    'Check - just in case - to make sure the path isn't empty
    If tmpString <> "" Then
        
        Message "Preparing to load recent file entry..."
        
        'Because PreLoadImage requires a string array, create an array to pass it
        Dim sFile(0) As String
        sFile(0) = tmpString
        
        PreLoadImage sFile
    End If
    
End Sub

Private Sub MnuRedo_Click()
    Process Redo
End Sub

Private Sub MnuRepeatLast_Click()
    Process LastCommand
End Sub

Private Sub MnuResample_Click()
    Process ImageSize, , , , , , , , , , True
End Sub

Private Sub MnuAutoEnhance_Click()
    Process AutoEnhance
End Sub

Private Sub MnuRightPanel_Click()
    ChangeRightPane VISIBILITY_TOGGLE
End Sub

Private Sub MnuRotate180_Click()
    Process Rotate180
End Sub

Private Sub MnuRotate270Clockwise_Click()
    Process Rotate270Clockwise
End Sub

Private Sub MnuRotateArbitrary_Click()
    Process FreeRotate, , , , , , , , , , True
End Sub

Private Sub MnuRotateClockwise_Click()
    Process Rotate90Clockwise
End Sub

Private Sub MnuSave_Click()
    Process FileSave
End Sub

Private Sub MnuSaveAs_Click()
    Process FileSaveAs
End Sub

Private Sub MnuScanImage_Click()
    Process ScanImage
End Sub

Private Sub MnuScreenCapture_Click()
    Process capScreen
End Sub

Private Sub MnuSelectScanner_Click()
    Process SelectScanner
End Sub

'All sharpen filters are handled here
Private Sub MnuSharpen_Click(Index As Integer)

    Select Case Index
    
        'Sharpen
        Case 0
            Process Sharpen
            
        'Sharpen More
        Case 1
            Process SharpenMore
        
        'Separator bar
        Case 2
        
        'Unsharp mask
        Case 3
            Process Unsharp, , , , , , , , , , True
            
    End Select

End Sub

'These menu items correspond to specific zoom values
Private Sub MnuSpecificZoom_Click(Index As Integer)

    Select Case Index
    
        Case 0
            If FormMain.CmbZoom.Enabled Then FormMain.CmbZoom.ListIndex = 2
        Case 1
            If FormMain.CmbZoom.Enabled Then FormMain.CmbZoom.ListIndex = 4
        Case 2
            If FormMain.CmbZoom.Enabled Then FormMain.CmbZoom.ListIndex = 8
        Case 3
            If FormMain.CmbZoom.Enabled Then FormMain.CmbZoom.ListIndex = 10
        Case 4
            If FormMain.CmbZoom.Enabled Then FormMain.CmbZoom.ListIndex = ZoomIndex100
        Case 5
            If FormMain.CmbZoom.Enabled Then FormMain.CmbZoom.ListIndex = 14
        Case 6
            If FormMain.CmbZoom.Enabled Then FormMain.CmbZoom.ListIndex = 16
        Case 7
            If FormMain.CmbZoom.Enabled Then FormMain.CmbZoom.ListIndex = 19
        Case 8
            If FormMain.CmbZoom.Enabled Then FormMain.CmbZoom.ListIndex = 21
    End Select

End Sub

Private Sub MnuStartMacroRecording_Click()
    Process MacroStartRecording
End Sub

Private Sub MnuStopMacroRecording_Click()
    Process MacroStopRecording
End Sub

'All stylize filters are handled here
Private Sub MnuStylize_Click(Index As Integer)

    Select Case Index
    
        'Diffuse
        Case 0
            Process CustomDiffuse, , , , , , , , , , True
        
        'Dilate (maximum rank)
        Case 1
        Process MaximumRank, , , , , , , , , , True
        
        'Erode (minimum rank)
        Case 2
        Process MinimumRank, , , , , , , , , , True

        'Solarize
        Case 3
            Process Solarize, , , , , , , , , , True

        'Twins
        Case 4
            Process Twins, , , , , , , , , , True
            
        'Vignetting
        Case 5
            Process Vignetting, , , , , , , , , , True
    
    End Select

End Sub

Private Sub MnuSynthesize_Click()
    Process Synthesize
End Sub

Private Sub MnuTest_Click()
    'FormShear.Show vbModal, FormMain
    'FormPerspective.Show vbModal, FormMain
    FormTruePerspective.Show vbModal, FormMain
    'MenuTest
End Sub

Private Sub MnuTile_Click()
    Process Tile, , , , , , , , , , True
End Sub

'All tool menu items are launched from here
Private Sub mnuTool_Click(Index As Integer)

    Select Case Index
    
        'Options
        Case 4
            If Not FormPreferences.Visible Then FormPreferences.Show 1, FormMain
            
        'Plugin manager
        Case 5
            If Not FormPluginManager.Visible Then FormPluginManager.Show 1, FormMain
            
    End Select

End Sub

Private Sub MnuUndo_Click()
    Process Undo
End Sub

Private Sub MnuUnfade_Click()
    Process UnFade
End Sub

Private Sub MnuVibrate_Click()
    Process Vibrate
End Sub

'Because VB doesn't allow key tracking in MDIForms, we have to hook keypresses via this method.
' Many thanks to Steve McMahon for the usercontrol that helps implement this
Private Sub ctlAccelerator_Accelerator(ByVal nIndex As Long, bCancel As Boolean)

    'Don't process accelerators when the main form is disabled (e.g. if a modal form is present)
    If FormMain.Enabled = False Then Exit Sub

    'Accelerators can be fired multiple times by accident.  Don't allow the user to press accelerators
    ' faster than one quarter-second apart.
    Static lastAccelerator As Double
    
    If Timer - lastAccelerator < 0.25 Then Exit Sub

    'Import from Internet
    If ctlAccelerator.Key(nIndex) = "Internet_Import" Then
        If FormInternetImport.Visible = False Then FormInternetImport.Show vbModal, FormMain
    End If
    
    'Save As...
    If ctlAccelerator.Key(nIndex) = "Save_As" Then
        If FormMain.MnuSaveAs.Enabled = True Then Process FileSaveAs
    End If
    
    'Capture the screen
    If ctlAccelerator.Key(nIndex) = "Screen_Capture" Then Process capScreen
    
    'Import from FRX
    If ctlAccelerator.Key(nIndex) = "Import_FRX" Then
        On Error Resume Next
        If FormImportFrx.Visible = False Then FormImportFrx.Show vbModal, FormMain
    End If

    'Open program preferences
    If ctlAccelerator.Key(nIndex) = "Preferences" Then
        If FormPreferences.Visible = False Then FormPreferences.Show vbModal, FormMain
    End If
    
    'Redo
    If ctlAccelerator.Key(nIndex) = "Redo" Then
        If FormMain.MnuRedo.Enabled = True Then Process Redo
    End If
    
    'Empty clipboard
    If ctlAccelerator.Key(nIndex) = "Empty_Clipboard" Then Process cEmpty
    
    'Fit on screen
    If ctlAccelerator.Key(nIndex) = "FitOnScreen" Then FitOnScreen
    
    'Zoom in
    If ctlAccelerator.Key(nIndex) = "Zoom_In" Then
        If FormMain.CmbZoom.Enabled = True And FormMain.CmbZoom.ListIndex > 0 Then FormMain.CmbZoom.ListIndex = FormMain.CmbZoom.ListIndex - 1
    End If
    
    'Zoom out
    If ctlAccelerator.Key(nIndex) = "Zoom_Out" Then
        If FormMain.CmbZoom.Enabled = True And FormMain.CmbZoom.ListIndex < (FormMain.CmbZoom.ListCount - 1) Then FormMain.CmbZoom.ListIndex = FormMain.CmbZoom.ListIndex + 1
    End If
    
    'Actual size
    If ctlAccelerator.Key(nIndex) = "Actual_Size" Then
        If FormMain.CmbZoom.Enabled Then FormMain.CmbZoom.ListIndex = ZoomIndex100
    End If
    
    'Various zoom values
    If ctlAccelerator.Key(nIndex) = "Zoom_161" Then
        If FormMain.CmbZoom.Enabled Then FormMain.CmbZoom.ListIndex = 2
    End If
    
    If ctlAccelerator.Key(nIndex) = "Zoom_81" Then
        If FormMain.CmbZoom.Enabled Then FormMain.CmbZoom.ListIndex = 4
    End If
    
    If ctlAccelerator.Key(nIndex) = "Zoom_41" Then
        If FormMain.CmbZoom.Enabled Then FormMain.CmbZoom.ListIndex = 8
    End If
    
    If ctlAccelerator.Key(nIndex) = "Zoom_21" Then
        If FormMain.CmbZoom.Enabled Then FormMain.CmbZoom.ListIndex = 10
    End If
    
    If ctlAccelerator.Key(nIndex) = "Zoom_12" Then
        If FormMain.CmbZoom.Enabled Then FormMain.CmbZoom.ListIndex = 14
    End If
    
    If ctlAccelerator.Key(nIndex) = "Zoom_14" Then
        If FormMain.CmbZoom.Enabled Then FormMain.CmbZoom.ListIndex = 16
    End If
    
    If ctlAccelerator.Key(nIndex) = "Zoom_18" Then
        If FormMain.CmbZoom.Enabled Then FormMain.CmbZoom.ListIndex = 19
    End If
    
    If ctlAccelerator.Key(nIndex) = "Zoom_116" Then
        If FormMain.CmbZoom.Enabled Then FormMain.CmbZoom.ListIndex = 21
    End If
    
    'Escape - right now it's only used to cancel batch conversions, but it could be applied elsewhere
    If ctlAccelerator.Key(nIndex) = "Escape" Then
        If MacroStatus = MacroBATCH Then MacroStatus = MacroCANCEL
    End If
    
    'Brightness/Contrast
    If ctlAccelerator.Key(nIndex) = "Bright_Contrast" Then
        Process BrightnessAndContrast, , , , , , , , , , True
    End If
    
    'Color balance
    If ctlAccelerator.Key(nIndex) = "Color_Balance" Then
        Process AdjustColorBalance, , , , , , , , , , True
    End If
    
    'Shadows / Highlights
    If ctlAccelerator.Key(nIndex) = "Shadow_Highlight" Then
        Process ShadowHighlight, , , , , , , , , , True
    End If
    
    'Rotate Right / Left
    If ctlAccelerator.Key(nIndex) = "Rotate_Left" Then Process Rotate270Clockwise
    If ctlAccelerator.Key(nIndex) = "Rotate_Right" Then Process Rotate90Clockwise
    
    'Crop to selection
    If ctlAccelerator.Key(nIndex) = "Crop_Selection" Then
        If pdImages(CurrentImage).selectionActive Then Process CropToSelection
    End If
    
    'Next / Previous image hotkeys ("Page Down" and "Page Up", respectively)
    If ctlAccelerator.Key(nIndex) = "Prev_Image" Or ctlAccelerator.Key(nIndex) = "Next_Image" Then
    
        'If one (or zero) images are loaded, ignore this accelerator
        If NumOfWindows <= 1 Then Exit Sub
    
        'Get the handle to the MDIClient area of FormMain; note that the "5" used is GW_CHILD per MSDN documentation
        Dim MDIClient As Long
        MDIClient = GetWindow(FormMain.hWnd, 5)
        
        'Use the API to instruct the MDI window to move one window forward or back
        If ctlAccelerator.Key(nIndex) = "Prev_Image" Then
            SendMessage MDIClient, ByVal &H224, vbNullString, ByVal 0&
        Else
            SendMessage MDIClient, ByVal &H224, vbNullString, ByVal 1&
        End If
    
    End If
    
    'MRU files
    Dim i As Integer
    For i = 0 To 9
        If ctlAccelerator.Key(nIndex) = ("MRU_" & i) Then
            If FormMain.mnuRecDocs.Count > i Then
                If FormMain.mnuRecDocs(i).Enabled = True Then
                    FormMain.mnuRecDocs_Click i
                End If
            End If
        End If
    Next i
    
    lastAccelerator = Timer
    
End Sub

'All "Window" menu items are handled here
Private Sub MnuWindow_Click(Index As Integer)

    Dim i As Long
    Dim MDIClient As Long

    Select Case Index
    
        'Next image
        Case 0
            'If one (or zero) images are loaded, ignore this option
            If NumOfWindows <= 1 Then Exit Sub
            
            'Get the handle to the MDIClient area of FormMain; note that the "5" used is GW_CHILD per MSDN documentation
            MDIClient = GetWindow(FormMain.hWnd, 5)
                
            'Use the API to instruct the MDI window to move one window forward or back
            SendMessage MDIClient, ByVal &H224, vbNullString, ByVal 1&
    
        'Previous image
        Case 1
            'If one (or zero) images are loaded, ignore this command
            If NumOfWindows <= 1 Then Exit Sub
            
            'Get the handle to the MDIClient area of FormMain; note that the "5" used is GW_CHILD per MSDN documentation
            MDIClient = GetWindow(FormMain.hWnd, 5)
                
            'Use the API to instruct the MDI window to move one window forward or back
            SendMessage MDIClient, ByVal &H224, vbNullString, ByVal 0&
    
        '<separator>
        Case 2
        
        'Arrange icons
        Case 3
            Me.Arrange vbArrangeIcons
        
        'Cascade
        Case 4
            Me.Arrange vbCascade
    
            'Rebuild the scroll bars for each window, since they will now be irrelevant (and each form's "Resize" event
            ' may not get triggered - it's a particular VB quirk)
            For i = 0 To NumOfImagesLoaded
                If (Not pdImages(i) Is Nothing) Then
                    If pdImages(i).IsActive Then PrepareViewport pdImages(i).containingForm, "Cascade"
                End If
            Next i
        
        'Tile horizontally
        Case 5
            Me.Arrange vbTileHorizontal
    
            'Rebuild the scroll bars for each window, since they will now be irrelevant (and each form's "Resize" event
            ' may not get triggered - it's a particular VB quirk)
            For i = 0 To NumOfImagesLoaded
                If (Not pdImages(i) Is Nothing) Then
                    If pdImages(i).IsActive Then PrepareViewport pdImages(i).containingForm, "Tile horizontally"
                End If
            Next i
    
        'Tile vertically
        Case 6
            Me.Arrange vbTileVertical
    
            'Rebuild the scroll bars for each window, since they will now be irrelevant (and each form's "Resize" event
            ' may not get triggered - it's a particular VB quirk)
            For i = 0 To NumOfImagesLoaded
                If (Not pdImages(i) Is Nothing) Then
                    If pdImages(i).IsActive Then PrepareViewport pdImages(i).containingForm, "Tile vertically"
                End If
            Next i
    
        '<separator>
        Case 7
        
        'Minimize all windows
        Case 8
            'Run a loop through every child form and minimize it
            Dim tForm As Form
            For Each tForm In VB.Forms
                If tForm.Name = "FormImage" Then tForm.WindowState = vbMinimized
            Next
        
        'Restore all windows
        Case 9
            'Rebuild the scroll bars for each window, since they will now be irrelevant (and each form's "Resize" event
            ' may not get triggered - it's a particular VB quirk)
            For i = 0 To NumOfImagesLoaded
                If (Not pdImages(i) Is Nothing) Then
                    If pdImages(i).IsActive Then PrepareViewport pdImages(i).containingForm, "Restore all windows"
                End If
            Next i
    
    End Select
    

End Sub

Private Sub MnuZoomIn_Click()
    If FormMain.CmbZoom.Enabled = True And FormMain.CmbZoom.ListIndex > 0 Then FormMain.CmbZoom.ListIndex = FormMain.CmbZoom.ListIndex - 1
End Sub

Private Sub MnuZoomOut_Click()
    If FormMain.CmbZoom.Enabled = True And FormMain.CmbZoom.ListIndex < (FormMain.CmbZoom.ListCount - 1) Then FormMain.CmbZoom.ListIndex = FormMain.CmbZoom.ListIndex + 1
End Sub

'When the form is resized, the left-hand bar needs to be manually redrawn.  Unfortunately, VB doesn't trigger
' the Resize() event properly for MDI parent forms, so we use the picLeftPane resize event instead.
Private Sub picLeftPane_Resize()
    
    'When this main form is resized, reapply any custom visual styles
    If FormMain.Visible Then RedrawMainForm
    
End Sub


'When the form is resized, the progress bar at bottom needs to be manually redrawn.  Unfortunately, VB doesn't trigger
' the Resize() event properly for MDI parent forms, so we use the pic_ProgBar resize event instead.
Private Sub picProgBar_Resize()
    
    'When this main form is resized, reapply any custom visual styles
    If FormMain.Visible Then RedrawMainForm
    
End Sub

Private Sub sltCornerRounding_Change()
    If selectionsAllowed Then pdImages(CurrentImage).mainSelection.setRoundedCornerAmount sltCornerRounding.Value
End Sub

'When the selection text boxes are updated, change the scrollbars to match
Private Sub tudSelHeight_Change(Index As Integer)
    updateSelectionsValuesViaText Index
End Sub

Private Sub tudSelLeft_Change(Index As Integer)
    updateSelectionsValuesViaText Index
End Sub

Private Sub tudSelTop_Change(Index As Integer)
    updateSelectionsValuesViaText Index
End Sub

Private Sub tudSelWidth_Change(Index As Integer)
    updateSelectionsValuesViaText Index
End Sub

Private Sub updateSelectionsValuesViaText(ByVal textBoxIndex As Long)
    If selectionsAllowed Then
        If Not pdImages(CurrentImage).mainSelection.rejectRefreshRequests Then pdImages(CurrentImage).mainSelection.updateViaTextBox textBoxIndex
    End If
End Sub

Private Function selectionsAllowed() As Boolean
    If NumOfWindows > 0 Then
        If pdImages(CurrentImage).selectionActive And (Not pdImages(CurrentImage).mainSelection Is Nothing) Then
            selectionsAllowed = True
        Else
            selectionsAllowed = False
        End If
    Else
        selectionsAllowed = False
    End If
End Function
