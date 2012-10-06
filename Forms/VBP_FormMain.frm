VERSION 5.00
Begin VB.MDIForm FormMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H80000010&
   Caption         =   "PhotoDemon by Tanner Helland - www.tannerhelland.com"
   ClientHeight    =   9375
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   15045
   Icon            =   "VBP_FormMain.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScrollBars      =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Top             =   9000
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
      Height          =   9000
      Left            =   0
      ScaleHeight     =   598
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   147
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   2235
      Begin VB.TextBox txtSelHeight 
         Alignment       =   2  'Center
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
         Left            =   1110
         MaxLength       =   5
         TabIndex        =   27
         Top             =   7620
         Width           =   645
      End
      Begin VB.VScrollBar vsSelHeight 
         Height          =   465
         Left            =   1740
         Min             =   1
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   7560
         Value           =   15000
         Width           =   270
      End
      Begin VB.TextBox txtSelWidth 
         Alignment       =   2  'Center
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
         Left            =   180
         MaxLength       =   5
         TabIndex        =   25
         Top             =   7620
         Width           =   645
      End
      Begin VB.VScrollBar vsSelWidth 
         Height          =   465
         Left            =   810
         Min             =   1
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   7560
         Value           =   15000
         Width           =   270
      End
      Begin VB.TextBox txtSelTop 
         Alignment       =   2  'Center
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
         Left            =   1110
         MaxLength       =   5
         TabIndex        =   22
         Top             =   6660
         Width           =   645
      End
      Begin VB.VScrollBar vsSelTop 
         Height          =   465
         Left            =   1740
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   6600
         Value           =   15000
         Width           =   270
      End
      Begin VB.TextBox txtSelLeft 
         Alignment       =   2  'Center
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
         Left            =   180
         MaxLength       =   5
         TabIndex        =   20
         Top             =   6660
         Width           =   645
      End
      Begin VB.VScrollBar vsSelLeft 
         Height          =   465
         Left            =   810
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   6600
         Value           =   15000
         Width           =   270
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
         ItemData        =   "VBP_FormMain.frx":058A
         Left            =   180
         List            =   "VBP_FormMain.frx":058C
         Style           =   2  'Dropdown List
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Click to change the way selections are rendered"
         Top             =   5730
         Width           =   1845
      End
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
         Picture         =   "VBP_FormMain.frx":058E
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
         ItemData        =   "VBP_FormMain.frx":A6C6
         Left            =   900
         List            =   "VBP_FormMain.frx":A6C8
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Click to adjust image zoom"
         Top             =   3930
         Width           =   1155
      End
      Begin PhotoDemon.jcbutton cmdOpen 
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   465
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1085
         ButtonStyle     =   13
         ShowFocusRect   =   -1  'True
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
         PictureNormal   =   "VBP_FormMain.frx":A6CA
         DisabledPictureMode=   1
         CaptionEffects  =   0
         TooltipType     =   1
         TooltipTitle    =   "Open"
      End
      Begin PhotoDemon.jcbutton cmdSave 
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   1560
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1085
         ButtonStyle     =   13
         ShowFocusRect   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         HandPointer     =   -1  'True
         PictureNormal   =   "VBP_FormMain.frx":B71C
         DisabledPictureMode=   1
         CaptionEffects  =   0
         TooltipType     =   1
         TooltipTitle    =   "Save"
      End
      Begin PhotoDemon.jcbutton cmdUndo 
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   2880
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1085
         ButtonStyle     =   13
         ShowFocusRect   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         HandPointer     =   -1  'True
         PictureNormal   =   "VBP_FormMain.frx":C76E
         DisabledPictureMode=   1
         CaptionEffects  =   0
         TooltipType     =   1
         TooltipTitle    =   "Undo"
         TooltipBackColor=   -2147483643
      End
      Begin PhotoDemon.jcbutton cmdRedo 
         Height          =   615
         Left            =   1155
         TabIndex        =   4
         Top             =   2880
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1085
         ButtonStyle     =   13
         ShowFocusRect   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         HandPointer     =   -1  'True
         PictureNormal   =   "VBP_FormMain.frx":D7C0
         DisabledPictureMode=   1
         CaptionEffects  =   0
         TooltipType     =   1
         TooltipTitle    =   "Redo"
         TooltipBackColor=   -2147483643
      End
      Begin PhotoDemon.jcbutton cmdClose 
         Height          =   615
         Left            =   1155
         TabIndex        =   12
         Top             =   465
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1085
         ButtonStyle     =   13
         ShowFocusRect   =   -1  'True
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
         PictureNormal   =   "VBP_FormMain.frx":E812
         DisabledPictureMode=   1
         CaptionEffects  =   0
         TooltipType     =   1
         TooltipTitle    =   "Close"
      End
      Begin PhotoDemon.jcbutton cmdSaveAs 
         Height          =   615
         Left            =   1155
         TabIndex        =   13
         Top             =   1560
         Width           =   900
         _ExtentX        =   1588
         _ExtentY        =   1085
         ButtonStyle     =   13
         ShowFocusRect   =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   ""
         HandPointer     =   -1  'True
         PictureNormal   =   "VBP_FormMain.frx":F864
         DisabledPictureMode=   1
         CaptionEffects  =   0
         TooltipType     =   1
         TooltipTitle    =   "Save As"
      End
      Begin VB.Label lblSelSize 
         Appearance      =   0  'Flat
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
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   7200
         Width           =   1935
      End
      Begin VB.Label lblSelPosition 
         Appearance      =   0  'Flat
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
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   6240
         Width           =   1935
      End
      Begin VB.Label lblSelStyle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "selection style"
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
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   5310
         Width           =   1695
      End
      Begin VB.Label lblZoom 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "zoom:"
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
         Left            =   135
         TabIndex        =   16
         Top             =   3945
         Width           =   675
      End
      Begin VB.Label lblUndoRedo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "undo / redo"
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
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   2505
         Width           =   1695
      End
      Begin VB.Label lblSaveSaveas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "save / save as"
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
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1185
         Width           =   1695
      End
      Begin VB.Label lblOpenClose 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "open / close"
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
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   90
         Width           =   1695
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000002&
         X1              =   5
         X2              =   142
         Y1              =   344
         Y2              =   344
      End
      Begin VB.Label lblRecording 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "macro recording in progress..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   8160
         Visible         =   0   'False
         Width           =   1935
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
         Left            =   120
         TabIndex        =   8
         Top             =   4800
         Width           =   1845
      End
      Begin VB.Label lblImgSize 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "size: WidthxHeight"
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
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   4440
         Width           =   1845
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000002&
         X1              =   5
         X2              =   142
         Y1              =   247
         Y2              =   247
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000002&
         X1              =   5
         X2              =   142
         Y1              =   158
         Y2              =   158
      End
   End
   Begin VB.Menu MnuFile 
      Caption         =   "&File"
      Begin VB.Menu MnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu MnuRecent 
         Caption         =   "Open &Recent..."
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
         Caption         =   "&Import..."
         Begin VB.Menu MnuScanImage 
            Caption         =   "From Scanner/Camera..."
            Shortcut        =   ^I
         End
         Begin VB.Menu MnuSelectScanner 
            Caption         =   "Select Scanner/Camera Source..."
         End
         Begin VB.Menu MnuImportSepBar0 
            Caption         =   "-"
         End
         Begin VB.Menu MnuImportFromInternet 
            Caption         =   "From Internet..."
         End
         Begin VB.Menu MnuImportSepBar1 
            Caption         =   "-"
         End
         Begin VB.Menu MnuScreenCapture 
            Caption         =   "Capture the Screen..."
         End
         Begin VB.Menu MnuImportSepBar2 
            Caption         =   "-"
         End
         Begin VB.Menu MnuImportFrx 
            Caption         =   "From Visual Basic Binary File..."
         End
      End
      Begin VB.Menu MnuFileSepBar2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu MnuSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu MnuFileSepBar3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuClose 
         Caption         =   "&Close"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu MnuFileSepBar5 
         Caption         =   "-"
      End
      Begin VB.Menu MnuBatchConvert 
         Caption         =   "&Batch Convert..."
         Shortcut        =   ^B
      End
      Begin VB.Menu MnuFileSepBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPrint 
         Caption         =   "&Print"
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
         Caption         =   "Repeat &Last Action"
         Enabled         =   0   'False
         Shortcut        =   ^F
      End
      Begin VB.Menu MnuEditSepBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCopy 
         Caption         =   "&Copy to Clipboard"
         Shortcut        =   ^C
      End
      Begin VB.Menu MnuPaste 
         Caption         =   "&Paste as New Image"
         Shortcut        =   ^V
      End
      Begin VB.Menu MnuEmptyClipboard 
         Caption         =   "&Empty Clipboard"
      End
      Begin VB.Menu MnuEditSepBar2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPreferences 
         Caption         =   "Program Preferences..."
      End
   End
   Begin VB.Menu MnuView 
      Caption         =   "&View"
      Begin VB.Menu MnuFitOnScreen 
         Caption         =   "&Fit Image on Screen"
      End
      Begin VB.Menu MnuFitWindowToImage 
         Caption         =   "Fit Viewport Around &Image"
      End
      Begin VB.Menu MnuViewSepBar0 
         Caption         =   "-"
      End
      Begin VB.Menu MnuZoomIn 
         Caption         =   "Zoom &In"
      End
      Begin VB.Menu MnuZoomOut 
         Caption         =   "Zoom &Out"
      End
      Begin VB.Menu MnuViewSepBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuZoom161 
         Caption         =   "16:1 (1600%)"
      End
      Begin VB.Menu MnuZoom81 
         Caption         =   "8:1 (800%)"
      End
      Begin VB.Menu MnuZoom41 
         Caption         =   "4:1 (400%)"
      End
      Begin VB.Menu MnuZoom21 
         Caption         =   "2:1 (200%)"
      End
      Begin VB.Menu MnuActualSize 
         Caption         =   "1:1 (Actual Size, 100%)"
      End
      Begin VB.Menu MnuZoom12 
         Caption         =   "1:2 (50%)"
      End
      Begin VB.Menu MnuZoom14 
         Caption         =   "1:4 (25%)"
      End
      Begin VB.Menu MnuZoom18 
         Caption         =   "1:8 (12.5%)"
      End
      Begin VB.Menu MnuZoom116 
         Caption         =   "1:16 (6.25%)"
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
      Begin VB.Menu MnuResample 
         Caption         =   "Resize..."
         Shortcut        =   ^R
      End
      Begin VB.Menu MnuCropSelection 
         Caption         =   "Crop to Selection"
      End
      Begin VB.Menu MnuImageSepBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuMirror 
         Caption         =   "Mirror (Horizontal)"
      End
      Begin VB.Menu MnuFlip 
         Caption         =   "Flip (Vertical)"
      End
      Begin VB.Menu MnuImageSepBar2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRotateClockwise 
         Caption         =   "Rotate 90° Clockwise"
      End
      Begin VB.Menu MnuRotate270Clockwise 
         Caption         =   "Rotate 90° Counter-clockwise"
      End
      Begin VB.Menu MnuRotate180 
         Caption         =   "Rotate 180°"
      End
      Begin VB.Menu MnuRotateArbitrary 
         Caption         =   "Arbitrary..."
         Visible         =   0   'False
      End
      Begin VB.Menu MnuImageSepBar3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuIsometric 
         Caption         =   "Convert to Isometric View"
      End
      Begin VB.Menu MnuTile 
         Caption         =   "Tile..."
      End
   End
   Begin VB.Menu MnuColor 
      Caption         =   "&Color"
      Begin VB.Menu MnuBrightness 
         Caption         =   "Brightness/Contrast..."
      End
      Begin VB.Menu MnuGamma 
         Caption         =   "Gamma..."
         Shortcut        =   ^G
      End
      Begin VB.Menu MnuImageLevels 
         Caption         =   "Levels..."
         Shortcut        =   ^L
      End
      Begin VB.Menu MnuTemperature 
         Caption         =   "Temperature..."
         Shortcut        =   ^T
      End
      Begin VB.Menu MnuWhiteBalance 
         Caption         =   "White Balance..."
         Shortcut        =   ^W
      End
      Begin VB.Menu MnuSepBarColor2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuHistogramTop 
         Caption         =   "Histogram"
         Begin VB.Menu MnuHistogram 
            Caption         =   "Display Histogram"
            Shortcut        =   ^H
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
      Begin VB.Menu MnuColorSepBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuColorShift 
         Caption         =   "Color Shift"
         Begin VB.Menu MnuCShiftR 
            Caption         =   "Shift Right (r -> g -> b -> r)"
         End
         Begin VB.Menu MnuCShiftL 
            Caption         =   "Shift Left (r -> b -> g -> r)"
         End
      End
      Begin VB.Menu MnuRechannel 
         Caption         =   "Rechannel..."
      End
      Begin VB.Menu MnuSepBarColor1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuBlackAndWhite 
         Caption         =   "Black and White..."
      End
      Begin VB.Menu MnuColorize 
         Caption         =   "Colorize..."
      End
      Begin VB.Menu MnuAutoEnhanceTop 
         Caption         =   "Enhance"
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
      Begin VB.Menu MnuFade 
         Caption         =   "Fade"
         Begin VB.Menu MnuFadeLow 
            Caption         =   "Low Fade"
         End
         Begin VB.Menu MnuFadeMedium 
            Caption         =   "Medium Fade"
         End
         Begin VB.Menu MnuFadeHigh 
            Caption         =   "High Fade"
         End
         Begin VB.Menu MnuCustomFade 
            Caption         =   "Custom Fade..."
         End
         Begin VB.Menu MnuFadeSepBar1 
            Caption         =   "-"
         End
         Begin VB.Menu MnuUnfade 
            Caption         =   "Unfade"
         End
      End
      Begin VB.Menu MnuGrayscale 
         Caption         =   "Grayscale..."
      End
      Begin VB.Menu MnuInvertTop 
         Caption         =   "Invert"
         Begin VB.Menu MnuNegative 
            Caption         =   "Invert CMYK (Film Negative)"
         End
         Begin VB.Menu MnuInvertHue 
            Caption         =   "Invert Hue"
         End
         Begin VB.Menu mnuInvert 
            Caption         =   "Invert RGB"
         End
         Begin VB.Menu mnuInvertSepBar0 
            Caption         =   "-"
         End
         Begin VB.Menu MnuCompoundInvert 
            Caption         =   "Compound Invert"
         End
      End
      Begin VB.Menu MnuPosterize 
         Caption         =   "Posterize..."
      End
      Begin VB.Menu MnuSepia 
         Caption         =   "Sepia"
      End
      Begin VB.Menu MnuColorSepBarPreCountColors 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCountColors 
         Caption         =   "Count unique colors"
      End
      Begin VB.Menu MnuR255 
         Caption         =   "Reduce unique colors..."
      End
   End
   Begin VB.Menu MnuFilter 
      Caption         =   "Filte&rs"
      Begin VB.Menu MnuFadeLastEffect 
         Caption         =   "Fade last effect"
         Enabled         =   0   'False
      End
      Begin VB.Menu MnuFilterSepBar2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuArtisticUpper 
         Caption         =   "Artistic"
         Begin VB.Menu MnuAntique 
            Caption         =   "Antique"
         End
         Begin VB.Menu MnuComicBook 
            Caption         =   "Comic Book"
         End
         Begin VB.Menu MnuMosaic 
            Caption         =   "Mosaic..."
         End
         Begin VB.Menu MnuPencil 
            Caption         =   "Pencil Drawing"
         End
         Begin VB.Menu MnuRelief 
            Caption         =   "Relief"
         End
      End
      Begin VB.Menu MnuBlurUpper 
         Caption         =   "Blur"
         Begin VB.Menu MnuAntialias 
            Caption         =   "Antialias"
         End
         Begin VB.Menu BlurSepBar1 
            Caption         =   "-"
         End
         Begin VB.Menu MnuSoften 
            Caption         =   "Soften"
         End
         Begin VB.Menu MnuSoftenMore 
            Caption         =   "Soften More"
         End
         Begin VB.Menu MnuBlur 
            Caption         =   "Blur"
         End
         Begin VB.Menu MnuBlurMore 
            Caption         =   "Blur More"
         End
         Begin VB.Menu MnuGaussianBlur 
            Caption         =   "Gaussian Blur"
         End
         Begin VB.Menu MnuGaussianBlurMore 
            Caption         =   "Gaussian Blur More"
         End
         Begin VB.Menu BlurSepBar2 
            Caption         =   "-"
         End
         Begin VB.Menu MnuGridBlur 
            Caption         =   "Grid Blur"
         End
      End
      Begin VB.Menu MnuDiffuseUpper 
         Caption         =   "Diffuse"
         Begin VB.Menu MnuDiffuse 
            Caption         =   "Diffuse"
         End
         Begin VB.Menu MnuDiffuseMore 
            Caption         =   "Diffuse More"
         End
         Begin VB.Menu MnuDiffuseSepBar0 
            Caption         =   "-"
         End
         Begin VB.Menu MnuCustomDiffuse 
            Caption         =   "Custom Diffuse..."
         End
      End
      Begin VB.Menu MnuEdge 
         Caption         =   "Edge"
         Begin VB.Menu MnuEmbossEngrave 
            Caption         =   "Emboss/Engrave..."
         End
         Begin VB.Menu MnuEdgeEnhance 
            Caption         =   "Enhance Edges"
         End
         Begin VB.Menu MnuFindEdges 
            Caption         =   "Find Edges..."
         End
      End
      Begin VB.Menu MnuNaturalFilters 
         Caption         =   "Natural"
         Begin VB.Menu MnuAtmosperic 
            Caption         =   "Atmosphere"
         End
         Begin VB.Menu MnuBurn 
            Caption         =   "Burn"
         End
         Begin VB.Menu MnuFogEffect 
            Caption         =   "Fog"
         End
         Begin VB.Menu MnuFrozen 
            Caption         =   "Freeze"
         End
         Begin VB.Menu MnuLava 
            Caption         =   "Lava"
         End
         Begin VB.Menu MnuOcean 
            Caption         =   "Ocean"
         End
         Begin VB.Menu MnuRainbow 
            Caption         =   "Rainbow"
         End
         Begin VB.Menu MnuSteel 
            Caption         =   "Steel"
         End
         Begin VB.Menu MnuWater 
            Caption         =   "Water"
         End
      End
      Begin VB.Menu MnuNoiseFilters 
         Caption         =   "Noise"
         Begin VB.Menu MnuNoise 
            Caption         =   "Add Noise..."
         End
         Begin VB.Menu MnuNoiseSepBar1 
            Caption         =   "-"
         End
         Begin VB.Menu MnuCustomDespeckle 
            Caption         =   "Despeckle..."
         End
         Begin VB.Menu MnuDespeckle 
            Caption         =   "Remove Orphan Pixels"
         End
      End
      Begin VB.Menu MnuOtherFilters 
         Caption         =   "Other filters"
         Begin VB.Menu MnuAlien 
            Caption         =   "Alien"
         End
         Begin VB.Menu MnuBlackLight 
            Caption         =   "Black Light..."
         End
         Begin VB.Menu MnuDream 
            Caption         =   "Dream"
         End
         Begin VB.Menu MnuRadioactive 
            Caption         =   "Radioactive"
         End
         Begin VB.Menu MnuSolarize 
            Caption         =   "Solarize..."
         End
         Begin VB.Menu MnuSynthesize 
            Caption         =   "Synthesize"
         End
         Begin VB.Menu MnuTwins 
            Caption         =   "Twins..."
         End
         Begin VB.Menu MnuVibrate 
            Caption         =   "Vibrate"
         End
      End
      Begin VB.Menu MnuRank 
         Caption         =   "Rank"
         Begin VB.Menu MnuMaximum 
            Caption         =   "Dilate (Maximum)"
         End
         Begin VB.Menu MnuMinimum 
            Caption         =   "Erode (Minimum)"
         End
         Begin VB.Menu MnuExtreme 
            Caption         =   "Extreme (Furthest value)"
         End
         Begin VB.Menu MnuRankSepBar0 
            Caption         =   "-"
         End
         Begin VB.Menu MnuCustomRank 
            Caption         =   "Custom Rank..."
         End
      End
      Begin VB.Menu MnuSharpenUpper 
         Caption         =   "Sharpen"
         Begin VB.Menu MnuUnsharp 
            Caption         =   "Remove Blur (Unsharp)"
         End
         Begin VB.Menu MnuSharpenSepBar0 
            Caption         =   "-"
         End
         Begin VB.Menu MnuSharpen 
            Caption         =   "Sharpen"
         End
         Begin VB.Menu MnuSharpenMore 
            Caption         =   "Sharpen More"
         End
      End
      Begin VB.Menu MnuFilterSepBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCustomFilter 
         Caption         =   "Custom Filter..."
      End
      Begin VB.Menu MnuTest 
         Caption         =   "Test"
      End
   End
   Begin VB.Menu MnuMacro 
      Caption         =   "&Macro"
      Begin VB.Menu MnuPlayMacroRecording 
         Caption         =   "Play Saved Macro..."
      End
      Begin VB.Menu MnuMacroSepBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuStartMacroRecording 
         Caption         =   "&Start Recording"
      End
      Begin VB.Menu MnuStopMacroRecording 
         Caption         =   "Sto&p Recording..."
      End
   End
   Begin VB.Menu MnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu MnuNextImage 
         Caption         =   "Next Image"
      End
      Begin VB.Menu MnuPreviousImage 
         Caption         =   "Previous Image"
      End
      Begin VB.Menu MnuWindowSepBar0 
         Caption         =   "-"
      End
      Begin VB.Menu MnuArrangeIcons 
         Caption         =   "&Arrange Icons"
      End
      Begin VB.Menu MnuCascadeWindows 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu MnuTileHorizontally 
         Caption         =   "Tile &Horizontally"
      End
      Begin VB.Menu MnuTileVertically 
         Caption         =   "Tile &Vertically"
      End
      Begin VB.Menu MnuWindowSepBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuMinimizeAllWindows 
         Caption         =   "&Minimize All Windows"
      End
      Begin VB.Menu MnuRestoreAllWindows 
         Caption         =   "&Restore All Windows"
      End
   End
   Begin VB.Menu MnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu MnuDonate 
         Caption         =   "Support PhotoDemon with a small donation (thank you!)"
      End
      Begin VB.Menu MnuHelpSepBar0 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCheckUpdates 
         Caption         =   "Check for &Updates..."
      End
      Begin VB.Menu MnuEmailAuthor 
         Caption         =   "Submit Feedback..."
      End
      Begin VB.Menu MnuBugReport 
         Caption         =   "Submit Bug Report..."
      End
      Begin VB.Menu MnuHelpSepBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuVisitWebsite 
         Caption         =   "&Visit the PhotoDemon Website"
      End
      Begin VB.Menu MnuDownloadSource 
         Caption         =   "Download PhotoDemon's Source Code"
      End
      Begin VB.Menu MnuReadLicense 
         Caption         =   "Read PhotoDemon's License and Terms of Use"
      End
      Begin VB.Menu MnuHelpSepBar2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAbout 
         Caption         =   "&About PhotoDemon"
      End
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Please see the included README.txt file for additional information regarding licensing and redistribution.

'PhotoDemon is Copyright ©1999-2012 by Tanner Helland, www.tannerhelland.com

'***************************************************************************
'Main Program MDI Form
'Copyright ©2000-2012 by Tanner Helland
'Created: 15/September/02
'Last updated: 30/July/12
'Last update: new accelerators added (page up, page down for previous, next image respectively)
'
'This is PhotoDemon's main form.  In actuality, it contains relatively little code.  Its
' primary purpose is sending parameters to other, more interesting sections of the program.
'
'***************************************************************************

Option Explicit

'These functions are used to scroll through consecutive MDI windows without flickering
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Any, lParam As Any) As Long
Private Declare Function GetWindow Lib "user32" (ByVal HWnd As Long, ByVal wCmd As Long) As Long

'Use to prevent scroll bar / text box combos from getting stuck in update loops
Dim suspendSelTextBoxUpdates As Boolean
Private updateSelLeftBar As Boolean, updateSelTopBar As Boolean
Private updateSelWidthBar As Boolean, updateSelHeightBar As Boolean

'When the selection type is changed, update the corresponding preference and redraw all selections
Private Sub cmbSelRender_Click()
    
    selectionRenderPreference = FormMain.cmbSelRender.ListIndex
    
    If NumOfWindows > 0 Then
    
        Dim i As Long
        For i = 1 To NumOfImagesLoaded
            If pdImages(i).IsActive And pdImages(i).selectionActive Then RenderViewport pdImages(i).containingForm
        Next i
    
    End If
    
End Sub

'When the zoom combo box is changed, redraw the image using the new zoom value
Private Sub CmbZoom_Click()
    
    'Track the current zoom value
    If NumOfWindows > 0 Then
        pdImages(FormMain.ActiveForm.Tag).CurrentZoomValue = FormMain.CmbZoom.ListIndex
        PrepareViewport FormMain.ActiveForm, "Zoom changed by user"
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

Private Sub cmdUndo_Click()
    Process Undo
End Sub

'THE BEGINNING OF EVERYTHING
'This form is actually loaded first!  Everything starts here!!
Private Sub MDIForm_Load()

    'Use a global variable to store the command-line parameters
    CommandLine = Command$
    
    'Temporarily exit from here and run the load program subroutine (Loading module)
    LoadTheProgram
    
    'After the program has been successfully loaded, change the focus to the Open Image button
    Me.Visible = True
    If FormMain.Enabled Then cmdOpen.SetFocus
    
    'If the user wants us to check for updates, now's the time to do it
    Dim tmpString As String
    
    'Start by seeing if we're allowed to check for software updates
    tmpString = GetFromIni("General Preferences", "CheckForUpdates")
    Dim allowedToUpdate As Boolean
    If Val(tmpString) = 0 Then allowedToUpdate = False Else allowedToUpdate = True
    
    'If updates ARE allowed, see when we last checked.  To be polite, only check once every 10 days.
    If allowedToUpdate = True Then
    
        Dim lastCheckDate As String
        lastCheckDate = GetFromIni("General Preferences", "LastUpdateCheck")
        
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
    
    'If we're STILL allowed to update, do so now
    If allowedToUpdate = True Then
    
        Message "Checking for software updates (this feature can be disabled from the Edit -> Preferences menu)..."
    
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
                WriteToIni "General Preferences", "LastUpdateCheck", Format$(Now, "Medium Date")
                
            Case 2
                Message "Software update found!  Launching update notifier..."
                FormSoftwareUpdate.Show 1, Me
            
        End Select
            
    End If
    
    'Last but not least, if any core plugin files were marked as "missing," offer to download them
    ' (NOTE: this check is superceded by the update check - since a full program update will include the missing plugins -
    '        so ignore this request if the user was already notified of an update.)
    If (updateNeeded = False) And ((zLibEnabled = False) Or (ScanEnabled = False) Or (FreeImageEnabled = False)) Then
    
        Message "Some core plugins could not be found. Preparing updater..."
        
        'As a courtesy, if the user has asked us to stop bugging them about downloading plugins, obey their request
        tmpString = GetFromIni("General Preferences", "PromptForPluginDownload")
        Dim promptToDownload As Boolean
        If Val(tmpString) = 0 Then promptToDownload = False Else promptToDownload = True
        
        'Finally, if allowed, we can prompt the user to download the recommended plugin set
        If promptToDownload = True Then
            FormPluginDownloader.Show 1, FormMain
        Else
            Message "Ignoring plugin update request per user's INI settings"
        End If
    
    End If
    
    Message "Please load an image.  (The large 'Open Image' button at the top-left should do the trick!)"
    
    'Render the main form with any extra visual styles we've decided to apply
    RedrawMainForm
    
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
    'Allow the selection scroll bars to be updated
    updateSelLeftBar = True
    updateSelTopBar = True
    updateSelWidthBar = True
    updateSelHeightBar = True
    
    'Hide the selection tools
    tInit tSelection, False
    
End Sub

'Allow the user to drag-and-drop files from Windows Explorer onto the main MDI form
Private Sub MDIForm_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    'Make sure the form is available (e.g. a modal form hasn't stolen focus)
    If FormMain.Enabled = False Then Exit Sub

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
    If FormMain.Enabled = False Then Exit Sub

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
    
End Sub

'UNLOAD EVERYTHING
Private Sub MDIForm_Unload(Cancel As Integer)
        
    'By this point, all the child forms should have taken care of their Undo clearing-out.
    ' Just in case, however, prompt a final cleaning.
    ClearALLUndo
    
    'Release GDIPlus (if applicable)
    If GDIPlusEnabled Then releaseGDIPlus
    
    'Release the scanner (if applicable)
    If ScanEnabled Then UnloadScanner
    
    'Destroy all custom-created form icons
    destroyAllIcons
    
    'Release the hand cursor we use for all clickable objects
    unloadAllCursors

    'Save the MRU list to the INI file.  (I've considered doing this as files are loaded, but the
    ' only time that would be an improvement is if the program crashes, and if it does crash, the user
    ' wouldn't want to re-load the problematic image anyway.)
    MRU_SaveToINI
    
End Sub

'Display the "About" form
Private Sub MnuAbout_Click()
    
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
    FormAbout.Show 1, FormMain
    
End Sub

Private Sub MnuActualSize_Click()
    If FormMain.CmbZoom.Enabled Then FormMain.CmbZoom.ListIndex = zoomIndex100
End Sub

'Private Sub MnuAnimate_Click()
'    Process Animate
'End Sub

Private Sub MnuAntialias_Click()
    Process Antialias
End Sub

Private Sub MnuAntique_Click()
    Process Antique
End Sub

Private Sub MnuArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub MnuAtmosperic_Click()
    Process Atmospheric
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
    FormBatchConvert.Show 1, FormMain
End Sub

Private Sub MnuBlackAndWhite_Click()
    Process BWImpressionist, , , , , , , , , , True
End Sub

Private Sub MnuBlackLight_Click()
    Process BlackLight, , , , , , , , , , True
End Sub

Private Sub MnuBlur_Click()
    Process Blur
End Sub

Private Sub MnuBlurMore_Click()
    Process BlurMore
End Sub

Public Sub MnuBrightness_Click()
    Process BrightnessAndContrast, , , , , , , , , , True
End Sub

Private Sub MnuBugReport_Click()
    
    'GitHub requires a login for submitting Issues; check for that first
    Dim msgReturn As VbMsgBoxResult
    
    msgReturn = MsgBox("Thank you for submitting a bug report.  To make sure your bug is addressed as quickly as possible, PhotoDemon needs to know where to send it." & vbCrLf & vbCrLf & "Do you have a GitHub account? (If you have no idea what this means, answer ""No"".)", vbQuestion + vbApplicationModal + vbYesNo, "Thanks for making " & PROGRAMNAME & " better")
    
    'If they have a GitHub account, let them submit the bug there.  Otherwise, send them to the tannerhelland.com contact form
    If msgReturn = vbYes Then
        'Shell a browser window with the GitHub issue report form
        ShellExecute FormMain.HWnd, "Open", "https://github.com/tannerhelland/PhotoDemon/issues/new", "", 0, SW_SHOWNORMAL
    Else
        'Shell a browser window with the tannerhelland.com PhotoDemon contact form
        ShellExecute FormMain.HWnd, "Open", "http://www.tannerhelland.com/photodemon-contact/", "", 0, SW_SHOWNORMAL
    End If

End Sub

Private Sub MnuBurn_Click()
    Process Burn
End Sub

Private Sub MnuCascadeWindows_Click()
    Me.Arrange vbCascade
    
    'Rebuild the scroll bars for each window, since they will now be irrelevant (and each form's "Resize" event
    ' may not get triggered - it's a particular VB quirk)
    Dim i As Long
    For i = 1 To NumOfImagesLoaded
        If pdImages(i).IsActive = True Then PrepareViewport pdImages(i).containingForm, "Cascade"
    Next i
    
End Sub

'This allows the user to manually check for updates.
Private Sub MnuCheckUpdates_Click()
        
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
            Message "This copy of PhotoDemon is the newest available.  (Version " & App.Major & "." & App.Minor & "." & App.Revision & ")"
                
            'Because the software is up-to-date, we can mark this as a successful check in the INI file
            WriteToIni "General Preferences", "LastUpdateCheck", Format$(Now, "Medium Date")
                
        Case 2
            Message "Software update found!  Launching update notifier..."
            FormSoftwareUpdate.Show 1, Me
            
    End Select
    
End Sub

Private Sub MnuClearMRU_Click()
    MRU_ClearList
End Sub

Private Sub MnuClose_Click()
    Unload Me.ActiveForm
End Sub

Private Sub MnuColorize_Click()
    Process Colorize, , , , , , , , , , True
End Sub

Private Sub MnuComicBook_Click()
    Process ComicBook
End Sub

Private Sub MnuCompoundInvert_Click()
    Process CompoundInvert, 128
End Sub

Private Sub MnuCountColors_Click()
    Process CountColors
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

Private Sub MnuCustomDespeckle_Click()
    Process CustomDespeckle, , , , , , , , , , True
End Sub

Private Sub MnuCustomDiffuse_Click()
    Process CustomDiffuse, , , , , , , , , , True
End Sub

Private Sub MnuCustomFade_Click()
    Process Fade, , , , , , , , , , True
End Sub

Private Sub MnuCustomFilter_Click()
    Process CustomFilter, , , , , , , , , , True
End Sub

Private Sub MnuCustomRank_Click()
    Process CustomRank, , , , , , , , , , True
End Sub

Private Sub MnuDespeckle_Click()
    Process Despeckle
End Sub

Private Sub MnuDiffuse_Click()
    Process Diffuse
End Sub

Private Sub MnuDiffuseMore_Click()
    Process DiffuseMore
End Sub

Private Sub MnuDonate_Click()
    'Launch the default web browser with the tannerhelland.com donation page
    ShellExecute FormMain.HWnd, "Open", "http://www.tannerhelland.com/donate", "", 0, SW_SHOWNORMAL
End Sub

Private Sub MnuDownloadSource_Click()
    'Launch the default web browser with PhotoDemon's GitHub page
    ShellExecute FormMain.HWnd, "Open", "https://github.com/tannerhelland/PhotoDemon", "", 0, SW_SHOWNORMAL
End Sub

Private Sub MnuDream_Click()
    Process Dream
End Sub

'Duplicate the current image
Private Sub MnuDuplicate_Click()
    
    'This sub can be found in the "Loading" module
    DuplicateCurrentImage
    
End Sub

Private Sub MnuEdgeEnhance_Click()
    Process EdgeEnhance
End Sub

Private Sub MnuEmailAuthor_Click()
    
    'Shell a browser window with the tannerhelland.com contact form
    ShellExecute FormMain.HWnd, "Open", "http://www.tannerhelland.com/photodemon-contact/", "", 0, SW_SHOWNORMAL

End Sub

Private Sub MnuEmbossEngrave_Click()
    Process EmbossToColor, , , , , , , , , , True
End Sub

Private Sub MnuExtreme_Click()
    Process CustomRank, 1, 2
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

Private Sub MnuFindEdges_Click()
    Process Laplacian, , , , , , , , , , True
End Sub

Private Sub MnuFitOnScreen_Click()
    FitOnScreen
End Sub

Private Sub MnuFitWindowToImage_Click()
    FitWindowToImage
End Sub

Private Sub MnuFogEffect_Click()
    Process FogEffect
End Sub

Private Sub MnuFrozen_Click()
    Process Frozen
End Sub

Private Sub MnuGamma_Click()
    Process GammaCorrection, , , , , , , , , , True
End Sub

Private Sub MnuGaussianBlur_Click()
    Process GaussianBlur
End Sub

Private Sub MnuGaussianBlurMore_Click()
    Process GaussianBlurMore
End Sub

Private Sub MnuGrayscale_Click()
    Process GrayScale, , , , , , , , , , True
End Sub

Private Sub MnuGridBlur_Click()
    Process GridBlur
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

Private Sub MnuImageLevels_Click()
    Process ImageLevels, , , , , , , , , , True
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

Private Sub MnuLava_Click()
    Process Lava
End Sub

Private Sub MnuMaximum_Click()
    Process CustomRank, 1, 0
End Sub

Private Sub MnuMinimizeAllWindows_Click()
    'Run a loop through every child form and minimize it
    Dim tForm As Form
    For Each tForm In VB.Forms
        If tForm.Name = "FormImage" Then tForm.WindowState = vbMinimized
    Next
End Sub

Private Sub MnuMinimum_Click()
    Process CustomRank, 1, 1
End Sub

Private Sub MnuMosaic_Click()
    Process Mosaic, , , , , , , , , , True
End Sub

Private Sub MnuNegative_Click()
    Process Negative
End Sub

Private Sub MnuNextImage_Click()
    
    'If one (or zero) images are loaded, ignore this option
    If NumOfWindows <= 1 Then Exit Sub
    
    'Get the handle to the MDIClient area of FormMain; note that the "5" used is GW_CHILD per MSDN documentation
    Dim MDIClient As Long
    MDIClient = GetWindow(FormMain.HWnd, 5)
        
    'Use the API to instruct the MDI window to move one window forward or back
    SendMessage MDIClient, ByVal &H224, vbNullString, ByVal 1&
    
End Sub

Private Sub MnuNoise_Click()
    Process Noise, , , , , , , , , , True
End Sub

Private Sub MnuOcean_Click()
    Process Ocean
End Sub

Private Sub MnuPencil_Click()
    Process Pencil
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

'Once I get this working (working WELL :), it will be re-enabled
Private Sub MnuFreeRotate_Click()
    Process FreeRotate, , , , , , , , , , True
End Sub

Private Sub MnuInvert_Click()
    Process Invert
End Sub

Private Sub MnuMirror_Click()
    Process Mirror
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

Private Sub MnuPosterize_Click()
    Process Posterize, , , , , , , , , , True
End Sub

Private Sub MnuPreferences_Click()
    If FormPreferences.Visible = False Then FormPreferences.Show 1, FormMain
End Sub

Private Sub MnuPreviousImage_Click()
    
    'If one (or zero) images are loaded, ignore this command
    If NumOfWindows <= 1 Then Exit Sub
    
    'Get the handle to the MDIClient area of FormMain; note that the "5" used is GW_CHILD per MSDN documentation
    Dim MDIClient As Long
    MDIClient = GetWindow(FormMain.HWnd, 5)
        
    'Use the API to instruct the MDI window to move one window forward or back
    SendMessage MDIClient, ByVal &H224, vbNullString, ByVal 0&
    
End Sub

Private Sub MnuPrint_Click()
    If FormPrint.Visible = False Then FormPrint.Show 1, FormMain
End Sub

Private Sub MnuR255_Click()
    Process ReduceColors, , , , , , , , , , True
End Sub

Private Sub MnuRadioactive_Click()
    Process Radioactive
End Sub

Private Sub MnuRainbow_Click()
    Process Rainbow
End Sub

Private Sub MnuReadLicense_Click()
    'Launch the default web browser with PhotoDemon's license page on tannerhelland.com
    ShellExecute FormMain.HWnd, "Open", "http://www.tannerhelland.com/photodemon/#license", "", 0, SW_SHOWNORMAL
End Sub

'This is triggered whenever a user clicks on one of the "Most Recent Files" entries
Public Sub mnuRecDocs_Click(Index As Integer)
    
    'Check - just in case - to make sure the path isn't empty
    If mnuRecDocs(Index).Caption <> "" Then
        Message "Preparing to load MRU entry..."
        
        'Strip the accelerator caption from this recent menu entry
        Dim tmpString As String
        tmpString = mnuRecDocs(Index).Caption
        StripAcceleratorFromCaption tmpString
        
        'Because PreLoadImage requires a string array, create an array to pass it
        Dim sFile(0) As String
        sFile(0) = tmpString
        
        PreLoadImage sFile
    End If
    
End Sub

Private Sub MnuRechannel_Click()
    Process Rechannel, , , , , , , , , , True
End Sub

Private Sub MnuRedo_Click()
    Process Redo
End Sub

Private Sub MnuRepeatLast_Click()
    Process LastCommand
End Sub

Private Sub MnuRelief_Click()
    Process Relief
End Sub

Private Sub MnuResample_Click()
    Process ImageSize, , , , , , , , , , True
End Sub

Private Sub MnuAutoEnhance_Click()
    Process AutoEnhance
End Sub

Private Sub MnuRestoreAllWindows_Click()
    'Run a loop through every child form and un-minimize it
    Dim tForm As Form
    For Each tForm In VB.Forms
        If tForm.Name = "FormImage" Then
            tForm.WindowState = vbNormal
            'Rebuild the scroll bars for each window, since they will now be irrelevant (and each form's "Resize" event
            ' may not get triggered - VB is quirky about triggering that event reliably)
            If pdImages(tForm.Tag).IsActive = True Then PrepareViewport tForm, "Restore all windows"
        End If
    Next
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

Private Sub MnuSepia_Click()
    Process Sepia
End Sub

Private Sub MnuSharpen_Click()
    Process Sharpen
End Sub

Private Sub MnuSharpenMore_Click()
    Process SharpenMore
End Sub

Private Sub MnuSoften_Click()
    Process Soften
End Sub

Private Sub MnuSoftenMore_Click()
    Process SoftenMore
End Sub

Private Sub MnuSolarize_Click()
    Process Solarize, , , , , , , , , , True
End Sub

Private Sub MnuStartMacroRecording_Click()
    Process MacroStartRecording
End Sub

Private Sub MnuSteel_Click()
    Process Steel
End Sub

Private Sub MnuStopMacroRecording_Click()
    Process MacroStopRecording
End Sub

Private Sub MnuSynthesize_Click()
    Process Synthesize
End Sub

Private Sub MnuTemperature_Click()
    Process AdjustTemperature, , , , , , , , , , True
End Sub

Private Sub MnuTest_Click()
    MenuTest
End Sub

Private Sub MnuTile_Click()
    Process Tile, , , , , , , , , , True
End Sub

Private Sub MnuTileHorizontally_Click()
    Me.Arrange vbTileHorizontal
    
    'Rebuild the scroll bars for each window, since they will now be irrelevant (and each form's "Resize" event
    ' may not get triggered - it's a particular VB quirk)
    Dim i As Long
    For i = 1 To NumOfImagesLoaded
        If pdImages(i).IsActive = True Then PrepareViewport pdImages(i).containingForm, "Tile horizontally"
    Next i
    
End Sub

Private Sub MnuTileVertically_Click()
    Me.Arrange vbTileVertical
    
    'Rebuild the scroll bars for each window, since they will now be irrelevant (and each form's "Resize" event
    ' may not get triggered - it's a particular VB quirk)
    Dim i As Long
    For i = 1 To NumOfImagesLoaded
        If pdImages(i).IsActive = True Then PrepareViewport pdImages(i).containingForm, "Tile vertically"
    Next i
    
End Sub

Private Sub MnuTwins_Click()
    Process Twins, , , , , , , , , , True
End Sub

Private Sub MnuUndo_Click()
    Process Undo
End Sub

Private Sub MnuUnfade_Click()
    Process Unfade
End Sub

Private Sub MnuUnsharp_Click()
    Process Unsharp
End Sub

Private Sub MnuVibrate_Click()
    Process Vibrate
End Sub

Private Sub MnuVisitWebsite_Click()
    'Nothing special here - just launch the default web browser with PhotoDemon's page on tannerhelland.com
    ShellExecute FormMain.HWnd, "Open", "http://www.tannerhelland.com/photodemon", "", 0, SW_SHOWNORMAL
End Sub

Private Sub MnuWater_Click()
    Process Water
End Sub

Private Sub MnuWhiteBalance_Click()
    Process WhiteBalance, , , , , , , , , , True
End Sub

'Because VB doesn't allow key tracking in MDIForms, we have to hook keypresses via this method.
' Many thanks to Steve McMahon for the usercontrol that helps implement this
Private Sub ctlAccelerator_Accelerator(ByVal nIndex As Long, bCancel As Boolean)

    'Don't process accelerators when the main form is disabled (e.g. if a modal form is present)
    If FormMain.Enabled = False Then Exit Sub

    'Accelerators can be fired multiple times by accident.  Don't allow the user to press accelerators
    ' faster than one quarter-second apart.
    Static lastAccelerator As Single
    
    If Timer - lastAccelerator < 0.25 Then Exit Sub

    'Import from Internet
    If ctlAccelerator.Key(nIndex) = "Internet_Import" Then
        If FormInternetImport.Visible = False Then FormInternetImport.Show 1, FormMain
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
        If FormImportFrx.Visible = False Then FormImportFrx.Show 1, FormMain
    End If

    'Open program preferences
    If ctlAccelerator.Key(nIndex) = "Preferences" Then
        If FormPreferences.Visible = False Then FormPreferences.Show 1, FormMain
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
        If FormMain.CmbZoom.Enabled Then FormMain.CmbZoom.ListIndex = zoomIndex100
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
    
    'Rotate Right / Left
    If ctlAccelerator.Key(nIndex) = "Rotate_Left" Then Process Rotate270Clockwise
    If ctlAccelerator.Key(nIndex) = "Rotate_Right" Then Process Rotate90Clockwise
    
    'Crop to selection
    If pdImages(CurrentImage).selectionActive Then Process CropToSelection
    
    'Next / Previous image hotkeys ("Page Down" and "Page Up", respectively)
    If ctlAccelerator.Key(nIndex) = "Prev_Image" Or ctlAccelerator.Key(nIndex) = "Next_Image" Then
    
        'If one (or zero) images are loaded, ignore this accelerator
        If NumOfWindows <= 1 Then Exit Sub
    
        'Get the handle to the MDIClient area of FormMain; note that the "5" used is GW_CHILD per MSDN documentation
        Dim MDIClient As Long
        MDIClient = GetWindow(FormMain.HWnd, 5)
        
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

Private Sub MnuZoom116_Click()
    If FormMain.CmbZoom.Enabled Then FormMain.CmbZoom.ListIndex = 21
End Sub

Private Sub MnuZoom12_Click()
    If FormMain.CmbZoom.Enabled Then FormMain.CmbZoom.ListIndex = 14
End Sub

Private Sub MnuZoom14_Click()
    If FormMain.CmbZoom.Enabled Then FormMain.CmbZoom.ListIndex = 16
End Sub

Private Sub MnuZoom161_Click()
    If FormMain.CmbZoom.Enabled Then FormMain.CmbZoom.ListIndex = 2
End Sub

Private Sub MnuZoom18_Click()
    If FormMain.CmbZoom.Enabled Then FormMain.CmbZoom.ListIndex = 19
End Sub

Private Sub MnuZoom21_Click()
    If FormMain.CmbZoom.Enabled Then FormMain.CmbZoom.ListIndex = 10
End Sub

Private Sub MnuZoom41_Click()
    If FormMain.CmbZoom.Enabled Then FormMain.CmbZoom.ListIndex = 8
End Sub

Private Sub MnuZoom81_Click()
    If FormMain.CmbZoom.Enabled Then FormMain.CmbZoom.ListIndex = 4
End Sub

Private Sub MnuZoomIn_Click()
    If FormMain.CmbZoom.Enabled = True And FormMain.CmbZoom.ListIndex > 0 Then FormMain.CmbZoom.ListIndex = FormMain.CmbZoom.ListIndex - 1
End Sub

Private Sub MnuZoomOut_Click()
    If FormMain.CmbZoom.Enabled = True And FormMain.CmbZoom.ListIndex < (FormMain.CmbZoom.ListCount - 1) Then FormMain.CmbZoom.ListIndex = FormMain.CmbZoom.ListIndex + 1
End Sub

'When the form is resized, the progress bar at bottom needs to be manually redrawn.  Unfortunately, VB doesn't trigger
' the Resize() event properly for MDI parent forms, so we use the picProgBar resize event instead.
Private Sub picProgBar_Resize()
    
    'When this main form is resized, reapply any custom visual styles
    RedrawMainForm
    
End Sub

Private Sub txtSelHeight_GotFocus()
    AutoSelectText txtSelHeight
End Sub

'When the selection text boxes are updated, change the scrollbars to match
Private Sub txtSelHeight_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtSelHeight
    changeToSelHeight
End Sub

Private Sub txtSelLeft_GotFocus()
    AutoSelectText txtSelLeft
End Sub

Private Sub txtSelLeft_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtSelLeft
    changeToSelLeft
End Sub

Private Sub txtSelTop_GotFocus()
    AutoSelectText txtSelTop
End Sub

Private Sub txtSelTop_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtSelTop
    changeToSelTop
End Sub

Private Sub txtSelWidth_GotFocus()
    AutoSelectText txtSelWidth
End Sub

Private Sub txtSelWidth_KeyUp(KeyCode As Integer, Shift As Integer)
    textValidate txtSelWidth
    changeToSelWidth
End Sub

'When the selection scroll bars are updated, change the text boxes to match
Private Sub vsSelLeft_Change()
    If updateSelLeftBar = True Then
        txtSelLeft = Abs(32767 - CStr(vsSelLeft.Value))
        txtSelLeft.Refresh
        changeToSelLeft
    End If
End Sub

Private Sub vsSelTop_Change()
    If updateSelTopBar = True Then
        txtSelTop = Abs(32767 - CStr(vsSelTop.Value))
        txtSelTop.Refresh
        changeToSelTop
    End If
End Sub

Private Sub vsSelWidth_Change()
    If updateSelWidthBar = True Then
        txtSelWidth = Abs(32767 - CStr(vsSelWidth.Value))
        txtSelWidth.Refresh
        changeToSelWidth
    End If
End Sub

Private Sub vsSelHeight_Change()
    If updateSelHeightBar = True Then
        txtSelHeight = Abs(32767 - CStr(vsSelHeight.Value))
        txtSelHeight.Refresh
        changeToSelHeight
    End If
End Sub

'The next four routines are used to keep the selection text boxes and scrollbars in sync
Public Sub changeToSelLeft()
    If EntryValid(txtSelLeft, 0, 32767, False, True) Then
        updateSelLeftBar = False
        vsSelLeft.Value = Abs(32767 - CInt(txtSelLeft))
        updateSelLeftBar = True
    End If
    If pdImages(CurrentImage).selectionActive And (pdImages(CurrentImage).mainSelection.rejectRefreshRequests = False) Then pdImages(CurrentImage).mainSelection.updateViaTextBox
End Sub

Public Sub changeToSelTop()
    If EntryValid(txtSelTop, 0, 32767, False, True) Then
        updateSelTopBar = False
        vsSelTop.Value = Abs(32767 - CInt(txtSelTop))
        updateSelTopBar = True
    End If
    If pdImages(CurrentImage).selectionActive And (pdImages(CurrentImage).mainSelection.rejectRefreshRequests = False) Then pdImages(CurrentImage).mainSelection.updateViaTextBox
End Sub

Public Sub changeToSelWidth()
    If EntryValid(txtSelWidth, 1, 32767, False, True) Then
        updateSelWidthBar = False
        vsSelWidth.Value = Abs(32767 - CInt(txtSelWidth))
        updateSelWidthBar = True
    End If
    If pdImages(CurrentImage).selectionActive And (pdImages(CurrentImage).mainSelection.rejectRefreshRequests = False) Then pdImages(CurrentImage).mainSelection.updateViaTextBox
End Sub

Public Sub changeToSelHeight()
    If EntryValid(txtSelHeight, 1, 32767, False, True) Then
        updateSelHeightBar = False
        vsSelHeight.Value = Abs(32767 - CInt(txtSelHeight))
        updateSelHeightBar = True
    End If
    If pdImages(CurrentImage).selectionActive And (pdImages(CurrentImage).mainSelection.rejectRefreshRequests = False) Then pdImages(CurrentImage).mainSelection.updateViaTextBox
End Sub

