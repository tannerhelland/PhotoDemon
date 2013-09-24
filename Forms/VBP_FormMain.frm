VERSION 5.00
Begin VB.MDIForm FormMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H80000010&
   Caption         =   "PhotoDemon by Tanner Helland - www.tannerhelland.com"
   ClientHeight    =   9480
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   16065
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
      Height          =   9105
      Left            =   13095
      ScaleHeight     =   607
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   198
      TabIndex        =   14
      Top             =   0
      Width           =   2970
      Begin VB.CommandButton cmdTools 
         Caption         =   "Select (Line)"
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
         Index           =   2
         Left            =   1995
         TabIndex        =   35
         ToolTipText     =   "Line Selection tool"
         Top             =   465
         Width           =   900
      End
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
         Left            =   75
         TabIndex        =   25
         ToolTipText     =   "Rectangular Selection tool"
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
         Left            =   1035
         TabIndex        =   24
         ToolTipText     =   "Elliptical (Oval) Selection tool"
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
         Height          =   6135
         Index           =   0
         Left            =   0
         ScaleHeight     =   409
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   198
         TabIndex        =   15
         Top             =   1440
         Visible         =   0   'False
         Width           =   2970
         Begin VB.ComboBox cmbSelSmoothing 
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
            Left            =   120
            List            =   "VBP_FormMain.frx":000E
            Style           =   2  'Dropdown List
            TabIndex        =   34
            TabStop         =   0   'False
            ToolTipText     =   "Use this option to change the way selections blend with their surroundings."
            Top             =   3000
            Width           =   2685
         End
         Begin VB.ComboBox cmbSelType 
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
            ItemData        =   "VBP_FormMain.frx":0010
            Left            =   120
            List            =   "VBP_FormMain.frx":0012
            Style           =   2  'Dropdown List
            TabIndex        =   30
            TabStop         =   0   'False
            ToolTipText     =   "This option controls the selection's area.  You can switch between the three settings without losing the current selection."
            Top             =   4320
            Width           =   2685
         End
         Begin PhotoDemon.sliderTextCombo sltCornerRounding 
            CausesValidation=   0   'False
            Height          =   495
            Left            =   0
            TabIndex        =   27
            Top             =   5640
            Width           =   3000
            _ExtentX        =   5318
            _ExtentY        =   873
            Max             =   10000
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
            ItemData        =   "VBP_FormMain.frx":0014
            Left            =   180
            List            =   "VBP_FormMain.frx":0016
            Style           =   2  'Dropdown List
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   540
            Width           =   2685
         End
         Begin PhotoDemon.textUpDown tudSel 
            Height          =   405
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   1320
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   714
            Min             =   -30000
            Max             =   30000
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
         Begin PhotoDemon.textUpDown tudSel 
            Height          =   405
            Index           =   1
            Left            =   1560
            TabIndex        =   18
            Top             =   1320
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   714
            Min             =   -30000
            Max             =   30000
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
         Begin PhotoDemon.textUpDown tudSel 
            Height          =   405
            Index           =   2
            Left            =   120
            TabIndex        =   19
            Top             =   2160
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   714
            Min             =   -30000
            Max             =   30000
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
         Begin PhotoDemon.textUpDown tudSel 
            Height          =   405
            Index           =   3
            Left            =   1560
            TabIndex        =   20
            Top             =   2160
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   714
            Min             =   -30000
            Max             =   30000
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
         Begin PhotoDemon.sliderTextCombo sltSelectionBorder 
            CausesValidation=   0   'False
            Height          =   495
            Left            =   0
            TabIndex        =   31
            Top             =   4710
            Visible         =   0   'False
            Width           =   3000
            _ExtentX        =   5318
            _ExtentY        =   873
            Min             =   1
            Max             =   10000
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
         Begin PhotoDemon.sliderTextCombo sltSelectionFeathering 
            CausesValidation=   0   'False
            Height          =   495
            Left            =   0
            TabIndex        =   33
            Top             =   3390
            Width           =   3000
            _ExtentX        =   5318
            _ExtentY        =   873
            Max             =   100
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
         Begin PhotoDemon.sliderTextCombo sltSelectionLineWidth 
            CausesValidation=   0   'False
            Height          =   495
            Left            =   0
            TabIndex        =   36
            Top             =   5640
            Width           =   3000
            _ExtentX        =   5318
            _ExtentY        =   873
            Min             =   1
            Max             =   10000
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
         Begin VB.Label lblSelection 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "smoothing"
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
            TabIndex        =   32
            Top             =   2640
            Width           =   1125
         End
         Begin VB.Label lblSelection 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "selection type"
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
            Index           =   4
            Left            =   120
            TabIndex        =   29
            Top             =   3960
            Width           =   1440
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
            Index           =   5
            Left            =   120
            TabIndex        =   28
            Top             =   5280
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
            TabIndex        =   23
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
            TabIndex        =   22
            Top             =   960
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
            TabIndex        =   21
            Top             =   1800
            Width           =   1380
         End
      End
      Begin VB.Line lineMain 
         BorderColor     =   &H80000002&
         Index           =   1
         X1              =   5
         X2              =   192
         Y1              =   92
         Y2              =   92
      End
      Begin VB.Label lblTools 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "selection tools"
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
         TabIndex        =   26
         Top             =   90
         Width           =   1500
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
      ScaleWidth      =   1071
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   9105
      Width           =   16065
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
      Height          =   9105
      Left            =   0
      ScaleHeight     =   607
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   70
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   1050
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
         ItemData        =   "VBP_FormMain.frx":0018
         Left            =   60
         List            =   "VBP_FormMain.frx":001A
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
         PictureNormal   =   "VBP_FormMain.frx":001C
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
         PictureNormal   =   "VBP_FormMain.frx":146E
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
         PictureNormal   =   "VBP_FormMain.frx":26D0
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
         PictureNormal   =   "VBP_FormMain.frx":3722
         DisabledPictureMode=   1
         CaptionEffects  =   0
         TooltipTitle    =   "Redo"
         TooltipBackColor=   -2147483643
      End
      Begin PhotoDemon.jcbutton cmdClose 
         Height          =   615
         Left            =   60
         TabIndex        =   10
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
         PictureNormal   =   "VBP_FormMain.frx":4774
         DisabledPictureMode=   1
         CaptionEffects  =   0
         TooltipTitle    =   "Close"
      End
      Begin PhotoDemon.jcbutton cmdSaveAs 
         Height          =   615
         Left            =   60
         TabIndex        =   11
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
         PictureNormal   =   "VBP_FormMain.frx":57C6
         DisabledPictureMode=   1
         CaptionEffects  =   0
         TooltipTitle    =   "Save As"
      End
      Begin PhotoDemon.jcbutton cmdZoomIn 
         Height          =   450
         Left            =   525
         TabIndex        =   12
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
         PictureNormal   =   "VBP_FormMain.frx":6A28
         DisabledPictureMode=   1
         CaptionEffects  =   0
         ToolTip         =   "Use this button to increase image zoom."
         TooltipTitle    =   "Zoom In"
      End
      Begin PhotoDemon.jcbutton cmdZoomOut 
         Height          =   450
         Left            =   45
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
         PictureNormal   =   "VBP_FormMain.frx":6E7A
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
         Height          =   1620
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
   Begin VB.Menu MnuImageTop 
      Caption         =   "&Image"
      Begin VB.Menu MnuImage 
         Caption         =   "&Duplicate"
         Index           =   0
         Shortcut        =   ^D
      End
      Begin VB.Menu MnuImage 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Transparency"
         Index           =   2
         Begin VB.Menu MnuTransparency 
            Caption         =   "Add basic transparency..."
            Index           =   0
         End
         Begin VB.Menu MnuTransparency 
            Caption         =   "Make color transparent..."
            Index           =   1
         End
         Begin VB.Menu MnuTransparency 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu MnuTransparency 
            Caption         =   "Remove transparency..."
            Index           =   3
         End
      End
      Begin VB.Menu MnuImage 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Resize..."
         Index           =   4
         Shortcut        =   ^R
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Canvas size..."
         Index           =   5
      End
      Begin VB.Menu MnuImage 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Crop to selection"
         Index           =   7
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Autocrop"
         Index           =   8
      End
      Begin VB.Menu MnuImage 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Rotate"
         Index           =   10
         Begin VB.Menu MnuRotate 
            Caption         =   "90° clockwise"
            Index           =   0
         End
         Begin VB.Menu MnuRotate 
            Caption         =   "90° counter-clockwise"
            Index           =   1
         End
         Begin VB.Menu MnuRotate 
            Caption         =   "180°"
            Index           =   2
         End
         Begin VB.Menu MnuRotate 
            Caption         =   "Arbitrary..."
            Index           =   3
         End
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Flip horizontal"
         Index           =   11
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Flip vertical"
         Index           =   12
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Convert to isometric view"
         Index           =   13
      End
      Begin VB.Menu MnuImage 
         Caption         =   "-"
         Index           =   14
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Indexed color..."
         Index           =   15
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Tile..."
         Index           =   16
      End
      Begin VB.Menu MnuImage 
         Caption         =   "-"
         Index           =   17
      End
      Begin VB.Menu MnuImage 
         Caption         =   "Metadata"
         Index           =   18
         Begin VB.Menu MnuMetadata 
            Caption         =   "Browse image metadata..."
            Index           =   0
         End
         Begin VB.Menu MnuMetadata 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu MnuMetadata 
            Caption         =   "Count unique colors"
            Index           =   2
         End
         Begin VB.Menu MnuMetadata 
            Caption         =   "Map photo location..."
            Index           =   3
         End
      End
   End
   Begin VB.Menu MnuSelectTop 
      Caption         =   "&Select"
      Begin VB.Menu MnuSelect 
         Caption         =   "All"
         Index           =   0
         Shortcut        =   ^A
      End
      Begin VB.Menu MnuSelect 
         Caption         =   "None"
         Index           =   1
      End
      Begin VB.Menu MnuSelect 
         Caption         =   "Invert"
         Index           =   2
      End
      Begin VB.Menu MnuSelect 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu MnuSelect 
         Caption         =   "Grow..."
         Index           =   4
      End
      Begin VB.Menu MnuSelect 
         Caption         =   "Shrink..."
         Index           =   5
      End
      Begin VB.Menu MnuSelect 
         Caption         =   "Border..."
         Index           =   6
      End
      Begin VB.Menu MnuSelect 
         Caption         =   "Feather..."
         Index           =   7
      End
      Begin VB.Menu MnuSelect 
         Caption         =   "Sharpen..."
         Index           =   8
      End
      Begin VB.Menu MnuSelect 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu MnuSelect 
         Caption         =   "Load selection..."
         Index           =   10
      End
      Begin VB.Menu MnuSelect 
         Caption         =   "Save current selection..."
         Index           =   11
      End
   End
   Begin VB.Menu MnuColorTop 
      Caption         =   "&Adjustments"
      Begin VB.Menu MnuColor 
         Caption         =   "Brightness and contrast..."
         Index           =   0
      End
      Begin VB.Menu MnuColor 
         Caption         =   "Curves..."
         Index           =   1
      End
      Begin VB.Menu MnuColor 
         Caption         =   "Exposure..."
         Index           =   2
      End
      Begin VB.Menu MnuColor 
         Caption         =   "Gamma..."
         Index           =   3
         Shortcut        =   ^G
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
         Caption         =   "White balance..."
         Index           =   6
         Shortcut        =   ^W
      End
      Begin VB.Menu MnuColor 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu MnuColor 
         Caption         =   "Color balance..."
         Index           =   8
      End
      Begin VB.Menu MnuColor 
         Caption         =   "Hue and saturation..."
         Index           =   9
         Shortcut        =   ^H
      End
      Begin VB.Menu MnuColor 
         Caption         =   "Photo filters..."
         Index           =   10
      End
      Begin VB.Menu MnuColor 
         Caption         =   "Temperature..."
         Index           =   11
         Shortcut        =   ^T
      End
      Begin VB.Menu MnuColor 
         Caption         =   "Vibrance..."
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
         Caption         =   "Grayscale..."
         Index           =   15
      End
      Begin VB.Menu MnuColor 
         Caption         =   "Invert"
         Index           =   16
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
         Index           =   17
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
         Index           =   18
      End
      Begin VB.Menu MnuColor 
         Caption         =   "-"
         Index           =   19
      End
      Begin VB.Menu MnuColor 
         Caption         =   "Components"
         Index           =   20
         Begin VB.Menu MnuColorComponents 
            Caption         =   "Channel mixer..."
            Index           =   0
         End
         Begin VB.Menu MnuColorComponents 
            Caption         =   "Rechannel..."
            Index           =   1
         End
         Begin VB.Menu MnuColorComponents 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu MnuColorComponents 
            Caption         =   "Shift channels right"
            Index           =   3
         End
         Begin VB.Menu MnuColorComponents 
            Caption         =   "Shift channels left"
            Index           =   4
         End
         Begin VB.Menu MnuColorComponents 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu MnuColorComponents 
            Caption         =   "Maximum channel..."
            Index           =   6
         End
         Begin VB.Menu MnuColorComponents 
            Caption         =   "Minimum channel..."
            Index           =   7
         End
      End
      Begin VB.Menu MnuColor 
         Caption         =   "-"
         Index           =   21
      End
      Begin VB.Menu MnuColor 
         Caption         =   "Histogram"
         Index           =   22
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
   End
   Begin VB.Menu MnuFilter 
      Caption         =   "Effe&cts"
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
            Caption         =   "Comic book"
            Index           =   0
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Figured glass (dents)..."
            Index           =   1
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Film noir"
            Index           =   2
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Kaleiodoscope..."
            Index           =   3
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Modern art..."
            Index           =   4
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Oil painting..."
            Index           =   5
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Pencil drawing"
            Index           =   6
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Posterize..."
            Index           =   7
         End
         Begin VB.Menu MnuArtistic 
            Caption         =   "Relief"
            Index           =   8
         End
      End
      Begin VB.Menu MnuEffectUpper 
         Caption         =   "Blur"
         Index           =   1
         Begin VB.Menu MnuBlurFilter 
            Caption         =   "Box blur..."
            Index           =   0
         End
         Begin VB.Menu MnuBlurFilter 
            Caption         =   "Gaussian blur..."
            Index           =   1
         End
         Begin VB.Menu MnuBlurFilter 
            Caption         =   "Grid blur"
            Index           =   2
         End
         Begin VB.Menu MnuBlurFilter 
            Caption         =   "Motion blur..."
            Index           =   3
         End
         Begin VB.Menu MnuBlurFilter 
            Caption         =   "Pixelate..."
            Index           =   4
         End
         Begin VB.Menu MnuBlurFilter 
            Caption         =   "Radial blur..."
            Index           =   5
         End
         Begin VB.Menu MnuBlurFilter 
            Caption         =   "Smart blur..."
            Index           =   6
         End
         Begin VB.Menu MnuBlurFilter 
            Caption         =   "Zoom blur..."
            Index           =   7
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
            Caption         =   "Miscellaneous..."
            Index           =   2
         End
         Begin VB.Menu MnuDistortFilter 
            Caption         =   "Pan and zoom..."
            Index           =   3
         End
         Begin VB.Menu MnuDistortFilter 
            Caption         =   "Perspective..."
            Index           =   4
         End
         Begin VB.Menu MnuDistortFilter 
            Caption         =   "Pinch and whirl..."
            Index           =   5
         End
         Begin VB.Menu MnuDistortFilter 
            Caption         =   "Poke..."
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
            Caption         =   "Rotate..."
            Index           =   9
         End
         Begin VB.Menu MnuDistortFilter 
            Caption         =   "Shear..."
            Index           =   10
         End
         Begin VB.Menu MnuDistortFilter 
            Caption         =   "Spherize..."
            Index           =   11
         End
         Begin VB.Menu MnuDistortFilter 
            Caption         =   "Squish..."
            Index           =   12
         End
         Begin VB.Menu MnuDistortFilter 
            Caption         =   "Swirl..."
            Index           =   13
         End
         Begin VB.Menu MnuDistortFilter 
            Caption         =   "Waves..."
            Index           =   14
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
            Caption         =   "Median..."
            Index           =   3
         End
      End
      Begin VB.Menu MnuEffectUpper 
         Caption         =   "Sharpen"
         Index           =   8
         Begin VB.Menu MnuSharpen 
            Caption         =   "Sharpen..."
            Index           =   0
         End
         Begin VB.Menu MnuSharpen 
            Caption         =   "Unsharp masking..."
            Index           =   1
         End
      End
      Begin VB.Menu MnuEffectUpper 
         Caption         =   "Stylize"
         Index           =   9
         Begin VB.Menu MnuStylize 
            Caption         =   "Antique"
            Index           =   0
         End
         Begin VB.Menu MnuStylize 
            Caption         =   "Diffuse..."
            Index           =   1
         End
         Begin VB.Menu MnuStylize 
            Caption         =   "Dilate..."
            Index           =   2
         End
         Begin VB.Menu MnuStylize 
            Caption         =   "Erode..."
            Index           =   3
         End
         Begin VB.Menu MnuStylize 
            Caption         =   "Solarize..."
            Index           =   4
         End
         Begin VB.Menu MnuStylize 
            Caption         =   "Twins..."
            Index           =   5
         End
         Begin VB.Menu MnuStylize 
            Caption         =   "Vignetting..."
            Index           =   6
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
         Caption         =   "Language editor..."
         Index           =   1
      End
      Begin VB.Menu mnuTool 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuTool 
         Caption         =   "&Macros"
         Index           =   3
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
         Index           =   4
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Options"
         Index           =   5
      End
      Begin VB.Menu mnuTool 
         Caption         =   "Plugin manager"
         Index           =   6
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

'PhotoDemon is Copyright ©1999-2013 by Tanner Helland, tannerhelland.com

'Please visit photodemon.org for updates and additional downloads

'***************************************************************************
'Main Program MDI Form
'Copyright ©2002-2013 by Tanner Helland
'Created: 15/September/02
'Last updated: 04/July/13
'Last update: overhauled much of the selection-related control code to better separate interface/domain handling
'
'This is PhotoDemon's main form.  In actuality, it contains relatively little code.  Its
' primary purpose is sending parameters to other, more interesting sections of the program.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
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
Private m_ToolTip As clsToolTip

'The value of all controls on this form are saved and loaded to file by this class
Private WithEvents lastUsedSettings As pdLastUsedSettings
Attribute lastUsedSettings.VB_VarHelpID = -1

'When the selection type is changed, update the corresponding preference and redraw all selections
Private Sub cmbSelRender_Click(Index As Integer)
            
    If NumOfWindows > 0 Then
    
        Dim i As Long
        For i = 0 To NumOfImagesLoaded
            If (Not pdImages(i) Is Nothing) Then
                If pdImages(i).IsActive And pdImages(i).selectionActive Then RenderViewport pdImages(i).containingForm
            End If
        Next i
    
    End If
    
End Sub

'Change selection smoothing (e.g. none, antialiased, fully feathered)
Private Sub cmbSelSmoothing_Click(Index As Integer)
    
    updateSelectionPanelLayout
    
    'If a selection is already active, change its type to match the current selection, then redraw it
    If selectionsAllowed(False) Then
        pdImages(FormMain.ActiveForm.Tag).mainSelection.setSmoothingType cmbSelSmoothing(Index).ListIndex
        pdImages(FormMain.ActiveForm.Tag).mainSelection.setFeatheringRadius sltSelectionFeathering.Value
        RenderViewport FormMain.ActiveForm
    End If
    
End Sub

'Change selection type (e.g. interior, exterior, bordered)
Private Sub cmbSelType_Click(Index As Integer)

    updateSelectionPanelLayout
    
    'If a selection is already active, change its type to match the current selection, then redraw it
    If selectionsAllowed(False) Then
        pdImages(FormMain.ActiveForm.Tag).mainSelection.setSelectionType cmbSelType(Index).ListIndex
        pdImages(FormMain.ActiveForm.Tag).mainSelection.setBorderSize sltSelectionBorder.Value
        RenderViewport FormMain.ActiveForm
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
    Process "Open", True
End Sub

Private Sub cmdRedo_Click()
    Process "Redo", , , False
End Sub

Private Sub cmdSave_Click()
    Process "Save", , , False
End Sub

Private Sub cmdSaveAs_Click()
    Process "Save as", True, , False
End Sub

'External functions can use this to request the selection of a new tool (for example, Select All uses this to set the
' rectangular tool selector as the current tool)
Public Sub selectNewTool(ByVal newToolID As Long)
    g_PreviousTool = g_CurrentTool
    g_CurrentTool = newToolID
    resetToolButtonStates
End Sub

Private Sub cmdTools_Click(Index As Integer)
    g_PreviousTool = g_CurrentTool
    g_CurrentTool = Index
    resetToolButtonStates
End Sub

'Private Sub cmdTools_LostFocus(Index As Integer)
    'g_CurrentTool = Index
    'resetToolButtonStates
'End Sub

Private Sub cmdTools_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    g_CurrentTool = Index
    resetToolButtonStates
End Sub

'When a new tool button is selected, we need to raise all the others and display the proper options box
Public Sub resetToolButtonStates()
    
    'Start by depressing the selected button and raising all unselected ones
    Dim i As Long
    For i = 0 To cmdTools.Count - 1
        SendMessageA cmdTools(i).hWnd, BM_SETSTATE, False, 0
    Next i
    SendMessageA cmdTools(g_CurrentTool).hWnd, BM_SETSTATE, True, 0
    
    'Next, we need to display the correct tool options panel.  There is no set pattern to this; some tools share
    ' panels, but show/hide certain controls as necessary.  Other tools require their own unique panel.  I've tried
    ' to strike a balance between "as few panels as possible" without going overboard.
    Dim activeToolPanel As Long
    
    Select Case g_CurrentTool
        
        'Rectangular, Elliptical selections
        Case SELECT_RECT, SELECT_CIRC, SELECT_LINE
            activeToolPanel = 0
        
        Case Else
        
    End Select
    
    'If tools share the same panel, they may need to show or hide a few additional controls.  (For example,
    ' "corner rounding", which is needed for rectangular selections but not elliptical ones, despite the two
    ' sharing the same tool panel.)  Do this before showing or hiding the tool panel.
    Select Case g_CurrentTool
    
        'For rectangular selections, show the rounded corners option
        Case SELECT_RECT
            FormMain.lblSelection(5).Visible = True
            FormMain.sltCornerRounding.Visible = True
            FormMain.sltSelectionLineWidth.Visible = False
            
        'For elliptical selections, hide the rounded corners option
        Case SELECT_CIRC
            FormMain.lblSelection(5).Visible = False
            FormMain.sltCornerRounding.Visible = False
            FormMain.sltSelectionLineWidth.Visible = False
            
        'Line selections also show the rounded corners slider, though they repurpose it for line width
        Case SELECT_LINE
            FormMain.lblSelection(5).Visible = True
            FormMain.sltCornerRounding.Visible = False
            FormMain.sltSelectionLineWidth.Visible = True
        
    End Select
    
    'Even if tools share the same panel, they may name controls differently, or use different max/min values.
    ' Check for this, and apply new text and max/min settings as necessary.
    Select Case g_CurrentTool
    
        'Rectangular and elliptical selections use rectangular bounding boxes and potential corner rounding
        Case SELECT_RECT, SELECT_CIRC
            lblSelection(1).Caption = g_Language.TranslateMessage("selection position")
            lblSelection(2).Caption = g_Language.TranslateMessage("selection size")
            lblSelection(5).Caption = g_Language.TranslateMessage("corner rounding")
            'If (g_PreviousTool <> SELECT_RECT) And (g_PreviousTool <> SELECT_CIRC) Then
                'If selectionsAllowed And (Not g_UndoRedoActive) Then Process "Remove selection", , , 2, g_PreviousTool
                'If g_CurrentTool = SELECT_RECT Then sltCornerRounding.Value = 0
            'End If
            
        'Line selections use two points, and the corner rounding slider gets repurposed as line width.
        Case SELECT_LINE
            'If selectionsAllowed And (Not g_UndoRedoActive) Then Process "Remove selection", , , 2, g_PreviousTool
            lblSelection(1).Caption = g_Language.TranslateMessage("first point (x, y)")
            lblSelection(2).Caption = g_Language.TranslateMessage("second point (x, y)")
            lblSelection(5).Caption = g_Language.TranslateMessage("line width")
            
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
        Case SELECT_RECT
                
            'If a similar selection is already active, change its shape to match the current tool, then redraw it
            If selectionsAllowed(True) And (Not g_UndoRedoActive) Then
                If (g_PreviousTool = SELECT_CIRC) And (pdImages(CurrentImage).mainSelection.getSelectionShape = sCircle) Then
                    pdImages(FormMain.ActiveForm.Tag).mainSelection.setSelectionShape g_CurrentTool
                    RenderViewport FormMain.ActiveForm
                Else
                    If pdImages(CurrentImage).mainSelection.getSelectionShape = sRectangle Then
                        metaToggle tSelectionTransform, True
                    Else
                        metaToggle tSelectionTransform, False
                    End If
                End If
            End If
            
        Case SELECT_CIRC
        
            'If a similar selection is already active, change its shape to match the current tool, then redraw it
            If selectionsAllowed(True) And (Not g_UndoRedoActive) Then
                If (g_PreviousTool = SELECT_RECT) And (pdImages(CurrentImage).mainSelection.getSelectionShape = sRectangle) Then
                    pdImages(FormMain.ActiveForm.Tag).mainSelection.setSelectionShape g_CurrentTool
                    RenderViewport FormMain.ActiveForm
                Else
                    If pdImages(CurrentImage).mainSelection.getSelectionShape = sCircle Then
                        metaToggle tSelectionTransform, True
                    Else
                        metaToggle tSelectionTransform, False
                    End If
                End If
            End If
            
        'Line selections
        Case SELECT_LINE
        
            'Deactivate the position text boxes - those shouldn't be accessible unless a line selection is presently active
            If selectionsAllowed(True) Then
                If pdImages(CurrentImage).mainSelection.getSelectionShape = sLine Then
                    metaToggle tSelectionTransform, True
                Else
                    metaToggle tSelectionTransform, False
                End If
            Else
                metaToggle tSelectionTransform, False
            End If
            
        Case Else
        
    End Select
    
End Sub

Private Sub cmdUndo_Click()
    Process "Undo", , , False
End Sub

Private Sub cmdZoomIn_Click()
    FormMain.CmbZoom.ListIndex = FormMain.CmbZoom.ListIndex - 1
End Sub

Private Sub cmdZoomOut_Click()
    FormMain.CmbZoom.ListIndex = FormMain.CmbZoom.ListIndex + 1
End Sub

Private Sub lastUsedSettings_AddCustomPresetData()
    
    'Write the currently selected selection tool to file
    lastUsedSettings.addPresetData "ActiveSelectionTool", g_CurrentTool
    
End Sub

Private Sub lastUsedSettings_ReadCustomPresetData()

    'Restore the last-used selection tool (which will be saved in the main form's preset file, if it exists)
    g_PreviousTool = -1
    If Len(lastUsedSettings.retrievePresetData("ActiveSelectionTool")) > 0 Then
        g_CurrentTool = CLng(lastUsedSettings.retrievePresetData("ActiveSelectionTool"))
    Else
        g_CurrentTool = SELECT_RECT
    End If
    cmdTools_Click CInt(g_CurrentTool)
    
    'The save/load preset control will save/load values of a number of controls that don't need to be saved.
    ' Reset those controls now.
    CmbZoom.ListIndex = ZoomIndex100
    
    Dim i As Long
    For i = 0 To tudSel.Count - 1
        tudSel(i) = 0
    Next i

End Sub

'THE BEGINNING OF EVERYTHING
' Actually, Sub "Main" in the module "modMain" is loaded first, but all it does is set up native theming.  Once it has done that, FormMain is loaded.
Private Sub MDIForm_Load()

    'Use a global variable to store any command line parameters we may have been passed
    g_CommandLine = Command$
    
    'Instantiate the themed tooltip class
    Set m_ToolTip = New clsToolTip
    
    'The bulk of the loading code actually takes place inside the LoadTheprogram subroutine (which can be found in the "Loading" module)
    LoadTheProgram
        
    'Hide the selection tools
    metaToggle tSelection, False
    
    'Render images to the toolbox command buttons
    Dim i As Long
    
    'Extract relevant icons from the resource file, and render them onto the buttons at run-time.
    ' (NOTE: because the icons require manifest theming, they will not appear in the IDE.)
    Set cImgCtl = New clsControlImage
    If g_IsProgramCompiled Then
        
        'Remove all tool button captions
        For i = 0 To cmdTools.Count - 1
            cmdTools(i).Caption = ""
        Next i
        
        With cImgCtl
            
            'Load the tool images (in PNG format) from the resource file
            .LoadImageFromStream cmdTools(0).hWnd, LoadResData("T_SELRECT", "CUSTOM"), 22, 22
            .LoadImageFromStream cmdTools(1).hWnd, LoadResData("T_SELCIRCLE", "CUSTOM"), 22, 22
            .LoadImageFromStream cmdTools(2).hWnd, LoadResData("T_SELLINE", "CUSTOM"), 22, 22
            
            'Center-align the images in their respective buttons
            For i = 0 To cmdTools.Count - 1
                .SetMargins cmdTools(i).hWnd, 0
                .Align(cmdTools(i).hWnd) = Icon_Center
                
                'On XP, the tool command button images aren't aligned properly until the buttons are hovered.  No one
                ' knows why.  We can imitate a hover with a click - do so now.
                If Not g_IsVistaOrLater Then cmdTools_Click CInt(i)
            Next i
            
        End With
    End If
            
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
    allowedToUpdate = g_UserPreferences.GetPref_Boolean("Updates", "Check For Updates", True)
        
    'If updates ARE allowed, see when we last checked.  To be polite, only check once every 10 days.
    If allowedToUpdate Then
    
        Dim lastCheckDate As String
        lastCheckDate = g_UserPreferences.GetPref_String("Updates", "Last Update Check", "")
        
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
    
        Dim updateNeeded As UpdateCheck
        updateNeeded = CheckForSoftwareUpdate
        
        'CheckForSoftwareUpdate can return one of three values:
        ' UPDATE_ERROR - something went wrong (no Internet connection, etc)
        ' UPDATE_NOT_NEEDED - the check was successful, but this version is up-to-date
        ' UPDATE_AVAILABLE - the check was successful, and an update is available
        
        Select Case updateNeeded
        
            Case UPDATE_ERROR
                Message "An error occurred while checking for updates.  Please make sure you have an active Internet connection."
            
            Case UPDATE_NOT_NEEDED
                Message "Software is up-to-date."
                
                'Because the software is up-to-date, we can mark this as a successful check in the preferences file
                g_UserPreferences.SetPref_String "Updates", "Last Update Check", Format$(Now, "Medium Date")
                
            Case UPDATE_AVAILABLE
                Message "Software update found!  Launching update notifier..."
                FormSoftwareUpdate.Show vbModal, Me
            
        End Select
            
    End If
    
    'Last but not least, if any core plugin files were marked as "missing," offer to download them
    ' (NOTE: this check is superceded by the update check - since a full program update will include the missing plugins -
    '        so ignore this request if the user was already notified of an update.)
    If (updateNeeded <> UPDATE_AVAILABLE) And ((Not isZLibAvailable) Or (Not isEZTwainAvailable) Or (Not isFreeImageAvailable) Or (Not isPngnqAvailable) Or (Not isExifToolAvailable)) Then
    
        Message "Some core plugins could not be found. Preparing updater..."
        
        'As a courtesy, if the user has asked us to stop bugging them about downloading plugins, obey their request
        Dim promptToDownload As Boolean
        promptToDownload = g_UserPreferences.GetPref_Boolean("Updates", "Prompt For Plugin Download", True)
                
        'Finally, if allowed, we can prompt the user to download the recommended plugin set
        If promptToDownload Then
            FormPluginDownloader.Show vbModal, FormMain
            
            'Since plugins may have been downloaded, update the interface to match any new features that may be available.
            LoadPlugins
            ApplyAllMenuIcons
            ResetMenuIcons
            g_ImageFormats.generateInputFormats
            g_ImageFormats.generateOutputFormats
            
        Else
            Message "Ignoring plugin update request per user's saved preference"
        End If
    
    End If
    
    'Load the last-used settings for the main form
    Set lastUsedSettings = New pdLastUsedSettings
    lastUsedSettings.setParentForm Me
    lastUsedSettings.loadAllControlValues
    
    Message "Please load an image.  (The large 'Open Image' button at the top-left should do the trick!)"
    
    'Render the main form with any extra visual styles we've decided to apply
    RedrawMainForm
        
    'Because people may be using this code in the IDE, warn them about the consequences of doing so
    If (Not g_IsProgramCompiled) And (g_UserPreferences.GetPref_Boolean("Core", "Display IDE Warning", True)) Then displayIDEWarning
    
    'Finally, return focus to the main form
    'FormMain.SetFocus
     
End Sub

'Allow the user to drag-and-drop files from Windows Explorer onto the main MDI form
Private Sub MDIForm_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

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

Private Sub MDIForm_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)

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
    
    'Save all last-used settings to file
    lastUsedSettings.saveAllControlValues
    
    'If the histogram form is open, close it
    Unload FormHistogram
    
    'If the user wants us to remember the program's last-used location, store those values to file now
    If g_UserPreferences.GetPref_Boolean("Core", "Remember Window Location", True) Then
    
        g_UserPreferences.SetPref_Long "Core", "Last Window State", Me.WindowState
        g_UserPreferences.SetPref_Long "Core", "Last Window Left", Me.Left / Screen.TwipsPerPixelX
        g_UserPreferences.SetPref_Long "Core", "Last Window Top", Me.Top / Screen.TwipsPerPixelY
        g_UserPreferences.SetPref_Long "Core", "Last Window Width", Me.Width / Screen.TwipsPerPixelX
        g_UserPreferences.SetPref_Long "Core", "Last Window Height", Me.Height / Screen.TwipsPerPixelY
    
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

    'Save the MRU list to the preferences file.  (I've considered doing this as files are loaded, but the only time
    ' that would be an improvement is if the program crashes, and if it does crash, the user wouldn't want to re-load
    ' the problematic image anyway.)
    MRU_SaveToFile
        
    'Restore the user's font smoothing setting as necessary
    handleClearType False
    
    ReleaseFormTheming Me
    
End Sub

'All artistic filters are launched here
Private Sub MnuArtistic_Click(Index As Integer)

    Select Case Index
            
        'Comic book
        Case 0
            Process "Comic book"
            
        'Figured glass
        Case 1
            Process "Figured glass", True
            
        'Film noir
        Case 2
            Process "Film noir"
        
        'Kaleidoscope
        Case 3
            Process "Kaleidoscope", True
        
        'Modern art
        Case 4
            Process "Modern art", True
        
        'Oil painting
        Case 5
            Process "Oil painting", True
            
        'Pencil drawing
        Case 6
            Process "Pencil drawing"
                
        'Posterize
        Case 7
            Process "Posterize", True
            
        'Relief
        Case 8
            Process "Relief"
    
    End Select

End Sub

Private Sub MnuBatchConvert_Click()
    g_AllowDragAndDrop = False
    FormBatchWizard.Show vbModal, FormMain
    g_AllowDragAndDrop = True
End Sub

Private Sub MnuBlackLight_Click()
    Process "Black light", True
End Sub

'All blur filters are handled here
Private Sub MnuBlurFilter_Click(Index As Integer)

    Select Case Index
        
        'Box blur
        Case 0
            Process "Box blur", True
            
        'Gaussian blur
        Case 1
            Process "Gaussian blur", True
                
        'Grid blur
        Case 2
            Process "Grid blur"
            
        'Motion blur
        Case 3
            Process "Motion blur", True
            
        'Pixelate (mosaic)
        Case 4
            Process "Pixelate", True
        
        'Radial blur
        Case 5
            Process "Radial blur", True
            
        'Smart Blur
        Case 6
            Process "Smart blur", True
        
        'Zoom Blur
        Case 7
            Process "Zoom blur", True
            
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
    
        'The first block of entries involves lighting fixes
        
        'Brightness/Contrast
        Case 0
            Process "Brightness and contrast", True
        
        'Curves
        Case 1
            'MsgBox "Still under construction!  The new Curves dialog should be finished in a week or so..."
            Process "Curves", True
        
        'Exposure
        Case 2
            Process "Exposure", True
            
        'Gamma correction
        Case 3
            Process "Gamma", True
            
        'Levels
        Case 4
            Process "Levels", True
            
        'Shadows/Midtones/Highlights
        Case 5
            Process "Shadows and highlights", True
            
        'White balance
        Case 6
            Process "White balance", True
            
        'The second block of entries involves color fixes
            
        'Separator
        Case 7
        
        'Color balance
        Case 8
            Process "Color balance", True
            
        'HSL
        Case 9
            Process "Hue and saturation", True
            
        'Photo filters
        Case 10
            Process "Photo filter", True
        
        'Temperature
        Case 11
            Process "Temperature", True
        
        'Vibrance
        Case 12
            Process "Vibrance", True
        
        'The third block of entries involves wholesale image color changes (grayscale, colorize, etc)
        
        '<separator>
        Case 13
        
        'Colorize
        Case 14
            Process "Colorize", True
        
        'Grayscale (black and white)
        Case 15
            Process "Grayscale", True
            
        '<Invert top-level>
        Case 16
        
        '<Monochrome top-level>
        Case 17
            
        'Sepia
        Case 18
            Process "Sepia"
            
        '<separator>
        Case 19
        
        '<Components top-level>
        Case 20
        
        '<separator>
        Case 21
        
        '<Histogram top-level>
        Case 22

        
    End Select

End Sub

'All entries in the "Color -> Components sub-menu are handled here"
Private Sub MnuColorComponents_Click(Index As Integer)

    Select Case Index
    
        'Channel mixer
        Case 0
            Process "Channel mixer", True
        
        'Rechannel
        Case 1
            Process "Rechannel", True
            
        '<separator>
        Case 2
        
        'Shift colors right
        Case 3
            Process "Shift colors (right)"
            
        'Shift colors left
        Case 4
            Process "Shift colors (left)"
            
        '<separator>
        Case 5
        
        'Max channel
        Case 6
            Process "Maximum channel"
        
        'Min channel
        Case 7
            Process "Minimum channel"
    
    End Select
    
End Sub

Private Sub MnuCompoundInvert_Click()
    Process "Compound invert", False, "128"
End Sub

Private Sub MnuCustomFilter_Click()
    Process "Custom filter", True
End Sub

'All distortion filters happen here
Private Sub MnuDistortFilter_Click(Index As Integer)

    Select Case Index
    
        'Apply lens distort
        Case 0
            Process "Apply lens distortion", True
        
        'Remove lens distort
        Case 1
            Process "Correct lens distortion", True
                
        'Miscellaneous
        Case 2
            Process "Miscellaneous distort", True
            
        'Pan and zoom
        Case 3
            Process "Pan and zoom", True
            
        'Perspective (free)
        Case 4
            Process "Perspective", True
            
        'Pinch and whirl
        Case 5
            Process "Pinch and whirl", True
        
        'Poke
        Case 6
            Process "Poke", True
            
        'Polar conversion
        Case 7
            Process "Polar conversion", True
        
        'Ripple
        Case 8
            Process "Ripple", True
        
        'Rotate
        Case 9
            Process "Rotate", True
        
        'Shear
        Case 10
            Process "Shear", True
            
        'Spherize
        Case 11
            Process "Spherize", True
            
        'Squish (formerly Fixed Perspective)
        Case 12
            Process "Squish", True
        
        'Swirl
        Case 13
            Process "Swirl", True
        
        'Waves
        Case 14
            Process "Waves", True
    
    End Select

End Sub

Private Sub MnuDream_Click()
    Process "Dream"
End Sub

Private Sub MnuEdge_Click(Index As Integer)

    Select Case Index
        
        'Emboss/engrave
        Case 0
            Process "Emboss or engrave", True
         
        'Enhance edges
        Case 1
            Process "Edge enhance"
        
        'Find edges
        Case 2
            Process "Find edges", True
        
        'Trace contour
        Case 3
            Process "Trace contour", True
    
    End Select

End Sub

Private Sub MnuFadeLastEffect_Click()
    Process "Fade last effect"
End Sub

Private Sub MnuFitOnScreen_Click()
    FitOnScreen
End Sub

Private Sub MnuFitWindowToImage_Click()
    If (FormMain.ActiveForm.WindowState = vbMaximized) Or (FormMain.ActiveForm.WindowState = vbMinimized) Then FormMain.ActiveForm.WindowState = vbNormal
    FitWindowToImage
End Sub

Private Sub MnuHeatmap_Click()
    Process "Thermograph (heat map)"
End Sub

'All help menu entries are launched from here
Private Sub MnuHelp_Click(Index As Integer)

    Select Case Index
        
        'Donations are so very, very welcome!
        Case 0
            OpenURL "http://photodemon.org/donate"
            
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
                        
                    'Because the software is up-to-date, we can mark this as a successful check in the preferences file
                    g_UserPreferences.SetPref_String "Updates", "Last Update Check", Format$(Now, "Medium Date")
                        
                Case 2
                    Message "Software update found!  Launching update notifier..."
                    FormSoftwareUpdate.Show vbModal, Me
                
            End Select
        
        'Submit feedback
        Case 3
            OpenURL "http://photodemon.org/about/contact/"
        
        'Submit bug report
        Case 4
            'GitHub requires a login for submitting Issues; check for that first
            Dim msgReturn As VbMsgBoxResult
            
            'If the user has previously been prompted about having a GitHub account, use their previous answer
            If g_UserPreferences.doesValueExist("Core ", "Has GitHub Account") Then
            
                Dim hasGitHub As Boolean
                hasGitHub = g_UserPreferences.GetPref_Boolean("Core", "Has GitHub Account", False)
                
                If hasGitHub Then msgReturn = vbYes Else msgReturn = vbNo
            
            'If this is the first time they are submitting feedback, ask them if they have a GitHub account
            Else
            
                msgReturn = pdMsgBox("Thank you for submitting a bug report.  To make sure your bug is addressed as quickly as possible, PhotoDemon needs to know where to send it." & vbCrLf & vbCrLf & "Do you have a GitHub account? (If you have no idea what this means, answer ""No"".)", vbQuestion + vbApplicationModal + vbYesNoCancel, "Thanks for fixing PhotoDemon")
                
                'If their answer was anything but "Cancel", store that answer to file
                If msgReturn = vbYes Then g_UserPreferences.SetPref_Boolean "Core", "Has GitHub Account", True
                If msgReturn = vbNo Then g_UserPreferences.SetPref_Boolean "Core", "Has GitHub Account", False
                
            End If
            
            'If they have a GitHub account, let them submit the bug there.  Otherwise, send them to the photodemon.org contact form
            If msgReturn = vbYes Then
                'Shell a browser window with the GitHub issue report form
                OpenURL "https://github.com/tannerhelland/PhotoDemon/issues/new"
            ElseIf msgReturn = vbNo Then
                'Shell a browser window with the photodemon.org contact form
                OpenURL "http://photodemon.org/about/contact/"
            End If
            
        'PhotoDemon's homepage
        Case 6
            OpenURL "http://www.photodemon.org"
            
        'Download source code
        Case 7
            OpenURL "https://github.com/tannerhelland/PhotoDemon"
        
        'Read terms and license agreement
        Case 8
            OpenURL "http://photodemon.org/about/license/"
            
        'Display About page
        Case 10
            FormAbout.Show vbModal, FormMain
        
    End Select

End Sub

Private Sub MnuHistogram_Click()
    'Process "Display histogram", True
    FormHistogram.Show vbModal, Me
End Sub

Private Sub MnuHistogramEqualize_Click()
    Process "Equalize", True
End Sub

Private Sub MnuHistogramStretch_Click()
    Process "Stretch histogram"
End Sub

'All top-level Image menu actions are handled here
Private Sub MnuImage_Click(Index As Integer)

    Select Case Index
    
        'Duplicate
        Case 0
        
            'It may seem odd, but the Duplicate function can be found in the "Loading" module; I do this because
            ' we effectively LOAD a copy of the original image, so all loading operations (create a form, catalog
            ' metadata, initialize properties) have to be repeated.
            DuplicateCurrentImage
        
        'Separator
        Case 1
        
        'Transparency top-level
        Case 2
        
        'Separator
        Case 3
        
        'Resize
        Case 4
            Process "Resize", True
        
        'Canvas resize
        Case 5
            Process "Canvas size", True
                
        'Separator
        Case 6
            
        'Crop to selection
        Case 7
            Process "Crop"
        
        'Autocrop
        Case 8
            Process "Autocrop"
        
        'Separator
        Case 9
        
        'Top-level Rotate
        Case 10
        
        'Flip horizontal (mirror)
        Case 11
            Process "Flip horizontal"
        
        'Flip vertical
        Case 12
            Process "Flip vertical"
        
        'Isometric view
        Case 13
            Process "Isometric conversion"
            
        'Separator
        Case 14
        
        'Indexed color
        Case 15
            Process "Reduce colors", True
        
        'Tile
        Case 16
            Process "Tile", True
            
        'Separator
        Case 17
        
        'Metadata top-level
        Case 18
    
    End Select

End Sub

'This is the exact same thing as "Paste as New Image".  It is provided in two locations for convenience.
Private Sub MnuImportClipboard_Click()
    Process "Paste as new image", , , False
End Sub

'Attempt to import an image from the Internet
Private Sub MnuImportFromInternet_Click()
    If FormInternetImport.Visible = False Then FormInternetImport.Show vbModal, FormMain
End Sub

Private Sub MnuAlien_Click()
    Process "Alien"
End Sub

Private Sub MnuInvertHue_Click()
    Process "Invert hue"
End Sub

'When a language is clicked, immediately activate it
Private Sub mnuLanguages_Click(Index As Integer)

    Screen.MousePointer = vbHourglass
    
    'Because loading a language can take some time, display a wait screen to discourage attempted interaction
    displayWaitScreen g_Language.TranslateMessage("Please wait while the new language is applied..."), Me
    
    'Remove the existing translation
    Message "Removing existing translation..."
    g_Language.undoTranslations FormMain, True
    
    'Apply the new translation
    Message "Applying new translation..."
    g_Language.activateNewLanguage Index, True
    
    Message "Language changed successfully."
    
    hideWaitScreen
    
    Screen.MousePointer = vbDefault
    
End Sub

'The user can toggle the appearance of the left-hand panel from this menu.  This toggle is also stored in the preferences file.
Private Sub MnuLeftPanel_Click()
    ChangeLeftPane VISIBILITY_TOGGLE
End Sub

'All metadata sub-menu options are handled here
Private Sub MnuMetadata_Click(Index As Integer)

    Select Case Index
    
        'Browse metadata
        Case 0
        
            'Before doing anything else, see if we've already loaded metadata.  If we haven't, do so now.
            If Not pdImages(CurrentImage).imgMetadata.hasXMLMetadata Then
                pdImages(CurrentImage).imgMetadata.loadAllMetadata pdImages(CurrentImage).LocationOnDisk, pdImages(CurrentImage).OriginalFileFormat
                
                'If the image contains GPS metadata, enable that option now
                metaToggle tGPSMetadata, pdImages(CurrentImage).imgMetadata.hasGPSMetadata()
            End If
            
            'If the image STILL doesn't have metadata, warn the user and exit.
            If Not pdImages(CurrentImage).imgMetadata.hasXMLMetadata Then
                Message "No metadata available."
                pdMsgBox "This image does not contain any metadata.", vbInformation + vbOKOnly + vbApplicationModal, "No metadata available"
                Exit Sub
            End If
            
            FormMetadata.Show vbModal, Me
        
        'Separator
        Case 1
        
        'Count colors
        Case 2
            Process "Count image colors", True
        
        'Map photo location
        Case 3
            
            'Note that mapping can only be performed if GPS metadata exists for this image.  If the user clicks this option while
            ' using the on-demand model for metadata caching, we must now attempt to load metadata.
            If Not pdImages(CurrentImage).imgMetadata.hasXMLMetadata Then
            
                'Attempt to load it now...
                Message "Loading metadata for this image..."
                pdImages(CurrentImage).imgMetadata.loadAllMetadata pdImages(CurrentImage).LocationOnDisk, pdImages(CurrentImage).OriginalFileFormat
                
                'Determine whether metadata is present, and dis/enable metadata menu items accordingly
                metaToggle tMetadata, pdImages(CurrentImage).imgMetadata.hasXMLMetadata
                metaToggle tGPSMetadata, pdImages(CurrentImage).imgMetadata.hasGPSMetadata()
            
            End If
            
            If Not pdImages(CurrentImage).imgMetadata.hasGPSMetadata Then
                pdMsgBox "This image does not contain any GPS metadata.", vbOKOnly + vbApplicationModal + vbInformation, "No GPS data found"
                Exit Sub
            End If
            
            Dim gMapsURL As String, latString As String, lonString As String
            If pdImages(CurrentImage).imgMetadata.fillLatitudeLongitude(latString, lonString) Then
                
                'Build a valid Google maps URL (you can use Google to see what the various parameters mean)
                
                'Note: I find a zoom of 18 ideal, as that is a common level for switching to an "aerial"
                ' view instead of a satellite view.  Much higher than that and you run the risk of not
                ' having data available at that high of zoom.
                gMapsURL = "https://maps.google.com/maps?f=q&z=18&t=h&q=" & latString & "%2c+" & lonString
                
                'As a convenience, request Google Maps in the current language
                If g_Language.translationActive Then
                    gMapsURL = gMapsURL & "&hl=" & g_Language.getCurrentLanguage()
                Else
                    gMapsURL = gMapsURL & "&hl=en"
                End If
                
                'Launch Google maps in the user's browser
                OpenURL gMapsURL
                
            End If
            
    End Select
    
End Sub

Private Sub MnuMonochrome_Click(Index As Integer)
    
    Select Case Index
        
        'Convert color to monochrome
        Case 0
            Process "Color to monochrome", True
        
        'Convert monochrome to grayscale
        Case 1
            Process "Monochrome to grayscale", True
        
    End Select
    
End Sub

Private Sub MnuNatureFilter_Click(Index As Integer)

    Select Case Index
    
        'Atmosphere
        Case 0
            Process "Atmosphere"
            
        'Burn
        Case 1
            Process "Burn"
        
        'Fog
        Case 2
            Process "Fog"
        
        'Freeze
        Case 3
            Process "Freeze"
        
        'Lava
        Case 4
            Process "Lava"
                
        'Rainbow
        Case 5
            Process "Rainbow"
        
        'Steel
        Case 6
            Process "Steel"
        
        'Water
        Case 7
            Process "Water"
    
    End Select

End Sub

Private Sub MnuNegative_Click()
    Process "Film negative"
End Sub

Private Sub MnuCopy_Click()
    Process "Copy to clipboard", , , False
End Sub

Private Sub MnuEmptyClipboard_Click()
    Process "Empty clipboard", , , False
End Sub

Private Sub MnuExit_Click()
    Unload FormMain
End Sub

Private Sub MnuInvert_Click()
    Process "Invert RGB"
End Sub

'All noise filters are handled here
Private Sub MnuNoise_Click(Index As Integer)

    Select Case Index
    
        'Film grain
        Case 0
            Process "Add film grain", True
        
        'RGB Noise
        Case 1
            Process "Add RGB noise", True
            
        'Separator
        Case 2
        
        'Median
        Case 3
            Process "Median", True
            
    End Select
        
End Sub

Private Sub MnuOpen_Click()
    Process "Open", True
End Sub

Private Sub MnuPaste_Click()
    Process "Paste as new image", , , False
End Sub

Private Sub MnuPlayMacroRecording_Click()
    Process "Play macro", True
End Sub

Private Sub MnuPrint_Click()
    If Not FormPrint.Visible Then FormPrint.Show vbModal, FormMain
End Sub

Private Sub MnuRadioactive_Click()
    Process "Radioactive"
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
    Process "Redo", , , False
End Sub

Private Sub MnuRepeatLast_Click()
    Process "Repeat last action"
End Sub

Private Sub MnuRightPanel_Click()
    ChangeRightPane VISIBILITY_TOGGLE
End Sub

'All rotation actions are initiated here
Private Sub MnuRotate_Click(Index As Integer)

    Select Case Index
    
        'Rotate 90
        Case 0
            Process "Rotate 90° clockwise"
        
        'Rotate 270
        Case 1
            Process "Rotate 90° Counter-Clockwise"
        
        'Rotate 180
        Case 2
            Process "Rotate 180°"
        
        'Rotate arbitrary
        Case 3
            Process "Arbitrary rotation", True
            
    End Select
            
End Sub

Private Sub MnuSave_Click()
    Process "Save", , , False
End Sub

Private Sub MnuSaveAs_Click()
    Process "Save as", True
End Sub

Private Sub MnuScanImage_Click()
    Process "Scan image", , , False
End Sub

Private Sub MnuScreenCapture_Click()
    Process "Screen capture", , , False
End Sub

'All select menu items are handled here
Private Sub MnuSelect_Click(Index As Integer)

    Select Case Index
    
        'Select all.  (Note that Square Selection is passed as the relevant tool for this action.)
        Case 0
            Process "Select all", , , 2, 0
        
        'Select none
        Case 1
            Process "Remove selection", , pdImages(CurrentImage).mainSelection.getSelectionParamString, 2
        
        'Invert
        Case 2
            Process "Invert selection", , , 2
        
        '<separator>
        Case 3
        
        'Grow selection
        Case 4
            Process "Grow selection", True, , 0
        
        'Shrink selection
        Case 5
            Process "Shrink selection", True, , 0
        
        'Border selection
        Case 6
            Process "Border selection", True, , 0
        
        'Feather selection
        Case 7
            Process "Feather selection", True, , 0
        
        'Sharpen selection
        Case 8
            Process "Sharpen selection", True, , 0
        
        '<separator>
        Case 9
        
        'Load selection
        Case 10
            Process "Load selection", True, , 0
        
        'Save current selection
        Case 11
            Process "Save selection", True, , 0
        
    End Select

End Sub

Private Sub MnuSelectScanner_Click()
    Process "Select scanner or camera", True
End Sub

'All sharpen filters are handled here
Private Sub MnuSharpen_Click(Index As Integer)

    Select Case Index
            
        'Sharpen
        Case 0
            Process "Sharpen", True
        
        'Unsharp mask
        Case 1
            Process "Unsharp mask", True
            
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
    Process "Start macro recording", , , False
End Sub

Private Sub MnuStopMacroRecording_Click()
    Process "Stop macro recording", True
End Sub

'All stylize filters are handled here
Private Sub MnuStylize_Click(Index As Integer)

    Select Case Index
    
        'Antique
        Case 0
            Process "Antique"
    
        'Diffuse
        Case 1
            Process "Diffuse", True
        
        'Dilate (maximum rank)
        Case 2
            Process "Dilate (maximum rank)", True
        
        'Erode (minimum rank)
        Case 3
            Process "Erode (minimum rank)", True
        
        'Solarize
        Case 4
            Process "Solarize", True

        'Twins
        Case 5
            Process "Twins", True
            
        'Vignetting
        Case 6
            Process "Vignetting", True
    
    End Select

End Sub

Private Sub MnuSynthesize_Click()
    Process "Synthesize"
End Sub

Private Sub MnuTest_Click()
    MenuTest
End Sub

'All tool menu items are launched from here
Private Sub mnuTool_Click(Index As Integer)

    Select Case Index
    
        'Language editor
        Case 1
            If Not FormLanguageEditor.Visible Then FormLanguageEditor.Show vbModeless, FormMain
    
        'Options
        Case 5
            If Not FormPreferences.Visible Then FormPreferences.Show vbModal, FormMain
            
        'Plugin manager
        Case 6
            If Not FormPluginManager.Visible Then FormPluginManager.Show vbModal, FormMain
            
    End Select

End Sub

'Add / Remove / Modify an image's alpha channel with this menu
Private Sub MnuTransparency_Click(Index As Integer)

    Select Case Index
    
        'Add alpha channel
        Case 0
            
            'Ignore if the current image is already in 32bpp mode
            Process "Add alpha channel", True
            
        'Color to alpha
        Case 1
        
            'Can be used even if the image already has an alpha channel
            Process "Color to alpha", True
        
        '<separator>
        Case 2

        'Remove alpha channel
        Case 3

            'Ignore if the current image is already in 24bpp mode
            If pdImages(CurrentImage).mainLayer.getLayerColorDepth = 24 Then Exit Sub
            Process "Remove alpha channel", True
    
    End Select

End Sub

Private Sub MnuUndo_Click()
    Process "Undo", , , False
End Sub

Private Sub MnuVibrate_Click()
    Process "Vibrate"
End Sub

'Because VB doesn't allow key tracking in MDIForms, we have to hook keypresses via this method.
' Many thanks to Steve McMahon for the usercontrol that helps implement this
Private Sub ctlAccelerator_Accelerator(ByVal nIndex As Long, bCancel As Boolean)

    'Don't process accelerators when the main form is disabled (e.g. if a modal form is present, or if a previous
    ' action is in the middle of execution)
    If Not FormMain.Enabled Then Exit Sub

    'Accelerators can be fired multiple times by accident.  Don't allow the user to press accelerators
    ' faster than one quarter-second apart.
    Static lastAccelerator As Double
    
    If Timer - lastAccelerator < 0.25 Then Exit Sub

    'Accelerators are divided into two groups: those that can be fired if no images are present (e.g. Open, Paste),
    ' and those that require an image.

    '***********************************************************
    'Accelerators that DO NOT require at least one loaded image:

    'Import from Internet
    If ctlAccelerator.Key(nIndex) = "Internet_Import" Then
        If Not FormInternetImport.Visible Then FormInternetImport.Show vbModal, FormMain
    End If
    
    'Capture the screen
    If ctlAccelerator.Key(nIndex) = "Screen_Capture" Then Process "Screen capture", , , False
    
    'Open program preferences
    If ctlAccelerator.Key(nIndex) = "Preferences" Then
        If Not FormPreferences.Visible Then FormPreferences.Show vbModal, FormMain
    End If
    
    'Empty clipboard
    If ctlAccelerator.Key(nIndex) = "Empty_Clipboard" Then Process "Empty clipboard", , , False
    
    'Escape - a separate function is used to cancel currently running filters.  This accelerator is only used
    ' to cancel batch conversions, but in the future it should be applied elsewhere.
    If ctlAccelerator.Key(nIndex) = "Escape" Then
        If MacroStatus = MacroBATCH Then MacroStatus = MacroCANCEL
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
    
    '***********************************************************
    'Accelerators that DO require at least one loaded image:
    
    'If no images are loaded, or another form is active, exit.
    If FormLanguageEditor.Visible Then Exit Sub
    
    If NumOfWindows = 0 Then Exit Sub
    
    
    'Save As...
    If ctlAccelerator.Key(nIndex) = "Save_As" Then
        If FormMain.MnuSaveAs.Enabled Then Process "Save as", True
    End If
    
    'Redo
    If ctlAccelerator.Key(nIndex) = "Redo" Then
        If FormMain.MnuRedo.Enabled Then Process "Redo", , , False
    End If
    
    'Fit on screen
    If ctlAccelerator.Key(nIndex) = "FitOnScreen" Then FitOnScreen
    
    'Zoom in
    If ctlAccelerator.Key(nIndex) = "Zoom_In" Then
        If FormMain.CmbZoom.Enabled And FormMain.CmbZoom.ListIndex > 0 Then FormMain.CmbZoom.ListIndex = FormMain.CmbZoom.ListIndex - 1
    End If
    
    'Zoom out
    If ctlAccelerator.Key(nIndex) = "Zoom_Out" Then
        If FormMain.CmbZoom.Enabled And FormMain.CmbZoom.ListIndex < (FormMain.CmbZoom.ListCount - 1) Then FormMain.CmbZoom.ListIndex = FormMain.CmbZoom.ListIndex + 1
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
    
    'Brightness/Contrast
    If ctlAccelerator.Key(nIndex) = "Bright_Contrast" Then
        Process "Brightness and contrast", True
    End If
    
    'Color balance
    If ctlAccelerator.Key(nIndex) = "Color_Balance" Then
        Process "Color balance", True
    End If
    
    'Shadows / Highlights
    If ctlAccelerator.Key(nIndex) = "Shadow_Highlight" Then
        Process "Shadows and highlights", True
    End If
    
    'Rotate Right / Left
    If ctlAccelerator.Key(nIndex) = "Rotate_Left" Then Process "Rotate 90° counter-clockwise"
    If ctlAccelerator.Key(nIndex) = "Rotate_Right" Then Process "Rotate 90° clockwise"
    
    'Crop to selection
    If ctlAccelerator.Key(nIndex) = "Crop_Selection" Then
        If pdImages(CurrentImage).selectionActive Then Process "Crop"
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
                    If pdImages(i).IsActive Then
                        pdImages(i).containingForm.WindowState = vbNormal
                        PrepareViewport pdImages(i).containingForm, "Restore all windows"
                    End If
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
    If selectionsAllowed(True) Then
        pdImages(FormMain.ActiveForm.Tag).mainSelection.setRoundedCornerAmount sltCornerRounding.Value
        RenderViewport FormMain.ActiveForm
    End If
End Sub

Private Sub sltSelectionBorder_Change()
    If selectionsAllowed(False) Then
        pdImages(FormMain.ActiveForm.Tag).mainSelection.setBorderSize sltSelectionBorder.Value
        RenderViewport FormMain.ActiveForm
    End If
End Sub

Private Sub sltSelectionFeathering_Change()
    If selectionsAllowed(False) Then
        pdImages(FormMain.ActiveForm.Tag).mainSelection.setFeatheringRadius sltSelectionFeathering.Value
        RenderViewport FormMain.ActiveForm
    End If
End Sub

Private Sub sltSelectionLineWidth_Change()
    If selectionsAllowed(True) Then
        pdImages(FormMain.ActiveForm.Tag).mainSelection.setSelectionLineWidth sltSelectionLineWidth.Value
        RenderViewport FormMain.ActiveForm
    End If
End Sub

Private Function selectionsAllowed(ByVal transformableMatters As Boolean) As Boolean
    If NumOfWindows > 0 Then
        If pdImages(FormMain.ActiveForm.Tag).selectionActive And (Not pdImages(FormMain.ActiveForm.Tag).mainSelection Is Nothing) And (Not pdImages(FormMain.ActiveForm.Tag).mainSelection.rejectRefreshRequests) Then
            
            If transformableMatters Then
                If pdImages(FormMain.ActiveForm.Tag).mainSelection.isTransformable Then
                    selectionsAllowed = True
                Else
                    selectionsAllowed = False
                End If
            Else
                selectionsAllowed = True
            End If
            
        Else
            selectionsAllowed = False
        End If
    Else
        selectionsAllowed = False
    End If
End Function

'Because we want tooltips preserved, outside functions should use THIS sub to request FormMain rethemes
Public Sub requestMakeFormPretty(Optional ByVal useDoEvents As Boolean = False)
    makeFormPretty Me, m_ToolTip, , useDoEvents
End Sub

'When certain selection settings are enabled or disabled, corresponding controls are shown or hidden.  To keep the
' panel concise and clean, we move other controls up or down depending on what controls are visible.
Private Sub updateSelectionPanelLayout()

    'Display the feathering slider as necessary
    If cmbSelSmoothing(0).ListIndex = sFullyFeathered Then
        sltSelectionFeathering.Visible = True
        lblSelection(4).Top = sltSelectionFeathering.Top + fixDPI(38)
    Else
        sltSelectionFeathering.Visible = False
        lblSelection(4).Top = cmbSelSmoothing(0).Top + fixDPI(34)
    End If
    cmbSelType(0).Top = lblSelection(4).Top + fixDPI(24)
    sltSelectionBorder.Top = cmbSelType(0).Top + fixDPI(26)

    'Display the border slider as necessary
    If cmbSelType(0).ListIndex = sBorder Then
        sltSelectionBorder.Visible = True
        lblSelection(5).Top = sltSelectionBorder.Top + fixDPI(38)
    Else
        sltSelectionBorder.Visible = False
        lblSelection(5).Top = cmbSelType(0).Top + fixDPI(34)
    End If
    sltCornerRounding.Top = lblSelection(5).Top + fixDPI(24)
    sltSelectionLineWidth.Top = lblSelection(5).Top + fixDPI(24)

End Sub

'When the selection text boxes are updated, change the scrollbars to match
Private Sub tudSel_Change(Index As Integer)
    updateSelectionsValuesViaText
End Sub

Private Sub updateSelectionsValuesViaText()
    If selectionsAllowed(True) Then
        If Not pdImages(FormMain.ActiveForm.Tag).mainSelection.rejectRefreshRequests Then
            pdImages(FormMain.ActiveForm.Tag).mainSelection.updateViaTextBox
            RenderViewport FormMain.ActiveForm
        End If
    End If
End Sub
