VERSION 5.00
Begin VB.Form toolbar_Tools 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Tools"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   14205
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
   ScaleHeight     =   157
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   947
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.jcbutton cmdTools 
      Height          =   600
      Index           =   2
      Left            =   1950
      TabIndex        =   16
      Top             =   315
      Width           =   720
      _extentx        =   1270
      _extenty        =   1058
      buttonstyle     =   7
      font            =   "VBP_ToolbarTools.frx":0000
      backcolor       =   -2147483643
      caption         =   ""
      handpointer     =   -1
      picturenormal   =   "VBP_ToolbarTools.frx":0028
      pictureeffectondown=   0
      captioneffects  =   0
      mode            =   1
      colorscheme     =   3
   End
   Begin PhotoDemon.jcbutton cmdTools 
      Height          =   600
      Index           =   3
      Left            =   2670
      TabIndex        =   17
      Top             =   315
      Width           =   720
      _extentx        =   1270
      _extenty        =   1058
      buttonstyle     =   7
      font            =   "VBP_ToolbarTools.frx":0C0A
      backcolor       =   -2147483643
      caption         =   ""
      handpointer     =   -1
      picturenormal   =   "VBP_ToolbarTools.frx":0C32
      pictureeffectondown=   0
      captioneffects  =   0
      mode            =   1
      colorscheme     =   3
   End
   Begin PhotoDemon.jcbutton cmdTools 
      Height          =   600
      Index           =   4
      Left            =   3390
      TabIndex        =   18
      Top             =   315
      Width           =   720
      _extentx        =   1270
      _extenty        =   1058
      buttonstyle     =   7
      font            =   "VBP_ToolbarTools.frx":1814
      backcolor       =   -2147483643
      caption         =   ""
      handpointer     =   -1
      picturenormal   =   "VBP_ToolbarTools.frx":183C
      pictureeffectondown=   0
      captioneffects  =   0
      mode            =   1
      colorscheme     =   3
   End
   Begin PhotoDemon.jcbutton cmdTools 
      Height          =   600
      Index           =   0
      Left            =   120
      TabIndex        =   23
      Top             =   315
      Width           =   720
      _extentx        =   1270
      _extenty        =   1058
      buttonstyle     =   7
      font            =   "VBP_ToolbarTools.frx":241E
      backcolor       =   -2147483643
      caption         =   ""
      handpointer     =   -1
      picturenormal   =   "VBP_ToolbarTools.frx":2446
      pictureeffectondown=   0
      captioneffects  =   0
      mode            =   1
      colorscheme     =   3
   End
   Begin PhotoDemon.jcbutton cmdTools 
      Height          =   600
      Index           =   1
      Left            =   840
      TabIndex        =   24
      Top             =   315
      Width           =   720
      _extentx        =   1270
      _extenty        =   1058
      buttonstyle     =   7
      font            =   "VBP_ToolbarTools.frx":3198
      backcolor       =   -2147483643
      caption         =   ""
      handpointer     =   -1
      picturenormal   =   "VBP_ToolbarTools.frx":31C0
      pictureeffectondown=   0
      captioneffects  =   0
      mode            =   1
      colorscheme     =   3
   End
   Begin PhotoDemon.jcbutton cmdTools 
      Height          =   600
      Index           =   5
      Left            =   4560
      TabIndex        =   33
      Top             =   315
      Width           =   720
      _extentx        =   1270
      _extenty        =   1058
      buttonstyle     =   7
      font            =   "VBP_ToolbarTools.frx":3DA2
      backcolor       =   -2147483643
      caption         =   ""
      handpointer     =   -1
      picturenormal   =   "VBP_ToolbarTools.frx":3DCA
      pictureeffectondown=   0
      captioneffects  =   0
      mode            =   1
      colorscheme     =   3
   End
   Begin VB.PictureBox picTools 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1575
      Index           =   2
      Left            =   0
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   950
      TabIndex        =   34
      Top             =   1020
      Visible         =   0   'False
      Width           =   14250
      Begin PhotoDemon.sliderTextCombo sltQuickFix 
         CausesValidation=   0   'False
         Height          =   495
         Index           =   0
         Left            =   1380
         TabIndex        =   36
         Top             =   90
         Width           =   2670
         _extentx        =   4710
         _extenty        =   873
         font            =   "VBP_ToolbarTools.frx":4B1C
         min             =   -2
         max             =   2
         sigdigits       =   2
         slidertrackstyle=   2
      End
      Begin PhotoDemon.sliderTextCombo sltQuickFix 
         CausesValidation=   0   'False
         Height          =   495
         Index           =   1
         Left            =   1380
         TabIndex        =   37
         Top             =   705
         Width           =   2670
         _extentx        =   4710
         _extenty        =   873
         font            =   "VBP_ToolbarTools.frx":4B44
         min             =   -100
         max             =   100
      End
      Begin PhotoDemon.sliderTextCombo sltQuickFix 
         CausesValidation=   0   'False
         Height          =   495
         Index           =   2
         Left            =   5640
         TabIndex        =   39
         Top             =   90
         Width           =   2670
         _extentx        =   4710
         _extenty        =   873
         font            =   "VBP_ToolbarTools.frx":4B6C
         min             =   -100
         max             =   100
      End
      Begin PhotoDemon.sliderTextCombo sltQuickFix 
         CausesValidation=   0   'False
         Height          =   495
         Index           =   3
         Left            =   5640
         TabIndex        =   41
         Top             =   705
         Width           =   2670
         _extentx        =   4710
         _extenty        =   873
         font            =   "VBP_ToolbarTools.frx":4B94
         min             =   -100
         max             =   100
      End
      Begin PhotoDemon.jcbutton cmdQuickFix 
         Height          =   570
         Index           =   0
         Left            =   13080
         TabIndex        =   43
         Top             =   75
         Width           =   660
         _extentx        =   1164
         _extenty        =   1005
         buttonstyle     =   13
         font            =   "VBP_ToolbarTools.frx":4BBC
         backcolor       =   -2147483643
         caption         =   ""
         handpointer     =   -1
         picturenormal   =   "VBP_ToolbarTools.frx":4BE4
         pictureeffectondown=   0
         captioneffects  =   0
         colorscheme     =   3
      End
      Begin PhotoDemon.jcbutton cmdQuickFix 
         Height          =   570
         Index           =   1
         Left            =   13080
         TabIndex        =   44
         Top             =   705
         Width           =   660
         _extentx        =   1164
         _extenty        =   1005
         buttonstyle     =   13
         font            =   "VBP_ToolbarTools.frx":5936
         backcolor       =   -2147483643
         caption         =   ""
         handpointer     =   -1
         picturenormal   =   "VBP_ToolbarTools.frx":595E
         pictureeffectondown=   0
         captioneffects  =   0
         colorscheme     =   3
      End
      Begin PhotoDemon.sliderTextCombo sltQuickFix 
         CausesValidation=   0   'False
         Height          =   495
         Index           =   4
         Left            =   9960
         TabIndex        =   45
         Top             =   90
         Width           =   2670
         _extentx        =   4710
         _extenty        =   873
         font            =   "VBP_ToolbarTools.frx":66B0
         min             =   -100
         max             =   100
         slidertrackstyle=   3
         gradientcolorleft=   16752699
         gradientcolorright=   2990335
         gradientcolormiddle=   16777215
         gradientmiddlevalue=   0
      End
      Begin PhotoDemon.sliderTextCombo sltQuickFix 
         CausesValidation=   0   'False
         Height          =   495
         Index           =   5
         Left            =   9960
         TabIndex        =   46
         Top             =   705
         Width           =   2670
         _extentx        =   4710
         _extenty        =   873
         font            =   "VBP_ToolbarTools.frx":66D8
         min             =   -100
         max             =   100
         slidertrackstyle=   3
         gradientcolorleft=   15102446
         gradientcolorright=   8253041
         gradientcolormiddle=   16777215
         gradientmiddlevalue=   0
      End
      Begin VB.Label lblOptions 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "temperature:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   240
         Index           =   7
         Left            =   8775
         TabIndex        =   48
         Top             =   195
         Width           =   1140
      End
      Begin VB.Label lblOptions 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "tint:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   240
         Index           =   6
         Left            =   9570
         TabIndex        =   47
         Top             =   810
         Width           =   345
      End
      Begin VB.Label lblOptions 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "vibrance:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   240
         Index           =   5
         Left            =   4800
         TabIndex        =   42
         Top             =   810
         Width           =   795
      End
      Begin VB.Label lblOptions 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "clarity:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   240
         Index           =   4
         Left            =   5010
         TabIndex        =   40
         Top             =   195
         Width           =   585
      End
      Begin VB.Label lblOptions 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "contrast:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   240
         Index           =   3
         Left            =   570
         TabIndex        =   38
         Top             =   810
         Width           =   765
      End
      Begin VB.Label lblOptions 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "exposure:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   240
         Index           =   2
         Left            =   480
         TabIndex        =   35
         Top             =   195
         Width           =   855
      End
   End
   Begin VB.PictureBox picTools 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1575
      Index           =   1
      Left            =   15
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   950
      TabIndex        =   25
      Top             =   1020
      Visible         =   0   'False
      Width           =   14250
      Begin PhotoDemon.smartCheckBox chkLayerBorder 
         Height          =   480
         Left            =   6480
         TabIndex        =   27
         Top             =   360
         Width           =   2025
         _extentx        =   3572
         _extenty        =   847
         caption         =   "show layer borders"
         font            =   "VBP_ToolbarTools.frx":6700
         value           =   1
      End
      Begin PhotoDemon.smartCheckBox chkLayerNodes 
         Height          =   480
         Left            =   6480
         TabIndex        =   28
         Top             =   780
         Width           =   2775
         _extentx        =   4895
         _extenty        =   847
         caption         =   "show layer transform nodes"
         font            =   "VBP_ToolbarTools.frx":6728
         value           =   1
      End
      Begin PhotoDemon.smartCheckBox chkAutoActivateLayer 
         Height          =   480
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   4080
         _extentx        =   7197
         _extenty        =   847
         caption         =   "automatically activate layer beneath mouse"
         font            =   "VBP_ToolbarTools.frx":6750
         value           =   1
      End
      Begin PhotoDemon.smartCheckBox chkIgnoreTransparent 
         Height          =   480
         Left            =   120
         TabIndex        =   31
         Top             =   780
         Width           =   4920
         _extentx        =   8678
         _extenty        =   847
         caption         =   "ignore transparent pixels when auto-activating layers"
         font            =   "VBP_ToolbarTools.frx":6778
         value           =   1
      End
      Begin VB.Label lblOptions 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "interaction options:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   29
         Top             =   60
         Width           =   1650
      End
      Begin VB.Label lblOptions 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "display options:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   240
         Index           =   0
         Left            =   6480
         TabIndex        =   26
         Top             =   60
         Width           =   1335
      End
   End
   Begin VB.PictureBox picTools 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1575
      Index           =   0
      Left            =   15
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   950
      TabIndex        =   0
      Top             =   1020
      Visible         =   0   'False
      Width           =   14250
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
         ItemData        =   "VBP_ToolbarTools.frx":67A0
         Left            =   120
         List            =   "VBP_ToolbarTools.frx":67A2
         Style           =   2  'Dropdown List
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   390
         Width           =   2250
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
         ItemData        =   "VBP_ToolbarTools.frx":67A4
         Left            =   8340
         List            =   "VBP_ToolbarTools.frx":67A6
         Style           =   2  'Dropdown List
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "This option controls the selection's area.  You can switch between the three settings without losing the current selection."
         Top             =   390
         Width           =   2445
      End
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
         ItemData        =   "VBP_ToolbarTools.frx":67A8
         Left            =   5640
         List            =   "VBP_ToolbarTools.frx":67AA
         Style           =   2  'Dropdown List
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Use this option to change the way selections blend with their surroundings."
         Top             =   390
         Width           =   2445
      End
      Begin PhotoDemon.sliderTextCombo sltCornerRounding 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   10860
         TabIndex        =   3
         Top             =   345
         Width           =   2670
         _extentx        =   4710
         _extenty        =   873
         font            =   "VBP_ToolbarTools.frx":67AC
         max             =   10000
      End
      Begin PhotoDemon.textUpDown tudSel 
         Height          =   405
         Index           =   0
         Left            =   2520
         TabIndex        =   4
         Top             =   390
         Width           =   1320
         _extentx        =   2328
         _extenty        =   714
         font            =   "VBP_ToolbarTools.frx":67D4
         min             =   -30000
         max             =   30000
      End
      Begin PhotoDemon.textUpDown tudSel 
         Height          =   405
         Index           =   1
         Left            =   2520
         TabIndex        =   5
         Top             =   840
         Width           =   1320
         _extentx        =   2328
         _extenty        =   714
         font            =   "VBP_ToolbarTools.frx":67FC
         min             =   -30000
         max             =   30000
      End
      Begin PhotoDemon.textUpDown tudSel 
         Height          =   405
         Index           =   2
         Left            =   4080
         TabIndex        =   6
         Top             =   390
         Width           =   1320
         _extentx        =   2328
         _extenty        =   714
         font            =   "VBP_ToolbarTools.frx":6824
         min             =   -30000
         max             =   30000
      End
      Begin PhotoDemon.textUpDown tudSel 
         Height          =   405
         Index           =   3
         Left            =   4080
         TabIndex        =   7
         Top             =   840
         Width           =   1320
         _extentx        =   2328
         _extenty        =   714
         font            =   "VBP_ToolbarTools.frx":684C
         min             =   -30000
         max             =   30000
      End
      Begin PhotoDemon.sliderTextCombo sltSelectionBorder 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   8220
         TabIndex        =   8
         Top             =   840
         Visible         =   0   'False
         Width           =   2670
         _extentx        =   4710
         _extenty        =   873
         font            =   "VBP_ToolbarTools.frx":6874
         min             =   1
         max             =   10000
         value           =   1
      End
      Begin PhotoDemon.sliderTextCombo sltSelectionFeathering 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   5520
         TabIndex        =   9
         Top             =   840
         Width           =   2670
         _extentx        =   4710
         _extenty        =   873
         font            =   "VBP_ToolbarTools.frx":689C
         max             =   100
      End
      Begin PhotoDemon.sliderTextCombo sltSelectionLineWidth 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   10860
         TabIndex        =   10
         Top             =   345
         Width           =   2670
         _extentx        =   4710
         _extenty        =   873
         font            =   "VBP_ToolbarTools.frx":68C4
         min             =   1
         max             =   10000
         value           =   10
      End
      Begin VB.Label lblSelection 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "appearance:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   60
         Width           =   1080
      End
      Begin VB.Label lblSelection 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "size (w, h)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   240
         Index           =   2
         Left            =   4080
         TabIndex        =   15
         Top             =   60
         Width           =   915
      End
      Begin VB.Label lblSelection 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "position (x, y)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   240
         Index           =   1
         Left            =   2520
         TabIndex        =   14
         Top             =   60
         Width           =   1170
      End
      Begin VB.Label lblSelection 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "corner rounding"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   240
         Index           =   5
         Left            =   10980
         TabIndex        =   13
         Top             =   60
         Width           =   1365
      End
      Begin VB.Label lblSelection 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "area"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   240
         Index           =   4
         Left            =   8340
         TabIndex        =   12
         Top             =   60
         Width           =   390
      End
      Begin VB.Label lblSelection 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "smoothing"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00606060&
         Height          =   240
         Index           =   3
         Left            =   5640
         TabIndex        =   11
         Top             =   60
         Width           =   885
      End
   End
   Begin VB.Label lblCategory 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "quick fix"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00606060&
      Height          =   240
      Index           =   2
      Left            =   4560
      TabIndex        =   32
      Top             =   30
      Width           =   690
   End
   Begin VB.Label lblCategory 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "nav"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00606060&
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   22
      Top             =   30
      Width           =   300
   End
   Begin VB.Line lineMain 
      BorderColor     =   &H80000002&
      Index           =   2
      X1              =   0
      X2              =   5000
      Y1              =   67
      Y2              =   67
   End
   Begin VB.Label lblCategory 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "selections"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00606060&
      Height          =   240
      Index           =   1
      Left            =   1920
      TabIndex        =   21
      Top             =   30
      Width           =   840
   End
   Begin VB.Line lineMain 
      BorderColor     =   &H80000002&
      Index           =   1
      X1              =   0
      X2              =   5000
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line lineMain 
      BorderColor     =   &H80000002&
      Index           =   0
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   2000
   End
End
Attribute VB_Name = "toolbar_Tools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Tools Toolbox
'Copyright ©2013-2014 by Tanner Helland
'Created: 03/October/13
'Last updated: 26/June/14
'Last update: add temperature and tint to the Quick Fix tool selection; minor UI adjustments
'
'This form was initially integrated into the main MDI form.  In fall 2013, PhotoDemon left behind the MDI model,
' and all toolbars were moved to their own forms.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'Used to toggle the command button state of the toolbar buttons
Private Const BM_SETSTATE = &HF3
Private Declare Function SendMessageA Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'The value of all controls on this form are saved and loaded to file by this class
Private WithEvents lastUsedSettings As pdLastUsedSettings
Attribute lastUsedSettings.VB_VarHelpID = -1

'Whether or not non-destructive FX can be applied to the image
Private m_NonDestructiveFXAllowed As Boolean

'If external functions want to disable automatic non-destructive FX syncing, then can do so via this function
Public Sub setNDFXControlState(ByVal newNDFXState As Boolean)
    m_NonDestructiveFXAllowed = newNDFXState
End Sub

Private Sub chkAutoActivateLayer_Click()
    If CBool(chkAutoActivateLayer) Then
        If Not chkIgnoreTransparent.Enabled Then chkIgnoreTransparent.Enabled = True
    Else
        If chkIgnoreTransparent.Enabled Then chkIgnoreTransparent.Enabled = False
    End If
End Sub

'Show/hide layer borders while using the move tool
Private Sub chkLayerBorder_Click()
    PrepareViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0), "Layer border toggle"
End Sub

'Show/hide layer transform nodes while using the move tool
Private Sub chkLayerNodes_Click()
    PrepareViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0), "Layer nodes toggle"
End Sub

Private Sub cmdQuickFix_Click(Index As Integer)

    'Do nothing unless an image has been loaded
    If pdImages(g_CurrentImage) Is Nothing Then Exit Sub
    If Not pdImages(g_CurrentImage).loadedSuccessfully Then Exit Sub

    Dim i As Long

    'Regardless of the action we're applying, we start by disabling all auto-refreshes
    setNDFXControlState False
    
    Select Case Index
    
        'Reset quick-fix settings
        Case 0
            
            'Resetting does not affect the Undo/Redo chain, so simply reset all sliders, then redraw the screen
            For i = 0 To sltQuickFix.Count - 1
                
                sltQuickFix(i).Value = 0
                pdImages(g_CurrentImage).getActiveLayer.setLayerNonDestructiveFXState i, 0
                
            Next i
            
        'Make quick-fix settings permanent
        Case 1
            
            'First, make sure at least one or more quick-fixes are active
            If pdImages(g_CurrentImage).getActiveLayer.getLayerNonDestructiveFXState Then
                
                'Back-up the current quick-fix settings (because they will be reset after being applied to the image)
                evaluateImageCheckpoint
                
                'Now we do something odd; we reset all sliders, then forcibly set an image checkpoint.  This prevents PD's internal
                ' processor from auto-detecting the slider resets and applying *another* entry to the Undo/Redo chain.
                For i = 0 To sltQuickFix.Count - 1
                    sltQuickFix(i).Value = 0
                Next i
                
                setImageCheckpoint
                
                'Ask the central processor to permanently apply the quick-fix changes
                Process "Make quick-fixes permanent", , , UNDO_LAYER
                                
            End If
    
    End Select
    
    'After one of these buttons has been used, all quick-fix values will be reset - so we can disable the buttons accordingly.
    For i = 0 To cmdQuickFix.Count - 1
        If cmdQuickFix(i).Enabled Then cmdQuickFix(i).Enabled = False
    Next i
    
    'Re-enable auto-refreshes
    setNDFXControlState True
    
    'Redraw the viewport
    ScrollViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)

End Sub

Private Sub cmdTools_Click(Index As Integer)
    
    'Before changing to the new tool, see if the previously active layer has had any non-destructive changes made.
    If Processor.evaluateImageCheckpoint() Then syncInterfaceToCurrentImage
    
    'Update the previous and current tool entries
    g_PreviousTool = g_CurrentTool
    g_CurrentTool = Index
    
    'Update the tool options area to match the newly selected tool
    resetToolButtonStates
    
    'Set a new image checkpoint (necessary to do this manually, as we haven't invoked PD's central processor)
    Processor.setImageCheckpoint
        
End Sub

Private Sub Form_Load()

    Dim i As Long
    
    'Because line controls aren't automatically made DPI-aware by VB, we must manually move this dialog's line
    ' control into place.
    lineMain(2).y1 = picTools(0).Top - fixDPI(2)
    lineMain(2).y2 = lineMain(2).y1
    
    'INITIALIZE ALL TOOLS
    
        'Tool button tooltips
        cmdTools(NAV_DRAG).ToolTip = g_Language.TranslateMessage("Hand (click-and-drag image scrolling)")
        cmdTools(NAV_MOVE).ToolTip = g_Language.TranslateMessage("Move and resize image layers")
        cmdTools(SELECT_RECT).ToolTip = g_Language.TranslateMessage("Rectangular Selection")
        cmdTools(SELECT_CIRC).ToolTip = g_Language.TranslateMessage("Elliptical (Oval) Selection")
        cmdTools(SELECT_LINE).ToolTip = g_Language.TranslateMessage("Line Selection")
        cmdTools(QUICK_FIX_LIGHTING).ToolTip = g_Language.TranslateMessage("Apply non-destructive lighting adjustments")
    
        'Selection visual styles (currently lightbox or highlight)
        toolbar_Tools.cmbSelRender(0).ToolTipText = g_Language.TranslateMessage("Click to change the way selections are rendered onto the image canvas.  This has no bearing on selection contents - only the way they appear while editing.")
        For i = 0 To toolbar_Tools.cmbSelRender.Count - 1
            toolbar_Tools.cmbSelRender(i).AddItem "Lightbox", 0
            toolbar_Tools.cmbSelRender(i).AddItem "Highlight (Blue)", 1
            toolbar_Tools.cmbSelRender(i).AddItem "Highlight (Red)", 2
            toolbar_Tools.cmbSelRender(i).ListIndex = 0
        Next i
        
        'Selection smoothing (currently none, antialiased, fully feathered)
        toolbar_Tools.cmbSelSmoothing(0).ToolTipText = g_Language.TranslateMessage("This option controls how smoothly a selection blends with its surroundings.")
        toolbar_Tools.cmbSelSmoothing(0).AddItem "None", 0
        toolbar_Tools.cmbSelSmoothing(0).AddItem "Antialiased", 1
        
        'Previously, live feathering was disallowed on XP or Vista for performance reasons (GDI+ can't be used to blur
        ' the selection mask, and our own code was too slow).  As of 17 Oct '13, I have reinstated live selection
        ' feathering on these OSes using PD's very fast horizontal and vertical blur.  While not perfect, this should
        ' still provide "good enough" performance for smaller images and/or slight feathering.
        toolbar_Tools.cmbSelSmoothing(0).AddItem "Feathered", 2
        toolbar_Tools.cmbSelSmoothing(0).ListIndex = 1
        
        'Selection types (currently interior, exterior, border)
        toolbar_Tools.cmbSelType(0).ToolTipText = g_Language.TranslateMessage("These options control the area affected by a selection.  The selection can be modified on-canvas while any of these settings are active.  For more advanced selection adjustments, use the Select menu.")
        toolbar_Tools.cmbSelType(0).AddItem "Interior", 0
        toolbar_Tools.cmbSelType(0).AddItem "Exterior", 1
        toolbar_Tools.cmbSelType(0).AddItem "Border", 2
        toolbar_Tools.cmbSelType(0).ListIndex = 0
        
        toolbar_Tools.sltSelectionFeathering.assignTooltip "This feathering slider allows for immediate feathering adjustments.  For performance reasons, it is limited to small radii.  For larger feathering radii, please use the Select -> Feathering menu."
        toolbar_Tools.sltCornerRounding.assignTooltip "This option adjusts the roundness of a rectangular selection's corners."
        toolbar_Tools.sltSelectionLineWidth.assignTooltip "This option adjusts the width of a line selection."
        toolbar_Tools.sltSelectionBorder.assignTooltip "This option adjusts the width of the selection border."
        
        'Quick-fix tools
        cmdQuickFix(0).ToolTip = g_Language.TranslateMessage("Reset all quick-fix adjustment values")
        cmdQuickFix(1).ToolTip = g_Language.TranslateMessage("Make quick-fix adjustments permanent.  This action is never required, but if viewport rendering is sluggish and many quick-fix adjustments are active, it may improve performance.")
        
    'Load any last-used settings for this form
    Set lastUsedSettings = New pdLastUsedSettings
    lastUsedSettings.setParentForm Me
    lastUsedSettings.loadAllControlValues
    
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'Allow non-destructive effects
    m_NonDestructiveFXAllowed = True

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'Save all last-used settings to file
    lastUsedSettings.saveAllControlValues

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
        g_CurrentTool = NAV_DRAG
    End If
    resetToolButtonStates
    
    'Reset the selection coordinate boxes to 0
    Dim i As Long
    For i = 0 To tudSel.Count - 1
        tudSel(i) = 0
    Next i

End Sub

'When the selection type is changed, update the corresponding preference and redraw all selections
Private Sub cmbSelRender_Click(Index As Integer)
            
    If g_OpenImageCount > 0 Then
    
        Dim i As Long
        For i = 0 To g_NumOfImagesLoaded
            If (Not pdImages(i) Is Nothing) Then
                If pdImages(i).IsActive And pdImages(i).selectionActive Then RenderViewport pdImages(i), FormMain.mainCanvas(0)
            End If
        Next i
    
    End If
    
End Sub

'Change selection smoothing (e.g. none, antialiased, fully feathered)
Private Sub cmbSelSmoothing_Click(Index As Integer)
    
    updateSelectionPanelLayout
    
    'If a selection is already active, change its type to match the current selection, then redraw it
    If selectionsAllowed(False) Then
        pdImages(g_CurrentImage).mainSelection.setSmoothingType cmbSelSmoothing(Index).ListIndex
        pdImages(g_CurrentImage).mainSelection.setFeatheringRadius sltSelectionFeathering.Value
        RenderViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
    
End Sub

'Change selection type (e.g. interior, exterior, bordered)
Private Sub cmbSelType_Click(Index As Integer)

    updateSelectionPanelLayout
    
    'If a selection is already active, change its type to match the current selection, then redraw it
    If selectionsAllowed(False) Then
        pdImages(g_CurrentImage).mainSelection.setSelectionType cmbSelType(Index).ListIndex
        pdImages(g_CurrentImage).mainSelection.setBorderSize sltSelectionBorder.Value
        RenderViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
    
End Sub

'External functions can use this to request the selection of a new tool (for example, Select All uses this to set the
' rectangular tool selector as the current tool)
Public Sub selectNewTool(ByVal newToolID As PDTools)
    g_PreviousTool = g_CurrentTool
    g_CurrentTool = newToolID
    resetToolButtonStates
End Sub

'When a new tool button is selected, we need to raise all the others and display the proper options box
Public Sub resetToolButtonStates()
    
    'Start by depressing the selected button and raising all unselected ones
    Dim catID As Long
    For catID = 0 To cmdTools.Count - 1
        If catID = g_CurrentTool Then
            cmdTools(catID).Value = True
        Else
            cmdTools(catID).Value = False
        End If
    Next catID
    
    Dim i As Long
    
    'Next, we need to display the correct tool options panel.  There is no set pattern to this; some tools share
    ' panels, but show/hide certain controls as necessary.  Other tools require their own unique panel.  I've tried
    ' to strike a balance between "as few panels as possible" without going overboard.
    Dim activeToolPanel As Long
    
    Select Case g_CurrentTool
        
        'Move/size tool
        Case NAV_MOVE
            activeToolPanel = 1
        
        'Rectangular, Elliptical, Line selections
        Case SELECT_RECT, SELECT_CIRC, SELECT_LINE
            activeToolPanel = 0
            
        '"Quick fix" tool(s)
        Case QUICK_FIX_LIGHTING
            activeToolPanel = 2
        
        Case Else
            activeToolPanel = -1
        
    End Select
    
    'If tools share the same panel, they may need to show or hide a few additional controls.  (For example,
    ' "corner rounding", which is needed for rectangular selections but not elliptical ones, despite the two
    ' sharing the same tool panel.)  Do this before showing or hiding the tool panel.
    Select Case g_CurrentTool
    
        'For rectangular selections, show the rounded corners option
        Case SELECT_RECT
            toolbar_Tools.lblSelection(5).Visible = True
            toolbar_Tools.sltCornerRounding.Visible = True
            toolbar_Tools.sltSelectionLineWidth.Visible = False
            
        'For elliptical selections, hide the rounded corners option
        Case SELECT_CIRC
            toolbar_Tools.lblSelection(5).Visible = False
            toolbar_Tools.sltCornerRounding.Visible = False
            toolbar_Tools.sltSelectionLineWidth.Visible = False
            
        'Line selections also show the rounded corners slider, though they repurpose it for line width
        Case SELECT_LINE
            toolbar_Tools.lblSelection(5).Visible = True
            toolbar_Tools.sltCornerRounding.Visible = False
            toolbar_Tools.sltSelectionLineWidth.Visible = True
        
    End Select
    
    'Even if tools share the same panel, they may name controls differently, or use different max/min values.
    ' Check for this, and apply new text and max/min settings as necessary.
    Select Case g_CurrentTool
    
        'Rectangular and elliptical selections use rectangular bounding boxes and potential corner rounding
        Case SELECT_RECT, SELECT_CIRC
            lblSelection(1).Caption = g_Language.TranslateMessage("position (x, y)")
            lblSelection(2).Caption = g_Language.TranslateMessage("size (x, y)")
            lblSelection(5).Caption = g_Language.TranslateMessage("corner rounding")
            
        'Line selections use two points, and the corner rounding slider gets repurposed as line width.
        Case SELECT_LINE
            lblSelection(1).Caption = g_Language.TranslateMessage("1st point (x, y)")
            lblSelection(2).Caption = g_Language.TranslateMessage("2nd point (x, y)")
            lblSelection(5).Caption = g_Language.TranslateMessage("line width")
            
    End Select
    
    'Display the current tool options panel, while hiding all inactive ones
    For i = 0 To picTools.Count - 1
        If i = activeToolPanel Then
            If Not picTools(i).Visible Then
                picTools(i).Visible = True
                setArrowCursor picTools(i)
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
                If (g_PreviousTool = SELECT_CIRC) And (pdImages(g_CurrentImage).mainSelection.getSelectionShape = sCircle) Then
                    pdImages(g_CurrentImage).mainSelection.setSelectionShape sRectangle
                    RenderViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
                Else
                    
                    If pdImages(g_CurrentImage).mainSelection.getSelectionShape = sRectangle Then
                        metaToggle tSelectionTransform, True
                    Else
                    
                        'Remove any existing selections
                        If g_OpenImageCount > 0 Then Process "Remove selection", , , UNDO_SELECTION
                    
                        metaToggle tSelectionTransform, False
                        
                    End If
                End If
            End If
            
        Case SELECT_CIRC
        
            'If a similar selection is already active, change its shape to match the current tool, then redraw it
            If selectionsAllowed(True) And (Not g_UndoRedoActive) Then
                If (g_PreviousTool = SELECT_RECT) And (pdImages(g_CurrentImage).mainSelection.getSelectionShape = sRectangle) Then
                    pdImages(g_CurrentImage).mainSelection.setSelectionShape sCircle
                    RenderViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
                Else
                    
                    If pdImages(g_CurrentImage).mainSelection.getSelectionShape = sCircle Then
                        metaToggle tSelectionTransform, True
                    Else
                        
                        'Remove any existing selections
                        If g_OpenImageCount > 0 Then Process "Remove selection", , , UNDO_SELECTION
                        
                        metaToggle tSelectionTransform, False
                        
                    End If
                End If
            End If
            
        'Line selections
        Case SELECT_LINE
        
            'Deactivate the position text boxes - those shouldn't be accessible unless a line selection is presently active
            If selectionsAllowed(True) Then
                If pdImages(g_CurrentImage).mainSelection.getSelectionShape = sLine Then
                    metaToggle tSelectionTransform, True
                Else
                
                    'Remove any existing selections
                    If g_OpenImageCount > 0 Then Process "Remove selection", , , UNDO_SELECTION
                
                    metaToggle tSelectionTransform, False
                    
                End If
            Else
                metaToggle tSelectionTransform, False
            End If
            
        Case Else
        
    End Select
    
    'Finally, because tools may do some custom rendering atop the image canvas, now is a good time to redraw the canvas
    RenderViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

'Toolbars can never be unloaded, EXCEPT when the whole program is going down.  Check for the program-wide closing flag prior
' to exiting; if it is not found, cancel the unload and simply hide this form.  (Note that the toggleToolbarVisibility sub
' will also keep this toolbar's Window menu entry in sync with the form's current visibility.)
Private Sub Form_Unload(Cancel As Integer)
    
    If g_ProgramShuttingDown Then
        ReleaseFormTheming Me
        g_WindowManager.unregisterForm Me
    Else
        Cancel = True
        toggleToolbarVisibility TOOLS_TOOLBOX
    End If
    
End Sub

Private Sub sltCornerRounding_Change()
    If selectionsAllowed(True) Then
        pdImages(g_CurrentImage).mainSelection.setRoundedCornerAmount sltCornerRounding.Value
        RenderViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
End Sub

'Non-destructive effect changes will force an immediate redraw of the viewport
Private Sub sltQuickFix_Change(Index As Integer)

    If (Not pdImages(g_CurrentImage) Is Nothing) And m_NonDestructiveFXAllowed Then
        
        'Check the state of the layer's non-destructive FX tracker before making any changes
        Dim initFXState As Boolean
        initFXState = pdImages(g_CurrentImage).getActiveLayer.getLayerNonDestructiveFXState
        
        'The index of sltQuickFix controls aligns exactly with PD's constants for non-destructive effects.  This is by design.
        pdImages(g_CurrentImage).getActiveLayer.setLayerNonDestructiveFXState Index, sltQuickFix(Index).Value
        
        'Redraw the viewport
        ScrollViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
        
        'If the layer now has non-destructive effects active, enable the quick fix buttons (if they aren't already)
        Dim i As Long
        
        If pdImages(g_CurrentImage).getActiveLayer.getLayerNonDestructiveFXState Then
        
            For i = 0 To cmdQuickFix.Count - 1
                If Not cmdQuickFix(i).Enabled Then cmdQuickFix(i).Enabled = True
            Next i
        
        Else
            
            For i = 0 To cmdQuickFix.Count - 1
                If cmdQuickFix(i).Enabled Then cmdQuickFix(i).Enabled = False
            Next i
        
        End If
        
        'Even though this action is not destructive, we want to allow the user to save after making non-destructive changes.
        If pdImages(g_CurrentImage).getSaveState(pdSE_AnySave) And (pdImages(g_CurrentImage).getActiveLayer.getLayerNonDestructiveFXState <> initFXState) Then
            pdImages(g_CurrentImage).setSaveState False, pdSE_AnySave
            syncInterfaceToCurrentImage
        End If
        
    End If

End Sub

Private Sub sltSelectionBorder_Change()
    If selectionsAllowed(False) Then
        pdImages(g_CurrentImage).mainSelection.setBorderSize sltSelectionBorder.Value
        RenderViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
End Sub

Private Sub sltSelectionFeathering_Change()
    If selectionsAllowed(False) Then
        pdImages(g_CurrentImage).mainSelection.setFeatheringRadius sltSelectionFeathering.Value
        RenderViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
End Sub

Private Sub sltSelectionLineWidth_Change()
    If selectionsAllowed(True) Then
        pdImages(g_CurrentImage).mainSelection.setSelectionLineWidth sltSelectionLineWidth.Value
        RenderViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
End Sub

Private Function selectionsAllowed(ByVal transformableMatters As Boolean) As Boolean
    If g_OpenImageCount > 0 Then
        If pdImages(g_CurrentImage).selectionActive And (Not pdImages(g_CurrentImage).mainSelection Is Nothing) And (Not pdImages(g_CurrentImage).mainSelection.rejectRefreshRequests) Then
            
            If transformableMatters Then
                If pdImages(g_CurrentImage).mainSelection.isTransformable Then
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

'When certain selection settings are enabled or disabled, corresponding controls are shown or hidden.  To keep the
' panel concise and clean, we move other controls up or down depending on what controls are visible.
Private Sub updateSelectionPanelLayout()

    'Display the feathering slider as necessary
    If cmbSelSmoothing(0).ListIndex = sFullyFeathered Then
        sltSelectionFeathering.Visible = True
    Else
        sltSelectionFeathering.Visible = False
    End If
    
    'Display the border slider as necessary
    If cmbSelType(0).ListIndex = sBorder Then
        sltSelectionBorder.Visible = True
    Else
        sltSelectionBorder.Visible = False
    End If
    
End Sub

'When the selection text boxes are updated, change the scrollbars to match
Private Sub tudSel_Change(Index As Integer)
    updateSelectionsValuesViaText
End Sub

Private Sub updateSelectionsValuesViaText()
    If selectionsAllowed(True) Then
        If Not pdImages(g_CurrentImage).mainSelection.rejectRefreshRequests Then
            pdImages(g_CurrentImage).mainSelection.updateViaTextBox
            RenderViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
        End If
    End If
End Sub

'External functions can use this to re-theme this form at run-time (important when changing languages, for example)
Public Sub requestMakeFormPretty()
    makeFormPretty Me, m_ToolTip
End Sub
