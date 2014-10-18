VERSION 5.00
Begin VB.Form toolbar_Tools 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Tools"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   18435
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
   ScaleHeight     =   96
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1229
   ShowInTaskbar   =   0   'False
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
      ScaleWidth      =   1230
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      Visible         =   0   'False
      Width           =   18450
      Begin VB.ComboBox cboWandCompare 
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
         ItemData        =   "VBP_ToolbarTools.frx":0000
         Left            =   8790
         List            =   "VBP_ToolbarTools.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   47
         ToolTipText     =   "This option controls the selection's area.  You can switch between the three settings without losing the current selection."
         Top             =   840
         Width           =   2445
      End
      Begin PhotoDemon.colorSelector csSelectionHighlight 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   40
         Top             =   840
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   661
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
         ItemData        =   "VBP_ToolbarTools.frx":0004
         Left            =   120
         List            =   "VBP_ToolbarTools.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   390
         Width           =   2250
      End
      Begin VB.ComboBox cmbSelArea 
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
         ItemData        =   "VBP_ToolbarTools.frx":0008
         Left            =   5460
         List            =   "VBP_ToolbarTools.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   2
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
         ItemData        =   "VBP_ToolbarTools.frx":000C
         Left            =   2640
         List            =   "VBP_ToolbarTools.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Use this option to change the way selections blend with their surroundings."
         Top             =   390
         Width           =   2445
      End
      Begin PhotoDemon.sliderTextCombo sltCornerRounding 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   8100
         TabIndex        =   3
         Top             =   345
         Visible         =   0   'False
         Width           =   2670
         _ExtentX        =   4710
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
         Max             =   1
      End
      Begin PhotoDemon.textUpDown tudSel 
         Height          =   405
         Index           =   0
         Left            =   11520
         TabIndex        =   4
         Top             =   390
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
         Left            =   11520
         TabIndex        =   5
         Top             =   840
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
         Left            =   13080
         TabIndex        =   6
         Top             =   390
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
         Left            =   13080
         TabIndex        =   7
         Top             =   840
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
         Left            =   5340
         TabIndex        =   8
         Top             =   840
         Visible         =   0   'False
         Width           =   2670
         _ExtentX        =   4710
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
         Min             =   1
         Max             =   10000
         Value           =   1
      End
      Begin PhotoDemon.sliderTextCombo sltSelectionFeathering 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   2520
         TabIndex        =   9
         Top             =   840
         Visible         =   0   'False
         Width           =   2670
         _ExtentX        =   4710
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
         Max             =   100
      End
      Begin PhotoDemon.sliderTextCombo sltSelectionLineWidth 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   8160
         TabIndex        =   10
         Top             =   345
         Visible         =   0   'False
         Width           =   2670
         _ExtentX        =   4710
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
         Min             =   1
         Max             =   10000
         Value           =   10
      End
      Begin PhotoDemon.sliderTextCombo sltSmoothStroke 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   8160
         TabIndex        =   41
         Top             =   360
         Visible         =   0   'False
         Width           =   2670
         _ExtentX        =   4710
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
         Max             =   1
         SigDigits       =   2
      End
      Begin PhotoDemon.sliderTextCombo sltPolygonCurvature 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   8040
         TabIndex        =   42
         Top             =   345
         Visible         =   0   'False
         Width           =   2670
         _ExtentX        =   4710
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
         Max             =   1
         SigDigits       =   2
      End
      Begin PhotoDemon.buttonStrip btsWandMerge 
         Height          =   825
         Left            =   11520
         TabIndex        =   43
         Top             =   390
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1455
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
      Begin PhotoDemon.sliderTextCombo sltWandTolerance 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   8640
         TabIndex        =   44
         Top             =   345
         Visible         =   0   'False
         Width           =   2670
         _ExtentX        =   4710
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
         Max             =   255
         SigDigits       =   1
      End
      Begin PhotoDemon.buttonStrip btsWandArea 
         Height          =   825
         Left            =   5460
         TabIndex        =   45
         Top             =   390
         Width           =   2895
         _ExtentX        =   4366
         _ExtentY        =   1455
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
         Caption         =   "tolerance"
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
         Left            =   8760
         TabIndex        =   46
         Top             =   60
         Width           =   795
      End
      Begin VB.Label lblSelection 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "appearance"
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
         TabIndex        =   16
         Top             =   60
         Width           =   1005
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
         Left            =   13080
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
         Left            =   11520
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
         Left            =   8220
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
         Left            =   5460
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
         Left            =   2640
         TabIndex        =   11
         Top             =   60
         Width           =   885
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
      Index           =   2
      Left            =   15
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   950
      TabIndex        =   25
      Top             =   15
      Visible         =   0   'False
      Width           =   14250
      Begin PhotoDemon.sliderTextCombo sltQuickFix 
         CausesValidation=   0   'False
         Height          =   495
         Index           =   0
         Left            =   1380
         TabIndex        =   27
         Top             =   90
         Width           =   2670
         _ExtentX        =   4710
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
         Min             =   -2
         Max             =   2
         SigDigits       =   2
         SliderTrackStyle=   2
      End
      Begin PhotoDemon.sliderTextCombo sltQuickFix 
         CausesValidation=   0   'False
         Height          =   495
         Index           =   1
         Left            =   1380
         TabIndex        =   28
         Top             =   705
         Width           =   2670
         _ExtentX        =   4710
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
         Min             =   -100
         Max             =   100
      End
      Begin PhotoDemon.sliderTextCombo sltQuickFix 
         CausesValidation=   0   'False
         Height          =   495
         Index           =   2
         Left            =   5640
         TabIndex        =   30
         Top             =   90
         Width           =   2670
         _ExtentX        =   4710
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
         Min             =   -100
         Max             =   100
      End
      Begin PhotoDemon.sliderTextCombo sltQuickFix 
         CausesValidation=   0   'False
         Height          =   495
         Index           =   3
         Left            =   5640
         TabIndex        =   32
         Top             =   705
         Width           =   2670
         _ExtentX        =   4710
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
         Min             =   -100
         Max             =   100
      End
      Begin PhotoDemon.jcbutton cmdQuickFix 
         Height          =   570
         Index           =   0
         Left            =   13080
         TabIndex        =   34
         Top             =   75
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   1005
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
         BackColor       =   -2147483643
         Caption         =   ""
         HandPointer     =   -1  'True
         PictureNormal   =   "VBP_ToolbarTools.frx":0010
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         ColorScheme     =   3
      End
      Begin PhotoDemon.jcbutton cmdQuickFix 
         Height          =   570
         Index           =   1
         Left            =   13080
         TabIndex        =   35
         Top             =   705
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   1005
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
         BackColor       =   -2147483643
         Caption         =   ""
         HandPointer     =   -1  'True
         PictureNormal   =   "VBP_ToolbarTools.frx":0D62
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         ColorScheme     =   3
      End
      Begin PhotoDemon.sliderTextCombo sltQuickFix 
         CausesValidation=   0   'False
         Height          =   495
         Index           =   4
         Left            =   9960
         TabIndex        =   36
         Top             =   90
         Width           =   2670
         _ExtentX        =   4710
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
         Min             =   -100
         Max             =   100
         SliderTrackStyle=   3
         GradientColorLeft=   16752699
         GradientColorRight=   2990335
         GradientColorMiddle=   16777215
      End
      Begin PhotoDemon.sliderTextCombo sltQuickFix 
         CausesValidation=   0   'False
         Height          =   495
         Index           =   5
         Left            =   9960
         TabIndex        =   37
         Top             =   705
         Width           =   2670
         _ExtentX        =   4710
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
         Min             =   -100
         Max             =   100
         SliderTrackStyle=   3
         GradientColorLeft=   15102446
         GradientColorRight=   8253041
         GradientColorMiddle=   16777215
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   33
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
         TabIndex        =   31
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
         TabIndex        =   29
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
         TabIndex        =   26
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
      TabIndex        =   18
      Top             =   15
      Visible         =   0   'False
      Width           =   14250
      Begin PhotoDemon.smartCheckBox chkLayerBorder 
         Height          =   330
         Left            =   7080
         TabIndex        =   20
         Top             =   360
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   582
         Caption         =   "show layer borders"
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
      Begin PhotoDemon.smartCheckBox chkLayerNodes 
         Height          =   330
         Left            =   7080
         TabIndex        =   21
         Top             =   780
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   582
         Caption         =   "show layer transform nodes"
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
      Begin PhotoDemon.smartCheckBox chkAutoActivateLayer 
         Height          =   330
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   6810
         _ExtentX        =   12012
         _ExtentY        =   582
         Caption         =   "automatically activate layer beneath mouse"
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
      Begin PhotoDemon.smartCheckBox chkIgnoreTransparent 
         Height          =   330
         Left            =   120
         TabIndex        =   24
         Top             =   780
         Width           =   6810
         _ExtentX        =   12012
         _ExtentY        =   582
         Caption         =   "ignore transparent pixels when auto-activating layers"
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
         TabIndex        =   22
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
         Left            =   7080
         TabIndex        =   19
         Top             =   60
         Width           =   1335
      End
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
'Last updated: 16/October/14
'Last update: rework all selection interface code to use the new property dictionary functions
'
'This form was initially integrated into the main MDI form.  In fall 2013, PhotoDemon left behind the MDI model,
' and all toolbars were moved to their own forms.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

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

Private Sub btsWandArea_Click(ByVal buttonIndex As Long)
    
    'If a selection is already active, change its type to match the current option, then redraw it
    If selectionsAllowed(False) Then
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_WAND_SEARCH_MODE, buttonIndex
        RenderViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
    
End Sub

Private Sub btsWandMerge_Click(ByVal buttonIndex As Long)

    'If a selection is already active, change its type to match the current option, then redraw it
    If selectionsAllowed(False) Then
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_WAND_SAMPLE_MERGED, buttonIndex
        RenderViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If

End Sub

Private Sub cboWandCompare_Click()
    
    'Limit the accuracy of the tolerance for certain comparison methods.
    If cboWandCompare.ListIndex > 1 Then
        sltWandTolerance.SigDigits = 0
    Else
        sltWandTolerance.SigDigits = 1
    End If
    
    'If a selection is already active, change its type to match the current option, then redraw it
    If selectionsAllowed(False) Then
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_WAND_COMPARE_METHOD, cboWandCompare.ListIndex
        RenderViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
    
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
            For i = 0 To sltQuickFix.count - 1
                
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
                For i = 0 To sltQuickFix.count - 1
                    sltQuickFix(i).Value = 0
                Next i
                
                setImageCheckpoint
                
                'Ask the central processor to permanently apply the quick-fix changes
                Process "Make quick-fixes permanent", , , UNDO_LAYER
                                
            End If
    
    End Select
    
    'After one of these buttons has been used, all quick-fix values will be reset - so we can disable the buttons accordingly.
    For i = 0 To cmdQuickFix.count - 1
        If cmdQuickFix(i).Enabled Then cmdQuickFix(i).Enabled = False
    Next i
    
    'Re-enable auto-refreshes
    setNDFXControlState True
    
    'Redraw the viewport
    ScrollViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)

End Sub

Private Sub csSelectionHighlight_ColorChanged(Index As Integer)
    
    'Redraw the viewport
    If selectionsAllowed(False) Then RenderViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub Form_Load()

    Dim i As Long
        
    'INITIALIZE ALL TOOLS
    
        
    
        'Selection visual styles (currently lightbox or highlight)
        toolbar_Tools.cmbSelRender(0).ToolTipText = g_Language.TranslateMessage("Click to change the way selections are rendered onto the image canvas.  This has no bearing on selection contents - only the way they appear while editing.")
        For i = 0 To toolbar_Tools.cmbSelRender.count - 1
            toolbar_Tools.cmbSelRender(i).AddItem " Highlight", 0
            toolbar_Tools.cmbSelRender(i).AddItem " Lightbox", 1
            toolbar_Tools.cmbSelRender(i).AddItem " Outline", 2
            toolbar_Tools.cmbSelRender(i).ListIndex = 0
        Next i
        csSelectionHighlight(0).Color = RGB(255, 58, 72)
        csSelectionHighlight(0).Visible = True
        
        'Selection smoothing (currently none, antialiased, fully feathered)
        toolbar_Tools.cmbSelSmoothing(0).ToolTipText = g_Language.TranslateMessage("This option controls how smoothly a selection blends with its surroundings.")
        toolbar_Tools.cmbSelSmoothing(0).AddItem " None", 0
        toolbar_Tools.cmbSelSmoothing(0).AddItem " Antialiased", 1
        toolbar_Tools.cmbSelSmoothing(0).AddItem " Feathered", 2
        toolbar_Tools.cmbSelSmoothing(0).ListIndex = 1
        
        'Selection types (currently interior, exterior, border)
        toolbar_Tools.cmbSelArea(0).ToolTipText = g_Language.TranslateMessage("These options control the area affected by a selection.  The selection can be modified on-canvas while any of these settings are active.  For more advanced selection adjustments, use the Select menu.")
        toolbar_File.setSelectionAreaOptions True, 0
        
        toolbar_Tools.sltSelectionFeathering.assignTooltip "This feathering slider allows for immediate feathering adjustments.  For performance reasons, it is limited to small radii.  For larger feathering radii, please use the Select -> Feathering menu."
        toolbar_Tools.sltCornerRounding.assignTooltip "This option adjusts the roundness of a rectangular selection's corners."
        toolbar_Tools.sltSelectionLineWidth.assignTooltip "This option adjusts the width of a line selection."
        toolbar_Tools.sltSelectionBorder.assignTooltip "This option adjusts the width of the selection border."
        toolbar_Tools.sltPolygonCurvature.assignTooltip "This option adjusts the curvature, if any, of a polygon selection's sides."
        toolbar_Tools.sltSmoothStroke.assignTooltip "This option increases the smoothness of a hand-drawn lasso selection."
        toolbar_Tools.sltWandTolerance.assignTooltip "Tolerance controls how similar two pixels must be before adding them to a magic wand selection."
        
        'Magic wand options
        btsWandMerge.AddItem "image", 0
        btsWandMerge.AddItem "layer", 1
        btsWandMerge.ListIndex = 0
        btsWandMerge.ToolTipText = g_Language.TranslateMessage("The magic wand can operate on the entire image, or just the active layer.")
        
        btsWandArea.AddItem "contiguous", 0
        btsWandArea.AddItem "global", 1
        btsWandArea.ListIndex = 0
        btsWandArea.ToolTipText = g_Language.TranslateMessage("Normally, the magic wand will spread out from the target pixel, adding neighboring pixels to the selection as it goes.  You can alternatively set it to search the entire image, without regards for continuuity.")
        
        cboWandCompare.Clear
        cboWandCompare.AddItem " Composite", 0
        cboWandCompare.AddItem " Luminance", 1
        cboWandCompare.AddItem " Red", 2
        cboWandCompare.AddItem " Green", 3
        cboWandCompare.AddItem " Blue", 4
        cboWandCompare.AddItem " Alpha", 5
        cboWandCompare.ListIndex = 0
        cboWandCompare.ToolTipText = g_Language.TranslateMessage("This option controls which criteria the magic wand uses to determine whether a pixel should be added to the current selection.")
        
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

Private Sub lastUsedSettings_ReadCustomPresetData()
    
    'Reset the selection coordinate boxes to 0
    Dim i As Long
    For i = 0 To tudSel.count - 1
        tudSel(i) = 0
    Next i

End Sub

'When the selection render type is changed, we must redraw the viewport to match
Private Sub cmbSelRender_Click(Index As Integer)
    
    'Show or hide the color selector, as appropriate
    If cmbSelRender(Index).ListIndex = SELECTION_RENDER_HIGHLIGHT Then
        csSelectionHighlight(Index).Visible = True
    Else
        csSelectionHighlight(Index).Visible = False
    End If
    
    'Redraw the viewport
    If selectionsAllowed(False) Then RenderViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

'Change selection smoothing (e.g. none, antialiased, fully feathered)
Private Sub cmbSelSmoothing_Click(Index As Integer)
    
    updateSelectionPanelLayout
    
    'If a selection is already active, change its type to match the current selection, then redraw it
    If selectionsAllowed(False) Then
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_SMOOTHING, cmbSelSmoothing(Index).ListIndex
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_FEATHERING_RADIUS, sltSelectionFeathering.Value
        RenderViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
    
End Sub

'Change selection type (e.g. interior, exterior, bordered)
Private Sub cmbSelArea_Click(Index As Integer)

    updateSelectionPanelLayout
    
    'If a selection is already active, change its type to match the current selection, then redraw it
    If selectionsAllowed(False) Then
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_AREA, cmbSelArea(Index).ListIndex
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_BORDER_WIDTH, sltSelectionBorder.Value
        RenderViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
    
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
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_ROUNDED_CORNER_RADIUS, sltCornerRounding.Value
        RenderViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
End Sub

Private Sub sltPolygonCurvature_Change()
    If selectionsAllowed(True) Then
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_POLYGON_CURVATURE, sltPolygonCurvature.Value
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
        
            For i = 0 To cmdQuickFix.count - 1
                If Not cmdQuickFix(i).Enabled Then cmdQuickFix(i).Enabled = True
            Next i
        
        Else
            
            For i = 0 To cmdQuickFix.count - 1
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
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_BORDER_WIDTH, sltSelectionBorder.Value
        RenderViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
End Sub

Private Sub sltSelectionFeathering_Change()
    If selectionsAllowed(False) Then
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_FEATHERING_RADIUS, sltSelectionFeathering.Value
        RenderViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
End Sub

Private Sub sltSelectionLineWidth_Change()
    If selectionsAllowed(True) Then
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_LINE_WIDTH, sltSelectionLineWidth.Value
        RenderViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
End Sub

'When certain selection settings are enabled or disabled, corresponding controls are shown or hidden.  To keep the
' panel concise and clean, we move other controls up or down depending on what controls are visible.
Public Sub updateSelectionPanelLayout()

    'Display the feathering slider as necessary
    If cmbSelSmoothing(0).ListIndex = sFullyFeathered Then
        sltSelectionFeathering.Visible = True
    Else
        sltSelectionFeathering.Visible = False
    End If
    
    'Display the border slider as necessary
    If cmbSelArea(0).ListIndex = sBorder Then
        sltSelectionBorder.Visible = True
    Else
        sltSelectionBorder.Visible = False
    End If
    
End Sub

Private Sub sltSmoothStroke_Change()
    If selectionsAllowed(False) Then
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_SMOOTH_STROKE, sltSmoothStroke.Value
        RenderViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
End Sub

Private Sub sltWandTolerance_Change()
    If selectionsAllowed(False) Then
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_WAND_TOLERANCE, sltWandTolerance.Value
        RenderViewport pdImages(g_CurrentImage), FormMain.mainCanvas(0)
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
