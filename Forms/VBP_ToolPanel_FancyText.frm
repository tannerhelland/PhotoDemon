VERSION 5.00
Begin VB.Form toolpanel_FancyText 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   1515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18435
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   101
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1229
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.buttonStripVertical btsCategory 
      Height          =   1380
      Left            =   6240
      TabIndex        =   1
      Top             =   30
      Width           =   2175
      _ExtentX        =   4048
      _ExtentY        =   2434
   End
   Begin PhotoDemon.pdTextBox txtTextTool 
      Height          =   1380
      Left            =   840
      TabIndex        =   0
      Top             =   30
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2434
      FontSize        =   9
      Multiline       =   -1  'True
   End
   Begin PhotoDemon.pdLabel lblText 
      Height          =   240
      Index           =   1
      Left            =   120
      Top             =   60
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   503
      Alignment       =   1
      Caption         =   "text:"
      ForeColor       =   0
   End
   Begin VB.PictureBox picCategory 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   1500
      Index           =   0
      Left            =   8520
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   729
      TabIndex        =   2
      Top             =   0
      Width           =   10935
      Begin VB.PictureBox picCharCategory 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   1500
         Index           =   0
         Left            =   1920
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   729
         TabIndex        =   40
         Top             =   60
         Visible         =   0   'False
         Width           =   10935
         Begin PhotoDemon.textUpDown tudTextFontSize 
            Height          =   345
            Left            =   1320
            TabIndex        =   41
            Top             =   450
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   609
            Min             =   1
            Max             =   1000
            Value           =   16
         End
         Begin PhotoDemon.pdLabel lblText 
            Height          =   240
            Index           =   3
            Left            =   0
            Top             =   60
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   503
            Alignment       =   1
            Caption         =   "font face:"
            ForeColor       =   0
         End
         Begin PhotoDemon.pdLabel lblText 
            Height          =   240
            Index           =   4
            Left            =   0
            Top             =   510
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   503
            Alignment       =   1
            Caption         =   "font size:"
            ForeColor       =   0
         End
         Begin PhotoDemon.pdLabel lblText 
            Height          =   240
            Index           =   2
            Left            =   0
            Top             =   960
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   503
            Alignment       =   1
            Caption         =   "font style:"
            ForeColor       =   0
         End
         Begin PhotoDemon.pdButtonToolbox btnFontStyles 
            Height          =   435
            Index           =   1
            Left            =   1800
            TabIndex        =   42
            Top             =   870
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   767
            StickyToggle    =   -1  'True
         End
         Begin PhotoDemon.pdButtonToolbox btnFontStyles 
            Height          =   435
            Index           =   2
            Left            =   2280
            TabIndex        =   43
            Top             =   870
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   767
            StickyToggle    =   -1  'True
         End
         Begin PhotoDemon.pdButtonToolbox btnFontStyles 
            Height          =   435
            Index           =   3
            Left            =   2760
            TabIndex        =   44
            Top             =   870
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   767
            StickyToggle    =   -1  'True
         End
         Begin PhotoDemon.smartCheckBox chkHinting 
            Height          =   330
            Left            =   4200
            TabIndex        =   45
            Top             =   450
            Width           =   1815
            _ExtentX        =   2990
            _ExtentY        =   582
            Caption         =   "hinting"
            Value           =   0
         End
         Begin PhotoDemon.pdComboBox cboTextRenderingHint 
            Height          =   375
            Left            =   5400
            TabIndex        =   46
            Top             =   0
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   635
         End
         Begin PhotoDemon.pdLabel lblText 
            Height          =   240
            Index           =   5
            Left            =   3840
            Top             =   60
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            Alignment       =   1
            Caption         =   "antialiasing:"
            ForeColor       =   0
         End
         Begin PhotoDemon.pdComboBox_Font cboTextFontFace 
            Height          =   375
            Left            =   1320
            TabIndex        =   47
            Top             =   0
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
         End
         Begin PhotoDemon.pdButtonToolbox btnFontStyles 
            Height          =   435
            Index           =   0
            Left            =   1320
            TabIndex        =   48
            Top             =   870
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   767
            StickyToggle    =   -1  'True
         End
      End
      Begin PhotoDemon.buttonStripVertical btsCharCategory 
         Height          =   1380
         Left            =   0
         TabIndex        =   39
         Top             =   30
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   2434
      End
      Begin VB.PictureBox picCharCategory 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   1500
         Index           =   1
         Left            =   1920
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   729
         TabIndex        =   49
         Top             =   60
         Visible         =   0   'False
         Width           =   10935
         Begin PhotoDemon.pdLabel lblText 
            Height          =   240
            Index           =   26
            Left            =   0
            Top             =   60
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   503
            Alignment       =   1
            Caption         =   "OpenType:"
            ForeColor       =   0
         End
      End
   End
   Begin VB.PictureBox picCategory 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   1500
      Index           =   3
      Left            =   8520
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   729
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   10935
   End
   Begin VB.PictureBox picCategory 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   1500
      Index           =   2
      Left            =   8520
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   729
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   10935
      Begin PhotoDemon.buttonStripVertical btsAppearanceCategory 
         Height          =   1380
         Left            =   30
         TabIndex        =   9
         Top             =   30
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   2434
      End
      Begin VB.PictureBox picAppearanceCategory 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   1500
         Index           =   2
         Left            =   1920
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   729
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   10935
         Begin PhotoDemon.smartCheckBox chkBackground 
            Height          =   330
            Left            =   915
            TabIndex        =   26
            Top             =   105
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   582
            Caption         =   "fill background"
            Value           =   0
         End
         Begin PhotoDemon.pdLabel lblText 
            Height          =   240
            Index           =   16
            Left            =   3480
            Top             =   120
            Width           =   885
            _ExtentX        =   1984
            _ExtentY        =   503
            Alignment       =   1
            Caption         =   "opacity:"
            ForeColor       =   0
         End
         Begin PhotoDemon.sliderTextCombo sltBackgroundOpacity 
            CausesValidation=   0   'False
            Height          =   495
            Left            =   4440
            TabIndex        =   25
            Top             =   30
            Width           =   2760
            _ExtentX        =   4868
            _ExtentY        =   873
            Max             =   100
            Value           =   100
            NotchPosition   =   2
            NotchValueCustom=   100
         End
         Begin PhotoDemon.colorSelector csBackground 
            Height          =   840
            Left            =   960
            TabIndex        =   27
            Top             =   540
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   1508
         End
         Begin PhotoDemon.pdLabel lblText 
            Height          =   720
            Index           =   15
            Left            =   0
            Top             =   600
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   1270
            Alignment       =   1
            Caption         =   "color:"
            ForeColor       =   0
            Layout          =   1
         End
      End
      Begin VB.PictureBox picAppearanceCategory 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   1500
         Index           =   1
         Left            =   1920
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   729
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   10935
         Begin PhotoDemon.pdLabel lblText 
            Height          =   240
            Index           =   17
            Left            =   0
            Top             =   120
            Width           =   885
            _ExtentX        =   1984
            _ExtentY        =   503
            Alignment       =   1
            Caption         =   "mode:"
            ForeColor       =   0
         End
         Begin PhotoDemon.pdComboBox cboOutlineMode 
            Height          =   375
            Left            =   960
            TabIndex        =   28
            Top             =   75
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   635
         End
         Begin PhotoDemon.pdLabel lblText 
            Height          =   240
            Index           =   18
            Left            =   3480
            Top             =   120
            Width           =   885
            _ExtentX        =   1984
            _ExtentY        =   503
            Alignment       =   1
            Caption         =   "opacity:"
            ForeColor       =   0
         End
         Begin PhotoDemon.sliderTextCombo sltOutlineOpacity 
            CausesValidation=   0   'False
            Height          =   495
            Left            =   4440
            TabIndex        =   29
            Top             =   30
            Width           =   2760
            _ExtentX        =   4868
            _ExtentY        =   873
            Max             =   100
            Value           =   100
            NotchPosition   =   2
            NotchValueCustom=   100
         End
         Begin PhotoDemon.colorSelector csOutline 
            Height          =   375
            Left            =   960
            TabIndex        =   30
            Top             =   540
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            curColor        =   4210752
         End
         Begin PhotoDemon.pdLabel lblText 
            Height          =   240
            Index           =   19
            Left            =   0
            Top             =   600
            Width           =   885
            _ExtentX        =   1984
            _ExtentY        =   503
            Alignment       =   1
            Caption         =   "color:"
            ForeColor       =   0
         End
         Begin PhotoDemon.pdLabel lblText 
            Height          =   240
            Index           =   20
            Left            =   3480
            Top             =   600
            Width           =   885
            _ExtentX        =   1984
            _ExtentY        =   503
            Alignment       =   1
            Caption         =   "width:"
            ForeColor       =   0
         End
         Begin PhotoDemon.sliderTextCombo sltOutlineWidth 
            CausesValidation=   0   'False
            Height          =   495
            Left            =   4440
            TabIndex        =   31
            Top             =   510
            Width           =   2760
            _ExtentX        =   4868
            _ExtentY        =   873
            Min             =   1
            Max             =   20
            SigDigits       =   1
            Value           =   1
            NotchPosition   =   1
            NotchValueCustom=   100
         End
         Begin PhotoDemon.pdLabel lblText 
            Height          =   240
            Index           =   21
            Left            =   0
            Top             =   1080
            Width           =   885
            _ExtentX        =   1984
            _ExtentY        =   503
            Alignment       =   1
            Caption         =   "corners:"
            ForeColor       =   0
         End
         Begin PhotoDemon.pdComboBox cboOutlineCorner 
            Height          =   375
            Left            =   960
            TabIndex        =   32
            Top             =   1035
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   635
         End
         Begin PhotoDemon.pdLabel lblText 
            Height          =   240
            Index           =   22
            Left            =   3480
            Top             =   1080
            Width           =   885
            _ExtentX        =   1984
            _ExtentY        =   503
            Alignment       =   1
            Caption         =   "caps:"
            ForeColor       =   0
         End
         Begin PhotoDemon.pdComboBox cboOutlineCaps 
            Height          =   375
            Left            =   4680
            TabIndex        =   33
            Top             =   1035
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   661
         End
      End
      Begin VB.PictureBox picAppearanceCategory 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   1500
         Index           =   0
         Left            =   1920
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   729
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   10935
         Begin PhotoDemon.pdLabel lblText 
            Height          =   240
            Index           =   7
            Left            =   0
            Top             =   120
            Width           =   885
            _ExtentX        =   1984
            _ExtentY        =   503
            Alignment       =   1
            Caption         =   "mode:"
            ForeColor       =   0
         End
         Begin PhotoDemon.pdComboBox cboFillMode 
            Height          =   375
            Left            =   960
            TabIndex        =   12
            Top             =   75
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   635
         End
         Begin PhotoDemon.pdLabel lblText 
            Height          =   240
            Index           =   9
            Left            =   3480
            Top             =   120
            Width           =   885
            _ExtentX        =   1984
            _ExtentY        =   503
            Alignment       =   1
            Caption         =   "opacity:"
            ForeColor       =   0
         End
         Begin PhotoDemon.sliderTextCombo sltFillOpacity 
            CausesValidation=   0   'False
            Height          =   495
            Left            =   4440
            TabIndex        =   14
            Top             =   30
            Width           =   2760
            _ExtentX        =   4868
            _ExtentY        =   873
            Max             =   100
            Value           =   100
            NotchPosition   =   2
            NotchValueCustom=   100
         End
         Begin VB.PictureBox picFillCategory 
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Height          =   1020
            Index           =   2
            Left            =   0
            ScaleHeight     =   68
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   729
            TabIndex        =   19
            Top             =   510
            Visible         =   0   'False
            Width           =   10935
            Begin PhotoDemon.colorSelector csPattern 
               Height          =   375
               Index           =   0
               Left            =   4440
               TabIndex        =   20
               Top             =   30
               Width           =   2775
               _ExtentX        =   4895
               _ExtentY        =   661
            End
            Begin PhotoDemon.pdLabel lblText 
               Height          =   240
               Index           =   11
               Left            =   3480
               Top             =   90
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               Alignment       =   1
               Caption         =   "color 1:"
               ForeColor       =   0
            End
            Begin PhotoDemon.colorSelector csPattern 
               Height          =   375
               Index           =   1
               Left            =   4440
               TabIndex        =   23
               Top             =   480
               Width           =   2775
               _ExtentX        =   4895
               _ExtentY        =   661
               curColor        =   0
            End
            Begin PhotoDemon.pdLabel lblText 
               Height          =   240
               Index           =   13
               Left            =   3480
               Top             =   510
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   503
               Alignment       =   1
               Caption         =   "color 2:"
               ForeColor       =   0
            End
            Begin PhotoDemon.pdLabel lblText 
               Height          =   240
               Index           =   14
               Left            =   0
               Top             =   90
               Width           =   885
               _ExtentX        =   1984
               _ExtentY        =   503
               Alignment       =   1
               Caption         =   "style:"
               ForeColor       =   0
            End
            Begin PhotoDemon.pdComboBox cboFillPattern 
               Height          =   375
               Left            =   960
               TabIndex        =   24
               Top             =   45
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   635
            End
         End
         Begin VB.PictureBox picFillCategory 
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Height          =   1020
            Index           =   1
            Left            =   0
            ScaleHeight     =   68
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   729
            TabIndex        =   17
            Top             =   510
            Visible         =   0   'False
            Width           =   10935
            Begin PhotoDemon.colorSelector csPlaceholder 
               Height          =   840
               Index           =   1
               Left            =   960
               TabIndex        =   18
               Top             =   30
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   1508
               curColor        =   0
            End
            Begin PhotoDemon.pdLabel lblText 
               Height          =   720
               Index           =   10
               Left            =   0
               Top             =   90
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   1270
               Alignment       =   1
               Caption         =   "style:"
               ForeColor       =   0
               Layout          =   1
            End
         End
         Begin VB.PictureBox picFillCategory 
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Height          =   1020
            Index           =   0
            Left            =   0
            ScaleHeight     =   68
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   729
            TabIndex        =   15
            Top             =   510
            Visible         =   0   'False
            Width           =   10935
            Begin PhotoDemon.colorSelector csFillColor 
               Height          =   840
               Left            =   960
               TabIndex        =   16
               Top             =   30
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   1508
               curColor        =   0
            End
            Begin PhotoDemon.pdLabel lblText 
               Height          =   240
               Index           =   6
               Left            =   0
               Top             =   90
               Width           =   885
               _ExtentX        =   1984
               _ExtentY        =   503
               Alignment       =   1
               Caption         =   "color:"
               ForeColor       =   0
            End
         End
         Begin VB.PictureBox picFillCategory 
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Height          =   1020
            Index           =   3
            Left            =   0
            ScaleHeight     =   68
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   729
            TabIndex        =   21
            Top             =   510
            Visible         =   0   'False
            Width           =   10935
            Begin PhotoDemon.colorSelector csPlaceholder 
               Height          =   840
               Index           =   0
               Left            =   960
               TabIndex        =   22
               Top             =   30
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   1508
               curColor        =   0
            End
            Begin PhotoDemon.pdLabel lblText 
               Height          =   720
               Index           =   12
               Left            =   0
               Top             =   90
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   1270
               Alignment       =   1
               Caption         =   "texture:"
               ForeColor       =   0
               Layout          =   1
            End
         End
      End
   End
   Begin VB.PictureBox picCategory 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   1500
      Index           =   1
      Left            =   8520
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   729
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   10935
      Begin PhotoDemon.textUpDown tudLineSpacing 
         Height          =   345
         Left            =   5160
         TabIndex        =   38
         Top             =   1020
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   609
         Min             =   -10
         SigDigits       =   2
      End
      Begin PhotoDemon.textUpDown tudMargin 
         Height          =   345
         Index           =   0
         Left            =   5160
         TabIndex        =   34
         Top             =   90
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   609
         Min             =   -1000
         Max             =   1000
      End
      Begin PhotoDemon.buttonStrip btsHAlignment 
         Height          =   435
         Left            =   1320
         TabIndex        =   4
         Top             =   60
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   767
         ColorScheme     =   1
      End
      Begin PhotoDemon.pdLabel lblText 
         Height          =   240
         Index           =   8
         Left            =   0
         Top             =   150
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         Alignment       =   1
         Caption         =   "alignment:"
         ForeColor       =   0
      End
      Begin PhotoDemon.buttonStrip btsVAlignment 
         Height          =   435
         Left            =   1320
         TabIndex        =   5
         Top             =   510
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   767
         ColorScheme     =   1
      End
      Begin PhotoDemon.pdLabel lblText 
         Height          =   240
         Index           =   0
         Left            =   0
         Top             =   1080
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   503
         Alignment       =   1
         Caption         =   "line wrap:"
         ForeColor       =   0
      End
      Begin PhotoDemon.pdComboBox cboWordWrap 
         Height          =   375
         Left            =   1320
         TabIndex        =   8
         Top             =   1020
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   661
      End
      Begin PhotoDemon.pdLabel lblText 
         Height          =   240
         Index           =   23
         Left            =   3360
         Top             =   150
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   503
         Alignment       =   1
         Caption         =   "h. padding:"
         ForeColor       =   0
      End
      Begin PhotoDemon.textUpDown tudMargin 
         Height          =   345
         Index           =   1
         Left            =   6120
         TabIndex        =   35
         Top             =   90
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   609
         Min             =   -1000
         Max             =   1000
      End
      Begin PhotoDemon.textUpDown tudMargin 
         Height          =   345
         Index           =   2
         Left            =   5160
         TabIndex        =   36
         Top             =   570
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   609
         Min             =   -1000
         Max             =   1000
      End
      Begin PhotoDemon.textUpDown tudMargin 
         Height          =   345
         Index           =   3
         Left            =   6120
         TabIndex        =   37
         Top             =   570
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   609
         Min             =   -1000
         Max             =   1000
      End
      Begin PhotoDemon.pdLabel lblText 
         Height          =   240
         Index           =   24
         Left            =   3360
         Top             =   630
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   503
         Alignment       =   1
         Caption         =   "v. padding:"
         ForeColor       =   0
      End
      Begin PhotoDemon.pdLabel lblText 
         Height          =   240
         Index           =   25
         Left            =   3480
         Top             =   1080
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   503
         Alignment       =   1
         Caption         =   "line spacing:"
         ForeColor       =   0
      End
   End
End
Attribute VB_Name = "toolpanel_FancyText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Advanced Typography Tool Panel
'Copyright 2013-2015 by Tanner Helland
'Created: 02/Oct/13
'Last updated: 13/May/15
'Last update: finish migrating all relevant controls to this dedicated form
'
'This form includes all user-editable settings for PD's Advanced Typography text tool.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'The value of all controls on this form are saved and loaded to file by this class
Private WithEvents lastUsedSettings As pdLastUsedSettings
Attribute lastUsedSettings.VB_VarHelpID = -1

'Current list of fonts, in pdStringStack format
Private userFontList As pdStringStack

Private Sub btnFontStyles_Click(Index As Integer)
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tool_Support.getToolBusyState
    If Not Tool_Support.canvasToolsAllowed Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.setToolBusyState True
    
    'Update whichever style was toggled
    Select Case Index
    
        'Bold
        Case 0
            pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_FontBold, btnFontStyles(Index).Value
        
        'Italic
        Case 1
            pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_FontItalic, btnFontStyles(Index).Value
        
        'Underline
        Case 2
            pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_FontUnderline, btnFontStyles(Index).Value
        
        'Strikeout
        Case 3
            pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_FontStrikeout, btnFontStyles(Index).Value
    
    End Select
    
    'Free the tool engine
    Tool_Support.setToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)

End Sub

Private Sub btnFontStyles_GotFocusAPI(Index As Integer)
    
    'Non-destructive effects are obviously not tracked if no images are loaded
    If g_OpenImageCount = 0 Then Exit Sub
    
    'Set Undo/Redo markers for whichever button was toggled
    Select Case Index
    
        'Bold
        Case 0
            Processor.flagInitialNDFXState_Text ptp_FontBold, btnFontStyles(Index).Value, pdImages(g_CurrentImage).getActiveLayerID
            
        'Italic
        Case 1
            Processor.flagInitialNDFXState_Text ptp_FontItalic, btnFontStyles(Index).Value, pdImages(g_CurrentImage).getActiveLayerID
        
        'Underline
        Case 2
            Processor.flagInitialNDFXState_Text ptp_FontUnderline, btnFontStyles(Index).Value, pdImages(g_CurrentImage).getActiveLayerID
        
        'Strikeout
        Case 3
            Processor.flagInitialNDFXState_Text ptp_FontStrikeout, btnFontStyles(Index).Value, pdImages(g_CurrentImage).getActiveLayerID
    
    End Select
    
End Sub

Private Sub btnFontStyles_LostFocusAPI(Index As Integer)
    
    'Evaluate Undo/Redo markers for whichever button was toggled
    Select Case Index
    
        'Bold
        Case 0
            If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_FontBold, btnFontStyles(Index).Value
            
        'Italic
        Case 1
            If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_FontItalic, btnFontStyles(Index).Value
        
        'Underline
        Case 2
            If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_FontUnderline, btnFontStyles(Index).Value
        
        'Strikeout
        Case 3
            If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_FontStrikeout, btnFontStyles(Index).Value
    
    End Select
    
End Sub

Private Sub btsAppearanceCategory_Click(ByVal buttonIndex As Long)
    
    'When the current category is changed, show the relevant panel and hide all others
    Dim i As Long
    For i = 0 To btsAppearanceCategory.ListCount - 1
        picAppearanceCategory(i).Visible = CBool(i = buttonIndex)
    Next i
    
End Sub

Private Sub btsCategory_Click(ByVal buttonIndex As Long)
    
    'When the current category is changed, show the relevant panel and hide all others
    Dim i As Long
    For i = 0 To btsCategory.ListCount - 1
        picCategory(i).Visible = CBool(i = buttonIndex)
    Next i
    
End Sub

Private Sub btsCharCategory_Click(ByVal buttonIndex As Long)
    
    'When the current category is changed, show the relevant panel and hide all others
    Dim i As Long
    For i = 0 To btsCharCategory.ListCount - 1
        picCharCategory(i).Visible = CBool(i = buttonIndex)
    Next i
    
End Sub

Private Sub btsHAlignment_Click(ByVal buttonIndex As Long)
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tool_Support.getToolBusyState
    If Not Tool_Support.canvasToolsAllowed Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.setToolBusyState True
        
    'Update the current layer text alignment
    pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_HorizontalAlignment, buttonIndex
    
    'Free the tool engine
    Tool_Support.setToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub btsHAlignment_GotFocusAPI()
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.flagInitialNDFXState_Text ptp_HorizontalAlignment, btsHAlignment.ListIndex, pdImages(g_CurrentImage).getActiveLayerID
End Sub

Private Sub btsHAlignment_LostFocusAPI()
    If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_HorizontalAlignment, btsHAlignment.ListIndex
End Sub

Private Sub btsVAlignment_Click(ByVal buttonIndex As Long)
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tool_Support.getToolBusyState
    If Not Tool_Support.canvasToolsAllowed Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.setToolBusyState True
        
    'Update the current layer text alignment
    pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_VerticalAlignment, buttonIndex
    
    'Free the tool engine
    Tool_Support.setToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub btsVAlignment_GotFocusAPI()
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.flagInitialNDFXState_Text ptp_VerticalAlignment, btsVAlignment.ListIndex, pdImages(g_CurrentImage).getActiveLayerID
End Sub

Private Sub btsVAlignment_LostFocusAPI()
    If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_VerticalAlignment, btsVAlignment.ListIndex
End Sub

Private Sub cboFillMode_Click()
    
    'When the current fill mode is changed, show the relevant panel and hide all others
    Dim i As Long
    For i = 0 To cboFillMode.ListCount - 2
        picFillCategory(i).Visible = CBool((i + 1) = cboFillMode.ListIndex)
    Next i
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tool_Support.getToolBusyState
    If Not Tool_Support.canvasToolsAllowed Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.setToolBusyState True
        
    'Update the current layer text alignment
    pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_FillMode, cboFillMode.ListIndex
    
    'Free the tool engine
    Tool_Support.setToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub cboFillMode_GotFocusAPI()
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.flagInitialNDFXState_Text ptp_FillMode, cboFillMode.ListIndex, pdImages(g_CurrentImage).getActiveLayerID
End Sub

Private Sub cboFillMode_LostFocusAPI()
    If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_FillMode, cboFillMode.ListIndex
End Sub

Private Sub cboFillPattern_Click()
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tool_Support.getToolBusyState
    If Not Tool_Support.canvasToolsAllowed Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.setToolBusyState True
        
    'Update the current layer text alignment
    pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_FillPattern, cboFillPattern.ListIndex
    
    'Free the tool engine
    Tool_Support.setToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub cboFillPattern_GotFocusAPI()
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.flagInitialNDFXState_Text ptp_FillPattern, cboFillPattern.ListIndex, pdImages(g_CurrentImage).getActiveLayerID
End Sub

Private Sub cboFillPattern_LostFocusAPI()
    If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_FillPattern, cboFillPattern.ListIndex
End Sub

Private Sub cboOutlineCaps_Click()

    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tool_Support.getToolBusyState
    If Not Tool_Support.canvasToolsAllowed Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.setToolBusyState True
        
    'Update the current layer text alignment
    pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_OutlineCaps, cboOutlineCaps.ListIndex
    
    'Free the tool engine
    Tool_Support.setToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)

End Sub

Private Sub cboOutlineCaps_GotFocusAPI()
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.flagInitialNDFXState_Text ptp_OutlineCaps, cboOutlineCaps.ListIndex, pdImages(g_CurrentImage).getActiveLayerID
End Sub

Private Sub cboOutlineCaps_LostFocusAPI()
    If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_OutlineCaps, cboOutlineCaps.ListIndex
End Sub

Private Sub cboOutlineCorner_Click()
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tool_Support.getToolBusyState
    If Not Tool_Support.canvasToolsAllowed Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.setToolBusyState True
        
    'Update the current layer text alignment
    pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_OutlineCorner, cboOutlineCorner.ListIndex
    
    'Free the tool engine
    Tool_Support.setToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub cboOutlineCorner_GotFocusAPI()
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.flagInitialNDFXState_Text ptp_OutlineCorner, cboOutlineCorner.ListIndex, pdImages(g_CurrentImage).getActiveLayerID
End Sub

Private Sub cboOutlineCorner_LostFocusAPI()
    If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_OutlineCorner, cboOutlineCorner.ListIndex
End Sub

Private Sub cboOutlineMode_Click()
    
    Dim otherOptionsVisible As Boolean
    otherOptionsVisible = CBool(cboOutlineMode.ListIndex <> 0)
    
    'Show/hide other outline options depending on the current mode.
    Dim i As Long
    
    For i = 18 To 22
        lblText(i).Visible = otherOptionsVisible
    Next i
    
    sltOutlineOpacity.Visible = otherOptionsVisible
    csOutline.Visible = otherOptionsVisible
    sltOutlineWidth.Visible = otherOptionsVisible
    cboOutlineCorner.Visible = otherOptionsVisible
    cboOutlineCaps.Visible = otherOptionsVisible
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tool_Support.getToolBusyState
    If Not Tool_Support.canvasToolsAllowed Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.setToolBusyState True
        
    'Update the current layer text alignment
    pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_OutlineMode, cboOutlineMode.ListIndex
    
    'Free the tool engine
    Tool_Support.setToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub cboOutlineMode_GotFocusAPI()
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.flagInitialNDFXState_Text ptp_OutlineMode, cboOutlineMode.ListIndex, pdImages(g_CurrentImage).getActiveLayerID
End Sub

Private Sub cboOutlineMode_LostFocusAPI()
    If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_OutlineMode, cboOutlineMode.ListIndex
End Sub

Private Sub cboTextFontFace_Click()
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tool_Support.getToolBusyState
    If Not Tool_Support.canvasToolsAllowed Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.setToolBusyState True
    
    'Update the current layer font size
    pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_FontFace, cboTextFontFace.List(cboTextFontFace.ListIndex)
    
    'Free the tool engine
    Tool_Support.setToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub cboTextFontFace_GotFocusAPI()
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.flagInitialNDFXState_Text ptp_FontFace, cboTextFontFace.List(cboTextFontFace.ListIndex), pdImages(g_CurrentImage).getActiveLayerID
End Sub

Private Sub cboTextFontFace_LostFocusAPI()
    If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_FontFace, cboTextFontFace.List(cboTextFontFace.ListIndex)
End Sub

Private Sub cboTextRenderingHint_Click()
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tool_Support.getToolBusyState
    If Not Tool_Support.canvasToolsAllowed Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.setToolBusyState True
    
    'Update the current layer text
    pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_TextAntialiasing, cboTextRenderingHint.ListIndex
    
    'Free the tool engine
    Tool_Support.setToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub cboTextRenderingHint_GotFocusAPI()
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.flagInitialNDFXState_Text ptp_TextAntialiasing, cboTextRenderingHint.ListIndex, pdImages(g_CurrentImage).getActiveLayerID
End Sub

Private Sub cboTextRenderingHint_LostFocusAPI()
    If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_TextAntialiasing, cboTextRenderingHint.ListIndex
End Sub

Private Sub cboWordWrap_Click()
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tool_Support.getToolBusyState
    If Not Tool_Support.canvasToolsAllowed Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.setToolBusyState True
    
    'Update the current layer text
    pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_WordWrap, cboWordWrap.ListIndex
    
    'Free the tool engine
    Tool_Support.setToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub cboWordWrap_GotFocusAPI()
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.flagInitialNDFXState_Text ptp_WordWrap, cboWordWrap.ListIndex, pdImages(g_CurrentImage).getActiveLayerID
End Sub

Private Sub cboWordWrap_LostFocusAPI()
    If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_WordWrap, cboWordWrap.ListIndex
End Sub

Private Sub chkBackground_Click()
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tool_Support.getToolBusyState
    If Not Tool_Support.canvasToolsAllowed Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.setToolBusyState True
        
    'Update the current layer text alignment
    pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_BackgroundMode, chkBackground.Value
    
    'Free the tool engine
    Tool_Support.setToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub chkBackground_GotFocusAPI()
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.flagInitialNDFXState_Text ptp_BackgroundMode, chkBackground.Value, pdImages(g_CurrentImage).getActiveLayerID
End Sub

Private Sub chkBackground_LostFocusAPI()
    If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_BackgroundMode, chkBackground.Value
End Sub

Private Sub chkHinting_Click()
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tool_Support.getToolBusyState
    If Not Tool_Support.canvasToolsAllowed Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.setToolBusyState True
    
    'Update the current layer text
    pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_TextHinting, CBool(chkHinting.Value)
    
    'Free the tool engine
    Tool_Support.setToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub chkHinting_GotFocusAPI()
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.flagInitialNDFXState_Text ptp_TextHinting, CBool(chkHinting.Value), pdImages(g_CurrentImage).getActiveLayerID
End Sub

Private Sub chkHinting_LostFocusAPI()
    If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_TextHinting, CBool(chkHinting.Value)
End Sub

Private Sub csBackground_ColorChanged()
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tool_Support.getToolBusyState
    If Not Tool_Support.canvasToolsAllowed Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.setToolBusyState True
        
    'Update the current layer text alignment
    pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_BackgroundColor, csBackground.Color
    
    'Free the tool engine
    Tool_Support.setToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub csBackground_GotFocusAPI()
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.flagInitialNDFXState_Text ptp_BackgroundColor, csBackground.Color, pdImages(g_CurrentImage).getActiveLayerID
End Sub

Private Sub csBackground_LostFocusAPI()
    If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_BackgroundColor, csBackground.Color
End Sub

Private Sub csFillColor_ColorChanged()
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tool_Support.getToolBusyState
    If Not Tool_Support.canvasToolsAllowed Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.setToolBusyState True
    
    'Update the current layer text
    pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_FontColor, csFillColor.Color
    
    'Free the tool engine
    Tool_Support.setToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub csFillColor_GotFocusAPI()
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.flagInitialNDFXState_Text ptp_FontColor, csFillColor.Color, pdImages(g_CurrentImage).getActiveLayerID
End Sub

Private Sub csFillColor_LostFocusAPI()
    If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_FontColor, csFillColor.Color
End Sub

Private Sub csOutline_ColorChanged()
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tool_Support.getToolBusyState
    If Not Tool_Support.canvasToolsAllowed Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.setToolBusyState True
        
    'Update the current layer text alignment
    pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_OutlineColor, csOutline.Color
    
    'Free the tool engine
    Tool_Support.setToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub csOutline_GotFocusAPI()
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.flagInitialNDFXState_Text ptp_OutlineColor, csOutline.Color, pdImages(g_CurrentImage).getActiveLayerID
End Sub

Private Sub csOutline_LostFocusAPI()
    If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_OutlineColor, csOutline.Color
End Sub

Private Sub csPattern_ColorChanged(Index As Integer)
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tool_Support.getToolBusyState
    If Not Tool_Support.canvasToolsAllowed Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.setToolBusyState True
        
    'Update the current layer text alignment
    If Index = 0 Then
        pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_PatternColor1, csPattern(Index).Color
    Else
        pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_PatternColor2, csPattern(Index).Color
    End If
    
    'Free the tool engine
    Tool_Support.setToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub csPattern_GotFocusAPI(Index As Integer)
    
    If g_OpenImageCount = 0 Then Exit Sub
    
    If Index = 0 Then
        Processor.flagInitialNDFXState_Text ptp_PatternColor1, csPattern(Index).Color, pdImages(g_CurrentImage).getActiveLayerID
    Else
        Processor.flagInitialNDFXState_Text ptp_PatternColor2, csPattern(Index).Color, pdImages(g_CurrentImage).getActiveLayerID
    End If
    
End Sub

Private Sub csPattern_LostFocusAPI(Index As Integer)
    If Tool_Support.canvasToolsAllowed Then
        If Index = 0 Then
            Processor.flagFinalNDFXState_Text ptp_PatternColor1, csPattern(Index).Color
        Else
            Processor.flagFinalNDFXState_Text ptp_PatternColor2, csPattern(Index).Color
        End If
    End If
End Sub

Private Sub Form_Load()

    'Generate a list of fonts
    If g_IsProgramRunning Then
        
        'Initialize the font list
        cboTextFontFace.initializeFontList
        
        'Set the system font as the default
        cboTextFontFace.setListIndexByString g_InterfaceFont, vbBinaryCompare
        
    End If
    
    'Draw the primary category selector
    btsCategory.AddItem "character", 0
    btsCategory.AddItem "paragraph", 1
    btsCategory.AddItem "visual", 2
    
    'I've already stubbed out a 4th options panel, but the vertical button list is *really* cramped, so another solution might be necessary
    
    'Draw the character sub-category selector
    btsCharCategory.AddItem "font", 0
    If g_IsVistaOrLater Then btsCharCategory.AddItem "OpenType", 1
    btsCharCategory.ListIndex = 0
    
    'Fill AA options
    cboTextRenderingHint.Clear
    cboTextRenderingHint.AddItem "none", 0
    cboTextRenderingHint.AddItem "normal", 1
    cboTextRenderingHint.AddItem "crisp", 2
    cboTextRenderingHint.ListIndex = 1
    
    'Draw font style buttons
    btnFontStyles(0).AssignImage "TEXT_BOLD"
    btnFontStyles(1).AssignImage "TEXT_ITALIC"
    btnFontStyles(2).AssignImage "TEXT_UNDERLINE"
    btnFontStyles(3).AssignImage "TEXT_STRIKE"
    
    'Draw alignment buttons
    btsHAlignment.AddItem "", 0
    btsHAlignment.AddItem "", 1
    btsHAlignment.AddItem "", 2
    
    btsHAlignment.AssignImageToItem 0, "TEXT_ALIGN_LEFT"
    btsHAlignment.AssignImageToItem 1, "TEXT_ALIGN_HCENTER"
    btsHAlignment.AssignImageToItem 2, "TEXT_ALIGN_RIGHT"
    
    btsVAlignment.AddItem "", 0
    btsVAlignment.AddItem "", 1
    btsVAlignment.AddItem "", 2
    
    btsVAlignment.AssignImageToItem 0, "TEXT_ALIGN_TOP"
    btsVAlignment.AssignImageToItem 1, "TEXT_ALIGN_VCENTER"
    btsVAlignment.AssignImageToItem 2, "TEXT_ALIGN_BOTTOM"
    
    'Fill wordwrap options
    cboWordWrap.Clear
    cboWordWrap.AddItem "none", 0
    cboWordWrap.AddItem "manual only", 1
    cboWordWrap.AddItem "characters", 2
    cboWordWrap.AddItem "words", 3
    cboWordWrap.ListIndex = 3
    
    'Draw the appearance sub-category selector
    btsAppearanceCategory.AddItem "fill", 0
    btsAppearanceCategory.AddItem "outline", 1
    btsAppearanceCategory.AddItem "background", 2
    btsAppearanceCategory.ListIndex = 0
    
    'Fill various appearance options
    cboFillMode.Clear
    cboFillMode.AddItem "none", 0
    cboFillMode.AddItem "color", 1
    cboFillMode.AddItem "gradient", 2
    cboFillMode.AddItem "pattern", 3
    cboFillMode.AddItem "texture", 4
    cboFillMode.ListIndex = 1
    
    'TODO: custom pattern dropdown, since we'll be using it elsewhere!
    cboFillPattern.Clear
    cboFillPattern.AddItem "horizontal"
    cboFillPattern.AddItem "vertical"
    cboFillPattern.AddItem "forward diagonal"
    cboFillPattern.AddItem "backward diagonal"
    cboFillPattern.AddItem "cross"
    cboFillPattern.AddItem "diagonal cross"
    cboFillPattern.ListIndex = 0
    
    cboOutlineMode.Clear
    cboOutlineMode.AddItem "invisible"
    cboOutlineMode.AddItem "solid"
    cboOutlineMode.AddItem "dashes"
    cboOutlineMode.AddItem "dots"
    cboOutlineMode.AddItem "dash + dot"
    cboOutlineMode.AddItem "dash + dot + dot"
    cboOutlineMode.ListIndex = 0
    
    cboOutlineCorner.Clear
    cboOutlineCorner.AddItem "miter"
    cboOutlineCorner.AddItem "bevel"
    cboOutlineCorner.AddItem "round"
    cboOutlineCorner.ListIndex = 0
    
    cboOutlineCaps.Clear
    cboOutlineCaps.AddItem "flat"
    cboOutlineCaps.AddItem "square"
    cboOutlineCaps.AddItem "round"
    cboOutlineCaps.AddItem "triangle"
    cboOutlineCaps.ListIndex = 0
        
    'Load any last-used settings for this form
    Set lastUsedSettings = New pdLastUsedSettings
    lastUsedSettings.setParentForm Me
    lastUsedSettings.loadAllControlValues
    
    'Update everything against the current theme.  This will also set tooltips for various controls.
    updateAgainstCurrentTheme

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'Save all last-used settings to file
    lastUsedSettings.saveAllControlValues
    lastUsedSettings.setParentForm Nothing
    
End Sub

Private Sub lastUsedSettings_ReadCustomPresetData()

    'Make sure the correct panels are shown
    btsCategory_Click btsCategory.ListIndex
    btsAppearanceCategory_Click btsAppearanceCategory.ListIndex
    btsCharCategory_Click btsCharCategory.ListIndex

End Sub

Private Sub sltBackgroundOpacity_Change()
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tool_Support.getToolBusyState
    If Not Tool_Support.canvasToolsAllowed Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.setToolBusyState True
        
    'Update the current layer text alignment
    pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_BackgroundOpacity, sltBackgroundOpacity.Value
    
    'Free the tool engine
    Tool_Support.setToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub sltBackgroundOpacity_GotFocusAPI()
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.flagInitialNDFXState_Text ptp_BackgroundOpacity, sltBackgroundOpacity.Value, pdImages(g_CurrentImage).getActiveLayerID
End Sub

Private Sub sltBackgroundOpacity_LostFocusAPI()
    If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_BackgroundOpacity, sltBackgroundOpacity.Value
End Sub

Private Sub sltFillOpacity_Change()

    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tool_Support.getToolBusyState
    If Not Tool_Support.canvasToolsAllowed Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.setToolBusyState True
        
    'Update the current layer text alignment
    pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_FillOpacity, sltFillOpacity.Value
    
    'Free the tool engine
    Tool_Support.setToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)

End Sub

Private Sub sltFillOpacity_GotFocusAPI()
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.flagInitialNDFXState_Text ptp_FillOpacity, sltFillOpacity.Value, pdImages(g_CurrentImage).getActiveLayerID
End Sub

Private Sub sltFillOpacity_LostFocusAPI()
    If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_FillOpacity, sltFillOpacity.Value
End Sub

Private Sub sltOutlineOpacity_Change()

    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tool_Support.getToolBusyState
    If Not Tool_Support.canvasToolsAllowed Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.setToolBusyState True
        
    'Update the current layer text alignment
    pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_OutlineOpacity, sltOutlineOpacity.Value
    
    'Free the tool engine
    Tool_Support.setToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)

End Sub

Private Sub sltOutlineOpacity_GotFocusAPI()
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.flagInitialNDFXState_Text ptp_OutlineOpacity, sltOutlineOpacity.Value, pdImages(g_CurrentImage).getActiveLayerID
End Sub

Private Sub sltOutlineOpacity_LostFocusAPI()
    If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_OutlineOpacity, sltOutlineOpacity.Value
End Sub

Private Sub sltOutlineWidth_Change()
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tool_Support.getToolBusyState
    If Not Tool_Support.canvasToolsAllowed Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.setToolBusyState True
        
    'Update the current setting
    pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_OutlineWidth, sltOutlineWidth.Value
    
    'Free the tool engine
    Tool_Support.setToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub sltOutlineWidth_GotFocusAPI()
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.flagInitialNDFXState_Text ptp_OutlineWidth, sltOutlineWidth.Value, pdImages(g_CurrentImage).getActiveLayerID
End Sub

Private Sub sltOutlineWidth_LostFocusAPI()
    If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_OutlineWidth, sltOutlineWidth.Value
End Sub

Private Sub tudLineSpacing_Change()
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tool_Support.getToolBusyState
    If Not Tool_Support.canvasToolsAllowed Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.setToolBusyState True
    
    'Update the setting
    pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_LineSpacing, tudLineSpacing.Value
    
    'Free the tool engine
    Tool_Support.setToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub tudLineSpacing_GotFocusAPI()
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.flagInitialNDFXState_Text ptp_LineSpacing, tudLineSpacing.Value, pdImages(g_CurrentImage).getActiveLayerID
End Sub

Private Sub tudLineSpacing_LostFocusAPI()
    If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_LineSpacing, tudLineSpacing.Value
End Sub

Private Sub tudMargin_Change(Index As Integer)
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tool_Support.getToolBusyState
    If Not Tool_Support.canvasToolsAllowed Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.setToolBusyState True
    
    'Update the current setting
    Select Case Index
    
        Case 0
            pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_MarginLeft, tudMargin(Index).Value
        
        Case 1
            pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_MarginRight, tudMargin(Index).Value
        
        Case 2
            pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_MarginTop, tudMargin(Index).Value
        
        Case 3
            pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_MarginBottom, tudMargin(Index).Value
    
    End Select
        
    'Free the tool engine
    Tool_Support.setToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub tudMargin_GotFocusAPI(Index As Integer)

    If g_OpenImageCount = 0 Then Exit Sub
    
    Select Case Index
    
        Case 0
            Processor.flagInitialNDFXState_Text ptp_MarginLeft, tudMargin(Index).Value, pdImages(g_CurrentImage).getActiveLayerID
        
        Case 1
            Processor.flagInitialNDFXState_Text ptp_MarginRight, tudMargin(Index).Value, pdImages(g_CurrentImage).getActiveLayerID
        
        Case 2
            Processor.flagInitialNDFXState_Text ptp_MarginTop, tudMargin(Index).Value, pdImages(g_CurrentImage).getActiveLayerID
        
        Case 3
            Processor.flagInitialNDFXState_Text ptp_MarginBottom, tudMargin(Index).Value, pdImages(g_CurrentImage).getActiveLayerID
        
    End Select
    
End Sub

Private Sub tudMargin_LostFocusAPI(Index As Integer)
    
    If Tool_Support.canvasToolsAllowed Then
        
        Select Case Index
        
            Case 0
                Processor.flagFinalNDFXState_Text ptp_MarginLeft, tudMargin(Index).Value
            
            Case 1
                Processor.flagFinalNDFXState_Text ptp_MarginRight, tudMargin(Index).Value
            
            Case 2
                Processor.flagFinalNDFXState_Text ptp_MarginTop, tudMargin(Index).Value
            
            Case 3
                Processor.flagFinalNDFXState_Text ptp_MarginBottom, tudMargin(Index).Value
        
        End Select
        
    End If
    
End Sub

Private Sub tudTextFontSize_Change()
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tool_Support.getToolBusyState
    If Not Tool_Support.canvasToolsAllowed Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.setToolBusyState True
    
    'Update the current layer font size
    pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_FontSize, tudTextFontSize.Value
    
    'Free the tool engine
    Tool_Support.setToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub tudTextFontSize_GotFocusAPI()
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.flagInitialNDFXState_Text ptp_FontSize, tudTextFontSize.Value, pdImages(g_CurrentImage).getActiveLayerID
End Sub

Private Sub tudTextFontSize_LostFocusAPI()
    If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_FontSize, tudTextFontSize.Value
End Sub

Private Sub txtTextTool_Change()
    
    'If tool changes are not allowed, exit.
    ' NOTE: this will also check tool busy status, via Tool_Support.getToolBusyState
    If Not Tool_Support.canvasToolsAllowed Then Exit Sub
    
    'Mark the tool engine as busy
    Tool_Support.setToolBusyState True
    
    'Update the current layer text
    pdImages(g_CurrentImage).getActiveLayer.setTextLayerProperty ptp_Text, txtTextTool.Text
    
    'Free the tool engine
    Tool_Support.setToolBusyState False
    
    'Redraw the viewport
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
        
End Sub

Private Sub txtTextTool_GotFocusAPI()
    If g_OpenImageCount = 0 Then Exit Sub
    Processor.flagInitialNDFXState_Text ptp_Text, txtTextTool.Text, pdImages(g_CurrentImage).getActiveLayerID
End Sub

Private Sub txtTextTool_LostFocusAPI()
    If Tool_Support.canvasToolsAllowed Then Processor.flagFinalNDFXState_Text ptp_Text, txtTextTool.Text
End Sub


'Updating against the current theme accomplishes a number of things:
' 1) All user-drawn controls are redrawn according to the current g_Themer settings.
' 2) All tooltips and captions are translated according to the current language.
' 3) MakeFormPretty is called, which redraws the form itself according to any theme and/or system settings.
'
'This function is called at least once, at Form_Load, but can be called again if the active language or theme changes.
Public Sub updateAgainstCurrentTheme()

    'Start by redrawing the form according to current theme and translation settings.  (This function also takes care of
    ' any common controls that may still exist in the program.)
    makeFormPretty Me

End Sub
