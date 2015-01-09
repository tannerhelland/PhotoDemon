VERSION 5.00
Begin VB.Form toolbar_Options 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Tools"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15045
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
   ScaleHeight     =   97
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1003
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picTools 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   3
      Left            =   15
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1230
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   15
      Visible         =   0   'False
      Width           =   18450
      Begin PhotoDemon.pdComboBox cboSelSmoothing 
         Height          =   375
         Left            =   2760
         TabIndex        =   77
         Top             =   390
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
      End
      Begin PhotoDemon.pdComboBox cboSelRender 
         Height          =   375
         Left            =   120
         TabIndex        =   76
         Top             =   390
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   635
      End
      Begin PhotoDemon.colorSelector csSelectionHighlight 
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   2445
         _ExtentX        =   3916
         _ExtentY        =   661
      End
      Begin PhotoDemon.sliderTextCombo sltSelectionFeathering 
         CausesValidation=   0   'False
         Height          =   495
         Left            =   2640
         TabIndex        =   25
         Top             =   840
         Visible         =   0   'False
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   873
         Max             =   100
      End
      Begin VB.PictureBox picSelectionSubcontainer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1470
         Index           =   5
         Left            =   5340
         ScaleHeight     =   98
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   665
         TabIndex        =   61
         Top             =   0
         Width           =   9975
         Begin PhotoDemon.pdComboBox cboWandCompare 
            Height          =   375
            Left            =   3270
            TabIndex        =   83
            Top             =   855
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   635
         End
         Begin PhotoDemon.buttonStrip btsWandArea 
            Height          =   825
            Left            =   120
            TabIndex        =   63
            Top             =   405
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
         Begin PhotoDemon.sliderTextCombo sltWandTolerance 
            CausesValidation=   0   'False
            Height          =   495
            Left            =   3120
            TabIndex        =   65
            Top             =   360
            Width           =   2670
            _ExtentX        =   4710
            _ExtentY        =   873
            Max             =   255
            SigDigits       =   1
         End
         Begin PhotoDemon.buttonStrip btsWandMerge 
            Height          =   825
            Left            =   6120
            TabIndex        =   67
            Top             =   405
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
         Begin VB.Label lblSelection 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "sampling area"
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
            Index           =   16
            Left            =   6120
            TabIndex        =   66
            Top             =   60
            Width           =   1215
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
            Left            =   3240
            TabIndex        =   64
            Top             =   60
            Width           =   795
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
            Index           =   14
            Left            =   120
            TabIndex        =   62
            Top             =   60
            Width           =   390
         End
      End
      Begin VB.PictureBox picSelectionSubcontainer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1470
         Index           =   4
         Left            =   5340
         ScaleHeight     =   98
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   665
         TabIndex        =   58
         Top             =   0
         Width           =   9975
         Begin PhotoDemon.pdComboBox cboSelArea 
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   81
            Top             =   390
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   635
         End
         Begin PhotoDemon.sliderTextCombo sltSelectionBorder 
            CausesValidation=   0   'False
            Height          =   495
            Index           =   4
            Left            =   0
            TabIndex        =   59
            Top             =   840
            Width           =   2670
            _ExtentX        =   4710
            _ExtentY        =   873
            Min             =   1
            Max             =   10000
            Value           =   1
         End
         Begin PhotoDemon.sliderTextCombo sltSmoothStroke 
            CausesValidation=   0   'False
            Height          =   495
            Left            =   2760
            TabIndex        =   74
            Top             =   360
            Visible         =   0   'False
            Width           =   2670
            _ExtentX        =   4710
            _ExtentY        =   873
            Max             =   1
            SigDigits       =   2
         End
         Begin VB.Label lblSelection 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "stroke smoothing"
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
            Index           =   19
            Left            =   2910
            TabIndex        =   75
            Top             =   60
            Visible         =   0   'False
            Width           =   1470
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
            Index           =   12
            Left            =   120
            TabIndex        =   60
            Top             =   60
            Width           =   390
         End
      End
      Begin VB.PictureBox picSelectionSubcontainer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1470
         Index           =   3
         Left            =   5340
         ScaleHeight     =   98
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   665
         TabIndex        =   55
         Top             =   0
         Width           =   9975
         Begin PhotoDemon.pdComboBox cboSelArea 
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   82
            Top             =   390
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   635
         End
         Begin PhotoDemon.sliderTextCombo sltSelectionBorder 
            CausesValidation=   0   'False
            Height          =   495
            Index           =   3
            Left            =   0
            TabIndex        =   56
            Top             =   840
            Width           =   2670
            _ExtentX        =   4710
            _ExtentY        =   873
            Min             =   1
            Max             =   10000
            Value           =   1
         End
         Begin PhotoDemon.sliderTextCombo sltPolygonCurvature 
            CausesValidation=   0   'False
            Height          =   495
            Left            =   2760
            TabIndex        =   70
            Top             =   360
            Width           =   2670
            _ExtentX        =   4710
            _ExtentY        =   873
            Max             =   1
            SigDigits       =   2
         End
         Begin VB.Label lblSelection 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "curvature"
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
            Index           =   17
            Left            =   2910
            TabIndex        =   71
            Top             =   60
            Width           =   810
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
            Index           =   15
            Left            =   120
            TabIndex        =   57
            Top             =   60
            Width           =   390
         End
      End
      Begin VB.PictureBox picSelectionSubcontainer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1470
         Index           =   2
         Left            =   5340
         ScaleHeight     =   98
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   665
         TabIndex        =   46
         Top             =   0
         Width           =   9975
         Begin PhotoDemon.pdComboBox cboSelArea 
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   78
            Top             =   390
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   635
         End
         Begin PhotoDemon.sliderTextCombo sltSelectionBorder 
            CausesValidation=   0   'False
            Height          =   495
            Index           =   2
            Left            =   0
            TabIndex        =   47
            Top             =   840
            Width           =   2670
            _ExtentX        =   4710
            _ExtentY        =   873
            Min             =   1
            Max             =   10000
            Value           =   1
         End
         Begin PhotoDemon.textUpDown tudSel 
            Height          =   345
            Index           =   8
            Left            =   2820
            TabIndex        =   48
            Top             =   375
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   714
            Min             =   -30000
            Max             =   30000
         End
         Begin PhotoDemon.textUpDown tudSel 
            Height          =   345
            Index           =   9
            Left            =   2820
            TabIndex        =   49
            Top             =   885
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   714
            Min             =   -30000
            Max             =   30000
         End
         Begin PhotoDemon.textUpDown tudSel 
            Height          =   345
            Index           =   10
            Left            =   4380
            TabIndex        =   50
            Top             =   375
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   714
            Min             =   -30000
            Max             =   30000
         End
         Begin PhotoDemon.textUpDown tudSel 
            Height          =   345
            Index           =   11
            Left            =   4380
            TabIndex        =   51
            Top             =   885
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   714
            Min             =   -30000
            Max             =   30000
         End
         Begin PhotoDemon.sliderTextCombo sltSelectionLineWidth 
            CausesValidation=   0   'False
            Height          =   495
            Left            =   5880
            TabIndex        =   72
            Top             =   360
            Width           =   2670
            _ExtentX        =   4710
            _ExtentY        =   873
            Min             =   1
            Max             =   10000
            Value           =   1
         End
         Begin VB.Label lblSelection 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "line width"
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
            Index           =   18
            Left            =   6000
            TabIndex        =   73
            Top             =   60
            Width           =   825
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
            Index           =   11
            Left            =   120
            TabIndex        =   54
            Top             =   60
            Width           =   390
         End
         Begin VB.Label lblSelection 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "2nd point (x, y)"
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
            Index           =   10
            Left            =   4380
            TabIndex        =   53
            Top             =   60
            Width           =   1305
         End
         Begin VB.Label lblSelection 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "1st point (x, y)"
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
            Index           =   9
            Left            =   2820
            TabIndex        =   52
            Top             =   60
            Width           =   1245
         End
      End
      Begin VB.PictureBox picSelectionSubcontainer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1470
         Index           =   1
         Left            =   5340
         ScaleHeight     =   98
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   665
         TabIndex        =   37
         Top             =   0
         Width           =   9975
         Begin PhotoDemon.pdComboBox cboSelArea 
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   79
            Top             =   390
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   635
         End
         Begin PhotoDemon.sliderTextCombo sltSelectionBorder 
            CausesValidation=   0   'False
            Height          =   495
            Index           =   1
            Left            =   0
            TabIndex        =   38
            Top             =   840
            Width           =   2670
            _ExtentX        =   4710
            _ExtentY        =   873
            Min             =   1
            Max             =   10000
            Value           =   1
         End
         Begin PhotoDemon.textUpDown tudSel 
            Height          =   345
            Index           =   4
            Left            =   2820
            TabIndex        =   39
            Top             =   375
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   714
            Min             =   -30000
            Max             =   30000
         End
         Begin PhotoDemon.textUpDown tudSel 
            Height          =   345
            Index           =   5
            Left            =   2820
            TabIndex        =   40
            Top             =   885
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   714
            Min             =   -30000
            Max             =   30000
         End
         Begin PhotoDemon.textUpDown tudSel 
            Height          =   345
            Index           =   6
            Left            =   4380
            TabIndex        =   41
            Top             =   375
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   714
            Min             =   -30000
            Max             =   30000
         End
         Begin PhotoDemon.textUpDown tudSel 
            Height          =   345
            Index           =   7
            Left            =   4380
            TabIndex        =   42
            Top             =   885
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   714
            Min             =   -30000
            Max             =   30000
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
            Index           =   7
            Left            =   120
            TabIndex        =   45
            Top             =   60
            Width           =   390
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
            Index           =   3
            Left            =   4380
            TabIndex        =   44
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
            Index           =   0
            Left            =   2820
            TabIndex        =   43
            Top             =   60
            Width           =   1170
         End
      End
      Begin VB.PictureBox picSelectionSubcontainer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1470
         Index           =   0
         Left            =   5340
         ScaleHeight     =   98
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   665
         TabIndex        =   28
         Top             =   0
         Width           =   9975
         Begin PhotoDemon.pdComboBox cboSelArea 
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   80
            Top             =   390
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   635
         End
         Begin PhotoDemon.sliderTextCombo sltSelectionBorder 
            CausesValidation=   0   'False
            Height          =   495
            Index           =   0
            Left            =   0
            TabIndex        =   29
            Top             =   840
            Width           =   2670
            _ExtentX        =   4710
            _ExtentY        =   873
            Min             =   1
            Max             =   10000
            Value           =   1
         End
         Begin PhotoDemon.textUpDown tudSel 
            Height          =   345
            Index           =   0
            Left            =   2820
            TabIndex        =   31
            Top             =   375
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   714
            Min             =   -30000
            Max             =   30000
         End
         Begin PhotoDemon.textUpDown tudSel 
            Height          =   345
            Index           =   1
            Left            =   2820
            TabIndex        =   32
            Top             =   885
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   714
            Min             =   -30000
            Max             =   30000
         End
         Begin PhotoDemon.textUpDown tudSel 
            Height          =   345
            Index           =   2
            Left            =   4380
            TabIndex        =   33
            Top             =   375
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   714
            Min             =   -30000
            Max             =   30000
         End
         Begin PhotoDemon.textUpDown tudSel 
            Height          =   345
            Index           =   3
            Left            =   4380
            TabIndex        =   34
            Top             =   885
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   714
            Min             =   -30000
            Max             =   30000
         End
         Begin PhotoDemon.sliderTextCombo sltCornerRounding 
            CausesValidation=   0   'False
            Height          =   495
            Left            =   5760
            TabIndex        =   68
            Top             =   345
            Width           =   2670
            _ExtentX        =   4710
            _ExtentY        =   873
            Max             =   1
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
            Left            =   5880
            TabIndex        =   69
            Top             =   60
            Width           =   1365
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
            Left            =   2820
            TabIndex        =   36
            Top             =   60
            Width           =   1170
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
            Left            =   4380
            TabIndex        =   35
            Top             =   60
            Width           =   915
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
            Left            =   120
            TabIndex        =   30
            Top             =   60
            Width           =   390
         End
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
         Index           =   13
         Left            =   2760
         TabIndex        =   27
         Top             =   60
         Width           =   885
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
         Index           =   8
         Left            =   120
         TabIndex        =   26
         Top             =   60
         Width           =   1005
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
      ScaleWidth      =   1230
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      Visible         =   0   'False
      Width           =   18450
   End
   Begin VB.PictureBox picTools 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Index           =   2
      Left            =   15
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   950
      TabIndex        =   8
      Top             =   15
      Visible         =   0   'False
      Width           =   14250
      Begin PhotoDemon.sliderTextCombo sltQuickFix 
         CausesValidation=   0   'False
         Height          =   495
         Index           =   0
         Left            =   1380
         TabIndex        =   10
         Top             =   90
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   873
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
         TabIndex        =   11
         Top             =   705
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   873
         Min             =   -100
         Max             =   100
      End
      Begin PhotoDemon.sliderTextCombo sltQuickFix 
         CausesValidation=   0   'False
         Height          =   495
         Index           =   2
         Left            =   5640
         TabIndex        =   13
         Top             =   90
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   873
         Min             =   -100
         Max             =   100
      End
      Begin PhotoDemon.sliderTextCombo sltQuickFix 
         CausesValidation=   0   'False
         Height          =   495
         Index           =   3
         Left            =   5640
         TabIndex        =   15
         Top             =   705
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   873
         Min             =   -100
         Max             =   100
      End
      Begin PhotoDemon.jcbutton cmdQuickFix 
         Height          =   570
         Index           =   0
         Left            =   13080
         TabIndex        =   17
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
         PictureNormal   =   "VBP_ToolbarTools.frx":0000
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         ColorScheme     =   3
      End
      Begin PhotoDemon.jcbutton cmdQuickFix 
         Height          =   570
         Index           =   1
         Left            =   13080
         TabIndex        =   18
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
         PictureNormal   =   "VBP_ToolbarTools.frx":0D52
         PictureEffectOnDown=   0
         CaptionEffects  =   0
         ColorScheme     =   3
      End
      Begin PhotoDemon.sliderTextCombo sltQuickFix 
         CausesValidation=   0   'False
         Height          =   495
         Index           =   4
         Left            =   9960
         TabIndex        =   19
         Top             =   90
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   873
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
         TabIndex        =   20
         Top             =   705
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   873
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   16
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
         TabIndex        =   14
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
         TabIndex        =   12
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
         TabIndex        =   9
         Top             =   195
         Width           =   855
      End
   End
   Begin VB.PictureBox picTools 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Index           =   1
      Left            =   15
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   950
      TabIndex        =   1
      Top             =   15
      Visible         =   0   'False
      Width           =   14250
      Begin PhotoDemon.smartCheckBox chkLayerBorder 
         Height          =   330
         Left            =   7080
         TabIndex        =   3
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
         TabIndex        =   4
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
         TabIndex        =   6
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
         TabIndex        =   7
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
         TabIndex        =   5
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
         TabIndex        =   2
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
Attribute VB_Name = "toolbar_Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Tools Toolbox
'Copyright 2013-2015 by Tanner Helland
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
        Viewport_Engine.Stage3_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
    
End Sub

Private Sub btsWandMerge_Click(ByVal buttonIndex As Long)

    'If a selection is already active, change its type to match the current option, then redraw it
    If selectionsAllowed(False) Then
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_WAND_SAMPLE_MERGED, buttonIndex
        Viewport_Engine.Stage3_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If

End Sub

Private Sub cboSelArea_Click(Index As Integer)

    If cboSelArea(Index).ListIndex = sBorder Then
        sltSelectionBorder(Index).Visible = True
    Else
        sltSelectionBorder(Index).Visible = False
    End If
    
    'If a selection is already active, change its type to match the current selection, then redraw it
    If selectionsAllowed(False) Then
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_AREA, cboSelArea(Index).ListIndex
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_BORDER_WIDTH, sltSelectionBorder(Index).Value
        Viewport_Engine.Stage3_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
    
End Sub

Private Sub cboSelRender_Click()

    'Show or hide the color selector, as appropriate
    If cboSelRender.ListIndex = SELECTION_RENDER_HIGHLIGHT Then
        csSelectionHighlight.Visible = True
    Else
        csSelectionHighlight.Visible = False
    End If
    
    'Redraw the viewport
    If selectionsAllowed(False) Then Viewport_Engine.Stage3_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)

End Sub

Private Sub cboSelSmoothing_Click()

    updateSelectionPanelLayout
    
    'If a selection is already active, change its type to match the current selection, then redraw it
    If selectionsAllowed(False) Then
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_SMOOTHING, cboSelSmoothing.ListIndex
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_FEATHERING_RADIUS, sltSelectionFeathering.Value
        Viewport_Engine.Stage3_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
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
        Viewport_Engine.Stage3_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
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
    Viewport_Engine.Stage1_InitializeBuffer pdImages(g_CurrentImage), FormMain.mainCanvas(0), "Layer border toggle"
End Sub

'Show/hide layer transform nodes while using the move tool
Private Sub chkLayerNodes_Click()
    Viewport_Engine.Stage1_InitializeBuffer pdImages(g_CurrentImage), FormMain.mainCanvas(0), "Layer nodes toggle"
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
    Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)

End Sub

Private Sub csSelectionHighlight_ColorChanged()
    
    'Redraw the viewport
    If selectionsAllowed(False) Then Viewport_Engine.Stage3_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    
End Sub

Private Sub Form_Load()

    Dim i As Long
        
    'INITIALIZE ALL TOOLS
    
        'Selection visual styles (Highlight, Lightbox, or Outline)
        toolbar_Options.cboSelRender.assignTooltip "Click to change the way selections are rendered onto the image canvas.  This has no bearing on selection contents - only the way they appear while editing."
        toolbar_Options.cboSelRender.AddItem " Highlight", 0
        toolbar_Options.cboSelRender.AddItem " Lightbox", 1
        toolbar_Options.cboSelRender.AddItem " Outline", 2
        toolbar_Options.cboSelRender.ListIndex = 0
        
        csSelectionHighlight.Color = RGB(255, 58, 72)
        csSelectionHighlight.Visible = True
        
        'Selection smoothing (currently none, antialiased, fully feathered)
        toolbar_Options.cboSelSmoothing.assignTooltip "This option controls how smoothly a selection blends with its surroundings."
        toolbar_Options.cboSelSmoothing.AddItem " None", 0
        toolbar_Options.cboSelSmoothing.AddItem " Antialiased", 1
        toolbar_Options.cboSelSmoothing.AddItem " Feathered", 2
        toolbar_Options.cboSelSmoothing.ListIndex = 1
        
        'Selection types (currently interior, exterior, border)
        For i = 0 To cboSelArea.Count - 1
            toolbar_Options.cboSelArea(i).AddItem " Interior", 0
            toolbar_Options.cboSelArea(i).AddItem " Exterior", 1
            toolbar_Options.cboSelArea(i).AddItem " Border", 2
            toolbar_Options.cboSelArea(i).ListIndex = 0
            
            toolbar_Options.cboSelArea(i).assignTooltip "These options control the area affected by a selection.  The selection can be modified on-canvas while any of these settings are active.  For more advanced selection adjustments, use the Select menu."
            toolbar_Options.sltSelectionBorder(i).assignTooltip "This option adjusts the width of the selection border."
        Next i
        
        toolbar_Options.sltSelectionFeathering.assignTooltip "This feathering slider allows for immediate feathering adjustments.  For performance reasons, it is limited to small radii.  For larger feathering radii, please use the Select -> Feathering menu."
        toolbar_Options.sltCornerRounding.assignTooltip "This option adjusts the roundness of a rectangular selection's corners."
        toolbar_Options.sltSelectionLineWidth.assignTooltip "This option adjusts the width of a line selection."
                
        toolbar_Options.sltPolygonCurvature.assignTooltip "This option adjusts the curvature, if any, of a polygon selection's sides."
        toolbar_Options.sltSmoothStroke.assignTooltip "This option increases the smoothness of a hand-drawn lasso selection."
        toolbar_Options.sltWandTolerance.assignTooltip "Tolerance controls how similar two pixels must be before adding them to a magic wand selection."
        
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
        cboWandCompare.AddItem " Color", 1
        cboWandCompare.AddItem " Luminance", 2, True
        cboWandCompare.AddItem " Red", 3
        cboWandCompare.AddItem " Green", 4
        cboWandCompare.AddItem " Blue", 5
        cboWandCompare.AddItem " Alpha", 6
        cboWandCompare.ListIndex = 1
        cboWandCompare.assignTooltip "This option controls which criteria the magic wand uses to determine whether a pixel should be added to the current selection."
        
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
    For i = 0 To tudSel.Count - 1
        tudSel(i) = 0
    Next i

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
        Viewport_Engine.Stage3_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
End Sub

Private Sub sltPolygonCurvature_Change()
    If selectionsAllowed(True) Then
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_POLYGON_CURVATURE, sltPolygonCurvature.Value
        Viewport_Engine.Stage3_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
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
        Viewport_Engine.Stage2_CompositeAllLayers pdImages(g_CurrentImage), FormMain.mainCanvas(0)
        
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

Private Sub sltSelectionBorder_Change(Index As Integer)
    If selectionsAllowed(False) Then
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_BORDER_WIDTH, sltSelectionBorder(Index).Value
        Viewport_Engine.Stage3_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
End Sub

Private Sub sltSelectionFeathering_Change()
    If selectionsAllowed(False) Then
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_FEATHERING_RADIUS, sltSelectionFeathering.Value
        Viewport_Engine.Stage3_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
End Sub

Private Sub sltSelectionLineWidth_Change()
    If selectionsAllowed(True) Then
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_LINE_WIDTH, sltSelectionLineWidth.Value
        Viewport_Engine.Stage3_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
End Sub

'When certain selection settings are enabled or disabled, corresponding controls are shown or hidden.  To keep the
' panel concise and clean, we move other controls up or down depending on what controls are visible.
Public Sub updateSelectionPanelLayout()

    'Display the feathering slider as necessary
    If cboSelSmoothing.ListIndex = sFullyFeathered Then
        sltSelectionFeathering.Visible = True
    Else
        sltSelectionFeathering.Visible = False
    End If
    
    'Display the border slider as necessary
    If (Selection_Handler.getSelectionSubPanelFromCurrentTool < cboSelArea.Count - 1) And (Selection_Handler.getSelectionSubPanelFromCurrentTool > 0) Then
        If cboSelArea(Selection_Handler.getSelectionSubPanelFromCurrentTool).ListIndex = sBorder Then
            sltSelectionBorder(Selection_Handler.getSelectionSubPanelFromCurrentTool).Visible = True
        Else
            sltSelectionBorder(Selection_Handler.getSelectionSubPanelFromCurrentTool).Visible = False
        End If
    End If
    
    'Finally, the magic wand selection type is unique because it cannot display an outline.  (This might someday be possible,
    ' but we would need to construct the border region ourselves - and I'm not a huge fan of the work involved.)
    ' As such, when activating that tool, we need to remove the Outline option, and when switching to a different tool, we need
    ' to restore the option.
    If g_CurrentTool = SELECT_WAND Then
    
        'See if the combo box is already modified
        If cboSelRender.ListCount = 3 Then
            
            'Remove the "outline" option
            If toolbar_Options.cboSelRender.ListIndex = 2 Then toolbar_Options.cboSelRender.ListIndex = 0
            toolbar_Options.cboSelRender.RemoveItem 2
            
        End If
    
    Else
    
        'See if the combo box is missing an entry
        If cboSelRender.ListCount = 2 Then
            toolbar_Options.cboSelRender.AddItem " Outline", 2
        End If
    
    End If
    
End Sub

Private Sub sltSmoothStroke_Change()
    If selectionsAllowed(False) Then
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_SMOOTH_STROKE, sltSmoothStroke.Value
        Viewport_Engine.Stage3_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
    End If
End Sub

Private Sub sltWandTolerance_Change()
    If selectionsAllowed(False) Then
        pdImages(g_CurrentImage).mainSelection.setSelectionProperty SP_WAND_TOLERANCE, sltWandTolerance.Value
        Viewport_Engine.Stage3_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
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
            Viewport_Engine.Stage3_CompositeCanvas pdImages(g_CurrentImage), FormMain.mainCanvas(0)
        End If
    End If
End Sub

'External functions can use this to re-theme this form at run-time (important when changing languages, for example)
Public Sub requestMakeFormPretty()
    makeFormPretty Me, m_ToolTip
End Sub
