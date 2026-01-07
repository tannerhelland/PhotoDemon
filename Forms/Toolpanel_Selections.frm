VERSION 5.00
Begin VB.Form toolpanel_Selections 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   6855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13740
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   Icon            =   "Toolpanel_Selections.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   457
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   916
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin PhotoDemon.pdCheckBox chkAppearance 
      Height          =   345
      Index           =   0
      Left            =   120
      TabIndex        =   56
      Top             =   420
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   609
      Caption         =   "animate"
   End
   Begin PhotoDemon.pdButtonStrip btsCombine 
      Height          =   375
      Left            =   2160
      TabIndex        =   55
      Top             =   375
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
   End
   Begin PhotoDemon.pdDropDown cboSelSmoothing 
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   375
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdTitle ttlPanel 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   661
      Caption         =   "appearance"
      Value           =   0   'False
   End
   Begin PhotoDemon.pdContainer cntrPopOut 
      Height          =   885
      Index           =   1
      Left            =   0
      Top             =   4320
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1561
      Begin PhotoDemon.pdButtonToolbox cmdFlyoutLock 
         Height          =   390
         Index           =   1
         Left            =   2640
         TabIndex        =   6
         Top             =   330
         Width           =   390
         _ExtentX        =   1111
         _ExtentY        =   1111
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdSlider sltSelectionFeathering 
         CausesValidation=   0   'False
         Height          =   765
         Left            =   150
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   1349
         Caption         =   "radius"
         FontSizeCaption =   10
         Max             =   100
      End
      Begin PhotoDemon.pdLabel lblNoOptions 
         Height          =   375
         Index           =   1
         Left            =   0
         Top             =   180
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         Alignment       =   2
         Caption         =   "(no additional options)"
      End
   End
   Begin PhotoDemon.pdContainer cntrPopOut 
      Height          =   2475
      Index           =   0
      Left            =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   4366
      Begin PhotoDemon.pdDropDown ddAppearance 
         Height          =   705
         Index           =   0
         Left            =   120
         TabIndex        =   57
         Top             =   0
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1244
         Caption         =   "fill interior"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdButtonToolbox cmdFlyoutLock 
         Height          =   390
         Index           =   0
         Left            =   2880
         TabIndex        =   7
         Top             =   1980
         Width           =   390
         _ExtentX        =   1111
         _ExtentY        =   1111
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdSpinner spnOpacity 
         Height          =   375
         Index           =   0
         Left            =   1440
         TabIndex        =   9
         Top             =   780
         Width           =   1125
         _ExtentX        =   1931
         _ExtentY        =   661
         DefaultValue    =   50
         Min             =   1
         Max             =   100
         Value           =   50
      End
      Begin PhotoDemon.pdColorSelector csSelection 
         Height          =   330
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   780
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   582
      End
      Begin PhotoDemon.pdSpinner spnOpacity 
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   11
         Top             =   1980
         Width           =   1125
         _ExtentX        =   1931
         _ExtentY        =   661
         DefaultValue    =   50
         Min             =   1
         Max             =   100
         Value           =   50
      End
      Begin PhotoDemon.pdColorSelector csSelection 
         Height          =   330
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   1980
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   582
      End
      Begin PhotoDemon.pdDropDown ddAppearance 
         Height          =   705
         Index           =   1
         Left            =   120
         TabIndex        =   58
         Top             =   1200
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1244
         Caption         =   "fill exterior"
         FontSizeCaption =   10
      End
   End
   Begin PhotoDemon.pdContainer cntrPopOut 
      Height          =   2160
      Index           =   2
      Left            =   3240
      Top             =   1320
      Visible         =   0   'False
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   3810
      Begin PhotoDemon.pdCheckBox chkAutoDrop 
         Height          =   375
         Index           =   0
         Left            =   210
         TabIndex        =   33
         Top             =   1680
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   661
         Caption         =   "open panel automatically"
         Value           =   0   'False
      End
      Begin PhotoDemon.pdButtonToolbox cmdFlyoutLock 
         Height          =   390
         Index           =   2
         Left            =   6120
         TabIndex        =   8
         Top             =   1650
         Width           =   390
         _ExtentX        =   1111
         _ExtentY        =   1111
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdDropDown cboSelArea 
         Height          =   735
         Index           =   0
         Left            =   3360
         TabIndex        =   15
         Top             =   810
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1296
         Caption         =   "area"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdSlider sltSelectionBorder 
         CausesValidation=   0   'False
         Height          =   405
         Index           =   0
         Left            =   3360
         TabIndex        =   16
         Top             =   1635
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   714
         Min             =   1
         Max             =   1000
         ScaleStyle      =   2
         Value           =   1
         DefaultValue    =   1
      End
      Begin PhotoDemon.pdLabel lblColon 
         Height          =   375
         Index           =   0
         Left            =   1320
         Top             =   360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Alignment       =   2
         Caption         =   ":"
         FontSize        =   12
      End
      Begin PhotoDemon.pdButtonToolbox cmdLock 
         Height          =   360
         Index           =   2
         Left            =   2895
         TabIndex        =   17
         Top             =   360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   635
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdSpinner tudSel 
         Height          =   345
         Index           =   2
         Left            =   240
         TabIndex        =   18
         Top             =   375
         Width           =   1080
         _ExtentX        =   2328
         _ExtentY        =   714
         DefaultValue    =   1
         Min             =   1
         Max             =   32000
         Value           =   1
         ShowResetButton =   0   'False
      End
      Begin PhotoDemon.pdSpinner tudSel 
         Height          =   345
         Index           =   3
         Left            =   1815
         TabIndex        =   19
         Top             =   375
         Width           =   1080
         _ExtentX        =   2328
         _ExtentY        =   714
         DefaultValue    =   1
         Min             =   1
         Max             =   32000
         Value           =   1
         ShowResetButton =   0   'False
      End
      Begin PhotoDemon.pdSlider sltCornerRounding 
         CausesValidation=   0   'False
         Height          =   735
         Left            =   3360
         TabIndex        =   32
         Top             =   0
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1296
         Caption         =   "round corners"
         FontSizeCaption =   10
         Max             =   100
         SigDigits       =   1
      End
      Begin PhotoDemon.pdSpinner tudSel 
         Height          =   345
         Index           =   4
         Left            =   240
         TabIndex        =   34
         Top             =   1215
         Width           =   1080
         _ExtentX        =   2328
         _ExtentY        =   714
         Min             =   -32000
         Max             =   32000
         ShowResetButton =   0   'False
      End
      Begin PhotoDemon.pdSpinner tudSel 
         Height          =   345
         Index           =   5
         Left            =   1815
         TabIndex        =   35
         Top             =   1215
         Width           =   1080
         _ExtentX        =   2328
         _ExtentY        =   714
         Min             =   -32000
         Max             =   32000
         ShowResetButton =   0   'False
      End
      Begin PhotoDemon.pdLabel lblNoOptions 
         Height          =   345
         Index           =   3
         Left            =   120
         Top             =   0
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   609
         Caption         =   "aspect ratio"
      End
      Begin PhotoDemon.pdLabel lblNoOptions 
         Height          =   345
         Index           =   4
         Left            =   120
         Top             =   840
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   609
         Caption         =   "position (x, y)"
      End
   End
   Begin PhotoDemon.pdTitle ttlPanel 
      Height          =   375
      Index           =   1
      Left            =   4200
      TabIndex        =   13
      Top             =   0
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      Caption         =   "smoothing"
      Value           =   0   'False
   End
   Begin PhotoDemon.pdContainer cntrPopOut 
      Height          =   1335
      Index           =   4
      Left            =   9960
      Top             =   1440
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   2355
      Begin PhotoDemon.pdButtonToolbox cmdFlyoutLock 
         Height          =   390
         Index           =   4
         Left            =   2880
         TabIndex        =   21
         Top             =   810
         Width           =   390
         _ExtentX        =   1111
         _ExtentY        =   1111
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdDropDown cboSelArea 
         Height          =   735
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   0
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1296
         Caption         =   "area"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdSlider sltSelectionBorder 
         CausesValidation=   0   'False
         Height          =   405
         Index           =   2
         Left            =   180
         TabIndex        =   23
         Top             =   780
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   714
         Min             =   1
         Max             =   1000
         ScaleStyle      =   2
         Value           =   1
         DefaultValue    =   1
      End
   End
   Begin PhotoDemon.pdContainer cntrPopOut 
      Height          =   495
      Index           =   5
      Left            =   10080
      Top             =   2880
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   873
      Begin PhotoDemon.pdButtonToolbox cmdFlyoutLock 
         Height          =   390
         Index           =   5
         Left            =   2880
         TabIndex        =   24
         Top             =   30
         Width           =   390
         _ExtentX        =   1111
         _ExtentY        =   1111
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdSlider sltSelectionBorder 
         CausesValidation=   0   'False
         Height          =   405
         Index           =   3
         Left            =   180
         TabIndex        =   25
         Top             =   0
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   714
         Min             =   1
         Max             =   1000
         ScaleStyle      =   2
         Value           =   1
         DefaultValue    =   1
      End
      Begin PhotoDemon.pdLabel lblNoOptions 
         Height          =   375
         Index           =   2
         Left            =   120
         Top             =   0
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         Alignment       =   2
         Caption         =   "(no additional options)"
      End
   End
   Begin PhotoDemon.pdContainer cntrPopOut 
      Height          =   3015
      Index           =   6
      Left            =   9960
      Top             =   3480
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   5318
      Begin PhotoDemon.pdButtonToolbox cmdFlyoutLock 
         Height          =   390
         Index           =   6
         Left            =   3120
         TabIndex        =   27
         Top             =   2430
         Width           =   390
         _ExtentX        =   1111
         _ExtentY        =   1111
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdDropDown cboWandCompare 
         Height          =   735
         Left            =   120
         TabIndex        =   28
         Top             =   30
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1296
         Caption         =   "compare pixels by"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdButtonStrip btsWandArea 
         Height          =   930
         Left            =   120
         TabIndex        =   29
         Top             =   870
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1640
         Caption         =   "mode"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdButtonStrip btsWandMerge 
         Height          =   930
         Left            =   120
         TabIndex        =   31
         Top             =   1890
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1640
         Caption         =   "sample from"
         FontSizeCaption =   10
      End
   End
   Begin PhotoDemon.pdContainer cntrPopOut 
      Height          =   1800
      Index           =   3
      Left            =   3240
      Top             =   4080
      Visible         =   0   'False
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   3175
      Begin PhotoDemon.pdCheckBox chkAutoDrop 
         Height          =   375
         Index           =   1
         Left            =   3435
         TabIndex        =   46
         Top             =   1260
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   661
         Caption         =   "open panel automatically"
         Value           =   0   'False
      End
      Begin PhotoDemon.pdButtonToolbox cmdFlyoutLock 
         Height          =   390
         Index           =   3
         Left            =   6120
         TabIndex        =   47
         Top             =   1230
         Width           =   390
         _ExtentX        =   1111
         _ExtentY        =   1111
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdDropDown cboSelArea 
         Height          =   735
         Index           =   1
         Left            =   3360
         TabIndex        =   48
         Top             =   0
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1296
         Caption         =   "area"
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdSlider sltSelectionBorder 
         CausesValidation=   0   'False
         Height          =   405
         Index           =   1
         Left            =   3480
         TabIndex        =   49
         Top             =   780
         Width           =   2610
         _ExtentX        =   4604
         _ExtentY        =   714
         Min             =   1
         Max             =   1000
         ScaleStyle      =   2
         Value           =   1
         DefaultValue    =   1
      End
      Begin PhotoDemon.pdLabel lblColon 
         Height          =   375
         Index           =   1
         Left            =   1320
         Top             =   375
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Alignment       =   2
         Caption         =   ":"
         FontSize        =   12
      End
      Begin PhotoDemon.pdButtonToolbox cmdLock 
         Height          =   360
         Index           =   5
         Left            =   2895
         TabIndex        =   50
         Top             =   375
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   635
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdSpinner tudSel 
         Height          =   345
         Index           =   8
         Left            =   240
         TabIndex        =   51
         Top             =   375
         Width           =   1080
         _ExtentX        =   2328
         _ExtentY        =   714
         DefaultValue    =   1
         Min             =   1
         Max             =   32000
         Value           =   1
         ShowResetButton =   0   'False
      End
      Begin PhotoDemon.pdSpinner tudSel 
         Height          =   345
         Index           =   9
         Left            =   1815
         TabIndex        =   52
         Top             =   375
         Width           =   1080
         _ExtentX        =   2328
         _ExtentY        =   714
         DefaultValue    =   1
         Min             =   1
         Max             =   32000
         Value           =   1
         ShowResetButton =   0   'False
      End
      Begin PhotoDemon.pdSpinner tudSel 
         Height          =   345
         Index           =   10
         Left            =   240
         TabIndex        =   53
         Top             =   1215
         Width           =   1080
         _ExtentX        =   2328
         _ExtentY        =   714
         Min             =   -32000
         Max             =   32000
         ShowResetButton =   0   'False
      End
      Begin PhotoDemon.pdSpinner tudSel 
         Height          =   345
         Index           =   11
         Left            =   1815
         TabIndex        =   54
         Top             =   1215
         Width           =   1080
         _ExtentX        =   2328
         _ExtentY        =   714
         Min             =   -32000
         Max             =   32000
         ShowResetButton =   0   'False
      End
      Begin PhotoDemon.pdLabel lblNoOptions 
         Height          =   375
         Index           =   5
         Left            =   120
         Top             =   0
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         Caption         =   "aspect ratio"
      End
      Begin PhotoDemon.pdLabel lblNoOptions 
         Height          =   375
         Index           =   6
         Left            =   120
         Top             =   840
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         Caption         =   "position (x, y)"
      End
   End
   Begin PhotoDemon.pdContainer ctlGroupSelectionSubcontainer 
      Height          =   855
      Index           =   0
      Left            =   6780
      Top             =   0
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1508
      Begin PhotoDemon.pdTitle ttlPanel 
         Height          =   375
         Index           =   2
         Left            =   0
         TabIndex        =   36
         Top             =   0
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   661
         Caption         =   "size (w, h)"
         Value           =   0   'False
      End
      Begin PhotoDemon.pdButtonToolbox cmdLock 
         Height          =   360
         Index           =   1
         Left            =   2775
         TabIndex        =   37
         Top             =   390
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   635
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdSpinner tudSel 
         Height          =   345
         Index           =   0
         Left            =   120
         TabIndex        =   38
         Top             =   405
         Width           =   1080
         _ExtentX        =   2328
         _ExtentY        =   714
         DefaultValue    =   1
         Min             =   1
         Max             =   32000
         Value           =   1
         ShowResetButton =   0   'False
      End
      Begin PhotoDemon.pdSpinner tudSel 
         Height          =   345
         Index           =   1
         Left            =   1695
         TabIndex        =   39
         Top             =   405
         Width           =   1080
         _ExtentX        =   2328
         _ExtentY        =   714
         DefaultValue    =   1
         Min             =   1
         Max             =   32000
         Value           =   1
         ShowResetButton =   0   'False
      End
      Begin PhotoDemon.pdButtonToolbox cmdLock 
         Height          =   360
         Index           =   0
         Left            =   1200
         TabIndex        =   40
         Top             =   390
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   635
         StickyToggle    =   -1  'True
      End
   End
   Begin PhotoDemon.pdContainer ctlGroupSelectionSubcontainer 
      Height          =   855
      Index           =   2
      Left            =   6780
      Top             =   0
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1508
      Begin PhotoDemon.pdSlider sltPolygonCurvature 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   375
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         FontSizeCaption =   10
         Max             =   1
         SigDigits       =   2
      End
      Begin PhotoDemon.pdTitle ttlPanel 
         Height          =   375
         Index           =   4
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   661
         Caption         =   "curvature"
         Value           =   0   'False
      End
   End
   Begin PhotoDemon.pdContainer ctlGroupSelectionSubcontainer 
      Height          =   855
      Index           =   4
      Left            =   6780
      Top             =   0
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1508
      Begin PhotoDemon.pdSlider sltWandTolerance 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   120
         TabIndex        =   1
         Top             =   375
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   741
         FontSizeCaption =   10
         Max             =   100
         SigDigits       =   1
         ScaleStyle      =   1
         Value           =   15
         DefaultValue    =   15
      End
      Begin PhotoDemon.pdTitle ttlPanel 
         Height          =   375
         Index           =   6
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   661
         Caption         =   "tolerance"
         Value           =   0   'False
      End
   End
   Begin PhotoDemon.pdContainer ctlGroupSelectionSubcontainer 
      Height          =   855
      Index           =   3
      Left            =   6780
      Top             =   0
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1508
      Begin PhotoDemon.pdDropDown cboSelArea 
         Height          =   360
         Index           =   3
         Left            =   120
         TabIndex        =   2
         Top             =   375
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   635
         FontSizeCaption =   10
      End
      Begin PhotoDemon.pdSlider sltSmoothStroke 
         CausesValidation=   0   'False
         Height          =   735
         Left            =   2760
         TabIndex        =   3
         Top             =   30
         Visible         =   0   'False
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   1296
         Caption         =   "stroke smoothing"
         FontSizeCaption =   10
         Max             =   1
         SigDigits       =   2
      End
      Begin PhotoDemon.pdTitle ttlPanel 
         Height          =   375
         Index           =   5
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   661
         Caption         =   "area"
         Value           =   0   'False
      End
   End
   Begin PhotoDemon.pdContainer ctlGroupSelectionSubcontainer 
      Height          =   855
      Index           =   1
      Left            =   6780
      Top             =   0
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1508
      Begin PhotoDemon.pdTitle ttlPanel 
         Height          =   375
         Index           =   3
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   661
         Caption         =   "size (w, h)"
         Value           =   0   'False
      End
      Begin PhotoDemon.pdButtonToolbox cmdLock 
         Height          =   360
         Index           =   4
         Left            =   2775
         TabIndex        =   42
         Top             =   390
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   635
         StickyToggle    =   -1  'True
      End
      Begin PhotoDemon.pdSpinner tudSel 
         Height          =   345
         Index           =   6
         Left            =   120
         TabIndex        =   43
         Top             =   405
         Width           =   1080
         _ExtentX        =   2328
         _ExtentY        =   714
         Min             =   -32000
         Max             =   32000
         ShowResetButton =   0   'False
      End
      Begin PhotoDemon.pdSpinner tudSel 
         Height          =   345
         Index           =   7
         Left            =   1695
         TabIndex        =   44
         Top             =   405
         Width           =   1080
         _ExtentX        =   2328
         _ExtentY        =   714
         Min             =   -32000
         Max             =   32000
         ShowResetButton =   0   'False
      End
      Begin PhotoDemon.pdButtonToolbox cmdLock 
         Height          =   360
         Index           =   3
         Left            =   1200
         TabIndex        =   45
         Top             =   390
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   635
         StickyToggle    =   -1  'True
      End
   End
   Begin PhotoDemon.pdLabel lblNoOptions 
      Height          =   345
      Index           =   7
      Left            =   2040
      Top             =   30
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   609
      Caption         =   "combine"
   End
End
Attribute VB_Name = "toolpanel_Selections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Selection Tool Panel
'Copyright 2013-2026 by Tanner Helland
'Created: 02/Oct/13
'Last updated: 17/February/22
'Last update: new UI elements for novel features in new selection rendering engine
'
'This form includes all user-editable settings for PD's various selection tools.
' Yes, all selection tools share a single options panel.  (This decision was made
' many years ago, when these tools shared most of their settings.  In the years since,
' they have grown more distinct, but this single shared options panel remains.)
'
'The code involved in this form has become much more complex since upgrading the UI
' in v9.0 to a new "flyout panel" layout.  This greatly improved information density
' on the form, but it also meant that things like tab order need to be handled
' manually at run-time (because control visibility is no longer fixed, but may change
' constantly within a single panel).
'
'As such, the bulk of this dialog's code now deals specifically with UI handling.
' Flyouts in particular need to behave nicely, and tab order must always work
' intuitively regardless of flyout panel state.  I think the current result is
' highly useable, but a lot of code is necessary to make that work!
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Flyout manager
Private WithEvents m_Flyout As pdFlyout
Attribute m_Flyout.VB_VarHelpID = -1

'The value of all controls on this form are saved and loaded to file by this class
Private WithEvents m_lastUsedSettings As pdLastUsedSettings
Attribute m_lastUsedSettings.VB_VarHelpID = -1

Private Sub btsCombine_Click(ByVal buttonIndex As Long)
    
    'Changing combine mode does *not* change the active selection.  This property is only read
    ' when a *new* selection is created.
    
    'As such, we don't need to set any properties here.
    
End Sub

Private Sub btsCombine_MouseMoveInfoOnly(ByVal buttonIndex As Long)
    
    Dim ttString As String
    
    Select Case buttonIndex
        Case pdsm_Replace
            ttString = g_Language.TranslateMessage("New selection")
        Case pdsm_Add
            ttString = g_Language.TranslateMessage("Add to selection")
            ttString = ttString & vbCrLf & g_Language.TranslateMessage("Shortcut key: %1", Hotkeys.GetGenericMenuText(cmt_Shift))
        Case pdsm_Subtract
            ttString = g_Language.TranslateMessage("Subtract from selection")
            ttString = ttString & vbCrLf & g_Language.TranslateMessage("Shortcut key: %1", Hotkeys.GetGenericMenuText(cmt_Alt))
        Case pdsm_Intersect
            ttString = g_Language.TranslateMessage("Intersect with selection")
            ttString = ttString & vbCrLf & g_Language.TranslateMessage("Shortcut key: %1", Hotkeys.GetGenericMenuText(cmt_Shift) & "+" & Hotkeys.GetGenericMenuText(cmt_Alt))
    End Select
    
    btsCombine.AssignTooltip ttString
    
End Sub

Private Sub btsCombine_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then newTargetHwnd = Me.cmdFlyoutLock(0).hWnd Else newTargetHwnd = Me.ttlPanel(1).hWnd
End Sub

Private Sub btsWandArea_Click(ByVal buttonIndex As Long)
    
    'If a selection is already active, change its type to match the current option, then redraw it
    If SelectionsAllowed(False) And (g_CurrentTool = SelectionUI.GetRelevantToolFromSelectShape()) Then
        PDImages.GetActiveImage.MainSelection.SetSelectionProperty sp_WandSearchMode, buttonIndex
        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    End If
    
End Sub

Private Sub btsWandArea_GotFocusAPI()
    UpdateFlyout 6, True
End Sub

Private Sub btsWandArea_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.cboWandCompare.hWnd
    Else
        newTargetHwnd = Me.btsWandMerge.hWnd
    End If
End Sub

Private Sub btsWandMerge_Click(ByVal buttonIndex As Long)

    'If a selection is already active, change its type to match the current option, then redraw it
    If SelectionsAllowed(False) And (g_CurrentTool = SelectionUI.GetRelevantToolFromSelectShape()) Then
        PDImages.GetActiveImage.MainSelection.SetSelectionProperty sp_WandSampleMerged, buttonIndex
        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    End If

End Sub

Private Sub btsWandMerge_GotFocusAPI()
    UpdateFlyout 6, True
End Sub

Private Sub btsWandMerge_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.btsWandArea.hWnd
    Else
        newTargetHwnd = Me.cmdFlyoutLock(6).hWnd
    End If
End Sub

Private Sub cboSelArea_Click(Index As Integer)

    sltSelectionBorder(Index).Visible = (cboSelArea(Index).ListIndex = sa_Border)
    
    'The flyout design of index 3 area box necessitates a special message if a border selection
    ' is *not* visible.
    If (Index = 3) Then
        Me.lblNoOptions(2).Visible = (Not sltSelectionBorder(Index).Visible)
    End If
    
    'If a selection is already active, change its type to match the current selection, then redraw it
    If SelectionsAllowed(False) And (g_CurrentTool = SelectionUI.GetRelevantToolFromSelectShape()) Then
        PDImages.GetActiveImage.MainSelection.SetSelectionProperty sp_Area, cboSelArea(Index).ListIndex
        PDImages.GetActiveImage.MainSelection.SetSelectionProperty sp_BorderWidth, sltSelectionBorder(Index).Value
        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    End If
    
End Sub

Private Sub cboSelArea_GotFocusAPI(Index As Integer)
    UpdateFlyout Index + 2, True
End Sub

Private Sub cboSelArea_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    Select Case Index
        Case 0
            If shiftTabWasPressed Then
                newTargetHwnd = Me.sltCornerRounding.hWndSpinner
            Else
                If Me.sltSelectionBorder(0).Visible Then
                    newTargetHwnd = Me.sltSelectionBorder(0).hWnd
                Else
                    newTargetHwnd = Me.cmdFlyoutLock(2).hWnd
                End If
            End If
        Case 1
            If shiftTabWasPressed Then
                If Me.tudSel(11).Enabled Then
                    newTargetHwnd = Me.tudSel(11).hWnd
                Else
                    newTargetHwnd = Me.cmdLock(5).hWnd
                End If
            Else
                If Me.sltSelectionBorder(1).Visible Then
                    newTargetHwnd = Me.sltSelectionBorder(1).hWnd
                Else
                    newTargetHwnd = Me.chkAutoDrop(1).hWnd
                End If
            End If
        Case 2
            If shiftTabWasPressed Then
                newTargetHwnd = Me.sltPolygonCurvature.hWndSpinner
            Else
                If Me.sltSelectionBorder(Index).Visible Then newTargetHwnd = Me.sltSelectionBorder(Index).hWnd Else newTargetHwnd = Me.cmdFlyoutLock(Index + 2).hWnd
            End If
        Case 3
            If shiftTabWasPressed Then
                newTargetHwnd = Me.ttlPanel(5).hWnd
            Else
                If Me.sltSelectionBorder(Index).Visible Then newTargetHwnd = Me.sltSelectionBorder(Index).hWnd Else newTargetHwnd = Me.cmdFlyoutLock(Index + 2).hWnd
            End If
    End Select
End Sub

'Selection smoothing is handled universally, even if the current selection shape does not match the active
' selection tool.  (This is done because antialiasing/feathering are universally supported across all types.)
Private Sub cboSelSmoothing_Click()

    UpdateSelectionPanelLayout
    
    'If a selection is already active, change its type to match the current selection, then redraw it
    If SelectionsAllowed(False) Then
        PDImages.GetActiveImage.MainSelection.SetSelectionProperty sp_Smoothing, cboSelSmoothing.ListIndex
        PDImages.GetActiveImage.MainSelection.SetSelectionProperty sp_FeatheringRadius, sltSelectionFeathering.Value
        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    End If

End Sub

Private Sub cboSelSmoothing_GotFocusAPI()
    UpdateFlyout 1, True
End Sub

Private Sub cboSelSmoothing_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.ttlPanel(1).hWnd
    Else
        
        'Feathered selections display a radius slider
        If (Me.cboSelSmoothing.ListIndex = 2) Then
            newTargetHwnd = Me.sltSelectionFeathering.hWndSlider
        Else
            newTargetHwnd = Me.cmdFlyoutLock(1).hWnd
        End If
        
    End If
End Sub

Private Sub cboWandCompare_Click()
    
    'Limit the accuracy of the tolerance for certain comparison methods.
    If (cboWandCompare.ListIndex > 1) Then sltWandTolerance.SigDigits = 0 Else sltWandTolerance.SigDigits = 1
    
    'If a selection is already active, change its type to match the current option, then redraw it
    If SelectionsAllowed(False) And (g_CurrentTool = SelectionUI.GetRelevantToolFromSelectShape()) Then
        PDImages.GetActiveImage.MainSelection.SetSelectionProperty sp_WandCompareMethod, cboWandCompare.ListIndex
        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    End If
    
End Sub

Private Sub cboWandCompare_GotFocusAPI()
    UpdateFlyout 6, True
End Sub

Private Sub cboWandCompare_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.sltWandTolerance.hWndSpinner
    Else
        newTargetHwnd = Me.btsWandArea.hWnd
    End If
End Sub

Private Sub chkAppearance_Click(Index As Integer)
    
    SelectionUI.NotifySelectionRenderChange pdsr_Animate, chkAppearance(0).Value
    
    'Note: appearance changes require a viewport redraw to reflect the new setting(s)
    If SelectionsAllowed(False) Then Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)

End Sub

Private Sub chkAppearance_GotFocusAPI(Index As Integer)
    UpdateFlyout 0, True
End Sub

Private Sub chkAppearance_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If (Index = 0) Then
        If shiftTabWasPressed Then
            newTargetHwnd = Me.ttlPanel(0).hWnd
        Else
            newTargetHwnd = Me.ddAppearance(0).hWnd
        End If
    End If
End Sub

Private Sub chkAutoDrop_Click(Index As Integer)
    
    'Clicking this checkbox prevents the "auto-open" behavior of the selection dropdown.
    ' (This may prove to be a temporary fix since this panel is being redesigned in light of
    ' new layout opportunities thanks to flyout panels, but I can easily remove this checkbox
    ' if it proves superfluous in a new design.)
    '
    'Anyway, because a user is likely to click this after the panel has already auto-dropped,
    ' as a convenience we'll also deactivate the pin button if it's active.  (This relationship
    ' may not be intuitive to beginners, and I don't want to frustrate them by having them
    ' unclick this button, then the flyout not disappearing.)
    If (Not Me.chkAutoDrop(Index).Value) Then
        Me.cmdFlyoutLock(Index + 2).Value = False
        If (Not m_Flyout Is Nothing) Then
            m_Flyout.UpdateLockStatus Me.cntrPopOut(Index + 2).hWnd, False, cmdFlyoutLock(Index + 2)
        End If
    End If
    
End Sub

Private Sub chkAutoDrop_GotFocusAPI(Index As Integer)
    UpdateFlyout Index + 2, True
End Sub

Private Sub chkAutoDrop_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If (Index = 0) Then
        If shiftTabWasPressed Then
            If Me.tudSel(5).Enabled Then
                newTargetHwnd = Me.tudSel(5).hWnd
            Else
                newTargetHwnd = Me.cmdLock(2).hWnd
            End If
        Else
            newTargetHwnd = Me.sltCornerRounding.hWndSlider
        End If
    Else
        If shiftTabWasPressed Then
            If Me.sltSelectionBorder(1).Visible Then
                newTargetHwnd = Me.sltSelectionBorder(1).hWnd
            Else
                newTargetHwnd = Me.cboSelArea(1).hWnd
            End If
        Else
            newTargetHwnd = Me.cmdFlyoutLock(3).hWnd
        End If
    End If
End Sub

Private Sub cmdFlyoutLock_Click(Index As Integer, ByVal Shift As ShiftConstants)
    If (Not m_Flyout Is Nothing) Then
        m_Flyout.UpdateLockStatus Me.cntrPopOut(Index).hWnd, cmdFlyoutLock(Index).Value, cmdFlyoutLock(Index)
    End If
End Sub

Private Sub cmdFlyoutLock_GotFocusAPI(Index As Integer)
    UpdateFlyout Index, True
End Sub

Private Sub cmdFlyoutLock_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    
    Select Case Index
        
        Case 0
        
            If shiftTabWasPressed Then
                newTargetHwnd = Me.spnOpacity(1).hWnd
            Else
                newTargetHwnd = Me.btsCombine.hWnd
            End If
            
        Case 1
            If shiftTabWasPressed Then
                If Me.sltSelectionFeathering.Visible Then
                    newTargetHwnd = Me.sltSelectionFeathering.hWndSpinner
                Else
                    newTargetHwnd = Me.cboSelSmoothing.hWnd
                End If
            Else
                
                'Next control varies by selection type
                If (g_CurrentTool = SELECT_RECT) Then
                    newTargetHwnd = Me.ttlPanel(2).hWnd
                ElseIf (g_CurrentTool = SELECT_CIRC) Then
                    newTargetHwnd = Me.ttlPanel(3).hWnd
                ElseIf (g_CurrentTool = SELECT_POLYGON) Then
                    newTargetHwnd = Me.ttlPanel(4).hWnd
                ElseIf (g_CurrentTool = SELECT_LASSO) Then
                    newTargetHwnd = Me.ttlPanel(5).hWnd
                ElseIf (g_CurrentTool = SELECT_WAND) Then
                    newTargetHwnd = Me.ttlPanel(6).hWnd
                End If
                
            End If
            
        Case 2
            If shiftTabWasPressed Then
                If Me.sltSelectionBorder(0).Visible Then
                    newTargetHwnd = Me.sltSelectionBorder(0).hWndSpinner
                Else
                    newTargetHwnd = Me.cboSelArea(0).hWnd
                End If
            Else
                newTargetHwnd = Me.ttlPanel(0).hWnd
            End If
            
        Case 3
            If shiftTabWasPressed Then
                newTargetHwnd = Me.chkAutoDrop(1).hWnd
            Else
                newTargetHwnd = Me.ttlPanel(0).hWnd
            End If
            
        Case 4
            If shiftTabWasPressed Then
                If Me.sltSelectionBorder(2).Visible Then
                    newTargetHwnd = Me.sltSelectionBorder(2).hWndSpinner
                Else
                    newTargetHwnd = Me.cboSelArea(2).hWnd
                End If
            Else
                newTargetHwnd = Me.ttlPanel(0).hWnd
            End If
        
        Case 5
            If shiftTabWasPressed Then
                If Me.sltSelectionBorder(3).Visible Then
                    newTargetHwnd = Me.sltSelectionBorder(3).hWndSpinner
                Else
                    newTargetHwnd = Me.cboSelArea(3).hWnd
                End If
            Else
                newTargetHwnd = Me.ttlPanel(0).hWnd
            End If
            
        Case 6
            If shiftTabWasPressed Then
                newTargetHwnd = Me.btsWandMerge.hWnd
            Else
                newTargetHwnd = Me.ttlPanel(0).hWnd
            End If
        
    End Select
    
End Sub

Private Sub cmdLock_Click(Index As Integer, ByVal Shift As ShiftConstants)
    
    'Ignore lock actions unless a selection is active, *and* the current selection tool matches the currently
    ' active selection.
    If SelectionsAllowed(False) And (g_CurrentTool = SelectionUI.GetRelevantToolFromSelectShape()) Then
        
        Dim lockedValue As Variant
        
        'Because of the way the cmdLock buttons are structured (with *two* instances per button, one for the
        ' rectangular selection tool and another for the elliptical selection tool), we have to perform some
        ' manual remapping of indices based on the active tool and the active selection attribute
        ' (position/size/aspect ratio).
        Dim relevantIndex As PD_SelectionLockable
        If (g_CurrentTool = SELECT_RECT) Then
            relevantIndex = Index
        ElseIf (g_CurrentTool = SELECT_CIRC) Then
            relevantIndex = Index - 3
        End If
        
        'TODO: verify working for ellipse selections
        If (relevantIndex = pdsl_Width) Or (relevantIndex = pdsl_Height) Then
            lockedValue = tudSel(Index).Value
        Else
            If (tudSel(1).Value <> 0) Then lockedValue = tudSel(0).Value / tudSel(1).Value
        End If
        
        'In the case of aspect ratio vs width/height locks, we don't see both controls at the same time so we
        ' don't have to manually synchronize any UI elements.  Width and height changes are different, however,
        ' because locking one necessarily unlocks the other.
        If cmdLock(Index).Value Then
            If (relevantIndex = pdsl_Width) Then
                cmdLock(Index + 1).Value = False
                cmdLock(Index + 2).Value = False
            ElseIf (relevantIndex = pdsl_Height) Then
                cmdLock(Index - 1).Value = False
                cmdLock(Index + 1).Value = False
            ElseIf (relevantIndex = pdsl_AspectRatio) Then
                cmdLock(Index - 1).Value = False
                cmdLock(Index - 2).Value = False
            End If
        End If
        
        If cmdLock(Index).Value Then
            PDImages.GetActiveImage.MainSelection.LockProperty relevantIndex, lockedValue
        Else
            PDImages.GetActiveImage.MainSelection.UnlockProperty relevantIndex
        End If
        
    End If
    
End Sub

Private Sub cmdLock_GotFocusAPI(Index As Integer)
    If (Index <= 2) Then
        UpdateFlyout 2, True
    Else
        UpdateFlyout 3, True
    End If
End Sub

Private Sub cmdLock_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    
    'Similar mappings can be used for both the rectangular and elliptical lock buttons
    Dim baseIndex As Long
    If (Index >= 3) Then baseIndex = 6 Else baseIndex = 0
    
    Select Case Index
        
        Case 0, 3
            If shiftTabWasPressed Then
                If Me.tudSel(baseIndex).Enabled Then
                    newTargetHwnd = Me.tudSel(baseIndex).hWnd
                Else
                    If (g_CurrentTool = SELECT_RECT) Then
                        newTargetHwnd = Me.ttlPanel(2).hWnd
                    Else
                        newTargetHwnd = Me.ttlPanel(3).hWnd
                    End If
                End If
            Else
                If Me.tudSel(baseIndex + 1).Enabled Then
                    newTargetHwnd = Me.tudSel(baseIndex + 1).hWnd
                Else
                    newTargetHwnd = Me.cmdLock(Index + 1).hWnd
                End If
            End If
            
        Case 1, 4
            If shiftTabWasPressed Then
                If Me.tudSel(baseIndex + 1).Enabled Then
                    newTargetHwnd = Me.tudSel(baseIndex + 1).hWnd
                Else
                    newTargetHwnd = Me.cmdLock(Index - 1).hWnd
                End If
            Else
                If Me.tudSel(baseIndex + 2).Enabled Then
                    newTargetHwnd = Me.tudSel(baseIndex + 2).hWnd
                Else
                    newTargetHwnd = Me.cmdLock(Index + 1).hWnd
                End If
            End If
        
        Case 2, 5
            If shiftTabWasPressed Then
                If Me.tudSel(baseIndex + 3).Enabled Then
                    newTargetHwnd = Me.tudSel(baseIndex + 3).hWnd
                Else
                    newTargetHwnd = Me.cmdLock(Index - 1).hWnd
                End If
            Else
                If Me.tudSel(baseIndex + 4).Enabled Then
                    newTargetHwnd = Me.tudSel(baseIndex + 4).hWnd
                Else
                    If (g_CurrentTool = SELECT_RECT) Then
                        newTargetHwnd = Me.chkAutoDrop(0).hWnd
                    Else
                        newTargetHwnd = Me.cboSelArea(1).hWnd
                    End If
                End If
            End If
            
    End Select
    
End Sub

Private Sub csSelection_ColorChanged(Index As Integer)
    
    If (Index = 0) Then
        SelectionUI.NotifySelectionRenderChange pdsr_InteriorFillColor, csSelection(Index).Color
    ElseIf (Index = 1) Then
        SelectionUI.NotifySelectionRenderChange pdsr_ExteriorFillColor, csSelection(Index).Color
    End If
    
    If SelectionsAllowed(False) Then Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub csSelection_GotFocusAPI(Index As Integer)
    UpdateFlyout 0, True
End Sub

Private Sub csSelection_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    
    If shiftTabWasPressed Then
        newTargetHwnd = Me.ddAppearance(Index).hWnd
    Else
        newTargetHwnd = Me.spnOpacity(Index).hWnd
    End If
    
End Sub

Private Sub ddAppearance_Click(Index As Integer)
    
    If (Index = 0) Then
        SelectionUI.NotifySelectionRenderChange pdsr_InteriorFillMode, ddAppearance(Index).ListIndex
    ElseIf (Index = 1) Then
        SelectionUI.NotifySelectionRenderChange pdsr_ExteriorFillMode, ddAppearance(Index).ListIndex
    End If
    
    If SelectionsAllowed(False) Then Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub ddAppearance_GotFocusAPI(Index As Integer)
    UpdateFlyout 0, True
End Sub

Private Sub ddAppearance_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        If (Index = 0) Then
            newTargetHwnd = Me.chkAppearance(0).hWnd
        Else
            newTargetHwnd = Me.spnOpacity(0).hWnd
        End If
    Else
        newTargetHwnd = Me.csSelection(Index).hWnd
    End If
End Sub

Private Sub Form_Load()
    
    'Suspend any visual updates while the form is being loaded
    Viewport.DisableRendering
    
    Dim suspendActive As Boolean
    If PDImages.IsImageActive() Then
        suspendActive = True
        PDImages.GetActiveImage.MainSelection.SuspendAutoRefresh True
    End If
    
    'Initialize various selection tool settings
    
    'Selection visual styles (Highlight, Lightbox, or Outline)
    Dim i As Long
    For i = 0 To 1
        ddAppearance(i).SetAutomaticRedraws False
        ddAppearance(i).Clear
        ddAppearance(i).AddItem "always", 0
        ddAppearance(i).AddItem "when combining", 1
        ddAppearance(i).AddItem "never", 2
        ddAppearance(i).SetAutomaticRedraws True
    Next i
    
    ddAppearance(0).ListIndex = 1
    ddAppearance(1).ListIndex = 2
    
    csSelection(0).Color = RGB(110, 230, 255)
    csSelection(0).Visible = True
    spnOpacity(0).Value = 50
    spnOpacity(0).Visible = True
    
    csSelection(1).Color = RGB(255, 60, 80)
    csSelection(1).Visible = True
    spnOpacity(1).Value = 50
    spnOpacity(1).Visible = True
    
    'Selection combine modes.  (These use icons so we do not need to specify captions.)
    For i = 0 To 3
        btsCombine.AddItem vbNullString
    Next i
    btsCombine.ListIndex = 0
    
    'Selection smoothing (currently none, antialiased, fully feathered)
    cboSelSmoothing.SetAutomaticRedraws False
    cboSelSmoothing.Clear
    cboSelSmoothing.AddItem "none", 0
    cboSelSmoothing.AddItem "antialiased", 1
    cboSelSmoothing.AddItem "feathered", 2
    cboSelSmoothing.SetAutomaticRedraws True
    cboSelSmoothing.ListIndex = 1
    
    'Selection types (currently interior, exterior, border)
    For i = 0 To cboSelArea.Count - 1
        cboSelArea(i).SetAutomaticRedraws False
        cboSelArea(i).AddItem "interior", 0
        cboSelArea(i).AddItem "exterior", 1
        cboSelArea(i).AddItem "border", 2
        cboSelArea(i).ListIndex = 0
        cboSelArea(i).SetAutomaticRedraws True
    Next i
    
    'Magic wand options
    btsWandMerge.AddItem "image", 0
    btsWandMerge.AddItem "layer", 1
    btsWandMerge.ListIndex = 0
    
    btsWandArea.AddItem "contiguous", 0
    btsWandArea.AddItem "global", 1
    btsWandArea.ListIndex = 0
    
    Interface.PopulateFloodFillTypes cboWandCompare
    
    'Load any last-used settings for this form
    Set m_lastUsedSettings = New pdLastUsedSettings
    m_lastUsedSettings.SetParentForm Me
    m_lastUsedSettings.LoadAllControlValues
    
    If suspendActive Then PDImages.GetActiveImage.MainSelection.SuspendAutoRefresh False
    
    'If a selection is already active, synchronize all UI elements to match
    If suspendActive Then
        If PDImages.GetActiveImage.IsSelectionActive Then SelectionUI.SyncTextToCurrentSelection PDImages.GetActiveImageID()
    End If
    
    Viewport.EnableRendering
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'Save all last-used settings to file
    If (Not m_lastUsedSettings Is Nothing) Then
        m_lastUsedSettings.SaveAllControlValues
        m_lastUsedSettings.SetParentForm Nothing
    End If
    
    'Failsafe only
    If (Not m_Flyout Is Nothing) Then m_Flyout.HideFlyout
    Set m_Flyout = Nothing
    
End Sub

'Whenever an active flyout panel is closed, we need to reset the matching titlebar to "closed" state
Private Sub m_Flyout_FlyoutClosed(origTriggerObject As Control)
    If (Not origTriggerObject Is Nothing) Then origTriggerObject.Value = False
End Sub

Private Sub m_LastUsedSettings_ReadCustomPresetData()
    
    'Reset the selection coordinate boxes to 0
    Dim i As Long
    For i = 0 To tudSel.Count - 1
        tudSel(i).Value = 0
    Next i
    
    'Selection properties always default to *unlocked*
    cmdLock(0).Value = False
    
    'Pull certain universal selection settings from PD's main preferences file
    If UserPrefs.IsReady Then
        'cboSelRender.ListIndex = SelectionUI.GetSelectionRenderMode()
        'csSelection(0).Color = SelectionUI.GetSelectionColor_Highlight()
        'spnOpacity(0).Value = SelectionUI.GetSelectionOpacity_Highlight()
        'csSelection(1).Color = SelectionUI.GetSelectionColor_Lightbox()
        'spnOpacity(1).Value = SelectionUI.GetSelectionOpacity_Lightbox()
    End If
    
End Sub

Private Sub sltCornerRounding_Change()
    If SelectionsAllowed(True) And (g_CurrentTool = SelectionUI.GetRelevantToolFromSelectShape()) Then
        PDImages.GetActiveImage.MainSelection.SetSelectionProperty sp_RoundedCornerRadius, sltCornerRounding.Value
        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    End If
End Sub

Private Sub sltCornerRounding_GotFocusAPI()
    UpdateFlyout 2, True
End Sub

Private Sub sltCornerRounding_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.chkAutoDrop(0).hWnd
    Else
        newTargetHwnd = Me.cboSelArea(0).hWnd
    End If
End Sub

Private Sub sltPolygonCurvature_Change()
    If SelectionsAllowed(True) And (g_CurrentTool = SelectionUI.GetRelevantToolFromSelectShape()) Then
        PDImages.GetActiveImage.MainSelection.SetSelectionProperty sp_PolygonCurvature, sltPolygonCurvature.Value
        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    End If
End Sub

Private Sub sltPolygonCurvature_GotFocusAPI()
    UpdateFlyout 4, True
End Sub

Private Sub sltPolygonCurvature_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.ttlPanel(4).hWnd
    Else
        newTargetHwnd = Me.cboSelArea(2).hWnd
    End If
End Sub

Private Sub sltSelectionBorder_Change(Index As Integer)
    If SelectionsAllowed(False) And (g_CurrentTool = SelectionUI.GetRelevantToolFromSelectShape()) Then
        PDImages.GetActiveImage.MainSelection.SetSelectionProperty sp_BorderWidth, sltSelectionBorder(Index).Value
        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    End If
End Sub

Private Sub sltSelectionBorder_GotFocusAPI(Index As Integer)
    UpdateFlyout Index + 2, True
End Sub

Private Sub sltSelectionBorder_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.cboSelArea(Index).hWnd
    Else
        If (Index = 1) Then
            newTargetHwnd = Me.chkAutoDrop(1).hWnd
        Else
            newTargetHwnd = Me.cmdFlyoutLock(Index + 2).hWnd
        End If
    End If
End Sub

Private Sub sltSelectionFeathering_Change()
    If SelectionsAllowed(False) Then
        PDImages.GetActiveImage.MainSelection.SetSelectionProperty sp_FeatheringRadius, sltSelectionFeathering.Value
        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    End If
End Sub

Private Sub sltSelectionFeathering_GotFocusAPI()
    UpdateFlyout 1, True
End Sub

Private Sub sltSelectionFeathering_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then newTargetHwnd = Me.cboSelSmoothing.hWnd Else newTargetHwnd = Me.cmdFlyoutLock(1).hWnd
End Sub

'When certain selection settings are enabled or disabled, corresponding controls are shown or hidden.  To keep the
' panel concise and clean, we move other controls up or down depending on what controls are visible.
Public Sub UpdateSelectionPanelLayout()

    'Display the feathering slider as necessary
    sltSelectionFeathering.Visible = (cboSelSmoothing.ListIndex = es_FullyFeathered)
    lblNoOptions(1).Visible = (cboSelSmoothing.ListIndex <> es_FullyFeathered)
    
    'Display the border slider as necessary
    If (SelectionUI.GetSelectionSubPanelFromCurrentTool < cboSelArea.Count - 1) And (SelectionUI.GetSelectionSubPanelFromCurrentTool > 0) Then
        sltSelectionBorder(SelectionUI.GetSelectionSubPanelFromCurrentTool).Visible = (cboSelArea(SelectionUI.GetSelectionSubPanelFromCurrentTool).ListIndex = sa_Border)
    End If
    
End Sub

'Smooth stroke for lasso selections is currently disabled, pending further testing
Private Sub sltSmoothStroke_Change()
    If SelectionsAllowed(False) And (g_CurrentTool = SelectionUI.GetRelevantToolFromSelectShape()) Then
        PDImages.GetActiveImage.MainSelection.SetSelectionProperty sp_SmoothStroke, sltSmoothStroke.Value
        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    End If
End Sub

Private Sub sltWandTolerance_Change()
    If SelectionsAllowed(False) And (g_CurrentTool = SelectionUI.GetRelevantToolFromSelectShape()) Then
        PDImages.GetActiveImage.MainSelection.SetSelectionProperty sp_WandTolerance, sltWandTolerance.Value
        Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    End If
End Sub

Private Sub sltWandTolerance_GotFocusAPI()
    UpdateFlyout 6, True
End Sub

Private Sub sltWandTolerance_SetCustomTabTarget(ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    If shiftTabWasPressed Then
        newTargetHwnd = Me.ttlPanel(6).hWnd
    Else
        newTargetHwnd = Me.cboWandCompare.hWnd
    End If
End Sub

Private Sub spnOpacity_Change(Index As Integer)

    If (Index = 0) Then
        SelectionUI.NotifySelectionRenderChange pdsr_InteriorFillOpacity, spnOpacity(Index).Value
    ElseIf (Index = 1) Then
        SelectionUI.NotifySelectionRenderChange pdsr_ExteriorFillOpacity, spnOpacity(Index).Value
    End If
    
    If SelectionsAllowed(False) Then Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
    
End Sub

Private Sub spnOpacity_GotFocusAPI(Index As Integer)
    UpdateFlyout 0, True
End Sub

Private Sub spnOpacity_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    
    If shiftTabWasPressed Then
        newTargetHwnd = Me.csSelection(Index).hWnd
    Else
        If (Index = 0) Then
            newTargetHwnd = Me.ddAppearance(1).hWnd
        Else
            newTargetHwnd = Me.cmdFlyoutLock(0).hWnd
        End If
    End If
    
End Sub

Private Sub ttlPanel_Click(Index As Integer, ByVal newState As Boolean)
    
    'Normally, clicking the title bar exposes a flyout panel... but this one is a little strange,
    ' because not all appearance options have additional options!  What might be nice is to
    ' *give* them options (like marching ant speed and outline color), but in the meantime,
    ' we just display a blank flyout if options are not available.
    UpdateFlyout Index, newState
    
End Sub

Private Sub ttlPanel_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    
    Select Case Index
    
        '1st titlebar: "appearance"
        Case 0
            If shiftTabWasPressed Then
                
                'Target hWnd varies by current subtool panel
                If (g_CurrentTool = SELECT_RECT) Then
                    newTargetHwnd = Me.cmdFlyoutLock(2).hWnd
                ElseIf (g_CurrentTool = SELECT_CIRC) Then
                    newTargetHwnd = Me.cmdFlyoutLock(3).hWnd
                ElseIf (g_CurrentTool = SELECT_POLYGON) Then
                    newTargetHwnd = Me.cmdFlyoutLock(4).hWnd
                ElseIf (g_CurrentTool = SELECT_LASSO) Then
                    newTargetHwnd = Me.cmdFlyoutLock(5).hWnd
                ElseIf (g_CurrentTool = SELECT_WAND) Then
                    newTargetHwnd = Me.cmdFlyoutLock(6).hWnd
                End If
                
            Else
                newTargetHwnd = Me.chkAppearance(0).hWnd
            End If
        
        '2nd titlebar: "smoothing"
        Case 1
            If shiftTabWasPressed Then
                newTargetHwnd = Me.btsCombine.hWnd
            Else
                newTargetHwnd = Me.cboSelSmoothing.hWnd
            End If
            
        '3rd titlebar: rectangular selections, "size and position"
        Case 2
            If shiftTabWasPressed Then
                newTargetHwnd = Me.cmdFlyoutLock(1).hWnd
            Else
                If Me.tudSel(0).Enabled Then
                    newTargetHwnd = Me.tudSel(0).hWnd
                Else
                    newTargetHwnd = Me.cmdLock(0).hWnd
                End If
            End If
        
        '4th titlebar: ellipse selections, "size and position"
        Case 3
            If shiftTabWasPressed Then
                newTargetHwnd = Me.cmdFlyoutLock(1).hWnd
            Else
                If Me.tudSel(6).Enabled Then
                    newTargetHwnd = Me.tudSel(6).hWnd
                Else
                    newTargetHwnd = Me.cmdLock(3).hWnd
                End If
            End If
            
        '5th titlebar: polygon selections, "curvature"
        Case 4
            If shiftTabWasPressed Then
                newTargetHwnd = Me.cmdFlyoutLock(1).hWnd
            Else
                newTargetHwnd = Me.sltPolygonCurvature.hWndSlider
            End If
            
        '6th titlebar: lasso selections, "area"
        Case 5
            If shiftTabWasPressed Then
                newTargetHwnd = Me.cmdFlyoutLock(1).hWnd
            Else
                newTargetHwnd = Me.cboSelArea(3).hWnd
            End If
            
        '7th titlebar: wand selections, "tolerance"
        Case 6
            If shiftTabWasPressed Then
                newTargetHwnd = Me.cmdFlyoutLock(1).hWnd
            Else
                newTargetHwnd = Me.sltWandTolerance.hWndSlider
            End If
            
    End Select
    
End Sub

'When the selection text boxes are updated, change the scrollbars to match
Private Sub tudSel_Change(Index As Integer)
    UpdateSelectionsValuesViaText Index
End Sub

Private Sub tudSel_GotFocusAPI(Index As Integer)
    If (Index < 6) Then
        UpdateFlyout 2, True
    Else
        UpdateFlyout 3, True
    End If
End Sub

Private Sub tudSel_SetCustomTabTarget(Index As Integer, ByVal shiftTabWasPressed As Boolean, newTargetHwnd As Long)
    
    'We can mirror the same tab order for both rectangular and elliptical selections, with some trickery.
    Dim baseIndex As Long
    If (Index >= 6) Then baseIndex = 6 Else baseIndex = 0
    
    Select Case Index
    
        Case 0, 6
            If shiftTabWasPressed Then
                If (Index = 0) Then
                    newTargetHwnd = Me.ttlPanel(2).hWnd
                Else
                    newTargetHwnd = Me.ttlPanel(3).hWnd
                End If
            Else
                newTargetHwnd = Me.cmdLock(baseIndex \ 2 + 0).hWnd
            End If
        
        Case 1, 7
            If shiftTabWasPressed Then
                newTargetHwnd = Me.cmdLock(baseIndex \ 2 + 0).hWnd
            Else
                newTargetHwnd = Me.cmdLock(baseIndex \ 2 + 1).hWnd
            End If
        
        Case 2, 8
            If shiftTabWasPressed Then
                newTargetHwnd = Me.cmdLock(baseIndex \ 2 + 1).hWnd
            Else
                newTargetHwnd = Me.tudSel(baseIndex + 3).hWnd
            End If
        
        Case 3, 9
            If shiftTabWasPressed Then
                newTargetHwnd = Me.tudSel(baseIndex + 2).hWnd
            Else
                newTargetHwnd = Me.cmdLock(baseIndex \ 2 + 2).hWnd
            End If
            
        Case 4, 10
            If shiftTabWasPressed Then
                newTargetHwnd = Me.cmdLock(baseIndex \ 2 + 2).hWnd
            Else
                newTargetHwnd = Me.tudSel(Index + 1).hWnd
            End If
            
        Case 5, 11
            If shiftTabWasPressed Then
                newTargetHwnd = Me.tudSel(Index - 1).hWnd
            Else
                If (Index = 5) Then
                    newTargetHwnd = Me.chkAutoDrop(0).hWnd
                Else
                    newTargetHwnd = Me.cboSelArea(1).hWnd
                End If
            End If
        
    End Select
    
End Sub

'All text boxes wrap this function.  Note that text box changes are not relayed unless the current selection shape
' matches the current selection tool.
Private Sub UpdateSelectionsValuesViaText(ByVal Index As Integer)
    If SelectionsAllowed(True) Then
        If (Not PDImages.GetActiveImage.MainSelection.GetAutoRefreshSuspend) And (g_CurrentTool = SelectionUI.GetRelevantToolFromSelectShape()) Then
            PDImages.GetActiveImage.MainSelection.UpdateViaTextBox Index
            Viewport.Stage3_CompositeCanvas PDImages.GetActiveImage(), FormMain.MainCanvas(0)
        End If
    End If
End Sub

'Updating against the current theme accomplishes a number of things:
' 1) All user-drawn controls are redrawn according to the current g_Themer settings.
' 2) All tooltips and captions are translated according to the current language.
' 3) ApplyThemeAndTranslations is called, which redraws the form itself according to any theme and/or system settings.
'
'This function is called at least once, at Form_Load, but can be called again if the active language or theme changes.
Public Sub UpdateAgainstCurrentTheme()
    
    Dim buttonSize As Long
    buttonSize = Interface.FixDPI(16)
    
    Dim i As Long
    For i = 0 To cmdLock.Count - 1
        cmdLock(i).AssignImage "generic_unlock", , buttonSize, buttonSize
        cmdLock(i).AssignImage_Pressed "generic_lock", , buttonSize, buttonSize
        cmdLock(i).AssignTooltip "Lock this value.  (Only one value can be locked at a time.  If you lock a new value, previously locked values will unlock.)"
    Next i
    
    buttonSize = Interface.FixDPI(18)
    btsCombine.AssignImageToItem 0, "select_combine_replace", imgWidth:=buttonSize, imgHeight:=buttonSize, resampleAlgorithm:=GP_IM_NearestNeighbor
    btsCombine.AssignImageToItem 1, "select_combine_add", imgWidth:=buttonSize, imgHeight:=buttonSize, resampleAlgorithm:=GP_IM_NearestNeighbor
    btsCombine.AssignImageToItem 2, "select_combine_subtract", imgWidth:=buttonSize, imgHeight:=buttonSize, resampleAlgorithm:=GP_IM_NearestNeighbor
    btsCombine.AssignImageToItem 3, "select_combine_intersect", imgWidth:=buttonSize, imgHeight:=buttonSize, resampleAlgorithm:=GP_IM_NearestNeighbor
    
    'Redrawing the form according to current theme and translation settings.
    ApplyThemeAndTranslations Me
    
    'Tooltips must be manually re-assigned according to the current language.  This is a necessary evil, if the user switches
    ' between two non-English languages at run-time.
    'TODO: tooltips for new rendering options
    'cboSelRender.AssignTooltip "This changes the rendering style of your selection tools.  This setting does not affect selection behavior; it only affects how selections appear on your screen."
    cboSelSmoothing.AssignTooltip "This option controls how smoothly selection edges blend with their surroundings."
    
    For i = 0 To cboSelArea.Count - 1
        cboSelArea(i).AssignTooltip "Selections normally select the area inside their boundaries, but you can modify this behavior here.  (For advanced selection area adjustments, use the Select menu.)"
        sltSelectionBorder(i).AssignTooltip "This option adjusts the width of the selection border."
    Next i
    
    sltSelectionFeathering.AssignTooltip "This feathering slider allows for immediate feathering adjustments.  For performance reasons, it is limited to small radii.  For larger feathering radii, use the Select -> Feathering menu."
    sltCornerRounding.AssignTooltip "This option adjusts the roundness of a rectangular selection's corners."
    For i = chkAutoDrop.lBound To chkAutoDrop.UBound
        chkAutoDrop(i).AssignTooltip "This panel can automatically open when you edit or create a selection.  On smaller screens, this may not be helpful if obscures too much of the canvas, so you can disable this behavior here."
    Next i
    
    sltPolygonCurvature.AssignTooltip "This option adjusts the curvature, if any, of a polygon selection's sides."
    sltSmoothStroke.AssignTooltip "This option increases the smoothness of a hand-drawn lasso selection."
    sltWandTolerance.AssignTooltip "Tolerance controls how similar a pixel must be to the target color before it is added to a magic wand selection."
    
    btsWandMerge.AssignTooltip "The magic wand tool can operate on the entire image, or just the active layer."
    btsWandArea.AssignTooltip "Normally, the magic wand will spread out from the target pixel, adding neighboring pixels to the selection as it goes.  Alternatively, you can have it search the entire image, without regards for continuity."
    
    cboWandCompare.AssignTooltip "This option controls which criteria the magic wand uses to compare pixels to the target color."
    
    'Flyout lock controls use the same behavior across all instances
    UserControls.ThemeFlyoutControls cmdFlyoutLock
    
End Sub

'Update the actively displayed flyout (if any).  Note that the flyout manager will automatically
' hide any other open flyouts, as necessary.
Private Sub UpdateFlyout(ByVal flyoutIndex As Long, Optional ByVal newState As Boolean = True)
    
    'Ensure we have a flyout manager
    If (m_Flyout Is Nothing) Then Set m_Flyout = New pdFlyout
    
    'Exit if we're already in the process of synchronizing
    If m_Flyout.GetFlyoutSyncState() Then Exit Sub
    m_Flyout.SetFlyoutSyncState True
    
    'X offset varies by selection sub-panel (each tool may have slightly different requirements here)
    Dim xOffset As Long
    If (flyoutIndex > 1) Then xOffset = -8
    
    'Ensure we have a flyout manager, then raise the corresponding panel
    If newState Then
        If (flyoutIndex <> m_Flyout.GetFlyoutTrackerID()) Then m_Flyout.ShowFlyout Me, ttlPanel(flyoutIndex), cntrPopOut(flyoutIndex), flyoutIndex, xOffset
    Else
        If (flyoutIndex = m_Flyout.GetFlyoutTrackerID()) Then m_Flyout.HideFlyout
    End If
    
    'Update titlebar state(s) to match
    Dim i As Long
    For i = ttlPanel.lBound To ttlPanel.UBound
        If (i = m_Flyout.GetFlyoutTrackerID()) Then
            If (Not ttlPanel(i).Value) Then ttlPanel(i).Value = True
        Else
            If ttlPanel(i).Value Then ttlPanel(i).Value = False
        End If
    Next i
    
    'Clear the synchronization flag before exiting
    m_Flyout.SetFlyoutSyncState False
    
End Sub

'Some selection flyouts are auto-raised when the user performs a selection action relevant to that flyout.
' The RequestDefaultFlyout sub, below, forwards requests here; this function raises the flyout AND locks
' it in the open position (so it doesn't instantly close due to the user interacting elsewhere).
Private Sub RaiseAndLockFlyout(ByVal flyoutIndex As Long)
        
    If chkAutoDrop(flyoutIndex - 2).Value Then
        
        UpdateFlyout flyoutIndex, True
        
        'By design, pdButtonToolbox controls do not trigger events when their value is set programmatically
        ' (this has to do with the way they were originally designed for the left toolbox only, which is
        ' constantly turning buttons on/off as the user switches between tools).  So we need to manually
        ' set lock status after setting the correct button value.
        cmdFlyoutLock(flyoutIndex).Value = True
        m_Flyout.UpdateLockStatus Me.cntrPopOut(flyoutIndex).hWnd, cmdFlyoutLock(flyoutIndex).Value, cmdFlyoutLock(flyoutIndex)
    
    End If
    
End Sub

'Some selection tools are more useful if we allow their flyout to display automatically.  Obviously the
' panel ID varies by current selection tool.  The canvas control can request flyout behavior via this sub,
' and it can also pass "hints" about behavior (which may enable us to display more useful information in
' the relevant flyout).
Public Sub RequestDefaultFlyout(ByVal xInScreenCoords As Long, ByVal yInScreenCoords As Long, Optional ByVal selectionBeingCreated As Boolean = False, Optional ByVal selectionBeingResized As Boolean = False, Optional ByVal selectionBeingMoved As Boolean = False)
    
    'Only raise a flyout if one isn't already visible.  (Regardless of current selection mode,
    ' we can always check for no flyouts being raised yet this session, or one having been raised
    ' but is closed now.)
    Dim okToRaiseFlyout As Boolean
    okToRaiseFlyout = (m_Flyout Is Nothing)
    If (Not m_Flyout Is Nothing) Then okToRaiseFlyout = okToRaiseFlyout Or (m_Flyout.GetFlyoutTrackerID() < 0)
    
    If (g_CurrentTool = SELECT_RECT) Then
        
        'Display the position/size flyout, and auto-set the dropdown to reflect the relevant setting for
        ' the current action.
        If (Not m_Flyout Is Nothing) Then okToRaiseFlyout = okToRaiseFlyout Or (m_Flyout.GetFlyoutTrackerID = 2)
        If okToRaiseFlyout Then
            
            'One last check - we need to compare the current flyout position (in screen coordinates) to
            ' the current mouse position (in screen coordinates).  If they overlap, we want to *hide* the
            ' flyout because it's impeding the user's view of the mouse cursor.
            Dim hideFlyoutInstead As Boolean
            If (Not m_Flyout Is Nothing) Then
                
                'Note that - by design - this call will retrieve the rect of the last flyout displayed
                ' by the flyout manager.  For selection tools, this will be the auto-dropped panel
                ' showing selection position and/or size.
                Dim curFlyoutRect As RectL
                m_Flyout.GetFlyoutRect curFlyoutRect
                
                'Manually check for overlap with the passed (x, y) coordinates
                If (xInScreenCoords >= curFlyoutRect.Left) And (xInScreenCoords <= curFlyoutRect.Right) Then
                    hideFlyoutInstead = (yInScreenCoords >= curFlyoutRect.Top) And (yInScreenCoords <= curFlyoutRect.Bottom)
                End If
                
            End If
            
            'With all UI settings synced, we can display and lock the corresponding flyout
            If hideFlyoutInstead Then
                UserControls.HideOpenFlyouts 0&
            Else
                RaiseAndLockFlyout 2
            End If
            
        End If
        
    ElseIf (g_CurrentTool = SELECT_CIRC) Then
        
        'TODO: mirror behavior of rectangular selections
        
        'Display the position/size flyout, and auto-set the dropdown to reflect the relevant setting for
        ' the current action.
        If (Not m_Flyout Is Nothing) Then okToRaiseFlyout = okToRaiseFlyout Or (m_Flyout.GetFlyoutTrackerID = 3)
        If okToRaiseFlyout Then
            
            'With all UI settings synced, we can display and lock the corresponding flyout
            RaiseAndLockFlyout 3
            
        End If
    
    ElseIf (g_CurrentTool = SELECT_POLYGON) Then
        'Flyout is not relevant here
    ElseIf (g_CurrentTool = SELECT_LASSO) Then
        'Flyout is not relevant here
    ElseIf (g_CurrentTool = SELECT_WAND) Then
        'Flyout is not relevant here
    End If
    
End Sub
