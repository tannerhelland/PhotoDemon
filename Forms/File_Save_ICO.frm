VERSION 5.00
Begin VB.Form dialog_ExportICO 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   7455
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12630
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   Icon            =   "File_Save_ICO.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   497
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   842
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   6705
      Width           =   12630
      _ExtentX        =   22278
      _ExtentY        =   1323
      DontAutoUnloadParent=   -1  'True
   End
   Begin PhotoDemon.pdDropDown ddIcon 
      Height          =   735
      Left            =   6000
      TabIndex        =   64
      Top             =   1020
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1296
      Caption         =   "icon purpose"
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   6465
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   11404
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   4695
      Index           =   0
      Left            =   5880
      Top             =   1920
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   8281
      Begin PhotoDemon.pdCheckBox chk768 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   4
         Top             =   360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   345
         Index           =   2
         Left            =   2400
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   609
         Caption         =   "32-bpp"
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   345
         Index           =   3
         Left            =   3240
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   609
         Caption         =   "24-bpp"
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   345
         Index           =   4
         Left            =   4080
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   609
         Caption         =   "8-bpp"
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   345
         Index           =   5
         Left            =   4920
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   609
         Caption         =   "4-bpp"
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   345
         Index           =   6
         Left            =   5760
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   609
         Caption         =   "1-bpp"
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   345
         Index           =   7
         Left            =   1650
         Top             =   0
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   609
         Caption         =   "PNG"
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   8
         Left            =   240
         Top             =   1455
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         Alignment       =   1
         Caption         =   "128x128"
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   9
         Left            =   240
         Top             =   375
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         Alignment       =   1
         Caption         =   "768x768"
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   10
         Left            =   240
         Top             =   1095
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         Alignment       =   1
         Caption         =   "256x256"
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   11
         Left            =   240
         Top             =   735
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         Alignment       =   1
         Caption         =   "512x512"
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   12
         Left            =   240
         Top             =   1815
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         Alignment       =   1
         Caption         =   "96x96"
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   13
         Left            =   240
         Top             =   2175
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         Alignment       =   1
         Caption         =   "64x64"
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   14
         Left            =   240
         Top             =   2535
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         Alignment       =   1
         Caption         =   "48x48"
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   15
         Left            =   240
         Top             =   2895
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         Alignment       =   1
         Caption         =   "40x40"
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   16
         Left            =   240
         Top             =   3255
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         Alignment       =   1
         Caption         =   "32x32"
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   17
         Left            =   240
         Top             =   3615
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         Alignment       =   1
         Caption         =   "24x24"
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   18
         Left            =   240
         Top             =   4335
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         Alignment       =   1
         Caption         =   "16x16"
      End
      Begin PhotoDemon.pdCheckBox chk512 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   5
         Top             =   720
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk256 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   6
         Top             =   1080
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
      End
      Begin PhotoDemon.pdCheckBox chk256 
         Height          =   285
         Index           =   1
         Left            =   2520
         TabIndex        =   7
         Top             =   1080
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk256 
         Height          =   285
         Index           =   2
         Left            =   3360
         TabIndex        =   8
         Top             =   1080
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk256 
         Height          =   285
         Index           =   3
         Left            =   4200
         TabIndex        =   9
         Top             =   1080
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk256 
         Height          =   285
         Index           =   4
         Left            =   5040
         TabIndex        =   10
         Top             =   1080
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk256 
         Height          =   285
         Index           =   5
         Left            =   5880
         TabIndex        =   11
         Top             =   1080
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk128 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   12
         Top             =   1440
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk128 
         Height          =   285
         Index           =   1
         Left            =   2520
         TabIndex        =   13
         Top             =   1440
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk128 
         Height          =   285
         Index           =   2
         Left            =   3360
         TabIndex        =   14
         Top             =   1440
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk128 
         Height          =   285
         Index           =   3
         Left            =   4200
         TabIndex        =   15
         Top             =   1440
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk128 
         Height          =   285
         Index           =   4
         Left            =   5040
         TabIndex        =   16
         Top             =   1440
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk128 
         Height          =   285
         Index           =   5
         Left            =   5880
         TabIndex        =   17
         Top             =   1440
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk96 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   18
         Top             =   1800
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk96 
         Height          =   285
         Index           =   1
         Left            =   2520
         TabIndex        =   19
         Top             =   1800
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk96 
         Height          =   285
         Index           =   2
         Left            =   3360
         TabIndex        =   20
         Top             =   1800
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk96 
         Height          =   285
         Index           =   3
         Left            =   4200
         TabIndex        =   21
         Top             =   1800
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk96 
         Height          =   285
         Index           =   4
         Left            =   5040
         TabIndex        =   22
         Top             =   1800
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk96 
         Height          =   285
         Index           =   5
         Left            =   5880
         TabIndex        =   23
         Top             =   1800
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk64 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   24
         Top             =   2160
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk64 
         Height          =   285
         Index           =   1
         Left            =   2520
         TabIndex        =   25
         Top             =   2160
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk64 
         Height          =   285
         Index           =   2
         Left            =   3360
         TabIndex        =   26
         Top             =   2160
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk64 
         Height          =   285
         Index           =   3
         Left            =   4200
         TabIndex        =   27
         Top             =   2160
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk64 
         Height          =   285
         Index           =   4
         Left            =   5040
         TabIndex        =   28
         Top             =   2160
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk64 
         Height          =   285
         Index           =   5
         Left            =   5880
         TabIndex        =   29
         Top             =   2160
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk48 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   30
         Top             =   2520
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk48 
         Height          =   285
         Index           =   1
         Left            =   2520
         TabIndex        =   31
         Top             =   2520
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk48 
         Height          =   285
         Index           =   2
         Left            =   3360
         TabIndex        =   32
         Top             =   2520
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk48 
         Height          =   285
         Index           =   3
         Left            =   4200
         TabIndex        =   33
         Top             =   2520
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk48 
         Height          =   285
         Index           =   4
         Left            =   5040
         TabIndex        =   34
         Top             =   2520
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk48 
         Height          =   285
         Index           =   5
         Left            =   5880
         TabIndex        =   35
         Top             =   2520
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk40 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   36
         Top             =   2880
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk40 
         Height          =   285
         Index           =   1
         Left            =   2520
         TabIndex        =   37
         Top             =   2880
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk40 
         Height          =   285
         Index           =   2
         Left            =   3360
         TabIndex        =   38
         Top             =   2880
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk40 
         Height          =   285
         Index           =   3
         Left            =   4200
         TabIndex        =   39
         Top             =   2880
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk40 
         Height          =   285
         Index           =   4
         Left            =   5040
         TabIndex        =   40
         Top             =   2880
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk40 
         Height          =   285
         Index           =   5
         Left            =   5880
         TabIndex        =   41
         Top             =   2880
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk32 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   42
         Top             =   3240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk32 
         Height          =   285
         Index           =   1
         Left            =   2520
         TabIndex        =   43
         Top             =   3240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk32 
         Height          =   285
         Index           =   2
         Left            =   3360
         TabIndex        =   44
         Top             =   3240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk32 
         Height          =   285
         Index           =   3
         Left            =   4200
         TabIndex        =   45
         Top             =   3240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk32 
         Height          =   285
         Index           =   4
         Left            =   5040
         TabIndex        =   46
         Top             =   3240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk32 
         Height          =   285
         Index           =   5
         Left            =   5880
         TabIndex        =   47
         Top             =   3240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk24 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   48
         Top             =   3600
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk24 
         Height          =   285
         Index           =   1
         Left            =   2520
         TabIndex        =   49
         Top             =   3600
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk24 
         Height          =   285
         Index           =   2
         Left            =   3360
         TabIndex        =   50
         Top             =   3600
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk24 
         Height          =   285
         Index           =   3
         Left            =   4200
         TabIndex        =   51
         Top             =   3600
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk24 
         Height          =   285
         Index           =   4
         Left            =   5040
         TabIndex        =   52
         Top             =   3600
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk24 
         Height          =   285
         Index           =   5
         Left            =   5880
         TabIndex        =   53
         Top             =   3600
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk16 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   54
         Top             =   4320
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk16 
         Height          =   285
         Index           =   1
         Left            =   2520
         TabIndex        =   55
         Top             =   4320
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk16 
         Height          =   285
         Index           =   2
         Left            =   3360
         TabIndex        =   56
         Top             =   4320
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk16 
         Height          =   285
         Index           =   3
         Left            =   4200
         TabIndex        =   57
         Top             =   4320
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk16 
         Height          =   285
         Index           =   4
         Left            =   5040
         TabIndex        =   58
         Top             =   4320
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk16 
         Height          =   285
         Index           =   5
         Left            =   5880
         TabIndex        =   59
         Top             =   4320
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   0
         Left            =   240
         Top             =   3975
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         Alignment       =   1
         Caption         =   "20x20"
      End
      Begin PhotoDemon.pdCheckBox chk20 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   60
         Top             =   3960
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk20 
         Height          =   285
         Index           =   1
         Left            =   2520
         TabIndex        =   61
         Top             =   3960
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk20 
         Height          =   285
         Index           =   2
         Left            =   3360
         TabIndex        =   62
         Top             =   3960
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk20 
         Height          =   285
         Index           =   3
         Left            =   4200
         TabIndex        =   63
         Top             =   3960
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk20 
         Height          =   285
         Index           =   4
         Left            =   5040
         TabIndex        =   2
         Top             =   3960
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
      Begin PhotoDemon.pdCheckBox chk20 
         Height          =   285
         Index           =   5
         Left            =   5880
         TabIndex        =   3
         Top             =   3960
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         Caption         =   ""
         Value           =   0   'False
      End
   End
   Begin PhotoDemon.pdDropDown ddSource 
      Height          =   735
      Left            =   6000
      TabIndex        =   65
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1296
      Caption         =   "icon source"
   End
End
Attribute VB_Name = "dialog_ExportICO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Windows Icon (ICO) Export Dialog
'Copyright 2020-2026 by Tanner Helland
'Created: 11/May/20
'Last updated: 27/October/22
'Last update: allow the user to switch between "merged image" and "match layers to frames" as sources
'
'This dialog works as a simple relay to the pdICO class. Look there for specific encoding details.
'
'I have tried to pare down the UI toggles to only the most essential elements.  Compatibility with
' various OS versions is the big one, especially given that some VB6 users use PD for everything -
' which may mean they want to produce legacy icons.  (Similarly, some WinForms elements, like toolbars,
' still suggest using 24-bpp + 1-bpp alpha, so they have unique requirements.)
'
'I am open to suggestions for improving the feature set and layout of this dialog.  It went through
' many, many prototypes before arriving at its current form, and I think the current layout is an
' optimal combination of simplicity and power... but this is a complicated topic, and most icon
' editors have garbage UIs so they're not exactly a helpful reference!
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This form can (and should!) be notified of the image being exported.  The only exception to this rule
' is invoking the dialog from the batch process dialog, as no image is associated with that preview.
Private m_SrcImage As pdImage

'A composite of the current image, 32-bpp, fully composited.  This is only regenerated if the source
' image changes.
Private m_CompositedImage As pdDIB

'OK or CANCEL result
Private m_UserDialogAnswer As VbMsgBoxResult

'Final format-specific XML packet, with all format-specific settings defined as tag+value pairs
Private m_FormatParamString As String

'Final metadata XML packet, with all metadata settings defined as tag+value pairs
Private m_MetadataParamString As String

'Selecting a different icon target will select a bunch of checkboxes automatically.
' When doing this, we suspend user interaction to prevent infinite loops.
Private m_CheckboxesChanging As Boolean

'When mass-setting the checkbox grid (according to a user selection in the dropdown),
' we set flags to this array; a dedicated function then translates these into
' actual checkbox indices.
Private Enum ICO_ColorDepths
    cd_PNG = 0
    cd_32bpp = 1
    cd_24bpp = 2
    cd_8bpp = 3
    cd_4bpp = 4
    cd_1bpp = 5
    cd_Unknown = 6
End Enum

#If False Then
    Private Const cd_PNG = 0, cd_32bpp = 1, cd_24bpp = 2, cd_8bpp = 3, cd_4bpp = 4, cd_1bpp = 5, cd_Unknown = 6
#End If

Private Enum ICO_Sizes
    sz_768 = 0
    sz_512 = 1
    sz_256 = 2
    sz_128 = 3
    sz_96 = 4
    sz_64 = 5
    sz_48 = 6
    sz_40 = 7
    sz_32 = 8
    sz_24 = 9
    sz_20 = 10
    sz_16 = 11
    sz_Unknown = 12
End Enum

#If False Then
    Private Const sz_768 = 0, sz_512 = 1, sz_256 = 2, sz_128 = 3, sz_96 = 4, sz_64 = 5, sz_48 = 6, sz_40 = 7, sz_32 = 8, sz_24 = 9, sz_20 = 10, sz_16 = 11, sz_Unknown = 12
#End If

Private m_Grid(0 To 11, 0 To 5) As Boolean

'If the ICO being exported was originally loaded from an ICO file, we'll attempt to retrieve
' the original embedding details (size and color-depth).  This allows users to load an ICO,
' make changes, and save out an identical frame list of sizes + color-depths.
Private Type PD_OrigIcon
    icoWidth As Long
    icoHeight As Long
    icoSize As ICO_Sizes
    icoCD As Long
    icoColorDepth As ICO_ColorDepths
    origLayerIndex As Long
End Type

Private m_numOrigIcons As Long, m_OrigIcons() As PD_OrigIcon

'The user's answer is returned via this property
Public Function GetDialogResult() As VbMsgBoxResult
    GetDialogResult = m_UserDialogAnswer
End Function

Public Function GetFormatParams() As String
    GetFormatParams = m_FormatParamString
End Function

Public Function GetMetadataParams() As String
    GetMetadataParams = m_MetadataParamString
End Function

Private Sub chk128_Click(Index As Integer)
    If (Not m_CheckboxesChanging) Then ddIcon.ListIndex = 5
End Sub

Private Sub chk16_Click(Index As Integer)
    If (Not m_CheckboxesChanging) Then ddIcon.ListIndex = 5
End Sub

Private Sub chk20_Click(Index As Integer)
    If (Not m_CheckboxesChanging) Then ddIcon.ListIndex = 5
End Sub

Private Sub chk24_Click(Index As Integer)
    If (Not m_CheckboxesChanging) Then ddIcon.ListIndex = 5
End Sub

Private Sub chk256_Click(Index As Integer)
    If (Not m_CheckboxesChanging) Then ddIcon.ListIndex = 5
End Sub

Private Sub chk32_Click(Index As Integer)
    If (Not m_CheckboxesChanging) Then ddIcon.ListIndex = 5
End Sub

Private Sub chk40_Click(Index As Integer)
    If (Not m_CheckboxesChanging) Then ddIcon.ListIndex = 5
End Sub

Private Sub chk48_Click(Index As Integer)
    If (Not m_CheckboxesChanging) Then ddIcon.ListIndex = 5
End Sub

Private Sub chk512_Click(Index As Integer)
    If (Not m_CheckboxesChanging) Then ddIcon.ListIndex = 5
End Sub

Private Sub chk64_Click(Index As Integer)
    If (Not m_CheckboxesChanging) Then ddIcon.ListIndex = 5
End Sub

Private Sub chk768_Click(Index As Integer)
    If (Not m_CheckboxesChanging) Then ddIcon.ListIndex = 5
End Sub

Private Sub chk96_Click(Index As Integer)
    If (Not m_CheckboxesChanging) Then ddIcon.ListIndex = 5
End Sub

Private Sub cmdBar_CancelClick()
    m_UserDialogAnswer = vbCancel
    Me.Visible = False
End Sub

Private Sub cmdBar_OKClick()

    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    
    'Source as merged image vs match-each-layer-to-an-output-frame
    cParams.AddParam "use-merged-image", (ddSource.ListIndex = 0), True
    
    'Now, this dialog only needs to add one more thing to its parameter list:
    ' a list of all the checked checkboxes.  This is as simple as iterating
    ' through the checkbox collection and writing predictable strings out to
    ' the parameter list.
    Dim icoCD As ICO_ColorDepths, icoSize As ICO_Sizes, writeICO As Boolean, numIcons As Long
    For icoSize = sz_768 To sz_16
        For icoCD = cd_PNG To cd_1bpp
            
            writeICO = False
            
            '768 and 512 icons are PNG format only
            If (icoSize = sz_768) Then
                If (icoCD = cd_PNG) Then writeICO = chk768(cd_PNG).Value
            ElseIf (icoSize = sz_512) Then
                If (icoCD = cd_PNG) Then writeICO = chk512(cd_PNG).Value
            Else
                
                'All other sizes can use variable color-depths
                Select Case icoSize
                    Case sz_256
                        writeICO = chk256(icoCD).Value
                    Case sz_128
                        writeICO = chk128(icoCD).Value
                    Case sz_96
                        writeICO = chk96(icoCD).Value
                    Case sz_64
                        writeICO = chk64(icoCD).Value
                    Case sz_48
                        writeICO = chk48(icoCD).Value
                    Case sz_40
                        writeICO = chk40(icoCD).Value
                    Case sz_32
                        writeICO = chk32(icoCD).Value
                    Case sz_24
                        writeICO = chk24(icoCD).Value
                    Case sz_20
                        writeICO = chk20(icoCD).Value
                    Case sz_16
                        writeICO = chk16(icoCD).Value
                End Select
                
            End If
            
            'Write TRUE values to param string
            If writeICO Then
                
                numIcons = numIcons + 1
                
                Dim strSize As String, strCD As String
                
                Select Case icoSize
                    Case sz_768
                        strSize = "768"
                    Case sz_512
                        strSize = "512"
                    Case sz_256
                        strSize = "256"
                    Case sz_128
                        strSize = "128"
                    Case sz_96
                        strSize = "96"
                    Case sz_64
                        strSize = "64"
                    Case sz_48
                        strSize = "48"
                    Case sz_40
                        strSize = "40"
                    Case sz_32
                        strSize = "32"
                    Case sz_24
                        strSize = "24"
                    Case sz_20
                        strSize = "20"
                    Case sz_16
                        strSize = "16"
                End Select
                
                Select Case icoCD
                    Case cd_PNG
                        strCD = "64"   'Any number > 32 represents "PNG" for purposes of this function
                    Case cd_32bpp
                        strCD = "32"
                    Case cd_24bpp
                        strCD = "24"
                    Case cd_8bpp
                        strCD = "8"
                    Case cd_4bpp
                        strCD = "4"
                    Case cd_1bpp
                        strCD = "1"
                End Select
                
                cParams.AddParam "ico-" & numIcons, strSize & "-" & strSize & "-" & strCD, True, True
                
            End If
        
        Next icoCD
    Next icoSize
    
    'Finally, add the total count of icons in our request list
    cParams.AddParam "icon-count", numIcons, True, True
    m_FormatParamString = cParams.GetParamString()
    
    'ICO files don't support metadata
    m_MetadataParamString = vbNullString
    
    'Free resources that are no longer required
    Set m_CompositedImage = Nothing
    Set m_SrcImage = Nothing
    
    'Hide but *DO NOT UNLOAD* the form.  The dialog manager needs to retrieve the setting strings before unloading us
    m_UserDialogAnswer = vbOK
    Me.Visible = False

End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    ddIcon.ListIndex = 0
    ChooseBestFrameSource
End Sub

Private Sub ddIcon_Click()

    'If m_CheckboxesChanging is already TRUE, we are currently initializing the dialog
    If m_CheckboxesChanging Then Exit Sub
    m_CheckboxesChanging = True
    
    Dim icoCD As ICO_ColorDepths, icoSize As ICO_Sizes

    'Reset all flags
    FillMemory VarPtr(m_Grid(0, 0)), 12 * 6 * 2, 0
    
    'Original file settings (if any)
    If (ddIcon.ListIndex = 6) Then
        
        For icoSize = sz_768 To sz_16
            For icoCD = cd_PNG To cd_1bpp
                m_Grid(icoSize, icoCD) = DoesPrevIcoExist(icoSize, icoCD)
            Next icoCD
        Next icoSize
        
    'Favicon
    ElseIf (ddIcon.ListIndex = 4) Then
    
        'Favicon settings assumed with thanks to
        ' https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/samples/gg491740(v=vs.85)
        ' https://en.wikipedia.org/wiki/Favicon
        ' https://web.archive.org/web/20160306075956/http://www.jonathantneal.com/blog/understand-the-favicon/
        ' https://stackoverflow.com/questions/48956465/favicon-standard-2020-svg-ico-png-and-dimensions
        m_Grid(sz_48, cd_32bpp) = True
        m_Grid(sz_48, cd_8bpp) = True
        m_Grid(sz_32, cd_32bpp) = True
        m_Grid(sz_32, cd_8bpp) = True
        m_Grid(sz_16, cd_32bpp) = True
        m_Grid(sz_16, cd_8bpp) = True
        
    'Other icon groups are standard Windows icons
    Else
    
        'Add legacy Windows icons to all categories
        ' https://docs.microsoft.com/en-us/previous-versions/ms997538(v=msdn.10)
        m_Grid(sz_16, cd_4bpp) = True
        m_Grid(sz_16, cd_8bpp) = True
        m_Grid(sz_32, cd_4bpp) = True
        m_Grid(sz_32, cd_8bpp) = True
        m_Grid(sz_48, cd_8bpp) = True
        
        'Windows XP added support for 32-bpp icons
        If (ddIcon.ListIndex < 3) Then
            m_Grid(sz_16, cd_32bpp) = True
            m_Grid(sz_32, cd_32bpp) = True
            m_Grid(sz_48, cd_32bpp) = True
            m_Grid(sz_256, cd_32bpp) = True
        End If
        
        'Vista added support for PNG icons; these are a lot smaller than uncompressed ones!
        If (ddIcon.ListIndex < 2) Then
            m_Grid(sz_256, cd_32bpp) = False
            m_Grid(sz_256, cd_PNG) = True
            m_Grid(sz_64, cd_32bpp) = True
        End If
        
        'Win 10 added support for mammoth ultra-high-res icons
        If (ddIcon.ListIndex < 1) Then
            m_Grid(sz_512, cd_PNG) = True
        End If
        
    End If
    
    'Populate the checkboxes to match the settings in our flag array
    If (ddIcon.ListIndex <> 5) Then
        
        For icoSize = sz_768 To sz_16
            
            '768 and 512 icons are PNG format only
            If (icoSize = sz_768) Then
                chk768(cd_PNG).Value = m_Grid(icoSize, cd_PNG)
            ElseIf (icoSize = sz_512) Then
                chk512(cd_PNG).Value = m_Grid(icoSize, cd_PNG)
            Else
                
                'All other sizes can use variable color-depths
                For icoCD = cd_PNG To cd_1bpp
                    Select Case icoSize
                        Case sz_256
                            chk256(icoCD).Value = m_Grid(icoSize, icoCD)
                        Case sz_128
                            chk128(icoCD).Value = m_Grid(icoSize, icoCD)
                        Case sz_96
                            chk96(icoCD).Value = m_Grid(icoSize, icoCD)
                        Case sz_64
                            chk64(icoCD).Value = m_Grid(icoSize, icoCD)
                        Case sz_48
                            chk48(icoCD).Value = m_Grid(icoSize, icoCD)
                        Case sz_40
                            chk40(icoCD).Value = m_Grid(icoSize, icoCD)
                        Case sz_32
                            chk32(icoCD).Value = m_Grid(icoSize, icoCD)
                        Case sz_24
                            chk24(icoCD).Value = m_Grid(icoSize, icoCD)
                        Case sz_20
                            chk20(icoCD).Value = m_Grid(icoSize, icoCD)
                        Case sz_16
                            chk16(icoCD).Value = m_Grid(icoSize, icoCD)
                    End Select
                Next icoCD
                
            End If
            
        Next icoSize
        
    End If
    
    m_CheckboxesChanging = False
    
End Sub

Private Function DoesPrevIcoExist(ByVal icoSize As ICO_Sizes, ByVal icoColor As ICO_ColorDepths, Optional ByRef dstIndex As Long = 0) As Boolean
    
    DoesPrevIcoExist = False
    
    If (m_numOrigIcons > 0) Then
    
        Dim i As Long
        For i = 0 To m_numOrigIcons - 1
        
            If (m_OrigIcons(i).icoSize = icoSize) And (m_OrigIcons(i).icoColorDepth = icoColor) Then
                DoesPrevIcoExist = True
                dstIndex = i
                Exit Function
            End If
        
        Next i
    
    End If

End Function

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

'When exporting an ICO file, PD can either use the merged image, or it can attempt to assign each layer to its
' most appropriate icon frame equivalent.  We attempt to choose a "smart" option based on the original file format
' (if one exists).
Private Sub ChooseBestFrameSource()
    
    'When preparing for a batch process, the user may be setting export options but *not* actually exporting
    ' a file right now.  No source image will exist in this case.
    If (Not m_SrcImage Is Nothing) Then
        
        'If the original image was an icon, default to "match each layer to an appropriate output frame"
        If (m_SrcImage.GetOriginalFileFormat = PDIF_ICO) Then
            ddSource.ListIndex = 1
        
        'Otherwise, use the merged image
        Else
            ddSource.ListIndex = 0
        End If
    
    'In batch mode, default to matching layers to "match each layer to an appropriate output frame"
    Else
        ddSource.ListIndex = 1
    End If
    
End Sub

'The ShowDialog routine presents the user with this form.
Public Sub ShowDialog(Optional ByRef srcImage As pdImage = Nothing)
    
    'Suspend UI refreshes until we've initialized
    m_CheckboxesChanging = True
    
    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    m_UserDialogAnswer = vbCancel
    
    'Make sure that the proper cursor is set
    Screen.MousePointer = 0
    Message "Waiting for user to specify export options... "
    
    'Populate the UI
    ddSource.SetAutomaticRedraws False
    ddSource.Clear
    ddSource.AddItem "use merged image", 0
    ddSource.AddItem "treat each layer as a unique icon frame", 1
    
    'Choose the best frame-matching method, but note that this may be overridden later if the user saves
    ' custom preset(s)
    ChooseBestFrameSource
    ddSource.SetAutomaticRedraws True, True
    
    ddIcon.SetAutomaticRedraws False
    ddIcon.Clear
    ddIcon.AddItem "app - Windows 10", 0
    ddIcon.AddItem "app - Windows Vista, 7, 8", 1
    ddIcon.AddItem "app - Windows XP", 2
    ddIcon.AddItem "app - Windows 95, 98, ME", 3
    ddIcon.AddItem "web - favicon", 4, True
    ddIcon.SetAutomaticRedraws True, True
    
    'We now want to see if this image was originally stored in ICO format.  If it was,
    ' we'll give the user a way to save the icon in its current format.
    If (Not srcImage Is Nothing) Then
        
        If (srcImage.GetOriginalFileFormat = PDIF_ICO) Then
            ddIcon.AddItem "custom icon", 5, True
            ddIcon.AddItem "use original file settings", 6
            
            'We now want to retrieve the original ICO file's frame settings (size and color-depth)
            m_numOrigIcons = 0
            
            Const INIT_SIZE As Long = 4
            ReDim m_OrigIcons(0 To INIT_SIZE - 1) As PD_OrigIcon
            
            Dim i As Long, lyrName As String, lyrSettings() As Long, numSettings As Long, lyrIsPNG As Boolean
            For i = 0 To srcImage.GetNumOfLayers - 1
                
                'Parse all integers out of the layer name.  This function returns the number of
                ' individual integers found; for a standard icon loaded into PD, layer names are
                ' formatted as e.g. "32x32 (8-bpp)" so this function will return "32, 32, 8"
                lyrName = srcImage.GetLayerByIndex(i).GetLayerName()
                numSettings = Strings.SplitIntegers(lyrName, lyrSettings, False)
                lyrIsPNG = False
                
                'Note that we only look for 2 settings (width and height).  PNG-encoded frames
                ' won't return a numeric color-depth; we'll check this special case in a moment.
                If (numSettings >= 2) Then
                    
                    'Quick validation of size numbers
                    Dim okToWrite As Boolean
                    okToWrite = (lyrSettings(0) > 0) And (lyrSettings(1) > 0)
                    If okToWrite Then
                        If (numSettings > 2) Then
                            okToWrite = (lyrSettings(2) >= 1)
                        Else
                            lyrIsPNG = (Strings.StrStrI(StrPtr(lyrName), StrPtr("PNG")) > 0)
                            okToWrite = lyrIsPNG
                        End If
                    End If
                    
                    If okToWrite Then
                        
                        'Make sure we have storage space before storing this to our module-level array
                        If (m_numOrigIcons > UBound(m_OrigIcons)) Then ReDim Preserve m_OrigIcons(0 To m_numOrigIcons * 2 - 1) As PD_OrigIcon
                        
                        With m_OrigIcons(m_numOrigIcons)
                            
                            .icoWidth = lyrSettings(0)
                            .icoHeight = lyrSettings(1)
                            
                            'Attempt to match the retrieved sizes against our internal enum list
                            If (.icoWidth = .icoHeight) Then
                                Select Case .icoWidth
                                    Case 768
                                        .icoSize = sz_768
                                    Case 512
                                        .icoSize = sz_512
                                    Case 256
                                        .icoSize = sz_256
                                    Case 128
                                        .icoSize = sz_128
                                    Case 96
                                        .icoSize = sz_96
                                    Case 64
                                        .icoSize = sz_64
                                    Case 48
                                        .icoSize = sz_48
                                    Case 40
                                        .icoSize = sz_40
                                    Case 32
                                        .icoSize = sz_32
                                    Case 24
                                        .icoSize = sz_24
                                    Case 20
                                        .icoSize = sz_20
                                    Case 16
                                        .icoSize = sz_16
                                    Case Else
                                        .icoSize = sz_Unknown
                                End Select
                            Else
                                .icoSize = sz_Unknown
                            End If
                            
                            If lyrIsPNG Then .icoCD = 64 Else .icoCD = lyrSettings(2)
                            
                            'Attempt to match color depth against our internal enum list
                            Select Case .icoCD
                                Case 64
                                    .icoColorDepth = cd_PNG
                                Case 32
                                    .icoColorDepth = cd_32bpp
                                Case 24
                                    .icoColorDepth = cd_24bpp
                                Case 8
                                    .icoColorDepth = cd_8bpp
                                Case 4
                                    .icoColorDepth = cd_4bpp
                                Case 1
                                    .icoColorDepth = cd_1bpp
                                Case Else
                                    .icoColorDepth = cd_Unknown
                            End Select
                            
                        End With
                        
                        'Increment found frame count
                        m_numOrigIcons = m_numOrigIcons + 1
                        
                    End If
                    
                End If
                
            Next i
            
        'Source image isn't ICO
        Else
            ddIcon.AddItem "custom icon", 5
        End If
            
    Else
        ddIcon.AddItem "custom icon", 5
    End If
    
    ddIcon.ListIndex = 0
    
    'Prep a preview (if any)
    Set m_SrcImage = srcImage
    If (Not m_SrcImage Is Nothing) Then
        m_SrcImage.GetCompositedImage m_CompositedImage, True
        pdFxPreview.NotifyNonStandardSource m_CompositedImage.GetDIBWidth, m_CompositedImage.GetDIBHeight
    End If
    If (m_SrcImage Is Nothing) Then Interface.ShowDisabledPreviewImage pdFxPreview
    
    UpdatePreview
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    Interface.SetFormCaptionW Me, g_Language.TranslateMessage("%1 options", "ICO")
    
    'Allow UI refreshes
    m_CheckboxesChanging = False
    
    'Display the dialog
    ShowPDDialog vbModal, Me, True
    
End Sub

Private Sub UpdatePreview()

    If cmdBar.PreviewsAllowed And (Not m_SrcImage Is Nothing) And (Not m_CompositedImage Is Nothing) Then
        
        'Because the user can change the preview viewport, we can't guarantee that the preview region
        ' hasn't changed since the last preview.  Prep a new preview base image now.
        Dim tmpSafeArray As SafeArray2D
        EffectPrep.PreviewNonStandardImage tmpSafeArray, m_CompositedImage, pdFxPreview, True
        EffectPrep.FinalizeNonstandardPreview pdFxPreview, True
        
    End If

End Sub
