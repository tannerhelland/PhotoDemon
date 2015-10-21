VERSION 5.00
Begin VB.Form FormBatchWizard 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Batch Process Wizard"
   ClientHeight    =   9060
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15360
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
   ScaleHeight     =   604
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdButton cmdPrevious 
      Height          =   615
      Left            =   10080
      TabIndex        =   110
      Top             =   8355
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   1085
      Caption         =   "&Previous"
   End
   Begin VB.PictureBox picDragAllow 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   720
      Picture         =   "File_BatchWizard.frx":0000
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   4
      Top             =   7680
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.PictureBox picDragDisallow 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   120
      Picture         =   "File_BatchWizard.frx":09F6
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   3
      Top             =   7680
      Visible         =   0   'False
      Width           =   540
   End
   Begin PhotoDemon.pdButton cmdNext 
      Height          =   615
      Left            =   11880
      TabIndex        =   111
      Top             =   8355
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   1085
      Caption         =   "&Next"
   End
   Begin PhotoDemon.pdButton cmdCancel 
      Height          =   615
      Left            =   13860
      TabIndex        =   112
      Top             =   8355
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1085
      Caption         =   "&Cancel"
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7455
      Index           =   0
      Left            =   3480
      ScaleHeight     =   497
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   785
      TabIndex        =   6
      Top             =   720
      Width           =   11775
      Begin PhotoDemon.pdButton cmdSelectMacro 
         Height          =   615
         Left            =   8640
         TabIndex        =   114
         Top             =   6570
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   1085
         Caption         =   "Select macro..."
         FontSize        =   9
      End
      Begin PhotoDemon.pdTextBox txtMacro 
         Height          =   315
         Left            =   1080
         TabIndex        =   101
         Top             =   6720
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   556
         Text            =   "no macro selected"
      End
      Begin VB.PictureBox picResizeDemo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   7680
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   191
         TabIndex        =   91
         Top             =   5385
         Width           =   2865
      End
      Begin VB.ComboBox cmbResizeFit 
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
         Left            =   3390
         Style           =   2  'Dropdown List
         TabIndex        =   87
         Top             =   5460
         Width           =   4095
      End
      Begin PhotoDemon.smartCheckBox chkActions 
         Height          =   300
         Index           =   2
         Left            =   600
         TabIndex        =   84
         Top             =   6150
         Width           =   10020
         _ExtentX        =   17674
         _ExtentY        =   582
         Caption         =   "custom actions from a saved macro file"
         Value           =   0
      End
      Begin PhotoDemon.smartCheckBox chkActions 
         Height          =   300
         Index           =   1
         Left            =   600
         TabIndex        =   85
         Top             =   2040
         Width           =   10020
         _ExtentX        =   17674
         _ExtentY        =   582
         Caption         =   "resize images"
         Value           =   0
      End
      Begin PhotoDemon.smartOptionButton optActions 
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   88
         Top             =   120
         Width           =   10500
         _ExtentX        =   18521
         _ExtentY        =   582
         Caption         =   "do not apply photo editing actions"
         Value           =   -1  'True
      End
      Begin PhotoDemon.smartOptionButton optActions 
         Height          =   360
         Index           =   1
         Left            =   120
         TabIndex        =   89
         Top             =   1080
         Width           =   10500
         _ExtentX        =   18521
         _ExtentY        =   582
         Caption         =   "apply one or more photo editing actions"
      End
      Begin PhotoDemon.smartCheckBox chkActions 
         Height          =   300
         Index           =   0
         Left            =   600
         TabIndex        =   92
         Top             =   1560
         Width           =   10020
         _ExtentX        =   17674
         _ExtentY        =   582
         Caption         =   "fix exposure and lighting problems"
         Value           =   0
      End
      Begin PhotoDemon.smartResize ucResize 
         Height          =   2850
         Left            =   1080
         TabIndex        =   96
         Top             =   2520
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   5027
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UnknownSizeMode =   -1  'True
      End
      Begin VB.Label lblExplanation 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "if you only want to rename images or change image formats, use this option "
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
         Index           =   1
         Left            =   600
         TabIndex        =   90
         Top             =   540
         Width           =   6615
      End
      Begin VB.Label lblFit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "resize image by:"
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
         Left            =   1875
         TabIndex        =   86
         Top             =   5520
         Width           =   1425
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7500
      Index           =   1
      Left            =   3480
      ScaleHeight     =   500
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   785
      TabIndex        =   5
      Top             =   720
      Width           =   11775
      Begin PhotoDemon.pdButton cmdSaveList 
         Height          =   615
         Left            =   9960
         TabIndex        =   109
         Top             =   6600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         Caption         =   "Save current list..."
         FontSize        =   8
      End
      Begin PhotoDemon.pdButton cmdLoadList 
         Height          =   615
         Left            =   8160
         TabIndex        =   108
         Top             =   6600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         Caption         =   "Load list..."
         FontSize        =   8
      End
      Begin PhotoDemon.pdButton cmdRemoveAll 
         Height          =   615
         Left            =   9960
         TabIndex        =   107
         Top             =   5400
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
         Caption         =   "Remove all images"
         FontSize        =   8
      End
      Begin PhotoDemon.pdButton cmdRemove 
         Height          =   615
         Left            =   8160
         TabIndex        =   106
         Top             =   5400
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
         Caption         =   "Remove selected image(s)"
         FontSize        =   8
      End
      Begin PhotoDemon.pdButton cmdUseCD 
         Height          =   615
         Left            =   8160
         TabIndex        =   105
         Top             =   360
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   1085
         Caption         =   "Add images using ""Open Image"" dialog..."
         FontSize        =   8
      End
      Begin PhotoDemon.pdButton cmdAddFiles 
         Height          =   615
         Left            =   4200
         TabIndex        =   104
         Top             =   4140
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   1085
         Caption         =   "Add selected image(s) to batch list"
         FontSize        =   8
      End
      Begin PhotoDemon.pdButton cmdSelectNone 
         Height          =   615
         Left            =   6120
         TabIndex        =   103
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
         Caption         =   "Select none"
         FontSize        =   8
      End
      Begin PhotoDemon.pdButton cmdSelectAll 
         Height          =   615
         Left            =   4200
         TabIndex        =   102
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
         Caption         =   "Select all"
         FontSize        =   8
      End
      Begin VB.ComboBox cmbPattern 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   510
         Width           =   3645
      End
      Begin VB.DriveListBox Drive1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   3645
      End
      Begin VB.DirListBox Dir1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   2565
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   3615
      End
      Begin VB.ListBox lstFiles 
         ForeColor       =   &H00800000&
         Height          =   2205
         Left            =   240
         MultiSelect     =   2  'Extended
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   11
         Top             =   5040
         Width           =   7575
      End
      Begin VB.PictureBox picPreview 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2445
         Left            =   8160
         ScaleHeight     =   161
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   231
         TabIndex        =   9
         Top             =   1080
         Width           =   3495
      End
      Begin VB.ListBox lstSource 
         ForeColor       =   &H00400000&
         Height          =   2940
         IntegralHeight  =   0   'False
         Left            =   4200
         MultiSelect     =   2  'Extended
         TabIndex        =   8
         Top             =   1080
         Width           =   3615
      End
      Begin PhotoDemon.smartCheckBox chkEnablePreview 
         Height          =   330
         Left            =   8160
         TabIndex        =   10
         Top             =   3600
         Width           =   3510
         _ExtentX        =   6191
         _ExtentY        =   582
         Caption         =   "show image previews"
      End
      Begin VB.Label lblFiles 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "potential source images:"
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
         Left            =   120
         TabIndex        =   18
         Top             =   0
         Width           =   2595
      End
      Begin VB.Label lblTargetFiles 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "batch list:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label lblModify 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "modify batch list:"
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
         Left            =   8040
         TabIndex        =   16
         Top             =   5040
         Width           =   1845
      End
      Begin VB.Label lblLoadSaveList 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "load / save batch list:"
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
         Left            =   8040
         TabIndex        =   15
         Top             =   6240
         Width           =   2265
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000D&
         Index           =   2
         X1              =   8
         X2              =   264
         Y1              =   296
         Y2              =   296
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000D&
         Index           =   1
         X1              =   536
         X2              =   776
         Y1              =   296
         Y2              =   296
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7455
      Index           =   4
      Left            =   3480
      ScaleHeight     =   497
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   785
      TabIndex        =   93
      Top             =   720
      Width           =   11775
      Begin VB.PictureBox picBatchProgress 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   753
         TabIndex        =   94
         Top             =   3360
         Width           =   11295
      End
      Begin VB.Label lblBatchProgress 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(batch conversion process will appear here at run-time)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   645
         Left            =   285
         TabIndex        =   95
         Top             =   2640
         Width           =   11205
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7455
      Index           =   3
      Left            =   3480
      ScaleHeight     =   497
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   785
      TabIndex        =   19
      Top             =   720
      Width           =   11775
      Begin PhotoDemon.pdButton cmdSelectOutputPath 
         Height          =   615
         Left            =   8280
         TabIndex        =   113
         Top             =   435
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1085
         Caption         =   "Select destination folder..."
         FontSize        =   9
      End
      Begin PhotoDemon.pdTextBox txtRenameRemove 
         Height          =   315
         Left            =   840
         TabIndex        =   100
         Top             =   4560
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   556
      End
      Begin PhotoDemon.pdTextBox txtAppendBack 
         Height          =   315
         Left            =   6120
         TabIndex        =   99
         Top             =   3480
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   556
      End
      Begin PhotoDemon.pdTextBox txtAppendFront 
         Height          =   315
         Left            =   840
         TabIndex        =   98
         Top             =   3480
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         Text            =   "NEW_"
      End
      Begin PhotoDemon.pdTextBox txtOutputPath 
         Height          =   315
         Left            =   480
         TabIndex        =   97
         Top             =   600
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   556
         Text            =   "C:\"
      End
      Begin PhotoDemon.smartOptionButton optCase 
         Height          =   330
         Index           =   0
         Left            =   840
         TabIndex        =   32
         Top             =   5640
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   582
         Caption         =   "lowercase"
         Value           =   -1  'True
      End
      Begin PhotoDemon.smartCheckBox chkRenamePrefix 
         Height          =   330
         Left            =   480
         TabIndex        =   28
         Top             =   3000
         Width           =   4650
         _ExtentX        =   8202
         _ExtentY        =   582
         Caption         =   "add a prefix to each filename:"
         Value           =   0
      End
      Begin VB.ComboBox cmbOutputOptions 
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
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   1800
         Width           =   7455
      End
      Begin PhotoDemon.smartCheckBox chkRenameSuffix 
         Height          =   330
         Left            =   5760
         TabIndex        =   29
         Top             =   3000
         Width           =   4650
         _ExtentX        =   8202
         _ExtentY        =   582
         Caption         =   "add a suffix to each filename:"
         Value           =   0
      End
      Begin PhotoDemon.smartCheckBox chkRenameRemove 
         Height          =   330
         Left            =   480
         TabIndex        =   30
         Top             =   4080
         Width           =   6780
         _ExtentX        =   11959
         _ExtentY        =   582
         Caption         =   "remove the following text (if found) from each filename:"
         Value           =   0
      End
      Begin PhotoDemon.smartCheckBox chkRenameCase 
         Height          =   330
         Left            =   480
         TabIndex        =   31
         Top             =   5160
         Width           =   11205
         _ExtentX        =   19764
         _ExtentY        =   582
         Caption         =   "force each filename, including extension, to the following case:"
         Value           =   0
      End
      Begin PhotoDemon.smartOptionButton optCase 
         Height          =   330
         Index           =   1
         Left            =   3240
         TabIndex        =   33
         Top             =   5640
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   582
         Caption         =   "UPPERCASE"
      End
      Begin PhotoDemon.smartCheckBox chkRenameSpaces 
         Height          =   330
         Left            =   480
         TabIndex        =   34
         Top             =   6240
         Width           =   11205
         _ExtentX        =   19764
         _ExtentY        =   582
         Caption         =   "replace spaces in filenames with underscores"
         Value           =   0
      End
      Begin PhotoDemon.smartCheckBox chkRenameCaseSensitive 
         Height          =   330
         Left            =   7560
         TabIndex        =   35
         Top             =   4560
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   582
         Caption         =   "use case-sensitive matching"
         Value           =   0
      End
      Begin VB.Label lblDstFilename 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "after images are processed, save them with the following name:"
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
         Left            =   120
         TabIndex        =   26
         Top             =   1320
         Width           =   6795
      End
      Begin VB.Label lblOptionalText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "additional rename options"
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
         Left            =   120
         TabIndex        =   25
         Top             =   2520
         Width           =   2760
      End
      Begin VB.Label lblDstFolder 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "output images to this folder:"
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
         Left            =   120
         TabIndex        =   24
         Top             =   120
         Width           =   3030
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7455
      Index           =   2
      Left            =   3480
      ScaleHeight     =   497
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   785
      TabIndex        =   7
      Top             =   720
      Width           =   11775
      Begin VB.ComboBox cmbOutputFormat 
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
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1920
         Width           =   7335
      End
      Begin PhotoDemon.smartOptionButton optFormat 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   10515
         _ExtentX        =   18547
         _ExtentY        =   582
         Caption         =   "keep images in their original format"
         Value           =   -1  'True
      End
      Begin PhotoDemon.smartOptionButton optFormat 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   1320
         Width           =   10515
         _ExtentX        =   18547
         _ExtentY        =   582
         Caption         =   "convert all images to a new format"
      End
      Begin VB.PictureBox picFileContainer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4575
         Index           =   6
         Left            =   600
         ScaleHeight     =   305
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   737
         TabIndex        =   69
         Tag             =   "GIF - Graphics Interchange Format"
         Top             =   2520
         Width           =   11055
         Begin PhotoDemon.sliderTextCombo sltThreshold 
            Height          =   405
            Left            =   360
            TabIndex        =   82
            Top             =   1080
            Width           =   8055
            _ExtentX        =   14208
            _ExtentY        =   873
            Max             =   255
            Value           =   127
            NotchPosition   =   2
            NotchValueCustom=   127
         End
         Begin VB.Label lblGIFExplanation 
            BackStyle       =   0  'Transparent
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
            Height          =   1920
            Left            =   480
            TabIndex        =   74
            Top             =   2280
            Width           =   9015
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblInterfaceTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "GIF options"
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
            TabIndex        =   73
            Top             =   120
            Width           =   1230
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "transparency threshold for images with complex (32bpp) alpha channels:"
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
            Index           =   1
            Left            =   360
            TabIndex        =   72
            Top             =   720
            Width           =   6270
         End
         Begin VB.Label lblAfter 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "no transparency "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   0
            Left            =   720
            TabIndex        =   71
            Top             =   1680
            Width           =   1230
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "maximum transparency "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   5520
            TabIndex        =   70
            Top             =   1680
            Width           =   1710
         End
      End
      Begin VB.PictureBox picFileContainer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4575
         Index           =   2
         Left            =   600
         ScaleHeight     =   305
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   737
         TabIndex        =   41
         Tag             =   "PPM - Portable Pixel Map"
         Top             =   2520
         Width           =   11055
         Begin VB.ComboBox cmbPPMFormat 
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
            Left            =   480
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   960
            Width           =   6975
         End
         Begin VB.Label lblPPMEncoding 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "export PPM files using:"
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
            TabIndex        =   44
            Top             =   600
            Width           =   1950
         End
         Begin VB.Label lblInterfaceTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PPM options"
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
            Index           =   12
            Left            =   120
            TabIndex        =   43
            Top             =   120
            Width           =   1305
         End
      End
      Begin VB.PictureBox picFileContainer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4575
         Index           =   4
         Left            =   600
         ScaleHeight     =   305
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   721
         TabIndex        =   36
         Tag             =   "TIFF - Tagged Image File Format"
         Top             =   2520
         Width           =   10815
         Begin VB.ComboBox cmbTIFFCompression 
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
            TabIndex        =   38
            Top             =   960
            Width           =   7095
         End
         Begin PhotoDemon.smartCheckBox chkTIFFCMYK 
            Height          =   330
            Left            =   360
            TabIndex        =   37
            Top             =   1560
            Width           =   7125
            _ExtentX        =   12568
            _ExtentY        =   582
            Caption         =   " save TIFFs as separated CMYK (for printing)"
         End
         Begin VB.Label lblFileStuff 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "when saving, compress TIFFs using:"
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
            Index           =   0
            Left            =   360
            TabIndex        =   40
            Top             =   645
            Width           =   3135
         End
         Begin VB.Label lblInterfaceTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TIFF options"
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
            TabIndex        =   39
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.PictureBox picFileContainer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4575
         Index           =   7
         Left            =   600
         ScaleHeight     =   305
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   705
         TabIndex        =   75
         Tag             =   "JP2 - JPEG 2000"
         Top             =   2520
         Width           =   10575
         Begin PhotoDemon.sliderTextCombo sltJP2Quality 
            Height          =   405
            Left            =   480
            TabIndex        =   83
            Top             =   1650
            Width           =   7935
            _ExtentX        =   13996
            _ExtentY        =   873
            Min             =   1
            Max             =   256
            Value           =   16
            NotchPosition   =   1
         End
         Begin VB.ComboBox cmbJP2SaveQuality 
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
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   77
            Top             =   1110
            Width           =   6855
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "image compression ratio:"
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
            Index           =   2
            Left            =   360
            TabIndex        =   80
            Top             =   720
            Width           =   2190
         End
         Begin VB.Label lblAfter 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "low quality, small file"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   1
            Left            =   5730
            TabIndex        =   79
            Top             =   2160
            Width           =   1470
         End
         Begin VB.Label lblBefore 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "high quality, large file"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   840
            TabIndex        =   78
            Top             =   2160
            Width           =   1545
         End
         Begin VB.Label lblInterfaceTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "JPEG-2000 options"
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
            TabIndex        =   76
            Top             =   120
            Width           =   2025
         End
      End
      Begin VB.PictureBox picFileContainer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4575
         Index           =   0
         Left            =   600
         ScaleHeight     =   305
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   705
         TabIndex        =   56
         Tag             =   "BMP - Windows Bitmap"
         Top             =   2520
         Width           =   10575
         Begin PhotoDemon.smartCheckBox chkBMPRLE 
            Height          =   330
            Left            =   360
            TabIndex        =   57
            Top             =   600
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   582
            Caption         =   "use RLE compression when saving 8bpp BMP images"
            Value           =   0
         End
         Begin VB.Label lblInterfaceTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BMP options"
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
            Index           =   19
            Left            =   120
            TabIndex        =   58
            Top             =   120
            Width           =   1305
         End
      End
      Begin VB.PictureBox picFileContainer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4695
         Index           =   3
         Left            =   600
         ScaleHeight     =   313
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   713
         TabIndex        =   53
         Tag             =   "TGA - Truevision (TARGA)"
         Top             =   2520
         Width           =   10695
         Begin PhotoDemon.smartCheckBox chkTGARLE 
            Height          =   330
            Left            =   360
            TabIndex        =   54
            Top             =   600
            Width           =   7125
            _ExtentX        =   12568
            _ExtentY        =   582
            Caption         =   "use RLE compression when saving TGA images"
         End
         Begin VB.Label lblInterfaceTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TGA options"
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
            Index           =   21
            Left            =   120
            TabIndex        =   55
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.PictureBox picFileContainer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4575
         Index           =   5
         Left            =   600
         ScaleHeight     =   305
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   729
         TabIndex        =   59
         TabStop         =   0   'False
         Tag             =   "JPG - Joint Photographic Experts Group"
         Top             =   2520
         Width           =   10935
         Begin PhotoDemon.sliderTextCombo sltQuality 
            Height          =   405
            Left            =   2640
            TabIndex        =   81
            Top             =   945
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   873
            Min             =   1
            Max             =   99
            Value           =   90
            NotchPosition   =   1
         End
         Begin VB.ComboBox cmbSubsample 
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
            TabIndex        =   68
            ToolTipText     =   "Subsampling affects the way the JPEG encoder compresses image luminance.  4:2:0 (moderate) is the default value."
            Top             =   3840
            Width           =   6735
         End
         Begin VB.ComboBox cmbJPEGSaveQuality 
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
            Left            =   480
            Style           =   2  'Dropdown List
            TabIndex        =   62
            Top             =   990
            Width           =   2055
         End
         Begin PhotoDemon.smartCheckBox chkOptimize 
            Height          =   330
            Left            =   480
            TabIndex        =   61
            Top             =   1920
            Width           =   7050
            _ExtentX        =   12435
            _ExtentY        =   582
            Caption         =   "optimize compression tables"
         End
         Begin PhotoDemon.smartCheckBox chkThumbnail 
            Height          =   330
            Left            =   480
            TabIndex        =   63
            Top             =   2400
            Width           =   7050
            _ExtentX        =   12435
            _ExtentY        =   582
            Caption         =   "embed thumbnail image"
            Value           =   0
         End
         Begin PhotoDemon.smartCheckBox chkProgressive 
            Height          =   330
            Left            =   480
            TabIndex        =   64
            Top             =   2880
            Width           =   7050
            _ExtentX        =   12435
            _ExtentY        =   582
            Caption         =   "use progressive encoding"
            Value           =   0
         End
         Begin PhotoDemon.smartCheckBox chkSubsample 
            Height          =   330
            Left            =   480
            TabIndex        =   65
            Top             =   3360
            Width           =   7050
            _ExtentX        =   12435
            _ExtentY        =   582
            Caption         =   "use specific subsampling:"
            Value           =   0
         End
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "image quality:"
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
            Index           =   0
            Left            =   240
            TabIndex        =   67
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblAdvancedJpegSettings 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "advanced JPEG settings:"
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
            TabIndex        =   66
            Top             =   1560
            Width           =   2070
         End
         Begin VB.Label lblInterfaceTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "JPEG options"
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
            TabIndex        =   60
            Top             =   120
            Width           =   1395
         End
      End
      Begin VB.PictureBox picFileContainer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4575
         Index           =   1
         Left            =   600
         ScaleHeight     =   305
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   729
         TabIndex        =   45
         TabStop         =   0   'False
         Tag             =   "PNG - Portable Network Graphic"
         Top             =   2520
         Width           =   10935
         Begin VB.HScrollBar hsPNGCompression 
            Height          =   330
            Left            =   360
            Max             =   9
            TabIndex        =   48
            Top             =   1080
            Value           =   9
            Width           =   7095
         End
         Begin PhotoDemon.smartCheckBox chkPNGBackground 
            Height          =   330
            Left            =   360
            TabIndex        =   46
            Top             =   2520
            Width           =   7125
            _ExtentX        =   12568
            _ExtentY        =   582
            Caption         =   "preserve file's original background color, if available"
         End
         Begin PhotoDemon.smartCheckBox chkPNGInterlacing 
            Height          =   330
            Left            =   360
            TabIndex        =   47
            Top             =   2040
            Width           =   7125
            _ExtentX        =   12568
            _ExtentY        =   582
            Caption         =   "use interlacing (Adam7)"
            Value           =   0
         End
         Begin VB.Label lblInterfaceTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PNG options"
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
            Index           =   20
            Left            =   120
            TabIndex        =   52
            Top             =   120
            Width           =   1320
         End
         Begin VB.Label lblFileStuff 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "when saving, compress PNG files at the following level:"
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
            Index           =   1
            Left            =   360
            TabIndex        =   51
            Top             =   720
            Width           =   4725
         End
         Begin VB.Label lblPNGCompression 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "no compression"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   0
            Left            =   600
            TabIndex        =   50
            Top             =   1560
            Width           =   1110
         End
         Begin VB.Label lblPNGCompression 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "maximum compression"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   1
            Left            =   5625
            TabIndex        =   49
            Top             =   1560
            Width           =   1590
         End
      End
      Begin VB.Label lblExplanationFormat 
         BackStyle       =   0  'Transparent
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
         Height          =   600
         Left            =   720
         TabIndex        =   22
         Top             =   540
         Width           =   10980
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      Index           =   0
      X1              =   224
      X2              =   224
      Y1              =   48
      Y2              =   544
   End
   Begin VB.Label lblExplanation 
      BackStyle       =   0  'Transparent
      Caption         =   "(text populated at run-time)"
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
      Height          =   7365
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   780
      Width           =   3135
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblWizardTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Step 1: select the photo editing action(s) to apply to each image"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7875
   End
   Begin VB.Label lblBackground 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   -240
      TabIndex        =   0
      Top             =   8280
      Width           =   17415
   End
End
Attribute VB_Name = "FormBatchWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Batch Conversion Form
'Copyright 2007-2015 by Tanner Helland
'Created: 3/Nov/07
'Last updated: 03/September/15
'Last update: convert all buttons to pdButton and overhaul related UI code
'
'PhotoDemon's batch process wizard is one of its most unique - and in my opinion, most impressive - features.  It integrates
' tightly with the macro recording feature to allow any combination of actions to be applied to any set of images.
'
'The process is broken into four steps.
'
'1) Select which photo editing operations (if any) to apply to the images.  This step is optional; if no photo editing actions
'    are selected, a simple format conversion will be applied.
'
'2) Build the batch list, e.g. the list of files to be processed.  This is by far the most complicated section of the wizard.
'    I have revisited the design of this page many times, and I think the current incarnation is pretty damn good.  It exposes
'    a lot of functionality without being overwhelming, and the user has many tools at their disposal to build an ideal list
'    of images from any number of source directories.  (Many batch tools limit you to just one source folder, which I do not
'    like.)
'
'3) Select output file format.  There are three choices: retain original format (e.g. "rename only", which allows the user to
'    use the tool as a batch renamer), pick optimal format for web (which will intermix JPEG and PNG intelligently) - POSTPONED
'    UNTIL 6.2 - or the user can pick their own format.  A comprehensive selection of PhotoDemon's many file format options is
'    also provided.
'
'4) Choose where the files will go and what they will be named.  This includes a number of renaming options, which is a big
'    step up from the batch process tool of earlier versions.  I am open to suggestions for other renaming features, but at
'    present I think the selection is sufficiently comprehensive.
'
'Due to the complexity of this tool, there may be odd combinations of things that don't work quite right - I'm hoping
' others can help test and provide feedback to ensure that everything runs smoothly.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'API constant to add a horizontal scroll bar as necessary - see http://support.microsoft.com/default.aspx?scid=kb%3Ben-us%3B192184
Private Const LB_SETHORIZONTALEXTENT = &H194

'Current active page in the wizard
Dim m_currentPage As Long

'Has the current list of images been saved?
Dim m_ImageListSaved As Boolean

'Current list of image format parameters
Dim m_FormatParams As String

'Currently rendered image preview, and which box it came from (top or bottom)
Dim m_CurImagePreview As String
Dim m_LastPreviewSource As Long

'Because these words are used frequently, if we have to translate them every time they're used, it slows down the
' process considerably.  So cache them in advance.
Dim m_wordForBatchList As String, m_wordForItem As String, m_wordForItems As String

'System progress bar control
Private sysProgBar As cProgressBarOfficial

Private Sub chkActions_Click(Index As Integer)
    
    'If a new action has been selected, activate the "apply photo editing actions" option button
    Dim i As Long
    
    For i = 0 To chkActions.Count - 1
        If CBool(chkActions(i)) Then optActions(1).Value = True
    Next i
    
End Sub

Private Sub chkEnablePreview_Click()
    
    picPreview.Picture = LoadPicture("")
    
    'If the user is disabling previews, clear the picture box and display a notice
    If Not CBool(chkEnablePreview) Then
        Dim strToPrint As String
        strToPrint = g_Language.TranslateMessage("Previews disabled")
        picPreview.CurrentX = (picPreview.ScaleWidth - picPreview.textWidth(strToPrint)) \ 2
        picPreview.CurrentY = (picPreview.ScaleHeight - picPreview.textHeight(strToPrint)) \ 2
        picPreview.Print strToPrint
    'If the user is enabling previews, try to display the last item the user selected in the SOURCE list box
    Else
        If lstSource.Selected(lstSource.ListIndex) Then updatePreview Dir1 & "\" & lstSource.List(lstSource.ListIndex)
    End If
    
End Sub

'By default, neither case-related option button is selected.  Default to lowercase when the RenameCase checkbox is used.
Private Sub chkRenameCase_Click()
    If (Not optCase(0).Value) And (Not optCase(1).Value) Then optCase(0).Value = True
End Sub

'Keep the JPEG-2000 combo box and quality scroll bar in sync
Private Sub cmbJP2SaveQuality_Click()

    Select Case cmbJP2SaveQuality.ListIndex
        
        Case 0
            sltJP2Quality.Value = 1
                
        Case 1
            sltJP2Quality.Value = 16
                
        Case 2
            sltJP2Quality = 32
                
        Case 3
            sltJP2Quality = 64
                
        Case 4
            sltJP2Quality = 256
                
    End Select

End Sub

Private Sub cmbOutputFormat_Click()

    Dim i As Long
    
    For i = 0 To picFileContainer.Count - 1
        If Trim(cmbOutputFormat.List(cmbOutputFormat.ListIndex)) = picFileContainer(i).Tag Then
            picFileContainer(i).Visible = True
        Else
            picFileContainer(i).Visible = False
        End If
    Next i
    
    optFormat(1).Value = True
    
End Sub

'cmbPattern controls the file pattern of the "add images to batch list" box
Private Sub cmbPattern_Click()
    If Me.Visible Then updateSourceImageList
End Sub

Private Sub cmbJPEGSaveQuality_Click()

    Select Case cmbJPEGSaveQuality.ListIndex
        
        Case 0
            sltQuality.Value = 99
                
        Case 1
            sltQuality.Value = 92
                
        Case 2
            sltQuality = 80
                
        Case 3
            sltQuality = 65
                
        Case 4
            sltQuality = 40
                
    End Select
    
End Sub

Private Sub cmbResizeFit_Click()
    
    'Display a sample image of the selected resize method
    Dim tmpDIB As pdDIB
    Set tmpDIB = New pdDIB
    
    'Load the proper sample image to our temporary DIB
    Select Case cmbResizeFit.ListIndex
    
        'Stretch
        Case 0
            loadResourceToDIB "RSZ_STRETCH", tmpDIB
        
        'Fit inclusive
        Case 1
            loadResourceToDIB "RSZ_FITIN", tmpDIB
        
        'Fit exclusive
        Case 2
            loadResourceToDIB "RSZ_FITEX", tmpDIB
    
    End Select
    
    'Paint the sample image to the screen
    picResizeDemo.Picture = LoadPicture("")
    tmpDIB.alphaBlendToDC picResizeDemo.hDC
    picResizeDemo.Picture = picResizeDemo.Image

End Sub

'cmdAddFiles allows the user to move files from the source image list box to the batch list box
Private Sub cmdAddFiles_Click()
    
    Screen.MousePointer = vbHourglass
    Dim x As Long
    For x = 0 To lstSource.ListCount - 1
        If lstSource.Selected(x) Then addFileToBatchList Dir1.Path & "\" & lstSource.List(x)
    Next x
    fixHorizontalListBoxScrolling lstFiles, 16
    Screen.MousePointer = vbDefault
    
End Sub

'Cancel and exit the dialog, with optional prompts as necessary (see Form_QueryUnload)
Private Sub CmdCancel_Click()
    
    If m_currentPage = picContainer.Count - 1 Then
        
        If MacroStatus <> MacroSTOP Then
        
            Dim msgReturn As VbMsgBoxResult
            msgReturn = PDMsgBox("Are you sure you want to cancel the current batch process?", vbApplicationModal + vbYesNoCancel + vbInformation, "Cancel batch processing")
            
            If msgReturn = vbYes Then
                MacroStatus = MacroCANCEL
            End If
            
        Else
            Unload Me
        End If
        
    Else
        Unload Me
    End If
    
End Sub

Private Function allowedToExit() As Boolean

    'If the user has created a list of images to process and they attempt to exit without saving the list,
    ' give them a chance to save it.
    If m_currentPage < picContainer.Count - 1 Then
    
        If (Not m_ImageListSaved) Then
        
            If (lstFiles.ListCount > 0) Then
                Dim msgReturn As VbMsgBoxResult
                msgReturn = PDMsgBox("If you exit now, your batch list (the list of images to be processed) will be lost.  By saving your list, you can easily resume this batch operation at a later date." & vbCrLf & vbCrLf & "Would you like to save your batch list before exiting?", vbApplicationModal + vbExclamation + vbYesNoCancel, "Unsaved image list")
                
                Select Case msgReturn
                    
                    Case vbYes
                        If saveCurrentBatchList() Then allowedToExit = True Else allowedToExit = False
                    
                    Case vbNo
                        allowedToExit = True
                    
                    Case vbCancel
                        allowedToExit = False
                            
                End Select
            Else
                allowedToExit = True
            End If
            
        Else
            allowedToExit = True
        End If
        
    Else
        allowedToExit = True
    End If
    
End Function

'Load a list of images (previously saved from within PhotoDemon) to the batch list
Private Sub cmdLoadList_Click()
    
    Dim sFile As String
    
    'Get the last "open/save image list" path from the preferences file
    Dim tempPathString As String
    tempPathString = g_UserPreferences.GetPref_String("Batch Process", "List Folder", "")
    
    Dim cdFilter As String
    cdFilter = g_Language.TranslateMessage("Batch Image List") & " (.pdl)|*.pdl"
    cdFilter = cdFilter & "|" & g_Language.TranslateMessage("All files") & "|*.*"
    
    Dim cdTitle As String
    cdTitle = g_Language.TranslateMessage("Load a list of images")
    
    Dim openDialog As pdOpenSaveDialog
    Set openDialog = New pdOpenSaveDialog
    
    If openDialog.GetOpenFileName(sFile, , True, False, cdFilter, 1, tempPathString, cdTitle, ".pdl", FormBatchWizard.hWnd) Then
        
        'Save this new directory as the default path for future usage
        Dim listPath As String
        listPath = sFile
        StripDirectory listPath
        g_UserPreferences.SetPref_String "Batch Process", "List Folder", listPath
        
        'Load the file using pdFSO, which is Unicode-compatible
        Dim cFile As pdFSO
        Set cFile = New pdFSO
        
        Dim fileContents As String
        If cFile.LoadTextFileAsString(sFile, fileContents) And (InStr(1, fileContents, vbCrLf) > 0) Then
            
            'The file was originally delimited by vbCrLf.  Parse it now.
            Dim fileLines() As String
            fileLines = Split(fileContents, vbCrLf)
            
            If UBound(fileLines) > 0 Then
                
                'Validate the first line of the file
                If StrComp(fileLines(0), "<" & PROGRAMNAME & " BATCH CONVERSION LIST>", vbTextCompare) = 0 Then
                    
                    'If the user has already created a list of files to process, ask if they want to replace or append
                    ' the loaded entries to their current list.
                    If lstFiles.ListCount > 0 Then
                
                    Dim msgReturn As VbMsgBoxResult
                    msgReturn = PDMsgBox("You have already created a list of images for processing.  The list of images inside this file will be appended to the bottom of your current list.", vbOKCancel + vbApplicationModal + vbInformation, "Batch process notification")
                    
                    If msgReturn = vbCancel Then Exit Sub
                    
                End If
                            
                Screen.MousePointer = vbHourglass
            
                'Now that everything is in place, load the entries from the previously saved file
                Dim numOfEntries As Long
                numOfEntries = CLng(fileLines(1))
                
                Dim suppressDuplicatesCheck As Boolean
                If numOfEntries > 100 Then suppressDuplicatesCheck = True
                
                Dim i As Long
                For i = 2 To numOfEntries + 1
                    addFileToBatchList fileLines(i), suppressDuplicatesCheck
                Next i
                
                fixHorizontalListBoxScrolling lstFiles, 16
                lstFiles.Refresh
                
                Screen.MousePointer = vbDefault
                        
                Else
                    PDMsgBox "This is not a valid list of images. Please try a different file.", vbExclamation + vbApplicationModal + vbOKOnly, "Invalid list file"
                    Exit Sub
                End If
                
            Else
                PDMsgBox "This is not a valid list of images. Please try a different file.", vbExclamation + vbApplicationModal + vbOKOnly, "Invalid list file"
                Exit Sub
            End If
            
        Else
            PDMsgBox "This is not a valid list of images. Please try a different file.", vbExclamation + vbApplicationModal + vbOKOnly, "Invalid list file"
            Exit Sub
        End If
        
        'Note that the current list has been saved (technically it hasn't, I realize, but it exists in a file in this exact state
        ' so close enough!)
        m_ImageListSaved = True
        
    End If
    
End Sub

Private Sub cmdNext_Click()
    changeBatchPage True
End Sub

Private Sub cmdPrevious_Click()
    changeBatchPage False
End Sub

'This function is used to advance (TRUE) or retreat (FALSE) the active wizard panel
Private Sub changeBatchPage(ByVal moveForward As Boolean)

    'Before doing anything else, see if the user is on the final step.  If they are, initiate the batch conversion.
    If moveForward And m_currentPage = picContainer.Count - 2 Then
        m_currentPage = picContainer.Count - 1
        updateWizardText
        prepareForBatchConversion
        Exit Sub
    End If
    
    'Before moving to the next page, validate the current one
    Select Case m_currentPage
    
        'Select photo editing options
        Case 0
        
            'If the user is not applying any photo editing actions, skip to the next step.  If the user IS applying photo editing
            ' actions, additional validations must be applied.
            If optActions(1) Then
            
                'If the user wants to resize the image, make sure the width and height values are valid
                If CBool(chkActions(1)) Then
                    If Not ucResize.IsValid(True) Then Exit Sub
                End If
                
                'If the user wants us to apply a macro, ensure that the macro text box has a macro file specified
                If CBool(chkActions(2)) And ((txtMacro.Text = "no macro selected") Or (Len(txtMacro.Text) = 0)) Then
                    PDMsgBox "You have requested that a macro be applied to each image, but no macro file has been selected.  Please select a valid macro file.", vbExclamation + vbOKOnly + vbApplicationModal, "No macro file selected"
                    txtMacro.selectAll
                    Exit Sub
                End If
                
            End If
            
        'Add images to batch list
        Case 1
        
            'If no images have been added to the batch list, make the user add some!
            If moveForward And lstFiles.ListCount = 0 Then
                PDMsgBox "You have not selected any images to process!  Please place one or more images in the batch list (at the bottom of the screen) before moving to the next step.", vbExclamation + vbOKOnly + vbApplicationModal, "No images selected"
                Exit Sub
            End If
        
        'Select output format
        Case 2
        
            'If the user has asked us to convert all images to a new format, we need to build a parameter string that
            ' contains all of the user's selected image format options (JPEG quality, etc)
            If optFormat(1) Then
            
                Select Case g_ImageFormats.getOutputFIF(cmbOutputFormat.ListIndex)
                
                    Case FIF_BMP
                        m_FormatParams = buildParams(CBool(chkBMPRLE))
                    
                    Case FIF_GIF
                        If sltThreshold.IsValid Then
                            m_FormatParams = buildParams(sltThreshold.Value)
                        Else
                            Exit Sub
                        End If
                    
                    Case FIF_JP2
                            'Determine the compression ratio for the JPEG-2000 wavelet transformation
                            If sltJP2Quality.IsValid Then
                                m_FormatParams = buildParams(sltQuality.Value)
                            Else
                                Exit Sub
                            End If
                    
                    Case FIF_JPEG
                        
                        'JPEG options are complicated, on account of some params being required (quality) but others not (thumbnail, etc)
                        
                        'First, determine the compression quality for the quantization tables
                        If sltQuality.IsValid Then
                            m_FormatParams = buildParams(sltQuality)
                        Else
                            Exit Sub
                        End If
                        
                        'Determine any extra flags based on the advanced settings
                        Dim tmpJPEGFlags As Long
                        tmpJPEGFlags = 0
                        
                        'Optimize
                        If CBool(chkOptimize) Then tmpJPEGFlags = tmpJPEGFlags Or JPEG_OPTIMIZE
        
                        'Progressive scan
                        If CBool(chkProgressive) Then tmpJPEGFlags = tmpJPEGFlags Or JPEG_PROGRESSIVE
        
                        'Subsampling
                        If CBool(chkSubsample) Then
    
                            Select Case cmbSubsample.ListIndex
                                Case 0
                                    tmpJPEGFlags = tmpJPEGFlags Or JPEG_SUBSAMPLING_444
                                Case 1
                                    tmpJPEGFlags = tmpJPEGFlags Or JPEG_SUBSAMPLING_422
                                Case 2
                                    tmpJPEGFlags = tmpJPEGFlags Or JPEG_SUBSAMPLING_420
                                Case 3
                                    tmpJPEGFlags = tmpJPEGFlags Or JPEG_SUBSAMPLING_411
                            End Select
            
                        End If
                        
                        m_FormatParams = m_FormatParams & "|" & Trim$(Str(tmpJPEGFlags))
        
                        'Finally, determine whether or not a thumbnail version of the file should be embedded inside
                        If CBool(chkThumbnail) Then
                            m_FormatParams = m_FormatParams & "|1"
                        Else
                            m_FormatParams = m_FormatParams & "|0"
                        End If
                        
                        'FOR NOW, disable automatic JPEG quality calculations.  This must be done manually on a per-image basis.
                        m_FormatParams = m_FormatParams & "|0"
                                        
                    Case FIF_PNG
                        m_FormatParams = Trim$(Str(hsPNGCompression))
                        m_FormatParams = m_FormatParams & "|" & Trim$(Str(chkPNGInterlacing))
                        m_FormatParams = m_FormatParams & "|" & Trim$(Str(chkPNGBackground))
                    
                    Case FIF_PPM
                        m_FormatParams = Trim$(Str(cmbPPMFormat.ListIndex))
                        
                    Case FIF_TARGA
                        m_FormatParams = Trim$(Str(CBool(chkTGARLE)))
                    
                    Case FIF_TIFF
                        m_FormatParams = Trim$(Str(cmbTIFFCompression.ListIndex))
                        m_FormatParams = m_FormatParams & "|" & Trim$(Str(CBool(chkTIFFCMYK)))
                
                End Select
            
            End If
        
        'Select output directory and file name
        Case 3
            
            Dim cFile As pdFSO
            Set cFile = New pdFSO
            
            'Make sure we have write access to the output folder.  If we don't, cancel and warn the user.
            If Not cFile.FolderExist(txtOutputPath) Then
                
                If Not cFile.CreateFolder(txtOutputPath) Then
                    PDMsgBox "PhotoDemon cannot access the requested output folder.  Please select a non-system, unrestricted folder for the batch process.", vbExclamation + vbOKOnly + vbApplicationModal, "Folder access unavailable"
                    txtOutputPath.selectAll
                    Exit Sub
                End If
                
            End If
    
    End Select

    'True means move forward; false means move backward
    If moveForward Then m_currentPage = m_currentPage + 1 Else m_currentPage = m_currentPage - 1
        
    'Hide all inactive panels (and show the active one)
    Dim i As Long
    For i = 0 To picContainer.Count - 1
        If i = m_currentPage Then
            picContainer(i).Visible = True
        Else
            picContainer(i).Visible = False
        End If
    Next i
    
    'If we are at the beginning, disable the previous button
    If m_currentPage = 0 Then cmdPrevious.Enabled = False Else cmdPrevious.Enabled = True
    
    'If we are at the end, change the text of the "next" button; otherwise, make sure it says "next"
    If m_currentPage = picContainer.Count - 2 Then
        cmdNext.Caption = g_Language.TranslateMessage("Start processing!")
    Else
        If cmdNext.Caption <> g_Language.TranslateMessage("Next") Then cmdNext.Caption = g_Language.TranslateMessage("Next")
    End If
    
    'Finally, update all the label captions that change according to the active panel
    updateWizardText
    
End Sub

'Used to display unique text for each page of the wizard.  The value of m_currentPage is used to determine what text to display.
Private Sub updateWizardText()

    Dim sideText As String
    sideText = "(description forthcoming)"

    Select Case m_currentPage
        
        'Step 1: choose what photo editing you will apply to each image
        Case 0
        
            lblWizardTitle.Caption = g_Language.TranslateMessage("Step 1: select the photo editing action(s) to apply to each image")
            
            sideText = g_Language.TranslateMessage("Welcome to PhotoDemon's batch wizard.  This tool can be used to edit multiple images at once, in what is called a ""batch process"".")
            sideText = sideText & vbCrLf & vbCrLf & g_Language.TranslateMessage("Start by selecting the photo editing action(s) you want to apply.  If multiple actions are selected, they will be applied in the order they appear on this page.")
            sideText = sideText & vbCrLf & vbCrLf & g_Language.TranslateMessage("Note: a ""macro"" is simply a list of photo editing actions.  It can include any adjustment, filter, or effect in the main program.  You can create a new macro by using the ""Tools -> Macros -> Record new macro"" menu in the main PhotoDemon window.")
            sideText = sideText & vbCrLf & vbCrLf & g_Language.TranslateMessage("In the next step, you will select the images you want to process.")
            
        'Step 2: add images to list
        Case 1
        
            lblWizardTitle.Caption = g_Language.TranslateMessage("Step 2: prepare the batch list (the list of images to be processed)")
            
            sideText = g_Language.TranslateMessage("You can add files to the batch list in several ways:")
            sideText = sideText & vbCrLf & vbCrLf & "  " & g_Language.TranslateMessage("1) The folder and file lists at the top of this page.  Use the ""Add selected image(s) to batch list"" button to move images to the batch list, or use the right mouse button to drag-and-drop images.")
            sideText = sideText & vbCrLf & vbCrLf & "  " & g_Language.TranslateMessage("2) The ""Add images using Open Image dialog..."" button.")
            sideText = sideText & vbCrLf & vbCrLf & "  " & g_Language.TranslateMessage("3) Drag-and-drop files directly onto the batch list from Windows Explorer or your desktop.")
            sideText = sideText & vbCrLf & vbCrLf & g_Language.TranslateMessage("Each of these methods supports use of the Ctrl and Shift keys to select multiple files at once.")
            sideText = sideText & vbCrLf & vbCrLf & g_Language.TranslateMessage("In the next step, you will choose how you want the processed images saved.")
        
        'Step 3: choose the output image format
        Case 2
        
            lblWizardTitle.Caption = g_Language.TranslateMessage("Step 3: choose a destination image format")
            
            sideText = g_Language.TranslateMessage("PhotoDemon needs to know which format to use when saving the images in your batch list.")
            sideText = sideText & vbCrLf & vbCrLf & g_Language.TranslateMessage("If ""keep images in their original format"" is selected, PhotoDemon will attempt to save each image in its original format.  If the original format is not supported, a standard format (JPEG or PNG, depending on color depth) will be used.")
            sideText = sideText & vbCrLf & vbCrLf & g_Language.TranslateMessage("If you choose to save images to a new format, please make sure the format you have selected is appropriate for all images in your list.  (For example, images with transparency should be saved to a format that supports transparency!)")
            sideText = sideText & vbCrLf & vbCrLf & g_Language.TranslateMessage("In the final step, you will choose how you want the saved files to be named.")
            
        'Step 4: choose where processed images will be placed and named
        Case 3
        
            lblWizardTitle.Caption = g_Language.TranslateMessage("Step 4: provide a destination folder and any renaming options")
            
            sideText = g_Language.TranslateMessage("In this final step, PhotoDemon needs to know where to save the processed images, and what name to give the new files.")
            sideText = sideText & vbCrLf & vbCrLf & g_Language.TranslateMessage("For your convenience, a number of standard renaming options are also provided.  Note that all items under ""additional rename options"" are optional.")
            sideText = sideText & vbCrLf & vbCrLf & g_Language.TranslateMessage("Finally, if two or more images in the batch list have the same filename, and the ""original filenames"" option is selected, such files will automatically be given unique filenames upon saving (e.g. ""original-filename (2)"").")
        
        'Step 5: process!
        Case 4
            lblWizardTitle.Caption = g_Language.TranslateMessage("Step 5: wait for batch processing to finish")
            
            sideText = g_Language.TranslateMessage("Batch processing is now underway.")
            sideText = sideText & vbCrLf & vbCrLf & g_Language.TranslateMessage("Once the batch processor has processed several images, it will display an estimated time remaining.")
            sideText = sideText & vbCrLf & vbCrLf & g_Language.TranslateMessage("You can cancel batch processing at any time by pressing the ""Cancel"" button in the bottom-right corner.  If you choose to cancel, any processed images will still be present in the output folder, so you may need to remove them manually.")
            
    End Select
    
    lblExplanation(0).Caption = sideText
    
    'If translations are active, the translated text may not fit the label.  Automatically adjust it to fit.
    FitWordwrapLabel lblExplanation(0), Me
    
End Sub

'Remove all selected items from the batch conversion list
Private Sub cmdRemove_Click()
        
    Dim x As Long
    For x = lstFiles.ListCount - 1 To 0 Step -1
        If lstFiles.Selected(x) Then lstFiles.RemoveItem x
    Next x
    
    'Because there are no longer any selected entries, disable the "remove selected images" button
    cmdRemove.Enabled = False
    
    'And if all files were removed, disable actions that require at least one image
    If lstFiles.ListCount = 0 Then
        cmdRemoveAll.Enabled = False
        cmdSaveList.Enabled = False
        'cmdNext.Enabled = False
    End If
    
    'Note that the current list has NOT been saved
    m_ImageListSaved = False
    
    'Update the label that displays the number of items in the list
    updateBatchListCount
    
    'If the lower box was the source of the current image preview, erase the preview now
    If m_LastPreviewSource = 1 Then updatePreview ""
        
End Sub

'Remove ALL items from the batch conversion list
Private Sub cmdRemoveAll_Click()
    
    lstFiles.Clear
    fixHorizontalListBoxScrolling lstFiles
    
    'If the lower box was the source of the current image preview, erase the preview now
    If m_LastPreviewSource = 1 Then updatePreview ""
    
    'Because all entries have been removed, disable actions that require at least one image to be present
    cmdRemove.Enabled = False
    cmdRemoveAll.Enabled = False
    cmdSaveList.Enabled = False
    'cmdNext.Enabled = False
    
    'Note that the current list has NOT been saved
    m_ImageListSaved = False
    
    'Update the label that displays the number of items in the list
    updateBatchListCount
    
End Sub

Private Function saveCurrentBatchList() As Boolean

    'Get the last "open/save image list" path from the preferences file
    Dim tempPathString As String
    tempPathString = g_UserPreferences.GetPref_String("Batch Process", "List Folder", "")
    
    Dim cdFilter As String
    cdFilter = g_Language.TranslateMessage("Batch Image List") & " (.pdl)|*.pdl"
    
    Dim cdTitle As String
    cdTitle = g_Language.TranslateMessage("Save the current list of images")
    
    Dim saveDialog As pdOpenSaveDialog
    Set saveDialog = New pdOpenSaveDialog
    
    Dim sFile As String
    If saveDialog.GetSaveFileName(sFile, , True, cdFilter, 1, tempPathString, cdTitle, ".pdl", FormBatchWizard.hWnd) Then
        
        'Save this new directory as the default path for future usage
        Dim listPath As String
        listPath = sFile
        StripDirectory listPath
        g_UserPreferences.SetPref_String "Batch Process", "List Folder", listPath
        
        'Assemble the output string, which basically just contains the currently selected list of files.
        Dim outputText As String
        
        outputText = "<" & PROGRAMNAME & " BATCH CONVERSION LIST>" & vbCrLf
        outputText = outputText & Trim$(Str(lstFiles.ListCount)) & vbCrLf
        
        Dim i As Long
        For i = 0 To lstFiles.ListCount - 1
            outputText = outputText & lstFiles.List(i) & vbCrLf
        Next i
        
        outputText = outputText & "<END OF LIST>" & vbCrLf
        
        'Write the text out to file using a pdFSO instance
        Dim cFile As pdFSO
        Set cFile = New pdFSO
        
        saveCurrentBatchList = cFile.SaveStringToTextFile(outputText, sFile)
                
    Else
        saveCurrentBatchList = False
    End If

End Function

Private Sub cmdSaveList_Click()
    
    'Before attempting to save, make sure at least one image has been placed in the list
    If lstFiles.ListCount = 0 Then
        PDMsgBox "You haven't selected any image files.  Please add one or more files to the batch list before saving.", vbExclamation + vbOKOnly + vbApplicationModal, "Empty image list"
        Exit Sub
    End If
        
    saveCurrentBatchList
    
    'Note that the current list has been saved
    m_ImageListSaved = True
    
End Sub

'Select every image file currently displayed in the source files box
Private Sub cmdSelectAll_Click()

    Screen.MousePointer = vbHourglass
    LockWindowUpdate lstSource.hWnd

    'If image previews are currently enabled, disable them before selecting all (to speed up processing)
    Dim enablePreviews As Boolean
    If CBool(chkEnablePreview) Then
        enablePreviews = True
        chkEnablePreview.Value = vbUnchecked
    End If

    Dim x As Long
    For x = 0 To lstSource.ListCount - 1
        lstSource.Selected(x) = True
    Next x

    'Restore the user's preference upon completion
    If enablePreviews Then chkEnablePreview.Value = vbChecked
    
    LockWindowUpdate 0
    Screen.MousePointer = vbDefault
    
End Sub

'Open a common dialog and allow the user to select a macro file (to apply to each image in the batch list)
Private Sub cmdSelectMacro_Click()
    
    'Get the last macro-related path from the preferences file
    Dim tempPathString As String
    tempPathString = g_UserPreferences.GetPref_String("Paths", "Macro", "")
    
    Dim cdFilter As String
    cdFilter = PROGRAMNAME & " " & g_Language.TranslateMessage("Macro Data") & " (." & MACRO_EXT & ")|*." & MACRO_EXT & ";*.thm"
    cdFilter = cdFilter & "|" & g_Language.TranslateMessage("All files") & "|*.*"
    
    'Prepare a common dialog object
    Dim openDialog As pdOpenSaveDialog
    Set openDialog = New pdOpenSaveDialog
    
    Dim sFile As String
   
    'If the user provides a valid macro file, use that as part of the batch process
    If openDialog.GetOpenFileName(sFile, , True, False, cdFilter, 1, tempPathString, g_Language.TranslateMessage("Open Macro File"), "." & MACRO_EXT, Me.hWnd) Then
        
        'As a convenience to the user, save this directory as the default macro path
        tempPathString = sFile
        StripDirectory tempPathString
        g_UserPreferences.SetPref_String "Paths", "Macro", tempPathString
        
        'Display the selected macro location in the relevant text box
        txtMacro.Text = sFile
        
        'Also, select the macro option button by default
        chkActions(2).Value = vbChecked
        
    End If

End Sub

'Unselect all files in the top-center list box
Private Sub cmdSelectNone_Click()

    Dim enablePreviews As Boolean
    If CBool(chkEnablePreview) Then
        enablePreviews = True
        chkEnablePreview.Value = vbUnchecked
    End If

    Dim x As Long
    For x = 0 To lstSource.ListCount - 1
        lstSource.Selected(x) = False
    Next x
    
    If enablePreviews Then chkEnablePreview.Value = vbChecked

End Sub

'Use "shell32.dll" to select a folder
Private Sub cmdSelectOutputPath_Click()
    Dim tString As String
    tString = BrowseForFolder(FormBatchWizard.hWnd)
    If tString <> "" Then
        txtOutputPath.Text = FixPath(tString)
    
        'Save this new directory as the default path for future usage
        g_UserPreferences.SetPref_String "Batch Process", "Output Folder", tString
    End If
End Sub

'Allow the user to use an "Open Image" dialog to add files to the batch convert list
Private Sub cmdUseCD_Click()
    'String returned from the common dialog wrapper
    Dim sFile() As String
    
    If PhotoDemon_OpenImageDialog(sFile, Me.hWnd) Then
        
        Dim x As Long
        For x = 0 To UBound(sFile)
            addFileToBatchList sFile(x)
        Next x
        fixHorizontalListBoxScrolling lstFiles, 16
    End If
End Sub

Private Sub Dir1_Change()
    If Me.Visible Then updateSourceImageList
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1
End Sub

'Dragged image files must be placed on the batch listbox - not anywhere else.
Private Sub Form_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    If Source = lstSource Then lstSource.DragIcon = picDragDisallow.Picture
End Sub

Private Sub Form_Load()
    
    Dim i As Long
    
    'Populate all photo-editing-action-related combo boxes, tooltip, and options
        
        'Resize fit types
            cmbResizeFit.Clear
            cmbResizeFit.AddItem "stretching to fit", 0
            cmbResizeFit.AddItem "fit inclusively", 1
            cmbResizeFit.AddItem "fit exclusively", 2
            cmbResizeFit.ListIndex = 0
        
        'For convenience, change the default resize width and height to the current screen resolution
            ucResize.setInitialDimensions Screen.Width / TwipsPerPixelXFix, Screen.Height / TwipsPerPixelYFix
            
        'By default, select "apply no photo editing actions"
            For i = 0 To chkActions.Count - 1
                chkActions(i).Value = vbUnchecked
            Next i
            optActions(0).Value = True
                
    'Populate all file-format-related combo boxes, tooltips, and options
    
        'GIF
            lblGIFExplanation.Caption = g_Language.TranslateMessage("GIF images only support simple transparency, meaning each pixel in a GIF image must be fully transparent or fully opaque.  If your batch list includes 32bpp images with complex alpha channels, PhotoDemon must simplify them before saving to GIF format." & vbCrLf & vbCrLf & "By default, transparency will be split down the middle, or you can specify a custom cut-off value.")
    
        'JP2 (JPEG-2000)
        
            'Populate the quality drop-down box with presets corresponding to the JPEG-2000 file format
            cmbJP2SaveQuality.Clear
            cmbJP2SaveQuality.AddItem " Lossless (1:1)", 0
            cmbJP2SaveQuality.AddItem " Low compression, good image quality (16:1)", 1
            cmbJP2SaveQuality.AddItem " Moderate compression, medium image quality (32:1)", 2
            cmbJP2SaveQuality.AddItem " High compression, poor image quality (64:1)", 3
            cmbJP2SaveQuality.AddItem " Super compression, very poor image quality (256:1)", 4
            cmbJP2SaveQuality.AddItem " Custom ratio (X:1)", 5
            cmbJP2SaveQuality.ListIndex = 0
                    
        'JPEG
    
            'Populate the quality drop-down box with presets corresponding to the JPEG format
            cmbJPEGSaveQuality.Clear
            cmbJPEGSaveQuality.AddItem " Perfect (99)", 0
            cmbJPEGSaveQuality.AddItem " Excellent (92)", 1
            cmbJPEGSaveQuality.AddItem " Good (80)", 2
            cmbJPEGSaveQuality.AddItem " Average (65)", 3
            cmbJPEGSaveQuality.AddItem " Low (40)", 4
            cmbJPEGSaveQuality.AddItem " Custom value", 5
            cmbJPEGSaveQuality.ListIndex = 1
                
            'Populate the custom subsampling combo box as well
            cmbSubsample.Clear
            cmbSubsample.AddItem " 4:4:4 (best quality)", 0
            cmbSubsample.AddItem " 4:2:2 (good quality)", 1
            cmbSubsample.AddItem " 4:2:0 (moderate quality)", 2
            cmbSubsample.AddItem " 4:1:1 (low quality)", 3
            cmbSubsample.ListIndex = 2
            
            'If FreeImage is not available, disable all the advanced JPEG settings
            If Not g_ImageFormats.FreeImageEnabled Then
                chkOptimize.Enabled = False
                chkProgressive.Enabled = False
                chkSubsample.Enabled = False
                chkThumbnail.Enabled = False
                cmbSubsample.AddItem "n/a", 4
                cmbSubsample.ListIndex = 4
                cmbSubsample.Enabled = False
                lblAdvancedJpegSettings.Caption = g_Language.TranslateMessage("advanced settings require the FreeImage plugin")
            End If
        
            'Apply some tooltips manually (so the translation engine can find them)
            chkOptimize.AssignTooltip "Optimization is highly recommended.  This option allows the JPEG encoder to compute an optimal Huffman coding table for the file.  It does not affect image quality - only file size."
            chkThumbnail.AssignTooltip "Embedded thumbnails increase file size, but they help previews of the image appear more quickly in other software (e.g. Windows Explorer)."
            chkProgressive.AssignTooltip "Progressive encoding is sometimes used for JPEG files that will be used on the Internet.  It saves the image in three steps, which can be used to gradually fade-in the image on a slow Internet connection."
            
        'PPM export
        
            cmbPPMFormat.Clear
            cmbPPMFormat.AddItem " binary encoding (faster, smaller file size)", 0
            cmbPPMFormat.AddItem " ASCII encoding (human-readable, multi-platform)", 1
            cmbPPMFormat.ListIndex = 0
        
        'TIFF export
        
            cmbTIFFCompression.Clear
            cmbTIFFCompression.AddItem " default settings - CCITT Group 4 for 1bpp, LZW for all others", 0
            cmbTIFFCompression.AddItem " no compression", 1
            cmbTIFFCompression.AddItem " Macintosh PackBits (RLE)", 2
            cmbTIFFCompression.AddItem " Official DEFLATE ('Adobe-style')", 3
            cmbTIFFCompression.AddItem " PKZIP DEFLATE (also known as zLib DEFLATE)", 4
            cmbTIFFCompression.AddItem " LZW", 5
            cmbTIFFCompression.AddItem " JPEG - 8bpp grayscale or 24bpp color only", 6
            cmbTIFFCompression.AddItem " CCITT Group 3 fax encoding - 1bpp only", 7
            cmbTIFFCompression.AddItem " CCITT Group 4 fax encoding - 1bpp only", 8
            cmbTIFFCompression.ListIndex = 0
        
        'Misc file-related tooltips that are too long to add at design-time
        
            chkBMPRLE.AssignTooltip "Bitmap files only support one type of compression, and they only support it for certain color depths.  PhotoDemon can apply simple RLE compression when saving 8bpp images."
            chkTGARLE.AssignTooltip "TGA files only support one type of compression.  PhotoDemon can apply simple RLE compression when saving TGA images."
            chkTIFFCMYK.AssignTooltip "TIFFs support both RGB and CMYK color spaces.  RGB is used by default, but if a TIFF file is going to be used in printed document, CMYK is sometimes required."
            cmbTIFFCompression.ToolTipText = g_Language.TranslateMessage("TIFFs support a variety of compression techniques.  Some of these techniques are limited to specific color depths, so make sure you pick one that matches the images you plan on saving.")
            chkPNGInterlacing.AssignTooltip "PNG interlacing is similar to ""progressive scan"" on JPEGs.  Interlacing slightly increases file size, but an interlaced image can ""fade-in"" while it downloads."
            chkPNGBackground.AssignTooltip "PNG files can contain a background color parameter.  This takes up extra space in the file, so feel free to disable it if you don't need background colors."
            cmbPPMFormat.ToolTipText = g_Language.TranslateMessage("Binary encoding of PPM files is strongly suggested.  (In other words, don't change this setting unless you are certain that ASCII encoding is what you want. :)")
            
    'Build default paths from preference file values
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    Dim tempPathString As String
    tempPathString = g_UserPreferences.GetPref_String("Batch Process", "Drive Box", "")
    If (tempPathString <> "") And (cFile.FolderExist(tempPathString)) Then Drive1 = tempPathString
    tempPathString = g_UserPreferences.GetPref_String("Batch Process", "Input Folder", "")
    If (tempPathString <> "") And (cFile.FolderExist(tempPathString)) Then Dir1.Path = tempPathString Else Dir1.Path = Drive1
    tempPathString = g_UserPreferences.GetPref_String("Batch Process", "Output Folder", "")
    If (tempPathString <> "") And (cFile.FolderExist(tempPathString)) Then txtOutputPath.Text = tempPathString Else txtOutputPath.Text = Dir1
        
    'Populate a combo box that will display user-friendly summaries of all possible input image types
    Dim x As Long
    For x = 0 To g_ImageFormats.getNumOfInputFormats
        cmbPattern.AddItem g_ImageFormats.getInputFormatDescription(x), x
    Next x
    cmbPattern.ListIndex = 0
    
    'Populate a combo box that displays user-friendly summaries of all possible output filetypes
    For x = 0 To g_ImageFormats.getNumOfOutputFormats
        cmbOutputFormat.AddItem g_ImageFormats.getOutputFormatDescription(x), x
    Next x
    
    'Save JPEGs by default
    For x = 0 To cmbOutputFormat.ListCount
        If g_ImageFormats.getOutputFormatExtension(x) = "jpg" Then
            cmbOutputFormat.ListIndex = x
            Exit For
        End If
    Next x
    
    'By default, offer to save processed images in their original format
    optFormat(0).Value = True
    
    'Populate the combo box for file rename options
    cmbOutputOptions.AddItem "Original filenames"
    cmbOutputOptions.AddItem "Ascending numbers (1, 2, 3, etc.)"
    cmbOutputOptions.ListIndex = 0
        
    'Extract relevant icons from the resource file, and render them onto the buttons at run-time.
    cmdNext.AssignImage "ARROWRIGHT"
    cmdPrevious.AssignImage "ARROWLEFT"
    cmdAddFiles.AssignImage "ARROWDOWN"
    
    'Set the current page number to 0
    m_currentPage = 0
    
    'Mark the current image list as "not saved"
    m_ImageListSaved = False
    
    'Display appropriate help text and wizard title
    updateWizardText
    
    'Display some text manually to make sure translations are handled correctly
    txtMacro.Text = g_Language.TranslateMessage("no macro selected")
    lblExplanationFormat.Caption = g_Language.TranslateMessage("if PhotoDemon does not support an image's original format, a standard format will be used")
    lblExplanationFormat.Caption = lblExplanationFormat.Caption & vbCrLf & " " & g_Language.TranslateMessage("( specifically, JPEG at 92% quality for photographs, and lossless PNG for non-photographs )")
        
    'Hide all inactive wizard panes
    For i = 1 To picContainer.Count - 1
        picContainer(i).Visible = False
    Next i
        
    'Apply visual themes and translations
    MakeFormPretty Me
    
    'For some reason, the container picture boxes automatically acquire the cursor of children objects.
    ' Manually force those cursors to arrows to prevent this.
    For i = 0 To picContainer.Count - 1
        setArrowCursor picContainer(i)
    Next i
    
    'Cache the translations for words used in high-performance processes
    m_wordForBatchList = g_Language.TranslateMessage("batch list")
    m_wordForItem = g_Language.TranslateMessage("item")
    m_wordForItems = g_Language.TranslateMessage("items")
    
    'Finally, update the available list of images.  We must do this after translation - otherwise, the translation engine
    ' attempts to translate all the filenames and it takes forever!
    updateSourceImageList
    
End Sub

'When the source drive, directory, or file pattern is changed, the image listbox needs to be rebuilt.
Private Sub updateSourceImageList()

    lstSource.Clear

    'Parse the incoming list according to the current pattern specified by the user.  Because that pattern can be quite
    ' complex, a file listbox won't suffice - instead, we use a regular listbox and populate it ourselves.
    Dim validExtensions As String
    validExtensions = g_ImageFormats.getInputFormatExtensions(cmbPattern.ListIndex)
    
    Dim chkFile As String, chkFileExt As String
    chkFile = Dir(Dir1 & "\" & "*.*", vbNormal)
        
    Do While chkFile <> ""
        
        chkFileExt = GetExtension(chkFile)
        If Len(chkFileExt) <> 0 Then
            
            chkFileExt = "." & LCase(chkFileExt)
            
            'Compare the extension against the current list of acceptable extensions
            If validExtensions <> "*.*" Then
                If InStr(1, validExtensions, chkFileExt) Then lstSource.AddItem chkFile
            Else
                lstSource.AddItem chkFile
            End If
            
        End If
        
        'Retrieve the next file and repeat
        chkFile = Dir
        
    Loop
    
    'Enable or disable the "select all" and "select none" boxes contingent on whether images are visible in the list or not
    If lstSource.ListCount > 0 Then
        cmdSelectAll.Enabled = True
        cmdSelectNone.Enabled = True
    Else
        cmdSelectAll.Enabled = False
        cmdSelectNone.Enabled = False
    End If
    
    'Because this function forcibly clears the list box, we know that no items will be selected - so disable the "add files" button
    cmdAddFiles.Enabled = False
    
    'Quickly loop through the contents of the list box.  If any are longer than the listbox itself, display a horizontal scrollbar
    fixHorizontalListBoxScrolling lstSource
            
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Cancel = Not allowedToExit()
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub sltJP2Quality_Change()
    updateJP2ComboBox
End Sub

Private Sub sltQuality_Change()
    updateJPEGComboBox
End Sub

Private Sub lstFiles_Click()
    
    updatePreview lstFiles.List(lstFiles.ListIndex)
    m_LastPreviewSource = 1
    
    'See if any files are selected.  If they are, enable the "remove selected images" button
    Dim enableRemoveButton As Boolean
    Dim i As Long
    For i = 0 To lstFiles.ListCount - 1
        If lstFiles.Selected(i) Then
            enableRemoveButton = True
            Exit For
        End If
    Next i
    
    If enableRemoveButton Then
        If Not cmdRemove.Enabled Then cmdRemove.Enabled = True
    Else
        If cmdRemove.Enabled Then cmdRemove.Enabled = False
    End If
    
End Sub

'Allow dropping of files from the source file list box
Private Sub lstFiles_DragDrop(Source As Control, x As Single, y As Single)
    
    If Source Is lstSource Then
        Dim i As Long
        For i = 0 To lstSource.ListCount - 1
            If lstSource.Selected(i) Then addFileToBatchList Dir1 & "\" & lstSource.List(i)
        Next i
        fixHorizontalListBoxScrolling lstFiles, 16
    End If
    
End Sub

Private Sub lstFiles_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    If Source = lstSource Then lstSource.DragIcon = picDragAllow.Picture
End Sub

'This latest version of the batch wizard now supports full drag-and-drop from both Explorer and common dialogs
Private Sub lstFiles_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'Verify that the object being dragged is some sort of file or file list
    If Data.GetFormat(vbCFFiles) Then
        
        'Copy the filenames into the list box as necessary
        Dim oleFilename
        Dim tmpString As String
        
        Dim cFile As pdFSO
        Set cFile = New pdFSO
        
        For Each oleFilename In Data.Files
            
            tmpString = CStr(oleFilename)
            
            If Len(tmpString) <> 0 Then
                If cFile.FileExist(tmpString) Then addFileToBatchList tmpString
            End If
            
        Next oleFilename
        
        fixHorizontalListBoxScrolling lstFiles, 16
        
    End If
    
End Sub

Private Sub lstFiles_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    
    'Check to make sure the type of OLE object is files
    If Data.GetFormat(vbCFFiles) Then
        'Inform the source (Explorer, in this case) that the files will be treated as "copied"
        Effect = vbDropEffectCopy And Effect
    Else
        'If it's not files, don't allow a drop
        Effect = vbDropEffectNone
    End If
    
End Sub

Private Sub lstSource_Click()
    
    'If at least one file is selected, enable the "Add to batch process" button
    Dim somethingSelected As Boolean
    
    Dim i As Long
    For i = 0 To lstSource.ListCount - 1
        If lstSource.Selected(i) Then somethingSelected = True
    Next i
    cmdAddFiles.Enabled = somethingSelected
    
    'Redraw the preview
    updatePreview Dir1.Path & "\" & lstSource.List(lstSource.ListIndex)
    m_LastPreviewSource = 0
    
End Sub

Private Sub fixHorizontalListBoxScrolling(ByRef srcListBox As ListBox, Optional ByVal lenModifier As Long = 0)
    
    Dim i As Long, lenText As Long, maxWidth As Long
    maxWidth = Me.textWidth(srcListBox.List(0) & "     ")
    For i = 0 To srcListBox.ListCount - 1
        lenText = Me.textWidth(srcListBox.List(i) & "     ")
        If lenText > maxWidth Then maxWidth = lenText
    Next i
    
    SendMessage srcListBox.hWnd, LB_SETHORIZONTALEXTENT, maxWidth + lenModifier, 0
    LockWindowUpdate 0
    
End Sub

'Update the active image preview in the top-right
Private Sub updatePreview(ByVal srcImagePath As String)
    
    'Only redraw the preview if it doesn't match the last image we previewed
    If CBool(chkEnablePreview) And (StrComp(m_CurImagePreview, srcImagePath, vbTextCompare) <> 0) Then
    
        'Use PD's central load function to load a copy of the requested image
        Dim tmpImagePath(0) As String
        tmpImagePath(0) = srcImagePath
        
        Dim tmpImage As pdImage
        Set tmpImage = New pdImage
        
        Dim tmpDIB As pdDIB
        Set tmpDIB = New pdDIB
        
        LoadFileAsNewImage tmpImagePath, False, "", "", False, tmpImage, tmpDIB, 0, True
        
        'Try to ascertain if the image loaded correctly
        Dim loadFailed As Boolean
        loadFailed = False
        
        If tmpDIB Is Nothing Then
            loadFailed = True
        Else
            If (tmpDIB.getDIBWidth = 0) Or (tmpDIB.getDIBHeight = 0) Then loadFailed = True
        End If
        
        If Not tmpImage.loadedSuccessfully Then loadFailed = True
        
        'If the image load failed, display a placeholder message; otherwise, render the image to the picture box
        If loadFailed Then
            picPreview.Picture = LoadPicture("")
            Dim strToPrint As String
            strToPrint = g_Language.TranslateMessage("Preview not available")
            picPreview.CurrentX = (picPreview.ScaleWidth - picPreview.textWidth(strToPrint)) \ 2
            picPreview.CurrentY = (picPreview.ScaleHeight - picPreview.textHeight(strToPrint)) \ 2
            picPreview.Print strToPrint
        Else
            tmpDIB.renderToPictureBox picPreview
        End If
        
        'Remember the name of the current preview; this saves us having to reload the preview any more than
        ' is absolutely necessary
        m_CurImagePreview = srcImagePath
    
    End If
    
End Sub

'Add a file to a batch list.  This separate routine is used so that duplicates and invalid files can be removed prior to addition.
Private Sub addFileToBatchList(ByVal srcFile As String, Optional ByVal suppressDuplicatesCheck As Boolean = False)
    
    LockWindowUpdate lstFiles.hWnd
    
    Dim novelAddition As Boolean
    novelAddition = True
    
    If Not suppressDuplicatesCheck Then
        Dim x As Long
        For x = 0 To lstFiles.ListCount - 1
            If StrComp(lstFiles.List(x), srcFile, vbTextCompare) = 0 Then
                novelAddition = False
                Exit For
            End If
        Next x
    End If
    
    'Only add this file to the list if a) it doesn't already appear there, and b) the file actually exists (important when loading
    ' a previously saved batch list from file)
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    If novelAddition Then
    
        If cFile.FileExist(srcFile) Then
            lstFiles.AddItem srcFile
            updateBatchListCount
        End If
        
    End If
    
    'Enable the "remove all images" button if at least one image exists in the processing list
    If lstFiles.ListCount > 0 Then
        If Not cmdRemoveAll.Enabled Then cmdRemoveAll.Enabled = True
        If Not cmdSaveList.Enabled Then cmdSaveList.Enabled = True
        'If Not cmdNext.Enabled Then cmdNext.Enabled = True
    End If
    
    'Note that the current list has NOT been saved
    m_ImageListSaved = False
    
End Sub

Private Sub updateBatchListCount()
    lblTargetFiles.Caption = m_wordForBatchList
    Select Case lstFiles.ListCount
    
        Case 0
            lblTargetFiles.Caption = lblTargetFiles.Caption & ":"
        Case 1
            lblTargetFiles.Caption = lblTargetFiles.Caption & " (" & lstFiles.ListCount & " " & m_wordForItem & "):"
        Case Else
            lblTargetFiles.Caption = lblTargetFiles.Caption & " (" & lstFiles.ListCount & " " & m_wordForItems & "):"
            
    End Select
    
End Sub

Private Sub lstSource_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    'Allow drag operations via the RIGHT mouse button
    If Button = vbRightButton Then
        lstSource.Drag vbBeginDrag
        lstSource.DragIcon = picDragDisallow.Picture
    End If
    
End Sub

Private Sub lstSource_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    lstSource.DragIcon = LoadPicture("")
End Sub

Private Sub optCase_Click(Index As Integer)
    chkRenameCase.Value = vbChecked
End Sub


'Used to keep the "image quality" text box, scroll bar, and combo box in sync
Private Sub updateJPEGComboBox()
    
    Select Case sltQuality.Value
        
        Case 40
            If cmbJPEGSaveQuality.ListIndex <> 4 Then cmbJPEGSaveQuality.ListIndex = 4
                            
        Case 65
            If cmbJPEGSaveQuality.ListIndex <> 3 Then cmbJPEGSaveQuality.ListIndex = 3
                
        Case 80
            If cmbJPEGSaveQuality.ListIndex <> 2 Then cmbJPEGSaveQuality.ListIndex = 2
                
        Case 92
            If cmbJPEGSaveQuality.ListIndex <> 1 Then cmbJPEGSaveQuality.ListIndex = 1
                
        Case 99
            If cmbJPEGSaveQuality.ListIndex <> 0 Then cmbJPEGSaveQuality.ListIndex = 0
                
        Case Else
            If cmbJPEGSaveQuality.ListIndex <> 5 Then cmbJPEGSaveQuality.ListIndex = 5
                
    End Select
    
End Sub

'Used to keep the JPEG-2000 "compression ratio" text box, scroll bar, and combo box in sync
Private Sub updateJP2ComboBox()
    
    Select Case sltJP2Quality.Value
        
        Case 1
            If cmbJP2SaveQuality.ListIndex <> 0 Then cmbJP2SaveQuality.ListIndex = 0
                
        Case 16
            If cmbJP2SaveQuality.ListIndex <> 1 Then cmbJP2SaveQuality.ListIndex = 1
                
        Case 32
            If cmbJP2SaveQuality.ListIndex <> 2 Then cmbJP2SaveQuality.ListIndex = 2
                
        Case 64
            If cmbJP2SaveQuality.ListIndex <> 3 Then cmbJP2SaveQuality.ListIndex = 3
                
        Case 256
            If cmbJP2SaveQuality.ListIndex <> 4 Then cmbJP2SaveQuality.ListIndex = 4
                
        Case Else
            If cmbJP2SaveQuality.ListIndex <> 5 Then cmbJP2SaveQuality.ListIndex = 5
                
    End Select
    
End Sub

'When the user presses "Start Conversion", this routine is triggered.
Private Sub prepareForBatchConversion()

    batchConvertMessage g_Language.TranslateMessage("Preparing batch processing engine...")
    
    'Display the progress panel
    Dim i As Long
    
    picContainer(picContainer.Count - 1).Visible = True
    
    For i = 0 To picContainer.Count - 2
        picContainer(i).Visible = False
    Next i
    
    'Hide the back/forward buttons
    cmdPrevious.Visible = False
    cmdNext.Visible = False
    
    'Before doing anything, save relevant folder locations to the preferences file
    g_UserPreferences.SetPref_String "Batch Process", "Drive Box", Drive1
    g_UserPreferences.SetPref_String "Batch Process", "Input Folder", Dir1.Path

    'Let the rest of the program know that batch processing has begun
    MacroStatus = MacroBATCH
    
    Dim curBatchFile As Long
    Dim tmpFilename As String, tmpFileExtension As String
    
    Dim totalNumOfFiles As Long
    totalNumOfFiles = lstFiles.ListCount
    
    'LoadFileAsNewImage requires an array.  This array will be used to send it individual filenames
    Dim sFile(0) As String
    
    'Prepare the folder that will receive the processed images
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    Dim outputPath As String
    outputPath = cFile.EnforcePathSlash(txtOutputPath)
    If Not cFile.FolderExist(outputPath) Then cFile.CreateFolder outputPath, True
    
    'Prepare the progress bar, which will keep the user updated on our progress.
    Set sysProgBar = New cProgressBarOfficial
    sysProgBar.CreateProgressBar picBatchProgress.hWnd, 0, 0, picBatchProgress.ScaleWidth, picBatchProgress.ScaleHeight, True, True, True, True
    sysProgBar.Max = totalNumOfFiles
    sysProgBar.Min = 0
    sysProgBar.Value = 0
    sysProgBar.Refresh
    
    'Let's also give the user an estimate of how long this is going to take.  We estimate time by determining an
    ' approximate "time-per-image" value, then multiplying that by the number of images remaining.  The progress bar
    ' will display this, automatically updated, as each image is completed.
    Dim timeStarted As Double, timeElapsed As Double, timeRemaining As Double, timePerFile As Double
    Dim numFilesProcessed As Long, numFilesRemaining As Long
    Dim minutesRemaining As Long, secondsRemaining As Long
    Dim timeMsg As String
    Dim lastTimeCalculation As Long
    lastTimeCalculation = &H7FFFFFFF
    
    timeStarted = GetTickCount
    timeMsg = ""
    
    'This is where the fun begins.  Loop through every file in the list, and process them one-by-one using the options requested
    ' by the user.
    For curBatchFile = 0 To totalNumOfFiles
    
        'Pause for keypresses - this allows the user to press "Escape" to cancel the operation
        DoEvents
        If MacroStatus = MacroCANCEL Then GoTo MacroCanceled
    
        tmpFilename = lstFiles.List(curBatchFile)
        
        'Give the user a progress update
        MacroMessage = g_Language.TranslateMessage("Processing image # %1 of %2. %3", (curBatchFile + 1), totalNumOfFiles, timeMsg)
        batchConvertMessage MacroMessage
        sysProgBar.Value = curBatchFile
        sysProgBar.Refresh
        
        'As a failsafe, check to make sure the current input file exists before attempting to load it
        If cFile.FileExist(tmpFilename) Then
            
            sFile(0) = tmpFilename
            
            'Check to see if the image file is a multipage file
            Dim howManyPages As Long
            
            howManyPages = isMultiImage(tmpFilename)
            
            'TODO: integrate this with future support for exporting multipage files.  At present, to avoid complications,
            ' PD will only load the first page/frame of a multipage file during conversion.  (This is why the code below
            ' looks so goofy.)
            
            'Check the user's preference regarding multipage images.  If they have specifically requested that we load
            ' only the first page of the image, ignore any subsequent pages.
            If howManyPages > 0 Then
                howManyPages = 1
            '    If g_UserPreferences.GetPref_Long("Loading", "Multipage Image Prompt", 0) = 1 Then howManyPages = 1
            Else
                howManyPages = 1
            End If
            
            'Now, loop through each page or frame (if applicable), load the image, and process it.
            Dim curPage As Long
            For curPage = 0 To howManyPages - 1
            
                'Load the current image
                LoadFileAsNewImage sFile, False, , , , , , curPage
                
                'Make sure the image loaded correctly
                If Not (pdImages(g_CurrentImage) Is Nothing) Then
                If pdImages(g_CurrentImage).loadedSuccessfully Then
                    
                    'With the image loaded, it is time to apply any requested photo editing actions.
                    If optActions(1) Then
                    
                        'If the user has requested automatic lighting fixes, apply it now
                        If CBool(chkActions(0)) Then
                            Process "White balance", , buildParams("0.1"), UNDO_LAYER
                        End If
                    
                        'If the user has requested an image resize, apply it now
                        If CBool(chkActions(1)) Then
                            Process "Resize image", , buildParams(ucResize.imgWidth, ucResize.imgHeight, RESIZE_LANCZOS, cmbResizeFit.ListIndex, RGB(255, 255, 255), ucResize.unitOfMeasurement, ucResize.imgDPIAsPPI, PD_AT_WHOLEIMAGE)
                        End If
                        
                        'If the user has requested a macro, play it now
                        If CBool(chkActions(2)) Then PlayMacroFromFile txtMacro
                        
                    End If
                
                    'With the macro complete, prepare the file for saving
                    tmpFilename = lstFiles.List(curBatchFile)
                    StripOffExtension tmpFilename
                    StripFilename tmpFilename
                
                    'Build a full file path using the options the user specified
                    If cmbOutputOptions.ListIndex = 0 Then
                        If CBool(chkRenamePrefix) Then tmpFilename = txtAppendFront & tmpFilename
                        If CBool(chkRenameSuffix) Then tmpFilename = tmpFilename & txtAppendBack
                    Else
                        tmpFilename = curBatchFile + 1
                        If CBool(chkRenamePrefix) Then tmpFilename = txtAppendFront & tmpFilename
                        If CBool(chkRenameSuffix) Then tmpFilename = tmpFilename & txtAppendBack
                    End If
                    
                    'Add the page number if necessary
                    If curPage > 0 Then tmpFilename = tmpFilename & " (" & curPage & ")"
                    
                    'If requested, remove any specified text from the filename
                    If CBool(chkRenameRemove) And (Len(txtRenameRemove) <> 0) Then
                    
                        'Use case-sensitive or case-insensitive matching as requested
                        If CBool(chkRenameCaseSensitive) Then
                            If InStr(1, tmpFilename, txtRenameRemove, vbBinaryCompare) Then
                                tmpFilename = Replace(tmpFilename, txtRenameRemove, "", , , vbBinaryCompare)
                            End If
                        Else
                            If InStr(1, tmpFilename, txtRenameRemove, vbTextCompare) Then
                                tmpFilename = Replace(tmpFilename, txtRenameRemove, "", , , vbTextCompare)
                            End If
                        End If
                        
                    End If
                    
                    'Replace spaces with underscores if requested
                    If CBool(chkRenameSpaces) Then
                        If InStr(1, tmpFilename, " ") Then
                            tmpFilename = Replace(tmpFilename, " ", "_")
                        End If
                    End If
                    
                    'Change the full filename's case if requested
                    If CBool(chkRenameCase) Then
                    
                        If optCase(0) Then
                            tmpFilename = LCase(tmpFilename)
                        Else
                            tmpFilename = UCase(tmpFilename)
                        End If
                    
                    End If
                    
                    'Attach a proper image format file extension and save format ID number based off the user's
                    ' requested output format
                    
                    'Possibility 1: use original file format
                    If optFormat(0) Then
                        
                        m_FormatParams = ""
                        
                        'See if this image's file format is supported by the export engine
                        If g_ImageFormats.getIndexOfOutputFIF(pdImages(g_CurrentImage).currentFileFormat) = -1 Then
                            
                            'If it isn't, save as JPEG or PNG contingent on color depth
                            
                            '24bpp images default to JPEG
                            If pdImages(g_CurrentImage).getCompositeImageColorDepth = 24 Then
                                tmpFileExtension = g_ImageFormats.getExtensionFromFIF(FIF_JPEG)
                                pdImages(g_CurrentImage).currentFileFormat = FIF_JPEG
                            
                            '32bpp images default to PNG
                            Else
                                tmpFileExtension = g_ImageFormats.getExtensionFromFIF(FIF_JPEG)
                                pdImages(g_CurrentImage).currentFileFormat = FIF_PNG
                            End If
                            
                        Else
                            
                            'This format IS supported, so use the default extension
                            tmpFileExtension = g_ImageFormats.getExtensionFromFIF(pdImages(g_CurrentImage).currentFileFormat)
                        
                        End If
                        
                    'Possibility 2: force all images to a single file format
                    Else
                        tmpFileExtension = g_ImageFormats.getOutputFormatExtension(cmbOutputFormat.ListIndex)
                        pdImages(g_CurrentImage).currentFileFormat = g_ImageFormats.getOutputFIF(cmbOutputFormat.ListIndex)
                    End If
                    
                    'If the user has requested lower- or upper-case, we now need to convert the extension as well
                    If CBool(chkRenameCase) Then
                    
                        If optCase(0) Then
                            tmpFileExtension = LCase(tmpFileExtension)
                        Else
                            tmpFileExtension = UCase(tmpFileExtension)
                        End If
                    
                    End If
                    
                    'Because removing specified text from filenames may lead to files with the same name, call the incrementFilename
                    ' function to find a unique filename of the "filename (n+1)" variety if necessary.  This will also prepend the
                    ' drive and directory structure.
                    tmpFilename = outputPath & incrementFilename(outputPath, tmpFilename, tmpFileExtension) & "." & tmpFileExtension
                                    
                    'Request a save from the PhotoDemon_SaveImage method, and pass it a specialized string containing
                    ' any extra information for the requested format (JPEG quality, etc)
                    If Len(m_FormatParams) <> 0 Then
                        PhotoDemon_SaveImage pdImages(CLng(g_CurrentImage)), tmpFilename, CLng(g_CurrentImage), False, m_FormatParams
                    Else
                        PhotoDemon_SaveImage pdImages(CLng(g_CurrentImage)), tmpFilename, CLng(g_CurrentImage), False
                    End If
                
                End If
                End If
            
                'Unload the active form
                FullPDImageUnload g_CurrentImage
                
            Next curPage
            
            'If a good number of images have been processed, start estimating the amount of time remaining
            If (curBatchFile > 10) Then
            
                timeElapsed = GetTickCount - timeStarted
                numFilesProcessed = curBatchFile + 1
                numFilesRemaining = totalNumOfFiles - numFilesProcessed
                timePerFile = timeElapsed / numFilesProcessed
                timeRemaining = timePerFile * numFilesRemaining
                
                'Convert timeRemaining to seconds (it is currently in milliseconds)
                timeRemaining = timeRemaining / 1000
                
                minutesRemaining = Int(timeRemaining / 60)
                secondsRemaining = Int(timeRemaining) Mod 60
                
                'Only update the time remaining message if it is LESS than our previous result, the seconds are a multiple
                ' of 5, or there is 0 minutes remaining (in which case we can display an exact seconds estimate).
                If (timeRemaining < lastTimeCalculation) And ((secondsRemaining Mod 5 = 0) Or (minutesRemaining = 0)) Then
                
                    lastTimeCalculation = timeRemaining
                
                    'This lets us format our time nicely (e.g. "minute" vs "minutes")
                    Select Case minutesRemaining
                        'No minutes remaining - only seconds
                        Case 0
                            timeMsg = g_Language.TranslateMessage("Estimated time remaining") & ": "
                        Case 1
                            timeMsg = g_Language.TranslateMessage("Estimated time remaining") & ": " & minutesRemaining
                            timeMsg = timeMsg & " " & g_Language.TranslateMessage("minute") & " "
                        Case Else
                            timeMsg = g_Language.TranslateMessage("Estimated time remaining") & ": " & minutesRemaining
                            timeMsg = timeMsg & " " & g_Language.TranslateMessage("minutes") & " "
                    End Select
                    
                    Select Case secondsRemaining
                        Case 1
                            timeMsg = timeMsg & "1 " & g_Language.TranslateMessage("second")
                        Case Else
                            timeMsg = timeMsg & secondsRemaining & " " & g_Language.TranslateMessage("seconds")
                    End Select
                
                End If

            ElseIf (curBatchFile > 20) And (totalNumOfFiles > 50) Then
                timeMsg = g_Language.TranslateMessage("Estimating time remaining") & "..."
            End If
        
        End If
                
    'Carry on
    Next curBatchFile
    
    MacroStatus = MacroSTOP
    
    Screen.MousePointer = vbDefault
    
    'Change the "Cancel" button to "Exit"
    cmdCancel.Caption = g_Language.TranslateMessage("Exit")
    
    'Max out the progess bar and display a success message
    sysProgBar.Value = sysProgBar.Max
    sysProgBar.Refresh
    batchConvertMessage g_Language.TranslateMessage("%1 files were successfully processed!", totalNumOfFiles)
    
    'Finally, there is no longer any need for the user to save their batch list, as the batch process is complete.
    m_ImageListSaved = True
    
    Exit Sub
    
MacroCanceled:

    MacroStatus = MacroSTOP
    
    Screen.MousePointer = vbDefault
    
    'Reset the progress bar
    sysProgBar.Value = 0
    sysProgBar.Refresh
    
    Dim cancelMsg As String
    cancelMsg = g_Language.TranslateMessage("Batch conversion canceled.") & " " & curBatchFile & " "
    
    'Properly display "image" or "images" depending on how many files were processed
    If curBatchFile <> 1 Then
        cancelMsg = cancelMsg & g_Language.TranslateMessage("images were")
    Else
        cancelMsg = cancelMsg & g_Language.TranslateMessage("image was")
    End If
    
    cancelMsg = cancelMsg & " "
    cancelMsg = cancelMsg & g_Language.TranslateMessage("processed before cancelation. Last processed image was ""%1"".", lstFiles.List(curBatchFile))
    
    batchConvertMessage cancelMsg
    
    'Change the "Cancel" button to "Exit"
    cmdCancel.Caption = g_Language.TranslateMessage("Exit")
    
    m_ImageListSaved = True
    
End Sub

'Display a progress update to the user
Private Sub batchConvertMessage(ByVal newMessage As String)
    lblBatchProgress.Caption = newMessage
    lblBatchProgress.Refresh
End Sub
