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
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "&Previous"
      Enabled         =   0   'False
      Height          =   615
      Left            =   10080
      TabIndex        =   7
      Top             =   8355
      Width           =   1725
   End
   Begin VB.PictureBox picDragAllow 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   720
      Picture         =   "VBP_FormBatchWizard.frx":0000
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   6
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
      Picture         =   "VBP_FormBatchWizard.frx":09F6
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   5
      Top             =   7680
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Height          =   615
      Left            =   13860
      TabIndex        =   1
      Top             =   8355
      Width           =   1365
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   615
      Left            =   11880
      TabIndex        =   0
      Top             =   8355
      Width           =   1725
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7455
      Index           =   1
      Left            =   3480
      ScaleHeight     =   497
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   785
      TabIndex        =   9
      Top             =   720
      Width           =   11775
      Begin VB.CommandButton cmdSelectMacro 
         Caption         =   "Select macro file..."
         Height          =   525
         Left            =   8400
         TabIndex        =   34
         Top             =   1440
         Width           =   3015
      End
      Begin VB.TextBox txtMacro 
         Enabled         =   0   'False
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
         Height          =   315
         Left            =   720
         TabIndex        =   33
         Text            =   "no macro selected"
         Top             =   1560
         Width           =   7455
      End
      Begin PhotoDemon.smartOptionButton optActions 
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   30
         Top             =   120
         Width           =   4110
         _ExtentX        =   7250
         _ExtentY        =   635
         Caption         =   "do not apply any photo editing actions"
         Value           =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin PhotoDemon.smartOptionButton optActions 
         Height          =   360
         Index           =   1
         Left            =   120
         TabIndex        =   32
         Top             =   1080
         Width           =   4140
         _ExtentX        =   7303
         _ExtentY        =   635
         Caption         =   "apply a recorded macro to the images"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         Left            =   720
         TabIndex        =   31
         Top             =   540
         Width           =   6615
      End
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7500
      Index           =   0
      Left            =   3480
      ScaleHeight     =   500
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   785
      TabIndex        =   8
      Top             =   720
      Width           =   11775
      Begin VB.ComboBox cmbPattern 
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   510
         Width           =   3645
      End
      Begin VB.DriveListBox Drive1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   240
         TabIndex        =   24
         Top             =   1080
         Width           =   3645
      End
      Begin VB.DirListBox Dir1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   2565
         Left            =   240
         TabIndex        =   23
         Top             =   1440
         Width           =   3615
      End
      Begin VB.ListBox lstFiles 
         ForeColor       =   &H00800000&
         Height          =   2400
         Left            =   240
         MultiSelect     =   2  'Extended
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   22
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
         TabIndex        =   20
         Top             =   1080
         Width           =   3495
      End
      Begin VB.CommandButton cmdAddFiles 
         Caption         =   "Add selected image(s) to batch list"
         Enabled         =   0   'False
         Height          =   615
         Left            =   4200
         TabIndex        =   19
         Top             =   4110
         Width           =   3615
      End
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "Select all"
         Enabled         =   0   'False
         Height          =   615
         Left            =   4200
         TabIndex        =   18
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdSelectNone 
         Caption         =   "Select none"
         Enabled         =   0   'False
         Height          =   615
         Left            =   6120
         TabIndex        =   17
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdUseCD 
         Caption         =   "Add images using ""Open Image"" dialog..."
         Height          =   615
         Left            =   8145
         TabIndex        =   16
         Top             =   360
         Width           =   3525
      End
      Begin VB.CommandButton cmdLoadList 
         Caption         =   "Load list..."
         Height          =   615
         Left            =   8160
         TabIndex        =   15
         Top             =   6600
         Width           =   1695
      End
      Begin VB.CommandButton cmdSaveList 
         Caption         =   "Save current list..."
         Enabled         =   0   'False
         Height          =   615
         Left            =   9960
         TabIndex        =   14
         Top             =   6600
         Width           =   1695
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove selected image(s)"
         Enabled         =   0   'False
         Height          =   615
         Left            =   8160
         TabIndex        =   13
         Top             =   5400
         Width           =   1695
      End
      Begin VB.CommandButton cmdRemoveAll 
         Caption         =   "Remove all images"
         Enabled         =   0   'False
         Height          =   615
         Left            =   9960
         TabIndex        =   12
         Top             =   5400
         Width           =   1695
      End
      Begin VB.ListBox lstSource 
         ForeColor       =   &H00400000&
         Height          =   2940
         IntegralHeight  =   0   'False
         Left            =   4200
         MultiSelect     =   2  'Extended
         TabIndex        =   11
         Top             =   1080
         Width           =   3615
      End
      Begin PhotoDemon.smartCheckBox chkEnablePreview 
         Height          =   480
         Left            =   8160
         TabIndex        =   21
         Top             =   3600
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   847
         Caption         =   "show image previews"
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
      Index           =   3
      Left            =   3480
      ScaleHeight     =   497
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   785
      TabIndex        =   35
      Top             =   720
      Width           =   11775
      Begin PhotoDemon.smartOptionButton optCase 
         Height          =   330
         Index           =   0
         Left            =   840
         TabIndex        =   53
         Top             =   5640
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   582
         Caption         =   "lowercase"
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
      Begin VB.TextBox txtRenameRemove 
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
         TabIndex        =   51
         Top             =   4560
         Width           =   4260
      End
      Begin VB.TextBox txtAppendBack 
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
         Left            =   6120
         TabIndex        =   50
         Top             =   3480
         Width           =   4260
      End
      Begin VB.TextBox txtAppendFront 
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
         TabIndex        =   49
         Text            =   "NEW_"
         Top             =   3480
         Width           =   4260
      End
      Begin PhotoDemon.smartCheckBox chkRenamePrefix 
         Height          =   480
         Left            =   480
         TabIndex        =   46
         Top             =   3000
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   847
         Caption         =   "add a prefix to each filename:"
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
         TabIndex        =   45
         Top             =   1800
         Width           =   7455
      End
      Begin VB.CommandButton cmdSelectOutputPath 
         Caption         =   "Select destination folder..."
         Height          =   525
         Left            =   8280
         TabIndex        =   42
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox txtOutputPath 
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
         Height          =   315
         Left            =   480
         TabIndex        =   41
         Text            =   "C:\"
         Top             =   600
         Width           =   7455
      End
      Begin PhotoDemon.smartCheckBox chkRenameSuffix 
         Height          =   480
         Left            =   5760
         TabIndex        =   47
         Top             =   3000
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   847
         Caption         =   "add a suffix to each filename:"
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
      Begin PhotoDemon.smartCheckBox chkRenameRemove 
         Height          =   480
         Left            =   480
         TabIndex        =   48
         Top             =   4080
         Width           =   5205
         _ExtentX        =   9181
         _ExtentY        =   847
         Caption         =   "remove the following text (if found) from each filename:"
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
      Begin PhotoDemon.smartCheckBox chkRenameCase 
         Height          =   480
         Left            =   480
         TabIndex        =   52
         Top             =   5160
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   847
         Caption         =   "force each filename, including extension, to the following case:"
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
      Begin PhotoDemon.smartOptionButton optCase 
         Height          =   330
         Index           =   1
         Left            =   3240
         TabIndex        =   54
         Top             =   5640
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         Caption         =   "UPPERCASE"
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
      Begin PhotoDemon.smartCheckBox chkRenameSpaces 
         Height          =   480
         Left            =   480
         TabIndex        =   55
         Top             =   6240
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   847
         Caption         =   "replace spaces in filenames with underscores"
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
      Begin VB.Label lblDstFilename 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "the base name of processed files will be:"
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
         TabIndex        =   44
         Top             =   1320
         Width           =   4305
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
         TabIndex        =   43
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
         TabIndex        =   40
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
      ScaleHeight     =   7455
      ScaleWidth      =   11775
      TabIndex        =   10
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
         TabIndex        =   39
         Top             =   1920
         Width           =   6735
      End
      Begin PhotoDemon.smartOptionButton optFormat 
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   120
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   635
         Caption         =   "keep images in their original format"
         Value           =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin PhotoDemon.smartOptionButton optFormat 
         Height          =   360
         Index           =   1
         Left            =   120
         TabIndex        =   37
         Top             =   1320
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   635
         Caption         =   "convert images to a single format"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         TabIndex        =   38
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
      Caption         =   "You can add files to the batch process list in several ways: "
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
      Left            =   240
      TabIndex        =   4
      Top             =   780
      Width           =   2895
   End
   Begin VB.Label lblWizardTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Step 1: prepare the batch list (the list of images to be processed)"
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
      TabIndex        =   3
      Top             =   120
      Width           =   7980
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
      TabIndex        =   2
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
'Copyright ©2007-2013 by Tanner Helland
'Created: 3/Nov/07
'Last updated: 21/February/13
'Last update: over two weeks, I rewrote the entire dialog from scratch.  Now it's a batch *wizard*.  Ooooo...
'
'PhotoDemon's batch process wizard is one of its most unique - and in my opinion, most impressive - features.  It integrates
' tightly with the macro recording feature to allow any combination of actions to be applied to any set of images.
'
'The process is broken into three steps.
'
'1) Build the batch list, e.g. the list of files to be processed.  This is by far the most complicated section of the wizard.
'    I have revisited the design of this page many times, and I think the current incarnation is pretty damn good.  It exposes
'    a lot of functionality without being overwhelming, and the user has many tools at their disposal to build an ideal list
'    of images from any number of source directories.  (Many batch tools limit you to just one source folder, which I hate.)
'
'2) Select the operation to apply to the images.  This is really two pages in one, condensed for simplicity.  First, the user
'    must select an output file format.  There are three choices: retain original format (e.g. "rename only", which allows the
'    user to use the tool as a batch renamer), pick optimal format for web (which will intermix JPEG and PNG intelligently),
'    or the user can pick their own format.  Some mechanism must be provided for the user to adjust certain settings, such as
'    JPEG quality.  I need to investigate how to handle that.  Additionally, the user can select what image editing options
'    to apply to each image.  A list of presets would be ideal, or they can optionally supply a recorded macro of any
'    complexity.  This feature places PD among the most comprehensive batch processing tools available anywhere.
'
'3) Choose where the files will go and what they will be named.  This includes a number of renaming options, which is a big
'    step up from the original batch process tool in earlier versions.  I am open to suggestions for other renaming features,
'    but at present I think the selection is sufficiently comprehensive.
'
'Due to the complexity of this tool, there may be some odd combinations of things that don't work quite right - I'm hoping
' others can help test and provide feedback to ensure that everything runs smoothly.
'
'***************************************************************************

Option Explicit

'API to add a horizontal scroll bar as necessary - see http://support.microsoft.com/default.aspx?scid=kb%3Ben-us%3B192184
Private Declare Function SendMessageA Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const LB_SETHORIZONTALEXTENT = &H194

'API to add items to a listbox without a refresh occurring
Private Declare Function SendMessageByString Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Const WM_SETREDRAW = &HB
Private Const LB_ADDSTRING = &H401
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Long) As Long

'Used to render images onto command buttons at run-time
Dim cImgCtl As clsControlImage

Dim m_currentPage As Long

Private Sub chkEnablePreview_Click()
    
    picPreview.Picture = LoadPicture("")
    
    'If the user is disabling previews, clear the picture box and display a notice
    If Not CBool(chkEnablePreview) Then
        Dim strToPrint As String
        strToPrint = g_Language.TranslateMessage("Previews disabled")
        picPreview.CurrentX = (picPreview.ScaleWidth - picPreview.TextWidth(strToPrint)) \ 2
        picPreview.CurrentY = (picPreview.ScaleHeight - picPreview.TextHeight(strToPrint)) \ 2
        picPreview.Print strToPrint
    'If the user is enabling previews, try to display the last item the user selected in the SOURCE list box
    Else
        If lstSource.Selected(lstSource.ListIndex) Then updatePreview Dir1 & "\" & lstSource.List(lstSource.ListIndex)
    End If
    
End Sub

Private Sub cmbPattern_Click()
    updateSourceImageList
End Sub

Private Sub cmdAddFiles_Click()
    Screen.MousePointer = vbHourglass
    Dim x As Long
    For x = 0 To lstSource.ListCount - 1
        If lstSource.Selected(x) Then addFileToBatchList Dir1.Path & "\" & lstSource.List(x)
    Next x
    fixHorizontalListBoxScrolling lstFiles, 16
    Screen.MousePointer = vbDefault
    makeFormPretty Me
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

'Load a list of images (previously saved from within PhotoDemon) to the batch list
Private Sub cmdLoadList_Click()
    
    Dim sFile As String
    
    'Get the last "open/save image list" path from the INI file
    Dim tempPathString As String
    tempPathString = g_UserPreferences.GetPreference_String("Batch Preferences", "ListFolder", "")
    
    Dim CC As cCommonDialog
    Set CC = New cCommonDialog
    
    If CC.VBGetOpenFileName(sFile, , True, False, False, True, "Batch Image List (.pdl)|*.pdl|All files|*.*", 0, tempPathString, "Load a list of images", ".pdl", FormBatchWizard.hWnd, OFN_HIDEREADONLY) Then
        
        'Save this new directory as the default path for future usage
        Dim listPath As String
        listPath = sFile
        StripDirectory listPath
        g_UserPreferences.SetPreference_String "Batch Preferences", "ListFolder", listPath
        
        Dim fileNum As Integer
        fileNum = FreeFile
    
        Open sFile For Input As #fileNum
            Dim tmpLine As String
            Input #fileNum, tmpLine
            If tmpLine <> ("<" & PROGRAMNAME & " BATCH CONVERSION LIST>") Then
                pdMsgBox "This is not a valid list of images. Please try a different file.", vbExclamation + vbApplicationModal + vbOKOnly, "Invalid list file"
                Exit Sub
            End If
            
            'Check to see if the user wants to append this list to the current list,
            ' or if they want to load just the list data
            If lstFiles.ListCount > 0 Then
                Dim msgReturn As VbMsgBoxResult
                'NOTE TO TANNER FOR v5.4: a translation may already exist for this text
                msgReturn = pdMsgBox("You have already created a list of images for processing.  The list of images inside this file will be appended to the bottom of your current list.", vbOKCancel + vbApplicationModal + vbInformation, "Batch process notification")
                If msgReturn = vbCancel Then Exit Sub
            End If
            Screen.MousePointer = vbHourglass
            'Now that everything is in place, load the entries from the file
            Input #fileNum, tmpLine
            Dim numOfEntries As Long
            numOfEntries = CLng(tmpLine)
            
            Dim suppressDuplicatesCheck As Boolean
            If numOfEntries > 100 Then suppressDuplicatesCheck = True
            Dim i As Long
            For i = 0 To numOfEntries - 1
                Input #fileNum, tmpLine
                addFileToBatchList tmpLine, suppressDuplicatesCheck
            Next i
            fixHorizontalListBoxScrolling lstFiles, 16
            lstFiles.Refresh
            Screen.MousePointer = vbDefault
            makeFormPretty Me
        Close #fileNum
    End If
    
End Sub

Private Sub cmdNext_Click()
    changeBatchPage True
End Sub

Private Sub cmdPrevious_Click()
    changeBatchPage False
End Sub

'This function is used to advance (or retreat) the active wizard panel
Private Sub changeBatchPage(ByVal moveForward As Boolean)

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
    If m_currentPage = picContainer.Count - 1 Then
        cmdNext.Caption = g_Language.TranslateMessage("Start processing!")
    Else
        If cmdNext.Caption <> g_Language.TranslateMessage("Next") Then cmdNext.Caption = g_Language.TranslateMessage("Next")
    End If
    
    'Finally, update all the label captions that change according to the active panel
    updateWizardText
    
End Sub

'Used to display unique text for each page of the wizard.  The value of m_currentPage is used to determine what text to display.
Private Sub updateWizardText()

    Select Case m_currentPage
        
        'Step 1: add images to list
        Case 0
        
            lblWizardTitle.Caption = g_Language.TranslateMessage("Step 1: prepare the batch list (the list of images to be processed)")
            
            lblExplanation(0).Caption = g_Language.TranslateMessage("You can add files to the batch list in several ways:")
            lblExplanation(0).Caption = lblExplanation(0).Caption & vbCrLf & vbCrLf & "  " & g_Language.TranslateMessage("1) The folder and file lists at the top of this page.  Use the ""Add selected image(s) to batch list"" button to move images to the batch list, or use the right mouse button to drag-and-drop one or more items.")
            lblExplanation(0).Caption = lblExplanation(0).Caption & vbCrLf & vbCrLf & "  " & g_Language.TranslateMessage("2) The ""Add images using Open Image dialog..."" button.  You can then select one or more image files to be processed.")
            lblExplanation(0).Caption = lblExplanation(0).Caption & vbCrLf & vbCrLf & "  " & g_Language.TranslateMessage("3) Drag-and-drop files directly onto the batch list from Windows Explorer or your desktop.")
            lblExplanation(0).Caption = lblExplanation(0).Caption & vbCrLf & vbCrLf & g_Language.TranslateMessage("Each of these methods supports use of the Ctrl and Shift keys to select multiple files at once.")
        
        'Step 2: choose what photo editing you will apply to each image
        Case 1
        
            lblWizardTitle.Caption = g_Language.TranslateMessage("Step 2: select the photo editing action(s) to apply to each image")
            
            lblExplanation(0).Caption = g_Language.TranslateMessage("(description forthcoming)")
        
        'Step 3: choose the output image format
        Case 2
        
            lblWizardTitle.Caption = g_Language.TranslateMessage("Step 3: choose a destination image format")
            
            lblExplanation(0).Caption = g_Language.TranslateMessage("(description forthcoming)")
        
        'Step 4: choose where processed images will be placed and named
        Case 3
        
            lblWizardTitle.Caption = g_Language.TranslateMessage("Step 4: provide a destination folder and any renaming options")
            
            lblExplanation(0).Caption = g_Language.TranslateMessage("(description forthcoming)")
        
    End Select

End Sub

'Remove all selected items from the batch conversion list
Private Sub cmdRemove_Click()
    
    Dim x As Long
    Do While (x <= lstFiles.ListCount - 1) And (x >= 0)
        If lstFiles.Selected(x) Then
            lstFiles.RemoveItem x
            x = x - 1
        Else
            x = x + 1
        End If
    Loop
    
    'Because there are no longer any selected entries, disable the "remove selected images" button
    cmdRemove.Enabled = False
    
    'And if all files were removed, disable actions that require at least one image
    If lstFiles.ListCount = 0 Then
        cmdRemoveAll.Enabled = False
        cmdSaveList.Enabled = False
        cmdNext.Enabled = False
    End If
        
End Sub

'Remove ALL items from the batch conversion list
Private Sub cmdRemoveAll_Click()
    lstFiles.Clear
    fixHorizontalListBoxScrolling lstFiles
    
    'Because all entries have been removed, disable actions that require at least one image to be present
    cmdRemove.Enabled = False
    cmdRemoveAll.Enabled = False
    cmdSaveList.Enabled = False
    cmdNext.Enabled = False
    
End Sub

Private Sub cmdSaveList_Click()
    
    'Before attempting to save, make sure at least one image has been placed in the list
    If lstFiles.ListCount = 0 Then
        pdMsgBox "You haven't selected any image files.  Please add one or more files to the batch list before saving.", vbExclamation + vbOKOnly + vbApplicationModal, "Empty image list"
        Exit Sub
    End If
        
    'Get the last "open/save image list" path from the INI file
    Dim tempPathString As String
    tempPathString = g_UserPreferences.GetPreference_String("Batch Preferences", "ListFolder", "")
    
    Dim CC As cCommonDialog
    Set CC = New cCommonDialog
    
    Dim sFile As String
    If CC.VBGetSaveFileName(sFile, , True, "Batch Image List (.pdl)|*.pdl|All files|*.*", 0, tempPathString, "Save the current list of images", ".pdl", FormBatchWizard.hWnd, OFN_HIDEREADONLY) Then
        
        'Save this new directory as the default path for future usage
        Dim listPath As String
        listPath = sFile
        StripDirectory listPath
        g_UserPreferences.SetPreference_String "Batch Preferences", "ListFolder", listPath
        
        If FileExist(sFile) Then Kill sFile
        Dim fileNum As Integer
        fileNum = FreeFile
        
        Dim x As Long
        
        Open sFile For Output As #fileNum
            Print #fileNum, "<" & PROGRAMNAME & " BATCH CONVERSION LIST>"
            Print #fileNum, Trim(CStr(lstFiles.ListCount))
            For x = 0 To lstFiles.ListCount - 1
                Print #fileNum, lstFiles.List(x)
            Next x
            Print #fileNum, "<END OF LIST>"
        Close #fileNum
    End If
    
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
    makeFormPretty Me
    
End Sub

'Open a common dialog and allow the user to select a macro file (to apply to each image in the batch list)
Private Sub cmdSelectMacro_Click()
    
    'Get the last macro-related path from the INI file
    Dim tempPathString As String
    tempPathString = g_UserPreferences.GetPreference_String("Program Paths", "Macro", "")
    
    'Prepare a common dialog object
    Dim cDialog As cCommonDialog
    Set cDialog = New cCommonDialog
    
    Dim sFile As String
   
    'If the user provides a valid macro file, use that as part of the batch process
    If cDialog.VBGetOpenFileName(sFile, , , , , True, PROGRAMNAME & " Macro Data (." & MACRO_EXT & ")|*." & MACRO_EXT & "|All files|*.*", , tempPathString, "Open Macro File", "." & MACRO_EXT, Me.hWnd, OFN_HIDEREADONLY) Then
        
        'As a convenience to the user, save this directory as the default macro path
        tempPathString = sFile
        StripDirectory tempPathString
        g_UserPreferences.SetPreference_String "Program Paths", "Macro", tempPathString
        
        'Display the selected macro location in the relevant text box
        txtMacro.Text = sFile
        
        'Also, select the macro option button by default
        optActions(1).Value = True
        
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
    updateSourceImageList
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1
End Sub

Private Sub Form_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    If Source = lstSource Then lstSource.DragIcon = picDragDisallow.Picture
End Sub

Private Sub Form_Load()
        
    'TESTING ONLY: warn the user that this form is not yet complete
    MsgBox "This tool is under heavy construction.  It may not work at all in this development build." & vbCrLf & vbCrLf & "It will be finished by the time 5.4 releases.", vbApplicationModal + vbExclamation + vbOKOnly, "Development still active"
        
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
            'jpegFormatIndex = x
            Exit For
        End If
    Next x
    
    'Populate the combo box for file rename options
    cmbOutputOptions.AddItem "Original filenames"
    cmbOutputOptions.AddItem "Ascending numbers (1, 2, 3, etc.)"
    cmbOutputOptions.ListIndex = 0
    
    updateSourceImageList
    
    'Extract relevant icons from the resource file, and render them onto the buttons at run-time.
    ' (NOTE: because the icons require manifest theming, they will not appear in the IDE.)
    Set cImgCtl = New clsControlImage
    With cImgCtl
        .LoadImageFromStream cmdNext.hWnd, LoadResData("ARROWRIGHT", "CUSTOM"), 32, 32
        .LoadImageFromStream cmdPrevious.hWnd, LoadResData("ARROWLEFT", "CUSTOM"), 32, 32
        .LoadImageFromStream cmdAddFiles.hWnd, LoadResData("ARROWDOWN", "CUSTOM"), 32, 32
        
        .SetMargins cmdNext.hWnd, , , 4
        .Align(cmdNext.hWnd) = Icon_Right
        .SetMargins cmdPrevious.hWnd, 4
        .Align(cmdPrevious.hWnd) = Icon_Left
        .SetMargins cmdAddFiles.hWnd, 4
        .Align(cmdAddFiles.hWnd) = Icon_Left
    End With
    
    'Set the current page number to 0
    m_currentPage = 0
    
    'Display appropriate help text and wizard title
    updateWizardText
    
    'Display some text manually to make sure translations are handled correctly
    txtMacro.Text = g_Language.TranslateMessage("no macro selected")
    lblExplanationFormat.Caption = g_Language.TranslateMessage("if PhotoDemon does not support a particular output format, a standard format will be used instead")
    lblExplanationFormat.Caption = lblExplanationFormat.Caption & vbCrLf & " " & g_Language.TranslateMessage("( JPEG @ 92% quality will be used for photographs, and lossless PNG will be used for non-photographs )")
        
    'Hide all inactive wizard panes
    Dim i As Long
    For i = 1 To picContainer.Count - 1
        picContainer(i).Visible = False
    Next i
    
    'Assign the system hand cursor to all relevant objects
    makeFormPretty Me
    
    'For some reason, the container picture boxes automatically acquire the cursor of children objects.
    ' Manually force those cursors to arrows to prevent this.
    For i = 0 To picContainer.Count - 1
        setArrowCursorToObject picContainer(i)
    Next i
    
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
        If Len(chkFileExt) > 0 Then
            
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

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub lstFiles_Click()
    updatePreview lstFiles.List(lstFiles.ListIndex)
    
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
        
        For Each oleFilename In Data.Files
            tmpString = CStr(oleFilename)
            If tmpString <> "" Then
                If FileExist(tmpString) Then addFileToBatchList tmpString
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
    
End Sub

Private Sub fixHorizontalListBoxScrolling(ByRef srcListBox As ListBox, Optional ByVal lenModifier As Long = 0)
    
    Dim i As Long, lenText As Long, maxWidth As Long
    maxWidth = Me.TextWidth(srcListBox.List(0) & "     ")
    For i = 0 To srcListBox.ListCount - 1
        lenText = Me.TextWidth(srcListBox.List(i) & "     ")
        If lenText > maxWidth Then maxWidth = lenText
    Next i
    
    SendMessageA srcListBox.hWnd, LB_SETHORIZONTALEXTENT, maxWidth + lenModifier, 0
    'SendMessageA srcListBox.hWnd, WM_SETREDRAW, 1, 0
    LockWindowUpdate 0
    
End Sub

'Update the active image preview in the top-right
Private Sub updatePreview(ByVal srcImagePath As String)

    Static lastPreview As String
    
    'Only redraw the preview if it doesn't match the last image we previewed
    If CBool(chkEnablePreview) And (StrComp(lastPreview, srcImagePath, vbTextCompare) <> 0) Then
    
        'Display a preview of the selected image
        Dim tmpImagePath(0) As String
        tmpImagePath(0) = srcImagePath
        
        Dim tmpImage As pdImage
        Set tmpImage = New pdImage
        PreLoadImage tmpImagePath, False, "", "", False, tmpImage, tmpImage.mainLayer, -1
        
        If Not (tmpImage.mainLayer Is Nothing) And (tmpImage.mainLayer.getLayerWidth > 0) And (tmpImage.mainLayer.getLayerHeight > 0) Then
            
            If (tmpImage.mainLayer.getLayerWidth > picPreview.ScaleWidth) Or (tmpImage.mainLayer.getLayerHeight > picPreview.ScaleHeight) Then
                DrawPreviewImage picPreview, True, tmpImage.mainLayer
                tmpImage.mainLayer.eraseLayer
            Else
                'Center the image in the sample area
                Dim ImgWidth As Long, ImgHeight As Long
                ImgWidth = tmpImage.mainLayer.getLayerWidth
                ImgHeight = tmpImage.mainLayer.getLayerHeight
                picPreview.Picture = LoadPicture("")
                If tmpImage.mainLayer.getLayerColorDepth = 32 Then tmpImage.mainLayer.compositeBackgroundColor
                BitBlt picPreview.hDC, (picPreview.ScaleWidth \ 2) - (ImgWidth \ 2), (picPreview.ScaleHeight \ 2) - (ImgHeight \ 2), ImgWidth, ImgHeight, tmpImage.mainLayer.getLayerDC, 0, 0, vbSrcCopy
                picPreview.Picture = picPreview.Image
                picPreview.Refresh
            End If
        Else
            picPreview.Picture = LoadPicture("")
            Dim strToPrint As String
            strToPrint = g_Language.TranslateMessage("Preview not available")
            picPreview.CurrentX = (picPreview.ScaleWidth - picPreview.TextWidth(strToPrint)) \ 2
            picPreview.CurrentY = (picPreview.ScaleHeight - picPreview.TextHeight(strToPrint)) \ 2
            picPreview.Print strToPrint
        End If
    
        lastPreview = srcImagePath
    
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
    If novelAddition Then
        If FileExist(srcFile) Then lstFiles.AddItem srcFile
    End If
    
    'Enable the "remove all images" button if at least one image exists in the processing list
    If lstFiles.ListCount > 0 Then
        If Not cmdRemoveAll.Enabled Then cmdRemoveAll.Enabled = True
        If Not cmdSaveList.Enabled Then cmdSaveList.Enabled = True
        If Not cmdNext.Enabled Then cmdNext.Enabled = True
    End If
    
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
