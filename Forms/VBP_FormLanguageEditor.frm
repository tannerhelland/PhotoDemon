VERSION 5.00
Begin VB.Form FormLanguageEditor 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Language Editor"
   ClientHeight    =   9000
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrProgBar 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   14760
      Top             =   120
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11880
      TabIndex        =   1
      Top             =   8310
      Width           =   1725
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13860
      TabIndex        =   2
      Top             =   8310
      Width           =   1365
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "&Previous"
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
      Height          =   615
      Left            =   10080
      TabIndex        =   0
      Top             =   8310
      Width           =   1725
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
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   720
      Width           =   11775
      Begin VB.CommandButton cmdNextPhrase 
         Caption         =   "Save this translation and proceed to the next phrase"
         Height          =   615
         Left            =   5040
         TabIndex        =   18
         Top             =   6600
         Width           =   6615
      End
      Begin PhotoDemon.smartCheckBox chkGoogleTranslate 
         Height          =   480
         Left            =   5040
         TabIndex        =   16
         Top             =   5520
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   847
         Caption         =   "automatically estimate missing translations (via Google Translate)"
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
      Begin VB.TextBox txtTranslation 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   5040
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   3120
         Width           =   6615
      End
      Begin VB.TextBox txtOriginal 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   5040
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   360
         Width           =   6615
      End
      Begin VB.ComboBox cmbPhraseFilter 
         BackColor       =   &H00FFFFFF&
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   6825
         Width           =   4500
      End
      Begin VB.ListBox lstPhrases 
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
         Height          =   5820
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   4500
      End
      Begin PhotoDemon.smartCheckBox chkShortcut 
         Height          =   480
         Left            =   5040
         TabIndex        =   17
         Top             =   6000
         Width           =   5520
         _ExtentX        =   9737
         _ExtentY        =   847
         Caption         =   "ENTER key automatically saves and proceeds to next phrase"
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
      Begin VB.Label lblTranslatedPhrase 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "translated phrase:"
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
         Left            =   4920
         TabIndex        =   39
         Top             =   2760
         Width           =   1905
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "original phrase:"
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
         Index           =   9
         Left            =   4920
         TabIndex        =   38
         Top             =   0
         Width           =   1665
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "phrases to display:"
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
         Left            =   0
         TabIndex        =   27
         Top             =   6465
         Width           =   1995
      End
      Begin VB.Label lblPhraseBox 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "list of phrases (%1 items):"
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
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   2835
      End
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
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   720
      Width           =   11775
      Begin VB.CommandButton cmdDeleteLanguage 
         Caption         =   "Delete selected language file"
         Enabled         =   0   'False
         Height          =   615
         Left            =   8385
         TabIndex        =   6
         Top             =   6360
         Width           =   3135
      End
      Begin VB.ListBox lstLanguages 
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
         Height          =   4620
         Left            =   840
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1560
         Width           =   10695
      End
      Begin PhotoDemon.smartOptionButton optBaseLanguage 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   661
         Caption         =   "start a new language file from scratch"
         Value           =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin PhotoDemon.smartOptionButton optBaseLanguage 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   3510
         _ExtentX        =   6191
         _ExtentY        =   661
         Caption         =   "edit an existing language file:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "language files currently available:"
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
         Left            =   840
         TabIndex        =   24
         Top             =   1080
         Width           =   3540
      End
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
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   720
      Width           =   11775
      Begin VB.PictureBox picProgBar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   785
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   3000
         Width           =   11775
      End
      Begin VB.Label lblPleaseWait 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "please wait..."
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
         Height          =   405
         Left            =   0
         TabIndex        =   36
         Top             =   2400
         Width           =   11760
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
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   720
      Width           =   11775
      Begin VB.TextBox txtLangAuthor 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   240
         TabIndex        =   12
         Text            =   "enter your name here"
         Top             =   5160
         Width           =   11415
      End
      Begin VB.TextBox txtLangStatus 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   240
         TabIndex        =   10
         Text            =   "incomplete"
         Top             =   3240
         Width           =   2895
      End
      Begin VB.TextBox txtLangVersion 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   240
         TabIndex        =   11
         Text            =   "1.0.0"
         Top             =   4200
         Width           =   2895
      End
      Begin VB.TextBox txtLangName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   240
         TabIndex        =   9
         Text            =   "English (US)"
         Top             =   2280
         Width           =   2895
      End
      Begin VB.TextBox txtLangID 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Text            =   "US"
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtLangID 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Text            =   "en"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblSubheader 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "e.g. ""1.0.0"".  Please use Major.Minor.Revision format."
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
         Index           =   4
         Left            =   3360
         TabIndex        =   44
         Top             =   4290
         Width           =   4620
      End
      Begin VB.Label lblSubheader 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "e.g. ""complete"", ""unfinished"", etc.  Any descriptive text is acceptable."
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
         Index           =   3
         Left            =   3360
         TabIndex        =   43
         Top             =   3330
         Width           =   5910
      End
      Begin VB.Label lblSubheader 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "e.g. ""English"" or ""English (US)"".  This text will be displayed in PhotoDemon's Language menu."
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
         Left            =   3360
         TabIndex        =   42
         Top             =   2370
         Width           =   7995
      End
      Begin VB.Label lblSubheader 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "e.g. ""US"" for ""United States"".  Please use the official 2-character ISO 3166-1 format."
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
         Left            =   1080
         TabIndex        =   41
         Top             =   1410
         Width           =   7245
      End
      Begin VB.Label lblSubheader 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "e.g. ""en"" for ""English"".  Please use the official 2-character ISO 639-1 format."
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
         Left            =   1080
         TabIndex        =   40
         Top             =   435
         Width           =   6570
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "author name(s)"
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
         Index           =   8
         Left            =   0
         TabIndex        =   34
         Top             =   4800
         Width           =   1635
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "translation status"
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
         Index           =   4
         Left            =   0
         TabIndex        =   33
         Top             =   2880
         Width           =   1800
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "translation version"
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
         Index           =   3
         Left            =   0
         TabIndex        =   32
         Top             =   3840
         Width           =   1950
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "language name"
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
         Left            =   0
         TabIndex        =   31
         Top             =   1920
         Width           =   1620
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "country ID"
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
         Index           =   6
         Left            =   0
         TabIndex        =   30
         Top             =   960
         Width           =   1125
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "language ID"
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
         Index           =   5
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   1290
      End
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
      TabIndex        =   22
      Top             =   8235
      Width           =   17415
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
      Height          =   7320
      Left            =   120
      TabIndex        =   21
      Top             =   780
      Width           =   3135
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      Index           =   0
      X1              =   224
      X2              =   224
      Y1              =   48
      Y2              =   544
   End
   Begin VB.Label lblWizardTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Step 1: select a language file"
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
      TabIndex        =   20
      Top             =   120
      Width           =   3555
   End
End
Attribute VB_Name = "FormLanguageEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Interactive Language (Translation) Editor
'Copyright ©2013-2014 by Frank Donckers and Tanner Helland
'Created: 28/August/13
'Last updated: 12/September/13
'Last update: completed initial build
'
'Thanks to the incredible work of Frank Donckers, PhotoDemon includes a fully functional language translation engine.
' Many thanks to Frank for taking the initiative on not only implementing the translation engine prototype, but also
' for taking the time to translate the entire PhotoDemon text collection into multiple languages. (This was a huge
' project, as PhotoDemon contains a LOT of text.)
'
'During the translation process, Frank pointed out that translating PhotoDemon's 1,000+ unique phrases takes a loooong
' time.  This new language editor aims to greatly speed up the process.  I have borrowed many concepts and code pieces
' from a similar project by Frank, which he used to create the original translation files.
'
'This integrated language editor requires a source language file to start.  This can either be a blank English
' language file (provided with all PD downloads) or an existing language file.
'
'Data retention is a key focus of the current implementation.  As a safeguard against crashes, two autosaves are
' maintained for any active project.  Every time a phrase is edited or added, an autosave is created.  (Same goes for
' language metadata.)  This should guarantee that even in the event of a crash, nothing more than the last-modified
' phrase will be lost.
'
'To accelerate the translation process, Google Translate can be used to automatically populate an "estimated"
' translation.  This was Frank's idea and Frank's code - many thanks to him for such a clever feature!
'
'Note: for the Google Translate© Terms of Use, please visit http://www.google.com/policies/terms/
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit
Option Compare Text

'Custom tooltip class allows for things like multiline, theming, and multiple monitor support
Dim m_ToolTip As clsToolTip

'The current list of available languages.  This list is not currently updated with the language the user is working on.
' It only contains a list of languages already stored in the /App/PhotoDemon/Languages and Data/Languages folders.
Private listOfAvailableLanguages() As pdLanguageFile

'The language currently being edited.  This curLanguage variable will contain all metadata for the language file.
Private curLanguage As pdLanguageFile

'All phrases that need to be translated will be stored in this array
Private Type Phrase
    Original As String
    Translation As String
    Length As Long
    ListBoxEntry As String
End Type
Private numOfPhrases As Long
Private allPhrases() As Phrase

'Has the source XML language file been loaded yet?
Private xmlLoaded As Boolean

'The current wizard page
Private curWizardPage As Long

'System progress bar control (used on the "please wait" screen)
Private sysProgBar As cProgressBarOfficial

'A Google Translate interface, which we use to auto-populate missing translations
Private autoTranslate As clsGoogleTranslate

'An XML engine is used to parse and update the actual language file contents
Private xmlEngine As pdXML

'To minimize the chance of data loss, PhotoDemon backs up translation data to two alternating files.  In the event of a crash anywhere in
' the editing or export stages, this guarantees that we will never lose more than the last-edited phrase.
Private curBackupFile As Long
Private Const backupFileName As String = "PD_LANG_EDIT_BACKUP_"

'During phrase editing, the user can choose to display all phrases, only translated phrases, or only untranslated phrases.
Private Sub cmbPhraseFilter_Click()

    lstPhrases.Clear
    LockWindowUpdate lstPhrases.hWnd
    
    Dim i As Long
                
    Select Case cmbPhraseFilter.ListIndex
    
        'All phrases
        Case 0
            For i = 0 To numOfPhrases - 1
                lstPhrases.AddItem allPhrases(i).ListBoxEntry
                lstPhrases.ItemData(lstPhrases.newIndex) = i
            Next i
        
        'Translated phrases
        Case 1
            For i = 0 To numOfPhrases - 1
                If Len(allPhrases(i).Translation) > 0 Then
                    lstPhrases.AddItem allPhrases(i).ListBoxEntry
                    lstPhrases.ItemData(lstPhrases.newIndex) = i
                End If
            Next i
        
        'Untranslated phrases
        Case 2
            For i = 0 To numOfPhrases - 1
                If Len(allPhrases(i).Translation) = 0 Then
                    lstPhrases.AddItem allPhrases(i).ListBoxEntry
                    lstPhrases.ItemData(lstPhrases.newIndex) = i
                End If
            Next i
    
    End Select
                
    LockWindowUpdate 0
    lstPhrases.Refresh
    
    updatePhraseBoxTitle
    
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

'Allow the user to delete the selected language file, if they so desire.
Private Sub cmdDeleteLanguage_Click()

    'Make sure a language is selected
    If lstLanguages.ListIndex < 0 Then Exit Sub
    
    Dim msgReturn As VbMsgBoxResult

    'Display different warnings for official languages (which can be restored) and user languages (which cannot)
    If listOfAvailableLanguages(lstLanguages.ItemData(lstLanguages.ListIndex)).langType = "Official" Then
        
        'Make sure we have write access to this folder before attempting to delete anything
        If DirectoryHasWriteAccess(getDirectory(listOfAvailableLanguages(lstLanguages.ItemData(lstLanguages.ListIndex)).FileName)) Then
        
            msgReturn = pdMsgBox("Are you sure you want to delete %1?" & vbCrLf & vbCrLf & "(Even though this is an official PhotoDemon language file, you can safely delete it.)", vbYesNo + vbApplicationModal + vbInformation, "Delete language file", lstLanguages.List(lstLanguages.ListIndex))
            
            If msgReturn = vbYes Then
                Kill listOfAvailableLanguages(lstLanguages.ItemData(lstLanguages.ListIndex)).FileName
                lstLanguages.RemoveItem lstLanguages.ListIndex
                cmdDeleteLanguage.Enabled = False
            End If
        
        'Write access not available
        Else
            pdMsgBox "You do not have access to this folder.  Please log in as an administrator and try again.", vbOKOnly + vbInformation + vbApplicationModal, "Administrator access required"
        End If
    
    'User-folder languages are gone forever once deleted, so change the wording of the deletion confirmation.
    Else
    
        msgReturn = pdMsgBox("Are you sure you want to delete %1?" & vbCrLf & vbCrLf & "(Unless you have manually backed up this language file, this action cannot be undone.)", vbYesNo + vbApplicationModal + vbInformation, "Delete language file", lstLanguages.List(lstLanguages.ListIndex))
        
        If msgReturn = vbYes Then
            Kill listOfAvailableLanguages(lstLanguages.ItemData(lstLanguages.ListIndex)).FileName
            lstLanguages.RemoveItem lstLanguages.ListIndex
            cmdDeleteLanguage.Enabled = False
        End If
        
    End If

End Sub

Private Sub cmdNext_Click()
    changeWizardPage True
End Sub

Private Sub cmdNextPhrase_Click()

    If lstPhrases.ListIndex < 0 Then Exit Sub
    
    'Store this translation to the phrases array
    allPhrases(lstPhrases.ItemData(lstPhrases.ListIndex)).Translation = txtTranslation
    
    'Insert this translation into the original XML file
    xmlEngine.updateTagAtLocation "translation", txtTranslation, xmlEngine.getLocationOfParentTag("phrase", "original", allPhrases(lstPhrases.ItemData(lstPhrases.ListIndex)).Original)
    
    'Write an alternating backup out to file
    performAutosave
        
    'If a specific type of phrase list is displayed, refresh it as necessary
    Dim newIndex As Long
    
    Select Case cmbPhraseFilter.ListIndex
    
        'All phrases
        Case 0
        
            newIndex = lstPhrases.ListIndex + 1
            
            'Attempt to automatically move to the next item in the list
            If newIndex <= lstPhrases.ListCount - 1 Then
                lstPhrases.ListIndex = newIndex
            Else
                If lstPhrases.ListCount > 0 Then lstPhrases.ListIndex = lstPhrases.ListCount - 1
            End If
        
        'Translated phrases
        Case 1
            
            'If the translation has been erased, this item is no longer part of the "translated phrases" group
            If Len(txtTranslation) = 0 Then
                
                newIndex = lstPhrases.ListIndex
                lstPhrases.RemoveItem lstPhrases.ListIndex
                
                'Attempt to automatically move to the next item in the list
                If newIndex <= lstPhrases.ListCount - 1 Then
                    lstPhrases.ListIndex = newIndex
                Else
                    If lstPhrases.ListCount > 0 Then lstPhrases.ListIndex = lstPhrases.ListCount - 1
                End If
                
            End If
        
        'Untranslated phrases
        Case 2
        
            'If a translation has been provided, this item is no longer part of the "untranslated phrases" group
            If Len(txtTranslation) > 0 Then
                
                newIndex = lstPhrases.ListIndex
                lstPhrases.RemoveItem lstPhrases.ListIndex
                
                'Attempt to automatically move to the next item in the list
                If newIndex <= lstPhrases.ListCount - 1 Then
                    lstPhrases.ListIndex = newIndex
                Else
                    If lstPhrases.ListCount > 0 Then lstPhrases.ListIndex = lstPhrases.ListCount - 1
                End If
                
            End If
    
    End Select
    
    updatePhraseBoxTitle

End Sub

Private Sub cmdPrevious_Click()
    changeWizardPage False
End Sub

'Change the active wizard page.  If moveForward is set to TRUE, the wizard page will be advanced; otherwise, it will move
' to the previous page.
Private Sub changeWizardPage(ByVal moveForward As Boolean)

    Dim i As Long

    Dim unloadFormNow As Boolean
    unloadFormNow = False

    'Before changing the page, maek sure all user input on the current page is valid
    Select Case curWizardPage
    
        'The first page is the language selection page.  When the user leaves this page, we must load the language they've selected
        ' into memory - so this validation step is quite large.
        Case 0
        
            'If the user wants to edit an existing language, make sure they've selected one.  (I hate OK-only message boxes, but am
            ' currently too lazy to write a more elegant warning!)
            If optBaseLanguage(1) And (lstLanguages.ListIndex = -1) Then
                pdMsgBox "Please select a language before continuing to the next step.", vbOKOnly + vbInformation + vbApplicationModal, "Please select a language"
                Exit Sub
            End If
            
            'Display the "please wait" panel
            For i = 0 To picContainer.Count - 1
                If i = 1 Then
                    picContainer(i).Visible = True
                Else
                    picContainer(i).Visible = False
                End If
            Next i
            
            'Force a refresh of the visible container picture boxes
            picContainer(1).Refresh
            DoEvents
            
            'Prepare a marquee-style system progress bar
            Set sysProgBar = New cProgressBarOfficial
            sysProgBar.CreateProgressBar picProgBar.hWnd, 0, 0, picProgBar.ScaleWidth, picProgBar.ScaleHeight, True, True, True, True
            sysProgBar.Max = 100
            sysProgBar.Min = 0
            sysProgBar.Value = 0
            sysProgBar.Marquee = True
            sysProgBar.Value = 0

            'Turn on the progress bar timer, which is used to move the marquee progress bar
            tmrProgBar.Enabled = True
            Screen.MousePointer = vbHourglass
                        
            'If they want to start a new language file from scratch, set the load path to the MASTER English language file (which is
            ' hopefully present... if not, there's not much we can do.)
            If optBaseLanguage(0) Then
                                
                If loadAllPhrasesFromFile(g_UserPreferences.getLanguagePath & "Master\MASTER.xml") Then
                    
                    'Populate the current language's metadata container with some default values
                    With curLanguage
                        .FileName = g_UserPreferences.getLanguagePath(True) & "new language.xml"
                        .langID = "en-US"
                        .langName = g_Language.TranslateMessage("New Language")
                        .langStatus = g_Language.TranslateMessage("incomplete")
                        .langType = "Unofficial"
                        .langVersion = "1.0.0"
                        .Author = g_Language.TranslateMessage("enter your name here")
                    End With
                                        
                'For some reason, we failed to load the master language file.  Tell them to download a fresh copy of PD.
                Else
                    Screen.MousePointer = vbDefault
                    tmrProgBar.Enabled = False
                    pdMsgBox "Unfortunately, the master language file could not be located on this PC.  This file is included with the official release of PhotoDemon, but it may not be included with development or beta builds." & vbCrLf & vbCrLf & "To start a new translation, please download a fresh copy of PhotoDemon from photodemon.org.", vbOKOnly + vbInformation + vbApplicationModal, "Master language file missing"
                    Unload Me
                End If
            
            'They want to edit an existing language.  Follow the same general pattern as for the master language file (above).
            Else
            
                'Fill the current language metadata container with matching information from the selected language,
                ' with a few changes
                curLanguage = listOfAvailableLanguages(lstLanguages.ItemData(lstLanguages.ListIndex))
                curLanguage.FileName = g_UserPreferences.getLanguagePath(True) & getFilename(listOfAvailableLanguages(lstLanguages.ItemData(lstLanguages.ListIndex)).FileName)
                
                'Attempt to load the selected language from file
                If loadAllPhrasesFromFile(listOfAvailableLanguages(lstLanguages.ItemData(lstLanguages.ListIndex)).FileName) Then
                    
                    'No further action is necessary!
                    
                'For some reason, we failed to load the master language file.  Tell them to download a fresh copy of PD.
                Else
                    Screen.MousePointer = vbDefault
                    tmrProgBar.Enabled = False
                    pdMsgBox "Unfortunately, this language file could not be loaded.  It's possible the copy on this PC is out-of-date." & vbCrLf & vbCrLf & "To continue, please download a fresh copy of PhotoDemon from photodemon.org.", vbOKOnly + vbInformation + vbApplicationModal, "Language file could not be loaded"
                    Unload Me
                End If
            
            End If
            
            'Advance to the next page
            Screen.MousePointer = vbDefault
            tmrProgBar.Enabled = False
            Set sysProgBar = Nothing
            curWizardPage = curWizardPage + 1
            
        'The second page is the metadata editing page.
        Case 2
        
            'When leaving the metadata page, automatically copy all text box entries into the metadata holder
            With curLanguage
                .langID = Trim$(txtLangID(0)) & "-" & Trim$(txtLangID(1))
                .langName = Trim$(txtLangName)
                .langStatus = Trim$(txtLangStatus)
                .langVersion = Trim$(txtLangVersion)
                .Author = Trim$(txtLangAuthor)
            End With
            
            'Also, automatically set the destination language of the Google Translate interface
            autoTranslate.setDstLanguage Trim$(txtLangID(0))
            
            'Write these updated tags into the original XML text
            With curLanguage
                xmlEngine.updateTagAtLocation "langid", .langID
                xmlEngine.updateTagAtLocation "langname", .langName
                xmlEngine.updateTagAtLocation "langstatus", .langStatus
                xmlEngine.updateTagAtLocation "langversion", .langVersion
                xmlEngine.updateTagAtLocation "author", .Author
            End With
            
            'Update the autosave file
            performAutosave
        
        'The third page is the phrase editing page.  This is the most important page in the wizard.
        Case 3
        
            If moveForward Then
            
                'On this page, the "Next" button is relabeled as "Save and Exit".  It does exactly what it claims!
                Dim CC As cCommonDialog
                Set CC = New cCommonDialog
                
                'If the user is working from an official file or an autosave, the folder and/or extension of the original filename
                ' may not be usable.  Strip just the original filename, and append our own folder and extension.
                Dim sFile As String
                
                If curLanguage.langType = "Autosave" Then
                    sFile = curLanguage.langName
                    makeValidWindowsFilename sFile
                    sFile = getDirectory(curLanguage.FileName) & sFile & ".xml"
                Else
                    sFile = getDirectory(curLanguage.FileName) & getFilenameWithoutExtension(curLanguage.FileName) & ".xml"
                End If
                
                Dim cdFilter As String
                cdFilter = g_Language.TranslateMessage("XML file") & " (.xml)|*.xml"
                
                If CC.VBGetSaveFileName(sFile, , True, cdFilter, , getDirectory(sFile), g_Language.TranslateMessage("Save current language file"), ".xml", Me.hWnd, 0) Then
                
                    'Write the current XML file out to the user's requested path
                    xmlEngine.writeXMLToFile sFile
                    unloadFormNow = True
                    
                Else
                    Exit Sub
                End If
                
            End If
    
    End Select
    
    If unloadFormNow Then
        Unload Me
        Exit Sub
    End If
    
    'Everything has successfully validated, so go ahead and advance (or decrement) the page count
    If moveForward Then
        curWizardPage = curWizardPage + 1
    Else
        curWizardPage = curWizardPage - 1
        If curWizardPage = 1 Then curWizardPage = 0
    End If
    
        
    'We can now apply any entrance-timed panel changes
    Select Case curWizardPage
    
        'Language selection
        Case 0
        
            'Fill the available languages list box with any language files on this system
            populateAvailableLanguages
        
        '"Please wait" panel
        Case 1
        
        'Metadata editor
        Case 2
        
            'When entering the metadata page, automatically fill all boxes with the currently stored metadata entries
            With curLanguage
            
                'Language ID is the most complex, because we must parse the two halves into individual text boxes
                If InStr(1, .langID, "-") > 0 Then
                    txtLangID(0) = Left$(.langID, InStr(1, .langID, "-") - 1)
                    txtLangID(1) = Mid$(.langID, InStr(1, .langID, "-") + 1, Len(.langID) - InStr(1, .langID, "-"))
                Else
                    txtLangID(0) = .langID
                    txtLangID(1) = ""
                End If
                
                'Everything else can be copied directly
                txtLangName = .langName
                txtLangStatus = .langStatus
                txtLangVersion = .langVersion
                txtLangAuthor = .Author
                
            End With
        
        'Phrase editor
        Case 3
        
            'If an XML file was successfully loaded, add its contents to the list box
            If Not xmlLoaded Then
            
                xmlLoaded = True
                
                'Setting the ListIndex property will fire the _Click event, which will handle the actual phrase population
                cmbPhraseFilter.ListIndex = 0
                cmbPhraseFilter_Click
                
            End If
                
    End Select
    
    'Hide all inactive panels (and show the active one)
    For i = 0 To picContainer.Count - 1
        If i = curWizardPage Then picContainer(i).Visible = True Else picContainer(i).Visible = False
    Next i
    
    'If we are at the beginning, disable the previous button
    If curWizardPage = 0 Then cmdPrevious.Enabled = False Else cmdPrevious.Enabled = True
    
    'If we are at the end, change the text of the "next" button; otherwise, make sure it says "next"
    If curWizardPage = picContainer.Count - 1 Then
        cmdNext.Caption = g_Language.TranslateMessage("&Save and Exit")
    Else
        cmdNext.Caption = g_Language.TranslateMessage("&Next")
    End If
    
    'Finally, change the top title caption and left-hand help text to match the current step
    If curWizardPage < 1 Then
        lblWizardTitle.Caption = g_Language.TranslateMessage("Step %1:", curWizardPage + 1)
    Else
        lblWizardTitle.Caption = g_Language.TranslateMessage("Step %1:", curWizardPage)
    End If
    lblWizardTitle.Caption = lblWizardTitle.Caption & " "
    
    Dim helpText As String
    
    Select Case curWizardPage
    
        Case 0
            lblWizardTitle.Caption = lblWizardTitle.Caption & g_Language.TranslateMessage("select a language file")
            
            helpText = g_Language.TranslateMessage("This tool allows you to create and edit PhotoDemon language files.")
            helpText = helpText & vbCrLf & vbCrLf & g_Language.TranslateMessage("Please start by selecting a base language file.  If the selected file already contains translation data, you will be able to edit any existing translations, as well as add translations that may be missing.")
            helpText = helpText & vbCrLf & vbCrLf & g_Language.TranslateMessage("This page also allows you to delete unused language files.  Note that there is no Undo when deleting language files, so please be careful!")
            helpText = helpText & vbCrLf & vbCrLf & g_Language.TranslateMessage("Upon clicking Next, the selected file will automatically be validated and parsed.  Depending on the number of translations present, this process may take a few seconds.")
            If Not g_IsProgramCompiled Then helpText = helpText & vbCrLf & vbCrLf & g_Language.TranslateMessage("(For best results, do not use this editor in the IDE!)")
            
        Case 2
            lblWizardTitle.Caption = lblWizardTitle.Caption & g_Language.TranslateMessage("add language metadata")
            
            helpText = g_Language.TranslateMessage("In this step, please provide a bit of metadata regarding this language.  This information helps PhotoDemon know how to handle this language file.")
            helpText = helpText & vbCrLf & vbCrLf & g_Language.TranslateMessage("The most important items on this page are the language ID and language name.  If these items are missing or invalid, PhotoDemon won't be able to use the language file.")
            helpText = helpText & vbCrLf & vbCrLf & g_Language.TranslateMessage("If multiple translators have worked on this language file, please separate their names with commas.  If this language file is based on an existing language file, please include the original author's name.")
            helpText = helpText & vbCrLf & vbCrLf & g_Language.TranslateMessage("(NOTE: changes made to this page won't be auto-saved unless you click the Next or Previous button.)")
            
        Case 3
            lblWizardTitle.Caption = lblWizardTitle.Caption & g_Language.TranslateMessage("translate all phrases")
            
            helpText = g_Language.TranslateMessage("This final step allows you to edit existing translations, and add missing ones.")
            helpText = helpText & vbCrLf & vbCrLf & g_Language.TranslateMessage("Every time a phrase is modified, an autosave will automatically be created in PhotoDemon's user language folder.  This means you can exit the program at any time without losing your work.")
            helpText = helpText & vbCrLf & vbCrLf & g_Language.TranslateMessage("When you are done translating, you may use the Save and Exit button to save your work to a file of your choosing.  (Note that autosave data will be preserved either way.)")
            helpText = helpText & vbCrLf & vbCrLf & g_Language.TranslateMessage("When you are finished editing this language, please consider sharing it!  Contact me by visiting:")
            helpText = helpText & vbCrLf & g_Language.TranslateMessage("photodemon.org/about/contact/")
            helpText = helpText & vbCrLf & g_Language.TranslateMessage("so we can discuss adding your translation to the official list of supported languages.  Even partial translations are helpful!")
    
    End Select
    
    lblExplanation.Caption = helpText
    lblExplanation.Refresh
    
    'If translations are active, the translated text may not fit the label.  Automatically adjust it to fit.
    fitWordwrapLabel lblExplanation, Me
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    'If the down arrow key is pressed while the user is in the phrase-editing panel, automatically save the current
    ' phrase and move to the next one.
    If CBool(chkShortcut) And (KeyCode = vbKeyReturn) And (curWizardPage = 3) Then
        cmdNextPhrase_Click
        KeyCode = 0
    End If
    
End Sub

Private Sub Form_Load()
    
    'Mark the XML file as not loaded
    xmlLoaded = False
    curBackupFile = 0
    
    'By default, the first wizard page is displayed.  (We start at -1 because we will incerement the page count by +1 with our first
    ' call to changeWizardPage in Form_Activate)
    curWizardPage = -1
        
    'Fill the "phrases to display" combo box
    cmbPhraseFilter.Clear
    cmbPhraseFilter.AddItem "All phrases", 0
    cmbPhraseFilter.AddItem "Translated phrases", 1
    cmbPhraseFilter.AddItem "Untranslated phrases", 2
    cmbPhraseFilter.ListIndex = 0
    
    'Initialize the Google Translate interface
    Set autoTranslate = New clsGoogleTranslate
    autoTranslate.setSrcLanguage "en"
        
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'Advance to the first page
    changeWizardPage True
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Given a source language file, find all phrase tags, and load them into a specialized phrase array
Private Function loadAllPhrasesFromFile(ByVal srcLangFile As String) As Boolean

    Set xmlEngine = New pdXML
    
    'Attempt to load the language file
    If xmlEngine.loadXMLFile(srcLangFile) Then
    
        'Validate the language file's contents
        If xmlEngine.isPDDataType("Translation") And xmlEngine.validateLoadedXMLData("phrase") Then
        
            lblPleaseWait.Caption = g_Language.TranslateMessage("Please wait while the language file is validated...")
            lblPleaseWait.Refresh
        
            'Attempt to load all phrase tag location occurrences
            Dim phraseLocations() As Long
            If xmlEngine.findAllTagLocations(phraseLocations, "phrase", True) Then
            
                lblPleaseWait.Caption = g_Language.TranslateMessage("Validation successful!  Loading all phrases and preparing translation engine...")
                lblPleaseWait.Refresh
                
                numOfPhrases = UBound(phraseLocations) + 1
                ReDim allPhrases(0 To numOfPhrases - 1) As Phrase
                
                Dim tmpString As String
                
                Dim i As Long
                For i = 0 To numOfPhrases - 1
                    tmpString = xmlEngine.getUniqueTag_String("original", , phraseLocations(i))
                    allPhrases(i).Original = tmpString
                    allPhrases(i).Length = Len(tmpString)
                    allPhrases(i).Translation = xmlEngine.getUniqueTag_String("translation", , phraseLocations(i))
                    
                    'We also need a modified version of the string to add to the phrase list box.  This text can't include line breaks,
                    ' and it can't be so long that it overflows the list box.
                    If InStr(1, tmpString, vbCrLf) Then tmpString = Replace(tmpString, vbCrLf, "")
                    If InStr(1, tmpString, vbCr) Then tmpString = Replace(tmpString, vbCr, "")
                    If InStr(1, tmpString, vbLf) Then tmpString = Replace(tmpString, vbLf, "")
                    If allPhrases(i).Length > 35 Then tmpString = Left$(tmpString, 35) & "..."
                    
                    allPhrases(i).ListBoxEntry = tmpString
                    
                    'I don't like using DoEvents, but we need a way to refresh the progress bar.
                    If (i And 3) = 0 Then DoEvents
                    
                Next i
                
                loadAllPhrasesFromFile = True
            
            Else
                loadAllPhrasesFromFile = False
            End If
        
        Else
            loadAllPhrasesFromFile = False
        End If
    
    Else
        loadAllPhrasesFromFile = False
    End If

End Function

Private Sub lstLanguages_Click()
    If Not optBaseLanguage(1) Then optBaseLanguage(1) = True
    If lstLanguages.ListIndex >= 0 Then cmdDeleteLanguage.Enabled = True Else cmdDeleteLanguage.Enabled = False
End Sub

'When the phrase box is clicked, display the original and translated (if available) text in the right-hand text boxes
Private Sub lstPhrases_Click()
    
    lblTranslatedPhrase.Caption = g_Language.TranslateMessage("translated phrase:")
    lblTranslatedPhrase.ForeColor = RGB(64, 64, 64)
    
    txtOriginal = allPhrases(lstPhrases.ItemData(lstPhrases.ListIndex)).Original
    
    'If a translation exists for this phrase, load it.  If it does not, use Google Translate to estimate a translation
    ' (contingent on the relevant check box setting)
    lblTranslatedPhrase.Caption = g_Language.TranslateMessage("translated phrase")
    
    If Len(allPhrases(lstPhrases.ItemData(lstPhrases.ListIndex)).Translation) > 0 Then
        txtTranslation = allPhrases(lstPhrases.ItemData(lstPhrases.ListIndex)).Translation
        lblTranslatedPhrase = lblTranslatedPhrase & " " & g_Language.TranslateMessage("(saved):")
    Else
    
        lblTranslatedPhrase = lblTranslatedPhrase & " " & g_Language.TranslateMessage("(NOT YET SAVED):")
        lblTranslatedPhrase.ForeColor = RGB(208, 52, 52)
    
        If CBool(chkGoogleTranslate) Then
            txtTranslation = g_Language.TranslateMessage("waiting for Google Translate...")
            
            'I've had trouble with the text boxes not clearing properly (no idea why), so manually clear them before
            ' assigning new text.
            Dim retString As String
            retString = autoTranslate.getGoogleTranslation(allPhrases(lstPhrases.ItemData(lstPhrases.ListIndex)).Original)
            If Len(retString) > 0 Then
                txtTranslation = ""
                txtTranslation = retString
            Else
                txtTranslation = ""
                txtTranslation = g_Language.TranslateMessage("translation failed!")
            End If
        Else
            txtTranslation = ""
        End If
            
    End If
        
End Sub

Private Sub optBaseLanguage_Click(Index As Integer)

    If lstLanguages.ListIndex >= 0 Then
        cmdDeleteLanguage.Enabled = True
    Else
        cmdDeleteLanguage.Enabled = False
    End If

End Sub

Private Sub tmrProgBar_Timer()
    
    sysProgBar.Value = sysProgBar.Value + 1
    If sysProgBar.Value = sysProgBar.Max Then sysProgBar.Value = sysProgBar.Min
    
    sysProgBar.Refresh
    
End Sub

'The phrase list box label will automatically be updated with the current count of list items
Private Sub updatePhraseBoxTitle()
    If lstPhrases.ListCount > 0 Then
        lblPhraseBox.Caption = g_Language.TranslateMessage("list of phrases (%1 items)", lstPhrases.ListCount - 1)
    Else
        lblPhraseBox.Caption = g_Language.TranslateMessage("list of phrases (%1 items)", 0)
    End If
End Sub

'Call this function whenever we want the in-memory XML data saved to an autosave file
Private Sub performAutosave()

    'We keep two autosaves at all times; simply alternate between them each time a save is requested
    If curBackupFile = 1 Then curBackupFile = 0 Else curBackupFile = 1
    
    'Generate an autosave filename.  The language ID is appended to the name, so separate autosaves will exist for each edited language
    ' (assuming they have different language IDs).
    Dim backupFile As String
    backupFile = g_UserPreferences.getLanguagePath(True) & backupFileName & curLanguage.langID & "_" & CStr(curBackupFile) & ".tmpxml"
    
    'The XML engine handles the actual writing to file.  For performance reasons, auto-tabbing is suppressed.
    xmlEngine.writeXMLToFile backupFile, True

End Sub

'Fill the first panel ("select a language file") with all available language files on this system
Private Sub populateAvailableLanguages()

    'Retrieve a list of available languages from the translation engine
    g_Language.copyListOfLanguages listOfAvailableLanguages
    
    'We now do a bit of additional work.  Look for any autosave files (with extension .tmpxml) in the user language folder.  Allow the
    ' user to load these if available.
    Dim chkFile As String
    chkFile = Dir(g_UserPreferences.getLanguagePath(True) & "*.tmpxml", vbNormal)
        
    Do While (chkFile <> "")
        
        'Use PD's XML engine to load the file
        Dim tmpXMLEngine As pdXML
        Set tmpXMLEngine = New pdXML
        If tmpXMLEngine.loadXMLFile(g_UserPreferences.getLanguagePath(True) & chkFile) Then
        
            'Use the XML engine to validate this file, and to make sure it contains at least a language ID, name, and one (or more) translated phrase
            If tmpXMLEngine.isPDDataType("Translation") And tmpXMLEngine.validateLoadedXMLData("langid", "langname", "phrase") Then
            
                ReDim Preserve listOfAvailableLanguages(0 To UBound(listOfAvailableLanguages) + 1) As pdLanguageFile
                
                With listOfAvailableLanguages(UBound(listOfAvailableLanguages))
                    'Get the language ID and name - these are the most important values, and technically the only REQUIRED ones.
                    .langID = tmpXMLEngine.getUniqueTag_String("langid")
                    .langName = tmpXMLEngine.getUniqueTag_String("langname")
    
                    'Version, status, and author information should also be present, but the file will still be loaded even if they don't exist
                    .langVersion = tmpXMLEngine.getUniqueTag_String("langversion")
                    .langStatus = tmpXMLEngine.getUniqueTag_String("langstatus")
                    .Author = tmpXMLEngine.getUniqueTag_String("author")
                    
                    'Finally, add some internal metadata
                    .FileName = g_UserPreferences.getLanguagePath(True) & chkFile
                    .langType = "Autosave"
                    
                End With
                
            End If
            
        End If
        
        'Retrieve the next file and repeat
        chkFile = Dir
    
    Loop
    
    'All autosave files have now been loaded as well
    
    'Add the contents of that array to the list box on the opening panel (the list of available languages, from which the user
    ' can select a language file as the "starting point" for their own translation).
    lstLanguages.Clear
    
    Dim i As Long
    For i = 0 To UBound(listOfAvailableLanguages)
    
        'Note that we DO NOT add the English language entry - that is used by the "start a new language file from scratch" option.
        If StrComp(listOfAvailableLanguages(i).langType, "Default", vbTextCompare) <> 0 Then
            Dim listEntry As String
            listEntry = listOfAvailableLanguages(i).langName
            
            'For official translations, an author name will always be provided.  Include the author's name in the list.
            If listOfAvailableLanguages(i).langType = "Official" Then
                listEntry = listEntry & " ("
                listEntry = listEntry & g_Language.TranslateMessage("official translation by")
                listEntry = listEntry & " " & listOfAvailableLanguages(i).Author
                listEntry = listEntry & ")"
            
            'For unofficial translations, an author name may not be provided.  Include the author's name only if it's available.
            ElseIf listOfAvailableLanguages(i).langType = "Unofficial" Then
                listEntry = listEntry & " "
                listEntry = listEntry & g_Language.TranslateMessage("by")
                listEntry = listEntry & " "
                If Len(listOfAvailableLanguages(i).Author) > 0 Then
                    listEntry = listEntry & listOfAvailableLanguages(i).Author
                Else
                    listEntry = listEntry & g_Language.TranslateMessage("unknown author")
                End If
                
            'Anything else is an autosave.
            Else
            
                'Include author name if available
                listEntry = listEntry & " "
                listEntry = listEntry & g_Language.TranslateMessage("by")
                listEntry = listEntry & " "
                If Len(listOfAvailableLanguages(i).Author) > 0 Then
                    listEntry = listEntry & listOfAvailableLanguages(i).Author
                Else
                    listEntry = listEntry & g_Language.TranslateMessage("unknown author")
                End If
                
                'Display autosave time and date
                listEntry = listEntry & " ("
                listEntry = listEntry & g_Language.TranslateMessage("autosaved on")
                listEntry = listEntry & " "
                listEntry = listEntry & Format(FileDateTime(listOfAvailableLanguages(i).FileName), "hh:mm:ss AM/PM, dd-mmm-yy")
                listEntry = listEntry & ") "
            
            End If
            
            'To save us time in the future, use the .ItemData property of this entry to store the language's original index position
            ' in our listOfAvailableLanguages array.
            lstLanguages.AddItem listEntry
            lstLanguages.ItemData(lstLanguages.newIndex) = i
            
        Else
            'Ignore the default language entry entirely
        End If
    Next i
    
    'By default, no language is selected for the user
    lstLanguages.ListIndex = -1
    
End Sub
