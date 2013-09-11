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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      TabIndex        =   4
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
      TabIndex        =   3
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
      TabIndex        =   2
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
      TabIndex        =   11
      Top             =   720
      Width           =   11775
      Begin PhotoDemon.smartCheckBox chkGoogleTranslate 
         Height          =   480
         Left            =   5040
         TabIndex        =   36
         Top             =   5640
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   847
         Caption         =   "automatically estimate missing translations with Google Translate"
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
         TabIndex        =   35
         Top             =   3240
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
         TabIndex        =   34
         Top             =   480
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
         TabIndex        =   15
         Top             =   6840
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
         Height          =   5580
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   4500
      End
      Begin VB.Label lblTitle 
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
         Index           =   10
         Left            =   4920
         TabIndex        =   33
         Top             =   2880
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
         TabIndex        =   32
         Top             =   120
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
         TabIndex        =   14
         Top             =   6360
         Width           =   1995
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "list of phrases:"
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
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   1560
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
      TabIndex        =   6
      Top             =   720
      Width           =   11775
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
         TabIndex        =   10
         Top             =   1560
         Width           =   10695
      End
      Begin PhotoDemon.smartOptionButton optBaseLanguage 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   8
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
         TabIndex        =   9
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
         TabIndex        =   7
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
      TabIndex        =   29
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
         TabIndex        =   31
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
         TabIndex        =   30
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
      TabIndex        =   16
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
         TabIndex        =   28
         Text            =   "enter your name here"
         Top             =   5160
         Width           =   10935
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
         TabIndex        =   26
         Text            =   "incomplete"
         Top             =   3240
         Width           =   10935
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
         TabIndex        =   24
         Text            =   "1.0.0"
         Top             =   4200
         Width           =   10935
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
         TabIndex        =   22
         Text            =   "English (US)"
         Top             =   2280
         Width           =   10935
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
         TabIndex        =   20
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
         TabIndex        =   17
         Text            =   "en"
         Top             =   360
         Width           =   615
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
         TabIndex        =   27
         Top             =   4800
         Width           =   1635
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "translation status (e.g. ""complete"", ""unfinished"", etc - any descriptive text is acceptable)"
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
         TabIndex        =   25
         Top             =   2880
         Width           =   9240
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "translation version (in Major.Minor.Revision format, e.g. ""1.0.0"")"
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
         TabIndex        =   23
         Top             =   3840
         Width           =   6855
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "full language name (e.g. ""English (US)""; this will be displayed in PhotoDemon's language menu)"
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
         TabIndex        =   21
         Top             =   1920
         Width           =   10215
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "country ID (2 characters in ISO 3166-1 alpha-2 format, e.g. ""US"" for ""United States"")"
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
         TabIndex        =   19
         Top             =   960
         Width           =   9030
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "language ID (2 characters in ISO 639-1 format, e.g. ""en"" for ""English"")"
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
         TabIndex        =   18
         Top             =   0
         Width           =   7515
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
      TabIndex        =   5
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
      Index           =   0
      Left            =   120
      TabIndex        =   1
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
      TabIndex        =   0
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
'Copyright ©2012-2013 by Frank Donckers and Tanner Helland
'Created: 28/August/13
'Last updated: 03/September/13
'Last update: initial build
'
'Thanks to the incredible work of Frank Donckers, PhotoDemon provides a fully functional language translation engine.
' Many thanks to Frank for taking the initiative on not only implementing the translation engine prototype, but also
' for taking the time to translate the entire PhotoDemon text collection into multiple languages. (This was a huge
' project, as PhotoDemon contains a LOT of text.)
'
'During the translation process, Frank pointed out that translating PhotoDemon's 1,000+ unique phrases takes a loooong
' time.  This new language editor aims to take some of the tediousness out of the process.  I have borrowed many
' concepts and code pieces from a similar project by Frank, which he used to create the original translation files.
'
'The editor requires a source language file to start.  This can either be a blank English language file (provided
' with all PD downloads) or an existing language file.  The user's work will be saved to a new language file of
' their choosing in the User subfolder (currently /Data/Languages).
'
'Google Translate can be used to automatically populate an "estimated" translation.  This was Frank's idea and
' Frank's code - many thanks to him for such a clever feature!
'
'Note: for the Google Translate© Terms of Use, please visit http://www.google.com/policies/terms/
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
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
                lstPhrases.ItemData(lstPhrases.NewIndex) = i
            Next i
        
        'Translated phrases
        Case 1
            For i = 0 To numOfPhrases - 1
                If Len(allPhrases(i).Translation) > 0 Then
                    lstPhrases.AddItem allPhrases(i).ListBoxEntry
                    lstPhrases.ItemData(lstPhrases.NewIndex) = i
                End If
            Next i
        
        'Untranslated phrases
        Case 2
            For i = 0 To numOfPhrases - 1
                If Len(allPhrases(i).Translation) = 0 Then
                    lstPhrases.AddItem allPhrases(i).ListBoxEntry
                    lstPhrases.ItemData(lstPhrases.NewIndex) = i
                End If
            Next i
    
    End Select
                
    LockWindowUpdate 0
    lstPhrases.Refresh
    
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    changeWizardPage True
End Sub

Private Sub cmdPrevious_Click()
    changeWizardPage False
End Sub

'Change the active wizard page.  If moveForward is set to TRUE, the wizard page will be advanced; otherwise, it will move
' to the previous page.
Private Sub changeWizardPage(ByVal moveForward As Boolean)

    Dim i As Long

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
                        .langID = "en-US"
                        .langName = "English (US)"
                        .langStatus = g_Language.TranslateMessage("incomplete")
                        .langType = "Unofficial"
                        .langVersion = "1.0.0"
                        .Author = g_Language.TranslateMessage("enter your name here")
                    End With
                                        
                'For some reason, we failed to load the master language file.  Tell them to download a fresh copy of PD.
                Else
                    Screen.MousePointer = vbDefault
                    tmrProgBar.Enabled = False
                    pdMsgBox "Unfortunately, the master language file could not be located on this PC.  This file is included with the official release of PhotoDemon, but it may not be included with development or beta builds." & vbCrLf & vbCrLf & "To start a new translation, please download a fresh copy of PhotoDemon from tannerhelland.com/photodemon.", vbOKOnly + vbInformation + vbApplicationModal, "Master language file missing"
                    Unload Me
                End If
            
            'They want to edit an existing language.  Follow the same general pattern as for the master language file (above).
            Else
            
                'Fill the current language metadata container with matching information from the selected language,
                ' with a few changes
                curLanguage = listOfAvailableLanguages(lstLanguages.ItemData(lstLanguages.ListIndex))
                curLanguage.FileName = ""
                curLanguage.langType = "Unofficial"
                
                'Attempt to load the selected language from file
                If loadAllPhrasesFromFile(listOfAvailableLanguages(lstLanguages.ItemData(lstLanguages.ListIndex)).FileName) Then
                    
                    'No further action is necessary!
                    
                'For some reason, we failed to load the master language file.  Tell them to download a fresh copy of PD.
                Else
                    Screen.MousePointer = vbDefault
                    tmrProgBar.Enabled = False
                    pdMsgBox "Unfortunately, this language file could not be loaded.  It's possible the copy on this PC is out-of-date." & vbCrLf & vbCrLf & "To continue, please download a fresh copy of PhotoDemon from tannerhelland.com/photodemon.", vbOKOnly + vbInformation + vbApplicationModal, "Language file could not be loaded"
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
        
        'The third page is the phrase editing page.  This is the most important page in the wizard.
        Case 3
    
    End Select
    

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
        cmdNext.Enabled = False
    Else
        cmdNext.Enabled = True
    End If
    
    'Finally, change the top title caption to match the current step
    If curWizardPage < 1 Then
        lblWizardTitle.Caption = g_Language.TranslateMessage("Step %1:", curWizardPage + 1)
    Else
        lblWizardTitle.Caption = g_Language.TranslateMessage("Step %1:", curWizardPage)
    End If
    lblWizardTitle.Caption = lblWizardTitle.Caption & " "
    
    Select Case curWizardPage
    
        Case 0
            lblWizardTitle.Caption = lblWizardTitle.Caption & g_Language.TranslateMessage("select a language file")
            
        Case 2
            lblWizardTitle.Caption = lblWizardTitle.Caption & g_Language.TranslateMessage("add language metadata")
            
        Case 3
            lblWizardTitle.Caption = lblWizardTitle.Caption & g_Language.TranslateMessage("translate all phrases")
    
    End Select
    
End Sub

Private Sub Form_Load()
    
    'Mark the XML file as not loaded
    xmlLoaded = False
    
    'By default, the first wizard page is displayed.  (We start at -1 because we will incerement the page count by +1 with our first
    ' call to changeWizardPage in Form_Activate)
    curWizardPage = -1
    
    'Retrieve a list of available languages from the translation engine
    g_Language.copyListOfLanguages listOfAvailableLanguages
    
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
            Else
                listEntry = listEntry & " ("
                listEntry = listEntry & g_Language.TranslateMessage("unofficial translation by")
                listEntry = listEntry & " "
                If Len(listOfAvailableLanguages(i).Author) > 0 Then
                    listEntry = listEntry & listOfAvailableLanguages(i).Author
                Else
                    listEntry = listEntry & g_Language.TranslateMessage("unknown author")
                End If
                listEntry = listEntry & ")"
            End If
            
            'To save us time in the future, use the .ItemData property of this entry to store the language's original index position
            ' in our listOfAvailableLanguages array.
            lstLanguages.AddItem listEntry
            lstLanguages.ItemData(lstLanguages.NewIndex) = i
            
        Else
            'Ignore the default language entry entirely
        End If
    Next i
    
    'By default, no language is selected for the user
    lstLanguages.ListIndex = -1
    
    'Fill the "phrases to display" combo box
    cmbPhraseFilter.Clear
    cmbPhraseFilter.AddItem g_Language.TranslateMessage("All phrases"), 0
    cmbPhraseFilter.AddItem g_Language.TranslateMessage("Translated phrases"), 1
    cmbPhraseFilter.AddItem g_Language.TranslateMessage("Untranslated phrases"), 2
    cmbPhraseFilter.ListIndex = 0
    
    'Initialize the Google Translate interface
    Set autoTranslate = New clsGoogleTranslate
    autoTranslate.setSrcLanguage "en"
        
    'Assign the system hand cursor to all relevant objects
    Set m_ToolTip = New clsToolTip
    makeFormPretty Me, m_ToolTip
    
    'Advance to the first page
    changeWizardPage True
    
    'DEV WARNING - REMOVE WHEN FINISHED!
    MsgBox "This tool is currently under heavy development.  It may not work as expected (or at all).", vbInformation + vbOKOnly + vbApplicationModal, "Development warning"
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Given a source language file, find all phrase tags, and load them into a specialized phrase array
Private Function loadAllPhrasesFromFile(ByVal srcLangFile As String) As Boolean

    Dim xmlEngine As New pdXML
    Set xmlEngine = New pdXML
    
    'Attempt to load the language file
    If xmlEngine.loadXMLFile(srcLangFile) Then
    
        'Validate the language file's contents
        If xmlEngine.isPDDataType("Translation") And xmlEngine.validateLoadedXMLData("phrase") Then
        
            lblPleaseWait.Caption = "Please wait while the original language file is validated..."
            lblPleaseWait.Refresh
        
            'Attempt to load all phrase tag location occurrences
            Dim phraseLocations() As Long
            If xmlEngine.findAllTagLocations(phraseLocations, "phrase", True) Then
            
                lblPleaseWait.Caption = "Validation successful!  Loading all phrases and preparing translation engine..."
                lblPleaseWait.Refresh
                
                numOfPhrases = UBound(phraseLocations) + 1
                ReDim allPhrases(0 To numOfPhrases - 1) As Phrase
                
                Dim tmpString As String
                
                Dim i As Long
                For i = 0 To numOfPhrases - 1
                    tmpString = xmlEngine.getUniqueTag_String("original", , phraseLocations(i))
                    allPhrases(i).Original = tmpString
                    allPhrases(i).Length = LenB(tmpString)
                    allPhrases(i).Translation = xmlEngine.getUniqueTag_String("translation", , phraseLocations(i))
                    
                    'We also need a modified version of the string to add to the phrase list box.  This text can't include line breaks,
                    ' and it can't be so long that it overflows the list box.
                    If InStr(1, tmpString, vbCrLf) Then tmpString = Replace(tmpString, vbCrLf, "")
                    If InStr(1, tmpString, vbCr) Then tmpString = Replace(tmpString, vbCr, "")
                    If InStr(1, tmpString, vbLf) Then tmpString = Replace(tmpString, vbLf, "")
                    If allPhrases(i).Length > 70 Then tmpString = Left$(tmpString, 35) & "..."
                    
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
End Sub

'When the phrase box is clicked, display the original and translated (if available) text in the right-hand text boxes
Private Sub lstPhrases_Click()
    
    txtOriginal = allPhrases(lstPhrases.ItemData(lstPhrases.ListIndex)).Original
    
    'If a translation exists for this phrase, load it.  If it does not, use Google Translate to estimate a translation
    ' (contingent on the relevant check box setting)
    If Len(allPhrases(lstPhrases.ItemData(lstPhrases.ListIndex)).Translation) > 0 Then
        txtTranslation = allPhrases(lstPhrases.ItemData(lstPhrases.ListIndex)).Translation
    Else
    
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

Private Sub tmrProgBar_Timer()
    
    sysProgBar.Value = sysProgBar.Value + 1
    If sysProgBar.Value = sysProgBar.Max Then sysProgBar.Value = sysProgBar.Min
    
    sysProgBar.Refresh
    
End Sub
