VERSION 5.00
Begin VB.Form FormLanguageEditor 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Language editor"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15360
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdButton cmdPrevious 
      Height          =   615
      Left            =   10080
      TabIndex        =   4
      Top             =   8310
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   1085
      Caption         =   "Previous"
   End
   Begin PhotoDemon.pdButton cmdNext 
      Height          =   615
      Left            =   11880
      TabIndex        =   18
      Top             =   8310
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   1085
      Caption         =   "Next"
   End
   Begin PhotoDemon.pdButton cmdCancel 
      Height          =   615
      Left            =   13860
      TabIndex        =   19
      Top             =   8310
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1085
      Caption         =   "Cancel"
   End
   Begin PhotoDemon.pdLabel lblExplanation 
      Height          =   7320
      Left            =   120
      Top             =   780
      Width           =   3135
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   ""
      ForeColor       =   4210752
      Layout          =   1
   End
   Begin PhotoDemon.pdLabel lblWizardTitle 
      Height          =   405
      Left            =   120
      Top             =   120
      Width           =   14475
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "Step 1: select a language file"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   7455
      Index           =   2
      Left            =   3480
      Top             =   720
      Width           =   11775
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdButton cmdUseReference 
         Height          =   735
         Left            =   11040
         TabIndex        =   25
         Top             =   5520
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1296
         Caption         =   "use"
      End
      Begin PhotoDemon.pdTextBox txtReference 
         Height          =   735
         Left            =   5040
         TabIndex        =   24
         Top             =   5520
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1296
         Multiline       =   -1  'True
      End
      Begin PhotoDemon.pdListBox lstPhrases 
         Height          =   5175
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   9128
      End
      Begin PhotoDemon.pdDropDown cboPhraseFilter 
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   6000
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   661
      End
      Begin PhotoDemon.pdButton cmdNextPhrase 
         Height          =   615
         Left            =   5040
         TabIndex        =   5
         Top             =   6720
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   1085
         Caption         =   "Save this translation and proceed to the next phrase"
      End
      Begin PhotoDemon.pdTextBox txtTranslation 
         Height          =   1965
         Left            =   5040
         TabIndex        =   7
         Top             =   2760
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   3466
         Multiline       =   -1  'True
      End
      Begin PhotoDemon.pdTextBox txtOriginal 
         Height          =   1995
         Left            =   5040
         TabIndex        =   9
         Top             =   360
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   3519
         Multiline       =   -1  'True
      End
      Begin PhotoDemon.pdCheckBox chkOnlineTranslate 
         Height          =   330
         Left            =   5040
         TabIndex        =   2
         Top             =   4800
         Width           =   6600
         _ExtentX        =   11642
         _ExtentY        =   582
         Caption         =   "use an online service to estimate missing translations"
      End
      Begin PhotoDemon.pdCheckBox chkShortcut 
         Height          =   330
         Left            =   5040
         TabIndex        =   3
         Top             =   6360
         Width           =   6600
         _ExtentX        =   11642
         _ExtentY        =   582
         Caption         =   "CTRL+ENTER automatically saves and proceeds to next phrase"
      End
      Begin PhotoDemon.pdLabel lblTranslatedPhrase 
         Height          =   285
         Left            =   4920
         Top             =   2400
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   503
         Caption         =   "translated phrase"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   9
         Left            =   4920
         Top             =   0
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   503
         Caption         =   "original phrase"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   2
         Left            =   0
         Top             =   5625
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   503
         Caption         =   "phrases to display"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblPhraseBox 
         Height          =   285
         Left            =   0
         Top             =   0
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   503
         Caption         =   "list of phrases (%1 items)"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdButton cmdAutoTranslate 
         Height          =   615
         Left            =   240
         TabIndex        =   20
         Top             =   6720
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   1085
         Caption         =   "Initiate auto-translation of all missing phrases"
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   11
         Left            =   5040
         Top             =   5160
         Width           =   6585
         _ExtentX        =   11615
         _ExtentY        =   503
         Caption         =   "reference translation (if a .po file was provided)"
         ForeColor       =   4210752
      End
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   7455
      Index           =   0
      Left            =   3480
      Top             =   720
      Width           =   11775
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdButton cmdPO 
         Height          =   330
         Left            =   7680
         TabIndex        =   23
         Top             =   7080
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
         Caption         =   "..."
      End
      Begin PhotoDemon.pdListBox lstLanguages 
         Height          =   4215
         Left            =   840
         TabIndex        =   6
         Top             =   1560
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   8070
      End
      Begin PhotoDemon.pdButton cmdDeleteLanguage 
         Height          =   690
         Left            =   8400
         TabIndex        =   10
         Top             =   5880
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1217
         Caption         =   "Delete selected language file"
      End
      Begin PhotoDemon.pdRadioButton optBaseLanguage 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   11325
         _ExtentX        =   19976
         _ExtentY        =   582
         Caption         =   "start a new language file"
         Value           =   -1  'True
      End
      Begin PhotoDemon.pdRadioButton optBaseLanguage 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   11325
         _ExtentX        =   19976
         _ExtentY        =   582
         Caption         =   "edit an existing language file:"
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   0
         Left            =   840
         Top             =   1200
         Width           =   10650
         _ExtentX        =   18785
         _ExtentY        =   503
         Caption         =   "language files currently available"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdTextBox txtApiKey 
         Height          =   345
         Left            =   1080
         TabIndex        =   21
         Top             =   6240
         Width           =   7095
         _ExtentX        =   12303
         _ExtentY        =   609
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   1
         Left            =   840
         Top             =   5880
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   503
         Caption         =   "(optional) free DeepL.com API key for translation suggestions"
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   10
         Left            =   840
         Top             =   6720
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   503
         Caption         =   "(optional) 3rd-party language file (.po) to compare"
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdTextBox txtPO 
         Height          =   345
         Left            =   1080
         TabIndex        =   22
         Top             =   7080
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   609
      End
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   7455
      Index           =   1
      Left            =   3480
      Top             =   720
      Width           =   11775
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdTextBox txtLangID 
         Height          =   345
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   1335
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   609
         FontSize        =   11
         Text            =   "US"
      End
      Begin PhotoDemon.pdTextBox txtLangID 
         Height          =   345
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   375
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   609
         FontSize        =   11
         Text            =   "en"
      End
      Begin PhotoDemon.pdTextBox txtLangName 
         Height          =   345
         Left            =   240
         TabIndex        =   14
         Top             =   2295
         Width           =   2910
         _ExtentX        =   5133
         _ExtentY        =   609
         FontSize        =   11
         Text            =   "English (US)"
      End
      Begin PhotoDemon.pdTextBox txtLangStatus 
         Height          =   345
         Left            =   240
         TabIndex        =   15
         Top             =   3255
         Width           =   2910
         _ExtentX        =   5133
         _ExtentY        =   609
         FontSize        =   11
         Text            =   "incomplete"
      End
      Begin PhotoDemon.pdTextBox txtLangVersion 
         Height          =   345
         Left            =   240
         TabIndex        =   16
         Top             =   4215
         Width           =   2910
         _ExtentX        =   5133
         _ExtentY        =   609
         FontSize        =   11
         Text            =   "1.0.0"
      End
      Begin PhotoDemon.pdTextBox txtLangAuthor 
         Height          =   345
         Left            =   240
         TabIndex        =   17
         Top             =   5190
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   609
         FontSize        =   11
         Text            =   "enter your name here"
      End
      Begin PhotoDemon.pdLabel lblSubheader 
         Height          =   240
         Index           =   4
         Left            =   3360
         Top             =   4290
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   423
         Caption         =   "e.g. ""1.0.0"".  Please use Major.Minor.Revision format."
         FontSize        =   9
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblSubheader 
         Height          =   240
         Index           =   3
         Left            =   3360
         Top             =   3330
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   423
         Caption         =   "e.g. ""complete"", ""unfinished"", etc.  Any descriptive text is acceptable."
         FontSize        =   9
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblSubheader 
         Height          =   240
         Index           =   2
         Left            =   3360
         Top             =   2370
         Width           =   7995
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "e.g. ""English"" or ""English (US)"".  This text will be displayed in PhotoDemon's Language menu."
         FontSize        =   9
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblSubheader 
         Height          =   240
         Index           =   1
         Left            =   1080
         Top             =   1410
         Width           =   10125
         _ExtentX        =   17859
         _ExtentY        =   423
         Caption         =   "e.g. ""US"" for ""United States"".  Please use the official 2-character ISO 3166-1 format."
         FontSize        =   9
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblSubheader 
         Height          =   240
         Index           =   0
         Left            =   1080
         Top             =   435
         Width           =   10050
         _ExtentX        =   17727
         _ExtentY        =   423
         Caption         =   "e.g. ""en"" for ""English"".  Please use the official 2-character ISO 639-1 format."
         FontSize        =   9
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   8
         Left            =   0
         Top             =   4800
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   503
         Caption         =   "author name(s)"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   4
         Left            =   0
         Top             =   2880
         Width           =   11520
         _ExtentX        =   20320
         _ExtentY        =   503
         Caption         =   "translation status"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   3
         Left            =   0
         Top             =   3840
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   503
         Caption         =   "translation version"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   7
         Left            =   0
         Top             =   1920
         Width           =   11460
         _ExtentX        =   20214
         _ExtentY        =   503
         Caption         =   "language name"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   6
         Left            =   0
         Top             =   960
         Width           =   11445
         _ExtentX        =   20188
         _ExtentY        =   503
         Caption         =   "country ID"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   5
         Left            =   0
         Top             =   0
         Width           =   11490
         _ExtentX        =   20267
         _ExtentY        =   503
         Caption         =   "language ID"
         FontSize        =   12
         ForeColor       =   4210752
      End
   End
End
Attribute VB_Name = "FormLanguageEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Interactive Language (Translation) Editor
'Copyright 2013-2022 by Frank Donckers and Tanner Helland
'Created: 28/August/13
'Last updated: 30/June/22
'Last update: new option for specifying a reference localization file (.po format) from other software.
'             I am using this to compare machine-translated text to GIMP and Krita in hopes of minimizing
'             egregious errors.
'
'Thanks to the incredible work of Frank Donckers, PhotoDemon ships with a custom-built text localization engine.
' Thank you to Frank for implementing the original translation engine prototype, and for taking the time to
' translate all of PhotoDemon's text into multiple languages. (This was a huge project, as PhotoDemon contains
' a LOT of text.)
'
'During the translation process, Frank pointed out that translating PhotoDemon's 1,000+ unique phrases takes
' a loooong time.  This new language editor aims to accelerate the process.  I have borrowed many concepts
' and code pieces from a similar project by Frank, which he used to create PhotoDemon's first translation files.
'
'This integrated language editor requires a source language file to start.  This can either be a blank English
' language file (provided with all PD downloads) or an existing language file.
'
'Data retention is a key focus of the current implementation.  As a safeguard against crashes, two autosaves are
' maintained for any active project.  Every time a phrase is edited or added, an autosave is created.
' This should guarantee that even in the event of a catastrophic failure (power failure or similar), only the
' last-modified phrase would ever risk being lost.
'
'To accelerate the translation process, DeepL.com can be used to automatically populate an "estimated"
' translation of a given phrase.  This uses the official DeepL translation API (via curl) and if you want to use
' it too, you will need to generate a free DeepL API key:
'
' https://www.deepl.com/pro-api?cta=header-pro-api
'
'(Scroll down to the "Free" box and click "sign up for free".)  This feature previously used Google Translate,
' but after receiving a ton of feedback that the translation results from Google were poor, I have since migrated
' to DeepL, which has received better reviews from users.  Note that a human will always need to review
' localizations for best results, but since I am not a polyglot I am not much help here - feedback from native
' speakers is *always* welcome.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit
Option Compare Text

'The current list of available languages.  This list is not currently updated with the language the user is working on.
' It only contains a list of languages already stored in the /App/PhotoDemon/Languages and Data/Languages folders.
Private m_ListOfLanguages() As PDLanguageFile

'The language currently being edited.  This m_curLanguage variable will contain all metadata for the language file.
Private m_curLanguage As PDLanguageFile

'All phrases that need to be translated will be stored in this array
Private Type PD_Phrase
    Original As String
    Translation As String
    Length As Long
    ListBoxEntry As String
    IsMachineTranslation As Boolean
End Type
Private m_NumOfPhrases As Long
Private m_AllPhrases() As PD_Phrase

'Has the source XML language file been loaded yet?
Private m_xmlLoaded As Boolean

'The current wizard page
Private m_WizardPage As Long

'A Google Translate interface, which we use to auto-populate missing translations
Private m_AutoTranslate As pdAutoLocalize

'An XML engine is used to parse and update the actual language file contents
Private m_XMLEngine As pdXML

'To minimize the chance of data loss, PhotoDemon backs up translation data to two alternating files.  In the event of a crash anywhere in
' the editing or export stages, this guarantees that we will never lose more than the last-edited phrase.
Private m_curBackupFile As Long
Private Const backupFileName As String = "PD_LANG_EDIT_BACKUP_"

'Hacky fix for specialized Ctrl+Enter detection
Private m_InKeyEvent As Boolean

'The user can (optionally) point at a target .po file for comparison.  This is very helpful for comparing
' phrases to their equivalent in other open-source software, which reduces the chance of us using translations
' different from what everyone else is using.
Private m_PoComparison As pdStringHash

'During phrase editing, the user can choose to display all phrases, only translated phrases, or only untranslated phrases.
Private Sub cboPhraseFilter_Click()

    lstPhrases.Clear
    lstPhrases.SetAutomaticRedraws False
    
    Dim i As Long
                
    Select Case cboPhraseFilter.ListIndex
    
        'All phrases
        Case 0
            For i = 0 To m_NumOfPhrases - 1
                lstPhrases.AddItem m_AllPhrases(i).ListBoxEntry
            Next i
        
        'Translated phrases
        Case 1
            For i = 0 To m_NumOfPhrases - 1
                If (LenB(m_AllPhrases(i).Translation) <> 0) Then
                    lstPhrases.AddItem m_AllPhrases(i).ListBoxEntry
                End If
            Next i
        
        'Untranslated phrases
        Case 2
            For i = 0 To m_NumOfPhrases - 1
                If (LenB(m_AllPhrases(i).Translation) = 0) Then
                    lstPhrases.AddItem m_AllPhrases(i).ListBoxEntry
                End If
            Next i
            
        '(Optional) phrases that don't match the supplied reference .po
        Case 3
            If (Not m_PoComparison Is Nothing) Then
                If (m_PoComparison.GetNumOfItems > 0) Then
                    
                    Dim tmpString As String
                    
                    For i = 0 To m_NumOfPhrases - 1
                        If m_PoComparison.GetItemByKey(LCase$(m_AllPhrases(i).Original), tmpString) Then
                            If Strings.StringsNotEqual(tmpString, m_AllPhrases(i).Translation, True) Then
                                lstPhrases.AddItem m_AllPhrases(i).ListBoxEntry
                            End If
                        End If
                    Next i
                    
                End If
            End If
    
    End Select
                
    lstPhrases.SetAutomaticRedraws True, True
    
    UpdatePhraseBoxTitle
    
End Sub

'Use Google Translate to auto-translate all untranslated messages.  Note that this is not a great implementation, but it
' should be "good enough" for PD's purposes.
Private Sub cmdAutoTranslate_Click()
    
    'If the program is interrupted while auto-translations are taking place, the IE object will stall and the function will crash.
    On Error GoTo AutoTranslateFailure
    
    'Because this process can take a very long time, warn the user in advance.
    Dim msgReturn As VbMsgBoxResult
    msgReturn = PDMsgBox("This action can take a very long time to complete.  Once started, it cannot be canceled.  Are you sure you want to continue?", vbYesNo Or vbInformation, "Automatic translation warning")

    If (msgReturn <> vbYes) Then Exit Sub
    
    'The user has given the go-ahead, so start translating!
    Dim i As Long
    
    'Start by counting the number of untranslated phrases (so we can provide a status report to the user)
    Dim totalUntranslated As Long, totalTranslated As Long
    totalUntranslated = 0
    totalTranslated = 0
    
    For i = 0 To m_NumOfPhrases - 1
        If (LenB(m_AllPhrases(i).Translation) = 0) Then totalUntranslated = totalUntranslated + 1
    Next i
    
    Dim srcPhrase As String, retString As String
    
    'Iterate through all untranslated phrases, requesting Google translates as we go
    For i = 0 To m_NumOfPhrases - 1
        If (LenB(m_AllPhrases(i).Translation) = 0) Then
        
            'Regardless of whether or not we succeed, increment the counter
            totalTranslated = totalTranslated + 1
            cmdAutoTranslate.Caption = g_Language.TranslateMessage("Processing phrase %1 of %2", totalTranslated, totalUntranslated)
            
            m_AllPhrases(i).IsMachineTranslation = True
            
            'This phrase is not translated.  Apply a minimal amount of preprocessing, then request a translation from Google.
            srcPhrase = m_AllPhrases(i).Original
            
            'Remove ampersands, as Google will treat these as the word "and"
            If InStr(1, srcPhrase, "&", vbBinaryCompare) Then srcPhrase = Replace$(srcPhrase, "&", vbNullString, , , vbBinaryCompare)
            
            'Request a translation from DeepL
            retString = m_AutoTranslate.GetDeepLTranslation(srcPhrase)
            
            'If Google succeeded, store the new translation
            If (LenB(retString) <> 0) Then
                
                'Do a "quick and dirty" case fix for titlecase text
                retString = GetFixedTitlecase(srcPhrase, retString)
                
                'Store the translation
                m_AllPhrases(i).Translation = retString
                
                'Insert this translation into the original XML file
                m_XMLEngine.UpdateTagAtLocation "translation", m_AllPhrases(i).Translation, m_XMLEngine.GetLocationOfParentTag("phrase", "original", m_AllPhrases(i).Original)
    
            End If
            
            'Every sixteen translations, perform an autosave
            If (i And 15) = 0 Then PerformAutosave
            
            'Translations can sometimes get "stuck" (for reasons unknown), so forcibly refresh them after attempting a translation
            srcPhrase = vbNullString
            retString = vbNullString
            
        End If
        
    Next i
    
    cmdAutoTranslate.Caption = g_Language.TranslateMessage("Automatic translation complete!")
    
    'Select the "show untranslated phrases" option, which will refresh the list of untranslated phrases
    cboPhraseFilter.ListIndex = 2
    
    Exit Sub
    
AutoTranslateFailure:
    
    'Auto-save whatever we've translated so far
    PerformAutosave
    
    'Notify the user, then exit
    PDMsgBox "Automatic translations were interrupted (the translation object stopped responding).  Any existing work has been auto-saved.", vbCritical Or vbOKOnly, "Translations interrupted"
    
End Sub

Private Sub cmdCancel_Click()
    
    'Before exiting, save some preference-like values for the user
    UpdateStoredUserValues
    
    Unload Me
    
End Sub

'Allow the user to delete the selected language file, if they so desire.
Private Sub cmdDeleteLanguage_Click()
    
    'Make sure a language is selected
    If (lstLanguages.ListIndex < 0) Then Exit Sub
    
    Dim msgReturn As VbMsgBoxResult

    'Display different warnings for official languages (which can be restored) and user languages (which cannot)
    If Strings.StringsEqual(m_ListOfLanguages(GetLanguageIndexFromListIndex()).LangType, "Official", True) Then
        
        'Make sure we have write access to this folder before attempting to delete anything
        If Files.PathExists(Files.FileGetPath(m_ListOfLanguages(GetLanguageIndexFromListIndex()).FileName), True) Then
        
            msgReturn = PDMsgBox("Are you sure you want to delete %1?" & vbCrLf & vbCrLf & "(This action cannot be undone.  To restore a deleted language file, you must download a fresh copy of PhotoDemon from photodemon.org.)", vbYesNo Or vbExclamation, "Delete language file", lstLanguages.List(lstLanguages.ListIndex))
            
            If (msgReturn = vbYes) Then
                Files.FileDeleteIfExists m_ListOfLanguages(GetLanguageIndexFromListIndex()).FileName
                lstLanguages.RemoveItem lstLanguages.ListIndex
                cmdDeleteLanguage.Enabled = False
            End If
        
        'Write access not available
        Else
            PDMsgBox "You do not have access to this folder.  Please log in as an administrator and try again.", vbOKOnly Or vbExclamation, "Administrator access required"
        End If
    
    'User-folder languages are gone forever once deleted, so change the wording of the deletion confirmation.
    Else
    
        msgReturn = PDMsgBox("Are you sure you want to delete %1?" & vbCrLf & vbCrLf & "(This action cannot be undone.)", vbYesNo Or vbExclamation, "Delete language file", lstLanguages.List(lstLanguages.ListIndex))
        
        If (msgReturn = vbYes) Then
            Files.FileDeleteIfExists m_ListOfLanguages(GetLanguageIndexFromListIndex()).FileName
            lstLanguages.RemoveItem lstLanguages.ListIndex
            cmdDeleteLanguage.Enabled = False
        End If
        
    End If

End Sub

Private Sub cmdNext_Click()
    ChangeWizardPage True
End Sub

Private Sub cmdNextPhrase_Click()
    PhraseFinished
End Sub

Private Sub PhraseFinished()

    If (lstPhrases.ListIndex < 0) Then Exit Sub
    
    'Store this translation to the phrases array
    m_AllPhrases(GetPhraseIndexFromListIndex()).Translation = txtTranslation.Text
    
    'Insert this translation into the original XML file
    m_XMLEngine.UpdateTagAtLocation "translation", txtTranslation, m_XMLEngine.GetLocationOfParentTag("phrase", "original", m_AllPhrases(GetPhraseIndexFromListIndex()).Original)
    
    'Write an alternating backup out to file
    PerformAutosave
        
    'If a specific type of phrase list is displayed, refresh it as necessary
    Dim newIndex As Long
    
    Select Case cboPhraseFilter.ListIndex
    
        'All phrases
        Case 0
        
            newIndex = lstPhrases.ListIndex + 1
            
            'Attempt to automatically move to the next item in the list
            If (newIndex <= lstPhrases.ListCount - 1) Then
                lstPhrases.ListIndex = newIndex
            Else
                If (lstPhrases.ListCount > 0) Then lstPhrases.ListIndex = lstPhrases.ListCount - 1
            End If
        
        'Translated phrases
        Case 1
            
            'If the translation has been erased, this item is no longer part of the "translated phrases" group
            If (LenB(txtTranslation.Text) = 0) Then
                
                newIndex = lstPhrases.ListIndex
                lstPhrases.RemoveItem lstPhrases.ListIndex
                
                'Attempt to automatically move to the next item in the list
                If (newIndex <= lstPhrases.ListCount - 1) Then
                    lstPhrases.ListIndex = newIndex
                Else
                    If (lstPhrases.ListCount > 0) Then lstPhrases.ListIndex = lstPhrases.ListCount - 1
                End If
                
            End If
        
        'Untranslated phrases
        Case 2
        
            'If a translation has been provided, this item is no longer part of the "untranslated phrases" group
            If (LenB(txtTranslation.Text) <> 0) Then
                
                newIndex = lstPhrases.ListIndex
                lstPhrases.RemoveItem lstPhrases.ListIndex
                
                'Attempt to automatically move to the next item in the list
                If (newIndex <= lstPhrases.ListCount - 1) Then
                    lstPhrases.ListIndex = newIndex
                Else
                    If (lstPhrases.ListCount > 0) Then lstPhrases.ListIndex = lstPhrases.ListCount - 1
                End If
                
            End If
        
        '(optional) Phrases that don't match a reference .po
        Case 3
    
            'If the current translation now matches the reference phrases, this item is no longer part of
            ' the "mismatched phrases" group
            If (Not m_PoComparison Is Nothing) Then
                    
                Dim tmpString As String
                m_PoComparison.GetItemByKey LCase$(m_AllPhrases(GetPhraseIndexFromListIndex()).Original), tmpString
                
                If Strings.StringsEqual(tmpString, txtTranslation.Text, True) Then
                    newIndex = lstPhrases.ListIndex
                    lstPhrases.RemoveItem lstPhrases.ListIndex
                Else
                    newIndex = lstPhrases.ListIndex + 1
                End If
                
                'Attempt to automatically move to the next item in the list
                If (newIndex <= lstPhrases.ListCount - 1) Then
                    lstPhrases.ListIndex = newIndex
                Else
                    If (lstPhrases.ListCount > 0) Then lstPhrases.ListIndex = lstPhrases.ListCount - 1
                End If
                
            End If
                
    End Select
    
    UpdatePhraseBoxTitle

End Sub

Private Sub cmdPO_Click()
    
    Dim srcFile As String
    
    'Pull the last-used path from the preferences file
    Dim tempPathString As String
    tempPathString = UserPrefs.GetPref_String("Paths", "reference-po", vbNullString, True)
    
    'If no path was found, default to... idk.  The current PD path?  (There's not an obvious suggestion here.)
    If (LenB(tempPathString) = 0) Then tempPathString = UserPrefs.GetProgramPath()
    
    'Prepare and show a common dialog
    Dim cdFilter As String
    cdFilter = "Language data (.po)|*.po"
    cdFilter = cdFilter & "|" & g_Language.TranslateMessage("All files") & "|*.*"
    
    Dim cdTitle As String
    cdTitle = "Please select a language file"
    
    Dim openDialog As pdOpenSaveDialog
    Set openDialog = New pdOpenSaveDialog
    If openDialog.GetOpenFileName(srcFile, vbNullString, True, False, cdFilter, 1, tempPathString, cdTitle, ".po", Me.hWnd) Then
        
        'Save this new folder as the default path for future usage
        Dim newFolder As String
        newFolder = Files.FileGetPath(srcFile)
        UserPrefs.SetPref_String "Paths", "reference-po", newFolder
        
        'Set the text box to match the selected file, and save said file to the user's pref folder
        txtPO.Text = srcFile
        UserPrefs.SetPref_String "Core", "i18n-po-ref", srcFile
        
    End If
    
    'Re-enable user input
    Interface.EnableUserInput

End Sub

Private Sub cmdPrevious_Click()
    ChangeWizardPage False
End Sub

'Change the active wizard page.  If moveForward is set to TRUE, the wizard page will be advanced; otherwise, it will move
' to the previous page.
Private Sub ChangeWizardPage(ByVal moveForward As Boolean)
    
    Dim i As Long

    Dim unloadFormNow As Boolean
    unloadFormNow = False

    'Before changing the page, maek sure all user input on the current page is valid
    Select Case m_WizardPage
    
        'The first page is the language selection page.  When the user leaves this page, we must load the language
        ' they've selected into memory.
        Case 0
            
            'Before doing anything, save the user's DeepL API key (if any) and reference PO (if any)
            UpdateStoredUserValues
            
            'If the user wants to edit an existing language, make sure they've selected one.  (I hate OK-only message boxes, but am
            ' currently too lazy to write a more elegant warning!)
            If (optBaseLanguage(1).Value And (lstLanguages.ListIndex < 0)) Then
                PDMsgBox "You must select a language file to edit.", vbOKOnly Or vbInformation, "Please select a language"
                Exit Sub
            End If
            
            'Show a brief hourglass while we load and validate the source language file
            Screen.MousePointer = vbHourglass
            
            'If they want to start a new language file from scratch, set the load path to the MASTER English language file
            ' (which is hopefully present... if not, there's not much we can do.)
            If optBaseLanguage(0).Value Then
                                
                If LoadAllPhrasesFromFile(UserPrefs.GetLanguagePath & "Master\MASTER.xml") Then
                    
                    'Populate the current language's metadata container with some default values
                    With m_curLanguage
                        .FileName = UserPrefs.GetLanguagePath(True) & "new language.xml"
                        .langID = "en-US"
                        .LangName = g_Language.TranslateMessage("New Language")
                        .LangStatus = g_Language.TranslateMessage("incomplete")
                        .LangType = "Unofficial"
                        .langVersion = "1.0.0"
                        .Author = g_Language.TranslateMessage("enter your name here")
                    End With
                                        
                'For some reason, we failed to load the master language file.  Tell them to download a fresh copy of PD.
                Else
                    Screen.MousePointer = vbDefault
                    PDMsgBox "Unfortunately, PhotoDemon's en-US language file could not be located on this PC.  This file is included with the official release of PhotoDemon, but it may not be included with development or beta builds." & vbCrLf & vbCrLf & "To start a new translation, please download a fresh copy of PhotoDemon from photodemon.org.", vbOKOnly Or vbExclamation, "Master language file missing"
                    Unload Me
                End If
            
            'They want to edit an existing language.  Follow the same general pattern as for the master language file (above).
            Else
            
                'Fill the current language metadata container with matching information from the selected language,
                ' with a few changes
                m_curLanguage = m_ListOfLanguages(GetLanguageIndexFromListIndex())
                m_curLanguage.FileName = UserPrefs.GetLanguagePath(True) & Files.FileGetName(m_ListOfLanguages(GetLanguageIndexFromListIndex()).FileName)
                
                'Attempt to load the selected language from file
                If LoadAllPhrasesFromFile(m_ListOfLanguages(GetLanguageIndexFromListIndex()).FileName) Then
                    
                    'No further action is necessary!
                    
                'For some reason, we failed to load the master language file.  Tell them to download a fresh copy of PD.
                Else
                    Screen.MousePointer = vbDefault
                    PDMsgBox "Unfortunately, this language file could not be loaded.  It's possible the copy on this PC is out-of-date." & vbCrLf & vbCrLf & "To continue, please download a fresh copy of PhotoDemon from photodemon.org.", vbOKOnly Or vbExclamation, "Language file could not be loaded"
                    Unload Me
                End If
            
            End If
            
            'If the user selected a 3rd-party .po file, parse it now so we can quickly compare translations
            LoadReferencePO
            
            'If the user supplied a reference .po, and 1+ phrases were loaded from it, add a new listbox
            ' option in the translation panel for "phrases that don't match reference".
            If (Not m_PoComparison Is Nothing) Then
                If (m_PoComparison.GetNumOfItems > 0) Then
                    If (cboPhraseFilter.ListCount <= 3) Then
                        cboPhraseFilter.AddItem "phrases that don't match reference"
                    End If
                End If
            End If
            
            'If the user didn't supply a DeepL API key, hide the "auto translate" checkbox
            chkOnlineTranslate.Value = (LenB(txtApiKey.Text) <> 0)
            
            'Reset the mouse pointer
            Screen.MousePointer = vbDefault
            
        'The second page is the metadata editing page.
        Case 1
        
            'When leaving the metadata page, automatically copy all text box entries into the metadata holder
            With m_curLanguage
                .langID = Trim$(txtLangID(0)) & "-" & Trim$(txtLangID(1))
                .LangName = Trim$(txtLangName)
                .LangStatus = Trim$(txtLangStatus)
                .langVersion = Trim$(txtLangVersion)
                .Author = Trim$(txtLangAuthor)
            End With
            
            'Also, automatically set the destination language of the Google Translate interface
            m_AutoTranslate.SetDstLanguage Trim$(txtLangID(0))
            
            'Write these updated tags into the original XML text
            With m_curLanguage
                m_XMLEngine.UpdateTagAtLocation "langid", .langID
                m_XMLEngine.UpdateTagAtLocation "langname", .LangName
                m_XMLEngine.UpdateTagAtLocation "langstatus", .LangStatus
                m_XMLEngine.UpdateTagAtLocation "langversion", .langVersion
                m_XMLEngine.UpdateTagAtLocation "author", .Author
            End With
            
            'Update the autosave file
            PerformAutosave
            
            'Hide the reference translation if a .po wasn't provided
            Dim showReference As Boolean
            showReference = (Not m_PoComparison Is Nothing)
            If showReference Then showReference = (m_PoComparison.GetNumOfItems() > 0)
            lblTitle(11).Visible = showReference
            txtReference.Visible = showReference
            cmdUseReference.Visible = showReference
        
        'The third page is the phrase editing page.  This is the most important page in the wizard.
        Case 2
        
            If moveForward Then
                
                'If the user is working from an official file or an autosave, the folder and/or extension of the
                ' original filename may not be usable.  Strip just the original filename, and append our own
                ' folder and extension.
                Dim sFile As String
                
                If m_curLanguage.LangType = "Autosave" Then
                    sFile = Files.FileMakeNameValid(m_curLanguage.LangName)
                    sFile = Files.FileGetPath(m_curLanguage.FileName) & sFile & ".xml"
                Else
                    sFile = Files.FileGetPath(m_curLanguage.FileName) & Files.FileGetName(m_curLanguage.FileName, True) & ".xml"
                End If
                
                Dim cdFilter As String
                cdFilter = g_Language.TranslateMessage("XML file") & " (.xml)|*.xml"
                
                'On this page, the "Next" button is relabeled as "Save and Exit".  It does exactly what it claims!
                Dim saveDialog As pdOpenSaveDialog
                Set saveDialog = New pdOpenSaveDialog
                
                If saveDialog.GetSaveFileName(sFile, , True, cdFilter, , Files.FileGetPath(sFile), g_Language.TranslateMessage("Save current language file"), ".xml", Me.hWnd) Then
                
                    'Write the current XML file out to the user's requested path
                    m_XMLEngine.WriteXMLToFile sFile, True
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
        m_WizardPage = m_WizardPage + 1
    Else
        m_WizardPage = m_WizardPage - 1
    End If
    
    'We can now apply any entrance-timed panel changes
    Select Case m_WizardPage
    
        'Language selection
        Case 0
        
            'Fill the available languages list box with any language files on this system
            PopulateAvailableLanguages
        
        'Metadata editor
        Case 1
        
            'When entering the metadata page, automatically fill all boxes with the currently stored metadata entries
            With m_curLanguage
            
                'Language ID is the most complex, because we must parse the two halves into individual text boxes
                If (InStr(1, .langID, "-") > 0) Then
                    txtLangID(0) = Left$(.langID, InStr(1, .langID, "-") - 1)
                    txtLangID(1) = Mid$(.langID, InStr(1, .langID, "-") + 1, Len(.langID) - InStr(1, .langID, "-"))
                Else
                    txtLangID(0) = .langID
                    txtLangID(1) = vbNullString
                End If
                
                'Everything else can be copied directly
                txtLangName = .LangName
                txtLangStatus = .LangStatus
                txtLangVersion = .langVersion
                txtLangAuthor = .Author
                
            End With
        
        'Phrase editor
        Case 2
        
            'If an XML file was successfully loaded, add its contents to the list box
            If (Not m_xmlLoaded) Then
            
                m_xmlLoaded = True
                
                'Setting the ListIndex property will fire the _Click event, which will handle the actual phrase population
                cboPhraseFilter.ListIndex = 0
                cboPhraseFilter_Click
                
            End If
                
    End Select
    
    'Hide all inactive panels (and show the active one)
    For i = 0 To picContainer.Count - 1
        picContainer(i).Visible = (i = m_WizardPage)
    Next i
    
    'If we are at the beginning, disable the previous button
    cmdPrevious.Enabled = (m_WizardPage <> 0)
    
    'If we are at the end, change the text of the "next" button; otherwise, make sure it says "next"
    If (m_WizardPage = picContainer.Count - 1) Then
        cmdNext.Caption = g_Language.TranslateMessage("Save and Exit")
    Else
        cmdNext.Caption = g_Language.TranslateMessage("Next")
    End If
    
    'Finally, change the top title caption and left-hand help text to match the current step
    lblWizardTitle.Caption = g_Language.TranslateMessage("Step %1:", m_WizardPage + 1)
    lblWizardTitle.Caption = lblWizardTitle.Caption & " "
    
    Dim helpText As pdString
    Set helpText = New pdString
    
    Select Case m_WizardPage
    
        Case 0
            lblWizardTitle.Caption = "step 1: select a language file"
            
            helpText.AppendLine "This tool allows you to create and edit PhotoDemon language files." & vbCrLf
            helpText.AppendLine "Please start by selecting a base language file.  If the selected file already contains translation data, you can edit existing translations or fill-in missing ones." & vbCrLf
            helpText.AppendLine "This page also allows you to delete unused language files.  There is no Undo option when deleting language files, so please be careful!" & vbCrLf
            helpText.Append "(When you click ""Next"", the selected language file will parsed and validated.  This process may take several seconds.)"
            
        Case 1
            lblWizardTitle.Caption = "step 2: add language metadata"
            
            helpText.AppendLine "PhotoDemon needs a little bit of metadata for each language file.  This allows it to auto-select the most relevant language based on the locale of a user's PC." & vbCrLf
            helpText.AppendLine "The most important items on this page are the language ID and language name.  Please double-check to ensure these are correct." & vbCrLf
            helpText.AppendLine "If multiple translators have worked on this language file, please separate their names with commas.  If this language file is based on an existing language file, please retain the original author's name." & vbCrLf
            helpText.Append "(NOTE: changes made to this page won't be auto-saved until you click the Next or Previous button.)"
            
        Case 2
            lblWizardTitle.Caption = "step 3: localize phrases"
            
            helpText.AppendLine "This final step allows you to edit individual phrases." & vbCrLf
            helpText.AppendLine "Every time a phrase is modified, an autosave is automatically created in PhotoDemon's /Data/Languages folder.  This means you can exit at any time without losing your work." & vbCrLf
            helpText.AppendLine "When you are done translating, use the Save and Exit button to save your work.  (Autosave data will be preserved either way.)" & vbCrLf
            helpText.AppendLine "When you finish editing this language, please send it to me!  My current contact information is available at:" & vbCrLf
            helpText.AppendLine "https://photodemon.org/about/" & vbCrLf
            helpText.Append "Even partial translations are helpful.  Thank you."
    
    End Select
    
    lblExplanation.Caption = helpText.ToString()
        
End Sub

Private Sub cmdUseReference_Click()
    If Strings.StringsNotEqual(txtTranslation.Text, txtReference.Text, False) Then lblTranslatedPhrase.Caption = "translated phrase (NOT YET SAVED):"
    txtTranslation.Text = txtReference.Text
End Sub

Private Sub Form_Load()
    
    'Mark the XML file as not loaded
    m_xmlLoaded = False
    m_curBackupFile = 0
    
    'By default, the first wizard page is displayed.  (We start at -1 because we will incerement the page count by +1 with our first
    ' call to changeWizardPage in Form_Activate)
    m_WizardPage = -1
    
    'Fill the "phrases to display" combo box
    cboPhraseFilter.Clear
    cboPhraseFilter.AddItem "all phrases", 0
    cboPhraseFilter.AddItem "translated phrases", 1
    cboPhraseFilter.AddItem "untranslated phrases", 2
    cboPhraseFilter.ListIndex = 0
    
    'Initialize an online translation interface
    Set m_AutoTranslate = New pdAutoLocalize
    m_AutoTranslate.SetSrcLanguage "en"
    
    'Note that the user must supply their own API key; I do not ship mine with the project
    Dim userKey As String
    userKey = UserPrefs.GetPref_String("Core", "DeepL-API", vbNullString, False)
    If (LenB(userKey) <> 0) Then
        txtApiKey.Text = userKey
        m_AutoTranslate.SetAPIKey Trim$(userKey)
    End If
    
    'Same goes for a reference .po, if any
    userKey = UserPrefs.GetPref_String("Core", "i18n-po-ref", vbNullString, False)
    If (LenB(userKey) <> 0) Then txtPO.Text = userKey
    
    'Apply translations and visual styles
    ApplyThemeAndTranslations Me
    
    'Advance to the first page
    ChangeWizardPage True
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

'Given a source language file, find all phrase tags, and load them into a specialized phrase array
Private Function LoadAllPhrasesFromFile(ByVal srcLangFile As String) As Boolean
    
    LoadAllPhrasesFromFile = False
    
    Set m_XMLEngine = New pdXML
    
    'Attempt to load the language file
    If m_XMLEngine.LoadXMLFile(srcLangFile) Then
    
        'Validate the language file's contents
        If m_XMLEngine.IsPDDataType("Translation") And m_XMLEngine.ValidateLoadedXMLData("phrase") Then
        
            'New as of August '14 is the ability to set text comparison mode.  To ensure output matches
            ' the rest of PD, the language editor now uses binary comparison mode exclusively.
            m_XMLEngine.SetTextCompareMode vbBinaryCompare
        
            'Attempt to load all phrase tag location occurrences
            Dim phraseLocations() As Long
            If m_XMLEngine.FindAllTagLocations(phraseLocations, "phrase", True) Then
                
                m_NumOfPhrases = UBound(phraseLocations) + 1
                ReDim m_AllPhrases(0 To m_NumOfPhrases - 1) As PD_Phrase
                
                Dim tmpString As String
                
                Dim i As Long
                For i = 0 To m_NumOfPhrases - 1
                    tmpString = m_XMLEngine.GetUniqueTag_String("original", vbNullString, phraseLocations(i))
                    m_AllPhrases(i).Original = tmpString
                    m_AllPhrases(i).Length = Len(tmpString)
                    m_AllPhrases(i).Translation = m_XMLEngine.GetUniqueTag_String("translation", vbNullString, phraseLocations(i) + Len(tmpString))
                    
                    'We also need a modified version of the string to add to the phrase list box.  This text can't include line breaks,
                    ' and it can't be so long that it overflows the list box.
                    If (InStr(1, tmpString, vbCr) > 0) Then tmpString = Replace(tmpString, vbCr, vbNullString)
                    If (InStr(1, tmpString, vbLf) > 0) Then tmpString = Replace(tmpString, vbLf, vbNullString)
                    m_AllPhrases(i).ListBoxEntry = tmpString
                    
                Next i
                
                LoadAllPhrasesFromFile = True
            
            '/Failed to find any phrases in the file
            End If
        
        '/Failed to validate XML
        End If
    
    '/Failed to load XML
    End If

End Function

Private Sub lstLanguages_Click()
    If (Not optBaseLanguage(1).Value) Then optBaseLanguage(1).Value = True
    cmdDeleteLanguage.Enabled = (lstLanguages.ListIndex >= 0)
End Sub

'When the phrase box is clicked, display the original and translated (if available) text in the right-hand text boxes
Private Sub lstPhrases_Click()
    
    txtOriginal.Text = m_AllPhrases(GetPhraseIndexFromListIndex()).Original
    
    'If a translation exists for this phrase, load it.  If it does not, use Google Translate to estimate a translation
    ' (contingent on the relevant check box setting)
    lblTranslatedPhrase.Caption = "translated phrase"
    
    If (LenB(m_AllPhrases(GetPhraseIndexFromListIndex()).Translation) <> 0) Then
        txtTranslation.Text = m_AllPhrases(GetPhraseIndexFromListIndex()).Translation
        lblTranslatedPhrase.Caption = lblTranslatedPhrase.Caption & " " & g_Language.TranslateMessage("(saved):")
    Else
    
        lblTranslatedPhrase.Caption = lblTranslatedPhrase.Caption & " " & g_Language.TranslateMessage("(NOT YET SAVED):")
        
        If chkOnlineTranslate.Value Then
        
            txtTranslation.Text = g_Language.TranslateMessage("waiting...")
            
            'Pull a phrase from the online service, if one is available
            Dim retString As String
            retString = m_AutoTranslate.GetDeepLTranslation(m_AllPhrases(GetPhraseIndexFromListIndex()).Original)
            
            If (LenB(retString) <> 0) Then
                txtTranslation.Text = GetFixedTitlecase(m_AllPhrases(GetPhraseIndexFromListIndex()).Original, retString)
            Else
                If (LenB(m_AutoTranslate.GetAPIKey()) > 0) Then
                    txtTranslation.Text = "[translation failed]"
                Else
                    txtTranslation.Text = vbNullString
                End If
            End If
            
        Else
            txtTranslation.Text = vbNullString
        End If
            
    End If
    
    'If a .po reference was provided, look for this text there too.  Do this regardless of whether we have an
    ' existing translation or not
    If (Not m_PoComparison Is Nothing) And txtReference.Visible Then
        
        Dim strReference As String
        If m_PoComparison.GetItemByKey(LCase$(m_AllPhrases(GetPhraseIndexFromListIndex()).Original), strReference) Then
            txtReference.Text = strReference
        Else
            txtReference.Text = "[phrase not in file]"
        End If
        
    End If
    
        
End Sub

Private Sub optBaseLanguage_Click(Index As Integer)
    cmdDeleteLanguage.Enabled = (lstLanguages.ListIndex >= 0)
End Sub

'The phrase list box label will automatically be updated with the current count of list items
Private Sub UpdatePhraseBoxTitle()
    Dim numPhrasesDisplay As Long
    If (lstPhrases.ListCount > 0) Then numPhrasesDisplay = lstPhrases.ListCount - 1 Else numPhrasesDisplay = 0
    lblPhraseBox.Caption = g_Language.TranslateMessage("list of phrases (%1 items)", numPhrasesDisplay)
End Sub

'Call this function whenever we want the in-memory XML data saved to an autosave file
Private Sub PerformAutosave()

    'We keep two autosaves at all times; simply alternate between them each time a save is requested
    If (m_curBackupFile = 1) Then m_curBackupFile = 0 Else m_curBackupFile = 1
    
    'Generate an autosave filename.  The language ID is appended to the name, so separate autosaves will exist for each edited language
    ' (assuming they have different language IDs).
    Dim backupFile As String
    backupFile = UserPrefs.GetLanguagePath(True) & backupFileName & m_curLanguage.langID & "_" & Trim$(Str$(m_curBackupFile)) & ".tmpxml"
    
    'The XML engine handles the actual writing to file.  For performance reasons, auto-tabbing is suppressed.
    m_XMLEngine.WriteXMLToFile backupFile, True

End Sub

'Fill the first panel ("select a language file") with all available language files on this system
Private Sub PopulateAvailableLanguages()
    
    'Retrieve a list of available languages from the translation engine
    g_Language.CopyListOfLanguages m_ListOfLanguages
    
    'We now do a bit of additional work.  Look for any autosave files (with extension .tmpxml) in the user language folder.  Allow the
    ' user to load these if available.
    Dim listOfTmpXML As pdStringStack
    Set listOfTmpXML = New pdStringStack
    If Files.RetrieveAllFiles(UserPrefs.GetLanguagePath(True), listOfTmpXML, False, True, "tmpxml") Then
        
        Dim chkFile As String
        Do While listOfTmpXML.PopString(chkFile)
            
            'Use PD's XML engine to load the file
            Dim tmpXML As pdXML
            Set tmpXML = New pdXML
            If tmpXML.LoadXMLFile(UserPrefs.GetLanguagePath(True) & chkFile) Then
            
                'Use the XML engine to validate this file, and to make sure it contains at least a language ID, name, and one (or more) translated phrase
                If tmpXML.IsPDDataType("Translation") And tmpXML.ValidateLoadedXMLData("langid", "langname", "phrase") Then
                
                    ReDim Preserve m_ListOfLanguages(0 To UBound(m_ListOfLanguages) + 1) As PDLanguageFile
                    
                    With m_ListOfLanguages(UBound(m_ListOfLanguages))
                    
                        'Get the language ID and name - these are the most important values, and technically the only REQUIRED ones.
                        .langID = tmpXML.GetUniqueTag_String("langid")
                        .LangName = tmpXML.GetUniqueTag_String("langname")
        
                        'Version, status, and author information should also be present, but the file will still be loaded even if they don't exist
                        .langVersion = tmpXML.GetUniqueTag_String("langversion")
                        .LangStatus = tmpXML.GetUniqueTag_String("langstatus")
                        .Author = tmpXML.GetUniqueTag_String("author")
                        
                        'Finally, add some internal metadata
                        .FileName = UserPrefs.GetLanguagePath(True) & chkFile
                        .LangType = "Autosave"
                        
                    End With
                    
                End If
                
            End If
            
        Loop
        
    End If
    
    'All autosave files have now been loaded as well
    
    'Add the contents of that array to the list box on the opening panel (the list of available languages, from which the user
    ' can select a language file as the "starting point" for their own translation).
    lstLanguages.Clear
    
    Dim i As Long
    For i = 0 To UBound(m_ListOfLanguages)
    
        'Note that we DO NOT add the English language entry - that is used by the "start a new language file from scratch" option.
        If Strings.StringsNotEqual(m_ListOfLanguages(i).LangType, "DEFAULT", True) Then
            Dim listEntry As String
            listEntry = m_ListOfLanguages(i).LangName
            
            'For official translations, an author name will always be provided.  Include the author's name in the list.
            If (m_ListOfLanguages(i).LangType = "Official") Then
                listEntry = listEntry & " ("
                listEntry = listEntry & g_Language.TranslateMessage("official translation by")
                listEntry = listEntry & " " & m_ListOfLanguages(i).Author
                listEntry = listEntry & ")"
            
            'For unofficial translations, an author name may not be provided.  Include the author's name only if it's available.
            ElseIf (m_ListOfLanguages(i).LangType = "Unofficial") Then
                listEntry = listEntry & " "
                listEntry = listEntry & g_Language.TranslateMessage("by")
                listEntry = listEntry & " "
                If (LenB(m_ListOfLanguages(i).Author) <> 0) Then
                    listEntry = listEntry & m_ListOfLanguages(i).Author
                Else
                    listEntry = listEntry & g_Language.TranslateMessage("unknown author")
                End If
                
            'Anything else is an autosave.
            Else
            
                'Include author name if available
                listEntry = listEntry & " "
                listEntry = listEntry & g_Language.TranslateMessage("by")
                listEntry = listEntry & " "
                If (LenB(m_ListOfLanguages(i).Author) <> 0) Then
                    listEntry = listEntry & m_ListOfLanguages(i).Author
                Else
                    listEntry = listEntry & g_Language.TranslateMessage("unknown author")
                End If
                
                'Display autosave time and date
                listEntry = listEntry & " ("
                listEntry = listEntry & g_Language.TranslateMessage("autosaved on")
                listEntry = listEntry & " "
                listEntry = listEntry & Format$(FileDateTime(m_ListOfLanguages(i).FileName), "hh:mm:ss AM/PM, dd-mmm-yy")
                listEntry = listEntry & ") "
            
            End If
            
            'To save us time in the future, use the .ItemData property of this entry to store the language's original index position
            ' in our m_ListOfLanguages array.
            lstLanguages.AddItem listEntry
            m_ListOfLanguages(i).InternalDisplayName = listEntry
            
        Else
            'Ignore the default language entry entirely
        End If
    Next i
    
    'By default, no language is selected for the user
    lstLanguages.ListIndex = -1
    
End Sub

Private Function GetLanguageIndexFromListIndex() As Long
    Dim i As Long
    For i = LBound(m_ListOfLanguages) To UBound(m_ListOfLanguages)
        If Strings.StringsEqual(lstLanguages.List(lstLanguages.ListIndex), m_ListOfLanguages(i).InternalDisplayName) Then
            GetLanguageIndexFromListIndex = i
            Exit For
        End If
    Next i
End Function

Private Function GetPhraseIndexFromListIndex() As Long
    Dim i As Long
    For i = LBound(m_AllPhrases) To UBound(m_AllPhrases)
        If Strings.StringsEqual(lstPhrases.List(lstPhrases.ListIndex), m_AllPhrases(i).ListBoxEntry) Then
            GetPhraseIndexFromListIndex = i
            Exit For
        End If
    Next i
End Function

'On Win 7+, we attempt to automatically handle titlecase of translated text.  If the original English string used titlecase,
' we'll set titlecase to the translated string as well.
Private Function GetFixedTitlecase(ByVal origString As String, ByVal translatedString As String) As String
    
    On Error GoTo TitlecaseFail
    
    If (LenB(origString) <> 0) And (LenB(translatedString) <> 0) Then
    
        If OS.IsWin7OrLater Then
            
            Dim origStringTitlecase As Boolean
            
            'Split the original string into individual words
            Dim strOrig() As String, strTranslated() As String
            strOrig = Split(origString, " ", , vbBinaryCompare)
            strTranslated = Split(translatedString, " ", , vbBinaryCompare)
            
            'Only proceed with automatic casing if *both* strings contain multiple words.  (Some translations may not
            ' result in 1:1 word counts.)
            Dim multWords As Boolean
            multWords = (UBound(strOrig) <> 0)
            If multWords Then multWords = (UBound(strTranslated) <> 0)
            
            'If the text involves multiple words, we only want to titlecase the first word in the string
            If multWords Then
            
                'Split out the first word in the string
                Dim firstWord As String, firstWordIndex As Long
                
                Dim i As Long
                For i = LBound(strOrig) To UBound(strOrig)
                    If (LenB(Trim$(strOrig(i))) <> 0) Then
                        firstWord = strOrig(i)
                        firstWordIndex = i
                        Exit For
                    End If
                Next i
                
                'See if the first word used titlecase
                origStringTitlecase = Strings.StringsEqual(firstWord, Strings.StringRemap(firstWord, sr_Titlecase), False)
                
                'If it did, apply titlecase to the first word of the translated string as well
                If origStringTitlecase Then
                    
                    'Find the first word in the translation and titlecase it
                    For i = LBound(strTranslated) To UBound(strTranslated)
                        If (LenB(Trim$(strTranslated(i))) <> 0) Then
                            firstWord = strTranslated(i)
                            firstWordIndex = i
                            Exit For
                        End If
                    Next i
                    
                    Dim tmpString As String
                    tmpString = Strings.StringRemap(firstWord, sr_Titlecase)
                    
                    If (LenB(tmpString) <> 0) Then
                    
                        strTranslated(firstWordIndex) = tmpString
                        
                        'Reassemble the translated string
                        GetFixedTitlecase = vbNullString
                        
                        For i = LBound(strTranslated) To UBound(strTranslated)
                            GetFixedTitlecase = GetFixedTitlecase & strTranslated(i)
                            If (i < UBound(strTranslated)) Then GetFixedTitlecase = GetFixedTitlecase & " "
                        Next i
                        
                    Else
                        GetFixedTitlecase = translatedString
                    End If
                    
                Else
                    GetFixedTitlecase = translatedString
                End If
                
            'Single-word case is quite a bit easier to handle
            Else
            
                'See if the original string used titlecase
                origStringTitlecase = Strings.StringsEqual(origString, Strings.StringRemap(origString, sr_Titlecase), False)
                
                'If it did, apply titlecase to the translated string as well
                If origStringTitlecase Then
                    GetFixedTitlecase = Strings.StringRemap(translatedString, sr_Titlecase)
                Else
                    GetFixedTitlecase = translatedString
                End If
                
            End If
            
        Else
            GetFixedTitlecase = translatedString
        End If
        
    End If
    
    Exit Function
    
TitlecaseFail:

    Debug.Print "WARNING!  Titlecase failed on string: " & origString
    Debug.Print "Attempted translation was: " & translatedString
    GetFixedTitlecase = translatedString

End Function

Private Sub txtTranslation_KeyDown(ByVal Shift As ShiftConstants, ByVal vKey As Long, preventFurtherHandling As Boolean)

    If (vKey = vbKeyReturn) And (Shift And vbCtrlMask = vbCtrlMask) Then
        preventFurtherHandling = True
        m_InKeyEvent = True
        PhraseFinished
        txtTranslation.SelStart = Len(txtTranslation.Text)
    Else
        m_InKeyEvent = False
    End If

End Sub

Private Sub txtTranslation_KeyPress(ByVal Shift As ShiftConstants, ByVal vKey As Long, preventFurtherHandling As Boolean)
    preventFurtherHandling = m_InKeyEvent
End Sub

'Call this to save any relevant form-level data to the prefs file (current includes the user's DeepL API key
' and reference po, if any)
Private Sub UpdateStoredUserValues()
    
    'DeepL API key
    If Strings.StringsNotEqual(Trim$(txtApiKey.Text), UserPrefs.GetPref_String("Core", "DeepL-API", vbNullString, True), True) Then
        UserPrefs.SetPref_String "Core", "DeepL-API", Trim$(txtApiKey.Text)
    End If
    
    'Reference .po
    If Strings.StringsNotEqual(Trim$(txtPO.Text), UserPrefs.GetPref_String("Core", "i18n-po-ref", vbNullString, True), True) Then
        UserPrefs.SetPref_String "Core", "i18n-po-ref", Trim$(txtPO.Text)
    End If
    
End Sub

'Manually load all phrases from a target .po file.  Translations from e.g. GIMP are extremely helpful when trying
' to figure out which of several translations is the "best" one for a given term.  It only works for exact phrase
' matches right now, but fuzzy matches could be an interesting project in the future.
'
'This .po parser is a quick-and-dirty implementation, but it's really damn fast, retrieving and hashing
' 5,500+ phrase files in < 0.1 s.
Private Sub LoadReferencePO()
    
    'The target .po file, if any, needs to be validated before parsing
    Dim srcFile As String
    srcFile = txtPO.Text
    
    If (LenB(srcFile) = 0) Then Exit Sub
    If (Not Files.FileExists(srcFile)) Then Exit Sub
    
    'All modern .po files should be UTF-8; PD doesn't attempt to handle other varieties
    Dim srcText As String
    If (Not Files.FileLoadAsString(srcFile, srcText, True)) Then Exit Sub
    
    'You know my motto: profile everything!
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    'Quick sanity check for expected gettext markers
    Const MSG_ID As String = "msgid """
    Const MSG_STR As String = "msgstr """
    If (InStr(1, srcText, MSG_ID) = 0) Then Exit Sub
    If (InStr(1, srcText, MSG_STR) = 0) Then Exit Sub
    
    Const QUOTE_CHAR As String = """"
    Const SPACE_CHAR As String = " "
    Const DOUBLE_LINEBREAK As String = vbCrLf & vbCrLf
    Const UNDERSCORE_CHAR As String = "_"
    Const ELLIPSIS As String = "..."
    Dim ELLIPSIS_CHAR As String: ELLIPSIS_CHAR = ChrW$(&H2026)
    
    'Looks like this is a normal .po file.  We now want to load all phrases (ugh) into some kind of dictionary,
    ' so we can query our own translations against theirs.
    
    'For now, a simple hash table will suffice.  If I decide to explore fuzzy matching in the future,
    ' a different solution may be better.
    Set m_PoComparison = New pdStringHash
    
    'Parsing the .po text currently uses a "fast and dirty" approach.  We basically just scan, looking for
    ' key/value pairs and ignoring any clues, whitespace, etc.
    Dim idxStart As Long, idxEnd As Long, curPhrase As String
    idxStart = InStr(1, srcText, MSG_ID, vbBinaryCompare)
    
    Dim msgID As String, msgStr As String
    
    Do While (idxStart > 0)
        
        'Null strings are important for detecting parser failures
        msgID = vbNullString
        msgStr = vbNullString
        
        'idxStart points at the start of a msgid line.  Find the msgstr that follows it.
        idxStart = idxStart + Len(MSG_ID)
        idxEnd = InStr(idxStart, srcText, MSG_STR, vbBinaryCompare)
        
        'We now have enough info to construct this phrase
        If (idxEnd > idxStart) Then
            
            'Grab everything between the id tags
            curPhrase = Mid$(srcText, idxStart, (idxEnd - idxStart) - 3)
            
            'Messages always start and end with quotation marks.  Remove 'em.
            If (Left$(curPhrase, 1) = QUOTE_CHAR) Then curPhrase = Right$(curPhrase, Len(curPhrase) - 1)
            If (Right$(curPhrase, 1) = QUOTE_CHAR) Then curPhrase = Left$(curPhrase, Len(curPhrase) - 1)
            
            'Messages can be multiline.  Look for line breaks in the text and remove them if found.
            If (InStr(1, curPhrase, vbCrLf, vbBinaryCompare) <> 0) Then
                
                'Replace valid quotes with a placeholder
                curPhrase = Replace$(curPhrase, "\""", "&quot;")
                
                'Remove linebreaks
                curPhrase = Replace$(curPhrase, vbCrLf, vbNullString)
                
                'Remove any remaining quotes
                curPhrase = Replace$(curPhrase, """", vbNullString)
                
                'Restore valid quotes (that we hacked out at the beginning)
                curPhrase = Replace$(curPhrase, "&quot;", """")
                
            End If
            
        End If
        
        'We now have the ID phrase.  Store it, because we've got more parsing to do.
        msgID = curPhrase
        
        '(Note that many .po files start with a blank tag followed by metadata; ignore those tags.)
        If (LenB(msgID) > 0) Then
            
            'Time to repeat the above steps, but this time for the translated text.
            idxStart = InStr(idxStart + Len(msgID), srcText, MSG_STR, vbBinaryCompare) + Len(MSG_STR)
            idxEnd = InStr(idxStart, srcText, DOUBLE_LINEBREAK, vbBinaryCompare) - 1
            If (idxEnd < 0) Then idxEnd = Len(srcText)
            
            'We now have enough info to construct this phrase
            If (idxEnd > idxStart) Then
                
                curPhrase = Mid$(srcText, idxStart, (idxEnd - idxStart))
                
                'Messages always start and end with quotation marks.  Remove 'em.
                If (Left$(curPhrase, 1) = QUOTE_CHAR) Then curPhrase = Right$(curPhrase, Len(curPhrase) - 1)
                If (Right$(curPhrase, 1) = QUOTE_CHAR) Then curPhrase = Left$(curPhrase, Len(curPhrase) - 1)
                
                'Messages can be multiline.  Look for line breaks in the text and remove them if found.
                If (InStr(1, curPhrase, vbCrLf, vbBinaryCompare) <> 0) Then
                    
                    'Replace valid quotes with a placeholder
                    curPhrase = Replace$(curPhrase, "\""", "&quot;")
                    
                    'Remove linebreaks
                    curPhrase = Replace$(curPhrase, vbCrLf, vbNullString)
                    
                    'Remove any remaining quotes
                    curPhrase = Replace$(curPhrase, """", vbNullString)
                    
                    'Restore valid quotes (that we hacked out at the beginning)
                    curPhrase = Replace$(curPhrase, "&quot;", """")
                    
                End If
                
                'Debug.Print "-----" & vbCrLf & curPhrase & vbCrLf & "-----"
                
            End If
            
            'We now have the translation.
            msgStr = curPhrase
            
            'Only store key+value pairs where both entities exist
            If (LenB(msgID) > 0) Then
                
                'Do a little pre-processing to both strings.  In particular, we don't want trailing ellipses
                ' or markers for hotkeys (typically _); we're only interested in the actual text
                If (InStr(1, msgID, ELLIPSIS) <> 0) Then msgID = Replace$(msgID, ELLIPSIS, vbNullString)
                If (InStr(1, msgID, ELLIPSIS_CHAR) <> 0) Then msgID = Replace$(msgID, ELLIPSIS_CHAR, vbNullString)
                If (InStr(1, msgID, UNDERSCORE_CHAR) <> 0) Then msgID = Replace$(msgID, UNDERSCORE_CHAR, vbNullString)
                
                If (InStr(1, msgStr, ELLIPSIS) <> 0) Then msgStr = Replace$(msgStr, ELLIPSIS, vbNullString)
                If (InStr(1, msgStr, ELLIPSIS_CHAR) <> 0) Then msgStr = Replace$(msgStr, ELLIPSIS_CHAR, vbNullString)
                If (InStr(1, msgStr, UNDERSCORE_CHAR) <> 0) Then msgStr = Replace$(msgStr, UNDERSCORE_CHAR, vbNullString)
                
                'We want case-insensitive matching, so deliberately lcase all keys
                m_PoComparison.AddItem LCase$(msgID), msgStr
                
            End If
            
        End If
        
        'Look for the next phrase
        If (idxEnd > 0) Then idxStart = InStr(idxEnd + 1, srcText, MSG_ID, vbBinaryCompare) Else idxStart = -1
        
    Loop
    
    If (Not m_PoComparison Is Nothing) Then PDDebug.LogAction "Loaded " & m_PoComparison.GetNumOfItems() & " phrases from the reference .po in " & VBHacks.GetTimeDiffNowAsString(startTime)
    
End Sub
