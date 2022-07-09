VERSION 5.00
Begin VB.Form FormLanguageEditor 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Language editor"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12060
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
   ScaleWidth      =   804
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdHyperlink hypReadme 
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   8445
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   661
      Alignment       =   2
      Caption         =   "click here for detailed instructions"
      URL             =   "https://github.com/tannerhelland/PhotoDemon/tree/main/App/PhotoDemon/Languages#readme"
   End
   Begin PhotoDemon.pdButton cmdPrevious 
      Height          =   615
      Left            =   6720
      TabIndex        =   2
      Top             =   8310
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   1085
      Caption         =   "Previous"
   End
   Begin PhotoDemon.pdButton cmdNext 
      Height          =   615
      Left            =   8520
      TabIndex        =   16
      Top             =   8310
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   1085
      Caption         =   "Next"
   End
   Begin PhotoDemon.pdButton cmdCancel 
      Height          =   615
      Left            =   10500
      TabIndex        =   17
      Top             =   8310
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1085
      Caption         =   "Cancel"
   End
   Begin PhotoDemon.pdLabel lblWizardTitle 
      Height          =   405
      Left            =   120
      Top             =   120
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   714
      Caption         =   ""
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   7455
      Index           =   2
      Left            =   120
      Top             =   720
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   13150
      Begin PhotoDemon.pdButton cmdUseReference 
         Height          =   1095
         Left            =   11040
         TabIndex        =   20
         Top             =   5160
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1931
         Caption         =   "use"
      End
      Begin PhotoDemon.pdTextBox txtReference 
         Height          =   1095
         Left            =   5040
         TabIndex        =   19
         Top             =   5160
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1931
         Multiline       =   -1  'True
      End
      Begin PhotoDemon.pdListBox lstPhrases 
         Height          =   5295
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   9340
      End
      Begin PhotoDemon.pdDropDown cboPhraseFilter 
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   661
      End
      Begin PhotoDemon.pdButton cmdNextPhrase 
         Height          =   615
         Left            =   5040
         TabIndex        =   3
         Top             =   6720
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   1085
         Caption         =   "Save this translation and proceed to the next phrase"
      End
      Begin PhotoDemon.pdTextBox txtTranslation 
         Height          =   1965
         Left            =   5040
         TabIndex        =   5
         Top             =   2760
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   3466
         Multiline       =   -1  'True
      End
      Begin PhotoDemon.pdTextBox txtOriginal 
         Height          =   1995
         Left            =   5040
         TabIndex        =   7
         Top             =   360
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   3519
         Multiline       =   -1  'True
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
         Top             =   0
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   503
         Caption         =   "phrase groups"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblPhraseBox 
         Height          =   285
         Left            =   0
         Top             =   840
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
         TabIndex        =   18
         Top             =   6720
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   1085
         Caption         =   "Auto-translate all missing phrases"
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   11
         Left            =   4920
         Top             =   4800
         Width           =   6705
         _ExtentX        =   11827
         _ExtentY        =   503
         Caption         =   "reference translation from .po file"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   6
         Left            =   5040
         Top             =   6360
         Width           =   6585
         _ExtentX        =   11615
         _ExtentY        =   503
         Caption         =   "(NOTE: CTRL+ENTER automatically saves and proceeds to next phrase.)"
         ForeColor       =   4210752
      End
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   7455
      Index           =   1
      Left            =   120
      Top             =   720
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   13150
      Begin PhotoDemon.pdCheckBox chkUserLocale 
         Height          =   360
         Left            =   8280
         TabIndex        =   27
         Top             =   405
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   635
         Caption         =   "copy system locale"
         Value           =   0   'False
      End
      Begin PhotoDemon.pdHyperlink hypISO 
         Height          =   375
         Index           =   0
         Left            =   6120
         TabIndex        =   25
         Top             =   840
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   661
         Caption         =   "official ISO language codes at Wikipedia"
         URL             =   "https://en.wikipedia.org/wiki/List_of_ISO_639-1_codes"
      End
      Begin PhotoDemon.pdTextBox txtLangID 
         Height          =   345
         Index           =   1
         Left            =   6840
         TabIndex        =   10
         Top             =   375
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   609
         FontSize        =   11
         Text            =   "US"
      End
      Begin PhotoDemon.pdTextBox txtLangID 
         Height          =   345
         Index           =   0
         Left            =   6120
         TabIndex        =   11
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
         TabIndex        =   12
         Top             =   360
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   609
         FontSize        =   11
         Text            =   "English (US)"
      End
      Begin PhotoDemon.pdTextBox txtLangStatus 
         Height          =   345
         Left            =   240
         TabIndex        =   13
         Top             =   2280
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   609
         FontSize        =   11
         Text            =   "incomplete"
      End
      Begin PhotoDemon.pdTextBox txtLangVersion 
         Height          =   345
         Left            =   240
         TabIndex        =   14
         Top             =   1320
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   609
         FontSize        =   11
         Text            =   "1.0.0"
      End
      Begin PhotoDemon.pdTextBox txtLangAuthor 
         Height          =   345
         Left            =   240
         TabIndex        =   15
         Top             =   3240
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   609
         FontSize        =   11
         Text            =   "enter your name here"
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   8
         Left            =   0
         Top             =   2880
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
         Top             =   1920
         Width           =   11610
         _ExtentX        =   20479
         _ExtentY        =   503
         Caption         =   "language status"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   3
         Left            =   0
         Top             =   960
         Width           =   5730
         _ExtentX        =   10107
         _ExtentY        =   503
         Caption         =   "language version"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   7
         Left            =   0
         Top             =   0
         Width           =   5730
         _ExtentX        =   10107
         _ExtentY        =   503
         Caption         =   "language name"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   5
         Left            =   5880
         Top             =   0
         Width           =   5730
         _ExtentX        =   10107
         _ExtentY        =   503
         Caption         =   "language and country ID"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   0
         Left            =   0
         Top             =   4440
         Width           =   11580
         _ExtentX        =   20426
         _ExtentY        =   503
         Caption         =   "optional translation aids"
         FontSize        =   12
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdTextBox txtApiKey 
         Height          =   345
         Left            =   360
         TabIndex        =   21
         Top             =   5640
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   609
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   1
         Left            =   240
         Top             =   5280
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   503
         Caption         =   "free DeepL.com API key for automatic translation suggestions"
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdLabel lblTitle 
         Height          =   285
         Index           =   10
         Left            =   240
         Top             =   6240
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   503
         Caption         =   "language file (.po) from any other software, as a reference"
         ForeColor       =   4210752
      End
      Begin PhotoDemon.pdTextBox txtPO 
         Height          =   345
         Left            =   360
         TabIndex        =   22
         Top             =   6600
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   609
      End
      Begin PhotoDemon.pdButton cmdPO 
         Height          =   330
         Left            =   11160
         TabIndex        =   23
         Top             =   6600
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
         Caption         =   "..."
      End
      Begin PhotoDemon.pdHyperlink hypISO 
         Height          =   375
         Index           =   1
         Left            =   6120
         TabIndex        =   26
         Top             =   1200
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   661
         Caption         =   "official ISO country codes at Wikipedia"
         URL             =   "https://en.wikipedia.org/wiki/ISO_3166-1_alpha-2#Officially_assigned_code_elements"
      End
      Begin PhotoDemon.pdLabel lblSubheader 
         Height          =   360
         Index           =   0
         Left            =   240
         Top             =   4845
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   635
         Caption         =   "These optional settings can accelerate the translation process.  They are not saved to the language file."
         ForeColor       =   4210752
         Layout          =   1
      End
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   7455
      Index           =   0
      Left            =   120
      Top             =   720
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   13150
      Begin PhotoDemon.pdListBox lstLanguages 
         Height          =   5055
         Left            =   840
         TabIndex        =   4
         Top             =   1080
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   8916
      End
      Begin PhotoDemon.pdButton cmdDeleteLanguage 
         Height          =   690
         Left            =   8400
         TabIndex        =   8
         Top             =   6360
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   1217
         Caption         =   "delete selected language file"
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
         Caption         =   "start new language"
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
         Caption         =   "edit existing language:"
      End
   End
End
Attribute VB_Name = "FormLanguageEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Interactive Language (i18n) Editor
'Copyright 2013-2022 by Tanner Helland
'Created: 28/August/13
'Last updated: 08/July/22
'Last update: rework UI, underlying code as part of broader modernization
'
'This tool can simplify the PhotoDemon localization process.  The original version (built in 2013) was
' heavily influenced by feedback from Frank Donckers.  Many thanks to Frank for his contributions to
' PhotoDemon i18n.  (Frank also contributed the first three language files to the project!)  You can see
' Frank's original, unaltered contributions in the old commit logs for the original version of this tool:
'
'https://github.com/tannerhelland/PhotoDemon/commits/c5d55af4ba3683eec49efc9c6e3d0e5bfc6d2395/Forms/VBP_FormLanguageEditor.frm
'
'Data retention is a key focus of the current editor.  As a safeguard against crashes, two autosaves
' are always maintained.  Autosaves are generated every time a phrase is edited. This (should) guarantee
' that even in the event of a catastrophic failure, only the last-modified phrase will ever risk being lost.
'
'To accelerate the translation process, DeepL.com can automatically populate "estimated" translations.
' This feature uses the official DeepL translation API and as such, requires a free DeepL API key:
'
' https://www.deepl.com/pro-api?cta=header-pro-api
'
'(Scroll down to the "Free" box and click "sign up for free".)
'
'Obviously, a human should always review localizations for best results, but the online service is very
' helpful for accelerating the process, especially on lengthy and/or highly technical text.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'The current list of available languages (e.g. XML files stored in the App - /App/PhotoDemon/Languages - and
' user - Data/Languages - folders).
Private m_listOfLanguages() As PDLanguageFile

'The language currently being edited.
Private m_curLanguage As PDLanguageFile

'All phrases from the existing language file
Private Type PD_Phrase
    txtOriginal As String
    txtTranslation As String
    txtForListBox As String
End Type
Private m_numOfPhrases As Long, m_Phrases() As PD_Phrase

'The current wizard page
Private m_wizardPage As Long

'An interface to an online service is used to auto-populate missing translations (if the user provides
' an API key)
Private m_AutoTranslate As pdAutoLocalize

'An internal XML engine is used to parse and update the actual language file contents
Private m_XMLEngine As pdXML

'To minimize the chance of data loss, PhotoDemon backs up translation data to two alternating files.
' In the event of a crash, this guarantees that we never lose more than the last-edited phrase.
Private m_curBackupFile As Long
Private Const BACKUP_FILE_PREFIX As String = "PD_LANG_EDIT_BACKUP_"

'Hacky fix for specialized Ctrl+Enter detection
Private m_inKeyEvent As Boolean

'The user can (optionally) point at a target .po file for comparison.  This is very helpful for comparing
' phrases to their equivalent in other open-source software, which reduces the chance of us using translations
' different from what everyone else is using.
Private m_ReferencePO As pdStringHash

'During phrase editing, the user can choose to display specific groups of phrases (e.g. "all phrases",
' "only untranslated phrases").  Available options vary according to user settings; for example, phrases matching
' a reference PO are only available if a reference PO file was supplied.
Private Sub cboPhraseFilter_Click()

    lstPhrases.Clear
    lstPhrases.SetAutomaticRedraws False
    
    Dim i As Long
                
    Select Case cboPhraseFilter.ListIndex
    
        'All phrases
        Case 0
            For i = 0 To m_numOfPhrases - 1
                lstPhrases.AddItem m_Phrases(i).txtForListBox
            Next i
        
        'Translated phrases
        Case 1
            For i = 0 To m_numOfPhrases - 1
                If (LenB(m_Phrases(i).txtTranslation) <> 0) Then
                    lstPhrases.AddItem m_Phrases(i).txtForListBox
                End If
            Next i
        
        'Untranslated phrases
        Case 2
            For i = 0 To m_numOfPhrases - 1
                If (LenB(m_Phrases(i).txtTranslation) = 0) Then
                    lstPhrases.AddItem m_Phrases(i).txtForListBox
                End If
            Next i
            
        '(Optional) phrases that don't match the supplied reference .po
        Case 3
            If (Not m_ReferencePO Is Nothing) Then
                If (m_ReferencePO.GetNumOfItems > 0) Then
                    
                    Dim tmpString As String
                    
                    For i = 0 To m_numOfPhrases - 1
                        If m_ReferencePO.GetItemByKey(LCase$(m_Phrases(i).txtOriginal), tmpString) Then
                            If Strings.StringsNotEqual(tmpString, m_Phrases(i).txtTranslation, True) Then
                                lstPhrases.AddItem m_Phrases(i).txtForListBox
                            End If
                        End If
                    Next i
                    
                End If
            End If
    
    End Select
                
    'Redraw the listbox *now*
    lstPhrases.SetAutomaticRedraws True, True
    UpdatePhraseBoxTitle
    
End Sub

'Locale can be pulled from the OS; useful for users creating a new language, so they don't have to look up
' the lang and region IDs manually
Private Sub chkUserLocale_Click()
    If (Not g_Language Is Nothing) And chkUserLocale.Value Then
        txtLangID(0).Text = g_Language.GetSystemUserLanguage()
        txtLangID(1).Text = g_Language.GetSystemUserCountry()
    End If
End Sub

'Use an online service to auto-translate *all* untranslated messages.  This is never ideal (online translations
' need to be human-reviewed), but for languages that don't have an active maintainer, it's sometimes better than
' nothing.
Private Sub cmdAutoTranslate_Click()
    
    'If the internet goes down while auto-translations are processing, errors may be raised by the underlying
    ' winHTTP object.
    On Error GoTo AutoTranslateFailure
    
    'Some strings on this page are error-specific and not intended for average users.  To avoid unnecessary
    ' localization burdens, I hide them from the auto-translator by using a pdString object.
    Dim cString As pdString
    Set cString = New pdString
    
    'Because this process can take a very long time, warn the user in advance.
    Dim msgReturn As VbMsgBoxResult
    cString.AppendLine "Once started, this process cannot be canceled.  It may take a very long time to complete."
    cString.AppendLineBreak
    cString.Append "Are you sure you want to continue?"
    msgReturn = PDMsgBox(cString.ToString(), vbYesNo Or vbInformation, "Automatic translations")
    If (msgReturn <> vbYes) Then Exit Sub
    
    'Count the number of untranslated phrases (so we can provide ongoing status reports)
    Dim totalUntranslated As Long, totalTranslated As Long
    totalUntranslated = 0
    totalTranslated = 0
    
    Dim i As Long
    For i = 0 To m_numOfPhrases - 1
        If (LenB(m_Phrases(i).txtTranslation) = 0) Then totalUntranslated = totalUntranslated + 1
    Next i
    
    Dim srcPhrase As String, retString As String
    
    'Iterate through all untranslated phrases, requesting online translations as we go
    For i = 0 To m_numOfPhrases - 1
        
        'Skip already translated phrases
        If (LenB(m_Phrases(i).txtTranslation) = 0) Then
        
            'Regardless of whether or not we succeed, increment the counter
            totalTranslated = totalTranslated + 1
            cmdAutoTranslate.Caption = g_Language.TranslateMessage("Processing phrase %1 of %2", totalTranslated, totalUntranslated)
            
            'Retrieve the original text, then request a translation from the online service
            srcPhrase = m_Phrases(i).txtOriginal
            
            retString = vbNullString
            retString = m_AutoTranslate.GetDeepLTranslation(srcPhrase)
            
            'If the translation succeeded, store it
            If (LenB(retString) <> 0) Then
                
                'Do a "quick and dirty" case fix for titlecase text
                retString = GetFixedTitlecase(srcPhrase, retString)
                
                'Store the translation, then insert it into the original XML file
                m_Phrases(i).txtTranslation = retString
                m_XMLEngine.UpdateTagAtLocation "translation", m_Phrases(i).txtTranslation, m_XMLEngine.GetLocationOfParentTag("phrase", "original", m_Phrases(i).txtOriginal)
    
            End If
            
            'Every sixteen translations, perform an autosave
            If (i And 15) = 0 Then PerformAutosave
            
        End If
        
    Next i
    
    cmdAutoTranslate.Caption = g_Language.TranslateMessage("Automatic translation complete!")
    
    'Select the "show untranslated phrases" option, which will refresh the list of untranslated phrases
    cboPhraseFilter.ListIndex = 2
    
    Exit Sub
    
AutoTranslateFailure:
    
    'Auto-save everything translated so far
    PerformAutosave
    
    'Notify the user, then exit
    cString.Reset
    cString.Append "The automatic translation process was interrupted.  Any completed translations have been auto-saved.  Please check your internet connection and try again."
    PDMsgBox cString.ToString(), vbCritical Or vbOKOnly, "Automatic translations"
    
End Sub

Private Sub cmdCancel_Click()
    
    'Before exiting, save any preference-like values (like the user's DeepL API key)
    UpdateStoredUserValues
    Unload Me
    
End Sub

'Allow the user to delete a selected language file, if they so desire.
Private Sub cmdDeleteLanguage_Click()
    
    'Make sure a language is selected
    If (lstLanguages.ListIndex < 0) Then Exit Sub
    
    'Make sure we have write access to the target folder.  (This is relevant for people who extract PD
    ' to system-protected folders.)
    If Files.PathExists(Files.FileGetPath(m_listOfLanguages(GetLanguageIndexFromListIndex()).FileName), True) Then
        
        Dim msgReturn As VbMsgBoxResult
        msgReturn = PDMsgBox("Are you sure you want to delete %1?  (This action cannot be undone.)", vbYesNo Or vbExclamation, "Delete language file", lstLanguages.List(lstLanguages.ListIndex))
        
        If (msgReturn = vbYes) Then
            Files.FileDeleteIfExists m_listOfLanguages(GetLanguageIndexFromListIndex()).FileName
            lstLanguages.RemoveItem lstLanguages.ListIndex
            cmdDeleteLanguage.Enabled = False
        End If
    
    'Write access not available
    Else
        PDMsgBox "You do not have access to this folder.  Please log in as an administrator and try again.", vbOKOnly Or vbExclamation, "Error"
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
    
    'Store this translation both locally, and in the original XML file
    m_Phrases(GetPhraseIndexFromListIndex()).txtTranslation = txtTranslation.Text
    m_XMLEngine.UpdateTagAtLocation "translation", txtTranslation, m_XMLEngine.GetLocationOfParentTag("phrase", "original", m_Phrases(GetPhraseIndexFromListIndex()).txtOriginal)
    
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
            If (Not m_ReferencePO Is Nothing) Then
                    
                Dim tmpString As String
                m_ReferencePO.GetItemByKey LCase$(m_Phrases(GetPhraseIndexFromListIndex()).txtOriginal), tmpString
                
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
    
End Sub

Private Sub cmdPrevious_Click()
    ChangeWizardPage False
End Sub

'Change the active wizard page.
Private Sub ChangeWizardPage(ByVal moveForward As Boolean)
    
    Dim unloadFormNow As Boolean
    unloadFormNow = False
    
    'To minimize localization requirements of this tool, some text is handled via pdString objects to avoid it
    ' being marked for localization.
    Dim cString As pdString
    Set cString = New pdString
    
    'Before changing the page, maek sure all user input on the current page is valid
    Select Case m_wizardPage
    
        'The first page is the language selection page.  When the user leaves this page,
        ' we must load the language they've selected into memory and parse all phrases.
        Case 0
            
            'If the user wants to edit an existing language, make sure they've selected one.
            ' (I hate OK-only message boxes, but am currently too lazy to write a more elegant solution.)
            If (optBaseLanguage(1).Value And (lstLanguages.ListIndex < 0)) Then
                cString.Reset
                cString.Append "You must select a language file to edit."
                PDMsgBox cString.ToString(), vbOKOnly Or vbInformation, "Error"
                Exit Sub
            End If
            
            'When starting a new language file (not editing an existing one), set the load path to match
            ' PD's master English language file.
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
                m_curLanguage = m_listOfLanguages(GetLanguageIndexFromListIndex())
                m_curLanguage.FileName = UserPrefs.GetLanguagePath(True) & Files.FileGetName(m_listOfLanguages(GetLanguageIndexFromListIndex()).FileName)
                
                'Attempt to load the selected language from file
                If LoadAllPhrasesFromFile(m_listOfLanguages(GetLanguageIndexFromListIndex()).FileName) Then
                    
                    'No further action is necessary!
                    
                'For some reason, we failed to load the master language file.  Tell them to download a fresh copy of PD.
                Else
                    Screen.MousePointer = vbDefault
                    PDMsgBox "Unfortunately, this language file could not be loaded.  It's possible the copy on this PC is out-of-date." & vbCrLf & vbCrLf & "To continue, please download a fresh copy of PhotoDemon from photodemon.org.", vbOKOnly Or vbExclamation, "Language file could not be loaded"
                    Unload Me
                End If
            
            End If
            
            'Reset the mouse pointer
            Screen.MousePointer = vbDefault
            
        'The second page is the metadata editing page.
        Case 1
            
            'Before doing anything, save the user's DeepL API key (if any) and reference PO (if any)
            UpdateStoredUserValues
            
            'Also, automatically set the destination language of the online translation service
            ' (and the API key, if one was provided)
            m_AutoTranslate.SetDstLanguage Trim$(txtLangID(0))
            If (LenB(Trim$(Me.txtApiKey.Text)) <> 0) Then m_AutoTranslate.SetAPIKey Trim$(Me.txtApiKey.Text) Else m_AutoTranslate.SetAPIKey vbNullString
            
            'When leaving the metadata page, automatically copy all text box entries into the metadata holder
            With m_curLanguage
                .langID = Trim$(txtLangID(0)) & "-" & Trim$(txtLangID(1))
                .LangName = Trim$(txtLangName)
                .LangStatus = Trim$(txtLangStatus)
                .langVersion = Trim$(txtLangVersion)
                .Author = Trim$(txtLangAuthor)
            End With
            
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
            
            'If the user selected a 3rd-party .po file, parse it now so we can quickly compare translations
            LoadReferencePO
            
            'If the user supplied a reference .po, and 1+ phrases were loaded from it, add a new listbox
            ' option in the translation panel for "phrases that don't match reference".
            Dim showReference As Boolean
            showReference = (Not m_ReferencePO Is Nothing)
            If showReference Then showReference = (m_ReferencePO.GetNumOfItems() > 0)
            
            If showReference And (cboPhraseFilter.ListCount <= 3) Then
                cboPhraseFilter.AddItem "phrases that don't match reference"
            End If
            
            'Hide the reference translation if a .po wasn't provided
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
        m_wizardPage = m_wizardPage + 1
    Else
        m_wizardPage = m_wizardPage - 1
    End If
    
    'We can now apply any entrance-timed panel changes
    Select Case m_wizardPage
    
        'Language selection
        Case 0
        
            'Fill the available languages list box with any language files on this system
            PopulateAvailableLanguages
        
        'Metadata editor
        Case 1
        
            'When entering the metadata page, automatically fill all boxes with the currently stored metadata entries
            With m_curLanguage
            
                'Language ID is the most complex, because we must parse the two halves into individual text boxes
                If (InStr(1, .langID, "-", vbBinaryCompare) > 0) Then
                    txtLangID(0) = Left$(.langID, InStr(1, .langID, "-", vbBinaryCompare) - 1)
                    txtLangID(1) = Mid$(.langID, InStr(1, .langID, "-", vbBinaryCompare) + 1, Len(.langID) - InStr(1, .langID, "-", vbBinaryCompare))
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
        
            'By default, select the "show untranslated phrases" setting
            cboPhraseFilter.ListIndex = 2
                
    End Select
    
    'Hide all inactive panels (and show the active one)
    Dim i As Long
    For i = 0 To picContainer.Count - 1
        picContainer(i).Visible = (i = m_wizardPage)
    Next i
    
    'If we are at the beginning, disable the previous button
    cmdPrevious.Enabled = (m_wizardPage <> 0)
    
    'If we are at the end, change the text of the "next" button; otherwise, make sure it says "next"
    If (m_wizardPage = picContainer.Count - 1) Then
        cmdNext.Caption = g_Language.TranslateMessage("Save and exit")
    Else
        cmdNext.Caption = g_Language.TranslateMessage("Next")
    End If
    
    'Finally, change the top title caption to match the current step
    Dim wzTitle As pdString
    Set wzTitle = New pdString
    wzTitle.Append g_Language.TranslateMessage("Step %1:", m_wizardPage + 1)
    wzTitle.Append " "
    
    Select Case m_wizardPage
    
        Case 0
            wzTitle.Append g_Language.TranslateMessage("select language")
            
        Case 1
            wzTitle.Append g_Language.TranslateMessage("apply language and translation settings")
            
        Case 2
            wzTitle.Append g_Language.TranslateMessage("localize phrases")
            
    End Select
    
    lblWizardTitle.Caption = wzTitle.ToString()
        
End Sub

Private Sub cmdUseReference_Click()
    If Strings.StringsNotEqual(txtTranslation.Text, txtReference.Text, False) Then
        lblTranslatedPhrase.Caption = g_Language.TranslateMessage("translated phrase") & " " & g_Language.TranslateMessage("(NOT YET SAVED)")
    End If
    txtTranslation.Text = txtReference.Text
End Sub

Private Sub Form_Load()
    
    m_curBackupFile = 0
    
    'By default, the first wizard page is displayed.  (We start at -1 because we will incerement the page count by +1 with our first
    ' call to changeWizardPage in Form_Activate)
    m_wizardPage = -1
    
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

'Given a source language file, find all phrase tags and load them into a specialized phrase array.
Private Function LoadAllPhrasesFromFile(ByVal srcLangFile As String) As Boolean
    
    LoadAllPhrasesFromFile = False
    
    Set m_XMLEngine = New pdXML
    
    'Attempt to load the language file
    If m_XMLEngine.LoadXMLFile(srcLangFile) Then
    
        'Validate the language file's contents
        If m_XMLEngine.IsPDDataType("Translation") And m_XMLEngine.ValidateLoadedXMLData("phrase") Then
        
            m_XMLEngine.SetTextCompareMode vbBinaryCompare
            
            'Attempt to load all phrase tag location occurrences
            Dim phraseLocations() As Long
            If m_XMLEngine.FindAllTagLocations(phraseLocations, "phrase") Then
                
                m_numOfPhrases = UBound(phraseLocations) + 1
                ReDim m_Phrases(0 To m_numOfPhrases - 1) As PD_Phrase
                
                Dim tmpString As String
                
                Dim i As Long
                For i = 0 To m_numOfPhrases - 1
                    tmpString = m_XMLEngine.GetUniqueTag_String("original", vbNullString, phraseLocations(i))
                    m_Phrases(i).txtOriginal = tmpString
                    m_Phrases(i).txtTranslation = m_XMLEngine.GetUniqueTag_String("translation", vbNullString, phraseLocations(i) + Len(tmpString))
                    
                    'We also need a modified version of the string to add to the phrase list box.
                    ' (This text can't include line breaks.)
                    If (InStr(1, tmpString, vbCr, vbBinaryCompare) > 0) Then tmpString = Replace(tmpString, vbCr, vbNullString, 1, -1, vbBinaryCompare)
                    If (InStr(1, tmpString, vbLf, vbBinaryCompare) > 0) Then tmpString = Replace(tmpString, vbLf, vbNullString, 1, -1, vbBinaryCompare)
                    m_Phrases(i).txtForListBox = tmpString
                    
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
    
    Dim origText As String
    origText = m_Phrases(GetPhraseIndexFromListIndex()).txtOriginal
    
    txtOriginal.Text = origText
    
    'If a translation exists for this phrase, load it.  If it does not, and we have an online service available,
    ' query that online service for a translation.
    lblTranslatedPhrase.Caption = g_Language.TranslateMessage("translated phrase")
    
    If (LenB(m_Phrases(GetPhraseIndexFromListIndex()).txtTranslation) <> 0) Then
        txtTranslation.Text = m_Phrases(GetPhraseIndexFromListIndex()).txtTranslation
        lblTranslatedPhrase.Caption = lblTranslatedPhrase.Caption & " " & g_Language.TranslateMessage("(saved)")
    Else
    
        lblTranslatedPhrase.Caption = lblTranslatedPhrase.Caption & " " & g_Language.TranslateMessage("(NOT YET SAVED)")
        
        If (LenB(m_AutoTranslate.GetAPIKey) <> 0) Then
        
            txtTranslation.Text = g_Language.TranslateMessage("waiting...")
            
            'Query the online service for a translation
            Dim retString As String
            retString = m_AutoTranslate.GetDeepLTranslation(origText)
            
            'Apply title case (as relevant) to the returned string
            If (LenB(retString) <> 0) Then
                txtTranslation.Text = GetFixedTitlecase(origText, retString)
            Else
                txtTranslation.Text = g_Language.TranslateMessage("[translation failed]")
            End If
            
        Else
            txtTranslation.Text = vbNullString
        End If
            
    End If
    
    'If a .po reference was provided, look for this text there too.
    ' (Do this regardless of whether this phrase has already been translated or not)
    If (Not m_ReferencePO Is Nothing) And txtReference.Visible Then
        
        Dim strReference As String
        If m_ReferencePO.GetItemByKey(LCase$(origText), strReference) Then
            txtReference.Text = strReference
        Else
            txtReference.Text = g_Language.TranslateMessage("[phrase not in file]")
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
    
    'Generate an autosave filename.  The language ID is appended to the name, so separate autosaves exist for each
    ' edited language (assuming they have different language IDs).
    Dim backupFile As String
    backupFile = UserPrefs.GetLanguagePath(True) & BACKUP_FILE_PREFIX & m_curLanguage.langID & "_" & Trim$(Str$(m_curBackupFile)) & ".tmpxml"
    
    'The XML engine handles the actual writing to file.  For performance reasons, auto-tabbing is suppressed.
    m_XMLEngine.WriteXMLToFile backupFile, True

End Sub

'Fill the first panel ("select a language file") with all available language files on this system
Private Sub PopulateAvailableLanguages()
    
    'Retrieve a list of available languages from the translation engine
    g_Language.CopyListOfLanguages m_listOfLanguages
    
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
            
                'Use the XML engine to validate this file, and to make sure it contains at least a language ID, name,
                ' and 1+ phrases
                If tmpXML.IsPDDataType("Translation") And tmpXML.ValidateLoadedXMLData("langid", "langname", "phrase") Then
                
                    ReDim Preserve m_listOfLanguages(0 To UBound(m_listOfLanguages) + 1) As PDLanguageFile
                    
                    With m_listOfLanguages(UBound(m_listOfLanguages))
                    
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
    
    'Add the contents of that array to the list box on the opening panel (the list of available languages,
    ' from which the user can select a file as the "starting point" for their own translation).
    lstLanguages.Clear
    lstLanguages.SetAutomaticRedraws False, False
    
    Dim i As Long
    For i = 0 To UBound(m_listOfLanguages)
    
        'Note that we DO NOT add the English language entry - that is used by the "start a new language file from scratch" option.
        If Strings.StringsNotEqual(m_listOfLanguages(i).LangType, "DEFAULT", True) Then
            
            Dim listEntry As String
            listEntry = m_listOfLanguages(i).LangName & " "
            
            'Use the author name embedded in the file, if any
            Dim authName As String
            If (LenB(m_listOfLanguages(i).Author) <> 0) Then
                authName = m_listOfLanguages(i).Author
            Else
                authName = g_Language.TranslateMessage("unknown author")
            End If
            
            'For official translations, an author name will always be provided.  Include the author's name in the list.
            If (m_listOfLanguages(i).LangType = "Official") Then
                listEntry = listEntry & g_Language.TranslateMessage("(official translation by %1)", authName)
                
            'For unofficial translations, an author name may not be provided.  Include the author's name only if it's available.
            ElseIf (m_listOfLanguages(i).LangType = "Unofficial") Then
                listEntry = listEntry & g_Language.TranslateMessage("by %1", authName)
                
            'Anything else is an autosave; on these we'll also append the autosave date
            Else
            
                'Include author name if available
                listEntry = listEntry & g_Language.TranslateMessage("by %1", authName) & " "
                
                'Add autosave time and date
                listEntry = listEntry & g_Language.TranslateMessage("(autosaved on %1)", Format$(FileDateTime(m_listOfLanguages(i).FileName), "hh:mm:ss AM/PM, dd-mmm-yy"))
                
            End If
            
            'Add the finished text to the listbox
            lstLanguages.AddItem listEntry
            m_listOfLanguages(i).InternalDisplayName = listEntry
            
        Else
            'Ignore the default language entry entirely
        End If
    Next i
    
    'By default, no language is selected for the user
    lstLanguages.SetAutomaticRedraws True, True
    lstLanguages.ListIndex = -1
    
End Sub

'Mapping functions between internal arrays and on-screen listboxes
Private Function GetLanguageIndexFromListIndex() As Long
    Dim i As Long
    For i = LBound(m_listOfLanguages) To UBound(m_listOfLanguages)
        If Strings.StringsEqual(lstLanguages.List(lstLanguages.ListIndex), m_listOfLanguages(i).InternalDisplayName) Then
            GetLanguageIndexFromListIndex = i
            Exit For
        End If
    Next i
End Function

Private Function GetPhraseIndexFromListIndex() As Long
    Dim i As Long
    For i = LBound(m_Phrases) To UBound(m_Phrases)
        If Strings.StringsEqual(lstPhrases.List(lstPhrases.ListIndex), m_Phrases(i).txtForListBox) Then
            GetPhraseIndexFromListIndex = i
            Exit For
        End If
    Next i
End Function

'On Win 7+, we attempt to automatically handle titlecase of translated text.
'
'If the original English string used titlecase, we'll set titlecase to the translated string as well, *if* the
' translated string came from an online service.  (This ensures grammar uniformity across languages, even if the
' online service doesn't attempt to match casing.)
Private Function GetFixedTitlecase(ByVal origString As String, ByVal translatedString As String) As String
    
    On Error GoTo TitlecaseFail
    
    If (LenB(origString) <> 0) And (LenB(translatedString) <> 0) Then
    
        If OS.IsWin7OrLater Then
            
            Dim origStringTitlecase As Boolean
            
            'Split the original string into individual words
            Dim strOrig() As String, strTranslated() As String
            strOrig = Split(origString, " ", , vbBinaryCompare)
            strTranslated = Split(translatedString, " ", , vbBinaryCompare)
            
            'Only proceed with automatic casing if *both* strings contain multiple words.
            ' (Some translations may not result in 1:1 word counts.)
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

'Handle Ctrl+Enter specially (save current phrase and proceed to next one)
Private Sub txtTranslation_KeyDown(ByVal Shift As ShiftConstants, ByVal vKey As Long, preventFurtherHandling As Boolean)

    If (vKey = vbKeyReturn) And (Shift And vbCtrlMask = vbCtrlMask) Then
        preventFurtherHandling = True
        m_inKeyEvent = True
        PhraseFinished
        txtTranslation.SelStart = Len(txtTranslation.Text)
    Else
        m_inKeyEvent = False
    End If

End Sub

Private Sub txtTranslation_KeyPress(ByVal Shift As ShiftConstants, ByVal vKey As Long, preventFurtherHandling As Boolean)
    preventFurtherHandling = m_inKeyEvent
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
    If (InStr(1, srcText, MSG_ID, vbBinaryCompare) = 0) Then Exit Sub
    If (InStr(1, srcText, MSG_STR, vbBinaryCompare) = 0) Then Exit Sub
    
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
    Set m_ReferencePO = New pdStringHash
    
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
                curPhrase = Replace$(curPhrase, "\""", "&quot;", 1, -1, vbBinaryCompare)
                
                'Remove linebreaks
                curPhrase = Replace$(curPhrase, vbCrLf, vbNullString, 1, -1, vbBinaryCompare)
                
                'Remove any remaining quotes
                curPhrase = Replace$(curPhrase, """", vbNullString, 1, -1, vbBinaryCompare)
                
                'Restore valid quotes (that we hacked out at the beginning)
                curPhrase = Replace$(curPhrase, "&quot;", """", 1, -1, vbBinaryCompare)
                
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
                    curPhrase = Replace$(curPhrase, "\""", "&quot;", 1, -1, vbBinaryCompare)
                    
                    'Remove linebreaks
                    curPhrase = Replace$(curPhrase, vbCrLf, vbNullString, 1, -1, vbBinaryCompare)
                    
                    'Remove any remaining quotes
                    curPhrase = Replace$(curPhrase, """", vbNullString, 1, -1, vbBinaryCompare)
                    
                    'Restore valid quotes (that we hacked out at the beginning)
                    curPhrase = Replace$(curPhrase, "&quot;", """", 1, -1, vbBinaryCompare)
                    
                End If
                
                'Debug.Print "-----" & vbCrLf & curPhrase & vbCrLf & "-----"
                
            End If
            
            'We now have the translation.
            msgStr = curPhrase
            
            'Only store key+value pairs where both entities exist
            If (LenB(msgID) > 0) Then
                
                'Do a little pre-processing to both strings.  In particular, we don't want trailing ellipses
                ' or markers for hotkeys (typically _); we're only interested in the actual text
                If (InStr(1, msgID, ELLIPSIS, vbBinaryCompare) <> 0) Then msgID = Replace$(msgID, ELLIPSIS, vbNullString, 1, -1, vbBinaryCompare)
                If (InStr(1, msgID, ELLIPSIS_CHAR, vbBinaryCompare) <> 0) Then msgID = Replace$(msgID, ELLIPSIS_CHAR, vbNullString, 1, -1, vbBinaryCompare)
                If (InStr(1, msgID, UNDERSCORE_CHAR, vbBinaryCompare) <> 0) Then msgID = Replace$(msgID, UNDERSCORE_CHAR, vbNullString, 1, -1, vbBinaryCompare)
                
                If (InStr(1, msgStr, ELLIPSIS, vbBinaryCompare) <> 0) Then msgStr = Replace$(msgStr, ELLIPSIS, vbNullString, 1, -1, vbBinaryCompare)
                If (InStr(1, msgStr, ELLIPSIS_CHAR, vbBinaryCompare) <> 0) Then msgStr = Replace$(msgStr, ELLIPSIS_CHAR, vbNullString, 1, -1, vbBinaryCompare)
                If (InStr(1, msgStr, UNDERSCORE_CHAR, vbBinaryCompare) <> 0) Then msgStr = Replace$(msgStr, UNDERSCORE_CHAR, vbNullString, 1, -1, vbBinaryCompare)
                
                'We want case-insensitive matching, so deliberately lcase all keys
                m_ReferencePO.AddItem LCase$(msgID), msgStr
                
            End If
            
        End If
        
        'Look for the next phrase
        If (idxEnd > 0) Then idxStart = InStr(idxEnd + 1, srcText, MSG_ID, vbBinaryCompare) Else idxStart = -1
        
    Loop
    
    If (Not m_ReferencePO Is Nothing) Then PDDebug.LogAction "Loaded " & m_ReferencePO.GetNumOfItems() & " phrases from the reference .po in " & VBHacks.GetTimeDiffNowAsString(startTime)
    
End Sub
