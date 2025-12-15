VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000005&
   Caption         =   " PhotoDemon i18n manager"
   ClientHeight    =   9915
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14535
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   661
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   969
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdForeignMerge 
      Caption         =   "3) Replace missing translations in (1) with any matching translations in (2)"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   8280
      TabIndex        =   17
      Top             =   8040
      Width           =   5775
   End
   Begin VB.CommandButton cmdForeignMerge 
      Caption         =   "2) Select XML file to use for missing translations..."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   4560
      TabIndex        =   16
      Top             =   8040
      Width           =   3255
   End
   Begin VB.CommandButton cmdForeignMerge 
      Caption         =   "1) Select base XML file..."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   840
      TabIndex        =   15
      Top             =   8040
      Width           =   3255
   End
   Begin VB.CommandButton cmdMergeAll 
      Caption         =   "2a (Optional) Automatically merge all PhotoDemon localizations against the latest en-US data..."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      TabIndex        =   13
      Top             =   6600
      Width           =   9495
   End
   Begin VB.CheckBox chkRemoveDuplicates 
      BackColor       =   &H80000005&
      Caption         =   " Remove duplicate entries"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   3960
      Value           =   1  'Checked
      Width           =   3015
   End
   Begin VB.CommandButton cmdMerge 
      Caption         =   "3) Merge the files into an updated localized XML file        (NOTE: this will not modify the source files)"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8280
      TabIndex        =   7
      Top             =   5640
      Width           =   5775
   End
   Begin VB.CommandButton cmdOldLanguage 
      Caption         =   "2) Select localized XML file..."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      TabIndex        =   6
      Top             =   5640
      Width           =   3255
   End
   Begin VB.CommandButton cmdEnUsFile 
      Caption         =   "1) Select en-US XML file..."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   5
      Top             =   5640
      Width           =   3255
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Begin processing"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   3240
      Width           =   3015
   End
   Begin VB.ListBox lstProjectFiles 
      Height          =   2310
      Left            =   4200
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   1920
      Width           =   10095
   End
   Begin VB.CommandButton cmdSelectVBP 
      Caption         =   "Select VBP file..."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   1920
      Width           =   3015
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "merge two language files together:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Index           =   2
      Left            =   480
      TabIndex        =   14
      Top             =   7560
      Width           =   3615
   End
   Begin VB.Label lblWarning 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmGenerateI18N.frx":0000
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   975
      Left            =   240
      TabIndex        =   12
      Top             =   120
      Width           =   14055
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblUpdates 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   4080
      TabIndex        =   11
      Top             =   4680
      Width           =   10215
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "merge old translation files with new data:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Index           =   1
      Left            =   480
      TabIndex        =   10
      Top             =   5160
      Width           =   4275
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "extra language support tools"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   4680
      Width           =   3300
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      X1              =   8
      X2              =   960
      Y1              =   296
      Y2              =   296
   End
   Begin VB.Label lblExtract 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "step 2: process all files in project"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   240
      TabIndex        =   8
      Top             =   2760
      Width           =   3345
   End
   Begin VB.Label lblVBP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "step 1: select VBP file"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   2175
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Master English Language File (XML) Generator
'Copyright 2013-2022 by Tanner Helland
'Created: 23/January/13
'Last updated: 01/August/22
'Last update: dump a micro "phrase database" alongside the primary en-US file
'
'This project is designed to scan through all project files in PhotoDemon, extract any user-facing English text, and compile
' it into an XML file which can be used as the basis for translations into other languages.  It reads the master PhotoDemon.vbp
' file, compiles a list of all project files, then analyzes them individually.  Control text is extracted (unless the text is
' in an FRX file - in that case the text needs to be manually rewritten so this project can find it).  Message box and
' progress/status bar text is also extracted, but this project relies on some particular PhotoDemon implementation quirks to
' do so.
'
'Basic statistics and organization information are added as comments to the final XML file.
'
'This project also supports merging an updated English language XML file with an outdated non-English language file, the result
' of which can be used to fill-in missing translations while keeping any that are still valid.  I typically use this a week or
' two before a formal release, so I can hand off new XML files to translators for them to update with any new or modified text.
'
'NOTE: this project is intended only as a support tool for PhotoDemon.  It is not designed or tested for general-purpose use.
'      As such, the code is pretty ugly.  Organization is minimal.  Read at your own risk.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit
Option Compare Binary

'PD currently uses an XML-like structure for its language files.  This means the same tags are
' written over and over and over.
Private Const XML_PHRASE_OPEN As String = "<phrase>"
Private Const XML_PHRASE_CLOSE As String = "</phrase>"
Private Const XML_ORIGINAL_OPEN As String = "<original>"
Private Const XML_ORIGINAL_CLOSE As String = "</original>"
Private Const XML_TRANSLATION_PAIR As String = "<translation></translation>"

'Variables used to generate the master translation file
Private m_VBPFile As String, m_VBPPath As String
Private m_FormName As String, m_ObjectName As String, m_FileName As String
Private m_NumOfPhrasesFound As Long, m_NumOfPhrasesWritten As Long, m_numOfWords As Long
Private vbpText() As String, m_vbpFiles() As String
Private m_outputText As pdString, outputFile As String

'Variables used to merge old language files with new ones
Private m_AllEnUsText As String, m_OldLanguageText As String, m_NewLanguageText As String
Private m_OldLanguagePath As String

'Supply missing phrases in a translation with phrases from another, similar translation.
' (Useful when languages are similar but not identical, e.g. Flemish > Dutch, Spanish-MX > Spanish-ES)
Private m_Alli18nFile As String, m_Alli18nText As String

'Variables used to build a blacklist of text that does not need to be translated
Private m_Blacklist As pdStringHash

'String to store the version of the current VBP file (which will be written out to the master XML file for guidance)
Private m_VersionString As String

'If silent mode has been activated via command line, this will be set to TRUE.
Private m_SilentMode As Boolean

'A pdXML instance provides UTF-8 support.
Private m_XML As pdXML

'If duplicates are assigned for removal, this flag is set to TRUE
Private m_RemoveDuplicates As Boolean

'When creating the default en-US language file, we track phrases we've already detected.  Duplicates
' are flagged and removed automatically.
Private m_enUSPhrases As pdStringHash

'When creating the full en-US phrase collection, we also collect detailed stats on each phrase.
' These stats allow the user to sort phrases by category, which is a large help when starting a new
' language file (so the localizer can concentrate on important phrases first).
'
'Importantly, this list needs to be sorted by importance (with the *most* important value at the
' top of the list, i.e. the lowest enum value).  This affects phrase sorting in PD's language editor,
' and localizers need it to "estimate" which phrases are most important (meaning which phrases should
' be translated first, vs which ones might be left to machine translation).
Private Enum PD_PhraseType
    pt_UIElement = 1
    pt_ActionName = 2
    pt_MsgBox = 4
    pt_StatusBar = 8
    pt_Tooltip = 16
    pt_Miscellaneous = 32
End Enum

#If False Then
    Private Const pt_UIElement = 1, pt_ActionName = 2, pt_MsgBox = 4, pt_StatusBar = 8, pt_Tooltip = 16, pt_Miscellaneous = 32
#End If

Private Type PD_PhraseInfo
    phraseType As PD_PhraseType
    origEnUSPhrase As String
    occursInFiles As String
End Type

Private m_phraseData() As PD_PhraseInfo, m_numPhraseData As Long

'During silent mode (used to synchronize localizations), we use a fast string hash table to update
' language files.  This greatly improves performance, especially given how many language files PD ships.
Private m_PhraseCollection As pdStringHash

Private Sub cmdEnUsFile_Click()

    Dim cDialog As pdOpenSaveDialog
    Set cDialog = New pdOpenSaveDialog
    
    'This project should be located in a sub-path of a normal PhotoDemon install.
    ' We can use shlwapi's PathCanonicalize function to automatically "guess" at the location of PD's
    ' base en-US language file.
    Dim likelyDefaultLocation As String
    If Files.PathCanonicalize(Files.AppPathW() & "..\..", likelyDefaultLocation) Then likelyDefaultLocation = Files.PathAddBackslash(likelyDefaultLocation)
    likelyDefaultLocation = likelyDefaultLocation & "App\PhotoDemon\Languages\Master\MASTER.xml"
    
    If cDialog.GetOpenFileName(likelyDefaultLocation, , True, False, "XML - PhotoDemon Language File|*.xml", 1, , "Please select a PhotoDemon language file (XML)", "xml", Me.hWnd) Then
        Files.FileLoadAsString likelyDefaultLocation, m_AllEnUsText, True
        
        'Remove tab chars, if any exist
        m_AllEnUsText = Replace$(m_AllEnUsText, vbTab, vbNullString, 1, -1, vbBinaryCompare)
    End If
    
End Sub

Private Sub cmdForeignMerge_Click(Index As Integer)

    Dim cDialog As pdOpenSaveDialog
    Set cDialog = New pdOpenSaveDialog
    
    'This project should be located in a sub-path of a normal PhotoDemon install.
    ' We can use shlwapi's PathCanonicalize function to automatically "guess" at the location of PD's
    ' base en-US language file.
    Dim likelyDefaultLocation As String
    If Files.PathCanonicalize(Files.AppPathW() & "..\..", likelyDefaultLocation) Then likelyDefaultLocation = Files.PathAddBackslash(likelyDefaultLocation)
    likelyDefaultLocation = likelyDefaultLocation & "App\PhotoDemon\Languages\"
    
    Dim srcFile As String
    
    Select Case Index
        
        'Choose base file
        Case 0
            
            If cDialog.GetOpenFileName(srcFile, , True, False, "XML - PhotoDemon Language File|*.xml", 1, likelyDefaultLocation, "Please select a PhotoDemon language file (XML)", "xml", Me.hWnd) Then
                
                'Load the file and remove tab chars, if any (sometimes added by 3rd-party editors)
                m_Alli18nFile = srcFile
                Files.FileLoadAsString srcFile, m_Alli18nText, True
                m_Alli18nText = Replace$(m_Alli18nText, vbTab, vbNullString, 1, -1, vbBinaryCompare)
                
            End If
            
        
        'Choose file that will supply missing translations
        Case 1
            
            If cDialog.GetOpenFileName(srcFile, , True, False, "XML - PhotoDemon Language File|*.xml", 1, likelyDefaultLocation, "Please select a PhotoDemon language file (XML)", "xml", Me.hWnd) Then
                
                'Load the file and remove tab chars, if any (sometimes added by 3rd-party editors)
                m_OldLanguagePath = srcFile
                Files.FileLoadAsString srcFile, m_OldLanguageText, True
                m_OldLanguageText = Replace$(m_OldLanguageText, vbTab, vbNullString, 1, -1, vbBinaryCompare)
                
            End If
        
        'Replace any missing translations in [base file] with matching translations in [other file]
        Case 2
                
            'Make sure I selected two files
            If (LenB(m_Alli18nText) = 0) Or (LenB(m_OldLanguageText) = 0) Then
                MsgBox "One or more source files are missing.  Supply those before attempting a merge."
                Exit Sub
            End If
            
            'Start by copying the contents of the master file into the destination string.
            ' We will use that as our base, and update it with the old translations as we go.
            m_NewLanguageText = m_Alli18nText
            
            Dim origText As String, translatedText As String
            Dim findText As String, replaceText As String
            
            Dim phrasesProcessed As Long, phrasesFound As Long, phrasesMissed As Long
            phrasesProcessed = 0
            phrasesFound = 0
            phrasesMissed = 0
            
            'Find the first occurence of a <phrase> tag
            Dim sPos As Long, sPosTranslation As Long
            sPos = InStr(1, m_NewLanguageText, XML_PHRASE_OPEN, vbBinaryCompare)
            sPosTranslation = 1
            
            'Start parsing the base text for <phrase> tags
            Do
            
                phrasesProcessed = phrasesProcessed + 1
                
                'Retrieve the original text associated with this phrase tag
                Const ORIG_TEXT_TAG As String = "original"
                origText = GetTextBetweenTags(m_Alli18nText, ORIG_TEXT_TAG, sPos)
                
                'Attempt to retrieve a translation for this phrase using the old language file.
                translatedText = GetTranslationTagFromCaption_CustomFile(origText, m_Alli18nText, sPosTranslation)
                
                'If no translation was found, and this string contains vbCrLf characters, replace them with plain vbLF characters and try again
                If (LenB(translatedText) = 0) Then
                    If (InStr(1, origText, vbCrLf) > 0) Then
                        translatedText = GetTranslationTagFromCaption_CustomFile(Replace$(origText, vbCrLf, vbLf), m_Alli18nText)
                    End If
                End If
                
                'If we still didn't find a translation, try to pull it from the backup file
                If (LenB(translatedText) = 0) Then
                    
                    'Attempt to retrieve a translation for this phrase using the old language file.
                    translatedText = GetTranslationTagFromCaption_CustomFile(origText, m_OldLanguageText, sPosTranslation)
                    
                    If (LenB(translatedText) <> 0) Then
                        
                        findText = XML_ORIGINAL_OPEN & origText & XML_ORIGINAL_CLOSE & vbCrLf & XML_TRANSLATION_PAIR
                        replaceText = XML_ORIGINAL_OPEN & origText & XML_ORIGINAL_CLOSE & vbCrLf & "<translation>" & translatedText & "</translation>"
                        m_NewLanguageText = Replace$(m_NewLanguageText, findText, replaceText)
                        
                        phrasesFound = phrasesFound + 1
                        
                    Else
                        Debug.Print "Couldn't find translation for: " & origText
                        phrasesMissed = phrasesMissed + 1
                    End If
                    
                Else
                    'Translation already exists; do nothing
                End If
            
                'Find the next occurrence of a <phrase> tag
                sPos = InStr(sPos + 1, m_Alli18nText, XML_PHRASE_OPEN, vbBinaryCompare)
                
                If ((phrasesProcessed And 15) = 0) Then
                    Message phrasesProcessed & " phrases processed.  (" & phrasesFound & " found, " & phrasesMissed & " missed)"
                End If
                
            Loop While sPos > 0
            
            '(This next code block is copied verbatim from cmdMergeAll.  It has only been tested *there*.)
            
            'Finally, look for any language-specific text+translation pairs.  This (optional) segment can be used
            ' to map phrases with identical English text (e.g. Color > Invert vs Selection > Invert) to unique
            ' phrases in a given translation.
            Dim posSpecialBlock As Long
            posSpecialBlock = InStrRev(m_Alli18nText, "<special-translations>", -1, vbBinaryCompare)
            If (posSpecialBlock > 0) Then
            
                'Copy over the entire <special-translations> XML block as-is.
                Const END_SPECIAL_BLOCK As String = "</special-translations>"
                Dim posSpecialBlockEnd As Long
                posSpecialBlockEnd = InStr(posSpecialBlock, m_Alli18nText, END_SPECIAL_BLOCK, vbBinaryCompare)
                If (posSpecialBlockEnd > posSpecialBlock) Then
                    
                    Dim insertPosition As Long
                    insertPosition = InStrRev(m_NewLanguageText, "</pdData>", -1, vbBinaryCompare)
                    If (insertPosition > 0) Then m_NewLanguageText = Replace$(m_NewLanguageText, "</pdData>", Mid$(m_Alli18nText, posSpecialBlock, (posSpecialBlockEnd - posSpecialBlock) + Len(END_SPECIAL_BLOCK)) & vbCrLf & vbCrLf & "</pdData>")
                    
                End If
                
            End If
            
            'Prompt the user to save the results
            Dim fPath As String
            fPath = m_Alli18nFile
            
            If cDialog.GetSaveFileName(fPath, , True, "XML - PhotoDemon Language File|*.xml", 1, , "Save the merged language file (XML)", "xml", Me.hWnd) Then
                
                'Worried about breaking something?  Enable strict overwrite checking:
                'If Files.FileExists(fPath) Then
                '    MsgBox "File already exists!  Too dangerous to overwrite - please perform the merge again."
                '    Exit Sub
                'End If
                
                'Use pdXML to write out a UTF-8 encoded XML file
                m_XML.LoadXMLFromString m_NewLanguageText
                m_XML.WriteXMLToFile fPath, True
                
            End If
            
            MsgBox "Merge complete." & vbCrLf & vbCrLf & phrasesProcessed & " phrases processed. " & phrasesFound & " translations found. " & phrasesMissed & " translations missing."
    
    End Select
    
End Sub

Private Sub cmdMerge_Click()

    'Make sure our source file strings are not empty
    If (LenB(m_AllEnUsText) = 0) Or (LenB(m_OldLanguageText) = 0) Then
        MsgBox "One or more source files are missing.  Supply those before attempting a merge."
        Exit Sub
    End If
    
    'Start by copying the contents of the master file into the destination string.
    ' We will use that as our base, and update it with the old translations as we go.
    m_NewLanguageText = m_AllEnUsText
    
    Dim origText As String, translatedText As String
    Dim findText As String, replaceText As String
    
    'Copy over all top-level language and author information
    ReplaceTopLevelTag "langid", m_AllEnUsText, m_OldLanguageText, m_NewLanguageText
    ReplaceTopLevelTag "langname", m_AllEnUsText, m_OldLanguageText, m_NewLanguageText
    ReplaceTopLevelTag "langversion", m_AllEnUsText, m_OldLanguageText, m_NewLanguageText
    ReplaceTopLevelTag "langstatus", m_AllEnUsText, m_OldLanguageText, m_NewLanguageText
    ReplaceTopLevelTag "author", m_AllEnUsText, m_OldLanguageText, m_NewLanguageText
        
    Dim phrasesProcessed As Long, phrasesFound As Long, phrasesMissed As Long
    phrasesProcessed = 0
    phrasesFound = 0
    phrasesMissed = 0
    
    'Find the first occurence of a <phrase> tag
    Dim sPos As Long, sPosTranslation As Long
    sPos = InStr(1, m_NewLanguageText, XML_PHRASE_OPEN, vbBinaryCompare)
    sPosTranslation = 1
    
    'Start parsing the master text for <phrase> tags
    Do
    
        phrasesProcessed = phrasesProcessed + 1
        
        'Retrieve the original text associated with this phrase tag
        Const ORIG_TEXT_TAG As String = "original"
        origText = GetTextBetweenTags(m_AllEnUsText, ORIG_TEXT_TAG, sPos)
        
        'Attempt to retrieve a translation for this phrase using the old language file.
        translatedText = GetTranslationTagFromCaption(origText, sPosTranslation)
        
        'If no translation was found, and this string contains vbCrLf characters, replace them with plain vbLF characters and try again
        If (LenB(translatedText) = 0) Then
            If (InStr(1, origText, vbCrLf) > 0) Then
                translatedText = GetTranslationTagFromCaption(Replace$(origText, vbCrLf, vbLf))
            End If
        End If
        
        'If a translation was found, insert it into the new file
        If (LenB(translatedText) <> 0) Then
            
            'As a failsafe, try the same thing without tabs
            findText = XML_ORIGINAL_OPEN & origText & XML_ORIGINAL_CLOSE & vbCrLf & XML_TRANSLATION_PAIR
            replaceText = XML_ORIGINAL_OPEN & origText & XML_ORIGINAL_CLOSE & vbCrLf & "<translation>" & translatedText & "</translation>"
            m_NewLanguageText = Replace$(m_NewLanguageText, findText, replaceText)
            
            phrasesFound = phrasesFound + 1
        Else
            Debug.Print "Couldn't find translation for: " & origText
            phrasesMissed = phrasesMissed + 1
        End If
    
        'Find the next occurrence of a <phrase> tag
        sPos = InStr(sPos + 1, m_AllEnUsText, XML_PHRASE_OPEN, vbBinaryCompare)
        
        If ((phrasesProcessed And 15) = 0) Then
            Message phrasesProcessed & " phrases processed.  (" & phrasesFound & " found, " & phrasesMissed & " missed)"
        End If
        
    Loop While sPos > 0
    
    '(This next code block is copied verbatim from cmdMergeAll.  It has only been tested *there*.)
    
    'Finally, look for any language-specific text+translation pairs.  This (optional) segment can be used
    ' to map phrases with identical English text (e.g. Color > Invert vs Selection > Invert) to unique
    ' phrases in a given translation.
    Dim posSpecialBlock As Long
    posSpecialBlock = InStrRev(m_OldLanguageText, "<special-translations>", -1, vbBinaryCompare)
    If (posSpecialBlock > 0) Then
    
        'Copy over the entire <special-translations> XML block as-is.
        Const END_SPECIAL_BLOCK As String = "</special-translations>"
        Dim posSpecialBlockEnd As Long
        posSpecialBlockEnd = InStr(posSpecialBlock, m_OldLanguageText, END_SPECIAL_BLOCK, vbBinaryCompare)
        If (posSpecialBlockEnd > posSpecialBlock) Then
            
            Dim insertPosition As Long
            insertPosition = InStrRev(m_NewLanguageText, "</pdData>", -1, vbBinaryCompare)
            If (insertPosition > 0) Then m_NewLanguageText = Replace$(m_NewLanguageText, "</pdData>", Mid$(m_OldLanguageText, posSpecialBlock, (posSpecialBlockEnd - posSpecialBlock) + Len(END_SPECIAL_BLOCK)) & vbCrLf & vbCrLf & "</pdData>")
            
        End If
        
    End If
        
    'Prompt the user to save the results
    Dim cDialog As pdOpenSaveDialog
    Set cDialog = New pdOpenSaveDialog
    
    Dim fPath As String
    fPath = m_OldLanguagePath
    
    If cDialog.GetSaveFileName(fPath, , True, "XML - PhotoDemon Language File|*.xml", 1, , "Save the merged language file (XML)", "xml", Me.hWnd) Then
        
        'Worried about breaking something?  Enable strict overwrite checking:
        'If Files.FileExists(fPath) Then
        '    MsgBox "File already exists!  Too dangerous to overwrite - please perform the merge again."
        '    Exit Sub
        'End If
        
        'Use pdXML to write out a UTF-8 encoded XML file
        m_XML.LoadXMLFromString m_NewLanguageText
        m_XML.WriteXMLToFile fPath, True
        
    End If
    
    MsgBox "Merge complete." & vbCrLf & vbCrLf & phrasesProcessed & " phrases processed. " & phrasesFound & " translations found. " & phrasesMissed & " translations missing."

End Sub

'Given a string, return the location of the <phrase> tag enclosing said string
Private Function GetPhraseTagLocation(ByRef srcString As String, Optional ByVal startPos As Long = 1) As Long
    
    GetPhraseTagLocation = 0
    
    Dim sLocation As Long
    sLocation = InStr(startPos, m_OldLanguageText, srcString, vbBinaryCompare)
    
    'If the source string was found, work backward to find the phrase tag location
    If (sLocation > 0) Then
        sLocation = InStrRev(m_OldLanguageText, "<phrase>", sLocation, vbBinaryCompare)
        If (sLocation > 0) Then GetPhraseTagLocation = sLocation
    
    'If given a start location, try searching again from that location, but backward
    Else
        
        If (startPos > 1) Then
            
            sLocation = InStrRev(m_OldLanguageText, srcString, startPos, vbBinaryCompare)
        
            If (sLocation > 0) Then
                sLocation = InStrRev(m_OldLanguageText, "<phrase>", sLocation, vbBinaryCompare)
                If (sLocation > 0) Then GetPhraseTagLocation = sLocation
            End If
            
        End If
        
    End If

End Function

'Given a string, return the location of the <phrase> tag enclosing said string
Private Function GetPhraseTagLocation_CustomFile(ByRef srcString As String, ByRef srcFullText As String, Optional ByVal startPos As Long = 1) As Long
    
    GetPhraseTagLocation_CustomFile = 0
    
    Dim sLocation As Long
    sLocation = InStr(startPos, srcFullText, srcString, vbBinaryCompare)
    
    'If the source string was found, work backward to find the phrase tag location
    If (sLocation > 0) Then
        sLocation = InStrRev(srcFullText, "<phrase>", sLocation, vbBinaryCompare)
        If (sLocation > 0) Then GetPhraseTagLocation_CustomFile = sLocation
    
    'If given a start location, try searching again from that location, but backward
    Else
        
        If (startPos > 1) Then
            
            sLocation = InStrRev(srcFullText, srcString, startPos, vbBinaryCompare)
        
            If (sLocation > 0) Then
                sLocation = InStrRev(srcFullText, "<phrase>", sLocation, vbBinaryCompare)
                If (sLocation > 0) Then GetPhraseTagLocation_CustomFile = sLocation
            End If
            
        End If
        
    End If

End Function

'Given the original caption of a message or control, return the matching translation from the active translation file
Private Function GetTranslationTagFromCaption(ByVal origCaption As String, Optional ByRef inOutWhereToStartSearch As Long = 1) As String
    
    GetTranslationTagFromCaption = vbNullString
    
    'Remove white space from the caption (if necessary, white space will be added back in after retrieving the translation from file)
    PreprocessText origCaption
    origCaption = XML_ORIGINAL_OPEN & origCaption & XML_ORIGINAL_CLOSE
    
    Dim phraseLocation As Long
    phraseLocation = GetPhraseTagLocation(origCaption, inOutWhereToStartSearch)
    
    'Make sure a phrase tag was found
    If (phraseLocation > 0) Then
        
        'Retrieve the <translation> tag inside this phrase tag
        Const TRANSLATION_TAG_NAME As String = "translation"
        GetTranslationTagFromCaption = GetTextBetweenTags(m_OldLanguageText, TRANSLATION_TAG_NAME, phraseLocation)
        inOutWhereToStartSearch = phraseLocation
    Else
        inOutWhereToStartSearch = 1
    End If

End Function

'Given the original caption of a message or control, return the matching translation from the active translation file
Private Function GetTranslationTagFromCaption_CustomFile(ByVal origCaption As String, ByRef srcFullText As String, Optional ByRef inOutWhereToStartSearch As Long = 1) As String
    
    GetTranslationTagFromCaption_CustomFile = vbNullString
    
    'Remove white space from the caption (if necessary, white space will be added back in after retrieving the translation from file)
    PreprocessText origCaption
    origCaption = XML_ORIGINAL_OPEN & origCaption & XML_ORIGINAL_CLOSE
    
    Dim phraseLocation As Long
    phraseLocation = GetPhraseTagLocation_CustomFile(origCaption, srcFullText, inOutWhereToStartSearch)
    
    'Make sure a phrase tag was found
    If (phraseLocation > 0) Then
        
        'Retrieve the <translation> tag inside this phrase tag
        Const TRANSLATION_TAG_NAME As String = "translation"
        GetTranslationTagFromCaption_CustomFile = GetTextBetweenTags(srcFullText, TRANSLATION_TAG_NAME, phraseLocation)
        inOutWhereToStartSearch = phraseLocation
    Else
        inOutWhereToStartSearch = 1
    End If

End Function

'Given a file (as a String) and a tag (without brackets), return the text between that tag.
' NOTE: this function will always return the first occurence of the specified tag, starting at the specified search position.
' If the tag is not found, a blank string will be returned.
Private Function GetTextBetweenTags(ByRef fileText As String, ByRef fTag As String, Optional ByVal searchLocation As Long = 1, Optional ByRef whereTagFound As Long = -1) As String
    
    GetTextBetweenTags = vbNullString
    
    Dim tagStart As Long, tagEnd As Long
    tagStart = InStr(searchLocation, fileText, "<" & fTag & ">", vbBinaryCompare)
    
    'If the tag was found in the file, we also need to find the closing tag.
    If (tagStart > 0) Then
    
        'If the closing tag exists, return everything between that and the opening tag
        tagEnd = InStr(tagStart, fileText, "</" & fTag & ">", vbBinaryCompare)
        If (tagEnd > tagStart) Then
            
            'Increment the tag start location by the length of the tag plus two (+1 for each bracket: <>)
            tagStart = tagStart + Len(fTag) + 2
            
            'If the user passed a long, they want to know where this tag was found - return the location just after the
            ' location where the closing tag was located.
            If (whereTagFound <> -1) Then whereTagFound = tagEnd + Len(fTag) + 2
            GetTextBetweenTags = Mid$(fileText, tagStart, tagEnd - tagStart)
            
        Else
            GetTextBetweenTags = "ERROR: specified tag wasn't properly closed!"
        End If
        
    End If

End Function

Private Sub PreprocessText(ByRef srcString As String)

    '1) Trim the string
    srcString = Trim$(srcString)
    
    '2) Check for a trailing "..."
    If (Right$(srcString, 3) = "...") Then
        srcString = Left$(srcString, Len(srcString) - 3)
    
    '3) Check for a trailing colon ":"
    Else
        If (Right$(srcString, 1) = ":") Then srcString = Left$(srcString, Len(srcString) - 1)
    End If
    
End Sub

'PD's build script uses this function to update all languages files against the latest en-US project text.
Private Sub cmdMergeAll_Click()
    
    Message "Loading en-US text file..."
    
    'Assuming this app instance is in its normal location (/Support/i18n-manager/), calculate a relevant
    ' neighboring folder where language files will be located.
    Dim baseFolder As String
    If Files.PathCanonicalize(Files.AppPathW() & "..\..", baseFolder) Then baseFolder = Files.PathAddBackslash(baseFolder)
    
    Dim srcFolder As String
    srcFolder = baseFolder & "App\PhotoDemon\Languages\"
    
    'Auto-load the latest master language file and remove tabstops from the text (if any exist)
    Files.FileLoadAsString srcFolder & "Master\MASTER.xml", m_AllEnUsText, True
    m_AllEnUsText = Replace$(m_AllEnUsText, vbTab, vbNullString, 1, -1, vbBinaryCompare)
    
    'Rather than backup the old files to the dev language folder (which is confusing),
    ' I now place them inside a dedicated backup folder.
    Dim backupFolder As String
    backupFolder = baseFolder & "no_sync\PD_Language_File_Tmp\dev_backup\"
    If (Not Files.PathExists(backupFolder, True)) Then Files.PathCreate backupFolder, True
    
    'As part of the merge, I want to try and save translations where the en-US text has only been
    ' slightly modified (e.g. fixing a typo).  The merger looks for identical matches to existing en-US text -
    ' - as it should - but sometimes these little changes cause translations to be lost.  Because there's no
    ' easy way to automate the determination of "minor enough change to warrant reusing existing translation",
    ' I instead dump near-matches to a text file and review it manually after modifying en-US text.
    '
    'Because each language has a different set of complete vs incomplete translations, we store all changes
    ' to a single text file *but* calculated on a per-language basis.  Also, we only do this for translations
    ' where a translation exists, *but* the corresponding en-US text is not used in the present file.
    
    'Per-language misses
    Dim curLangMisses As pdStringHash
    Set curLangMisses = New pdStringHash
    
    'Total misses (from *all* languages, stitched together into one file)
    Dim totalMisses As pdString
    Set totalMisses = New pdString
    
    'Iterate through every language file in the default PD directory
    'Scan the translation folder for .xml files.  Ignore anything that isn't XML.
    Dim chkFile As String
    chkFile = Dir$(srcFolder & "*.xml", vbNormal)
    
    'String constants to prevent constant allocations
    Const PHRASE_START As String = "<phrase>"
    Const AMPERSAND_CHAR As String = "&"
    
    Do While (LenB(chkFile) > 0)
        
        Message "Loading " & chkFile & "..."
        
        'On a new language, reset the current misses collection
        curLangMisses.Reset
        
        'Load the target language file into an XML parser, enforce Windows line-endings, and strip any
        ' tab stops from the text.  (PD never uses tab stops in official text because it can cause unpredictable
        ' layout issues, but 3rd-party editors may add tabs as a "convenience" when editing XML.)
        m_OldLanguagePath = srcFolder & chkFile
        Files.FileLoadAsString m_OldLanguagePath, m_OldLanguageText, True
        m_OldLanguageText = Replace$(m_OldLanguageText, vbTab, vbNullString, 1, -1, vbBinaryCompare)
        
        Dim oldLangXML As pdXML
        Set oldLangXML = New pdXML
        oldLangXML.LoadXMLFromString m_OldLanguageText
        oldLangXML.SetTextCompareMode vbBinaryCompare
        
        'Retrieve all phrase tag locations
        Dim phraseLocations() As Long
        oldLangXML.FindAllTagLocations phraseLocations, "phrase"
        
        Dim numOldPhrases As Long
        numOldPhrases = UBound(phraseLocations) + 1
        
        Dim origText As String, translatedText As String
        Dim findText As String, replaceText As String
        
        Const TAG_NAME_ORIG As String = "original"
        Const TAG_NAME_TRNS As String = "translation"
        
        'Build a collection of all phrases in the current translation file.  Some phrases may not be
        ' translated and that's fine - we'll leave them blank and simply plug-in the phrases we *do* have.
        Set m_PhraseCollection = New pdStringHash
        Dim i As Long
        
        If (numOldPhrases > 0) Then
            
            For i = 0 To numOldPhrases - 1
                
                origText = oldLangXML.GetUniqueTag_String(TAG_NAME_ORIG, vbNullString, phraseLocations(i))
                translatedText = oldLangXML.GetUniqueTag_String(TAG_NAME_TRNS, vbNullString, phraseLocations(i) + Len(origText))
                
                'Old PhotoDemon language files used manually inserted & characters for keyboard accelerators.
                ' Accelerators are now handled automatically on a per-language basis.  To ensure work isn't lost
                ' when upgrading these old files, strip any accelerators from the incoming text.
                ' (As of 2024, this change is no longer necessary.)
                'If (InStr(1, origText, AMPERSAND_CHAR, vbBinaryCompare) <> 0) Then origText = Replace$(origText, AMPERSAND_CHAR, vbNullString, 1, -1, vbBinaryCompare)
                'If (InStr(1, translatedText, AMPERSAND_CHAR, vbBinaryCompare) <> 0) Then translatedText = Replace$(translatedText, AMPERSAND_CHAR, vbNullString, 1, -1, vbBinaryCompare)
                
                If (LenB(translatedText) > 0) Then m_PhraseCollection.AddItem origText, translatedText
                
            Next i
            
        End If
        
        'BEGIN COPY OF CODE FROM cmdMerge (with changes to accelerate the process, since we don't need a UI)
        
        'Make sure our source file strings are not empty
        If (LenB(m_AllEnUsText) = 0) Or (numOldPhrases <= 0) Then
            Debug.Print "One or more source files are missing.  Supply those before attempting a merge."
            Exit Sub
        End If
        
        'Start by copying the contents of the master file into the destination string.
        ' We will use that as our base, and update it with the old translations as best we can.
        m_NewLanguageText = m_AllEnUsText
        
        'This table stores phrases that are successfully copied from the source file to the destination file.
        ' (From this, we can produce a list of phrases that were *not* successfully copied.)
        Dim localizedPhrasesHit As pdStringHash
        Set localizedPhrasesHit = New pdStringHash
        
        Dim sPos As Long
        sPos = InStr(1, m_NewLanguageText, PHRASE_START)
        
        'Copy over all top-level language and author information
        ReplaceTopLevelTag "langid", m_AllEnUsText, m_OldLanguageText, m_NewLanguageText
        ReplaceTopLevelTag "langname", m_AllEnUsText, m_OldLanguageText, m_NewLanguageText
        ReplaceTopLevelTag "langstatus", m_AllEnUsText, m_OldLanguageText, m_NewLanguageText
        ReplaceTopLevelTag "author", m_AllEnUsText, m_OldLanguageText, m_NewLanguageText
        ReplaceTopLevelTag "langversion", m_AllEnUsText, m_OldLanguageText, m_NewLanguageText, False
            
        Dim phrasesProcessed As Long, phrasesFound As Long, phrasesMissed As Long
        phrasesProcessed = 0
        phrasesFound = 0
        phrasesMissed = 0
        
        Const ORIG_TAG_CLOSE As String = "</original>" & vbCrLf & "<translation></translation>"
        Const TRANSLATE_TAG_INTERIOR As String = "</original>" & vbCrLf & "<translation>"
        Const TRANSLATE_TAG_CLOSE As String = "</translation>"
            
        'Start parsing the master text for <phrase> tags
        Do
        
            phrasesProcessed = phrasesProcessed + 1
            
            'Retrieve the original text associated with this phrase tag
            origText = GetTextBetweenTags(m_AllEnUsText, TAG_NAME_ORIG, sPos)
            
            'Attempt to retrieve a translation for this phrase using the old language file
            If m_PhraseCollection.GetItemByKey(origText, translatedText) Then
                
                'Remove any tab stops from the translated text (which may have been added by an outside editor)
                If (InStr(1, translatedText, vbTab, vbBinaryCompare) <> 0) Then translatedText = Replace$(translatedText, vbTab, vbNullString, 1, -1, vbBinaryCompare)
                
            Else
                translatedText = vbNullString
            End If
            
            'If a translation was found, insert it into the new file
            If (LenB(translatedText) <> 0) Then
                findText = XML_ORIGINAL_OPEN & origText & ORIG_TAG_CLOSE
                replaceText = XML_ORIGINAL_OPEN & origText & TRANSLATE_TAG_INTERIOR & translatedText & TRANSLATE_TAG_CLOSE
                m_NewLanguageText = Replace$(m_NewLanguageText, findText, replaceText, 1, -1, vbBinaryCompare)
                phrasesFound = phrasesFound + 1
                localizedPhrasesHit.AddItem origText, vbNullString
            Else
                phrasesMissed = phrasesMissed + 1
                
                'Store this phrase to the "miss" list of phrases; we'll try and find approximate matches for
                ' this phrase after we finish parsing this file.
                curLangMisses.AddItem origText, vbNullString
                
            End If
        
            'Find the next occurrence of a <phrase> tag
            sPos = InStr(sPos + 1, m_AllEnUsText, PHRASE_START, vbBinaryCompare)
            
            If ((phrasesProcessed And 15) = 0) Then
                Message chkFile & ": " & phrasesProcessed & " phrases processed (" & phrasesFound & " found, " & phrasesMissed & " missed)"
            End If
            
        Loop While sPos > 0
        
        'All translated phrases with exact en-US matches have now been merged into a new language file.
        
        'Next, we want to look for translations that exist in the old file but their corresponding en-US phrase
        ' does *not* appear in the current en-US language file.  This can happen if I fix a typo or make a trivial
        ' text change, and I do not want to lose translations like this.
        
        'So let's start by looking for any en-US phrases that exist in the translation file but
        ' *not* the latest master en-US phrase list.
        If (curLangMisses.GetNumOfItems > 0) Then
            
            'To ensure text is only written out if at least one phrase is "saved" (e.g. salvaged),
            ' we track how many potential phrases we've "saved".
            Dim numSavesMade As Long
            numSavesMade = 0
            
            'Retrieve a list of *all* phrases in the localized language file.
            ' (Note that we only need the phrases - the translations don't matter in this list.)
            Dim tmpL10nKeys() As String, tmpL10nItems() As String
            m_PhraseCollection.GetAllItems tmpL10nKeys, tmpL10nItems
            Erase tmpL10nItems
            
            'Next, we want to build a list of translations that existed in the source translation file
            ' but were *not* transferred to the new merged translation file.
            Dim unusedPhrases() As String, numUnusedPhrases As Long
            numUnusedPhrases = 0
            
            Const INIT_SIZE_UNUSED_PHRASES As Long = 128
            ReDim unusedPhrases(0 To INIT_SIZE_UNUSED_PHRASES - 1) As String
            
            For i = 0 To UBound(tmpL10nKeys)
                
                'See if this phrase (from the localized file) was matched to an active en-US phrase
                If (Not localizedPhrasesHit.GetItemByKey(tmpL10nKeys(i), vbNullString)) Then
                    If (numUnusedPhrases > UBound(unusedPhrases)) Then ReDim Preserve unusedPhrases(0 To numUnusedPhrases * 2 - 1) As String
                    unusedPhrases(numUnusedPhrases) = tmpL10nKeys(i)
                    numUnusedPhrases = numUnusedPhrases + 1
                End If
                
            Next i
            
            'Grab the list of en-US phrases that didn't have a matching translation in this language.
            ' (Note that the list of items is unused here - they will just be null strings, by design.)
            Dim untranslatedPhrases() As String
            curLangMisses.GetAllItems untranslatedPhrases, tmpL10nItems
            curLangMisses.Reset
            
            'Sometimes I need to fix typos or slightly reword en-US text.  Unfortunately, minor changes like this
            ' will cause the translated versions of these phrases - if any - to be lost, because the en-US "key" phrase
            ' in the translation file will no longer match the modified en-US text in the "master" language file.
            Dim numRescuedPhrases As Long
            numRescuedPhrases = 0
            
            'We're going to try to fix this by searching for phrases that are "similar but not identical".  If they
            ' exceed an arbitrary threshold, we'll dump them to file so I can manually review and copy+paste
            ' translations that should be retained.
            For i = 0 To UBound(untranslatedPhrases)
                
                Dim distMin As Long, distCur As Long, idxBestMatch As Long
                distMin = LONG_MAX
                idxBestMatch = -1
                
                'Failsafe check for non-null phrases
                Dim lenOrig As Long
                lenOrig = Len(untranslatedPhrases(i))
                
                If (lenOrig > 0) Then
                    
                    'We now want to iterate all phrases in the "unused phrases" list to find the best match.
                    Dim j As Long
                    For j = 0 To numUnusedPhrases - 1
                        
                        'Don't compare strings unless their total length is 80% similar
                        Const LENGTH_SIMILARITY_THRESHOLD As Double = 0.2
                        If ((Abs(lenOrig - Len(unusedPhrases(j))) / lenOrig) <= LENGTH_SIMILARITY_THRESHOLD) Then
                            
                            'Calculate distance, and only treat it as relevant if the target phrases are at least
                            ' 75% similar.  (This threshold is effectively arbitrary; it's meant to filter out
                            ' low-quality matches to avoid wasting my time during manual review.)
                            distCur = Strings.StringDistance(untranslatedPhrases(i), unusedPhrases(j), True)
                            
                            Const DISTANCE_SIMILARITY_THRESHOLD As Double = 0.25
                            If (distCur < distMin) Then
                                If ((distCur / lenOrig) < DISTANCE_SIMILARITY_THRESHOLD) Then
                                    distMin = distCur
                                    idxBestMatch = j
                                End If
                            End If
                            
                        End If
                            
                    Next j
                    
                End If
                    
                'Write the best match (and its associated translation) out to the merged report
                If (distMin < LONG_MAX) And (idxBestMatch >= 0) Then
                    
                    translatedText = vbNullString
                    If m_PhraseCollection.GetItemByKey(unusedPhrases(idxBestMatch), translatedText) And (LenB(translatedText) > 0) Then
                        
                        'If this is the first translation we've "saved" in this file, write a header first
                        If (numSavesMade = 0) Then
                            totalMisses.AppendLineBreak
                            totalMisses.AppendLine String$(32, "*")
                            totalMisses.AppendLine "Best-match report for " & chkFile
                            totalMisses.AppendLine String$(32, "*")
                            totalMisses.AppendLineBreak
                        End If
                        
                        'If the score is 0, it means the translation file has a translation for an en-US phrase
                        ' that is identical to the target phrase *except* for casing.  This is 100% okay -
                        ' I probably just tweaked something in the source code, and we should just use the existing
                        ' translation as-is.  (Note that we do not try to propagate the case change to the translated
                        ' text because the rules for this vary by language and it might affect meaning in unintended ways!)
                        '
                        'Similarly, if the score is 1, it often just means that a space was added somewhere to
                        ' the original text.  (This is common if the translator used an automated tool.)
                        ' Single-character differences shouldn't trigger any meaningful problems, so retain this
                        ' translation and replace the offending text with the expected one.
                        If (distMin = 0) Or (distMin = 1) Then
                        
                            findText = XML_ORIGINAL_OPEN & untranslatedPhrases(i) & ORIG_TAG_CLOSE
                            replaceText = XML_ORIGINAL_OPEN & untranslatedPhrases(i) & TRANSLATE_TAG_INTERIOR & translatedText & TRANSLATE_TAG_CLOSE
                            m_NewLanguageText = Replace$(m_NewLanguageText, findText, replaceText, 1, -1, vbBinaryCompare)
                            numRescuedPhrases = numRescuedPhrases + 1
                        
                        'If the score is *not* 0 or 1, it means a translation exists for a similar - but *not* identical -
                        ' en-US phrase.  These cases require manual review because it's not always obvious whether the
                        ' two phrases are close enough to matter.
                        Else
                            totalMisses.AppendLineBreak
                            totalMisses.AppendLine "New text: " & untranslatedPhrases(i)
                            totalMisses.AppendLine "Old text: " & unusedPhrases(idxBestMatch)
                            totalMisses.AppendLine "Match distance: " & distMin
                            totalMisses.AppendLine "<original>" & untranslatedPhrases(i) & "</original>"
                            totalMisses.AppendLine "<translation>" & translatedText & "</translation>"
                            totalMisses.AppendLineBreak
                        End If
                        
                        numSavesMade = numSavesMade + 1
                        
                    End If
                        
                End If
                
                If ((i And 15) = 0) Then
                    Message chkFile & ": " & (i + 1) & " of " & (UBound(untranslatedPhrases) + 1) & " missing phrases estimated, " & numSavesMade & " potentially salvaged)"
                End If
                
            Next i
            
        End If
        
        'All exact translations have now been merged, and near-exact translations have been written out to a
        ' text file for human review.
        If (numRescuedPhrases > 0) Then Debug.Print "NOTE: " & numRescuedPhrases & " near-identical phrases were rescued!"
        
        'Finally, look for any language-specific text+translation pairs.  This (optional) segment can be used
        ' to map phrases with identical English text (e.g. Color > Invert vs Selection > Invert) to unique
        ' phrases in a given translation.
        Dim posSpecialBlock As Long
        posSpecialBlock = InStrRev(m_OldLanguageText, "<special-translations>", -1, vbBinaryCompare)
        If (posSpecialBlock > 0) Then
        
            'Copy over the entire <special-translations> XML block as-is.
            Const END_SPECIAL_BLOCK As String = "</special-translations>"
            Dim posSpecialBlockEnd As Long
            posSpecialBlockEnd = InStr(posSpecialBlock, m_OldLanguageText, END_SPECIAL_BLOCK, vbBinaryCompare)
            If (posSpecialBlockEnd > posSpecialBlock) Then
                
                Dim insertPosition As Long
                insertPosition = InStrRev(m_NewLanguageText, "</pdData>", -1, vbBinaryCompare)
                If (insertPosition > 0) Then m_NewLanguageText = Replace$(m_NewLanguageText, "</pdData>", Mid$(m_OldLanguageText, posSpecialBlock, (posSpecialBlockEnd - posSpecialBlock) + Len(END_SPECIAL_BLOCK)) & vbCrLf & vbCrLf & "</pdData>")
                
            End If
            
        End If
        
        'We can now save the final merged text out to file.
        
        'See if the old and new language files are equal.  If they are, we won't bother writing the results out to file.
        If (LenB(Trim$(m_NewLanguageText)) = LenB(Trim$(m_OldLanguageText))) Then
            Debug.Print "New language file and old language file are identical for " & chkFile & ".  Merge abandoned."
        Else
            
            'Update the version number by 1
            ReplaceTopLevelTag "langversion", m_AllEnUsText, m_OldLanguageText, m_NewLanguageText
            
            'Unlike the normal merge option, we will automatically save the results to a new XML file
            
            'Start by backing up the old file
            Files.FileDeleteIfExists backupFolder & chkFile
            Files.FileCopyW m_OldLanguagePath, backupFolder & chkFile
            
            If Files.FileExists(m_OldLanguagePath) Then
                Debug.Print "Note - old file with same name (" & m_OldLanguagePath & ") will be erased.  Hope this is what you wanted!"
            End If
            
            'Use pdXML to write out a UTF-8 encoded XML file
            m_XML.LoadXMLFromString m_NewLanguageText
            m_XML.WriteXMLToFile m_OldLanguagePath, True
            
        End If
        
        'END COPY OF CODE FROM cmdMerge
        
        'Retrieve the next file and repeat
        chkFile = Dir
        
    Loop
    
    'If some translations were lost (usually due to me making changes to existing en-US text), dump any
    ' suggested translation replacements to file for later human review.
    If (LenB(totalMisses.ToString) > 0) Then Files.FileSaveAsText totalMisses.ToString, srcFolder & "final_report.txt", True, True
    
    Message "All language files processed successfully."
    
End Sub

Private Sub cmdOldLanguage_Click()
    
    Dim cDialog As pdOpenSaveDialog
    Set cDialog = New pdOpenSaveDialog
    
    'Assuming this app instance is in its normal location (/Support/i18n-manager/), calculate a relevant
    ' neighboring folder where language files will be located.
    Dim likelyDefaultLocation As String
    If Files.PathCanonicalize(Files.AppPathW() & "..\..", likelyDefaultLocation) Then likelyDefaultLocation = Files.PathAddBackslash(likelyDefaultLocation)
    likelyDefaultLocation = likelyDefaultLocation & "App\PhotoDemon\Languages\"
    
    Dim tmpLangFile As String
    If cDialog.GetOpenFileName(tmpLangFile, , True, False, "XML - PhotoDemon Language File|*.xml", 1, likelyDefaultLocation, "Please select a PhotoDemon language file (XML)", "xml", Me.hWnd) Then
        
        m_OldLanguagePath = tmpLangFile
        
        'Load the language file and strip tab chars from it
        Files.FileLoadAsString tmpLangFile, m_OldLanguageText, True
        m_OldLanguageText = Replace$(m_OldLanguageText, vbTab, vbNullString, 1, -1, vbBinaryCompare)
        
    End If
    
End Sub

'Process all files in a project file.  (NOTE: a VBP file must first be selected before running this step.)
Private Sub cmdProcess_Click()

    If (LenB(m_VBPFile) = 0) Then
        MsgBox "Select a VBP file first.", vbExclamation + vbApplicationModal + vbOKOnly, "Oops"
        Exit Sub
    End If
    
    'Note whether duplicate phrases are automatically removed.
    ' (In production, duplicate phrases are *always* removed.)
    m_RemoveDuplicates = CBool(chkRemoveDuplicates)
    
    'Reset the existing phrase collection, if any
    m_enUSPhrases.Reset
    m_numPhraseData = 0
    
    'Start by preparing a generic XML header
    Set m_outputText = New pdString
    m_outputText.AppendLine "<?xml version=""1.0"" encoding=""UTF-8""?>"
    m_outputText.AppendLineBreak
    m_outputText.AppendLine vbTab & "<pdData>"
    m_outputText.AppendLineBreak
    m_outputText.AppendLine vbTab & vbTab & "<pdDataType>Translation</pdDataType>"
    m_outputText.AppendLineBreak
    m_outputText.AppendLine vbTab & vbTab & "<langid>en-US</langid>"
    m_outputText.AppendLine vbTab & vbTab & "<langname>English (US) - MASTER COPY</langname>"
    m_outputText.AppendLine vbTab & vbTab & "<langversion>" & m_VersionString & "</langversion>"
    m_outputText.AppendLine vbTab & vbTab & "<langstatus>Automatically generated from PhotoDemon's source code</langstatus>"
    m_outputText.AppendLineBreak
    m_outputText.AppendLine vbTab & vbTab & "<author>Tanner Helland</author>"
    m_outputText.AppendLineBreak
    m_outputText.Append vbTab & vbTab & "<!-- BEGIN AUTOMATIC TEXT GENERATION -->"
    
    Dim numOfFiles As Long
    numOfFiles = UBound(m_vbpFiles)
    
    m_NumOfPhrasesFound = 0
    m_NumOfPhrasesWritten = 0
    m_numOfWords = 0
    
    Dim i As Long
    For i = 0 To numOfFiles
        cmdProcess.Caption = "Processing project file " & i + 1 & " of " & numOfFiles + 1
        ProcessFile m_vbpFiles(i)
    Next i
    
    'With processing complete, write out our final stats (just for fun)
    m_outputText.AppendLineBreak
    m_outputText.AppendLineBreak
    m_outputText.AppendLine vbTab & vbTab & "<!-- Automatic text generation complete. -->"
    m_outputText.AppendLineBreak
    m_outputText.AppendLine vbTab & vbTab & "<phrasecount>" & m_NumOfPhrasesWritten & "</phrasecount>"
    m_outputText.AppendLineBreak
    
    'Proceed with human-readable phrase statistics
    If CBool(chkRemoveDuplicates) Then
        m_outputText.AppendLine vbTab & "<!-- As of this build, PhotoDemon contains " & m_NumOfPhrasesFound & " phrases. -->"
        m_outputText.AppendLine vbTab & "<!-- " & CStr(m_NumOfPhrasesFound - m_NumOfPhrasesWritten) & " are duplicates, so only " & m_NumOfPhrasesWritten & " unique phrases have been written to file. -->"
        m_outputText.AppendLine vbTab & "<!-- These " & m_NumOfPhrasesWritten & " phrases contain approximately " & m_numOfWords & " total words. -->"
    Else
        m_outputText.AppendLine vbTab & "<!-- As of this build, PhotoDemon contains " & m_NumOfPhrasesWritten & " phrases (including duplicates). -->"
        m_outputText.AppendLine vbTab & "<!-- These " & m_NumOfPhrasesWritten & " phrases contain approximately " & m_numOfWords & " total words. -->"
    End If
    
    'Terminate the final language tag
    m_outputText.AppendLineBreak
    m_outputText.Append vbTab & "</pdData>"
    
    'Write the text out to file
    If CBool(chkRemoveDuplicates) Then
        outputFile = m_VBPPath & "App\PhotoDemon\Languages\Master\MASTER.xml"
    Else
        outputFile = m_VBPPath & "App\PhotoDemon\Languages\Master\MASTER (with duplicates).xml"
    End If
    
    'We are now going to compare the length of the old file and new file.
    ' If the lengths match, there's no reason to write out this new file.
    Dim oldFileString As String
    Files.FileLoadAsString outputFile, oldFileString, True
    
    Dim newFileLen As Long, oldFileLen As Long
    newFileLen = LenB(Trim$(Replace$(Replace$(m_outputText.ToString(), vbCrLf, vbNullString), vbTab, vbNullString)))
    oldFileLen = LenB(Trim$(Replace$(Replace$(oldFileString, vbCrLf, vbNullString), vbTab, vbNullString)))
    
    If (newFileLen <> oldFileLen) Then
        
        'Use pdXML to write a UTF-8 encoded text file
        m_XML.LoadXMLFromString m_outputText.ToString()
        m_XML.WriteXMLToFile outputFile, True
        
        cmdProcess.Caption = "Processing complete!"
        
    Else
        cmdProcess.Caption = "Processing complete (no changes made)"
    End If
    
    'Finally, write out an updated phrase database.
    If (m_numPhraseData = 0) Then Exit Sub
    
    Dim cStream As pdStream
    Set cStream = New pdStream
    
    Dim dbFilePath As String
    dbFilePath = m_VBPPath & "App\PhotoDemon\Languages\Master\Phrases.db"
    Files.FileDeleteIfExists dbFilePath
    
    If cStream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadWrite, dbFilePath) Then
        
        'Write the number of phrases, then all phrases in sequence
        Const PHRASE_DB_ID As String = "pdPhraseDB"
        cStream.WriteString_ASCII PHRASE_DB_ID
        
        Const PHRASE_DB_VERSION As Long = 1
        cStream.WriteLong PHRASE_DB_VERSION
        
        cStream.WriteLong m_numPhraseData
        
        For i = 0 To m_numPhraseData - 1
            With m_phraseData(i)
                cStream.WriteByte .phraseType
                cStream.WriteIntU Len(.origEnUSPhrase)
                cStream.WriteString_UTF8 .origEnUSPhrase
                cStream.WriteIntU Len(.occursInFiles)
                cStream.WriteString_UTF8 .occursInFiles
            End With
        Next i
        
        cStream.StopStream True
        
    End If
    
End Sub

'Given a VB file (form, module, class, user control), extract any relevant text from it
Private Sub ProcessFile(ByRef srcFile As String)

    If (LenB(srcFile) = 0) Then Exit Sub

    m_FileName = Files.FileGetName(srcFile)
    
    'Certain files can be ignored.  I generate this list manually, on account of knowing which files (classes, mostly) contain
    ' no special text.  I could probably add many more files to this list, but I primarily want to blacklist those that create
    ' parsing problems.  (The tooltip classes are particularly bad, since they use the phrase "tooltip" frequently, which messes
    ' up the parser as it thinks it's found hundreds of tooltips in each file.)
    Select Case m_FileName
    
        Case "pdToolTip.cls", "pdFilterSupport.cls", "pdParamString.cls"
            Exit Sub
            
        Case "pdButtonStrip.ctl", "pdButtonStripVertical.ctl"
            Exit Sub
            
        Case "Misc_Tooltip.frm"
            Exit Sub
            
        'Some developer-only dialogs do not need to be translated.
        Case "Tools_ThemeEditor.frm", "Tools_BuildPackage.frm"
            Exit Sub
    
    End Select
            
    
    'Start by copying all text from the file into a line-by-line array
    Dim fileContents As String
    Files.FileLoadAsString srcFile, fileContents, True
    
    Dim fileLines() As String
    fileLines = Split(fileContents, vbCrLf)
    
    'If this file is a form file, the second line of the file will contain the text: "Begin VB.FORM FormName",
    ' where FormName is the name of the form (as it appears in the VB IDE).  This engine inserts that form
    ' name into an XML comment to help translators know where to find text in the current build.
    Dim shortcutName As String
    shortcutName = vbNullString
    
    If Right$(m_FileName, 3) = "frm" Then
        Dim findName() As String
        findName = Split(fileLines(1), " ")
        shortcutName = findName(2)
    End If
    
    'Initialize the phrase metadata array (if it hasn't been initialized already)
    Const INIT_PHRASE_METADATA_COUNT As Long = 4096
    If (m_numPhraseData = 0) Then ReDim m_phraseData(0 To INIT_PHRASE_METADATA_COUNT - 1) As PD_PhraseInfo
    
    'For convenience, write the name of the source file into the translation file - this can be helpful when
    ' tracking down errors or incomplete text.
    'If (LenB(m_FileName) > 0) Then
    '    m_outputText.AppendLineBreak
    '    m_outputText.Append vbTab & vbTab
    '    If (LenB(shortcutName) <> 0) Then
    '        m_outputText.Append "<!-- BEGIN text for " & m_FileName & " (" & shortcutName & ") -->"
    '    Else
    '        m_outputText.Append "<!-- BEGIN text for " & m_FileName & " -->"
    '    End If
    'End If
    
    Dim curLineNumber As Long
    curLineNumber = 0
    
    Dim numOfPhrasesFound As Long, numOfPhrasesWritten As Long
    numOfPhrasesFound = 0
    numOfPhrasesWritten = 0
    
    Dim curLineText As String, ucCurLineText As String, processedText As String, processedTextSecondary As String, chkText As String
    m_FormName = vbNullString
    
    Dim toolTipSecondCheckNeeded As Boolean
        
    'Now, start processing the file one line at a time, searching for relevant text as we go
    Do
    
        processedText = vbNullString
        processedTextSecondary = vbNullString
    
        curLineText = fileLines(curLineNumber)
        
        'Before processing this line, make sure is isn't a comment.  (Comments are always ignored.)
        If (Left$(Trim$(curLineText), 1) = "'") Then GoTo nextLine
        
        'Make a copy of the upper-case version of the line; we'll use that for case-invariant comparisons
        ucCurLineText = UCase$(curLineText)
        
        'There are many ways that translatable text may appear in a VB source file.
        ' 1) As a form caption
        ' 2) As a control caption
        ' 3) As tooltip text
        ' 4) As text added to a combo box or list box control at run-time (e.g. "control.AddItem "xyz")
        ' 5) As a message call (e.g. Message "xyz")
        ' 6) As message box text, specifically pdMsgBox (e.g. one of either pdMsgBox("xyz"...) or pdMsgBox "xyz"...)
        ' 7) As a message box title caption (more convoluted to find - basically the 3rd parameter of a pdMsgBox call)
        ' 8) As miscellaneous text manually marked for translation (e.g. g_Language.translateMessage("xyz"))
        ' 9) As miscellaneous tooltip text manually marked for translation by the AssignTooltip function.
        ' 10) Process calls, which are relayed to the user in the Undo / Redo menus (e.g. "Undo Blur")
        ' (in some rare cases, text may appear that doesn't fit any of these cases - such text must be added manually)
        
        'Every one of these requires a unique mechanism for checking the text.  Note that we also track the
        ' current phrase type as part of building a phrase "database"; translators can use this to
        ' prioritize their translation work.
        Dim curPhraseType As PD_PhraseType
        curPhraseType = pt_UIElement
        
        'Note that some of these mechanisms will modify the current line number.  These require the line number, passed
        ' ByRef, for that purpose.
        
        'If any of the functions are successful, they will return the string that needs to be added to the XML file
        
        '1) Check for a form caption
        If InStr(1, ucCurLineText, "BEGIN VB.FORM", vbBinaryCompare) Then
            processedText = FindFormCaption(fileLines, curLineNumber)
            curPhraseType = pt_UIElement
                
        '2) Check for a control caption.  (This has to be handled slightly differently than form caption.)
        ElseIf ((InStr(1, ucCurLineText, "BEGIN VB.", vbBinaryCompare) > 0) Or (InStr(1, ucCurLineText, "BEGIN PHOTODEMON.", vbBinaryCompare) > 0)) And (InStr(1, ucCurLineText, "PICTUREBOX", vbBinaryCompare) = 0) And (InStr(1, curLineText, "ComboBox") = 0) And (InStr(1, curLineText, ".Shape") = 0) And (InStr(1, curLineText, "TextBox") = 0) And (InStr(1, curLineText, "HScrollBar") = 0) And (InStr(1, curLineText, "VScrollBar") = 0) Then
            processedText = FindControlCaption(fileLines, curLineNumber)
            curPhraseType = pt_UIElement
        
        '3) Check for tooltip text on PD controls (assigned via the custom .AssignTooltip function)
        ElseIf (InStr(1, ucCurLineText, ".ASSIGNTOOLTIP ") > 0) And (InStr(1, curLineText, "ByVal") = 0) Then
            
            'Process the tooltip text itself
            processedText = FindTooltipMessage(fileLines, curLineNumber, False, toolTipSecondCheckNeeded)
            
            'Process the title, if any
            If toolTipSecondCheckNeeded Then processedTextSecondary = FindMsgBoxTitle(fileLines, curLineNumber)
            
            curPhraseType = pt_Tooltip
            
        '4) Check for text added to a combo box or list box control at run-time
        ElseIf InStr(1, curLineText, ".AddItem """) <> 0 Then
            processedText = FindCaptionInComplexQuotes(fileLines, curLineNumber)
            curPhraseType = pt_UIElement
            
        '5) Check for message calls
        ElseIf InStr(1, curLineText, "Message """) <> 0 Then
            processedText = FindCaptionInQuotes(fileLines, curLineNumber)
            curPhraseType = pt_StatusBar
            
        '6) Check for message box text, including 7) message box titles (which must also be translated)
        ElseIf (InStr(1, ucCurLineText, "PDMSGBOX", vbTextCompare) <> 0) Then
        
            'First, retrieve the message box text itself
            processedText = FindMsgBoxText(fileLines, curLineNumber)
            
            'Next, retrieve the message box title
            processedTextSecondary = FindMsgBoxTitle(fileLines, curLineNumber)
            curPhraseType = pt_MsgBox
            
        '7) Specific to PhotoDemon - check for action names that may not be present elsewhere
        ElseIf InStr(1, curLineText, "Process """) <> 0 Then
            processedText = FindCaptionInQuotes(fileLines, curLineNumber, InStr(1, curLineText, "Process """))
            curPhraseType = pt_ActionName
            
        '7.5) Now that various PD-specific objects manage their own translations, we should also check for
        '     run-time caption assignments
        ElseIf InStr(1, curLineText, "Caption = """, vbBinaryCompare) <> 0 Then
            processedText = FindCaptionInQuotes(fileLines, curLineNumber, 1)
            curPhraseType = pt_UIElement
            
        End If
        
        '8) Check for text that has been manually marked for translation (e.g. g_Language.TranslateMessage("xyz"))
        '    NOTE: as of 07 June 2013, each line can contain two translation calls (instead of just one)
        '
        'Note that this check is performed regardless of previous checks, to make sure no translations are missed.
        If InStr(1, curLineText, "g_Language.TranslateMessage(""") Then
            
            processedText = FindMessage(fileLines, curLineNumber)
            processedTextSecondary = FindMessage(fileLines, curLineNumber, True)
            
            'Note that standalone translation requests like this are impossible to universally categorize,
            ' so we assume maximum importance (since these phrases *may* be UI elements)
            curPhraseType = pt_Miscellaneous
            
        End If
        
        'We now have text in potentially two places: processedText, and processedTextSecondary (for message box titles)
        chkText = Trim$(processedText)
        
        'Only pass the text along if it isn't blank, or a number, or a symbol, or a manually blacklisted phrase
        If (LenB(chkText) <> 0) Then
            If (Not IsNumeric(chkText)) And (Not IsNumericPercentage(chkText)) And (Not IsBlacklisted(chkText)) Then
                If (chkText <> ".") And (chkText <> "-") And (Not IsURL(chkText)) Then
                    numOfPhrasesFound = numOfPhrasesFound + 1
                    If AddPhrase(processedText, curPhraseType, m_FileName) Then numOfPhrasesWritten = numOfPhrasesWritten + 1
                End If
            End If
        End If
        
        chkText = Trim$(processedTextSecondary)
        
        'Do the same for the secondary text
        If (LenB(chkText) <> 0) Then
            If (Not IsNumeric(chkText)) And (Not IsNumericPercentage(chkText)) And (Not IsBlacklisted(chkText)) Then
                If (chkText <> ".") And (chkText <> "-") And (Not IsURL(chkText)) Then
                    numOfPhrasesFound = numOfPhrasesFound + 1
                    If AddPhrase(processedTextSecondary, curPhraseType, m_FileName) Then numOfPhrasesWritten = numOfPhrasesWritten + 1
                End If
            End If
        End If
    
nextLine:
        curLineNumber = curLineNumber + 1
    
    Loop While curLineNumber < UBound(fileLines)
    
    'Now that all phrases in this file have been processed, we can wrap up this section of XML
    
    'For fun, write some stats about our processing results into the translation file.
    ' (But only if we actually found text inside this file.)
    If (LenB(m_FileName) <> 0) And (numOfPhrasesFound > 0) Then
        
        m_outputText.AppendLineBreak
        m_outputText.AppendLineBreak
        m_outputText.Append vbTab & vbTab
        If (numOfPhrasesFound <> 1) Then
            m_outputText.Append "<!-- " & m_FileName & " contains " & numOfPhrasesFound & " phrases. "
        Else
            m_outputText.Append "<!-- " & m_FileName & " contains " & numOfPhrasesFound & " phrase. "
        End If
        
        If (numOfPhrasesFound > 0) Then
            If (numOfPhrasesWritten <> numOfPhrasesFound) Then
                
                Dim phraseDifference As Long
                phraseDifference = numOfPhrasesFound - numOfPhrasesWritten
                
                Select Case phraseDifference
                    Case 1
                        m_outputText.Append " One was a duplicate of an existing phrase, so only " & numOfPhrasesWritten & " new phrases were written to file. -->"
                    Case numOfPhrasesFound
                        m_outputText.Append " All were duplicates of existing phrases, so no new phrases were written to file. -->"
                    Case Else
                        If (numOfPhrasesWritten = 1) Then
                            m_outputText.Append CStr(phraseDifference) & " were duplicates of existing phrases, so only one new phrase was written to file. -->"
                        Else
                            m_outputText.Append CStr(phraseDifference) & " were duplicates of existing phrases, so only " & numOfPhrasesWritten & " new phrases were written to file. -->"
                        End If
                End Select
                
            Else
                Select Case numOfPhrasesFound
                    Case 1
                        m_outputText.Append " The phrase was unique, so 1 new phrase was written to file. -->"
                    Case 2
                        m_outputText.Append " Both phrases were unique, so " & numOfPhrasesFound & " new phrases were written to file. -->"
                    Case Else
                        m_outputText.Append " All " & numOfPhrasesFound & " were unique, so " & numOfPhrasesFound & " new phrases were written to file. -->"
                End Select
            End If
        Else
            m_outputText.Append "-->"
        End If
    End If
    
    'For convenience, once again write the name of the source file into the translation file - this can be helpful when
    ' tracking down errors or incomplete text.
    'If (LenB(m_FileName) <> 0) Then
    '    m_outputText.AppendLineBreak
    '    m_outputText.AppendLineBreak
    '    m_outputText.Append vbTab & vbTab & "<!-- END text for " & m_FileName & "-->"
    'End If
    
    'Add the number of phrases found and written to the master count
    m_NumOfPhrasesFound = m_NumOfPhrasesFound + numOfPhrasesFound
    m_NumOfPhrasesWritten = m_NumOfPhrasesWritten + numOfPhrasesWritten

End Sub

'Add a discovered phrase to the XML file.  If this phrase already exists in the file, ignore it.
Private Function AddPhrase(ByRef phraseText As String, ByRef curPhraseType As PD_PhraseType, ByRef parentFilename As String) As Boolean
    
    'Replace double double-quotes (which are required in code) with just one set of double-quotes
    If (InStr(1, phraseText, """""", vbBinaryCompare) <> 0) Then phraseText = Replace$(phraseText, """""", """", 1, -1, vbBinaryCompare)
    
    'Next, do the same pre-processing that we do in the translation engine
    
    '1) Trim the text.  Extra spaces will be handled by the translation engine.
    phraseText = Trim$(phraseText)
    
    '2) Check for a trailing "..." and remove it
    If (Right$(phraseText, 3) = "...") Then phraseText = Left$(phraseText, Len(phraseText) - 3)
    
    '3) Check for a trailing colon ":" and remove it
    If (Right$(phraseText, 1) = ":") Then phraseText = Left$(phraseText, Len(phraseText) - 1)
    
    'Perform a final failsafe check for null-length phrases.
    If (LenB(phraseText) = 0) Then
        AddPhrase = False
        Exit Function
    End If
    
    'This phrase is now ready to write to file.
    
    'Before writing the phrase out, check to see if it already exists.
    ' (By default, PD suppresses duplicate entries.)
    '
    'If the phrase has *already* been handled, this call will place the index of the item in the phrase
    ' tracking array into prevItemValue (which yes, is a string; this is just so I don't have a write a
    ' separate hash table implementation!)
    Dim prevItemValue As String, idxItem As Long
    If m_RemoveDuplicates Then AddPhrase = Not m_enUSPhrases.GetItemByKey(phraseText, prevItemValue)
    
    'If the phrase does not exist, add it now
    If AddPhrase Then
    
        'Physically place this tag in the output XML
        m_outputText.AppendLineBreak
        m_outputText.AppendLineBreak
        m_outputText.AppendLine XML_PHRASE_OPEN
        m_outputText.Append XML_ORIGINAL_OPEN
        m_outputText.Append phraseText
        m_outputText.AppendLine XML_ORIGINAL_CLOSE
        m_outputText.AppendLine XML_TRANSLATION_PAIR
        m_outputText.Append XML_PHRASE_CLOSE
        
        'Add the phrase to our running "duplicate phrase" detector, and make the associated item an index
        ' into the phrase data collection.
        m_enUSPhrases.AddItem phraseText, Trim$(Str$(m_numPhraseData))
        
        If (m_numPhraseData > UBound(m_phraseData)) Then ReDim Preserve m_phraseData(0 To m_numPhraseData * 2 - 1) As PD_PhraseInfo
        With m_phraseData(m_numPhraseData)
            .phraseType = curPhraseType
            .origEnUSPhrase = phraseText
            .occursInFiles = parentFilename
        End With
        
        m_numPhraseData = m_numPhraseData + 1
        
        'Keep a running tally of total word count (approximately)
        m_numOfWords = m_numOfWords + CountWordsInString(phraseText)
        
    'If this phrase is a duplicate, we still want to update its tracking data.  (This is especially
    ' important to ensure that a phrase's "type" collects all flags for appearance type, if it appears
    ' in multiple places throughout PD.)
    Else
    
        idxItem = Val(prevItemValue)
        With m_phraseData(idxItem)
            
            'Store the most-important phrase type encountered
            .phraseType = .phraseType Or curPhraseType
            
            'Update the list of files in which this phrase appears
            If (Right$(.occursInFiles, Len(parentFilename)) <> parentFilename) Then .occursInFiles = .occursInFiles & ", " & parentFilename
            
        End With
        
    End If
    
End Function

'Given a line number and the original file contents, search for a custom PhotoDemon translation request.
Private Function FindMessage(ByRef srcLines() As String, ByRef lineNumber As Long, Optional ByVal inReverse As Boolean = False) As String
    
    'Finding the text of the message is tricky, because it may be spliced between multiple quotations.  As an example, I frequently
    ' add manual line breaks to messages via " & vbCrLf & " - these need to be checked for and replaced.
    
    'The scan will work by looping through the string, and tracking whether or not we are currently inside quotation marks.
    'If we are outside a set of quotes, and we encounter a comma or closing parentheses, we know that we have reached the end of the
    ' first (and/or only) parameter.
    Const sQuot As String = """"
    
    Dim initPosition As Long
    If inReverse Then
        initPosition = InStrRev(srcLines(lineNumber), "g_Language.TranslateMessage(""")
    Else
        initPosition = InStr(1, srcLines(lineNumber), "g_Language.TranslateMessage(""")
    End If
    
    Dim startQuote As Long
    startQuote = InStr(initPosition, srcLines(lineNumber), sQuot)
    
    Dim endQuote As Long
    endQuote = -1
    
    Dim insideQuotes As Boolean
    insideQuotes = True
    
    Dim i As Long
    For i = startQuote + 1 To Len(srcLines(lineNumber))
    
        If Mid$(srcLines(lineNumber), i, 1) = sQuot Then insideQuotes = Not insideQuotes
        
        If ((Mid$(srcLines(lineNumber), i, 1) = ",") Or (Mid$(srcLines(lineNumber), i, 1) = ")")) And (Not insideQuotes) Then
            endQuote = i - 1
            Exit For
        End If
    
    Next i
    
    'If endQuote = -1, something went horribly wrong
    If endQuote = -1 Then
        FindMessage = "MANUAL FIX REQUIRED FOR MESSAGE PARSE ERROR AT LINE # " & lineNumber & " IN " & m_FileName
    Else
        FindMessage = Mid$(srcLines(lineNumber), startQuote + 1, endQuote - startQuote - 1)
    End If
    
    'We now need to replace line breaks in the text.  These can appear in a variety of ways.  Replace them all.
    Dim lineBreak As String
    lineBreak = """ & vbCrLf & """
    If (InStr(1, FindMessage, lineBreak, vbBinaryCompare) <> 0) Then FindMessage = Replace(FindMessage, lineBreak, vbCrLf)
    lineBreak = """ & vbCrLf & vbCrLf & """
    If (InStr(1, FindMessage, lineBreak, vbBinaryCompare) <> 0) Then FindMessage = Replace(FindMessage, lineBreak, vbCrLf & vbCrLf)
    
End Function

'Given a line number and the original file contents, search for a custom PhotoDemon tooltip assignment
Private Function FindTooltipMessage(ByRef srcLines() As String, ByRef lineNumber As Long, Optional ByVal inReverse As Boolean = False, Optional ByRef isSecondarySearchNecessary As Boolean) As String
    
    'Finding the text of the message is tricky, because it may be spliced between multiple quotations.  As an example, I frequently
    ' add manual line breaks to messages via " & vbCrLf & " - these need to be checked for and replaced.
    
    'The scan will work by looping through the string, and tracking whether or not we are currently inside quotation marks.
    'If we are outside a set of quotes, and we encounter a comma or closing parentheses, we know that we have reached the end of the
    ' first (and/or only) parameter.
    
    Dim initPosition As Long
    If inReverse Then
        initPosition = InStrRev(UCase$(srcLines(lineNumber)), ".ASSIGNTOOLTIP ")
    Else
        initPosition = InStr(1, UCase$(srcLines(lineNumber)), ".ASSIGNTOOLTIP ")
    End If
    
    Dim startQuote As Long
    startQuote = InStr(initPosition, srcLines(lineNumber), """")
    
    'Some tooltip assignments rely only on variables, not string text.  Ignore these, obviously, as their translation will be handled elsewhere.
    If (startQuote > 0) Then
        
        Dim endQuote As Long
        endQuote = -1
        
        Dim insideQuotes As Boolean
        insideQuotes = True
        
        Dim i As Long
        For i = startQuote + 1 To Len(srcLines(lineNumber))
            
            'Next, check for quotes
            If Mid$(srcLines(lineNumber), i, 1) = """" Then
                
                'Double-quotes are valid indicators, so manually check their appearance now
                If i < Len(srcLines(lineNumber)) - 1 Then
                
                    If Mid$(srcLines(lineNumber), i, 2) = """""" Then
                        
                        'Double quotes were found.  Increment i manually, and do not reset the insideQuotes marker
                        i = i + 1
                        
                    Else
                        insideQuotes = Not insideQuotes
                    End If
                
                Else
                    insideQuotes = Not insideQuotes
                End If
                
            End If
            
            If ((Mid$(srcLines(lineNumber), i, 1) = ",") Or (Mid$(srcLines(lineNumber), i, 1) = ")")) And (Not insideQuotes) Then
                endQuote = i - 1
                isSecondarySearchNecessary = True
                Exit For
            End If
            
            'See if we've reached the end of the line
            If (i = Len(srcLines(lineNumber))) And (Not insideQuotes) Then
                endQuote = i
                isSecondarySearchNecessary = False
                Exit For
            End If
                    
        Next i
        
        'If endQuote = -1, something went horribly wrong
        If endQuote = -1 Then
            Debug.Print "POTENTIAL MANUAL FIX REQUIRED FOR MESSAGE PARSE ERROR AT LINE # " & lineNumber & " IN " & m_FileName
            FindTooltipMessage = vbNullString
        Else
            FindTooltipMessage = Mid$(srcLines(lineNumber), startQuote + 1, endQuote - startQuote - 1)
        End If
        
        'We now need to replace line breaks in the text.  These can appear in a variety of ways.  Replace them all.
        Dim lineBreak As String
        lineBreak = """ & vbCrLf & """
        If InStr(1, FindTooltipMessage, lineBreak) Then FindTooltipMessage = Replace(FindTooltipMessage, lineBreak, vbCrLf)
        lineBreak = """ & vbCrLf & vbCrLf & """
        If InStr(1, FindTooltipMessage, lineBreak) Then FindTooltipMessage = Replace(FindTooltipMessage, lineBreak, vbCrLf & vbCrLf)
    
    Else
        FindTooltipMessage = vbNullString
    End If
    
End Function

'Given a line number and the original file contents, search for a message box title.
' (Note that this is cumbersome, as PD message boxes may have string-delimited ParamArray entries at the end of a PDMsgBox call,
'  entries which are dynamically inserted at run-time.  As such, we can't just look from the back of the string!)
Private Function FindMsgBoxTitle(ByRef srcLines() As String, ByRef lineNumber As Long) As String
    
    Const sQuot As String = """"
    Const sComma As String = ","
    
    Dim startQuote As Long
    startQuote = InStr(1, srcLines(lineNumber), sQuot, vbBinaryCompare)
    
    Dim startComma As Long
    startComma = InStr(1, srcLines(lineNumber), sComma, vbBinaryCompare)
    
    'If this string appears *after* the first comma, it's got to be the title - cool!
    ' If it appears *before* the first comma, however, we need to find the end of this initial message block,
    ' which is defined by a comma lying outside a quote block.
    If (startQuote > 0) And (startQuote < startComma) Then
    
        Dim insideQuotes As Boolean
        insideQuotes = True
        
        Dim i As Long
        For i = startQuote + 1 To Len(srcLines(lineNumber))
        
            If (Mid$(srcLines(lineNumber), i, 1) = sQuot) Then insideQuotes = Not insideQuotes
            
            If ((Mid$(srcLines(lineNumber), i, 1) = ",") Or (Mid$(srcLines(lineNumber), i, 1) = ")")) And (Not insideQuotes) Then
                startQuote = InStr(i, srcLines(lineNumber), sQuot, vbBinaryCompare)
                Exit For
            End If
        
        Next i
        
    End If
    
    Dim endQuote As Long
    If (startQuote > 0) Then endQuote = InStr(startQuote + 1, srcLines(lineNumber), sQuot, vbBinaryCompare)
    
    If (endQuote > 0) Then
        FindMsgBoxTitle = Mid$(srcLines(lineNumber), startQuote + 1, endQuote - startQuote - 1)
    Else
        FindMsgBoxTitle = vbNullString
    End If

End Function

'Given a line number and the original file contents, search for message box text
Private Function FindMsgBoxText(ByRef srcLines() As String, ByRef lineNumber As Long) As String

    'Before processing this message box, make sure that the text contains actual text and not just a reference to a string.
    ' If all it contains is a reference to a string variable, don't process it.
    If InStr(1, srcLines(lineNumber), "pdMsgBox(""", vbTextCompare) = 0 And InStr(1, srcLines(lineNumber), "pdMsgBox """, vbTextCompare) = 0 Then
        FindMsgBoxText = vbNullString
        Exit Function
    End If

    'Finding the text of the message is tricky, because it may be spliced between multiple quotations.  As an example, I frequently
    ' add manual line breaks to messages boxes via " & vbCrLf & " - these need to be checked for and replaced.
    
    'The scan will work by looping through the string, and tracking whether or not we are currently inside quotation marks.
    'If we are outside a set of quotes, and we encounter a comma, we know that we have reached the end of the first parameter.
    
    Dim startQuote As Long
    startQuote = InStr(1, srcLines(lineNumber), """")
    
    Dim endQuote As Long
    endQuote = -1
    
    Dim insideQuotes As Boolean
    insideQuotes = True
    
    Dim i As Long
    For i = startQuote + 1 To Len(srcLines(lineNumber))
    
        If Mid$(srcLines(lineNumber), i, 1) = """" Then insideQuotes = Not insideQuotes
        
        If (Mid$(srcLines(lineNumber), i, 1) = ",") And Not insideQuotes Then
            endQuote = i - 1
            Exit For
        End If
    
    Next i
    
    'If endQuote = -1, something went horribly wrong
    If endQuote = -1 Then
        FindMsgBoxText = "MANUAL FIX REQUIRED FOR MSGBOX PARSE ERROR AT LINE # " & lineNumber & " IN " & m_FileName
    Else
        FindMsgBoxText = Mid$(srcLines(lineNumber), startQuote + 1, endQuote - startQuote - 1)
    End If
    
    'We now need to replace line breaks in the text.  These can appear in a variety of ways.  Replace them all.
    Dim lineBreak As String
    lineBreak = """ & vbCrLf & """
    If InStr(1, FindMsgBoxText, lineBreak) Then FindMsgBoxText = Replace(FindMsgBoxText, lineBreak, vbCrLf)
    lineBreak = """ & vbCrLf & vbCrLf & """
    If InStr(1, FindMsgBoxText, lineBreak) Then FindMsgBoxText = Replace(FindMsgBoxText, lineBreak, vbCrLf & vbCrLf)

End Function

'Given a line number and the original file contents, search for arbitrary text between two quotation marks -
' but taking into account the complexities of extra mid-line quotes
Private Function FindCaptionInComplexQuotes(ByRef srcLines() As String, ByRef lineNumber As Long, Optional ByVal isTooltip As Boolean = False) As String

    '(This code is based off findMsgBoxText above - look there for more detailed comments)
    
    Dim startQuote As Long
    startQuote = InStr(1, srcLines(lineNumber), """")
    
    Dim endQuote As Long
    endQuote = -1
    
    Dim insideQuotes As Boolean
    insideQuotes = True
    
    Dim i As Long
    For i = startQuote + 1 To Len(srcLines(lineNumber))
    
        If Mid$(srcLines(lineNumber), i, 1) = """" Then insideQuotes = Not insideQuotes
        
        If ((Mid$(srcLines(lineNumber), i, 1) = ",") And Not insideQuotes) Then
            
            'Retreat backward until we find the last quotation mark, then report its location as the end of this text segment
            Dim j As Long
            For j = i To 1 Step -1
                If Mid$(srcLines(lineNumber), j, 1) = """" Then
                    endQuote = j
                    Exit For
                End If
            Next j
            
            Exit For
            
        End If
        
        If (i = Len(srcLines(lineNumber))) And Not insideQuotes Then
            endQuote = i
            Exit For
        End If
        
    Next i
    
    If isTooltip Then
        If InStr(1, srcLines(lineNumber), ".frx") > 0 Then
            FindCaptionInComplexQuotes = "MANUAL FIX REQUIRED FOR TOOLTIP (FRX REFERENCE) OF " & m_ObjectName & " IN " & m_FormName
            'MsgBox srcLines(lineNumber)
            Exit Function
        End If
    End If
    
    'If endQuote = -1, something went horribly wrong
    If endQuote = -1 Then
        If isTooltip Then
            FindCaptionInComplexQuotes = "MANUAL FIX REQUIRED FOR TOOLTIP ERROR FOR " & m_ObjectName & " IN " & m_FormName
            Debug.Print FindCaptionInComplexQuotes
            'MsgBox srcLines(lineNumber)
        Else
            FindCaptionInComplexQuotes = "MANUAL FIX REQUIRED FOR TEXT PARSE ERROR AT LINE # " & lineNumber & " IN " & m_FileName
            Debug.Print FindCaptionInComplexQuotes
        End If
    Else
        FindCaptionInComplexQuotes = Mid$(srcLines(lineNumber), startQuote + 1, endQuote - startQuote - 1)
    End If
    
    'We now need to replace line breaks in the text.  These can appear in a variety of ways.  Replace them all.
    Dim lineBreak As String
    lineBreak = """ & vbCrLf & """
    If InStr(1, FindCaptionInComplexQuotes, lineBreak) Then FindCaptionInComplexQuotes = Replace(FindCaptionInComplexQuotes, lineBreak, vbCrLf)
    lineBreak = """ & vbCrLf & vbCrLf & """
    If InStr(1, FindCaptionInComplexQuotes, lineBreak) Then FindCaptionInComplexQuotes = Replace(FindCaptionInComplexQuotes, lineBreak, vbCrLf & vbCrLf)

End Function

'Given a line number and the original file contents, search for arbitrary text between two quotation marks
Private Function FindCaptionInQuotes(ByRef srcLines() As String, ByRef lineNumber As Long, Optional ByVal startPosition As Long = 1) As String

    Dim startQuote As Long
    startQuote = InStr(startPosition, srcLines(lineNumber), """")
        
    Dim endQuote As Long
    endQuote = InStr(startQuote + 1, srcLines(lineNumber), """")
    
    If endQuote > 0 Then
        FindCaptionInQuotes = Mid$(srcLines(lineNumber), startQuote + 1, endQuote - startQuote - 1)
    Else
        FindCaptionInQuotes = vbNullString
    End If

End Function

'Given a line number and the original file contents, search for a "Caption" property.  Terminate if "End" is found.
Private Function FindControlCaption(ByRef srcLines() As String, ByRef lineNumber As Long) As String

    Dim foundCaption As Boolean
    foundCaption = True

    Dim originalLineNumber As Long
    originalLineNumber = lineNumber

    'Start by retrieving the name of this object and storing it in a module-level string.  The calling function may
    ' need this if the caption meets certain criteria.
    Dim objectName As String
    objectName = Trim$(srcLines(lineNumber))

    Dim sPos As Long
    sPos = Len(objectName)
    Do
        sPos = sPos - 1
    Loop While Mid$(objectName, sPos, 1) <> " "
    
    m_ObjectName = Right$(objectName, Len(objectName) - sPos)
    
    Do While InStr(1, UCase$(srcLines(lineNumber)), "CAPTION", vbBinaryCompare) = 0
        lineNumber = lineNumber + 1
        
        'Some controls may not have a caption.  If this occurs, exit.
        ' NOTE: we must use a binary comparison here, or entries with "End" in them will fail to work!
        If (InStr(1, srcLines(lineNumber), "End", vbBinaryCompare) > 0) And Not (InStr(1, srcLines(lineNumber), "EndProperty", vbBinaryCompare) > 0) Then
            foundCaption = False
            lineNumber = originalLineNumber
            Exit Do
        End If
        
    Loop
    
    'When we finally arrive here, line number has arrived at a line that contains the word "Caption"
    ' Grab whatever text is inside the quotation marks on that line
    If foundCaption Then
        
        'It's possible that the string is not actually a string, but a reference to a location in the relevant FRX file.
        ' I don't current have a way to retrieve this data, so do the next best thing - place a warning in the translation
        ' file.  I will later search for these and replace them with the relevant text.
        If InStr(1, srcLines(lineNumber), ".frx") Then
            FindControlCaption = "MANUAL FIX REQUIRED FOR " & m_ObjectName & " IN " & m_FormName
        Else
        
            Dim startQuote As Long
            startQuote = InStr(1, srcLines(lineNumber), """")
    
            Dim endQuote As Long
            endQuote = InStrRev(srcLines(lineNumber), """")
        
            If endQuote > 0 Then
                FindControlCaption = Mid$(srcLines(lineNumber), startQuote + 1, endQuote - startQuote - 1)
            Else
                FindControlCaption = vbNullString
            End If
            
        End If
        lineNumber = originalLineNumber + 1
                
    Else
        FindControlCaption = vbNullString
    End If

End Function

'Given a line number and the original file contents, search for a "Caption" property.  Terminate if "End" is found.
Private Function FindFormCaption(ByRef srcLines() As String, ByRef lineNumber As Long) As String

    Dim foundCaption As Boolean
    foundCaption = True

    'Start by retrieving the name of this form and storing it in a module-level string.  The calling function may
    ' need this if the caption meets certain criteria.
    Dim objectName As String
    objectName = Trim$(srcLines(lineNumber))

    Dim sPos As Long
    sPos = Len(objectName)
    Do
        sPos = sPos - 1
    Loop While Mid$(objectName, sPos, 1) <> " "
    
    m_FormName = Right$(objectName, Len(objectName) - sPos)
    
    Do While InStr(1, srcLines(lineNumber), "Caption") = 0
        lineNumber = lineNumber + 1
        
        'Some forms may not have a caption.  If this occurs, exit.
        If InStr(1, srcLines(lineNumber), "ClientHeight") Then
            foundCaption = False
            Exit Do
        End If
        
    Loop
    
    'When we finally arrive here, line number has arrived at a line that contains the word "Caption"
    ' Grab whatever text is inside the quotation marks on that line
    If foundCaption Then
        Dim startQuote As Long
        startQuote = InStr(1, srcLines(lineNumber), """")
        
        Dim endQuote As Long
        endQuote = InStrRev(srcLines(lineNumber), """")
        
        FindFormCaption = Mid$(srcLines(lineNumber), startQuote + 1, endQuote - startQuote - 1)
        
    Else
        FindFormCaption = vbNullString
    End If

End Function


'Extract a list of all project files from a VBP file
Private Sub cmdSelectVBP_Click()
    
    'This project should be located in a sub-path of a normal PhotoDemon install.
    ' We can use shlwapi's PathCanonicalize function to automatically "guess" at the location of PD's
    ' base en-US language file.
    Dim likelyDefaultLocation As String
    If Files.PathCanonicalize(Files.AppPathW() & "..\..", likelyDefaultLocation) Then likelyDefaultLocation = Files.PathAddBackslash(likelyDefaultLocation)
    m_VBPFile = likelyDefaultLocation & "PhotoDemon.vbp"
    
    lblVBP = "Active VBP: " & m_VBPFile
    m_VBPPath = Files.FileGetPath(m_VBPFile)
    
    'PD uses a hard-coded VBP location, but if you want to specify your own location, you can do so here
    'Dim cDialog As pdOpenSaveDialog
    'Set cDialog = New pdOpenSaveDialog
    'If cDialog.GetOpenFileName (m_VBPFile, , True, False, "VBP - Visual Basic Project|*.vbp", 1, , "Please select a Visual Basic project file (VBP)", "vbp", Me.hWnd) Then
    '    lblVBP = "Active VBP: " & m_VBPFile
    '    m_VBPPath = GetDirectory(m_VBPFile)
    'Else
    '    Exit Sub
    'End If
    
    'Load the file into a string array, split up line-by-line
    Dim vbpContents As String
    Files.FileLoadAsString m_VBPFile, vbpContents, True
    
    vbpText = Split(vbpContents, vbCrLf)
    ReDim m_vbpFiles(0 To UBound(vbpText)) As String
    Dim numOfFiles As Long
    numOfFiles = 0
    
    Dim majorVer As String, minorVer As String, buildVer As String
    
    'Extract only the relevant file paths
    Dim i As Long
    For i = 0 To UBound(vbpText)
    
        'Check for forms
        If InStr(1, vbpText(i), "Form=", vbBinaryCompare) = 1 Then
            m_vbpFiles(numOfFiles) = m_VBPPath & Right$(vbpText(i), Len(vbpText(i)) - 5)
            numOfFiles = numOfFiles + 1
        End If
        
        'Check for user controls
        If InStr(1, vbpText(i), "UserControl=", vbBinaryCompare) = 1 Then
            m_vbpFiles(numOfFiles) = m_VBPPath & Right$(vbpText(i), Len(vbpText(i)) - 12)
            numOfFiles = numOfFiles + 1
        End If
        
        'Check for modules
        If InStr(1, vbpText(i), "Module=", vbBinaryCompare) = 1 Then
            m_vbpFiles(numOfFiles) = m_VBPPath & Trim$(Right$(vbpText(i), Len(vbpText(i)) - InStr(1, vbpText(i), ";")))
            numOfFiles = numOfFiles + 1
        End If
        
        'Check for classes
        If InStr(1, vbpText(i), "Class=", vbBinaryCompare) = 1 Then
            m_vbpFiles(numOfFiles) = m_VBPPath & Trim$(Right$(vbpText(i), Len(vbpText(i)) - InStr(1, vbpText(i), ";")))
            numOfFiles = numOfFiles + 1
        End If
        
        'Check for version numbers
        If InStr(1, vbpText(i), "MajorVer=", vbBinaryCompare) = 1 Then
            majorVer = Trim$(Right$(vbpText(i), Len(vbpText(i)) - InStr(1, vbpText(i), "=")))
        End If
        If InStr(1, vbpText(i), "MinorVer=", vbBinaryCompare) = 1 Then
            minorVer = Trim$(Right$(vbpText(i), Len(vbpText(i)) - InStr(1, vbpText(i), "=")))
        End If
        If InStr(1, vbpText(i), "RevisionVer=", vbBinaryCompare) = 1 Then
            buildVer = Trim$(Right$(vbpText(i), Len(vbpText(i)) - InStr(1, vbpText(i), "=")))
        End If
    
    Next i
    
    ReDim Preserve m_vbpFiles(0 To numOfFiles) As String
    
    'To make sure everything worked, dump the contents of the array into the list box on the left
    lstProjectFiles.Clear
    
    For i = 0 To UBound(m_vbpFiles)
        If (LenB(m_vbpFiles(i)) <> 0) Then lstProjectFiles.AddItem m_vbpFiles(i)
    Next i
    
    'Build a complete version string
    m_VersionString = majorVer & "." & minorVer & "." & buildVer
    
    cmdProcess.Caption = "Begin processing"

End Sub

'Count the number of words in a string (will not be 100% accurate, but that's okay)
Private Function CountWordsInString(ByVal srcString As String) As Long
    
    CountWordsInString = 0
    
    srcString = Trim$(srcString)
    If (LenB(srcString) <> 0) Then
        
        Dim tmpArray() As String
        tmpArray = Split(srcString, " ")
        
        Dim i As Long
        For i = 0 To UBound(tmpArray)
            If IsAlpha(tmpArray(i)) Then CountWordsInString = CountWordsInString + 1
        Next i
        
    End If

End Function

'VB's IsNumeric function can't detect percentage text (e.g. "100%").  PhotoDemon includes text like this,
' but I don't want that text translated - so manually check for and reject it.
Private Function IsNumericPercentage(ByVal srcString As String) As Boolean
    
    IsNumericPercentage = False
    
    srcString = Trim$(srcString)

    'Start by checking for a percent in the right-most position
    Const PERCENT_SIGN As String = "%"
    If (Right$(srcString, 1) = PERCENT_SIGN) Then
        
        'If a percent was found, check the rest of the text to see if it's numeric
        IsNumericPercentage = IsNumeric(Left$(srcString, Len(srcString) - 1))
        
    End If

End Function

'URLs shouldn't be translated.  Check for them and reject as necessary.
Private Function IsURL(ByRef srcString As String) As Boolean
    Const FTP_PREFIX As String = "ftp"
    Const HTTP_PREFIX As String = "http"
    IsURL = (Left$(srcString, 6) = FTP_PREFIX) Or (Left$(srcString, 7) = HTTP_PREFIX)
End Function

Private Sub Message(ByRef msgText As String)
    If (Not m_SilentMode) Then
        lblUpdates.Caption = msgText
        lblUpdates.Refresh
        VBHacks.DoEvents_PaintOnly Me.hWnd, True
    End If
End Sub

Private Sub Form_Load()
    
    Set m_XML = New pdXML
    Set m_enUSPhrases = New pdStringHash
    
    'Build a blacklist of phrases that are in the software, but do not need to be translated.
    ' (These are complex phrases that may include things like proper nouns or mathematical terms,
    ' but the automatic text generator has no way of knowing that the text is non-translatable.)
    Set m_Blacklist = New pdStringHash
    
    AddBlacklist "*"
    AddBlacklist "("
    AddBlacklist ")"
    AddBlacklist ","
    AddBlacklist "(X, Y)"
    AddBlacklist "16:1 (1600%)"
    AddBlacklist "8:1 (800%)"
    AddBlacklist "4:1 (400%)"
    AddBlacklist "2:1 (200%)"
    AddBlacklist "1:2 (50%)"
    AddBlacklist "1:4 (25%)"
    AddBlacklist "1:8 (12.5%)"
    AddBlacklist "1:16 (6.25%)"
    AddBlacklist "GNU GPLv3"
    AddBlacklist "X.X"
    AddBlacklist "XX.XX.XX"
    AddBlacklist "tannerhelland.com/contact"
    AddBlacklist "photodemon.org/about/contact"
    AddBlacklist "photodemon.org/about/contact/"
    AddBlacklist "HTML / CSS"
    AddBlacklist "while it downloads."
    AddBlacklist "16x16"
    AddBlacklist "20x20"
    AddBlacklist "24x24"
    AddBlacklist "32x32"
    AddBlacklist "40x40"
    AddBlacklist "48x48"
    AddBlacklist "64x64"
    AddBlacklist "96x96"
    AddBlacklist "128x128"
    AddBlacklist "256x256"
    AddBlacklist "512x512"
    AddBlacklist "768x768"
    AddBlacklist "PackBits"
    AddBlacklist "LZW"
    AddBlacklist "ZIP"
    AddBlacklist "CCITT Fax 4"
    AddBlacklist "CCITT Fax 3"
    AddBlacklist "L*"
    AddBlacklist "a*"
    AddBlacklist "b*"
    AddBlacklist "Mitchell-Netravali"
    AddBlacklist "Catmull-Rom"
    AddBlacklist "Sinc (Lanczos)"
    AddBlacklist "Floyd-Steinberg"
    AddBlacklist "Stucki"
    AddBlacklist "Burkes"
    AddBlacklist "Sierra-3"
    AddBlacklist "DIB"
    AddBlacklist "DIB v5"
    AddBlacklist "PNG"
    AddBlacklist "IIR (Deriche)"
    AddBlacklist "Perlin"
    AddBlacklist "Simplex"
    AddBlacklist "OpenSimplex"
    AddBlacklist "Lab"
    AddBlacklist "PhotoDemon"
    AddBlacklist "Hilite"
    AddBlacklist "Laplacian"
    AddBlacklist "Prewitt"
    AddBlacklist "Sobel"
    AddBlacklist "Reinhard"
    AddBlacklist "Drago"
    AddBlacklist "WebP"
    AddBlacklist "fant"
    
    'Check the command line.  This project can be run in silent mode as part of my nightly build batch script.
    Dim chkCommandLine As String
    chkCommandLine = Command$
    
    If (LenB(Trim$(chkCommandLine)) <> 0) Then
        m_SilentMode = (InStr(1, chkCommandLine, "-s", vbTextCompare) <> 0)
    End If
    
    'If silent mode is activated, automatically "click" the relevant button
    If m_SilentMode Then
    
        'Load the current PhotoDemon VBP file
        Call cmdSelectVBP_Click
        
        'Generate a new master English file
        Call cmdProcess_Click
        
        'Forcibly merge all translation files with the latest English text
        Call cmdMergeAll_Click
        
        'If the program is running in silent mode, unload it now
        Unload Me
        
    End If
    
End Sub

Private Sub AddBlacklist(ByRef blString As String)
    m_Blacklist.AddItem LCase$(blString), vbNullString
End Sub

Private Function IsBlacklisted(ByRef blString As String) As Boolean
    IsBlacklisted = m_Blacklist.GetItemByKey(LCase$(blString), vbNullString)
End Function

'Used to estimate if a given string is an English word or not (where not means a number, standalone punctuation, etc).
' (PhotoDemon uses this to *VERY* roughly estimate word count in language files.)
Private Function IsAlpha(ByRef srcString As String) As Boolean
    
    IsAlpha = False
    
    If (Len(srcString) = 1) Then
        Const ALPHA_A As String = "A", ALPHA_I As String = "I"
        IsAlpha = (UCase$(srcString) = ALPHA_A) Or (UCase$(srcString) = ALPHA_I)
    Else
        
        Dim numAlphaChars As Long
        numAlphaChars = 0
        
        Dim i As Long
        For i = 1 To Len(srcString)
        
            Dim charID As Long
            charID = AscW(UCase$(Mid$(srcString, i, 1)))
            
            'Look for at least 2 alpha chars; that's good enough to assume this is a word
            If ((charID >= 65) And (charID <= 90)) Then numAlphaChars = numAlphaChars + 1
            If (numAlphaChars >= 2) Then
                IsAlpha = True
                Exit Function
            End If
            
        Next i
        
    End If
        
End Function

Private Sub ReplaceTopLevelTag(ByVal origTagName As String, ByRef sourceTextMaster As String, ByRef sourceTextTranslation As String, ByRef destinationText As String, Optional ByVal alsoIncrementVersion As Boolean = True)

    Dim openTagName As String, closeTagName As String
    openTagName = "<" & origTagName & ">"
    closeTagName = "</" & origTagName & ">"
    
    Dim findText As String, replaceText As String
    findText = openTagName & GetTextBetweenTags(sourceTextMaster, origTagName) & closeTagName
    
    'A special check is applied to the "langversion" tag.  Whenever this function is used, a merge is taking place;
    ' as such, we want to auto-increment the language's version number to trigger an update on client machines.
    If (Strings.StringsEqual(origTagName, "langversion", True) And alsoIncrementVersion) Then
        
        findText = openTagName & GetTextBetweenTags(sourceTextTranslation, origTagName) & closeTagName
        
        'Retrieve the current language version
        Dim curVersion As String
        curVersion = GetTextBetweenTags(sourceTextTranslation, origTagName)
        
        'Parse the current version into two discrete chunks: the major/minor value, and the revision value
        Dim curMajorMinor As String, curRevision As Long
        curMajorMinor = RetrieveVersionMajorMinorAsString(curVersion)
        curRevision = RetrieveVersionRevisionAsLong(curVersion)
        
        'Increment the revision value by 1, then assemble the modified replacement text
        curRevision = curRevision + 1
        replaceText = openTagName & curMajorMinor & "." & Trim$(Str$(curRevision)) & closeTagName
            
    Else
        replaceText = openTagName & GetTextBetweenTags(sourceTextTranslation, origTagName) & closeTagName
    End If
    
    destinationText = Replace$(destinationText, findText, replaceText)

End Sub
