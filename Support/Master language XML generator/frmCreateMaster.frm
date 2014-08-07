VERSION 5.00
Begin VB.Form frmCreateMaster 
   BackColor       =   &H80000005&
   Caption         =   " PhotoDemon Master Language XML Generator"
   ClientHeight    =   7665
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14535
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
   ScaleHeight     =   511
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   969
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMergeAll 
      Caption         =   "2a (Optional) Automatically merge all language files in default PhotoDemon folder with newest Master XML file..."
      Height          =   735
      Left            =   4800
      TabIndex        =   13
      Top             =   6600
      Width           =   9495
   End
   Begin VB.CheckBox chkRemoveDuplicates 
      BackColor       =   &H80000005&
      Caption         =   " Remove duplicate entries"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   3960
      Value           =   1  'Checked
      Width           =   3015
   End
   Begin VB.CommandButton cmdMerge 
      Caption         =   "3) Merge the files into an updated non-English XML file (NOTE: this will not modify the source files in any way)"
      Height          =   735
      Left            =   8520
      TabIndex        =   7
      Top             =   5640
      Width           =   5775
   End
   Begin VB.CommandButton cmdOldLanguage 
      Caption         =   "2) Select old non-English XML file..."
      Height          =   735
      Left            =   4800
      TabIndex        =   6
      Top             =   5640
      Width           =   3255
   End
   Begin VB.CommandButton cmdMaster 
      Caption         =   "1) Select master English XML file..."
      Height          =   735
      Left            =   1080
      TabIndex        =   5
      Top             =   5640
      Width           =   3255
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Begin processing"
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   3240
      Width           =   3015
   End
   Begin VB.ListBox lstProjectFiles 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   4200
      TabIndex        =   4
      Top             =   1920
      Width           =   10095
   End
   Begin VB.CommandButton cmdSelectVBP 
      Caption         =   "Select VBP file..."
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   1920
      Width           =   3015
   End
   Begin VB.Label lblWarning 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmCreateMaster.frx":0000
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
      Height          =   1335
      Left            =   240
      TabIndex        =   12
      Top             =   120
      Width           =   14055
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblUpdates 
      BackStyle       =   0  'Transparent
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
      Height          =   495
      Left            =   4080
      TabIndex        =   11
      Top             =   4560
      Width           =   10215
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "merge old translation files with new data:"
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
      Left            =   720
      TabIndex        =   10
      Top             =   5160
      Width           =   4410
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "extra language support tools"
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
      Left            =   240
      TabIndex        =   9
      Top             =   4560
      Width           =   3030
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      X1              =   8
      X2              =   960
      Y1              =   288
      Y2              =   288
   End
   Begin VB.Label lblExtract 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "step 2: process all files in project"
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
      Left            =   240
      TabIndex        =   8
      Top             =   2760
      Width           =   3495
   End
   Begin VB.Label lblVBP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "step 1: select VBP file"
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
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   2280
   End
End
Attribute VB_Name = "frmCreateMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Master English Language File (XML) Generator
'Copyright ©2012-2013 by Tanner Helland
'Created: 23/January/13
'Last updated: 27/September/13
'Last update: add detection code for new tooltip types; fix vbCrLf/vbLf issue the same way we fixed it in PD (necessary because
'             localization settings on a PC will sometimes use one in place of the other, so Frank's language files are always
'             the opposite of mine!)
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
'       I do not have any intention of supporting this tool outside its intended use, so please do not submit bug reports
'       regarding this project unless they directly relate to its intended purpose (generating a PhotoDemon XML language file).
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'Used to quickly check if a file (or folder) exists
Private Const ERROR_SHARING_VIOLATION As Long = 32
Private Declare Function GetFileAttributesW Lib "kernel32" (ByVal lpFileName As Long) As Long

'Variables used to generate the master translation file
Dim m_VBPFile As String, m_VBPPath As String
Dim m_FormName As String, m_ObjectName As String, m_FileName As String
Dim m_NumOfPhrasesFound As Long, m_NumOfPhrasesWritten As Long, m_numOfWords As Long
Dim vbpText() As String, vbpFiles() As String
Dim outputText As String, outputFile As String

'Variables used to merge old language files with new ones
Dim m_MasterText As String, m_OldLanguageText As String, m_NewLanguageText As String
Dim m_OldLanguagePath As String

'Variables used to build a blacklist of text that does not need to be translated
Dim m_Blacklist() As String
Dim m_numOfBlacklistEntries As Long

'String to store the version of the current VBP file (which will be written out to the master XML file for guidance)
Dim versionString As String

Private Sub cmdMaster_Click()

    Dim cDlg As cCommonDialog
    Set cDlg = New cCommonDialog
    
    Dim fPath As String
    fPath = "C:\PhotoDemon v4\PhotoDemon\App\PhotoDemon\Languages\Master\MASTER.xml"
    
    If cDlg.VBGetOpenFileName(fPath, , True, False, False, True, "XML - PhotoDemon Language File|*.xml", , , "Please select a PhotoDemon language file (XML)", "xml", Me.hWnd) Then
    
        'Load the file into a string
        m_MasterText = getFileAsString(fPath)
                
    End If
    
End Sub

Private Sub replaceTopLevelTag(ByVal origTagName As String, ByRef sourceTextMaster As String, ByRef sourceTextTranslation As String, ByRef destinationText As String)

    Dim openTagName As String, closeTagName As String
    openTagName = "<" & origTagName & ">"
    closeTagName = "</" & origTagName & ">"
    
    Dim findText As String, replaceText As String
    findText = openTagName & getTextBetweenTags(sourceTextMaster, origTagName) & closeTagName
    replaceText = openTagName & getTextBetweenTags(sourceTextTranslation, origTagName) & closeTagName
    destinationText = Replace(destinationText, findText, replaceText)

End Sub

Private Sub cmdMerge_Click()

    'Make sure our source file strings are not empty
    If m_MasterText = "" Or m_OldLanguageText = "" Then
        MsgBox "One or more source files are missing.  Supply those before attempting a merge."
        Exit Sub
    End If
    
    'Start by copying the contents of the master file into the destination string.  We will use that as our base, and update it
    ' with the old translations as best we can.
    m_NewLanguageText = m_MasterText
        
    Dim sPos As Long
    sPos = InStr(1, m_NewLanguageText, "<phrase>")
    
    Dim origText As String, translatedText As String
    Dim findText As String, replaceText As String
    
    'Copy over all top-level language and author information
    replaceTopLevelTag "langid", m_MasterText, m_OldLanguageText, m_NewLanguageText
    replaceTopLevelTag "langname", m_MasterText, m_OldLanguageText, m_NewLanguageText
    replaceTopLevelTag "langversion", m_MasterText, m_OldLanguageText, m_NewLanguageText
    replaceTopLevelTag "langstatus", m_MasterText, m_OldLanguageText, m_NewLanguageText
    replaceTopLevelTag "author", m_MasterText, m_OldLanguageText, m_NewLanguageText
        
    Dim phrasesProcessed As Long, phrasesFound As Long, phrasesMissed As Long
    phrasesProcessed = 0
    phrasesFound = 0
    phrasesMissed = 0
    
    'Start parsing the master text for <phrase> tags
    Do
    
        phrasesProcessed = phrasesProcessed + 1
    
        'Retrieve the original text associated with this phrase tag
        origText = getTextBetweenTags(m_MasterText, "original", sPos)
        
        'Attempt to retrieve a translation for this phrase using the old language file
        translatedText = getTranslationTagFromCaption(origText)
                
        'If no translation was found, and this string contains vbCrLf characters, replace them with plain vbLF characters and try again
        If Len(translatedText) = 0 Then
            If (InStr(1, origText, vbCrLf) > 0) Then
                translatedText = getTranslationTagFromCaption(Replace$(origText, vbCrLf, vbLf))
            End If
        End If
                
        'If a translation was found, insert it into the new file
        If Len(translatedText) > 0 Then
            findText = "<original>" & origText & "</original>" & vbCrLf & vbTab & vbTab & "<translation></translation>"
            replaceText = "<original>" & origText & "</original>" & vbCrLf & vbTab & vbTab & "<translation>" & translatedText & "</translation>"
            m_NewLanguageText = Replace(m_NewLanguageText, findText, replaceText)
            
            'As a failsafe, try the same thing without tabs
            findText = "<original>" & origText & "</original>" & vbCrLf & "<translation></translation>"
            replaceText = "<original>" & origText & "</original>" & vbCrLf & "<translation>" & translatedText & "</translation>"
            m_NewLanguageText = Replace(m_NewLanguageText, findText, replaceText)
            
            phrasesFound = phrasesFound + 1
        Else
            phrasesMissed = phrasesMissed + 1
        End If
    
        'Find the next occurrence of a <phrase> tag
        sPos = InStr(sPos + 1, m_MasterText, "<phrase>")
        
        lblUpdates.Caption = phrasesProcessed & " phrases processed.  (" & phrasesFound & " found, " & phrasesMissed & " missed)"
        lblUpdates.Refresh
        DoEvents
    
    Loop While sPos > 0
    
    'Prompt the user to save the results
    Dim cDlg As cCommonDialog
    Set cDlg = New cCommonDialog
    
    Dim fPath As String
    fPath = m_OldLanguagePath
    
    If cDlg.VBGetSaveFileName(fPath, , True, "XML - PhotoDemon Language File|*.xml", , , "Save the merged language file (XML)", "xml", Me.hWnd) Then
    
        If FileExist(fPath) Then
            MsgBox "File already exists!  Too dangerous to overwrite - please perform the merge again."
            Exit Sub
        End If
        
        Dim fileNum As Integer
        fileNum = FreeFile
        
        Open fPath For Output As #fileNum
            Print #fileNum, m_NewLanguageText
        Close #fileNum
        
    End If
    
    MsgBox "Merge complete." & vbCrLf & vbCrLf & phrasesProcessed & " phrases processed. " & phrasesFound & " translations found. " & phrasesMissed & " translations missing."

End Sub

'Given a string, return the location of the <phrase> tag enclosing said string
Private Function getPhraseTagLocation(ByRef srcString As String) As Long
    
    Dim sLocation As Long
    sLocation = InStr(1, m_OldLanguageText, srcString, vbTextCompare)
    
    'If the source string was found, work backward to find the phrase tag location
    If sLocation > 0 Then
        sLocation = InStrRev(m_OldLanguageText, "<phrase>", sLocation, vbTextCompare)
        If sLocation > 0 Then
            getPhraseTagLocation = sLocation
        Else
            getPhraseTagLocation = 0
        End If
    Else
        getPhraseTagLocation = 0
    End If

End Function

'Given the original caption of a message or control, return the matching translation from the active translation file
Private Function getTranslationTagFromCaption(ByVal origCaption As String) As String

    'Remove white space from the caption (if necessary, white space will be added back in after retrieving the translation from file)
    preProcessText origCaption
    origCaption = "<original>" & origCaption & "</original>"
    
    Dim phraseLocation As Long
    phraseLocation = getPhraseTagLocation(origCaption)
    
    'Make sure a phrase tag was found
    If phraseLocation > 0 Then
        
        'Retrieve the <translation> tag inside this phrase tag
        origCaption = getTextBetweenTags(m_OldLanguageText, "translation", phraseLocation)
        'postProcessText origCaption
        getTranslationTagFromCaption = origCaption
        
    Else
        getTranslationTagFromCaption = ""
    End If

End Function

'Given a file (as a String) and a tag (without brackets), return the text between that tag.
' NOTE: this function will always return the first occurence of the specified tag, starting at the specified search position.
' If the tag is not found, a blank string will be returned.
Private Function getTextBetweenTags(ByRef fileText As String, ByRef fTag As String, Optional ByVal searchLocation As Long = 1, Optional ByRef whereTagFound As Long = -1) As String

    Dim tagStart As Long, tagEnd As Long
    tagStart = InStr(searchLocation, fileText, "<" & fTag & ">", vbTextCompare)

    'If the tag was found in the file, we also need to find the closing tag.
    If tagStart > 0 Then
    
        tagEnd = InStr(tagStart, fileText, "</" & fTag & ">", vbTextCompare)
        
        'If the closing tag exists, return everything between that and the opening tag
        If tagEnd > tagStart Then
            
            'Increment the tag start location by the length of the tag plus two (+1 for each bracket: <>)
            tagStart = tagStart + Len(fTag) + 2
            
            'If the user passed a long, they want to know where this tag was found - return the location just after the
            ' location where the closing tag was located.
            If whereTagFound <> -1 Then whereTagFound = tagEnd + Len(fTag) + 2
            getTextBetweenTags = Mid$(fileText, tagStart, tagEnd - tagStart)
            
        Else
            getTextBetweenTags = "ERROR: specified tag wasn't properly closed!"
        End If
        
    Else
        getTextBetweenTags = ""
    End If

End Function

Private Sub preProcessText(ByRef srcString As String)

    '1) Trim the string
    srcString = Trim$(srcString)
    
    '2) Check for a trailing "..."
    If Right$(srcString, 3) = "..." Then srcString = Left$(srcString, Len(srcString) - 3)
    
    '3) Check for a trailing colon ":"
    If Right$(srcString, 1) = ":" Then srcString = Left$(srcString, Len(srcString) - 1)
    
End Sub

'New option added 09 September 2013 - Merge all language files automatically.  This will save me some trouble in the future.
Private Sub cmdMergeAll_Click()

    Dim srcFolder As String
    srcFolder = "C:\PhotoDemon v4\PhotoDemon\App\PhotoDemon\Languages\"
    
    'Auto-load the latest master language file
    m_MasterText = getFileAsString(srcFolder & "Master\MASTER.xml")
    
    'Iterate through every language file in the default PD directory
    'Scan the translation folder for .xml files.  Ignore anything that isn't XML.
    Dim chkFile As String
    chkFile = Dir(srcFolder & "*.xml", vbNormal)
        
    Do While (chkFile <> "")
        
        'Load the file into a string
        m_OldLanguageText = getFileAsString(srcFolder & chkFile)
        m_OldLanguagePath = srcFolder & chkFile
        
        'MsgBox m_OldLanguageText
        
        'BEGIN COPY OF CODE FROM cmdMerge
        
            'Make sure our source file strings are not empty
            If m_MasterText = "" Or m_OldLanguageText = "" Then
                Debug.Print "One or more source files are missing.  Supply those before attempting a merge."
                Exit Sub
            End If
            
            'Start by copying the contents of the master file into the destination string.  We will use that as our base, and update it
            ' with the old translations as best we can.
            m_NewLanguageText = m_MasterText
                
            Dim sPos As Long
            sPos = InStr(1, m_NewLanguageText, "<phrase>")
            
            Dim origText As String, translatedText As String
            Dim findText As String, replaceText As String
            
            'Copy over all top-level language and author information
            replaceTopLevelTag "langid", m_MasterText, m_OldLanguageText, m_NewLanguageText
            replaceTopLevelTag "langname", m_MasterText, m_OldLanguageText, m_NewLanguageText
            replaceTopLevelTag "langversion", m_MasterText, m_OldLanguageText, m_NewLanguageText
            replaceTopLevelTag "langstatus", m_MasterText, m_OldLanguageText, m_NewLanguageText
            replaceTopLevelTag "author", m_MasterText, m_OldLanguageText, m_NewLanguageText
                
            Dim phrasesProcessed As Long, phrasesFound As Long, phrasesMissed As Long
            phrasesProcessed = 0
            phrasesFound = 0
            phrasesMissed = 0
            
            'Start parsing the master text for <phrase> tags
            Do
            
                phrasesProcessed = phrasesProcessed + 1
            
                'Retrieve the original text associated with this phrase tag
                origText = getTextBetweenTags(m_MasterText, "original", sPos)
                
                'Attempt to retrieve a translation for this phrase using the old language file
                translatedText = getTranslationTagFromCaption(origText)
                
                'If no translation was found, and this string contains vbCrLf characters, replace them with plain vbLF characters and try again
                If translatedText = "" Then
                    If (InStr(1, origText, vbCrLf) > 0) Then
                        translatedText = getTranslationTagFromCaption(Replace$(origText, vbCrLf, vbLf))
                    End If
                End If
                
                'Remove any tab stops from the translated text (which may have been added by an outside editor)
                translatedText = Replace(translatedText, vbTab, "", , , vbBinaryCompare)
                                
                'If a translation was found, insert it into the new file
                If translatedText <> "" Then
                    'findText = "<original>" & origText & "</original>" & vbCrLf & vbTab & vbTab & vbTab & "<translation></translation>"
                    'replaceText = "<original>" & origText & "</original>" & vbCrLf & vbTab & vbTab & vbTab & "<translation>" & translatedText & "</translation>"
                    findText = "<original>" & origText & "</original>" & vbCrLf & "<translation></translation>"
                    replaceText = "<original>" & origText & "</original>" & vbCrLf & "<translation>" & translatedText & "</translation>"
                    m_NewLanguageText = Replace(m_NewLanguageText, findText, replaceText)
                    phrasesFound = phrasesFound + 1
                Else
                    phrasesMissed = phrasesMissed + 1
                End If
            
                'Find the next occurrence of a <phrase> tag
                sPos = InStr(sPos + 1, m_MasterText, "<phrase>")
                
                If (phrasesProcessed And 15) = 0 Then
                    lblUpdates.Caption = chkFile & ": " & phrasesProcessed & " phrases processed (" & phrasesFound & " found, " & phrasesMissed & " missed)"
                    lblUpdates.Refresh
                    DoEvents
                End If
            
            Loop While sPos > 0
            
            'Unlike the normal merge option, we will automatically save the results to a new XML file
            
            'Start by backing up the old file
            FileCopy m_OldLanguagePath, m_OldLanguagePath & ".backup"
                        
            If FileExist(m_OldLanguagePath) Then
                Kill m_OldLanguagePath
                Debug.Print "Note - old file with same name (" & m_OldLanguagePath & ") was erased.  Hope this is what you wanted!"
            End If
            
            Dim fileNum As Integer
            fileNum = FreeFile
            
            Open m_OldLanguagePath For Output As #fileNum
                Print #fileNum, m_NewLanguageText
            Close #fileNum
            
        
        'END COPY OF CODE FROM cmdMerge
        
        'Retrieve the next file and repeat
        chkFile = Dir
        
    Loop
    
    lblUpdates.Caption = "All language files processed successfully."
    lblUpdates.Refresh
    DoEvents

End Sub

Private Sub cmdOldLanguage_Click()
    
    Dim cDlg As cCommonDialog
    Set cDlg = New cCommonDialog
    
    Dim fPath As String
    fPath = "C:\PhotoDemon v4\PhotoDemon\App\PhotoDemon\Languages\"
    
    Dim tmpLangFile As String
    
    If cDlg.VBGetOpenFileName(tmpLangFile, , True, False, False, True, "XML - PhotoDemon Language File|*.xml", , fPath, "Please select a PhotoDemon language file (XML)", "xml", Me.hWnd) Then
    
        'Load the file into a string
        m_OldLanguageText = getFileAsString(tmpLangFile)
        m_OldLanguagePath = tmpLangFile
                
    End If
    
End Sub

'Process all files in a project file.  (NOTE: a VBP file must first be selected before running this step.)
Private Sub cmdProcess_Click()

    If Len(m_VBPFile) = 0 Then
        MsgBox "Select a VBP file first.", vbExclamation + vbApplicationModal + vbOKOnly, "Oops"
        Exit Sub
    End If
    
    'Start by preparing the XML header
    outputText = "<?xml version=""1.0"" encoding=""windows-1252""?>"
    outputText = outputText & vbCrLf & vbCrLf
    outputText = outputText & vbTab & "<pdData>"
    outputText = outputText & vbCrLf & vbCrLf
    outputText = outputText & vbTab & vbTab & "<pdDataType>Translation</pdDataType>"
    outputText = outputText & vbCrLf & vbCrLf
    outputText = outputText & vbTab & vbTab & "<langid>en-US</langid>" & vbCrLf
    outputText = outputText & vbTab & vbTab & "<langname>English (US) - MASTER COPY</langname>" & vbCrLf
    outputText = outputText & vbTab & vbTab & "<langversion>" & versionString & "</langversion>" & vbCrLf
    outputText = outputText & vbTab & vbTab & "<langstatus>Autogenerated - manual inspection still required</langstatus>" & vbCrLf
    outputText = outputText & vbCrLf & vbTab & vbTab & "<author>VBP Text Extraction App (by Tanner Helland)</author>"
    outputText = outputText & vbCrLf & vbCrLf
    outputText = outputText & vbTab & vbTab & "<!-- BEGIN AUTOMATIC TEXT GENERATION -->"
    
    Dim numOfFiles As Long
    numOfFiles = UBound(vbpFiles)
    
    m_NumOfPhrasesFound = 0
    m_NumOfPhrasesWritten = 0
    m_numOfWords = 0
    
    Dim i As Long
    For i = 0 To numOfFiles
        cmdProcess.Caption = "Processing project file " & i + 1 & " of " & numOfFiles + 1
        processFile vbpFiles(i)
    Next i
    
    'With processing complete, write out our final stats (just for fun)
    outputText = outputText & vbCrLf & vbCrLf
    outputText = outputText & vbTab & vbTab & "<!-- Automatic text extraction complete. -->" & vbCrLf & vbCrLf
    
    'Updated 09 September 2013: write out phrase count as an actual tag, which PD's new language editor can use to approximate a max
    ' value for its progress bar when loading the language file.
    outputText = outputText & vbTab & vbTab & "<phrasecount>" & m_NumOfPhrasesWritten & "</phrasecount>"
    outputText = outputText & vbCrLf & vbCrLf
    
    'Proceed with human-readable phrase statistics
    If CBool(chkRemoveDuplicates) Then
        outputText = outputText & vbTab & "<!-- As of this build, PhotoDemon contains " & m_NumOfPhrasesFound & " phrases. -->"
        outputText = outputText & vbCrLf
        outputText = outputText & vbTab & "<!-- " & CStr(m_NumOfPhrasesFound - m_NumOfPhrasesWritten) & " are duplicates, so only " & m_NumOfPhrasesWritten & " unique phrases have been written to file. -->"
        outputText = outputText & vbCrLf
        outputText = outputText & vbTab & "<!-- These " & m_NumOfPhrasesWritten & " phrases contain approximately " & m_numOfWords & " total words. -->"
    Else
        outputText = outputText & vbTab & "<!-- As of this build, PhotoDemon contains " & m_NumOfPhrasesWritten & " phrases (including duplicates). -->"
        outputText = outputText & vbCrLf
        outputText = outputText & vbTab & "<!-- These " & m_NumOfPhrasesWritten & " phrases contain approximately " & m_numOfWords & " total words. -->"
    End If
    
    
    'Terminate the final language tag
    outputText = outputText & vbCrLf & vbCrLf & vbTab & "</pdData>"
    
    'Write the text out to file
    If CBool(chkRemoveDuplicates) Then
        outputFile = m_VBPPath & "App\PhotoDemon\Languages\Master\MASTER.xml"
    Else
        outputFile = m_VBPPath & "App\PhotoDemon\Languages\Master\MASTER (with duplicates).xml"
    End If
    
    If FileExist(outputFile) Then Kill outputFile
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open outputFile For Output As #fileNum
        Print #fileNum, outputText
    Close #fileNum
    
    cmdProcess.Caption = "Processing complete!"
    'MsgBox "Text extraction complete!"
    
End Sub

'Given a VB file (form, module, class, user control), extract any relevant text from it
Private Sub processFile(ByVal srcFile As String)

    If srcFile = "" Then Exit Sub

    m_FileName = getFilename(srcFile)
    
    'Certain files can be ignored.  I generate this list manually, on account of knowing which files (classes, mostly) contain
    ' no special text.
    If (m_FileName = "clsToolTip.cls") Or (m_FileName = "clsControlImages.cls") Or (m_FileName = "pdFilterSupport.cls") Or (m_FileName = "cSelfSubHookCallback.cls") Or (m_FileName = "jcButton.ctl") Or (m_FileName = "VBP_PublicVariables.bas") Or (m_FileName = "pdParamString.cls") Then Exit Sub
    
    'Start by copying all text from the file into a line-by-line array
    Dim fileContents As String
    fileContents = getFileAsString(srcFile)
    Dim fileLines() As String
    fileLines = Split(fileContents, vbCrLf)
    
    'If this file is a form file, the second line of the file will contain the text: "Begin VB.FORM FormName", where FormName
    ' is the name of the form. By inserting the form's name into our translation file, the translation engine can use it to quickly
    ' locate all translations on that form.
    Dim shortcutName As String
    shortcutName = ""
    
    If Right(m_FileName, 3) = "frm" Then
        Dim findName() As String
        findName = Split(fileLines(1), " ")
        shortcutName = findName(2)
    End If
    
    'For convenience, write the name of the source file into the translation file - this can be helpful when
    ' tracking down errors or incomplete text.
    If LenB(m_FileName) > 0 Then
        outputText = outputText & vbCrLf & vbCrLf & vbTab & vbTab
        If Len(shortcutName) > 0 Then
            outputText = outputText & "<!-- BEGIN text for " & m_FileName & " (" & shortcutName & ") -->"
        Else
            outputText = outputText & "<!-- BEGIN text for " & m_FileName & " -->"
        End If
    End If
    
    Dim curLineNumber As Long
    curLineNumber = 0
    
    Dim numOfPhrasesFound As Long, numOfPhrasesWritten As Long
    numOfPhrasesFound = 0
    numOfPhrasesWritten = 0
    
    Dim curLineText As String, processedText As String, processedTextSecondary As String, chkText As String
    m_FormName = ""
    
    'Now, start processing the file one line at a time, searching for relevant text as we go
    Do
    
        processedText = ""
        processedTextSecondary = ""
    
        curLineText = fileLines(curLineNumber)
        
        'Before processing this line, make sure is isn't a comment.  (Comments are always ignored.)
        If Left(Trim(curLineText), 1) = "'" Then GoTo nextLine
        
        'There are many ways that translatable text may appear in a VB source file.
        ' 1) As a form caption
        ' 2) As a control caption
        ' 3) As tooltip text
        ' 4) As text added to a combo box or list box control at run-time (e.g. "control.AddItem "xyz")
        ' 5) As a message call (e.g. Message "xyz")
        ' 6) As message box text, specifically pdMsgBox (e.g. one of either pdMsgBox("xyz"...) or pdMsgBox "xyz"...)
        ' 7) As a message box title caption (more convoluted to find - basically the 3rd parameter of a pdMsgBox call)
        ' 8) As miscellaneous text manually marked for translation (e.g. g_Language.translateMessage("xyz"))
        ' 9) As miscellaneous tooltip text manually marked for translation by the assignTooltip function.
        ' 10) Process calls, which are relayed to the user in the Undo / Redo menus (e.g. "Undo Blur")
        ' (in some rare cases, text may appear that doesn't fit any of these cases - such text must be added manually)
        
        'Every one of these requires a unique mechanism for checking the text.
        
        'Note that some of these mechanisms will modify the current line number.  These require the line number, passed
        ' ByRef, for that purpose.
        
        'If any of the functions are successful, they will return the string that needs to be added to the XML file
        
        '1) Check for a form caption
        If InStr(1, curLineText, "Begin VB.Form", vbTextCompare) Then
            processedText = findFormCaption(fileLines, curLineNumber)
                
        '2) Check for a control caption.  (This has to be handled slightly differently than form caption.)
        ElseIf ((InStr(1, curLineText, "Begin VB.", vbTextCompare) > 0) Or (InStr(1, curLineText, "Begin PhotoDemon.", vbTextCompare) > 0)) And (InStr(1, curLineText, "PictureBox", vbTextCompare) = 0) And (InStr(1, curLineText, "ComboBox") = 0) And (InStr(1, curLineText, ".Shape") = 0) And (InStr(1, curLineText, "TextBox") = 0) And (InStr(1, curLineText, "HScrollBar") = 0) And (InStr(1, curLineText, "VScrollBar") = 0) Then
            processedText = findControlCaption(fileLines, curLineNumber)
        
        '3) Check for tooltip text (several varations of this exist due to custom controls having unique tooltip property names)
        ElseIf InStr(1, curLineText, "ToolTipText", vbTextCompare) And (InStr(1, curLineText, ".ToolTipText", vbTextCompare) = 0) Then
            processedText = findCaptionInComplexQuotes(fileLines, curLineNumber, True)
                        
        ElseIf (InStr(1, curLineText, "ToolTip", vbTextCompare) > 0) And (InStr(1, curLineText, ".ToolTip", vbTextCompare) = 0) And (InStr(1, curLineText, "TooltipTitle", vbTextCompare) = 0) And (InStr(1, curLineText, "ToolTipText", vbTextCompare) = 0) And (InStr(1, curLineText, "TooltipBackColor", vbTextCompare) = 0) And (InStr(1, curLineText, "ToolTipType", vbTextCompare) = 0) And (InStr(1, curLineText, "m_ToolTip", vbTextCompare) = 0) And (InStr(1, curLineText, "clsToolTip", vbTextCompare) = 0) And (Not m_FileName = "jcButton.ctl") And (InStr(1, curLineText, "=") > 0) And (InStr(1, curLineText, "PD_MAX_TOOLTIP_WIDTH") = 0) And (InStr(1, curLineText, "delaytime", vbTextCompare) = 0) And (InStr(1, curLineText, "eControl.ToolTipText", vbTextCompare) = 0) And (InStr(1, curLineText, "tooltipBackup", vbTextCompare) = 0) And (InStr(1, curLineText, "newTooltip", vbTextCompare) = 0) And (InStr(1, curLineText, "SetThumbnailTooltip", vbTextCompare) = 0) Then
            processedText = findCaptionInComplexQuotes(fileLines, curLineNumber, True)
        
        ElseIf InStr(1, curLineText, "TooltipTitle", vbTextCompare) And (InStr(1, curLineText, ".TooltipTitle") = 0) And (Not m_FileName = "jcButton.ctl") Then
            processedText = findCaptionInComplexQuotes(fileLines, curLineNumber, True)
                        
        '4) Check for text added to a combo box or list box control at run-time
        ElseIf InStr(1, curLineText, ".AddItem """) Then
            processedText = findCaptionInComplexQuotes(fileLines, curLineNumber)
            
        '5) Check for message calls
        ElseIf InStr(1, curLineText, "Message """) Then
            processedText = findCaptionInQuotes(fileLines, curLineNumber)
        
        '6) Check for message box text, including 7) message box titles (which must also be translated)
        ElseIf InStr(1, curLineText, "pdMsgBox") Then
        
            'First, retrieve the message box text itself
            processedText = findMsgBoxText(fileLines, curLineNumber)
            
            'Next, retrieve the message box title
            processedTextSecondary = findMsgBoxTitle(fileLines, curLineNumber)
                        
        '8) Check for text that has been manually marked for translation (e.g. g_Language.TranslateMessage("xyz"))
        '    NOTE: as of 07 June 2013, each line can contain two translation calls (instead of just one)
        ElseIf InStr(1, curLineText, "g_Language.TranslateMessage(""") Then
            processedText = findMessage(fileLines, curLineNumber)
            processedTextSecondary = findMessage(fileLines, curLineNumber, True)
        
        '9) Check for tooltip text that has been manually assigned to a custom PhotoDemon object
        ElseIf InStr(1, curLineText, ".assignTooltip """) Then
            processedText = findTooltipMessage(fileLines, curLineNumber)
            
        '10) And finally, specific to PhotoDemon - check for action names that may not be present elsewhere
        'ElseIf InStr(1, curLineText, "GetNameOfProcess =") Then
        ElseIf InStr(1, curLineText, "Process """) Then
            processedText = findCaptionInQuotes(fileLines, curLineNumber, InStr(1, curLineText, "Process """))
            
        End If
        
        'DEBUG! Check for certain text entries here
        'If (shortcutName = "FormLens") And Len(Trim$(processedText)) > 0 Then MsgBox processedText
        
        'We now have text in potentially two places: processedText, and processedTextSecondary (for message box titles)
        chkText = Trim$(processedText)
        
        'Only pass the text along if it isn't blank, or a number, or a symbol, or a manually blacklisted phrase
        If (chkText <> "") And (Not IsNumeric(chkText)) And (Not IsNumericPercentage(chkText)) And (Not isBlacklisted(chkText)) Then
            If (chkText <> ".") And (chkText <> "-") And (Not IsURL(chkText)) Then
                numOfPhrasesFound = numOfPhrasesFound + 1
                If addPhrase(processedText) Then numOfPhrasesWritten = numOfPhrasesWritten + 1
            End If
        End If
        
        chkText = Trim$(processedTextSecondary)
        
        'Do the same for the secondary text
        If (chkText <> "") And (Not IsNumeric(chkText)) And (Not IsNumericPercentage(chkText)) And (Not isBlacklisted(chkText)) Then
            If (chkText <> ".") And (chkText <> "-") And (Not IsURL(chkText)) Then
                numOfPhrasesFound = numOfPhrasesFound + 1
                If addPhrase(processedTextSecondary) Then numOfPhrasesWritten = numOfPhrasesWritten + 1
            End If
        End If
    
nextLine:
        curLineNumber = curLineNumber + 1
    
    Loop While curLineNumber < UBound(fileLines)
    
    'Now that all phrases in this file have been processed, we can wrap up this section of XML
    
    'For fun, write some stats about our processing results into the translation file.
    If m_FileName <> "" Then
        
        outputText = outputText & vbCrLf & vbCrLf & vbTab & vbTab
        If numOfPhrasesFound <> 1 Then
            outputText = outputText & "<!-- " & m_FileName & " contains " & numOfPhrasesFound & " phrases. "
        Else
            outputText = outputText & "<!-- " & m_FileName & " contains " & numOfPhrasesFound & " phrase. "
        End If
        If numOfPhrasesFound > 0 Then
            If numOfPhrasesWritten <> numOfPhrasesFound Then
                
                Dim phraseDifference As Long
                phraseDifference = numOfPhrasesFound - numOfPhrasesWritten
                
                Select Case phraseDifference
                    Case 1
                        outputText = outputText & " One was a duplicate of an existing phrase, so only " & numOfPhrasesWritten & " new phrases were written to file. -->"
                    Case numOfPhrasesFound
                        outputText = outputText & " All were duplicates of existing phrases, so no new phrases were written to file. -->"
                    Case Else
                        If numOfPhrasesWritten = 1 Then
                            outputText = outputText & CStr(phraseDifference) & " were duplicates of existing phrases, so only one new phrase was written to file. -->"
                        Else
                            outputText = outputText & CStr(phraseDifference) & " were duplicates of existing phrases, so only " & numOfPhrasesWritten & " new phrases were written to file. -->"
                        End If
                End Select
                
            Else
                Select Case numOfPhrasesFound
                    Case 1
                        outputText = outputText & " The phrase was unique, so 1 new phrase was written to file. -->"
                    Case 2
                        outputText = outputText & " Both phrases were unique, so " & numOfPhrasesFound & " new phrases were written to file. -->"
                    Case Else
                        outputText = outputText & " All " & numOfPhrasesFound & " were unique, so " & numOfPhrasesFound & " new phrases were written to file. -->"
                End Select
            End If
        Else
            outputText = outputText & "-->"
        End If
    End If
    
    'For convenience, once again write the name of the source file into the translation file - this can be helpful when
    ' tracking down errors or incomplete text.
    If m_FileName <> "" Then
        outputText = outputText & vbCrLf & vbCrLf & vbTab & vbTab
        outputText = outputText & "<!-- END text for " & m_FileName & "-->"
    End If
    
    'Add the number of phrases found and written to the master count
    m_NumOfPhrasesFound = m_NumOfPhrasesFound + numOfPhrasesFound
    m_NumOfPhrasesWritten = m_NumOfPhrasesWritten + numOfPhrasesWritten

End Sub

'Add a discovered phrase to the XML file.  If this phrase already exists in the file, ignore it.
Private Function addPhrase(ByRef phraseText As String) As Boolean
                        
    'Replace double double-quotes (which are required in code) with just one set of double-quotes
    If InStr(1, phraseText, """""") Then phraseText = Replace(phraseText, """""", """")
            
    'Next, do the same pre-processing that we do in the translation engine
    
    '1) Trim the text.  Extra spaces will be handled by the translation engine.
    phraseText = Trim(phraseText)
    
    '2) Check for a trailing "..." and remove it
    If Right$(phraseText, 3) = "..." Then phraseText = Left$(phraseText, Len(phraseText) - 3)
    
    '3) Check for a trailing colon ":" and remove it
    If Right$(phraseText, 1) = ":" Then phraseText = Left$(phraseText, Len(phraseText) - 1)
    
    'This phrase is now ready to write to file.
    
    'Before writing the phrase out, check to see if it already exists
    If CBool(chkRemoveDuplicates) Then
                
        If InStr(1, outputText, "<original>" & phraseText & "</original>", vbTextCompare) > 0 Then
            addPhrase = False
        Else
            If phraseText <> "" Then
                addPhrase = True
            Else
                addPhrase = False
            End If
        End If
    Else
        If phraseText <> "" Then
            addPhrase = True
        Else
            addPhrase = False
        End If
    End If
    
    'If the phrase does not exist, add it now
    If addPhrase Then
        outputText = outputText & vbCrLf & vbCrLf
        outputText = outputText & vbTab & vbTab & "<phrase>"
        outputText = outputText & vbCrLf & vbTab & vbTab & vbTab & "<original>"
        outputText = outputText & phraseText
        outputText = outputText & "</original>"
        outputText = outputText & vbCrLf & vbTab & vbTab & vbTab & "<translation></translation>"
        outputText = outputText & vbCrLf & vbTab & vbTab & "</phrase>"
        m_numOfWords = m_numOfWords + countWordsInString(phraseText)
    End If
    
End Function

'Given a line number and the original file contents, search for a custom PhotoDemon translation request
Private Function findMessage(ByRef srcLines() As String, ByRef lineNumber As Long, Optional ByVal inReverse As Boolean = False) As String
    
    'Finding the text of the message is tricky, because it may be spliced between multiple quotations.  As an example, I frequently
    ' add manual line breaks to messages via " & vbCrLf & " - these need to be checked for and replaced.
    
    'The scan will work by looping through the string, and tracking whether or not we are currently inside quotation marks.
    'If we are outside a set of quotes, and we encounter a comma or closing parentheses, we know that we have reached the end of the
    ' first (and/or only) parameter.
    
    Dim initPosition As Long
    If inReverse Then
        initPosition = InStrRev(srcLines(lineNumber), "g_Language.TranslateMessage(""")
    Else
        initPosition = InStr(1, srcLines(lineNumber), "g_Language.TranslateMessage(""")
    End If
    
    Dim startQuote As Long
    startQuote = InStr(initPosition, srcLines(lineNumber), """")
    
    Dim endQuote As Long
    endQuote = -1
    
    Dim insideQuotes As Boolean
    insideQuotes = True
    
    Dim i As Long
    For i = startQuote + 1 To Len(srcLines(lineNumber))
    
        If Mid(srcLines(lineNumber), i, 1) = """" Then insideQuotes = Not insideQuotes
        
        If ((Mid(srcLines(lineNumber), i, 1) = ",") Or (Mid(srcLines(lineNumber), i, 1) = ")")) And (Not insideQuotes) Then
            endQuote = i - 1
            Exit For
        End If
    
    Next i
    
    'If endQuote = -1, something went horribly wrong
    If endQuote = -1 Then
        findMessage = "MANUAL FIX REQUIRED FOR MESSAGE PARSE ERROR AT LINE # " & lineNumber & " IN " & m_FileName
    Else
        findMessage = Mid(srcLines(lineNumber), startQuote + 1, endQuote - startQuote - 1)
    End If
    
    'We now need to replace line breaks in the text.  These can appear in a variety of ways.  Replace them all.
    Dim lineBreak As String
    lineBreak = """ & vbCrLf & """
    If InStr(1, findMessage, lineBreak) Then findMessage = Replace(findMessage, lineBreak, vbCrLf)
    lineBreak = """ & vbCrLf & vbCrLf & """
    If InStr(1, findMessage, lineBreak) Then findMessage = Replace(findMessage, lineBreak, vbCrLf & vbCrLf)

    
End Function

'Given a line number and the original file contents, search for a custom PhotoDemon tooltip assignment
Private Function findTooltipMessage(ByRef srcLines() As String, ByRef lineNumber As Long, Optional ByVal inReverse As Boolean = False) As String
    
    'Finding the text of the message is tricky, because it may be spliced between multiple quotations.  As an example, I frequently
    ' add manual line breaks to messages via " & vbCrLf & " - these need to be checked for and replaced.
    
    'The scan will work by looping through the string, and tracking whether or not we are currently inside quotation marks.
    'If we are outside a set of quotes, and we encounter a comma or closing parentheses, we know that we have reached the end of the
    ' first (and/or only) parameter.
    
    Dim initPosition As Long
    If inReverse Then
        initPosition = InStrRev(srcLines(lineNumber), ".assignTooltip """)
    Else
        initPosition = InStr(1, srcLines(lineNumber), ".assignTooltip """)
    End If
    
    Dim startQuote As Long
    startQuote = InStr(initPosition, srcLines(lineNumber), """")
    
    Dim endQuote As Long
    endQuote = -1
    
    Dim insideQuotes As Boolean
    insideQuotes = True
    
    Dim i As Long
    For i = startQuote + 1 To Len(srcLines(lineNumber))
    
        If Mid(srcLines(lineNumber), i, 1) = """" Then insideQuotes = Not insideQuotes
        
        If ((Mid(srcLines(lineNumber), i, 1) = ",") Or (Mid(srcLines(lineNumber), i, 1) = ")")) And (Not insideQuotes) Then
            endQuote = i - 1
            Exit For
        End If
        
        If (i = Len(srcLines(lineNumber))) And (Not insideQuotes) Then
            endQuote = i
            Exit For
        End If
    
    Next i
    
    'If endQuote = -1, something went horribly wrong
    If endQuote = -1 Then
        findTooltipMessage = "MANUAL FIX REQUIRED FOR MESSAGE PARSE ERROR AT LINE # " & lineNumber & " IN " & m_FileName
    Else
        findTooltipMessage = Mid(srcLines(lineNumber), startQuote + 1, endQuote - startQuote - 1)
    End If
    
    'We now need to replace line breaks in the text.  These can appear in a variety of ways.  Replace them all.
    Dim lineBreak As String
    lineBreak = """ & vbCrLf & """
    If InStr(1, findTooltipMessage, lineBreak) Then findTooltipMessage = Replace(findTooltipMessage, lineBreak, vbCrLf)
    lineBreak = """ & vbCrLf & vbCrLf & """
    If InStr(1, findTooltipMessage, lineBreak) Then findTooltipMessage = Replace(findTooltipMessage, lineBreak, vbCrLf & vbCrLf)

    
End Function

'Given a line number and the original file contents, search for a message box title
Private Function findMsgBoxTitle(ByRef srcLines() As String, ByRef lineNumber As Long) As String

    Dim endQuote As Long
    endQuote = InStrRev(srcLines(lineNumber), """", Len(srcLines(lineNumber)))
        
    Dim startQuote As Long
    startQuote = InStrRev(srcLines(lineNumber), """", endQuote - 1)
    
    If endQuote > 0 Then
        findMsgBoxTitle = Mid(srcLines(lineNumber), startQuote + 1, endQuote - startQuote - 1)
    Else
        findMsgBoxTitle = ""
    End If

End Function

'Given a line number and the original file contents, search for message box text
Private Function findMsgBoxText(ByRef srcLines() As String, ByRef lineNumber As Long) As String

    'Before processing this message box, make sure that the text contains actual text and not just a reference to a string.
    ' If all it contains is a reference to a string variable, don't process it.
    If InStr(1, srcLines(lineNumber), "pdMsgBox(""") = 0 And InStr(1, srcLines(lineNumber), "pdMsgBox """) = 0 Then
        findMsgBoxText = ""
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
    
        If Mid(srcLines(lineNumber), i, 1) = """" Then insideQuotes = Not insideQuotes
        
        If (Mid(srcLines(lineNumber), i, 1) = ",") And Not insideQuotes Then
            endQuote = i - 1
            Exit For
        End If
    
    Next i
    
    'If endQuote = -1, something went horribly wrong
    If endQuote = -1 Then
        findMsgBoxText = "MANUAL FIX REQUIRED FOR MSGBOX PARSE ERROR AT LINE # " & lineNumber & " IN " & m_FileName
    Else
        findMsgBoxText = Mid(srcLines(lineNumber), startQuote + 1, endQuote - startQuote - 1)
    End If
    
    'We now need to replace line breaks in the text.  These can appear in a variety of ways.  Replace them all.
    Dim lineBreak As String
    lineBreak = """ & vbCrLf & """
    If InStr(1, findMsgBoxText, lineBreak) Then findMsgBoxText = Replace(findMsgBoxText, lineBreak, vbCrLf)
    lineBreak = """ & vbCrLf & vbCrLf & """
    If InStr(1, findMsgBoxText, lineBreak) Then findMsgBoxText = Replace(findMsgBoxText, lineBreak, vbCrLf & vbCrLf)

End Function

'Given a line number and the original file contents, search for arbitrary text between two quotation marks -
' but taking into account the complexities of extra mid-line quotes
Private Function findCaptionInComplexQuotes(ByRef srcLines() As String, ByRef lineNumber As Long, Optional ByVal isTooltip As Boolean = False) As String

    '(This code is based off findMsgBoxText above - look there for more detailed comments)
    
    Dim startQuote As Long
    startQuote = InStr(1, srcLines(lineNumber), """")
    
    Dim endQuote As Long
    endQuote = -1
    
    Dim insideQuotes As Boolean
    insideQuotes = True
    
    Dim i As Long
    For i = startQuote + 1 To Len(srcLines(lineNumber))
    
        If Mid(srcLines(lineNumber), i, 1) = """" Then insideQuotes = Not insideQuotes
        
        If ((Mid(srcLines(lineNumber), i, 1) = ",") And Not insideQuotes) Then
            
            'Retreat backward until we find the last quotation mark, then report its location as the end of this text segment
            Dim j As Long
            For j = i To 1 Step -1
                If Mid(srcLines(lineNumber), j, 1) = """" Then
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
            findCaptionInComplexQuotes = "MANUAL FIX REQUIRED FOR TOOLTIP OF " & m_ObjectName & " IN " & m_FormName
            'MsgBox srcLines(lineNumber)
            Exit Function
        End If
    End If
    
    'If endQuote = -1, something went horribly wrong
    If endQuote = -1 Then
        If isTooltip Then
            findCaptionInComplexQuotes = "MANUAL FIX REQUIRED FOR TOOLTIP ERROR FOR " & m_ObjectName & " IN " & m_FormName
            'MsgBox srcLines(lineNumber)
        Else
            findCaptionInComplexQuotes = "MANUAL FIX REQUIRED FOR TEXT PARSE ERROR AT LINE # " & lineNumber & " IN " & m_FileName
        End If
    Else
        findCaptionInComplexQuotes = Mid(srcLines(lineNumber), startQuote + 1, endQuote - startQuote - 1)
    End If
    
    'We now need to replace line breaks in the text.  These can appear in a variety of ways.  Replace them all.
    Dim lineBreak As String
    lineBreak = """ & vbCrLf & """
    If InStr(1, findCaptionInComplexQuotes, lineBreak) Then findCaptionInComplexQuotes = Replace(findCaptionInComplexQuotes, lineBreak, vbCrLf)
    lineBreak = """ & vbCrLf & vbCrLf & """
    If InStr(1, findCaptionInComplexQuotes, lineBreak) Then findCaptionInComplexQuotes = Replace(findCaptionInComplexQuotes, lineBreak, vbCrLf & vbCrLf)

End Function

'Given a line number and the original file contents, search for arbitrary text between two quotation marks
Private Function findCaptionInQuotes(ByRef srcLines() As String, ByRef lineNumber As Long, Optional ByVal startPosition As Long = 1) As String

    Dim startQuote As Long
    startQuote = InStr(startPosition, srcLines(lineNumber), """")
        
    Dim endQuote As Long
    endQuote = InStr(startQuote + 1, srcLines(lineNumber), """")
    
    If endQuote > 0 Then
        findCaptionInQuotes = Mid(srcLines(lineNumber), startQuote + 1, endQuote - startQuote - 1)
    Else
        findCaptionInQuotes = ""
    End If

End Function

'Given a line number and the original file contents, search for a "Caption" property.  Terminate if "End" is found.
Private Function findControlCaption(ByRef srcLines() As String, ByRef lineNumber As Long) As String

    Dim foundCaption As Boolean
    foundCaption = True

    Dim originalLineNumber As Long
    originalLineNumber = lineNumber

    'Start by retrieving the name of this object and storing it in a module-level string.  The calling function may
    ' need this if the caption meets certain criteria.
    Dim objectName As String
    objectName = Trim(srcLines(lineNumber))

    Dim sPos As Long
    sPos = Len(objectName)
    Do
        sPos = sPos - 1
    Loop While Mid(objectName, sPos, 1) <> " "
    
    m_ObjectName = Right(objectName, Len(objectName) - sPos)
    'MsgBox "OBJECT NAME: " & objectName

    Do While InStr(1, srcLines(lineNumber), "Caption", vbTextCompare) = 0
        lineNumber = lineNumber + 1
        
        'Some controls may not have a caption.  If this occurs, exit.
        ' NOTE: we must use a binary comparison here, or entries with "End" in them will fail to work!
        If (InStr(1, srcLines(lineNumber), "End", vbBinaryCompare) > 0) And Not (InStr(1, srcLines(lineNumber), "EndProperty", vbBinaryCompare) > 0) Then
            foundCaption = False
            lineNumber = originalLineNumber '+ 1
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
            findControlCaption = "MANUAL FIX REQUIRED FOR " & m_ObjectName & " IN " & m_FormName
        Else
        
            Dim startQuote As Long
            startQuote = InStr(1, srcLines(lineNumber), """")
    
            Dim endQuote As Long
            endQuote = InStrRev(srcLines(lineNumber), """")
        
            If endQuote > 0 Then
                findControlCaption = Mid(srcLines(lineNumber), startQuote + 1, endQuote - startQuote - 1)
            Else
                findControlCaption = ""
            End If
            
        End If
        lineNumber = originalLineNumber + 1
                
    Else
        findControlCaption = ""
    End If

End Function

'Given a line number and the original file contents, search for a "Caption" property.  Terminate if "End" is found.
Private Function findFormCaption(ByRef srcLines() As String, ByRef lineNumber As Long) As String

    Dim foundCaption As Boolean
    foundCaption = True

    'Start by retrieving the name of this form and storing it in a module-level string.  The calling function may
    ' need this if the caption meets certain criteria.
    Dim objectName As String
    objectName = Trim(srcLines(lineNumber))

    Dim sPos As Long
    sPos = Len(objectName)
    Do
        sPos = sPos - 1
    Loop While Mid(objectName, sPos, 1) <> " "
    
    m_FormName = Right(objectName, Len(objectName) - sPos)
    'MsgBox "FORM NAME: " & objectName
    
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
        endQuote = InStr(startQuote + 1, srcLines(lineNumber), """")
        
        findFormCaption = Mid(srcLines(lineNumber), startQuote + 1, endQuote - startQuote - 1)
    Else
        findFormCaption = ""
    End If

End Function





'Extract a list of all project files from a VBP file
Private Sub cmdSelectVBP_Click()

    Dim cDlg As cCommonDialog
    Set cDlg = New cCommonDialog
    
    m_VBPFile = "C:\PhotoDemon v4\PhotoDemon\PhotoDemon.vbp"
    
    If cDlg.VBGetOpenFileName(m_VBPFile, , True, False, False, True, "VBP - Visual Basic Project|*.vbp", , , "Please select a Visual Basic project file (VBP)", "vbp", Me.hWnd) Then
        lblVBP = "Active VBP: " & m_VBPFile
        m_VBPPath = getDirectory(m_VBPFile)
    Else
        Exit Sub
    End If
    
    'Load the file into a string array, split up line-by-line
    Dim vbpContents As String
    vbpContents = getFileAsString(m_VBPFile)
    vbpText = Split(vbpContents, vbCrLf)
    ReDim vbpFiles(0 To UBound(vbpText)) As String
    Dim numOfFiles As Long
    numOfFiles = 0
    
    Dim majorVer As String, minorVer As String, buildVer As String
    
    'Extract only the relevant file paths
    Dim i As Long
    For i = 0 To UBound(vbpText)
    
        'Check for forms
        If InStr(1, vbpText(i), "Form=", vbBinaryCompare) = 1 Then
            vbpFiles(numOfFiles) = m_VBPPath & Right(vbpText(i), Len(vbpText(i)) - 5)
            numOfFiles = numOfFiles + 1
        End If
        
        'Check for user controls
        If InStr(1, vbpText(i), "UserControl=", vbBinaryCompare) = 1 Then
            vbpFiles(numOfFiles) = m_VBPPath & Right(vbpText(i), Len(vbpText(i)) - 12)
            numOfFiles = numOfFiles + 1
        End If
        
        'Check for modules
        If InStr(1, vbpText(i), "Module=", vbBinaryCompare) = 1 Then
            vbpFiles(numOfFiles) = m_VBPPath & Trim(Right(vbpText(i), Len(vbpText(i)) - InStr(1, vbpText(i), ";")))
            numOfFiles = numOfFiles + 1
        End If
        
        'Check for classes
        If InStr(1, vbpText(i), "Class=", vbBinaryCompare) = 1 Then
            vbpFiles(numOfFiles) = m_VBPPath & Trim(Right(vbpText(i), Len(vbpText(i)) - InStr(1, vbpText(i), ";")))
            numOfFiles = numOfFiles + 1
        End If
        
        'Check for version numbers
        If InStr(1, vbpText(i), "MajorVer=", vbBinaryCompare) = 1 Then
            majorVer = Trim(Right(vbpText(i), Len(vbpText(i)) - InStr(1, vbpText(i), "=")))
        End If
        If InStr(1, vbpText(i), "MinorVer=", vbBinaryCompare) = 1 Then
            minorVer = Trim(Right(vbpText(i), Len(vbpText(i)) - InStr(1, vbpText(i), "=")))
        End If
        If InStr(1, vbpText(i), "RevisionVer=", vbBinaryCompare) = 1 Then
            buildVer = Trim(Right(vbpText(i), Len(vbpText(i)) - InStr(1, vbpText(i), "=")))
        End If
    
    Next i
    
    ReDim Preserve vbpFiles(0 To numOfFiles) As String
    
    'To make sure everything worked, dump the contents of the array into the list box on the left
    lstProjectFiles.Clear
    
    For i = 0 To UBound(vbpFiles)
        lstProjectFiles.AddItem vbpFiles(i)
    Next i
    
    'Build a complete version string
    versionString = majorVer & "." & minorVer & "." & buildVer
    
    cmdProcess.Caption = "Begin processing"

End Sub

'Given a full file name, remove everything but the directory structure
Private Function getDirectory(ByVal sString As String) As String
    
    Dim x As Long
    
    For x = Len(sString) - 1 To 1 Step -1
        If (Mid(sString, x, 1) = "/") Or (Mid(sString, x, 1) = "\") Then
            getDirectory = Left(sString, x)
            Exit Function
        End If
    Next x
    
    'getDirectory = ""
    
End Function

'Retrieve an entire file and return it as a string.
Private Function getFileAsString(ByVal fName As String) As String
        
    'Ensure that the file exists before attempting to load it
    If FileExist(fName) Then
        
        Dim fileNum As Integer
        fileNum = FreeFile
    
        Open fName For Binary As #fileNum
            getFileAsString = Space$(LOF(fileNum))
            Get #fileNum, , getFileAsString
        Close #fileNum
    
        'Remove all tabs from the source file (which may have been added in by an XML editor, but are not relevant to the translation process)
        If InStr(1, getFileAsString, vbTab) <> 0 Then getFileAsString = Replace(getFileAsString, vbTab, "")
    
    Else
        getFileAsString = ""
    End If
    
End Function

'Returns a boolean as to whether or not a given file exists
Private Function FileExist(ByRef fName As String) As Boolean
    Select Case (GetFileAttributesW(StrPtr(fName)) And vbDirectory) = 0
        Case True: FileExist = True
        Case Else: FileExist = (Err.LastDllError = ERROR_SHARING_VIOLATION)
    End Select
End Function

'Return the filename chunk of a path
Public Function getFilename(ByVal sString As String) As String

    Dim i As Long
    
    For i = Len(sString) - 1 To 1 Step -1
        If (Mid(sString, i, 1) = "/") Or (Mid(sString, i, 1) = "\") Then
            getFilename = Right(sString, Len(sString) - i)
            Exit Function
        End If
    Next i
    
End Function

'Count the number of words in a string (will not be 100% accurate, but that's okay)
Private Function countWordsInString(ByVal srcString As String) As Long

    If Trim$(srcString) <> "" Then

        Dim tmpArray() As String
        tmpArray = Split(Trim$(srcString), " ")
        
        Dim tmpWordCount As Long
        tmpWordCount = 0
        
        Dim i As Long
        For i = 0 To UBound(tmpArray)
            If IsAlpha(tmpArray(i)) Then tmpWordCount = tmpWordCount + 1
        Next i
        
        countWordsInString = tmpWordCount
        
    Else
        countWordsInString = 0
    End If

End Function

'VB's IsNumeric function can't detect percentage text (e.g. "100%").  PhotoDemon includes text like this, but I don't want such
' text translated - so manually check for it and reject such text if found.
Private Function IsNumericPercentage(ByVal srcString As String) As Boolean

    srcString = Trim(srcString)

    'Start by checking for a percent in the right-most position
    If Right$(srcString, 1) = "%" Then
        
        'If a percent was found, check the rest of the text to see if it's numeric
        If IsNumeric(Left$(srcString, Len(srcString) - 1)) Then
            IsNumericPercentage = True
        Else
            IsNumericPercentage = False
        End If
        
    Else
        IsNumericPercentage = False
    End If

End Function

'URLs shouldn't be translated.  Check for them and reject as necessary.
Private Function IsURL(ByVal srcString As String) As Boolean

    If (Left$(srcString, 6) = "ftp://") Or (Left$(srcString, 7) = "http://") Then
        IsURL = True
    Else
        IsURL = False
    End If

End Function

Private Sub Form_Load()
    
    'Build a blacklist of phrases that are in the software, but do not need to be translated.  (These are complex phrases that
    ' may include things like names, but the automatic text generator has no way of knowing that the text is non-translatable.)
    ReDim m_Blacklist(0) As String
    m_numOfBlacklistEntries = 0
    
    addBlacklist "PhotoDemon by Tanner Helland - www.tannerhelland.com"
    addBlacklist "(X, Y)"
    addBlacklist "16:1 (1600%)"
    addBlacklist "8:1 (800%)"
    addBlacklist "4:1 (400%)"
    addBlacklist "2:1 (200%)"
    addBlacklist "1:2 (50%)"
    addBlacklist "1:4 (25%)"
    addBlacklist "1:8 (12.5%)"
    addBlacklist "1:16 (6.25%)"
    addBlacklist "pngquant 2.1.1"
    addBlacklist "zLib 1.2.8"
    addBlacklist "EZTwain 1.18"
    addBlacklist "FreeImage 3.16.0"
    addBlacklist "ExifTool 9.62"
    addBlacklist "X.X"
    addBlacklist "XX.XX.XX"
    addBlacklist "pngquant"
    addBlacklist "zLib"
    addBlacklist "EZTwain"
    addBlacklist "FreeImage"
    addBlacklist "ExifTool"
    addBlacklist "tannerhelland.com/contact"
    addBlacklist "photodemon.org/about/contact"
    addBlacklist "photodemon.org/about/contact/"
    addBlacklist "HTML / CSS"
    
End Sub

Private Sub addBlacklist(ByVal blString As String)

    m_Blacklist(m_numOfBlacklistEntries) = blString
    m_numOfBlacklistEntries = m_numOfBlacklistEntries + 1
    ReDim Preserve m_Blacklist(0 To m_numOfBlacklistEntries) As String

End Sub

Private Function isBlacklisted(ByVal blString As String) As Boolean

    isBlacklisted = False
    
    'Search the blacklist for this string.  If it is found, immediately return TRUE.
    Dim i As Long
    For i = 0 To m_numOfBlacklistEntries - 1
        If StrComp(blString, m_Blacklist(i), vbTextCompare) = 0 Then
            isBlacklisted = True
            Exit Function
        End If
    Next i

End Function

'Used to roughly estimate if a string is purely alphabetical (this project uses it to check if a statement is a "word" or not).
Private Function IsAlpha(ByRef srcString As String) As Boolean
    
    Dim charID As Byte
   
    IsAlpha = True
    
    Dim i As Long
    For i = 1 To Len(srcString)
        charID = Asc(UCase$(Mid$(srcString, i, 1)))
        
        'First, check to see if the character lies outside the ASCII alphabet range
        If ((charID < 65) Or (charID > 90)) Then
        
            'Next, if the length of the source string is greater than one, check to see if this is a hyphenated or punctuated word
            If Len(srcString) > 1 Then
                
                'Allow certain punctuation to still count as "alphabetical"
                If Not ((charID = 33) Or (charID = 34) Or (charID = 38) Or (charID = 40) Or (charID = 41) Or (charID = 44) Or (charID = 45) Or (charID = 46) Or (charID = 58) Or (charID = 59) Or (charID = 63) Or (charID = 64) Or (charID = 96)) Then
                    IsAlpha = False
                    Exit For
                End If
                
            Else
                IsAlpha = False
                Exit For
            End If
            
        End If
    Next i
    
End Function
