VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000005&
   Caption         =   " PhotoDemon i18n manager"
   ClientHeight    =   8190
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
   ScaleHeight     =   546
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   969
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConvertLabels 
      Caption         =   "Convert labels in selected project file to pdLabel format.  (This cannot be undone; use cautiously!)"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   2520
      TabIndex        =   15
      Top             =   4380
      Width           =   11775
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
      Left            =   4800
      TabIndex        =   13
      Top             =   7320
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
      Left            =   8520
      TabIndex        =   7
      Top             =   6360
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
      Left            =   4800
      TabIndex        =   6
      Top             =   6360
      Width           =   3255
   End
   Begin VB.CommandButton cmdMaster 
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
      Left            =   1080
      TabIndex        =   5
      Top             =   6360
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
      Caption         =   "other support tools:"
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
      Left            =   240
      TabIndex        =   14
      Top             =   4500
      Width           =   2040
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
      Height          =   495
      Left            =   4080
      TabIndex        =   11
      Top             =   5280
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
      Left            =   720
      TabIndex        =   10
      Top             =   5880
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   5280
      Width           =   2970
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      X1              =   8
      X2              =   960
      Y1              =   344
      Y2              =   344
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
'Last updated: 13/July/22
'Last update: switch to using various PhotoDemon code files directly (instead of using local, manually edited copies)
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
'      Also, given this project's purpose, the code is pretty ugly.  Organization is minimal.  Read at your own risk.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit
Option Compare Binary

'Variables used to generate the master translation file
Private m_VBPFile As String, m_VBPPath As String
Private m_FormName As String, m_ObjectName As String, m_FileName As String
Private m_NumOfPhrasesFound As Long, m_NumOfPhrasesWritten As Long, m_numOfWords As Long
Private vbpText() As String, vbpFiles() As String
Private m_outputText As pdString, outputFile As String

'Variables used to merge old language files with new ones
Private m_MasterText As String, m_OldLanguageText As String, m_NewLanguageText As String
Private m_OldLanguagePath As String

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

'During silent mode (used to synchronize localizations), we use a fast string hash table to update
' language files.  This greatly improves performance, especially given how many language files PD ships.
Private m_PhraseCollection As pdStringHash

'New support function for auto-converting old common control labels to PD's new pdLabel object.  If successful, this will save me a ton of time
' manually converting all the labels in the program.
Private Sub cmdConvertLabels_Click()

    'Make sure a file has been selected
    If (lstProjectFiles.ListIndex <> -1) Then

        'Read the file into a string array
        Dim srcFilename As String
        srcFilename = lstProjectFiles.List(lstProjectFiles.ListIndex)
        
        Dim fileContents As String
        Files.FileLoadAsString srcFilename, fileContents, True
        
        Dim fileLines() As String
        fileLines = Split(fileContents, vbCrLf)
        
        Dim curLineNumber As Long, curLineText As String
        curLineNumber = 0
        
        Dim startLine As Long, endLine As Long, i As Long, numLabelsReplaced As Long
        numLabelsReplaced = 0
        
        Dim ignoreThisLine As Boolean
                
        'Now, start processing the file one line at a time, searching for label entries as we go
        Do
        
            curLineText = fileLines(curLineNumber)
            
            'Before processing this line, make sure is isn't a comment.
            If Left$(Trim$(curLineText), 1) = "'" Then GoTo nextLine
            
            'Check for a VB label identifier.  The format of these declarations is always the same.
            If (InStr(1, UCase$(curLineText), "BEGIN VB.LABEL", vbBinaryCompare) > 0) Then
            
                'This line is the start of a VB label identifier.
                startLine = curLineNumber
                numLabelsReplaced = numLabelsReplaced + 1
                
                'Next, we are going to continue looping through the file, looking for the End identifier that marks the end of this label's properties.
                Do While (StrComp(Trim$(curLineText), "End", vbBinaryCompare) = 0)
                    curLineNumber = curLineNumber + 1
                    curLineText = fileLines(curLineNumber)
                Loop
                
                'curLineNumber now contains the index of the last line of this label's properties.  Mark it.
                endLine = curLineNumber
                
                'Convert the starting line to represent a pdLabel instead of a VB label
                fileLines(startLine) = Replace$(fileLines(startLine), "VB.Label", "PhotoDemon.pdLabel")
                
                'Now, we need to iterate through the properties this label might have.  pdLabel objects have a much smaller property
                ' list than standard VB labels.  Incompatible properties must be removed, or the IDE will throw errors.
                For i = startLine + 1 To endLine - 1
                    
                    'See if this line contains a valid pdLabel property
                    ignoreThisLine = IsValidPDLabelProperty(fileLines(i))
                    
                    'If this line does not contain a valid pdLabel property, perform some special checks for compatible properties
                    ' with different names.
                    If Not ignoreThisLine Then
                        
                        'Font size descriptions should be renamed from Size to FontSize
                        If InStr(1, Trim$(fileLines(i)), vbTab & "Size") Then
                            fileLines(i) = Replace$(fileLines(i), "Size", "FontSize")
                            ignoreThisLine = True
                        End If
                    
                    End If
                    
                    'If the line is still not compatible, replace its contents with a uniquely identifiable string
                    If Not ignoreThisLine Then fileLines(i) = "INVALID PROPERTY LINE"
                    
                Next i
            
            End If
            
            
        'Continue to the next line in the file
nextLine:
            curLineNumber = curLineNumber + 1
    
        Loop While curLineNumber < UBound(fileLines)
        
        'The fileLines array now contains the original file's contents, but with all invalid lines marked for removal.
        ' We are now going to overwrite the original file (gasp) with these new contents.
        
        'Start by killing the original copy
        Files.FileDeleteIfExists srcFilename
        
        'Open the file anew
        Dim fHandle As Integer
        fHandle = FreeFile
        
        Open srcFilename For Output As #fHandle
        
            'Write the modified file contents out to file
            For i = LBound(fileLines) To UBound(fileLines)
                
                If StrComp(fileLines(i), "INVALID PROPERTY LINE", vbBinaryCompare) <> 0 Then
                    Print #fHandle, fileLines(i)
                End If
                
            Next i
        
        Close #fHandle
        
        MsgBox numLabelsReplaced & " labels replaced successfully.", vbOKOnly + vbApplicationModal + vbInformation, "Label replacement complete"
            
    Else
        MsgBox "Select a file first.", vbOKOnly + vbApplicationModal + vbInformation, "No file selected"
    End If
    
End Sub

'See if a given line from a VB Form contains a valid pdLabel property
Private Function IsValidPDLabelProperty(ByVal srcString As String) As Boolean
    
    IsValidPDLabelProperty = False
    
    'Trim the source string to make comparisons easier
    srcString = Trim$(LCase$(srcString))
    
    'The list of valid properties is hardcoded.
    If InStr(1, srcString, "Alignment", vbBinaryCompare) > 0 Then IsValidPDLabelProperty = True
    If InStr(1, srcString, "BackColor", vbBinaryCompare) > 0 Then IsValidPDLabelProperty = True
    If InStr(1, srcString, "Caption", vbBinaryCompare) > 0 Then IsValidPDLabelProperty = True
    
    'Attempt a periodic exit (since we can't short-circuit VB code otherwise argh)
    If IsValidPDLabelProperty Then Exit Function
    
    If InStr(1, srcString, "Enabled", vbBinaryCompare) > 0 Then IsValidPDLabelProperty = True
    If InStr(1, srcString, "ForeColor", vbBinaryCompare) > 0 Then IsValidPDLabelProperty = True
    If InStr(1, srcString, "Height", vbBinaryCompare) > 0 Then IsValidPDLabelProperty = True
    If IsValidPDLabelProperty Then Exit Function
    
    If InStr(1, srcString, "Index", vbBinaryCompare) > 0 Then IsValidPDLabelProperty = True
    If InStr(1, srcString, "Layout", vbBinaryCompare) > 0 Then IsValidPDLabelProperty = True
    If InStr(1, srcString, "Left", vbBinaryCompare) > 0 Then IsValidPDLabelProperty = True
    If InStr(1, srcString, "Top", vbBinaryCompare) > 0 Then IsValidPDLabelProperty = True
    If InStr(1, srcString, "Width", vbBinaryCompare) > 0 Then IsValidPDLabelProperty = True
    
End Function

Private Sub cmdMaster_Click()

    Dim cDialog As pdOpenSaveDialog
    Set cDialog = New pdOpenSaveDialog
    
    'This project should be located in a sub-path of a normal PhotoDemon install.
    ' You could easily modify this to dynamically calculate the corresponding folder, but I hard-code
    ' a path to the typical folder on my dev PC.
    Dim fPath As String
    fPath = "C:\PhotoDemon v4\PhotoDemon\App\PhotoDemon\Languages\Master\MASTER.xml"
    
    If cDialog.GetOpenFileName(fPath, , True, False, "XML - PhotoDemon Language File|*.xml", 1, , "Please select a PhotoDemon language file (XML)", "xml", Me.hWnd) Then
        Files.FileLoadAsString fPath, m_MasterText, True
        
        'Remove tabstops, if any exist
        m_MasterText = Replace$(m_MasterText, vbTab, vbNullString, 1, -1, vbBinaryCompare)
    End If
    
End Sub

Private Sub ReplaceTopLevelTag(ByVal origTagName As String, ByRef sourceTextMaster As String, ByRef sourceTextTranslation As String, ByRef destinationText As String, Optional ByVal alsoIncrementVersion As Boolean = True)

    Dim openTagName As String, closeTagName As String
    openTagName = "<" & origTagName & ">"
    closeTagName = "</" & origTagName & ">"
    
    Dim findText As String, replaceText As String
    findText = openTagName & GetTextBetweenTags(sourceTextMaster, origTagName) & closeTagName
    
    'A special check is applied to the "langversion" tag.  Whenever this function is used, a merge is taking place; as such, we want to
    ' auto-increment the language's version number to trigger an update on client machines.
    If (StrComp(origTagName, "langversion", vbBinaryCompare) = 0) And alsoIncrementVersion Then
        
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

Private Sub cmdMerge_Click()

    'Make sure our source file strings are not empty
    If (LenB(m_MasterText) = 0) Or (LenB(m_OldLanguageText) = 0) Then
        MsgBox "One or more source files are missing.  Supply those before attempting a merge."
        Exit Sub
    End If
    
    'Start by copying the contents of the master file into the destination string.  We will use that as our base, and update it
    ' with the old translations as best we can.
    m_NewLanguageText = m_MasterText
        
    Dim sPos As Long
    sPos = InStr(1, m_NewLanguageText, "<phrase>", vbBinaryCompare)
    
    Dim origText As String, translatedText As String
    Dim findText As String, replaceText As String
    
    'Copy over all top-level language and author information
    ReplaceTopLevelTag "langid", m_MasterText, m_OldLanguageText, m_NewLanguageText
    ReplaceTopLevelTag "langname", m_MasterText, m_OldLanguageText, m_NewLanguageText
    ReplaceTopLevelTag "langversion", m_MasterText, m_OldLanguageText, m_NewLanguageText
    ReplaceTopLevelTag "langstatus", m_MasterText, m_OldLanguageText, m_NewLanguageText
    ReplaceTopLevelTag "author", m_MasterText, m_OldLanguageText, m_NewLanguageText
        
    Dim phrasesProcessed As Long, phrasesFound As Long, phrasesMissed As Long
    phrasesProcessed = 0
    phrasesFound = 0
    phrasesMissed = 0
    
    'Start parsing the master text for <phrase> tags
    Do
    
        phrasesProcessed = phrasesProcessed + 1
    
        'Retrieve the original text associated with this phrase tag
        origText = GetTextBetweenTags(m_MasterText, "original", sPos)
        
        'Attempt to retrieve a translation for this phrase using the old language file
        translatedText = GetTranslationTagFromCaption(origText)
                
        'If no translation was found, and this string contains vbCrLf characters, replace them with plain vbLF characters and try again
        If (LenB(translatedText) = 0) Then
            If (InStr(1, origText, vbCrLf) > 0) Then
                translatedText = GetTranslationTagFromCaption(Replace$(origText, vbCrLf, vbLf))
            End If
        End If
                
        'If a translation was found, insert it into the new file
        If (LenB(translatedText) <> 0) Then
            
            'As a failsafe, try the same thing without tabs
            findText = "<original>" & origText & "</original>" & vbCrLf & "<translation></translation>"
            replaceText = "<original>" & origText & "</original>" & vbCrLf & "<translation>" & translatedText & "</translation>"
            m_NewLanguageText = Replace$(m_NewLanguageText, findText, replaceText)
            
            phrasesFound = phrasesFound + 1
        Else
            phrasesMissed = phrasesMissed + 1
        End If
    
        'Find the next occurrence of a <phrase> tag
        sPos = InStr(sPos + 1, m_MasterText, "<phrase>", vbBinaryCompare)
        
        If ((phrasesProcessed And 7) = 0) Then
            lblUpdates.Caption = phrasesProcessed & " phrases processed.  (" & phrasesFound & " found, " & phrasesMissed & " missed)"
            lblUpdates.Refresh
            If ((phrasesProcessed And 63) = 0) Then VBHacks.DoEvents_PaintOnly Me.hWnd, True
        End If
        
    Loop While sPos > 0
    
    'Prompt the user to save the results
    Dim cDialog As pdOpenSaveDialog
    Set cDialog = New pdOpenSaveDialog
    
    Dim fPath As String
    fPath = m_OldLanguagePath
    
    If cDialog.GetSaveFileName(fPath, , True, "XML - PhotoDemon Language File|*.xml", 1, , "Save the merged language file (XML)", "xml", Me.hWnd) Then
    
        If Files.FileExists(fPath) Then
            MsgBox "File already exists!  Too dangerous to overwrite - please perform the merge again."
            Exit Sub
        End If
        
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
    End If

End Function

'Given the original caption of a message or control, return the matching translation from the active translation file
Private Function GetTranslationTagFromCaption(ByVal origCaption As String) As String

    'Remove white space from the caption (if necessary, white space will be added back in after retrieving the translation from file)
    PreprocessText origCaption
    origCaption = "<original>" & origCaption & "</original>"
    
    Dim phraseLocation As Long
    phraseLocation = GetPhraseTagLocation(origCaption)
    
    'Make sure a phrase tag was found
    If (phraseLocation > 0) Then
        
        'Retrieve the <translation> tag inside this phrase tag
        origCaption = GetTextBetweenTags(m_OldLanguageText, "translation", phraseLocation)
        GetTranslationTagFromCaption = origCaption
        
    Else
        GetTranslationTagFromCaption = vbNullString
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

'New option added 09 September 2013 - Merge all language files automatically.  This will save me some trouble in the future.
Private Sub cmdMergeAll_Click()

    Dim srcFolder As String
    srcFolder = "C:\PhotoDemon v4\PhotoDemon\App\PhotoDemon\Languages\"
    
    'Auto-load the latest master language file and remove tabstops from the text (if any exist)
    Files.FileLoadAsString srcFolder & "Master\MASTER.xml", m_MasterText, True
    m_MasterText = Replace$(m_MasterText, vbTab, vbNullString, 1, -1, vbBinaryCompare)
    
    'Rather than backup the old files to the dev language folder (which is confusing),
    ' I now place them inside a dedicated backup folder.
    Dim backupFolder As String
    backupFolder = "C:\PhotoDemon v4\PhotoDemon\no_sync\PD_Language_File_Tmp\dev_backup\"
    
    'Iterate through every language file in the default PD directory
    'Scan the translation folder for .xml files.  Ignore anything that isn't XML.
    Dim chkFile As String
    chkFile = Dir$(srcFolder & "*.xml", vbNormal)
    
    'String constants to prevent constant allocations
    Const PHRASE_START As String = "<phrase>"
    Const AMPERSAND_CHAR As String = "&"
    
    Do While (LenB(chkFile) > 0)
        
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
        
        'Build a collection of all phrases in the current translation file.  Some phrases may not be
        ' translated and that's fine - we'll leave them blank and simply plug-in the phrases we *do* have.
        Set m_PhraseCollection = New pdStringHash
        
        If (numOldPhrases > 0) Then
            
            Dim i As Long
            For i = 0 To numOldPhrases - 1
                
                origText = oldLangXML.GetUniqueTag_String("original", vbNullString, phraseLocations(i))
                translatedText = oldLangXML.GetUniqueTag_String("translation", vbNullString, phraseLocations(i) + Len(origText))
                
                'Old PhotoDemon language files used manually inserted & characters for keyboard accelerators.
                ' Accelerators are now handled automatically on a per-language basis.  To ensure work isn't lost
                ' when upgrading these old files, strip any accelerators from the incoming text.
                If (InStr(1, origText, AMPERSAND_CHAR, vbBinaryCompare) <> 0) Then origText = Replace$(origText, AMPERSAND_CHAR, vbNullString, 1, -1, vbBinaryCompare)
                If (InStr(1, translatedText, AMPERSAND_CHAR, vbBinaryCompare) <> 0) Then origText = Replace$(translatedText, AMPERSAND_CHAR, vbNullString, 1, -1, vbBinaryCompare)
                
                m_PhraseCollection.AddItem origText, translatedText
                
            Next i
            
        End If
        
        'BEGIN COPY OF CODE FROM cmdMerge (with changes to accelerate the process, since we don't need a UI)
        
            'Make sure our source file strings are not empty
            If (LenB(m_MasterText) = 0) Or (numOldPhrases <= 0) Then
                Debug.Print "One or more source files are missing.  Supply those before attempting a merge."
                Exit Sub
            End If
            
            'Start by copying the contents of the master file into the destination string.
            ' We will use that as our base, and update it with the old translations as best we can.
            m_NewLanguageText = m_MasterText
                
            Dim sPos As Long
            sPos = InStr(1, m_NewLanguageText, PHRASE_START)
            
            'Dim origText As String, translatedText As String
            'Dim findText As String, replaceText As String
            
            'Copy over all top-level language and author information
            ReplaceTopLevelTag "langid", m_MasterText, m_OldLanguageText, m_NewLanguageText
            ReplaceTopLevelTag "langname", m_MasterText, m_OldLanguageText, m_NewLanguageText
            ReplaceTopLevelTag "langstatus", m_MasterText, m_OldLanguageText, m_NewLanguageText
            ReplaceTopLevelTag "author", m_MasterText, m_OldLanguageText, m_NewLanguageText
            ReplaceTopLevelTag "langversion", m_MasterText, m_OldLanguageText, m_NewLanguageText, False
                
            Dim phrasesProcessed As Long, phrasesFound As Long, phrasesMissed As Long
            phrasesProcessed = 0
            phrasesFound = 0
            phrasesMissed = 0
            
            'Start parsing the master text for <phrase> tags
            Do
            
                phrasesProcessed = phrasesProcessed + 1
            
                'Retrieve the original text associated with this phrase tag
                origText = GetTextBetweenTags(m_MasterText, "original", sPos)
                
                'Attempt to retrieve a translation for this phrase using the old language file
                If (Not m_PhraseCollection.GetItemByKey(origText, translatedText)) Then
                
                    translatedText = vbNullString
                    
                    'If no translation was found, and this string contains vbCrLf characters,
                    ' replace them with plain vbLF characters and try again
                    If (InStr(1, origText, vbCrLf, vbBinaryCompare) > 0) Then
                        translatedText = GetTranslationTagFromCaption(Replace$(origText, vbCrLf, vbLf, 1, -1, vbBinaryCompare))
                    End If
                    
                End If
                
                'Remove any tab stops from the translated text (which may have been added by an outside editor)
                If (InStr(1, translatedText, vbTab, vbBinaryCompare) <> 0) Then translatedText = Replace$(translatedText, vbTab, vbNullString, 1, -1, vbBinaryCompare)
                
                'If a translation was found, insert it into the new file
                Const ORIG_TAG_OPEN As String = "<original>"
                Const ORIG_TAG_CLOSE As String = "</original>" & vbCrLf & "<translation></translation>"
                Const TRANSLATE_TAG_INTERIOR As String = "</original>" & vbCrLf & "<translation>"
                Const TRANSLATE_TAG_CLOSE As String = "</translation>"
                If (LenB(translatedText) <> 0) Then
                    findText = ORIG_TAG_OPEN & origText & ORIG_TAG_CLOSE
                    replaceText = ORIG_TAG_OPEN & origText & TRANSLATE_TAG_INTERIOR & translatedText & TRANSLATE_TAG_CLOSE
                    m_NewLanguageText = Replace$(m_NewLanguageText, findText, replaceText, 1, -1, vbBinaryCompare)
                    phrasesFound = phrasesFound + 1
                Else
                    phrasesMissed = phrasesMissed + 1
                End If
            
                'Find the next occurrence of a <phrase> tag
                sPos = InStr(sPos + 1, m_MasterText, PHRASE_START, vbBinaryCompare)
                
                If (Not m_SilentMode) Then
                    If ((phrasesProcessed And 127) = 0) Then
                        lblUpdates.Caption = chkFile & ": " & phrasesProcessed & " phrases processed (" & phrasesFound & " found, " & phrasesMissed & " missed)"
                        lblUpdates.Refresh
                        VBHacks.DoEvents_PaintOnly Me.hWnd, True
                    End If
                End If
            
            Loop While sPos > 0
            
            'See if the old and new language files are equal.  If they are, we won't bother writing the results out to file.
            If (LenB(Trim$(m_NewLanguageText)) = LenB(Trim$(m_OldLanguageText))) Then
                Debug.Print "New language file and old language file are identical for " & chkFile & ".  Merge abandoned."
            Else
                
                'Update the version number by 1
                ReplaceTopLevelTag "langversion", m_MasterText, m_OldLanguageText, m_NewLanguageText
                
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
    
    lblUpdates.Caption = "All language files processed successfully."
    lblUpdates.Refresh
    DoEvents

End Sub

Private Sub cmdOldLanguage_Click()
    
    Dim cDialog As pdOpenSaveDialog
    Set cDialog = New pdOpenSaveDialog
    
    Dim fPath As String
    fPath = "C:\PhotoDemon v4\PhotoDemon\App\PhotoDemon\Languages\"
    
    Dim tmpLangFile As String
    
    If cDialog.GetOpenFileName(tmpLangFile, , True, False, "XML - PhotoDemon Language File|*.xml", 1, fPath, "Please select a PhotoDemon language file (XML)", "xml", Me.hWnd) Then
        
        m_OldLanguagePath = tmpLangFile
        
        'Load the language file and strip tabstops from it
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
    
    'Note whether duplicate phrases are automatically removed
    m_RemoveDuplicates = CBool(chkRemoveDuplicates)
    
    'Start by preparing the XML header
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
    numOfFiles = UBound(vbpFiles)
    
    m_NumOfPhrasesFound = 0
    m_NumOfPhrasesWritten = 0
    m_numOfWords = 0
    
    Dim i As Long
    For i = 0 To numOfFiles
        cmdProcess.Caption = "Processing project file " & i + 1 & " of " & numOfFiles + 1
        ProcessFile vbpFiles(i)
    Next i
    
    'With processing complete, write out our final stats (just for fun)
    m_outputText.AppendLineBreak
    m_outputText.AppendLineBreak
    m_outputText.AppendLine vbTab & vbTab & "<!-- Automatic text extraction complete. -->"
    m_outputText.AppendLineBreak
    
    'Updated 09 September 2013: write out phrase count as an actual tag, which PD's new language editor can use to approximate a max
    ' value for its progress bar when loading the language file.
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
    
    'We are now going to compare the length of the old file and new file.  If the lengths match, there's no reason to write out this new file.
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
    
End Sub

'Given a VB file (form, module, class, user control), extract any relevant text from it
Private Sub ProcessFile(ByVal srcFile As String)

    If (LenB(srcFile) = 0) Then Exit Sub

    m_FileName = Files.FileGetName(srcFile)
    
    'Certain files can be ignored.  I generate this list manually, on account of knowing which files (classes, mostly) contain
    ' no special text.  I could probably add many more files to this list, but I primarily want to blacklist those that create
    ' parsing problems.  (The tooltip classes are particularly bad, since they use the phrase "tooltip" frequently, which messes
    ' up the parser as it thinks it's found hundreds of tooltips in each file.)
    Select Case m_FileName
    
        Case "clsToolTip.cls", "pdToolTip.cls", "clsControlImages.cls"
            Exit Sub
            
        Case "pdFilterSupport.cls", "cSelfSubHookCallback.cls"
            Exit Sub
            
        Case "VBP_PublicVariables.bas", "pdParamString.cls", "VBP_ToolbarDebug.frm"
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
    
    'If this file is a form file, the second line of the file will contain the text: "Begin VB.FORM FormName", where FormName
    ' is the name of the form. By inserting the form's name into our translation file, the translation engine can use it to quickly
    ' locate all translations on that form.
    Dim shortcutName As String
    shortcutName = vbNullString
    
    If Right$(m_FileName, 3) = "frm" Then
        Dim findName() As String
        findName = Split(fileLines(1), " ")
        shortcutName = findName(2)
    End If
    
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
        
        'Every one of these requires a unique mechanism for checking the text.
        
        'Note that some of these mechanisms will modify the current line number.  These require the line number, passed
        ' ByRef, for that purpose.
        
        'If any of the functions are successful, they will return the string that needs to be added to the XML file
        
        '1) Check for a form caption
        If InStr(1, ucCurLineText, "BEGIN VB.FORM", vbBinaryCompare) Then
            processedText = FindFormCaption(fileLines, curLineNumber)
                
        '2) Check for a control caption.  (This has to be handled slightly differently than form caption.)
        ElseIf ((InStr(1, ucCurLineText, "BEGIN VB.", vbBinaryCompare) > 0) Or (InStr(1, ucCurLineText, "BEGIN PHOTODEMON.", vbBinaryCompare) > 0)) And (InStr(1, ucCurLineText, "PICTUREBOX", vbBinaryCompare) = 0) And (InStr(1, curLineText, "ComboBox") = 0) And (InStr(1, curLineText, ".Shape") = 0) And (InStr(1, curLineText, "TextBox") = 0) And (InStr(1, curLineText, "HScrollBar") = 0) And (InStr(1, curLineText, "VScrollBar") = 0) Then
            processedText = FindControlCaption(fileLines, curLineNumber)
        
        '3) Check for tooltip text on PD controls (assigned via the custom .AssignTooltip function)
        ElseIf (InStr(1, ucCurLineText, ".ASSIGNTOOLTIP ") > 0) And (InStr(1, curLineText, "ByVal") = 0) Then
            
            'Process the tooltip text itself
            processedText = FindTooltipMessage(fileLines, curLineNumber, False, toolTipSecondCheckNeeded)
            
            'Process the title, if any
            If toolTipSecondCheckNeeded Then processedTextSecondary = FindMsgBoxTitle(fileLines, curLineNumber)
            
        '4) Check for text added to a combo box or list box control at run-time
        ElseIf InStr(1, curLineText, ".AddItem """) <> 0 Then
            processedText = FindCaptionInComplexQuotes(fileLines, curLineNumber)
            
        '5) Check for message calls
        ElseIf InStr(1, curLineText, "Message """) <> 0 Then
            processedText = FindCaptionInQuotes(fileLines, curLineNumber)
        
        '6) Check for message box text, including 7) message box titles (which must also be translated)
        ElseIf (InStr(1, ucCurLineText, "PDMSGBOX", vbTextCompare) <> 0) Then
        
            'First, retrieve the message box text itself
            processedText = FindMsgBoxText(fileLines, curLineNumber)
            
            'Next, retrieve the message box title
            processedTextSecondary = FindMsgBoxTitle(fileLines, curLineNumber)
        
        '7) Specific to PhotoDemon - check for action names that may not be present elsewhere
        ElseIf InStr(1, curLineText, "Process """) <> 0 Then
            processedText = FindCaptionInQuotes(fileLines, curLineNumber, InStr(1, curLineText, "Process """))
        
        '7.5) Now that pdLabel objects manage their own translations, we should also check for caption assignments
        ElseIf InStr(1, curLineText, "Caption = """, vbBinaryCompare) <> 0 Then
            processedText = FindCaptionInQuotes(fileLines, curLineNumber, 1)
        
        End If
        
        '8) Check for text that has been manually marked for translation (e.g. g_Language.TranslateMessage("xyz"))
        '    NOTE: as of 07 June 2013, each line can contain two translation calls (instead of just one)
        '
        'Note that this check is performed regardless of previous checks, to make sure no translations are missed.
        If InStr(1, curLineText, "g_Language.TranslateMessage(""") Then
            processedText = FindMessage(fileLines, curLineNumber)
            processedTextSecondary = FindMessage(fileLines, curLineNumber, True)
        End If
        
        'DEBUG! Check for certain text entries here
        'If (shortcutName = "FormLens") And Len(Trim$(processedText)) <> 0 Then MsgBox processedText
        
        'We now have text in potentially two places: processedText, and processedTextSecondary (for message box titles)
        chkText = Trim$(processedText)
        
        'Only pass the text along if it isn't blank, or a number, or a symbol, or a manually blacklisted phrase
        If (LenB(chkText) <> 0) Then
            If (Not IsNumeric(chkText)) And (Not IsNumericPercentage(chkText)) And (Not IsBlacklisted(chkText)) Then
                If (chkText <> ".") And (chkText <> "-") And (Not IsURL(chkText)) Then
                    numOfPhrasesFound = numOfPhrasesFound + 1
                    If AddPhrase(processedText) Then numOfPhrasesWritten = numOfPhrasesWritten + 1
                End If
            End If
        End If
        
        chkText = Trim$(processedTextSecondary)
        
        'Do the same for the secondary text
        If (LenB(chkText) <> 0) Then
            If (Not IsNumeric(chkText)) And (Not IsNumericPercentage(chkText)) And (Not IsBlacklisted(chkText)) Then
                If (chkText <> ".") And (chkText <> "-") And (Not IsURL(chkText)) Then
                    numOfPhrasesFound = numOfPhrasesFound + 1
                    If AddPhrase(processedTextSecondary) Then numOfPhrasesWritten = numOfPhrasesWritten + 1
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
Private Function AddPhrase(ByRef phraseText As String) As Boolean
                        
    'Replace double double-quotes (which are required in code) with just one set of double-quotes
    If InStr(1, phraseText, """""") Then phraseText = Replace$(phraseText, """""", """")
            
    'Next, do the same pre-processing that we do in the translation engine
    
    '1) Trim the text.  Extra spaces will be handled by the translation engine.
    phraseText = Trim$(phraseText)
    
    '2) Check for a trailing "..." and remove it
    If Right$(phraseText, 3) = "..." Then phraseText = Left$(phraseText, Len(phraseText) - 3)
    
    '3) Check for a trailing colon ":" and remove it
    If Right$(phraseText, 1) = ":" Then phraseText = Left$(phraseText, Len(phraseText) - 1)
    
    'This phrase is now ready to write to file.
    
    'Before writing the phrase out, check to see if it already exists
    If m_RemoveDuplicates Then
                
        If m_outputText.StrStr("<original>" & phraseText & "</original>") Then
            AddPhrase = False
        Else
            AddPhrase = (LenB(phraseText) <> 0)
        End If
        
    Else
        AddPhrase = (LenB(phraseText) <> 0)
    End If
    
    'If the phrase does not exist, add it now
    If AddPhrase Then
        m_outputText.AppendLineBreak
        m_outputText.AppendLineBreak
        m_outputText.AppendLine vbTab & vbTab & "<phrase>"
        m_outputText.Append vbTab & vbTab & vbTab & "<original>"
        m_outputText.Append phraseText
        m_outputText.AppendLine "</original>"
        m_outputText.AppendLine vbTab & vbTab & vbTab & "<translation></translation>"
        m_outputText.Append vbTab & vbTab & "</phrase>"
        m_numOfWords = m_numOfWords + CountWordsInString(phraseText)
    End If
    
End Function

'Given a line number and the original file contents, search for a custom PhotoDemon translation request
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
    If InStr(1, FindMessage, lineBreak) Then FindMessage = Replace(FindMessage, lineBreak, vbCrLf)
    lineBreak = """ & vbCrLf & vbCrLf & """
    If InStr(1, FindMessage, lineBreak) Then FindMessage = Replace(FindMessage, lineBreak, vbCrLf & vbCrLf)

    
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
    
    'If text is not found, try again, using a different tooltip assignment command
    If initPosition = 0 Then
        If inReverse Then
            initPosition = InStrRev(UCase$(srcLines(lineNumber)), ".SETTOOLTIP ")
        Else
            initPosition = InStr(1, UCase$(srcLines(lineNumber)), ".SETTOOLTIP ")
        End If
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
    'MsgBox "OBJECT NAME: " & objectName

    Do While InStr(1, UCase$(srcLines(lineNumber)), "CAPTION", vbBinaryCompare) = 0
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
        endQuote = InStrRev(srcLines(lineNumber), """")
        
        FindFormCaption = Mid$(srcLines(lineNumber), startQuote + 1, endQuote - startQuote - 1)
        
    Else
        FindFormCaption = vbNullString
    End If

End Function


'Extract a list of all project files from a VBP file
Private Sub cmdSelectVBP_Click()
    
    m_VBPFile = "C:\PhotoDemon v4\PhotoDemon\PhotoDemon.vbp"
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
    ReDim vbpFiles(0 To UBound(vbpText)) As String
    Dim numOfFiles As Long
    numOfFiles = 0
    
    Dim majorVer As String, minorVer As String, buildVer As String
    
    'Extract only the relevant file paths
    Dim i As Long
    For i = 0 To UBound(vbpText)
    
        'Check for forms
        If InStr(1, vbpText(i), "Form=", vbBinaryCompare) = 1 Then
            vbpFiles(numOfFiles) = m_VBPPath & Right$(vbpText(i), Len(vbpText(i)) - 5)
            numOfFiles = numOfFiles + 1
        End If
        
        'Check for user controls
        If InStr(1, vbpText(i), "UserControl=", vbBinaryCompare) = 1 Then
            vbpFiles(numOfFiles) = m_VBPPath & Right$(vbpText(i), Len(vbpText(i)) - 12)
            numOfFiles = numOfFiles + 1
        End If
        
        'Check for modules
        If InStr(1, vbpText(i), "Module=", vbBinaryCompare) = 1 Then
            vbpFiles(numOfFiles) = m_VBPPath & Trim$(Right$(vbpText(i), Len(vbpText(i)) - InStr(1, vbpText(i), ";")))
            numOfFiles = numOfFiles + 1
        End If
        
        'Check for classes
        If InStr(1, vbpText(i), "Class=", vbBinaryCompare) = 1 Then
            vbpFiles(numOfFiles) = m_VBPPath & Trim$(Right$(vbpText(i), Len(vbpText(i)) - InStr(1, vbpText(i), ";")))
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
    
    ReDim Preserve vbpFiles(0 To numOfFiles) As String
    
    'To make sure everything worked, dump the contents of the array into the list box on the left
    lstProjectFiles.Clear
    
    For i = 0 To UBound(vbpFiles)
        If Len(vbpFiles(i)) > 0 Then lstProjectFiles.AddItem vbpFiles(i)
    Next i
    
    'Build a complete version string
    m_VersionString = majorVer & "." & minorVer & "." & buildVer
    
    cmdProcess.Caption = "Begin processing"

End Sub

'Given a full file name, remove everything but the directory structure
Private Function GetDirectory(ByRef sString As String) As String
    
    Dim x As Long
    
    For x = Len(sString) - 1 To 1 Step -1
        If (Mid$(sString, x, 1) = "/") Or (Mid$(sString, x, 1) = "\") Then
            GetDirectory = Left$(sString, x)
            Exit Function
        End If
    Next x
    
End Function

'Count the number of words in a string (will not be 100% accurate, but that's okay)
Private Function CountWordsInString(ByVal srcString As String) As Long

    If (LenB(Trim$(srcString)) <> 0) Then

        Dim tmpArray() As String
        tmpArray = Split(Trim$(srcString), " ")
        
        Dim tmpWordCount As Long
        tmpWordCount = 0
        
        Dim i As Long
        For i = 0 To UBound(tmpArray)
            If IsAlpha(tmpArray(i)) Then tmpWordCount = tmpWordCount + 1
        Next i
        
        CountWordsInString = tmpWordCount
        
    Else
        CountWordsInString = 0
    End If

End Function

'VB's IsNumeric function can't detect percentage text (e.g. "100%").  PhotoDemon includes text like this,
' but I don't want that text translated - so manually check for and reject it.
Private Function IsNumericPercentage(ByVal srcString As String) As Boolean

    srcString = Trim$(srcString)

    'Start by checking for a percent in the right-most position
    If (Right$(srcString, 1) = "%") Then
        
        'If a percent was found, check the rest of the text to see if it's numeric
        IsNumericPercentage = IsNumeric(Left$(srcString, Len(srcString) - 1))
        
    Else
        IsNumericPercentage = False
    End If

End Function

'URLs shouldn't be translated.  Check for them and reject as necessary.
Private Function IsURL(ByRef srcString As String) As Boolean
    IsURL = (Left$(srcString, 6) = "ftp") Or (Left$(srcString, 7) = "http")
End Function

Private Sub Form_Load()
    
    Set m_XML = New pdXML
        
    'Build a blacklist of phrases that are in the software, but do not need to be translated.  (These are complex phrases that
    ' may include things like names, but the automatic text generator has no way of knowing that the text is non-translatable.)
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
        
        'Update the master langupdate.XML file, and generate new compressed language copies in their
        ' dedicated upload folders
        'NOTE: as of 23 October 2017 (just prior to 7.0's release), this feature has been disabled.
        ' PD no longer attempts to patch language files separately, which greatly simplifies the core
        ' program's update code and network access requirements.
        'Call cmdLangVersions_Click
        
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
        IsAlpha = (UCase$(srcString) = "A") Or (UCase$(srcString) = "I")
    Else
        
        Dim charID As Long
        
        Dim numAlphaChars As Long
        numAlphaChars = 0
        
        Dim i As Long
        For i = 1 To Len(srcString)
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
