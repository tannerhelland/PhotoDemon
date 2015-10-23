VERSION 5.00
Begin VB.Form frmCreateMaster 
   BackColor       =   &H80000005&
   Caption         =   " PhotoDemon Master Language XML Generator"
   ClientHeight    =   8100
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
   ScaleHeight     =   540
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   969
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLangVersions 
      Caption         =   "Generate master language update file(s)"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   16
      Top             =   7320
      Width           =   3975
   End
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
      Caption         =   "2a (Optional) Automatically merge all language files with newest Master XML file..."
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
      Caption         =   "3) Merge the files into an updated non-English XML file (NOTE: this will not modify the source files)"
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
      Caption         =   "2) Select old non-English XML file..."
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
      Caption         =   "1) Select master English XML file..."
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
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      Index           =   1
      X1              =   304
      X2              =   304
      Y1              =   480
      Y2              =   536
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      Index           =   0
      X1              =   8
      X2              =   304
      Y1              =   480
      Y2              =   480
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
      Caption         =   $"frmCreateMaster.frx":0000
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
      Y1              =   336
      Y2              =   336
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
Attribute VB_Name = "frmCreateMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Master English Language File (XML) Generator
'Copyright ©2013-2015 by Tanner Helland
'Created: 23/January/13
'Last updated: 27/January/15
'Last update: new support functions for generating PD's master language version file
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

'If silent mode has been activated via command line, this will be set to TRUE.
Dim m_SilentMode As Boolean

'A pdXML instance provides UTF-8 support.
Private m_XML As pdXML

'PD language file identifier.  IMPORTANT NOTE: this constant is shared with the main PhotoDemon project.  DO NOT CHANGE IT!
Private Const PD_LANG_IDENTIFIER As Long = &H414C4450    'pdLanguage data (ASCII characters "PDLA", as hex, little-endian)

'If duplicates are assigned for removal, this flag is set to TRUE
Private m_RemoveDuplicates As Boolean

'New support function for auto-converting old common control labels to PD's new pdLabel object.  If successful, this will save me a ton of time
' manually converting all the labels in the program.
Private Sub cmdConvertLabels_Click()

    'Make sure a file has been selected
    If lstProjectFiles.ListIndex <> -1 Then

        'Read the file into a string array
        Dim srcFilename As String
        srcFilename = lstProjectFiles.List(lstProjectFiles.ListIndex)
        
        Dim fileContents As String
        fileContents = getFileAsString(srcFilename)
        
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
                    ignoreThisLine = isValidPDLabelProperty(fileLines(i))
                    
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
        If FileExist(srcFilename) Then Kill srcFilename
        
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
Private Function isValidPDLabelProperty(ByVal srcString As String) As Boolean
    
    isValidPDLabelProperty = False
    
    'Trim the source string to make comparisons easier
    srcString = Trim$(LCase$(srcString))
    
    'The list of valid properties is hardcoded.
    If InStr(1, srcString, "Alignment", vbBinaryCompare) > 0 Then isValidPDLabelProperty = True
    If InStr(1, srcString, "BackColor", vbBinaryCompare) > 0 Then isValidPDLabelProperty = True
    If InStr(1, srcString, "Caption", vbBinaryCompare) > 0 Then isValidPDLabelProperty = True
    If InStr(1, srcString, "Enabled", vbBinaryCompare) > 0 Then isValidPDLabelProperty = True
    If InStr(1, srcString, "ForeColor", vbBinaryCompare) > 0 Then isValidPDLabelProperty = True
    If InStr(1, srcString, "Height", vbBinaryCompare) > 0 Then isValidPDLabelProperty = True
    If InStr(1, srcString, "Index", vbBinaryCompare) > 0 Then isValidPDLabelProperty = True
    If InStr(1, srcString, "Layout", vbBinaryCompare) > 0 Then isValidPDLabelProperty = True
    If InStr(1, srcString, "Left", vbBinaryCompare) > 0 Then isValidPDLabelProperty = True
    If InStr(1, srcString, "Top", vbBinaryCompare) > 0 Then isValidPDLabelProperty = True
    If InStr(1, srcString, "Width", vbBinaryCompare) > 0 Then isValidPDLabelProperty = True
    
End Function

'This function scans all of PD's current language files, and generates a small XML file with their version numbers.
' (Note that two folders are scanned: the standard /App/PhotoDemon/Languages folder, which contains dev build values, and a separate
'  stable folder, which contains the latest stable build language files.)
'
'It also fills two temporary folders (one stable, one dev) with pdPackaged copies of the latest PD language files.  PD's nightly build script
' will then upload these files to photodemon.org, so individual PD instances can self-patch according to the user's preferences.
'
'Note that this function can be automatically run by specifying -s on the command line.  If -s is used, this function will close the program
' upon completion.
Private Sub cmdLangVersions_Click()
    
    Dim numOfLangFiles As Long
    numOfLangFiles = 0
    
    'Two folders must be iterated for existing language files: a stable language folder, and an unstable (development) language folder.
    ' We'll start with the development folder.
    Dim srcFolder As String
    srcFolder = "C:\PhotoDemon v4\PhotoDemon\App\PhotoDemon\Languages\"
    
    'Two folders are also required for exporting the compressed language file copies (again, stable and dev).
    Dim exportFolderDev As String, exportFolderStable As String
    exportFolderDev = "C:\PhotoDemon v4\PhotoDemon\no_sync\PD_Language_File_Tmp\dev\"
    exportFolderStable = "C:\PhotoDemon v4\PhotoDemon\no_sync\PD_Language_File_Tmp\stable\"
    
    'Lots of XML parsing will be going on here.
    Dim xmlInput As pdXML, xmlOutput As pdXML
    Set xmlInput = New pdXML
    Set xmlOutput = New pdXML
    
    'Prep xmlOutput in advance
    xmlOutput.prepareNewXML "Language versions"
    xmlOutput.writeBlankLine
    xmlOutput.writeComment "This language version file was automatically generated on " & Format(Now, "Medium date")
    xmlOutput.writeBlankLine
    
    'The trusty pdPackage class will be used to compress each language file as we process it.
    Dim cPackager As pdPackager
    Set cPackager = New pdPackager
    cPackager.init_ZLib App.Path & "\zlibwapi.dll"
    
    Dim compressedFilename As String
    
    'Iterate through every language file in this folder.  (Language files will always be XML format, so we can ignore anything
    ' that isn't XML.)
    Dim chkFile As String, chkFileNoExtension As String
    chkFile = Dir(srcFolder & "*.xml", vbNormal)
    
    Do While (chkFile <> "")
        
        numOfLangFiles = numOfLangFiles + 1
        lblUpdates.Caption = "Processing language file #" & numOfLangFiles
        
        'Attempt to add this file to the version list
        chkFileNoExtension = chkFile
        StripOffExtension chkFileNoExtension
        addFileToMasterVersionList xmlInput, xmlOutput, srcFolder & chkFile, chkFileNoExtension, False
        
        'This program is also responsible for compressing each language file and copying it to a temp folder,
        ' so the nightly build batch script can find it.
        compressedFilename = chkFileNoExtension & ".pdz"
        cPackager.prepareNewPackage 1, PD_LANG_IDENTIFIER
        cPackager.autoAddNodeFromFile srcFolder & chkFile
        cPackager.writePackageToFile exportFolderDev & compressedFilename
        
        'Retrieve the next file and repeat
        chkFile = Dir
        
    Loop
    
    'The MASTER language file is handled separately, on account of being held in a separate location.
        numOfLangFiles = numOfLangFiles + 1
        lblUpdates.Caption = "Processing language file #" & numOfLangFiles
        
        'Attempt to add this file to the version list
        chkFile = "Master\MASTER.xml"
        chkFileNoExtension = "MASTER"
        addFileToMasterVersionList xmlInput, xmlOutput, srcFolder & chkFile, chkFileNoExtension, False
        
        'This program is also responsible for compressing each language file and copying it to a temp folder,
        ' so the nightly build batch script can find it.
        compressedFilename = chkFileNoExtension & ".pdz"
        cPackager.prepareNewPackage 1, PD_LANG_IDENTIFIER
        cPackager.autoAddNodeFromFile srcFolder & chkFile
        cPackager.writePackageToFile exportFolderDev & compressedFilename
        
    
    'We are now going to repeat the above process, but for a separate folder of stable version language files.
    srcFolder = "C:\PhotoDemon v4\PhotoDemon\Support\Master language XML generator\stable build language files\"
    chkFile = Dir(srcFolder & "*.xml", vbNormal)
    
    Do While (chkFile <> "")
        
        numOfLangFiles = numOfLangFiles + 1
        lblUpdates.Caption = "Processing language file #" & numOfLangFiles
        
        'Attempt to add this file to the version list
        chkFileNoExtension = chkFile
        StripOffExtension chkFileNoExtension
        addFileToMasterVersionList xmlInput, xmlOutput, srcFolder & chkFile, chkFileNoExtension, True
        
        'This program is also responsible for compressing each language file and copying it to a temp folder,
        ' so the nightly build batch script can find it.
        compressedFilename = chkFileNoExtension & ".pdz"
        cPackager.prepareNewPackage 1, PD_LANG_IDENTIFIER
        cPackager.autoAddNodeFromFile srcFolder & chkFile
        cPackager.writePackageToFile exportFolderStable & compressedFilename
        
        'Retrieve the next file and repeat
        chkFile = Dir
        
    Loop
    
    'Once again, the MASTER language file is handled separately, on account of being held in a separate location.
        numOfLangFiles = numOfLangFiles + 1
        lblUpdates.Caption = "Processing language file #" & numOfLangFiles
        
        'Attempt to add this file to the version list
        chkFile = "Master\MASTER.xml"
        chkFileNoExtension = "MASTER"
        addFileToMasterVersionList xmlInput, xmlOutput, srcFolder & chkFile, chkFileNoExtension, True
        
        'This program is also responsible for compressing each language file and copying it to a temp folder,
        ' so the nightly build batch script can find it.
        compressedFilename = chkFileNoExtension & ".pdz"
        cPackager.prepareNewPackage 1, PD_LANG_IDENTIFIER
        cPackager.autoAddNodeFromFile srcFolder & chkFile
        cPackager.writePackageToFile exportFolderStable & compressedFilename
        
    
    'The master language version file is now complete.  Write it.
    Dim dstFile As String
    dstFile = "C:\PhotoDemon v4\langupdate.xml"
    
    xmlOutput.writeXMLToFile dstFile
    
    lblUpdates.Caption = numOfLangFiles & " languages successfully added to master language file."
    lblUpdates.Refresh
    DoEvents
    
    'If the program is running in silent mode, unload it now
    If m_SilentMode Then Unload Me

End Sub

'Given a full path to a language file, add the language file's information to an output XML object.
Private Sub addFileToMasterVersionList(ByRef xmlInput As pdXML, ByRef xmlOutput As pdXML, ByRef pathToFile As String, ByRef sourceFilename As String, ByVal isSourceStableVersion As Boolean)

    Dim langID As String, langVersion As String, langName As String
    
    Dim versionMajor As String, versionMinor As String, versionRevision As String
    Dim versionCheck() As String
    
    'A pdPackage class provides a convenient way to checksum files
    Dim cPackage As pdPackager
    Set cPackage = New pdPackager
    
    'Load the file into an XML parser
    If xmlInput.loadXMLFile(pathToFile) Then
    
        'Check the file for two things: the language identifier, and the current version number.
        xmlInput.setTextCompareMode vbTextCompare
        langID = xmlInput.getUniqueTag_String("langid")
        langVersion = xmlInput.getUniqueTag_String("langversion")
        langName = xmlInput.getUniqueTag_String("langname")
        
        'Make sure both returns are valid.  If they are not, skip this file.
        If (Len(langID) <> 0) And (Len(langVersion) <> 0) Then
        
            'This file is valid.  To make it easier to check for updates in the core PD program, we're going to modify the information a bit.
            ' PD's update code works by comparing the current program version (e.g. 6.4) to the listed versions of each language file.
            ' (Version comparisons are necessary because stable and nightly builds require unique language files.)
            '
            'Because language files are updated independent of the program itself, revision numbers between the software and language files
            ' are unlikely to match, and that's okay - we only care about the *latest* revision of each language file, that matches the base
            ' version of the current build.
            '
            'So what does that mean for us?  It means we need to separate the language version into two distinct parts:
            ' 1) A base version (e.g. "6.4")
            ' 2) A revision number (e.g. the "1" in "6.4.1")
            
            'Start by replacing comma delimiters, if present
            If InStr(1, langVersion, ",") Then langVersion = Replace$(langVersion, ",", ".")
            
            'Split the version into its component parts
            versionCheck = Split(langVersion, ".")
            
            'Make sure the version contains at least a major and minor value
            If UBound(versionCheck) >= 1 Then
                
                versionMajor = versionCheck(0)
                versionMinor = versionCheck(1)
                
                'If no revision is given, assume a revision of 0
                If UBound(versionCheck) > 1 Then
                    versionRevision = versionCheck(2)
                Else
                    versionRevision = "0"
                End If
                
                'We now have a major, minor, and revision value for this language file.  Write them out to file.
                If isSourceStableVersion Then
                    xmlOutput.writeTagWithAttribute "language", "updateID", langID & " stable", "", True
                Else
                    xmlOutput.writeTagWithAttribute "language", "updateID", langID & " dev", "", True
                End If
                
                xmlOutput.writeTag "name", langName
                xmlOutput.writeTag "id", langID
                xmlOutput.writeTag "filename", sourceFilename
                xmlOutput.writeTag "version", versionMajor & "." & versionMinor
                xmlOutput.writeTag "revision", versionRevision
                xmlOutput.writeTag "checksum", cPackage.checkSumArbitraryFile(pathToFile)
                If isSourceStableVersion Then
                    xmlOutput.writeTag "location", "stable"
                Else
                    xmlOutput.writeTag "location", "dev"
                End If
                xmlOutput.closeTag "language"
                xmlOutput.writeBlankLine
            
            End If
            
        End If
    
    End If

End Sub

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

Private Sub replaceTopLevelTag(ByVal origTagName As String, ByRef sourceTextMaster As String, ByRef sourceTextTranslation As String, ByRef destinationText As String, Optional ByVal alsoIncrementVersion As Boolean = True)

    Dim openTagName As String, closeTagName As String
    openTagName = "<" & origTagName & ">"
    closeTagName = "</" & origTagName & ">"
    
    Dim findText As String, replaceText As String
    findText = openTagName & getTextBetweenTags(sourceTextMaster, origTagName) & closeTagName
    
    'A special check is applied to the "langversion" tag.  Whenever this function is used, a merge is taking place; as such, we want to
    ' auto-increment the language's version number to trigger an update on client machines.
    If (StrComp(origTagName, "langversion", vbBinaryCompare) = 0) And alsoIncrementVersion Then
        
        findText = openTagName & getTextBetweenTags(sourceTextTranslation, origTagName) & closeTagName
        
        'Retrieve the current language version
        Dim curVersion As String
        curVersion = getTextBetweenTags(sourceTextTranslation, origTagName)
        
        'Parse the current version into two discrete chunks: the major/minor value, and the revision value
        Dim curMajorMinor As String, curRevision As Long
        curMajorMinor = retrieveVersionMajorMinorAsString(curVersion)
        curRevision = retrieveVersionRevisionAsLong(curVersion)
        
        'Increment the revision value by 1, then assemble the modified replacement text
        curRevision = curRevision + 1
        replaceText = openTagName & curMajorMinor & "." & Trim$(Str$(curRevision)) & closeTagName
            
    Else
        replaceText = openTagName & getTextBetweenTags(sourceTextTranslation, origTagName) & closeTagName
    End If
    
    destinationText = Replace$(destinationText, findText, replaceText)

End Sub

Private Sub cmdMerge_Click()

    'Make sure our source file strings are not empty
    If Len(m_MasterText) = 0 Or Len(m_OldLanguageText) = 0 Then
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
        If Len(translatedText) <> 0 Then
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
        
        'Use pdXML to write out a UTF-8 encoded XML file
        m_XML.loadXMLFromString m_NewLanguageText
        m_XML.writeXMLToFile m_OldLanguagePath, True
        
    End If
    
    MsgBox "Merge complete." & vbCrLf & vbCrLf & phrasesProcessed & " phrases processed. " & phrasesFound & " translations found. " & phrasesMissed & " translations missing."

End Sub

'Given a string, return the location of the <phrase> tag enclosing said string
Private Function getPhraseTagLocation(ByRef srcString As String) As Long
    
    Dim sLocation As Long
    sLocation = InStr(1, m_OldLanguageText, srcString, vbBinaryCompare)
    
    'If the source string was found, work backward to find the phrase tag location
    If sLocation > 0 Then
        sLocation = InStrRev(m_OldLanguageText, "<phrase>", sLocation, vbBinaryCompare)
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
    tagStart = InStr(searchLocation, fileText, "<" & fTag & ">", vbBinaryCompare)

    'If the tag was found in the file, we also need to find the closing tag.
    If tagStart > 0 Then
    
        tagEnd = InStr(tagStart, fileText, "</" & fTag & ">", vbBinaryCompare)
        
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
    
    'Rather than backup the old files to the dev language folder (which is confusing), I now place them inside a dedicated backup folder.
    Dim backupFolder As String
    backupFolder = "C:\PhotoDemon v4\PhotoDemon\no_sync\PD_Language_File_Tmp\dev_backup\"
    
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
            If Len(m_MasterText) = 0 Or Len(m_OldLanguageText) = 0 Then
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
            replaceTopLevelTag "langstatus", m_MasterText, m_OldLanguageText, m_NewLanguageText
            replaceTopLevelTag "author", m_MasterText, m_OldLanguageText, m_NewLanguageText
            replaceTopLevelTag "langversion", m_MasterText, m_OldLanguageText, m_NewLanguageText, False
                
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
                
                'Remove any tab stops from the translated text (which may have been added by an outside editor)
                If InStr(translatedText, vbTab) <> 0 Then translatedText = Replace(translatedText, vbTab, "", , , vbBinaryCompare)
                
                'If a translation was found, insert it into the new file
                If Len(translatedText) <> 0 Then
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
            
            'Find where the two files don't match
            'Dim i As Long
            'For i = 1 To Len(Trim$(m_NewLanguageText))
            '    If StrComp(Mid$(m_NewLanguageText, i, 1), Mid$(m_OldLanguageText, i, 1), vbBinaryCompare) <> 0 Then
            '        MsgBox i & vbCrLf & Mid$(m_NewLanguageText, i, 10) & vbCrLf & Mid$(m_OldLanguageText, i, 10)
            '    End If
            'Next i
            '
            'MsgBox Len(Trim$(m_NewLanguageText)) & vbCrLf & Len(Trim$(m_OldLanguageText))
            
            'See if the old and new language files are equal.  If they are, we won't bother writing the results out to file.
            If Len(Trim$(m_NewLanguageText)) = Len(Trim$(m_OldLanguageText)) Then
                Debug.Print "New language file and old language file are identical for " & chkFile & ".  Merge abandoned."
            Else
                
                'Update the version number by 1
                replaceTopLevelTag "langversion", m_MasterText, m_OldLanguageText, m_NewLanguageText
                
                'Unlike the normal merge option, we will automatically save the results to a new XML file
                
                'Start by backing up the old file
                FileCopy m_OldLanguagePath, backupFolder & chkFile
                
                If FileExist(m_OldLanguagePath) Then
                    Debug.Print "Note - old file with same name (" & m_OldLanguagePath & ") was erased.  Hope this is what you wanted!"
                End If
                
                'Use pdXML to write out a UTF-8 encoded XML file
                m_XML.loadXMLFromString m_NewLanguageText
                m_XML.writeXMLToFile m_OldLanguagePath, True
                
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
    
    'Note whether duplicate phrases are automatically removed
    m_RemoveDuplicates = CBool(chkRemoveDuplicates)
    
    'Start by preparing the XML header
    outputText = "<?xml version=""1.0"" encoding=""UTF-8""?>"
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
    
    'We are now going to compare the length of the old file and new file.  If the lengths match, there's no reason to write out this new file.
    Dim oldFileString As String
    oldFileString = getFileAsString(outputFile)
    
    Dim newFileLen As Long, oldFileLen As Long
    
    newFileLen = Len(Trim$(Replace$(Replace$(outputText, vbCrLf, ""), vbTab, "")))
    oldFileLen = Len(Trim$(Replace$(Replace$(oldFileString, vbCrLf, ""), vbTab, "")))
        
    If newFileLen <> oldFileLen Then
        
        'Use pdXML to write a UTF-8 encoded text file
        m_XML.loadXMLFromString outputText
        m_XML.writeXMLToFile outputFile, True
        
        cmdProcess.Caption = "Processing complete!"
        
    Else
        cmdProcess.Caption = "Processing complete (no changes made)"
    End If
    
End Sub

'Given a VB file (form, module, class, user control), extract any relevant text from it
Private Sub processFile(ByVal srcFile As String)

    If Len(srcFile) = 0 Then Exit Sub

    m_FileName = getFilename(srcFile)
    
    'Certain files can be ignored.  I generate this list manually, on account of knowing which files (classes, mostly) contain
    ' no special text.  I could probably add many more files to this list, but I primarily want to blacklist those that create
    ' parsing problems.  (The tooltip classes are particularly bad, since they use the phrase "tooltip" frequently, which messes
    ' up the parser as it thinks it's found hundreds of tooltips in each file.)
    Select Case m_FileName
    
        Case "clsToolTip.cls", "pdToolTip.cls", "clsControlImages.cls"
            Exit Sub
            
        Case "pdFilterSupport.cls", "cSelfSubHookCallback.cls", "jcButton.ctl"
            Exit Sub
            
        Case "VBP_PublicVariables.bas", "pdParamString.cls", "VBP_ToolbarDebug.frm"
            Exit Sub
            
        Case "buttonStrip.ctl", "buttonStripVertical.ctl"
            Exit Sub
            
        'TEMPORARILY: I am disabling the Theme Editor, as it only contains debug text at present.
        ' When PD has an actual theme editor, the form *will* need to be translated.
        Case "Tools_ThemeEditor.frm"
            Exit Sub
    
    End Select
            
    
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
    
    If Right$(m_FileName, 3) = "frm" Then
        Dim findName() As String
        findName = Split(fileLines(1), " ")
        shortcutName = findName(2)
    End If
    
    'For convenience, write the name of the source file into the translation file - this can be helpful when
    ' tracking down errors or incomplete text.
    If LenB(m_FileName) > 0 Then
        outputText = outputText & vbCrLf & vbCrLf & vbTab & vbTab
        If Len(shortcutName) <> 0 Then
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
    
    Dim toolTipSecondCheckNeeded As Boolean
        
    'Now, start processing the file one line at a time, searching for relevant text as we go
    Do
    
        processedText = ""
        processedTextSecondary = ""
    
        curLineText = fileLines(curLineNumber)
        
        'Before processing this line, make sure is isn't a comment.  (Comments are always ignored.)
        If Left$(Trim$(curLineText), 1) = "'" Then GoTo nextLine
        
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
        If InStr(1, UCase$(curLineText), "BEGIN VB.FORM", vbBinaryCompare) Then
            processedText = findFormCaption(fileLines, curLineNumber)
                
        '2) Check for a control caption.  (This has to be handled slightly differently than form caption.)
        ElseIf ((InStr(1, UCase$(curLineText), "BEGIN VB.", vbBinaryCompare) > 0) Or (InStr(1, UCase$(curLineText), "BEGIN PHOTODEMON.", vbBinaryCompare) > 0)) And (InStr(1, UCase$(curLineText), "PICTUREBOX", vbBinaryCompare) = 0) And (InStr(1, curLineText, "ComboBox") = 0) And (InStr(1, curLineText, ".Shape") = 0) And (InStr(1, curLineText, "TextBox") = 0) And (InStr(1, curLineText, "HScrollBar") = 0) And (InStr(1, curLineText, "VScrollBar") = 0) Then
            processedText = findControlCaption(fileLines, curLineNumber)
        
        '3) Check for tooltip text (several varations of this exist due to custom controls having unique tooltip property names)
        ElseIf InStr(1, UCase$(curLineText), "TOOLTIPTEXT", vbBinaryCompare) And (InStr(1, UCase$(curLineText), ".TOOLTIPTEXT", vbBinaryCompare) = 0) Then
            processedText = findCaptionInComplexQuotes(fileLines, curLineNumber, True)
                        
        ElseIf (InStr(1, UCase$(curLineText), "TOOLTIP", vbBinaryCompare) > 0) And (InStr(1, UCase$(curLineText), ".TOOLTIP", vbBinaryCompare) = 0) And (InStr(1, UCase$(curLineText), "TOOLTIPTITLE", vbBinaryCompare) = 0) And (InStr(1, UCase$(curLineText), "TOOLTIPTEXT", vbBinaryCompare) = 0) Then
            
            'Tooltips represent a complicated situation in PD.  They can appear in several forms, such as being set via the standard property dialog,
            ' or being manually assigned to a custom pdToolTip object.  Because the term "tooltip" appears so frequently, I have to go to rather elaborate
            ' lengths to make sure only valid tooltip text is parsed, and not false-positive lines that simply happen to contain the word "tooltip" in them.
            
            'The massive chunk of text below is designed to address this problem, when checking for tooltips set via VB's property window.
            
            '3a) Check for tooltip text embedded as a VB property
            If (InStr(1, UCase$(curLineText), "TOOLTIPBACKCOLOR", vbBinaryCompare) = 0) And (InStr(1, UCase$(curLineText), "TOOLTIPTYPE", vbBinaryCompare) = 0) And (InStr(1, UCase$(curLineText), "M_TOOLTIP", vbBinaryCompare) = 0) And (InStr(1, UCase$(curLineText), "CLSTOOLTIP", vbBinaryCompare) = 0) Then
            If (Not m_FileName = "jcButton.ctl") And (InStr(1, curLineText, "=") > 0) And (InStr(1, curLineText, "PD_MAX_TOOLTIP_WIDTH") = 0) And (InStr(1, UCase$(curLineText), "DELAYTIME", vbBinaryCompare) = 0) And (InStr(1, UCase$(curLineText), "ECONTROL.TOOLTIPTEXT", vbBinaryCompare) = 0) Then
            If (InStr(1, UCase$(curLineText), "TOOLTIPBACKUP", vbBinaryCompare) = 0) And (InStr(1, UCase$(curLineText), "NEWTOOLTIP", vbBinaryCompare) = 0) And (InStr(1, UCase$(curLineText), "SETTHUMBNAILTOOLTIP", vbBinaryCompare) = 0) And (InStr(1, UCase$(curLineText), "TOOLTIPMANAGER", vbBinaryCompare) = 0) Then
            If (InStr(1, UCase$(curLineText), "M_PREVIOUSTOOLTIP", vbBinaryCompare) = 0) And (InStr(1, UCase$(curLineText), "ASSIGNTOOLTIP", vbBinaryCompare) = 0) And (InStr(1, UCase$(curLineText), "SETTOOLTIP", vbBinaryCompare) = 0) And (InStr(1, UCase$(curLineText), "PDTOOLTIP", vbBinaryCompare) = 0) Then
                processedText = findCaptionInComplexQuotes(fileLines, curLineNumber, True)
                If InStr(1, processedText, "MANUAL FIX REQUIRED") Then Debug.Print "Tooltip error occurred on line " & curLineNumber & " of m_filename"
            End If
            End If
            End If
            End If
            
            'In current builds, the more likely place for tooltip text is assignment via a pdToolTip object.  These are much simpler to detect, as they
            ' will rely exclusively on an .assignToolTip request.  Note, however, that we must search for two pieces of translated text: the tooltip text,
            ' and a potential title.
            
            '3b) Check for tooltip text that has been manually assigned to a custom PhotoDemon object.  Note that we (obviously) avoid .assignToolTip
            '     function declarations themselves.
            If (InStr(1, UCase$(curLineText), ".ASSIGNTOOLTIP ") > 0) And (InStr(1, curLineText, "ByVal") = 0) Then
                
                'Process the tooltip text itself
                processedText = findTooltipMessage(fileLines, curLineNumber, False, toolTipSecondCheckNeeded)
                
                'Process the title, if any
                If toolTipSecondCheckNeeded Then processedTextSecondary = findMsgBoxTitle(fileLines, curLineNumber)
            
            End If
            
            '3b) Check for tooltip text that has been manually assigned to a PhotoDemon pdToolTip object.  Note that we (obviously) avoid
            '     .setToolTip function declarations themselves.
            If (InStr(1, UCase$(curLineText), ".SETTOOLTIP") > 0) And (InStr(1, curLineText, "ByVal") = 0) Then
                
                'Process the tooltip text itself
                processedText = findTooltipMessage(fileLines, curLineNumber, False, toolTipSecondCheckNeeded)
                
                'Process the title, if any
                If toolTipSecondCheckNeeded Then processedTextSecondary = findMsgBoxTitle(fileLines, curLineNumber)
            
            End If
            
            
        
        ElseIf InStr(1, UCase$(curLineText), "TOOLTIPTITLE", vbBinaryCompare) And (InStr(1, curLineText, ".TooltipTitle") = 0) And (InStr(1, UCase$(curLineText), "NEWTOOLTIPTITLE") = 0) And (Not m_FileName = "jcButton.ctl") Then
            processedText = findCaptionInComplexQuotes(fileLines, curLineNumber, True)
        
        '4) Check for text added to a combo box or list box control at run-time
        ElseIf InStr(1, curLineText, ".AddItem """) <> 0 Then
            processedText = findCaptionInComplexQuotes(fileLines, curLineNumber)
            
        '5) Check for message calls
        ElseIf InStr(1, curLineText, "Message """) <> 0 Then
            processedText = findCaptionInQuotes(fileLines, curLineNumber)
        
        '6) Check for message box text, including 7) message box titles (which must also be translated)
        ElseIf (InStr(1, UCase$(curLineText), "PDMSGBOX", vbTextCompare) <> 0) Then
        
            'First, retrieve the message box text itself
            processedText = findMsgBoxText(fileLines, curLineNumber)
            
            'Next, retrieve the message box title
            processedTextSecondary = findMsgBoxTitle(fileLines, curLineNumber)
        
        '7) Specific to PhotoDemon - check for action names that may not be present elsewhere
        ElseIf InStr(1, curLineText, "Process """) <> 0 Then
            processedText = findCaptionInQuotes(fileLines, curLineNumber, InStr(1, curLineText, "Process """))
        
        '7.5) Now that pdLabel objects manage their own translations, we should also check for caption assignments
        ElseIf InStr(1, curLineText, "Caption = """, vbBinaryCompare) <> 0 Then
            processedText = findCaptionInQuotes(fileLines, curLineNumber, 1)
        
        End If
        
        '8) Check for text that has been manually marked for translation (e.g. g_Language.TranslateMessage("xyz"))
        '    NOTE: as of 07 June 2013, each line can contain two translation calls (instead of just one)
        '
        'Note that this check is performed regardless of previous checks, to make sure no translations are missed.
        If InStr(1, curLineText, "g_Language.TranslateMessage(""") Then
            processedText = findMessage(fileLines, curLineNumber)
            processedTextSecondary = findMessage(fileLines, curLineNumber, True)
        End If
        
        'DEBUG! Check for certain text entries here
        'If (shortcutName = "FormLens") And Len(Trim$(processedText)) <> 0 Then MsgBox processedText
        
        'We now have text in potentially two places: processedText, and processedTextSecondary (for message box titles)
        chkText = Trim$(processedText)
        
        'Only pass the text along if it isn't blank, or a number, or a symbol, or a manually blacklisted phrase
        If Len(chkText) <> 0 Then
            If (Not IsNumeric(chkText)) And (Not IsNumericPercentage(chkText)) And (Not isBlacklisted(chkText)) Then
                If (chkText <> ".") And (chkText <> "-") And (Not IsURL(chkText)) Then
                    numOfPhrasesFound = numOfPhrasesFound + 1
                    If addPhrase(processedText) Then numOfPhrasesWritten = numOfPhrasesWritten + 1
                End If
            End If
        End If
        
        chkText = Trim$(processedTextSecondary)
        
        'Do the same for the secondary text
        If Len(chkText) <> 0 Then
            If (Not IsNumeric(chkText)) And (Not IsNumericPercentage(chkText)) And (Not isBlacklisted(chkText)) Then
                If (chkText <> ".") And (chkText <> "-") And (Not IsURL(chkText)) Then
                    numOfPhrasesFound = numOfPhrasesFound + 1
                    If addPhrase(processedTextSecondary) Then numOfPhrasesWritten = numOfPhrasesWritten + 1
                End If
            End If
        End If
    
nextLine:
        curLineNumber = curLineNumber + 1
    
    Loop While curLineNumber < UBound(fileLines)
    
    'Now that all phrases in this file have been processed, we can wrap up this section of XML
    
    'For fun, write some stats about our processing results into the translation file.
    If Len(m_FileName) <> 0 Then
        
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
    phraseText = Trim$(phraseText)
    
    '2) Check for a trailing "..." and remove it
    If Right$(phraseText, 3) = "..." Then phraseText = Left$(phraseText, Len(phraseText) - 3)
    
    '3) Check for a trailing colon ":" and remove it
    If Right$(phraseText, 1) = ":" Then phraseText = Left$(phraseText, Len(phraseText) - 1)
    
    'This phrase is now ready to write to file.
    
    'Before writing the phrase out, check to see if it already exists
    If m_RemoveDuplicates Then
                
        If InStr(1, outputText, "<original>" & phraseText & "</original>", vbBinaryCompare) > 0 Then
            addPhrase = False
        Else
            If Len(phraseText) <> 0 Then
                addPhrase = True
            Else
                addPhrase = False
            End If
        End If
        
    Else
        
        If Len(phraseText) <> 0 Then
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
    
        If Mid$(srcLines(lineNumber), i, 1) = """" Then insideQuotes = Not insideQuotes
        
        If ((Mid$(srcLines(lineNumber), i, 1) = ",") Or (Mid$(srcLines(lineNumber), i, 1) = ")")) And (Not insideQuotes) Then
            endQuote = i - 1
            Exit For
        End If
    
    Next i
    
    'If endQuote = -1, something went horribly wrong
    If endQuote = -1 Then
        findMessage = "MANUAL FIX REQUIRED FOR MESSAGE PARSE ERROR AT LINE # " & lineNumber & " IN " & m_FileName
    Else
        findMessage = Mid$(srcLines(lineNumber), startQuote + 1, endQuote - startQuote - 1)
    End If
    
    'We now need to replace line breaks in the text.  These can appear in a variety of ways.  Replace them all.
    Dim lineBreak As String
    lineBreak = """ & vbCrLf & """
    If InStr(1, findMessage, lineBreak) Then findMessage = Replace(findMessage, lineBreak, vbCrLf)
    lineBreak = """ & vbCrLf & vbCrLf & """
    If InStr(1, findMessage, lineBreak) Then findMessage = Replace(findMessage, lineBreak, vbCrLf & vbCrLf)

    
End Function

'Given a line number and the original file contents, search for a custom PhotoDemon tooltip assignment
Private Function findTooltipMessage(ByRef srcLines() As String, ByRef lineNumber As Long, Optional ByVal inReverse As Boolean = False, Optional ByRef isSecondarySearchNecessary As Boolean) As String
    
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
    If startQuote > 0 Then
        
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
            findTooltipMessage = ""
        Else
            findTooltipMessage = Mid$(srcLines(lineNumber), startQuote + 1, endQuote - startQuote - 1)
        End If
        
        'We now need to replace line breaks in the text.  These can appear in a variety of ways.  Replace them all.
        Dim lineBreak As String
        lineBreak = """ & vbCrLf & """
        If InStr(1, findTooltipMessage, lineBreak) Then findTooltipMessage = Replace(findTooltipMessage, lineBreak, vbCrLf)
        lineBreak = """ & vbCrLf & vbCrLf & """
        If InStr(1, findTooltipMessage, lineBreak) Then findTooltipMessage = Replace(findTooltipMessage, lineBreak, vbCrLf & vbCrLf)
    
    Else
        findTooltipMessage = ""
    End If
    
End Function

'Given a line number and the original file contents, search for a message box title
Private Function findMsgBoxTitle(ByRef srcLines() As String, ByRef lineNumber As Long) As String

    Dim endQuote As Long
    endQuote = InStrRev(srcLines(lineNumber), """", Len(srcLines(lineNumber)))
        
    Dim startQuote As Long
    startQuote = InStrRev(srcLines(lineNumber), """", endQuote - 1)
    
    If endQuote > 0 Then
        findMsgBoxTitle = Mid$(srcLines(lineNumber), startQuote + 1, endQuote - startQuote - 1)
    Else
        findMsgBoxTitle = ""
    End If

End Function

'Given a line number and the original file contents, search for message box text
Private Function findMsgBoxText(ByRef srcLines() As String, ByRef lineNumber As Long) As String

    'Before processing this message box, make sure that the text contains actual text and not just a reference to a string.
    ' If all it contains is a reference to a string variable, don't process it.
    If InStr(1, srcLines(lineNumber), "pdMsgBox(""", vbTextCompare) = 0 And InStr(1, srcLines(lineNumber), "pdMsgBox """, vbTextCompare) = 0 Then
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
    
        If Mid$(srcLines(lineNumber), i, 1) = """" Then insideQuotes = Not insideQuotes
        
        If (Mid$(srcLines(lineNumber), i, 1) = ",") And Not insideQuotes Then
            endQuote = i - 1
            Exit For
        End If
    
    Next i
    
    'If endQuote = -1, something went horribly wrong
    If endQuote = -1 Then
        findMsgBoxText = "MANUAL FIX REQUIRED FOR MSGBOX PARSE ERROR AT LINE # " & lineNumber & " IN " & m_FileName
    Else
        findMsgBoxText = Mid$(srcLines(lineNumber), startQuote + 1, endQuote - startQuote - 1)
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
            findCaptionInComplexQuotes = "MANUAL FIX REQUIRED FOR TOOLTIP (FRX REFERENCE) OF " & m_ObjectName & " IN " & m_FormName
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
        findCaptionInComplexQuotes = Mid$(srcLines(lineNumber), startQuote + 1, endQuote - startQuote - 1)
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
        findCaptionInQuotes = Mid$(srcLines(lineNumber), startQuote + 1, endQuote - startQuote - 1)
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
            findControlCaption = "MANUAL FIX REQUIRED FOR " & m_ObjectName & " IN " & m_FormName
        Else
        
            Dim startQuote As Long
            startQuote = InStr(1, srcLines(lineNumber), """")
    
            Dim endQuote As Long
            endQuote = InStrRev(srcLines(lineNumber), """")
        
            If endQuote > 0 Then
                findControlCaption = Mid$(srcLines(lineNumber), startQuote + 1, endQuote - startQuote - 1)
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
        
        findFormCaption = Mid$(srcLines(lineNumber), startQuote + 1, endQuote - startQuote - 1)
        
    Else
        findFormCaption = ""
    End If

End Function


'Extract a list of all project files from a VBP file
Private Sub cmdSelectVBP_Click()

    Dim cDlg As cCommonDialog
    Set cDlg = New cCommonDialog
    
    m_VBPFile = "C:\PhotoDemon v4\PhotoDemon\PhotoDemon.vbp"
    lblVBP = "Active VBP: " & m_VBPFile
    m_VBPPath = getDirectory(m_VBPFile)
    
    'PD uses a hard-coded VBP location, but if you want to specify your own location, you can do so here
    'If cDlg.VBGetOpenFileName(m_VBPFile, , True, False, False, True, "VBP - Visual Basic Project|*.vbp", , , "Please select a Visual Basic project file (VBP)", "vbp", Me.hWnd) Then
    '    lblVBP = "Active VBP: " & m_VBPFile
    '    m_VBPPath = getDirectory(m_VBPFile)
    'Else
    '    Exit Sub
    'End If
    
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
    versionString = majorVer & "." & minorVer & "." & buildVer
    
    cmdProcess.Caption = "Begin processing"

End Sub

'Given a full file name, remove everything but the directory structure
Private Function getDirectory(ByVal sString As String) As String
    
    Dim x As Long
    
    For x = Len(sString) - 1 To 1 Step -1
        If (Mid$(sString, x, 1) = "/") Or (Mid$(sString, x, 1) = "\") Then
            getDirectory = Left(sString, x)
            Exit Function
        End If
    Next x
    
End Function

'Retrieve an entire file and return it as a string.  pdXML is used to support UTF-8 encodings (which PD's language files default to).
Private Function getFileAsString(ByVal fName As String) As String
           
    'Attempt to load the file as an XML object; if this fails, we'll assume it's not XML, and just load it as plain ol' ANSI text
    If m_XML.loadXMLFile(fName) Then
        getFileAsString = m_XML.returnCurrentXMLString(True)
        
    Else
        
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
            Debug.Print "File does not exist; exiting."
            getFileAsString = ""
        End If
            
    End If
    
End Function

'Count the number of words in a string (will not be 100% accurate, but that's okay)
Private Function countWordsInString(ByVal srcString As String) As Long

    If Len(Trim$(srcString)) <> 0 Then

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

    srcString = Trim$(srcString)

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
    
    Set m_XML = New pdXML
        
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
    addBlacklist "PNGQuant 2.1.1"
    addBlacklist "zLib 1.2.8"
    addBlacklist "EZTwain 1.18"
    addBlacklist "FreeImage 3.16.0"
    addBlacklist "ExifTool 9.62"
    addBlacklist "X.X"
    addBlacklist "XX.XX.XX"
    addBlacklist "PNGQuant"
    addBlacklist "zLib"
    addBlacklist "EZTwain"
    addBlacklist "FreeImage"
    addBlacklist "ExifTool"
    addBlacklist "tannerhelland.com/contact"
    addBlacklist "photodemon.org/about/contact"
    addBlacklist "photodemon.org/about/contact/"
    addBlacklist "HTML / CSS"
    addBlacklist "jcbutton"
    addBlacklist "while it downloads."
    addBlacklist "*"
    addBlacklist "("
    addBlacklist ")"
    addBlacklist ","
    
    'Check the command line.  This project can be run in silent mode as part of my nightly build batch script.
    Dim chkCommandLine As String
    chkCommandLine = Command$
    
    If Len(Trim$(chkCommandLine)) <> 0 Then
        If InStr(1, chkCommandLine, "-s", vbTextCompare) Then m_SilentMode = True Else m_SilentMode = False
    End If
    
    'If silent mode is activated, automatically "click" the relevant button
    If m_SilentMode Then
    
        'Load the current PhotoDemon VBP file
        Call cmdSelectVBP_Click
        
        'Generate a new master English file
        Call cmdProcess_Click
        
        'Forcibly merge all translation files with the latest English text
        Call cmdMergeAll_Click
        
        'Update the master langupdate.XML file, and generate new compressed language copies in their dedicated upload folders
        Call cmdLangVersions_Click
        
    End If
    
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

