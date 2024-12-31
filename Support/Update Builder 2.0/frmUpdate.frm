VERSION 5.00
Begin VB.Form frmUpdate 
   BackColor       =   &H80000005&
   Caption         =   "PhotoDemon Update Generator"
   ClientHeight    =   6915
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12630
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   461
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   842
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAction 
      Caption         =   "Calculate version, checksum, and release announcement details"
      Height          =   615
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   3240
      Width           =   12135
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Assemble stable and beta build packages (dedicated folders)"
      Height          =   615
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   12135
   End
   Begin VB.CommandButton cmdAction 
      Caption         =   "Assemble nightly build package (direct from current development folder)"
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   12135
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Step 3: build a versioning XML file, which PD will download first (to determine if an update is necessary)"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00DC7032&
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   11265
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Step 2: assemble the stable and beta update packages"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00DC7032&
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   5925
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Step 1: assemble the nightly build update package"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00DC7032&
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5505
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Program Auto-Update Generator
'Copyright 2015-2018 by Tanner Helland
'Created: 28/January/15
'Last updated: 15/June/18
'Last update: upgrade to v2.0, total overhaul to use pdPackager2 and zstd for compression.  Other fixes to improve
'             reliability, and start work on support for delta patches.
'
'This project was built to help assemble automatic update information for PhotoDemon.  It is run by the nightly build batch
' script, and it assembles a few different things:
' - New pdPackage archives for the latest stable, beta (if relevant), and nightly build entries.
' - A central update file with version numbers and checksums for each of the pdPackage files
'
'NOTE: this project is intended only as a support tool for PhotoDemon.  It is not designed or tested for general-purpose use.
'       I do not have any intention of supporting this tool outside its intended use, so please do not submit bug reports
'       regarding this project unless they directly relate to its intended purpose (generating PhotoDemon update files).
'
'       Also, given this project's purpose, the code is pretty ugly.  Organization is minimal.  Read at your own risk.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'PD update patch identifier.  IMPORTANT NOTE: this constant is shared with the main PhotoDemon project.  DO NOT CHANGE IT!
Private Const PD_PATCH_IDENTIFIER As Long = &H50554450   'PD update patch data (ASCII characters "PDUP", as hex, little-endian)

'A module-level pdFSO object is provided, for convenience.
Private m_File As pdFSO

'If silent mode has been activated via command line, this will be set to TRUE.
Private m_SilentMode As Boolean

'I previously built PD automatic update files only on an extremely old desktop PC.
' These days it's more convenient for me to build it wherever I'm working,
' which means the repo may not be in the same location.
Private m_basePath As String

Private Sub cmdAction_Click(Index As Integer)
    
    Select Case Index
    
        'Assemble nightly build files
        Case 0
            AssembleNightlyBuild
            
        'Assemble stable + beta build files
        Case 1
            AssembleStableAndBetaBuilds
            
        'Build version and checksum file
        Case 2
            MakeVersionFile
            
    End Select
    
End Sub

'The nightly build is unique, because we generate it directly from the current PD development folder.  As such, it uses a
' different series of assembly steps (compared to the stable and beta builds).
Private Sub AssembleNightlyBuild()
    
    'This list of relevant files is hardcoded to match the nightly build script's instructions for 7zip.
    Dim nightlyList As pdStringStack
    Set nightlyList = New pdStringStack
    
    nightlyList.AddString m_basePath & "PhotoDemon\PhotoDemon.exe"
    nightlyList.AddString m_basePath & "PhotoDemon\README.md"
    nightlyList.AddString m_basePath & "PhotoDemon\LICENSE.md"
    nightlyList.AddString m_basePath & "PhotoDemon\AUTHORS.md"
    nightlyList.AddString m_basePath & "PhotoDemon\CODE_OF_CONDUCT.md"
    nightlyList.AddString m_basePath & "PhotoDemon\Donate to PhotoDemon.url"
    
    'For the /App subfolder, we forcibly restrict which extensions are allowed, to avoid copying any backup files
    ' or other unwanted entries.
    m_File.RetrieveAllFiles m_basePath & "PhotoDemon\App\", nightlyList, True, False, "exe|txt|TXT|dll|xml|pdrc"
    
    'Manually remove any files we don't want to include in nightly downloads
    nightlyList.RemoveStringByText m_basePath & "PhotoDemon\App\PhotoDemon\Plugins\avifdec.exe", True
    nightlyList.RemoveStringByText m_basePath & "PhotoDemon\App\PhotoDemon\Plugins\avifenc.exe", True
    nightlyList.RemoveStringByText m_basePath & "PhotoDemon\App\PhotoDemon\Plugins\brotlicommon.dll", True
    nightlyList.RemoveStringByText m_basePath & "PhotoDemon\App\PhotoDemon\Plugins\brotlidec.dll", True
    nightlyList.RemoveStringByText m_basePath & "PhotoDemon\App\PhotoDemon\Plugins\brotlienc.dll", True
    nightlyList.RemoveStringByText m_basePath & "PhotoDemon\App\PhotoDemon\Plugins\cjxl.exe", True
    nightlyList.RemoveStringByText m_basePath & "PhotoDemon\App\PhotoDemon\Plugins\djxl.exe", True
    nightlyList.RemoveStringByText m_basePath & "PhotoDemon\App\PhotoDemon\Plugins\jxl.dll", True
    nightlyList.RemoveStringByText m_basePath & "PhotoDemon\App\PhotoDemon\Plugins\jxl_threads.dll", True
    nightlyList.RemoveStringByText m_basePath & "PhotoDemon\App\PhotoDemon\Plugins\jxlinfo.exe", True
    nightlyList.RemoveStringByText m_basePath & "PhotoDemon\App\PhotoDemon\Plugins\libjxl-LICENSE.brotli", True
    nightlyList.RemoveStringByText m_basePath & "PhotoDemon\App\PhotoDemon\Plugins\libjxl-LICENSE.giflib", True
    nightlyList.RemoveStringByText m_basePath & "PhotoDemon\App\PhotoDemon\Plugins\libjxl-LICENSE.highway", True
    nightlyList.RemoveStringByText m_basePath & "PhotoDemon\App\PhotoDemon\Plugins\libjxl-LICENSE.libjpeg-turbo", True
    nightlyList.RemoveStringByText m_basePath & "PhotoDemon\App\PhotoDemon\Plugins\libjxl-LICENSE.libjxl", True
    nightlyList.RemoveStringByText m_basePath & "PhotoDemon\App\PhotoDemon\Plugins\libjxl-LICENSE.libpng", True
    nightlyList.RemoveStringByText m_basePath & "PhotoDemon\App\PhotoDemon\Plugins\libjxl-LICENSE.libwebp", True
    nightlyList.RemoveStringByText m_basePath & "PhotoDemon\App\PhotoDemon\Plugins\libjxl-LICENSE.sjpeg", True
    nightlyList.RemoveStringByText m_basePath & "PhotoDemon\App\PhotoDemon\Plugins\libjxl-LICENSE.skcms", True
    nightlyList.RemoveStringByText m_basePath & "PhotoDemon\App\PhotoDemon\Plugins\libjxl-LICENSE.zlib", True
    
    'Make a copy of the current file list for post-compression verification
    Dim verifyFiles As pdStringStack
    Set verifyFiles = New pdStringStack
    verifyFiles.CloneStack nightlyList
    
    'Assemble the corresponding pdPackage
    Dim nightlyPackage As pdPackager
    Set nightlyPackage = New pdPackager
    nightlyPackage.PrepareNewPackage 4, PD_PATCH_IDENTIFIER
    
    nightlyPackage.AutoAddNodesFromStringStack nightlyList, m_basePath & "PhotoDemon\", , True
    
    'We also want to add the update patching program itself
    nightlyPackage.AutoAddNodeFromFile m_basePath & "PhotoDemon\Support\Update Patcher 2.0\PD_Update_Patcher.exe", 99, "\PD_Update_Patcher.exe"
    
    'Write the completed package out to the updates folder
    nightlyPackage.WritePackageToFile m_basePath & "PhotoDemon\no_sync\PD_Updates\nightly.pdz2", , True
    
    'Next, we're going to extract all packaged files to a temp folder.  This serves two purposes: it lets us verify that the packaging went
    ' as expected, and it also gives us a dedicated folder we can scan for assembling version and checksum data.
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    nightlyPackage.ReadPackageFromFile m_basePath & "PhotoDemon\no_sync\PD_Updates\nightly.pdz2", PD_PATCH_IDENTIFIER
    nightlyPackage.AutoExtractAllFiles m_basePath & "PhotoDemon\no_sync\PD_Updates\nightly_auto_extract_test\"
    Debug.Print "Verified nightly build file extraction in " & VBHacks.GetTimeDiffNowAsString(startTime)
    
    'We now want to manually verify file contents, to ensure they are byte-for-byte identical
    Debug.Print "Testing all files for equality..."
    Dim cFSO As pdFSO
    Set cFSO = New pdFSO
    
    Dim failedFiles As pdStringStack
    Set failedFiles = New pdStringStack
    
    Dim testFilePath As String
    Do While verifyFiles.PopString(testFilePath)
        If (Not Files.FilesEqual(testFilePath, m_basePath & "PhotoDemon\no_sync\PD_Updates\nightly_auto_extract_test\" & cFSO.GenerateRelativePath(m_basePath & "PhotoDemon\", testFilePath))) Then
            failedFiles.AddString testFilePath
        End If
    Loop
    
    If (failedFiles.GetNumOfStrings = 0) Then
        Debug.Print "All files verified successfully!"
    Else
        MsgBox "WARNING!  One or more nightly build files reported a change to file contents post-extraction.  Filenames will be shown in order.", vbOKOnly
        Do While failedFiles.PopString(testFilePath)
            MsgBox testFilePath, vbOKOnly
        Loop
    End If
    
End Sub

'Stable and Beta update channels use custom, dedicated folders.  The contents of these folders are updated manually,
' only when necessary (as opposed to the nightly channel which is built directly from the current PD codebase).
Private Sub AssembleStableAndBetaBuilds()

    'Stable and beta builds can be constructed directly from their folders, no special work required.
    
    'Assemble a basic pdPackage instance
    Dim cPackage As pdPackager
    Set cPackage = New pdPackager
    
    'Build the stable update file directly from its folder.
    cPackage.PrepareNewPackage 4, PD_PATCH_IDENTIFIER
    cPackage.AutoAddNodesFromFolder m_basePath & "PhotoDemon\no_sync\PD_Updates\stable\", 0
    cPackage.AutoAddNodeFromFile m_basePath & "PhotoDemon\Support\Update Patcher 2.0\PD_Update_Patcher.exe", 99, "\PD_Update_Patcher.exe"
    cPackage.WritePackageToFile m_basePath & "PhotoDemon\no_sync\PD_Updates\stable.pdz2", , True
    
    'Repeat the above steps for the beta update folder
    Set cPackage = New pdPackager
    cPackage.PrepareNewPackage 4, PD_PATCH_IDENTIFIER
    cPackage.AutoAddNodesFromFolder m_basePath & "PhotoDemon\no_sync\PD_Updates\beta\", 0
    cPackage.AutoAddNodeFromFile m_basePath & "PhotoDemon\Support\Update Patcher 2.0\PD_Update_Patcher.exe", 99, "\PD_Update_Patcher.exe"
    cPackage.WritePackageToFile m_basePath & "PhotoDemon\no_sync\PD_Updates\beta.pdz2", , True
    
    'Want to test extraction, to verify everything was stored correctly?  Use these lines:
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    Set cPackage = New pdPackager
    cPackage.ReadPackageFromFile m_basePath & "PhotoDemon\no_sync\PD_Updates\beta.pdz2", PD_PATCH_IDENTIFIER
    cPackage.AutoExtractAllFiles m_basePath & "PhotoDemon\no_sync\PD_Updates\beta_auto_extract_test\"
    Debug.Print "Verified beta build file extraction in " & VBHacks.GetTimeDiffNowAsString(startTime)
    
    VBHacks.GetHighResTime startTime
    Set cPackage = New pdPackager
    cPackage.ReadPackageFromFile m_basePath & "PhotoDemon\no_sync\PD_Updates\stable.pdz2", PD_PATCH_IDENTIFIER
    cPackage.AutoExtractAllFiles m_basePath & "PhotoDemon\no_sync\PD_Updates\stable_auto_extract_test\"
    Debug.Print "Verified stable build file extraction in " & VBHacks.GetTimeDiffNowAsString(startTime)
    
End Sub

'Generate a central version XML file, by reading the version numbers from each .exe.
Private Sub MakeVersionFile()
    
    'Retrieve stable, beta, developer build versions
    Dim vStable As String, vBeta As String, vDev As String
    vStable = GetFileVersion_Modified(m_basePath & "PhotoDemon\no_sync\PD_updates\stable\PhotoDemon.exe")
    vBeta = GetFileVersion_Modified(m_basePath & "PhotoDemon\no_sync\PD_updates\beta\PhotoDemon.exe")
    vDev = GetFileVersion_Modified(m_basePath & "PhotoDemon\no_sync\PD_updates\nightly_auto_extract_test\PhotoDemon.exe")
    
    'We now want to write these version numbers out to file - specifically, the YAML file that describes
    ' the PhotoDemon update server homepage.
    Dim targetFile As String
    targetFile = m_basePath & "PhotoDemon-Updates-v2\_config.yml"
    
    Dim srcYML As String
    If Files.FileLoadAsString(targetFile, srcYML) Then
    
        'Split the file by newline
        Dim srcLines() As String
        srcLines = Split(srcYML, vbCrLf)
        
        'Search for the lines that describe version number
        Dim i As Long
        For i = LBound(srcLines) To UBound(srcLines)
            
            If (InStr(1, srcLines(i), "pd-stable-v:", vbBinaryCompare) <> 0) Then
                srcLines(i) = "pd-stable-v: " & vStable
            ElseIf (InStr(1, srcLines(i), "pd-beta-v:", vbBinaryCompare) <> 0) Then
                srcLines(i) = "pd-beta-v: " & vBeta
            ElseIf (InStr(1, srcLines(i), "pd-nightly-v:", vbBinaryCompare) <> 0) Then
                srcLines(i) = "pd-nightly-v: " & vDev
                
                'Nightly version data should be the last line in the file; exit now
                ReDim Preserve srcLines(LBound(srcLines) To i + 1) As String
                Exit For
                
            End If
            
        Next i
        
        'Reassemble the text, then overwrite the old file
        Dim cString As pdString
        Set cString = New pdString
        
        For i = LBound(srcLines) To UBound(srcLines)
            cString.AppendLine srcLines(i)
        Next i
        
        If Files.FileExists(targetFile) Then Files.FileDelete targetFile
        Files.FileSaveAsText cString.ToString, targetFile, True, False
        
    End If
    
    Exit Sub
    
    'Prep an XML object.
    Dim xmlOutput As pdXML
    Set xmlOutput = New pdXML
    
    xmlOutput.prepareNewXML "Program version"
    xmlOutput.writeBlankLine
    xmlOutput.writeComment "This program version file was automatically generated on " & Format(Now, "Medium date")
    xmlOutput.writeBlankLine
    
    'For each build, we're going to generate some key pieces of information.  Start with the stable build.
    xmlOutput.writeTagWithAttribute "update", "track", "stable", "", True
    AddVersionGroupToXML xmlOutput, m_basePath & "PhotoDemon\no_sync\PD_updates\stable\"
    xmlOutput.closeTag "update"
    xmlOutput.writeBlankLine
    
    'Next comes beta (which is often the same as the stable release)
    xmlOutput.writeTagWithAttribute "update", "track", "beta", "", True
    AddVersionGroupToXML xmlOutput, m_basePath & "PhotoDemon\no_sync\PD_updates\beta\"
    xmlOutput.closeTag "update"
    xmlOutput.writeBlankLine
    
    'Last comes nightly.  Note that the nightly files will be out of date unless Step 1 (AssembleNightlyBuild) has been run during this session.
    xmlOutput.writeTagWithAttribute "update", "track", "nightly", "", True
    AddVersionGroupToXML xmlOutput, m_basePath & "PhotoDemon\no_sync\PD_updates\nightly\"
    xmlOutput.closeTag "update"
    xmlOutput.writeBlankLine
    
    'Also, write out release announcement links.  These are stored in a custom local XML file.
    AddReleaseAnnouncementLinks xmlOutput, m_basePath & "PhotoDemon\no_sync\PD_updates\release_announcements.xml"
    
    'Write the XML out to file
    Dim dstFile As String
    dstFile = m_basePath & "pdupdate2.xml"
    
    xmlOutput.writeXMLToFile dstFile
    
End Sub

'Given a path to the release announcement URL file, copy those links into the central language version XML file
Private Sub AddReleaseAnnouncementLinks(ByRef xmlOutput As pdXML, ByRef srcPath As String)

    'Create an XML engine to parse the source document
    Dim xmlSource As pdXML
    Set xmlSource = New pdXML
    
    If xmlSource.loadXMLFile(srcPath) Then
    
        xmlOutput.writeTag "raurl-stable", xmlSource.getUniqueTag_String("raurl-stable")
        xmlOutput.writeTag "raurl-beta", xmlSource.getUniqueTag_String("raurl-beta")
        xmlOutput.writeTag "raurl-nightly", xmlSource.getUniqueTag_String("raurl-nightly")
        xmlOutput.writeBlankLine
        xmlOutput.writeTag "releasenumber-beta", xmlSource.getUniqueTag_String("releasenumber-beta")
    
    Else
        MsgBox "Something went wrong with the release announcement URL file.  You should probably investigate.", vbOKOnly + vbApplicationModal + vbCritical, "Release announcement XML failure"
    End If
    
    xmlOutput.writeBlankLine

End Sub

'Helpful wrapper to add version and checksum data to an output XML object
Private Sub AddVersionGroupToXML(ByRef xmlOutput As pdXML, ByRef srcPath As String)
    
    'A pdPackage instance is used to generate checksums
    Dim cPackager As pdPackager
    Set cPackager = New pdPackager
    
    'We're now going to assemble a list of files that need to be parsed.  This list is universal for every program group (at present; we can add
    ' custom code in the future if required files change between versions).
    Dim buildFiles As pdStringStack
    Set buildFiles = New pdStringStack
    
    'Note that XML files are currently ignored, as they're handled by the separate language file update protocol
    m_File.RetrieveAllFiles srcPath, buildFiles, True, False
    
    Dim curFile As String, vString As String
    
    'Iterate through each file, adding its version and checksum to the update file as we go
    Do While buildFiles.PopString(curFile)
        
        'Retrieve the file's version (if any)
        vString = GetFileVersion_Modified(curFile)
        
        'If version isn't available, we must fall back to checksums for updating files.
        If (StrComp(vString, "unknown", vbBinaryCompare) <> 0) Then
            xmlOutput.writeTagWithAttribute "version", "component", m_File.GenerateRelativePath(srcPath, curFile), vString
        End If
        
    Loop
    
End Sub

'Small convenience wrapper, so we can plug in "unknown" when the version number is, actually, unknown
Private Function GetFileVersion_Modified(ByRef srcFilename As String, Optional ByVal useThisIfVersionDoesntExist As String = "unknown") As String
    
    Dim vString As String
    
    If m_File.FileGetVersionAsString(srcFilename, vString) Then
        GetFileVersion_Modified = vString
    Else
        GetFileVersion_Modified = useThisIfVersionDoesntExist
    End If
    
End Function

Private Sub Form_Load()
    
    Set m_File = New pdFSO
    
    If Files.PathExists("C:\PhotoDemon v4", False) Then
        m_basePath = "C:\PhotoDemon v4\"
    Else
        m_basePath = "C:\tanner-dev\"
    End If
    
    'Initialize compression engines
    Compression.InitializeCompressionEngine PD_CE_Zstd, App.Path & "\"
    
    'Check the command line.  This project can be run in silent mode as part of my nightly build batch script.
    Dim chkCommandLine As String
    chkCommandLine = Command$
    
    If (LenB(Trim$(chkCommandLine)) <> 0) Then
        m_SilentMode = (InStr(1, chkCommandLine, "-s", vbTextCompare) <> 0)
    End If
    
    'If silent mode is activated, automatically "click" the relevant button
    If m_SilentMode Then
    
        'Assemble the nightly build update package
        Call cmdAction_Click(0)
        
        'Assemble the stable and beta build update packages
        Call cmdAction_Click(1)
        
        'Generate the central version and checksum file
        Call cmdAction_Click(2)
        
        'If the program is running in silent mode, unload it now
        Unload Me
        
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    'Release compression engines
    Compression.ShutDownCompressionEngine PD_CE_Zstd

End Sub
