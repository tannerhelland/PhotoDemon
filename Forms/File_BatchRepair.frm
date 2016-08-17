VERSION 5.00
Begin VB.Form FormBatchRepair 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Batch repair"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8550
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   440
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   570
   ShowInTaskbar   =   0   'False
   Begin PhotoDemon.pdButtonStrip btsMoveCopy 
      Height          =   855
      Left            =   360
      TabIndex        =   7
      Top             =   2640
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   1508
      Caption         =   "after a successful repair"
      FontSizeCaption =   10
   End
   Begin PhotoDemon.pdCheckBox chkRecurseFolders 
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1080
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   450
      Caption         =   "include subfolders"
   End
   Begin PhotoDemon.pdCommandBarMini cmdBar 
      Align           =   2  'Align Bottom
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   5985
      Width           =   8550
      _ExtentX        =   15081
      _ExtentY        =   1085
      DontAutoUnloadParent=   -1  'True
   End
   Begin PhotoDemon.pdLabel lblProgress 
      Height          =   495
      Left            =   120
      Top             =   5160
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   873
      Alignment       =   2
      Caption         =   "click OK to begin the repair operation"
      FontBold        =   -1  'True
   End
   Begin PhotoDemon.pdCheckBox chkRepairs 
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   4080
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   873
      Caption         =   "ensure file extensions match repaired image types"
   End
   Begin PhotoDemon.pdButton cmdSrcFolder 
      Height          =   450
      Left            =   7800
      TabIndex        =   0
      Top             =   555
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   794
      Caption         =   "..."
   End
   Begin PhotoDemon.pdTextBox txtSrcFolder 
      Height          =   315
      Left            =   360
      TabIndex        =   1
      Top             =   630
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   556
      Text            =   "automatically generated at run-time"
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   1
      Left            =   120
      Top             =   120
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   503
      Caption         =   "source folder"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdButton cmdDstFolder 
      Height          =   450
      Left            =   7800
      TabIndex        =   2
      Top             =   2115
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   794
      Caption         =   "..."
   End
   Begin PhotoDemon.pdTextBox txtDstFolder 
      Height          =   315
      Left            =   360
      TabIndex        =   3
      Top             =   2190
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   556
      Text            =   "automatically generated at run-time"
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   0
      Left            =   120
      Top             =   1680
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   503
      Caption         =   "destination folder (for repaired files only)"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   2
      Left            =   120
      Top             =   3600
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   503
      Caption         =   "repair operations"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   3
      Left            =   120
      Top             =   4800
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   503
      Caption         =   "progress"
      FontSize        =   12
      ForeColor       =   4210752
   End
End
Attribute VB_Name = "FormBatchRepair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBar_OKClick()
    
    Dim tmpDIB As pdDIB
    Dim tmpExtension As String, newFilename As String
    
    'Make sure the destination folder exists
    Dim cFSO As pdFSO
    Set cFSO = New pdFSO
    
    Dim dstFolder As String
    dstFolder = FixPath(txtDstFolder.Text)
    
    If (Not cFSO.FolderExist(dstFolder, False)) Then cFSO.CreateFolder dstFolder, True
    
    'Start by preparing the list of files to be processed
    Dim listOfFiles As pdStringStack
    Set listOfFiles = New pdStringStack
    
    UpdateProgress g_Language.TranslateMessage("generating list of files to repair...")
    
    Dim numOfFiles As Long, curFileNumber As Long, numFilesRepaired As Long
    Dim fileWasRepaired As Boolean
    
    'For improved performance, copy all user options into local variables
    Dim identifyFileType As Boolean
    identifyFileType = CBool(chkRepairs(0).Value = vbChecked)
    
    Dim eraseOriginal As Boolean
    eraseOriginal = CBool(btsMoveCopy.ListIndex = 1)
    
    'cFSO returns TRUE if at least one file is found; this is good enough for us to attempt repairs
    If cFSO.RetrieveAllFiles(txtSrcFolder.Text, listOfFiles, CBool(chkRecurseFolders.Value), False) Then
            
        numOfFiles = listOfFiles.GetNumOfStrings
        curFileNumber = 1
            
        Dim srcFilename As String
        Do While listOfFiles.PopString(srcFilename)
            
            UpdateProgress g_Language.TranslateMessage("processing file %1 of %2 (%3 repairs performed)...", curFileNumber, numOfFiles, numFilesRepaired)
            fileWasRepaired = False
            
            'Attempt to load the file as an image
            If identifyFileType Then
                If Loading.QuickLoadImageToDIB(srcFilename, tmpDIB, False, False, False) Then
                
                    'This is a valid image file.  Determine the file's correct extension.
                    tmpExtension = g_ImageFormats.GetExtensionFromPDIF(tmpDIB.GetOriginalFormat())
                    
                    'If the file already has the correct extension, ignore it
                    If (StrComp(tmpExtension, cFSO.GetFileExtension(srcFilename), vbBinaryCompare) <> 0) Then
                        
                        newFilename = dstFolder & cFSO.GetFilename(srcFilename, True) & "." & tmpExtension
                        
                        'Move the file - with its new extension - to the repaired folder
                        If cFSO.CopyFile(srcFilename, newFilename) Then
                            If eraseOriginal Then cFSO.KillFile srcFilename
                            fileWasRepaired = True
                        End If
                        
                    End If
                    
                    'Free the temporary DIB
                    Set tmpDIB = Nothing
                    
                End If
            End If
            
            'If one or more repair steps was applied,
            If fileWasRepaired Then
                numFilesRepaired = numFilesRepaired + 1
            End If
            
            'Regardless of what happened to this file, increment the current file count
            curFileNumber = curFileNumber + 1
            If ((curFileNumber And 31) = 0) Then DoEvents
            
        Loop
        
        PDMsgBox "%1 image(s) repaired." & vbCrLf & vbCrLf & "Repaired images have been saved to the destination folder.  Unrepaired images remain in their original locations.", vbOKOnly Or vbInformation, "Repair complete", numFilesRepaired
        
    Else
        PDMsgBox "The source folder does not contain any files.  Please try another folder.", vbOKOnly Or vbInformation, "No files found"
    End If
    
End Sub

Private Sub UpdateProgress(ByVal newMessage As String)
    lblProgress.Caption = newMessage
    lblProgress.RequestRefresh
End Sub

Private Sub cmdDstFolder_Click()
    Dim folderPath As String
    folderPath = FileSystem.BrowseForFolder(Me.hWnd, txtDstFolder.Text)
    If (Len(folderPath) <> 0) Then
        txtDstFolder.Text = FixPath(folderPath)
        g_UserPreferences.SetPref_String "BatchProcess", "RepairDstFolder", txtDstFolder.Text
    End If
End Sub

Private Sub cmdSrcFolder_Click()
    Dim folderPath As String
    folderPath = FileSystem.BrowseForFolder(Me.hWnd, txtSrcFolder.Text)
    If (Len(folderPath) <> 0) Then
        txtSrcFolder.Text = FixPath(folderPath)
        g_UserPreferences.SetPref_String "BatchProcess", "RepairSrcFolder", txtSrcFolder.Text
    End If
End Sub

Private Sub Form_Load()

    'Load default source/dest folders.  If previously saved paths are not found, default to the user's current
    ' open/save image paths.
    If g_UserPreferences.DoesValueExist("BatchProcess", "RepairSrcFolder") Then
        txtSrcFolder.Text = g_UserPreferences.GetPref_String("BatchProcess", "RepairSrcFolder", g_UserPreferences.GetPref_String("Paths", "Open Image", ""))
    Else
        txtSrcFolder.Text = g_UserPreferences.GetPref_String("Paths", "Open Image", "")
    End If
    
    If g_UserPreferences.DoesValueExist("BatchProcess", "RepairDstFolder") Then
        txtDstFolder.Text = g_UserPreferences.GetPref_String("BatchProcess", "RepairDstFolder", g_UserPreferences.GetPref_String("Paths", "Save Image", ""))
    Else
        txtDstFolder.Text = g_UserPreferences.GetPref_String("Paths", "Save Image", "")
    End If
    
    btsMoveCopy.AddItem "keep original (unrepaired) file", 0
    btsMoveCopy.AddItem "erase original file", 1
    btsMoveCopy.ListIndex = 0
    
    Interface.ApplyThemeAndTranslations Me

End Sub
