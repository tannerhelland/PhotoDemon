VERSION 5.00
Begin VB.Form FormBatchRepair 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Batch repair"
   ClientHeight    =   6765
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
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   451
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   570
   ShowInTaskbar   =   0   'False
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
      Top             =   6150
      Width           =   8550
      _ExtentX        =   15081
      _ExtentY        =   1085
      DontAutoUnloadParent=   -1  'True
   End
   Begin PhotoDemon.pdLabel lblProgress 
      Height          =   495
      Left            =   120
      Top             =   5520
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   873
      Alignment       =   2
      Caption         =   "click OK to begin the repair operation"
      FontBold        =   -1  'True
   End
   Begin PhotoDemon.pdCheckBox chkRepairs 
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   3360
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   873
      Caption         =   "attempt to repair image files"
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
      Top             =   2880
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   503
      Caption         =   "repair options"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdLabel lblTitle 
      Height          =   285
      Index           =   3
      Left            =   120
      Top             =   5160
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   503
      Caption         =   "progress"
      FontSize        =   12
      ForeColor       =   4210752
   End
   Begin PhotoDemon.pdCheckBox chkRepairs 
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   3720
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   661
      Caption         =   "attempt to repair video and audio files"
   End
   Begin PhotoDemon.pdCheckBox chkRepairs 
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   8
      Top             =   4080
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   661
      Caption         =   "overwrite matching filenames in the destination folder"
      Value           =   0   'False
   End
   Begin PhotoDemon.pdCheckBox chkRepairs 
      Height          =   375
      Index           =   3
      Left            =   360
      TabIndex        =   9
      Top             =   4440
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   661
      Caption         =   "after a successful recovery, erase the original (unrepaired) file"
      Value           =   0   'False
   End
End
Attribute VB_Name = "FormBatchRepair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Batch Repair dialog
'Copyright 2016-2026 by Tanner Helland
'Created: 16/August/16
'Last updated: 18/August/16
'Last update: add video repair capabilities
'
'Scandisk is a reasonably good tool for failed/failing hardware.  That said, it's very naive about file recovery;
' potential files (or fragments) are extracted as generic .chk files, and it's up to the user to figure out if
' any of the .chk files contain meaningful data.
'
'This dialog aims to help that process, at least when it comes to images and videos.  This dialog will search
' the results of a Scandisk folder (or any folder, really), and look for files whose extension does not match
' the actual file type.  All supported image types are scanned, and because mpeg video formats are so common on
' modern cameras and phones, rudimentary video file detection is also available.
'
'Recovered files are automatically assigned the correct file extension, and (optionally) moved to a destination
' folder of the user's choosing.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Declare Function mciSendStringW Lib "winmm" (ByVal lpstrCommand As Long, ByVal lpstrReturnString As Long, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Private m_OkayToProceed As Boolean

Private Sub cmdBar_CancelClick()
    m_OkayToProceed = False
End Sub

Private Sub cmdBar_OKClick()
    
    Dim tmpDIB As pdDIB
    Dim newExtension As String, newFilename As String
    
    'Make sure the destination folder exists
    Dim cFSO As pdFSO
    Set cFSO = New pdFSO
    
    Dim dstFolder As String
    dstFolder = Files.PathAddBackslash(txtDstFolder.Text)
    
    If (Not Files.PathExists(dstFolder, False)) Then Files.PathCreate dstFolder, True
    
    'Start by preparing the list of files to be processed
    Dim listOfFiles As pdStringStack
    Set listOfFiles = New pdStringStack
    
    UpdateProgress g_Language.TranslateMessage("generating list of files to repair...")
    
    Dim numOfFiles As Long, curFileNumber As Long, numFilesRepaired As Long
    Dim fileWasRepaired As Boolean
    
    'When testing broken video files, MCI command strings will be used
    Dim tmpFilename As String, mciString As String, mciResult As Long
    
    'For improved performance, copy all user options into local variables
    Dim identifyImageFiles As Boolean
    identifyImageFiles = chkRepairs(0).Value
    
    Dim identifyVideoFiles As Boolean
    identifyVideoFiles = chkRepairs(1).Value
    
    Dim eraseDestinationMatches As Boolean
    eraseDestinationMatches = chkRepairs(2).Value
    
    Dim eraseOriginal As Boolean
    eraseOriginal = chkRepairs(3).Value
    
    'cFSO returns TRUE if at least one file is found; this is good enough for us to attempt repairs
    If cFSO.RetrieveAllFiles(txtSrcFolder.Text, listOfFiles, chkRecurseFolders.Value, False) Then
            
        m_OkayToProceed = True
        
        numOfFiles = listOfFiles.GetNumOfStrings
        curFileNumber = 1
        
        Dim srcFilename As String
        Do While (listOfFiles.PopString(srcFilename) And m_OkayToProceed)
            
            UpdateProgress g_Language.TranslateMessage("processing file %1 of %2 (%3 repairs performed)...", curFileNumber, numOfFiles, numFilesRepaired)
            fileWasRepaired = False
            
            'Attempt to load the file as an image
            If identifyImageFiles Then
                If Loading.QuickLoadImageToDIB(srcFilename, tmpDIB, False, False) Then
                    
                    'This is a valid image file.  Determine the file's correct extension.
                    newExtension = ImageFormats.GetExtensionFromPDIF(tmpDIB.GetOriginalFormat())
                    fileWasRepaired = True
                    Set tmpDIB = Nothing
                    
                End If
            End If
            
            'Look for video files, if the user has requested it.  (This is significantly less sophisticated than our
            ' automated image detection, but it should catch most common video formats from modern cameras.)
            If identifyVideoFiles And (Not fileWasRepaired) Then
                
                'Don't proceed unless the file is more than 16 kb in size (this improves performance, as scandisk
                ' loves to create hordes of chunk-sized recovery files)
                If (FileLen(srcFilename) > 16384) Then
                    
                    'Filenames with spaces must be enclosed in quotes
                    If InStr(srcFilename, " ") Then
                        tmpFilename = ChrW$(34) & srcFilename & ChrW$(34)
                    Else
                        tmpFilename = srcFilename
                    End If
                    
                    'Couple of notes on this MCI command string:
                    ' 1) We want to just open (*not* play) the file
                    ' 2) We want to assign the file an alias so that we can close it after testing the open command
                    ' 3) If we don't explicitly test the file as an mpeg-type video, the file's extension will be used
                    '    to infer format (which is useless during a repair op!)
                    ' 4) Although slower, we want the command to return asynchronously so that we can check it's return.
                    '    Hypothetically, a callback function could be used, but that's outside the scope of the current
                    '    repair implementation.
                    mciString = "open " & tmpFilename & " alias pd" & curFileNumber & " type mpegvideo wait"
                    mciResult = mciSendStringW(StrPtr(mciString), 0&, 0&, 0&)
                    
                    'There are multiple potential failure values for testing a digital file; if success is returned,
                    ' attempt a repair
                    If (mciResult = 0) Then
                        mciString = "close pd" & curFileNumber
                        mciResult = mciSendStringW(StrPtr(mciString), 0&, 0&, 0&)
                        newExtension = "mp4"
                        fileWasRepaired = True
                    Else
                    
                        'Attempt again, but this time, use the AVI codec set
                        mciString = "open " & tmpFilename & " alias pd" & curFileNumber & " type avivideo wait"
                        mciResult = mciSendStringW(StrPtr(mciString), 0&, 0&, 0&)
                        
                        If (mciResult = 0) Then
                            mciString = "close pd" & curFileNumber
                            mciResult = mciSendStringW(StrPtr(mciString), 0&, 0&, 0&)
                            newExtension = "avi"
                            fileWasRepaired = True
                        Else
                            fileWasRepaired = False
                        End If
                        
                    End If
                    
                End If
                
            End If
            
            'If this file is repairable, the newExtension string will also be filled (so we know what kind
            ' of file to write!)
            If fileWasRepaired Then
                
                'If the file already has the correct extension, ignore it
                If Strings.StringsNotEqual(newExtension, Files.FileGetExtension(srcFilename), False) Then
                    
                    newFilename = dstFolder & Files.FileGetName(srcFilename, True) & "." & newExtension
                    
                    'The user can optionally request to overwrite files in the destination folder; if this option
                    ' is selected, check for matching filenames before writing.
                    If eraseDestinationMatches Then
                        fileWasRepaired = Not Files.FileExists(newFilename)
                    End If
                    
                    'Move the file - with its new extension - to the repaired folder
                    If fileWasRepaired Then
                        If cFSO.FileCopyW(srcFilename, newFilename) Then
                            If eraseOriginal Then cFSO.FileDelete srcFilename
                            fileWasRepaired = True
                        Else
                            fileWasRepaired = False
                        End If
                    End If
                        
                Else
                    fileWasRepaired = False
                End If
                    
            End If
            
            'If one or more repair steps was applied, increment the repair counter
            If fileWasRepaired Then numFilesRepaired = numFilesRepaired + 1
            
            'Regardless of what happened to this file, increment the current file count
            curFileNumber = curFileNumber + 1
            If ((curFileNumber And 7) = 0) Then DoEvents
            
        Loop
        
        UpdateProgress g_Language.TranslateMessage("Repairs complete.  %1 files repaired.", numFilesRepaired)
        PDMsgBox "%1 files(s) repaired." & vbCrLf & vbCrLf & "Repaired files have been saved to the destination folder.  Unrepaired files remain in their original locations.", vbOKOnly Or vbInformation, "Repairs complete", numFilesRepaired
        
    Else
        PDMsgBox "The source folder does not contain any files.  Please try another folder.", vbOKOnly Or vbInformation, "No files found"
    End If
    
    'Normally, an OK button press unloads the parent form.  This dialog is an exception.
    cmdBar.DoNotUnloadForm
    
End Sub

Private Sub UpdateProgress(ByVal newMessage As String)
    lblProgress.Caption = newMessage
    lblProgress.RequestRefresh
End Sub

Private Sub cmdDstFolder_Click()
    Dim folderPath As String
    folderPath = Files.PathBrowseDialog(Me.hWnd, txtDstFolder.Text)
    If (LenB(folderPath) <> 0) Then
        txtDstFolder.Text = Files.PathAddBackslash(folderPath)
        UserPrefs.SetPref_String "BatchProcess", "RepairDstFolder", txtDstFolder.Text
    End If
End Sub

Private Sub cmdSrcFolder_Click()
    Dim folderPath As String
    folderPath = Files.PathBrowseDialog(Me.hWnd, txtSrcFolder.Text)
    If (LenB(folderPath) <> 0) Then
        txtSrcFolder.Text = Files.PathAddBackslash(folderPath)
        UserPrefs.SetPref_String "BatchProcess", "RepairSrcFolder", txtSrcFolder.Text
    End If
End Sub

Private Sub Form_Load()

    'Load default source/dest folders.  If previously saved paths are not found, default to the user's current
    ' open/save image paths.
    If UserPrefs.DoesValueExist("BatchProcess", "RepairSrcFolder") Then
        txtSrcFolder.Text = UserPrefs.GetPref_String("BatchProcess", "RepairSrcFolder", UserPrefs.GetPref_String("Paths", "Open Image", vbNullString))
    Else
        txtSrcFolder.Text = UserPrefs.GetPref_String("Paths", "Open Image", vbNullString)
    End If
    
    If UserPrefs.DoesValueExist("BatchProcess", "RepairDstFolder") Then
        txtDstFolder.Text = UserPrefs.GetPref_String("BatchProcess", "RepairDstFolder", UserPrefs.GetPref_String("Paths", "Save Image", vbNullString))
    Else
        txtDstFolder.Text = UserPrefs.GetPref_String("Paths", "Save Image", vbNullString)
    End If
    
    Interface.ApplyThemeAndTranslations Me

End Sub
