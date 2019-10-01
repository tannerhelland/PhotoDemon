VERSION 5.00
Begin VB.Form FormPatch 
   BackColor       =   &H80000005&
   Caption         =   " PhotoDemon Update"
   ClientHeight    =   4695
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9885
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   313
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   659
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrCheck 
      Enabled         =   0   'False
      Interval        =   750
      Left            =   9360
      Top             =   120
   End
   Begin VB.TextBox txtOut 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   9135
   End
End
Attribute VB_Name = "FormPatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Program Update Patching App
'Copyright 2015-2018 by Tanner Helland
'Created: 01/Februrary/15
'Last updated: 13/June/18
'Last update: total overhaul against new patching strategy
'
'PhotoDemon's small update-patcher program.  This program is downloaded as part of an update file.  PD extracts it
' and shells it prior to closing; this file then handles the rest of the patching process.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Const PD_PATCH_IDENTIFIER As Long = &H50554450   'PD update patch data (ASCII characters "PDUP", as hex, little-endian)

Private Const MAX_PATH_LEN As Long = 260

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH_LEN
End Type

'APIs for making sure PhotoDemon.exe has terminated
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'APIs for restarting PhotoDemon.exe when we're done
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteW" (ByVal hWnd As Long, ByVal lpOperationStr As Long, ByVal lpFileStr As Long, ByVal lpParametersStr As Long, ByVal lpDirectoryStr As Long, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal byteLength As Long)

'Several paths are required for the update process: a path to PD's folder, the update subfolder, and the plugin folder (for accessing zLib)
Private m_PDPath As String, m_PDUpdatePath As String, m_PluginPath As String

'Once PhotoDemon.exe can no longer be detected as an active process, this will be set to TRUE
Private m_PDClosed As Boolean

'If the user asked PD to restart after patching, this app will be notified of the decision via command-line
Private m_RestartWhenDone As Boolean

'If the patch was successful, this will be set to TRUE
Private m_PatchSuccessful As Boolean

'PhotoDemon passes some values to us via command line:
Private m_TrackStartPosition As Long, m_TrackEndPosition As Long  'Start and end position of the relevant update track in the update XML file

'This program starts working as soon as it loads.  No user interaction is expected or handled.
Private Sub Form_Load()
    
    'Position the output text box
    txtOut.Width = FormPatch.ScaleWidth - txtOut.Left * 2
    
    'Replace the crappy default VB icon
    SetIcon Me.hWnd, "AAA", True
    
    'Display the window
    Me.Show
    
    'Check relevant command-line params; this function returns TRUE if the command line contains parameters
    If ParseCommandLine() Then
        
        'Wait for PD to close; when it does, the timer will initiate the rest of the patch process.
        txtOut.Text = "Waiting for PhotoDemon to terminate..."
        m_PDClosed = False
        tmrCheck.Enabled = True
        
    'If the command line is empty, the user somehow ran this independent of PD.  Terminate immediately.
    Else
    
        TextOut "Something other than PhotoDemon launched this program.", False
        TextOut "For security reasons, this update patcher will not run unless started by PhotoDemon itself.", False
        TextOut "(You may close this window now.)", False
        
    End If
    
End Sub

'Parse the command line for all relevant instructions.  PD handles some update tasks for us, and it relays its findings through
' the command line.
Private Function ParseCommandLine() As Boolean
    
    'Retrieve a complete list of input parameters, pre-parsed for us
    Dim allParams As pdStringStack
    Set allParams = New pdStringStack
    If OS.CommandW(allParams, True) Then
    
        'Make sure we were passed valid input params
        If (allParams.GetNumOfStrings = 0) Then
            TextOut "WARNING! Input parameters invalid; no arguments found.", False
            ParseCommandLine = False
    
        'Retrieve all parameters
        Else
        
            Dim curLine As String, pdSrcFound As Boolean
            pdSrcFound = False
            
            'Iterate through the params, looking for meaningful entries as we go
            Do While allParams.PopString(curLine)
                
                'Start checking instructions of interest
                If Strings.StringsEqual(Trim$(curLine), "/restart", True) Then
                    m_RestartWhenDone = True
                ElseIf Strings.StringsEqual(Trim$(curLine), "/sourceIsPD", True) Then
                    pdSrcFound = True
                End If
                
            Loop
            
            ParseCommandLine = pdSrcFound
            
        End If
        
    End If
    
End Function

Private Sub tmrCheck_Timer()

    'Check to see if PD has closed.
    If (Not m_PDClosed) Then
    
        Dim pdFound As Boolean
        pdFound = False
        
        'Prepare to iterate through all running processes
        Const TH32CS_SNAPPROCESS As Long = 2&
        Const PROCESS_ALL_ACCESS As Long = 0&
        Dim uProcess As PROCESSENTRY32
        Dim rProcessFound As Long, hSnapShot As Long, myProcess As Long
        Dim szExename As String
        Dim i As Long
        
        On Local Error GoTo PDDetectionError
    
        'Prepare a generic process reference
        uProcess.dwSize = Len(uProcess)
        hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
        If (hSnapShot <> 0) Then
            
            rProcessFound = ProcessFirst(hSnapShot, uProcess)
            
            'Iterate through all running processes, looking for PhotoDemon instances
            Do While (rProcessFound <> 0)
        
                'Retrieve the EXE name of this process
                i = InStr(1, uProcess.szExeFile, Chr(0))
                If (i > 1) Then
                    
                    szExename = LCase$(Left$(uProcess.szExeFile, i - 1))
                    
                    'If the process name is "PhotoDemon.exe", note it
                    If Right$(szExename, Len("PhotoDemon.exe")) = "PhotoDemon.exe" Then
                        pdFound = True
                        Exit Do
                    End If
                    
                End If
                
                'Find the next process, then continue
                rProcessFound = ProcessNext(hSnapShot, uProcess)
            
            Loop
        
            'Release our generic process snapshot
            CloseHandle hSnapShot
            
        End If
        
        'If PD was found, do nothing.  Otherwise, start patching the program.
        If (Not pdFound) Then
            
            'Disable this timer
            tmrCheck.Enabled = False
            
            'Start the patch process
            m_PDClosed = True
            StartPatching
            
        End If
    
        Exit Sub
    
    End If
    
PDDetectionError:
    TextOut "Error occurred while waiting for PhotoDemon to close (#" & Err.Number & ": " & Err.Description & ").  Checking again..."

End Sub

'Start the patch process
Private Function StartPatching() As Boolean
    
    TextOut "PhotoDemon shutdown detected.  Starting patch process."
    
    'This update patcher will have been extracted to PD's root folder.
    m_PDPath = Files.PathAddBackslash(App.Path)
    m_PDUpdatePath = m_PDPath & "Data\Updates\"
    m_PluginPath = m_PDPath & "App\PhotoDemon\Plugins\"
    
    'This function will only return TRUE if all files were patched successfully.
    Dim allFilesSuccessful As Boolean
    allFilesSuccessful = True
    
    'Temporary files are a necessary evil of this function, due to the ugliness of patching in-use binary files.
    ' As a courtesy, we'll hash to create arbitrary temp filenames.
    Randomize Timer
    
    Dim cHash As pdCrypto
    Set cHash = New pdCrypto
    
    'A pdFSO object helps with some extra file operations
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    'Initialize a zstd decompressor
    Compression.StartCompressionEngines m_PluginPath
    
    'The downloaded data is saved in the /Data/Updates folder.  Retrieve it directly into a pdPackager object.
    Dim cPackage As pdPackager
    Set cPackage = New pdPackager
    If cPackage.ReadPackageFromFile(m_PDUpdatePath & "PDPatch.tmp", PD_PATCH_IDENTIFIER) Then
            
        'The package appears to be intact.  Time to start enumerating and patching files.
        Dim rawNewFile() As Byte, newFilename As String, rndSrcString As String
        Dim rawOldFile() As Byte
        
        Dim numOfNodes As Long
        numOfNodes = cPackage.GetNumOfNodes
        
        'Iterate each file in turn, extracting as we go
        Dim i As Long
        For i = 0 To numOfNodes - 1
            
            'Grab the filename
            If cPackage.GetNodeDataByIndex_String(i, True, newFilename) Then
            
                'Ignore the patcher itself (it's already extracted and running, obviously!)
                If Strings.StringsNotEqual(newFilename, "\PD_Update_Patcher.exe", True) Then
                
                    'Grab the file's bits
                    If cPackage.GetNodeDataByIndex(i, False, rawNewFile) Then
                        
                        'We now want to overwrite the old binary file with its new copy.
                        
                        'First, we must write this file out to a temporary file.  The filename doesn't matter, but we'll hash it as a
                        ' privacy and security precaution.
                        rndSrcString = CStr(Rnd * Timer)
                        Dim tmpFilename As String
                        tmpFilename = cHash.QuickHash_AsString(StrPtr(rndSrcString), LenB(rndSrcString), 16, PDCA_SHA_256) & ".tmp"
                        
                        'Dump the bits out to that temp file
                        If cFile.FileCreateFromByteArray(rawNewFile, m_PDUpdatePath & tmpFilename, , True) Then
                            
                            'The temp file is ready to go.  Prepare a destination name, which we get by appending the embedded pdPackage name
                            ' and the current PD folder.
                            Dim dstFilename As String
                            dstFilename = m_PDPath & newFilename
                            
                            'Use a special patch function to replace the binary file in question
                            Dim patchResult As PD_FILE_PATCH_RESULT
                            patchResult = PatchArbitraryFile(dstFilename, m_PDUpdatePath & tmpFilename, , True, cPackage)
                            
                            If patchResult = FPR_SUCCESS Then
                                TextOut "Successfully patched " & newFilename, False
                            Else
                            
                                TextOut "WARNING! patchProgramFiles failed to patch " & newFilename
                                    
                                Select Case patchResult
                                
                                    Case FPR_FAIL_NOTHING_CHANGED
                                        TextOut "(However, patchProgramFiles was able to restore everything to its initial state.)"
                                        
                                    Case FPR_FAIL_BOTH_FILES_REMOVED
                                        TextOut "WARNING! Somehow, patchProgramFiles managed to kill both files while it was at it."
                                    
                                    Case FPR_FAIL_NEW_FILE_REMOVED
                                        TextOut "WARNING! Somehow, patchProgramFiles managed to kill the new file while it was at it."
                                    
                                    Case FPR_FAIL_OLD_FILE_REMOVED
                                        TextOut "WARNING! Somehow, patchProgramFiles managed to kill the old file while it was at it."
                                    
                                End Select
                                                
                                allFilesSuccessful = False
                                
                            'End PatchArbitraryFile success
                            End If
                        
                        'End writing temp file success
                        Else
                            TextOut "WARNING!  Failed to write temp file copy of " & newFilename
                        End If
                        
                    'End node data retrieval success
                    Else
                        TextOut "WARNING!  Failed to retrieve data node for #" & i
                    End If
                
                'End ignoring the patcher itself
                End If
                
            'End header data retrieval success
            Else
                TextOut "WARNING!  Failed to retrieve header node for #" & i
            End If
            
        Next i
                
        m_PatchSuccessful = allFilesSuccessful
    
    'End loading the update pdPackage file
    Else
        TextOut "Patch file is missing or corrupted.  Patching cannot proceed.", False
        m_PatchSuccessful = False
    End If
    
    'Regardless of outcome, perform some clean-up afterward.
    FinishPatching
    
End Function

'Advanced wrapper to pdFSO's ReplaceFile function.  This function adds several features to the ReplaceFile function:
' - A comprehensive backup system for the original file, which this function will use to undo the replace operation if anything
'    goes wrong during the replacement process.
' - Support for a newFile checksum (adler32, as generated by a pdPackage object, which the caller must have initiated with zLib support).
'    - It is assumed that the caller has already verified that this checksum is valid for the new file.
'    - This checksum is used in two places:
'       - First, if the checksum matches the oldFile, the replace step is skipped (as the files are identical)
'       - Second, if the replacement process reports success, this checksum is validated AGAIN on the destination file.  This verifies that
'          nothing hijacked the replacement process.
'
'Returns a FILE_PATCH_RESULT enum.  (FPR_SUCCESS means success; all other returns are various failures.)
Public Function PatchArbitraryFile(ByVal oldFile As String, ByVal newFile As String, Optional ByVal customBackupFile As String = vbNullString, Optional ByVal handleBackupsForMe As Boolean = True, Optional ByRef srcPackage As pdPackager = Nothing) As PD_FILE_PATCH_RESULT
    
    'Create a pdFSO instance
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    'If the user wants us to handle backups, we'll hash their incoming filename as our backup name.
    Dim cHash As pdCrypto
    Set cHash = New pdCrypto
    
    'Two paths are required: a more complicated one for backups that we handle, and a thin wrapper otherwise
    If handleBackupsForMe Then
        
        'We use the standard Data/Updates folder for backups when patching files
        customBackupFile = m_PDUpdatePath & cHash.QuickHash_AsString(StrPtr(oldFile), LenB(oldFile), 16, PDCA_SHA_256) & ".tmp"
        
        'Copy the contents of newFile to Backup file.  (Note that we can skip this step if the old file doesn't exist.)
        Dim copySuccess As Boolean
        If cFile.FileExists(oldFile) Then
            copySuccess = cFile.FileCopyW(oldFile, customBackupFile)
        Else
            copySuccess = True
        End If
        
        If copySuccess Then
        
            'With a backup successfully created, lean on the API to perform the actual patching
            Dim patchResult As PD_FILE_PATCH_RESULT
            If cFile.FileExists(oldFile) Then
                patchResult = cFile.FileReplace(oldFile, newFile)
            Else
                If cFile.FileCopyW(newFile, oldFile) Then
                    patchResult = FPR_SUCCESS
                    cFile.FileDelete newFile
                Else
                    patchResult = FPR_FAIL_NOTHING_CHANGED
                End If
            End If
            
            'If the patch succeeds, great!
            If (patchResult = FPR_SUCCESS) Then
                PatchArbitraryFile = FPR_SUCCESS
                
            'If the patch does not succeed, restore our backup as necessary
            Else
            
                'If the old file still exists, kill our backup, then return the appropriate fail state
                If cFile.FileExists(oldFile) Then
                    PatchArbitraryFile = FPR_FAIL_NOTHING_CHANGED
                
                'The old file is missing.  Restore it from our backup.
                Else
                    
                    If cFile.FileExists(customBackupFile) Then
                        If cFile.FileCopyW(customBackupFile, oldFile) Then
                            PatchArbitraryFile = FPR_FAIL_NOTHING_CHANGED
                        Else
                            PatchArbitraryFile = FPR_FAIL_OLD_FILE_REMOVED
                        End If
                    
                    'If we can't restore our backup, things are really messed up.  We have no choice but to exit.
                    Else
                        PatchArbitraryFile = FPR_FAIL_OLD_FILE_REMOVED
                    End If
                    
                End If
            
            End If
        
        'If the copy failed, try and get the API to copy the file for us.  This isn't ideal, as the API may leave behind a copy of the backup file,
        ' but it's better than nothing.
        Else
            
            TextOut "WARNING! PatchArbitraryFile was unable to create a manual backup prior to patching.", False
            
            'Leave it to the API from here...
            PatchArbitraryFile = cFile.FileReplace(oldFile, newFile, customBackupFile)
            
        End If
        
    'If the caller doesn't want us to handle backups, its up to them to
    Else
        PatchArbitraryFile = cFile.FileReplace(oldFile, newFile, customBackupFile)
    End If
    
    'By this point, the function has done everything it can to ensure one of two states:
    ' - A successful replacement operation
    ' - A failed replacement operation, but everything has been restored to its original state.
    
    'Regardless of outcome, we no longer need our backup file, so kill it
    If cFile.FileExists(customBackupFile) Then cFile.FileDelete customBackupFile
    
End Function

'Regardless of patch success or failure, this function is called.  If the user wants us to restart PD, we do so now.
Private Sub FinishPatching()
    
    TextOut "Update process complete.  Checking user's restart request."
    
    If m_RestartWhenDone Then
        
        TextOut "Restarting PhotoDemon, as requested."
        
        Dim actionString As String, fileString As String, pathString As String, paramString As String
        actionString = "open"
        fileString = "PhotoDemon.exe"
        pathString = m_PDPath
        paramString = ""
        
        ShellExecute 0&, StrPtr(actionString), StrPtr(fileString), 0&, StrPtr(pathString), SW_SHOWNORMAL
    
    End If
    
    'Shut down any open compression engines
    Compression.StopCompressionEngines
    
    TextOut "Writing final log and shutting down update patcher."
    
    On Error Resume Next
    
    Dim cFile As pdFSO: Set cFile = New pdFSO
    cFile.FileSaveAsText txtOut.Text, m_PDUpdatePath & "update_log.txt"
    
    Unload Me

End Sub

'Display basic update text
Public Sub TextOut(ByVal newText As String, Optional ByVal appendEllipses As Boolean = True)
    
    If appendEllipses Then
    
        If StrComp(Right$(newText, 1), ".", vbBinaryCompare) = 0 Then
            txtOut.Text = txtOut.Text & vbCrLf & newText & ".."
        Else
            txtOut.Text = txtOut.Text & vbCrLf & newText & "..."
        End If
        
    Else
        txtOut.Text = txtOut.Text & vbCrLf & newText
    End If
    
    'Stick the cursor at the end of the text, which looks more natural IMO
    txtOut.Refresh
    txtOut.SelStart = Len(txtOut.Text)
    
End Sub
