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
'Copyright 2015 by Tanner Helland
'Created: 01/Februrary/15
'Last updated: 03/March/15
'Last update: migrate various patch functions out of PhotoDemon and into this dedicated app
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Const PD_PATCH_IDENTIFIER As Long = &H50554450   'PD update patch data (ASCII characters "PDUP", as hex, little-endian)

Private Const MAX_PATH_LEN = 260

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
    szexeFile As String * MAX_PATH_LEN
End Type

'APIs for making sure PhotoDemon.exe has terminated
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
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
    If parseCommandLine() Then
        
        'Wait for PD to close; when it does, the timer will initiate the rest of the patch process.
        txtOut.Text = "Waiting for PhotoDemon to terminate..."
        m_PDClosed = False
        tmrCheck.Enabled = True
        
    'If the command line is empty, the user somehow ran this independent of PD.  Terminate immediately.
    Else
    
        textOut "It appears that something other than PhotoDemon launched this program.", False
        textOut "For security reasons, this update patcher will not run unless started by PhotoDemon itself.", False
        textOut "(You may close this window now.)", False
        
    End If
    
End Sub

'Parse the command line for all relevant instructions.  PD handles some update tasks for us, and it relays its findings through
' the command line.
Private Function parseCommandLine() As Boolean
    
    Dim cUnicode As pdUnicode
    Set cUnicode = New pdUnicode
    
    'Split params according to spaces
    Dim allParams() As String
    allParams = Split(cUnicode.CommandW, " ")
    
    'Check for an empty command line
    If UBound(allParams) <= LBound(allParams) Then
        textOut "WARNING! Input parameters invalid (" & cUnicode.CommandW & ").", False
        parseCommandLine = False
    
    'Retrieve all parameters
    Else
    
        Dim curLine As Long
        curLine = LBound(allParams)
        
        'Iterate through the params, looking for meaningful entries as we go
        Do While curLine <= UBound(allParams)
            
            'Start checking instructions of interest
            If StringsEqual(allParams(curLine), "/restart") Then
                m_RestartWhenDone = True
            
            ElseIf StringsEqual(allParams(curLine), "/start") Then
                
                'Retrieve the start position
                curLine = curLine + 1
                m_TrackStartPosition = CLng(allParams(curLine))
                
            ElseIf StringsEqual(allParams(curLine), "/end") Then
                
                'Retrieve the start position
                curLine = curLine + 1
                m_TrackEndPosition = CLng(allParams(curLine))
            
            End If
            
            'Increment to the next line and continue checking params
            curLine = curLine + 1
            
        Loop
        
        parseCommandLine = True
        
    End If
    
End Function

'Shortcut function for checking string equality
Private Function StringsEqual(ByVal strOne As String, ByVal strTwo As String) As Boolean
    StringsEqual = (StrComp(Trim$(strOne), Trim$(strTwo), vbBinaryCompare) = 0)
End Function

Private Sub tmrCheck_Timer()

    'Check to see if PD has closed.
    If (Not m_PDClosed) Then
    
        Dim pdFound As Boolean
        pdFound = False
        
        'Prepare to iterate through all running processes
        Const TH32CS_SNAPPROCESS As Long = 2&
        Const PROCESS_ALL_ACCESS = 0
        Dim uProcess As PROCESSENTRY32
        Dim rProcessFound As Long, hSnapshot As Long, myProcess As Long
        Dim szExename As String
        Dim i As Long
        
        On Local Error GoTo PDDetectionError
    
        'Prepare a generic process reference
        uProcess.dwSize = Len(uProcess)
        hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
        rProcessFound = ProcessFirst(hSnapshot, uProcess)
        
        'Iterate through all running processes, looking for PhotoDemon instances
        Do While rProcessFound
    
            'Retrieve the EXE name of this process
            i = InStr(1, uProcess.szexeFile, Chr(0))
            szExename = LCase$(Left$(uProcess.szexeFile, i - 1))
            
            'If the process name is "exiftool.exe", terminate it
            If Right$(szExename, Len("PhotoDemon.exe")) = "PhotoDemon.exe" Then
                
                pdFound = True
                Exit Do
                 
            End If
            
            'Find the next process, then continue
            rProcessFound = ProcessNext(hSnapshot, uProcess)
        
        Loop
    
        'Release our generic process snapshot
        CloseHandle hSnapshot
    
        'If PD was found, do nothing.  Otherwise, start patching the program.
        If Not pdFound Then
            
            'Disable this timer
            tmrCheck.Enabled = False
            
            'Start the patch process
            m_PDClosed = True
            startPatching
            
        End If
    
        Exit Sub
    
    End If
    
PDDetectionError:

    textOut "Unknown error occurred while waiting for PhotoDemon to close.  Checking again..."

End Sub

'Start the patch process
Private Function startPatching() As Boolean
    
    textOut "PhotoDemon shutdown detected.  Starting patch process."
    
    'This update patcher will have been extracted to PD's root folder.
    m_PDPath = App.Path
    
    If StrComp(Right$(m_PDPath, 1), "\", vbBinaryCompare) <> 0 Then m_PDPath = m_PDPath & "\"
    m_PDUpdatePath = m_PDPath & "Data\Updates\"
    m_PluginPath = m_PDPath & "App\PhotoDemon\Plugins\"
    
    'Retrieve the patch XML file from its hard-coded location
    Dim xmlEngine As pdXML
    Set xmlEngine = New pdXML
    
    If xmlEngine.loadXMLFile(m_PDUpdatePath & "patch.xml") Then
        
        If xmlEngine.isPDDataType("Program version") Then
            
            'This function will only return TRUE if all files were patched successfully.
            Dim allFilesSuccessful As Boolean
            allFilesSuccessful = True
            
            'Temporary files are a necessary evil of this function, due to the ugliness of patching in-use binary files.
            ' As a security precaution, we'll be hashing our temp filenames.
            Randomize Timer
            
            Dim cHash As CSHA256
            Set cHash = New CSHA256
            
            'A pdFSO object helps with some extra file operations
            Dim cFile As pdFSO
            Set cFile = New pdFSO
            
            'The downloaded data is saved in the /Data/Updates folder.  Retrieve it directly into a pdPackager object.
            Dim cPackage As pdPackager
            Set cPackage = New pdPackager
            cPackage.init_ZLib m_PluginPath & "zlibwapi.dll"
            
            If cPackage.readPackageFromFile(m_PDUpdatePath & "PDPatch.tmp", PD_PATCH_IDENTIFIER) Then
            
                'The package appears to be intact.  Time to start enumerating and patching files.
                Dim rawNewFile() As Byte, newFilenameArray() As Byte, newFilename As String, failsafeChecksum As Long
                Dim rawOldFile() As Byte
                
                Dim numOfNodes As Long
                numOfNodes = cPackage.getNumOfNodes
                
                'Iterate each file in turn, extracting as we go
                Dim i As Long
                For i = 0 To numOfNodes - 1
                
                    'Somewhat unconventionally, we extract the file's contents prior to extracting its name.  We want to verify that the contents
                    ' are intact (via pdPackage's internal checksum data) before proceeding with actual file patching.
                    If cPackage.getNodeDataByIndex(i, False, rawNewFile) Then
                    
                        'If we made it here, it means the internal pdPackage checksum passed successfully, meaning the post-compression file checksum
                        ' matches the original checksum calculated at creation time.  Because we are very cautious, we now apply a second checksum verification,
                        ' using the checksum value embedded within the original pdupdate.xml file.
            
                        'Start by retrieving the filename of the updated language file; we need this to look up the original checksum value in the update XML file.
                        If cPackage.getNodeDataByIndex(i, True, newFilenameArray) Then
                            
                            newFilename = Space$((UBound(newFilenameArray) + 1) \ 2)
                            CopyMemory ByVal StrPtr(newFilename), ByVal VarPtr(newFilenameArray(0)), UBound(newFilenameArray) + 1
                            
                            'Ignore the update patcher itself
                            If Not StringsEqual(newFilename, "PD_Update_Patcher.exe") Then
                            
                                'Retrieve the secondary failsafe checksum for this file
                                failsafeChecksum = getFailsafeChecksum(xmlEngine, newFilename)
                                
                                'Before proceeding with the write, compare the temp file array to our stored checksum
                                If failsafeChecksum = cPackage.checkSumArbitraryArray(rawNewFile) Then
                        
                                    'Checksums match!  We now want to overwrite the old binary file with its new copy.
                                     
                                    'First, we must write this file out to a temporary file.  The filename doesn't matter, but we'll hash it as a
                                    ' privacy and security precaution.
                                    Dim tmpFilename As String
                                    tmpFilename = Left$(cHash.SHA256(CStr(Rnd) & newFilename), 16) & ".tmp"
                                    
                                    'Write the temp file
                                    If cFile.SaveByteArrayToFile(rawNewFile, m_PDUpdatePath & tmpFilename) Then
                                    
                                        'The temp file is ready to go.  Prepare a destination name, which we get by appending the embedded pdPackage name
                                        ' and the current PD folder.
                                        Dim dstFilename As String
                                        dstFilename = m_PDPath & newFilename
                                    
                                        'Use a special patch function to replace the binary file in question
                                        Dim patchResult As FILE_PATCH_RESULT
                                        patchResult = patchArbitraryFile(dstFilename, m_PDUpdatePath & tmpFilename, , True, failsafeChecksum, cPackage)
                                        
                                        If patchResult = FPR_SUCCESS Then
                                        
                                            textOut "Successfully patched " & newFilename, False
                                            
                                        Else
                                        
                                            textOut "WARNING! patchProgramFiles failed to patch " & newFilename
                                                
                                            Select Case patchResult
                                            
                                                Case FPR_FAIL_NOTHING_CHANGED
                                                    textOut "(However, patchProgramFiles was able to restore everything to its initial state.)"
                                                    
                                                Case FPR_FAIL_BOTH_FILES_REMOVED
                                                    textOut "WARNING! Somehow, patchProgramFiles managed to kill both files while it was at it."
                                                
                                                Case FPR_FAIL_NEW_FILE_REMOVED
                                                    textOut "WARNING! Somehow, patchProgramFiles managed to kill the new file while it was at it."
                                                
                                                Case FPR_FAIL_OLD_FILE_REMOVED
                                                    textOut "WARNING! Somehow, patchProgramFiles managed to kill the old file while it was at it."
                                                
                                            End Select
                                            
                                            allFilesSuccessful = False
                                            
                                        'End patchArbitraryFile success
                                        End If
                                    
                                    'End writing temp file success
                                    End If
                                    
                                'End secondary checksum failsafe
                                End If
                                
                            'End ignoring the update patch program itself
                            End If
                        
                        'End node header data retrieval success
                        End If
                    
                    'End node data retrieval success
                    End If
                
                Next i
                
                m_PatchSuccessful = allFilesSuccessful
                
            Else
                textOut "Patch file is missing or corrupted.  Patching cannot proceed.", False
                m_PatchSuccessful = False
            End If
            
        Else
            textOut "Update XML file doesn't contain patch data.  Patching cannot proceed.", False
            m_PatchSuccessful = False
        End If
        
    Else
        textOut "Update XML file wasn't found.  Patching cannot proceed.", False
        m_PatchSuccessful = False
    End If
    
    'Regardless of outcome, perform some clean-up afterward.
    finishPatching
    
End Function

'When patching program files, we double-check checksums of both the temp files and the final binary copies.  This prevents hijackers from
' intercepting the files mid-transit, and replacing them with their own.
Private Function getFailsafeChecksum(ByRef xmlEngine As pdXML, ByVal relativePath As String) As Long

    'Find the position of this file's checksum
    Dim pdTagPosition As Long
    pdTagPosition = xmlEngine.getLocationOfTagPlusAttribute("checksum", "component", relativePath, m_TrackStartPosition)
    
    'Make sure the tag position is within the valid range.  (This should always be TRUE, but it doesn't hurt to check.)
    If (pdTagPosition >= m_TrackStartPosition) And (pdTagPosition <= m_TrackEndPosition) Then
    
        'This is the checksum tag we want!  Retrieve its value.
        Dim thisChecksum As String
        thisChecksum = xmlEngine.getTagValueAtPreciseLocation(pdTagPosition)
        
        'Convert the checksum to a long and return it
        getFailsafeChecksum = thisChecksum
        
    'If the checksum doesn't exist in the file, return 0
    Else
        getFailsafeChecksum = 0
    End If
    
    'Debug.Print pdTagPosition & " (" & m_TrackStartPosition & ", " & m_TrackEndPosition & "): " & relativePath
    
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
Public Function patchArbitraryFile(ByVal oldFile As String, ByVal newFile As String, Optional ByVal customBackupFile As String = "", Optional ByVal handleBackupsForMe As Boolean = True, Optional ByVal srcChecksum As Long = 0, Optional ByRef srcPackage As pdPackager = Nothing) As FILE_PATCH_RESULT
    
    'Create a pdFSO instance
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    'If the user wants us to handle backups, we'll hash their incoming filename as our backup name.
    Dim cHash As CSHA256
    Set cHash = New CSHA256
    
    'Before doing anything, look for an incoming checksum.  If one was provided, compare it against the original (old) file now.
    ' If the checksum matches the old file, it means the old and new files are identical, so we can skip the patch process.
    If (srcChecksum <> 0) Then
        
        'If the old file doesn't exist, we're installing a new file, so ignore this first checksum verification.
        ' (A second checksum verification will still be applied after the new file is written.)
        If cFile.FileExist(oldFile) Then
        
            'Compare old and new checksums
            If srcPackage.checkSumArbitraryFile(oldFile) = srcChecksum Then
                
                'Checksums are identical.  Patching is not required.  Report TRUE and exit now.
                #If DEBUGMODE = 1 Then
                    pdDebug.LogAction "patchArbitraryFile skipped patching of " & oldFile & " because it's identical to the new file."
                #End If
                
                patchArbitraryFile = FPR_SUCCESS
                Exit Function
                
            End If
            
        End If
        
    End If
    
    'Two paths are required: a more complicated one for backups that we handle, and a thin wrapper otherwise
    If handleBackupsForMe Then
        
        'We use the standard Data/Updates folder for backups when patching files
        customBackupFile = m_PDUpdatePath & Left$(cHash.SHA256(cFile.getFilename(oldFile) & CStr(Timer)), 16) & ".tmp"
        
        'Copy the contents of newFile to Backup file
        If cFile.CopyFile(oldFile, customBackupFile) Then
        
            'With a backup successfully created, lean on the API to perform the actual patching
            Dim patchResult As FILE_PATCH_RESULT
            patchResult = cFile.ReplaceFile(oldFile, newFile)
            
            'If the patch succeeds, great!
            If patchResult = FPR_SUCCESS Then
            
                patchArbitraryFile = FPR_SUCCESS
                
            'If the patch does not succeed, restore our backup as necessary
            Else
            
                'If the old file still exists, kill our backup, then return the appropriate fail state
                If cFile.FileExist(oldFile) Then
                    patchArbitraryFile = FPR_FAIL_NOTHING_CHANGED
                
                'The old file is missing.  Restore it from our backup.
                Else
                    
                    If cFile.CopyFile(customBackupFile, oldFile) Then
                        patchArbitraryFile = FPR_FAIL_NOTHING_CHANGED
                    
                    'If we can't restore our backup, things are really messed up.  We have no choice but to exit.
                    Else
                        patchArbitraryFile = FPR_FAIL_OLD_FILE_REMOVED
                    End If
                    
                End If
            
            End If
        
        'If the copy failed, try and get the API to copy the file for us.  This isn't ideal, as the API may leave behind a copy of the backup file,
        ' but it's better than nothing.
        Else
            
            textOut "WARNING! patchArbitraryFile was unable to create a manual backup prior to patching.", False
            
            'Leave it to the API from here...
            patchArbitraryFile = cFile.ReplaceFile(oldFile, newFile, customBackupFile)
            
        End If
        
    'If the caller doesn't want us to handle backups, its up to them to
    Else
        patchArbitraryFile = cFile.ReplaceFile(oldFile, newFile, customBackupFile)
    End If
    
    'If we made it all the way here, the replace operation completed.  If it thinks it was successful, and a checksum was provided, perform a final
    ' failsafe checksum verification on the new file.
    If (srcChecksum <> 0) And (patchArbitraryFile = FPR_SUCCESS) Then
    
        'Validate the oldFile (which now contains the contents of newFile)
        If srcPackage.checkSumArbitraryFile(oldFile) <> srcChecksum Then
        
            'The checksums don't match, which means something went horribly wrong.  Restore our backup now.
            If cFile.ReplaceFile(oldFile, customBackupFile) = FPR_SUCCESS Then
            
                'Any damage was undone.  Report a matching fail state.
                patchArbitraryFile = FPR_FAIL_NOTHING_CHANGED
            
            Else
                
                'We couldn't undo the damage.  This is an impossible outcome, IMO, but catch it anyway.
                cFile.KillFile oldFile
                patchArbitraryFile = FPR_FAIL_OLD_FILE_REMOVED
                
                textOut "WARNING! patchArbitraryFile detected a checksum mismatch, but it was unable to restore the backup file.", False
                textOut "WARNING! (File in question is " & oldFile & ")", False
                
            End If
        
        End If
    
    End If
    
    'By this point, the function has done everything it can to ensure one of two states:
    ' - A successful replacement operation
    ' - A failed replacement operation, but everything has been restored to its original state.
    
    'Regardless of outcome, we no longer need our backup file, so kill it
    cFile.KillFile customBackupFile
    
End Function

'Regardless of patch success or failure, this function is called.  If the user wants us to restart PD, we do so now.
Private Sub finishPatching()
    
    textOut "Update process complete.  Applying final validation to all updated files."
    
    If m_RestartWhenDone Then
        
        textOut "Restarting PhotoDemon, as requested."
        
        Dim actionString As String, fileString As String, pathString As String, paramString As String
        actionString = "open"
        fileString = "PhotoDemon.exe"
        pathString = m_PDPath
        paramString = ""
        
        ShellExecute 0&, StrPtr(actionString), StrPtr(fileString), 0&, StrPtr(pathString), SW_SHOWNORMAL
    
    End If
    
    textOut "Validation passed.  Shutting down update patcher.", False
    Unload Me

End Sub

'Display basic update text
Public Sub textOut(ByVal newText As String, Optional ByVal appendEllipses As Boolean = True)
    
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
