Attribute VB_Name = "Files"
'***************************************************************************
'Comprehensive wrapper for pdFSO (Unicode file and folder functions)
'Copyright 2001-2018 by Tanner Helland
'Created: 6/12/01
'Last updated: 12/July/17
'Last update: large code cleanup
'
'The pdFSO class provides Unicode file/folder interactions for PhotoDemon.  However, sometimes you just want to do
' something trivial, like checking whether a file exists, without instantiating a full class (especially because this
' is unnecessarily verbose in VB).  This module wraps a pdFSO instance and allows you to directly invoke common
' functions without worrying about the details.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'API calls for retrieving detailed date time for a given file
Private Const STARTF_USESHOWWINDOW As Long = &H1
Private Const SW_NORMAL As Long = 1
Private Const SW_HIDE As Long = 0
Private Const WAIT_INFINITE As Long = -1

Private Type WIN32_STARTUP_INFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type WIN32_PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

'Used to shell an external program, then wait until it completes
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreateProcessW Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As Long, ByRef lpProcessAttributes As Any, ByRef lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByRef lpEnvironment As Any, ByVal lpCurrentDriectory As Long, ByVal lpStartupInfo As Long, ByVal lpProcessInformation As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

'PD creates a lot of temp files.  To prevent hard drive thrashing, this module tracks created temp files,
' and deletes them en masse when PD terminates.  (You can also delete temp files prematurely by calling the
' associated function, but because some PD tasks run asynchronously, this behavior is not generally advised.)
Private m_NumOfTempFiles As Long
Private m_ListOfTempFiles() As String
Private Const INIT_TEMP_FILE_CACHE As Long = 16

'pdFSO is the primary workhorse for file-level interactions.  (I use a class because it is sometimes helpful to cache certain values internally,
' especially for operations that may be running in parallel; pdFSO makes this possible.)  However, for trivial operations that don't require a
' dedicated class instance, pdFSO can be a pain.  We work around that by caching a local pdFSO instance, and wrapping it for any relevant
' one-off operations.
Private m_FSO As pdFSO

'Free all temporary files generated this session.  Note that this (obviously) only deletes temp files generated via the
' RequestTempFile() function, below.
Public Sub DeleteTempFiles()

    If (m_NumOfTempFiles > 0) Then
    
        Dim i As Long
        For i = 0 To m_NumOfTempFiles - 1
            If (Len(m_ListOfTempFiles(i)) <> 0) Then Files.FileDeleteIfExists m_ListOfTempFiles(i)
        Next i
        
        m_NumOfTempFiles = 0
    
    End If
    
End Sub

'Quick and dirty "compare 2 files" function; returns TRUE if the files are byte-for-byte identical.
' Designed only for small files that can be fully cached in memory.
Public Function FilesEqual(ByRef srcFile1 As String, ByRef srcFile2 As String) As Boolean
    
    If InitializeFSO Then
    
        'Load both files into memory.
        Dim srcBytes1() As Byte, srcBytes2() As Byte
        If m_FSO.FileLoadAsByteArray(srcFile1, srcBytes1) And m_FSO.FileLoadAsByteArray(srcFile2, srcBytes2) Then
            If UBound(srcBytes1) = UBound(srcBytes2) Then
                FilesEqual = VBHacks.MemCmp(VarPtr(srcBytes1(0)), VarPtr(srcBytes2(0)), UBound(srcBytes1) + 1)
            End If
        End If
    
    End If

End Function

'If a file exists, this function can be used to intelligently increment the file name (e.g. "filename (n+1).ext")
' Note that the function returns the auto-incremented filename WITHOUT an extension and WITHOUT a prepended folder,
' by design, so that the result can be passed to a common dialog without further parsing.
'
'That said, an initial extension is still required, because this function should only be used if a file name with
' a matching extension exists (e.g. it is perfectly fine to have the same filename with DIFFERENT extensions in the
' target directory).
Public Function IncrementFilename(ByRef dstDirectory As String, ByRef srcFilename As String, ByRef desiredExtension As String) As String
    
    'First, check to see if a file with that name and extension appears in the destination directory.
    ' If it does, just return the filename we were passed.
    If (Not Files.FileExists(dstDirectory & srcFilename & "." & desiredExtension)) Then
        IncrementFilename = srcFilename
    Else

        'If we made it to this line of code, a file with that name and extension appears in the destination directory.
        
        'Start by figuring out if the file is already in the format: "filename (#).ext"
        Dim tmpFilename As String
        tmpFilename = Trim$(srcFilename)
        
        Dim numToAppend As Long
        
        'Check the trailing character.  If it is a closing parentheses ")", we need to analyze more
        If Strings.StringsEqual(Right$(tmpFilename, 1), ")", False) Then
        
            Dim i As Long
            For i = Len(tmpFilename) - 1 To 1 Step -1
                
                'If this character is a number, continue scanning leftward until we find a character that is *not* a number
                If (Not IsNumeric(Mid$(tmpFilename, i, 1))) Then
                    
                    'This character is non-numeric.  See if it's an opening parentheses.
                    If Strings.StringsEqual(Mid$(tmpFilename, i, 1), "(", False) Then
                        
                        'This filename adheres to the pattern "Filename (###).ext".  To spare us from auto-scanning
                        ' numbers that are likely taken, use this number as our starting point for auto-incrementing.
                        numToAppend = CLng(Mid$(tmpFilename, i + 1, Len(tmpFilename) - i - 1))
                        tmpFilename = Left$(tmpFilename, i - 2)
                        Exit For
                    
                    'If this character is non-numeric and NOT an initial parentheses, this filename is not in the format we want.
                    ' Treat it like any other filename and start by appending " (2)" to it
                    Else
                        numToAppend = 2
                        Exit For
                    End If
                    
                End If
            
            Next i
        
        'If this is not already a copy of the format "Filename (###).ext", start scanning at ### = 2
        Else
            numToAppend = 2
        End If
        
        'Loop through the folder, looking for the first "Filename (###).ext" variant that is not already taken.
        Do While Files.FileExists(dstDirectory & tmpFilename & " (" & CStr(numToAppend) & ")" & "." & desiredExtension)
            numToAppend = numToAppend + 1
        Loop
            
        'If the loop has terminated, a unique filename has been found.  Make that the recommended filename.
        IncrementFilename = tmpFilename & " (" & CStr(numToAppend) & ")"
        
    End If

End Function

'Request a temporary filename.  The filename will automatically be added to PD's internal cache, and deleted when
' PD exits.  The caller *can* delete the file if they want, but it is not necessary (as PD automatically clears all
' cached filenames at shutdown time).
Public Function RequestTempFile() As String

    If (m_NumOfTempFiles = 0) Then
        ReDim m_ListOfTempFiles(0 To INIT_TEMP_FILE_CACHE - 1) As String
    Else
        If (m_NumOfTempFiles > UBound(m_ListOfTempFiles)) Then ReDim Preserve m_ListOfTempFiles(0 To m_NumOfTempFiles * 2 - 1) As String
    End If
    
    Dim tmpFile As String
    tmpFile = OS.UniqueTempFilename()
    
    m_ListOfTempFiles(m_NumOfTempFiles) = tmpFile
    m_NumOfTempFiles = m_NumOfTempFiles + 1
    RequestTempFile = tmpFile

End Function

'Execute another program (in PhotoDemon's case, a plugin), then wait for it to finish running.
Public Function ShellAndWait(ByVal executablePath As String, Optional ByVal commandLineArguments As String = vbNullString, Optional ByVal showAppWindow As Boolean = False) As Boolean
    
    Dim startInfo As WIN32_STARTUP_INFO, procInfo As WIN32_PROCESS_INFORMATION
    With startInfo
        .cb = Len(startInfo)
        .dwFlags = STARTF_USESHOWWINDOW
        If showAppWindow Then .wShowWindow = SW_NORMAL Else .wShowWindow = SW_HIDE
    End With
    
    'Null strings are problematic here; generate pointers manually to account for their possible existence
    Dim ptrToExePath As Long, ptrToCmdArgs As Long
    If (Len(executablePath) <> 0) Then ptrToExePath = StrPtr(executablePath) Else ptrToExePath = 0
    If (Len(commandLineArguments) <> 0) Then ptrToCmdArgs = StrPtr(commandLineArguments) Else ptrToCmdArgs = 0
    
    If (CreateProcessW(ptrToExePath, ptrToCmdArgs, ByVal 0&, ByVal 0&, 1&, 0&, ByVal 0&, 0&, VarPtr(startInfo), VarPtr(procInfo)) <> 0) Then
        
        'Get a process handle from the returned ID
        If (procInfo.hProcess <> 0) Then
            ShellAndWait = (WaitForSingleObject(procInfo.hProcess, WAIT_INFINITE) <> &HFFFFFFFF)
            CloseHandle procInfo.hProcess
        End If
    
    Else
        Debug.Print "WARNING!  ShellAndWait failed to create target process: " & executablePath
    End If
    
End Function

'Some functions are just thin wrappers to pdFSO.  (This is desired behavior, as pdFSO uses some internal caches that allow for greater
' flexibility than a module-level implementation.)
'
'For non-cached functions, however, we can simply wrap a pdFSO instance for consistent results.
Private Function InitializeFSO() As Boolean
    If (m_FSO Is Nothing) Then Set m_FSO = New pdFSO
    InitializeFSO = True
End Function

Public Function FileCreateFromByteArray(ByRef srcArray() As Byte, ByVal pathToFile As String, Optional ByVal overwriteExistingIfPresent As Boolean = True, Optional ByVal fileIsTempFile As Boolean = False) As Boolean
    If InitializeFSO Then FileCreateFromByteArray = m_FSO.FileCreateFromByteArray(srcArray, pathToFile, overwriteExistingIfPresent, fileIsTempFile)
End Function

Public Function FileCreateFromPtr(ByVal ptrSrc As Long, ByVal dataLength As Long, ByVal pathToFile As String, Optional ByVal overwriteExistingIfPresent As Boolean = True, Optional ByVal fileIsTempFile As Boolean = False) As Boolean
    If InitializeFSO Then FileCreateFromPtr = m_FSO.FileCreateFromPtr(ptrSrc, dataLength, pathToFile, overwriteExistingIfPresent, fileIsTempFile)
End Function

Public Function FileDelete(ByRef srcFile As String) As Boolean
    If InitializeFSO Then FileDelete = m_FSO.FileDelete(srcFile)
End Function

'Thin wrapper around FileExists() and FileDelete().  Status is not returned, by design; if you want that, you should be calling those
' functions directly so you can deal with individual failure possibilities.
'
'TODO: look at adding a possible "rollback" option, where we cache the deleted file contents at module-level, then allow the user
' to restore it via a matching RestoreLastDelete() kinda thing.  (For some things, like failed image saves, this is preferable
' to deleting the target file, then doing nothing if the export mysteriously fails.)
Public Sub FileDeleteIfExists(ByRef srcFile As String)
    If InitializeFSO Then
        If m_FSO.FileExists(srcFile) Then m_FSO.FileDelete srcFile
    End If
End Sub

Public Function FileExists(ByRef srcFile As String) As Boolean
    If InitializeFSO Then FileExists = m_FSO.FileExists(srcFile)
End Function

Public Function FileGetExtension(ByRef srcFile As String) As String
    If InitializeFSO Then FileGetExtension = m_FSO.FileGetExtension(srcFile)
End Function

Public Function FileGetName(ByRef srcPath As String, Optional ByVal stripExtension As Boolean = False) As String
    If InitializeFSO Then FileGetName = m_FSO.FileGetName(srcPath, stripExtension)
End Function

Public Function FileGetPath(ByRef srcPath As String) As String
    If InitializeFSO Then FileGetPath = m_FSO.FileGetPath(srcPath)
End Function

Public Function FileGetTimeAsDate(ByRef srcFile As String, Optional ByVal typeOfTime As PD_FILE_TIME = PDFT_CreateTime) As Date
    If InitializeFSO Then FileGetTimeAsDate = m_FSO.FileGetTimeAsDate(srcFile, typeOfTime)
End Function

Public Function FileLenW(ByRef srcPath As String) As Long
    If InitializeFSO Then FileLenW = m_FSO.FileLenW(srcPath)
End Function

Public Function FileLoadAsByteArray(ByRef srcFile As String, ByRef dstArray() As Byte) As Boolean
    If InitializeFSO Then FileLoadAsByteArray = m_FSO.FileLoadAsByteArray(srcFile, dstArray)
End Function

Public Function FileLoadAsPDStream(ByRef srcFile As String, ByRef dstStream As pdStream) As Boolean
    If InitializeFSO Then FileLoadAsPDStream = m_FSO.FileLoadAsPDStream(srcFile, dstStream)
End Function

Public Function FileLoadAsString(ByRef srcFile As String, ByRef dstString As String, Optional ByVal forceWindowsLineEndings As Boolean = True) As Boolean
    If InitializeFSO Then FileLoadAsString = m_FSO.FileLoadAsString(srcFile, dstString, forceWindowsLineEndings)
End Function

Public Function FileMakeNameValid(ByRef srcFilename As String, Optional ByVal replacementChar As String = "_") As String
    If InitializeFSO Then FileMakeNameValid = m_FSO.MakeValidWindowsFilename(srcFilename, replacementChar)
End Function

Public Function FileReplace(ByVal oldFile As String, ByVal newFile As String, Optional ByVal customBackupFile As String = vbNullString) As PD_FILE_PATCH_RESULT
    If InitializeFSO Then FileReplace = m_FSO.FileReplace(oldFile, newFile, customBackupFile)
End Function

Public Function FileSaveAsText(ByRef srcString As String, ByRef dstFilename As String, Optional ByVal useUTF8 As Boolean = True, Optional ByVal useUTF8_BOM As Boolean = True) As Boolean
    If InitializeFSO Then FileSaveAsText = m_FSO.FileSaveAsText(srcString, dstFilename, useUTF8, useUTF8_BOM)
End Function

Public Function FileTestAccess_Read(ByVal srcFile As String, Optional ByRef dstLastDLLError As Long) As Boolean
    If InitializeFSO Then FileTestAccess_Read = m_FSO.FileTestAccess_Read(srcFile, dstLastDLLError)
End Function

Public Function FileTestAccess_Write(ByVal srcFile As String, Optional ByRef dstLastDLLError As Long) As Boolean
    If InitializeFSO Then FileTestAccess_Write = m_FSO.FileTestAccess_Write(srcFile, dstLastDLLError)
End Function

Public Function PathAddBackslash(ByRef srcPath As String) As String
    If InitializeFSO Then PathAddBackslash = m_FSO.PathAddBackslash(srcPath)
End Function

'Public Function PathBrowseDialog(ByVal srcHwnd As Long, Optional ByVal initFolder As String = vbNullString, Optional ByVal dialogTitleText As String = vbNullString) As String
'    If InitializeFSO Then PathBrowseDialog = m_FSO.PathBrowseDialog(srcHwnd, initFolder, dialogTitleText)
'End Function

Public Function PathCompact(ByRef srcString As String, ByVal newMaxLength As Long) As String
    If InitializeFSO Then PathCompact = m_FSO.PathCompact(srcString, newMaxLength)
End Function

Public Function PathCreate(ByVal fullPath As String, Optional ByVal createIntermediateFoldersAsNecessary As Boolean = False) As Boolean
    If InitializeFSO Then PathCreate = m_FSO.PathCreate(fullPath, createIntermediateFoldersAsNecessary)
End Function

Public Function PathExists(ByRef fullPath As String, Optional ByVal checkWriteAccessAsWell As Boolean = True) As Boolean
    If InitializeFSO Then PathExists = m_FSO.PathExists(fullPath, checkWriteAccessAsWell)
End Function

Public Function RetrieveAllFiles(ByVal srcFolder As String, ByRef dstFiles As pdStringStack, Optional ByVal recurseSubfolders As Boolean, Optional ByVal returnRelativeStrings As Boolean = True, Optional ByVal onlyAllowTheseExtensions As String = vbNullString, Optional ByVal doNotAllowTheseExtensions As String = vbNullString) As Boolean
    If InitializeFSO Then RetrieveAllFiles = m_FSO.RetrieveAllFiles(srcFolder, dstFiles, recurseSubfolders, returnRelativeStrings, onlyAllowTheseExtensions, doNotAllowTheseExtensions)
End Function
