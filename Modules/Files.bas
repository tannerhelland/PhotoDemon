Attribute VB_Name = "Files"
'***************************************************************************
'Comprehensive wrapper for pdFSO (Unicode file and folder functions)
'Copyright 2001-2026 by Tanner Helland
'Created: 6/12/01
'Last updated: 18/March/22
'Last update: add additional pdFSO function wrapper(s)
'
'The pdFSO class normally provides Unicode file/folder interactions for PhotoDemon.
'
'But sometimes you just want to do something trivial - like see if a file exists - and instantiating
' a full class for this is unnecessarily verbose in VB.  Thus the purpose of this module: to provide
' fast access to pdFSO functions without you needing to worry about the details.
'
'If you need persistent interactions with a file, use pdFSO.  If you need one-off access to a
' particular file-related function, use this module.  (All calls end up at a pdFSO instance eventually.)
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Because our file patching code affects critical PhotoDemon files, we need to make sure we return
' detailed success/failure information.
Public Enum PD_FILE_PATCH_RESULT
    FPR_SUCCESS = 0
    FPR_FAIL_NOTHING_CHANGED = 1
    FPR_FAIL_OLD_FILE_REMOVED = 2
    FPR_FAIL_NEW_FILE_REMOVED = 3
    FPR_FAIL_BOTH_FILES_REMOVED = 4
End Enum

#If False Then
    Private Const FPR_SUCCESS = 0, FPR_FAIL_NOTHING_CHANGED = 1, FPR_FAIL_OLD_FILE_REMOVED = 2, FPR_FAIL_NEW_FILE_REMOVED = 3, FPR_FAIL_BOTH_FILES_REMOVED = 4
#End If

Public Enum PD_FILE_ACCESS_OPTIMIZE
    OptimizeNone = 0
    OptimizeRandomAccess = 1
    OptimizeSequentialAccess = 2
    OptimizeTempFile = 3
End Enum

#If False Then
    Private Const OptimizeNone = 0, OptimizeRandomAccess = 1, OptimizeSequentialAccess = 2, OptimizeTempFile = 3
#End If

Public Enum FILE_POINTER_MOVE_METHOD
    FILE_BEGIN = 0
    FILE_CURRENT = 1
    FILE_END = 2
End Enum

#If False Then
    Private Const FILE_BEGIN = 0, FILE_CURRENT = 1, FILE_END = 2
#End If

Public Enum PD_FILE_TIME
    PDFT_CreateTime = 0
    PDFT_AccessTime = 1
    PDFT_WriteTime = 2
End Enum

#If False Then
    Private Const PDFT_CreateTime = 0, PDFT_AccessTime = 1, PDFT_WriteTime = 2
#End If

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

'Made public in Win 11 24H2 to allow aller to retrieve multiple file attributes; the translation engine was seeing
' random large perf hits post-hibernation or post-reboot when pulling file size and/or last access time, and cutting
' the number of accesses of those properties in half reduces the odds of triggering weird 24H2 issues like that.
Public Type WIN32_FILE_ATTRIBUTES_BASIC
    dwFileAttributes As Long
    ftCreationTime As Currency
    ftLastAccessTime As Currency
    ftLastWriteTime As Currency
    nFileSizeBig As Currency
End Type

'Used to shell an external program, then wait until it completes
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreateProcessW Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As Long, ByRef lpProcessAttributes As Any, ByRef lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByRef lpEnvironment As Any, ByVal lpCurrentDriectory As Long, ByVal lpStartupInfo As Long, ByVal lpProcessInformation As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)

Private Declare Function ILCreateFromPathW Lib "shell32" (ByVal lpFileName As Long) As Long
Private Declare Function SHOpenFolderAndSelectItems Lib "shell32" (ByVal pidlFolder As Long, ByVal cidl As Long, ByVal apidl As Long, ByVal dwFlags As Long) As Long

Private Declare Function StrFormatByteSizeW Lib "shlwapi" (ByVal srcSize As Currency, ByVal ptrDstString As Long, ByVal sizeDstStringInChars As Long) As Long

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
            If (LenB(m_ListOfTempFiles(i)) <> 0) Then Files.FileDeleteIfExists m_ListOfTempFiles(i)
        Next i
        
        m_NumOfTempFiles = 0
    
    End If
    
End Sub

'Given a file size (32-bit), return a properly-formatted string representation.  Uses shlwapi.
Public Function GetFormattedFileSize(ByVal srcSize As Long) As String
    Dim tmpLongLong As Currency
    VBHacks.PutMem4 VarPtr(tmpLongLong), srcSize
    GetFormattedFileSize = String$(64, 0)
    If (StrFormatByteSizeW(tmpLongLong, StrPtr(GetFormattedFileSize), Len(GetFormattedFileSize)) <> 0) Then
        GetFormattedFileSize = Strings.TrimNull(GetFormattedFileSize)
    Else
        GetFormattedFileSize = CStr(srcSize) & " bytes"
    End If
End Function

'Given a file size (64-bit), return a properly-formatted string representation.  Uses shlwapi.
Public Function GetFormattedFileSizeL(ByVal srcSize As Currency) As String
    GetFormattedFileSizeL = String$(64, 0)
    If (StrFormatByteSizeW(srcSize, StrPtr(GetFormattedFileSizeL), Len(GetFormattedFileSizeL)) <> 0) Then
        GetFormattedFileSizeL = Strings.TrimNull(GetFormattedFileSizeL)
    Else
        GetFormattedFileSizeL = CStr(srcSize * 10000) & " bytes"
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
    If (LenB(executablePath) <> 0) Then ptrToExePath = StrPtr(executablePath) Else ptrToExePath = 0
    If (LenB(commandLineArguments) <> 0) Then ptrToCmdArgs = StrPtr(commandLineArguments) Else ptrToCmdArgs = 0
    
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

Public Function AppPathW() As String
    If InitializeFSO Then AppPathW = m_FSO.AppPathW()
End Function

Public Function FileCopyW(ByRef srcFilename As String, ByRef dstFilename As String) As Boolean
    If InitializeFSO Then FileCopyW = m_FSO.FileCopyW(srcFilename, dstFilename)
End Function

Public Function FileCreateFromByteArray(ByRef srcArray() As Byte, ByVal pathToFile As String, Optional ByVal overwriteExistingIfPresent As Boolean = True, Optional ByVal fileIsTempFile As Boolean = False, Optional ByVal sizeOfData As Long = -1) As Boolean
    If InitializeFSO Then FileCreateFromByteArray = m_FSO.FileCreateFromByteArray(srcArray, pathToFile, overwriteExistingIfPresent, fileIsTempFile, sizeOfData)
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
Public Function FileDeleteIfExists(ByRef srcFile As String) As Boolean
    If InitializeFSO Then
        If m_FSO.FileExists(srcFile) Then FileDeleteIfExists = m_FSO.FileDelete(srcFile) Else FileDeleteIfExists = True
    End If
End Function

Public Function FileExists(ByRef srcFile As String) As Boolean
    If InitializeFSO Then FileExists = m_FSO.FileExists(srcFile)
End Function

Public Function FileGetAttributesBasic(ByRef srcFile As String, ByRef dstAttributes As WIN32_FILE_ATTRIBUTES_BASIC) As Boolean
    If InitializeFSO Then FileGetAttributesBasic = m_FSO.FileGetAttributesBasic(srcFile, dstAttributes)
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

Public Function FileGetTimeAsCurrency(ByRef srcFile As String, Optional ByVal typeOfTime As PD_FILE_TIME = PDFT_CreateTime) As Currency
    If InitializeFSO Then FileGetTimeAsCurrency = m_FSO.FileGetTimeAsCurrency(srcFile, typeOfTime)
End Function

Public Function FileGetTimeAsDate(ByRef srcFile As String, Optional ByVal typeOfTime As PD_FILE_TIME = PDFT_CreateTime) As Date
    If InitializeFSO Then FileGetTimeAsDate = m_FSO.FileGetTimeAsDate(srcFile, typeOfTime)
End Function

'Retrieve the version number of an .exe or .dll file.
' The passed "version index" correlates to 0 = Major, 1 = Minor, 2 = Build, 3 = Revision
Public Function FileGetVersionAsLong(ByRef srcFile As String, ByVal versionIndex As Long, Optional ByVal getProductVersionInstead As Boolean = True) As Long
    If InitializeFSO Then
        If (versionIndex >= 0) And (versionIndex <= 3) Then
            Dim lVersion(0 To 3) As Long
            If m_FSO.FileGetVersion(srcFile, lVersion(0), lVersion(1), lVersion(2), lVersion(3), getProductVersionInstead) Then FileGetVersionAsLong = lVersion(versionIndex)
        Else
            PDDebug.LogAction "WARNING: Files.FileGetVersionAsLong was passed a bad version index: " & CStr(versionIndex)
        End If
    End If
End Function

Public Function FileLenW(ByRef srcPath As String) As Long
    If InitializeFSO Then FileLenW = m_FSO.FileLenW(srcPath)
End Function

Public Function FileLenLargeW(ByRef srcPath As String) As Currency
    If InitializeFSO Then FileLenLargeW = m_FSO.FileLenW_Large(srcPath)
End Function

Public Function FileLoadAsByteArray(ByRef srcFile As String, ByRef dstArray() As Byte) As Boolean
    If InitializeFSO Then FileLoadAsByteArray = m_FSO.FileLoadAsByteArray(srcFile, dstArray)
End Function

Public Function FileLoadAsString(ByRef srcFile As String, ByRef dstString As String, Optional ByVal forceWindowsLineEndings As Boolean = True) As Boolean
    If InitializeFSO Then FileLoadAsString = m_FSO.FileLoadAsString(srcFile, dstString, forceWindowsLineEndings)
End Function

Public Function FileMakeNameValid(ByRef srcFilename As String, Optional ByVal replacementChar As String = "_") As String
    If InitializeFSO Then FileMakeNameValid = m_FSO.MakeValidWindowsFilename(srcFilename, replacementChar)
End Function

Public Function FileMove(ByVal oldFile As String, ByVal newFile As String, Optional ByVal delNewFileFirstIfExists As Boolean = False) As Boolean
    If InitializeFSO Then FileMove = m_FSO.FileMove(oldFile, newFile, delNewFileFirstIfExists)
End Function

Public Function FileReplace(ByVal oldFile As String, ByVal newFile As String, Optional ByVal customBackupFile As String = vbNullString) As PD_FILE_PATCH_RESULT
    If InitializeFSO Then FileReplace = m_FSO.FileReplace(oldFile, newFile, customBackupFile)
End Function

Public Function FileSaveAsText(ByRef srcString As String, ByRef dstFilename As String, Optional ByVal useUTF8 As Boolean = True, Optional ByVal useUTF8_BOM As Boolean = True) As Boolean
    If InitializeFSO Then FileSaveAsText = m_FSO.FileSaveAsText(srcString, dstFilename, useUTF8, useUTF8_BOM)
End Function

'Open a new Explorer window and select the file in question
Public Sub FileSelectInExplorer(ByRef srcFile As String)
    If Files.FileExists(srcFile) Then
        Dim pItemIDList As Long
        pItemIDList = ILCreateFromPathW(StrPtr(srcFile))
        If (pItemIDList <> 0) Then
            SHOpenFolderAndSelectItems pItemIDList, 0&, 0&, 0&
            CoTaskMemFree pItemIDList
        End If
    End If
End Sub

Public Function FileTestAccess_Read(ByVal srcFile As String, Optional ByRef dstLastDLLError As Long) As Boolean
    If InitializeFSO Then FileTestAccess_Read = m_FSO.FileTestAccess_Read(srcFile, dstLastDLLError)
End Function

Public Function FileTestAccess_Write(ByVal srcFile As String, Optional ByRef dstLastDLLError As Long) As Boolean
    If InitializeFSO Then FileTestAccess_Write = m_FSO.FileTestAccess_Write(srcFile, dstLastDLLError)
End Function

Public Function PathAddBackslash(ByRef srcPath As String) As String
    If InitializeFSO Then PathAddBackslash = m_FSO.PathAddBackslash(srcPath)
End Function

Public Function PathBrowseDialog(ByVal srcHWnd As Long, Optional ByVal initFolder As String = vbNullString, Optional ByVal dialogTitleText As String = vbNullString) As String
    If InitializeFSO Then PathBrowseDialog = m_FSO.PathBrowseDialog(srcHWnd, initFolder, dialogTitleText)
End Function

Public Function PathCanonicalize(ByVal srcPath As String, ByRef dstPath As String) As Boolean
    If InitializeFSO Then PathCanonicalize = m_FSO.PathCanonicalize(srcPath, dstPath)
End Function

Public Function PathCommonPrefix(ByRef srcPath1 As String, ByRef srcPath2 As String, ByRef dstCommonPrefix As String) As Boolean
    If InitializeFSO Then PathCommonPrefix = m_FSO.PathCommonPrefix(srcPath1, srcPath2, dstCommonPrefix)
End Function

Public Function PathCompact(ByRef srcString As String, ByVal newMaxLength As Long) As String
    If InitializeFSO Then PathCompact = m_FSO.PathCompact(srcString, newMaxLength)
End Function

Public Function PathCreate(ByVal fullPath As String, Optional ByVal createIntermediateFoldersAsNecessary As Boolean = False) As Boolean
    If InitializeFSO Then PathCreate = m_FSO.PathCreate(fullPath, createIntermediateFoldersAsNecessary)
End Function

Public Function PathExists(ByRef fullPath As String, Optional ByVal checkWriteAccessAsWell As Boolean = True) As Boolean
    If InitializeFSO Then PathExists = m_FSO.PathExists(fullPath, checkWriteAccessAsWell)
End Function

Public Function PathGetLargestCommonPrefix(ByRef listOfPaths As pdStringStack, ByRef dstCommonPrefix As String) As Boolean
    If InitializeFSO Then PathGetLargestCommonPrefix = m_FSO.PathGetLargestCommonPrefix(listOfPaths, dstCommonPrefix)
End Function

Public Function PathRebaseListOnNewPath(ByRef srcListOfPaths As pdStringStack, ByRef dstListOfRebasedPaths As pdStringStack, ByRef newBasePath As String) As Boolean
    If InitializeFSO Then PathRebaseListOnNewPath = m_FSO.PathRebaseListOnNewPath(srcListOfPaths, dstListOfRebasedPaths, newBasePath)
End Function

Public Function RetrieveAllFiles(ByVal srcFolder As String, ByRef dstFiles As pdStringStack, Optional ByVal recurseSubfolders As Boolean, Optional ByVal returnRelativeStrings As Boolean = True, Optional ByVal onlyAllowTheseExtensions As String = vbNullString, Optional ByVal doNotAllowTheseExtensions As String = vbNullString) As Boolean
    If InitializeFSO Then RetrieveAllFiles = m_FSO.RetrieveAllFiles(srcFolder, dstFiles, recurseSubfolders, returnRelativeStrings, onlyAllowTheseExtensions, doNotAllowTheseExtensions)
End Function
