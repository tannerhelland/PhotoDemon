VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdFSO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon File System Object Interface
'Copyright 2014-2015 by Tanner Helland
'Created: 04/February/15
'Last updated: 05/February/15
'Last update: ongoing work on folder and file iterators
'Dependencies: pdStringStack (used internally for performance-friendly string collection management)
'
'This class is currently under heavy construction.
'
'The end goal is to provide convenient, Unicode-friendly replacements for VB's various file and folder functions.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Const MAX_PATH = 260

Private Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As Currency
   ftLastAccessTime As Currency
   ftLastWriteTime As Currency
   nFileSizeBig As Currency
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

Private Declare Function FindFirstFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal lpFindFileData As Long) As Long
Private Declare Function FindNextFileW Lib "kernel32" (ByVal hFindFile As Long, ByVal lpFindFileData As Long) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const ERROR_FILE_NOT_FOUND As Long = 2
Private Const ERROR_PATH_NOT_FOUND As Long = 3
Private Const ERROR_NO_MORE_FILES As Long = 18

'ReplaceFile is a handy API that combines the following steps into one smooth procedure:
' Saving data to a new file, renaming the original file using a temporary name, renaming the new file to have the same name as the
' original file, then deleting the original file.  As a nice bonus, it also preserves DACLs, object ID, security descriptors, and various
' other attributes.  This makes it perfect for in-place file updating.
Private Declare Function ReplaceFileW Lib "kernel32" (ByVal lpReplacedFileName As Long, ByVal lpReplacementFileName As Long, ByVal lpBackupFileName As Long, ByVal dwReplaceFlags As Long, ByVal lpExclude As Long, ByVal lpReserved As Long) As Long
Private Const REPLACEFILE_IGNORE_MERGE_ERRORS As Long = &H2

'ReplaceFileW returns a non-zero value if successful, or zero if it fails.  GetLastError can be used to ascertain specific failures, including:
Private Const ERROR_UNABLE_TO_MOVE_REPLACEMENT As Long = &H498  'The replacement file could not be renamed. If lpBackupFileName was specified, the replaced
                                                                ' and replacement files retain their original file names. Otherwise, the replaced file no
                                                                ' longer exists and the replacement file exists under its original name.
Private Const ERROR_UNABLE_TO_MOVE_REPLACEMENT_2 As Long = &H499    'The replacement file could not be moved. The replacement file still exists under its
                                                                    ' original name; however, it has inherited the file streams and attributes from the file
                                                                    ' it is replacing. The file to be replaced still exists with a different name. If
                                                                    ' lpBackupFileName is specified, it will be the name of the replaced file.
Private Const ERROR_UNABLE_TO_REMOVE_REPLACED As Long = &H497       'The replaced file could not be deleted. The replaced and replacement files retain their
                                                                    ' original file names.

'ReplaceFile has some rare (but not impossible) failure states that could leave us in a nasty lurch.  It also fails if the target file does not exist.
' We can work around some of these states with a MoveFile fallback.
Private Declare Function MoveFileExW Lib "kernel32" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long, ByVal dwFlags As Long) As Long
Private Const MOVEFILE_REPLACE_EXISTING As Long = &H1   'If a file named lpNewFileName exists, the function replaces its contents with the contents of the
                                                        ' lpExistingFileName file, provided that security requirements regarding access control lists (ACLs)
                                                        ' are met. (This value cannot be used if lpNewFileName or lpExistingFileName names a directory.)
Private Const MOVEFILE_COPY_ALLOWED As Long = &H2   'If the file is to be moved to a different volume, the function simulates the move by using the CopyFile
                                                    ' and DeleteFile functions. If the file is successfully copied to a different volume and the original file
                                                    ' is unable to be deleted, the function succeeds leaving the source file intact.

'Because our file patching code affects critical PhotoDemon files, we need to make sure we return detailed success/failure information.
Public Enum FILE_PATCH_RESULT
    FPR_SUCCESS = 0
    FPR_FAIL_NOTHING_CHANGED = 1
    FPR_FAIL_OLD_FILE_REMOVED = 2
    FPR_FAIL_NEW_FILE_REMOVED = 3
    FPR_FAIL_BOTH_FILES_REMOVED = 4
End Enum

#If False Then
Private Const FPR_SUCCESS = 0, FPR_FAIL_NOTHING_CHANGED = 1, FPR_FAIL_OLD_FILE_REMOVED = 2, FPR_FAIL_NEW_FILE_REMOVED = 3, FPR_FAIL_BOTH_FILES_REMOVED = 4
#End If

'Folder creation.  NULL can be passed as the security attributes pointer; this results in default permissions.
Private Declare Function CreateDirectoryW Lib "kernel32" (ByVal lpPathName As Long, ByVal ptrToSecurityAttributes As Long) As Long

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

'Used to quickly check if a file (or folder) exists.  Thanks to Bonnie West's "Optimum FileExists Function"
' for this technique: http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=74264&lngWId=1
Private Const ERROR_SHARING_VIOLATION As Long = 32
Private Declare Function GetFileAttributesW Lib "kernel32" (ByVal lpFileName As Long) As Long

'Our Unicode-friendly DirW function may be iterated many many times.  To prevent the reallocation of huge WIN32_FIND_DATA structs,
' we use a single module-level entry.
Private m_FileDataReturn As WIN32_FIND_DATA
Private m_FileDataReturnPtr As Long
Private m_SearchHandle As Long

'Returns a VB boolean indicating whether a given file exists.  This should also work on system files that prevent direct access;
' the ERROR_SHARING_VIOLATION check below is meant to capture such files.
' (IMPORTANT NOTE: wildcards are not supported by this function.)
Public Function FileExist(ByRef fName As String) As Boolean
    Select Case (GetFileAttributesW(StrPtr(fName)) And vbDirectory) = 0
        Case True: FileExist = True
        Case Else: FileExist = (Err.LastDllError = ERROR_SHARING_VIOLATION)
    End Select
End Function

'Returns a boolean as to whether or not a given directory exists.  Write access can optionally be tested, too.
Public Function FolderExist(ByRef fullPath As String, Optional ByVal checkWriteAccessAsWell As Boolean = True) As Boolean
    
    'First, make sure the directory exists
    Dim chkExistence As Boolean
    chkExistence = Abs(GetFileAttributesW(StrPtr(fullPath))) And vbDirectory
        
    'If the caller doesn't care about write access, or the folder wasn't found, exit now.
    If (Not chkExistence) Or (Not checkWriteAccessAsWell) Then
        FolderExist = chkExistence
        
    'The folder was found, but the caller wants to know about write access, so a few more checks are needed.
    Else
        
        'Writing to an inaccessible folder will raise an error; this is an ugly way to check, but security descriptors are a hassle...
        On Error GoTo noWriteAccess
        
        Dim tmpFilename As String
        tmpFilename = EnforcePathSlash(fullPath) & "tmp.tmp"
        
        Dim fileNum As Integer
        fileNum = FreeFile
    
        'Attempt to create a file within this directory.  If we succeed, delete the file and return "true".
        ' If we fail, we do not have access rights.
        Open tmpFilename For Binary As #fileNum
            Put #fileNum, 1, 1
        Close #fileNum
        
        If FileExist(tmpFilename) Then Kill tmpFilename
        
        'If we made it all the way here, the file was written and deleted successfully.
        FolderExist = True
        
    End If
    
    Exit Function
    
noWriteAccess:

    FolderExist = False
    
End Function

'Create a given directory.  Returns TRUE if folder already exists, or folder was successfully created.  FALSE otherwise.
'
'An optional parameter, "createIntermediateFoldersAsNecessary", can be specified.  This instructs the function to create as many
' missing subfolders as are necessary in order to ensure the full source path exists.
Public Function CreateFolder(ByVal fullPath As String, Optional ByVal createIntermediateFoldersAsNecessary As Boolean = False) As Boolean
    
    'If the folder already exists, great!
    If FolderExist(fullPath, False) Then
        CreateFolder = True
        
    'The folder doesn't exist, so we need to make it
    Else
    
        Dim cSuccess As Long
        cSuccess = CreateDirectoryW(StrPtr(fullPath), 0&)
        
        If (cSuccess <> 0) Then
            CreateFolder = True
        Else
            CreateFolder = False
            
            'Check relevant errors
            
            'PATH NOT FOUND means one or more parent folders do not exist.
            If Err.LastDllError = ERROR_PATH_NOT_FOUND Then
            
                'If the user wants us to construct intermediate folders, this is our chance to do so.
                If createIntermediateFoldersAsNecessary Then
                    
                    'Starting at the base and working our way forward, create all necessary folders.
                    
                    'First, normalize against backslashes
                    If InStr(1, fullPath, "/", vbBinaryCompare) Then fullPath = Replace$(fullPath, "/", "\")
                    fullPath = EnforcePathSlash(fullPath)
                    
                    'Assume a drive letter is present.  We'll use that as our starting point.
                    Dim startPos As Long, endPos As Long
                    startPos = InStr(1, fullPath, ":", vbBinaryCompare)
                    
                    'Find the first "\"
                    startPos = InStr(startPos, fullPath, "\")
                    
                    If startPos > 0 Then
                    
                        'Extract the next folder in line
                        endPos = InStr(startPos + 1, fullPath, "\", vbBinaryCompare)
                        
                        Do While endPos > 0
                            
                            'We're going to ignore return values during creation of these child folders, because permissions may not exist
                            ' for creating some parent folders.  Instead, to verify that everything worked, we'll re-test the full folder
                            ' existence after attempting to create the entire hierarchy.
                            CreateFolder Left$(fullPath, endPos), False
                            
                            'Find the next folder in the path and repeat!
                            startPos = endPos
                            endPos = InStr(startPos + 1, fullPath, "\", vbBinaryCompare)
                            
                        Loop
                        
                        'Now we check to see if the full path exists; if it doesn't, we're SOL.
                        CreateFolder = FolderExist(fullPath, True)
                    
                    Else
                        Debug.Print "WARNING! CreateFolder failed; no folders were specified in the path parameter...?"
                        CreateFolder = False
                    End If
                    
                Else
                    Debug.Print "WARNING! CreateFolder failed; intermediate folders were missing."
                End If
            
            Else
                Debug.Print "WARNING! CreateFolder failed on an unknown error (#" & Err.LastDllError & ")."
            End If
            
        End If
        
    End If

End Function

'Given a base folder, return all files within that folder (including subfolders).  Subfolder recursion is assumed, but can be waived by setting
' recurseSubfolders to FALSE.
'
'If returnRelativeStrings is true, strings are (obviously) returned relative to the base folder.  So for e.g. base folder "C:\Folder",
' "C:\Folder\Subfolder\file.txt" will be returned as "Subfolder\file.txt".
'
'Returns TRUE if at least one file is found; FALSE otherwise.  If the incoming dstFiles parameter already contains strings, TRUE will
' always be returned.
'
'As an additional convenience, files can be restricted in one of two ways:
' 1) by restricting allowed extensions, using a pipe-delimited list of acceptable extensions (e.g. "jpg|bmp|gif")
' 2) by avoiding disallowed extensions, using a pipe-delimited list of unacceptable extensions (e.g. "bak|tmp")
'
'Use either the whitelist or the blacklist option, but not both (obviously).
Public Function retrieveAllFiles(ByVal srcFolder As String, ByRef dstFiles As pdStringStack, Optional ByVal recurseSubfolders As Boolean, Optional ByVal returnRelativeStrings As Boolean = True, Optional ByVal onlyAllowTheseExtensions As String = "", Optional ByVal doNotAllowTheseExtensions As String = "") As Boolean
    
    'Enforce strict trailing slash formatting of the base folder
    srcFolder = EnforcePathSlash(srcFolder)
    
    'Initialize the destination stack as necessary.  Note that nothing happens if dstFiles is already initialized; this is by design, so the caller
    ' can concatenate multiple search results together if desired.
    If dstFiles Is Nothing Then Set dstFiles = New pdStringStack
    
    'This function was first used in PD as part of pdPackage's zip-like interface.  The goal was to create a convenient way
    ' to generate a folder-preserved list of files.  Consider a file tree like the following:
    ' C:\Folder\file.txt
    ' C:\Folder\SubFolder\subfile1.txt
    ' C:\Folder\SubFolder\AnotherFolder\subfile2.txt
    '
    'By calling this function with "C:\Folder\" as the base, this function can return an array with the following entries:
    ' file.txt
    ' SubFolder\subfile1.txt
    ' SubFolder\AnotherFolder\subfile2.txt
    '
    'This structure makes it very easy to duplicate a full file and folder structure in a new directory, which is exactly
    ' what pdPackage does when behaving like a .zip container.
    '
    'To that end, this function uses DirW to assemble a list of files relative to some base folder.  Subfolders are not explicitly
    ' returned; the hope is that you can implicitly determine them from the relative filenames provided.  If you want a list of
    ' subfolders, make a separate call to retrieveAllFolders, below.
    '
    'Finally, two optional pattern parameters are available.  Use one or the other (or neither), but not both.
    '
    'onlyAllowTheseExtensions: used when the set of desired files uses a small subset of extensions.  Separate valid extensions
    ' by pipe, e.g. "exe|txt".  Do not include ".".  If non-extension files are desired, YOU CANNOT USE THIS PARAMETER, as "||" doesn't
    ' parse correctly.
    '
    'doNotAllowTheseExtensions: used to blacklist unwanted file types.  Same rules as onlyAllowTheseExtensions applies, e.g.
    ' "bak|tmp" would be used to exclude .bak and .tmp files.
    
    'Because white/blacklisting is computationally expensive, we prepare separate boolean checks in advance
    Dim whiteListInUse As Boolean, blackListInUse As Boolean
    whiteListInUse = (Len(onlyAllowTheseExtensions) <> 0)
    blackListInUse = (Len(doNotAllowTheseExtensions) <> 0)
    
    'To further simplify searching for extension matches in the whitelist, we enforce strict formatting of the string
    If whiteListInUse Then
        If Not StringsEqual(Left$(onlyAllowTheseExtensions, 1), "|") Then onlyAllowTheseExtensions = "|" & onlyAllowTheseExtensions
        If Not StringsEqual(Right$(onlyAllowTheseExtensions, 1), "|") Then onlyAllowTheseExtensions = onlyAllowTheseExtensions & "|"
    End If
    
    If blackListInUse Then
        If Not StringsEqual(Left$(doNotAllowTheseExtensions, 1), "|") Then doNotAllowTheseExtensions = "|" & doNotAllowTheseExtensions
        If Not StringsEqual(Right$(doNotAllowTheseExtensions, 1), "|") Then doNotAllowTheseExtensions = doNotAllowTheseExtensions & "|"
    End If
    
    'To eliminate the need for expensive recursion, we use a stack structure to store subfolders that still need to be searched.
    ' Valid files are directly added to dstFiles as they are encountered.
    Dim cFoldersToCheck As pdStringStack
    Set cFoldersToCheck = New pdStringStack
    
    'Add the base folder to the collection, then start searching for subfolders.
    cFoldersToCheck.AddString srcFolder
    
    Dim curFolder As String, chkFile As String
    Dim fileValid As Boolean
        
    'The first folder doesn't get added to the destination collection, but all other folders do.
    Dim isNotFirstFolder As Boolean
    isNotFirstFolder = False
    
    Do While cFoldersToCheck.PopString(curFolder)
        
        'DirW does all the heavy lifting on the actual folder iteration steps
        chkFile = DirW(curFolder)
    
        Do While Len(chkFile) <> 0
            
            'See if this chkFile iteration contains a folder
            If (GetAttributesOfLastDirWReturn And vbDirectory) <> 0 Then
            
                'This is a folder.  Add it to the "folders to check" collection if subfolders are being parsed.
                If recurseSubfolders Then cFoldersToCheck.AddString EnforcePathSlash(curFolder & chkFile)
                
            'This is not a folder, but a file.  Add it to the destination file list if it meets any white/blacklisted criteria.
            Else
                
                'White-listing check required
                If whiteListInUse Then
                    
                    'Look for this file's extension in the white list
                    If InStr(1, onlyAllowTheseExtensions, "|" & GetFileExtension(chkFile) & "|", vbBinaryCompare) > 0 Then
                        fileValid = True
                    Else
                        fileValid = False
                    End If
                
                'Black-listing check required
                ElseIf blackListInUse Then
                
                    'Look for this file's extension in the black list
                    If InStr(1, doNotAllowTheseExtensions, "|" & GetFileExtension(chkFile) & "|", vbBinaryCompare) > 0 Then
                        fileValid = False
                    Else
                        fileValid = True
                    End If
                
                'All files are allowed
                Else
                    fileValid = True
                End If
                
                'If we are allowed to add this file, do so now
                If fileValid Then
                    If returnRelativeStrings Then
                        dstFiles.AddString GenerateRelativePath(srcFolder, curFolder & chkFile, True)
                    Else
                        dstFiles.AddString curFolder & chkFile
                    End If
                End If
            
            End If
            
            'Get the next file
            chkFile = DirW()
        
        Loop
        
        'If recursion is enabled, any subfolders in this folder have now been added to cFolderToCheck, while all files in this folder
        ' are already present in the dstFiles collection.
        
    Loop
    
    'dstFiles now contains a full collection of files with the given base folder (and subfolders, if recursion is enabled).
    ' If at least one file was found, return TRUE. (Note that this value might be incorrect if the user sent us an already-populated
    ' string stack, but that's okay - the assumption is that they'll be processing the stack as one continuous list, so this return
    ' value won't be relevant anyway.)
    If dstFiles.getNumOfStrings > 0 Then
        retrieveAllFiles = True
    Else
        retrieveAllFiles = False
    End If
    

End Function

'Given a base folder, return all subfolders.  Recursion is assumed, but can be waived by setting recurseSubfolders to FALSE.
'
'If returnRelativeStrings is true, strings are (obviously) returned relative to the base folder.  So for e.g. base folder "C:\Folder",
' "C:\Folder\Subfolder" will be returned as just "Subfolder", while "C:\Folder\Subfolder\etc\" will return as "Subfolder\etc".
'
'Returns TRUE if at least one subfolder is found; FALSE otherwise.  If the incoming dstFolders parameter already contains strings, TRUE will
' always be returned.
Public Function retrieveAllFolders(ByVal srcFolder As String, ByRef dstFolders As pdStringStack, Optional ByVal recurseSubfolders As Boolean = True, Optional ByVal returnRelativeStrings As Boolean = True) As Boolean

    'Enforce strict trailing slash formatting of the base folder
    srcFolder = EnforcePathSlash(srcFolder)
    
    'Initialize dstStrings as necessary.  Note that nothing happens if dstStrings is already initialized; this is by design, so the caller
    ' can concatenate multiple results together if desired.
    If dstFolders Is Nothing Then Set dstFolders = New pdStringStack
    
    'To eliminate the need for expensive recursion, we use a stack structure to store subfolders that still need to be searched.
    Dim cFoldersToCheck As pdStringStack
    Set cFoldersToCheck = New pdStringStack
    
    'Add the base folder to the collection, then start searching for subfolders.
    cFoldersToCheck.AddString srcFolder
    
    Dim curFolder As String, chkFile As String
    
    'The first folder doesn't get added to the destination collection, but all other folders do.
    Dim isNotFirstFolder As Boolean
    isNotFirstFolder = False
    
    Do While cFoldersToCheck.PopString(curFolder)
        
        'DirW does all the heavy lifting on the actual folder iteration steps
        chkFile = DirW(curFolder)
    
        Do While Len(chkFile) <> 0
            
            'See if this chkFile iteration contains a folder
            If (GetAttributesOfLastDirWReturn And vbDirectory) <> 0 Then
            
                'This is a folder.  Add it to the "folders to check" collection.
                If recurseSubfolders Then
                    cFoldersToCheck.AddString EnforcePathSlash(curFolder & chkFile)
                Else
                    If returnRelativeStrings Then
                        dstFolders.AddString GenerateRelativePath(srcFolder, EnforcePathSlash(curFolder & chkFile), True)
                    Else
                        dstFolders.AddString EnforcePathSlash(curFolder & chkFile)
                    End If
                End If
            
            End If
            
            'Get the next file
            chkFile = DirW()
        
        Loop
        
        'Any subfolders in this folder have now been added to cFolderToCheck.  With this folder successfully processed, we can move it to
        ' the "folders checked" stack.
        If isNotFirstFolder Then
            
            If returnRelativeStrings Then
                dstFolders.AddString GenerateRelativePath(srcFolder, EnforcePathSlash(curFolder & chkFile), True)
            Else
                dstFolders.AddString curFolder
            End If
            
        Else
            isNotFirstFolder = True
        End If
        
    Loop
    
    'dstFolders now contains a full collection of subfolders for the given base folder.  If subfolders were found, return TRUE.
    ' (Note that this value might be incorrect if the user sent us an already-populated string stack, but that's okay - the assumption is that
    ' they'll be processing the stack as one continuous list, so individual returns don't matter.)
    If dstFolders.getNumOfStrings > 0 Then
        retrieveAllFolders = True
    Else
        retrieveAllFolders = False
    End If
    
End Function

'Given a base folder and some other path (with or without filename), generate a relative path between the two.
' It's assumed that thisFolder contains baseFolder within its path; if it doesn't, a copy of the full thisFolder string is returned.
Public Function GenerateRelativePath(ByVal baseFolder As String, ByVal thisFolder As String, Optional ByVal normalizationCanBeSkipped As Boolean = False) As String
    
    'Start by forcing each string to have a trailing path
    If Not normalizationCanBeSkipped Then baseFolder = EnforcePathSlash(baseFolder)
    
    'Make sure a relative path is possible
    If InStr(1, thisFolder, baseFolder) = 1 Then
        
        'Check equality first; equal strings mean no relative folder is required
        If StringsEqual(baseFolder, thisFolder) Then
            GenerateRelativePath = ""
        
        'Strings are not equal, but baseFolder is contained within thisFolder.  Perfect!
        Else
            GenerateRelativePath = Right$(thisFolder, Len(thisFolder) - Len(baseFolder))
        End If
            
    Else
        GenerateRelativePath = thisFolder
    End If

End Function

'Return the .dwFileAttributes parameter of the last DirW() return.  This provides on-demand file attribute access, without any
' performance penalty to the DirW loop.
Public Function GetAttributesOfLastDirWReturn() As Long
    GetAttributesOfLastDirWReturn = m_FileDataReturn.dwFileAttributes
End Function

'Unicode-friendly Dir() replacement.  Original version developed by vbForums user "Elroy"
' (http://www.vbforums.com/showthread.php?736735-How-to-mamage-files-with-non-English-names&p=4779595&viewfull=1#post4779595)
' ...but heavily modified for use in PD.  Many thanks to Elroy for sharing his code.
Public Function DirW(Optional ByVal sParam As String = "") As String
    
    Dim allFilesFound As Boolean, keepSearching As Boolean, handleJustCreated As Boolean
    Dim retValue As Long
    Dim retString As String
    
    'IMPORTANT NOTE!  Because this function has been designed to work like VB's Dir() function, it has an important caveat:
    ' you should continue to call it until no more files exist.  (Unlike VB's Dir() function, it can't auto-detect when its
    ' caller goes out of scope, so its file handle will remain open.)  As a failsafe, any open file handles will be released
    ' when the class is closed, but it's not ideal to leave search handles open any longer than you need them.

    'Start by applying some modifications to sParam.  FindFirstFile fails under conditions that VB's own Dir() doese not.
    If Len(sParam) > 0 Then
    
        'First, prepend "\\?\" to sParam.  This enables long file paths.
        If Not StringsEqual(Left$(sParam, 4), "\\?\") Then sParam = "\\?\" & sParam
    
        'FindFirstFile fails if the requested path has a trailing slash.  If the user hands us a bare path, assume that they
        ' want to iterate all files and folders within that folder.
        If StringsEqual(Right$(sParam, 1), "\") Or StringsEqual(Right$(sParam, 1), "/") Then
            sParam = sParam & "*"
        End If
        
    End If
    
    'Next, we need to separate our handling into two cases: when a parameter is passed (meaning initiate a search),
    ' vs no parameter (meaning iterate upon the previous search).
    
    'Parameter provided: initiate a new search
    If Len(sParam) Then
        
        'Close any previous searches
        If m_SearchHandle <> 0 Then FindClose m_SearchHandle
        
        'Retrieve the first file in the new search; this returns the search handle we'll use for subsequent searches
        handleJustCreated = True
        m_SearchHandle = FindFirstFileW(StrPtr(sParam), m_FileDataReturnPtr)
        
        'Check for failure.  Failure can occur for multiple reasons: bad inputs, no files meeting the criteria, etc.
        If m_SearchHandle = INVALID_HANDLE_VALUE Then
            
            'No files found is fine, but if the caller screwed up the input path, we want to print some debug info.
            If Err.LastDllError <> ERROR_FILE_NOT_FOUND Then
                Debug.Print "WARNING! DirW was possibly handed a bad path (" & sParam & "). Please investigate."
            End If
            
            Exit Function
        
        End If
        
    End If
      
    'Now it's time to return an actual file to the caller.
    
    'Make sure a valid search handle exists
    If m_SearchHandle <> 0 Then
        
        'Prepare to retrieve the next file.  Some extra work is required to cover the case of ".." and ".", which are
        ' not relevant for PD's purposes.
        allFilesFound = False
        keepSearching = False
                    
        Do
            
            'FindNextFile will return a non-zero value if successful, but in the case of the *first* retrieved file,
            ' we already pulled its info from FindFirstFileW, above.
            If handleJustCreated Then
                retValue = 1
                handleJustCreated = False
            Else
                retValue = FindNextFileW(m_SearchHandle, m_FileDataReturnPtr)
            End If
                
            If retValue <> 0 Then
                
                'If the return value is "." or "..", ignore it and keep looking for the next file
                retString = TrimNull(m_FileDataReturn.cFileName)
                
                If Len(retString) <= 2 Then
                    
                    If StringsEqual(retString, ".") Or StringsEqual(retString, "..") Then
                        keepSearching = True
                    Else
                        keepSearching = False
                    End If
                
                Else
                    keepSearching = False
                End If
                
            Else
                keepSearching = False
                allFilesFound = True
            End If
        
        Loop While keepSearching
        
        'If all files were found, it's time to exist.  (Note that this value is triggered by a 0 return from FindNextFileW,
        ' which can indicate other error states as well - we'll check this momentarily.)
        If allFilesFound Then
            
            'Start by closing the search handle
            FindClose m_SearchHandle
            m_SearchHandle = 0
            
            'Check for unexpected errors
            If Err.LastDllError <> ERROR_NO_MORE_FILES Then
                Debug.Print "WARNING! DirW terminated for a reason other than ERROR_NO_MORE_FILES. Please investigate."
            End If
            
        Else
            DirW = retString
        End If
    
    Else
        Debug.Print "WARNING! DirW tried to iterate a previous search, but no search handle exists.  Please investigate."
    End If
    
    
End Function

'Replace an arbitrary file with another arbitrary file, with optional backup file of the original created during the process.
'
'All files must reside on the same volume.  The target file should exist, but to be safe, I have added some custom code to address instances
' where the target file does not exist (this function will silently fall back to a "Move file" action instead).
'
'IMPORTANT NOTE: the ReplaceFileW function has rare-but-not-impossible failure outcomes where the original file is deleted, but the new file is
'                 not renamed.  Because this failure could be catastrophic during an update process, PD will forcibly create backups of any files
'                 patched via this function, and restore them as necessary.  The "customBackupFile" parameter only exists if you want to specify
'                 a CUSTOM backup filename + location; it does not affect whether or not a backup IS created.  (TL;DR - backups are always created.)
Public Function ReplaceFile(ByVal oldFile As String, ByVal newFile As String, Optional ByVal customBackupFile As String = "") As FILE_PATCH_RESULT
    
    'Because this function is capable of botching an existing PD install if catastrophic failure occurs, error handling must be comprehensive.
    On Error GoTo replaceFailure
    
    Dim backupFile As String
    Dim rfReturn As Long, mfReturn As Long
    
    'Make sure the target file exists.  If it doesn't, we need to use a different set of APIs
    If FileExist(oldFile) Then
    
        'Forcibly specify a backup path if the user didn't provide their own
        If Len(customBackupFile) = 0 Then
        
            'IMPORTANT NOTE!  If the files you're patching are critical, I STRONGLY recommend forcibly creating a backup here.
            ' Otherwise, you risk a scenario where something goes wrong during the replace step, and the original file is gone forever.
            
            'For example, in PhotoDemon I use the standard Data/Updates folder for backups when patching, like this:
            'backupFile = getFilename(oldFile)
            'backupFile = g_UserPreferences.getUpdatePath & backupFile
            
            'This ensure that we always have a way to undo the replace step if something goes wrong
        
        Else
            backupFile = customBackupFile
        End If
        
        'ReplaceFileW returns a non-zero value if successful, or zero if it fails.  Because this function is used to patch PhotoDemon.exe,
        ' there is no margin for failure, so detailed handling of every possible outcome is required.
        rfReturn = ReplaceFileW(StrPtr(oldFile), StrPtr(newFile), StrPtr(backupFile), REPLACEFILE_IGNORE_MERGE_ERRORS, 0&, 0&)
        
        'Success means we can exit immediately
        If rfReturn <> 0 Then
            
            ReplaceFile = FPR_SUCCESS
            Exit Function
        
        'Failure is a problem.  Retrieve the error cause, and do whatever we can to resolve it.
        Else
        
            Select Case Err.LastDllError
            
                'The replacement file could not be renamed. If lpBackupFileName was specified, the replaced and replacement files retain their original
                ' file names. Otherwise, the replaced file no longer exists and the replacement file exists under its original name.
                Case ERROR_UNABLE_TO_MOVE_REPLACEMENT
                    
                    'See if the old file exists
                    If FileExist(oldFile) Then
                    
                        'The old file exists.  Remove the backup file (if any), then exit.  Because the system is in an identical place as it was prior
                        ' to this function being invoked, the caller is free to try again.
                        If FileExist(customBackupFile) Then Kill customBackupFile
                        
                        ReplaceFile = FPR_FAIL_NOTHING_CHANGED
                        Exit Function
                       
                    Else
                    
                        'The old file no longer exists.  Restore the backup file, if at all possible.
                        If FileExist(backupFile) Then
                        
                            'The backup file exists.  Try to restore it to the location of the original file.
                            mfReturn = MoveFileExW(StrPtr(backupFile), StrPtr(oldFile), MOVEFILE_REPLACE_EXISTING Or MOVEFILE_COPY_ALLOWED)
                            
                            'The move succeeded!  Perform a failsafe deletion of the backup file, then notify the caller that the system has been
                            ' restored to its original state.
                            If mfReturn <> 0 Then
                                
                                If FileExist(backupFile) Then Kill backupFile
                                ReplaceFile = FPR_FAIL_NOTHING_CHANGED
                                Exit Function
                            
                            'The move failed.  This should never happen.
                            Else
                                
                                Debug.Print "WARNING!  Impossible outcome 1 occurred in ReplaceFile.  Please investigate."
                                ReplaceFile = FPR_FAIL_OLD_FILE_REMOVED
                                Exit Function
                            
                            End If
                        
                        'The backup file does not exist.  This should never happen.
                        Else
                        
                            Debug.Print "WARNING!  Impossible outcome 2 occurred in ReplaceFile.  Please investigate."
                            ReplaceFile = FPR_FAIL_OLD_FILE_REMOVED
                            Exit Function
                        
                        End If
                    
                    End If
                
                'The replacement file could not be moved. The replacement file still exists under its original name; however, it has inherited the file streams
                ' and attributes from the file it is replacing. The file to be replaced still exists with a different name. If lpBackupFileName is specified,
                ' it will be the name of the replaced file.
                Case ERROR_UNABLE_TO_MOVE_REPLACEMENT_2
                
                    'See if the old file exists; this should never happen.
                    If FileExist(oldFile) Then
                    
                        'The old file exists.  Remove the backup file (if any), then exit.  Because the system is in an identical place as it was prior
                        ' to this function being invoked, the caller is free to try again.
                        If FileExist(customBackupFile) Then Kill customBackupFile
                        
                        Debug.Print "WARNING!  Impossible outcome 3 occurred in ReplaceFile.  Please investigate."
                        ReplaceFile = FPR_FAIL_NOTHING_CHANGED
                        Exit Function
                       
                    Else
                    
                        'The old file no longer exists.  Restore the backup file.  (The backup file is supposedly guaranteed to exist with this error code.)
                        If FileExist(backupFile) Then
                        
                            'The backup file exists.  Try to restore it to the location of the original file.
                            mfReturn = MoveFileExW(StrPtr(backupFile), StrPtr(oldFile), MOVEFILE_REPLACE_EXISTING Or MOVEFILE_COPY_ALLOWED)
                            
                            'The move succeeded!  Perform a failsafe deletion of the backup file, then notify the caller that the system has been
                            ' restored to its original state.
                            If mfReturn <> 0 Then
                                
                                If FileExist(backupFile) Then Kill backupFile
                                ReplaceFile = FPR_FAIL_NOTHING_CHANGED
                                Exit Function
                            
                            'The move failed.  This should never happen.
                            Else
                                
                                Debug.Print "WARNING!  Impossible outcome 4 occurred in ReplaceFile.  Please investigate."
                                ReplaceFile = FPR_FAIL_OLD_FILE_REMOVED
                                Exit Function
                            
                            End If
                        
                        'The backup file does not exist.  This should never happen.
                        Else
                        
                            Debug.Print "WARNING!  Impossible outcome 5 occurred in ReplaceFile.  Please investigate."
                            ReplaceFile = FPR_FAIL_OLD_FILE_REMOVED
                            Exit Function
                        
                        End If
                    
                    End If
                
                'The replaced file could not be deleted. The replaced and replacement files retain their original file names.
                Case ERROR_UNABLE_TO_REMOVE_REPLACED
                
                    'MSDN is unclear on the state of the backup file on this error.  As a failsafe, remove it if it exists.
                    If FileExist(backupFile) Then Kill backupFile
                    ReplaceFile = FPR_FAIL_NOTHING_CHANGED
                    Exit Function
                
                'Other failure states are more problematic.  Defer to the primary error handler, which will try to restore the
                ' original file setup as best as it can.
                Case Else
                    GoTo replaceFailure
            
            End Select
        
        End If
    
    'The file-to-be-replaced does *not* exist.  Use MoveFile instead of ReplaceFile.
    Else
    
        'Use the API to attempt a move.  (Note that a backup file isn't required in this case, as nothing will be deleted unless the
        ' operation is successful.)
        mfReturn = MoveFileExW(StrPtr(oldFile), StrPtr(newFile), MOVEFILE_REPLACE_EXISTING Or MOVEFILE_COPY_ALLOWED)
        
        'The move succeeded!  Exit immediately.
        If mfReturn <> 0 Then
            
            ReplaceFile = FPR_SUCCESS
            Exit Function
        
        'The move failed.  This should not happen (as the target file is known to not exist), but if it does, no harm is done as we haven't
        ' changed the filesystem at all.
        Else
            
            Debug.Print "WARNING!  Impossible outcome 10 occurred in ReplaceFile.  Please investigate."
            ReplaceFile = FPR_FAIL_NOTHING_CHANGED
            Exit Function
        
        End If
    
    End If
    
    Exit Function
    
    
'Arbitrary failure states are problematic.  Basically, our goal is to restore the original file setup as best as we can.  We do this by attempting
' to restore the backup file (if it exists), and applying any renames or moves required to get everything back to the way it was.
replaceFailure:

    'See if the old file exists
    If FileExist(oldFile) Then
    
        'The old file exists.  Remove the backup file (if any), then exit.  Because the system is in an identical place as it was prior
        ' to this function being invoked, the caller is free to try again.
        If FileExist(customBackupFile) Then Kill customBackupFile
        
        If FileExist(newFile) Then
            ReplaceFile = FPR_FAIL_NOTHING_CHANGED
        Else
            Debug.Print "WARNING!  Impossible outcome 6 occurred in ReplaceFile.  Please investigate."
            ReplaceFile = FPR_FAIL_NEW_FILE_REMOVED
        End If
        Exit Function
       
    Else
    
        'The old file no longer exists.  Restore the backup file, if at all possible.
        If FileExist(backupFile) Then
        
            'The backup file exists.  Try to restore it to the location of the original file.
            mfReturn = MoveFileExW(StrPtr(backupFile), StrPtr(oldFile), MOVEFILE_REPLACE_EXISTING Or MOVEFILE_COPY_ALLOWED)
            
            'The move succeeded!  Perform a failsafe deletion of the backup file, then notify the caller that the system has been
            ' restored to its original state.
            If mfReturn <> 0 Then
                
                If FileExist(backupFile) Then Kill backupFile
                If FileExist(newFile) Then
                    ReplaceFile = FPR_FAIL_NOTHING_CHANGED
                Else
                    Debug.Print "WARNING!  Impossible outcome 7 occurred in ReplaceFile.  Please investigate."
                    ReplaceFile = FPR_FAIL_NEW_FILE_REMOVED
                End If
                Exit Function
            
            'The move failed.  This should never happen.
            Else
                
                Debug.Print "WARNING!  Impossible outcome 8 occurred in ReplaceFile.  Please investigate."
                ReplaceFile = FPR_FAIL_OLD_FILE_REMOVED
                Exit Function
            
            End If
        
        'The backup file does not exist.  This should never happen.
        Else
        
            Debug.Print "WARNING!  Impossible outcome 9 occurred in ReplaceFile.  Please investigate."
            ReplaceFile = FPR_FAIL_OLD_FILE_REMOVED
            Exit Function
        
        End If
    
    End If

End Function

'Given a path, make sure the right backslash is existant
Public Function EnforcePathSlash(ByRef srcPath As String) As String
    
    If (Not StringsEqual(Right$(srcPath, 1), "\")) And (Not StringsEqual(Right$(srcPath, 1), "/")) Then
        EnforcePathSlash = srcPath & "\"
    Else
        EnforcePathSlash = srcPath
    End If
    
End Function

'Return just the filename from a full path
Public Function getFilename(ByRef sString As String) As String

    Dim i As Long
    
    For i = Len(sString) - 1 To 1 Step -1
        If (Mid$(sString, i, 1) = "/") Or (Mid$(sString, i, 1) = "\") Then
            getFilename = Right$(sString, Len(sString) - i)
            Exit Function
        End If
    Next i
    
    'If we made it here, no slash was found, so presume that the entire incoming string was the filename
    getFilename = sString
    
End Function

'Function to return the extension from a filename.
Public Function GetFileExtension(sFile As String) As String
    
    Dim i As Long
    For i = Len(sFile) To 1 Step -1
    
        'If we find a path before we find an extension, return a blank string
        If StringsEqual(Mid(sFile, i, 1), "\") Or StringsEqual(Mid(sFile, i, 1), "/") Then
            GetFileExtension = ""
            Exit Function
        End If
        
        If StringsEqual(Mid(sFile, i, 1), ".") Then
            GetFileExtension = Right$(sFile, Len(sFile) - i)
            Exit Function
        End If
    Next i
    
    'If we reach this point, no extension was found
    GetFileExtension = ""
            
End Function

'Given a full path+filename string, return only the folder portion.
' Returns a blank string if a folder cannot be extracted.
Public Function GetPathOnly(ByRef sFilename As String) As String
    
    Dim slashPosition As Long
    slashPosition = InStrRev(sFilename, "\", , vbBinaryCompare)
    
    If slashPosition <> 0 Then GetPathOnly = Left$(sFilename, slashPosition)
    
End Function

'When passing file and path strings to WAPI functions, we often have to pre-initialize them to some arbitrary buffer length
' (typically MAX_PATH).  When finished, the string needs to be trimmed to remove unused null chars.
Public Function TrimNull(ByRef origString As String) As String
    
    'Start by double-checking that null chars actually exist in the string
    Dim nullPosition As Long
    nullPosition = InStr(origString, Chr$(0))
    
    'Remove null chars if present; otherwise, return a copy of the input string
    If nullPosition <> 0 Then
       TrimNull = Left$(origString, nullPosition - 1)
    Else
       TrimNull = origString
    End If
    
End Function

'VB's default string comparison function is overly cumbersome.  This wrapper makes string equality checks much tidier.
Private Function StringsEqual(ByRef srcString1 As String, ByRef srcString2 As String) As Boolean
    StringsEqual = (StrComp(srcString1, srcString2, vbBinaryCompare) = 0)
End Function

Private Sub Class_Initialize()
    
    'Search functions require a pointer to a WIN32_FIND_DATA struct.  Generate a persistent pointer now.
    m_FileDataReturnPtr = VarPtr(m_FileDataReturn)
    
End Sub

Private Sub Class_Terminate()
    
    'If an active search is still in progress, terminate it now
    If m_SearchHandle <> 0 Then FindClose m_SearchHandle
    
End Sub
