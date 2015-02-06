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
Private Const ERROR_NO_MORE_FILES As Long = 18

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
            If (getAttributesOfLastDirWReturn And vbDirectory) <> 0 Then
            
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
                        dstFiles.AddString generateRelativePath(srcFolder, curFolder & chkFile, True)
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
            If (getAttributesOfLastDirWReturn And vbDirectory) <> 0 Then
            
                'This is a folder.  Add it to the "folders to check" collection.
                If recurseSubfolders Then
                    cFoldersToCheck.AddString EnforcePathSlash(curFolder & chkFile)
                Else
                    If returnRelativeStrings Then
                        dstFolders.AddString generateRelativePath(srcFolder, EnforcePathSlash(curFolder & chkFile), True)
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
                dstFolders.AddString generateRelativePath(srcFolder, EnforcePathSlash(curFolder & chkFile), True)
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

'Given a base folder and some other path, generate a relative path between the two.
' It's assumed that thisFolder contains baseFolder within its path; if it doesn't, a copy of the full thisFolder string is returned.
Private Function generateRelativePath(ByVal baseFolder As String, ByVal thisFolder As String, Optional ByVal normalizationCanBeSkipped As Boolean = False) As String
    
    'Start by forcing each string to have a trailing path
    If Not normalizationCanBeSkipped Then
        baseFolder = EnforcePathSlash(baseFolder)
        thisFolder = EnforcePathSlash(thisFolder)
    End If
    
    'Make sure a relative path is possible
    If InStr(1, thisFolder, baseFolder) = 1 Then
        
        'Check equality first; equal strings mean no relative folder is required
        If StringsEqual(baseFolder, thisFolder) Then
            generateRelativePath = ""
        
        'Strings are not equal, but baseFolder is contained within thisFolder.  Perfect!
        Else
            generateRelativePath = Right$(thisFolder, Len(thisFolder) - Len(baseFolder))
        End If
            
    Else
        generateRelativePath = thisFolder
    End If

End Function

'Return the .dwFileAttributes parameter of the last DirW() return.  This provides on-demand file attribute access, without any
' performance penalty to the DirW loop.
Public Function getAttributesOfLastDirWReturn() As Long
    getAttributesOfLastDirWReturn = m_FileDataReturn.dwFileAttributes
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

'Given a path, make sure the right backslash is existant
Public Function EnforcePathSlash(ByRef srcPath As String) As String
    
    If (Not StringsEqual(Right$(srcPath, 1), "\")) And (Not StringsEqual(Right$(srcPath, 1), "/")) Then
        EnforcePathSlash = srcPath & "\"
    Else
        EnforcePathSlash = srcPath
    End If
    
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

'Given a full path+filename string, return only the folder portion
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
