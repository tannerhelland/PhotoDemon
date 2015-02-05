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
'Last updated: 04/February/15
'Last update: started moving various functions from across PD into this single class
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

'Given a base folder, retrieve the names of all files within that folder (and optionally, its subfolders)
Public Function retrieveAllFiles_Relative(ByVal srcFolder As String, ByRef dstArray() As String, Optional ByVal recurseSubfolders As Boolean, Optional ByVal onlyAllowTheseExtensions As String) As Boolean

    'TODO!

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
        If StrComp(Left$(sParam, 4), "\\?\", vbBinaryCompare) <> 0 Then sParam = "\\?\" & sParam
    
        'FindFirstFile fails if the requested path has a trailing slash.  If the user hands us a bare path, assume that they
        ' want to iterate all files and folders within that folder.
        If (StrComp(Right$(sParam, 1), "\", vbBinaryCompare) = 0) Or (StrComp(Right$(sParam, 1), "/", vbBinaryCompare) = 0) Then
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
                retString = Trim$(TrimNull(m_FileDataReturn.cFileName))
                
                If Len(retString) <= 2 Then
                    
                    If (StrComp(retString, ".", vbBinaryCompare) = 0) Or (StrComp(retString, "..", vbBinaryCompare) = 0) Then
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

Private Sub Class_Initialize()
    
    'Search functions require a pointer to a WIN32_FIND_DATA struct.  Generate a persistent pointer now.
    m_FileDataReturnPtr = VarPtr(m_FileDataReturn)
    
End Sub

Private Sub Class_Terminate()
    
    'If an active search is still in progress, terminate it now
    If m_SearchHandle <> 0 Then FindClose m_SearchHandle
    
End Sub
