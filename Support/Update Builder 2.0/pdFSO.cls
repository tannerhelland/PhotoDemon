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
'Copyright 2014-2018 by Tanner Helland
'Created: 04/February/15
'Last updated: 06/July/18
'Last update: rewrite file and folder search code for improved performance
'Dependencies: pdStringStack (used internally for performance-friendly string collection management)
'              Strings module (handles lots of Unicode-related transforms)
'
'This class provides convenient, Unicode-friendly replacements for VB's various file and folder functions.  In some cases, the VB
' equivalents have been directly reproduced.  In others, the functions have been rewritten to make fail states easier to detect
' (e.g. returning Boolean's instead of raising errors).
'
'Pay careful attention to the requirements of each function.  Some functions that operate on bare file handles require that the
' handle was created with certain permissions (e.g. you can't check file size using an hFile unless the hFile was created with
' read access).
'
'Thank you to the many invaluable references I used while constructing this class, particularly:
' - Dana Seaman's UnicodeTutorialVB (http://www.cyberactivex.com/UnicodeTutorialVb.htm)
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Const CREATE_ALWAYS As Long = &H2
Private Const DAY_ZERO_BIAS As Double = 109205# 'Difference between day zero for VB dates and Win32 dates (or #12-30-1899# - #01-01-1601#)
Private Const ERROR_FILE_NOT_FOUND As Long = 2
Private Const ERROR_NO_MORE_FILES As Long = 18
Private Const ERROR_PATH_NOT_FOUND As Long = 3
Private Const ERROR_SHARING_VIOLATION As Long = 32
Private Const ERROR_UNABLE_TO_MOVE_REPLACEMENT As Long = &H498  'The replacement file could not be renamed. If lpBackupFileName was specified, the replaced
                                                                ' and replacement files retain their original file names. Otherwise, the replaced file no
                                                                ' longer exists and the replacement file exists under its original name.
Private Const ERROR_UNABLE_TO_MOVE_REPLACEMENT_2 As Long = &H499    'The replacement file could not be moved. The replacement file still exists under its
                                                                    ' original name; however, it has inherited the file streams and attributes from the file
                                                                    ' it is replacing. The file to be replaced still exists with a different name. If
                                                                    ' lpBackupFileName is specified, it will be the name of the replaced file.
Private Const ERROR_UNABLE_TO_REMOVE_REPLACED As Long = &H497       'The replaced file could not be deleted. The replaced and replacement files retain their
                                                                    ' original file names.

Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80&
Private Const FILE_ATTRIBUTE_TEMPORARY As Long = &H100&
Private Const FILE_FLAG_SEQUENTIAL_SCAN As Long = &H8000000
Private Const FILE_FLAG_RANDOM_ACCESS As Long = &H10000000
Private Const FILE_MAP_WRITE As Long = 2
Private Const FILE_SHARE_READ As Long = &H1
Private Const FILE_SHARE_WRITE As Long = &H2
Private Const FILE_SHARE_DELETE As Long = &H4
Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const INVALID_SET_FILE_POINTER As Long = -1
Private Const MAX_PATH As Long = 260
Private Const MILLISECONDS_PER_DAY As Double = 10000000# * 60# * 60# * 24# / 10000# '10000000 nanoseconds * 60 seconds * 60 minutes * 24 hours / 10000 comes to 86400000 (the 10000 adjusts for fixed point in Currency)
Private Const MOVEFILE_REPLACE_EXISTING As Long = &H1   'If a file named lpNewFileName exists, the function replaces its contents with the contents of the
                                                        ' lpExistingFileName file, provided that security requirements regarding access control lists (ACLs)
                                                        ' are met. (This value cannot be used if lpNewFileName or lpExistingFileName names a directory.)
Private Const MOVEFILE_COPY_ALLOWED As Long = &H2   'If the file is to be moved to a different volume, the function simulates the move by using the CopyFile
                                                    ' and FileDelete functions. If the file is successfully copied to a different volume and the original file
                                                    ' is unable to be deleted, the function succeeds leaving the source file intact.
Private Const OPEN_ALWAYS As Long = &H4
Private Const OPEN_EXISTING As Long = &H3
Private Const PAGE_READWRITE As Long = 4
Private Const REPLACEFILE_IGNORE_MERGE_ERRORS As Long = &H2

'Because our file patching code affects critical PhotoDemon files, we need to make sure we return detailed success/failure information.
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

'Type comments derived from http://allapi.mentalis.org/apilist/FileGetVersionInfo.shtml
' Thank you to those authors for demystifying obscure API calls.
Private Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersionl As Integer     ' e.g. = &h0000 = 0
    dwStrucVersionh As Integer     ' e.g. = &h0042 = .42
    dwFileVersionMSl As Integer    ' e.g. = &h0003 = 3
    dwFileVersionMSh As Integer    ' e.g. = &h0075 = .75
    dwFileVersionLSl As Integer    ' e.g. = &h0000 = 0
    dwFileVersionLSh As Integer    ' e.g. = &h0031 = .31
    dwProductVersionMSl As Integer ' e.g. = &h0003 = 3
    dwProductVersionMSh As Integer ' e.g. = &h0010 = .1
    dwProductVersionLSl As Integer ' e.g. = &h0000 = 0
    dwProductVersionLSh As Integer ' e.g. = &h0031 = .31
    dwFileFlagsMask As Long        ' = &h3F for version "0.42"
    dwFileFlags As Long            ' e.g. VFF_DEBUG Or VFF_PRERELEASE
    dwFileOS As Long               ' e.g. VOS_DOS_WINDOWS16
    dwFileType As Long             ' e.g. VFT_DRIVER
    dwFileSubtype As Long          ' e.g. VFT2_DRV_KEYBOARD
    dwFileDateMS As Long           ' e.g. 0
    dwFileDateLS As Long           ' e.g. 0
End Type

Private Type WAPI_FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

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

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CopyFileW Lib "kernel32" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long, ByVal bFailIfExists As Long) As Long
Private Declare Function CreateDirectoryW Lib "kernel32" (ByVal lpPathName As Long, ByVal ptrToSecurityAttributes As Long) As Long
Private Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingW" (ByVal hFile As Long, ByVal lpFileMappingAttributes As Long, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal ptrToNameString As Long) As Long
Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function DeleteFileW Lib "kernel32" (ByVal lpFileName As Long) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (ByRef lpFileTime As Currency, ByRef lpLocalFileTime As Currency) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindFirstFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal lpFindFileData As Long) As Long
Private Declare Function FindNextFileW Lib "kernel32" (ByVal hFindFile As Long, ByVal lpFindFileData As Long) As Long
Private Declare Function FlushViewOfFile Lib "kernel32" (ByVal lpBaseAddress As Long, ByVal dwNumberOfBytesToFlush As Long) As Long
Private Declare Function GetFileAttributesW Lib "kernel32" (ByVal lpFileName As Long) As Long
Private Declare Function GetFileSizeEx Lib "kernel32" (ByVal hFile As Long, ByRef lpFileSize As Currency) As Long
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, ByRef lpCreationTime As WAPI_FILETIME, ByRef lpLastAccessTime As WAPI_FILETIME, ByRef lpLastWriteTime As WAPI_FILETIME) As Long
Private Declare Function GetModuleFileNameW Lib "kernel32" (ByVal hModule As Long, ByVal ptrToFileNameBuffer As Long, ByVal nSize As Long) As Long
Private Declare Function GetShortPathNameW Lib "kernel32" (ByVal lpLongPath As Long, ByVal lpShortPath As Long, ByVal nBufLen As Long) As Long
Private Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function MoveFileExW Lib "kernel32" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long, ByVal dwFlags As Long) As Long
Private Declare Function OpenFileMapping Lib "kernel32" Alias "OpenFileMappingW" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal ptrToNameString As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal ptrToDstBuffer As Long, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function ReplaceFileW Lib "kernel32" (ByVal lpReplacedFileName As Long, ByVal lpReplacementFileName As Long, ByVal lpBackupFileName As Long, ByVal dwReplaceFlags As Long, ByVal lpExclude As Long, ByVal lpReserved As Long) As Long
Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByRef lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As FILE_POINTER_MOVE_METHOD) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32" (ByVal lpBaseAddress As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal ptrToSourceBuffer As Long, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal ptrToOverlappedStruct As Long) As Long

Private Declare Function PathCompactPathEx Lib "shlwapi" Alias "PathCompactPathExW" (ByVal pszOutPointer As Long, ByVal pszSrcPointer As Long, ByVal cchMax As Long, ByVal dwFlags As Long) As Long

Private Declare Function GetFileVersionInfoW Lib "version" (ByVal ptrToFilename As Long, ByVal dwHandle As Long, ByVal dwLen As Long, ByVal ptrToDstData As Long) As Long
Private Declare Function GetFileVersionInfoSizeW Lib "version" (ByVal ptrToFilename As Long, ByRef dstHandle As Long) As Long
Private Declare Function VerQueryValueW Lib "version" (ByVal ptrToSrcBlock As Long, ByVal ptrToSubblockName As Long, ByRef dstBufferPtr As Long, ByRef dstBufferSize As Long) As Long

'Use the API to retrieve a Unicode-friendly version of App.Path
Friend Function AppPathW() As String
    
    'When running from the IDE, App.Path is our only option
    If (Not OS.IsProgramCompiled) Then
        AppPathW = App.Path
        AppPathW = Me.PathAddBackslash(AppPathW)
    Else
    
        'MAX_PATH no longer applies, but the docs are unclear on a reasonable buffer size.  Because buffer behavior is sketchy on XP
        ' (see https://msdn.microsoft.com/en-us/library/windows/desktop/ms683197%28v=vs.85%29.aspx) it's easier to just go with a
        ' huge buffer, then manually trim the result.
        Dim TEMPORARY_LARGE_BUFFER As Long
        TEMPORARY_LARGE_BUFFER = 1024
        
        Dim tmpString As String
        tmpString = String$(TEMPORARY_LARGE_BUFFER, 0)
        
        If (GetModuleFileNameW(0&, StrPtr(tmpString), TEMPORARY_LARGE_BUFFER \ 2) <> 0) Then
            AppPathW = Strings.TrimNull(tmpString)
            AppPathW = Me.FileGetPath(AppPathW)
            FSOFeedback "AppPathW", "App.Path equivalent: " & AppPathW
        Else
            FSOError "AppPathW", "couldn't retrieve Unicode-friendly version of App.Path.  Falling back to the default VB path."
            AppPathW = App.Path
            AppPathW = Me.PathAddBackslash(AppPathW)
        End If
        
    End If
    
End Function

'Append binary data to an arbitrary file.  The caller is responsible for supplying a source pointer and length (in bytes).
'
'Returns TRUE if successful, FALSE otherwise.
' (FALSE generally only occurs if write access to the destination folder is restricted.)
Friend Function FileAppendBinaryData(ByVal srcPointer As Long, ByVal srcLengthInBytes As Long, ByRef dstFilename As String) As Boolean
    
    On Error GoTo StopAppendBinary
    
    'Open the file with append rights only
    Dim hFile As Long
    If Me.FileCreateAppendHandle(dstFilename, hFile) Then
        
        'Write the byte array and exit!
        FileAppendBinaryData = Me.FileWriteData(hFile, srcPointer, srcLengthInBytes)
        Me.FileCloseHandle hFile
        
    'hFile should technically never be zero, FYI.  An invalid handle value may be returned for generic filesystem errors,
    ' but if a previous "file exists and is writable" check passed, that outcome is extremely unlikely.
    Else
        FileAppendBinaryData = False
    End If
    
    Exit Function
    
StopAppendBinary:
    FSOError "FileAppendBinaryData", "internal error on " & dstFilename, Err.Number
    FileAppendBinaryData = False
End Function

'Append string data to an existing text file.  Because this function is PhotoDemon-specific, UTF-8 output is enforced.
'
'Returns TRUE if successful, FALSE otherwise.
' (FALSE generally only occurs if write access to the destination folder is restricted.)
Friend Function FileAppendText(ByVal srcString As String, ByRef dstFilename As String) As Boolean
    
    On Error GoTo StopAppendText
    
    'Convert the incoming string to UTF-8
    Dim fileBytes() As Byte, lenFileBytes As Long
    Strings.UTF8FromString srcString, fileBytes, lenFileBytes
    
    'Open the file with append rights only
    Dim hFile As Long
    If Me.FileCreateAppendHandle(dstFilename, hFile) Then
        
        'Write the byte array and exit!
        Me.FileWriteData hFile, VarPtr(fileBytes(0)), lenFileBytes
        Me.FileCloseHandle hFile
        FileAppendText = True
        
    'hFile should technically never be zero, FYI.  An invalid handle value may be returned for generic filesystem errors,
    ' but if a previous "does file exist" check passed, that outcome is extremely unlikely.
    Else
        FileAppendText = False
    End If
    
    Exit Function
    
StopAppendText:
    FSOError "FileAppendText", "internal error on " & dstFilename, Err.Number
    FileAppendText = False
End Function

'Close an open file handle.  (Returns TRUE if successful.)
Friend Function FileCloseHandle(ByRef srcHandle As Long) As Boolean
    
    If (srcHandle <> 0) Then
        FileCloseHandle = (CloseHandle(srcHandle) <> 0)
    Else
        FSOError "FileCloseHandle", "was passed an empty handle.  Ask yourself why this handle was already closed!"
    End If
    
    'As a convenience to the caller, manually reset the handle to 0.  (This is why it's passed ByRef.)
    If FileCloseHandle Then srcHandle = 0
    
End Function

'Given a file handle successfully created by FileCreateHandle, above, create a memory-mapped view of the file and return
' a pointer to the start of the map.  This is effectively a "shortcut" function that maps the entire file view at once,
' enabling the caller to CopyMemory whatever they want directly into the file itself, which the system can then flush
' at its leisure (enabling a sort of poor man's asynchronicity).
'
'IMPORTANT: read/write access is required by this function, so you *must* specify both read+write attributes on the
' originally created file handle.  If you do not, the mmap handle creation will fail.
'
'Besides returning a basic pass/fail result, this function fills two output values: a handle to the mapped object,
' and a pointer for writing.  The pointer can be immediately used with CopyMemory, and the mapped object handle must
' be retained as we need it to close the map after writing has finished.  Do not confuse the two values!
'
'ALSO: you *must* remember the value of dstBaseAddress.  The *same* base address value must be passed to the UnmapView
' function upon completion, so for things like offset calculations, you should use a separate variable.
Friend Function FileConvertHandleToMMPtr(ByVal srcFileHandle As Long, ByRef dstMappedHandle As Long, ByRef dstBaseAddress As Long, ByVal sizeInBytes As Long, Optional ByVal nameOfMap As String = vbNullString) As Boolean
    
    Dim ptrName As Long
    If (Len(nameOfMap) <> 0) Then ptrName = StrPtr(nameOfMap) Else ptrName = 0
    dstMappedHandle = CreateFileMapping(srcFileHandle, 0&, PAGE_READWRITE, 0&, sizeInBytes, ptrName)
                                 
    If (dstMappedHandle = 0) Then
        FSOError "FileConvertHandleToMMPtr", " failed to create a handle for #" & srcFileHandle & ".  Relevant last error: " & Err.LastDllError
        FileConvertHandleToMMPtr = False
    
    'Handle creation appears to be successful
    Else
        
        'Immediately map a full-sized view of the file, and retrieve a pointer to the initial offset
        dstBaseAddress = MapViewOfFile(dstMappedHandle, FILE_MAP_WRITE, 0&, 0&, 0&)
        
        If (dstBaseAddress = 0) Then
            CloseHandle dstMappedHandle
            FSOError "FileConvertHandleToMMPtr", "failed to map the created handle for " & nameOfMap & ".  Relevant last error: " & Err.LastDllError
            FileConvertHandleToMMPtr = False
        Else
            FileConvertHandleToMMPtr = True
        End If
        
    End If
    
End Function

'Copy a given file.  Returns TRUE if successful; false otherwise.  The debug window will provide additional failure information.
Friend Function FileCopyW(ByVal srcFilename As String, ByVal dstFilename As String) As Boolean

    'Failsafe "source file exists" check
    If Me.FileExists(srcFilename) Then
        FileCopyW = (CopyFileW(StrPtr(srcFilename), StrPtr(dstFilename), 0) <> 0)
        If (Not FileCopyW) Then FSOError "CopyFile", "unable to copy " & srcFilename & " due to WAPI error " & Err.LastDllError
    Else
        FileCopyW = False
        FSOError "FileCopyW", "failed because the source file (" & srcFilename & ") doesn't exist."
    End If

End Function

'Given a file path, access indicators, and a destination long, retrieve an hFile handle that DOESN'T ERASE OR OTHERWISE
' MODIFY THE TARGET FILE, IF IT ALREADY EXISTS.  (If it doesn't exist, it will be created.)
'
'If the file *does* exist, the file's pointer will automatically be moved to the *end* of the file, so any subsequent
' write requests can be naturally appended without special work on the caller's part.
'
'Note that this function (by design) assumes normal file attributes and active sharing rights (e.g. other functions can
' access files opened via this function).  If this behavior is not desired, modify the attributeFlags and shareFlags
' settings below.
'
'RETURNS: TRUE if handle generated successfully; FALSE otherwise.
Friend Function FileCreateAppendHandle(ByRef srcFilePath As String, ByRef dstFileHandle As Long) As Boolean
    
    Dim accessFlags As Long
    accessFlags = GENERIC_READ Or GENERIC_WRITE
    
    'Regardless of what the user plans to do with this file, we enable all sharing flags.  This is important while working from the IDE,
    ' as a crash (or use of the STOP button) may cause a file handle to go unclosed, preventing all access to the file until the PC is rebooted.
    Dim shareFlags As Long
    shareFlags = FILE_SHARE_WRITE Or FILE_SHARE_READ Or FILE_SHARE_DELETE
    
    'The OPEN_ALWAYS flag is important because it lets us use the file as-is if it exists, or create the file automatically
    ' if it doesn't exist.
    Dim createFlags As Long
    createFlags = OPEN_ALWAYS
    
    Dim attributeFlags As Long
    attributeFlags = FILE_ATTRIBUTE_NORMAL
    
    'Because we're appending to an existing file, assume sequential access
    Dim optimizeFlags As Long
    optimizeFlags = FILE_FLAG_SEQUENTIAL_SCAN
    
    'If the file already exists, make a note of it, and we'll move the file pointer accordingly.
    Dim fileAlreadyExisted As Boolean
    fileAlreadyExisted = Me.FileExists(srcFilePath)
    
    'Create the handle!
    Dim hFile As Long
    hFile = CreateFileW(StrPtr(srcFilePath), accessFlags, shareFlags, 0&, createFlags, attributeFlags Or optimizeFlags, 0&)
    
    If (hFile = 0) Or (hFile = INVALID_HANDLE_VALUE) Then
        FSOError "FileCreateHandle", "failed to create a handle for " & srcFilePath & ".  Relevant last error: " & Err.LastDllError
        FileCreateAppendHandle = False
    Else
    
        dstFileHandle = hFile
        FileCreateAppendHandle = True
        
        'If the file existed prior to us accessing it, move the file pointer to the end of the file.
        If fileAlreadyExisted Then SetFilePointer dstFileHandle, 0&, 0&, FILE_END
        
    End If
    
End Function

'Shortcut function for dumping a byte array into a file.
Friend Function FileCreateFromByteArray(ByRef srcArray() As Byte, ByVal pathToFile As String, Optional ByVal overwriteExistingIfPresent As Boolean = True, Optional ByVal fileIsTempFile As Boolean = False) As Boolean
    
    On Error GoTo FileCreateFromByteArray_Failure
    FileCreateFromByteArray = FileCreateFromPtr(VarPtr(srcArray(LBound(srcArray))), UBound(srcArray) - LBound(srcArray) + 1, pathToFile, overwriteExistingIfPresent, fileIsTempFile)
    Exit Function
    
FileCreateFromByteArray_Failure:
    FSOError "FileCreateFromByteArray", "internal error.  Write abandoned.", Err.Number
    FileCreateFromByteArray = False
End Function

Friend Function FileCreateFromPtr(ByVal ptrSrc As Long, ByVal dataLength As Long, ByVal pathToFile As String, Optional ByVal overwriteExistingIfPresent As Boolean = True, Optional ByVal fileIsTempFile As Boolean = False) As Boolean
    
    On Error GoTo FileCreateFromPtr_Failure
    
    'See if the file exists, and delete as necessary
    If Me.FileExists(pathToFile) Then
    
        If (Not overwriteExistingIfPresent) Then
            FSOError "FileCreateFromPtr", "file passed already exists, and overwrites not allowed.  Write abandoned."
            Exit Function
        Else
            Me.FileDelete pathToFile
        End If
        
    End If
    
    'Create a file handle
    Dim hFile As Long, createSuccess As Boolean
    
    If fileIsTempFile Then
        createSuccess = FileCreateHandle(pathToFile, hFile, False, True, OptimizeTempFile)
    Else
        createSuccess = FileCreateHandle(pathToFile, hFile, False, True, OptimizeSequentialAccess)
    End If
    
    If createSuccess Then
        FileCreateFromPtr = FileWriteData(hFile, ptrSrc, dataLength)
        CloseHandle hFile
    End If
    
    Exit Function
    
FileCreateFromPtr_Failure:
    FSOError "FileCreateFromPtr", "internal error", Err.Number
    FileCreateFromPtr = False
End Function

'Given a file path, access indicators, and a destination long, retrieve a generic hFile handle.
' Note that this function (by design) assumes normal file attributes and active sharing rights (e.g. other functions can
' access files we currently have open).  If this behavior is not desired, modify the attributeFlags and shareFlags
' settings below.
'
'RETURNS: TRUE if handle generated successfully; FALSE otherwise.
Friend Function FileCreateHandle(ByRef srcFilePath As String, ByRef dstFileHandle As Long, Optional ByVal requestReadAccess As Boolean = False, Optional ByVal requestWriteAccess As Boolean = False, Optional ByVal optimizeAccess As PD_FILE_ACCESS_OPTIMIZE = OptimizeNone) As Boolean
    
    Dim accessFlags As Long
    If requestReadAccess Then accessFlags = accessFlags Or GENERIC_READ
    If requestWriteAccess Then accessFlags = accessFlags Or GENERIC_WRITE
    
    'Regardless of what the user plans to do with this file, we enable all sharing flags.  This is important while working from the IDE,
    ' as a crash (or use of the STOP button) may cause a file handle to go unclosed, preventing all access to the file until the PC is rebooted.
    Dim shareFlags As Long
    shareFlags = FILE_SHARE_WRITE Or FILE_SHARE_READ Or FILE_SHARE_DELETE
    
    Dim createFlags As Long
    If requestWriteAccess Then createFlags = CREATE_ALWAYS Else createFlags = OPEN_EXISTING
    
    Dim attributeFlags As Long
    attributeFlags = FILE_ATTRIBUTE_NORMAL
    
    Dim optimizeFlags As Long
    If (optimizeAccess = OptimizeNone) Then
        optimizeFlags = 0
    ElseIf (optimizeAccess = OptimizeRandomAccess) Then
        optimizeFlags = FILE_FLAG_RANDOM_ACCESS
    ElseIf (optimizeAccess = OptimizeSequentialAccess) Then
        optimizeFlags = FILE_FLAG_SEQUENTIAL_SCAN
    ElseIf (optimizeAccess = OptimizeTempFile) Then
        optimizeFlags = 0
        attributeFlags = attributeFlags Or FILE_ATTRIBUTE_TEMPORARY
    End If
    
    'Create the handle!
    Dim hFile As Long
    hFile = CreateFileW(StrPtr(srcFilePath), accessFlags, shareFlags, 0&, createFlags, attributeFlags Or optimizeFlags, 0&)
    
    If (hFile = 0) Or (hFile = INVALID_HANDLE_VALUE) Then
        FSOError "FileCreateHandle", "failed to create a handle for " & srcFilePath & ".  Relevant last error: " & Err.LastDllError
        FileCreateHandle = False
    Else
        dstFileHandle = hFile
        FileCreateHandle = True
    End If
    
End Function

'Kill a given file.  Returns TRUE if the file does not exist after the operation completes.  (Thus, if the file doesn't exist prior
' to calling this function, it will still return TRUE.)
'
'A FALSE return indicates the file still exists.  The debug window will contain additional debug data.
Friend Function FileDelete(ByVal srcFilename As String) As Boolean
    
    'If the file doesn't exist, return TRUE in advance.
    If Me.FileExists(srcFilename) Then
        FileDelete = (DeleteFileW(StrPtr(srcFilename)) <> 0)
        If (Not FileDelete) Then FSOError "FileDelete", "unable to kill " & srcFilename & " due to WAPI error " & Err.LastDllError
    Else
        FileDelete = True
    End If
    
End Function

'Returns a VB boolean indicating whether a given file exists.  This should also work on system files that prevent direct access;
' the ERROR_SHARING_VIOLATION check below is meant to capture such files.
' (IMPORTANT NOTE: wildcards are not supported by this function.)
'
'Thank you to Bonnie West for the original implementation of this function.
Friend Function FileExists(ByRef srcFile As String) As Boolean
    If ((GetFileAttributesW(StrPtr(srcFile)) And vbDirectory) = 0) Then
        FileExists = True
    Else
        FileExists = (Err.LastDllError = ERROR_SHARING_VIOLATION)
    End If
End Function

'Unicode-aware wrapper to FindFirst/NextFileW APIs.  Search parameters should match inputs to VB's Dir() function,
' and return values are also compatible with VB's Dir function.  Said another way, this is a loop-free, Unicode-aware
' replacement for VB's Dir() function.
'
'Results, if any, are returned as a stack of strings (via pdStringStack).
' Returns TRUE if FindFirst/NextFileW were invoked successfully; FALSE otherwise.
' Note that some special handling exists - e.g. "."/".." are *not* returned, by design.
Friend Function FileFind(ByRef searchParam As String, ByRef dstStack As pdStringStack) As Boolean
    
    'Start by applying some modifications to the incoming search parameter string.
    ' (FindFirstFile fails under conditions that VB's Dir() does not.)
    If (Len(searchParam) <> 0) Then
    
        If (dstStack Is Nothing) Then Set dstStack = New pdStringStack Else dstStack.resetStack
    
        'First, prepend "\\?\" to sParam.  This enables long file paths.
        If (Not Strings.StringsEqual(Left$(searchParam, 4), "\\?\")) Then searchParam = "\\?\" & searchParam
    
        'FindFirstFile fails if the requested path has a trailing slash.  If the user hands us a bare path,
        ' assume that they want to iterate all files and folders within that folder.
        If (Strings.StringsEqual(Right$(searchParam, 1), "\") Or Strings.StringsEqual(Right$(searchParam, 1), "/")) Then searchParam = searchParam & "*"
            
        'Retrieve the first file in the new search; this returns the search handle we'll use for subsequent searches
        Dim searchHandle As Long, findResult As WIN32_FIND_DATA
        searchHandle = FindFirstFileW(StrPtr(searchParam), VarPtr(findResult))
            
        'Check for failure.  Failure can occur for multiple reasons: bad inputs, no files meeting the criteria, etc.
        If (searchHandle = INVALID_HANDLE_VALUE) Then
            
            'No files found is fine, but if the caller screwed up the input path, we want to print some debug info.
            If (Err.LastDllError <> ERROR_FILE_NOT_FOUND) Then FSOError "FileFind", "possibly handed a bad path (" & searchParam & "). Please investigate."
            Exit Function
        
        'We obtained a good handle.  Store the first result, and continue searching until all results are discovered.
        Else
            
            Do
                
                Dim retString As String
                retString = Strings.TrimNull(findResult.cFileName)
                
                Const STR_DOT As String = "."
                Const STR_DOTDOT As String = ".."
                
                'Look for "." and ".." returns; these are *not* returned, by design.
                If (Len(retString) <> 0) And (Strings.StringsNotEqual(retString, STR_DOT, True) And Strings.StringsNotEqual(retString, STR_DOTDOT, True)) Then
                    dstStack.AddString retString
                End If
                
            Loop While (FindNextFileW(searchHandle, VarPtr(findResult)) <> 0)
                
            'Close the search handle and report any irregularities
            FindClose searchHandle
            searchHandle = 0
            
            'Check for unexpected errors
            Dim lastDllErr As Long
            lastDllErr = Err.LastDllError
            FileFind = (lastDllErr = 0) Or (lastDllErr = ERROR_NO_MORE_FILES)
            If (Not FileFind) Then FSOError "FileFind", "terminated for a reason other than ERROR_NO_MORE_FILES. Please investigate.  (FYI: error # is last DLL error.)", lastDllErr
            
        End If
        
    Else
        FSOError "FileFind", "null search string!"
    End If
    
End Function

'After a file has been opened, you can use this function to query the current file pointer position
Friend Function FileGetCurrentPointer(ByVal srcHandle As Long) As Long
    FileGetCurrentPointer = SetFilePointer(srcHandle, 0&, 0&, FILE_CURRENT)
End Function

'Function to return the extension from a filename.
Friend Function FileGetExtension(ByRef srcFile As String) As String

    'We want to find two different pieces of information in the target file:
    ' 1) The right-most position of a path separator (forward or backward slash)
    ' 2) The right-most position of an extension separator (".")
    '
    'TODO: normalize path before processing it
    
    Dim slashPos As Long
    slashPos = InStrRev(srcFile, "\", -1, vbBinaryCompare)
    If (slashPos = 0) Then slashPos = InStrRev(srcFile, "/", -1, vbBinaryCompare)
    
    Dim dotPos As Long
    dotPos = InStrRev(srcFile, ".", -1, vbBinaryCompare)
    
    'If a slash sits later in the string than the dot, there is no extension; return a blank string
    If (slashPos > dotPos) Then
        FileGetExtension = vbNullString
    
    'If the dot sits later than the path, we know where to find the file extension
    Else
        FileGetExtension = Right$(srcFile, Len(srcFile) - dotPos)
    End If
            
End Function

'Return just the filename from a full path.  Optionally, you can ask the function to strip the extension (if any)
' from the filename.  If stripExtension is FALSE, the full filename + extension will be returned (e.g. "file.txt").
Friend Function FileGetName(ByRef srcPath As String, Optional ByVal stripExtension As Boolean = False) As String
    
    'Normalize path
    'TODO!
    
    'First, find the last path indicator (if any) in the incoming string
    Dim slashPos As Long
    slashPos = InStrRev(srcPath, "\", -1, vbBinaryCompare)
    If (slashPos = 0) Then slashPos = InStrRev(srcPath, "/", -1, vbBinaryCompare)
    
    'If the function fails (because no slash is found), we'll simply return the entire incoming string, under the
    ' assumption that it's already just a filename.
    If (slashPos = 0) Then
        FileGetName = srcPath
        
    Else
    
        'Retrieve everything past the slash
        FileGetName = Right$(srcPath, Len(srcPath) - slashPos)
        
        'If the caller wants us to strip the extension, do so before exiting
        If stripExtension Then
            
            Dim dotPos As Long
            dotPos = InStrRev(FileGetName, ".")
            
            If (dotPos > 1) Then
                FileGetName = Left$(FileGetName, dotPos - 1)
            ElseIf (dotPos = 1) Then
                FileGetName = vbNullString
            
            'No final "Else" is required, because FileGetName has already been set, above
            End If
            
        End If
        
    End If
    
End Function

'Given a full path+filename string, return only the folder portion.
' Returns a blank string if a folder cannot be extracted.
Friend Function FileGetPath(ByRef srcPath As String) As String
    Dim slashPosition As Long
    slashPosition = InStrRev(srcPath, "\", -1, vbBinaryCompare)
    If (slashPosition <> 0) Then FileGetPath = Left$(srcPath, slashPosition)
End Function

'Return varies types of file times in currency format.  By default, WAPI functions return file times using a custom
' FILETIME struct.  This can be converted to system time objects via helper functions.  At present, this function
' simply collates the two values into a Currency and returns that, which is good enough for current PD usage
' (comparing dates, generating checksums of changed files, etc).
'
'RETURNS: non-zero value if successful.
Friend Function FileGetTimeAsCurrency(ByRef srcFile As String, Optional ByVal typeOfTime As PD_FILE_TIME = PDFT_CreateTime) As Currency
    
    FileGetTimeAsCurrency = 0
    
    Dim hFile As Long
    If FileCreateHandle(srcFile, hFile, True) Then
        
        Dim createTime As WAPI_FILETIME, accessTime As WAPI_FILETIME, writeTime As WAPI_FILETIME
        If (GetFileTime(hFile, createTime, accessTime, writeTime) <> 0) Then
        
            Dim srcTime As WAPI_FILETIME
            
            If (typeOfTime = PDFT_AccessTime) Then
                srcTime = accessTime
            ElseIf (typeOfTime = PDFT_WriteTime) Then
                srcTime = writeTime
            Else
                srcTime = createTime
            End If
            
            'At present, PD only needs file times for checksum purposes.  In the future, it would be nice to add
            ' comparison functions, but for now, you have to deal with times being returned as currency.
            CopyMemoryStrict VarPtr(FileGetTimeAsCurrency), VarPtr(srcTime.dwHighDateTime), 4&
            CopyMemoryStrict VarPtr(FileGetTimeAsCurrency) + 4, VarPtr(srcTime.dwLowDateTime), 4&
        
        End If
        
        CloseHandle hFile
    
    End If

End Function

'Return varies types of file times in currency format.  By default, WAPI functions return file times using a custom
' FILETIME struct.  This can be converted to VB Date objects via a few simple helper functions.
'
'RETURNS: the specified file time if successful, DATE_MINIMUM if unsuccessful.
Friend Function FileGetTimeAsDate(ByRef srcFile As String, Optional ByVal typeOfTime As PD_FILE_TIME = PDFT_CreateTime) As Date
    
    'Cache file properties in a local UDT
    Dim hFind As Long, tmpFindData As WIN32_FIND_DATA
    hFind = FindFirstFileW(StrPtr(srcFile), VarPtr(tmpFindData))
    If (hFind <> INVALID_HANDLE_VALUE) Then
        
        FindClose hFind
        
        'Extract the various date bits
        If (typeOfTime = PDFT_CreateTime) Then
            FileGetTimeAsDate = Win32ToVBTime(tmpFindData.ftCreationTime)
        ElseIf (typeOfTime = PDFT_AccessTime) Then
            FileGetTimeAsDate = Win32ToVBTime(tmpFindData.ftLastAccessTime)
        ElseIf (typeOfTime = PDFT_WriteTime) Then
            FileGetTimeAsDate = Win32ToVBTime(tmpFindData.ftLastWriteTime)
        End If
    
    Else
        FSOError "FileGetTimeAsDate", "failed to retrieve date/time values for: " & srcFile
    End If

End Function

'Given an arbitrary file path, return embeded versioning information.  The optional getProductVersionInstead value sets whether
' file or product version values are returned.  These are often the same, as Visual Studio inheritance rules typically make the
' Product value match the File value if it isn't explicitly given
' (http://all-things-pure.blogspot.com/2009/09/assembly-version-file-version-product.html).
' But this rule is not foolproof, so double-check 3rd-party sources in particular.
'
'Returns TRUE if successful; FALSE otherwise.  If the function fails, DLL error info will be dumped to the debug window.
Friend Function FileGetVersion(ByVal srcFile As String, ByRef verMajor As Long, ByRef verMinor As Long, ByRef verBuild As Long, ByRef verRevision As Long, Optional ByVal getProductVersionInstead As Boolean = True) As Boolean
    
    FileGetVersion = False
    
    Dim bufferSize As Long, tmpLong As Long
    
    'Start by retrieving the size of the required buffer
    bufferSize = GetFileVersionInfoSizeW(StrPtr(srcFile), tmpLong)
    If (bufferSize > 0) Then
     
        'Retrieve the root versioning block, without any localization edits (https://msdn.microsoft.com/en-us/library/windows/desktop/ms647464%28v=vs.85%29.aspx)
        Dim tmpBuffer() As Byte
        ReDim tmpBuffer(0 To bufferSize - 1) As Byte
        If (GetFileVersionInfoW(StrPtr(srcFile), 0&, bufferSize, VarPtr(tmpBuffer(0))) <> 0) Then
            
            'Finally, we need to query the actual value(s) we want.  For versioning, we only require the root block.
            Dim subBlockName As String
            subBlockName = "\"
            
            Dim ptrDstBuffer As Long, dstBufferLen As Long
            If VerQueryValueW(VarPtr(tmpBuffer(0)), StrPtr(subBlockName), ptrDstBuffer, dstBufferLen) Then
                
                'Copy the raw bytes into a usable struct
                Dim finalFileInfo As VS_FIXEDFILEINFO
                CopyMemoryStrict VarPtr(finalFileInfo), ptrDstBuffer, LenB(finalFileInfo)
                
                'Return file or product version number, depending on the optional getProductVersionInstead parameter.
                With finalFileInfo
                    If getProductVersionInstead Then
                        verMajor = (.dwProductVersionMSh And &HFFFF&)
                        verMinor = (.dwProductVersionMSl And &HFFFF&)
                        verBuild = (.dwProductVersionLSh And &HFFFF&)
                        verRevision = (.dwProductVersionLSl And &HFFFF&)
                    Else
                        verMajor = (.dwFileVersionMSh And &HFFFF&)
                        verMinor = (.dwFileVersionMSl And &HFFFF&)
                        verBuild = (.dwFileVersionLSh And &HFFFF&)
                        verRevision = (.dwFileVersionLSl And &HFFFF&)
                    End If
                End With
                
                FileGetVersion = True
                
            Else
                FSOError "FileGetVersion", "Could not retrieve version information for (" & srcFile & ") due to VerQueryValue error.", Err.Last
            End If
            
        Else
            FSOError "FileGetVersion", "Could not retrieve version information for (" & srcFile & ") due to GetFileVersionInfo error.", Err.LastDllError
        End If
        
    Else
        FSOError "FileGetVersion", "Could not retrieve version information for (" & srcFile & ") due to buffer allocation failure.", Err.LastDllError
    End If

End Function

'Convenience wrapper to FileGetVersion, above.
Friend Function FileGetVersionAsString(ByVal FullFileName As String, ByRef dstVersionString As String, Optional ByVal getProductVersionInstead As Boolean = True) As Boolean
    Dim tmpLong1 As Long, tmpLong2 As Long, tmpLong3 As Long, tmpLong4 As Long
    FileGetVersionAsString = FileGetVersion(FullFileName, tmpLong1, tmpLong2, tmpLong3, tmpLong4, getProductVersionInstead)
    dstVersionString = CStr(tmpLong1) & "." & CStr(tmpLong2) & "." & CStr(tmpLong3) & "." & CStr(tmpLong4)
End Function

'Return the size of a file with a Unicode filename.  For now, only the low 4-bytes (signed!) of the full 8-byte return are used.
' This is all PD requires at present, since it can't work with 2 GB+ images anyway.
' IMPORTANT!  It's assumed that the caller checked the file's existence PRIOR to calling this function.
Friend Function FileLenW(ByRef srcPath As String) As Long
    
    Dim hFile As Long
    If Me.FileCreateHandle(srcPath, hFile, True) Then
        FileLenW = CLng(FileLenW_Large(hFile))
        CloseHandle hFile
    Else
        FileLenW = 0
    End If
    
End Function

'Unicode-compatible LOF variant.  Note that the hFile must have generic read attributes active; write-alone won't work.
' Note also the specific currency-type workaround; DO NOT cast the result directly into a long int.
Friend Function FileLenW_Large(ByVal hFile As Long) As Currency
    
    Dim tmpVal As Currency
    tmpVal = 0
    
    If (GetFileSizeEx(hFile, tmpVal) <> 0) Then
        FileLenW_Large = tmpVal * 10000
    Else
        FSOError "FileLenW_Large", "failed on hFile " & hFile & ".  LastDllErr = " & Err.LastDllError
        FileLenW_Large = 0
    End If
    
End Function

'Dump a specified file directly into a VB byte array.  No additional processing is performed, so it's up to you to interpret
' the bytes correctly.  (Note also that the *entire* file is brought into memory - so this is a bad function for extremely
' large files!)
'
'Returns TRUE if successful; FALSE otherwise.  FALSE generally means the file is missing, or read access was denied.
Friend Function FileLoadAsByteArray(ByRef srcFile As String, ByRef dstArray() As Byte) As Boolean

    On Error GoTo StopFileLoad
    
    FileLoadAsByteArray = False
    
    'Before doing anything else, make sure the file exists
    If (Not Me.FileExists(srcFile)) Then
        FSOError "FileLoadAsByteArray", "can't find " & srcFile & ".  Abandoning load."
    Else
        
        'Dump file contents into a byte array.
        Dim hFile As Long
        If FileCreateHandle(srcFile, hFile, True, False, OptimizeSequentialAccess) Then
            
            'Prep the destination buffer.  (Note that we check buffer size in advance; there's no reason to resize it
            ' if we're just going to overwrite it with new data!)
            Dim lenOfFile As Long
            lenOfFile = CLng(FileLenW_Large(hFile))
            
            'Only proceed if the file is not zero-length
            If (lenOfFile > 0) Then
                
                If VBHacks.IsArrayInitialized(dstArray) Then
                    If (LBound(dstArray) <> 0) Or (UBound(dstArray) <> lenOfFile - 1) Then ReDim dstArray(0 To lenOfFile - 1) As Byte
                Else
                    ReDim dstArray(0 To lenOfFile - 1) As Byte
                End If
                
                'Retrieve the full file contents
                Dim tmpVerify As Long
                If (ReadFile(hFile, VarPtr(dstArray(0)), lenOfFile, tmpVerify, 0&) <> 0) Then
                    FileLoadAsByteArray = True
                Else
                    FSOError "FileLoadAsByteArray", "failure when calling ReadFile on " & srcFile & ". Abandoning load."
                End If
                
            Else
                FSOError "FileLoadAsByteArray", "source file is zero-length: " & srcFile
                ReDim dstArray(0) As Byte
                FileLoadAsByteArray = False
            End If
            
            CloseHandle hFile
            
        Else
            FSOError "FileLoadAsByteArray", "failed to create a valid handle for " & srcFile & ". Abandoning load."
        End If
        
    End If
    
    Exit Function
    
StopFileLoad:
    FSOError "FileLoadAsByteArray", "internal error on " & srcFile & ". Abandoning load.", Err.Number
    FileLoadAsByteArray = False
End Function

'Convenience wrapper to pdStream's file load functions.  For file-backed access (where the original bytes are
' *not* loaded into memory) you must manually invoke the pdStream object, as this function - by design -
' loads the entire source file into memory.
Friend Function FileLoadAsPDStream(ByRef srcFile As String, ByRef dstStream As pdStream) As Boolean
    
    If (dstStream Is Nothing) Then Set dstStream = New pdStream Else dstStream.StopStream True
    Dim streamSize As Long
    streamSize = Me.FileLenW(srcFile)
    FileLoadAsPDStream = dstStream.StartStream(PD_SM_MemoryBacked, PD_SA_ReadWrite, , streamSize)
    
    Dim tmpHandle As Long
    If Me.FileCreateHandle(srcFile, tmpHandle, True, False, OptimizeSequentialAccess) Then
        Me.FileReadData tmpHandle, dstStream.Peek_PointerOnly(0), streamSize
        Me.FileCloseHandle tmpHandle
        dstStream.SetSizeExternally streamSize
    End If
    
End Function

'Load a text file into a VB string.  Any valid encoding is supported.  Heuristics are used for "best-guess" encoding detection when
' BOM and other markers are unclear.
'
'The optional forceWindowsLineEndings parameter will apply additional checks, to make sure line endings are consistently vbCrLf
' (vs Unix-style vbLf only).
'
'RETURNS: TRUE if successful; FALSE otherwise.  FALSE generally means the file is missing, or read access was denied.
'
'Thank you to Dana Seaman and cyberactivex.com, whose "GenericFileRead" function served as the original inspiration for this work:
' http://www.cyberactivex.com/UnicodeTutorialVb.htm
Friend Function FileLoadAsString(ByRef srcFile As String, ByRef dstString As String, Optional ByVal forceWindowsLineEndings As Boolean = True) As Boolean
    
    On Error GoTo StopTextFileRead
    
    FileLoadAsString = False
    
    'Use a wrapper function to dump the file's contents into a byte array.
    Dim fileBytes() As Byte
    If Me.FileLoadAsByteArray(srcFile, fileBytes) Then FileLoadAsString = Strings.StringFromMysteryBytes(fileBytes, dstString, forceWindowsLineEndings)
    
    Exit Function
    
StopTextFileRead:
    FSOError "FileLoadAsString", " internal error on " & srcFile & ". Abandoning load.", Err.Number
    FileLoadAsString = False
End Function

'After a file has been opened, you can use this function to manually move the file pointer around
Friend Function FileMovePointer(ByVal srcHandle As Long, ByVal newPosition As Long, Optional ByVal startingPosition As FILE_POINTER_MOVE_METHOD = FILE_BEGIN) As Boolean
    FileMovePointer = (SetFilePointer(srcHandle, newPosition, 0&, startingPosition) <> INVALID_SET_FILE_POINTER)
End Function

'Reads arbitrary binary data from file.  It is up to the caller to make sure the requested number of bytes are valid and accurate,
' and most importantly, *that the destination memory is already sized appropriately*.
' - The read will be performed synchronously, FYI.
' - An optional newFilePointer value can be passed.  A pointer move method can also be specified.  Note that by default, the pointer
'   will be reset *relative to the start of the file*, *not relative to its current position*.  This is meant to mimic file positions
'   from old VB file access code, but you can change it to be relative to the current pointer position by modifying the
'   filePointerMoveMethod param.
'
' Returns: TRUE if successful; FALSE otherwise.
Friend Function FileReadData(ByVal srcHandle As Long, ByVal ptrToDestinationObject As Long, ByVal numOfBytesToRead As Long, Optional ByVal newFilePointer As Long = -1, Optional ByVal filePointerMoveMethod As FILE_POINTER_MOVE_METHOD = FILE_BEGIN) As Boolean
    
    'Manually move the pointer only if required
    If (newFilePointer <> -1) Then SetFilePointer srcHandle, newFilePointer, ByVal 0&, filePointerMoveMethod
    
    'Read the data!
    Dim tmpReturn As Long
    FileReadData = (ReadFile(srcHandle, ptrToDestinationObject, numOfBytesToRead, tmpReturn, 0&) <> 0)
    
    'As a convenience to the caller, print out scenarios where the actual amount read differs from the amount of data requested.
    If (tmpReturn <> numOfBytesToRead) Then FSOError "FileReadData", "retrieved " & CStr(tmpReturn) & " bytes instead of the " & CStr(numOfBytesToRead) & " bytes you requested.  Check your math."
    
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
Friend Function FileReplace(ByVal oldFile As String, ByVal newFile As String, Optional ByVal customBackupFile As String = vbNullString) As PD_FILE_PATCH_RESULT
    
    'Because this function is capable of botching an existing PD install if catastrophic failure occurs, error handling must be comprehensive.
    On Error GoTo ReplaceFailure
    
    Dim backupFile As String
    Dim rfReturn As Long, mfReturn As Long
    
    'Make sure the target file exists.  If it doesn't, we need to use a different set of APIs
    If Me.FileExists(oldFile) Then
    
        'Consider forcibly specifying a backup path if the user didn't provide their own
        If (Len(customBackupFile) = 0) Then
        
            'IMPORTANT NOTE!  If the files you're patching are critical, I STRONGLY recommend forcibly creating a backup here.
            ' Otherwise, you risk a scenario where something goes wrong during the replace step, and the original file is gone forever.
            
            'For example, in PhotoDemon I use the standard Data/Updates folder for backups when patching, like this:
            'backupFile = UserPrefs.GetUpdatePath & FileGetName(oldFile)
            
            'This ensures that we always have a way to undo the replace step if something goes wrong
        
        Else
            backupFile = customBackupFile
        End If
        
        'ReplaceFileW returns a non-zero value if successful, or zero if it fails.  Because this function is used to patch PhotoDemon.exe,
        ' there is no margin for failure, so detailed handling of every possible outcome is required.
        rfReturn = ReplaceFileW(StrPtr(oldFile), StrPtr(newFile), StrPtr(backupFile), REPLACEFILE_IGNORE_MERGE_ERRORS, 0&, 0&)
        
        'Success means we can exit immediately
        If (rfReturn <> 0) Then
            
            'Remove the backup file, then report success
            Me.FileDelete backupFile
            FileReplace = FPR_SUCCESS
            Exit Function
        
        'Failure is a problem.  Retrieve the error cause, and do whatever we can to resolve it.
        Else
        
            Select Case Err.LastDllError
            
                'The replacement file could not be renamed. If lpBackupFileName was specified, the replaced and replacement files retain their original
                ' file names. Otherwise, the replaced file no longer exists and the replacement file exists under its original name.
                Case ERROR_UNABLE_TO_MOVE_REPLACEMENT
                    
                    'See if the old file exists
                    If Me.FileExists(oldFile) Then
                    
                        'The old file exists.  Remove the backup file (if any), then exit.  Because the system is in an identical place as it was prior
                        ' to this function being invoked, the caller is free to try again.
                        Me.FileDelete customBackupFile
                        FileReplace = FPR_FAIL_NOTHING_CHANGED
                        Exit Function
                       
                    Else
                    
                        'The old file no longer exists.  Restore the backup file, if at all possible.
                        If Me.FileExists(backupFile) Then
                        
                            'The backup file exists.  Try to restore it to the location of the original file.
                            mfReturn = MoveFileExW(StrPtr(backupFile), StrPtr(oldFile), MOVEFILE_REPLACE_EXISTING Or MOVEFILE_COPY_ALLOWED)
                            
                            'The move succeeded!  Perform a failsafe deletion of the backup file, then notify the caller that the system has been
                            ' restored to its original state.
                            If (mfReturn <> 0) Then
                                Me.FileDelete backupFile
                                FileReplace = FPR_FAIL_NOTHING_CHANGED
                                Exit Function
                            
                            'The move failed.  This should never happen.
                            Else
                                FSOError "FileReplace", "Impossible outcome 1 occurred in FileReplace.  Please investigate."
                                FileReplace = FPR_FAIL_OLD_FILE_REMOVED
                                Exit Function
                            End If
                        
                        'The backup file does not exist.  This should never happen.
                        Else
                            FSOError "FileReplace", "Impossible outcome 2 occurred in FileReplace.  Please investigate."
                            FileReplace = FPR_FAIL_OLD_FILE_REMOVED
                            Exit Function
                        End If
                    
                    End If
                
                'The replacement file could not be moved. The replacement file still exists under its original name; however, it has inherited the file streams
                ' and attributes from the file it is replacing. The file to be replaced still exists with a different name. If lpBackupFileName is specified,
                ' it will be the name of the replaced file.
                Case ERROR_UNABLE_TO_MOVE_REPLACEMENT_2
                
                    'See if the old file exists; this should never happen.
                    If Me.FileExists(oldFile) Then
                    
                        'The old file exists.  Remove the backup file (if any), then exit.  Because the system is in an identical place as it was prior
                        ' to this function being invoked, the caller is free to try again.
                        Me.FileDelete customBackupFile
                        
                        FSOError "FileReplace", "Impossible outcome 3 occurred in FileReplace.  Please investigate."
                        FileReplace = FPR_FAIL_NOTHING_CHANGED
                        Exit Function
                       
                    Else
                    
                        'The old file no longer exists.  Restore the backup file.  (The backup file is supposedly guaranteed to exist with this error code.)
                        If Me.FileExists(backupFile) Then
                        
                            'The backup file exists.  Try to restore it to the location of the original file.
                            mfReturn = MoveFileExW(StrPtr(backupFile), StrPtr(oldFile), MOVEFILE_REPLACE_EXISTING Or MOVEFILE_COPY_ALLOWED)
                            
                            'The move succeeded!  Perform a failsafe deletion of the backup file, then notify the caller that the system has been
                            ' restored to its original state.
                            If (mfReturn <> 0) Then
                                Me.FileDelete backupFile
                                FileReplace = FPR_FAIL_NOTHING_CHANGED
                                Exit Function
                            
                            'The move failed.  This should never happen.
                            Else
                                FSOError "FileReplace", "Impossible outcome 4 occurred in FileReplace.  Please investigate."
                                FileReplace = FPR_FAIL_OLD_FILE_REMOVED
                                Exit Function
                            End If
                        
                        'The backup file does not exist.  This should never happen.
                        Else
                            FSOError "FileReplace", "Impossible outcome 5 occurred in FileReplace.  Please investigate."
                            FileReplace = FPR_FAIL_OLD_FILE_REMOVED
                            Exit Function
                        End If
                    
                    End If
                
                'The replaced file could not be deleted. The replaced and replacement files retain their original file names.
                Case ERROR_UNABLE_TO_REMOVE_REPLACED
                
                    'MSDN is unclear on the state of the backup file on this error.  As a failsafe, remove it if it exists.
                    Me.FileDelete backupFile
                    FileReplace = FPR_FAIL_NOTHING_CHANGED
                    Exit Function
                
                'Other failure states are more problematic.  Defer to the primary error handler, which will try to restore the
                ' original file setup as best as it can.
                Case Else
                    GoTo ReplaceFailure
            
            End Select
        
        End If
    
    'The file-to-be-replaced does *not* exist.  Use MoveFile instead of ReplaceFile.
    Else
    
        'Use the API to attempt a move.  (Note that a backup file isn't required in this case, as nothing will be deleted unless the
        ' operation is successful.)
        mfReturn = MoveFileExW(StrPtr(oldFile), StrPtr(newFile), MOVEFILE_REPLACE_EXISTING Or MOVEFILE_COPY_ALLOWED)
        
        'The move succeeded!  Exit immediately.
        If (mfReturn <> 0) Then
            
            FileReplace = FPR_SUCCESS
            Exit Function
        
        'The move failed.  This should not happen (as the target file is known to not exist), but if it does, no harm is done as we haven't
        ' changed the filesystem at all.
        Else
            FSOError "FileReplace", "Impossible outcome 10 occurred in FileReplace.  Please investigate."
            FileReplace = FPR_FAIL_NOTHING_CHANGED
            Exit Function
        End If
    
    End If
    
    Exit Function
    
    
'Arbitrary failure states are problematic.  Basically, our goal is to restore the original file setup as best as we can.  We do this by attempting
' to restore the backup file (if it exists), and applying any renames or moves required to get everything back to the way it was.
ReplaceFailure:

    'See if the old file exists
    If Me.FileExists(oldFile) Then
    
        'The old file exists.  Remove the backup file (if any), then exit.  Because the system is in an identical place as it was prior
        ' to this function being invoked, the caller is free to try again.
        Me.FileDelete customBackupFile
        
        If Me.FileExists(newFile) Then
            FileReplace = FPR_FAIL_NOTHING_CHANGED
        Else
            FSOError "ReplaceFile", "Impossible outcome 6 occurred in ReplaceFile.  Please investigate."
            FileReplace = FPR_FAIL_NEW_FILE_REMOVED
        End If
        
        Exit Function
       
    Else
    
        'The old file no longer exists.  Restore the backup file, if at all possible.
        If Me.FileExists(backupFile) Then
        
            'The backup file exists.  Try to restore it to the location of the original file.
            mfReturn = MoveFileExW(StrPtr(backupFile), StrPtr(oldFile), MOVEFILE_REPLACE_EXISTING Or MOVEFILE_COPY_ALLOWED)
            
            'The move succeeded!  Perform a failsafe deletion of the backup file, then notify the caller that the system has been
            ' restored to its original state.
            If (mfReturn <> 0) Then
                
                Me.FileDelete backupFile
                
                If Me.FileExists(newFile) Then
                    FileReplace = FPR_FAIL_NOTHING_CHANGED
                Else
                    FSOError "FileReplace", "Impossible outcome 7 occurred in FileReplace.  Please investigate."
                    FileReplace = FPR_FAIL_NEW_FILE_REMOVED
                End If
                
                Exit Function
            
            'The move failed.  This should never happen.
            Else
                FSOError "FileReplace", "Impossible outcome 8 occurred in FileReplace.  Please investigate."
                FileReplace = FPR_FAIL_OLD_FILE_REMOVED
                Exit Function
            End If
        
        'The backup file does not exist.  This should never happen.
        Else
            FSOError "FileReplace", "Impossible outcome 9 occurred in FileReplace.  Please investigate."
            FileReplace = FPR_FAIL_OLD_FILE_REMOVED
            Exit Function
        End If
    
    End If

End Function

'Write a string out to text file.  UTF-8 is assumed, but old-style ANSI can also be used by setting the optional useUTF8 parameter to FALSE.
' If UTF-8 is requested, the optional parameter useUTF8_BOM specifies whether to preface the file with a BOM.
'
'Returns TRUE if successful, FALSE otherwise.  FALSE generally only occurs if write access to the destination folder is restricted.
' That said, if UTF-8 conversion fails, this function (by design) will silently fall back to regular ANSI output.  It will still return TRUE,
' however, so if you need to be absolutely certain that the file saved as UTF-8, you will want to perform manual validation post-saving.
Friend Function FileSaveAsText(ByRef srcString As String, ByRef dstFilename As String, Optional ByVal useUTF8 As Boolean = True, Optional ByVal useUTF8_BOM As Boolean = True) As Boolean
    
    On Error GoTo StopFileSaveAsText
    
    FileSaveAsText = False
    
    'Remove the existing output file, if any
    If Me.FileExists(dstFilename) Then Me.FileDelete dstFilename
    
    'If UTF-8 output is requested, generate it now.  (If the conversion fails, we can fall back to a non-UTF8 write.)
    Dim fileBytes() As Byte, lenFileBytes As Long
    If useUTF8 Then useUTF8 = Strings.UTF8FromStrPtr(StrPtr(srcString), Len(srcString), fileBytes, lenFileBytes, 0&) 'Strings.UTF8FromString(srcString, fileBytes, lenFileBytes)
    
    If useUTF8 Then
        
        'Open the file
        Dim hFile As Long
        If Me.FileCreateHandle(dstFilename, hFile, False, True, OptimizeSequentialAccess) Then
        
            'If requested, write a BOM first.  This is desirable for PD's internal files, as it makes it very quick to identify the new UTF-8
            ' ones vs old Windows-1252 ones.
            If useUTF8_BOM Then
                Dim bomMarker(0 To 2) As Byte
                bomMarker(0) = &HEF: bomMarker(1) = &HBB: bomMarker(2) = &HBF
                Me.FileWriteData hFile, VarPtr(bomMarker(0)), 3
            End If
            
            'Write the byte array and exit!
            Me.FileWriteData hFile, VarPtr(fileBytes(0)), lenFileBytes
            Me.FileCloseHandle hFile
            FileSaveAsText = True
            
        Else
            FSOError "FileSaveAsText", "could not create file handle"
        End If
                
    'If UTF-8 is not requested, allow VB to perform its own innate SBCS conversion
    Else
        
        'Since we're using a cheap ASCII conversion, we may as well use native VB methods too, by wrapping the
        ' (potentially Unicode) srcFile with GetShortPathNameW.
        Dim fileNum As Integer
        fileNum = FreeFile
        
        Open Me.GetShortPath(dstFilename) For Output As #fileNum
            Print #fileNum, srcString
        Close #fileNum
        
        FileSaveAsText = True
        
    End If
    
    Exit Function
    
StopFileSaveAsText:
    FSOError "FileSaveAsText", "internal error on " & dstFilename, Err.Number
    FileSaveAsText = False
End Function

'Fast write to an open file handle.  The passed data will be written to the beginning of the file, and the file will
' be truncated to match the written length.  The file handle *will not* be closed after the write.
'
'This is helpful in limited scenarios, such as XML files that are updated frequently, using the same filename.
' An open file handle can be cached, and the data can repeatedly be overwritten at will.
Friend Function FileSave_FastOverwrite(ByVal hFile As Long, ByVal srcPointer As Long, ByVal srcDataLength As Long) As Boolean
    
    On Error GoTo FastFileSaveError
    
    If (Not Me.FileMovePointer(hFile, 0, FILE_BEGIN)) Then FSOError "FileSave_FastOverwrite", "FileMovePointer failed", Err.LastDllError
    Me.FileWriteData hFile, srcPointer, srcDataLength
    If (SetEndOfFile(hFile) = 0) Then FSOError "FileSave_FastOverwrite", "SetEndOfFile failed", Err.LastDllError
    
    FileSave_FastOverwrite = True
    Exit Function
    
FastFileSaveError:
    FSOError "FileSave_FastOverwrite", "internal error on " & hFile, Err.Number
    FileSave_FastOverwrite = False
End Function

'Thin wrapper to SetEndOfFile API
Friend Function FileSetEOF(ByVal hFile As Long) As Boolean
    FileSetEOF = (SetEndOfFile(hFile) <> 0)
End Function

'Reverse function of FileLen.  Note that special considerations are required when using memory-mapped files;
' see MSDN for details: https://msdn.microsoft.com/en-us/library/windows/desktop/aa365531(v=vs.85).aspx
Friend Function FileSetLength(ByVal hFile As Long, ByVal newLength As Long) As Boolean
    If (hFile <> 0) Then
        If Me.FileMovePointer(hFile, newLength, FILE_BEGIN) Then FileSetLength = (SetEndOfFile(hFile) <> 0)
    End If
End Function

'Test a file for READ access.  This is helpful before attempting to open a file, to see if some other process has
' locked the file.
'
'RETURNS: TRUE if file is readable; FALSE otherwise.  Optionally, dstLastDLLError will be filled with the last
' DLL error, if any.
Friend Function FileTestAccess_Read(ByVal srcFile As String, Optional ByRef dstLastDLLError As Long) As Boolean
    
    FileTestAccess_Read = False
    
    'First, ensure the file exists
    If Me.FileExists(srcFile) Then
        
        'Try to create a read-only access handle
        Dim hFile As Long
        hFile = CreateFileW(StrPtr(srcFile), GENERIC_READ, FILE_SHARE_READ, 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0&)
        
        If (hFile <> 0) And (hFile <> INVALID_HANDLE_VALUE) Then
            CloseHandle hFile
            FileTestAccess_Read = True
        Else
            dstLastDLLError = Err.LastDllError
        End If
        
    End If
    
End Function

'Test a file for WRITE access.  This is helpful before attempting to save a file, to see if some other process has
' locked the file.
'
'RETURNS: TRUE if file is writable; FALSE otherwise.  Optionally, dstLastDLLError will be filled with the last
' DLL error, if any.
Friend Function FileTestAccess_Write(ByVal srcFile As String, Optional ByRef dstLastDLLError As Long) As Boolean
    
    FileTestAccess_Write = False
    
    'First, see if the file exists; if it doesn't, assume we're okay to write it
    If Me.FileExists(srcFile) Then
        
        'Try to create a read/write access handle
        Dim hFile As Long
        hFile = CreateFileW(StrPtr(srcFile), GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE, 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0&)
        
        If (hFile <> 0) And (hFile <> INVALID_HANDLE_VALUE) Then
            CloseHandle hFile
            FileTestAccess_Write = True
        Else
            dstLastDLLError = Err.LastDllError
        End If
        
    Else
        FileTestAccess_Write = True
    End If
    
    
End Function

'Write arbitrary binary data to file.  It is up to the caller to make sure the passed pointer and number of bytes are valid and accurate.
' (The write will be performed synchronously, FYI.)
' Returns: TRUE if successful; FALSE otherwise.
Friend Function FileWriteData(ByVal dstHandle As Long, ByVal srcDataPointer As Long, ByVal numOfBytesToWrite As Long) As Boolean
    Dim tmpReturn As Long
    FileWriteData = (WriteFile(dstHandle, srcDataPointer, numOfBytesToWrite, tmpReturn, 0&) <> 0)
    If (Not FileWriteData) Then FSOError "FileWriteData", " failed to write to handle " & dstHandle & ".  Relevant error is " & Err.LastDllError & "."
End Function

'Given a base folder and some other path (with or without filename), generate a relative path between the two.
' It's assumed that thisFolder contains baseFolder within its path; if it doesn't, a copy of the full thisFolder string is returned.
Friend Function GenerateRelativePath(ByVal baseFolder As String, ByRef thisFolder As String, Optional ByVal normalizationCanBeSkipped As Boolean = False) As String
    
    'Start by forcing each string to have a trailing path
    If (Not normalizationCanBeSkipped) Then baseFolder = Me.PathAddBackslash(baseFolder)
    
    'Make sure a relative path is possible
    If (InStr(1, thisFolder, baseFolder, vbBinaryCompare) = 1) Then
        
        'Check equality first; equal strings mean no relative folder is required
        If Strings.StringsEqual(baseFolder, thisFolder) Then
            GenerateRelativePath = vbNullString
        
        'Strings are not equal, but baseFolder is contained within thisFolder.  Perfect!
        Else
            GenerateRelativePath = Right$(thisFolder, Len(thisFolder) - Len(baseFolder))
        End If
            
    Else
        GenerateRelativePath = thisFolder
    End If

End Function

'Given a (potentially) long path, with or without Unicode chars, return the short path.  This provides an easy way to use internal VB file functions,
' even on Unicode pathnames.
Friend Function GetShortPath(ByRef longPath As String) As String
    
    Dim tmpString As String
    tmpString = Space$(MAX_PATH)
    
    Dim lenFinalBuffer As Long
    lenFinalBuffer = GetShortPathNameW(StrPtr(longPath), StrPtr(tmpString), Len(tmpString))
    
    If (lenFinalBuffer > 0) Then
        GetShortPath = Left$(tmpString, lenFinalBuffer)
    Else
        GetShortPath = longPath
    End If
    
End Function

'Given a potential filename, scan and replace any invalid characters with a character of the caller's choosing
Friend Function MakeValidWindowsFilename(ByVal srcFilename As String, Optional ByVal replacementChar As String = "_") As String
    
    Dim strInvalidChars As String
    strInvalidChars = "\/*?""<>|:"
    
    'Replace each invalid character occurrence with whatever char the user specified
    Dim i As Long
    For i = 1 To Len(strInvalidChars)
        If InStr(1, srcFilename, Mid$(strInvalidChars, i, 1), vbBinaryCompare) Then srcFilename = Replace$(srcFilename, Mid$(strInvalidChars, i, 1), replacementChar)
    Next i
    
    MakeValidWindowsFilename = srcFilename

End Function

'Given a path (*without* a trailing filename), make sure the path ends in a slash character.  If it already
' ends in a slash, no change is made.  (Note that this function does *not* wrap the shlwapi function of the
' same name; it is custom-built.)
Friend Function PathAddBackslash(ByRef srcPath As String) As String
    If (Not Strings.StringsEqual(Right$(srcPath, 1), "\")) And (Not Strings.StringsEqual(Right$(srcPath, 1), "/")) Then
        PathAddBackslash = srcPath & "\"
    Else
        PathAddBackslash = srcPath
    End If
End Function
'
''Generate a "browse for folder" dialog.  Vista+ users get a fancy new interface; XP users get the shitty old interface.
'' Many thanks to vbForums user "LaVolpe", who wrote both the interfaces used here.  For implementation details and
'' links to the original, unmodified classes, please read the header comments in the referenced classes.
'Friend Function PathBrowseDialog(ByVal srcHwnd As Long, Optional ByVal initFolder As String = vbNullString, Optional ByVal dialogTitleText As String = vbNullString) As String
'
'    If (Len(dialogTitleText) = 0) Then dialogTitleText = g_Language.TranslateMessage("Please select a folder")
'
'    'Vista+ users get the fancy new "browse for folder" interface
'    If OS.IsVistaOrLater Then
'
'        Dim cBrowseNew As cFileDialogVista
'        Set cBrowseNew = New cFileDialogVista
'
'        With cBrowseNew
'            .propFlags = FOS__BrowseFoldersDefaults
'            If (Len(initFolder) <> 0) Then .propStartupFolder_Set initFolder, ppType_AsString
'
'            'For reasons I don't understand, a strange magic number is used to report cancellation; see the
'            ' function documentation for additional details.
'            If (.DialogShow(srcHwnd, FDLG_BROWSEFOLDERS, dialogTitleText) <> -2147023673) Then
'                PathBrowseDialog = .IShellItem_GetDisplayName(ObjPtr(.ResultsItem(1)), SIGDN_FILESYSPATH, False)
'            Else
'                PathBrowseDialog = vbNullString
'            End If
'        End With
'
'    'XP users get the crappy old browse interface.
'    Else
'
'        Dim cBrowse As cUnicodeBrowseFolders
'        Set cBrowse = New cUnicodeBrowseFolders
'
'        With cBrowse
'
'            If (Len(initFolder) <> 0) Then .InitialDirectory = initFolder
'            .dialogTitle = dialogTitleText
'            .Flags = BIF_RETURNONLYFSDIRS Or BIF_NEWDIALOGSTYLE
'
'            If .ShowBrowseForFolder(srcHwnd) Then
'                PathBrowseDialog = cBrowse.SelectedFolder
'            Else
'                PathBrowseDialog = vbNullString
'            End If
'
'        End With
'
'    End If
'
'    'PD folder functions enforce a trailing slash, to simplify subsequent concatenations
'    If (Len(PathBrowseDialog) <> 0) Then PathBrowseDialog = Me.PathAddBackslash(PathBrowseDialog)
'
'End Function

Friend Function PathCompact(ByRef srcString As String, ByVal newMaxLength As Long) As String
    
    'Limit length to MAX_PATH
    If (newMaxLength > MAX_PATH) Then newMaxLength = MAX_PATH
    
    'This API is weird because regardless of original length, the incoming string must be, per MSDN,
    ' "A pointer to a null-terminated string of length MAX_PATH that contains the path to be altered."
    ' So we must copy the incoming string into a MAX_PATH buffer
    Dim tmpStringSrc As String, copyLength As Long
    tmpStringSrc = String$(MAX_PATH, 0)
    
    copyLength = Len(srcString)
    If (copyLength > MAX_PATH) Then copyLength = MAX_PATH
    CopyMemoryStrict StrPtr(tmpStringSrc), StrPtr(srcString), copyLength * 2
    
    'Now, prep an output buffer of size MAX_PATH
    Dim tmpStringDst As String
    tmpStringDst = String$(MAX_PATH, 0)
    
    'Use the API to shrink the path
    If (PathCompactPathEx(StrPtr(tmpStringDst), StrPtr(tmpStringSrc), newMaxLength + 1, 0&) <> 0) Then
        PathCompact = Strings.TrimNull(tmpStringDst)
    Else
        PathCompact = tmpStringDst
    End If
    
End Function

'See if an arbitrary path contains a trailing filename.  Note that this *does not work* for detecting
' extension-less files, as they will be treated as folder paths.  This function is for fast validation of
' internal PD paths and files - if you are working with files directly, simply validate via FileExists().
Friend Function PathContainsFilename(ByVal fullPath As String) As Boolean
    
    'Perform a quick and dirty check for a "." later in the path than a "\" character.
    If (InStr(1, fullPath, "/", vbBinaryCompare) <> 0) Then fullPath = Replace$(fullPath, "/", "\", , , vbBinaryCompare)
    PathContainsFilename = (InStrRev(fullPath, ".", , vbBinaryCompare) > InStrRev(fullPath, "\", , vbBinaryCompare))
    
End Function

'Create a given directory.  Returns TRUE if folder already exists, or the folder was successfully created.  FALSE otherwise.
'
'An optional parameter, "createIntermediateFoldersAsNecessary", can be specified.  This instructs the function to create as many
' missing subfolders as are necessary in order to ensure the full source path exists.
Friend Function PathCreate(ByVal fullPath As String, Optional ByVal createIntermediateFoldersAsNecessary As Boolean = False) As Boolean
    
    PathCreate = False
    
    'If the folder already exists, great!
    If Me.PathExists(fullPath, False) Then
        PathCreate = True
        
    'The folder doesn't exist, so we need to make it
    Else
    
        Dim cSuccess As Long
        cSuccess = CreateDirectoryW(StrPtr(fullPath), 0&)
        PathCreate = (cSuccess <> 0)
        
        'Check for any relevant errors
        If (Not PathCreate) Then
            
            'PATH NOT FOUND means one or more parent folders do not exist.
            If (Err.LastDllError = ERROR_PATH_NOT_FOUND) Then
            
                'If the user wants us to construct intermediate folders, this is our chance to do so.
                If createIntermediateFoldersAsNecessary Then
                    
                    'Starting at the base and working our way forward, create all necessary folders.
                    
                    'First, normalize against backslashes
                    If (InStr(1, fullPath, "/", vbBinaryCompare) <> 0) Then fullPath = Replace$(fullPath, "/", "\")
                    fullPath = Me.PathAddBackslash(fullPath)
                    
                    'TODO: use a system-level normalize as well??
                    
                    'Assume a drive letter is present.  We'll use that as our starting point.
                    Dim startPos As Long, endPos As Long
                    startPos = InStr(1, fullPath, ":", vbBinaryCompare)
                    
                    'Find the first "\"
                    startPos = InStr(startPos, fullPath, "\", vbBinaryCompare)
                    
                    If (startPos > 0) Then
                    
                        'Extract the next folder in line
                        endPos = InStr(startPos + 1, fullPath, "\", vbBinaryCompare)
                        
                        Do While (endPos > 0)
                            
                            'We're going to ignore return values during creation of these child folders, because permissions may not exist
                            ' for creating some parent folders.  Instead, to verify that everything worked, we'll re-test the full folder
                            ' existence after attempting to create the entire hierarchy.
                            Me.PathCreate Left$(fullPath, endPos), False
                            
                            'Find the next folder in the path and repeat!
                            startPos = endPos
                            endPos = InStr(startPos + 1, fullPath, "\", vbBinaryCompare)
                            
                        Loop
                        
                        'Now we check to see if the full path exists; if it doesn't, we're SOL.
                        PathCreate = Me.PathExists(fullPath, True)
                    
                    Else
                        FSOError "PathCreate", "failed; no folders were specified in the path parameter...?"
                    End If
                    
                Else
                    FSOError "PathCreate", "failed; intermediate folders were missing."
                End If
            
            Else
                FSOError "PathCreate", "failed on an unknown error (#" & Err.LastDllError & ")."
            End If
            
        End If
        
    End If

End Function

'Returns a boolean as to whether or not a given directory exists.  Write access can optionally be tested, too.
Friend Function PathExists(ByRef fullPath As String, Optional ByVal checkWriteAccessAsWell As Boolean = True) As Boolean
    
    'First, make sure the directory exists
    Dim chkExistence As Boolean
    chkExistence = (Abs(GetFileAttributesW(StrPtr(fullPath))) And vbDirectory)
    
    'The folder was found, but the caller wants to know about write access, so a few more checks are needed.
    If (chkExistence And checkWriteAccessAsWell) Then
        
        Dim tmpFilename As String
        tmpFilename = Me.PathAddBackslash(fullPath) & "writecheck.tmp"
        
        'Attempt to open a temp file handle inside the given folder
        Dim hFile As Long
        If Me.FileCreateHandle(tmpFilename, hFile, True, True, OptimizeTempFile) Then
            
            'The folder is writable!  Close and kill the temp file.
            Me.FileCloseHandle hFile
            Me.FileDelete tmpFilename
            PathExists = True
            
        Else
            PathExists = False
        End If
        
    'If the caller doesn't care about write access, or the folder wasn't found, exit now.
    Else
        PathExists = chkExistence
    End If
    
End Function

'Partner function to FileConvertHandleToMMPtr, above.  The same values output by that function must be passed here,
' including an identical value for the base address.  Also, note that the flushImmediately parameter is asynchronous;
' it will not wait for the flush to cpmlete before returning, so you should take that into account if you plan to
' immediately access the written file via a separate method like ReadFile.  (If you want to enforce an immediate full
' write to the actual physical media, you'd need to follow this function with a manual call to FlushFileBuffers.)
Friend Function ReleaseMMHandle(ByVal srcMappedHandle As Long, ByVal srcBaseAddress As Long, Optional ByVal srcFileHandleToo As Long = 0, Optional ByVal flushImmediately As Boolean = False) As Boolean
    
    If (srcMappedHandle <> 0) And (srcBaseAddress <> 0) Then
        
        'Unmapping the view frees up the corresponding memory chunk in our address space; do this first
        UnmapViewOfFile srcBaseAddress
        
        'If the caller wants a flush, initiate it before we release the mapping handle
        If flushImmediately Then FlushViewOfFile srcBaseAddress, 0&
        
        'Release the mapped file object
        ReleaseMMHandle = (CloseHandle(srcMappedHandle) <> 0)
        
        'As a convenience to the caller, we can close the original file handle now as well.
        If (srcFileHandleToo <> 0) Then ReleaseMMHandle = ReleaseMMHandle And Me.FileCloseHandle(srcFileHandleToo)
        
    Else
        FSOError "ReleaseMMHandle", "passed null values.  Fix this!"
        ReleaseMMHandle = False
    End If
    
End Function

'Given a base folder, return all files within that folder (including subfolders).  Subfolder recursion is assumed,
' but can be waived by setting recurseSubfolders to FALSE.
'
'If returnRelativeStrings is true, strings are (obviously) returned relative to the base folder.
' So for e.g. base folder "C:\Folder", "C:\Folder\Subfolder\file.txt" will be returned as "Subfolder\file.txt".
'
'Returns TRUE if at least one file is found; FALSE otherwise.
' If the incoming dstFiles parameter already contains strings, TRUE will always be returned (by design).
'
'As an additional convenience, files can be restricted in one of two ways:
' 1) by restricting allowed extensions, using a pipe-delimited list of acceptable extensions (e.g. "jpg|bmp|gif")
' 2) by avoiding disallowed extensions, using a pipe-delimited list of unacceptable extensions (e.g. "bak|tmp")
'
'Obviously, you should use either the whitelist or the blacklist option, but never both.
Friend Function RetrieveAllFiles(ByVal srcFolder As String, ByRef dstFiles As pdStringStack, Optional ByVal recurseSubfolders As Boolean, Optional ByVal returnRelativeStrings As Boolean = True, Optional ByVal onlyAllowTheseExtensions As String = vbNullString, Optional ByVal doNotAllowTheseExtensions As String = vbNullString) As Boolean
    
    On Error GoTo RetrievalProblem
    
    'Retrieve just the base folder, and enforce strict trailing slash formatting if we were passed
    ' a bare path.
    Dim basePath As String
    If (Not Me.PathContainsFilename(srcFolder)) Then
        srcFolder = Me.PathAddBackslash(srcFolder)
        basePath = srcFolder
    Else
        basePath = Me.FileGetPath(srcFolder)
    End If
    
    'Initialize the destination stack as necessary.  Note that nothing happens if dstFiles is already initialized;
    ' this is by design, so the caller can concatenate multiple search results together if desired.
    If (dstFiles Is Nothing) Then Set dstFiles = New pdStringStack
    
    'This function was first used in PD as part of pdPackage's zip-like interface.  The goal was to create a
    ' convenient way to generate a folder-preserved list of files.  Consider a file tree like the following:
    ' C:\Folder\file.txt
    ' C:\Folder\SubFolder\subfile1.txt
    ' C:\Folder\SubFolder\AnotherFolder\subfile2.txt
    '
    'By calling this function with "C:\Folder\" as the base, this function will return a pdStringStack containing
    ' the following strings:
    ' file.txt
    ' SubFolder\subfile1.txt
    ' SubFolder\AnotherFolder\subfile2.txt
    '
    'This structure makes it very easy to duplicate a full file and folder structure into a new directory,
    ' which is exactly what pdPackage does when behaving like a .zip container.
    '
    'To that end, this function uses FindFirst/NextFileW to assemble a list of files relative to some base folder.
    ' Subfolders are not explicitly returned; the hope is that you can implicitly determine them from the relative
    ' filenames provided.  If you want a list of subfolders only, make a separate call to RetrieveAllFolders, below.
    '
    'Finally, two optional pattern parameters are available.  Use one or the other (or neither), but not both.
    '
    'onlyAllowTheseExtensions: used when the set of desired files uses a small subset of extensions.
    ' Separate valid extensions by pipe, e.g. "exe|txt".  Do not include ".".  If non-extension files are desired,
    ' YOU CANNOT USE THIS PARAMETER, as "||" doesn't parse correctly.
    '
    'doNotAllowTheseExtensions: used to blacklist unwanted file types.  Same rules as onlyAllowTheseExtensions applies,
    ' e.g. "bak|tmp" would be used to exclude .bak and .tmp files.
    
    'Because white/blacklisting is computationally expensive, we prepare separate boolean checks in advance
    Dim whiteListInUse As Boolean, blackListInUse As Boolean
    whiteListInUse = (LenB(onlyAllowTheseExtensions) <> 0)
    blackListInUse = (LenB(doNotAllowTheseExtensions) <> 0)
        
    'To further simplify searching for extension matches in the whitelist, we enforce strict formatting of the string
    If whiteListInUse Then
        If Not Strings.StringsEqual(Left$(onlyAllowTheseExtensions, 1), "|") Then onlyAllowTheseExtensions = "|" & onlyAllowTheseExtensions
        If Not Strings.StringsEqual(Right$(onlyAllowTheseExtensions, 1), "|") Then onlyAllowTheseExtensions = onlyAllowTheseExtensions & "|"
    End If
        
    If blackListInUse Then
        If Not Strings.StringsEqual(Left$(doNotAllowTheseExtensions, 1), "|") Then doNotAllowTheseExtensions = "|" & doNotAllowTheseExtensions
        If Not Strings.StringsEqual(Right$(doNotAllowTheseExtensions, 1), "|") Then doNotAllowTheseExtensions = doNotAllowTheseExtensions & "|"
    End If
        
    'To eliminate the need for expensive recursion, we use a stack structure to store subfolders that still need to be searched.
    ' Valid files are directly added to dstFiles as they are encountered.
    Dim cFoldersToCheck As pdStringStack
    Set cFoldersToCheck = New pdStringStack
    
    'Add the base folder to the collection, then start searching for subfolders.
    cFoldersToCheck.AddString srcFolder
        
    Dim curFolder As String, chkFile As String, modifiedPath As String
    Dim fileValid As Boolean
    
    Dim searchHandle As Long, findResult As WIN32_FIND_DATA
    Const STR_DOT As String = "."
    Const STR_DOTDOT As String = ".."
    Const LONG_PATH_PREPEND As String = "\\?\"
                            
    'The first folder doesn't get added to the destination collection, but all other folders do.
    Dim isNotFirstFolder As Boolean
    isNotFirstFolder = False
    
    Do While cFoldersToCheck.PopString(curFolder)
        
        If (LenB(curFolder) <> 0) Then
        
            'Prepend "\\?\" to sParam.  This enables long file paths.
            If Strings.StringsEqual(Left$(curFolder, 4), LONG_PATH_PREPEND) Then modifiedPath = curFolder Else modifiedPath = LONG_PATH_PREPEND & curFolder
            
            'FindFirstFile fails if the requested path has a trailing slash.
            If (Strings.StringsEqual(Right$(modifiedPath, 1), "\") Or Strings.StringsEqual(Right$(modifiedPath, 1), "/")) Then modifiedPath = modifiedPath & "*"
            
            'Retrieve the first file in the new search; this returns the search handle we'll use for subsequent searches
            searchHandle = FindFirstFileW(StrPtr(modifiedPath), VarPtr(findResult))
            
            'Ensure the search was initiated correctly.
            If (searchHandle = INVALID_HANDLE_VALUE) Then
                
                'No files found is fine, but if the caller screwed up the input path, we want to print some debug info.
                If (Err.LastDllError <> ERROR_FILE_NOT_FOUND) Then FSOError "RetrieveAllFiles", "possibly handed a bad path (" & modifiedPath & "). Please investigate."
                Exit Function
            
            'We obtained a good handle.  Store the first result, and continue searching until all results are discovered.
            Else
        
                Do
                    
                    chkFile = Strings.TrimNull(findResult.cFileName)
                    
                    'Ignore empty, ".", and ".." returns.
                    If (LenB(chkFile) <> 0) And (Strings.StringsNotEqual(chkFile, STR_DOT, True) And Strings.StringsNotEqual(chkFile, STR_DOTDOT, True)) Then
                    
                        'See if this chkFile iteration contains a folder
                        If ((findResult.dwFileAttributes And vbDirectory) <> 0) Then
                        
                            'This is a folder.  Add it to the "folders to check" collection if subfolders are being parsed.
                            If recurseSubfolders Then cFoldersToCheck.AddString Me.PathAddBackslash(curFolder & chkFile)
                            
                        'This is not a folder, but a file.  Add it to the destination file list if it meets any white/blacklisted criteria.
                        Else
                        
                            'White-listing check required
                            If whiteListInUse Then
                                fileValid = (InStr(1, onlyAllowTheseExtensions, "|" & Me.FileGetExtension(chkFile) & "|", vbBinaryCompare) <> 0)
                                
                            'Black-listing check required
                            ElseIf blackListInUse Then
                                fileValid = (InStr(1, doNotAllowTheseExtensions, "|" & Me.FileGetExtension(chkFile) & "|", vbBinaryCompare) = 0)
                                
                            'All files are allowed
                            Else
                                fileValid = True
                            End If
                        
                            'If we are allowed to add this file, perform a few final checks
                            If fileValid Then
                                If returnRelativeStrings Then
                                    dstFiles.AddString Me.GenerateRelativePath(basePath, Me.FileGetPath(curFolder) & chkFile, True)
                                Else
                                    dstFiles.AddString curFolder & chkFile
                                End If
                            End If
                        
                        'End file vs folder check
                        End If
                    
                    'End null, ".", ".." check
                    End If
                    
                Loop While (FindNextFileW(searchHandle, VarPtr(findResult)) <> 0)
                    
                'Close the search handle for this folder and report any irregularities
                If (searchHandle <> 0) Then FindClose searchHandle
                searchHandle = 0
                
                'Check for unexpected errors
                Dim lastDllErr As Long
                lastDllErr = Err.LastDllError
                If (lastDllErr <> 0) And (lastDllErr <> ERROR_NO_MORE_FILES) Then FSOError "RetrieveAllFiles", "terminated for a reason other than ERROR_NO_MORE_FILES. Please investigate.  (FYI: error # is last DLL error.)", lastDllErr
            
            End If
        
        End If
        
        'If recursion is enabled, any subfolders in this folder have now been added to cFolderToCheck, while all files in this folder
        ' are already present in the dstFiles collection.
        
    Loop
    
    'dstFiles now contains a full collection of files with the given base folder (and subfolders, if recursion is enabled).
    ' If at least one file was found, return TRUE. (Note that this value might be incorrect if the user sent us an already-populated
    ' string stack, but that's okay - the assumption is that they'll be processing the stack as one continuous list, so this return
    ' value won't be relevant anyway.)
    RetrieveAllFiles = (dstFiles.getNumOfStrings > 0)
    Exit Function
    
RetrievalProblem:
    
    'Clean up any outstanding handles
    If (searchHandle <> 0) Then FindClose searchHandle
    searchHandle = 0
    FSOError "RetrieveAllFiles", "VB error occurred: " & Err.Description, Err.Number
    
End Function

'Given a base folder, return all subfolders.  Recursion is assumed, but can be waived by setting recurseSubfolders to FALSE.
'
'If returnRelativeStrings is true, strings are (obviously) returned relative to the base folder.  So for e.g. base folder "C:\Folder",
' "C:\Folder\Subfolder" will be returned as just "Subfolder", while "C:\Folder\Subfolder\etc\" will return as "Subfolder\etc".
'
'Returns TRUE if at least one subfolder is found; FALSE otherwise.  If the incoming dstFolders parameter already contains strings,
' TRUE will always be returned.
Friend Function RetrieveAllFolders(ByVal srcFolder As String, ByRef dstFolders As pdStringStack, Optional ByVal recurseSubfolders As Boolean = True, Optional ByVal returnRelativeStrings As Boolean = True) As Boolean
    
    On Error GoTo RetrievalProblem
    
    'Enforce strict trailing slash formatting of the base folder
    srcFolder = Me.PathAddBackslash(srcFolder)
    
    'Initialize dstStrings as necessary.  Note that nothing happens if dstStrings is already initialized; this is by design, so the caller
    ' can concatenate multiple results together if desired.
    If (dstFolders Is Nothing) Then Set dstFolders = New pdStringStack
    
    'To eliminate the need for expensive recursion, we use a stack structure to store subfolders that still need to be searched.
    Dim cFoldersToCheck As pdStringStack
    Set cFoldersToCheck = New pdStringStack
    
    'Add the base folder to the collection, then start searching for subfolders.
    cFoldersToCheck.AddString srcFolder
    
    Dim curFolder As String, chkFile As String, modifiedPath As String
    Dim searchHandle As Long, findResult As WIN32_FIND_DATA
    Const STR_DOT As String = "."
    Const STR_DOTDOT As String = ".."
    Const LONG_PATH_PREPEND As String = "\\?\"
    
    'The first folder doesn't get added to the destination collection, but all other folders do.
    Dim isNotFirstFolder As Boolean
    isNotFirstFolder = False
    
    Do While cFoldersToCheck.PopString(curFolder)
        
        If (LenB(curFolder) <> 0) Then
        
            'Prepend "\\?\" to enable long file paths.
            If Strings.StringsEqual(Left$(curFolder, 4), LONG_PATH_PREPEND) Then modifiedPath = curFolder Else modifiedPath = LONG_PATH_PREPEND & curFolder
            
            'FindFirstFile fails if the requested path has a trailing slash.
            If (Strings.StringsEqual(Right$(modifiedPath, 1), "\") Or Strings.StringsEqual(Right$(modifiedPath, 1), "/")) Then modifiedPath = modifiedPath & "*"
            
            'Retrieve the first file in the new search; this returns the search handle we'll use for subsequent searches
            searchHandle = FindFirstFileW(StrPtr(modifiedPath), VarPtr(findResult))
            
            'Ensure the search was initiated correctly.
            If (searchHandle = INVALID_HANDLE_VALUE) Then
                
                'No files found is fine, but if the caller screwed up the input path, we want to print some debug info.
                If (Err.LastDllError <> ERROR_FILE_NOT_FOUND) Then FSOError "RetrieveAllFolders", "possibly handed a bad path (" & modifiedPath & "). Please investigate."
                Exit Function
            
            'We obtained a good handle.  Store the first result, and continue searching until all results are discovered.
            Else
        
                Do
                    
                    chkFile = Strings.TrimNull(findResult.cFileName)
                    
                    'Ignore empty, ".", and ".." returns.
                    If (LenB(chkFile) <> 0) And (Strings.StringsNotEqual(chkFile, STR_DOT, True) And Strings.StringsNotEqual(chkFile, STR_DOTDOT, True)) Then
                        
                        'See if this chkFile iteration contains a folder
                        If ((findResult.dwFileAttributes And vbDirectory) <> 0) Then
                            
                            'This is a folder.  Add it to the "folders to check" collection.
                            If recurseSubfolders Then
                                cFoldersToCheck.AddString Me.PathAddBackslash(curFolder & chkFile)
                            Else
                                If returnRelativeStrings Then
                                    dstFolders.AddString Me.GenerateRelativePath(srcFolder, Me.PathAddBackslash(curFolder & chkFile), True)
                                Else
                                    dstFolders.AddString Me.PathAddBackslash(curFolder & chkFile)
                                End If
                            End If
                            
                        End If
                        
                    'End null, ".", ".." check
                    End If
                    
                Loop While (FindNextFileW(searchHandle, VarPtr(findResult)) <> 0)
                    
                'Close the search handle for this folder and report any irregularities
                If (searchHandle <> 0) Then FindClose searchHandle
                searchHandle = 0
                
                'Check for unexpected errors
                Dim lastDllErr As Long
                lastDllErr = Err.LastDllError
                If (lastDllErr <> 0) And (lastDllErr <> ERROR_NO_MORE_FILES) Then FSOError "RetrieveAllFolders", "terminated for a reason other than ERROR_NO_MORE_FILES. Please investigate.  (FYI: error # is last DLL error.)", lastDllErr
                
                'Any subfolders in this folder have now been added to cFolderToCheck.  With this folder successfully processed, we can move it to
                ' the "folders checked" stack.
                If isNotFirstFolder Then
                    
                    If returnRelativeStrings Then
                        dstFolders.AddString Me.GenerateRelativePath(srcFolder, Me.PathAddBackslash(curFolder & chkFile), True)
                    Else
                        dstFolders.AddString curFolder
                    End If
                    
                Else
                    isNotFirstFolder = True
                End If
            
            'End valid FindFirstFileW handle check
            End If
        
        'End null path check
        End If
        
    Loop
    
    'dstFolders now contains a full collection of subfolders for the given base folder.  If subfolders were found, return TRUE.
    ' (Note that this value might be incorrect if the user sent us an already-populated string stack, but that's okay -
    ' the assumption is they'll process the stack as one continuous list, so individual returns don't matter.)
    RetrieveAllFolders = (dstFolders.getNumOfStrings > 0)
    Exit Function
    
RetrievalProblem:
    
    'Clean up any outstanding handles
    If (searchHandle <> 0) Then FindClose searchHandle
    searchHandle = 0
    FSOError "RetrieveAllFolders", "VB error occurred: " & Err.Description, Err.Number
    
End Function

'Helper function for FileGetTimeAsDate, above.  Once again, thank you to http://vb.mvps.org/hardcore/html/filedatestimes.htm
' for the original verison of this code.
Private Function Win32ToVBTime(ByRef fileTimeUTC As Currency) As Date
    
    'Convert from UTC time to local time
    Dim fileTimeLocal As Currency
    If FileTimeToLocalFileTime(fileTimeUTC, fileTimeLocal) Then
        
        'Local time is described as the number of nanoseconds since 01-01-1601.  Our use of a Currency value means we can
        ' treat these as milliseconds, instead.
        
        'Thus we divide by milliseconds per day to get the number of days elapsed since 1601; then we subtract the days from
        ' 1601 to 1899 to give us a corresponding value in VB's native "Date" format.
        Win32ToVBTime = CDate((fileTimeLocal / MILLISECONDS_PER_DAY) - DAY_ZERO_BIAS)
        
    Else
        Debug.Print "WARNING!  Files.Win32ToVbTime failed to convert the UTC value it was passed."
    End If
    
End Function

'INTERNAL FUNCTIONS FOLLOW

'Internal error messages should be passed through this function.  Note that individual functions may choose to handle the error
' and continue operation; that's fine.  (This function is primarily used for reporting, to reduce the number of internal #ifdefs.)
Private Sub FSOError(ByVal functionName As String, ByVal errText As String, Optional ByVal errNumber As Long = 0)
    
    If (LenB(functionName) <> 0) Then
        If (errNumber <> 0) Then
            Debug.Print "WARNING!  pdFSO." & functionName & " error #" & CStr(errNumber) & ": " & errText
        Else
            Debug.Print "WARNING!  pdFSO." & functionName & " problem: " & errText
        End If
    Else
        If (errNumber <> 0) Then
            Debug.Print "WARNING!  pdFSO error #" & CStr(errNumber) & ": " & errText
        Else
            Debug.Print "WARNING!  pdFSO problem: " & errText
        End If
    End If
        
End Sub

'Non-error messages containing potentially useful nightly build info can be routed here.  Feel free to disregard at your leisure.
Private Sub FSOFeedback(ByVal functionName As String, ByVal msgFeedback As String)
    Debug.Print "pdFSO." & functionName & " reports: " & msgFeedback
End Sub
