VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdFileMM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon Memory-Mapped File Interface
'Copyright 2014-2026 by Tanner Helland
'Created: 04/February/15
'Last updated: 06/June/19
'Last update: spin off from pdFSO
'
'This class provides a convenient interface to memory-mapped files.  It assumes that you already have
' experience with MM file APIs, including their constraints (particularly when writing to them).
'
'The class should be leak-proof but code reviews are always welcome.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Standard file handle, preserved for the life of the class.  This handle is required to produce a
' memory-mapped handle.
Private m_FileHandle As Long

'Memory-mapped file handle.  This is different from m_FileHandle, obviously.
Private m_FileHandleMM As Long

'pdFSO handles the actual API calls; we just wrap it in a user-friendly way
Private m_FSO As pdFSO

'When reading data from file, read-only access can improve performance, as data doesn't need to
' be flushed out to disk when map view as closed.  The caller specifies this when initializing
' this class.
Private m_WriteAccessRequired As Boolean

'Memory allocation granularity; set at class creation time and treated as a const thereafter
Private m_constAllocationGranularity As Long

'Minimum amount to allocate on each call to MapViewOfFile.  If the user requests more than this,
' we'll allocate whatever they require, but for small reads, we are forced to allocate at least
' the system allocation granularity anyway - and if we're traveling out to the HDD, we may as well
' make it worth our while and grab a good chunk of data while we're there.  At present, this is
' calculated as a multiplier (16x) of the allocation granularity (usually 64kb) - so at present,
' it will map 1 MB of data at a time (unless the source file is smaller than this, or the caller
' explicitly requests more than 1 MB).
Private m_constAllocationMinSize As Long
Private Const ALLOCATION_SIZE_MULT As Long = 16

'When starting a new file, this is the default file size to allocate.
Private Const DEFAULT_NEW_FILE_SIZE As Long = &HFFFF&   '64k

'Base address returned by the last call to MapViewOfFile, and its corresponding offset within
' the file (0-based).  If either value is set to UNMAPPED_VIEW, no view is currently mapped;
' use this to determine if a previous mapping needs to be released.
Private Const UNMAPPED_VIEW As Long = &HFFFFFFFF
Private m_MapAddress As Long, m_MapOffset As Long

'Size of the current mapping.  If set to UNMAPPED_VIEW, no view is currently mapped.
Private m_MapSize As Long

'Maximum size of the current file.  Mappings can *never* exceed this value.
Private m_MapLimit As Long

'Close all open views, close the memory-mapped handle, and close the underlying file itself.
Friend Sub CloseFile(Optional ByVal finalFileSize As Long = -1)
    ResetMapTrackers
    If (m_FileHandleMM <> 0) Then m_FSO.FileCloseHandle m_FileHandleMM
    If (m_FileHandle <> 0) Then
        If (finalFileSize >= 0) And m_WriteAccessRequired Then m_FSO.FileSetLength m_FileHandle, finalFileSize
        m_FSO.FileCloseHandle m_FileHandle
    End If
End Sub

'Before performing a read/write operation, call this function to ensure all necessary bytes have been
' mapped into view.  The required offset must always be absolute (e.g. 0-based, not relative to any
' previous calls).
'
'RETURNS: boolean TRUE if the map was successful; FALSE otherwise.  FALSE likely means the map is
' too large and you need to break your data into smaller pieces, or you have requested an
' offset or offset+size combination that extends beyond the limits of the initial map.
' (Remember: memory-mapped files require you to specify the upper limit of the file size IN ADVANCE.)
'
'Three other pieces of data are returned (and all are important!).  First is the pointer you can use
' to interact with the memory-mapped file.  IMPORTANTLY, this address is automatically scaled so
' that it corresponds to 0 in your source offset scale (as determined by offsetRequired).  This is a
' deliberate design decision to simplify adding offsets to the returned pointer; use it to avoid
' calling this function every time you read/write data.  (Instead, only call this function when
' writing data that extends beyond the dimensions of the last map you created.)  Note that any/all
' boundary alignment requirements were already been taken care of by this function.
'
'Besides the returned address, this function (optionally) returns two absolute positions (0-based),
' defining the start and ending offsets between which you can safely write.  These are *not* pointers
' - they are offsets, using the same scale as your offsetRequired value.  As long as your current
' offset (for the start position) and offset + length (for the end position) are within these values,
' , you do not need to call this function, because the map is guaranteed to cover that data.
' Again, this provides a way to avoid repeat calls to this function during sequential read/write
' operations of short lengths.
'
'Finally, note that you can pass any sizeRequired value you want to this function; if you pass a
' negative value, the function will attempt to map the entire underlying file into view.
Friend Function EnsureMapAvailable(ByVal offsetRequired As Long, ByVal sizeRequired As Long, ByRef dstPtr As Long, Optional ByRef dstMinOffsetForThisMap As Long, Optional ByRef dstMaxOffsetForThisMap As Long) As Boolean
    
    'If the user passes an invalid size, assume they want the remainder of the file mapped into view.
    If (sizeRequired <= 0) Then sizeRequired = m_MapLimit - offsetRequired
    
    'Is a map already open?  If so, start by seeing if it covers the requested offset+size combination.
    ' (If a map exists, but it *doesn't* cover the requested range, we'll free the current map and allocate
    ' a new one.)
    If (m_MapAddress <> UNMAPPED_VIEW) Then
        
        'Ensure the required offset falls within the current map.
        Dim curMapOK As Boolean
        curMapOK = (offsetRequired >= m_MapOffset) And (offsetRequired < m_MapOffset + m_MapSize)
        
        'If the starting offset is okay, check the offset + size as well
        If curMapOK Then curMapOK = (offsetRequired + sizeRequired >= m_MapOffset) And (offsetRequired + sizeRequired < m_MapOffset + m_MapSize)
        
        'If the current map is acceptable, return our current map values and exit immediately
        If curMapOK Then
            dstPtr = m_MapAddress - m_MapOffset
            dstMinOffsetForThisMap = m_MapOffset
            dstMaxOffsetForThisMap = m_MapOffset + m_MapSize - 1
            EnsureMapAvailable = True
            Exit Function
        
        '...otherwise, unmap the current view because we need to map a new portion of the file into view
        Else
            m_FSO.FileUnmapView m_MapAddress
            m_MapAddress = UNMAPPED_VIEW
        End If
        
    End If
    
    'If we're still here, the current map (if any) is insufficient.  We need to allocate a new one.
    
    'Convert the required offset to a multiple of the current allocation granularity.
    Dim offsetDifference As Long
    m_MapOffset = Int(offsetRequired \ m_constAllocationGranularity) * m_constAllocationGranularity
    offsetDifference = offsetRequired - m_MapOffset
    
    'Determine how much size we need to allocate (accounting for any extra bytes incurred by
    ' the base offset calculation, above).  Sizes don't matter so much here, because the OS
    ' automatically handles padding for us.
    sizeRequired = sizeRequired + offsetDifference
    
    'We now have a decision to make - if the caller is attempting to write past the end of the
    ' current file, MapViewOfFile is guaranteed to fail.  The only way to allow this is to
    ' close our current memory map handle, forcibly resize the underlying file, then re-create
    ' our map from scratch.
    If (m_MapOffset + sizeRequired > m_MapLimit) Then
        
        'Note that we can only auto-resize the target file if it's backed by an actual file!
        ' Otherwise, the map only exists in memory, and it will be deleted (permanently)
        ' when freed!
        If (m_FileHandle <> 0) Then
        
            'We have no choice but to recreate our map from scratch.  Start by releasing our current handle.
            m_FSO.FileCloseHandle m_FileHandleMM
            
            'Using our existing file handle, resize the file to the required size
            m_MapLimit = m_MapOffset + sizeRequired
            If (Not m_FSO.FileSetLength(m_FileHandle, m_MapLimit)) Then InternalError "EnsureMapAvailable", "couldn't automatically resize underlying file", Err.LastDllError
            
            'Re-create our map
            If (Not m_FSO.FileGetMMHandle_Local(m_FileHandle, m_FileHandleMM, 0&, m_WriteAccessRequired)) Then InternalError "EnsureMapAvailable", "couldn't reacquire map handle", Err.LastDllError
            
        Else
            InternalError "EnsureMapAvailable", "page-backed file cannot be auto-resized - you must back with an *actual* file for this feature!"
        End If
        
    Else
        
        'As a "bonus" optimization, use a minimum allocation size (m_constAllocationMinSize)
        ' *unless* that would require us to resize the underlying file.
        If (sizeRequired < m_constAllocationMinSize) And (m_MapOffset + m_constAllocationMinSize < m_MapLimit) Then sizeRequired = m_constAllocationMinSize
        
    End If
    
    'Map limits should be safely set above, but perform one last failsafe check
    If (m_MapLimit - m_MapOffset < sizeRequired) Then InternalError "EnsureMapAvailable", "consider: sizeRequired = m_MapLimit - m_MapOffset"
    m_MapSize = sizeRequired
    
    'Create the map view
    m_MapAddress = m_FSO.FileMapView(m_FileHandleMM, m_MapOffset, m_MapSize, m_WriteAccessRequired)
    
    If (m_MapAddress = 0) Then m_MapAddress = UNMAPPED_VIEW
    EnsureMapAvailable = (m_MapAddress <> UNMAPPED_VIEW)
    
    'If we were successful, rebase the returned pointer against offset 0, and return it alongside
    ' min/max allowable offset values.
    If EnsureMapAvailable Then
    
        'The returned pointer is for m_MapOffset.  Rescale it against the offset the user supplied.
        dstPtr = m_MapAddress - m_MapOffset
        
        'Tell the user the minimum/maximum allowable offsets to which they can write.
        dstMinOffsetForThisMap = m_MapOffset
        dstMaxOffsetForThisMap = m_MapOffset + m_MapSize - 1
    
        'Debug.Print "Mapped region [" & dstMinOffsetForThisMap & ", " & dstMaxOffsetForThisMap & "]"
    
    'Mapping failed; offset or size is likely bad
    Else
        InternalError "EnsureMapAvailable", "FileMapView failed", Err.LastDllError
    End If
    
End Function

'Use a memory-mapped interface to quickly dump a pdStream object to file.  This typically provides
' improved throughput (as the memory-to-file portion becomes asynchronous).  If chunkSize is negative,
' the function will attempt to map the entire target file into memory; the success of this is obviously
' contingent on available memory vs the size of the source stream.  (By default, this function attempts
' to write the target file in 32-mb increments; limited local testing showed this to be a relatively
' safe sweet spot between memory consumption and performance.)
Friend Function FastDumpStreamToFile(ByRef srcStream As pdStream, ByRef dstFile As String, Optional ByVal chunkSize As Long = 2 ^ 25) As Boolean
    
    Dim dataTotal As Long
    dataTotal = srcStream.GetStreamSize()
    
    'If no chunk size is specified, write out the stream in 32 MB chunks.
    ' (Limited local testing showed this to be a relatively safe "sweet spot" between
    ' memory consumption and mapping performance.)
    If (chunkSize < 0) Then
        chunkSize = dataTotal
    ElseIf (chunkSize > dataTotal) Then
        chunkSize = dataTotal
    End If
    
    'We don't need to worry about deleting the target file, because we're about to overwrite
    ' and/or truncate it with the source stream's data.
    'If m_FSO.FileExists(dstFile) Then m_FSO.FileDelete dstFile
    
    FastDumpStreamToFile = Me.StartNewFile(dstFile, dataTotal, OptimizeSequentialAccess)
    If FastDumpStreamToFile Then
        
        FastDumpStreamToFile = True
        
        Dim dataWritten As Long
        dataWritten = 0
        
        Dim writeSize As Long, dstPtr As Long, mapOK As Boolean
        
        Do While (dataWritten < dataTotal)
        
            'Write either a full chunk of data, or the remaining data left in the stream
            ' (if that's less than a full chunk)
            If (dataWritten + chunkSize < dataTotal) Then
                writeSize = chunkSize
            Else
                writeSize = dataTotal - dataWritten
            End If
            
            'Attempt the map, and if it fails, reduce chunk size to the nearest power of two.
            mapOK = Me.EnsureMapAvailable(dataWritten, writeSize, dstPtr)
            Do While (Not mapOK)
                writeSize = NearestPowerOfTwo(writeSize \ 2)
                If (writeSize <= 2) Then Exit Do
            Loop
            
            'If the map worked, copy our data over and advance the source pointer
            If mapOK Then
                CopyMemoryStrict dstPtr + dataWritten, srcStream.Peek_PointerOnly(dataWritten, writeSize), writeSize
                dataWritten = dataWritten + writeSize
                chunkSize = writeSize
            Else
                InternalError "FastDumpStreamToFile", "failed to allocate a map at any size"
                Exit Do
            End If
            
            FastDumpStreamToFile = FastDumpStreamToFile And mapOK
            
        Loop
        
        'Release all open handles and exit
        Me.CloseFile dataTotal
        
    Else
        InternalError "FastDumpStreamToFile", "couldn't open target file; double-check path!"
    End If

End Function

Friend Function GetMMHandle() As Long
    GetMMHandle = m_FileHandleMM
End Function

'Open an existing file for memory-mapped file access.  If you don't need write access, specify as much
' because performance will likely improve (as pages don't need to be written out to file when you're
' done with them).
Friend Function OpenExistingFile(ByRef srcFile As String, Optional ByVal writeAccessRequired As Boolean = False, Optional ByVal optimizeFileAccess As PD_FILE_ACCESS_OPTIMIZE = OptimizeNone) As Boolean
    
    'Always start by closing our existing file, if any
    Me.CloseFile
    
    'Ensure the source file exists!
    If Files.FileExists(srcFile) Then
        m_MapLimit = m_FSO.FileLenW(srcFile)
        m_WriteAccessRequired = writeAccessRequired
        OpenExistingFile = m_FSO.FileCreateHandle(srcFile, m_FileHandle, True, writeAccessRequired, optimizeFileAccess)
        If OpenExistingFile Then
            OpenExistingFile = m_FSO.FileGetMMHandle_Local(m_FileHandle, m_FileHandleMM, 0&, writeAccessRequired)
            If (Not OpenExistingFile) Then InternalError "Couldn't get mm handle", "DLL error #" & Err.LastDllError
        Else
            InternalError "OpenExistingFile", "couldn't get file handle", "DLL error #" & Err.LastDllError
        End If
    End If
    
End Function

Friend Sub UnmapAllViews()
    ResetMapTrackers
End Sub

'Start a new memory-mapped file for writing.  Note that the maximum size of the file *must* be specified
' in advance if the destination file does not exist (e.g. if you are using a temporary page-backed file).
' Per the API docs, "After a file mapping object is created, the size of the file must not exceed the
' size of the file mapping object; if it does, not all of the file contents are available for sharing."
' This makes memory-mapping inconvenient when writing files, as you should know an upper limit of
' required space in advance - and when you are done with the file, you may need to forcibly set a
' byte-accurate EOF to replace the "maxPossibleSize" passed here.
'
'For maps backed by an actual on-disk file, this class can work around the "know max size in advance"
' restriction by auto-remapping the file if/when it grows.  See EnsureMapAvailable() for details on how
' we achieve this.
Friend Function StartNewFile(ByRef dstFile As String, ByVal maxPossibleSize As Long, Optional ByVal optimizeFileAccess As PD_FILE_ACCESS_OPTIMIZE = OptimizeNone) As Boolean
    
    'Always start by closing our existing file, if any
    Me.CloseFile
    
    'Create the new file and establish its maximum length
    m_WriteAccessRequired = True
    If (maxPossibleSize <= 0) Then maxPossibleSize = DEFAULT_NEW_FILE_SIZE
    
    'Note that a filename is *not* required; per MSDN, you can pass INVALID_FILE_HANDLE as the
    ' source file handle, and the system will use the page file as backing (instead of an
    ' actual file out on disk).
    If (LenB(dstFile) <> 0) Then
        StartNewFile = m_FSO.FileCreateHandle(dstFile, m_FileHandle, True, True, optimizeFileAccess)
        If StartNewFile Then StartNewFile = m_FSO.FileSetLength(m_FileHandle, maxPossibleSize) Else InternalError "StartNewFile", "couldn't allocate initial file handle", Err.LastDllError
    Else
        m_FileHandle = 0
    End If
    
    m_MapLimit = maxPossibleSize
    
    'Create a memory-mapped handle to wrap the file we just created
    StartNewFile = m_FSO.FileGetMMHandle_Local(m_FileHandle, m_FileHandleMM, maxPossibleSize, True)
    
End Function

'Internal error messages should be passed through this function.  Note that individual functions may choose to handle the error
' and continue operation; that's fine.  (This function is primarily used for reporting, to reduce the number of internal #ifdefs.)
Private Sub InternalError(ByRef functionName As String, ByRef errText As String, Optional ByVal errNumber As Long = 0)
    
    If (LenB(functionName) <> 0) Then
        If (errNumber <> 0) Then
            PDDebug.LogAction "WARNING!  pdFileMM." & functionName & " error #" & CStr(errNumber) & ": " & errText
        Else
            PDDebug.LogAction "WARNING!  pdFileMM." & functionName & " problem: " & errText
        End If
    Else
        If (errNumber <> 0) Then
            PDDebug.LogAction "WARNING!  pdFileMM error #" & CStr(errNumber) & ": " & errText
        Else
            PDDebug.LogAction "WARNING!  pdFileMM problem: " & errText
        End If
    End If
    
End Sub

'Cheap and easy way to find the nearest power of two.  Note that you could also do this with logarithms
' (or even bitshifts maybe?) but I haven't thought about it hard enough lol
Private Function NearestPowerOfTwo(ByVal srcNumber As Long) As Long
    
    Dim curPower As Long
    curPower = 1
    
    Do While (curPower < srcNumber)
        curPower = curPower * 2
    Loop
    
    NearestPowerOfTwo = curPower
    
End Function

Private Sub ResetMapTrackers()
    If (m_MapAddress <> UNMAPPED_VIEW) Then m_FSO.FileUnmapView m_MapAddress
    m_MapAddress = UNMAPPED_VIEW
    m_MapOffset = UNMAPPED_VIEW
    m_MapSize = UNMAPPED_VIEW
End Sub

Private Sub Class_Initialize()
    Set m_FSO = New pdFSO
    m_constAllocationGranularity = m_FSO.GetSystemAllocationGranularity
    m_constAllocationMinSize = m_constAllocationGranularity * ALLOCATION_SIZE_MULT
End Sub

Private Sub Class_Terminate()
    CloseFile
End Sub
