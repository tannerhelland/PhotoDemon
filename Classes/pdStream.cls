VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'Generic Stream-like I/O Interface
'Copyright 2016-2026 by Tanner Helland
'Created: 05/December/16
'Last updated: 27/January/26
'Last update: relent and implement FindByte for memory-mapped streams
'
'Before I say anything else, please note that this class is based off work originally done by vbforums.com user
' "dilettante", whose original clsBasicBuffer class served PD well for a long time.  clsBasicBuffer is a nice,
' simple, memory-only stream interface that beginners will find much more accessible than this class.  You can
' download the original version of clsBasicBuffer from the following link (good as of December '16):
' http://www.vbforums.com/showthread.php?710155-VB6-BasicBuffer-Binary-Stream-Class
'
'After hacking the original clsBasicBuffer to pieces to improve performance, it finally reached a point where it
' was time to migrate to a new solution designed from the ground-up for performance.  In addition, PD really needs
' a generic stream object that can be either file-backed or memory-backed, and given the work involved in handling
' both scenarios efficiently, it was easier to rework the class's design to match.
'
'Significant upgrades to pdStream include:
' - Allocations are now controlled automatically, without any user input.  The only user input allowed is a hint for
'   the initial allocation size; this can be passed to the StartBuffer function, and it's very helpful if you have
'   some knowledge of how big the final buffer is likely to be.  If you don't have such knowledge, the default
'   allocator will now handle this intelligently based on the size of incoming write requests.
' - Byte arrays are no longer returned directly from functions.  Instead, the caller must pass their own arrays
'   as parameters.  This class will only resize those arrays as necessary, and the caller can specifically turn off
'   array trimming.  (Without trimming, the array will only be resized if it is currently too small to hold the
'   returned data - and it's up to the caller to check the returned data size and react accordingly.)  This saves a
'   ton of time if we're retrieving multiple large nodes, as we can simply reuse a single very large array.
' - A framework is now place for backing a stream by a file or a memory buffer.  File-backed buffers allow you to
'   stream data directly to a file on persistent media, without requiring a matching in-memory copy.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Declare Function htonl Lib "Ws2_32" (ByVal srcLong As Long) As Long
Private Declare Function htons Lib "Ws2_32" (ByVal srcShort As Integer) As Integer
Private Declare Function ntohl Lib "Ws2_32" (ByVal srcLong As Long) As Long
Private Declare Function ntohs Lib "Ws2_32" (ByVal srcShort As Integer) As Integer

Public Enum PD_STREAM_MODE
    PD_SM_MemoryBacked = 0
    PD_SM_FileBacked = 1
    PD_SM_ExternalPtrBacked = 2
    PD_SM_FileMemoryMapped = 3
End Enum

#If False Then
    Private Const PD_SM_MemoryBacked = 0, PD_SM_FileBacked = 1, PD_SM_ExternalPtrBacked = 2, PD_SM_FileMemoryMapped = 3
#End If

Public Enum PD_STREAM_ACCESS
    PD_SA_ReadWrite = 0
    PD_SA_ReadOnly = 1
End Enum

#If False Then
    Private Const PD_SA_ReadWrite = 0, PD_SA_ReadOnly = 1
#End If

'The current stream mode.  If it is file-backed, a filename will also be present.
Private m_StreamMode As PD_STREAM_MODE

'The current stream access mode.  Note that there is only "read/write" or "read-only".  There is no write-only.  This value
' really only matters for file streams, as read-only access can be accelerated thanks to memory mapping.
Private m_StreamAccess As PD_STREAM_ACCESS

'For in-memory streams, this array stores the actual stream data.
Private m_MemBuffer() As Byte

'If an in-memory stream is constructed as READ-ONLY around an EXTERNAL POINTER, we use this SafeArray to store
' our "artificially" constructed array header.  It should be freed manually before the class goes out of scope.
Private m_TemporarySA As SafeArray1D

'For file-backed streams, pdFSO is used for file interactions
Private m_FSO As pdFSO

'For memory-mapped file streams, pdFileMM is used for file interactions
Private m_FileMM As pdFileMM

'When writing a stream, this is the current stream size.  It is updated on each write.  (Importantly, for memory streams,
' this is *not* the UBound() of m_MemBuffer()!  m_MemBuffer's UBound() will always be >= this number.)
Private m_BufferSize As Long

'Current position inside the stream.  Starts at 0 and increases from there.  Cannot be negative.  Can be modified via Seek().
Private m_Pointer As Long

'Once a stream has been opened (by a call to the StartStream function), this will be set to TRUE.  This value is important
' for ensuring that all internals have been properly prepped prior to issuing read/write requests.
Private m_Open As Boolean

'While in file-backed mode, the caller can specify an offset to use as position "zero" inside the file.  This is very helpful
' for pdPackages, because we can ignore the file header and treat the beginning of the data chunk as "position 0".
Private m_FilePointerOffset As Long

'File-backed mode requires a number of other pointers and trackers
Private m_FileHandle As Long

'Returns TRUE if at least numBytesRequired exist in the current stream; FALSE otherwise.
Friend Function AreBytesAvailable(ByVal readLength As Long) As Boolean
    If m_Open Then
        AreBytesAvailable = Not (Me.GetPosition() + readLength > m_BufferSize)
    Else
        AreBytesAvailable = False
    End If
End Function

'Remove (n) bytes from the stream, starting at the front of the stream.  Any un-removed bytes will be shifted to
' the front of the stream, and the current pointer will be shifted accordingly.  (That said, if the current pointer
' would be moved before the front of the stream, it will be forcibly reset to 0.)
'
'NOTE: at present, only memory-backed streams support this function
Friend Sub DeleteFromStart(ByVal numBytesToDelete As Long)
    
    If (m_StreamAccess = PD_SA_ReadOnly) Then
        InternalStreamError "DeleteFromStart() cannot operate on read-only streams."
        Exit Sub
    End If
    
    'Ensure the number of bytes to delete is valid
    If (numBytesToDelete > m_BufferSize) Then numBytesToDelete = m_BufferSize
    If (numBytesToDelete <= 0) Then Exit Sub
    
    If (m_StreamMode = PD_SM_MemoryBacked) Then
    
        'If there is a difference in size between the number of bytes to remove, and the size of the buffer,
        ' shift any remaining bytes downward.  (IMPORTANTLY, note that this would normally require overlapped I/O,
        ' which carries huge performance penalties - to avoid this, we use a temp buffer.)
        If (numBytesToDelete < m_BufferSize) Then
        
            'Note that we allocate an extra byte, "just in case"
            Dim tmpBuffer() As Byte
            ReDim tmpBuffer(0 To m_BufferSize - numBytesToDelete) As Byte
            CopyMemoryStrict VarPtr(tmpBuffer(0)), VarPtr(m_MemBuffer(numBytesToDelete)), m_BufferSize - numBytesToDelete
            
            'Now, copy the same memory into place over the buffer itself
            CopyMemoryStrict VarPtr(m_MemBuffer(0)), VarPtr(tmpBuffer(0)), m_BufferSize - numBytesToDelete
            Erase tmpBuffer
            
            'Shift the pointer accordingly
            m_Pointer = m_Pointer - numBytesToDelete
            If (m_Pointer < 0) Then m_Pointer = 0
            
            'Also shift the current buffer size
            m_BufferSize = m_BufferSize - numBytesToDelete
            If (m_BufferSize < 0) Then m_BufferSize = 0
            
        'If we're deleting the entire buffer, just reset the pointer and size variables, but leave the
        ' existing memory untouched.
        Else
            m_Pointer = 0
            m_BufferSize = 0
        End If
        
    Else
        InternalStreamError "sorry, DeleteFromStart is only supported for memory-backed streams"
    End If

End Sub

'Return the current stream pointer position.
Friend Function GetPosition() As Long
    If (m_StreamMode = PD_SM_MemoryBacked) Then
        GetPosition = m_Pointer
    ElseIf (m_StreamMode = PD_SM_ExternalPtrBacked) Then
        GetPosition = m_Pointer
    ElseIf (m_StreamMode = PD_SM_FileBacked) Then
        If (m_FileHandle <> 0) And (Not m_FSO Is Nothing) Then GetPosition = m_FSO.FileGetCurrentPointer(m_FileHandle) - m_FilePointerOffset
    ElseIf (m_StreamMode = PD_SM_FileMemoryMapped) Then
        GetPosition = m_Pointer
    End If
End Function

'Set a new stream pointer position.
'
'RETURNS: TRUE if successful; FALSE otherwise.  FALSE may occur if you attempt to set the pointer outside the current stream range.
Friend Function SetPosition(ByVal newPosition As Long, Optional ByVal startingPosition As FILE_POINTER_MOVE_METHOD = FILE_BEGIN) As Boolean
    
    'In file-backed mode, we only need special handling for the FILE_BEGIN method, if the user has specified a custom offset
    ' into the file
    If (m_StreamMode = PD_SM_FileBacked) Then
        If (startingPosition = FILE_BEGIN) Then newPosition = newPosition + m_FilePointerOffset
    
    'All other modes must calculate the pointer manually
    Else
        If (startingPosition = FILE_BEGIN) Then
            'Do nothing; newPosition is an absolute value
        ElseIf (startingPosition = FILE_CURRENT) Then
            newPosition = m_Pointer + newPosition
        ElseIf (startingPosition = FILE_END) Then
            newPosition = m_BufferSize + newPosition
        End If
    End If
    
    'Now we can actually apply the new position
    If (newPosition >= 0) And (newPosition <= m_BufferSize) Then
        
        If (m_StreamMode = PD_SM_FileBacked) Then
            If (Not m_FSO Is Nothing) Then SetPosition = m_FSO.FileMovePointer(m_FileHandle, newPosition, startingPosition)
        Else
            m_Pointer = newPosition
            SetPosition = True
        End If
        
    Else
        InternalStreamError "SetPosition failed; bad position: " & newPosition & " requested vs " & m_BufferSize & " total buffer len"
        SetPosition = False
    End If
    
End Function

Friend Function GetStreamAccess() As PD_STREAM_ACCESS
    GetStreamAccess = m_StreamAccess
End Function

Friend Function GetStreamMode() As PD_STREAM_MODE
    GetStreamMode = m_StreamMode
End Function

'Return the number of bytes in the current stream.  Note that this is *not* necessarily the size of the in-memory buffer;
' rather, it is the number of bytes the caller has written.
' (Conversely, for a read-only stream, this is always the full size of the stream.)
Friend Function GetStreamSize() As Long
    GetStreamSize = m_BufferSize
End Function

'WARNING: DO NOT USE THIS FUNCTION IF YOU DON'T KNOW WHAT YOU'RE DOING
'
'PD sometimes writes buffer bytes using 3rd-party libraries and naked pointers (e.g. during compression/decompression).
' Because of this, the buffer's size may change without us accessing any internal stream functions.  (We still ensure a
' given stream size using EnsureBufferSpaceAvailable(), but then we write data to the stream without using any of this
' class's helper functions.)  To ensure that the buffer's reported size remains correct, call this function afterward.
Friend Sub SetSizeExternally(ByVal newSize As Long)
    m_BufferSize = newSize
End Sub

'If StartStream has been called, this will return TRUE.
Friend Function IsOpen() As Boolean
    IsOpen = m_Open
End Function

'To UNSAFELY wrap an external byte array around our data, use this sub.  When you are done - and before the external
' array goes out of scope - you MUST call UnwrapArray, below.  (Also, per the name, this only works on memory streams -
' file streams will fail!)
'
'An optional startPosition index can be passed.  The array will be wrapped around the stream *starting at this position*.
' (Note that its base index will still be 0.)  By default, the array's length is always set to the difference between the
' start position and the full size of the stream.
Friend Function WrapArrayAroundMemoryStream(ByRef srcArray() As Byte, ByRef srcSafeArray As SafeArray1D, Optional ByVal startPosition As Long = 0&) As Boolean
    
    WrapArrayAroundMemoryStream = False
    
    If m_Open And (m_StreamMode = PD_SM_MemoryBacked) Then
        
        With srcSafeArray
            .cbElements = 1
            .cDims = 1
            .cLocks = 1
            .lBound = 0
            .cElements = Me.GetStreamSize() - startPosition
            .pvData = VarPtr(m_MemBuffer(startPosition))
        End With
        
        PutMem4 VarPtrArray(srcArray()), VarPtr(srcSafeArray)
        WrapArrayAroundMemoryStream = True
        
    End If
    
End Function

Friend Sub UnwrapArrayFromMemoryStream(ByRef srcArray() As Byte)
    PutMem4 VarPtrArray(srcArray), 0&
End Sub

'Start a new stream.  The stream type *must* be specified.  If this is a file stream, a valid filename *must* be provided.
' Read/write access does not need to be specified, but if you only need a read-only buffer, specifying as much will yield
' improved performance, particularly for file streams.  (Similarly, when wrapping an external pointer, read-only access
' will be forcibly set.)
'
'Optionally, when creating a writable buffer, you can also specify a starting buffer size.  If you have some notion of the
' stream's net size in advance, an initial memory allocation that comes close to your required value can spare us intermediate
' allocations while the stream is constructed.  Similarly, if you have already been using this stream for other tasks,
' you can mark the "reuseExistingBuffer" parameter as TRUE; this will leave the stream in its current state instead of
' reallocating memory, potentially reducing memory thrashing.
'
'When wrapping an external pointer, use the optional startingBufferSize and baseFilePointerOffset values to the source
' data's length and base pointer, respectively.
'
'Returns TRUE if successful; FALSE otherwise.  FALSE should only occur if...
' 1) This is a file-backed stream and you supplied an invalid or inaccessible filename, or...
' 2) This is a memory-backed stream and the initial starting size is ridiculously huge.
' 3) This is an external pointer-backed stream and the initial size and/or pointer value is zero.
' 4) This is a memory-mapped file stream and the filename is bad, or you requested write access but supplied an initial size <= 0.
'
'If this stream object is already open, the existing stream will be closed and a new one will be started.
Friend Function StartStream(ByVal streamType As PD_STREAM_MODE, Optional ByVal streamAccess As PD_STREAM_ACCESS = PD_SA_ReadWrite, Optional ByVal srcFilename As String, Optional ByVal startingBufferSize As Long = 0, Optional ByVal baseFilePointerOffset As Long = 0, Optional ByVal optimizeAccess As PD_FILE_ACCESS_OPTIMIZE = OptimizeNone, Optional ByVal reuseExistingBuffer As Boolean = False) As Boolean
    
    'If we're already open, close the current stream before starting a new one.
    If m_Open Then Me.StopStream (Not reuseExistingBuffer)
    
    m_StreamMode = streamType
    m_StreamAccess = streamAccess
    
    If (streamType = PD_SM_MemoryBacked) Then
    
        'To improve performance when performing back-to-back stream read/write operations from the same instance,
        ' we attempt to reuse existing memory as much as possible.
        Dim needToAllocate As Boolean
        needToAllocate = Not VBHacks.IsArrayInitialized(m_MemBuffer)
        
        'The caller can specify a starting buffer size.  If they don't, we'll start small - just 4k.
        If (startingBufferSize <= 0) Then
            Const INITIAL_BUFFER_SIZE As Long = 4096
            startingBufferSize = INITIAL_BUFFER_SIZE
        End If
        
        'Allocation is performed if...
        ' (1) we haven't allocated anything yet (checked above), or...
        ' (2) our current allocation is smaller than the minimum initial size, or the user's required size
        If (Not needToAllocate) Then needToAllocate = (UBound(m_MemBuffer) < startingBufferSize - 1) Or (Not reuseExistingBuffer)
        If needToAllocate Then ReDim m_MemBuffer(0 To startingBufferSize - 1) As Byte
        
        m_Open = True
        
    ElseIf (streamType = PD_SM_FileBacked) Then
    
        If (m_FSO Is Nothing) Then Set m_FSO = New pdFSO
        
        If m_FSO.FileCreateHandle(srcFilename, m_FileHandle, True, (streamAccess = PD_SA_ReadWrite), optimizeAccess) Then
            
            m_Open = True
            m_FilePointerOffset = baseFilePointerOffset
            
            'Automatically set the current "buffer size" to match the size of the file, minus any base pointer offset
            If (m_FSO.FileLenW(srcFilename) > 0) Then
                m_BufferSize = m_FSO.FileLenW(srcFilename) - baseFilePointerOffset
            Else
                m_BufferSize = baseFilePointerOffset
            End If
            
            'Move the file pointer to the baseFilePointerOffset, if any
            If (baseFilePointerOffset <> 0) Then m_FSO.FileMovePointer m_FileHandle, baseFilePointerOffset, FILE_BEGIN
        
        End If
        
    ElseIf (streamType = PD_SM_ExternalPtrBacked) Then
    
        m_StreamAccess = PD_SA_ReadOnly
        m_BufferSize = startingBufferSize
        
        If (m_BufferSize <> 0) And (baseFilePointerOffset <> 0) Then
            
            'Populate a replacement sa header, and use it to "initialize" our internal memory buffer
            With m_TemporarySA
                .cbElements = 1
                .cDims = 1
                .cLocks = 1
                .lBound = 0
                .cElements = m_BufferSize
                .pvData = baseFilePointerOffset
            End With
            
            Erase m_MemBuffer
            PutMem4 VarPtrArray(m_MemBuffer()), VarPtr(m_TemporarySA)
            
            m_Open = True
            
        Else
            InternalStreamError "External pointer-backed streams require a non-zero buffer size and pointer."
            m_Open = False
        End If
        
    ElseIf (streamType = PD_SM_FileMemoryMapped) Then
        
        Set m_FileMM = New pdFileMM
        If (m_FSO Is Nothing) Then Set m_FSO = New pdFSO
        
        If (streamAccess = PD_SA_ReadOnly) Then
            m_BufferSize = m_FSO.FileLenW(srcFilename)
            m_Open = m_FileMM.OpenExistingFile(srcFilename, False, optimizeAccess)
        Else
            
            If m_FSO.FileExists(srcFilename) Then
                
                If (startingBufferSize > m_FSO.FileLenW(srcFilename)) Then
                    Dim hFileTmp As Long
                    If m_FSO.FileCreateHandle(srcFilename, hFileTmp, True, True, OptimizeSequentialAccess) Then
                        m_FSO.FileSetLength hFileTmp, startingBufferSize
                        m_FSO.FileCloseHandle hFileTmp
                    Else
                        InternalStreamError "failed to set startingbuffersize", Err.LastDllError
                    End If
                End If
                
                m_Open = m_FileMM.OpenExistingFile(srcFilename, True, optimizeAccess)
                
            Else
                m_Open = m_FileMM.StartNewFile(srcFilename, startingBufferSize, optimizeAccess)
            End If
            
        End If
            
    End If
    
    StartStream = m_Open
    If StartStream Then
        m_Pointer = 0
    Else
        InternalStreamError "Could not start stream"
    End If
    
End Function

'Shut down the current stream.  For file-backed streams, this closes the current file handle.  For memory streams, this just
' resets our internal buffers to null values.
Friend Sub StopStream(Optional ByVal alsoFreeMemory As Boolean = True)
    
    If m_Open Then
    
        If (m_StreamMode = PD_SM_MemoryBacked) Then
            If alsoFreeMemory Then Erase m_MemBuffer
        ElseIf (m_StreamMode = PD_SM_FileBacked) Then
            If (Not m_FSO Is Nothing) Then m_FSO.FileCloseHandle m_FileHandle
            m_FileHandle = 0
        ElseIf (m_StreamMode = PD_SM_ExternalPtrBacked) Then
            PutMem4 VarPtrArray(m_MemBuffer), 0&
        ElseIf (m_StreamMode = PD_SM_FileMemoryMapped) Then
            If (m_StreamAccess = PD_SA_ReadOnly) Then m_FileMM.CloseFile Else m_FileMM.CloseFile m_BufferSize
        End If
        
        m_Open = False
        m_Pointer = 0
        m_BufferSize = 0
        
    End If
    
End Sub

'When parsing certain filetypes, you sometimes have to look for things like null-terminated strings.
' This can be a pain in the ass to do manually, so you can ask this class do it for you.  Optionally, you can
' also ask it to move the current pointer to the first-found occurrence - BUT IMPORTANTLY, note that the
' pointer *will not be moved* if the target byte isn't found.
'
'For safety reasons, you may also want to pass a "maxSearchLength" value.  This value is calculated as an
' offset FROM THE CURRENT POINTER POSITION.  If the default -1 is passed, the search will continue until EOF.
' (This is especially important when searching a memory-mapped file, as the entire search space will be
'  mapped into memory.  The caller may want to manually search smaller chunks instead.)
'
'BY DESIGN, this feature is not implemented for file-backed streams.  It's too damn slow, so you're better
' off caching a chunk of the file and searching it manually.  (PhotoDemon almost always uses memory-mapped
' file interactions, so this isn't relevant there.)
'
'RETURNS: position >= 0 where the byte was found; -1 if it was not found (or if your starting position is invalid)
Friend Function FindByte(ByVal targetByte As Byte, Optional ByVal startPosition As Long = -1, Optional ByVal movePtrToMatch As Boolean = True, Optional ByVal maxSearchLength As Long = -1) As Long
    
    FindByte = -1
    
    If m_Open Then
    
        If (startPosition < 0) Then startPosition = Me.GetPosition()
        If (startPosition < m_BufferSize) Then
            
            Dim posFinal As Long
            
            'Memory-backed modes are easy
            If (m_StreamMode = PD_SM_MemoryBacked) Or (m_StreamMode = PD_SM_ExternalPtrBacked) Then
                
                If (maxSearchLength <= 0) Then
                    posFinal = m_BufferSize
                Else
                    posFinal = startPosition + maxSearchLength
                    If (posFinal > m_BufferSize) Then posFinal = m_BufferSize
                End If
                
                'TODO: look at wrapping an API for this; it'd potentially be faster
                Do
                    If (m_MemBuffer(startPosition) = targetByte) Then Exit Do
                    startPosition = startPosition + 1
                Loop While (startPosition < posFinal)
                
                If (startPosition < posFinal) Then
                    FindByte = startPosition
                    If movePtrToMatch Then m_Pointer = startPosition
                End If
            
            'Jan 2026: okay, okay - I relented and added this version for memory-mapped files.
            ' (A generic solution like this is never going to be as fast as a custom-built solution
            '  for a specific task, but this should work "well enough" for most use-cases, provided
            '  the search surface isn't enormously large.)
            ElseIf (m_StreamMode = PD_SM_FileMemoryMapped) Then
                
                'Failsafe check for valid underlying memory-map interface
                If (m_FileMM Is Nothing) Then
                    InternalStreamError "FindByte failed; m_FileMM invalid."
                    Exit Function
                End If
                
                'Find the correct end-point of the search (may require a large map!)
                If (maxSearchLength <= 0) Then
                    posFinal = m_BufferSize
                Else
                    posFinal = startPosition + maxSearchLength
                    If (posFinal > m_BufferSize) Then posFinal = m_BufferSize
                End If
                
                'Failsafe only to ensure a non-null map region
                Dim searchSize As Long
                searchSize = posFinal - startPosition
                If (searchSize <= 0) Then Exit Function
                
                'Map the search region into memory
                Dim srcPtr As Long, srcReadMax As Long, tmpArray() As Byte, tmpArraySA As SafeArray1D
                srcPtr = 0
                srcReadMax = 0
                
                If m_FileMM.EnsureMapAvailable(startPosition, searchSize, srcPtr, srcReadMax) Then
                    VBHacks.WrapArrayAroundPtr_Byte tmpArray, tmpArraySA, startPosition + srcPtr, searchSize
                Else
                    PDDebug.LogAction "WARNING: FindByte failed due to memory map failure"
                    Exit Function
                End If
                
                'Instead of searching by actual position, we must search by artificial 0-offset
                ' into the search region.
                Dim initPos As Long
                initPos = 0
                
                Do
                    If (tmpArray(initPos) = targetByte) Then Exit Do
                    initPos = initPos + 1
                Loop While (initPos < searchSize)
                
                'Always free the temporary array!
                VBHacks.UnwrapArrayFromPtr_Byte tmpArray
                
                'Remember to add the search offset back to the *original* pointer.
                If (initPos < searchSize) Then
                    FindByte = startPosition + initPos
                    If movePtrToMatch Then m_Pointer = FindByte
                End If
                
            'Do not use this in file-backed mode(s)!  Manually cache a chunk of the file and perform the
            ' search yourself!
            Else
                InternalStreamError "FindByte is not implemented for this stream mode"
            End If
        
        End If
        
    End If
    
End Function

'Return a copy of the stream, starting from the current stream position (which you can/should move prior to reading,
' as relevant) and extending the supplied readLength amount of bytes.  If readLength is not supplied, a full copy of
' the entire stream from the current position to the end will be returned.
'
'RETURNS: number of bytes read if successful; zero otherwise.
'
'NOTE: if the optional trimDestinationArray parameter is TRUE, the destination array will be exactly sized to
'      match the outgoing data.  If trimDestinationArray is FALSE, the destination array will only be resized if
'      necessary to fit the data.  This can be very helpful for performance, as you can reuse the same destination
'      array for reading data from the stream, saving expensive allocations.
Friend Function ReadBytes(ByRef dstBytes() As Byte, Optional ByVal readLength As Long = -1, Optional ByVal trimDestinationArray As Boolean = True) As Long
    
    If m_Open Then
        
        'Calculate the length we need to move the current position marker
        If (readLength < 0) Then
            readLength = m_BufferSize - Me.GetPosition()
        
        'If the caller supplied a length value, make sure it doesn't extend past the end of the buffer
        Else
            If (Me.GetPosition() + readLength > m_BufferSize) Then readLength = m_BufferSize - Me.GetPosition()
            If (readLength < 0) Then
                InternalStreamError "ReadBytes calculated a read length < 0 (" & readLength & ")"
                ReadBytes = False
                Exit Function
            End If
        End If
        
        'Check to see if our destination array is already initialized to the perfect size.  If it is, we don't need to
        ' allocate new memory.
        If VBHacks.IsArrayInitialized(dstBytes) Then
            
            Dim arrLBound As Long, arrUBound As Long
            arrLBound = LBound(dstBytes)
            arrUBound = UBound(dstBytes)
            
            Dim requiredSize As Long, mustResize As Long
            requiredSize = (arrUBound - arrLBound) + 1
            
            'If the user wants the destination array trimmed to exact size, see if we need to modify the destination
            ' array's size to fit.
            If trimDestinationArray Then
                mustResize = (requiredSize <> readLength)
                
            'If the user doesn't want the destination array trimmed to exact size, we only need to resize it if it's
            ' smaller than the amount of data being read.
            Else
                mustResize = (requiredSize < readLength)
            End If
            
            If mustResize Then ReDim dstBytes(arrLBound To arrLBound + readLength - 1) As Byte
            
        'If the destination array is *not* initialized, make it the perfect size regardless of trimDestinationArray.
        Else
            arrLBound = 0
            ReDim dstBytes(0 To readLength - 1) As Byte
        End If
        
        'Rely on the bare pointer function from here
        ReadBytes = ReadBytesToBarePointer(VarPtr(dstBytes(arrLBound)), readLength)
        
    Else
        InternalStreamError "stream closed"
    End If
    
End Function

'Return a copy of the stream, starting from the current stream position (which you can/should move prior to reading,
' as relevant) and extending the supplied readLength amount of bytes.  If readLength is not supplied, a full copy of
' the entire stream from the current position to the end will be returned.
'
'RETURNS: number of bytes read if successful; zero otherwise.
Friend Function ReadBytesToBarePointer(ByVal dstPointer As Long, Optional ByVal readLength As Long = -1) As Long

    If m_Open Then
        
        'Calculate the length we need to move the current position marker
        If (readLength < 0) Then
            readLength = m_BufferSize - Me.GetPosition()
        
        'If the caller supplied a length value, make sure it doesn't extend past the end of the buffer
        Else
            If (Me.GetPosition() + readLength > m_BufferSize) Then
                Debug.Print "WARNING!  Requested read extends past the end of the stream."
                readLength = m_BufferSize - Me.GetPosition()
            End If
        End If
        
        'Copy the relevant amount of bytes into position
        If (m_StreamMode = PD_SM_MemoryBacked) Or (m_StreamMode = PD_SM_ExternalPtrBacked) Then
            CopyMemoryStrict dstPointer, VarPtr(m_MemBuffer(m_Pointer)), readLength
            
            'In memory mode, we must manually move the pointer to match the new position
            m_Pointer = m_Pointer + readLength
            If (m_Pointer > m_BufferSize) Then m_Pointer = m_BufferSize
        
        ElseIf (m_StreamMode = PD_SM_FileBacked) Then
            If (Not m_FSO Is Nothing) Then m_FSO.FileReadData m_FileHandle, dstPointer, readLength
        ElseIf (m_StreamMode = PD_SM_FileMemoryMapped) Then
            If (Not m_FileMM Is Nothing) Then
                Dim srcPtr As Long, srcReadMax As Long
                If m_FileMM.EnsureMapAvailable(m_Pointer, readLength, srcPtr, srcReadMax) Then
                    CopyMemoryStrict dstPointer, srcPtr + m_Pointer, readLength
                    m_Pointer = m_Pointer + readLength
                Else
                    InternalStreamError "ReadBytesToBarePointer failed; const pdFileMM for details."
                End If
            End If
        End If
        
        ReadBytesToBarePointer = readLength
        
    End If
    
End Function

'Read (n) bytes to a destination stream object.  This is helpful if the bytes require further parsing.
' Note that the returned stream will be at position (0) by default.
Friend Function ReadBytesToStream(ByRef dstStream As pdStream, ByVal readLength As Long) As Boolean

    If m_Open Then
        
        'Make sure the read length doesn't extend past the end of the buffer
        If (Me.GetPosition() + readLength > m_BufferSize) Then readLength = m_BufferSize - Me.GetPosition()
        
        'Prep the destination stream
        If (dstStream Is Nothing) Then Set dstStream = New pdStream
        ReadBytesToStream = dstStream.StartStream(PD_SM_MemoryBacked, PD_SA_ReadWrite, startingBufferSize:=readLength, reuseExistingBuffer:=True)
        
        'Copy the relevant amount of bytes into position
        If (m_StreamMode = PD_SM_MemoryBacked) Or (m_StreamMode = PD_SM_ExternalPtrBacked) Then
            CopyMemoryStrict dstStream.Peek_PointerOnly(0), VarPtr(m_MemBuffer(m_Pointer)), readLength
            
            'In memory mode, we must manually move the pointer to match the new position
            m_Pointer = m_Pointer + readLength
            If (m_Pointer > m_BufferSize) Then m_Pointer = m_BufferSize
        
        ElseIf (m_StreamMode = PD_SM_FileBacked) Then
            If (Not m_FSO Is Nothing) Then m_FSO.FileReadData m_FileHandle, dstStream.Peek_PointerOnly(0), readLength
        ElseIf (m_StreamMode = PD_SM_FileMemoryMapped) Then
            If (Not m_FileMM Is Nothing) Then
                Dim srcPtr As Long, srcReadMax As Long
                If m_FileMM.EnsureMapAvailable(m_Pointer, readLength, srcPtr, srcReadMax) Then
                    CopyMemoryStrict dstStream.Peek_PointerOnly(0), srcPtr + m_Pointer, readLength
                    m_Pointer = m_Pointer + readLength
                Else
                    InternalStreamError "ReadBytesToStream failed; const pdFileMM for details."
                End If
            End If
        End If
        
        'Make sure the buffer size of the destination stream accurately reflects however many bytes we copied.
        dstStream.SetSizeExternally readLength
        
    End If
    
End Function

'Read bytes from the stream.  Instead of returning the bytes (which requires an expensive copy operation), this function
' simply returns a pointer to the relevant position in the stream, which the user can use however they want.  An optional
' readLength parameter controls how far the position pointer is moved within the stream, but it does not otherwise
' affect anything EXCEPT on memory-mapped files - for these, it is critical to supply a valid read length, because only
' that many bytes will get mapped into memory.
'
'NOTE: for file-backed streams, *if* the stream was opened in read-only mode, this function is still valid.  If the file
' stream is opened in read/write mode, this function *cannot be used*.  Instead, you must use the ReadBytes() function, above.
'
'RETURNS: non-zero pointer if successful; zero otherwise.  Zero will be returned if the current stream is file-backed,
' and the stream was opened with read/write access.  (Memory-mapped access is currently only supported for read-only access,
' for performance reasons.)
'
'IMPORTANT NOTE: adding data to the buffer may force it to allocate new memory, invalidating all previously returned
' pointers.  You must use the return of this function immediately, as its correctness is not guaranteed after any other
' class functions are called.
Friend Function ReadBytes_PointerOnly(Optional ByVal readLength As Long = -1) As Long
    
    If m_Open Then
        
        If (m_StreamMode = PD_SM_MemoryBacked) Or (m_StreamMode = PD_SM_ExternalPtrBacked) Then
        
            'Return a pointer to the current buffer position
            ReadBytes_PointerOnly = VarPtr(m_MemBuffer(m_Pointer))
            
            'Calculate the length we need to move the current position marker
            If (readLength < 0) Then readLength = m_BufferSize - m_Pointer
            
            'Move the pointer
            m_Pointer = m_Pointer + readLength
            
            'As a failsafe, check for invalid pointer positions
            If (m_Pointer > m_BufferSize) Then m_Pointer = m_BufferSize
        
        ElseIf (m_StreamMode = PD_SM_FileBacked) Then
            
            'This function is unsupported for file-backed streams, as there's no pointer
            ' to reference!
            InternalStreamError "ReadBytes_PointerOnly can't be used on file-backed streams"
            ReadBytes_PointerOnly = 0
            
        ElseIf (m_StreamMode = PD_SM_FileMemoryMapped) Then
            If (Not m_FileMM Is Nothing) Then
                
                'Calculate the length we need to move the current position marker
                If (readLength < 0) Then readLength = m_BufferSize - m_Pointer
                
                Dim srcPtr As Long, srcReadMax As Long
                If m_FileMM.EnsureMapAvailable(m_Pointer, readLength, srcPtr, srcReadMax) Then
                    
                    ReadBytes_PointerOnly = srcPtr + m_Pointer
                    
                    'Move the pointer
                    m_Pointer = m_Pointer + readLength
                    If (m_Pointer > m_BufferSize) Then m_Pointer = m_BufferSize
            
                Else
                    InternalStreamError "ReadBytes_PointerOnly failed; const pdFileMM for details."
                End If
                
            Else
                InternalStreamError "ReadBytes_PointerOnly doesn't have a valid m_FileMM object!"
            End If
        End If
        
    End If
    
End Function

'Retrieve a single byte from the source data.  There is no way to return success or failure from this function,
' so you will need to manually validate things like stream position vs length when using this.
Friend Function ReadByte() As Byte
    If m_Open And (Me.GetPosition() + 1& <= m_BufferSize) Then
        If (m_StreamMode = PD_SM_MemoryBacked) Or (m_StreamMode = PD_SM_ExternalPtrBacked) Then
            GetMem1 VarPtr(m_MemBuffer(m_Pointer)), ReadByte
            m_Pointer = m_Pointer + 1&
        ElseIf (m_StreamMode = PD_SM_FileBacked) Then
            If (Not m_FSO Is Nothing) Then m_FSO.FileReadData m_FileHandle, VarPtr(ReadByte), 1&
        ElseIf (m_StreamMode = PD_SM_FileMemoryMapped) Then
            If (Not m_FileMM Is Nothing) Then
                Dim srcPtr As Long, srcReadMax As Long
                If m_FileMM.EnsureMapAvailable(m_Pointer, 1, srcPtr, srcReadMax) Then
                    GetMem1 srcPtr + m_Pointer, ReadByte
                    m_Pointer = m_Pointer + 1&
                Else
                    InternalStreamError "ReadByte failed; const pdFileMM for details."
                End If
            End If
        End If
    End If
End Function

'Retrieve a 16-bit word from the source data.  There is no way to return success or failure from this function,
' so you will need to manually validate things like stream position vs length when using this.
Friend Function ReadInt() As Integer
    If m_Open And (Me.GetPosition() + 2& <= m_BufferSize) Then
        If (m_StreamMode = PD_SM_MemoryBacked) Or (m_StreamMode = PD_SM_ExternalPtrBacked) Then
            GetMem2 VarPtr(m_MemBuffer(m_Pointer)), ReadInt
            m_Pointer = m_Pointer + 2&
        ElseIf (m_StreamMode = PD_SM_FileBacked) Then
            If (Not m_FSO Is Nothing) Then m_FSO.FileReadData m_FileHandle, VarPtr(ReadInt), 2&
        ElseIf (m_StreamMode = PD_SM_FileMemoryMapped) Then
            If (Not m_FileMM Is Nothing) Then
                Dim srcPtr As Long, srcReadMax As Long
                If m_FileMM.EnsureMapAvailable(m_Pointer, 2, srcPtr, srcReadMax) Then
                    GetMem2 srcPtr + m_Pointer, ReadInt
                    m_Pointer = m_Pointer + 2&
                Else
                    InternalStreamError "ReadInt failed; const pdFileMM for details."
                End If
            End If
        End If
    End If
End Function

'Same as ReadInt(), above, but with endianness conversion.
Friend Function ReadInt_BE() As Integer
    Dim tmpInteger As Integer
    tmpInteger = Me.ReadInt()
    ReadInt_BE = ntohs(tmpInteger)
End Function

'Retrieve an unsigned 16-bit integer from the source data, but copy it into a VB Long to make processing easier.
' Note that there is no way to return success or failure from this function, so you will need to manually validate
' things like stream position vs length.
Friend Function ReadIntUnsigned() As Long
    If m_Open And (Me.GetPosition() + 2& <= m_BufferSize) Then
        If (m_StreamMode = PD_SM_MemoryBacked) Or (m_StreamMode = PD_SM_ExternalPtrBacked) Then
            CopyMemoryStrict VarPtr(ReadIntUnsigned), VarPtr(m_MemBuffer(m_Pointer)), 2&
            m_Pointer = m_Pointer + 2&
        ElseIf (m_StreamMode = PD_SM_FileBacked) Then
            If (Not m_FSO Is Nothing) Then m_FSO.FileReadData m_FileHandle, VarPtr(ReadInt) + 2&, 2&
        ElseIf (m_StreamMode = PD_SM_FileMemoryMapped) Then
            If (Not m_FileMM Is Nothing) Then
                Dim srcPtr As Long, srcReadMax As Long
                If m_FileMM.EnsureMapAvailable(m_Pointer, 2, srcPtr, srcReadMax) Then
                    CopyMemoryStrict VarPtr(ReadIntUnsigned), srcPtr + m_Pointer, 2&
                    m_Pointer = m_Pointer + 2&
                Else
                    InternalStreamError "ReadIntUnsigned failed; const pdFileMM for details."
                End If
            End If
        End If
    End If
End Function

'Same as ReadIntUnsigned(), above, but with endianness conversion.
Friend Function ReadIntUnsigned_BE() As Long
    Dim tmpInteger As Integer, tmpInteger2 As Integer
    tmpInteger = Me.ReadInt()
    tmpInteger2 = ntohs(tmpInteger)
    ReadIntUnsigned_BE = (tmpInteger2 And &HFFFF&)
End Function

Friend Function ReadLong() As Long
    If m_Open And (Me.GetPosition() + 4& <= m_BufferSize) Then
        If (m_StreamMode = PD_SM_MemoryBacked) Or (m_StreamMode = PD_SM_ExternalPtrBacked) Then
            GetMem4 VarPtr(m_MemBuffer(m_Pointer)), ReadLong
            m_Pointer = m_Pointer + 4&
        ElseIf (m_StreamMode = PD_SM_FileBacked) Then
            If (Not m_FSO Is Nothing) Then m_FSO.FileReadData m_FileHandle, VarPtr(ReadLong), 4&
        ElseIf (m_StreamMode = PD_SM_FileMemoryMapped) Then
            If (Not m_FileMM Is Nothing) Then
                Dim srcPtr As Long, srcReadMax As Long
                If m_FileMM.EnsureMapAvailable(m_Pointer, 4, srcPtr, srcReadMax) Then
                    GetMem4 srcPtr + m_Pointer, ReadLong
                    m_Pointer = m_Pointer + 4&
                Else
                    InternalStreamError "ReadLong failed; const pdFileMM for details."
                End If
            End If
        End If
    End If
End Function

'Same as ReadLong(), above, but with endianness conversion.
Friend Function ReadLong_BE() As Long
    Dim tmpLong As Long
    tmpLong = Me.ReadLong()
    ReadLong_BE = ntohl(tmpLong)
End Function

'Adobe files used fixed-point floats in some places.
Friend Function ReadFixed1616() As Single
    If m_Open And (Me.GetPosition() + 4& <= m_BufferSize) Then
        Dim intPart As Long, floatPart As Long
        intPart = Me.ReadInt()
        floatPart = Me.ReadIntUnsigned()
        ReadFixed1616 = CSng(intPart) + CSng(floatPart) / 65535!
    End If
End Function

Friend Function ReadFixed1616_BE() As Single
    If m_Open And (Me.GetPosition() + 4& <= m_BufferSize) Then
        Dim intPart As Long, floatPart As Long
        intPart = Me.ReadInt_BE()
        floatPart = Me.ReadIntUnsigned_BE()
        ReadFixed1616_BE = CSng(intPart) + CSng(floatPart) / 65536!
    End If
End Function

Friend Function ReadFloat() As Single
    If m_Open And (Me.GetPosition() + 4& <= m_BufferSize) Then
        If (m_StreamMode = PD_SM_MemoryBacked) Or (m_StreamMode = PD_SM_ExternalPtrBacked) Then
            GetMem4 VarPtr(m_MemBuffer(m_Pointer)), ByVal VarPtr(ReadFloat)
            m_Pointer = m_Pointer + 4&
        ElseIf (m_StreamMode = PD_SM_FileBacked) Then
            If (Not m_FSO Is Nothing) Then m_FSO.FileReadData m_FileHandle, VarPtr(ReadFloat), 4&
        ElseIf (m_StreamMode = PD_SM_FileMemoryMapped) Then
            If (Not m_FileMM Is Nothing) Then
                Dim srcPtr As Long, srcReadMax As Long
                If m_FileMM.EnsureMapAvailable(m_Pointer, 4, srcPtr, srcReadMax) Then
                    GetMem4 srcPtr + m_Pointer, ByVal VarPtr(ReadFloat)
                    m_Pointer = m_Pointer + 4&
                Else
                    InternalStreamError "ReadFloat failed; const pdFileMM for details."
                End If
            End If
        End If
    End If
End Function

'Same as ReadFloat(), above, but with endianness conversion.
Friend Function ReadFloat_BE() As Single
    Dim tmpLong As Long
    tmpLong = ntohl(Me.ReadLong())
    GetMem4_Ptr VarPtr(tmpLong), VarPtr(ReadFloat_BE)
End Function

Friend Function ReadDouble() As Double
    If m_Open And (Me.GetPosition() + 8& <= m_BufferSize) Then
        If (m_StreamMode = PD_SM_MemoryBacked) Or (m_StreamMode = PD_SM_ExternalPtrBacked) Then
            GetMem8_Ptr VarPtr(m_MemBuffer(m_Pointer)), ByVal VarPtr(ReadDouble)
            m_Pointer = m_Pointer + 8&
        ElseIf (m_StreamMode = PD_SM_FileBacked) Then
            If (Not m_FSO Is Nothing) Then m_FSO.FileReadData m_FileHandle, VarPtr(ReadDouble), 8&
        ElseIf (m_StreamMode = PD_SM_FileMemoryMapped) Then
            If (Not m_FileMM Is Nothing) Then
                Dim srcPtr As Long, srcReadMax As Long
                If m_FileMM.EnsureMapAvailable(m_Pointer, 8, srcPtr, srcReadMax) Then
                    GetMem8_Ptr srcPtr + m_Pointer, VarPtr(ReadDouble)
                    m_Pointer = m_Pointer + 8&
                Else
                    InternalStreamError "ReadDouble failed; const pdFileMM for details."
                End If
            End If
        End If
    End If
End Function

'Same as ReadDouble(), above, but with endianness conversion.
Friend Function ReadDouble_BE() As Double
    
    ReadDouble_BE = Me.ReadDouble()
    
    'There's no good way to reverse bytes in pure VB code, so we opt for an extremely ugly way to do it. Yay?
    Dim i As Long, tmpByte As Byte
    For i = 0 To 3
        GetMem1 VarPtr(ReadDouble_BE) + i, tmpByte
        GetMem1_Ptr VarPtr(ReadDouble_BE) + (7 - i), VarPtr(ReadDouble_BE) + i
        PutMem1 VarPtr(ReadDouble_BE) + (7 - i), tmpByte
    Next i
    
End Function

'Read a series of one-byte values from the stream, and treat the chars as if they are individual ASCII codes.
' A proper BSTR is returned.  If the requested length is invalid (i.e. it extends beyond the end of the stream),
' a null string will be returned, *not* a partial string.
Friend Function ReadString_ASCII(ByVal numOfChars As Long) As String

    If m_Open And (Me.GetPosition() + numOfChars <= m_BufferSize) Then
        If (m_StreamMode = PD_SM_MemoryBacked) Or (m_StreamMode = PD_SM_ExternalPtrBacked) Then
            ReadString_ASCII = Strings.StringFromCharPtr(VarPtr(m_MemBuffer(m_Pointer)), False, numOfChars, True)
            m_Pointer = m_Pointer + numOfChars
        ElseIf (m_StreamMode = PD_SM_FileBacked) Then
            If (Not m_FSO Is Nothing) Then
                Dim tmpBytes() As Byte
                ReDim tmpBytes(0 To numOfChars - 1) As Byte
                m_FSO.FileReadData m_FileHandle, VarPtr(tmpBytes(0)), numOfChars
                ReadString_ASCII = Strings.StringFromCharPtr(VarPtr(tmpBytes(0)), False, numOfChars, True)
            End If
        ElseIf (m_StreamMode = PD_SM_FileMemoryMapped) Then
            If (Not m_FileMM Is Nothing) Then
                Dim srcPtr As Long, srcReadMax As Long
                If m_FileMM.EnsureMapAvailable(m_Pointer, numOfChars, srcPtr, srcReadMax) Then
                    ReadString_ASCII = Strings.StringFromCharPtr(srcPtr + m_Pointer, False, numOfChars, True)
                    m_Pointer = m_Pointer + numOfChars
                Else
                    InternalStreamError "ReadString_ASCII failed; const pdFileMM for details."
                End If
            End If
        End If
    End If

End Function

'Read a series of two-byte values from the stream, and treat each 16-bit word as a UTF-16 code.
' If the requested length is invalid (i.e. it extends beyond the end of the stream), a null string
' will be returned, *not* a partial string.
Friend Function ReadString_Unicode(ByVal numOfChars As Long) As String

    If m_Open And (Me.GetPosition() + numOfChars * 2 <= m_BufferSize) Then
        If (m_StreamMode = PD_SM_MemoryBacked) Or (m_StreamMode = PD_SM_ExternalPtrBacked) Then
            ReadString_Unicode = Strings.StringFromCharPtr(VarPtr(m_MemBuffer(m_Pointer)), True, numOfChars, True)
            m_Pointer = m_Pointer + numOfChars * 2
        ElseIf (m_StreamMode = PD_SM_FileBacked) Then
            If (Not m_FSO Is Nothing) Then
                Dim tmpBytes() As Byte
                ReDim tmpBytes(0 To numOfChars * 2 - 1) As Byte
                m_FSO.FileReadData m_FileHandle, VarPtr(tmpBytes(0)), numOfChars * 2
                ReadString_Unicode = Strings.StringFromCharPtr(VarPtr(tmpBytes(0)), True, numOfChars, True)
            End If
        ElseIf (m_StreamMode = PD_SM_FileMemoryMapped) Then
            If (Not m_FileMM Is Nothing) Then
                Dim srcPtr As Long, srcReadMax As Long
                If m_FileMM.EnsureMapAvailable(m_Pointer, numOfChars * 2, srcPtr, srcReadMax) Then
                    ReadString_Unicode = Strings.StringFromCharPtr(srcPtr + m_Pointer, True, numOfChars, True)
                    m_Pointer = m_Pointer + numOfChars * 2&
                Else
                    InternalStreamError "ReadString_Unicode failed; const pdFileMM for details."
                End If
            End If
        End If
    End If

End Function

'Read a series of two-byte values from the stream, and treat each 16-bit word as a big-endian UTF-16 code.
Friend Function ReadString_Unicode_BE(ByVal numOfChars As Long) As String

    If m_Open And (Me.GetPosition() + numOfChars * 2 <= m_BufferSize) Then
        
        'Retrieve the string normally
        Dim tmpString As String
        tmpString = Me.ReadString_Unicode(numOfChars)
        
        'If string is not null-length, manually reverse endianness of all chars
        If (LenB(tmpString) <> 0) Then
            
            ReadString_Unicode_BE = String$(Len(tmpString), 0)
            Dim i As Long, tmpInt As Integer
            For i = 0 To numOfChars - 1
                GetMem2 StrPtr(tmpString) + (i * 2), tmpInt
                PutMem2 StrPtr(ReadString_Unicode_BE) + (i * 2), ntohs(tmpInt)
            Next i
            
        End If
        
    End If

End Function

'Read a series of multi-byte values from the stream, and treat the chars as if they are string bytes with
' unknown encoding.  Heuristics will be used to attempt to distinguish between common encodings.
' (This function is useful when you've wrapped a pdStream around a text file of unknown encoding, and you
' want to parse its contents using the full suite of generic string functions.)
'
'Regardless of original encoding, a proper BSTR will be returned.  If the requested length is invalid
' (i.e. it extends beyond the end of the stream), a partial string will be returned.
Friend Function ReadString_UnknownEncoding(ByVal numOfBytes As Long, Optional ByVal forceWindowsLineEndings As Boolean = True) As String
    
    If m_Open Then
        
        Dim tmpBytes() As Byte, trueByteCount As Long
        trueByteCount = Me.ReadBytes(tmpBytes, numOfBytes, True)
        
        If (trueByteCount > 0) Then
            If (Not Strings.StringFromMysteryBytes(tmpBytes, ReadString_UnknownEncoding, forceWindowsLineEndings)) Then ReadString_UnknownEncoding = vbNullString
        Else
            ReadString_UnknownEncoding = vbNullString
        End If
        
    End If

End Function

'Read a series of UTF-8 bytes from the stream, and return them as a standard BSTR.
' If the requested length is invalid (i.e. it extends beyond the end of the stream), a partial string will be returned.
' Optionally, the lengthIsEmbedded parameter can be set to TRUE; if it is, a 4-byte UTF-8 length will first be read,
' using little-endian notation, followed by that number of UTF-8 bytes.  The numOfBytes parameter will be ignored
' by this use-case (e.g. just pass zero).  This makes it symmetrical with the WriteString_UTF8 function if its optional
' embedded length parameter is also set.
Friend Function ReadString_UTF8(ByVal numOfBytes As Long, Optional ByVal lengthIsEmbedded As Boolean = False) As String
    
    If m_Open Then
        
        If lengthIsEmbedded Then numOfBytes = Me.ReadLong()
        If (Me.GetPosition + numOfBytes > Me.GetStreamSize) Then numOfBytes = Me.GetStreamSize - Me.GetPosition
        
        If (numOfBytes > 0) Then
            ReadString_UTF8 = Strings.StringFromUTF8Ptr(Me.ReadBytes_PointerOnly(numOfBytes), numOfBytes)
        Else
            ReadString_UTF8 = vbNullString
        End If
        
    End If

End Function

'Reduce the buffer to its exact size.  This is *strongly* discouraged if you intend on writing more data to the stream,
' as extra allocations are guaranteed after a trim.
'
'Returns: exact size of stream if successful; zero otherwise.
Friend Function TrimStream() As Long
    
    If m_Open Then
    
        If (m_StreamAccess = PD_SA_ReadOnly) Then
            InternalStreamError "TrimStream() cannot operate on read-only streams."
            Exit Function
        End If
    
        If (m_StreamMode = PD_SM_MemoryBacked) Then
            ReDim Preserve m_MemBuffer(0 To m_BufferSize - 1) As Byte
            TrimStream = m_BufferSize
        ElseIf (m_StreamMode = PD_SM_FileBacked) Then
            If (Not m_FSO Is Nothing) Then
                m_FSO.FileSetLength m_FileHandle, m_BufferSize + m_FilePointerOffset
            End If
        ElseIf (m_StreamMode = PD_SM_FileMemoryMapped) Then
            InternalStreamError "Not implemented"
        End If
    
    End If
    
End Function

'Write data from any arbitrary byte array, with an optional length parameter.  If the length parameter is not supplied,
' the entire array will be written.  If the length parameter *is* supplied, THE LENGTH WILL BE CALCULATED FROM THE
' SOURCE ARRAY'S LOWER BOUND.  If you want to write from some other starting position in the array, use the
' WriteBytesFromPointer() function, below.
'
'DO NOT PASS AN UNINITIALIZED ARRAY TO THIS FUNCTION.
'
'RETURNS: new file pointer position if successful; 0 if unsuccessful.  (Failure typically only happens if we are
'         supplied bad data, or if this stream hasn't been opened yet.)
Friend Function WriteByteArray(ByRef srcBytes() As Byte, Optional ByVal dataLength As Long = -1) As Long
    
    Dim srcPointer As Long, lowerBound As Long
    lowerBound = LBound(srcBytes)
    srcPointer = VarPtr(srcBytes(lowerBound))
    
    'Figure out how many bytes we'll actually be writing
    If (dataLength < 0) Then dataLength = UBound(srcBytes) - lowerBound + 1
        
    'Wrap the "write from pointer" function, which takes care of the rest (including checking for an open stream)
    WriteByteArray = WriteBytesFromPointer(srcPointer, dataLength)
    
End Function

'Copy bytes from some arbitrary pointer into the stream.  Most write functions ultimately wrap this function.
' If you call this function directly, please make sure your pointer and length values are valid!
'
'RETURNS: new file pointer position if successful; 0 if unsuccessful.  (Failure typically only happens if we are
'         supplied bad data, or if this stream hasn't been opened yet.)
Friend Function WriteBytesFromPointer(ByVal dataPointer As Long, ByVal dataLength As Long) As Long
    
    If m_Open Then
        
        If (m_StreamAccess = PD_SA_ReadOnly) Then
            InternalStreamError "Write() functions cannot operate on read-only streams."
            Exit Function
        End If
    
        If (m_StreamMode = PD_SM_MemoryBacked) Then
        
            'If we don't have room for this write, increase the buffer now.
            If ((m_Pointer + dataLength) > UBound(m_MemBuffer)) Then EnsureBufferSpaceAvailable dataLength
            
            'Copy the data into place
            CopyMemoryStrict VarPtr(m_MemBuffer(m_Pointer)), dataPointer, dataLength
            
            'Increment the buffer pointer to reflect its new position post-write
            m_Pointer = m_Pointer + dataLength
            
            'Increase the calculated size of the buffer to match.  (Note that we first check to see if the pointer
            ' exceeds the buffer size; because the caller can seek to an arbitrary position before writing, they may
            ' not actually be increasing the buffer's size.)
            If (m_Pointer > m_BufferSize) Then m_BufferSize = m_Pointer
            
            WriteBytesFromPointer = m_Pointer
        
        ElseIf (m_StreamMode = PD_SM_FileBacked) Then
        
            If (Not m_FSO Is Nothing) Then
            
                'Write the data out to file
                If m_FSO.FileWriteData(m_FileHandle, dataPointer, dataLength) Then
                
                    'Retrieve the new file pointer (which is opaquely managed by WAPI)
                    WriteBytesFromPointer = m_FSO.FileGetCurrentPointer(m_FileHandle)
                    
                    'Manually increase the buffer size to match, but only if we've moved past the end of the file
                    Dim tmpSize As Long
                    tmpSize = WriteBytesFromPointer - m_FilePointerOffset
                    If (m_BufferSize < tmpSize) Then m_BufferSize = tmpSize
                
                'Write failed
                Else
                    InternalStreamError "WriteBytesFromPointer failed to write to target file; see pdFSO error for more details."
                End If
            
            Else
                WriteBytesFromPointer = 0
            End If
        
        ElseIf (m_StreamMode = PD_SM_FileMemoryMapped) Then
            If (Not m_FileMM Is Nothing) Then
                Dim dstPtr As Long, dstWriteMin As Long, dstWriteMax As Long
                If m_FileMM.EnsureMapAvailable(m_Pointer, dataLength, dstPtr, dstWriteMin, dstWriteMax) Then
                    CopyMemoryStrict dstPtr + m_Pointer, dataPointer, dataLength
                    m_Pointer = m_Pointer + dataLength
                    If (m_Pointer > m_BufferSize) Then m_BufferSize = m_Pointer
                    WriteBytesFromPointer = m_Pointer
                Else
                    InternalStreamError "WriteBytesFromPointer failed; const pdFileMM for details."
                End If
            Else
                InternalStreamError "WriteBytesFromPointer failed; no pdFileMM object!"
            End If
        End If
    
    Else
        WriteBytesFromPointer = 0
    End If
    
End Function

'Write a single byte to the stream.
Friend Function WriteByte(ByVal srcByte As Byte) As Boolean
    WriteByte = WriteBytesFromPointer(VarPtr(srcByte), 1)
End Function

'Write a single 64-bit float to the stream
Friend Function WriteDouble(ByVal srcDouble As Double) As Boolean
    WriteDouble = WriteBytesFromPointer(VarPtr(srcDouble), 8)
End Function

'Write a 16-16 fixed-point value to the stream.
Friend Function WriteFixed1616(ByVal srcFloat As Single) As Boolean
    Dim intPart As Long, floatPart As Long
    intPart = Int(srcFloat)
    floatPart = Int(CSng(srcFloat - Int(srcFloat)) * 65536!)
    WriteFixed1616 = Me.WriteInt(intPart And &HFFFF&)
    If WriteFixed1616 Then WriteFixed1616 = Me.WriteIntU(floatPart And &HFFFF&)
End Function

'Write a 16-16 fixed-point value (big-endian) to the stream.
Friend Function WriteFixed1616_BE(ByVal srcFloat As Single) As Boolean
    Dim intPart As Long, floatPart As Long
    intPart = Int(srcFloat)
    floatPart = Int(CSng(srcFloat - Int(srcFloat)) * 65536!)
    WriteFixed1616_BE = Me.WriteInt_BE(intPart And &HFFFF&)
    If WriteFixed1616_BE Then WriteFixed1616_BE = Me.WriteIntU_BE(floatPart And &HFFFF&)
End Function

'Write a single (pun?) 32-bit float to the stream
Friend Function WriteFloat(ByVal srcSingle As Single) As Boolean
    WriteFloat = WriteBytesFromPointer(VarPtr(srcSingle), 4)
End Function

'Write a single 32-bit float to the stream, *after* converting it to big-endian format
Friend Function WriteFloat_BE(ByVal srcSingle As Single) As Boolean
    Dim tmpLong As Long
    CopyMemoryStrict VarPtr(tmpLong), VarPtr(srcSingle), 4
    WriteFloat_BE = Me.WriteLong_BE(tmpLong)
End Function

'Write a single integer to the stream.
Friend Function WriteInt(ByVal srcInt As Integer) As Boolean
    WriteInt = WriteBytesFromPointer(VarPtr(srcInt), 2)
End Function

'Write a single integer to the stream, *after* converting it to big-endian format.
Friend Function WriteInt_BE(ByVal srcInt As Integer) As Boolean
    WriteInt_BE = WriteInt(htons(srcInt))
End Function

'Unsigned int writes; pass a Long, per PD convention.
Friend Function WriteIntU(ByVal srcInt As Long) As Boolean
    WriteIntU = WriteBytesFromPointer(VarPtr(srcInt), 2)
End Function

'Big-endian unsigned int writes
Friend Function WriteIntU_BE(ByVal srcInt As Long) As Boolean
    Dim tmpInteger As Integer
    CopyMemoryStrict VarPtr(tmpInteger), VarPtr(srcInt), 2
    tmpInteger = htons(tmpInteger)
    WriteIntU_BE = WriteBytesFromPointer(VarPtr(tmpInteger), 2)
End Function

'Write a single long to the stream.
Friend Function WriteLong(ByVal srcLong As Long) As Boolean
    WriteLong = WriteBytesFromPointer(VarPtr(srcLong), 4)
End Function

'Write a single long to the stream, *after* converting it to big-endian format.
Friend Function WriteLong_BE(ByVal srcLong As Long) As Boolean
    WriteLong_BE = WriteLong(htonl(srcLong))
End Function

'Write null padding bytes.  Padding values are guaranteed to be 0.  (If you don't care about
' padding values, note that it's faster to just call SetFilePointer.)
Friend Function WritePadding(ByVal padLength As Long) As Boolean
    Dim nullPadding() As Byte
    ReDim nullPadding(0 To padLength - 1) As Byte
    WritePadding = (Me.WriteBytesFromPointer(VarPtr(nullPadding(0)), padLength) <> 0)
End Function

'Write the contents of another stream.  By default, the *full* source stream is written; if writing less than this
' is desired, you will need to perform the write yourself.  The source stream's pointer is not read or modified.
'
'IMPORTANT NOTE: this function currently only works when the source stream is a MEMORY STREAM, as it uses a
' fast pointer-only mechanism for copying stream contents.
Friend Function WriteStream(ByRef srcStream As pdStream) As Boolean
    If (Not srcStream Is Nothing) Then
        WriteStream = (Me.WriteBytesFromPointer(srcStream.Peek_PointerOnly(0), srcStream.GetStreamSize()) <> 0)
    End If
End Function

'Given a source VB string, write each char to the stream as if it were a one-byte ASCII value.
Friend Function WriteString_ASCII(ByRef srcString As String) As Boolean
    If (LenB(srcString) <> 0) Then
        Dim tmpBytes() As Byte
        tmpBytes = StrConv(srcString, vbFromUnicode)
        WriteString_ASCII = Me.WriteBytesFromPointer(VarPtr(tmpBytes(0)), UBound(tmpBytes) + 1)
    End If
End Function

'Given a source VB string, write each char to the stream as a two-byte UTF-16 character.
Friend Function WriteString_Unicode(ByRef srcString As String, Optional ByVal addTrailingNull As Boolean = False) As Boolean
    WriteString_Unicode = Me.WriteBytesFromPointer(StrPtr(srcString), LenB(srcString))
    If addTrailingNull And WriteString_Unicode Then WriteString_Unicode = Me.WriteInt(0)
End Function

'Given a source VB string, write each char to the stream as a two-byte UTF-16 character,
' using big-endian storage.
Friend Function WriteString_UnicodeBE(ByRef srcString As String, Optional ByVal addTrailingNull As Boolean = False) As Boolean
    
    WriteString_UnicodeBE = True
    
    'There's no good way to do this, but since performance isn't a huge deal in PD's current use-cases,
    ' just write each char one at a time.
    Dim i As Long
    For i = 1 To Len(srcString)
        WriteString_UnicodeBE = WriteString_UnicodeBE And Me.WriteInt_BE(AscW(Mid$(srcString, i, 1)))
    Next i
    
    If addTrailingNull Then WriteString_UnicodeBE = WriteString_UnicodeBE And Me.WriteInt(0)
    
End Function

'Given a source VB string, transparently convert it to UTF-8 and embed *that* in the stream.  Optionally,
' you can choose to also embed the string's length; this will preface the UTF-8 data with a 4-byte size
' descriptor.  If you choose to embed length, you MUST also specify "readEmbeddedLength=TRUE" when reading
' the UTF-8 string - otherwise the size descriptor will show up in the UTF-8 data!
Friend Function WriteString_UTF8(ByRef srcString As String, Optional ByVal embedLengthToo As Boolean = False) As Boolean
    Dim lenUTF8 As Long, tmpUTF8() As Byte
    If Strings.UTF8FromStrPtr(StrPtr(srcString), Len(srcString), tmpUTF8, lenUTF8) Then
        If embedLengthToo Then Me.WriteLong lenUTF8
        WriteString_UTF8 = (Me.WriteBytesFromPointer(VarPtr(tmpUTF8(0)), lenUTF8) <> 0)
    Else
        InternalStreamError "WriteString_Unicode", "couldn't convert string to UTF-8"
    End If
End Function

'Return a pointer to the stream data at peekPosition.  If peekPosition is -1, the current stream pointer
' will be used.
'
'Unlike the various ReadXYZ() functions, this function does *not* modify the current stream pointer.
'
'NOTE: for memory-mapped file streams, you *must* supply a length alongside the peek position.
'      A length is required to page the backing data into memory.
'
'NOTE: for file-backed streams, this function *cannot be used* because there is no underlying pointer for
'      the destination.
'
'RETURNS: non-zero pointer if successful; zero otherwise.  Zero will be returned if the current stream
'         is file-backed.
'
'IMPORTANT NOTE: adding data to the buffer may force it to allocate new memory, invalidating all
'                previously returned pointers.  You must use the return of this function immediately,
'                as its correctness is not guaranteed after any other functions are called.
Friend Function Peek_PointerOnly(Optional ByVal peekPosition As Long = -1, Optional ByVal peekLength As Long = 0) As Long
    
    'Return 0 in case of failure
    Peek_PointerOnly = 0
    
    If (peekPosition < 0) Then peekPosition = Me.GetPosition()
    
    If (m_StreamMode = PD_SM_MemoryBacked) Or (m_StreamMode = PD_SM_ExternalPtrBacked) Then
        Peek_PointerOnly = VarPtr(m_MemBuffer(peekPosition))
        
    'Unsupported
    ElseIf (m_StreamMode = PD_SM_FileBacked) Then
        InternalStreamError "You can't peek pointers on a file-based stream!"
        
    ElseIf (m_StreamMode = PD_SM_FileMemoryMapped) Then
        If (Not m_FileMM Is Nothing) And (peekLength > 0) Then
            Dim srcPtr As Long
            If m_FileMM.EnsureMapAvailable(peekPosition, peekLength, srcPtr) Then
                Peek_PointerOnly = srcPtr + peekPosition
            Else
                Peek_PointerOnly = 0
                InternalStreamError "Peek_PointerOnly failed; const pdFileMM for details."
            End If
        Else
            Peek_PointerOnly = 0
            InternalStreamError "Peek_PointerOnly failed - you need to supply a length value!"
        End If
    End If

End Function

'Ensure that (n) bytes are available for a new write.  In PD, we do a lot of writing by raw pointer(s), and this function
' lets us ensure that we have enough space available for unprotected copies.
'
'If (n) bytes are not available, the buffer will automatically be resized to ensure at least (n) bytes are available.
' (Note that allocations are likely to be larger than the requested amount, for performance reasons.)
'
'This function is currently designed against PhotoDemon's unique needs, where stream objects tend to wrap very large
' targets (like image files).  If you only use this class for very small streams, you may want to change the allocation
' strategy to better support tiny objects.  (Also, you'll want to look at the .StartStream function - by default,
' memory streams are allocated as 4k to start.  That may be too big for something like a string builder.)
'
'Returns: TRUE if the buffer was successfully resized to ensure (n) bytes are available; FALSE otherwise.  If this stream
'         is not open, returns FALSE.  If this stream was opened in file-backed mode, this function does nothing.
'
'For an explanation of the allocation strategy used, please see:
' - http://stackoverflow.com/questions/1100311/what-is-the-ideal-growth-rate-for-a-dynamically-allocated-array
' - https://github.com/facebook/folly/blob/master/folly/docs/FBVector.md
' - https://blog.mozilla.org/nnethercote/2014/11/04/please-grow-your-buffers-exponentially/
Friend Function EnsureBufferSpaceAvailable(ByVal newDataSizeInBytes As Long, Optional ByVal exactAllocationWanted As Boolean = False) As Boolean
    
    If m_Open Then
        
        Dim needAtLeastBytes As Long, bufferLimit As Long, newSize As Long
        
        If (m_StreamMode = PD_SM_MemoryBacked) Then
            
            'Calculate the buffer size required for the new data
            needAtLeastBytes = m_Pointer + newDataSizeInBytes
            bufferLimit = UBound(m_MemBuffer)
            
            If (needAtLeastBytes > bufferLimit) Then
                
                'Calculate a new buffer limit.  This calculation varies depending on the size of current buffer,
                ' the size of the requested data, and whether the user wants an exact trim or not.
                
                'Callers can request an exact buffer size.  You should *only* use this for things like filling the
                ' entire buffer at once from a single known source.  If you plan on appending data on the future,
                ' let the stream object use its default allocation strategy instead.
                If exactAllocationWanted Then
                    newSize = newDataSizeInBytes
                Else
                
                    Const ONE_MEBIBYTE = 1048576
                    
                    'If the current buffer is small (< 1 MiB), and we're writing less than 1 MiB, allow for more
                    ' granular allocations.
                    If (bufferLimit < ONE_MEBIBYTE) And (newDataSizeInBytes < ONE_MEBIBYTE) Then
                    
                        'If the write size is larger than 4k, increase size by 1.5x and lock to the nearest 256k boundary
                        If (newDataSizeInBytes > 4096) Then
                        
                            Const BYTES_256K As Long = 262144
                            newSize = needAtLeastBytes * 1.5 + BYTES_256K
                            
                            'Lock to 256k boundaries
                            newSize = Int(newSize \ BYTES_256K) * BYTES_256K
                        
                        'If the write size is smaller than 4k, use a basic doubling strategy
                        Else
                        
                            newSize = needAtLeastBytes * 2 + 4096
                            
                            'Lock to 4K boundaries
                            newSize = Int(newSize \ 4096) * 4096
                        
                        End If
                    
                    'Buffers or writes over 1 MiB use a more aggressive growth strategy, as allocations in this range are
                    ' quickly become very expensive.
                    Else
                        
                        'Use a 1.5x growth strategy, and because we'll be using integer division in the next step,
                        ' ensure that we round UP to the nearest mebibyte boundary.
                        newSize = CDbl(needAtLeastBytes) * 1.5 + ONE_MEBIBYTE
                            
                        'Lock to mebibyte boundaries, for convenience
                        newSize = Int(newSize \ ONE_MEBIBYTE) * ONE_MEBIBYTE
                        
                    End If
                    
                End If
                
                'ReDim the buffer to match the new size
                ReDim Preserve m_MemBuffer(0 To newSize - 1) As Byte
                
            End If
        
        ElseIf (m_StreamMode = PD_SM_FileBacked) Then
        
            'TODO?  I'm not really sure there's anything *to* do here, honestly, besides maybe check for available file space
            ' (but even that is unreliable, as Windows may not be able to write a file that consumes all that space).
        
        ElseIf (m_StreamMode = PD_SM_ExternalPtrBacked) Then
            InternalStreamError "EnsureBufferSpaceAvailable is not allowed on pointer-wrapped streams!"
            EnsureBufferSpaceAvailable = False
        
        ElseIf (m_StreamMode = PD_SM_FileMemoryMapped) Then
            Dim tmpPtrMap As Long
            If (Not m_FileMM Is Nothing) Then m_FileMM.EnsureMapAvailable m_Pointer, newDataSizeInBytes, tmpPtrMap
        End If
        
        EnsureBufferSpaceAvailable = True
    
    End If
    
End Function

Private Sub InternalStreamError(ByVal errMessage As String, Optional ByVal errNumber As Long = 0)
    If (errNumber <> 0) Then
        PDDebug.LogAction "WARNING!  pdStream error # " & errNumber & ": " & errMessage
    Else
        PDDebug.LogAction "WARNING!  Unspecified pdStream error: " & errMessage
    End If
End Sub

Private Sub Class_Terminate()
    If m_Open Then Me.StopStream
End Sub
