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
'Copyright 2016-2016 by Tanner Helland
'Created: 05/December/16
'Last updated: 05/December/16
'Last update: heavily streamline all memory methods for idealized performance
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
' - A framework is now place for backing a stream by a file or a memory buffer.  (TODO; only memory buffers are
'   implemented right now.)
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Declare Sub CopyMemory_Strict Lib "kernel32" Alias "RtlMoveMemory" (ByVal ptrToDestination As Long, ByVal ptrToSource As Long, ByVal memLength As Long)

Public Enum PD_STREAM_MODE
    PD_SM_MemoryBacked = 0
    PD_SM_FileBacked = 1
End Enum

#If False Then
    Private Const PD_SM_MemoryBacked = 0, PD_SM_FileBacked = 1
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

'When writing a stream, this is the current stream size.  It is updated on each write.  (Importantly, for memory streams,
' this is *not* the UBound() of m_MemBuffer()!  m_MemBuffer's UBound() will always be >= this number.)
Private m_BufferSize As Long

'Current position inside the stream.  Starts at 0 and increases from there.  Cannot be negative.  Can be modified via Seek().
Private m_Pointer As Long

'Once a stream has been opened (by a call to the StartStream function), this will be set to TRUE.  This value is important
' for ensuring that all internals have been properly prepped prior to issuing read/write requests.
Private m_Open As Boolean

'Return the current stream pointer position.
Friend Function GetPosition() As Long
    GetPosition = m_Pointer
End Function

'Set a new stream pointer position.
'
'RETURNS: TRUE if successful; FALSE otherwise.  FALSE may occur if you attempt to set the pointer outside the current stream range.
Friend Function SetPosition(ByVal newPosition As Long) As Boolean
    If (newPosition >= 0) And (newPosition <= m_BufferSize) Then
        
        If (m_StreamMode = PD_SM_MemoryBacked) Then
            m_Pointer = newPosition
            SetPosition = True
        ElseIf (m_StreamMode = PD_SM_FileBacked) Then
        
        End If
        
    Else
        SetPosition = False
    End If
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
' class's helper functions.)  To ensure that the buffer's reported size remains correct, we call this function afterward.
Friend Sub SetSizeExternally(ByVal newSize As Long)
    m_BufferSize = newSize
End Sub

'Start a new stream.  The stream type *must* be specified.  If this is a file stream, a valid filename *must* be provided.
' Read/write access does not need to be specified, but if you only need a read-only buffer, specifying as much will yield
' improved performance, particularly for file streams.
'
'Optionally, when creating a writable buffer, you can also specify a starting buffer size.  If you have some notion of the
' stream's net size in advance, an initial memory allocation that comes close to your required value can spare us intermediate
' allocations while the stream is constructed.
'
'Returns TRUE if successful; FALSE otherwise.  FALSE should only occur if...
' 1) This is a file-backed stream and you supplied an invalid or inaccessible filename, or...
' 2) This is a memory-backed stream and the initial starting size is ridiculously huge.
'
'If this stream object is already open, the existing stream will be closed and a new one will be started.
Friend Function StartStream(ByVal streamType As PD_STREAM_MODE, Optional ByVal streamAccess As PD_STREAM_ACCESS = PD_SA_ReadWrite, Optional ByVal srcFilename As String, Optional ByVal startingBufferSize As Long = 0) As Boolean

    'If we're already open, close the current stream before starting a new one.
    If m_Open Then Me.StopStream
    
    m_StreamMode = streamType
    m_StreamAccess = streamAccess
    
    If (streamType = PD_SM_MemoryBacked) Then
    
        'The caller can specify a starting buffer size.  If they don't, we'll start small - just 4k.
        If (startingBufferSize <= 0) Then
            
            Const INITIAL_BUFFER_SIZE As Long = 4096
            ReDim m_MemBuffer(0 To INITIAL_BUFFER_SIZE - 1) As Byte
            
        'If the caller has some knowledge of the stream's final size, we can use their hint instead.
        Else
            ReDim m_MemBuffer(0 To startingBufferSize - 1) As Byte
        End If
        
        m_Open = True
        
    ElseIf (streamType = PD_SM_FileBacked) Then
    
        'TODO: validate the filename, create a handle, etc
    
    End If
    
    StartStream = m_Open
    
End Function

'Shut down the current stream.  For file-backed streams, this closes the current file handle.  For memory streams, this just
' resets our internal buffers to null values.
Friend Sub StopStream()
    
    m_Open = False
    m_Pointer = 0
    m_BufferSize = 0
    
    If (m_StreamMode = PD_SM_MemoryBacked) Then
        Erase m_MemBuffer
    ElseIf (m_StreamMode = PD_SM_FileBacked) Then
    
    End If
    
End Sub

'Return a copy of the stream, starting from the current stream position (which you can/should move prior to reading,
' as relevant) and extending the supplied readLength amount of bytes.  If readLength is not supplied, a full copy of
' the entire stream from the current position to the end will be returned.
'
'RETURNS: number of bytes read if successful; zero otherwise.
'
'NOTE: if the optional trimDestinationArray parameter is specified, the destination array will be exactly sized to
'      match the outgoing data.  If trimDestinationArray is FALSE, the destination array will only be resized if
'      necessary to fit the data.  This can be very helpful for performance, as you can reuse the same destination
'      array for reading data from the stream, saving expensive allocations.
Friend Function ReadBytes(ByRef dstBytes() As Byte, Optional ByVal readLength As Long = -1, Optional ByVal trimDestinationArray As Boolean = True) As Long
    
    If m_Open Then
        
        'Calculate the length we need to move the current position marker
        If (readLength < 0) Then
            readLength = m_BufferSize - m_Pointer
        
        'If the caller supplied a length value, make sure it doesn't extend past the end of the buffer
        Else
            If (m_Pointer + readLength > m_BufferSize) Then readLength = m_BufferSize - m_Pointer
        End If
        
        Dim arrLBound As Long, arrUBound As Long
        
        'Check to see if our destination array is already initialized to the perfect size.  If it is, we don't need to
        ' allocate new memory.
        If VB_Hacks.IsArrayInitialized(dstBytes) Then
            
            arrLBound = LBound(dstBytes)
            arrUBound = UBound(dstBytes)
            
            'If the user wants the destination array trimmed to exact size, see if we need to modify the destination
            ' array's size to fit.
            If trimDestinationArray Then
                If ((arrUBound - arrLBound + 1) <> readLength) Then ReDim dstBytes(arrLBound To arrLBound + readLength - 1) As Byte
            
            'If the user doesn't want the destination array trimmed to exact size, we only need to resize it if it's
            ' smaller than the amount of data being read.
            Else
                If ((arrUBound - arrLBound + 1) < readLength) Then ReDim dstBytes(arrLBound To arrLBound + readLength - 1) As Byte
            End If
        
        'If the destination array is *not* initialized, make it the perfect size regardless of trimDestinationArray.
        Else
            arrLBound = 0
            ReDim dstBytes(0 To readLength - 1) As Byte
        End If
        
        'Copy the relevant amount of bytes into position
        CopyMemory_Strict VarPtr(dstBytes(arrLBound)), VarPtr(m_MemBuffer(m_Pointer)), readLength
        
        'Move the pointer
        m_Pointer = m_Pointer + readLength
        
        'As a failsafe, check for invalid pointer positions
        If (m_Pointer > m_BufferSize) Then m_Pointer = m_BufferSize
        
        ReadBytes = readLength
        
    End If
    
End Function

'Read bytes from the stream.  Instead of returning the bytes (which requires an expensive copy operation), this function
' simply returns a pointer to the relevant position in the stream, which the user can use however they want.  An optional
' readLength parameter controls how far the position pointer is moved within the stream, but it does not otherwise
' affect anything.
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
    
        'Return a pointer to the current buffer position
        ReadBytes_PointerOnly = VarPtr(m_MemBuffer(m_Pointer))
        
        'Calculate the length we need to move the current position marker
        If (readLength < 0) Then
            readLength = m_BufferSize - m_Pointer
        End If
        
        'Move the pointer
        m_Pointer = m_Pointer + readLength
        
        'As a failsafe, check for invalid pointer positions
        If (m_Pointer > m_BufferSize) Then m_Pointer = m_BufferSize
        
    End If
    
End Function

'Reduce the buffer to its exact size.  This is *strongly* discouraged if you intend on writing more data to the stream,
' as extra allocations are guaranteed after a trim.
'
'Returns: exact size of stream if successful; zero otherwise.
Friend Function TrimStream() As Long
    If m_Open Then
        If (m_StreamMode = PD_SM_MemoryBacked) Then
            ReDim Preserve m_MemBuffer(0 To m_BufferSize - 1) As Byte
            TrimStream = m_BufferSize
        ElseIf (m_StreamMode = PD_SM_FileBacked) Then
        
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
    If (dataLength < 0) Then
        dataLength = UBound(srcBytes) - lowerBound + 1
    End If
        
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
    
        If (m_StreamMode = PD_SM_MemoryBacked) Then
        
            'If we don't have room for this write, increase the buffer now.
            If ((m_Pointer + dataLength) > UBound(m_MemBuffer)) Then EnsureBufferSpaceAvailable dataLength
            
            'Copy the data into place
            CopyMemory_Strict VarPtr(m_MemBuffer(m_Pointer)), dataPointer, dataLength
    
            'Increment the buffer pointer to reflect its new position post-write
            m_Pointer = m_Pointer + dataLength
            
            'Increase the calculated size of the buffer to match.  (Note that we first check to see if the pointer
            ' exceeds the buffer size; because the caller can seek to an arbitrary position before writing, they may
            ' not actually be increasing the buffer's size.)
            If (m_Pointer > m_BufferSize) Then m_BufferSize = m_Pointer
            
            WriteBytesFromPointer = m_Pointer
        
        ElseIf (m_StreamMode = PD_SM_FileBacked) Then
        
        End If
    
    Else
        WriteBytesFromPointer = 0
    End If
    
End Function

'Return a pointer to the stream data at peekPosition.  If peekPosition is -1, the current stream pointer will be used.
' Unlike the various ReadXYZ() functions, this function does *not* modify the current stream pointer.
'
'NOTE: for file-backed streams, *if* the stream was opened in read-only mode, this function is still valid.  If the file
' stream is opened in read/write mode, this function *cannot be used*.
'
'RETURNS: non-zero pointer if successful; zero otherwise.  Zero will be returned if the current stream is file-backed,
' and the stream was opened with read/write access.  (Memory-mapped access is currently only supported for read-only access,
' for performance reasons.)
'
'IMPORTANT NOTE: adding data to the buffer may force it to allocate new memory, invalidating all previously returned
' pointers.  You must use the return of this function immediately, as its correctness is not guaranteed after any other
' class functions are called.
Friend Function Peek_PointerOnly(Optional ByVal peekPosition As Long = -1) As Long

    If (peekPosition < 0) Then peekPosition = m_Pointer
    
    If (m_StreamMode = PD_SM_MemoryBacked) Then
        Peek_PointerOnly = VarPtr(m_MemBuffer(peekPosition))
        
    ElseIf (m_StreamMode = PD_SM_FileBacked) Then
    
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
' strategy to better support tiny objects.  (Also, you'll want to look at the .StartStream function - by default, streams
' are allocated as 4k to start.  That may be too big for something like a string builder.)
'
'Returns: TRUE if the buffer was successfully resized to ensure (n) bytes are available; FALSE otherwise.  If this stream
'         was opened in read-only mode, this function will always return FALSE.  If this stream is not open, returns FALSE.
'
'For an explanation of the allocation strategy used, please see:
' - http://stackoverflow.com/questions/1100311/what-is-the-ideal-growth-rate-for-a-dynamically-allocated-array
' - https://github.com/facebook/folly/blob/master/folly/docs/FBVector.md
' - https://blog.mozilla.org/nnethercote/2014/11/04/please-grow-your-buffers-exponentially/
Friend Function EnsureBufferSpaceAvailable(ByVal newDataSizeInBytes As Long, Optional ByVal exactAllocationWanted As Boolean = False) As Boolean
    
    If (m_Open And (m_StreamAccess = PD_SA_ReadWrite)) Then
    
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
        
        End If
        
        EnsureBufferSpaceAvailable = True
    
    End If
    
End Function

