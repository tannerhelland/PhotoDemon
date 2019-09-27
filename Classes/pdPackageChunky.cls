VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdPackageChunky"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon "pdPackageChunky" v2.0 Interface (e.g. chunk-based archive handler)
'Copyright 2014-2019 by Tanner Helland
'Created: 03/June/19
'Last updated: 27/September/19
'Last update: finalize 2.0 API
'Dependencies: - pdPackaging module from photodemon.org (defines public enums and structs)
'              - pdStream class from photodemon.org (used to easily read/write memory and file buffers)
'              - pdFSO class from photodemon.org (Unicode-aware file operations)
'              - pdStringStack class from photodemon.org (optimized handling of large string collections)
'              - VBHacks module from photodemon.org (optimized workarounds for lacking VB functionality)
'              - If you want compression support, one or more 3rd-party compression libraries
'                (e.g. zstd, zlib) are also required.
'
'This class provides an interface for creating and reading chunk-based pdPackage files.  pdPackage is
' the file format originally developed for PhotoDemon (www.photodemon.org), and as the program
' developed increasingly complicated features, the need for a more flexible file format came with it.
'
'Enter this new chunk-based pdPackage system.  This class reads and writes files that use a simple
' chunk-based storage system, where each "item" in the file is stored as a single contiguous "chunk".
' (This design is similar to the PNG file format, among others.)
'
'Chunk-based systems carry many benefits:
' 1) Decoders can skip over any chunks they can't/don't support, making forward- and
'    backward-compatibility very straightforward.
' 2) Changes to one chunk do not affect neighboring chunks; nothing is based on relative offsets
'    (unlike JPEG/TIFF/etc)
' 3) No need for a central directory (like zip files)
' 4) Individual chunks can be individually compressed using custom per-chunk settings without
'    affecting neighboring chunks.  Compressed and uncompressed chunks can be freely intermixed.
'
'The file format itself is very simple.  A few universal notes:
' - All numeric data types are little-endian.
' - Individual chunks must be 2 GB or less in size.  To store larger amounts of data, spread it across
'   multiple chunks.
' - Unless otherwise noted, all string data is in UTF-8 format WITHOUT BOM.  (The only strings
'   documented here actually use ASCII-safe chars, but the expectation is that user-defined chunks
'   can use any valid UTF-8 string.)
'
'Files begin with a hard-coded 8-byte "magic string", identifying the file as a chunky pdPackage:
' "PDPKCHNK" (8-byte UTF-8 string)
'
'This is followed by a 4-byte unsigned integer describing the encoder software version.  This value can
' be used to identify files that may contain new, possibly unsupported features.  At the time of this
' writing, only a version value of "2" (0x00000002) is defined.
'
'The remainder of the file is comprised of one or more chunks.  (Zero chunks is not allowed; there must
' be at least one chunk in the file.)
'
'Chunks follow a predefined format:
' - 4 byte UTF-8 string defining the chunk's ID.  (This string is case-sensitive, and it is suggested
'   - but not required - that the string use uppercase characters.)
' - 4 byte signed integer defining the total size of this chunk.  For compatibility purposes, this
'   value must be >= 0, which limits its max size to 2 GB.  Importantly, note that this size value
'   does *not* include the 8 bytes required for the 4-byte ASCII chunk identifier and the 4 bytes
'   required for this chunk size value; thus an empty chunk (consisting of just a 4-byte ID and a
'   4-byte chunk size) would have a stated size of 0.
'
'It a chunk size is *not* zero, it MUST be >= 12.  This is because a non-zero-length chunk must
' - immediately following the chunk size described above - provide 12 bytes of data further describing
' the contents of the chunk.  (These 12 bytes are *not* used in a zero-length chunk, as that would be
' wasteful and silly).
'
'If a chunk's size is non-zero, the next 12 bytes are defined as follows.  (These 12 bytes are
' mandatory, meaning they must ALWAYS be present for a chunk of non-zero size.)
' - 4 byte UTF-8 string defining the chunk's compression format, if any.  This string is case-sensitive.
'   Currently defined compression formats include:
'   - "none" (no compression)
'   - "zstd" (zstandard)
'   - "defl" (deflate)
'   - "zlib" (zlib)
'   PhotoDemon uses some other compression types internally (e.g. "lz4 " for the lz4 compression
'   library, which PhotoDemon uses for real-time compression of large binary data), but such
'   compression types will never be used in externally accessible files.
' - 4 byte signed integer defining the *uncompressed* size of this chunk's data.  This value must be
'   greater than 0.
' - 4 byte signed integer defining the *compressed* size of the chunk's data.  This value must be
'   greater than 0, and less than or equal to "chunk size - 12 bytes".  Also, if the compressed size
'   would somehow be LARGER than the uncompressed size, please just write an uncompressed chunk;
'   this spec does not formally require this, but you are a bad developer if you do otherwise.
'
'Note that this last value - a 4-byte signed integer defining the compressed size of this chunk's data -
' could theoretically be inferred from the chunk's size, e.g. "chunk size - 8 bytes (for compression
' type and uncompressed size)".  However, it is the observation of this author that many data formats
' struggle to gracefully handle padding and/or arbitrary data alignment.  To provide flexibility, this
' spec requires that the compressed size always be stated explicitly.  This allows chunks to be
' arbitrarily padded to specific alignments as convenient for a specific use-case (e.g. perhaps to
' achieve 32-bit or 64-bit alignment).  This spec does NOT make any formal requirements regarding
' padding, and compatible decoders must handle any arbitrary alignment (or lack thereof) between chunks,
' even if inconvenient for their platform.
'
'Note that even if a chunk is NOT compressed (e.g. compression type = "none"), it MUST still supply
' valid uncompressed and compressed sizes.  These sizes will simply be equal to each other.  These
' sizes do NOT however need to align with the chunk size itself, which means uncompressed chunks can
' still be arbitrarily padded at your convenience.
'
'IMPORTANTLY: the offset between adjacent chunks MUST be calculated using the first chunk size value
' (the 4-byte signed integer that immediately follows the 4-byte chunk ID).  Any bytes in a chunk that
' lie beyond the stated *compressed* size (regardless of compression type) but before the next chunk
' in the file should be considered "undefined" and MUST be skipped during parsing.  (The author
' suggests that such bytes always be set to 0 by an encoder, but this not a requirement.)
'
'Chunks in the file must be iterated until an "PEND" (package end) chunk is reached.  This chunk always
' has a chunk size of 0.  No chunks may appear after it.  (In fact, any data whatsoever past the "PEND"
' chunk MUST be ignored.)
'
'Besides the "PEND" chunk, chunk names and definitions are left up to individual programs.
'
'Final thoughts:
'
'No other formal limitations are placed upon this format.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Current package manager version
Private Const PACKAGE_VERSION As Long = &H2&

'Hard-coded chunky pdPackage file identifier
Private Const CHUNKY_PDPACKAGE_ID As String = "PDPKCHNK"

'The currently assembled chunk data.  This may wrap a file or memory instance; both must be supported.
Private m_Stream As pdStream

'When writing chunks in pieces (whether compressed or otherwise), we need temporary holders for
' chunk data.  Persistent temporary streams are used, and by default, their memory is *not* deallocated
' between chunks.  This prevents cache thrashing when writing a bunch of similarly-sized chunks in a row
' (with or without compression).  Two streams are used because one stores the uncompressed data as
' received from the caller, while the second stream stores the compressed data stream.
Private m_TmpChunkSrc As pdStream, m_TmpChunkDst As pdStream

'When writing a chunk in pieces, we need to remember the chunk ID (because it's specified in the
' StartChunk() function, but not actually written until the EndChunk() function)
Private m_TmpChunkID As String

'When opening an existing package, this class automatically "peeks" at the next chunk in line.
' This means that the current stream pointer *ALWAYS* points at the data segment of the current chunk.
' The current chunk's metadata should be retrieved from m_PeekData, as required. (This provides
' meaningful perf improvements as we can read the file fully sequentially, without ever needing to
' back-up the stream pointer to maintain chunk alignment.)
Private Type PackageChunkPeek
    peekOffset As Long
    peekName As String
    peekSize As Long
    
    'The following members are only used if the chunk has a non-zero size
    peekCompType As PD_CompressionFormat
    peekCompressedSize As Long
    peekUncompressedSize As Long
    
End Type

Private m_PeekData As PackageChunkPeek

'Add a full chunk to the current package in one fell swoop.  This is the fastest way to write a chunk,
' but note that it requires the data to already exist externally in a contiguous, ready-to-go state.
Friend Function AddChunk_WholeFromPtr(ByVal chunkID As String, ByVal ptrChunkData As Long, ByVal chunkDataLength As Long, Optional ByVal cmpFormat As PD_CompressionFormat = cf_None, Optional ByVal cmpLevel As Long = -1) As Boolean
    
    If (Not m_Stream.IsOpen()) Then
        InternalError "AddChunk_WholeFromPtr", "you haven't called StartNewPackage yet!"
        Exit Function
    End If
    
    'Always start by validating the nodeID - it must be a 4-byte ASCII string.
    If (Len(chunkID) <> 4) Then chunkID = PDPackaging.ValidateChunkID(chunkID)
    
    'The rest of the write process is handled by a dedicated function
    AddChunk_WholeFromPtr = WriteChunkFromPtr_Safe(chunkID, ptrChunkData, chunkDataLength, cmpFormat, cmpLevel)
    
End Function

'PhotoDemon often needs to store multi-part data pieces - e.g. an image layer's XML header, followed by
' one or more binary blobs with actual layer pixel data.  This convenience-only function accepts two sets
' of data, and it will automatically create (n) valid chunks in the package, in the order the source headers
' are supplied.  The source array MUST have an LBound of 0; its UBound doesn't matter, but if you don't want
' the full array of structs written, you MUST pass the number of chunks to use (1-based).
Friend Function AddChunks(ByRef srcChunks() As PDPackage_ChunkData, Optional ByVal numOfChunks As Long = -1) As Boolean
    
    AddChunks = True
    If (numOfChunks <= 0) Then numOfChunks = UBound(srcChunks) + 1
    
    Dim tmpBytes() As Byte, tmpBytesSize As Long
    
    Dim i As Long
    For i = 0 To numOfChunks - 1
        
        With srcChunks(i)
            
            'Convert the source data into a standard byte stream, as required.
            
            'If the source data is a string, convert it from UTF-16 to UTF-8.
            If (.chunkDataFormat = dt_StringToUTF8) Then
                AddChunks = AddChunks And Strings.UTF8FromStrPtr(.ptrChunkData, .chunkDataLength \ 2, tmpBytes, tmpBytesSize)
                If AddChunks Then
                    .ptrChunkData = VarPtr(tmpBytes(0))
                    .chunkDataLength = tmpBytesSize
                End If
            End If
            
            'Add the chunk to the collection
            If AddChunks Then AddChunks = AddChunks And WriteChunkFromPtr_Safe(.chunkID, .ptrChunkData, .chunkDataLength, .chunkCompressionFormat, .chunkCompressionLevel)
            
        End With
        
    Next i
    
End Function

'If this function returns TRUE, there are more chunks in the file.  Only valid and relevant during
' package reading; it always returns FALSE during package writing.
Friend Function ChunksRemain() As Boolean
    ChunksRemain = (m_PeekData.peekName <> "PEND")
End Function

'After calling StartChunk() and GetInProgressChunk() one or more times, call EndChunk() to commit the
' finished chunk to file/memory.
Friend Function EndChunk(Optional ByVal cmpFormat As PD_CompressionFormat = cf_None, Optional ByVal cmpLevel As Long = -1) As Boolean
    
    If m_Stream.IsOpen() Then
        
        'Always start by validating the nodeID - it must be a 4-byte ASCII string.
        Dim chunkID As String
        chunkID = m_TmpChunkID
        If (Len(chunkID) <> 4) Then chunkID = PDPackaging.ValidateChunkID(chunkID)
        
        'The actual write is handled by a dedicated function
        EndChunk = WriteChunkFromPtr_Safe(chunkID, m_TmpChunkSrc.Peek_PointerOnly(0), m_TmpChunkSrc.GetStreamSize(), cmpFormat, cmpLevel)
        
    Else
        InternalError "EndChunk", "you haven't called StartNewPackage yet!"
        Exit Function
    End If
    
End Function

'Finish the current package.  If the current package is being written to file, the file will be closed.
' If the current package is being written to memory, you *must* pass your own pdStream object as the
' copyOfFinalStream parameters; your object will be set to this class's internal buffer, and you can
' then do whatever you please with it.
Friend Function FinishPackage(Optional ByRef copyOfFinalStream As pdStream = Nothing) As Boolean
    
    FinishPackage = True
    
    'Write a final "PEND" (package end) marker
    FinishPackage = FinishPackage And m_Stream.WriteString_ASCII("PEND")
    FinishPackage = FinishPackage And m_Stream.WriteLong(0)
    
    'For file-based streams, close the handle now
    If (m_Stream.GetStreamMode <> PD_SM_MemoryBacked) Then
        m_Stream.StopStream False
        
    'For memory-based streams, give the caller a copy of the data
    Else
        Set copyOfFinalStream = m_Stream
    End If

End Function

'When creating a chunk in pieces, you must call StartChunk first.  Then, call this function as many
' times as you need, using any available stream functions to write your data.  When the chunk's data
' is finished, call EndChunk() to commit the finished chunk (with any compression, encryption, etc).
Friend Function GetInProgressChunk() As pdStream
    If m_Stream.IsOpen() Then
        Set GetInProgressChunk = m_TmpChunkSrc
    Else
        InternalError "GetInProgressChunk", "you haven't called StartNewPackage yet!"
        Exit Function
    End If
End Function

'Return the name of the next chunk in line.  Note that this does *not* permanently move the
' stream pointer (by design)!
Friend Function GetChunkName() As String
    GetChunkName = m_PeekData.peekName
End Function

'Return the size of the next chunk in line.  Note that this does *not* permanently move the
' stream pointer (by design)!
Friend Function GetChunkSize() As Long
    GetChunkSize = m_PeekData.peekSize
End Function

Friend Function GetPackageSize() As Long
    GetPackageSize = m_Stream.GetStreamSize()
End Function

'Retrieve the next chunk in the file.
' RETURNS: boolean TRUE if another chunk exists, FALSE if no more chunks exist.  (Note that the package end chunk
' "PEND" results in a FALSE return, by design, even though dstChunkName/Size will still be set and returned correctly.)
' Note: the destination stream may not be the same size reported by dstChunkSize.  This is especially true if you're
' using the same pdStream object multiple times in a row, as this class will attempt to reuse previous memory
' allocations whenever possible.  You *must* use dstChunkSize to handle the data correctly.
'
'If the chunk is compressed, this function will automatically decompress it for you.
Friend Function GetNextChunk(ByRef dstChunkName As String, ByRef dstChunkSize As Long, Optional ByRef dstChunkStream As pdStream, Optional ByVal loadToThisPtrInstead As Long = 0) As Boolean
    
    GetNextChunk = GetNextChunk_Helper(dstChunkName, dstChunkSize)
    If GetNextChunk Then
        
        'Make sure there is actually data to retrieve
        If (dstChunkSize <= 0) Then Exit Function
            
        'Prep the receiving stream
        If (loadToThisPtrInstead = 0) Then
            If (dstChunkStream Is Nothing) Then Set dstChunkStream = New pdStream
            If dstChunkStream.IsOpen Then dstChunkStream.SetPosition 0, FILE_BEGIN Else dstChunkStream.StartStream PD_SM_MemoryBacked, PD_SA_ReadWrite
        End If
        
        'Compressed chunk (must be decompressed before returning it)
        If (m_PeekData.peekCompType <> cf_None) Then
            
            'Overwrite the returned size with the *uncompressed* size of the compressed chunk
            dstChunkSize = m_PeekData.peekUncompressedSize
            
            'Next is the compression format
            Dim cmpFormat As PD_CompressionFormat
            cmpFormat = m_PeekData.peekCompType
            
            'Prep the destination stream
            If (loadToThisPtrInstead = 0) Then
                dstChunkStream.EnsureBufferSpaceAvailable dstChunkSize
                dstChunkStream.SetSizeExternally dstChunkSize
            End If
            
            'Retrieve the compressed bytes; if they're sitting in a file, we have to read them into memory
            ' before decompressing, but if this is a memory-backed or memory-mapped-file-backed stream,
            ' we can simply pull the relevant pointer.
            Dim ptrSrcComp As Long
            
            If (m_Stream.GetStreamMode = PD_SM_FileBacked) Then
                ReadyTmpChunkSrc
                m_Stream.ReadBytesToStream m_TmpChunkSrc, m_PeekData.peekCompressedSize
                ptrSrcComp = m_TmpChunkSrc.Peek_PointerOnly(0)
            Else
                ptrSrcComp = m_Stream.Peek_PointerOnly(, m_PeekData.peekCompressedSize)
            End If
            
            'Perform decompression.  If a destination pointer was supplied, presume the caller knows what
            ' they're doing and they have already prepared sufficient space for us.
            If (loadToThisPtrInstead <> 0) Then
                GetNextChunk = Compression.DecompressPtrToPtr(loadToThisPtrInstead, dstChunkSize, ptrSrcComp, m_PeekData.peekCompressedSize, cmpFormat)
            Else
                GetNextChunk = Compression.DecompressPtrToPtr(dstChunkStream.Peek_PointerOnly(0), dstChunkSize, ptrSrcComp, m_PeekData.peekCompressedSize, cmpFormat)
            End If
            
        'Anything else is a normal chunk.  If the underlying stream is file-based, copy the relevant bytes into the
        ' destination stream.  If the underlying stream is memory-based, simply wrap a "fake" array around the source
        ' bytes, to avoid the need for a copy.
        Else
            
            If (loadToThisPtrInstead <> 0) Then
                GetNextChunk = m_Stream.ReadBytesToBarePointer(loadToThisPtrInstead, dstChunkSize)
            Else
                GetNextChunk = m_Stream.ReadBytesToStream(dstChunkStream, dstChunkSize)
            End If
            
        End If
        
        'Always end by ensuring our stream pointer is aligned, then peeking the next chunk's data.
        m_Stream.SetPosition m_PeekData.peekOffset + m_PeekData.peekSize + 8, FILE_BEGIN
        PeekNextChunk
        
    End If
    
End Function

'Internal helper for various GetNextChunk_ functions.
Private Function GetNextChunk_Helper(ByRef dstChunkName As String, ByRef dstChunkSize As Long) As Boolean
    
    GetNextChunk_Helper = True
    
    'Chunk data has already been "peeked" and the stream pointer already points at the
    ' data segment of this chunk.  Return name and size values from our cache.
    dstChunkName = m_PeekData.peekName
    dstChunkSize = m_PeekData.peekSize
    
    'Failsafe check to ensure chunk size doesn't extend past EOF
    If (m_PeekData.peekOffset + dstChunkSize + 8 > m_Stream.GetStreamSize()) Or (dstChunkSize < 0) Then
        InternalError "GetNextChunk", "Bad chunk size; it extends past EOF"
        GetNextChunk_Helper = False
        Exit Function
    End If
    
    'Handle invalid sizes
    If (m_PeekData.peekSize > 0) And (m_PeekData.peekSize < 12) Then
        InternalError "GetNextChunk", "Bad chunk size; non-zero chunks must be at least 12-bytes long"
        GetNextChunk_Helper = False
        Exit Function
    End If
    
    'Handle the special-case EOF chunk
    If (dstChunkName = "PEND") Then
        If (dstChunkSize <> 0) Then InternalError "GetNextChunk", "Bad PEND chunk size (" & dstChunkSize & ")"
        GetNextChunk_Helper = False
        Exit Function
    End If
    
    'Zero-sized chunks are allowed; negative chunk sizes are not
    If (dstChunkSize = 0) Then
        GetNextChunk_Helper = True
        PeekNextChunk
        Exit Function
    ElseIf (dstChunkSize < 0) Then
        InternalError "GetNextChunk", "chunk has a bad size (" & dstChunkSize & ")"
        GetNextChunk_Helper = False
        Exit Function
    End If
    
End Function

'Open an existing chunky pdPackage file.  The file will automatically be validated and the file pointer moved to the
' start of the first chunk in the file.  You can also pass a long to srcPackageVersion to verify the version of the
' embedded file (but for the most part, the package manager will handle that kind of low-level data for you).
Friend Function OpenPackage_File(ByRef srcFilename As String, Optional ByRef srcPackageVersion As Long = 1) As Boolean
    
    Me.Reset
    If m_Stream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, srcFilename, optimizeAccess:=OptimizeSequentialAccess) Then OpenPackage_File = ValidateExternalPackage(srcPackageVersion)
    
    'If the file validates, pre-load the first chunk's information
    If OpenPackage_File Then PeekNextChunk
    
End Function

Friend Function OpenPackage_Memory(ByVal ptrData As Long, ByVal dataLength As Long, Optional ByRef srcPackageVersion As Long = 1) As Boolean
    
    Me.Reset
    If m_Stream.StartStream(PD_SM_ExternalPtrBacked, PD_SA_ReadOnly, , dataLength, ptrData) Then OpenPackage_Memory = ValidateExternalPackage(srcPackageVersion)
    
    'If the file validates, pre-load the first chunk's information
    If OpenPackage_Memory Then PeekNextChunk
    
End Function

'Reset the current package to an uninitialized state.  For an in-memory package, this frees all associated
' memory.  For a file-based package, this closes the file handle - but the function should not be used
' for this purpose, as the resulting package will likely be invalid!  Use the .EndPackage() function instead!
Friend Sub Reset()
    Set m_Stream = New pdStream
End Sub

Friend Sub SkipToNextChunk()
    m_Stream.SetPosition m_PeekData.peekOffset + m_PeekData.peekSize + 8, FILE_BEGIN
    PeekNextChunk
End Sub

'Start a new chunk.  The chunk can be written to in segments, but you *must* finish an in-progress chunk
' before starting a new one.  (Starting a new chunk erases any data left-over from previous in-progress
' chunks.)  If you want compression and/or encryption, you will specify that in the EndChunk() function.
Friend Sub StartChunk(ByVal chunkID As String)
    m_TmpChunkID = chunkID
    ReadyTmpChunkSrc
End Sub

'Start a new pdPackage file.
' REQUIRED PARAMETERS:
' - dstFilename: Unicode-aware filename.  Will be created if it does not exist; erased if it does exist.
' OPTIONAL PARAMETERS:
' - isTempFile: set to TRUE for temp files; controls the FILE_ATTRIBUTE_TEMPORARY flag used with CreateFile
' RETURNS:
' - Boolean: TRUE if destination file handle was created successfully; FALSE otherwise.
Friend Function StartNewPackage_File(ByRef dstFilename As String, Optional ByVal isTempFile As Boolean = False, Optional ByVal estimateOfInitialSize As Long = 2 ^ 16) As Boolean
    Me.Reset
    Dim fFlags As PD_FILE_ACCESS_OPTIMIZE
    fFlags = OptimizeSequentialAccess
    If isTempFile Then fFlags = fFlags Or OptimizeTempFile
    StartNewPackage_File = m_Stream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadWrite, dstFilename, estimateOfInitialSize, optimizeAccess:=fFlags)
    If StartNewPackage_File Then WriteFileHeader
End Function

'Start a new in-memory pdPackage.
' REQUIRED PARAMETERS: none
' OPTIONAL PARAMETERS: none
' RETURNS:
' - Boolean: TRUE if an initial memory allocation was successful; FALSE otherwise.
Friend Function StartNewPackage_Memory(Optional ByVal initMemoryAllocation As Long = 0) As Boolean
    Me.Reset
    StartNewPackage_Memory = m_Stream.StartStream(PD_SM_MemoryBacked, PD_SA_ReadWrite, , initMemoryAllocation)
    If StartNewPackage_Memory Then WriteFileHeader
End Function

Private Sub WriteFileHeader()
    m_Stream.WriteString_ASCII CHUNKY_PDPACKAGE_ID
    m_Stream.WriteLong PACKAGE_VERSION
End Sub

Private Sub Class_Initialize()
    Me.Reset
End Sub

Private Function GetCompressionFormatFromName(ByRef srcName As String) As PD_CompressionFormat

    If (srcName = "none") Then
        GetCompressionFormatFromName = cf_None
    ElseIf (srcName = "zstd") Then
        GetCompressionFormatFromName = cf_Zstd
    ElseIf (srcName = "lz4 ") Then
        GetCompressionFormatFromName = cf_Lz4
    ElseIf (srcName = "defl") Then
        GetCompressionFormatFromName = cf_Deflate
    ElseIf (srcName = "zlib") Then
        GetCompressionFormatFromName = cf_Zlib
    ElseIf (srcName = "gzip") Then
        GetCompressionFormatFromName = cf_Gzip
    Else
        InternalError "GetCompressionFormatFromName", srcName & " is not a valid compression format name"
    End If

End Function

Private Function GetNameFromCompressionFormat(ByVal srcFormat As PD_CompressionFormat) As String

    If (srcFormat = cf_None) Then
        GetNameFromCompressionFormat = "none"
    ElseIf (srcFormat = cf_Zstd) Then
        GetNameFromCompressionFormat = "zstd"
    ElseIf (srcFormat = cf_Lz4) Or (srcFormat = cf_Lz4hc) Then
        GetNameFromCompressionFormat = "lz4 "
    ElseIf (srcFormat = cf_Deflate) Then
        GetNameFromCompressionFormat = "defl"
    ElseIf (srcFormat = cf_Zlib) Then
        GetNameFromCompressionFormat = "zlib"
    ElseIf (srcFormat = cf_Gzip) Then
        GetNameFromCompressionFormat = "gzip"
    Else
        InternalError "GetNameFromCompressionFormat", CStr(srcFormat) & " is not a valid compression format ID"
    End If

End Function

'"Peek" at the next chunk in line.  This will move the stream pointer to the start of the
' peeked chunk's data segment, so subsequent read functions will need to note this and
' adjust their parsing accordingly.
Private Sub PeekNextChunk()
    
    'Peek the new data
    With m_PeekData
        .peekOffset = m_Stream.GetPosition()
        .peekName = m_Stream.ReadString_ASCII(4)
        .peekSize = m_Stream.ReadLong()
        
        'Non-zero-sized chunks must report compression data
        If (.peekSize <> 0) Then
            .peekCompType = GetCompressionFormatFromName(m_Stream.ReadString_ASCII(4))
            .peekUncompressedSize = m_Stream.ReadLong()
            .peekCompressedSize = m_Stream.ReadLong()
        End If
        
    End With
    
End Sub

'Ensure reusable temporary chunk object(s) are instantiated
Private Sub ReadyTmpChunkDst(Optional ByVal ensureBytesAvailable As Long = 0)
    If (m_TmpChunkDst Is Nothing) Then
        Set m_TmpChunkDst = New pdStream
        m_TmpChunkDst.StartStream PD_SM_MemoryBacked, PD_SA_ReadWrite
    Else
        m_TmpChunkDst.SetPosition 0, FILE_BEGIN
    End If
    If (ensureBytesAvailable <> 0) Then m_TmpChunkDst.EnsureBufferSpaceAvailable ensureBytesAvailable
End Sub

Private Sub ReadyTmpChunkSrc(Optional ByVal ensureBytesAvailable As Long = 0)
    If (m_TmpChunkSrc Is Nothing) Then
        Set m_TmpChunkSrc = New pdStream
        m_TmpChunkSrc.StartStream PD_SM_MemoryBacked, PD_SA_ReadWrite
    Else
        m_TmpChunkSrc.SetPosition 0, FILE_BEGIN
    End If
    If (ensureBytesAvailable <> 0) Then m_TmpChunkSrc.EnsureBufferSpaceAvailable ensureBytesAvailable
End Sub

'After opening a package (either from memory or file), call this function to validate its contents.
' RETURNS: boolean TRUE if package appears valid; FALSE otherwise.
Private Function ValidateExternalPackage(Optional ByRef srcPackageVersion As Long = 2) As Boolean
    m_Stream.SetPosition 0, FILE_BEGIN
    ValidateExternalPackage = (m_Stream.ReadString_ASCII(8) = "PDPKCHNK")
    If ValidateExternalPackage Then srcPackageVersion = m_Stream.ReadLong()
End Function

'Given a source pointer, write the contents as a new chunk.  Any/all validation needs to be performed
' *prior* to calling this function, as it is speed-optimized and it performs absolutely no validation
' on the source data.
Private Function WriteChunkFromPtr_Safe(ByRef chunkID As String, ByVal ptrChunkData As Long, ByVal chunkDataLength As Long, Optional ByVal cmpFormat As PD_CompressionFormat = cf_None, Optional ByVal cmpLevel As Long = -1) As Boolean
    
    'Assume success; we'll && this with the result of individual writes to determine success/fail state
    WriteChunkFromPtr_Safe = True
    
    'Separate handling by compression type
    Dim useCompression As Boolean
    useCompression = (cmpFormat <> cf_None)
    
    If useCompression And (chunkDataLength > 0) And (ptrChunkData <> 0) Then
    
        'When compression *is* in use, we have to create a temporary copy of the source data, as we don't
        ' know its final, compressed size until actually compress it.
        
        'Make sure we have sufficient bytes available for a worst-case compression.
        Dim finalSize As Long
        finalSize = Compression.GetWorstCaseSize(chunkDataLength, cmpFormat, cmpLevel)
        ReadyTmpChunkDst finalSize
        
        'Perform compression and retrieve the final compressed size (which is hopefully much smaller
        ' than the worst-case size!)
        If Compression.CompressPtrToPtr(m_TmpChunkDst.Peek_PointerOnly(0), finalSize, ptrChunkData, chunkDataLength, cmpFormat, cmpLevel) Then
            
            'Compression worked!  If the data's size shrunk, write out a compressed data stream.
            If (finalSize < chunkDataLength) Then
                WriteChunkFromPtr_Safe = WriteChunkFromPtr_Safe And m_Stream.WriteString_ASCII(chunkID)
                WriteChunkFromPtr_Safe = WriteChunkFromPtr_Safe And m_Stream.WriteLong(finalSize + 12)
                WriteChunkFromPtr_Safe = WriteChunkFromPtr_Safe And m_Stream.WriteString_ASCII(GetNameFromCompressionFormat(cmpFormat))
                WriteChunkFromPtr_Safe = WriteChunkFromPtr_Safe And m_Stream.WriteLong(chunkDataLength)
                WriteChunkFromPtr_Safe = WriteChunkFromPtr_Safe And m_Stream.WriteLong(finalSize)
                WriteChunkFromPtr_Safe = WriteChunkFromPtr_Safe And (m_Stream.WriteBytesFromPointer(m_TmpChunkDst.Peek_PointerOnly(0), finalSize) <> 0)
            Else
                PDDebug.LogAction "Compression didn't improve chunk size (" & finalSize & " vs " & chunkDataLength & "); writing uncompressed chunk instead."
                useCompression = False
            End If
                
        Else
            InternalError "WriteChunkFromPtr_Safe", "Compression failed; writing node uncompressed instead"
            useCompression = False
        End If
        
    End If
    
    'If compression failed, or the user doesn't want compression, write the chunk data as-is
    If (Not useCompression) Then
    
        WriteChunkFromPtr_Safe = WriteChunkFromPtr_Safe And m_Stream.WriteString_ASCII(chunkID)
        
        If (chunkDataLength = 0) Then
            WriteChunkFromPtr_Safe = WriteChunkFromPtr_Safe And m_Stream.WriteLong(0)
        Else
            WriteChunkFromPtr_Safe = WriteChunkFromPtr_Safe And m_Stream.WriteLong(chunkDataLength + 12)
            WriteChunkFromPtr_Safe = WriteChunkFromPtr_Safe And m_Stream.WriteString_ASCII(GetNameFromCompressionFormat(cf_None))
            WriteChunkFromPtr_Safe = WriteChunkFromPtr_Safe And m_Stream.WriteLong(chunkDataLength) 'Uncompressed and compressed sizes are identical
            WriteChunkFromPtr_Safe = WriteChunkFromPtr_Safe And m_Stream.WriteLong(chunkDataLength)
            WriteChunkFromPtr_Safe = WriteChunkFromPtr_Safe And (m_Stream.WriteBytesFromPointer(ptrChunkData, chunkDataLength) <> 0)
        End If
        
    End If
    
    'If, in the future, we decide to align chunks, we must ensure the stream pointer is aligned correctly before exiting.
    'If (chunkDataLength <> 0) Then...
    
End Function

Private Sub InternalError(ByVal fncName As String, ByVal errDescription As String)
    PDDebug.LogAction "WARNING!  pdPackageChunky." & fncName & "() reported an error: " & errDescription
End Sub

Private Sub Class_Terminate()
    If (Not m_Stream Is Nothing) Then m_Stream.StopStream
    Set m_Stream = Nothing
End Sub
