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
'PhotoDemon "pdPackageChunky" v1.0 Interface (e.g. chunk-based archive handler)
'Copyright 2014-2019 by Tanner Helland
'Created: 05/April/14
'Last updated: 03/June/19
'Last update: spin off similar pdPackage ideas into a chunk-based format, which can be more easily updated
'             with new/modified features as PD evolves
'Dependencies: - pdStream class from photodemon.org (used to easily read/write memory and file buffers)
'              - pdFSO class from photodemon.org (Unicode-friendly file operations)
'              - pdStringStack class from photodemon.org (optimized handling of large string collections)
'              - VBHacks module from photodemon.org (optimized workarounds for lacking VB functionality)
'              - If you want compression support, one or more 3rd-party compression libraries
'                (e.g. zstd, libdeflate) are also required.
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
' 1) Decoders can skip over any chunks they can't/don't support, making forward- and backward-compatibility
'    very straightforward.
' 2) Changes to one chunk do not affect other chunks; nothing is based on relative offsets
' 3) No need for a central directory (like zip files) that needs to be updated during writes
' 4) Individual chunks can be individually compressed and/or encrypted using custom per-chunk settings,
'    without affecting any neighboring chunks.
'
'The file format itself is very simple.  A few universal notes before I document it:
' - All numeric data types are little-endian.
' - Individual chunks must be 2 GB or less in size.  To store larger amounts of data, spread it across
'   multiple chunks.
' - Unless otherwise noted, all string data is in UTF-8 format.  (The term "ASCII" is also used for
'   various bits, and as you hopefully know if you're a developer, UTF-8 and ASCII are identical
'   for code points on the range [0, 127] - so any distinction for chars on this range is primarily
'   pedantic.)
'
'Files begin with a hard-coded 8-byte "magic string", identifying the file as a chunky pdPackage:
' "PDPKCHNK" (8-byte ASCII string)
'
'This is followed by a 4-byte unsigned integer describing the encoder software version.  This value can
' be used to identify files that may contain new, possibly unsupported features.  At the time of this
' writing, only a version value of "1" (0x00000001) is defined.
'
'The remainder of the file is comprised of zero or more chunks.  Chunks follow a predefined format:
' - 4 byte ASCII string defining the chunk's type.  (This string is case-sensitive.)
' - 4 byte signed integer defining the size of this chunk.  For compatibility purposes, this value
'   must be >= 0, which limits its max size to 2 GB.  Importantly, note that this size value
'   does *not* include the 8 bytes required for the ASCII string identifier and chunk size; thus an
'   empty chunk (consisting of just a 4-byte ID and 4-byte size value) would have a size of 0.
'
'Chunks must be iterated until an "PEND" (package end) chunk is reached.  This chunk always has a size
' of 0, and no chunks may appear after it.  (In fact, any data whatsoever past the "PEND" chunk *must
' be ignored*.)
'
'For the most part, chunk names and definitions are left up to individual programs.  This spec only
' defines a handful of protected chunk names.
'
'Including the "PEND" chunk (described above), the following chunk names are protected and *must not*
' be used except as described in this document.
'
'"COMP" chunk: compression
'
'The "COMP" ID defines a compressed chunk.  "COMP" chunks are unique because they wrap another
' (entire) chunk inside themselves.  "COMP" chunks allow any individual chunk to be compressed,
' regardless of contents.
'
'After the required 4-byte ASCII ID and 4-byte size descriptor, a COMP chunk contains the following:
' 1) 4-byte ASCII ID of the chunk that has been compressed.  As with a normal chunk, decoders can use this
'    to determine whether or not they support (or care about) the contents of the compressed chunk -
'    and thus, whether to perform decompression or skip it entirely.
' 2) 4-byte signed integer describing the *uncompressed size* of the chunk that has been compressed.
'    This *does not* include the 8-bytes that would normally be required for the 4-byte ASCII ID and
'    4-byte chunk size descriptor, because those 8-bytes are *not* part of the compression stream.
'    (Said another way, only the data portion of the embedded chunk is compressed - the 4-byte ASCII ID
'    and chunk size are stored here, separate from the data.)
' 3) A 4-byte ASCII string defining the compression format used.  Valid compression types include:
'  - "none" (no compression)
'  - "zstd" (zstandard)
'  - "defl" (deflate)
'    Other compression types may be used, but encoders/decoders are not required to implement them.
'    For example, internal PhotoDemon files (like autosave data) use the lz4 compression format,
'    with an "lz4 " descriptor - but these files are not exposed externally, and 3rd-party support isn't
'    required or even desired.
' 4) A 4-byte signed integer describing the size of the compressed data that follows.  This value
'    could theoretically be inferred by taking the size of this "COMP" chunk minus 12 (for the preceding
'    3x4 byte requirements), but clarifying the value here allows writers to pad overall "COMP" chunk
'    length to a multiple or 4 (or any other value) as desired, independent of actual compressed
'    data size.
' 5) The compressed data stream follows.  It must be the same size described by (4), above, but it need
'    not consume the entire remaining length of the "COMP" chunk.
' 6) Any unused bytes beyond the end of (5) are considered "undefined" and should be ignored.
'
'Finally, to be absolutely clear: "COMP" chunks are allowed to store any type of chunk *except "PEND"
' end-of-package chunks*. "PEND" chunks must always be stored as-is, without compression.
'
'"ENCR" chunk: encryption
'
'"ENCR" chunks behave similarly to "COMP" chunks.  They wrap another chunk inside themselves, using some
' manner of encryption.
'
'The details of "ENCR" chunks have not been finalized, but the "ENCR" chunk name remains reserved for
' this purpose in the future.  Don't use it for your own chunks.
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
Private Const PACKAGE_VERSION As Long = &H1

'Hard-coded chunky pdPackage file identifier
Private Const CHUNKY_PDPACKAGE_ID As String = "PDPKCHNK"

'The currently assembled chunk data.  This may wrap a file or memory instance; both must be supported.
Private m_Stream As pdStream

'When writing chunks in pieces (whether compressed or otherwise), we need temporary holders for
' chunk data.  Persistent temporary streams are used, and by default, their memory is *not* deallocated
' between chunks.  This prevents cache thrashing when writing a bunch of similarly-sized chunks in a row
' (with or without compression).
Private m_TmpChunkSrc As pdStream, m_TmpChunkDst As pdStream

'When writing a chunk in pieces, we need to remember the chunk ID (because it's specified in the
' StartChunk() function, but not actually written until the EndChunk() function)
Private m_TmpChunkID As String

'When opening an existing package, this class automatically "peeks" at the next chunk in line.
' This means that the current stream pointer *ALWAYS* points at the data segment of the current
' chunk.  The current chunk's metadata should be retrieved from m_PeekData, as required.
' (This approach provides meaningful perf improvements as we can read the file fully
' sequentially, without ever needing to back-up the stream pointer to maintain chunk alignment.)
Private Type PackageChunkPeek
    peekActive As Boolean
    peekName As String
    peekSize As Long
    peekOffset As Long
    'The following members are only used for COMP (compression) chunks
    peekCompName As String
    peekCompType As PD_CompressionFormat
    peekCompCompressedSize As Long
    peekCompUncompressedSize As Long
    peekCompDataOffset As Long
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
    If (Len(chunkID) <> 4) Then chunkID = ValidateChunkID(chunkID)
    
    'The rest of the write process is handled by a dedicated function
    AddChunk_WholeFromPtr = WriteChunkFromPtr_Safe(chunkID, ptrChunkData, chunkDataLength, cmpFormat, cmpLevel)
    
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
        If (Len(chunkID) <> 4) Then chunkID = ValidateChunkID(chunkID)
        
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
    If (m_PeekData.peekName = "COMP") Then
        GetChunkName = m_PeekData.peekCompName
    Else
        GetChunkName = m_PeekData.peekName
    End If
End Function

'Return the size of the next chunk in line.  Note that this does *not* permanently move the
' stream pointer (by design)!
Friend Function GetChunkSize() As Long
    If (m_PeekData.peekName = "COMP") Then
        GetChunkSize = m_PeekData.peekCompUncompressedSize
    Else
        GetChunkSize = m_PeekData.peekSize
    End If
End Function

Friend Function GetPackageSize() As Long
    GetPackageSize = m_Stream.GetStreamSize()
End Function

'Retrieve the next chunk in the file.
' RETURNS: boolean TRUE if another chunk exists, FALSE if no more chunks exist.  (Note that the package end chunk
' "PEND" results in a FALSE return, by design, even though dstChunkName/Size will still be set and returned correctly.)
' Note: the destination stream may not be the same size reported by dstChunkSize.  This is especially true if you're
' using the same pdStream object multiple times in a row, as this class will attempt to reuse previous memory
' allocations whenever possible.
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
        If (dstChunkName = "COMP") Then
            
            'Overwrite the returned name and size with the name and *uncompressed* size of the compressed chunk
            dstChunkName = m_PeekData.peekCompName
            dstChunkSize = m_PeekData.peekCompUncompressedSize
            
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
                m_Stream.ReadBytesToStream m_TmpChunkSrc, m_PeekData.peekCompCompressedSize
                ptrSrcComp = m_TmpChunkSrc.Peek_PointerOnly(0)
            Else
                ptrSrcComp = m_Stream.Peek_PointerOnly(, m_PeekData.peekCompCompressedSize)
            End If
            
            If (loadToThisPtrInstead <> 0) Then
                GetNextChunk = Compression.DecompressPtrToPtr(loadToThisPtrInstead, dstChunkSize, ptrSrcComp, m_PeekData.peekCompCompressedSize, cmpFormat)
            Else
                GetNextChunk = Compression.DecompressPtrToPtr(dstChunkStream.Peek_PointerOnly(0), dstChunkSize, ptrSrcComp, m_PeekData.peekCompCompressedSize, cmpFormat)
            End If
            
            'Move the pointer in the source data to the next chunk, then peek the next chunk
            m_Stream.SetPosition m_PeekData.peekOffset + m_PeekData.peekSize + 8, FILE_BEGIN
            PeekNextChunk
            
        'Anything else is a normal chunk.  If the underlying stream is file-based, copy the relevant bytes into the
        ' destination stream.  If the underlying stream is memory-based, simply wrap a "fake" array around the source
        ' bytes, to avoid the need for a copy.
        Else
            If (dstChunkSize <> 0) Then
                If (loadToThisPtrInstead <> 0) Then
                    GetNextChunk = m_Stream.ReadBytesToBarePointer(loadToThisPtrInstead, dstChunkSize)
                Else
                    GetNextChunk = m_Stream.ReadBytesToStream(dstChunkStream, dstChunkSize)
                End If
            End If
            PeekNextChunk
        End If
        
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
Friend Function StartChunk(ByVal chunkID As String)
    m_TmpChunkID = chunkID
    ReadyTmpChunkSrc
End Function

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
        .peekActive = True
        .peekOffset = m_Stream.GetPosition()
        .peekName = m_Stream.ReadString_ASCII(4)
        .peekSize = m_Stream.ReadLong()
        
        'Compressed chunks have extra data; it is useful to read this in advance
        If (.peekName = "COMP") Then
            .peekCompName = m_Stream.ReadString_ASCII(4)
            .peekCompUncompressedSize = m_Stream.ReadLong()
            .peekCompType = GetCompressionFormatFromName(m_Stream.ReadString_ASCII(4))
            .peekCompCompressedSize = m_Stream.ReadLong()
            .peekCompDataOffset = m_Stream.GetPosition()
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

Private Function ValidateChunkID(ByRef srcID As String) As String
    If (Len(srcID) = 4) Then
        ValidateChunkID = srcID
    ElseIf (Len(srcID) > 4) Then
        InternalError "ValidateChunkID", srcID & " is an invalid chunk ID - only 4 chars allowed.  Truncating now."
        ValidateChunkID = Left$(srcID, 4)
    
    'Lengths less than 4 are fine; they're just extended with spaces
    Else
        ValidateChunkID = srcID & String$(4 - Len(srcID), " ")
    End If
End Function

'After opening a package (either from memory or file), call this function to validate its contents.
' RETURNS: boolean TRUE if package appears valid; FALSE otherwise.
Private Function ValidateExternalPackage(Optional ByRef srcPackageVersion As Long = 1) As Boolean
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
    
    If useCompression Then
    
        'When compression *is* in use, we have to create a temporary copy of the source data, as we don't
        ' know its final, compressed size until actually compress it.
        
        'Make sure we have sufficient bytes available for a worst-case compression.
        Dim finalSize As Long
        finalSize = Compression.GetWorstCaseSize(chunkDataLength, cmpFormat, cmpLevel)
        ReadyTmpChunkDst finalSize
        
        'Perform compression and retrieve the final compressed size (which is hopefully much smaller
        ' than the worst-case size!)
        If Compression.CompressPtrToPtr(m_TmpChunkDst.Peek_PointerOnly(0), finalSize, ptrChunkData, chunkDataLength, cmpFormat, cmpLevel) Then
            
            'Compression worked!  Write a full COMP chunk.
            WriteChunkFromPtr_Safe = WriteChunkFromPtr_Safe And m_Stream.WriteString_ASCII("COMP")
            WriteChunkFromPtr_Safe = WriteChunkFromPtr_Safe And m_Stream.WriteLong(finalSize + 16)
            WriteChunkFromPtr_Safe = WriteChunkFromPtr_Safe And m_Stream.WriteString_ASCII(chunkID)
            WriteChunkFromPtr_Safe = WriteChunkFromPtr_Safe And m_Stream.WriteLong(chunkDataLength)
            WriteChunkFromPtr_Safe = WriteChunkFromPtr_Safe And m_Stream.WriteString_ASCII(GetNameFromCompressionFormat(cmpFormat))
            WriteChunkFromPtr_Safe = WriteChunkFromPtr_Safe And m_Stream.WriteLong(finalSize)
            WriteChunkFromPtr_Safe = WriteChunkFromPtr_Safe And (m_Stream.WriteBytesFromPointer(m_TmpChunkDst.Peek_PointerOnly(0), finalSize) <> 0)
            
        Else
            InternalError "WriteChunkFromPtr_Safe", "Compression failed; writing node uncompressed instead"
            useCompression = False
        End If
        
    End If
    
    'If compression failed, or the user doesn't want compression, write the chunk data as-is
    If (Not useCompression) Then
        WriteChunkFromPtr_Safe = WriteChunkFromPtr_Safe And m_Stream.WriteString_ASCII(chunkID)
        WriteChunkFromPtr_Safe = WriteChunkFromPtr_Safe And m_Stream.WriteLong(chunkDataLength)
        WriteChunkFromPtr_Safe = WriteChunkFromPtr_Safe And (m_Stream.WriteBytesFromPointer(ptrChunkData, chunkDataLength) <> 0)
    End If
    
End Function

Private Function InternalError(ByVal fncName As String, ByVal errDescription As String)
    PDDebug.LogAction "WARNING!  pdPackageChunky." & fncName & "() reported an error: " & errDescription
End Function
