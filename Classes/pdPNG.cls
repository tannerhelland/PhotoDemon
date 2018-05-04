VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pdPNG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'PhotoDemon PNG Container and Parser
'Copyright 2018-2018 by Tanner Helland
'Created: 12/April/18
'Last updated: 23/April/18
'Last update: handle cHRM chunks (with the final corrections applied via LittleCMS)
'
'I have tried - *so* hard - to make 3rd-party PNG solutions work.  But they all suck.  Every last one of them.
' Like most standards attempts, PNG is an overcomplicated, hackish mess, and every library that attempts to
' wrap the format only makes things worse for developers.  (I mean, let's take LibPNG as an example.  It is
' the official PNG reference library, so it must be well-constructed, right?  It wouldn't do something asinine
' like implementing error-handling via setjmp/longjmp, would it?  Oh wait - it would?  @#$#$^@#$#^*^!!)
'
'After wasting dozens of hours fighting 3rd-party libraries, I gave up and wrote my own PNG parser.
' PhotoDemon needs fairly intricate access to PNG internals - during both reading and writing - and given
' the format's current ubiquity for images < 24-bpp, it makes sense to handle PNGs manually.  Since we ship
' zLib anyway due to legacy compression tasks, we have everything we need to do this, and it frees me from
' the nightmare of trying to make legacy C-libraries work with VB6.
'
'At present, this class only supports PNG importing.  Exporting is a WIP.  I may revisit the API design after
' finalizing export tasks.
'
'As of v7.2 nightly builds, this class *is* used as PD's primary PNG importer.  If you encounter bugs,
' please let me know so I can resolve them before 7.2 stable releases.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'PNG loading is complicated, and a lot of things can go wrong.  Instead of returning binary "success/fail"
' values, we return specific flags; "warnings" may be recoverable and you can still attempt to load the file.
' "Failure" returns are unrecoverable and processing *must* be abandoned.  (As a convenience, you can treat
' the "warning" and "failure" values as flags; specific warning/failure states in each category will share
' the same high flag bit.)
Public Enum PD_PNGResult
    png_Success = 0
    png_Warning = 256
    png_Failure = 65536
    png_FileNotPNG = 16777217
End Enum

#If False Then
    Private Const png_Success = 0, png_Warning = 256, png_Failure = 65536, png_FileNotPNG = 16777217
#End If

'PNGs use a standard header, which we convert into this VB-friendly version.  Note that not all information
' from the IHDR chunk is stored here; entries like Compression and Filter only have one allowed value, so while
' we validate these at load-time, we don't go to the trouble of storing them internally.  Similarly, some of
' these values (like bits-per-pixel) are not formally stored in the original header - instead, we infer them.
' (NOTE: the header type is declared publicly in PD, so we can transfer it to child classes.)
Private m_Header As PD_PNGHeader

'If warnings are encountered during processing, we push their messages onto a string stack.  (We may
' decide to report these to the user... I haven't decided yet.  Either way, it'll be helpful for debugging.)
Private m_Warnings As pdStringStack

'pdStream makes our life much easier!
Private m_Stream As pdStream

'At present, we require the caller to pass an identical source file path to every load function.
' (This is a cheap and easy way to ensure no funny business is afoot!)
Private m_SourceFilename As String

'In interlaced images, the number of pixels-per-reduced image is non-obvious.  We calculate these
' line lengths as part of the decompression pass, and rather than re-calculate them on subsequent passes,
' we simply cache the results and re-use them later.  (Note that there is no "y" byte count, because it
' is the same as the pixel count.)
Private m_XPixelCount() As Long, m_YPixelCount() As Long
Private m_XByteCount() As Long

'Chunk collection.  The size of the collection *may not match* the number of chunks actually loaded.
' (For performance reasons, the array is allocated using a power-of-2 strategy.)
Private Const INIT_NUM_OF_CHUNKS As Long = 8
Private m_NumOfChunks As Long
Private m_Chunks() As pdPNGChunk

'Simplified wrapper to load a PNG automatically.
Friend Function LoadPNG_Simple(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB, Optional ByVal checkExtension As Boolean = False) As PD_PNGResult
    
    'Try to validate the source file
    Dim keepLoading As PD_PNGResult
    keepLoading = Me.Step1_ValidatePNG(srcFile, checkExtension)
    
    'As long as warnings at critical, we'll continue to try and load the file.  (Only outright failures
    ' will cause us to abandon the attempt entirely.)
    If (keepLoading < png_Failure) Then
    
        'Parse the original file into chunks.  (Once this is done, we no longer need to maintain an
        ' in-memory copy of the original PNG data.)
        keepLoading = Me.Step2_PreLoadChunks(srcFile)
        
        'If all chunks were loaded successfully and minimal validation passed (e.g. a minimal amount of
        ' required chunks are present), go ahead and decompress all relevant chunks, including merging
        ' multiple IDAT chunks (if any) into a single IDAT instance.
        If (keepLoading < png_Failure) Then keepLoading = Me.Step3_Decompress(srcFile)
        
        'All compressed data has now been decompressed.  Our next task is to "un-filter" the IDAT chunk,
        ' which converts the pixel bytestream (which has been "filtered" into some other representation
        ' format) into a raw stream of actual pixel data.
        If (keepLoading < png_Failure) Then keepLoading = Me.Step4_UnfilterIDAT(srcFile)
        
        'Next, convert the unfiltered IDAT into actual pixel data.
        If (keepLoading < png_Failure) Then keepLoading = Me.Step5_ConstructImage(srcFile, dstDIB, dstImage)
        
        'Finally, perform any weird post-processing due to non-standard gamma or chromaticity chunks
        If (keepLoading < png_Failure) Then keepLoading = Me.Step6_PostProcessing(srcFile, dstDIB, dstImage)
        
    End If
    
    LoadPNG_Simple = keepLoading
    
End Function

'The first step toward loading a PNG is validating it.  Do this first, before attempting anything else.
Friend Function Step1_ValidatePNG(ByRef srcFile As String, Optional ByVal checkExtension As Boolean = False) As PD_PNGResult
    
    On Error GoTo InternalVBError
    
    ResetChunks
    m_SourceFilename = srcFile
    
    Dim okToProceed As PD_PNGResult
    okToProceed = png_Success
    
    'We always check the file extension.  If the user has *asked* us to check it, we treat extension
    ' mismatches as a failure state.  (Otherwise, it will only raise a warning.)
    If Strings.StringsNotEqual(Right$(srcFile, 3), "png", True) Then
        m_Warnings.AddString "File extension doesn't match PNG"
        If checkExtension Then okToProceed = png_FileNotPNG Else okToProceed = png_Warning
    End If
    
    'PNG files must have a certain minimum size (comprising a valid magic number - 8 bytes - and at least
    ' three valid chunks (IHDR, IDAT, IEND); 8 + 12 * 3 = 44).  An empty IDAT chunk would still mean a
    ' broken image, but that's okay; this is just a minor failsafe.
    If (okToProceed < png_Failure) Then
        If (Files.FileLenW(srcFile) < 44) Then
            m_Warnings.AddString "File size is too small to contain valid PNG data"
            okToProceed = png_Failure
        End If
    End If
    
    'If all pre-checks passed, open a stream.  (For performance reasons, we cache the entire stream
    ' in-memory instead of streaming it in pieces.)
    If (okToProceed < png_Failure) Then
        If (Not Files.FileLoadAsPDStream(srcFile, m_Stream)) Then
            m_Warnings.AddString "Can't read file; it may be locked or in an inaccessible location."
            okToProceed = png_Failure
        End If
    End If
    
    'The stream is open; validate the PNG's "magic number".
    If (okToProceed < png_Failure) Then
        
        Dim header1 As Long, header2 As Long
        header1 = m_Stream.ReadLong_BE()
        header2 = m_Stream.ReadLong_BE()
        If (header1 <> &H89504E47) Or (header2 <> &HD0A1A0A) Then
            m_Warnings.AddString "PNG header failed basic validation.  (This is not a PNG file.)"
            okToProceed = png_FileNotPNG
        End If
        
    End If
    
    'Note an outright failure state in the debugger, and importantly, don't write this to the debug log if we're
    ' running as a compiled exe (as the PNG's failure is simply due to not being a PNG at all, and not due to
    ' any internal problem in this class).
    If (okToProceed >= png_Failure) Then
        InternalError "Step1_ValidatePNG", "file is not a valid PNG file", False
        If m_Stream.IsOpen() Then m_Stream.StopStream True
    End If
    
    Step1_ValidatePNG = okToProceed
    
    Exit Function

'Internal VB errors are always treated as catastrophic failures.
InternalVBError:
    InternalError "Step1_ValidatePNG", "internal VB error #" & Err.Number & ": " & Err.Description
    If m_Stream.IsOpen Then m_Stream.StopStream True
    
    m_Warnings.AddString "Internal error in step 1, #" & Err.Number & ": " & Err.Description
    Step1_ValidatePNG = png_Failure
    
End Function

'After validating the PNG header, we want to preload all chunks into their own containers.  This simplifies
' further processing, and it allows us to immediately release the source file.
Friend Function Step2_PreLoadChunks(ByRef srcFile As String) As PD_PNGResult

    On Error GoTo InternalVBError
    
    Step2_PreLoadChunks = png_Success
    
    'Failsafe check
    If Strings.StringsNotEqual(m_SourceFilename, srcFile, False) Then
        InternalError "Step2_PreLoadChunks", "filename has changed since original validation!"
        Step2_PreLoadChunks = png_Failure
        Exit Function
    End If
    
    'If we're still here, the m_Stream object should already be open and pointing at the file in question.
    ' The file pointer has also moved past the 8-byte PNG signature and is now pointing at the first chunk.
    ' We basically want to iterate through all chunks, dumping their raw data into new chunk classes as
    ' we go.  (We'll validate internal chunks later - first, let's just get the data parsed and sorted
    ' into discrete chunks.)
    Dim chunkSize As Long, chunkType As String, testCRC As Long, embeddedCRC As Long
    
    'While we process chunks, we also want to make sure at least one IDAT entry is found.
    Dim idatFound As Boolean
    
    Do
        
        'Chunk size is *for the data segment only*!  Chunks always have 12 bytes worth of data,
        ' (4x size, 4x type, 4x CRC32), with a data segment that is allowed to be zero-length.
        ' (And in fact, for some required chunks - like IEND that marks the end of the file - the
        ' length is required to be zero.)
        chunkSize = m_Stream.ReadLong_BE()
        
        'Because the PNG format is dumb, CRCs are calculated over the chunk type *and* data portion.
        ' Before proceeding, calculate a CRC using zLib.  (Nonsense like this is why we cache the
        ' entire source file in memory; it's a hell of a lot faster that way.)
        testCRC = Plugin_zLib.ZLib_GetCRC32(m_Stream.Peek_PointerOnly(), chunkSize + 4)
        
        'With the CRC calculated, grab the chunk type next.  (Note that the CRC calculation only used
        ' a pointer peek - it didn't actually move the stream pointer!)
        chunkType = m_Stream.ReadString_ASCII(4)
        If (chunkType = "IDAT") Then
            
            'If this is not the first IDAT chunk we've encountered, make sure that the previous chunk
            ' was also IDAT.  (The spec requires IDAT chunks to be contiguous in the file.)
            If idatFound Then
                If (Not m_Chunks(m_NumOfChunks - 1).GetType = "IDAT") Then
                    Step2_PreLoadChunks = png_Warning
                    m_Warnings.AddString "IDAT chunks are not contiguous; this file is technically invalid!"
                End If
            Else
                idatFound = True
            End If
            
        End If
        
        'Create a new chunk to hold this chunk's data
        If (m_NumOfChunks > UBound(m_Chunks)) Then ReDim Preserve m_Chunks(0 To m_NumOfChunks * 2 - 1) As pdPNGChunk
        Set m_Chunks(m_NumOfChunks) = New pdPNGChunk
        m_Chunks(m_NumOfChunks).CreateChunk chunkType, chunkSize
        
        'The chunk object itself will handle loading the chunk's data
        If (chunkSize > 0) Then m_Chunks(m_NumOfChunks).LoadChunkData m_Stream
        
        'Finally, grab the embedded CRC, make sure it matches our internal one, then pass *both* CRCs
        ' to the chunk class.
        embeddedCRC = m_Stream.ReadLong_BE()
        m_Chunks(m_NumOfChunks).NotifyCRCs testCRC, embeddedCRC
        
        If (embeddedCRC <> testCRC) Then
            Step2_PreLoadChunks = png_Warning
            m_Warnings.AddString "Checksum validation failed on chunk " & CStr(m_NumOfChunks + 1) & " - " & chunkType & " (0x" & Hex$(embeddedCRC) & " expected, 0x" & Hex$(testCRC) & " found)"
        End If
        
        'For debug purposes, it can be helpful to see what's inside a file - this will tell you about
        ' the chunks we've encountered.
        'Debug.Print m_NumOfChunks, m_Chunks(m_NumOfChunks).GetType, m_Chunks(m_NumOfChunks).GetDataSize, embeddedCRC, testCRC
        
        'Increment the chunk counter and carry on!
        m_NumOfChunks = m_NumOfChunks + 1
        
    'Continue looping as long as...
    ' 1) There are more bytes in the file, and...
    ' 2) We haven't hit the IEND chunk yet.
    Loop While (m_Stream.GetStreamSize() >= m_Stream.GetPosition() + 1) And Strings.StringsNotEqual(chunkType, "IEND", False)
    
    'Because we have parsed all relevant information from the original source file, we can immediately
    ' free that memory.
    m_Stream.StopStream True
    
    'If at least three valid chunks (IHDR, IDAT, and IEND) were found - and they were found in
    ' correct order - return SUCCESS.
    If (m_NumOfChunks >= 3) Then
        If (m_Chunks(0).GetType <> "IHDR") Then
            Step2_PreLoadChunks = png_Failure
            InternalError "Step2_PreLoadChunks", "Required header chunk (IHDR) is missing or not listed as the first chunk in the file"
        ElseIf (m_Chunks(m_NumOfChunks - 1).GetType <> "IEND") Then
            Step2_PreLoadChunks = png_Failure
            InternalError "Step2_PreLoadChunks", "Required final chunk (IEND) is missing or not listed as the final chunk in the file"
        ElseIf (Not idatFound) Then
            Step2_PreLoadChunks = png_Failure
            InternalError "Step2_PreLoadChunks", "No pixel chunks (IDAT) found in the file!"
        End If
    End If
    
    Exit Function
    
InternalVBError:
    InternalError "Step2_PreLoadChunks", "internal VB error #" & Err.Number & ": " & Err.Description
    If m_Stream.IsOpen Then m_Stream.StopStream True
    
    m_Warnings.AddString "Internal error in step 2, #" & Err.Number & ": " & Err.Description
    Step2_PreLoadChunks = png_Failure

End Function

'After loading all chunks and releasing the source file, we want to decompress any/all compressed chunks.
' While we're at it, we must merge all IDAT blocks into a single instance - they represent a *single*
' compressed data stream.
Friend Function Step3_Decompress(ByRef srcFile As String) As PD_PNGResult
    
    On Error GoTo InternalVBError
    
    Step3_Decompress = png_Success
    
    'Failsafe check(s)
    If Strings.StringsNotEqual(m_SourceFilename, srcFile, False) Then
        InternalError "Step3_Decompress", "filename has changed since original validation!"
        Step3_Decompress = png_Failure
    End If
    
    If (m_NumOfChunks < 3) Then
        InternalError "Step3_Decompress", "not enough chunks to continue!"
        Step3_Decompress = png_Failure
    End If
    
    If (Step3_Decompress >= png_Failure) Then Exit Function
    
    'If we're still here, it means that we have (at a minimum) the three required chunks for
    ' constructing pixel data, so the file is (likely) valid.  We now want to tackle two tasks:
    ' 1) Merging multiple IDAT entries, if any, into a single IDAT chunk.  (IDAT data is a single
    '    zLib stream, but encoders are allowed to split it over multiple IDAT chunks if they want;
    '    the original goal was "streaming decompression" but this is just an annoyance today.)
    ' 2) Decompressing chunks whose entire datastream was compressed via zLib.  Besides IDAT, we're
    '    primarily concerned with ICC profiles (iCCP).  Two other chunks - zTXt and optionally iTXt -
    '    can compress a *portion* of their datastream, but you have to manually scan the stream for
    '    markers and offsets, and we don't care about that in this step.  We just want enough data
    '    to start assembling pixel data (which is the most time-consuming part of the load process).
    
    'Before doing any decompression, let's merge our various IDAT chunks into a single instance.
    Dim firstIDATIndex As Long, numIDATsRemoved As Long
    firstIDATIndex = -1
    numIDATsRemoved = 0
    
    Dim i As Long
    For i = 0 To m_NumOfChunks - 1
        
        If (m_Chunks(i).GetType = "IDAT") Then
            
            'If this is the first IDAT chunk we've found, flag its index and carry on
            If (firstIDATIndex = -1) Then
                firstIDATIndex = i
            
            'If this is *not* the first IDAT chunk we've found, merge the contents of this chunk into
            ' the first IDAT chunk.
            Else
                
                'As part of the merge step, this chunk will be freed; that's by design!
                m_Chunks(firstIDATIndex).MergeOtherChunk m_Chunks(i)
                
                'Increment the multiple IDAT count; we use these to shift other chunks "forward" in
                ' the chunk list as we erase the duplicate IDATs.
                numIDATsRemoved = numIDATsRemoved + 1
                
            End If
            
        'If this is *not* an IDAT chunk, and we've removed IDAT chunks prior to this one, shift this
        ' chunk forward in the list.
        Else
            If (numIDATsRemoved > 0) Then
                Set m_Chunks(i - numIDATsRemoved) = m_Chunks(i)
                Set m_Chunks(i) = Nothing
            End If
        End If
        
    Next i
    
    'If multiple IDAT chunks were condensed into a single chunk, update our net chunk count
    m_NumOfChunks = m_NumOfChunks - numIDATsRemoved
    
    'We now want to proceed with decompression, but because the PNG format is moronic, this step isn't
    ' as simple as asking each chunk to decompress itself.  zLib streams don't store the original,
    ' uncompressed size of their data stream.  You are expected to store that data on your own.  PNG files
    ' decided not to do that.  Instead, you have to manually infer each chunks decompressed size from
    ' data unique to that chunk.  (And for some chunks - like zTXT - your only option is to repeatedly
    ' attempt to decompress the chunk, with ever-larger buffers, until a full decompress works.
    ' That's how shittily they've designed it.)
    
    'IDAT chunks are the most obnoxious to decompress, because their inflated (normal) size is a
    ' function of the image's dimensions plus its color depth.  This means we need to parse the PNG
    ' header and retrieve some critical bits before continuing.
    Step3_Decompress = PopulateHeader()
    
    'If the header didn't validate, don't decompress anything as the entire file is invalid
    If (Step3_Decompress < png_Failure) Then
    
        'The header appears to be valid.  We now need to calculate how much space is required for decompressing.
        Dim reqIDATSize As Long
        
        'Interlaced images treat each "interlace pass" as its own "mini-PNG", which means each non-empty scanline
        ' in each mini-image gets its own filtering byte.  (I know, it's fucking confusing.)  Because of this,
        ' calculating a decompression buffer size is non-trivial.
        If m_Header.Interlaced Then
            
            reqIDATSize = 0
            
            'Because it's complicated to calculate scanline pixel and byte counts, we store our calculations
            ' in this step so that future steps can re-use them.
            ReDim m_XPixelCount(0 To 6) As Long
            ReDim m_YPixelCount(0 To 6) As Long
            ReDim m_XByteCount(0 To 6) As Long
            
            'A separate function calculates how many pixels exist in each interlacing pass; we can then use
            ' our standard "bit-depth" adjustment formula on each pass, and tally the results to find a
            ' "net" required decompression buffer size.
            Dim srcX As Long, srcY As Long
            For i = 0 To 6
                
                GetSizeInterlaced srcX, srcY, i + 1
                m_XPixelCount(i) = srcX
                m_YPixelCount(i) = srcY
                m_XByteCount(i) = ((srcX * m_Header.BitsPerPixel + 7) \ 8) + 1
                
                'Per the spec, 0-byte passes (only possible on very small images!) do not get a filter type byte,
                ' so we can skip such lines entirely.
                If (m_XByteCount(i) = 1) Or (srcY = 0) Then
                    m_XByteCount(i) = 0
                Else
                    reqIDATSize = reqIDATSize + ((srcX * m_Header.BitsPerPixel + 7) \ 8) * srcY + srcY
                End If
                
            Next i
        
        'Non-interlaced images are much simpler!
        Else
            
            ReDim m_XPixelCount(0) As Long
            ReDim m_YPixelCount(0) As Long
            ReDim m_XByteCount(0) As Long
            m_XPixelCount(0) = m_Header.Width
            m_YPixelCount(0) = m_Header.Height
            m_XByteCount(0) = ((m_Header.Width * m_Header.BitsPerPixel + 7) \ 8) + 1
            
            'To properly cover the case of bit-depths < 8 (e.g. 1, 2, 4), ensure that scanline width is always
            ' rounded up to the nearest byte alignment.
            reqIDATSize = ((m_Header.Width * m_Header.BitsPerPixel + 7) \ 8) * m_Header.Height + m_Header.Height
            
        End If
        
        'Ask all relevant chunks to decompress themselves; note that chunks without compressed data will
        ' just ignore this request.
        For i = 0 To m_NumOfChunks - 1
            If (Not m_Chunks(i).DecompressChunk(m_Warnings, reqIDATSize)) Then
                
                'Decompression failures in IDAT are critical; other decompression failures only raise warnings
                If (m_Chunks(i).GetType = "IDAT") Then
                    InternalError "Step3_Decompress", "IDAT decompression failed; PNG file is unreadable."
                    Step3_Decompress = png_Failure
                Else
                    m_Warnings.AddString "WARNING: " & m_Chunks(i).GetType & " could not be decompressed.  (I'll still try to salvage pixel data.)"
                    Step3_Decompress = png_Warning
                End If
                
            Else
                'If (m_Chunks(i).GetType = "IDAT") Then Debug.Print "Successfully decompressed chunk; " & reqIDATSize
            End If
        Next i
        
    End If
    
    Exit Function
    
InternalVBError:
    InternalError "Step3_Decompress", "internal VB error #" & Err.Number & ": " & Err.Description
    Step3_Decompress = png_Failure

End Function

'Once all relevant chunks have been decompressed, we next need to un-filter each scanline into its original,
' unfiltered pixel representation.  This step operates only on bytes, which means it is completely independent
' of color-depth or bit-depth.
Friend Function Step4_UnfilterIDAT(ByRef srcFile As String) As PD_PNGResult

    On Error GoTo InternalVBError
    
    Step4_UnfilterIDAT = png_Success
    
    'Failsafe check(s)
    If Strings.StringsNotEqual(m_SourceFilename, srcFile, False) Then
        InternalError "Step4_UnfilterIDAT", "filename has changed since original validation!"
        Step4_UnfilterIDAT = png_Failure
    End If
    
    If (Step4_UnfilterIDAT >= png_Failure) Then Exit Function
    
    'We are now ready to convert the raw IDAT stream into its original, "unfiltered pixels" state.
    ' IMPORTANTLY, this step operates only on pure bytes - so 1-bpp pixel data will still be 1-bpp after this step.
    ' All this step does is "translate" pixels from their current filtered state (which is a quirky PNG term
    ' that roughly correlates to "frequency transform") to their original, unfiltered state.
    
    'Also note that filter type 0 is totally valid, which means "unfiltered", in which case we won't do shit
    ' during this stage!
    
    'Make sure an IDAT chunk exists...
    Dim idatIndex As Long
    idatIndex = Me.GetIndexOfChunk("IDAT")
    If (idatIndex >= 0) Then
        
        'Ask the IDAT chunk to unfilter itself.  (As part of this step, we pass the x/y pixel counts we
        ' already computed as part of the interlacing step - if the image was interlaced.)
        m_Chunks(idatIndex).UnfilterChunk m_Warnings, m_Header, m_XPixelCount, m_YPixelCount, m_XByteCount
        
    End If
    
    Exit Function
    
InternalVBError:
    InternalError "Step4_UnfilterIDAT", "internal VB error #" & Err.Number & ": " & Err.Description
    Step4_UnfilterIDAT = png_Failure

End Function

'Once all relevant chunks have been decompressed, we next need to un-filter each scanline into its original,
' unfiltered pixel representation.  This step operates only on bytes, which means it is completely independent
' of color-depth or bit-depth.
Friend Function Step5_ConstructImage(ByRef srcFile As String, ByRef dstDIB As pdDIB, ByRef dstImage As pdImage) As PD_PNGResult

    On Error GoTo InternalVBError
    
    Step5_ConstructImage = png_Success
    
    'Failsafe check(s)
    If Strings.StringsNotEqual(m_SourceFilename, srcFile, False) Then
        InternalError "Step5_ConstructImage", "filename has changed since original validation!"
        Step5_ConstructImage = png_Failure
    End If
    
    Dim idatIndex As Long
    idatIndex = Me.GetIndexOfChunk("IDAT")
    If (idatIndex < 0) Or (idatIndex >= m_NumOfChunks) Then
        InternalError "Step5_ConstructImage", "no IDAT chunk!"
        Step5_ConstructImage = png_Failure
    End If
    
    If (Step5_ConstructImage >= png_Failure) Then Exit Function
    
    'If the image contains a palette (PLTE chunk), now's the time to grab it!
    Dim hasPalette As Boolean, plteIndex As Long, srcColors() As RGBQuad, numColors As Long
    plteIndex = Me.GetIndexOfChunk("PLTE")
    hasPalette = (plteIndex >= 0)
    If hasPalette Then hasPalette = m_Chunks(plteIndex).GetPalette(srcColors, numColors, m_Warnings)
    
    'If a tRNS chunk exists, retrieve its information, too.  (Note that we only check for tRNS under certain
    ' color type combinations; in particular, it is forbidden if the image contains a full alpha channel.)
    Dim hasTransparency As Boolean, trnsIndex As Long, trnsRed As Long, trnsGreen As Long, trnsBlue As Long
    trnsRed = -1: trnsGreen = -1: trnsBlue = -1
    
    If (m_Header.ColorType = png_Indexed) Or (m_Header.ColorType = png_Greyscale) Or (m_Header.ColorType = png_Truecolor) Then
        trnsIndex = Me.GetIndexOfChunk("tRNS")
        hasTransparency = (trnsIndex >= 0)
        If hasTransparency Then hasTransparency = m_Chunks(trnsIndex).GetTRNSData(srcColors, trnsRed, trnsGreen, trnsBlue, m_Header, m_Warnings)
    End If
    
    'Because the tRNS chunk may update the contents of the palette (because it can store a table of
    ' transparency values), only *now* can we notify the parent image of the original palette contents.
    If hasPalette Then dstImage.SetOriginalPalette srcColors, numColors
    
    'If the PNG contained ICC profile data, retrieve it now; otherwise, assume sRGB.  (Note that a
    ' transform may be generated either way, as LittleCMS can be used for fast swizzling.)
    Dim srcProfile As pdLCMSProfile, dstProfile As pdLCMSProfile, dstTransform As pdLCMSTransform, srcImgProfileGood As Boolean
    PrepICCData srcProfile, dstProfile, dstTransform, srcImgProfileGood
    If srcImgProfileGood Then
    
        'Add a copy of the profile to PD's central cache, and associate the target image with that profile
        Dim tmpProfile As pdICCProfile
        Set tmpProfile = New pdICCProfile
        If tmpProfile.LoadICCFromPtr(m_Chunks(Me.GetIndexOfChunk("iCCP")).BorrowData().GetStreamSize(), m_Chunks(Me.GetIndexOfChunk("iCCP")).BorrowData().Peek_PointerOnly(0)) Then
        
            Dim profHash As String
            profHash = ColorManagement.AddProfileToCache(tmpProfile, True, False, False)
            dstImage.SetColorProfile_Original profHash
            
            'IMPORTANT NOTE: at preset, the destination DIB - by the time we're done with it - will have been
            ' hard-converted to sRGB, so we don't want to associate the destination DIB with it's source profile.
            ' Instead, note that it is currently sRGB.
            profHash = ColorManagement.GetSRGBProfileHash()
            dstDIB.SetColorProfileHash profHash
            dstDIB.SetColorManagementState cms_ProfileConverted
            
        End If
        
    End If
    
    'We're close enough to final processing that we can finally construct the DIB at its target size.
    ' (Note that - regardless of PNG color depth - we *always* create the DIB as 32-bpp RGBA.)
    Dim okToProceed As Boolean
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    okToProceed = dstDIB.CreateBlank(m_Header.Width, m_Header.Height, 32, vbWhite, 255)
    
    'Ensure that all our previous steps completed successfully, and that we have a valid destination image to
    ' place our decoded data into.
    If okToProceed Then
        
        'To simplify our code, we now want to upsample any color-depths below 8-bpp.  (They present a
        ' unique challenge within VB, since we lack the bit-shift operators necessary to handle them
        ' elegantly.)  Note that the upsampler will reset our xStride value to match 8-bpp data
        ' it produces.  (Note that we also pass the special case of 8-bpp *interlaced* data to this
        ' function; that's because the function performs de-interlacing of the 8-bpp data produced
        ' by the upsampler, and it spares us a bit of code to reuse that functionality for 8-bpp
        ' interlaced images.
        If (m_Header.BitDepth < 8) Or ((m_Header.BitsPerPixel = 8) And m_Header.Interlaced) Then UpsampleLowBitDepthData
        
        '8-bpp interlaced images have now been handled correctly.  Interlaced images at higher
        ' bit-depths (16, 24, 32, 48, 64) need to be dealt with using a separate function.
        If ((m_Header.BitsPerPixel > 8) And m_Header.Interlaced) Then DeInterlaceNonPaletteData
        
        'If the image is 8-bpp grayscale, we now want to construct an artifical palette.  (This lets us process it
        ' in an identical manner to indexed color images.)
        If (m_Header.ColorType = png_Greyscale) Then ConstructGrayPalette srcColors, numColors, trnsRed
        
        'Certain bit-depths require us to handle transparency manually (but note that 8-bpp grayscale has already
        ' been covered by the ConstructGrayPalette() function, above!)
        Dim trnsHandlingRequired As Boolean, trnsCheck As Long
        trnsHandlingRequired = (m_Header.ColorType = png_Truecolor) And (trnsRed >= 0) And (trnsGreen >= 0) And (trnsBlue >= 0)
        If trnsHandlingRequired And (m_Header.BitDepth = 8) Then trnsCheck = RGB(trnsRed, trnsGreen, trnsBlue)
        
        'We can now assign actual pixel bits!
        Dim dstBytes() As Byte, dstSA As SafeArray2D
        dstDIB.WrapArrayAroundDIB dstBytes, dstSA
        
        'To improve performance, we use unsafe direct access to the chunk's contents
        Dim srcBytes() As Byte, srcSA As SafeArray1D
        m_Chunks(idatIndex).BorrowData.WrapArrayAroundMemoryStream srcBytes, srcSA, 0
        
        'Prep all the required loop values
        Dim imgWidth As Long, imgHeight As Long, xFinal As Long, yFinal As Long
        Dim x As Long, y As Long, xDIB As Long, srcOffset As Long
        imgWidth = m_Header.Width
        imgHeight = m_Header.Height
        xFinal = imgWidth - 1
        yFinal = imgHeight - 1
        
        Dim xStride As Long
        xStride = m_XByteCount(0)
        
        Dim r As Long, g As Long, b As Long, a As Long, palIndex As Long
            
        Select Case m_Header.ColorType
        
            Case png_Greyscale
                
                Dim normalCopy As Boolean: normalCopy = True
                
                'Grayscale present a unique challenge if an ICC profile is present.  ICC profiles operate on
                ' specific color spaces, so a grayscale image must have a grayscale profile attached, and that
                ' grayscale profile must be used on grayscale data.  This means that once we convert the data
                ' to 32-bpp RGBA, the color profile *cannot physically be used* on the resulting image - so we
                ' need to perform the ICC correction on the *original* grayscale data.
                If (Not dstTransform Is Nothing) Then
                    
                    'We now have two possibilities:
                    ' 1) The image is color-managed, but doesn't use a tRNS chunk for wacky transparency.
                    '    This is fine and easy to handle.
                    ' 2) The is color-managed, *and* it uses a tRNS chunk.  This is a rare possibility,
                    '    but it creates some real headaches, as we need to maintain a copy of the original
                    '    IDAT stream - because ICC-converted data may have changed the color identified by
                    '    the tRNS chunk as "transparent".
                    '
                    'Case (1) comes first.  Note that it just performs an in-place transform; we still use
                    ' the normal greyscale conversion code afterward.
                    If (trnsRed < 0) Then
                    
                        For y = 0 To yFinal
                            srcOffset = y * xStride + 1
                            dstTransform.ApplyTransformToScanline VarPtr(srcBytes(srcOffset)), VarPtr(srcBytes(srcOffset)), imgWidth
                        Next y
                        normalCopy = True
                        If srcImgProfileGood Then dstDIB.SetColorManagementState cms_ProfileConverted
                    
                    'This is horrible case (2).  Let's do a first pass through the image, marking alpha data,
                    ' after which we'll perform the ICC transform to the source data, then do a second pass
                    ' to assign RGB bytes.
                    Else
                        
                        'Assign alpha bytes, if any
                        For y = 0 To yFinal
                            srcOffset = y * xStride + 1
                        For x = 0 To xFinal
                            If (m_Header.BitDepth = 8) Then
                                g = srcBytes(srcOffset + x)
                                If (g = trnsRed) Then dstBytes(x * 4 + 3, y) = 0
                            Else
                                g = srcBytes(srcOffset + x * 2 + 1)
                                g = g * 256 + srcBytes(srcOffset + x * 2)
                                If (g = trnsRed) Then dstBytes(x * 4 + 3, y) = 0
                            End If
                        Next x
                        Next y
                        
                        'Perform ICC correction on the original source bytes
                        For y = 0 To yFinal
                            srcOffset = y * xStride + 1
                            dstTransform.ApplyTransformToScanline VarPtr(srcBytes(srcOffset)), VarPtr(srcBytes(srcOffset)), imgWidth
                        Next y
                        
                        'Copy ICC-transformed grey bytes into the final RGBA container
                        For y = 0 To yFinal
                            srcOffset = y * xStride + 1
                        For x = 0 To xFinal
                            If (m_Header.BitDepth = 8) Then
                                g = srcBytes(srcOffset + x)
                            Else
                                g = srcBytes(srcOffset + x * 2 + 1)
                                g = g * 256 + srcBytes(srcOffset + x * 2)
                                g = g \ 256
                            End If
                            dstBytes(x * 4, y) = g
                            dstBytes(x * 4 + 1, y) = g
                            dstBytes(x * 4 + 2, y) = g
                        Next x
                        Next y
                        
                        normalCopy = False
                        If srcImgProfileGood Then dstDIB.SetColorManagementState cms_ProfileConverted
                        
                    End If
                
                End If
                
                If normalCopy Then
                
                    'With any ICC handling successfully covered, we can now proceed with normal loading.
                    For y = 0 To yFinal
                    
                        srcOffset = y * xStride + 1
                        
                        'For performance reasons, separate out 8- and 16-bit handling
                        If (m_Header.BitDepth <= 8) Then
                            For x = 0 To xFinal
                                palIndex = srcBytes(srcOffset + x)
                                xDIB = x * 4
                                With srcColors(palIndex)
                                    dstBytes(xDIB, y) = .Blue
                                    dstBytes(xDIB + 1, y) = .Green
                                    dstBytes(xDIB + 2, y) = .Red
                                    dstBytes(xDIB + 3, y) = .Alpha
                                End With
                            Next x
                        Else
                            For x = 0 To xFinal
                                g = srcBytes(srcOffset + x * 2)
                                g = g * 256 + srcBytes(srcOffset + x * 2 + 1)
                                xDIB = x * 4
                                If (g = trnsRed) Then dstBytes(xDIB + 3, y) = 0
                                g = g \ 256
                                dstBytes(x * 4, y) = g
                                dstBytes(x * 4 + 1, y) = g
                                dstBytes(x * 4 + 2, y) = g
                            Next x
                        End If
                        
                    Next y
                    
                End If
                
                okToProceed = True
                
            Case png_GreyscaleAlpha
                
                'If an ICC profile is active, apply it in-place to the source data
                If (Not dstTransform Is Nothing) Then
                    For y = 0 To yFinal
                        srcOffset = y * xStride + 1
                        dstTransform.ApplyTransformToScanline VarPtr(srcBytes(srcOffset)), VarPtr(srcBytes(srcOffset)), imgWidth
                    Next y
                    If srcImgProfileGood Then dstDIB.SetColorManagementState cms_ProfileConverted
                End If
                
                'Copy ICC-transformed grey bytes into the final RGBA container
                For y = 0 To yFinal
                    srcOffset = y * xStride + 1
                For x = 0 To xFinal
                    If (m_Header.BitDepth = 8) Then
                        g = srcBytes(srcOffset + x * 2)
                        a = srcBytes(srcOffset + x * 2 + 1)
                    Else
                        g = srcBytes(srcOffset + x * 4)
                        g = g * 256 + srcBytes(srcOffset + x * 4 + 1)
                        g = g \ 256
                        a = srcBytes(srcOffset + x * 4 + 2)
                        a = a * 256 + srcBytes(srcOffset + x * 4 + 3)
                        a = a \ 256
                    End If
                    dstBytes(x * 4, y) = g
                    dstBytes(x * 4 + 1, y) = g
                    dstBytes(x * 4 + 2, y) = g
                    dstBytes(x * 4 + 3, y) = a
                Next x
                Next y
            
                okToProceed = True
                
            Case png_Indexed
                
                For y = 0 To yFinal
                    srcOffset = y * xStride + 1
                For x = 0 To xFinal
                    
                    palIndex = srcBytes(srcOffset + x)
                    xDIB = x * 4
                    
                    With srcColors(palIndex)
                        dstBytes(xDIB, y) = .Blue
                        dstBytes(xDIB + 1, y) = .Green
                        dstBytes(xDIB + 2, y) = .Red
                        dstBytes(xDIB + 3, y) = .Alpha
                    End With
                    
                Next x
                Next y
                
                'With the indexed image successfully constructed as a full 32-bit RGBA image, we now need to handle
                ' the ICC profile (if any).  We only do this if the image contained an embedded profile; otherwise,
                ' the pixel data is already good to go.
                If (Not dstTransform Is Nothing) Then
                    For y = 0 To yFinal
                        dstTransform.ApplyTransformToScanline VarPtr(dstBytes(0, y)), VarPtr(dstBytes(0, y)), imgWidth
                    Next y
                    If srcImgProfileGood Then dstDIB.SetColorManagementState cms_ProfileConverted
                End If
                
                okToProceed = True
            
            Case png_Truecolor
                
                'If LittleCMS is available, use it for faster byte swizzling
                If (Not dstTransform Is Nothing) Then
                    
                    For y = 0 To yFinal
                    
                        srcOffset = y * xStride + 1
                        
                        'Use LittleCMS to perform swizzling for us
                        dstTransform.ApplyTransformToScanline VarPtr(srcBytes(srcOffset)), VarPtr(dstBytes(0, y)), imgWidth
                        
                        'If a tRNS chunk specified a transparent color, we need to find and detect transparent colors -
                        ' but importantly, we need to scan the *original* bytes, not the transformed ones (as the
                        ' color may have changed!)
                        If trnsHandlingRequired Then
                            
                            If (m_Header.BitDepth = 8) Then
                                For x = 0 To xFinal
                                    r = srcBytes(srcOffset + x * 3)
                                    g = srcBytes(srcOffset + x * 3 + 1)
                                    b = srcBytes(srcOffset + x * 3 + 2)
                                    If (RGB(r, g, b) = trnsCheck) Then dstBytes(x * 4 + 3, y) = 0
                                Next x
                            Else
                                Dim newR As Long, newG As Long, newB As Long
                                For x = 0 To xFinal
                                    newR = srcBytes(srcOffset + x * 6)
                                    newR = (newR * 256) + srcBytes(srcOffset + x * 6 + 1)
                                    newG = srcBytes(srcOffset + x * 6 + 2)
                                    newG = (newG * 256) + srcBytes(srcOffset + x * 6 + 3)
                                    newB = srcBytes(srcOffset + x * 6 + 4)
                                    newB = (newB * 256) + srcBytes(srcOffset + x * 6 + 5)
                                    If (newR = trnsRed) And (newG = trnsGreen) And (newB = trnsBlue) Then dstBytes(x * 4 + 3, y) = 0
                                Next x
                            End If
                            
                        '/End tRNS handling
                        End If
                    
                    Next y
                    
                    'Because color-management has been handled, we need to mark the destination DIB accordingly.
                    ' (Note that we only do this if the color profile was embedded in the source image; if the
                    ' image was untagged, we leave it as such.)
                    If srcImgProfileGood Then dstDIB.SetColorManagementState cms_ProfileConverted
                    
                'Failsafe pure-VB transform, in the unlikely event LittleCMS is unavailable
                Else
                
                    For y = 0 To yFinal
                        srcOffset = y * xStride + 1
                    For x = 0 To xFinal
                        
                        'Byte order has to be swapped to BGRA
                        If (m_Header.BitDepth = 8) Then
                            r = srcBytes(srcOffset + x * 3)
                            g = srcBytes(srcOffset + x * 3 + 1)
                            b = srcBytes(srcOffset + x * 3 + 2)
                            If trnsHandlingRequired Then If (RGB(r, g, b) = trnsCheck) Then dstBytes(x * 4 + 3, y) = 0
                        Else
                            r = srcBytes(srcOffset + x * 6)
                            r = r * 256 + srcBytes(srcOffset + x * 6 + 1)
                            g = srcBytes(srcOffset + x * 6 + 2)
                            g = g * 256 + srcBytes(srcOffset + x * 6 + 3)
                            b = srcBytes(srcOffset + x * 6 + 4)
                            b = b * 256 + srcBytes(srcOffset + x * 6 + 5)
                            If trnsHandlingRequired Then If (r = trnsRed) And (g = trnsGreen) And (b = trnsBlue) Then dstBytes(x * 4 + 3, y) = 0
                        End If
                        
                        xDIB = x * 4
                        dstBytes(xDIB, y) = b \ 256
                        dstBytes(xDIB + 1, y) = g \ 256
                        dstBytes(xDIB + 2, y) = r \ 256
                        
                    Next x
                    Next y
                    
                End If
                
                okToProceed = True
            
            Case png_TruecolorAlpha
            
                'If LittleCMS is available, use it for faster byte swizzling
                If (Not dstTransform Is Nothing) Then
                    
                    For y = 0 To yFinal
                    
                        srcOffset = y * xStride + 1
                        dstTransform.ApplyTransformToScanline VarPtr(srcBytes(srcOffset)), VarPtr(dstBytes(0, y)), imgWidth
                        
                        'On 16-bit data, we have to do some housekeeping to compensate for a LittleCMS bug.
                        ' LittleCMS normally swaps endianness just fine - *except* for alpha channels, as they
                        ' are generally copied over as-is.  When reducing 16-bit to 8-bit data, this results in
                        ' the least-significant bit being copied over (which we obviously don't want).  As such,
                        ' we need to manually track back through the original data and copy over MSBs.
                        '
                        'NOTE!  I submitted this issue to Marti (https://github.com/mm2/Little-CMS/issues/159,
                        ' https://github.com/mm2/Little-CMS/issues/162) and with his help, I've manually patched
                        ' up PD's lcms2.dll.  Manual correction is no longer required - yay!
                        '
                        'If (m_Header.BitDepth = 16) Then
                        '    For x = 0 To xFinal
                        '        dstBytes(x * 4 + 3, y) = srcBytes(srcOffset + x * 8 + 6)
                        '    Next x
                        'End If
                        
                    Next y
                    
                    'Because color-management has been handled, we need to mark the destination DIB accordingly.
                    ' (Note that we only do this if the color profile was embedded in the source image; if the
                    ' image was untagged, we leave it as such.)
                    If srcImgProfileGood Then dstDIB.SetColorManagementState cms_ProfileConverted
                    
                'Failsafe pure-VB transform, in the unlikely event LittleCMS is unavailable
                Else
                    For y = 0 To yFinal
                        srcOffset = y * xStride + 1
                    For x = 0 To xFinal
                        
                        xDIB = x * 4
                        If (m_Header.BitDepth = 8) Then
                            dstBytes(xDIB, y) = srcBytes(srcOffset + x * 4 + 2)
                            dstBytes(xDIB + 1, y) = srcBytes(srcOffset + x * 4 + 1)
                            dstBytes(xDIB + 2, y) = srcBytes(srcOffset + x * 4)
                            dstBytes(xDIB + 3, y) = srcBytes(srcOffset + x * 4 + 3)
                        Else
                            dstBytes(xDIB, y) = srcBytes(srcOffset + x * 8 + 4)
                            dstBytes(xDIB + 1, y) = srcBytes(srcOffset + x * 8 + 2)
                            dstBytes(xDIB + 2, y) = srcBytes(srcOffset + x * 8)
                            dstBytes(xDIB + 3, y) = srcBytes(srcOffset + x * 8 + 6)
                        End If
                        
                    Next x
                    Next y
                End If
                
                okToProceed = True
            
        End Select
        
        'Release all unsafe array wrappers
        dstDIB.UnwrapArrayFromDIB dstBytes
        m_Chunks(idatIndex).BorrowData.UnwrapArrayFromMemoryStream srcBytes
        
        If okToProceed Then Step5_ConstructImage = png_Success Else Step5_ConstructImage = png_Failure
    
    End If
    
    'If we constructed any ICC profile or transform handles, make sure to free them before exiting
    Set srcProfile = Nothing
    Set dstProfile = Nothing
    Set dstTransform = Nothing
    
    Exit Function
    
InternalVBError:
    InternalError "Step5_ConstructImage", "internal VB error #" & Err.Number & ": " & Err.Description
    Step5_ConstructImage = png_Failure
End Function

'Once image data has been fully constructed, the last thing we need to do is check for any esoteric chunks that affect
' the final image presentation (like gamma or chromaticity conversions).  For 99% of images, this step won't do a thing.
Friend Function Step6_PostProcessing(ByRef srcFile As String, ByRef dstDIB As pdDIB, ByRef dstImage As pdImage) As PD_PNGResult

    On Error GoTo InternalVBError
    
    Step6_PostProcessing = png_Success
    
    'Failsafe check(s)
    If Strings.StringsNotEqual(m_SourceFilename, srcFile, False) Then
        InternalError "Step6_PostProcessing", "filename has changed since original validation!"
        Step6_PostProcessing = png_Failure
    End If
    
    'Look for DPI information
    Dim physIndex As Long
    physIndex = Me.GetIndexOfChunk("pHYs")
    If (physIndex >= 0) Then
        
        Dim physWidth As Long, physHeight As Long, physUnit As Byte
        m_Chunks(physIndex).BorrowData().SetPosition 0, FILE_BEGIN
        physWidth = m_Chunks(physIndex).BorrowData().ReadLong_BE()
        physHeight = m_Chunks(physIndex).BorrowData().ReadLong_BE()
        physUnit = m_Chunks(physIndex).BorrowData().ReadByte()
        
        'The only supported unit in the spec is "1", for "meters"
        If (physUnit = 1) Then
            Dim dpiWidth As Double, dpiHeight As Double
            dpiWidth = (CDbl(physWidth) / 100#) * 2.54
            dpiHeight = (CDbl(physHeight) / 100#) * 2.54
            If (Not dstImage Is Nothing) Then dstImage.SetDPI dpiWidth, dpiHeight, False
        End If
        
    End If
    
    'PNG presents a hierarchy of lighting and color adjustment chunks.  If present, they are to be dealt with
    ' in this order (and once one is dealt with, subsequent ones should be ignored, to avoid double-corrections):
    
    '1) iCCP
    '2) sRGB
    '3) gAMA and/or cHRM (equal priority, and should be handled together if possible)
    
    'ICC profiles or sRGB flags, if present, are dealt with during a previous decoding step (for performance reasons).
    ' As such, we don't need to deal with them here - and in fact, we can bail if they were handled previously,
    ' as they take precedence over all other corrective functions.
    If (dstDIB.GetColorManagementState <> cms_NoManagement) Then
        Step6_PostProcessing = png_Success
        Exit Function
    End If
    
    'As of 7.2, PD assumes sRGB for untagged files, so the sRGB chunk doesn't mean anything special to us.
    ' If it exists, however, it will keep us from manually adjusting gamma and chromaticity to match.
    ' (Gamma in particular is important, as sRGB doesn't use a standard exponent curve, so you can't use a
    ' gAMA chunk to replicate it correctly.)
    If (Me.GetIndexOfChunk("sRGB") >= 0) Then
        Step6_PostProcessing = png_Success
        Exit Function
    End If
    
    'Next, look for gAMA and cHRM chunks, if any
    Dim gammaIndex As Long, chrmIndex As Long
    gammaIndex = Me.GetIndexOfChunk("gAMA")
    chrmIndex = Me.GetIndexOfChunk("cHRM")
    
    Dim gammaOK As Boolean, chrmOK As Boolean
    
    'Handle gamma; this allows us to produce the same results as Firefox, Chrome, Edge, Internet Explorer,
    ' and (surprisingly) Windows Photo Viewer on Windows 10.  (For whatever reason, Safari does *not* handle
    ' gAMA chunks; given that PD is Windows-only, I haven't provided a toggle to change this behavior,
    ' at present.)  Note that photo editing software may handle gamma differently - I have *not* tested them
    ' thoroughly, and if they don't manage gamma correctly, I may need to revisit this design entirely.
    If (gammaIndex >= 0) Then
        
        Dim gammaInt As Long
        gammaInt = m_Chunks(gammaIndex).GetGammaData(m_Warnings)
        
        Dim gammaFloat As Double
        gammaFloat = CDbl(gammaInt) / 100000#
        
        'If the returned gamma value is valid, apply run-time gamma correction.
        gammaOK = (gammaFloat > 0#)
        If gammaOK Then gammaFloat = (1# / gammaFloat)
        
    End If
    
    'Ideally, gamma chunks should be accompanied by chromaticity data, to help us know the expected
    ' viewing conditions of the image.  If chromaticity *doesn't* exist, we'll just apply a bare
    ' gamma correction to the raw PNG pixel data.
    If (chrmIndex >= 0) Then
        
        Dim chrmData() As Long
        If m_Chunks(chrmIndex).GetChromaticityData(chrmData, m_Warnings) Then
            
            'Convert the chromaticity values to floating-point
            Dim wpValues() As Double, rgbValues() As Double
            ReDim wpValues(0 To 3) As Double
            ReDim rgbValues(0 To 9) As Double
            wpValues(0) = CDbl(chrmData(0)) / 100000#
            wpValues(1) = CDbl(chrmData(1)) / 100000#
            wpValues(2) = 1#
            
            rgbValues(0) = CDbl(chrmData(2)) / 100000#
            rgbValues(1) = CDbl(chrmData(3)) / 100000#
            rgbValues(2) = 1#
            rgbValues(3) = CDbl(chrmData(4)) / 100000#
            rgbValues(4) = CDbl(chrmData(5)) / 100000#
            rgbValues(5) = 1#
            rgbValues(6) = CDbl(chrmData(6)) / 100000#
            rgbValues(7) = CDbl(chrmData(7)) / 100000#
            rgbValues(8) = 1#
            
            If (Not gammaOK) Then gammaFloat = 1#
            
            'Our next goal is to assemble an ICC profile that matches the chromaticity and/or gamma data.
            ' With this, we can use our normal color-management pipeline to color-correct this DIB.
            Dim srcProfile As pdLCMSProfile
            Set srcProfile = New pdLCMSProfile
        
            'Make sure the chromaticity data passes internal LCMS validation
            chrmOK = srcProfile.CreateCustomRGBProfile(VarPtr(wpValues(0)), VarPtr(rgbValues(0)), gammaFloat)
            If chrmOK Then
                
                Dim dstProfile As pdLCMSProfile
                Set dstProfile = New pdLCMSProfile
                dstProfile.CreateSRGBProfile
                
                'Apply the profile in-place to the current RGB data
                Dim dstTransform As pdLCMSTransform
                Set dstTransform = New pdLCMSTransform
                If dstTransform.CreateInPlaceTransformForDIB(dstDIB, srcProfile, dstProfile, INTENT_PERCEPTUAL, cmsFLAGS_COPY_ALPHA) Then
                    
                    dstTransform.ApplyTransformToPDDib dstDIB
                    Set dstTransform = Nothing
                    
                    'If the transform was successful, let's embed this ICC profile into the destination image.
                    ' This ensures reasonable persistence for the data described by the original cHRM chunk.
                    Dim profHash As String
                    profHash = ColorManagement.AddLCMSProfileToCache(srcProfile, True)
                    dstDIB.SetColorProfileHash profHash
                    dstDIB.SetColorManagementState cms_ProfileConverted
                    
                    dstImage.SetColorProfile_Original profHash
                    
                End If
                
                'Free the destination LCMS profile before exiting
                Set dstProfile = Nothing
                
                'Mark both chrm and gamma as *not* OK, so that subsequent checks don't process them
                chrmOK = False
                gammaOK = False
                
            End If
            
            'Free the source profile handle, if any
            Set srcProfile = Nothing
            
        End If
        
    End If
    
    'If the cHRM chunk doesn't exist (or is invalid), but a gAMA chunk exists (and is valid),
    ' apply basic gamma correction before exiting.
    If gammaOK And (Not chrmOK) Then Filters_Layers.FastGammaDIB dstDIB, 2.2 / gammaFloat
    
    Exit Function
    
InternalVBError:
    InternalError "Step6_PostProcessing", "internal VB error #" & Err.Number & ": " & Err.Description
    Step6_PostProcessing = png_Failure
End Function

'For non-paletted images (e.g. 16-bpp or more), call this function to convert an interlaced IDAT stream
' to a de-interlaced one.  (8-bpp are handled separately, using a function optimized specifically for them.)
Private Sub DeInterlaceNonPaletteData()

    'This function only handles deinterlacing for bit-depths higher than 8-bpp
    If (m_Header.BitsPerPixel < 8) Or (Not m_Header.Interlaced) Then Exit Sub
    
    Dim idatIndex As Long
    idatIndex = Me.GetIndexOfChunk("IDAT")
    If (idatIndex < 0) Then Exit Sub
    
    'The goal of this function is to separate the original data stream into its component images
    ' (one for each interlacing pass), then reassemble them into a normal, de-interlaced stream
    ' that subsequent assembly functions can use without modification.  (This is not the most efficient
    ' way to handle interlaced images, but it saves us a *ton* of code.)
    Dim xFinal As Long, yFinal As Long
    
    'Grab a reference to the primary, interlaced IDAT stream
    m_Chunks(idatIndex).BorrowData.SetPosition 0, FILE_BEGIN
    
    'Because we can't lean on normal image dimensions (and stride), we'll instead calculate
    ' new dimensions and stride on-the-fly for each interlacing pass.
    Dim tmpChunks() As pdPNGChunk
    SeparateInterlacedIDATs tmpChunks, False
        
    'We now have seven separate "fake" IDAT chunks, one for each interlacing pass.  Next we must
    ' reassemble these into a single, contiguous, non-interlaced IDAT chunk; we'll copy this result
    ' over our original IDAT chunk, which will allow subsequent rendering functions to operate
    ' without consideration for interlacing.
    Dim finalStream As pdStream
    Set finalStream = New pdStream
    
    'Calculate a necessary final size for the stream.  This varies by bit-depth, obviously:
    ' n-bytes per pixel, plus 1 extra byte per scanline for the silly PNG "filter" byte indicators
    ' (which we don't need, as pixel data has already been unfiltered, but the render functions
    ' expect the bytes to be there).
    Dim BytesPerPixel As Long, bppLoop As Long
    BytesPerPixel = (m_Header.BitsPerPixel \ 8)
    bppLoop = BytesPerPixel - 1
    
    Dim newIdatSize As Long
    newIdatSize = (m_Header.Width * m_Header.Height * BytesPerPixel) + m_Header.Height
    finalStream.StartStream PD_SM_MemoryBacked, PD_SA_ReadWrite, , newIdatSize
    finalStream.SetSizeExternally newIdatSize
    
    'For performance reasons, wrap a standard array around the final stream's contents.  (This is
    ' much faster than writing individual bytes by moving around the stream pointer manually,
    ' and using the safe WriteByte functions.)
    Dim dstBytes() As Byte, dstSA As SafeArray1D
    finalStream.WrapArrayAroundMemoryStream dstBytes, dstSA, 0
    
    'We'll also wrap an array around each chunk's stream, but this is handled inside the loop
    Dim srcBytes() As Byte, srcSA As SafeArray1D
    
    Dim x As Long, y As Long, j As Long
    Dim startOffsetX As Long, startOffsetY As Long, xLineOffset As Long, yLineOffset As Long
    Dim srcPxOffset As Long, newStride As Long
    
    'The destination image bounds and stride are constant for each pass (only the *source* ones
    ' will change).
    xFinal = m_Header.Width - 1
    yFinal = m_Header.Height - 1
    newStride = (m_Header.Width * BytesPerPixel) + 1
        
    'Loop through each iteration, placing pixels as we go!
    Dim i As Long
    For i = 0 To 6
        
        'Tiny images can have 0-length passes, if an interlacing pass doesn't contain any pixels
        ' any pixels that lie inside image bounds.
        If (m_XByteCount(i) > 0) Then
        
            'Each interlacing pass requires different starting points and x/y offsets between scans
            Select Case i
                Case 0
                    startOffsetX = 0: startOffsetY = 0: xLineOffset = 8: yLineOffset = 8
                Case 1
                    startOffsetX = 4: startOffsetY = 0: xLineOffset = 8: yLineOffset = 8
                Case 2
                    startOffsetX = 0: startOffsetY = 4: xLineOffset = 4: yLineOffset = 8
                Case 3
                    startOffsetX = 2: startOffsetY = 0: xLineOffset = 4: yLineOffset = 4
                Case 4
                    startOffsetX = 0: startOffsetY = 2: xLineOffset = 2: yLineOffset = 4
                Case 5
                    startOffsetX = 1: startOffsetY = 0: xLineOffset = 2: yLineOffset = 2
                Case 6
                    startOffsetX = 0: startOffsetY = 1: xLineOffset = 1: yLineOffset = 2
            End Select
            
            'For performance reasons, wrap a temporary array around this chunk's stream
            tmpChunks(i).BorrowData.WrapArrayAroundMemoryStream srcBytes, srcSA, 0
            srcPxOffset = 0
            
            'Place all pixels in this chunk
            For y = startOffsetY To yFinal Step yLineOffset
                
                'On each y-line (including the first one), advance the pixel offset by 1 to account
                ' for the filter byte that prepends each PNG scanline.
                srcPxOffset = srcPxOffset + 1
                
            For x = startOffsetX To xFinal Step xLineOffset
                For j = 0 To bppLoop
                    dstBytes(1 + (y * newStride) + (x * BytesPerPixel) + j) = srcBytes(srcPxOffset)
                    srcPxOffset = srcPxOffset + 1
                Next j
            Next x
            Next y
            
            'Unwrap the temporary source array and free the source chunk - it's no longer needed!
            tmpChunks(i).BorrowData.UnwrapArrayFromMemoryStream srcBytes
            Set tmpChunks(i) = Nothing
            
        End If
    
    Next i
    
    'Release our wrapper around the destination stream's bytes, and reset its internal pointer
    finalStream.UnwrapArrayFromMemoryStream dstBytes
    finalStream.SetPosition 0, FILE_BEGIN
    
    'Replace our original, interlaced IDAT stream with the new, non-interlaced one
    m_Chunks(idatIndex).SubmitNewDataStream finalStream
    
    'With the final, un-interlaced data stream assembled, we can update the pixel counts and stride tracker
    ' to match normal, un-interlaced data.
    ReDim m_XPixelCount(0) As Long
    m_XPixelCount(0) = m_Header.Width
    
    ReDim m_YPixelCount(0) As Long
    m_YPixelCount(0) = m_Header.Height
    
    ReDim m_XByteCount(0) As Long
    m_XByteCount(0) = (m_Header.Width * BytesPerPixel) + 1
    
End Sub

'LittleCMS is an integral part of the PNG load process.  If PNG images contain ICC profiles, we use LittleCMS
' to translate colors between spaces.  If PNG images do *not* contain ICC profiles, we can still use LittleCMS
' for much faster swizzling than VB6 alone can provide.  (As such, we still prep generic ICC profiles as part
' of the preparation stage.)
Private Sub PrepICCData(ByRef srcProfile As pdLCMSProfile, ByRef dstProfile As pdLCMSProfile, ByRef dstTransform As pdLCMSTransform, ByRef srcImgProfileGood As Boolean)
    
    Set srcProfile = New pdLCMSProfile
    
    'In most (?) cases, we want to use LittleCMS to color-manage incoming PNG data.  There are exceptions, however;
    ' indexed images without embedded profiles don't require management, for example.  As such, we can skip specialized
    ' ICC handling under some scenarios.
    Dim iccRequired As Boolean: iccRequired = True
    
    Dim iccIndex As Long
    iccIndex = Me.GetIndexOfChunk("iCCP")
    If (iccIndex > 0) Then
        srcImgProfileGood = srcProfile.CreateFromPointer(m_Chunks(iccIndex).BorrowData().Peek_PointerOnly(0), m_Chunks(iccIndex).BorrowData().GetStreamSize())
        
        'If the image does not contain color profile data, or if the embedded profile is invalid, assumed sRGB.
        ' (TODO: expose this to the user as a preference; untagged images may reside in a different space.)
        If (Not srcImgProfileGood) Then
            m_Warnings.AddString "WARNING!  ICC profile found in file, but LittleCMS couldn't validate the profile.  Color management will *not* be applied!"
            iccIndex = -1
        End If
        
    End If
    
    'ICC processing failed.  Create a generic source profile.
    If (iccIndex <= 0) Then
        If (m_Header.ColorType = png_Greyscale) Or (m_Header.ColorType = png_GreyscaleAlpha) Then
            srcProfile.CreateGenericGrayscaleProfile
        Else
            srcProfile.CreateSRGBProfile
        End If
    End If
    
    'At present, sRGB destination profiles are hard-coded; this is a work in progress for future releases
    Set dstProfile = New pdLCMSProfile
    If (m_Header.ColorType = png_Greyscale) Or (m_Header.ColorType = png_GreyscaleAlpha) Then
        dstProfile.CreateGenericGrayscaleProfile
    Else
        dstProfile.CreateSRGBProfile
    End If
    
    'Because LittleCMS supports fast swizzling (while VB6 does not), we use LittleCMS to accelerate certain types
    ' of pixel transformations - even if the source data isn't explicitly color-managed.
    Dim srcLcmsFormat As LCMS_PIXEL_FORMAT, dstLcmsFormat As LCMS_PIXEL_FORMAT
        
    'A quick note on how this works: in most cases, PD will first decode the PNG pixel data to our standard
    ' 32-bpp BGRA layout, *then* apply color-management to it.  However, some color formats don't work with
    ' this approach (e.g. grayscale data, which cannot be expanded to RGB as grayscale profiles don't work
    ' with RGB data; similarly, high bit-depth data needs to be transformed *prior* to downsampling), so we
    ' sometimes use LittleCMS to combine color-management and swizzling/depth conversion into one action.
    Select Case m_Header.ColorType
        
        'Greyscale and grayscale+alpha images only receive color management if the PNG contains an embedded profile.
        ' (Note also - importantly - that we leave the data in its original endianness; this lets us reuse some
        ' code elsewhere, for both ICC-corrected and uncorrected data.)
        Case png_Greyscale
            iccRequired = srcImgProfileGood
            If iccRequired Then
                If (m_Header.BitDepth = 16) Then srcLcmsFormat = TYPE_GRAY_16_SE Else srcLcmsFormat = TYPE_GRAY_8
                dstLcmsFormat = srcLcmsFormat
            End If
            
        Case png_GreyscaleAlpha
            iccRequired = srcImgProfileGood
            If iccRequired Then
                If (m_Header.BitDepth = 16) Then srcLcmsFormat = TYPE_GRAYA_16_SE Else srcLcmsFormat = TYPE_GRAYA_8
                dstLcmsFormat = srcLcmsFormat
            End If
        
        'Indexed images only receive color management if the PNG contains an embedded profile.
        Case png_Indexed
            iccRequired = srcImgProfileGood
            If iccRequired Then
                 srcLcmsFormat = TYPE_BGRA_8
                 dstLcmsFormat = TYPE_BGRA_8
            End If
        
        'RGB images can always make use of LittleCMS
        Case png_Truecolor
            iccRequired = True
            If (m_Header.BitDepth = 8) Then
                srcLcmsFormat = TYPE_RGB_8
                dstLcmsFormat = TYPE_BGRA_8
            Else
                srcLcmsFormat = TYPE_RGB_16_SE
                dstLcmsFormat = TYPE_BGRA_8
            End If
        
        'RGBA images can always make use of LittleCMS
        Case png_TruecolorAlpha
            iccRequired = True
            If (m_Header.BitDepth = 8) Then
                srcLcmsFormat = TYPE_RGBA_8
                dstLcmsFormat = TYPE_BGRA_8
            Else
                srcLcmsFormat = TYPE_RGBA_16_SE
                dstLcmsFormat = TYPE_BGRA_8
            End If
            
    End Select
    
    'If an ICC transform is relevant, create it now
    If (srcLcmsFormat <> 0) And (dstLcmsFormat <> 0) And iccRequired Then
        Set dstTransform = New pdLCMSTransform
        If (Not dstTransform.CreateTwoProfileTransform(srcProfile, dstProfile, srcLcmsFormat, dstLcmsFormat, INTENT_PERCEPTUAL, cmsFLAGS_COPY_ALPHA)) Then
            m_Warnings.AddString "WARNING!  pdPNG.Step5_ConstructImage failed to create a valid ICC transform.  Color management disabled for this image."
            Set dstTransform = Nothing
        End If
    End If

End Sub

'1, 2, and 4-bit-per-pixel data must be upsampled to be easily handled in VB.  We do this prior to constructing a
' final image, as it allows us to use identical rendering code for 1/2/4/8-bpp indexed images.
' (NOTE: this function handles both interlaced and non-interlaced images just fine.)
Private Sub UpsampleLowBitDepthData()
    
    'We only need to upsample bit-depths less than 8-bpp
    If (m_Header.BitsPerPixel > 8) Or ((m_Header.BitsPerPixel = 8) And (Not m_Header.Interlaced)) Then Exit Sub
    
    Dim idatIndex As Long
    idatIndex = Me.GetIndexOfChunk("IDAT")
    If (idatIndex < 0) Then Exit Sub
    
    'For non-interlaced images, this sub is just a thin wrapper.  For interlaced images, however,
    ' we first need to separate the original data stream into its component images (one for each
    ' interlacing pass), up-sample each one separately, then reassemble the results into a normal,
    ' un-interlaced stream that subsequent assembly functions can more easily use.
    Dim xFinal As Long, yFinal As Long, xStride As Long
    
    'Anyway, the gist is that non-interlaced images are *way* easier to deal with.
    If (Not m_Header.Interlaced) Then
    
        xFinal = m_Header.Width - 1
        yFinal = m_Header.Height - 1
        xStride = m_XByteCount(0)
        m_XByteCount(0) = UpsampleIDATChunk(m_Chunks(idatIndex), xFinal, yFinal, xStride)
    
    'Interlaced handling follows...
    Else
    
        'Interlaced handling is the same idea as non-interlaced handling - but instead of doing
        ' one upsample for the entire IDAT chunk, we're instead going to break the chunk into
        ' seven parts - one for each interlace pass - and upsample each one individually.
        ' After all passes are upsampled, we'll merge their results into a new, non-interlaced
        ' stream that subsequent functions can deal with like any normal PNG.
        m_Chunks(idatIndex).BorrowData.SetPosition 0, FILE_BEGIN
        
        'Because we can't lean on normal image dimensions (and stride), we'll instead calculate
        ' new dimensions and stride on-the-fly for each interlacing pass.
        Dim tmpChunks() As pdPNGChunk
        SeparateInterlacedIDATs tmpChunks, (m_Header.BitDepth < 8)
        
        'We now have seven separate "fake" IDAT chunks, one for each interlacing pass.  Each has been upsampled
        ' from its original bit-depth of 1/2/4-bpp to 8-bpp.  Before exiting, let's reassemble these into a
        ' single, contiguous, non-interlaced IDAT chunk; we'll copy this result over our original IDAT chunk,
        ' which will allow subsequent rendering functions to operate without consideration for interlacing.
        Dim finalStream As pdStream
        Set finalStream = New pdStream
        
        'Calculate a necessary final size for the stream.  This is easy for 8-bit images: 1-byte per pixel,
        ' plus 1 extra byte per scanline for the silly PNG "filter" byte indicators (which we don't need,
        ' as pixel data has already been unfiltered, but the render functions expect the bytes to be there).
        Dim newIdatSize As Long
        newIdatSize = (m_Header.Width + 1) * m_Header.Height
        finalStream.StartStream PD_SM_MemoryBacked, PD_SA_ReadWrite, , newIdatSize
        finalStream.SetSizeExternally newIdatSize
        
        'For performance reasons, wrap a standard array around the stream's contents.  (This is obviously
        ' faster than writing individual bytes by moving around the stream pointer manually.)
        Dim dstBytes() As Byte, dstSA As SafeArray1D
        finalStream.WrapArrayAroundMemoryStream dstBytes, dstSA, 0
        
        'We'll also wrap an array around each chunk's stream, but this is handled inside the loop
        Dim srcBytes() As Byte, srcSA As SafeArray1D
        
        Dim x As Long, y As Long
        Dim startOffsetX As Long, startOffsetY As Long, xLineOffset As Long, yLineOffset As Long
        Dim srcPxOffset As Long, newStride As Long
        
        'The destination image bounds and stride are constant for each pass (only the *source* ones
        ' will change).
        xFinal = m_Header.Width - 1
        yFinal = m_Header.Height - 1
        newStride = m_Header.Width + 1
        
        'Loop through each iteration, placing pixels as we go!
        Dim i As Long
        For i = 0 To 6
            
            'As before, we need to skip 0-length passes.
            If (m_XByteCount(i) > 0) Then
            
                'Each interlacing pass requires different starting points and x/y offsets between scans
                Select Case i
                    Case 0
                        startOffsetX = 0: startOffsetY = 0: xLineOffset = 8: yLineOffset = 8
                    Case 1
                        startOffsetX = 4: startOffsetY = 0: xLineOffset = 8: yLineOffset = 8
                    Case 2
                        startOffsetX = 0: startOffsetY = 4: xLineOffset = 4: yLineOffset = 8
                    Case 3
                        startOffsetX = 2: startOffsetY = 0: xLineOffset = 4: yLineOffset = 4
                    Case 4
                        startOffsetX = 0: startOffsetY = 2: xLineOffset = 2: yLineOffset = 4
                    Case 5
                        startOffsetX = 1: startOffsetY = 0: xLineOffset = 2: yLineOffset = 2
                    Case 6
                        startOffsetX = 0: startOffsetY = 1: xLineOffset = 1: yLineOffset = 2
                End Select
                
                'For performance reasons, wrap a temporary array around this chunk's stream
                tmpChunks(i).BorrowData.WrapArrayAroundMemoryStream srcBytes, srcSA, 0
                srcPxOffset = 0
                
                'Place all pixels in this chunk
                For y = startOffsetY To yFinal Step yLineOffset
                    
                    'On each y-line (including the first one), advance the pixel offset by 1 to account
                    ' for the filter byte that prepends each PNG scanline.
                    srcPxOffset = srcPxOffset + 1
                    
                For x = startOffsetX To xFinal Step xLineOffset
                    dstBytes(1 + (y * newStride) + x) = srcBytes(srcPxOffset)
                    srcPxOffset = srcPxOffset + 1
                Next x
                Next y
                
                'Unwrap the temporary source array and free the source chunk - it's no longer needed!
                tmpChunks(i).BorrowData.UnwrapArrayFromMemoryStream srcBytes
                Set tmpChunks(i) = Nothing
                
            End If
        
        Next i
        
        'Release our wrapper around the destination stream's bytes, and reset its internal pointer
        finalStream.UnwrapArrayFromMemoryStream dstBytes
        finalStream.SetPosition 0, FILE_BEGIN
        
        'Replace our original, interlaced IDAT stream with the new, non-interlaced one
        m_Chunks(idatIndex).SubmitNewDataStream finalStream
        
        'With the final, un-interlaced data stream assembled, we can update the pixel counts and stride tracker
        ' to match normal, un-interlaced data.
        ReDim m_XPixelCount(0) As Long
        m_XPixelCount(0) = m_Header.Width
        
        ReDim m_YPixelCount(0) As Long
        m_YPixelCount(0) = m_Header.Height
        
        ReDim m_XByteCount(0) As Long
        m_XByteCount(0) = m_Header.Width + 1
        
    End If
    
End Sub

'Interlaced images are a giant PITA.  To reuse as much code as possible, it's easiest to break them down into
' seven separate IDAT streams (one for each interlaced image), which we can then use with standard functions.
' If you call this function, that's exactly what happens: the current IDAT chunk will be converted to seven
' standalone IDAT chunks.  Importantly, this function will *not* free the original IDAT - that's up to you.
' To keep memory usage down, however, I strongly recommend freeing it, as it's completely redundant once these
' new chunks exist.
'
'As part of the separation process, you can optionally request to up-sample each chunk; PD does this to
' 1/2/4-bit data to make it *much* easier to work with during a subsequent de-interlacing pass.
Private Sub SeparateInterlacedIDATs(ByRef dstChunks() As pdPNGChunk, Optional ByVal alsoUpSampleBits As Boolean = True)
    
    Dim idatIndex As Long
    idatIndex = Me.GetIndexOfChunk("IDAT")
    If (idatIndex < 0) Then Exit Sub
    
    ReDim dstChunks(0 To 6) As pdPNGChunk
    Dim sizeOfPass As Long, sizeOfPassCheck As Long
    Dim xFinal As Long, yFinal As Long, xStride As Long
        
    'Per the name (Adam7), there are seven interlacing passes...
    Dim i As Long
    For i = 0 To 6
        
        'We can entirely skip 0-length passes.  (These only exist in tiny images, where the image
        ' dimensions are so small that some interlacing passes comprise 0 pixels.)
        If (m_XByteCount(i) > 0) Then
        
            'Figure out the total size of this pass, in bytes.
            sizeOfPass = m_XByteCount(i) * m_YPixelCount(i)
            
            'Initialize the new chunk and prep its internal buffer to the required size
            Set dstChunks(i) = New pdPNGChunk
            dstChunks(i).CreateChunk "IDAT", sizeOfPass
            
            'Use the size of this pass to "extract" the relevant portion of the original IDAT stream
            ' into the temporary chunk's internal buffer.  (Note that, like all .Read* functions, this call
            ' will increment the internal pointer of the m_Data stream.)  As a failsafe, we'll also
            ' double-check that the read size matches the size we requested - this ensures we never attempt
            ' to read past the end of the stream.
            sizeOfPassCheck = m_Chunks(idatIndex).BorrowData.ReadBytesToBarePointer(dstChunks(i).BorrowData.Peek_PointerOnly(0), sizeOfPass)
            If (sizeOfPassCheck = sizeOfPass) Then
                dstChunks(i).BorrowData.SetSizeExternally sizeOfPass
            Else
                dstChunks(i).BorrowData.SetSizeExternally sizeOfPassCheck
                m_Warnings.AddString "WARNING!  pdPNG.UpsampleLowBitDepthData failed to read an interlacing pass correctly (" & sizeOfPass & ", " & sizeOfPassCheck & ")"
            End If
            
            'With the temporary chunk assembled, we now need to calculate loop parameters for it.
            xFinal = m_XPixelCount(i) - 1
            yFinal = m_YPixelCount(i) - 1
            xStride = m_XByteCount(i)
            
            'If this image is > 8-bpp, upsample the chunk!  (We pass 8-bpp data to this function so we
            ' can take advantage of its 8-bpp deinterlacing.)
            If (m_Header.BitDepth < 8) And alsoUpSampleBits Then m_XByteCount(i) = UpsampleIDATChunk(dstChunks(i), xFinal, yFinal, xStride)
            
            'Leave the upsampled bytes in the temporary container; since we still need access to the
            ' original IDAT stream (for the next interlacing pass), we don't want to overwrite
            ' anything yet.
            
        End If
    
    Next i
        
End Sub

'Upsample the target IDAT chunk.  This function will return the *new* stride value for the chunk
' (x/y dimensions obviously don't change, but stride *does* as a result of the upsampling).  Note also
' that the target chunk will have a new data stream; the old one is IRRETRIEVABLY GONE - so if you want
' to process it further, back it up before calling this!
Private Function UpsampleIDATChunk(ByRef srcChunk As pdPNGChunk, ByVal xFinal As Long, ByVal yFinal As Long, ByVal xStride As Long) As Long
    
    Dim x As Long, y As Long, i As Long, curByte As Byte
    
    'We'll use pdStream to generate a new, 8-bpp version of the image; then we can use our standard 8-bpp
    ' decode functions on 1, 2, 4-bpp data.
    Dim tmpStream As pdStream
    Set tmpStream = New pdStream
    tmpStream.StartStream PD_SM_MemoryBacked, PD_SA_ReadWrite
    
    Dim srcStream As pdStream
    Set srcStream = srcChunk.BorrowData()
    srcStream.SetPosition 0, FILE_BEGIN
    
    Dim numPixelsProcessed As Long
    
    Select Case m_Header.BitDepth
        
        'There's not a performance-friendly way to mask flags in VB, so we just use a byte array for clarity
        Case 1
        
            Dim bitFlags() As Byte
            ReDim bitFlags(0 To 7) As Byte
            bitFlags(0) = 128
            bitFlags(1) = 64
            bitFlags(2) = 32
            bitFlags(3) = 16
            bitFlags(4) = 8
            bitFlags(5) = 4
            bitFlags(6) = 2
            bitFlags(7) = 1
                    
            'We want to focus on the bytes relevant to the current scanline only
            For y = 0 To yFinal
            
                'Move past the filter byte
                tmpStream.WriteByte srcStream.ReadByte()
                numPixelsProcessed = 0
                        
                'Read through (numOfBytesPerLine) entries, pushing values into the new stream as we go
                For x = 0 To xStride - 2
                
                    curByte = srcStream.ReadByte()
                    
                    'Parse each bit in turn
                    For i = 0 To 7
                        
                        'Ignore empty bytes at the end of each scanline
                        If (numPixelsProcessed <= xFinal) Then
                            If (bitFlags(i) = (curByte And bitFlags(i))) Then tmpStream.WriteByte 1 Else tmpStream.WriteByte 0
                            numPixelsProcessed = numPixelsProcessed + 1
                        End If
                        
                    Next i
                Next x
                
            Next y
                    
        Case 2
            
            Dim shiftTable() As Byte
            ReDim shiftTable(0 To 3) As Byte
            shiftTable(0) = 2 ^ 6
            shiftTable(1) = 2 ^ 4
            shiftTable(2) = 2 ^ 2
            shiftTable(3) = 1
                    
            'Unlike other methods, we want to extract the bytes relevant to the current scanline only
            For y = 0 To yFinal
            
                'Move past the filter byte (which has already been dealt with in a previous step)
                tmpStream.WriteByte srcStream.ReadByte()
                numPixelsProcessed = 0
                
                'Read through (numOfBytesPerLine) entries, pushing values into the new stream as we go
                For x = 0 To xStride - 2
                
                    curByte = srcStream.ReadByte()
                    
                    'Parse each two-bit pair in turn
                    For i = 0 To 3
                        
                        'Ignore empty bytes at the end of each scanline
                        If (numPixelsProcessed <= xFinal) Then
                            tmpStream.WriteByte (curByte \ shiftTable(i)) And &H3
                            numPixelsProcessed = numPixelsProcessed + 1
                        End If
                        
                    Next i
                Next x
            Next y
            
        Case 4
        
            'Unlike other methods, we want to extract the bytes relevant to the current scanline only
            For y = 0 To yFinal
            
                'Move past the filter byte (which has already been dealt with in a previous step)
                tmpStream.WriteByte srcStream.ReadByte()
                numPixelsProcessed = 0
                        
                'Read through (numOfBytesPerLine) entries, pushing values into the new stream as we go
                For x = 0 To xStride - 2
                
                    curByte = srcStream.ReadByte()
                    
                    'Parse each two-bit pair in turn, while ignoring empty bytes at the end of each scanline
                    tmpStream.WriteByte (curByte \ 16) And &HF
                    numPixelsProcessed = numPixelsProcessed + 1
                    
                    If (numPixelsProcessed <= xFinal) Then
                        tmpStream.WriteByte curByte And &HF
                        numPixelsProcessed = numPixelsProcessed + 1
                    End If
                    
                Next x
            Next y
    
    End Select
    
    'Overwrite the IDAT chunk's original data stream with the new one, and return the new xStride value.
    ' (This is just the number of horizontal pixels in the image - xFinal + 1 - plus an extra 1 to account
    ' for the prepended PNG filter value.)
    srcChunk.SubmitNewDataStream tmpStream
    UpsampleIDATChunk = xFinal + 2
    
End Function

'To simplify rendering, PD treates grayscale data identically to palette data; we just generate a grayscale palette
' on-the-fly (one matching the current color depth), then render accordingly.
Private Sub ConstructGrayPalette(ByRef srcPalette() As RGBQuad, ByRef srcPaletteSize As Long, ByVal srcTrnsRed As Long)

    Dim palOK As Boolean: palOK = True
    
    Select Case m_Header.BitDepth
        Case 1
            srcPaletteSize = 2
        Case 2
            srcPaletteSize = 4
        Case 4
            srcPaletteSize = 16
        Case 8
            srcPaletteSize = 256
        Case Else
            palOK = False
    End Select
            
    If palOK Then
    
        ReDim srcPalette(0 To srcPaletteSize - 1) As RGBQuad
        
        Dim tmpCalcInt As Long, tmpMax As Long
        tmpMax = srcPaletteSize - 1
        
        Dim i As Long
        For i = 0 To tmpMax
        
            tmpCalcInt = Int(CDbl(i / tmpMax) * 255# + 0.5)
            
            With srcPalette(i)
                .Red = tmpCalcInt
                .Green = tmpCalcInt
                .Blue = tmpCalcInt
                .Alpha = 255
            End With
            
        Next i
        
        'If tRNS is present, assign it now
        If (srcTrnsRed >= 0) And (srcTrnsRed <= 255) Then
            For i = 0 To tmpMax
                If (srcPalette(i).Red = srcTrnsRed) Then srcPalette(i).Alpha = 0
            Next i
        End If
        
    End If
            
End Sub

'TODO!  We can now assemble the image's underlying pixel data.  Note that
' the way we do this varies by color-mode and bit-depth.  If the image has an ICC profile, we want to leave data
' in its original format in some circumstances (e.g. 16-bit-per-channel images must remain high-bit-depth prior
' to ICC conversion), while in others (e.g. 1-bit images) there is no advantage to leaving pixels in their
' original state.  (In fact, we *must* convert them to another color-depth in order to ICC-correct them.)

'NOTE: this function only works if m_Header has been populated!  We need access to color-mode and bit-depth data
' in order to calculate interlacing pixel counts.
Private Function GetSizeInterlaced(ByRef dstX As Long, ByRef dstY As Long, ByVal intPass As Long)

    'Per the spec, each interlacing pass encodes the following pixels from each 8x8 block in the image:
    ' Offset: 0 1 2 3 4 5 6 7
    '------------------------
    '         1 6 4 6 2 6 4 6
    '         7 7 7 7 7 7 7 7
    '         5 6 5 6 5 6 5 6
    '         7 7 7 7 7 7 7 7
    '         3 6 4 6 3 6 4 6
    '         7 7 7 7 7 7 7 7
    '         5 6 5 6 5 6 5 6
    '         7 7 7 7 7 7 7 7
    
    'To make things a little simpler, we can break down the problem into two parts:
    ' 1) How many pixels \ 8 are in this line?  This uses a fixed calculation and is easy to generalize.
    ' 2) How many pixels % 8 are in this line?  These trailing pixels are a bigger pain to deal with,
    '    and while we could hardcode a table (based on x/y direction and interlacing pass), VB doesn't
    '    make this easy - so let's just do a bit of math on each one.  This has the added bonus of being
    '    much easier to visualize.
    
    'Also note that I'm deliberately explicit with my use of Int(), even though integer division is involved,
    ' to clarify the process.
    
    Select Case intPass
        
        'First pass is easily generalized
        Case 1
            dstX = Int((m_Header.Width + 7) \ 8)
            dstY = Int((m_Header.Height + 7) \ 8)
        
        'Second pass is a little weirder in the x-direction
        Case 2
            dstX = Int((m_Header.Width + 3) \ 8)
            dstY = Int((m_Header.Height + 7) \ 8)
        
        'Third pass introduces twice as many pixels in each scanline
        Case 3
            dstX = Int((m_Header.Width + 7) \ 8) + Int((m_Header.Width + 3) \ 8)
            dstY = Int((m_Header.Height + 3) \ 8)
        
        'Fourth pass introduces twice as many y-pixels in each 8x8 block
        Case 4
            dstX = Int((m_Header.Width + 5) \ 8) + Int((m_Header.Width + 1) \ 8)
            dstY = Int((m_Header.Height + 7) \ 8) + Int((m_Header.Height + 3) \ 8)
        
        'Fifth pass once again doubles the pixels in each scanline
        Case 5
            dstX = Int((m_Header.Width + 7) \ 8) + Int((m_Header.Width + 5) \ 8) + Int((m_Header.Width + 3) \ 8) + Int((m_Header.Width + 1) \ 8)
            dstY = Int((m_Header.Height + 5) \ 8) + Int((m_Header.Height + 1) \ 8)
        
        'Sixth pass now has 4 pixels in the x and y direction of each 8x8 block
        Case 6
            dstX = Int((m_Header.Width + 6) \ 8) + Int((m_Header.Width + 4) \ 8) + Int((m_Header.Width + 2) \ 8) + Int(m_Header.Width \ 8)
            dstY = Int((m_Header.Height + 7) \ 8) + Int((m_Header.Height + 5) \ 8) + Int((m_Header.Height + 3) \ 8) + Int((m_Header.Height + 1) \ 8)
        
        'Seventh (and final) pass has full pixels in the x-direction, and 4 pixels per 8x8 block in the y-direction
        Case 7
            dstX = m_Header.Width
            dstY = Int((m_Header.Height + 6) \ 8) + Int((m_Header.Height + 4) \ 8) + Int((m_Header.Height + 2) \ 8) + Int(m_Header.Height \ 8)
    
    End Select

End Function

'Want to access individual chunks?  Use this to return the index of a given chunk type.  (Note that some
' chunk types can appear multiple times, so you may need to iterate this function more than once!  That's the
' point of the "starting index" value - PD will first check *that* index, then move upward.)
'RETURNS: -1 if the chunk doesn't exist or the starting index is invalid; some value >= 0 if the chunk does exist
Friend Function GetIndexOfChunk(ByRef chunkType As String, Optional ByVal startIndex As Long = 0) As Long
    
    GetIndexOfChunk = -1
    If (startIndex < 0) Then startIndex = 0
    
    Do While (startIndex < m_NumOfChunks)
    
        If (m_Chunks(startIndex).GetType = chunkType) Then
            GetIndexOfChunk = startIndex
            Exit Do
        End If
        
        startIndex = startIndex + 1
    
    Loop
    
End Function

'Returns TRUE if the requested chunk (case-sensitive!) appears in the file
Friend Function HasChunk(ByVal chunkName As String) As Boolean
    HasChunk = (GetIndexOfChunk(chunkName) >= 0)
End Function

'Want data on warnings?  Use these helper functions.
Friend Function Warnings_GetCount() As Long
    Warnings_GetCount = m_Warnings.GetNumOfStrings()
End Function

Friend Sub Warnings_CopyList(ByRef dstStack As pdStringStack)
    Set dstStack = m_Warnings
End Sub

Friend Sub Warnings_DumpToDebugger()
    If (m_Warnings.GetNumOfStrings() > 0) Then
        Dim i As Long
        For i = 0 To m_Warnings.GetNumOfStrings() - 1
            PDDebug.LogAction "WARNING: pdPNG reported: " & m_Warnings.GetString(i)
        Next i
    End If
End Sub

'Some wrapper functions are provided to help deal with complicated child chunks.

'Return the image's background color.  Various color types are handled automatically, and converted to 24-bpp RGB for you.
' RETURNS: -1 if no chunk exists; otherwise, a valid 24-bpp RGB triple is returned.
Friend Function GetBackgroundColor() As Long

    If Me.HasChunk("bKGD") Then
        
        GetBackgroundColor = m_Chunks(Me.GetIndexOfChunk("bKGD")).GetBackgroundColor((m_Header.BitDepth > 8), m_Warnings)
        
        'If this image is a paletted image, the background color is a palette index; convert it to an actual RGB triple
        If (m_Header.ColorType = png_Indexed) Then
            
            Dim palColors() As RGBQuad, numColors As Long
            If m_Chunks(Me.GetIndexOfChunk("PLTE")).GetPalette(palColors, numColors, m_Warnings) Then
                GetBackgroundColor = RGB(palColors(GetBackgroundColor).Red, palColors(GetBackgroundColor).Green, palColors(GetBackgroundColor).Blue)
            End If
            
        End If
    
    Else
        GetBackgroundColor = -1
    End If
    
End Function

'Request a copy of this PNG's header; returns a blank header if no PNG is loaded
Friend Sub GetHeader(ByRef dstHeader As PD_PNGHeader)
    dstHeader = m_Header
End Sub

Friend Function GetBytesPerPixel() As Long
    GetBytesPerPixel = m_Header.BitsPerPixel
End Function

Friend Function GetColorType() As PD_PNGColorType
    GetColorType = m_Header.ColorType
End Function

Friend Function HasAlpha() As Boolean
    HasAlpha = (m_Header.ColorType = png_GreyscaleAlpha) Or (m_Header.ColorType = png_TruecolorAlpha) Or (Me.GetIndexOfChunk("tRNS") >= 0)
End Function

'After all chunks have been loaded and parsed, you can call this function to convert the embedded IHDR chunk
' into a VB-friendly "png header" type.  Bad header values (e.g. width/height = 0) will return failure; you need
' to check for failure states and respond accordingly.
Private Function PopulateHeader() As PD_PNGResult
    
    'Failsafe checks
    PopulateHeader = png_Failure
    If (m_NumOfChunks < 3) Then Exit Function
    If (m_Chunks(0).GetType <> "IHDR") Then Exit Function
    
    'Assume valid values from here on out.  If we encounter a (rare) critical failure value,
    ' we'll reset this value as necessary.
    PopulateHeader = png_Success
    
    'Grab a reference to the underlying chunk stream and reset its pointer to the start of the stream
    Dim tmpStream As pdStream
    Set tmpStream = m_Chunks(0).BorrowData()
    tmpStream.SetPosition 0, FILE_BEGIN
    
    'PNG files use a fixed-length header:
    '   Width               4 bytes
    '   Height              4 bytes
    '   Bit depth           1 byte
    '   Colour type         1 byte
    '   Compression method  1 byte
    '   Filter method       1 byte
    '   Interlace method    1 byte
    
    'We're going to parse each of these values in turn, and if one experiences a critical failure,
    ' we'll suspend further processing.
    m_Header.Width = tmpStream.ReadLong_BE()
    m_Header.Height = tmpStream.ReadLong_BE()
    
    If (m_Header.Width <= 0) Or (m_Header.Height <= 0) Then
        m_Warnings.AddString "Invalid width or height value (" & CStr(m_Header.Width) & "x" & CStr(m_Header.Height) & ")"
        PopulateHeader = png_Failure
    End If
    
    If (PopulateHeader < png_Failure) Then
    
        'Bit-depth and color type need to be handled together, as the value of one restricts the allowed
        ' values of the other.
        m_Header.BitDepth = tmpStream.ReadByte()
        m_Header.ColorType = tmpStream.ReadByte()
        
        'Validating bit-depth and color type requires an ugly table:
        'Colour type             Allowed depths     Interpretation
        '0 - Greyscale           1, 2, 4, 8, 16     Each pixel is a greyscale sample
        '2 - Truecolour          8, 16              Each pixel is an R,G,B triple
        '3 - Indexed             1, 2, 4, 8         Each pixel is a palette index; a PLTE chunk shall appear.
        '4 - Greyscale + alpha   8, 16              Each pixel is a greyscale sample followed by an alpha sample.
        '6 - Truecolour + alpha  8, 16              Each pixel is an R,G,B triple followed by an alpha sample.
        
        'Note that we don't validate the palette chunk, if any, in this function.
        Dim mismatchedDepth As Boolean: mismatchedDepth = False
        
        With m_Header
            
            If (.ColorType = png_Greyscale) Then
                mismatchedDepth = (.BitDepth <> 1) And (.BitDepth <> 2) And (.BitDepth <> 4) And (.BitDepth <> 8) And (.BitDepth <> 16)
            ElseIf (.ColorType = png_Truecolor) Then
                mismatchedDepth = (.BitDepth <> 8) And (.BitDepth <> 16)
            ElseIf (.ColorType = png_Indexed) Then
                mismatchedDepth = (.BitDepth <> 1) And (.BitDepth <> 2) And (.BitDepth <> 4) And (.BitDepth <> 8)
            ElseIf (.ColorType = png_GreyscaleAlpha) Then
                mismatchedDepth = (.BitDepth <> 8) And (.BitDepth <> 16)
            ElseIf (.ColorType = png_TruecolorAlpha) Then
                mismatchedDepth = (.BitDepth <> 8) And (.BitDepth <> 16)
                
            'Any other color type is invalid
            Else
                m_Warnings.AddString "Invalid color type: " & CStr(.ColorType)
                PopulateHeader = png_Failure
            End If
            
            If mismatchedDepth Then
                m_Warnings.AddString "Color type &" & CStr(.ColorType) & ") and bit-depth (" & CStr(.BitDepth) & ") combination is invalid."
                PopulateHeader = png_Failure
            End If
            
        End With
        
    End If
    
    'Compression and filter method only have one supported value, so we don't store them; we simply validate
    Dim tmpByte As Byte
    If (PopulateHeader < png_Failure) Then
        
        'Ensure compression = 0
        tmpByte = tmpStream.ReadByte()
        If (tmpByte <> 0) Then
            m_Warnings.AddString "Invalid compression type in header: " & CStr(tmpByte)
            PopulateHeader = png_Failure
        End If
        
        'Ensure filter = 0
        tmpByte = tmpStream.ReadByte()
        If (tmpByte <> 0) Then
            m_Warnings.AddString "Invalid filter type in header: " & CStr(tmpByte)
            PopulateHeader = png_Failure
        End If
        
    End If
    
    'Finally, check interlacing.  This is only allowed to be 0 (no interlacing) or 1 (interlaced).
    If (PopulateHeader < png_Failure) Then
    
        tmpByte = tmpStream.ReadByte()
        If (tmpByte < 2) Then
            m_Header.Interlaced = (tmpByte = 1)
        Else
            m_Warnings.AddString "Invalid interlaced type in header: " & CStr(tmpByte)
            PopulateHeader = png_Failure
        End If
    
    End If
    
    'If the header was loaded successfully, populate some "convenience" header bits for subsequent functions.
    If (PopulateHeader < png_Failure) Then
    
        'Convert the current color type + bit-depth into a usable "bits per pixel" value.  We need this to
        ' know how much space to allocate for the decompressed pixel data (IDAT chunk).
        With m_Header
        
            'Greyscale is easy - the bit-depth is also the bits-per-pixel value
            If (.ColorType = png_Greyscale) Then
                .BitsPerPixel = .BitDepth
                
            'Truecolor must be multipled by 3, since there are three channels (RGB)
            ElseIf (.ColorType = png_Truecolor) Then
                .BitsPerPixel = .BitDepth * 3
                
            'Indexed is like grayscale; bit-depth is also bits-per-pixel
            ElseIf (.ColorType = png_Indexed) Then
                .BitsPerPixel = .BitDepth
            
            'Greyscale + alpha is bit-depth * 2 (because there are two channels in the image)
            ElseIf (.ColorType = png_GreyscaleAlpha) Then
                .BitsPerPixel = .BitDepth * 2
            
            'Finally, truecolor + alpha is bit-depth * 4 (RGBA channels)
            ElseIf (.ColorType = png_TruecolorAlpha) Then
                .BitsPerPixel = .BitDepth * 4
                
            End If
                    
        End With
        
    End If

End Function

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String, Optional ByVal writeDebugLog As Boolean = True)
    If UserPrefs.GenerateDebugLogs Then
        If writeDebugLog Then PDDebug.LogAction "pdPNG." & funcName & "() reported an error on file """ & m_SourceFilename & """: " & errDescription
    Else
        Debug.Print "pdPNG." & funcName & "() reported an error on file """ & m_SourceFilename & """: " & errDescription
    End If
End Sub

Private Sub ResetChunks()
    m_NumOfChunks = 0
    ReDim m_Chunks(0 To INIT_NUM_OF_CHUNKS - 1) As pdPNGChunk
End Sub

Private Sub Class_Initialize()
    Set m_Stream = New pdStream
    Set m_Warnings = New pdStringStack
End Sub

Private Sub Class_Terminate()
    If (Not m_Stream Is Nothing) Then
        If m_Stream.IsOpen() Then m_Stream.StopStream True
    End If
End Sub
