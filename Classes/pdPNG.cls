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
'Last updated: 12/April/18
'Last update: initial build
'
'I have tried - *so* hard - to make 3rd-party PNG solutions work.  But they all suck.  Every last one of them.
' Like most standards attempts, PNG is an overcomplicated, hackish mess, and every library that attempts to
' wrap the format only makes things worse for developers.  (I mean, let's take LibPNG as an example.  It is
' the official PNG reference library, so it must be well-constructed, right?  It wouldn't do something asinine
' like implementing error-handling via setjmp/longjmp, would it?  Oh wait - it would?  @#$#$^@#$#^*^!!)
'
'After wasting dozens of hours fighting these damn libraries, I am giving up and writing my own PNG parser.
' PD needs fairly intricate access to PNG internals - during both reading and writing - and given the format's
' current ubiquity for images < 24-bpp, it makes sense to handle this unique format manually.  Since we ship
' zLib anyway due to legacy compression tasks, we may as well just use it for this, too, and forgo the
' nightmare of trying to make legacy C-libraries work with VB6.
'
'At present, this class is designed around making it easy to test PNG parsing.  I'll revisit the API design
' after getting everything to a "working" state.
'
'As of v7.2 nightly builds, this class is still under construction and it is *not* currently used as PD's
' main PNG parser.  (That "honor", dubious as it is, still falls to FreeImage.)
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

'Like any format, PNGs have some "magic number" values.
Private Enum PD_PNGColorType
    png_Greyscale = 0
    png_Truecolor = 2
    png_Indexed = 3
    png_GrayscaleAlpha = 4
    png_TruecolorAlpha = 6
End Enum

#If False Then
    Private Const png_Greyscale = 0, png_Truecolor = 2, png_Indexed = 3, png_GrayscaleAlpha = 4, png_TruecolorAlpha = 6
#End If

'PNGs use a standard header, which we convert into this VB-friendly version.  Note that not all information
' from the IHDR chunk is stored here; entries like Compression and Filter only have one allowed value, so while
' we validate these at load-time, we don't go to the trouble of storing them internally.  Similarly, some of
' these values (like bits-per-pixel) are not formally stored in the original header - instead, we infer them.
Private Type PD_PNGHeader
    Width As Long
    Height As Long
    ColorType As PD_PNGColorType
    Interlaced As Boolean
    BitDepth As Byte
    BitsPerPixel As Byte
End Type

Private m_Header As PD_PNGHeader

'If warnings are encountered during processing, we push their messages onto a string stack.  (We may
' decide to report these to the user... I haven't decided yet.  Either way, it'll be helpful for debugging.)
Private m_Warnings As pdStringStack

'pdStream makes our life much easier!
Private m_Stream As pdStream

'At present, we require the caller to pass an identical source file path to every load function.
' (This is a cheap and easy way to ensure no funny business is afoot!)
Private m_SourceFilename As String

'Chunk collection.  The size of the collection *may not match* the number of chunks actually loaded.
' (For performance reasons, the array is allocated using a power-of-2 strategy.)
Private Const INIT_NUM_OF_CHUNKS As Long = 8
Private m_NumOfChunks As Long
Private m_Chunks() As pdPNGChunk

'The first step toward loading a PNG is validating it.  Do this first, before attempting anything else.
Friend Function Step1_ValidatePNG(ByRef srcFile As String, Optional ByVal checkExtension As Boolean) As PD_PNGResult
    
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
    ' three valid chunks (IHDR, IDAT, IEND); 8 + 12 * 3 = 44).
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
    
    'Note an outright failure state in the debugger
    If (okToProceed >= png_Failure) Then
        InternalError "Step1_ValidatePNG", "file is not a valid PNG file"
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
            m_Warnings.AddString "Checksum validation failed on chunk " & CStr(m_NumOfChunks + 1) & " - " & chunkType & " (" & Hex$(embeddedCRC) & " expected, " & Hex$(testCRC) & " found)"
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
    
    'Want to debug the new chunk list?  Here you go:
    'For i = 0 To m_NumOfChunks - 1
    '    Debug.Print i, m_Chunks(i).GetType, m_Chunks(i).GetDataSize
    'Next i
    
    'We now want to proceed with decompression, but because the PNG format is moronic, this step isn't
    ' as simple as asking each chunk to decompress itself.  zLib streams don't store the original,
    ' uncompressed size of their data stream.  You are expected to store that data on your own.  PNG files
    ' decided not to do that.  Instead, you have to manually infer each chunks decompressed size from
    ' data unique to that chunk.  (And for some chunks - like zTXT - your only option is to repeatedly
    ' attempt to decompress the chunk, with ever-larger buffers, until a full decompress works.
    ' That's how shittily they've designed this.)
    
    'IDAT chunks are the most obnoxious to decompress, because their inflated (normal) size is a
    ' function of the image's dimensions plus its color depth.  This means we need to parse the PNG
    ' header and retrieve some critical bits before continuing.
    Step3_Decompress = PopulateHeader()
    
    'If the header didn't validate, don't decompress anything as the entire file is invalid
    If (Step3_Decompress < png_Failure) Then
    
        'The header appears to be valid.  We now know enough to decompress the IDAT chunk.
        
        'The required size of the uncompressed IDAT chunk is simple: width * (BitsPerPixel \ 8, rounded up) * height + height.
        ' Height is added once more to the total because each scanline has a 1-byte "filter" value prepended to it.
        ' (That single-byte filter value is *not* multiplied by bit-depth, obviously.)
        Dim reqIDATSize As Long
        
        'If individual pixels require byte-aligned sizes, this step is easy
        If (m_Header.BitsPerPixel >= 8) Then
            reqIDATSize = m_Header.Width * m_Header.Height * Int(m_Header.BitsPerPixel \ 8) + m_Header.Height
        
        'If individual pixels require less than one-byte, round each scanline up to the nearest byte boundary
        Else
            reqIDATSize = Int((m_Header.Width * m_Header.BitsPerPixel + 7) \ 8) * m_Header.Height + m_Header.Height
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
                
            End If
        Next i
        
    End If
    
    Exit Function
    
InternalVBError:
    InternalError "Step3_Decompress", "internal VB error #" & Err.Number & ": " & Err.Description
    Step3_Decompress = png_Failure

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
            ElseIf (.ColorType = png_GrayscaleAlpha) Then
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
    Dim testByte As Byte
    If (PopulateHeader < png_Failure) Then
        
        'Ensure compression = 0
        testByte = m_Stream.ReadByte()
        If (testByte <> 0) Then
            m_Warnings.AddString "Invalid compression type in header: " & CStr(testByte)
            PopulateHeader = png_Failure
        End If
        
        'Ensure filter = 0
        testByte = m_Stream.ReadByte()
        If (testByte <> 0) Then
            m_Warnings.AddString "Invalid filter type in header: " & CStr(testByte)
            PopulateHeader = png_Failure
        End If
        
    End If
    
    'Finally, check interlacing.  This is only allowed to be 0 (no interlacing) or 1 (interlaced).
    If (PopulateHeader < png_Failure) Then
    
        testByte = m_Stream.ReadByte()
        If (testByte < 2) Then
            m_Header.Interlaced = (testByte = 1)
        Else
            m_Warnings.AddString "Invalid interlaced type in header: " & CStr(testByte)
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
            ElseIf (.ColorType = png_GrayscaleAlpha) Then
                .BitsPerPixel = .BitDepth * 2
            
            'Finally, truecolor + alpha is bit-depth * 4 (RGBA channels)
            ElseIf (.ColorType = png_TruecolorAlpha) Then
                .BitsPerPixel = .BitDepth * 4
                
            End If
                    
        End With
        
    End If

End Function

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String)
    If UserPrefs.GenerateDebugLogs Then
        PDDebug.LogAction "pdPNG." & funcName & "() reported an error on file """ & m_SourceFilename & """: " & errDescription
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
