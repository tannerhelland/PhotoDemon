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
'Copyright 2018-2026 by Tanner Helland
'Created: 12/April/18
'Last updated: 28/August/25
'Last update: add coverage for new cICP HDR chunk (see https://w3c.github.io/png/Implementation_Report_3e/)
'
'I have tried - *so* hard - to make 3rd-party PNG solutions work.  But they all suck.  Every last one of them.
' Like most standards attempts, PNG is an overcomplicated, hackish mess, and every library that attempts to
' wrap the format only makes things worse for developers.  (I mean, let's take LibPNG as an example.  It is
' the official PNG reference library, so it must be well-constructed, right?  It wouldn't do something asinine
' like implementing error-handling via setjmp/longjmp, would it?  Oh wait - it would?  @#$#$^@#$#^*^!!)
'
'After wasting dozens of hours fighting 3rd-party libraries, I gave up and wrote my own PNG import/export
' library. PhotoDemon needs fairly intricate access to PNG internals - during both reading and writing -
' and given the format's current ubiquity for images < 24-bpp, it makes sense to handle PNGs manually.
' As an added bonus, PD can use libdeflate instead of zlib for de/compression tasks, allowing it to both
' create smaller PNG files, and have faster de/compression times than libpng offers.
'
'As of v8.0 nightly builds, this class *is* used as PD's primary PNG importer and exporter.
' If you encounter bugs or surprises, please let me know so I can resolve them before the next
' stable release.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'PNG files can contain a lot of extra information.  To aid debugging, you can activate "verbose" output;
' this will dump all kinds of diagnostic information to the debug log.  (Note that other PNG-adjacent
' classes have their own version of this constant.)
Private Const PNG_DEBUG_VERBOSE As Boolean = False

'Even the PNG designers admit that their CRC validation is "overkill"
' (see http://www.dalkescientific.com/writings/diary/archive/2014/07/10/png_checksum.html)
' As such, PD does not enable it by default.  For strict CRC validation, set this compile-time constant to TRUE.
Private Const REQUIRE_CRC_VALIDATION As Boolean = False

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

Public Enum PD_PNG_FilterStrategy
    png_FilterUndefined = -1    'If you don't understand filter strategy, pass this; PD will handle the rest for you
    png_FilterAuto = 0
    png_FilterFast = 1
    png_FilterOptimal = 2
    png_FilterNone = 3
    png_FilterSubOnly = 4
    png_FilterUpOnly = 5
    png_FilterAverageOnly = 6
    png_FilterPaethOnly = 7
End Enum

#If False Then
    Private Const png_FilterUndefined = -1, png_FilterAuto = 0, png_FilterFast = 1, png_FilterOptimal = 2, png_FilterNone = 3, png_FilterSubOnly = 4, png_FilterUpOnly = 5, png_FilterAverageOnly = 6, png_FilterPaethOnly = 7
#End If

'APNG enums.  Frame disposal and blend ops are defined by mozilla: https://developer.mozilla.org/en-US/docs/Mozilla/Tech/APNG
Private Enum PD_APNG_FrameDisposal
    APNG_DisposeOp_None = 0
    APNG_DisposeOp_Background = 1
    APNG_DisposeOp_Previous = 2
End Enum

#If False Then
    Private Const APNG_DisposeOp_None = 0, APNG_DisposeOp_Background = 1, APNG_DisposeOp_Previous = 2
#End If

Private Enum PD_APNG_BlendOp
    APNG_BlendOp_Source = 0
    APNG_BlendOp_Over = 1
End Enum

#If False Then
    Private Const APNG_BlendOp_Source = 0, APNG_BlendOp_Over = 1
#End If

'PNGs use a standard header, which we convert into this VB-friendly version.  Note that not all information
' from the IHDR chunk is stored here; entries like Compression and Filter only have one allowed value, so while
' we validate these at load-time, we don't go to the trouble of storing them internally.  Similarly, some of
' these values (like bits-per-pixel) are not formally stored in the original header - instead, we infer them.
' (NOTE: the header type is declared publicly in PD, so we can transfer it to child classes.)
Private m_Header As PD_PNGHeader

'If warnings are encountered during processing, we push their messages onto a string stack.  Loading/saving
' is typically allowed to continue despite warnings, but these can be helpful for further debugging.
Private m_Warnings As pdStringStack

'pdStream handles actual byte traversal; it also makes our life much easier!
Private m_Stream As pdStream

'At present, we require the caller to pass an identical source file path to every load function.
' (This is a cheap and easy way to ensure no funny business.)  If the PNG is loaded directly from memory,
' we flag this with a special name.
Private m_SourceFilename As String
Private Const PNG_LOADED_FROM_MEMORY As String = "LoadFromPtr*"
Private m_SourcePtr As Long, m_SourcePtrLen As Long

'In interlaced images, the number of pixels-per-reduced image is non-obvious.  We calculate these
' line lengths as part of the decompression pass, and rather than re-calculate them on subsequent passes,
' we simply cache the results and re-use them later.  (Note that there is no "y" byte count, because it
' is the same as the pixel count; x requires a separate value, due to potential multi-byte-per-pixel
' color depths.)
Private m_XPixelCount() As Long, m_YPixelCount() As Long
Private m_XByteCount() As Long

'Chunk collection.  The size of the collection *may not match* the number of chunks actually loaded.
' (For performance reasons, the array is allocated using a power-of-2 strategy.)
Private Const INIT_NUM_OF_CHUNKS As Long = 8
Private m_NumOfChunks As Long
Private m_Chunks() As pdPNGChunk

'The vast majority of PNGs are single-frame images, but this class also support animated PNGs.
' Frame sizes (and x/y offsets) of every frame are variable; they do not need to match the base image
' size or be painted at a [0, 0] offset.  We cache frame sizes during the initial parsing step,
' to simplify the process of decompressing and rendering frames in later steps.
Private m_FrameRects() As RectL_WH

'To save us the trouble of tracking down frame data chunks more than once, we cache their position
' as they are encountered during the decompression stage.
Private m_fcTLIndices() As Long
Private m_fdATIndices() As Long

'Animated PNGs do *not* require the first frame (the normal IDAT-based PNG frame) to be part of
' the animation.  Instead, it can be a fallback image for use in viewers that don't support APNGs.
' To help this class know how to deal with that first frame, we flag it during initial chunk parsing.
Private m_IDATIsFrame As Boolean

'Animated PNGs cache some extra animation data, to simplify loading+rendering.  Note also that the
' structure of PD's image loader always attempts to load a single frame first, and only if it succeeds
' does it come back to us and request additional frames.  As such, it's simpler for us to cache all
' multiframe DIBs internally during the load process, then simply pass them out to new layers later.
Private m_AnimationDataCached As Boolean
Private m_FrameCount As Long, m_LoopCount As Long
Private m_FrameDIBs() As pdDIB
Private m_FrameSequence As Long

'Private type used during APNG exporting to collect frame data.  Note that some members are
' only used during normal APNG export, while others are used by the specialized "streaming" interface
' (used for screen recording).
Private Type PD_APNGFrame
    frameBlend As PD_APNG_BlendOp
    frameDIB As pdDIB
    frameDisposal As PD_APNG_FrameDisposal
    frameSeq As Long
    frameTime As Long
    rectOfInterest As RectF
    frameIsDuplicateOrEmpty As Boolean
End Type

'When exporting the image, we need to cache various export parameters to share between export functions.
' (This is required to maintain chunk order, as some chunk data must be written in specific places -
' e.g. "after the palette but before the pixel data".)
Private m_EmbedBkgdColor As Boolean, m_BkgdColor As Long, m_CompositingColor As Long
Private m_ImgPalette() As RGBQuad, m_ImgColorCount As Long, m_ImgMaxColorCount As Long, m_ImgPaletteSafe As Boolean
Private m_WriteTrns As Boolean, m_TrnsValueR As Long, m_TrnsValueG As Long, m_TrnsValueB As Long

'For APNGs only, we set a flag specifying whether full transparency is available for each frame;
' if it isn't (e.g. if the image uses all 256 colors with no room left over for a transparent index),
' this will be set to FALSE; optimizations like frame differentials then need to be deactivated.
Private m_apngTrnsAvailable As Boolean

'When writing an APNG file in a streaming fashion (where the file is started, then left open as
' we wait for frames to arrive), we need to flag some positions in the file as we don't know
' their values until the animation is complete (e.g. frame count).
Const INIT_STREAMING_BUFFER As Long = 32
Private m_flagActlPos As Long, m_streamingFrames() As PD_APNGFrame

'Similarly, streaming APNGs need to cache various frame references so it can optimize frame
' data nicely.  (APNG frames support multiple frame disposal methods, and we have to manually
' run per-frame tests to figure out which ones produce the smallest final file.)
Private m_apngPrevFrame As pdDIB, m_apngBufferFrame As pdDIB, m_apngCurFrameBackup As pdDIB
Private m_cmpTestBuffer() As Byte, m_cmpTestBufferSize As Long
Private m_lastGoodFrame As Long

'PD always attempts to use embedded color management data.  However, when converting from some source formats
' (like DDS > PNG), color management data isn't useful.  The caller can forcibly request that we ignore it.
Private m_IgnoreEmbeddedColorData As Boolean

'Simplified wrapper to load a PNG automatically.
Friend Function LoadPNG_Simple(ByRef srcFile As String, ByRef dstImage As pdImage, ByRef dstDIB As pdDIB, Optional ByVal checkExtension As Boolean = False, Optional ByVal loadFromPtr As Long = 0, Optional ByVal loadFromPtrLen As Long = 0, Optional ByVal offsetInSrcFile As Long = 0) As PD_PNGResult
    
    'We want to support PNG loading from both file and memory; if the passed loadFromPtr value is non-zero,
    ' treat it as a pointer and we'll wrap our stream around it instead.
    If (loadFromPtr <> 0) And (loadFromPtrLen <> 0) Then
        srcFile = PNG_LOADED_FROM_MEMORY
        m_SourcePtr = loadFromPtr
        m_SourcePtrLen = loadFromPtrLen
    Else
        m_SourcePtr = 0
        m_SourcePtrLen = 0
    End If
    
    Dim startTime As Currency, firstStartTime As Currency
    VBHacks.GetHighResTime startTime
    firstStartTime = startTime
    
    'Try to validate the source file
    Dim keepLoading As PD_PNGResult
    keepLoading = Me.ImportStage1_ValidatePNG(srcFile, checkExtension, offsetInSrcFile)
    
    'As long as warnings at critical, we'll continue to try and load the file.  (Only outright failures
    ' will cause us to abandon the attempt entirely.)
    If (keepLoading < png_Failure) Then
    
        If PNG_DEBUG_VERBOSE Then PDDebug.LogAction "File is PNG; running internal parser..."
        
        If PNG_DEBUG_VERBOSE Then
            PDDebug.LogAction "PNG > validation took " & VBHacks.GetTimeDiffNowAsString(startTime)
            VBHacks.GetHighResTime startTime
        End If
        
        'Parse the original file into chunks.  (Once this is done, we no longer need to maintain an
        ' copy of the original PNG data.)
        keepLoading = Me.ImportStage2_PreLoadChunks(srcFile)
        
        If PNG_DEBUG_VERBOSE Then
            PDDebug.LogAction "PNG > chunk parsing took " & VBHacks.GetTimeDiffNowAsString(startTime)
            VBHacks.GetHighResTime startTime
        End If
        
        'If all chunks were loaded successfully and minimal validation passed (e.g. a minimal amount of
        ' required chunks are present), go ahead and decompress all relevant chunks, including merging
        ' multiple IDAT chunks (if any) into a single IDAT instance.
        If (keepLoading < png_Failure) Then keepLoading = Me.ImportStage3_Decompress(srcFile)
        
        If PNG_DEBUG_VERBOSE Then
            PDDebug.LogAction "PNG > inflate took " & VBHacks.GetTimeDiffNowAsString(startTime)
            VBHacks.GetHighResTime startTime
        End If
        
        'All compressed data has now been decompressed.  Our next task is to "un-filter" the IDAT chunk,
        ' which converts the pixel bytestream (which has been "filtered" into some other representation
        ' format) into a raw stream of actual pixel data.
        If (keepLoading < png_Failure) Then keepLoading = Me.ImportStage4_UnfilterIDAT(srcFile)
        
        If PNG_DEBUG_VERBOSE Then
            PDDebug.LogAction "PNG > unfiltering IDAT took " & VBHacks.GetTimeDiffNowAsString(startTime)
            VBHacks.GetHighResTime startTime
        End If
        
        'Next, convert the unfiltered IDAT into actual pixel data.
        If (keepLoading < png_Failure) Then keepLoading = Me.ImportStage5_ConstructImage(srcFile, dstDIB, dstImage)
        
        If PNG_DEBUG_VERBOSE Then
            PDDebug.LogAction "PNG > constructing DIB from IDAT took " & VBHacks.GetTimeDiffNowAsString(startTime)
            VBHacks.GetHighResTime startTime
        End If
        
        'Finally, perform any weird post-processing due to non-standard gamma or chromaticity chunks
        If (keepLoading < png_Failure) Then keepLoading = Me.ImportStage6_PostProcessing(srcFile, dstDIB, dstImage)
        
        If PNG_DEBUG_VERBOSE Then
            PDDebug.LogAction "PNG > post-processing (color management, etc) took " & VBHacks.GetTimeDiffNowAsString(startTime)
            PDDebug.LogAction "Total PNG import time took " & VBHacks.GetTimeDiffNowAsString(firstStartTime)
            VBHacks.GetHighResTime startTime
        End If
        
    End If
    
    LoadPNG_Simple = keepLoading
    
End Function

'The first step toward loading a PNG is validating it.  Do this first, before attempting anything else.
Friend Function ImportStage1_ValidatePNG(ByRef srcFile As String, Optional ByVal checkExtension As Boolean = False, Optional ByVal startingOffset As Long = 0) As PD_PNGResult
    
    On Error GoTo InternalVBError
    
    ResetChunks
    
    'If the passed path is zero, assume the caller is loading the PSD from memory.
    If (LenB(srcFile) = 0) Then m_SourceFilename = PNG_LOADED_FROM_MEMORY Else m_SourceFilename = srcFile
    
    Dim okToProceed As PD_PNGResult
    okToProceed = png_Success
    
    'We always check the file extension.  If the user has *asked* us to check it, we treat extension
    ' mismatches as a failure state.  (Otherwise, it will only raise a warning.)  This step is skipped
    ' if the PNG is loaded directly from memory
    If (m_SourceFilename <> PNG_LOADED_FROM_MEMORY) Then
        If Strings.StringsNotEqual(Right$(m_SourceFilename, 3), "png", True) And checkExtension Then
            m_Warnings.AddString "File extension doesn't match PNG"
            okToProceed = png_FileNotPNG
        End If
    End If
    
    'PNG files must have a certain minimum size (comprising a valid magic number - 8 bytes - and at least
    ' three valid chunks (IHDR, IDAT, IEND); 8 + 12 * 3 = 44).  An empty IDAT chunk would still mean a
    ' broken image, but that's okay; this is just a minor failsafe.
    If (okToProceed < png_Failure) Then
        If (m_SourceFilename = PNG_LOADED_FROM_MEMORY) Then
            If (m_SourcePtrLen < 44) Then
                m_Warnings.AddString "File size is too small to contain valid PNG data"
                okToProceed = png_Failure
            End If
        Else
            If (Files.FileLenW(m_SourceFilename) < 44) Then
                m_Warnings.AddString "File size is too small to contain valid PNG data"
                okToProceed = png_Failure
            End If
        End If
    End If
    
    'If all pre-checks passed, open a stream on the file.
    If (okToProceed < png_Failure) Then
        Set m_Stream = New pdStream
        If (m_SourceFilename = PNG_LOADED_FROM_MEMORY) Then
            If Not m_Stream.StartStream(PD_SM_ExternalPtrBacked, PD_SA_ReadOnly, vbNullString, m_SourcePtrLen, m_SourcePtr) Then
                m_Warnings.AddString "Couldn't start in-memory stream against passed pointer: " & m_SourcePtr
                okToProceed = png_Failure
            End If
        Else
            If Not m_Stream.StartStream(PD_SM_FileMemoryMapped, PD_SA_ReadOnly, m_SourceFilename, , , OptimizeSequentialAccess) Then
                m_Warnings.AddString "Can't read file; it may be locked or in an inaccessible location."
                okToProceed = png_Failure
            End If
        End If
    End If
    
    'The stream is open; validate the PNG's "magic number".
    If (okToProceed < png_Failure) Then
    
        'If the caller wants us to use a different offset, move the file pointer now
        If (startingOffset <> 0) Then m_Stream.SetPosition startingOffset, FILE_BEGIN
        
        Dim header1 As Long, header2 As Long
        header1 = m_Stream.ReadLong_BE()
        header2 = m_Stream.ReadLong_BE()
        If (header1 <> &H89504E47) Or (header2 <> &HD0A1A0A) Then
            m_Warnings.AddString "PNG header failed basic validation.  (This is not a PNG file.)"
            okToProceed = png_FileNotPNG
        Else
            okToProceed = png_Success
        End If
        
    End If
    
    'Note an outright failure state in the debugger, and importantly, don't write this to the debug log if we're
    ' running as a compiled exe (as the PNG's failure is simply due to not being a PNG at all, and not due to
    ' any internal problem in this class).
    If (okToProceed >= png_Failure) Then
        If PNG_DEBUG_VERBOSE Then InternalError "ImportStage1_ValidatePNG", "file is not a valid PNG file", False
        If m_Stream.IsOpen() Then m_Stream.StopStream True
    End If
    
    ImportStage1_ValidatePNG = okToProceed
    
    Exit Function

'Internal VB errors are always treated as catastrophic failures.
InternalVBError:
    InternalError "ImportStage1_ValidatePNG", "internal VB error #" & Err.Number & ": " & Err.Description
    If m_Stream.IsOpen Then m_Stream.StopStream True
    
    m_Warnings.AddString "Internal error in step 1, #" & Err.Number & ": " & Err.Description
    ImportStage1_ValidatePNG = png_Failure
    
End Function

'After validating the PNG header, we want to preload all chunks into their own containers.  This simplifies
' further processing, and it allows us to immediately release the source file.
Friend Function ImportStage2_PreLoadChunks(ByRef srcFile As String) As PD_PNGResult

    On Error GoTo InternalVBError
    
    ImportStage2_PreLoadChunks = png_Success
    
    'Failsafe check
    If Strings.StringsNotEqual(m_SourceFilename, srcFile, False) Then
        InternalError "ImportStage2_PreLoadChunks", "filename has changed since original validation!"
        ImportStage2_PreLoadChunks = png_Failure
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
        ' Before proceeding, calculate a CRC using libdeflate.
        If REQUIRE_CRC_VALIDATION Then
            Dim posBackup As Long, tmpCrcBytes() As Byte
            posBackup = m_Stream.GetPosition()
            m_Stream.ReadBytes tmpCrcBytes, chunkSize + 4
            m_Stream.SetPosition posBackup
            testCRC = Plugin_libdeflate.GetCrc32(VarPtr(tmpCrcBytes(0)), chunkSize + 4)
        End If
        
        'With the CRC calculated, grab the chunk type next.  (Note that the CRC calculation only used
        ' a pointer peek - it didn't actually move the stream pointer!)
        chunkType = m_Stream.ReadString_ASCII(4)
        If (chunkType = "IDAT") Then
            
            'If this is not the first IDAT chunk we've encountered, make sure that the previous chunk
            ' was also IDAT.  (The spec requires IDAT chunks to be contiguous in the file.)
            If idatFound Then
                If (Not m_Chunks(m_NumOfChunks - 1).GetType = "IDAT") Then
                    ImportStage2_PreLoadChunks = png_Warning
                    m_Warnings.AddString "IDAT chunks are not contiguous; this file is technically invalid!"
                End If
            Else
                idatFound = True
            End If
            
        End If
        
        'Create a new chunk to hold this chunk's data
        If (m_NumOfChunks > UBound(m_Chunks)) Then ReDim Preserve m_Chunks(0 To m_NumOfChunks * 2 - 1) As pdPNGChunk
        Set m_Chunks(m_NumOfChunks) = New pdPNGChunk
        m_Chunks(m_NumOfChunks).CreateChunkForImport chunkType, chunkSize
        
        'The chunk object itself will handle loading the chunk's data
        If (chunkSize > 0) Then m_Chunks(m_NumOfChunks).LoadChunkData m_Stream
        
        'Finally, grab the embedded CRC, make sure it matches our internal one, then pass *both* CRCs
        ' to the chunk class.
        embeddedCRC = m_Stream.ReadLong_BE()
        
        If REQUIRE_CRC_VALIDATION Then
            m_Chunks(m_NumOfChunks).NotifyCRCs testCRC, embeddedCRC
            If (embeddedCRC <> testCRC) Then
                ImportStage2_PreLoadChunks = png_Warning
                m_Warnings.AddString "Checksum validation failed on chunk " & CStr(m_NumOfChunks + 1) & " - " & chunkType & " (0x" & Hex$(embeddedCRC) & " expected, 0x" & Hex$(testCRC) & " found)"
            End If
        End If
        
        'For debug purposes, it can be helpful to see what's inside a file - this will tell you about
        ' the chunks we've encountered.
        'PDDebug.LogAction m_NumOfChunks & ", " & m_Chunks(m_NumOfChunks).GetType & ", " & m_Chunks(m_NumOfChunks).GetDataSize & ", " & Hex$(embeddedCRC) & ", " & Hex$(testCRC)
        
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
            ImportStage2_PreLoadChunks = png_Failure
            InternalError "ImportStage2_PreLoadChunks", "Required header chunk (IHDR) is missing or not listed as the first chunk in the file"
        ElseIf (m_Chunks(m_NumOfChunks - 1).GetType <> "IEND") Then
            ImportStage2_PreLoadChunks = png_Failure
            InternalError "ImportStage2_PreLoadChunks", "Required final chunk (IEND) is missing or not listed as the final chunk in the file"
        ElseIf (Not idatFound) Then
            ImportStage2_PreLoadChunks = png_Failure
            InternalError "ImportStage2_PreLoadChunks", "No pixel chunks (IDAT) found in the file!"
        End If
    End If
    
    Exit Function
    
InternalVBError:
    InternalError "ImportStage2_PreLoadChunks", "internal VB error #" & Err.Number & ": " & Err.Description
    If m_Stream.IsOpen Then m_Stream.StopStream True
    
    m_Warnings.AddString "Internal error in step 2, #" & Err.Number & ": " & Err.Description
    ImportStage2_PreLoadChunks = png_Failure

End Function

'After loading all chunks and releasing the source file, we want to decompress any/all compressed chunks.
' While we're at it, we must merge all IDAT blocks into a single instance - they represent a *single*
' compressed data stream.
Friend Function ImportStage3_Decompress(ByRef srcFile As String) As PD_PNGResult
    
    On Error GoTo InternalVBError
    
    ImportStage3_Decompress = png_Success
    
    'Failsafe check(s)
    If Strings.StringsNotEqual(m_SourceFilename, srcFile, False) Then
        InternalError "ImportStage3_Decompress", "filename has changed since original validation!"
        ImportStage3_Decompress = png_Failure
    End If
    
    If (m_NumOfChunks < 3) Then
        InternalError "ImportStage3_Decompress", "not enough chunks to continue!"
        ImportStage3_Decompress = png_Failure
    End If
    
    If (ImportStage3_Decompress >= png_Failure) Then Exit Function
    
    Dim i As Long
    Dim lastFrameNumber As Long
    
    'If we're still here, it means that we have (at a minimum) the three required chunks for
    ' constructing pixel data, so the file is (likely) valid.  We now want to tackle two tasks:
    ' 1) Merging multiple IDAT entries, if any, into a single IDAT chunk.  (IDAT data is a single
    '    zLib stream, but encoders are allowed to split it over multiple IDAT chunks if they want;
    '    the original goal was "streaming decompression" but this is just an annoyance today.)
    ' 1b) For animated PNGs, merging multiple fdAT chunks for each frame into a single fdAT chunk
    '     for each frame.
    ' 2) Decompressing chunks whose entire datastream is compressed in zLib format.  Besides IDAT,
    '    we're primarily concerned with ICC profiles (iCCP), and in animated PNGs, fdAT chunks.
    '    Two other chunks - zTXt and optionally iTXt - can compress a *portion* of their datastream,
    '    but you have to manually scan the stream for markers and offsets, and we don't care about
    '    that in this step.  We just want enough data to start assembling pixel data (which is the
    '    most time-consuming part of the load process).
    
    'Before proceeding, construct an array of frame offsets and sizes.  (For single-frame PNGs,
    ' this will obviously represent just the first frame.)
    If Me.IsAnimated() Then
        
        'The IsAnimated() function call, above, will cache frame count for the image.  Note that we
        ' actually allocate one extra space for frame information; we do this because PNGs allow the
        ' standalone PNG data to *not* be part of the animation - it can be a separate placeholder
        ' image, distinct from all animation frames.  As such, there may be m_FrameCount + 1 images
        ' inside an animated PNG.
        ReDim m_FrameRects(0 To m_FrameCount) As RectL_WH
        
        'We're also going to store indices of the fcTL and fdAT chunks for each frame
        ReDim m_fcTLIndices(0 To m_FrameCount) As Long
        ReDim m_fdATIndices(0 To m_FrameCount) As Long

        lastFrameNumber = 0
        
        'Iterate all chunks, looking for frame control chunks (fcTL).
        
        ' When found, cache frame offsets and dimensions, which we need in order to know how to size
        ' the decompression buffer for each frame.
        For i = 0 To m_NumOfChunks - 1
            
            If (m_Chunks(i).GetType = "fcTL") Then
                
                If (lastFrameNumber <= UBound(m_fcTLIndices)) Then
                    m_Chunks(i).GetFrameRect m_FrameRects(lastFrameNumber)
                    m_fcTLIndices(lastFrameNumber) = i
                    lastFrameNumber = lastFrameNumber + 1
                Else
                    InternalError "ImportStage3_Decompress", "frame number exceeds file's frame count", True
                End If
            End If
            
        Next i
        
        'TBD: the default image does *not* have to be the first frame of the file; instead, the first frame
        ' can come after the default image (which can be an entirely separate placeholder image).
        'Ensure that the first frame's data matches the data for the image as a whole
        'With m_FrameRects(0)
        '    If (.Left <> 0) Or (.Top <> 0) Then InternalError "ImportStage3_Decompress", "APNG first frame contains an illegal offset", True
        '    .Left = 0
        '    .Top = 0
        '    If (.Width <> m_Header.Width) Or (.Height <> m_Header.Height) Then InternalError "ImportStage3_Decompress", "APNG first frame contains illegal dimensions", True
        '    .Width = m_Header.Width
        '    .Height = m_Header.Height
        'End With
        
    Else
        ReDim m_FrameRects(0) As RectL_WH
        With m_FrameRects(0)
            .Left = 0
            .Top = 0
            .Width = m_Header.Width
            .Height = m_Header.Height
        End With
    End If
    
    'Before doing any decompression, let's merge our various IDAT (and fdAT chunks) into single instances.
    Dim firstIDATIndex As Long, numChunksRemoved As Long
    firstIDATIndex = -1
    numChunksRemoved = 0
    
    Dim lastfdATIndex As Long, lastSequenceNumber As Long
    lastfdATIndex = -1
    lastSequenceNumber = -1
    lastFrameNumber = -1
    
    'Animated PNGs require different handling; frames will be identified by a frame control number
    For i = 0 To m_NumOfChunks - 1
        
        'Regular PNGs use IDAT chunks
        If (m_Chunks(i).GetType = "IDAT") Then
            
            'If this is the first IDAT chunk we've found, flag its index and carry on
            If (firstIDATIndex = -1) Then
                firstIDATIndex = i
                
                'If this is an animated PNG, and we've already encountered an fcTL chunk, flag this
                ' as the position of the first frame's data.
                If Me.IsAnimated Then
                    
                    'Ensure this file contains at least one valid animation control chunk
                    If (m_fcTLIndices(0) <> 0) Then
                        
                        'See if the IDAT chunk is part of the animation, or if it is a standalone "placeholder" frame
                        m_IDATIsFrame = (m_fcTLIndices(0) < i)
                        
                        'If this frame *is* part of the animation, store its index
                        If m_IDATIsFrame Then m_fdATIndices(0) = i
                        
                    End If
                    
                End If
            
            'If this is *not* the first IDAT chunk we've found, merge the contents of this chunk into
            ' the first IDAT chunk.
            Else
                
                'As part of the merge step, this chunk will be freed; that's by design!
                m_Chunks(firstIDATIndex).MergeOtherChunk m_Chunks(i)
                Set m_Chunks(i) = Nothing
                
                'Increment the multiple IDAT count; we use these to shift other chunks "forward" in
                ' the chunk list as we erase the duplicate IDATs.
                numChunksRemoved = numChunksRemoved + 1
                
            End If
        
        'Animated PNGs use fdAT chunks (for frames *after* the first one)
        ElseIf (m_Chunks(i).GetType = "fdAT") Then
        
            'If this is the first fdAT chunk we've found for this frame, flag its index and carry on
            If (lastfdATIndex = -1) Then
                
                lastfdATIndex = i
                
                'Double-check the sequence value
                If (m_Chunks(i).GetSeqNumber <> lastSequenceNumber + 1) Then InternalError "ImportStage3_Decompress", "bad sequence number!", True
                lastSequenceNumber = lastSequenceNumber + 1
                
                'Convert the sequence number to a frame number
                m_Chunks(i).SetSeqNumber lastFrameNumber
                m_fdATIndices(lastFrameNumber) = i
                
                'Shift this chunk forward, as necessary
                If (numChunksRemoved > 0) Then
                    Set m_Chunks(i - numChunksRemoved) = m_Chunks(i)
                    Set m_Chunks(i) = Nothing
                    lastfdATIndex = lastfdATIndex - numChunksRemoved
                    m_fdATIndices(lastFrameNumber) = lastfdATIndex
                End If
            
            'If this is *not* the first fdAT chunk we've found for this frame, merge the contents of
            ' this chunk into the last fdAT chunk.
            Else
                
                'Double-check the sequence value
                If (m_Chunks(i).GetSeqNumber <> lastSequenceNumber + 1) Then InternalError "ImportStage3_Decompress", "bad sequence number!", True
                lastSequenceNumber = lastSequenceNumber + 1
                
                'Convert the sequence number to a frame number
                m_Chunks(i).SetSeqNumber lastFrameNumber
                
                'As part of the merge step, this chunk will be freed; that's by design!
                m_Chunks(lastfdATIndex).MergeOtherChunk m_Chunks(i)
                Set m_Chunks(i) = Nothing
                
                'Increment the multiple IDAT count; we use these to shift other chunks "forward" in
                ' the chunk list as we erase the duplicate IDATs.
                numChunksRemoved = numChunksRemoved + 1
                
            End If
            
        'If this is *not* an IDAT or fdAT chunk, and we've removed IDAT or fdAT chunks prior to this one,
        ' shift this chunk forward in the list.
        Else
            
            'If this is an fcTL chunk, check (and increment) sequence count
            If (m_Chunks(i).GetType = "fcTL") Then
                If (m_Chunks(i).GetSeqNumber <> lastSequenceNumber + 1) Then InternalError "ImportStage3_Decompress", "bad sequence number!", True
                lastSequenceNumber = lastSequenceNumber + 1
                lastFrameNumber = lastFrameNumber + 1
                m_Chunks(i).SetSeqNumber lastFrameNumber
                m_fcTLIndices(lastFrameNumber) = i
                If (numChunksRemoved > 0) Then m_fcTLIndices(lastFrameNumber) = i - numChunksRemoved
            End If
            
            If (numChunksRemoved > 0) Then
                Set m_Chunks(i - numChunksRemoved) = m_Chunks(i)
                Set m_Chunks(i) = Nothing
            End If
            
            'While here, reset the fdAT tracker
            lastfdATIndex = -1
            
        End If
        
    Next i
    
    'If multiple IDAT chunks were condensed into a single chunk, update our net chunk count
    m_NumOfChunks = m_NumOfChunks - numChunksRemoved
    
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
    ImportStage3_Decompress = PopulateHeader()
    
    'If the header didn't validate, don't decompress anything as the entire file is invalid
    If (ImportStage3_Decompress < png_Failure) Then
    
        'The header appears to be valid.  We now need to calculate how much space is required for decompressing.
        Dim reqIDATSize As Long
        
        If m_Header.Interlaced Then
            ReDim m_XPixelCount(0 To 6) As Long
            ReDim m_YPixelCount(0 To 6) As Long
            ReDim m_XByteCount(0 To 6) As Long
        Else
            ReDim m_XPixelCount(0) As Long
            ReDim m_YPixelCount(0) As Long
            ReDim m_XByteCount(0) As Long
        End If
        
        reqIDATSize = CalculateInterlaceArrays(m_Header.Width, m_Header.Height)
        
        'Ask all relevant chunks to decompress themselves; note that chunks without compressed data will
        ' just ignore this request.
        For i = 0 To m_NumOfChunks - 1
            
            'If this is an animated PNG, we need to update reqIDATSize every time a new frame is encountered.
            If (m_Chunks(i).GetType = "fcTL") Then
                
                Dim curFrame As Long
                curFrame = m_Chunks(i).GetSeqNumber()
                If (curFrame < m_FrameCount) Then
                    reqIDATSize = CalculateInterlaceArrays(m_FrameRects(curFrame).Width, m_FrameRects(curFrame).Height)
                End If
            
            End If
            
            If (Not m_Chunks(i).DecompressChunk(m_Warnings, reqIDATSize)) Then
                
                'Decompression failures in IDAT are critical; other decompression failures only raise warnings
                If (m_Chunks(i).GetType = "IDAT") Then
                    InternalError "ImportStage3_Decompress", "IDAT decompression failed; PNG file is unreadable."
                    ImportStage3_Decompress = png_Failure
                Else
                    m_Warnings.AddString "WARNING: " & m_Chunks(i).GetType & " could not be decompressed.  (I'll still try to salvage pixel data.)"
                    ImportStage3_Decompress = png_Warning
                End If
                
            End If
            
        Next i
        
    End If
    
    Exit Function
    
InternalVBError:
    InternalError "ImportStage3_Decompress", "internal VB error #" & Err.Number & ": " & Err.Description
    ImportStage3_Decompress = png_Failure

End Function

'Populate the module-level interlacing arrays with correct pixel counts for a given pixel width/height count.
' RETURNS: required size for a decompressed IDAT chunk; 0 if failure
Private Function CalculateInterlaceArrays(ByVal pxWidth As Long, ByVal pxHeight As Long) As Long
    
    CalculateInterlaceArrays = 0
    
    Dim srcX As Long, srcY As Long
    
    'Interlaced images treat each "interlace pass" as its own "mini-PNG", which means each non-empty scanline
    ' in each mini-image gets its own filtering byte.  (I know, it's fucking confusing.)  Because of this,
    ' calculating a decompression buffer size is non-trivial.
    If m_Header.Interlaced Then
        
        'A separate function calculates how many pixels exist in each interlacing pass; we can then use
        ' our standard "bit-depth" adjustment formula on each pass, and tally the results to find a
        ' "net" required decompression buffer size.
        Dim i As Long
        For i = 0 To 6
            
            GetSizeInterlaced pxWidth, pxHeight, srcX, srcY, i + 1
            m_XPixelCount(i) = srcX
            m_YPixelCount(i) = srcY
            m_XByteCount(i) = ((srcX * m_Header.BitsPerPixel + 7) \ 8) + 1
            
            'Per the spec, 0-byte passes (only possible on very small images!) do not get a filter type byte,
            ' so we can skip such lines entirely.
            If (m_XByteCount(i) = 1) Or (srcY = 0) Then
                m_XByteCount(i) = 0
            Else
                CalculateInterlaceArrays = CalculateInterlaceArrays + ((srcX * m_Header.BitsPerPixel + 7) \ 8) * srcY + srcY
            End If
            
        Next i
    
    'Non-interlaced images are much simpler!
    Else
        
        m_XPixelCount(0) = pxWidth
        m_YPixelCount(0) = pxHeight
        m_XByteCount(0) = ((pxWidth * m_Header.BitsPerPixel + 7) \ 8) + 1
        
        'To properly cover the case of bit-depths < 8 (e.g. 1, 2, 4), ensure that scanline width is always
        ' rounded up to the nearest byte alignment.
        CalculateInterlaceArrays = ((pxWidth * m_Header.BitsPerPixel + 7) \ 8) * pxHeight + pxHeight
        
    End If
    
End Function

'Once all relevant chunks have been decompressed, we next need to un-filter each scanline into its original,
' unfiltered pixel representation.  This step operates only on bytes, which means it is completely independent
' of color-depth or bit-depth.
Friend Function ImportStage4_UnfilterIDAT(ByRef srcFile As String) As PD_PNGResult

    On Error GoTo InternalVBError
    
    ImportStage4_UnfilterIDAT = png_Success
    
    'Failsafe check(s)
    If Strings.StringsNotEqual(m_SourceFilename, srcFile, False) Then
        InternalError "ImportStage4_UnfilterIDAT", "filename has changed since original validation!"
        ImportStage4_UnfilterIDAT = png_Failure
    End If
    
    If (ImportStage4_UnfilterIDAT >= png_Failure) Then Exit Function
    
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
        
        'Make sure interlacing data has been correctly prepped
        CalculateInterlaceArrays m_Header.Width, m_Header.Height
        
        'Ask the IDAT chunk to unfilter itself.  (As part of this step, we pass the x/y pixel counts we
        ' already computed as part of the interlacing step - if the image was interlaced.)
        m_Chunks(idatIndex).UnfilterChunk m_Warnings, m_Header, m_XPixelCount, m_YPixelCount, m_XByteCount
        
    End If
    
    'If this is an animated PNG, repeat the step for any fdAT chunks
    If Me.IsAnimated() Then
        
        Dim fakeHeader As PD_PNGHeader
        
        Dim i As Long
        For i = 0 To m_NumOfChunks - 1
            If (m_Chunks(i).GetType = "fdAT") And (m_Chunks(i).GetSeqNumber < m_FrameCount) Then
            
                'Prep a "fake" header with the dimensions of this particular frame
                fakeHeader = m_Header
                fakeHeader.Width = m_FrameRects(m_Chunks(i).GetSeqNumber).Width
                fakeHeader.Height = m_FrameRects(m_Chunks(i).GetSeqNumber).Height
                
                'Make sure interlacing data, if any, is correctly calculated
                CalculateInterlaceArrays fakeHeader.Width, fakeHeader.Height
                
                'Perform the actual unfilter
                m_Chunks(i).UnfilterChunk m_Warnings, fakeHeader, m_XPixelCount, m_YPixelCount, m_XByteCount
                
            End If
        Next i
    
    End If
    
    Exit Function
    
InternalVBError:
    InternalError "ImportStage4_UnfilterIDAT", "internal VB error #" & Err.Number & ": " & Err.Description
    ImportStage4_UnfilterIDAT = png_Failure

End Function

'Once all relevant chunks have been decompressed, we next need to un-filter each scanline into its original,
' unfiltered pixel representation.  This step operates only on bytes, which means it is completely independent
' of color-depth or bit-depth.
Friend Function ImportStage5_ConstructImage(ByRef srcFile As String, ByRef dstDIB As pdDIB, ByRef dstImage As pdImage) As PD_PNGResult
    
    On Error GoTo InternalVBError
    
    ImportStage5_ConstructImage = png_Success
    
    'Failsafe check(s)
    If Strings.StringsNotEqual(m_SourceFilename, srcFile, False) Then
        InternalError "ImportStage5_ConstructImage", "filename has changed since original validation!"
        ImportStage5_ConstructImage = png_Failure
    End If
    
    Dim idatIndex As Long
    idatIndex = Me.GetIndexOfChunk("IDAT")
    If (idatIndex < 0) Or (idatIndex >= m_NumOfChunks) Then
        InternalError "ImportStage5_ConstructImage", "no IDAT chunk!"
        ImportStage5_ConstructImage = png_Failure
    End If
    
    If (ImportStage5_ConstructImage >= png_Failure) Then Exit Function
    
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
        
        'In grayscale images, the transparent color is scaled to the current bit-depth.  So e.g. a 4-bit grayscale
        ' image will have 15 as its maximum transparency index.  When constructing a grayscale palette for upsampling,
        ' we want to use the *final* grayscale value - so upscale the corresponding transparency value now
        If hasTransparency And (m_Header.ColorType = png_Greyscale) Then
            
            Dim grayScaleFactor As Long, grayScaleLimit As Long, grayScaleWarning As Boolean
            grayScaleFactor = 1
            
            If (m_Header.BitDepth = 1) Then
                grayScaleLimit = 1
                grayScaleFactor = 255
            ElseIf (m_Header.BitDepth = 2) Then
                grayScaleLimit = 3
                grayScaleFactor = 85
            ElseIf (m_Header.BitDepth = 4) Then
                grayScaleLimit = 15
                grayScaleFactor = 17
            End If
            
            'Check for badly encoded tRNS values; malformed writers may write an unscaled value, and if found,
            ' we want to mimic libpng and attempt to salvage it with a warning
            If (trnsRed <= grayScaleLimit) Then trnsRed = trnsRed * grayScaleFactor Else grayScaleWarning = True
            If (trnsGreen <= grayScaleLimit) Then trnsGreen = trnsGreen * grayScaleFactor Else grayScaleWarning = True
            If (trnsBlue <= grayScaleLimit) Then trnsBlue = trnsBlue * grayScaleFactor Else grayScaleWarning = True
            If grayScaleWarning Then PDDebug.LogAction "WARNING: PNG has malformed tRNS chunk; attempting to load anyway..."
            
        End If
        
    End If
    
    'Because the tRNS chunk may update the contents of the palette (because it can store a table of
    ' transparency values), only *now* can we notify the parent image of the original palette contents.
    If hasPalette And (Not dstImage Is Nothing) Then dstImage.SetOriginalPalette srcColors, numColors
    
    'We're close enough to final processing that we can finally construct the DIB at its target size.
    ' (Note that - regardless of PNG color depth - we *always* create the DIB as 32-bpp RGBA.)
    Dim okToProceed As Boolean
    If (dstDIB Is Nothing) Then Set dstDIB = New pdDIB
    okToProceed = dstDIB.CreateBlank(m_Header.Width, m_Header.Height, 32, vbWhite, 255)
    
    'If the PNG contained ICC profile data, retrieve it now; otherwise, assume sRGB.  (Note that a
    ' transform may be generated either way, as LittleCMS can be used for fast swizzling.)
    Dim imgHasEmbeddedICC As Boolean, iccEmbeddedV2 As Boolean, iccEmbeddedV4 As Boolean
    If ColorManagement.UseEmbeddedICCProfiles() Then
        iccEmbeddedV2 = (Me.GetIndexOfChunk("iCCP") >= 0)
        iccEmbeddedV4 = (Me.GetIndexOfChunk("iCCN") >= 0)
        imgHasEmbeddedICC = iccEmbeddedV2 Or iccEmbeddedV4
    End If
    
    Dim srcProfile As pdLCMSProfile, dstProfile As pdLCMSProfile, dstTransform As pdLCMSTransform, srcImgProfileGood As Boolean
    
    'The PrepICCData function will mark this as TRUE if the embedded profile passes validation
    srcImgProfileGood = False
    PrepICCData srcProfile, dstProfile, dstTransform, srcImgProfileGood
    
    If srcImgProfileGood Then
    
        'Add a copy of the profile to PD's central cache, and associate the target image with that profile
        Dim tmpProfile As pdICCProfile
        Set tmpProfile = New pdICCProfile
        
        If (Me.GetIndexOfChunk("cICP") >= 0) And srcImgProfileGood Then
            If PNG_DEBUG_VERBOSE Then PDDebug.LogAction "Using cICP data for color management..."
            tmpProfile.LoadICCFromLCMSProfile srcProfile
        Else
            If iccEmbeddedV4 Then
                srcImgProfileGood = tmpProfile.LoadICCFromPtr(m_Chunks(Me.GetIndexOfChunk("iCCN")).BorrowData().GetStreamSize(), m_Chunks(Me.GetIndexOfChunk("iCCN")).BorrowData().Peek_PointerOnly(0))
                If PNG_DEBUG_VERBOSE And srcImgProfileGood Then PDDebug.LogAction "Using iCCN data for color management..."
            ElseIf iccEmbeddedV2 Then
                srcImgProfileGood = tmpProfile.LoadICCFromPtr(m_Chunks(Me.GetIndexOfChunk("iCCP")).BorrowData().GetStreamSize(), m_Chunks(Me.GetIndexOfChunk("iCCP")).BorrowData().Peek_PointerOnly(0))
                If PNG_DEBUG_VERBOSE And srcImgProfileGood Then PDDebug.LogAction "Using iCCP data for color management..."
            End If
        End If
        
        If srcImgProfileGood Then
            
            Dim profHash As String
            profHash = ColorManagement.AddProfileToCache(tmpProfile, True, False, False)
            If (Not dstImage Is Nothing) Then dstImage.SetColorProfile_Original profHash
            
            'IMPORTANT NOTE: at present, the destination DIB - by the time we're done with it - will have been
            ' hard-converted to sRGB, so we don't want to associate the destination DIB with its source profile.
            ' Instead, note that it is currently sRGB.
            profHash = ColorManagement.GetSRGBProfileHash()
            dstDIB.SetColorProfileHash profHash
            dstDIB.SetColorManagementState cms_ProfileConverted
            
        End If
        
    End If
    
    'Ensure that all our previous steps completed successfully, and that we have a valid destination image to
    ' place our decoded data into.
    If okToProceed Then
        
        'If the image is 8-bpp grayscale, we now want to construct an artifical palette.  (This lets us process it
        ' in an identical manner to indexed color images.)
        If (m_Header.ColorType = png_Greyscale) Then ConstructGrayPalette srcColors, numColors, trnsRed
        
        'Certain bit-depths require us to handle transparency manually (but note that 8-bpp grayscale has already
        ' been covered by the ConstructGrayPalette() function, above!)
        Dim trnsHandlingRequired As Boolean, trnsCheck As Long
        trnsHandlingRequired = (m_Header.ColorType = png_Truecolor) And (trnsRed >= 0) And (trnsGreen >= 0) And (trnsBlue >= 0)
        If trnsHandlingRequired And (m_Header.BitDepth = 8) Then
            
            'August 2025: per the latest PNG spec revision, we must explicitly handle trns values
            ' that exceed known valid values for this color space.
            If (trnsRed <= 255) And (trnsGreen <= 255) And (trnsBlue <= 255) Then
                trnsCheck = RGB(trnsRed, trnsGreen, trnsBlue)
            Else
                trnsCheck = &HFFFFFFFF
            End If
            
        End If
        
        'We now have everything we need to construct actual 32-bpp DIBs usable by PhotoDemon.
        
        '(For animated PNGs, note that we assemble *all* frames now, even though PD won't request them
        ' during this step - instead we just cache them locally.)
        If Me.IsAnimated() Then ReDim m_FrameDIBs(0 To m_FrameCount - 1) As pdDIB
        
        Dim curFrame As Long, endFrame As Long
        If Me.IsAnimated Then endFrame = m_FrameCount - 1 Else endFrame = 0
        
        For curFrame = 0 To endFrame
        
            Dim targetDIB As pdDIB, targetWidth As Long, targetHeight As Long
            
            If (curFrame = 0) Then
                
                Set targetDIB = dstDIB
                targetWidth = m_Header.Width
                targetHeight = m_Header.Height
                
                'On APNGs, we need to cover a weird case where the first frame in the animation is
                ' just a "placeholder" image to satisfy non-APNG-compatible readers, but it's *not*
                ' actually part of the animation - if this occurs, we want to skip the default image
                ' and move directly to the first animation frame.
                If Me.IsAnimated() Then
                    If (Not m_IDATIsFrame) Then idatIndex = m_fdATIndices(curFrame)
                End If
                
            Else
                
                'In an APNG, all subsequent frames need to be created in turn.  This is a large memory hit,
                ' but we'll incur it one way or another since we need to do this step when constructing the
                ' final pdImage object anyway (and it's more convenient to take the hit now).
                Set m_FrameDIBs(curFrame) = New pdDIB
                targetWidth = m_FrameRects(curFrame).Width
                targetHeight = m_FrameRects(curFrame).Height
                If PNG_DEBUG_VERBOSE Then PDDebug.LogAction "Created DIB " & targetWidth & "x" & targetHeight & " for frame #" & (curFrame + 1)
                
                m_FrameDIBs(curFrame).CreateBlank targetWidth, targetHeight, 32, vbWhite, 255
                Set targetDIB = m_FrameDIBs(curFrame)
                idatIndex = m_fdATIndices(curFrame)
                
            End If
            
            'To simplify our code, we now want to upsample any color-depths below 8-bpp.  (They present a
            ' unique challenge within VB, since we lack the bit-shift operators necessary to handle them
            ' elegantly.)  Note that the upsampler will reset our xStride value to match 8-bpp data
            ' it produces.  (Note that we also pass the special case of 8-bpp *interlaced* data to this
            ' function; that's because the function performs de-interlacing of the 8-bpp data produced
            ' by the upsampler, and it spares us a bit of code to reuse that functionality for 8-bpp
            ' interlaced images.)
            CalculateInterlaceArrays targetWidth, targetHeight
            If (m_Header.BitDepth < 8) Or ((m_Header.BitsPerPixel = 8) And m_Header.Interlaced) Then UpsampleLowBitDepthData idatIndex, targetWidth, targetHeight
            
            '8-bpp interlaced images have now been handled correctly.  Interlaced images at higher
            ' bit-depths (16, 24, 32, 48, 64) need to be dealt with using a separate function.
            If ((m_Header.BitsPerPixel > 8) And m_Header.Interlaced) Then DeInterlaceNonPaletteData idatIndex, targetWidth, targetHeight
            
            'Interlaced images have now been converted to standard, non-interlaced streams.  No further
            ' steps require special handling for interlaced images.
            
            'We can now assign actual pixel bits!
            Dim dstBytes() As Byte, dstSA As SafeArray2D
            targetDIB.WrapArrayAroundDIB dstBytes, dstSA
            
            'To improve performance, we use unsafe direct access to the chunk's contents
            Dim srcBytes() As Byte, srcSA As SafeArray1D
            m_Chunks(idatIndex).BorrowData.WrapArrayAroundMemoryStream srcBytes, srcSA, 0
                
            'Prep all the required loop values
            Dim imgWidth As Long, imgHeight As Long, xFinal As Long, yFinal As Long
            Dim x As Long, y As Long, xDIB As Long, srcOffset As Long
            imgWidth = targetWidth
            imgHeight = targetHeight
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
                            If srcImgProfileGood Then targetDIB.SetColorManagementState cms_ProfileConverted
                        
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
                            If srcImgProfileGood Then targetDIB.SetColorManagementState cms_ProfileConverted
                            
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
                        If srcImgProfileGood Then targetDIB.SetColorManagementState cms_ProfileConverted
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
                        If srcImgProfileGood Then targetDIB.SetColorManagementState cms_ProfileConverted
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
                        If srcImgProfileGood Then targetDIB.SetColorManagementState cms_ProfileConverted
                        
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
                
                    'If the image has an embedded color profile, combine swizzling and color management
                    ' into a single step.
                    If imgHasEmbeddedICC And (Not dstTransform Is Nothing) Then
                        
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
                        If srcImgProfileGood Then targetDIB.SetColorManagementState cms_ProfileConverted
                        
                    'If the image does *not* have an embedded ICC profile, we can do a faster manual swizzle.
                    Else
                        
                        '32-bit RGBA are the most common use-case for standard PNGs, so they are worthwhile
                        ' to manually accelerate.
                        If (m_Header.BitDepth = 8) Then
                            
                            Dim dstBytesLine() As Byte, dstSA1D As SafeArray1D
                            targetDIB.WrapArrayAroundScanline dstBytesLine, dstSA1D, 0
                            
                            Dim scanStart As Long, scanWidth As Long
                            scanStart = targetDIB.GetDIBPointer
                            scanWidth = targetDIB.GetDIBStride
                        
                            For y = 0 To yFinal
                                
                                srcOffset = y * xStride + 1
                                dstSA1D.pvData = scanStart + y * scanWidth
                                
                                'Start by copying the unfiltered IDAT as-is into the destination DIB.
                                ' (Green and alpha bytes are already in the correct position; we only need
                                ' to swizzle red and blue to the required BGRA order.)
                                CopyMemoryStrict VarPtr(dstBytesLine(0)), VarPtr(srcBytes(srcOffset)), imgWidth * 4
                                
                                'Manually swizzle r/b
                                For x = 0 To xFinal * 4 Step 4
                                    r = dstBytesLine(x)
                                    dstBytesLine(x) = dstBytesLine(x + 2)
                                    dstBytesLine(x + 2) = r
                                Next x
                                
                            Next y
                            
                            targetDIB.UnwrapArrayFromDIB dstBytesLine
                            
                        '64-bit is much rarer "in the wild", and it needs to be processed manually as VB
                        ' doesn't natively support ushorts.
                        Else
                        
                            For y = 0 To yFinal
                                srcOffset = y * xStride + 1
                            For x = 0 To xFinal
                                
                                xDIB = x * 4
                                dstBytes(xDIB, y) = srcBytes(srcOffset + x * 8 + 4)
                                dstBytes(xDIB + 1, y) = srcBytes(srcOffset + x * 8 + 2)
                                dstBytes(xDIB + 2, y) = srcBytes(srcOffset + x * 8)
                                dstBytes(xDIB + 3, y) = srcBytes(srcOffset + x * 8 + 6)
                                
                            Next x
                            Next y
                        
                        End If
                        
                    End If
                    
                    okToProceed = True
                
            End Select
            
            'Release all unsafe array wrappers
            targetDIB.UnwrapArrayFromDIB dstBytes
            m_Chunks(idatIndex).BorrowData.UnwrapArrayFromMemoryStream srcBytes
            
            'If this is an animated PNG, make a local copy of the target DIB on the first frame.
            ' We need to reuse this when constructing an output buffer if subsequent animation frames
            ' are only partial frames.
            If Me.IsAnimated And (curFrame = 0) Then
                Set m_FrameDIBs(curFrame) = New pdDIB
                m_FrameDIBs(curFrame).CreateFromExistingDIB dstDIB
            End If
            
            If okToProceed Then ImportStage5_ConstructImage = png_Success Else ImportStage5_ConstructImage = png_Failure
            
        Next curFrame
        
    End If
    
    'If we constructed any ICC profile or transform handles, make sure to free them before exiting
    Set srcProfile = Nothing
    Set dstProfile = Nothing
    Set dstTransform = Nothing
    
    Exit Function
    
InternalVBError:
    InternalError "ImportStage5_ConstructImage", "internal VB error #" & Err.Number & ": " & Err.Description
    ImportStage5_ConstructImage = png_Failure
End Function

'Once image data has been fully constructed, the last thing we need to do is check for any esoteric chunks that affect
' the final image presentation (like gamma or chromaticity conversions).  For 99% of images, this step won't do a thing.
Friend Function ImportStage6_PostProcessing(ByRef srcFile As String, ByRef dstDIB As pdDIB, ByRef dstImage As pdImage) As PD_PNGResult

    On Error GoTo InternalVBError
    
    ImportStage6_PostProcessing = png_Success
    
    'Failsafe check(s)
    If Strings.StringsNotEqual(m_SourceFilename, srcFile, False) Then
        InternalError "ImportStage6_PostProcessing", "filename has changed since original validation!"
        ImportStage6_PostProcessing = png_Failure
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
            If (Not dstImage Is Nothing) Then dstImage.SetDPI dpiWidth, dpiHeight
        End If
        
    End If
    
    'UPDATE AUGUST 2025
    ' - Changed hierarchy of color tags, and expanded support for new color management tags (iCCN and cICP)
    '
    'PNG presents a hierarchy of lighting and color adjustment chunks.  If present, they are to be dealt with
    ' in this order (and once one is dealt with, subsequent ones should be ignored, to avoid double-corrections):
    
    '1) cICP
    '2) iCCN
    '3) iCCP
    '4) sRGB
    '5) gAMA and/or cHRM (equal priority, and should be handled together if possible)
    
    'ICC profiles or sRGB flags, if present, are dealt with during a previous decoding step (for performance reasons).
    ' As such, we don't need to deal with them here - and in fact, we can bail if they were handled previously,
    ' as they take precedence over all other corrective functions.
    If (dstDIB.GetColorManagementState <> cms_NoManagement) Then
        ImportStage6_PostProcessing = png_Success
        Exit Function
    End If
    
    'As of v8.0, PD assumes sRGB for untagged files, so the sRGB chunk doesn't mean anything special to us.
    ' If it exists, however, it will keep us from manually adjusting gamma and chromaticity to match.
    ' (Gamma in particular is important, as sRGB doesn't use a standard exponent curve, so you can't use a
    ' gAMA chunk to replicate it correctly.)
    If (Me.GetIndexOfChunk("sRGB") >= 0) Then
        ImportStage6_PostProcessing = png_Success
        Exit Function
    End If
    
    'Next, look for gAMA and cHRM chunks, if any
    Dim gammaIndex As Long, chrmIndex As Long
    gammaIndex = Me.GetIndexOfChunk("gAMA")
    chrmIndex = Me.GetIndexOfChunk("cHRM")
    
    'Handle gamma; this allows us to produce the same results as Firefox, Chrome, Edge, Internet Explorer,
    ' and (surprisingly) Windows Photo Viewer on Windows 10.  (For whatever reason, Safari does *not* handle
    ' gAMA chunks; given that PD is Windows-only, I haven't provided a toggle to change this behavior,
    ' at present.)  Note that photo editing software may handle gamma differently - I have *not* tested them
    ' thoroughly, and if they don't manage gamma correctly, I may need to revisit this design entirely.
    Dim gammaOK As Boolean
    If (gammaIndex >= 0) And ColorManagement.UseEmbeddedLegacyProfiles() And (Not m_IgnoreEmbeddedColorData) Then
        
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
    Dim chrmOK As Boolean
    If (chrmIndex >= 0) And ColorManagement.UseEmbeddedLegacyProfiles() And (Not m_IgnoreEmbeddedColorData) Then
        
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
            
            'Per https://github.com/tannerhelland/PhotoDemon/issues/612, when chromaticity data is present
            ' but gamma is *not*, assume 2.2 gamma.  (This resolves color issues from images originating
            ' from ezgif.com.)
            If (Not gammaOK) Then
                If PNG_DEBUG_VERBOSE Then PDDebug.LogAction "WARNING: chromaticity present, but gamma missing; colors may be incorrect!"
                gammaFloat = 2.2
            End If
            
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
                    
                    If (Not dstImage Is Nothing) Then dstImage.SetColorProfile_Original profHash
                    
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
    InternalError "ImportStage6_PostProcessing", "internal VB error #" & Err.Number & ": " & Err.Description
    ImportStage6_PostProcessing = png_Failure
End Function

'Import stage 7 is unique to animated PNGs.  Animated PNGs can contain multiple sub-images within a single
' PNG file.  PD will always load the "default" image (which may or may not be the first frame) using standard
' PD load procedures.  Subsequent frames will only be loaded via this step, after the standard image importer
' has run.
Friend Function ImportStage7_LoadRemainingFrames(ByRef dstImage As pdImage) As PD_PNGResult

    On Error GoTo InternalVBError
    
    ImportStage7_LoadRemainingFrames = png_Success
    
    'Like animated GIFs, animated PNGs support multiple "disposal" methods - basically, what we do
    ' with the running on-screen buffer between frames.  Frames can choose to retain the on-screen
    ' buffer (useful for "layering" new data atop old data) or clear it, or use a special
    ' "restore to previous state" command.
    
    'To account for these various options, we keep a running composite of the on-screen buffer.
    Dim renderBuffer As pdDIB
    Set renderBuffer = New pdDIB
    
    'We have already cached all frame DIBs as part of the initial import step.  Now we simply need to
    ' composite them into full-sized frames, then add the results to the image.
    Dim curFrame As Long
    
    'Start by ensuring all frame DIBs meet internal PD requirements
    For curFrame = 0 To Me.NumAnimationFrames - 1
        m_FrameDIBs(curFrame).SetOriginalFormat PDIF_PNG
        If (Not m_FrameDIBs(curFrame).GetDIBColorDepth = 32) Then ImageImporter.ForceTo32bppMode m_FrameDIBs(curFrame)
        If (Not m_FrameDIBs(curFrame).GetAlphaPremultiplication) Then m_FrameDIBs(curFrame).SetAlphaPremultiplication True
    Next curFrame
    
    'Now we can actually composite the results
    For curFrame = 0 To Me.NumAnimationFrames - 1
        
        Message "Loading page %1 of %2...", CStr(curFrame + 1), Me.NumAnimationFrames, "DONOTLOG"
        
        'Frame 0 is a special case.  In most animated PNGs, it will also be the first frame - but this is
        ' *not* guaranteed.  Note that this has already been handled in a previous step, by pointing the
        ' original DIBs buffer at the first fdAT chunk instead of the IDAT chunk.
        
        'Always start the composite buffer by rendering the first image atop it, unless the first frame
        ' has the (weird but valid) dispose op DisposeOp_PREVIOUS - in which case the spec requires us
        ' to reset the entire buffer to transparent black.
        If (curFrame = 0) Then
            renderBuffer.CreateFromExistingDIB m_FrameDIBs(curFrame)
            If (m_Chunks(m_fcTLIndices(curFrame)).GetFrameDisposal = APNG_DisposeOp_Previous) Then renderBuffer.ResetDIB 0
        End If
        
        'If this frame's disposal method is "PREVIOUS", the spec requires us to perform the following:
        ' "The frame's region of the output buffer is to be reverted to the previous contents before
        ' rendering the next frame."
        ' Because of this, we need a backup copy of the current frame
        Dim frameBackup As pdDIB
        If (m_Chunks(m_fcTLIndices(curFrame)).GetFrameDisposal = APNG_DisposeOp_Previous) Then
            If (frameBackup Is Nothing) Then Set frameBackup = New pdDIB
            frameBackup.CreateFromExistingDIB renderBuffer
        End If
        
        If (curFrame > 0) Then
            
            'Composite the current frame against the render buffer
            Dim xOffset As Long, yOffset As Long
            xOffset = m_Chunks(m_fcTLIndices(curFrame)).GetFrameOffsetX
            yOffset = m_Chunks(m_fcTLIndices(curFrame)).GetFrameOffsetY
            
            Select Case m_Chunks(m_fcTLIndices(curFrame)).GetFrameBlend
                Case APNG_BlendOp_Source
                    GDI.BitBltWrapper renderBuffer.GetDIBDC, xOffset, yOffset, m_FrameDIBs(curFrame).GetDIBWidth, m_FrameDIBs(curFrame).GetDIBHeight, m_FrameDIBs(curFrame).GetDIBDC, 0, 0, vbSrcCopy
                Case APNG_BlendOp_Over
                    m_FrameDIBs(curFrame).AlphaBlendToDC renderBuffer.GetDIBDC, , xOffset, yOffset
            End Select
            
            'Create a blank layer in the receiving image, and copy our DIB into it
            Dim newLayerID As Long, newLayerName As String
            newLayerID = dstImage.CreateBlankLayer
            newLayerName = Layers.GenerateInitialLayerName(vbNullString, vbNullString, True, dstImage, renderBuffer, curFrame)
            dstImage.GetLayerByID(newLayerID).InitializeNewLayer PDL_Image, newLayerName, renderBuffer, True
            
        End If
        
        'Also, write out an offset x/y (if any) and append frame time
        With dstImage.GetLayerByID(newLayerID)
            .SetLayerFrameTimeInMS m_Chunks(m_fcTLIndices(curFrame)).GetFrameDelayMS
            .SetLayerName .GetLayerName & " (" & CStr(.GetLayerFrameTimeInMS()) & "ms)"
            .SetLayerOffsetX 0&
            .SetLayerOffsetY 0&
            .NotifyOfDestructiveChanges
        End With
        
        'Next, prepare a frame rectangle; we need this for calculating a "disposal region" - specifically,
        ' to prepare for the next frame, do we reset or revert the region where this frame was drawn?
        Dim dstSurface As pd2DSurface, srcSurface As pd2DSurface, cBrush As pd2DBrush, disposeRect As RectF
        disposeRect.Left = m_Chunks(m_fcTLIndices(curFrame)).GetFrameOffsetX
        disposeRect.Top = m_Chunks(m_fcTLIndices(curFrame)).GetFrameOffsetY
        If (curFrame = 0) Then
            disposeRect.Width = m_Header.Width + 1
            disposeRect.Height = m_Header.Height + 1
        Else
            disposeRect.Width = m_FrameDIBs(curFrame).GetDIBWidth
            disposeRect.Height = m_FrameDIBs(curFrame).GetDIBHeight
        End If
        
        'Update the running render buffer according to this frame's disposal requirements
        Select Case m_Chunks(m_fcTLIndices(curFrame)).GetFrameDisposal
            
            'nop
            Case APNG_DisposeOp_None
                
            'Fill w/ background color - but *only* the frame rect!
            Case APNG_DisposeOp_Background
                Drawing2D.QuickCreateSolidBrush cBrush, vbBlack, 0
                Drawing2D.QuickCreateSurfaceFromDIB dstSurface, renderBuffer, False
                dstSurface.SetSurfaceCompositing P2_CM_Overwrite
                PD2D.FillRectangleF_FromRectF dstSurface, cBrush, disposeRect
                
            'Fill w/ previous composite data - but *only* the frame rect!
            Case APNG_DisposeOp_Previous
                Drawing2D.QuickCreateSurfaceFromDIB dstSurface, renderBuffer, False
                Drawing2D.QuickCreateSurfaceFromDIB srcSurface, frameBackup, False
                dstSurface.SetSurfaceCompositing P2_CM_Overwrite
                PD2D.DrawSurfaceCroppedF dstSurface, disposeRect.Left, disposeRect.Top, disposeRect.Width, disposeRect.Height, srcSurface, disposeRect.Left, disposeRect.Top
                
        End Select
        
        Set dstSurface = Nothing: Set srcSurface = Nothing: Set cBrush = Nothing
        
        'The current frame DIB is no longer required - free it to reduce memory consumption
        Set m_FrameDIBs(curFrame) = Nothing
            
    Next curFrame
    
    'With all frames loaded, cache any other relevant animation metadata inside the destination image
    dstImage.ImgStorage.AddEntry "animation-loop-count", Trim$(Str$(Me.NumAnimationLoops))
    
    Exit Function
    
InternalVBError:
    InternalError "ImportStage7_LoadRemainingFrames", "internal VB error #" & Err.Number & ": " & Err.Description
    ImportStage7_LoadRemainingFrames = png_Failure
End Function

'APNG support functions
Friend Function IsAnimated() As Boolean
    IsAnimated = Me.HasChunk("acTL")
    If IsAnimated And (Not m_AnimationDataCached) Then CacheAnimationData
End Function

Friend Function NumAnimationFrames() As Long
    If Me.IsAnimated Then
        If (Not m_AnimationDataCached) Then CacheAnimationData
        NumAnimationFrames = m_FrameCount
    Else
        NumAnimationFrames = 1
    End If
End Function

Friend Function NumAnimationLoops() As Long
    If Me.IsAnimated Then
        If (Not m_AnimationDataCached) Then CacheAnimationData
        NumAnimationLoops = m_LoopCount
    Else
        NumAnimationLoops = 1
    End If
End Function

'Cache frame and loop count of an aPNG file
Private Sub CacheAnimationData()
    
    Dim acTLIndex As Long
    acTLIndex = Me.GetIndexOfChunk("acTL")
    
    If (acTLIndex >= 0) Then
        
        Dim origPosition As Long
        origPosition = m_Chunks(acTLIndex).BorrowData().GetPosition()
        
        m_Chunks(acTLIndex).BorrowData().SetPosition 0, FILE_BEGIN
        m_FrameCount = m_Chunks(acTLIndex).BorrowData().ReadLong_BE()
        
        m_LoopCount = m_Chunks(acTLIndex).BorrowData().ReadLong_BE()
        m_Chunks(acTLIndex).BorrowData().SetPosition origPosition, FILE_BEGIN
        
        PDDebug.LogAction "Animated PNG found: " & CStr(m_FrameCount) & " frames"
        
        m_AnimationDataCached = True
        
    End If
    
End Sub

'Simplified wrapper to save a PNG automatically.  I recommend setting outColorType to png_AutoColorType and
' outBitsPerChannel to 0; this allows the class to auto-detect the best output settings for the source
' image's data (and thus produce the smallest output file).
Friend Function SavePNG_ToFile(ByRef dstFile As String, ByRef srcImage As pdDIB, ByRef srcPDImage As pdImage, Optional ByVal outColorType As PD_PNGColorType = png_AutoColorType, Optional ByVal outBitsPerChannel As Long = 0, Optional ByVal cmpLevel As Long = -1, Optional ByRef fullParamString As String = vbNullString, Optional ByVal sourceDIBIsDisposable As Boolean = False) As PD_PNGResult
    
    On Error GoTo CouldNotSaveFile
    
    If (srcImage Is Nothing) Or (LenB(dstFile) = 0) Then
        InternalError "SavePNG", "invalid function parameters"
        SavePNG_ToFile = False
        Exit Function
    End If
    
    'Before proceeding, if the output requires image analysis (due to "automatic" color mode / bit-depth detection,
    ' or the presence of a full PD-compatible parameter string), perform those analyses before proceeding further.
    If (outColorType = png_AutoColorType) Or (outBitsPerChannel = 0) Or (LenB(fullParamString) <> 0) Then
        SetupExportSettings srcImage, srcPDImage, outColorType, outBitsPerChannel, fullParamString
        PDDebug.LogAction "Color type: " & outColorType & ", bpc: " & outBitsPerChannel
    End If
    
    'Failsafe check only; PD enforces safe overwriting in a parent function.
    If Files.FileExists(dstFile) Then Files.FileDelete dstFile
    
    'As usual, a pdStream object is used to simplify the process of saving to file *or* memory
    Dim cStream As pdStream
    Set cStream = New pdStream
    cStream.StartStream PD_SM_FileMemoryMapped, PD_SA_ReadWrite, dstFile, srcImage.GetDIBWidth * srcImage.GetDIBHeight, , OptimizeSequentialAccess
    
    Dim startTime As Currency, firstStartTime As Currency
    VBHacks.GetHighResTime startTime
    firstStartTime = startTime
    
    'Prep a header; this contains standard IHDR information, and storing it module-level prevents us from
    ' needing to pass it to individual functions over-and-over
    With m_Header
        
        .Width = srcImage.GetDIBWidth
        .Height = srcImage.GetDIBHeight
        .BitDepth = outBitsPerChannel
        .ColorType = outColorType
        
        PDDebug.LogAction "Writing PNG color type " & outColorType & " with channel bit-depth " & outBitsPerChannel
        
        'From color type and bit-depth, determine a "friendly" representation of bits-per-pixel
        Select Case outColorType
            Case png_Indexed
                .BitsPerPixel = outBitsPerChannel
            Case png_Greyscale
                .BitsPerPixel = outBitsPerChannel
            Case png_GreyscaleAlpha
                .BitsPerPixel = outBitsPerChannel * 2
            Case png_Truecolor
                .BitsPerPixel = outBitsPerChannel * 3
            Case png_TruecolorAlpha
                .BitsPerPixel = outBitsPerChannel * 4
        End Select
        
        'We never write interlaced PNGs, because interlaced PNGs are much larger and entirely pointless in
        ' the 21st century.
        .Interlaced = False
        
    End With
    
    'Start by writing out a basic PNG header.
    Dim keepGoing As PD_PNGResult
    keepGoing = Me.ExportStep1_StartFile(cStream)
    
    If PNG_DEBUG_VERBOSE Then
        PDDebug.LogAction "PNG export > initialization " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
        
    'All subsequent writes must be written as standalone chunks.  Note that chunks contain more than just "data" -
    ' they also embed length markers, chunk IDs (4-byte ASCII strings), and checksums.  Most of these tedious
    ' sub-tasks are automatically handled by the pdPNGChunk class.
    
    'PNG files must start with a special "IHDR" chunk
    If (keepGoing < png_Failure) Then keepGoing = Me.ExportStep2_WriteIHDR(cStream)
    
    If PNG_DEBUG_VERBOSE Then
        PDDebug.LogAction "PNG export > IHDR creation took " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'Pixel data must be filtered before compression.  We use a temporary pdPNGChunk object to do this for us,
    ' as it already contains a bunch of code related to un-filtering pixel data.
    Dim tmpChunk As pdPNGChunk, filterStrategy As PD_PNG_FilterStrategy
    If (cmpLevel = 0) Then filterStrategy = png_FilterNone Else filterStrategy = png_FilterUndefined
    If (keepGoing < png_Failure) Then keepGoing = Me.ExportStep3_FilterBytes(cStream, srcImage, tmpChunk, fullParamString, filterStrategy, True, sourceDIBIsDisposable)
    
    If PNG_DEBUG_VERBOSE Then
        PDDebug.LogAction "PNG export > byte filtering took " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'Filtered pixel data can now be added in the form of an IDAT chunk
    If (keepGoing < png_Failure) Then keepGoing = Me.ExportStep4_WriteIDAT(cStream, tmpChunk, cmpLevel)
    
    If PNG_DEBUG_VERBOSE Then
        PDDebug.LogAction "PNG export > deflate + write IDAT took " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'PNG files end with a special "IEND" chunk
    If (keepGoing < png_Failure) Then keepGoing = Me.ExportStep5_WriteIEND(cStream)
    
    'The write is finished.  Close the file handle and exit.
    Dim finalFileSize As Long
    finalFileSize = cStream.GetStreamSize()
    
    cStream.StopStream
    Set cStream = Nothing
    
    If PNG_DEBUG_VERBOSE Then
        PDDebug.LogAction "PNG export > IEND + finalize data took " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    PDDebug.LogAction "PNG export took: " & VBHacks.GetTimeDiffNowAsString(firstStartTime)
    PDDebug.LogAction "Final PNG size is " & Files.GetFormattedFileSize(finalFileSize)
    
    SavePNG_ToFile = keepGoing
    
    Exit Function
    
CouldNotSaveFile:
    InternalError "SavePNG_ToFile", "Internal VB error #" & Err.Number & ": " & Err.Description
    SavePNG_ToFile = False
    
End Function

'Simplified wrapper to save a PNG to a byte array.  Unlike the SavePNG_ToFile function, this one returns the size of the
' saved PNG data.  (For perf reasons, it will *not* reallocate dstArray() unless absolutely necessary.)
' I recommend setting outColorType to png_AutoColorType and outBitsPerChannel to 0; this allows the class to auto-detect
' the best output settings for the source image's data (and thus produce the smallest output file).
'
'Returns: 0 if unsuccessful, some other integer >= 0 describing the size of the data in dstArray() if successful.
Friend Function SavePNG_ToMemory(ByRef dstArray() As Byte, ByRef srcImage As pdDIB, ByRef srcPDImage As pdImage, Optional ByVal outColorType As PD_PNGColorType = png_AutoColorType, Optional ByVal outBitsPerChannel As Long = 0, Optional ByVal cmpLevel As Long = -1, Optional ByRef fullParamString As String = vbNullString, Optional ByVal useFilterStrategy As PD_PNG_FilterStrategy = png_FilterUndefined, Optional ByVal sourceDIBIsDisposable As Boolean = False) As Long
    
    On Error GoTo CouldNotSaveFile
    
    If (srcImage Is Nothing) Then
        InternalError "SavePNG", "invalid function parameters"
        SavePNG_ToMemory = False
        Exit Function
    End If
    
    'Before proceeding, if the output requires image analysis (due to "automatic" color mode / bit-depth detection,
    ' or the presence of a full PD-compatible parameter string), perform those analyses before proceeding further.
    If (outColorType = png_AutoColorType) Or (outBitsPerChannel = 0) Or (LenB(fullParamString) <> 0) Then
        SetupExportSettings srcImage, srcPDImage, outColorType, outBitsPerChannel, fullParamString
        PDDebug.LogAction "Color type: " & outColorType & ", bpc: " & outBitsPerChannel
    End If
    
    'As usual, a pdStream object is used to simplify the process of saving to file *or* memory
    Dim cStream As pdStream
    Set cStream = New pdStream
    cStream.StartStream PD_SM_MemoryBacked, PD_SA_ReadWrite, , , , OptimizeSequentialAccess
    
    Dim startTime As Currency, firstStartTime As Currency
    VBHacks.GetHighResTime startTime
    firstStartTime = startTime
    
    'Prep a header; this contains standard IHDR information, and storing it module-level prevents us from
    ' needing to pass it to individual functions over-and-over
    With m_Header
        
        .Width = srcImage.GetDIBWidth
        .Height = srcImage.GetDIBHeight
        .BitDepth = outBitsPerChannel
        .ColorType = outColorType
        
        'From color type and bit-depth, determine a "friendly" representation of bits-per-pixel
        Select Case outColorType
            Case png_Indexed
                .BitsPerPixel = outBitsPerChannel
            Case png_Greyscale
                .BitsPerPixel = outBitsPerChannel
            Case png_GreyscaleAlpha
                .BitsPerPixel = outBitsPerChannel * 2
            Case png_Truecolor
                .BitsPerPixel = outBitsPerChannel * 3
            Case png_TruecolorAlpha
                .BitsPerPixel = outBitsPerChannel * 4
        End Select
        
        'We never write interlaced PNGs, because interlaced PNGs are much larger and entirely pointless in
        ' the 21st century.
        .Interlaced = False
        
    End With
    
    'Start by writing out a basic PNG header.
    Dim keepGoing As PD_PNGResult
    keepGoing = Me.ExportStep1_StartFile(cStream)
    
    If PNG_DEBUG_VERBOSE Then
        PDDebug.LogAction "PNG export > initialization " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
        
    'All subsequent writes must be written as standalone chunks.  Note that chunks contain more than just "data" -
    ' they also embed length markers, chunk IDs (4-byte ASCII strings), and checksums.  Most of these tedious
    ' sub-tasks are automatically handled by the pdPNGChunk class.
    
    'PNG files must start with a special "IHDR" chunk
    If (keepGoing < png_Failure) Then keepGoing = Me.ExportStep2_WriteIHDR(cStream)
    
    If PNG_DEBUG_VERBOSE Then
        PDDebug.LogAction "PNG export > IHDR creation took " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'Pixel data must be filtered before compression.  We use a temporary pdPNGChunk object to do this for us,
    ' as it already contains a bunch of code related to un-filtering pixel data.
    Dim tmpChunk As pdPNGChunk
    If (keepGoing < png_Failure) Then keepGoing = Me.ExportStep3_FilterBytes(cStream, srcImage, tmpChunk, fullParamString, useFilterStrategy, True, sourceDIBIsDisposable)
    
    If PNG_DEBUG_VERBOSE Then
        PDDebug.LogAction "PNG export > byte filtering took " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'Filtered pixel data can now be added in the form of an IDAT chunk
    If (keepGoing < png_Failure) Then keepGoing = Me.ExportStep4_WriteIDAT(cStream, tmpChunk, cmpLevel)
    
    If PNG_DEBUG_VERBOSE Then
        PDDebug.LogAction "PNG export > deflate + write IDAT took " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'PNG files end with a special "IEND" chunk
    If (keepGoing < png_Failure) Then keepGoing = Me.ExportStep5_WriteIEND(cStream)
    
    'The write is finished.  Copy the complete stream into the destination array.
    Dim finalFileSize As Long
    finalFileSize = cStream.GetStreamSize()
    
    If VBHacks.IsArrayInitialized(dstArray) Then
        If (UBound(dstArray) < finalFileSize - 1) Then ReDim dstArray(0 To finalFileSize - 1) As Byte
    Else
        ReDim dstArray(0 To finalFileSize - 1) As Byte
    End If
    
    CopyMemoryStrict VarPtr(dstArray(0)), cStream.Peek_PointerOnly(0), finalFileSize
    
    cStream.StopStream
    Set cStream = Nothing
    
    If PNG_DEBUG_VERBOSE Then
        PDDebug.LogAction "PNG export > IEND + finalize data took " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    PDDebug.LogAction "PNG export took: " & VBHacks.GetTimeDiffNowAsString(firstStartTime)
    PDDebug.LogAction "Final PNG size is " & Files.GetFormattedFileSize(finalFileSize)
    
    If (keepGoing < png_Failure) Then SavePNG_ToMemory = finalFileSize Else SavePNG_ToMemory = 0
    
    Exit Function
    
CouldNotSaveFile:
    InternalError "SavePNG_ToMemory", "Internal VB error #" & Err.Number & ": " & Err.Description
    SavePNG_ToMemory = False
    
End Function

'Simplified wrapper to save an animated PNG automatically.  At present, I don't intend to provide the same
' level of comprehensive output settings that still PNGs support; instead, output settings are largely
' designed to be auto-detected.  This is particularly critical for animated files as they are much more
' size-sensitive than single-frame ones.
Friend Function SaveAPNG_ToFile(ByRef dstFile As String, ByRef srcPDImage As pdImage, Optional ByVal outColorType As PD_PNGColorType = png_AutoColorType, Optional ByVal outBitsPerChannel As Long = 0, Optional ByVal cmpLevel As Long = -1, Optional ByRef fullParamString As String = vbNullString) As PD_PNGResult
    
    On Error GoTo CouldNotSaveFile
    
    If (srcPDImage Is Nothing) Or (LenB(dstFile) = 0) Then
        InternalError "SaveAPNG_ToFile", "invalid function parameters"
        SaveAPNG_ToFile = False
        Exit Function
    End If
    
    'Parse all relevant APNG parameters.  (See the APNG export dialog for details on how these are generated.)
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString fullParamString
    
    'In Nov 2021 an alternate option was provided for export: lossy optimization a la pngquant, but for
    ' animated PNGs.  When this setting is TRUE, PD skips automatic heuristics and instead forcibly
    ' converts the output image to palettized mode (but with PNG's variable alpha support, obviously).
    '
    'This option can greatly reduce file size, but obviously it may incur a significant quality penalty
    ' depending on the input image.  On basic animations that just require variable alpha - like logos or
    ' anything with simple transparency effects, like drop-shadows - the quality difference will be
    ' minimal and you can save a ton of space with this setting.  With complex, varied images like
    ' photo-derived animations or 3D-rendered scenes full of gradients, the quality loss will likely be
    ' unacceptable.  For this reason, I have currently marked this feature as EXPERIMENTAL and may not
    ' enable it in production builds pending further testing.
    '
    'TODO: create a UI that exposes this setting to users
    Const ALLOW_LOSSY_APNG_OPTIMIZATION As Boolean = False
    
    Dim lossyProgressBarOffset As Long
    If ALLOW_LOSSY_APNG_OPTIMIZATION Then lossyProgressBarOffset = srcPDImage.GetNumOfLayers Else lossyProgressBarOffset = 0
    
    'Initialize a progress bar
    ProgressBars.SetProgBarMax srcPDImage.GetNumOfLayers() * 3 + lossyProgressBarOffset
    
    'Unlike static PNGs, PD always enforces strict automatic output color model analysis for APNGs.
    ' (This is a necessity for producing optimized APNGs; user settings can't achieve the same results as
    ' automatic heuristics, alas.)  This value must be set early in the save process, as we will use its
    ' results to determine what animation properties we can/can't use.
    '
    'Note that this stage is fairly time-consuming, as all frames may need to be iterated.
    If (outColorType = png_AutoColorType) Or (outBitsPerChannel = 0) Or (LenB(fullParamString) <> 0) Then
        SetupExportSettings_APNG srcPDImage, outColorType, outBitsPerChannel, fullParamString
        PDDebug.LogAction "Color type: " & outColorType & ", bpc: " & outBitsPerChannel
    End If
    
    If ALLOW_LOSSY_APNG_OPTIMIZATION Then
    
        'Lossy optimization is only relevant if the output image intends to use full 32-bpp mode.
        ' (If the image is already going to use 8-bpp mode, it means the incoming data is 8-bpp
        '  and a lossy translation is unnecessary.)
        If ((outColorType = png_Truecolor) Or (outColorType = png_TruecolorAlpha)) And (outBitsPerChannel >= 8) Then
        
            'The PNG is going to be 32-bpp.  Generate a merged palette for *all* frames (yes *all* frames)
            ' and store it at module-level, while also setting all relevant flags.
            SetupExportSettings_APNG_Lossy srcPDImage, outColorType, outBitsPerChannel, fullParamString
        
        End If
    
    End If
    
    'If you want to simplify debugging, you can force all frames to output as full 32-bpp RGBA.
    ' (Do NOT enable these next two lines in production code; they still produce valid APNG files,
    ' but the file will likely be much larger than it needs to be.)
    'outColorType = png_TruecolorAlpha
    'outBitsPerChannel = 8
    
    'Retrieve loop count from the passed param string (0 = infinite loops)
    Dim loopCount As Long
    loopCount = cParams.GetLong("animation-loop-count", 1)
    
    'Failsafe check only; PD enforces safe overwriting in a parent function.
    If Files.FileExists(dstFile) Then Files.FileDelete dstFile
    
    'As usual, a pdStream object is used to simplify the process of saving to file *or* memory.
    ' For this use-case, we want to write to a memory-mapped file, which provides the best of
    ' both worlds (a persistent file object that we can address with memory pointers).
    Dim cStream As pdStream
    Set cStream = New pdStream
    cStream.StartStream PD_SM_FileMemoryMapped, PD_SA_ReadWrite, dstFile, srcPDImage.Width * srcPDImage.Height, , OptimizeSequentialAccess
    
    'PD's APNG exporter has been extensively profiled, and I've left the timing code in place
    ' because it could be very valuable if users experience issues.
    Dim startTime As Currency, firstStartTime As Currency
    VBHacks.GetHighResTime startTime
    firstStartTime = startTime
    
    'Prep a PNG header; this is a standard IHDR chunk, identical to a static PNG, and storing
    ' it module-level prevents us from needing to pass it to individual functions over-and-over.
    With m_Header
        
        .Width = srcPDImage.Width
        .Height = srcPDImage.Height
        .BitDepth = outBitsPerChannel
        .ColorType = outColorType
        
        PDDebug.LogAction "Writing animated PNG w/ color type " & outColorType & " with channel bit-depth " & outBitsPerChannel
        
        'From color type and bit-depth, determine a "friendly" representation of bits-per-pixel
        Select Case outColorType
            Case png_Indexed
                .BitsPerPixel = outBitsPerChannel
            Case png_Greyscale
                .BitsPerPixel = outBitsPerChannel
            Case png_GreyscaleAlpha
                .BitsPerPixel = outBitsPerChannel * 2
            Case png_Truecolor
                .BitsPerPixel = outBitsPerChannel * 3
            Case png_TruecolorAlpha
                .BitsPerPixel = outBitsPerChannel * 4
        End Select
        
        'We never write interlaced PNGs, because interlaced PNGs are much larger and entirely pointless in
        ' the 21st century, especially for animated PNGs.
        .Interlaced = False
        
    End With
    
    'Write out the magic "PNG" numbers to identify a PNG file
    Dim keepGoing As PD_PNGResult
    keepGoing = Me.ExportStep1_StartFile(cStream)
    
    If PNG_DEBUG_VERBOSE Then
        PDDebug.LogAction "PNG export > initialization " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
        
    'All subsequent writes must be written as standalone chunks.  Note that chunks contain more than
    ' just "data" - they also embed length markers, chunk IDs (4-byte ASCII strings), and CRC32 checksums.
    ' Most of these tedious sub-tasks are automatically handled by the pdPNGChunk class.
    
    'PNG files must start with a special "IHDR" chunk.  (This is identical to static PNGs.)
    ' Write one using the header we prepared earlier.
    If (keepGoing < png_Failure) Then keepGoing = Me.ExportStep2_WriteIHDR(cStream)
    
    If PNG_DEBUG_VERBOSE Then
        PDDebug.LogAction "PNG export > IHDR creation took " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'Now it is time to prepare and write individual animation frames.  Each frame consists of two parts:
    ' an fcTL ("frame control") chunk that describes frame size and timing, and an fdAT ("frame data")
    ' chunk that describes the pixel data.  fdAT chunks are nearly identical to regular IDAT chunks,
    ' and PD's write function is capable of producing both.
    
    'These chunks are unique among all PNG (and APNG) chunks because they both include a "sequence_number"
    ' value - an ascending value that ensures chunks are in the correct order (required because traditional
    ' PNGs only impose a small subset of ordering rules, with optional chunks - which APNG chunks technically
    ' are - being allowed to appear in any order).  This sequence number is individually tracked and
    ' incremented by the fcTL chunk writer, but it must always be reset before writing a new file.
    Dim numFrames As Long
    numFrames = srcPDImage.GetNumOfLayers - 1
    m_FrameSequence = 0
    
    'If we detect two identical back-to-back frames (surprisingly common in GIFs "in the wild"), we will
    ' simply merge their frame times into a single value and remove the duplicate frame.  This reduces
    ' file size "for free", but it requires more complicated tracking as the number of frames may
    ' decrease as optimization proceeds.
    Dim numGoodFrames As Long, lastGoodFrame As Long
    numGoodFrames = 0
    lastGoodFrame = 0
    
    'Retrieve frame delay defaults and overrides from the passed param string
    Dim useFixedFrameDelay As Boolean, frameDelayDefault As Long
    useFixedFrameDelay = cParams.GetBool("use-fixed-frame-delay", False)
    frameDelayDefault = cParams.GetLong("frame-delay-default", 100)
    
    'This (very important) array stores all frame data between optimization and write passes
    Dim allFrames() As PD_APNGFrame
    ReDim allFrames(0 To numFrames) As PD_APNGFrame
    
    'Frame contents will be heavily modified as we optimize them, so we (obviously) don't want to operate
    ' directly on the image's source layers - instead, we make temporary copies of frames as-we-go.
    Dim tmpLayer As pdLayer
    Set tmpLayer = New pdLayer
    
    'Transparency tables are used for a bunch of different optimization settings; they are one-byte arrays
    ' that track transparency data only, making them much more memory-friendly then full RGBA storage.
    Dim trnsTable() As Byte
    
    'Filtering an animated PNG is more time-consuming than filtering a static PNG (obviously), but it
    ' still yields compression dividends... as such, PD still uses a full optimization strategy.
    ' In the future, I may expose this to the user, so I've left the possibility for modification here.
    Dim tmpChunk As pdPNGChunk, filterStrategy As PD_PNG_FilterStrategy
    If (cmpLevel = 0) Then filterStrategy = png_FilterNone Else filterStrategy = png_FilterUndefined
    
    'Tracking individual stage times is very helpful, just as it is with static PNGs
    Dim timeTakenPrep As Currency, timeTakenFilter As Currency, timeTakenDeflate As Currency
    
    'Before proceeding with export, we must perform a pre-processing step to handle a bunch of
    ' optimization tasks (e.g. checking for duplicate frames, cropping relevant frame regions,
    ' determining optimal frame disposal methods, etc).
    
    'This is necessary because PNG files are *extremely* unpleasant to rewind and modify as-you-go,
    ' because chunks must be decompressed, modified, re-compressed then re-checksummed whenever their
    ' contents change.
    
    'As such, we consume a lot of memory during this step, and this could be avoided by writing out
    ' (suboptimal, unoptimized) frames in a single pass - but I consider the current arrangement a
    ' worthwhile tradeoff, as the PNGs it produces are *significantly* smaller (e.g. an order of
    ' magnitude or more).
    
    'As part of optimizing frames, we need to keep a running copy several different frames:
    ' 1) what the current frame looks like right now
    ' 2) what the previous frame looked like
    ' 3) what our current frame looked like before we started fucking with it
    
    'These three copies are important because APNGs provide three different frame disposal mechanisms,
    ' e.g. what you do with the current frame buffer after rendering the current frame, but before
    ' rendering the next one.  We can reset the frame buffer to transparent black, leave it as-is,
    ' or reset it to whatever it looked like *before* the current frame was rendered.  (FYI: these
    ' three mechanisms are taken from animated GIFs, which use an identical system.)  Each of these
    ' "frame disposal methods" have potential optimization trade-offs depending on the current frame's
    ' contents, and we want to find the combination that produces the smallest frame filesize - so we
    ' need to maintain running copies of all three states.
    Dim prevFrame As pdDIB, bufferFrame As pdDIB, curFrameBackup As pdDIB
    Set prevFrame = New pdDIB
    Set bufferFrame = New pdDIB
    Set curFrameBackup = New pdDIB
    
    'We also use a soft reference that we can point at whatever frame DIB we want; its contents are not
    ' maintained between frames...
    Dim refFrame As pdDIB
    
    '...as well as a temporary "test" DIB, which is used to test optimizations that may not pan out
    ' (because they actually harm compression ratio in the end).
    Dim testDIB As pdDIB
    Set testDIB = New pdDIB
    
    'Some parts of the optimization process are not guaranteed to improve file size (e.g. blanking out
    ' duplicate pixels between frames can actually hurt compression if the current frame is very noisy).
    ' To ensure we only apply beneficial optimizations, we test-compress the current frame after certain
    ' potentially problematic optimizations to double-check that our compression ratio improved.  (If it
    ' didn't, we'll roll back the changes we made - see the multiple DIB copies above!)
    
    'To reduce memory churn, we initialize a single worst-case-size buffer in advance, then reuse it
    ' for all compression test runs.
    Dim cmpTestBuffer() As Byte, cmpTestBufferSize As Long
    cmpTestBufferSize = Compression.GetWorstCaseSize(srcPDImage.Width * srcPDImage.Height * 4, cf_Zlib, 1)
    ReDim cmpTestBuffer(0 To cmpTestBufferSize - 1) As Byte
    
    If PNG_DEBUG_VERBOSE Then VBHacks.GetHighResTimeInMS startTime
    
    'It's time to start our second pass through all layers.  (Our first pass happened at the very start,
    ' when we analyzed frames to determine ideal color mode outputs.)
    Dim i As Long
    For i = 0 To numFrames
        
        'This stage can be time-consuming, so we provide UI feedback
        ProgressBars.SetProgBarVal srcPDImage.GetNumOfLayers() + i + lossyProgressBarOffset
        Message "Optimizing animation frame %1 of %2...", i + 1, numFrames + 1, "DONOTLOG"
        
        With allFrames(i)
            
            'Before dealing with pixel data, attempt to retrieve a frame time from the source layer's name.
            ' (If the layer name does not provide a frame time, or if the user has specified a fixed
            ' frame time, this value will be overwritten with their requsted value.)
            Dim finalFrameTime As Long
            finalFrameTime = GetFrameTimeFromLayerName(srcPDImage.GetLayerByIndex(i).GetLayerName, 0)
            If (useFixedFrameDelay Or (finalFrameTime = 0)) Then finalFrameTime = frameDelayDefault
            .frameTime = finalFrameTime
            
            'Remaining parameters are contingent on optimization passes; for now, populate with safe
            ' default parameters (e.g parameters that produce a function APNG even if optimization fails).
            ' Individual optimizations will modify these settings as necessary.
            .frameBlend = APNG_BlendOp_Source
            .frameDisposal = APNG_DisposeOp_None
            .frameIsDuplicateOrEmpty = False
            .rectOfInterest.Left = 0
            .rectOfInterest.Top = 0
            .rectOfInterest.Width = srcPDImage.Width
            .rectOfInterest.Height = srcPDImage.Height
            
        End With
        
        'Convert the source layer to an image-sized, zero-offsets, all-transforms-applied layer,
        ' then cache a local copy of it.  (This MUST be lossless for the source layer!)
        Set tmpLayer = New pdLayer
        tmpLayer.CopyExistingLayer srcPDImage.GetLayerByIndex(i)
        tmpLayer.ConvertToNullPaddedLayer srcPDImage.Width, srcPDImage.Height, True
        curFrameBackup.CreateFromExistingDIB tmpLayer.GetLayerDIB
        
        'Ensure we have a target DIB to operate on; the final, optimized frame will be stored here.
        Set allFrames(i).frameDIB = New pdDIB
        
        'The first frame in the file must always be full-size, per the spec.  We are not allowed to
        ' optimize it.
        If (i = 0) Then
        
            'Cache the temporary layer DIB as-is; it serves as both the first frame in the animation,
            ' and the fallback "static" image for decoders that don't understand APNGs.
            allFrames(i).frameDIB.CreateFromExistingDIB tmpLayer.GetLayerDIB
            
            'Initialize the frame buffer to be the same as the first frame...
            bufferFrame.CreateFromExistingDIB tmpLayer.GetLayerDIB
            
            '...and initialize the "previous" frame buffer to pure transparency (this is the
            ' recommended behavior per the spec, on the assumption that the buffer always begins as
            ' transparent black)
            prevFrame.CreateBlank tmpLayer.GetLayerDIB.GetDIBWidth, tmpLayer.GetLayerDIB.GetDIBHeight, 32, 0, 0
            
            'Finally, mark this as the "last good frame" (in case subsequent frames are duplicates
            ' of this one, we'll need to refer back to the "last good frame" index)
            lastGoodFrame = i
            numGoodFrames = i + 1
        
        'If this is *not* the first frame, there are many ways we can optimize frame contents.
        Else
        
            'First, we want to figure out how much of this frame needs to be used at all.  If this frame
            ' shares contents with the frame that just appeared (an extremely common use-case for
            ' animations!), we can simply crop out any frame borders that are identical to the previous
            ' frame.  (Said another way: APNGs only require the *first* frame to be full size - subsequent
            ' ones can be any size you want, so long as they remain within the first frame's borders
            ' as a whole.)
            
            '(If this check fails, it means that this frame is 100% identical to the frame that came
            ' before it.)
            Dim dupArea As RectF
            If DIBs.GetRectOfInterest_Overlay(tmpLayer.GetLayerDIB, bufferFrame, dupArea) Then
                
                'This frame contains at least one unique pixel, so it needs to be added to the file.
                
                'Before proceeding further, let's compare this frame to one other buffer - specifically, the
                ' frame buffer as it appeared *before* the previous frame was painted.  Like GIFs, APNGs
                ' define a disposal op called "DISPOSE_OP_PREVIOUS", which tells the previous frame to "undo"
                ' its rendering when it's done.  On certain frames and/or animation styles, this may allow
                ' for better compression if this frame is more identical to a frame further back in sequence.
                Dim prevFrameArea As RectF
                
                'As before, ensure that the previous check succeeded.  If it fails, it means this frame
                ' is 100% identical the the frame that preceded the previous frame.  Rather than encode this
                ' frame at all, we can simply store a lone transparent pixel and paint it "over" the
                ' corresponding frame buffer - maximum file size savings!
                If DIBs.GetRectOfInterest_Overlay(tmpLayer.GetLayerDIB, prevFrame, prevFrameArea) Then
                
                    'With an overlap rectangle calculated for both cases, determine a "winner"
                    ' (where "winner" equals "least number of pixels), store its frame rectangle,
                    ' and then mark the *previous* frame's disposal op accordingly.  (That's important -
                    ' disposal ops describe what you do *after* the current frame is painted, so if
                    ' we want a certain frame buffer state *before* rendering this frame, we set it
                    ' using the disposal op of the *previous* frame - an obnoxious quirk left over
                    ' from GIFs, presumably).
                    
                    'The frame *before* the frame *before* this one is smallest...
                    If (prevFrameArea.Width * prevFrameArea.Height) < (dupArea.Width * dupArea.Height) Then
                        Set refFrame = prevFrame
                        allFrames(i).rectOfInterest = prevFrameArea
                        allFrames(lastGoodFrame).frameDisposal = APNG_DisposeOp_Previous
                        
                    'The frame immediately preceding this one is smallest...
                    Else
                        Set refFrame = bufferFrame
                        allFrames(i).rectOfInterest = dupArea
                        allFrames(lastGoodFrame).frameDisposal = APNG_DisposeOp_None
                    End If
                    
                    'We now have the smallest possible rectangle that defines this frame, while accounting for
                    ' both DISPOSAL_OP_NONE and DISPOSAL_OP_PREVIOUS.
                    
                    'We have one more potential crop operation we can do, and it involves the third APNG
                    ' disposal op (DISPOSAL_OP_BACKGROUND).  This disposal op asks the previous frame to
                    ' erase itself completely after rendering.  For animations with large transparent borders,
                    ' it may actually be best to crop the current frame according to its transparent borders,
                    ' then use the "erase" disposal op before displaying it, thus forgoing any connection
                    ' whatsoever to preceding frames.
                    Dim trnsRect As RectF
                    If DIBs.GetRectOfInterest(tmpLayer.GetLayerDIB, trnsRect) Then
                    
                        'If this frame is smaller than the previous "winner", switch to using this op instead.
                        If (trnsRect.Width * trnsRect.Height) < (allFrames(i).rectOfInterest.Width * allFrames(i).rectOfInterest.Height) Then
                            allFrames(lastGoodFrame).frameDisposal = APNG_DisposeOp_Background
                            allFrames(i).frameBlend = APNG_BlendOp_Source
                            allFrames(i).rectOfInterest = trnsRect
                        End If
                        
                        'Crop the "winning" region into a separate DIB, and store it as the formal pixel buffer
                        ' for this frame.
                        With allFrames(i).rectOfInterest
                            allFrames(i).frameDIB.CreateBlank Int(.Width), Int(.Height), 32, 0, 0
                            GDI.BitBltWrapper allFrames(i).frameDIB.GetDIBDC, 0, 0, Int(.Width), Int(.Height), tmpLayer.GetLayerDIB.GetDIBDC, Int(.Left), Int(.Top), vbSrcCopy
                        End With
                    
                    'This weird (but valid) branch means that the current frame is 100% transparent.  For this
                    ' special case, request that the previous frame erase the running buffer, then store a 1 px
                    ' transparent pixel for this frame.
                    Else
                        
                        allFrames(i).frameDIB.CreateBlank 1, 1, 32, 0, 0
                        With allFrames(i).rectOfInterest
                            .Left = 0
                            .Top = 0
                            .Width = 1
                            .Height = 1
                        End With
                        allFrames(i).frameBlend = APNG_BlendOp_Source
                        allFrames(lastGoodFrame).frameDisposal = APNG_DisposeOp_None
                        
                    End If
                    
                    'Because the current frame came from a premultiplied source, we can safely mark it
                    ' as premultiplied as well.
                    If (Not allFrames(i).frameDIB Is Nothing) Then allFrames(i).frameDIB.SetInitialAlphaPremultiplicationState True
                    
                    'If the previous frame is not being blanked, we have additional optimization strategies
                    ' to attempt.  (If, however, the previous frame *is* being blanked, we are done with
                    ' preprocessing.  Any further gains will have to come from the PNG encoder itself.)
                    If (allFrames(lastGoodFrame).frameDisposal <> APNG_DisposeOp_Background) Then
                        
                        'The next optimization we want to attempt is duplicate pixel blanking, which takes
                        ' pixels in the current frame that are identical to the previous frame and turns them
                        ' transparent, allowing the previous frame to "show through" in those regions (instead
                        ' of storing all that pixel data again in the current frame).
                        
                        'The more pixels we can turn transparent, the better the resulting buffer will compress,
                        ' but note that there are two major caveats to this optimization.  Specifically:
                        
                        '1) The previous frame must use DisposeOp_None or DisposeOp_Previous.  If it erases the
                        '   frame (DisposeOp_Background), the previous frame's pixels aren't around for us to
                        '   use.  (We caught this possibility with the If statement above, FYI.)
                        '2) Because this approach requires us to alphablend the current frame "over" the previous
                        '   one (instead of simply *replacing* the previous frame's contents with this one), we
                        '   need to ensure there are no transparency mismatches.  Specifically, if this frame has
                        '   variable transparency where the previous frame DOES NOT, we can't use this frame
                        '   blanking strategy (as this frame's transparent regions will allow the previous frame to
                        '   "show through" where it shouldn't).
                        
                        'Because (1) has already been taken care of by the frame disposal If/Then statement above,
                        ' we now want to address case (2).
                        If DIBs.RetrieveTransparencyTable(allFrames(i).frameDIB, trnsTable) Then
                        If DIBs.CheckAlpha_DuplicatePixels(refFrame, allFrames(i).frameDIB, trnsTable, Int(allFrames(i).rectOfInterest.Left), Int(allFrames(i).rectOfInterest.Top)) Then
                            
                            'This frame contains variable alpha transparency where the previous frame
                            ' does not. This means that we cannot pixel-blank this frame, and we *must*
                            ' specify replace mode for pixels in the current frame (to ensure our
                            ' transparent pixels *overwrite* the previous frame's non-transparent ones,
                            ' instead of simply "overlaying" atop whatever was there before).
                            allFrames(i).frameBlend = APNG_BlendOp_Source
                            
                        'The current frame can be safely alpha-blended "over the top" of the previous one
                        Else
                            
                            'This frame is a candidate for frame differentials!
                            
                            'Note that some esoteric color mode settings may *still* prevent us from calculating
                            ' frame differentials (e.g. if an image has a full 256-color palette, there's no room
                            ' left for a transparent index without bumping the color model up to full RGBA output,
                            ' which would likely harm compression more than it helps).  This state will be flagged
                            ' by the color model analyzer at the very start of this function; check that flag now,
                            ' and if found, skip frame differentials entirely.
                            If m_apngTrnsAvailable Then
                                
                                'A transparent index is available, meaning we have a mechanism for "erasing" parts
                                ' of this frame so that the previous frame can show through.  (Note that the exact
                                ' mechanism varies by color model - RGB output may use a transparent color, while
                                ' paletted images use a transparent index.  Those details don't matter here.)
                                
                                'Before proceeding, let's get a rough idea of the current frame's entropy.  If
                                ' subsequent optimizations increase entropy (and thus decrease compression ratio),
                                ' we'll just revert to the current frame as-is.
                                
                                'An easy (and fast) way to estimate entropy is to just compress the current frame!
                                ' Better compression ratios correlate with lower source entropy, and this is
                                ' faster than a formal entropy calculation (in VB6 code, anyway).
                                
                                '(This is also where use our persistent compression buffer; note that we don't
                                ' need to clear or prep it in any way - the compression engine will overwrite
                                ' whatever it needs to.)
                                Dim initSize As Long, initCmpSize As Long
                                initCmpSize = cmpTestBufferSize
                                initSize = allFrames(i).frameDIB.GetDIBStride * allFrames(i).frameDIB.GetDIBHeight
                                Plugin_libdeflate.CompressPtrToPtr VarPtr(cmpTestBuffer(0)), initCmpSize, allFrames(i).frameDIB.GetDIBPointer, initSize, 1, cf_Zlib
                                
                                'With current entropy established (well... "estimated"), we're next going to try
                                ' blanking out any pixels that are identical between this frame and the current frame
                                ' buffer (as calculated in a previous step).  On many animations, this will create
                                ' large patches of pure transparency that compress *brilliantly* - but note that very
                                ' noisy images - like full-color images originally converted w/dithering to animated
                                ' GIFs - this will produce equally noisy results, which can actually *harm* our
                                ' compression ratio. This is why we estimated entropy for the untouched frame data
                                ' (above), and why we're gonna perform all our tests on a temporary frame copy (in
                                ' case we have to throw the copy away).
                                testDIB.CreateFromExistingDIB allFrames(i).frameDIB
                                
                                If DIBs.RetrieveTransparencyTable(testDIB, trnsTable) Then
                                If DIBs.ApplyAlpha_DuplicatePixels(testDIB, refFrame, trnsTable, allFrames(i).rectOfInterest.Left, allFrames(i).rectOfInterest.Top) Then
                                If DIBs.ApplyTransparencyTable(testDIB, trnsTable) Then
                                
                                    'The frame differential was produced successfully.  See if it compresses better
                                    ' than the original, untouched frame did.
                                    Dim testSize As Long
                                    testSize = cmpTestBufferSize
                                    Plugin_libdeflate.CompressPtrToPtr VarPtr(cmpTestBuffer(0)), testSize, testDIB.GetDIBPointer, initSize, 1, cf_Zlib
                                    
                                    'This frame compressed better!  Use it instead of the original frame, and set
                                    ' the frame blend to OVER (which will allow pixels from the previous frame to
                                    ' "show through" this one, regardless of which disposal op is used).
                                    If (testSize < initCmpSize) Then
                                        allFrames(i).frameDIB.CreateFromExistingDIB testDIB
                                        allFrames(i).frameBlend = APNG_BlendOp_Over
                                    
                                    'If this frame compressed worse (or identically), we can simply leave current
                                    ' frame settings as they are.
                                    End If
                                
                                '/end failsafe "retrieve and apply duplicate pixel test" checks
                                End If
                                End If
                                End If
                                
                            '/end m_apngTrnsAvailable = False
                            End If
                        
                        '/end "previous frame is opaque where this frame is transparent"
                        End If
                        End If
                    
                    '/end "previous frame is NOT being blanked, so we can attempt to optimize frame diffs"
                    End If
                    
                'This (rare) branch means that the current frame is identical to the animation as it appeared
                ' *before* the previous frame was rendered.  (This is not unheard of, especially for blinking-
                ' or spinning-style animations.)  A great, cheap optimization is to just ask the previous
                ' frame to dispose of itself using the DISPOSE_OP_PREVIOUS method (which restores the frame
                ' buffer to whatever it was *before* the previous frame was rendered), then store this frame
                ' as a transparent 1-px PNG.  This effectively gives us a copy of the frame two-frames-previous
                ' "for free".
                Else
                    
                    'Ensure transparency is available
                    If m_apngTrnsAvailable Then
                        
                        allFrames(i).frameDIB.CreateBlank 1, 1, 32, 0, 0
                        With allFrames(i).rectOfInterest
                            .Left = 0
                            .Top = 0
                            .Width = 1
                            .Height = 1
                        End With
                        allFrames(i).frameBlend = APNG_BlendOp_Source
                        allFrames(lastGoodFrame).frameDisposal = APNG_DisposeOp_Previous
                    
                    'Damn - this optimization "cheat" won't work, because we don't have access to transparency
                    ' in the current image.  Unfortunately, the full frame will need to be written out to file :(
                    Else
                        allFrames(i).frameDIB.CreateFromExistingDIB curFrameBackup
                    End If
                    
                End If
                
            'If the GetRectOfInterest() check failed, it means this frame is 100% identical to the
            ' frame that preceded it.  Rather than optimize this frame, let's just delete it from
            ' the animation and merge its frame time into the previous frame.
            Else
                allFrames(i).frameIsDuplicateOrEmpty = True
                allFrames(lastGoodFrame).frameTime = allFrames(lastGoodFrame).frameTime + allFrames(i).frameTime
                Set allFrames(i).frameDIB = Nothing
            End If
            
            'This frame is now optimized as well as we can possibly optimize it.
            
            'Before moving to the next frame, create backup copies of the buffer frames *we* were handed.
            ' The next frame can request that we reset our state to this frame, which may be closer to
            ' their frame's contents (and thus compress better).
            If (allFrames(lastGoodFrame).frameDisposal = APNG_DisposeOp_None) Then
                prevFrame.CreateFromExistingDIB bufferFrame
            ElseIf (allFrames(lastGoodFrame).frameDisposal = APNG_DisposeOp_Background) Then
                prevFrame.ResetDIB 0
            
            'We don't have to cover the case of DISPOSE_OP_PREVIOUS, as that's the state the prevFrame
            ' DIB is already in!
            'Else

            End If
            
            'Overwrite the *current* frame buffer with an (untouched) copy of this frame, as it appeared
            ' before we applied optimizations to it.
            bufferFrame.CreateFromExistingDIB curFrameBackup
            
            'If this frame is valid (e.g. not a duplicate of the previous frame), increment our current
            ' "good frame" count, and mark this frame as the "last good" index.
            If (Not allFrames(i).frameIsDuplicateOrEmpty) Then
                numGoodFrames = numGoodFrames + 1
                lastGoodFrame = i
            End If
        
        'i !/= 0 branch
        End If
    
    'Next frame
    Next i
    
    If PNG_DEBUG_VERBOSE Then
        PDDebug.LogAction "APNG export > optimization pass took " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'Clear all optimization-related objects, as they are no longer needed
    Set prevFrame = Nothing
    Set bufferFrame = Nothing
    Set curFrameBackup = Nothing
    Set refFrame = Nothing
    Set testDIB = Nothing
    Erase cmpTestBuffer
    
    'With all frames optimized, we can finally write an acTL - or "animation control" - chunk.
    ' This marks the file as an animated PNG, and it lists the number of frames in the file (which may
    ' have changed as a result of the optimization step, above, which finds and removes duplicate frames).
    keepGoing = ExportStep2a_WriteacTL(cStream, numGoodFrames, loopCount)
    
    For i = 0 To numFrames
        
        'Again, this step takes time - keep the UI active and relaying updates
        ProgressBars.SetProgBarVal srcPDImage.GetNumOfLayers() * 2 + i + lossyProgressBarOffset
        Message "Saving animation frame %1 of %2...", i + 1, numFrames + 1, "DONOTLOG"
        
        'Skip all frames marked as duplicates
        If (Not allFrames(i).frameIsDuplicateOrEmpty) Then
            
            If PNG_DEBUG_VERBOSE Then VBHacks.GetHighResTimeInMS startTime
            
            'Write this frame's animation metadata out (before writing any pixel data; order is
            ' critical here!)
            With allFrames(i)
                ExportStep3a_WritefcTL cStream, .rectOfInterest, .frameTime, .frameDisposal, .frameBlend
            End With
            
            If PNG_DEBUG_VERBOSE Then
                timeTakenPrep = timeTakenPrep + (VBHacks.GetHighResTimeInMSEx - startTime)
                VBHacks.GetHighResTimeInMS startTime
            End If
            
            'Next, we need to prepare the pixel data for this layer.  The precise steps involved in this
            ' vary depending on optimization settings.
            
            'If the saved PNG color mode is "true color", *and* we are using a transparent color flag,
            ' we need to convert the alpha channel of the incoming image to use that flag color to
            ' represent transparent pixels.
            If (outColorType = png_Truecolor) And m_apngTrnsAvailable Then
                DIBs.RetrieveTransparencyTable allFrames(i).frameDIB, trnsTable
                DIBs.ApplyBinaryTransparencyTableColor allFrames(i).frameDIB, trnsTable, RGB(m_TrnsValueR, m_TrnsValueG, m_TrnsValueB), m_CompositingColor
            End If
            
            'Unpremultiply the source DIB
            allFrames(i).frameDIB.SetAlphaPremultiplication False
            
            'Convert the optimized frame to a filtered PNG byte stream.  (Filtering reorders scanline data
            ' according to one of several optimization strategies; see the PNG spec for details.)
            Set tmpChunk = Nothing
            If (keepGoing < png_Failure) Then keepGoing = Me.ExportStep3_FilterBytes(cStream, allFrames(i).frameDIB, tmpChunk, fullParamString, filterStrategy, (i = 0), True)
            
            'Free our copy of this animation frame as it's no longer needed
            Set allFrames(i).frameDIB = Nothing
            
            If PNG_DEBUG_VERBOSE Then
                timeTakenFilter = timeTakenFilter + (VBHacks.GetHighResTimeInMSEx - startTime)
                VBHacks.GetHighResTimeInMS startTime
            End If
            
            'Filtered pixel data can now be finalized in the form of a compressed IDAT (first frame) or fdAT
            ' (all subsequent frames) chunk - this is where we lean on libdeflate to do the actual DEFLATE-ing.
            If (keepGoing < png_Failure) Then keepGoing = Me.ExportStep4_WriteIDAT(cStream, tmpChunk, cmpLevel, (i > 0))
            
            If PNG_DEBUG_VERBOSE Then
                timeTakenDeflate = timeTakenDeflate + (VBHacks.GetHighResTimeInMSEx - startTime)
                VBHacks.GetHighResTimeInMS startTime
            End If
            
        End If
            
    Next i
    
    'With all frames added, we can now finalize the file
    ProgressBars.SetProgBarVal ProgressBars.GetProgBarMax()
    Message "Finalizing image..."
    
    If PNG_DEBUG_VERBOSE Then
        PDDebug.LogAction "APNG export > prep took " & Format$(timeTakenPrep, "#,#0.0") & " ms"
        PDDebug.LogAction "APNG export > filter took " & Format$(timeTakenFilter, "#,#0.0") & " ms"
        PDDebug.LogAction "APNG export > deflate took " & Format$(timeTakenDeflate, "#,#0.0") & " ms"
        VBHacks.GetHighResTime startTime
    End If
    
    'Just like regular PNG files, APNGs end with a special "IEND" chunk
    If (keepGoing < png_Failure) Then keepGoing = Me.ExportStep5_WriteIEND(cStream)
    
    'The write is finished.  Close the file handle and exit.
    Dim finalFileSize As Long
    finalFileSize = cStream.GetStreamSize()
    
    cStream.StopStream
    Set cStream = Nothing
    
    If PNG_DEBUG_VERBOSE Then
        PDDebug.LogAction "APNG export > IEND + finalize data took " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    PDDebug.LogAction "APNG export took: " & VBHacks.GetTimeDiffNowAsString(firstStartTime)
    PDDebug.LogAction "Final APNG size is " & Files.GetFormattedFileSize(finalFileSize)
    
    'Free any remaining UI elements
    ProgressBars.SetProgBarVal 0
    ProgressBars.ReleaseProgressBar
    
    SaveAPNG_ToFile = keepGoing
    
    Exit Function
    
CouldNotSaveFile:
    InternalError "SaveAPNG_ToFile", "Internal VB error #" & Err.Number & ": " & Err.Description
    SaveAPNG_ToFile = False
    
End Function

'Streaming approach to saving APNGs.  Call this function to start the stream, then pass successive frames
' at will.  This function is perf-optimized for 24-bpp data.  32-bpp data will not work.  (TODO??)
' Note that the APNGs produced by this function may be larger than the standard SaveAPNG function.
' This function does not optimize as aggressively in order to keep performance high.  For the absolute
' smallest file, use this function to produce the initial PNG, then re-load the file into PhotoDemon
' and save it again.  That will produce the smallest possible file.
Friend Function SaveAPNG_Streaming_Start(ByRef dstFile As String, ByVal pxWidth As Long, ByVal pxHeight As Long) As PD_PNGResult
    
    On Error GoTo CouldNotStartAPNG
    
    If (LenB(dstFile) = 0) Then
        InternalError "SaveAPNG_Streaming_Start", "invalid function parameters"
        SaveAPNG_Streaming_Start = False
        Exit Function
    End If
    
    'Reset frame trackers
    m_FrameCount = 0
    m_FrameSequence = 0
    
    'For now, this function always assumes 24-bpp output.  This is fast for us to produce,
    ' and it enables easy per-frame heuristics.
    Dim outColorType As PD_PNGColorType
    outColorType = png_Truecolor
    
    Dim outBitsPerChannel As Long
    outBitsPerChannel = 8
    
    'If the target file exists, kill it before starting
    If Files.FileExists(dstFile) Then Files.FileDelete dstFile
    
    'As usual, a pdStream object is used to simplify the process of saving to file *or* memory.
    ' For this use-case, we want to write to a memory-mapped file, which provides the best of
    ' both worlds (a persistent file object that we can address with memory pointers).
    Set m_Stream = New pdStream
    m_Stream.StartStream PD_SM_FileMemoryMapped, PD_SA_ReadWrite, dstFile, pxWidth * pxHeight * 4, , OptimizeSequentialAccess
    
    'This APNG exporter has been extensively profiled, and I've left the timing code in place
    ' because it could prove valuable again if users experience issues.
    Dim startTime As Currency, firstStartTime As Currency
    VBHacks.GetHighResTime startTime
    firstStartTime = startTime
    
    'Prep a PNG header; this is a standard IHDR chunk, identical to a static PNG, and storing
    ' it module-level prevents us from needing to pass it to individual functions over-and-over.
    With m_Header
        
        .Width = pxWidth
        .Height = pxHeight
        .BitDepth = outBitsPerChannel
        .ColorType = outColorType
        
        PDDebug.LogAction "Starting animated PNG w/ color type " & outColorType & " with channel bit-depth " & outBitsPerChannel
        
        'From color type and bit-depth, determine a "friendly" representation of bits-per-pixel
        Select Case outColorType
            Case png_Indexed
                .BitsPerPixel = outBitsPerChannel
            Case png_Greyscale
                .BitsPerPixel = outBitsPerChannel
            Case png_GreyscaleAlpha
                .BitsPerPixel = outBitsPerChannel * 2
            Case png_Truecolor
                .BitsPerPixel = outBitsPerChannel * 3
            Case png_TruecolorAlpha
                .BitsPerPixel = outBitsPerChannel * 4
        End Select
        
        'We never write interlaced PNGs, because interlaced PNGs are much larger and entirely pointless in
        ' the 21st century, especially for animated PNGs.
        .Interlaced = False
        
    End With
    
    'Write out the magic "PNG" numbers to identify a PNG file
    Dim keepGoing As PD_PNGResult
    keepGoing = Me.ExportStep1_StartFile(m_Stream)
    
    If PNG_DEBUG_VERBOSE Then
        PDDebug.LogAction "APNG export > initialization " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
        
    'All subsequent writes must be written as standalone chunks.  Note that chunks contain more than
    ' just "data" - they also embed length markers, chunk IDs (4-byte ASCII strings), and CRC32 checksums.
    ' Most of these tedious sub-tasks are automatically handled by the pdPNGChunk class.
    
    'PNG files must start with a special "IHDR" chunk.  (This is identical to static PNGs.)
    ' Write one using the header we prepared earlier.
    If (keepGoing < png_Failure) Then keepGoing = Me.ExportStep2_WriteIHDR(m_Stream)
    
    If PNG_DEBUG_VERBOSE Then
        PDDebug.LogAction "APNG export > IHDR creation took " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'Next comes the acTL - or "animation control" - chunk.  This marks the file as an animated PNG,
    ' and it lists the number of frames in the file (which we won't know until the animation is complete),
    ' and the loop count.  Flag the current file position so we can return to it later, but advance the
    ' stream normally so that we're in position to write the first frame.
    m_flagActlPos = m_Stream.GetPosition()
    If (keepGoing < png_Failure) Then keepGoing = ExportStep2a_WriteacTL(m_Stream, 0, 0)
    
    If PNG_DEBUG_VERBOSE Then
        PDDebug.LogAction "APNG export > acTL creation took " & VBHacks.GetTimeDiffNowAsString(startTime)
        VBHacks.GetHighResTime startTime
    End If
    
    'Now it is time to prepare and write individual animation frames.  The caller sends these
    ' as they arrive, so we can't do anything yet!
    SaveAPNG_Streaming_Start = keepGoing
    
    Exit Function
    
CouldNotStartAPNG:
    InternalError "SaveAPNG_Streaming_Start", "Internal VB error #" & Err.Number & ": " & Err.Description
    SaveAPNG_Streaming_Start = False
    
End Function

'Streaming approach to saving APNGs.  After a stream is started, call this function to pass
' individual frames.  Optimization is handled automatically - just make sure every pdDIB object
' you pass has the same width, height, and color-depth.  If they don't, this will crash!
' Also, this function does not currently handle duplicate frames gracefully; that's up to the
' client to handle (where it makes more sense, *and* conserves memory).
Friend Function SaveAPNG_Streaming_Frame(ByRef srcDIB As pdDIB, ByVal msTimeWhenFrameArrived As Currency, ByVal pngCompressionLevel As Long) As PD_PNGResult
    
    'Each APNG frame consists of two parts:
    ' - an fcTL ("frame control") chunk that describes frame size and timing
    ' - an fdAT ("frame data") chunk that describes the pixel data.  (fdAT chunks are nearly
    '   identical to regular IDAT chunks.)
    
    'The fcTL chunk is problematic because it requires us to know things like frame delay,
    ' but we don't know that until the caller passed us the *next* frame in line.  As such,
    ' we don't actually write it at present; instead, we just add dummy padding where it will
    ' go in the future.
    
    'If this is the first frame, initialize all buffers
    If (m_FrameCount = 0) Then
        ReDim m_fcTLIndices(0 To INIT_STREAMING_BUFFER - 1) As Long
        ReDim m_fdATIndices(0 To INIT_STREAMING_BUFFER - 1) As Long
        ReDim m_streamingFrames(0 To INIT_STREAMING_BUFFER - 1) As PD_APNGFrame
    
    'If this is *not* the first frame, ensure all buffers are big enough for the next frame
    Else
        If (m_FrameCount > UBound(m_fcTLIndices)) Then
            ReDim Preserve m_fcTLIndices(0 To m_FrameCount * 2 - 1) As Long
            ReDim Preserve m_fdATIndices(0 To m_FrameCount * 2 - 1) As Long
            ReDim Preserve m_streamingFrames(0 To m_FrameCount * 2 - 1) As PD_APNGFrame
        End If
    End If
    
    'Make a note of the current m_fcTL position in the file
    m_fcTLIndices(m_FrameCount) = m_Stream.GetPosition()
    
    'Next, populate what we can of the m_fcTL chunk for this frame
    With m_streamingFrames(m_FrameCount)
        .frameSeq = m_FrameSequence
        m_FrameSequence = m_FrameSequence + 1
        .frameTime = msTimeWhenFrameArrived
        .frameBlend = APNG_BlendOp_Source
        .frameDisposal = APNG_DisposeOp_None
        .frameIsDuplicateOrEmpty = False
        .rectOfInterest.Left = 0
        .rectOfInterest.Top = 0
        .rectOfInterest.Width = srcDIB.GetDIBWidth
        .rectOfInterest.Height = srcDIB.GetDIBHeight
        Set .frameDIB = New pdDIB
    End With
    
    'Write out a dummy fcTL buffer; we'll come back in and fill this in later.
    ' (The size of an fcTL chunk is hard-coded; 26-bytes + 12-bytes of required PNG chunk data)
    Const fcTLChunkSize As Long = 26 + 12
    m_Stream.WritePadding fcTLChunkSize
    
    Dim keepGoing As PD_PNGResult, tmpChunk As pdPNGChunk
    
    'We also need a soft reference that we can point at whatever frame DIB we want; its contents
    ' are not maintained between frames...
    Dim refFrame As pdDIB
    
    '...as well as a temporary "test" DIB, which is used to test optimizations that may not pan out
    ' (because they actually harm compression ratio in the end).
    Dim testDIB As pdDIB
    Set testDIB = New pdDIB
    
    'Transparency tables are used for a bunch of different optimization settings; they are one-byte arrays
    ' that track transparency data only, making them much more memory-friendly then full RGBA storage.
    Dim trnsTable() As Byte
    
    'If this is the first frame, we always want to write it out as-is
    If (m_FrameCount = 0) Then
        
        'Before writing this first frame, we want to find an unused color that we can use to flag
        ' transparent pixels in frames (which may exist after we perform frame differencing).
        
        'We start with magic magenta and use it if unused in this first frame.
        Dim cColorCount As pdColorCount
        DIBs.GetDIBColorCountObject srcDIB, cColorCount, False
        
        m_TrnsValueR = 254
        m_TrnsValueG = 1
        m_TrnsValueB = 239
        
        'Next, query the object for an unused color, and start the search with the image's original
        ' transparent color (if any; if it didn't supply one, we'll just start at black).
        If cColorCount.DoesColorExist(m_TrnsValueR, m_TrnsValueG, m_TrnsValueB) Then
            
            'The original transparent color now appears in the image, which means we can't use it for export.
            ' We need to find a new one.  (This will only return false if all 16.7 million colors are already
            ' in-use in the target image.  This prevents tRNS usage, so per-frame blanking gets deactivated.)
            m_apngTrnsAvailable = cColorCount.GetUnusedColor(m_TrnsValueR, m_TrnsValueG, m_TrnsValueB)
            
        'The original transparent color does not appear in the image!  Use it as our transparent color.
        Else
            m_apngTrnsAvailable = True
        End If
        
        'Filter the full DIB
        m_WriteTrns = m_apngTrnsAvailable
        keepGoing = Me.ExportStep3_FilterBytes(m_Stream, srcDIB, tmpChunk, vbNullString, png_FilterOptimal, True, False)
        
        'Pass the full DIB as-is out to the IDAT writer
        If (keepGoing < png_Failure) Then keepGoing = Me.ExportStep4_WriteIDAT(m_Stream, tmpChunk, pngCompressionLevel, False)
        
        'Next, we need to initialize our various optimization DIBs.
        
        'Initialize the frame buffer to be the same as the first frame...
        Set m_apngBufferFrame = New pdDIB
        m_apngBufferFrame.CreateFromExistingDIB srcDIB
        
        '...and initialize the "previous" frame buffer to pure transparency (this is the
        ' recommended behavior per the spec, on the assumption that the buffer always begins as
        ' transparent black)
        Set m_apngPrevFrame = New pdDIB
        m_apngPrevFrame.CreateBlank srcDIB.GetDIBWidth, srcDIB.GetDIBHeight, 32, 0, 0
        
        Set m_apngCurFrameBackup = New pdDIB
        
        'Some parts of the optimization process are not guaranteed to improve file size (e.g. blanking out
        ' duplicate pixels between frames can actually hurt compression if the current frame is very noisy).
        ' To ensure we only apply beneficial optimizations, we test-compress the current frame after certain
        ' potentially problematic optimizations to double-check that our compression ratio improved.  (If it
        ' didn't, we'll roll back the changes we made - see the multiple DIB copies above!)
        
        'To reduce memory churn, we initialize a single worst-case-size buffer when the first frame arrives,
        ' then reuse it for all subsequent compression test runs.
        m_cmpTestBufferSize = Compression.GetWorstCaseSize(srcDIB.GetDIBWidth * srcDIB.GetDIBHeight * 4, cf_Zlib, 1)
        ReDim m_cmpTestBuffer(0 To m_cmpTestBufferSize - 1) As Byte
        
        'Finally, mark this as the "last good frame" (in case subsequent frames are duplicates
        ' of this one, we'll need to refer back to the "last good frame" index)
        m_lastGoodFrame = 0
        m_FrameCount = 1
        
    'If this is *not* the first frame, we want to run heuristics on it before embedding it.
    Else
        
        'Make a backup of the current frame, before we perform any heuristics
        m_apngCurFrameBackup.CreateFromExistingDIB srcDIB
        
        'First, we want to figure out how much of this frame needs to be used at all.  If this frame
        ' shares contents with the frame that just appeared (an extremely common use-case for
        ' animations!), we can simply crop out any frame borders that are identical to the previous
        ' frame.  (Said another way: APNGs only require the *first* frame to be full size - subsequent
        ' ones can be any size you want, so long as they remain within the first frame's borders.)
        
        '(If this check fails, it means that this frame is 100% identical to the frame that came
        ' before it.)
        Dim dupArea As RectF
        If DIBs.GetRectOfInterest_Overlay(srcDIB, m_apngBufferFrame, dupArea) Then
        
            'This frame contains at least one unique pixel, so it needs to be added to the file.
            
            'Before proceeding further, let's compare this frame to one other buffer - specifically, the
            ' frame buffer as it appeared *before* the previous frame was painted.  Like GIFs, APNGs
            ' define a disposal op called "DISPOSE_OP_PREVIOUS", which tells the previous frame to "undo"
            ' its rendering when it's done.  On certain frames and/or animation styles, this may allow
            ' for better compression if this frame is more identical to a frame further back in sequence.
            Dim prevFrameArea As RectF
            
            'As before, ensure that the previous check succeeded.  If it fails, it means this frame
            ' is 100% identical the the frame that preceded the previous frame.  Rather than encode this
            ' frame at all, we can simply store a lone transparent pixel and paint it "over" the
            ' corresponding frame buffer - maximum file size savings!
            If DIBs.GetRectOfInterest_Overlay(srcDIB, m_apngPrevFrame, prevFrameArea) Then
            
                'With an overlap rectangle calculated for both cases, determine a "winner"
                ' (where "winner" equals "least number of pixels").  Store the "winner's" frame
                ' rect, then mark the *previous* frame's disposal op accordingly.  (This is
                ' important - disposal ops describe what you do *after* the current frame is
                ' painted, so if we want a certain frame buffer state *before* rendering this
                ' frame, we set it using the disposal op of the *previous* frame - an obnoxious
                ' quirk left over from GIFs, presumably).
                
                'The frame *before* the frame *before* this one is smallest...
                If (prevFrameArea.Width * prevFrameArea.Height) < (dupArea.Width * dupArea.Height) Then
                    Set refFrame = m_apngPrevFrame
                    m_streamingFrames(m_FrameCount).rectOfInterest = prevFrameArea
                    m_streamingFrames(m_lastGoodFrame).frameDisposal = APNG_DisposeOp_Previous
                    
                'The frame immediately preceding this one is smallest...
                Else
                    Set refFrame = m_apngBufferFrame
                    m_streamingFrames(m_FrameCount).rectOfInterest = dupArea
                    m_streamingFrames(m_lastGoodFrame).frameDisposal = APNG_DisposeOp_None
                End If
                
                'We now have the smallest possible rectangle that defines this frame, while accounting for
                ' both DISPOSAL_OP_NONE and DISPOSAL_OP_PREVIOUS.
                
                'We have one more potential crop operation we can do, and it involves the third APNG
                ' disposal op (DISPOSAL_OP_BACKGROUND).  This disposal op asks the previous frame to
                ' erase itself completely after rendering.  For animations with large transparent borders,
                ' it may actually be best to crop the current frame according to its transparent borders,
                ' then use the "erase" disposal op before displaying it, thus forgoing any connection
                ' whatsoever to preceding frames.
                Dim trnsRect As RectF
                If DIBs.GetRectOfInterest(srcDIB, trnsRect) Then
                
                    'If this frame is smaller than the previous "winner", switch to using this op instead.
                    If (trnsRect.Width * trnsRect.Height) < (m_streamingFrames(m_FrameCount).rectOfInterest.Width * m_streamingFrames(m_FrameCount).rectOfInterest.Height) Then
                        m_streamingFrames(m_lastGoodFrame).frameDisposal = APNG_DisposeOp_Background
                        m_streamingFrames(m_FrameCount).frameBlend = APNG_BlendOp_Source
                        m_streamingFrames(m_FrameCount).rectOfInterest = trnsRect
                    End If
                    
                    'Crop the "winning" region into a separate DIB, and store it as the formal pixel buffer
                    ' for this frame.
                    With m_streamingFrames(m_FrameCount).rectOfInterest
                        m_streamingFrames(m_FrameCount).frameDIB.CreateBlank Int(.Width), Int(.Height), 32, 0, 0
                        GDI.BitBltWrapper m_streamingFrames(m_FrameCount).frameDIB.GetDIBDC, 0, 0, Int(.Width), Int(.Height), srcDIB.GetDIBDC, Int(.Left), Int(.Top), vbSrcCopy
                    End With
                
                'This weird (but valid) branch means that the current frame is 100% transparent.  For this
                ' special case, request that the previous frame erase the running buffer, then store a 1 px
                ' transparent pixel for this frame.
                Else
                    
                    m_streamingFrames(m_FrameCount).frameDIB.CreateBlank 1, 1, 32, 0, 0
                    With m_streamingFrames(m_FrameCount).rectOfInterest
                        .Left = 0
                        .Top = 0
                        .Width = 1
                        .Height = 1
                    End With
                    m_streamingFrames(m_FrameCount).frameBlend = APNG_BlendOp_Source
                    m_streamingFrames(m_lastGoodFrame).frameDisposal = APNG_DisposeOp_None
                    
                End If
                
                'Because the current frame came from a premultiplied source, we can safely mark it
                ' as premultiplied as well.
                If (Not m_streamingFrames(m_FrameCount).frameDIB Is Nothing) Then m_streamingFrames(m_FrameCount).frameDIB.SetInitialAlphaPremultiplicationState True
                
                'If the previous frame is not being blanked, we have additional optimization strategies
                ' to attempt.  (If, however, the previous frame *is* being blanked, we are done with
                ' preprocessing.  Any further gains will have to come from the PNG encoder itself.)
                If (m_streamingFrames(m_lastGoodFrame).frameDisposal <> APNG_DisposeOp_Background) Then
                
                    'The next optimization we want to attempt is duplicate pixel blanking, which takes
                    ' pixels in the current frame that are identical to the previous frame and turns them
                    ' transparent, allowing the previous frame to "show through" in those regions (instead
                    ' of storing all that pixel data again in the current frame).
                    
                    'The more pixels we can turn transparent, the better the resulting buffer will compress,
                    ' but note that there are two major caveats to this optimization.  Specifically:
                    
                    '1) The previous frame must use DisposeOp_None or DisposeOp_Previous.  If it erased the
                    '   frame (DisposeOp_Background), the previous frame's pixels aren't around for us to
                    '   use.  (We caught this possibility with the If statement above, FYI.)
                    '2) Because this approach requires us to alphablend the current frame "over" the previous
                    '   one (instead of simply *replacing* the previous frame's contents with this one), we
                    '   need to ensure there are no transparency mismatches.  Specifically, if this frame has
                    '   variable transparency where the previous frame DOES NOT, we can't use this frame
                    '   blanking strategy (as this frame's transparent regions will allow the previous frame to
                    '   "show through" where it shouldn't).
                    
                    'Because (1) has already been taken care of by the frame disposal If/Then statement above,
                    ' we now want to address case (2).
                    If DIBs.RetrieveTransparencyTable(m_streamingFrames(m_FrameCount).frameDIB, trnsTable) Then
                    If DIBs.CheckAlpha_DuplicatePixels(refFrame, m_streamingFrames(m_FrameCount).frameDIB, trnsTable, Int(m_streamingFrames(m_FrameCount).rectOfInterest.Left), Int(m_streamingFrames(m_FrameCount).rectOfInterest.Top)) Then
                        
                        'This frame contains variable alpha transparency where the previous frame does not.
                        ' This means the previous frame *must* be blanked.  Skip frame differentials entirely.
                        m_streamingFrames(m_FrameCount).frameDisposal = APNG_DisposeOp_Background
                        
                    'The current frame can be safely alpha-blended "over the top" of the previous one
                    Else
                        
                        'This frame is a candidate for frame differentials!
                        
                        'Note that some esoteric color mode settings may *still* prevent us from calculating
                        ' frame differentials (e.g. if the previous frame is enormous and uses all 16.7m
                        ' colors).  This state will be flagged by the color model analyzer that runs on the
                        ' first frame of the stream; check that flag now, and if found, skip frame
                        ' differentials entirely.
                        If m_apngTrnsAvailable Then
                            
                            'A transparent index is available, meaning we have a mechanism for "erasing" parts
                            ' of this frame so that the previous frame can show through.
                            
                            'Before proceeding, let's get a rough idea of the current frame's entropy.  If
                            ' subsequent optimizations increase entropy (and thus decrease compression ratio),
                            ' we'll just revert to the current frame as-is.
                            
                            'An easy (and fast) way to estimate entropy is to just compress the current frame!
                            ' Better compression ratios correlate with lower source entropy, and this is
                            ' faster than a formal entropy calculation (in VB6 code, anyway).
                            
                            '(This is also where use our persistent compression buffer; note that we don't
                            ' need to clear or prep it in any way - the compression engine will overwrite
                            ' whatever it needs to.)
                            Dim initSize As Long, initCmpSize As Long
                            initCmpSize = m_cmpTestBufferSize
                            initSize = m_streamingFrames(m_FrameCount).frameDIB.GetDIBStride * m_streamingFrames(m_FrameCount).frameDIB.GetDIBHeight
                            Plugin_libdeflate.CompressPtrToPtr VarPtr(m_cmpTestBuffer(0)), initCmpSize, m_streamingFrames(m_FrameCount).frameDIB.GetDIBPointer, initSize, 1, cf_Zlib
                            
                            'With current entropy established (well... "estimated"), we're next going to try
                            ' blanking out any pixels that are identical between this frame and the current frame
                            ' buffer (as calculated in a previous step).  On many animations, this will create
                            ' large patches of pure transparency that compress *brilliantly* - but note that very
                            ' noisy images can produce equally noisy results, which can actually *harm* our
                            ' compression ratio. This is why we estimated entropy for the untouched frame data
                            ' (above), and why we're gonna perform all our tests on a temporary frame copy (in
                            ' case we have to throw the copy away).
                            testDIB.CreateFromExistingDIB m_streamingFrames(m_FrameCount).frameDIB
                            
                            If DIBs.RetrieveTransparencyTable(testDIB, trnsTable) Then
                            If DIBs.ApplyAlpha_DuplicatePixels(testDIB, refFrame, trnsTable, m_streamingFrames(m_FrameCount).rectOfInterest.Left, m_streamingFrames(m_FrameCount).rectOfInterest.Top) Then
                            If DIBs.ApplyTransparencyTable(testDIB, trnsTable) Then
                                
                                'The frame differential was produced successfully.  See if it compresses better
                                ' than the original, untouched frame did.
                                Dim testSize As Long
                                testSize = m_cmpTestBufferSize
                                Plugin_libdeflate.CompressPtrToPtr VarPtr(m_cmpTestBuffer(0)), testSize, testDIB.GetDIBPointer, initSize, 1, cf_Zlib
                                
                                'This frame compressed better!  Use it instead of the original frame, and set
                                ' the frame blend to OVER (which will allow pixels from the previous frame to
                                ' "show through" this one, regardless of which disposal op is used).
                                If (testSize < initCmpSize) Then
                                    m_streamingFrames(m_FrameCount).frameDIB.CreateFromExistingDIB testDIB
                                    m_streamingFrames(m_FrameCount).frameBlend = APNG_BlendOp_Over
                                
                                'If this frame compressed worse (or identically), we can simply leave current
                                ' frame settings as they are.
                                End If
                            
                            '/end failsafe "retrieve and apply duplicate pixel test" checks
                            End If
                            End If
                            End If
                            
                        '/end m_apngTrnsAvailable = False
                        End If
                        
                    '/end "previous frame is opaque where this frame is transparent"
                    End If
                    End If
                        
                '/end "previous frame is NOT being blanked, so we can attempt to optimize frame diffs"
                End If
            
            'This (rare) branch means that the current frame is identical to the animation as it appeared
            ' *before* the previous frame was rendered.  (This is not unheard of, especially for blinking-
            ' or spinning-style animations.)  A great, cheap optimization is to just ask the previous
            ' frame to dispose of itself using the DISPOSE_OP_PREVIOUS method (which restores the frame
            ' buffer to whatever it was *before* the previous frame was rendered), then store this frame
            ' as a transparent 1-px PNG.  This effectively gives us a copy of the frame two-frames-previous
            ' "for free".
            Else
                
                'Ensure transparency is available
                If m_apngTrnsAvailable Then
                    
                    m_streamingFrames(m_FrameCount).frameDIB.CreateBlank 1, 1, 32, 0, 0
                    m_streamingFrames(m_FrameCount).frameDIB.SetInitialAlphaPremultiplicationState True
                    With m_streamingFrames(m_FrameCount).rectOfInterest
                        .Left = 0
                        .Top = 0
                        .Width = 1
                        .Height = 1
                    End With
                    m_streamingFrames(m_FrameCount).frameBlend = APNG_BlendOp_Source
                    m_streamingFrames(m_lastGoodFrame).frameDisposal = APNG_DisposeOp_Previous
                
                'Damn - this optimization "cheat" won't work, because we don't have access to transparency
                ' in the current image.  Unfortunately, the full frame will need to be written out to file :(
                Else
                    m_streamingFrames(m_FrameCount).frameDIB.CreateFromExistingDIB m_apngCurFrameBackup
                End If
                
            'end duplicate frame check against previous buffer
            End If
        
        'If the GetRectOfInterest() check failed, it means this frame is 100% identical to the
        ' frame that preceded it.  Rather than optimize this frame, let's just delete it from
        ' the animation and merge its frame time into the previous frame.
        Else
            m_streamingFrames(m_FrameCount).frameIsDuplicateOrEmpty = True
            Set m_streamingFrames(m_FrameCount).frameDIB = Nothing
        End If
        
        'This frame is now optimized as well as we can possibly optimize it.
        
        'Before moving to the next frame, create backup copies of the buffer frames *we* were handed.
        ' The next frame can request that we reset our state to this frame, which may be closer to
        ' their frame's contents (and thus compress better).
        If (m_streamingFrames(m_lastGoodFrame).frameDisposal = APNG_DisposeOp_None) Then
            m_apngPrevFrame.CreateFromExistingDIB m_apngBufferFrame
        ElseIf (m_streamingFrames(m_lastGoodFrame).frameDisposal = APNG_DisposeOp_Background) Then
            m_apngPrevFrame.ResetDIB 0
        
        'We don't have to cover the case of DISPOSE_OP_PREVIOUS, as that's the state the prevFrame
        ' DIB is already in!
        'Else

        End If
        
        'Overwrite the *current* frame buffer with an (untouched) copy of this frame, as it appeared
        ' before we applied optimizations to it.
        m_apngBufferFrame.CreateFromExistingDIB m_apngCurFrameBackup
        
        'If this frame is valid (e.g. not a duplicate of the previous frame), increment our current
        ' "good frame" count, and mark this frame as the "last good" index.
        If (Not m_streamingFrames(m_FrameCount).frameIsDuplicateOrEmpty) Then
            
            'If the saved PNG color mode is "true color", *and* we are using a transparent color flag,
            ' we need to convert the alpha channel of the incoming image to use that flag color to
            ' represent transparent pixels.
            If m_apngTrnsAvailable Then
                DIBs.RetrieveTransparencyTable m_streamingFrames(m_FrameCount).frameDIB, trnsTable
                DIBs.ApplyBinaryTransparencyTableColor m_streamingFrames(m_FrameCount).frameDIB, trnsTable, RGB(m_TrnsValueR, m_TrnsValueG, m_TrnsValueB), m_CompositingColor
            End If
            
            'Unpremultiply the source DIB
            m_streamingFrames(m_FrameCount).frameDIB.SetAlphaPremultiplication False
            
            'Write this frame to the destination file.  (At the end of the stream, we will loop back
            ' through the file and populate all frame data segments with their final values, but pixel
            ' values are already ready to go.)
            
            'Note that we also use a faster frame filter strategy while in the IDE; this is mostly to reduce
            ' annoyances while testing this interface.
            Dim curFrameStrategy As PD_PNG_FilterStrategy
            If OS.IsProgramCompiled Then curFrameStrategy = png_FilterOptimal Else curFrameStrategy = png_FilterFast
            keepGoing = Me.ExportStep3_FilterBytes(m_Stream, m_streamingFrames(m_FrameCount).frameDIB, tmpChunk, vbNullString, curFrameStrategy, False, True)
            If (keepGoing < png_Failure) Then keepGoing = Me.ExportStep4_WriteIDAT(m_Stream, tmpChunk, pngCompressionLevel, True)
            
            'Make a note of where the pointer is for the next frame *after* the IDAT/FDAT is written
            m_fdATIndices(m_FrameCount) = m_Stream.GetPosition()
            
            'Increment frame count
            m_lastGoodFrame = m_FrameCount
            m_FrameCount = m_FrameCount + 1
            
        Else
            keepGoing = png_Success
        End If
        
    'End first vs subsequent frames
    End If
    
    SaveAPNG_Streaming_Frame = keepGoing
        
End Function

'Streaming approach to saving APNGs.  After all frames have been added to a stream, call this function to
' finalize the file.  Note that the file will need to be traversed again to embed final frame data
' (as we can't know things like frame times until the end, due to interframe optimizations).
Friend Function SaveAPNG_Streaming_Stop(Optional ByVal loopCount As Long = 0) As PD_PNGResult
    
    Dim keepGoing As PD_PNGResult
    keepGoing = png_Success
    
    'Free a bunch of intermediary frame bits that we no longer need
    Set m_apngPrevFrame = Nothing
    Set m_apngBufferFrame = Nothing
    Set m_apngCurFrameBackup = Nothing
    Erase m_cmpTestBuffer
    
    'Rewind to the very start of the stream and note the final frame count
    m_Stream.SetPosition m_flagActlPos
    
    'Write a "real" acTL chunk
    If (keepGoing < png_Failure) Then keepGoing = ExportStep2a_WriteacTL(m_Stream, m_FrameCount, loopCount)
    
    'Next, we want to iterate through all frames and write out their final fcTL results
    Dim i As Long, curFrameTime As Long
    For i = 0 To m_FrameCount - 1
        
        'Move the file pointer to the position we saved for this frame
        m_Stream.SetPosition m_fcTLIndices(i)
        
        'Reset the frame sequence value
        m_FrameSequence = m_streamingFrames(i).frameSeq
        
        'Write out the frame control chunk
        If (keepGoing < png_Failure) Then
            
            'The final frame is written differently, since it doesn't represent a difference between frames
            If (i < m_FrameCount - 1) Then
                curFrameTime = m_streamingFrames(i + 1).frameTime - m_streamingFrames(i).frameTime
            
            'Display the last frame for some arbitrary amount of time (currently 3 seconds)
            Else
                curFrameTime = 3000
            End If
            
            Me.ExportStep3a_WritefcTL m_Stream, m_streamingFrames(i).rectOfInterest, curFrameTime, m_streamingFrames(i).frameDisposal, m_streamingFrames(i).frameBlend
            
        End If
        
    Next i
    
    'Move the pointer to the end of the file
    m_Stream.SetPosition m_fdATIndices(m_FrameCount - 1)
    
    'With all frames written, we can close out the file with a termination chunk
    If (keepGoing < png_Failure) Then keepGoing = Me.ExportStep5_WriteIEND(m_Stream)
    
    'Free the underlying stream object
    m_Stream.StopStream
    Set m_Stream = Nothing
    
    SaveAPNG_Streaming_Stop = keepGoing
    
End Function
    
'Streaming approach to saving APNGs.  If something goes wrong and you want to cancel, call this function.
' Note that the file is 100% guaranteed to be invalid (as it will lack a trailing termination chunk),
' so you *will* need to delete the destination file after calling this.
Friend Function SaveAPNG_Streaming_Cancel() As PD_PNGResult
    If (Not m_Stream Is Nothing) Then m_Stream.StopStream
    Set m_Stream = Nothing
    SaveAPNG_Streaming_Cancel = png_Success
End Function

'Prior to exporting a PNG, this function needs to be called to setup a bunch of image-specific export settings.
' Specifically, it'll analyze the source image to determine things like appropriate output color mode and
' bit-depth settings.
Private Sub SetupExportSettings(ByRef srcImage As pdDIB, ByRef srcPDImage As pdImage, ByRef outColorType As PD_PNGColorType, ByRef outBitsPerChannel As Long, ByRef fullParamString As String)
    
    'In 99.9% of cases, PD needs to analyze the output image and automatically determine the "best" output settings -
    ' e.g. the combination of PNG properties that produces the smallest output file.
    
    'As such, we want to check for the "do everything automatically" case up front, and switch to a dedicated,
    ' optimized function as early as possible.
    
    'Prep a parameter parser for the source XML.
    ' (See the PNG export dialog for details on how the XML string is generated.)
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString fullParamString
    
    'Some color-depth-specific parameters are stored in their own XML packet. (This is because they are handled
    ' by a standalone color depth control that's shared between multiple export dialogs.)
    Dim cParamsDepth As pdSerialize
    Set cParamsDepth = New pdSerialize
    cParamsDepth.SetParamString cParams.GetString("png-color-depth", , True)
    
    Dim outputColorModel As String, outputAlphaModel As String
    outputColorModel = cParamsDepth.GetString("cd-color-model", "auto", True)
    outputAlphaModel = cParamsDepth.GetString("cd-alpha-model", "auto", True)
    
    Dim autoColorModeActive As Boolean, autoTransparencyModeActive As Boolean
    autoColorModeActive = Strings.StringsEqual(outputColorModel, "auto", True)
    autoTransparencyModeActive = Strings.StringsEqual(outputAlphaModel, "auto", True)
    
    'If fully automatic settings have been requested, launch the auto-analyzer and exit when it finishes
    If autoColorModeActive And autoTransparencyModeActive Then
        SetupExportSettings_Auto srcImage, srcPDImage, outColorType, outBitsPerChannel, fullParamString
        Exit Sub
    End If
    
    'If we're still here, the user wants us to forcibly convert the image data to some specific format.
    ' This requires different handling, as the source data will likely need to be destructively modified
    ' to fit the user's request.
    
    'First, see if the "use original file settings" option was clicked.  If it was, we'll need
    ' to simply retrieve the original PNG file settings and use those for export.
    Dim useOrigMode As Boolean, origColorType As PD_PNGColorType
    origColorType = png_AutoColorType
    If (Not srcPDImage Is Nothing) Then
        useOrigMode = Strings.StringsEqual(cParamsDepth.GetString("cd-color-model", "auto", True), "original", True)
        If useOrigMode Then origColorType = srcPDImage.ImgStorage.GetEntry_Long("png-color-type", png_AutoColorType)
    End If
    
    'Retrieve some miscellaneous data from the param strings; these settings may not be used in all
    ' output modes (e.g. transparent color has limited usage, depending on the output alpha mode).
    If useOrigMode Then
        If srcPDImage.ImgStorage.DoesKeyExist("png-background-color") Then
            m_BkgdColor = srcPDImage.ImgStorage.GetEntry_Long("png-background-color")
        Else
            m_BkgdColor = cParamsDepth.GetLong("png-background-color", vbWhite)
        End If
    Else
        m_BkgdColor = cParamsDepth.GetLong("png-background-color", vbWhite)
    End If
    
    m_CompositingColor = m_BkgdColor
    m_EmbedBkgdColor = cParams.GetBool("png-create-bkgd", False)
    
    'A weird user may choose "automatic" mode for color or alpha mode, but not both.
    ' If they've done this, we need to analyze the image to determine proper automatic
    ' settings for either mode.
    Dim currentAlphaStatus As PD_ALPHA_STATUS, isGrayscale As Boolean
    If (autoColorModeActive Or autoTransparencyModeActive) Then
        Export_AnalyzeImage srcImage, m_ImgColorCount, isGrayscale, currentAlphaStatus, m_ImgPalette
    End If
    
    'If for some reason the user selected auto color mode but *not* auto transparency mode,
    ' set a corresponding color mode using the results of the previous scan.
    Dim outputHDR As Boolean, outputGrayscale As Boolean, outputMonochrome As Boolean
    Dim outputIndexed As Boolean, outputPaletteSize As Long
    
    If autoColorModeActive Then
        outputGrayscale = isGrayscale
        outputHDR = False
        outputMonochrome = False
        outputIndexed = False
        If (m_ImgColorCount <= 256) Then outputPaletteSize = m_ImgColorCount Else outputPaletteSize = 256
        m_ImgMaxColorCount = outputPaletteSize
    Else
        
        'If the caller wants us to use original file settings, pull those instead of relying
        ' on the image's current state.
        If useOrigMode Then
        
            outputGrayscale = (origColorType = png_Greyscale) Or (origColorType = png_GreyscaleAlpha)
            outputIndexed = (origColorType = png_Indexed)
            
            'Pull the original file's palette
            Dim tmpPalette As pdPalette
            If srcPDImage.HasOriginalPalette Then
                
                srcPDImage.GetOriginalPalette tmpPalette
                If (Not tmpPalette Is Nothing) Then
                    
                    'A palette exists!  Copy it, then immediately forcibly convert the source
                    ' image to use that palette.
                    m_ImgMaxColorCount = tmpPalette.GetPaletteColorCount()
                    outputPaletteSize = m_ImgMaxColorCount
                    tmpPalette.CopyPaletteToArray m_ImgPalette
                    Palettes.ApplyPaletteToImage_KDTree srcImage, m_ImgPalette, True
                    m_ImgPaletteSafe = True
                    
                End If
                
            Else
                outputPaletteSize = 256
            End If
            
            outputMonochrome = (outputGrayscale And (srcPDImage.GetOriginalColorDepth() = 1))
            outputHDR = (outputGrayscale And srcPDImage.GetOriginalColorDepth >= 16) Or ((Not outputGrayscale) And (srcPDImage.GetOriginalColorDepth > 32))
            
        Else
            outputGrayscale = Strings.StringsEqual(outputColorModel, "gray", True)
            outputPaletteSize = cParamsDepth.GetLong("cd-palette-size", 256)
        End If
        
    End If
    
    'Interlacing is currently forcibly disabled, as it is unsupported in major web browsers (the images will
    ' still load but they are not displayed progressively) and it tends to greatly increase file size.
    'Dim pngInterlacing As Boolean
    'pngInterlacing = cParams.GetBool("PNGInterlacing", False)
    
    'PD supports multiple alpha output modes; some of these modes (like "binary" alpha, which consists of only
    ' 0 or 255 values), require additional settings.  We always retrieve all values, even if we won't need to
    ' use them given the user's requested output settings.
    Dim outputAlphaCutoff As Long, outputAlphaColor As Long
    outputAlphaCutoff = cParamsDepth.GetLong("cd-alpha-cutoff", PD_DEFAULT_ALPHA_CUTOFF)
    outputAlphaColor = cParamsDepth.GetLong("cd-alpha-color", vbMagenta)
    
    'Look for any requests to forcibly output a specific color or grayscale color depth
    Dim outputColorDepthName As String
    If (Not useOrigMode) Then
        If outputGrayscale Then
            outputColorDepthName = cParamsDepth.GetString("cd-gray-depth", "gray-standard")
            outputHDR = Strings.StringsEqual(outputColorDepthName, "gray-hdr", True)
            outputMonochrome = Strings.StringsEqual(outputColorDepthName, "gray-monochrome", True)
        Else
            outputColorDepthName = cParamsDepth.GetString("cd-color-depth", "color-standard")
            outputHDR = Strings.StringsEqual(outputColorDepthName, "color-hdr", True)
            outputIndexed = Strings.StringsEqual(outputColorDepthName, "color-indexed", True)
            If outputIndexed Then m_ImgMaxColorCount = outputPaletteSize Else m_ImgMaxColorCount = 256
        End If
    End If
    
    'Finally, pull the requested alpha mode from the user's settings; if "auto" output mode was requested,
    ' we'll simply use the output mode detected by the auto-analysis function, above
    Dim outputAlphaMode As PD_ALPHA_STATUS
    If autoTransparencyModeActive Then
        outputAlphaMode = currentAlphaStatus
    Else
        
        'If the caller specified "use original file settings", construct an alpha model
        ' matching the original file
        If useOrigMode Then
        
            'For PNGs with full alpha channels, we want to enable full alpha channel output
            If (origColorType = png_GreyscaleAlpha) Or (origColorType = png_TruecolorAlpha) Then
                outputAlphaMode = PDAS_ComplicatedAlpha
            Else
                
                'If the file used some other form of transparency, assume binary transparency here
                If srcPDImage.GetOriginalAlpha Then
                    outputAlphaMode = PDAS_BinaryAlpha
                    outputAlphaCutoff = 127
                Else
                    outputAlphaMode = PDAS_NoAlpha
                End If
                
            End If
            
        Else
            
            If Strings.StringsEqual(outputAlphaModel, "full", True) Then
                outputAlphaMode = PDAS_ComplicatedAlpha
            
            'For non-standard alpha modes, we also want to use the compositing color specified by the color
            ' depth parameter string.
            ElseIf Strings.StringsEqual(outputAlphaModel, "none", True) Then
                outputAlphaMode = PDAS_NoAlpha
                m_CompositingColor = cParamsDepth.GetLong("cd-matte-color", m_BkgdColor)
            ElseIf Strings.StringsEqual(outputAlphaModel, "by-cutoff", True) Then
                outputAlphaMode = PDAS_BinaryAlpha
                m_CompositingColor = cParamsDepth.GetLong("cd-matte-color", m_BkgdColor)
            ElseIf Strings.StringsEqual(outputAlphaModel, "by-color", True) Then
                outputAlphaMode = PDAS_NewAlphaFromColor
                m_CompositingColor = cParamsDepth.GetLong("cd-matte-color", m_BkgdColor)
            Else
                outputAlphaMode = PDAS_ComplicatedAlpha
            End If
            
        End If
        
    End If
    
    'Some color modes may require us to produce palettes or custom transparency tables.
    Dim trnsTable() As Byte, cCount As pdColorCount
    
    'We now know enough (finally) to construct an output color model and suggested bit-depth.  Whew!
    
    'First, images without alpha data
    If (outputAlphaMode = PDAS_NoAlpha) Then
    
        If outputGrayscale Then
            outColorType = png_Greyscale
            If outputHDR Then
                outBitsPerChannel = 16
            ElseIf outputMonochrome Then
                outBitsPerChannel = 1
                m_ImgMaxColorCount = 2
            Else
                m_ImgMaxColorCount = outputPaletteSize
                If (outputPaletteSize = 2) Then
                    outBitsPerChannel = 1
                ElseIf (outputPaletteSize = 4) Then
                    outBitsPerChannel = 2
                ElseIf (outputPaletteSize = 16) Then
                    outBitsPerChannel = 4
                Else
                    outBitsPerChannel = 8
                End If
            End If
        Else
            outColorType = png_Truecolor
            If outputHDR Then
                outBitsPerChannel = 16
            ElseIf outputIndexed Then
                outColorType = png_Indexed
                m_ImgMaxColorCount = outputPaletteSize
                If (outputPaletteSize <= 2) Then
                    outBitsPerChannel = 1
                ElseIf (outputPaletteSize <= 4) Then
                    outBitsPerChannel = 2
                ElseIf (outputPaletteSize <= 16) Then
                    outBitsPerChannel = 4
                Else
                    outBitsPerChannel = 8
                End If
            Else
                outBitsPerChannel = 8
            End If
        End If
    
    'Complicated (full) alpha
    ElseIf (outputAlphaMode = PDAS_ComplicatedAlpha) Then
        
        If outputGrayscale Then
        
            outColorType = png_GreyscaleAlpha
            If outputHDR Then outBitsPerChannel = 16 Else outBitsPerChannel = 8
            
            'If the user's done something weird like request a set number of gray shades,
            ' we must apply that now, as the grayscale + alpha function won't modify
            ' an image to a non-standard number of gray shades.
            If outputMonochrome Then outputPaletteSize = 2
            If (Not outputHDR) And (outputPaletteSize <> 256) Then DIBs.MakeDIBGrayscale srcImage, outputPaletteSize, False
        
        Else
        
            outColorType = png_TruecolorAlpha
            If outputHDR Then outBitsPerChannel = 16 Else outBitsPerChannel = 8
            
            'If the user's done something weird like request a specific palette size, we must apply that
            ' in advance, as the regular truecolor+alpha function won't reduce colors for us.
            If outputIndexed Then
                If autoColorModeActive Then
                    
                    'In a previous step, m_ImgMaxColorCount was set to the number of colors that actually
                    ' occur in the image
                    outColorType = png_Indexed
                    outBitsPerChannel = 8
                    If (m_ImgMaxColorCount <= 2) Then
                        outBitsPerChannel = 1
                    ElseIf (m_ImgMaxColorCount <= 4) Then
                        outBitsPerChannel = 2
                    ElseIf (m_ImgMaxColorCount <= 16) Then
                        outBitsPerChannel = 4
                    End If
                
                'Auto color mode is *not* active; forcibly reduce the image to an optimal palette
                Else
                    outColorType = png_Indexed
                    outBitsPerChannel = 8
                    m_ImgPaletteSafe = True
                    If (Not useOrigMode) Then Palettes.GetNeuquantPalette_RGBA srcImage, m_ImgPalette, outputPaletteSize
                    Palettes.ApplyPaletteToImage_KDTree srcImage, m_ImgPalette, True
                End If
                
            End If
        
        End If
        
    'Alpha from color requires us to manually create a custom alpha channel for the image.
    ' (Note that this setting *cannot* occur with "use original file settings" mode; that mode
    ' will use a binary alpha cutoff, instead.)
    ElseIf (outputAlphaMode = PDAS_NewAlphaFromColor) Then
        
        'Request a tRNS chunk in the final image
        m_WriteTrns = True
        
        'First things first: before doing anything else, retrieve a transparency table of the image's
        ' new alpha channel, using the specified color as an alpha channel guide.
        DIBs.MakeColorTransparent_Ex srcImage, trnsTable, outputAlphaColor
        
        'Next, remap the image, and assign the specified transparency color to any transparent pixels
        DIBs.ApplyBinaryTransparencyTableColor srcImage, trnsTable, outputAlphaColor, m_CompositingColor
        
        'We must also remember to set a correct tRNS color before exiting, but we set this differently
        ' depending on the output color mode.
        If outputGrayscale Then
            m_TrnsValueG = Colors.GetLuminance(Colors.ExtractRed(outputAlphaColor), Colors.ExtractGreen(outputAlphaColor), Colors.ExtractBlue(outputAlphaColor))
            m_TrnsValueR = m_TrnsValueG
            m_TrnsValueB = m_TrnsValueG
            outColorType = png_Greyscale
            If outputHDR Then
                outBitsPerChannel = 16
            Else
                outBitsPerChannel = 8
            End If
        Else
            m_TrnsValueR = Colors.ExtractRed(outputAlphaColor)
            m_TrnsValueG = Colors.ExtractGreen(outputAlphaColor)
            m_TrnsValueB = Colors.ExtractBlue(outputAlphaColor)
            outColorType = png_Truecolor
            If outputHDR Then
                outBitsPerChannel = 16
            Else
                outBitsPerChannel = 8
            End If
        End If
        
    'The only option left is "Alpha by cutoff", which requires us to manually create a custom alpha
    ' channel for the image, using some specified threshold as our 0/255 alpha split.
    Else
        
        'Request a tRNS chunk in the final image
        m_WriteTrns = True
        
        'First things first: before doing anything else, retrieve a transparency table of the image's
        ' new alpha channel, using the specified cutoff as an alpha channel guide.
        DIBs.ApplyAlphaCutoff_Ex srcImage, trnsTable, outputAlphaCutoff
        
        'Next, we need to find a color that does *not* exist in the image; this will be used as our
        ' "transparent" flag color.  (Note: if all 16.7 million colors are used in the image, this
        ' step will fail; we cover this possibility in the "automatic" output mode detector, but
        ' because this step means the user has *forcibly* requested binary alpha, we have no choice
        ' but to honor their request, however stupid it may be.)
        DIBs.GetDIBColorCountObject srcImage, cCount, False
        cCount.GetUnusedColor m_TrnsValueR, m_TrnsValueG, m_TrnsValueB
        outputAlphaColor = RGB(m_TrnsValueR, m_TrnsValueG, m_TrnsValueB)
        
        'Next, remap the image, and assign the specified transparency color to any transparent pixels
        DIBs.ApplyBinaryTransparencyTableColor srcImage, trnsTable, outputAlphaColor, m_CompositingColor
        
        'We must remember to set a correct tRNS color before exiting, but we set this differently
        ' depending on the output color mode.
        If outputGrayscale Then
            m_TrnsValueG = Colors.GetLuminance(Colors.ExtractRed(outputAlphaColor), Colors.ExtractGreen(outputAlphaColor), Colors.ExtractBlue(outputAlphaColor))
            m_TrnsValueR = m_TrnsValueG
            m_TrnsValueB = m_TrnsValueG
            outColorType = png_Greyscale
            If outputHDR Then
                outBitsPerChannel = 16
            Else
                outBitsPerChannel = 8
            End If
        Else
            m_TrnsValueR = Colors.ExtractRed(outputAlphaColor)
            m_TrnsValueG = Colors.ExtractGreen(outputAlphaColor)
            m_TrnsValueB = Colors.ExtractBlue(outputAlphaColor)
            outColorType = png_Truecolor
            If outputHDR Then
                outBitsPerChannel = 16
            Else
                outBitsPerChannel = 8
            End If
        End If
        
    End If

End Sub

Private Sub SetupExportSettings_Auto(ByRef srcImage As pdDIB, ByRef srcPDImage As pdImage, ByRef outColorType As PD_PNGColorType, ByRef outBitsPerChannel As Long, ByRef fullParamString As String)

    'Start by analyzing the image to determine its current color makeup.
    ' (This analysis informs subsequent steps, and by performing it upfront we can reuse things like the
    ' image's palette (if any) because the auto-detector will produce a working copy - this is why
    ' some of the values passed to the analyzer are module-level variables.)
    Dim outputAlphaMode As PD_ALPHA_STATUS, isGrayscale As Boolean
    Export_AnalyzeImage srcImage, m_ImgColorCount, isGrayscale, outputAlphaMode, m_ImgPalette
    
    'Grayscale images have special requirements for 1/2/4-bit modes, so we don't care about their current color count.
    ' (A precise analysis of bit-depth will happen later in the function.)
    If isGrayscale Then
        m_ImgMaxColorCount = 256
    Else
        m_ImgMaxColorCount = m_ImgColorCount
    End If
    
    'Use the analyzer's findings to see if a truecolor output mode is required
    Dim isTrueColor As Boolean
    isTrueColor = (m_ImgColorCount > 256) And (Not isGrayscale)
    
    'PNGs support 2- and 4-bit grayscale data; if an image is already grayscale, quickly analyze the palette
    ' to determine if it meets 2- or 4-bit criteria.
    Dim isGrayscaleSpecialBitDepth As Long, numPaletteAlphaNon255Non0 As Long, numPaletteAlphaZero As Long
    If isGrayscale Then
        isGrayscaleSpecialBitDepth = GetPreciseGrayBitDepth()
        If (m_ImgColorCount <= 256) Then CountPaletteAlphaEntries numPaletteAlphaNon255Non0, numPaletteAlphaZero
    End If
    
    'We now know what the image's current color count is, whether it's true grayscale or monochrome, and what the
    ' status of its alpha channel is.  Keep an eye on these parameters as we progress through the function,
    ' because they'll be reused when making decisions about export settings.
    
    'Interlacing is currently forcibly disabled, as it is unsupported in major web browsers (the images will
    ' still load but they are not displayed progressively) and it tends to greatly increase file size.
    'Dim pngInterlacing As Boolean
    'pngInterlacing = cParams.GetBool("PNGInterlacing", False)
    
    'Background color and chunk creation need to be detected up front, as they must be added to the PNG file
    ' in a specific order ("After PLTE; before IDAT" per the spec)
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString fullParamString
    
    m_BkgdColor = cParams.GetLong("png-background-color", vbWhite)
    m_CompositingColor = m_BkgdColor
    m_EmbedBkgdColor = cParams.GetBool("png-create-bkgd", False)
    
    'Some color modes may require us to produce palettes or custom transparency tables.
    Dim i As Long, trnsTable() As Byte, cCount As pdColorCount
    
    'We now know enough (finally) to construct an output color model and suggested bit-depth.  Whew!
    
    'First, handle images without meaningful alpha data
    If (outputAlphaMode = PDAS_NoAlpha) Then
        
        'If the image doesn't have alpha data, we definitely don't want to write a transparency chunk!
        m_WriteTrns = False
        
        If isTrueColor Then
            outColorType = png_Truecolor
            outBitsPerChannel = 8
        Else
            If isGrayscale Then
                outColorType = png_Greyscale
                outBitsPerChannel = isGrayscaleSpecialBitDepth
            Else
                outColorType = png_Indexed
                m_ImgPaletteSafe = True
                If (m_ImgColorCount <= 2) Then
                    outBitsPerChannel = 1
                ElseIf (m_ImgColorCount <= 4) Then
                    outBitsPerChannel = 2
                ElseIf (m_ImgColorCount <= 16) Then
                    outBitsPerChannel = 4
                Else
                    outBitsPerChannel = 8
                End If
            End If
        End If
        
    'Complicated (full) alpha
    ElseIf (outputAlphaMode = PDAS_ComplicatedAlpha) Then
    
        If isTrueColor Then
            outColorType = png_TruecolorAlpha
            outBitsPerChannel = 8
        Else
        
            If isGrayscale Then
                
                outBitsPerChannel = 8
                
                'Grayscale can be written out to file in a few different ways, depending on the presence of
                ' alpha data.
                ' 1) If there are more than 257 grayscale + alpha combinations in the image, we must write the
                '    image as gray+alpha
                If (m_ImgColorCount > 256) Then
                    outColorType = png_GreyscaleAlpha
                
                'If there are 256 or less unique grayscale + alpha combinations, we can actually write this
                ' as a paletted image (instead of requiring a full additional alpha channel).
                Else
                    
                    outColorType = png_Indexed
                    
                    'Note that the module-level image palette is safe; the indexer uses this value to determine
                    ' whether it needs to forcibly downsample an image to 8-bit indexed color.
                    m_ImgPaletteSafe = True
                    
                    'Monochrome is a weird exception, as we can write the image as monochrome if either black or white
                    ' is transparent, but not both; note that we don't have to worry about this in PD, however,
                    ' because we premultiply alpha (so fully transparent pixels will *always* be black).
                    If (m_ImgColorCount <= 2) Then
                        outBitsPerChannel = 1
                    ElseIf (m_ImgColorCount <= 4) Then
                        outBitsPerChannel = 2
                    ElseIf (m_ImgColorCount <= 16) Then
                        outBitsPerChannel = 4
                    Else
                        outBitsPerChannel = 8
                    End If
                    
                End If
                
            'If the image is not grayscale, and not true color, it must be paletted
            Else
                outColorType = png_Indexed
                m_ImgPaletteSafe = True
                If (m_ImgColorCount <= 2) Then
                    outBitsPerChannel = 1
                ElseIf (m_ImgColorCount <= 4) Then
                    outBitsPerChannel = 2
                ElseIf (m_ImgColorCount <= 16) Then
                    outBitsPerChannel = 4
                Else
                    outBitsPerChannel = 8
                End If
            End If
            
        End If
    
    'Binary alpha (meaning the image only contains alpha values that are 0 or 255, nothing in-between).
    ' In many cases, we can write this data as 24-bit and simply "flag" a single color as transparent.
    ' Sometimes, however, we'll need to fall back to full alpha coverage (for example, if the image is
    ' grayscale but all 256 gray shades are in use in opaque pixels - then we have to use a full
    ' alpha channel).
    Else
        
        m_WriteTrns = False
        
        'Before proceeding, see if this image came from an existing PNG file with an existing tRNS chunk.
        ' If it did, we'll try to reuse the original file's settings as much as possible.
        Dim hasOrigTrnsColor As Boolean, origTrnsColor As Long
        If (Not srcPDImage Is Nothing) Then
            hasOrigTrnsColor = srcPDImage.ImgStorage.DoesKeyExist("png-transparent-color")
            If hasOrigTrnsColor Then
                origTrnsColor = srcPDImage.ImgStorage.GetEntry_Long("png-transparent-color", 0)
                m_TrnsValueR = Colors.ExtractRed(origTrnsColor)
                m_TrnsValueG = Colors.ExtractGreen(origTrnsColor)
                m_TrnsValueB = Colors.ExtractBlue(origTrnsColor)
            End If
        End If
            
        If isTrueColor Then
            
            'In true color images, we need to find a unique, unused color to use as our transparent color.
            ' In the (extremely rare) case that this proves impossible, we'll fall back to writing a
            ' standard 32-bpp image.
            outColorType = png_Truecolor
            outBitsPerChannel = 8
            
            'Start by creating a color count object for this image, and note that we *don't* want to track alpha.
            ' (Alpha = 0 always uses pure black in PD images because we require premultiplied alpha.)
            DIBs.GetDIBColorCountObject srcImage, cCount, False
            
            'Next, query the object for an unused color, and start the search with the image's original
            ' transparent color (if any; if it didn't supply one, we'll just start at black).
            If cCount.DoesColorExist(m_TrnsValueR, m_TrnsValueG, m_TrnsValueB) Then
                
                'The original transparent color now appears in the image, which means we can't use it for export.
                ' We need to find a new one.
                If cCount.GetUnusedColor(m_TrnsValueR, m_TrnsValueG, m_TrnsValueB) Then
                    
                    'A new, unused color was found.  Use it to add transparency to the image.
                    m_WriteTrns = True
                
                'By some (horrifying) miracle, all 16.7 million colors are used in this image.  We can't use a tRNS
                ' chunk to write transparency data; fall back to 32-bit RGBA
                Else
                    outColorType = png_TruecolorAlpha
                End If
            
            'The original transparent color does not appear in the image!  Use it as our transparent color.
            Else
                m_WriteTrns = True
            End If
            
            'If we can still use a transparent color, downsample the image now, and replace any transparent pixels
            ' with the designated color.
            If (outColorType = png_Truecolor) Then
                DIBs.RetrieveTransparencyTable srcImage, trnsTable
                DIBs.ApplyBinaryTransparencyTableColor srcImage, trnsTable, RGB(m_TrnsValueR, m_TrnsValueG, m_TrnsValueB), m_CompositingColor
            End If
        
        'Not truecolor, which leaves grayscale and indexed modes
        Else
        
            If isGrayscale Then
                
                outBitsPerChannel = 8
                
                'Grayscale can be written out to file in a few different ways, depending on the way the alpha
                ' data is presented.
                
                ' 1) If there are more than 256 grayscale + alpha combinations in the image, we must write the
                '    image as gray+full alpha channel.  (This can happen if all 256 shades of gray are used in
                '    an image, *and* full transparency is also used somewhere - no "unused" gray shades exist
                '    to use as a transparent color flag.)
                If (m_ImgColorCount > 256) Then
                    outColorType = png_GreyscaleAlpha
                
                'If there are 256 or less unique grayscale + alpha combinations, we can successfully write
                ' the grayscale image, because at least one unused shade exists which we can "steal" for
                ' our transparency flag.
                Else
                    
                    outColorType = png_Greyscale
                    outBitsPerChannel = 8
                    m_WriteTrns = True
                    
                    'The target grayscale bit-depth presents a problem, as we may need to "bump" up the
                    ' bit-depth by one to free up a shade that we can use as a transparency marker.
                    If (isGrayscaleSpecialBitDepth < 8) Then
                        
                        'Calculate a bit-depth table.  We need to see if there are any gray shades in the current
                        ' bit-depth table that are unused; if there are, we can steal them to use as a transparency
                        ' flag.
                        Dim numShadesAvailable As Long
                        numShadesAvailable = 2 ^ isGrayscaleSpecialBitDepth
                        
                        Dim gDivisor As Long
                        gDivisor = 255 \ (numShadesAvailable - 1)
                        
                        Dim colFound() As Byte
                        ReDim colFound(0 To numShadesAvailable - 1) As Byte
                        For i = 0 To m_ImgColorCount - 1
                            If (m_ImgPalette(i).Alpha > 0) Then colFound(m_ImgPalette(i).Green \ gDivisor) = 1
                        Next i
                        
                        'Test the original transparent color (if any); we want to give it priority.
                        If (colFound(m_TrnsValueG \ gDivisor) <> 0) Then
                            
                            Dim mustUpsample As Boolean
                            mustUpsample = True
                            
                            'See if another value is available for use.
                            For i = 0 To numShadesAvailable - 1
                                If (colFound(i) = 0) Then
                                    m_TrnsValueG = i * gDivisor
                                    mustUpsample = False
                                    Exit For
                                End If
                            Next i
                            
                            'If the requested color isn't available, we have no choice but to save at a higher
                            ' bit-depth.  In some cases we can just bump up the bit-depth to the next level,
                            ' but in others (e.g. 2-bits to 4-bits) we can't, as the available colors don't
                            ' map cleanly.
                            If mustUpsample Then
                            
                                If (isGrayscaleSpecialBitDepth = 1) Then
                                    outBitsPerChannel = 2
                                    m_TrnsValueG = 85
                                Else
                                    outBitsPerChannel = 8
                                    m_TrnsValueG = 1
                                End If
                            
                            'Upsampling isn't required; use the marked transparent color and exit
                            Else
                                outBitsPerChannel = isGrayscaleSpecialBitDepth
                            End If
                        
                        'The original transparent color does not appear in the image; reuse it!
                        Else
                            outBitsPerChannel = isGrayscaleSpecialBitDepth
                        End If
                        
                        'We now have a transparency color we can use.  Apply it now.
                        DIBs.RetrieveTransparencyTable srcImage, trnsTable
                        DIBs.ApplyBinaryTransparencyTableColor srcImage, trnsTable, RGB(m_TrnsValueG, m_TrnsValueG, m_TrnsValueG), m_CompositingColor
                        
                    '8-bit grayscale is simpler.
                    Else
                            
                        'As with a color image, we need to find an unused value in the grayscale palette.  Note that
                        ' the color of transparent pixels will always be black (due to premultiplication), so we
                        ' can't just grab the transparent shade from the existing palette (in case opaque black is
                        ' also used in the image).
                        ReDim colFound(0 To 255) As Byte
                        For i = 0 To m_ImgColorCount - 1
                            If (m_ImgPalette(i).Alpha > 0) Then colFound(m_ImgPalette(i).Green) = 1
                        Next i
                        
                        'Test the original transparent color (if any); we want to give it priority.
                        If (colFound(m_TrnsValueG) <> 0) Then
                        
                            'Grab the first value that *isn't* used
                            For i = 0 To 255
                                If (colFound(i) = 0) Then
                                    m_TrnsValueG = i
                                    Exit For
                                End If
                            Next i
                        
                        End If
                        
                        'We now have a transparency color we can use.  Apply it now.
                        DIBs.RetrieveTransparencyTable srcImage, trnsTable
                        DIBs.ApplyBinaryTransparencyTableColor srcImage, trnsTable, RGB(m_TrnsValueG, m_TrnsValueG, m_TrnsValueG), m_CompositingColor
                        
                    End If
                    
                End If
            
            'Indexed color images require no special treatment.
            Else
                outColorType = png_Indexed
                m_ImgPaletteSafe = True
                If (m_ImgColorCount <= 2) Then
                    outBitsPerChannel = 1
                ElseIf (m_ImgColorCount <= 4) Then
                    outBitsPerChannel = 2
                ElseIf (m_ImgColorCount <= 16) Then
                    outBitsPerChannel = 4
                Else
                    outBitsPerChannel = 8
                End If
            End If
            
        End If
    
    '/End alpha output mode checks
    End If
    
End Sub

'APNGs require a separate, specialized export analysis, because we have to analyze *all* output frames
' in order to determine which color mode(s) will work.
Private Sub SetupExportSettings_APNG(ByRef srcPDImage As pdImage, ByRef outColorType As PD_PNGColorType, ByRef outBitsPerChannel As Long, ByRef fullParamString As String)
    
    'Unlike a regular static PNG, we need to perform the (ugly) task of analyzing *every*
    ' frame of the animated image.  Note that we can shortcut the process if we determine that
    ' the image requires 32-bpp output (because the auto-exporter won't recommend any "higher"
    ' output depths than that, at least at present), but for paletted images, we need to check
    ' every frame to ensure that its colors fit into a pre-constructed palette.  This is
    ' especially critical when the source is an animated GIF, as animated GIFs allow each
    ' frame to have its *own* palette, but animated PNGs do not - so APNGs will have to
    ' work around this by writing 24-bpp data (with a bonus transparent color, as required).
    
    'Note that unlike the actual writing step, this step analyzes raw source layer data.
    ' It *does not* apply non-destructive transforms or crop regions, making it entirely lossless.
    ' We do this because it makes the analysis a hell of a lot faster (and less memory-intensive).
    Dim outputAlphaMode As PD_ALPHA_STATUS, frameAlphaStatus As PD_ALPHA_STATUS
    Dim frameIsGrayscale As Boolean, allFramesAreGrayscale As Boolean
    Dim localPalette() As RGBQuad
    Dim frameColorCount As Long
    
    allFramesAreGrayscale = True
    outputAlphaMode = PDAS_NoAlpha
    
    Dim cColorCount As pdColorCount
    Set cColorCount = New pdColorCount
    cColorCount.SetAlphaTracking True
    
    Dim i As Long
    For i = 0 To srcPDImage.GetNumOfLayers - 1
        
        ProgressBars.SetProgBarVal i
        Message "Analyzing animation frame %1 of %2...", i + 1, srcPDImage.GetNumOfLayers(), "DONOTLOG"
        
        'Analyze the current frame
        Export_AnalyzeImage srcPDImage.GetLayerByIndex(i).GetLayerDIB, frameColorCount, frameIsGrayscale, frameAlphaStatus, localPalette, cColorCount, True
        
        'If any one frame is *not* grayscale, no frames are grayscale.  (Note that PD doesn't currently
        ' do anything special with grayscale; instead, grayscale animations are deliberately identified
        ' as paletted images because paletted images are much easier to "drop" transparency into.)
        If (Not frameIsGrayscale) Then allFramesAreGrayscale = False
        
        'If this frame requires complicated alpha, and it has more than 256 colors (or any previous frames
        ' have more than 256 colors, or our running global palette now exceeds 256 colors), we know we have
        ' to use 32-bpp export mode - so we can exit immediately.
        If ((frameColorCount > 256) Or (cColorCount.GetUniqueRGBACount() > 256)) And ((frameAlphaStatus = PDAS_ComplicatedAlpha) Or (outputAlphaMode = PDAS_ComplicatedAlpha)) Then
            outColorType = png_TruecolorAlpha
            outBitsPerChannel = 8
            m_ImgColorCount = 257
            m_ImgMaxColorCount = m_ImgColorCount
            m_apngTrnsAvailable = True
            m_WriteTrns = False
            Exit Sub
        
        'If this frame does not require 32-bpp color mode, we need to merge any settings we found
        ' into image-wide trackers.
        Else
            
            'If any one frame requires complex alpha, the image as a whole requires complex alpha
            If (frameAlphaStatus = PDAS_ComplicatedAlpha) Then
                outputAlphaMode = PDAS_ComplicatedAlpha
            
            Else
                
                'If any one frame requires binary alpha, the image as a whole requires binary alpha,
                ' but note that this setting is not as critical as PD will attempt to add a mechanism
                ' for transparency to the current image using whatever means is appropriate for its
                ' particular color model, because we want a transparent color available for use in
                ' optimizing frame differentials.
                If (frameAlphaStatus = PDAS_BinaryAlpha) And (outputAlphaMode <> PDAS_ComplicatedAlpha) Then outputAlphaMode = PDAS_BinaryAlpha
                
            End If
            
        End If
        
    Next i
    
    'Perform one last check for "image requires full 32-bpp mode", as the final frame in the image
    ' may have triggered this (so the early-exit check didn't catch it).
    If (cColorCount.GetUniqueRGBACount > 256) And (outputAlphaMode = PDAS_ComplicatedAlpha) Then
        outColorType = png_TruecolorAlpha
        outBitsPerChannel = 8
        m_ImgColorCount = 257
        m_ImgMaxColorCount = m_ImgColorCount
        m_apngTrnsAvailable = True
        m_WriteTrns = False
        Exit Sub
    End If
    
    'If we're still here, this image probably doesn't need to be exported as 32-bpp RGBA.  Let's figure
    ' out the best combination of settings for it.
    
    'Start by cloning our final global palette into the module-level image palette.  (Other functions
    ' need to analyze this to determine special color mode settings.)
    m_ImgColorCount = cColorCount.GetUniqueRGBACount()
    If (m_ImgColorCount <= 256) Then
        
        ReDim m_ImgPalette(0 To m_ImgColorCount - 1) As RGBQuad
        cColorCount.GetPalette m_ImgPalette
        
        'Look for a fully transparent index in the palette.  If it doesn't have one, we want
        ' to add one (as that will allow us to optimize via frame differentials).
        Dim palHasZeroTransparency As Boolean
        For i = 0 To m_ImgColorCount - 1
            If (m_ImgPalette(i).Alpha = 0) Then
                palHasZeroTransparency = True
                Exit For
            End If
        Next i
        
        'If the palette does NOT have a zero-transparency index, and it has room for more colors,
        ' add a zero-transparency index now.
        If (Not palHasZeroTransparency) Then
        
            If (m_ImgColorCount < 256) Then
                
                ReDim Preserve m_ImgPalette(0 To m_ImgColorCount) As RGBQuad
                With m_ImgPalette(m_ImgColorCount)
                    .Red = 0
                    .Green = 0
                    .Blue = 0
                    .Alpha = 0
                End With
                m_ImgColorCount = m_ImgColorCount + 1
                m_apngTrnsAvailable = True
                
            'There is no room in the palette for transparency.  This sucks, but rather than
            ' "promote" the entire image to RGB output (which would require three bytes per pixel -
            ' a potentially big increase in space), simply mark the image as "transparency
            ' unavailable".  The optimizer will detect this case and skip frame differentials.
            Else
                m_apngTrnsAvailable = False
            End If
            
        Else
            m_apngTrnsAvailable = True
        End If
        
        m_WriteTrns = m_apngTrnsAvailable
        
        'Sort the palette by transparency index; PNGs can conserve a little space this way due to
        ' the way transparency is encoded.
        If m_apngTrnsAvailable Then
            Dim tmpPalette As pdPaletteChild
            Set tmpPalette = New pdPaletteChild
            tmpPalette.CreateFromRGBQuads m_ImgPalette
            tmpPalette.SortByChannel 3
            tmpPalette.CopyRGBQuadsToArray m_ImgPalette
        End If
    
    Else
        m_ImgColorCount = 257
    End If
    
    m_ImgMaxColorCount = m_ImgColorCount
    
    'Use the analyzer's findings to see if a truecolor output mode is required
    Dim isTrueColor As Boolean
    isTrueColor = (m_ImgColorCount > 256)
    
    'Some frame optimizations rely on a fully transparent color (or index), e.g. frame differentials.
    ' If the current image/palette has one available, set a corresponding flag.  If it *doesn't* have
    ' one available, we'll add one "for free" if we can (e.g. specifying an unused palette index as
    ' transparent, or prepping a tRNS chunk for a true-color image).
    If (outputAlphaMode = PDAS_ComplicatedAlpha) Or (outputAlphaMode = PDAS_BinaryAlpha) Then m_apngTrnsAvailable = True
    
    'If this image is true-color *and* it doesn't have transparency, add a transparent color flag now.
    If isTrueColor And (Not m_apngTrnsAvailable) Then
        
        'Before proceeding, see if this image came from an existing PNG file with an existing tRNS chunk.
        ' (If it did, we'll try to reuse the original file's settings as much as possible.)
        Dim hasOrigTrnsColor As Boolean, origTrnsColor As Long
        If (Not srcPDImage Is Nothing) Then
            hasOrigTrnsColor = srcPDImage.ImgStorage.DoesKeyExist("png-transparent-color")
            If hasOrigTrnsColor Then
                origTrnsColor = srcPDImage.ImgStorage.GetEntry_Long("png-transparent-color", 0)
                m_TrnsValueR = Colors.ExtractRed(origTrnsColor)
                m_TrnsValueG = Colors.ExtractGreen(origTrnsColor)
                m_TrnsValueB = Colors.ExtractBlue(origTrnsColor)
            Else
                m_TrnsValueR = 0
                m_TrnsValueG = 0
                m_TrnsValueB = 0
            End If
        End If

        'Next, query the object for an unused color, and start the search with the image's original
        ' transparent color (if any; if it didn't supply one, we'll just start at black).
        If cColorCount.DoesColorExist(m_TrnsValueR, m_TrnsValueG, m_TrnsValueB) Then
            
            'The original transparent color now appears in the image, which means we can't use it for export.
            ' We need to find a new one.
            If cColorCount.GetUnusedColor(m_TrnsValueR, m_TrnsValueG, m_TrnsValueB) Then
                
                'A new, unused color was found.  Use it to add transparency to the image.
                m_apngTrnsAvailable = True
            
            'By some (horrifying) miracle, all 16.7 million colors are used in this image.  We can't use a tRNS
            ' chunk to write transparency data; fall back to 32-bit RGBA
            Else
                m_apngTrnsAvailable = False
            End If
        
        'The original transparent color does not appear in the image!  Use it as our transparent color.
        Else
            m_apngTrnsAvailable = True
        End If
        
    End If
    
    'We now know whether the image is true-color or paletted, and what it's current alpha state is
    ' (e.g. whether it supports *any* alpha whatsoever).  Keep an eye on these parameters as we
    ' progress through the function, because they'll be reused when making decisions about export settings.
    
    'Interlacing is currently forcibly disabled, as it is unsupported in major web browsers (the images will
    ' still load but they are not displayed progressively) and it greatly increases file size, especially
    ' for APNGs.
    'Dim pngInterlacing As Boolean
    'pngInterlacing = cParams.GetBool("PNGInterlacing", False)
    
    'Background color and chunk creation need to be detected up front, as they must be added to the PNG file
    ' in a specific order ("After PLTE; before IDAT" per the spec)
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString fullParamString
    
    m_BkgdColor = cParams.GetLong("png-background-color", vbWhite)
    m_CompositingColor = m_BkgdColor
    m_EmbedBkgdColor = cParams.GetBool("png-create-bkgd", False)
    
    'We now know enough (finally) to construct an output color model and suggested bit-depth.  Whew!
    
    'Complicated (full) alpha is required
    If (outputAlphaMode = PDAS_ComplicatedAlpha) Then
        
        m_apngTrnsAvailable = True
        
        'Simplest combination: full RGBA output
        If isTrueColor Then
            outColorType = png_TruecolorAlpha
            outBitsPerChannel = 8
            m_apngTrnsAvailable = True
            m_WriteTrns = False
            
        'If true color isn't required, we can get away with a simpler color space
        Else
            
            'We could deal with grayscale images here, but because they have additional restrictions
            ' that make transparency a pain, it's easier to just identify them as paletted; the end
            ' display result is identical.
            outColorType = png_Indexed
            m_ImgPaletteSafe = True
            If (m_ImgColorCount <= 2) Then
                outBitsPerChannel = 1
            ElseIf (m_ImgColorCount <= 4) Then
                outBitsPerChannel = 2
            ElseIf (m_ImgColorCount <= 16) Then
                outBitsPerChannel = 4
            Else
                outBitsPerChannel = 8
            End If
            
            'Make sure a tRNS chunk gets written, as the frames themselves require it!
            m_apngTrnsAvailable = True
            m_WriteTrns = True
            
        End If
    
    'Next, handle images without alpha data in the frames themselves (but which likely has had
    ' transparency "added" by us, so we can use it during frame optimizations).
    ElseIf (outputAlphaMode = PDAS_NoAlpha) Then
        
        If isTrueColor Then
            outColorType = png_Truecolor
            outBitsPerChannel = 8
            m_WriteTrns = m_apngTrnsAvailable
        Else
            
            outColorType = png_Indexed
            m_ImgPaletteSafe = True
            If (m_ImgColorCount <= 2) Then
                outBitsPerChannel = 1
            ElseIf (m_ImgColorCount <= 4) Then
                outBitsPerChannel = 2
            ElseIf (m_ImgColorCount <= 16) Then
                outBitsPerChannel = 4
            Else
                outBitsPerChannel = 8
            End If
            
            m_WriteTrns = m_apngTrnsAvailable
            
        End If
        
    'Binary alpha (meaning the image only contains alpha values that are 0 or 255, nothing in-between).
    ' As APNGs often come from GIF sources, this is a likely outcome and PNG deals with it very well.
    Else
        
        m_WriteTrns = True
        
        If isTrueColor Then
            
            outColorType = png_Truecolor
            outBitsPerChannel = 8
            
            'This is an insane possibility, but if the animation uses all 16.7 million colors *and*
            ' has binary transparency, upgrade the image to full 32-bpp output as we don't have a spare
            ' color available for transparency flagging.
            If (Not m_apngTrnsAvailable) Then
                m_apngTrnsAvailable = True
                m_WriteTrns = True
                outColorType = png_TruecolorAlpha
            End If
            
        'Not truecolor, which leaves grayscale and indexed modes
        Else
            
            'Grayscale images are treated as indexed images, and indexed image require no
            ' special treatment.
            outColorType = png_Indexed
            m_ImgPaletteSafe = True
            If (m_ImgColorCount <= 2) Then
                outBitsPerChannel = 1
            ElseIf (m_ImgColorCount <= 4) Then
                outBitsPerChannel = 2
            ElseIf (m_ImgColorCount <= 16) Then
                outBitsPerChannel = 4
            Else
                outBitsPerChannel = 8
            End If
            
            m_apngTrnsAvailable = True
            m_WriteTrns = True
            
        End If
    
    '/End alpha output mode checks
    End If
    
End Sub

'New, experimental "lossy" APNG optimization.  This will reduce output frames to 8-bpp (meaning one shared
' palette for *all* frames) which can greatly reduce file size, but may incur unacceptable quality loss.
Private Sub SetupExportSettings_APNG_Lossy(ByRef srcPDImage As pdImage, ByRef outColorType As PD_PNGColorType, ByRef outBitsPerChannel As Long, ByRef fullParamString As String)
    
    'Before proceeding, forcibly set all meaningful output settings
    outColorType = png_Indexed
    
    '(This may need to change if the user is allowed to set output palette size)
    outBitsPerChannel = 8
    
    'We now want to produce a 255 color palette + 1 transparent index.  (Since the whole point of this
    ' exercise is to minimize APNG size, we want a transparent index available so pixel blanking is
    ' available as a size optimization strategy.)
    
    'What makes this a little weird is that we need to generate a unique palette across *all* incoming
    ' frames.  The only way to do this is to sample from each frame!
    
    'Because of the performance constraints, we're going to use a median cut algorithm instead of a
    ' neural network one.  (Random sampling from every input frame could work okay with a neural network,
    ' but I'd need to write a completely different class interface and I don't want to do that right now.)
    If PNG_DEBUG_VERBOSE Then PDDebug.LogAction "Analyzing all layers for lossy PNG palettization..."
    
    'The palette module can produce a full-image palette for us.
    Palettes.GetOptimizedPaletteIncAlpha_AllLayers srcPDImage, m_ImgPalette, 256, suppressMessages:=False, modifyProgBarMax:=ProgressBars.GetProgBarMax(), modifyProgBarOffset:=srcPDImage.GetNumOfLayers()
    
    m_ImgColorCount = 256
    m_ImgMaxColorCount = m_ImgColorCount
    m_ImgPaletteSafe = True
    
    'Make sure a tRNS chunk gets written, as the frames themselves require it!
    m_apngTrnsAvailable = True
    m_WriteTrns = True
    
    'Fill any other module-level values from the param string
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString fullParamString
    
    'Background color and chunk creation need to be detected up front, as they must be added to the PNG file
    ' in a specific order ("After PLTE; before IDAT" per the spec)
    m_BkgdColor = cParams.GetLong("png-background-color", vbWhite)
    m_CompositingColor = m_BkgdColor
    m_EmbedBkgdColor = cParams.GetBool("png-create-bkgd", False)
    
End Sub

'The module-level m_ImgColorCount and m_ImgPalette values must be populated before calling this function.
Private Sub CountPaletteAlphaEntries(ByRef numNon255Or0 As Long, ByRef numZero As Long)
    
    Dim i As Long
    For i = 0 To m_ImgColorCount - 1
        If (m_ImgPalette(i).Alpha < 255) Then
            If (m_ImgPalette(i).Alpha = 0) Then
                numZero = numZero + 1
            Else
                numNon255Or0 = numNon255Or0 + 1
            End If
        End If
    Next i

End Sub

'The module-level m_ImgColorCount and m_ImgPalette values must be populated before calling this function.
Private Function GetPreciseGrayBitDepth() As Long

    GetPreciseGrayBitDepth = 8
    Dim i As Long, badValueFound As Boolean
    
    'Check 1-bit
    If (m_ImgColorCount <= 2) Then
        badValueFound = False
        For i = 0 To m_ImgColorCount - 1
            If (m_ImgPalette(i).Red <> 0) And (m_ImgPalette(i).Red <> 255) Then
                badValueFound = True
                Exit For
            End If
        Next i
        If (Not badValueFound) Then GetPreciseGrayBitDepth = 1
    End If
    
    'Check 2-bit
    If (m_ImgColorCount <= 4) And (GetPreciseGrayBitDepth = 8) Then
        badValueFound = False
        For i = 0 To m_ImgColorCount - 1
            If (m_ImgPalette(i).Red Mod 85 <> 0) Then
                badValueFound = True
                Exit For
            End If
        Next i
        If (Not badValueFound) Then GetPreciseGrayBitDepth = 2
    End If
    
    'Check 4-bit
    If (m_ImgColorCount <= 16) And (GetPreciseGrayBitDepth = 8) Then
        badValueFound = False
        For i = 0 To m_ImgColorCount - 1
            If (m_ImgPalette(i).Red Mod 17 <> 0) Then
                badValueFound = True
                Exit For
            End If
        Next i
        If (Not badValueFound) Then GetPreciseGrayBitDepth = 4
    End If
    
End Function

'Given a 32-bpp source (the source *MUST BE 32-bpp*, but its alpha channel can be constant), fill various critical
' pieces of information about the image's color+opacity makeup:
' 1) netColorCount: an integer on the range [1, 257].  257 = more than 256 unique colors, so full RGB output is required
' 2) isGrayscale: TRUE if the image consists of only gray shades (PNG supports grayscale-only output modes)
' 3) currentAlphaStatus: custom enum describing the alpha channel contents of the image (determines tRNS viability)
' 4) uniqueColors(): if the image contains 256 unique color + opacity combinations (or less), this will return an exact palette
'                    (important because PNGs allow for full RGBA palettes, unlike GIF)
'
'The function as a whole returns TRUE if the source image was scanned correctly; FALSE otherwise.  (FALSE probably means you passed
' it a 24-bpp image!)
Private Sub Export_AnalyzeImage(ByRef srcDIB As pdDIB, ByRef netColorCount As Long, ByRef isGrayscale As Boolean, ByRef currentAlphaStatus As PD_ALPHA_STATUS, ByRef uniqueColors() As RGBQuad, Optional ByRef cColorTree As pdColorCount = Nothing, Optional ByVal countAllColors As Boolean = False)

    If PNG_DEBUG_VERBOSE And (Not countAllColors) Then PDDebug.LogAction "Analyzing image prior to PNG export..."
    
    Dim srcPixels() As Byte, tmpSA As SafeArray1D
    
    Dim x As Long, y As Long, finalX As Long, finalY As Long, pxSize As Long
    finalY = srcDIB.GetDIBHeight - 1
    finalX = srcDIB.GetDIBWidth - 1
    If (srcDIB.GetDIBColorDepth = 32) Then pxSize = 4 Else pxSize = 3
    finalX = finalX * pxSize
    
    'Use a dedicated color counting class to "collect" a palette for this image.
    ' (Note that in a PNG context, "color" means "RGBA quad", not "RGB triplet".)
    Dim numUniqueColors As Long
    If (cColorTree Is Nothing) Then
        Set cColorTree = New pdColorCount
        numUniqueColors = 0
    Else
        numUniqueColors = cColorTree.GetUniqueRGBACount()
    End If
    
    cColorTree.SetAlphaTracking True
    
    'Total number of unique colors counted so far
    Dim non255Alpha As Boolean, nonBinaryAlpha As Boolean
    non255Alpha = False
    nonBinaryAlpha = False
    isGrayscale = True
    
    'Finally, a bunch of variables used in color calculation
    Dim r As Long, g As Long, b As Long, a As Long
    
    Dim srcPtr As Long, srcStride As Long
    srcDIB.WrapArrayAroundScanline srcPixels, tmpSA, y
    srcPtr = tmpSA.pvData
    srcStride = tmpSA.cElements
    
    'Look for unique colors
    For y = 0 To finalY
        tmpSA.pvData = srcPtr + srcStride * y
    For x = 0 To finalX Step pxSize
        
        b = srcPixels(x)
        g = srcPixels(x + 1)
        r = srcPixels(x + 2)
        If (pxSize = 4) Then a = srcPixels(x + 3) Else a = 255
        
        'First, look for non-binary alpha
        If (Not nonBinaryAlpha) Then
            If (a < 255) Then
                non255Alpha = True
                If (a > 0) Then nonBinaryAlpha = True
            End If
        End If
        
        'Until we find at least 257 unique colors, we need to investigate colors more closely
        If (numUniqueColors <= 256) Or countAllColors Then
        
            If cColorTree.AddColor(r, g, b, a) Then numUniqueColors = numUniqueColors + 1
            
            'Once more than 256 colors have been found, we no longer need to count colors, because we
            ' already know the image must be exported as 24-bit (or higher).  Note, however, that we
            ' may need to continue analyzing the image to look for unique alpha values.
            If (numUniqueColors > 256) Then numUniqueColors = 257
            
        End If
        
        If (r <> g) Or (g <> b) Or (r <> b) Then isGrayscale = False
        
        'If an image looks grayscale (so far), or if we haven't found alpha values other than 0 and 255,
        ' or we haven't found more than 256 unique colors, we need to keep scanning the image.
        If (Not countAllColors) Then
            If (Not isGrayscale) And nonBinaryAlpha And (numUniqueColors > 256) Then Exit For
        End If
        
    Next x
        If (Not countAllColors) Then
            If (numUniqueColors > 256) And nonBinaryAlpha Then Exit For
        End If
    Next y
    
    srcDIB.UnwrapArrayFromDIB srcPixels
    
    netColorCount = numUniqueColors
    
    'Further checks are only relevant if the image contains 256 colors or less
    If (numUniqueColors <= 256) And (Not countAllColors) Then
        
        'Retrieve the current color palette for this image
        cColorTree.GetPalette uniqueColors
        
        'Sort the palette from "least opaque" to "most opaque"; PNGs only require transparency data
        ' for non-opaque palette entries, so we can shave a few bytes by ordering the palette table
        ' so that any/all 255 entries occur at the table's tail.
        Dim cPaletteSort As pdPaletteChild
        Set cPaletteSort = New pdPaletteChild
        cPaletteSort.CreateFromRGBQuads uniqueColors
        cPaletteSort.SortByChannel 3    '0 = red, 1 = green, 2 = blue, 3 = alpha
        cPaletteSort.CopyRGBQuadsToArray uniqueColors
        
    'End "If 256 colors or less..."
    End If
    
    'Convert our individual alpha trackers into the single "currentAlphaStatus" output, then exit
    If non255Alpha Then
        If nonBinaryAlpha Then
            currentAlphaStatus = PDAS_ComplicatedAlpha
        Else
            currentAlphaStatus = PDAS_BinaryAlpha
        End If
    Else
        currentAlphaStatus = PDAS_NoAlpha
    End If
    
End Sub

Friend Function ExportStep1_StartFile(ByRef cStream As pdStream) As PD_PNGResult
    
    'Start with the PNG "magic numbers".  From the spec:
    ' 5.2 PNG signature
    ' The first eight bytes of a PNG datastream always contain the following (decimal) values:
    ' 137 80 78 71 13 10 26 10
    cStream.WriteLong_BE &H89504E47
    cStream.WriteLong_BE &HD0A1A0A
    
    ExportStep1_StartFile = png_Success
    
End Function

'Make sure you have populated the module-level m_Header struct before calling this function!
Friend Function ExportStep2_WriteIHDR(ByRef cStream As pdStream) As PD_PNGResult

    'Initialize a chunk object
    Dim dstChunk As pdPNGChunk
    Set dstChunk = New pdPNGChunk
    
    'The header is always 13-bytes long.  (Note that we never include length, chunk type, or CRC lengths -
    ' the chunk class handles all that for us.)
    dstChunk.CreateChunkForExport "IHDR", 13
    
    'Write image width/height
    dstChunk.BorrowData.WriteLong_BE m_Header.Width
    dstChunk.BorrowData.WriteLong_BE m_Header.Height
    
    'Next come five PNG-specific bytes:
    
    'Bit depth (1, 2, 4, 8, 16)
    dstChunk.BorrowData.WriteByte m_Header.BitDepth
    
    'Colour type (0, 2, 3, 4, 6)
    dstChunk.BorrowData.WriteByte m_Header.ColorType
    
    'Compression method (always 0)
    dstChunk.BorrowData.WriteByte 0
    
    'Filter method (always 0)
    dstChunk.BorrowData.WriteByte 0
    
    'Interlace method (0 or 1)
    dstChunk.BorrowData.WriteByte IIf(m_Header.Interlaced, 1, 0)
    
    'Finalize the chunk and copy its contents into the final stream
    dstChunk.FinalizeChunk
    cStream.WriteStream dstChunk.BorrowData
    
    ExportStep2_WriteIHDR = png_Success
    
End Function

'Make sure you have populated the module-level m_Header struct before calling this function!
Friend Function ExportStep2a_WriteacTL(ByRef cStream As pdStream, ByVal numFrames As Long, ByVal numPlays As Long) As PD_PNGResult

    'Initialize a chunk object
    Dim dstChunk As pdPNGChunk
    Set dstChunk = New pdPNGChunk
    
    'The header is always 13-bytes long.  (Note that we never include length, chunk type, or CRC lengths -
    ' the chunk class handles all that for us.)
    dstChunk.CreateChunkForExport "acTL", 8
    
    'Write animation frame count
    dstChunk.BorrowData.WriteLong_BE numFrames
    
    'Write number of plays (0 = repeat infinitely)
    dstChunk.BorrowData.WriteLong_BE numPlays
    
    'Finalize the chunk and copy its contents into the destination stream
    dstChunk.FinalizeChunk
    cStream.WriteStream dstChunk.BorrowData
    
    ExportStep2a_WriteacTL = png_Success
    
End Function

'Make sure you have populated the module-level m_Header struct before calling this function!
Friend Function ExportStep3_FilterBytes(ByRef cStream As pdStream, ByRef srcImage As pdDIB, ByRef tmpChunk As pdPNGChunk, Optional ByRef fullParamString As String = vbNullString, Optional ByVal useFilterStrategy As PD_PNG_FilterStrategy = png_FilterUndefined, Optional ByVal writeAncillaryChunks As Boolean = True, Optional ByVal sourceDIBIsDisposable As Boolean = False) As PD_PNGResult
    
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    'First, we need to generate a PNG-compatible layout of image data.  In a normal 32-bpp RGBA image, this simply
    ' requires us to swizzle red and blue channels.  For other bit-depths, however, we need to produce the proper
    ' bit-stream here.
    Dim finalBytes() As Byte, sizeOfPngBytes As Long
    Dim dstStride As Long
    
    'Prep a parameter parser, as certain bit-depth and color mode combinations may require specialized pixel handling
    ' (e.g. compositing against a specific backcolor).
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.SetParamString fullParamString
    
    'At present, we produce a full copy of the image in the required bit-depth and layout.  In the future,
    ' it may be nice to handle this on a per-scanline basis, to avoid the need for massive up-front
    ' memory allocations (especially when working with HDR modes).  Note, however, that this would require
    ' a different solution than libdeflate as it does *not* expose a streaming API.
    If (m_Header.ColorType = png_TruecolorAlpha) Then
        sizeOfPngBytes = srcImage.GetDIBWidth * srcImage.GetDIBHeight * 4 * (m_Header.BitDepth \ 8)
    ElseIf (m_Header.ColorType = png_Truecolor) Then
        sizeOfPngBytes = srcImage.GetDIBWidth * srcImage.GetDIBHeight * 3 * (m_Header.BitDepth \ 8)
    ElseIf (m_Header.ColorType = png_GreyscaleAlpha) Then
        sizeOfPngBytes = srcImage.GetDIBWidth * srcImage.GetDIBHeight * 2 * (m_Header.BitDepth \ 8)
    
    'Only remaining options are grayscale and indexed, which can both be 1/2/4/8 bpp; grayscale can
    ' also be 16-bpp.
    Else
        
        '8/16 bpp is easy
        If (m_Header.BitDepth >= 8) Then
            dstStride = srcImage.GetDIBWidth * (m_Header.BitDepth \ 8)
            sizeOfPngBytes = dstStride * srcImage.GetDIBHeight
            
        '1/2/4 bpp requires us to check trailing bits for adequate space
        Else
            
            dstStride = (srcImage.GetDIBWidth * m_Header.BitDepth) \ 8
            If (dstStride < 1) Then
                dstStride = 1
            Else
                If ((dstStride * 8) \ m_Header.BitDepth < srcImage.GetDIBWidth) Then dstStride = dstStride + 1
            End If
            
            sizeOfPngBytes = dstStride * srcImage.GetDIBHeight
            
        End If
    
    End If
    
    'Prepare the buffer that will hold the color-mode-processed bytes
    ReDim finalBytes(0 To sizeOfPngBytes - 1) As Byte
    
    'RGBA data
    If (m_Header.ColorType = png_TruecolorAlpha) Then
        ExportStep3_FilterBytes = Step3Helper_TruecolorAlpha(srcImage, cStream, finalBytes, sizeOfPngBytes, dstStride, writeAncillaryChunks)
        
    'RGB data
    ElseIf (m_Header.ColorType = png_Truecolor) Then
        ExportStep3_FilterBytes = Step3Helper_Truecolor(srcImage, cStream, finalBytes, sizeOfPngBytes, dstStride, writeAncillaryChunks)
        
    'Grayscale
    ElseIf (m_Header.ColorType = png_Greyscale) Then
        ExportStep3_FilterBytes = Step3Helper_Grayscale(srcImage, cStream, finalBytes, sizeOfPngBytes, dstStride, writeAncillaryChunks)
        
    'Greyscale + alpha
    ElseIf (m_Header.ColorType = png_GreyscaleAlpha) Then
        ExportStep3_FilterBytes = Step3Helper_GrayscaleAlpha(srcImage, cStream, finalBytes, sizeOfPngBytes, dstStride, writeAncillaryChunks)
        
    'Indexed
    ElseIf (m_Header.ColorType = png_Indexed) Then
        ExportStep3_FilterBytes = Step3Helper_Indexed(srcImage, cStream, finalBytes, sizeOfPngBytes, dstStride, writeAncillaryChunks)
    
    End If
    
    'Note that we no longer require the original DIB; we can safely free it before continuing, so long as we note
    ' the image's dimensions (as the chunk filter code requires these)
    Dim srcImageWidth As Long, srcImageHeight As Long
    srcImageWidth = srcImage.GetDIBWidth
    srcImageHeight = srcImage.GetDIBHeight
    
    'Note that erasing the DIB carries some obvious assumptions - e.g. that the caller doesn't want it! -
    ' so the caller must *explicitly* notify us that the object is safe to free.  (If possible, we like to
    ' free it here, as the forthcoming compression work requires very large I/O buffers for both source
    ' and destination inputs - so every bit of free RAM is appreciated!)
    If sourceDIBIsDisposable Then srcImage.EraseDIB True
    
    If PNG_DEBUG_VERBOSE Then
        If writeAncillaryChunks Then PDDebug.LogAction "- IDAT creation: prepping pixel data took " & VBHacks.GetTimeDiffNowAsString(startTime)
    End If
    
    'From the original param string, determine a filter strategy (unless overridden by the caller).
    If (useFilterStrategy = png_FilterUndefined) Or (LenB(fullParamString) <> 0) Then useFilterStrategy = cParams.GetLong("png-filter-strategy", png_FilterAuto)
    
    'Per the spec, indexed-mode images don't usually benefit from automatic filtering.  In the future, perhaps we could
    ' still test fixed filters to see if compression improves?
    If (useFilterStrategy = png_FilterAuto) Then
        If (m_Header.ColorType = png_Indexed) Then
            useFilterStrategy = png_FilterNone
        Else
            useFilterStrategy = png_FilterOptimal
        End If
    End If
    
    'The temporary chunk handles the actual filtering for us
    If (tmpChunk Is Nothing) Then Set tmpChunk = New pdPNGChunk
    If tmpChunk.FilterChunk(m_Warnings, m_Header, VarPtr(finalBytes(0)), srcImageWidth, srcImageHeight, dstStride, useFilterStrategy) Then ExportStep3_FilterBytes = png_Success Else ExportStep3_FilterBytes = png_Failure
    
End Function

'This function auto-increments the module-level m_FrameSequence value.
Friend Function ExportStep3a_WritefcTL(ByRef cStream As pdStream, ByRef srcFrameRect As RectF, ByVal srcFrameDelayInMS As Long, ByVal frameDisposal As PD_APNG_FrameDisposal, ByVal frameBlend As PD_APNG_BlendOp) As PD_PNGResult
    
    'Initialize a chunk object
    Dim dstChunk As pdPNGChunk
    Set dstChunk = New pdPNGChunk
    
    'Frame control chunks are always 26-bytes long.  (Note that we never include length, chunk type, or CRC lengths -
    ' the chunk class handles all that for us.)
    dstChunk.CreateChunkForExport "fcTL", 16
    
    'Write the current frame sequence count, then increment it
    dstChunk.BorrowData.WriteLong_BE m_FrameSequence
    m_FrameSequence = m_FrameSequence + 1
    
    'Write out the frame rectangle (where this frame appears on the canvas)
    dstChunk.BorrowData.WriteLong_BE Int(srcFrameRect.Width + 0.5)
    dstChunk.BorrowData.WriteLong_BE Int(srcFrameRect.Height + 0.5)
    dstChunk.BorrowData.WriteLong_BE Int(srcFrameRect.Left + 0.5)
    dstChunk.BorrowData.WriteLong_BE Int(srcFrameRect.Top + 0.5)
    
    'Frame delay is written as (sigh) a ushort numerator/denominator pair.  Convert the incoming
    ' floating-point value to an approximate fraction equivalent.
    Dim delayNumerator As Long, delayDenominator As Long
    PDMath.ConvertToFraction CDbl(srcFrameDelayInMS) * 0.001, delayNumerator, delayDenominator, 0.0001
    If (delayNumerator > 65535) Then
        delayDenominator = Int(CDbl(delayDenominator) * (65535 / delayNumerator) + 0.5)
        delayNumerator = Int(CDbl(delayNumerator) * (65535 / delayNumerator) + 0.5)
    End If
    If (delayDenominator > 65535) Then
        delayNumerator = Int(CDbl(delayNumerator) * (65535 / delayDenominator) + 0.5)
        delayDenominator = Int(CDbl(delayDenominator) * (65535 / delayDenominator) + 0.5)
    End If
    dstChunk.BorrowData.WriteIntU_BE Int(delayNumerator + 0.5)
    dstChunk.BorrowData.WriteIntU_BE Int(delayDenominator + 0.5)
    
    'Finally, a one-byte frame disposal and blend op must be specified
    dstChunk.BorrowData.WriteByte frameDisposal
    dstChunk.BorrowData.WriteByte frameBlend
    
    'Finalize the chunk and copy its contents into the destination stream
    dstChunk.FinalizeChunk
    cStream.WriteStream dstChunk.BorrowData
    
    ExportStep3a_WritefcTL = png_Success
    
End Function

'Because the bKGD chunk must be placed in a specific order in the file (after the palette, if any, but before
' actual pixel data), we rely on color-format-specific orders to write it at the correct moment.
Private Sub WriteBkgdChunk(ByRef cStream As pdStream)
    
    If m_EmbedBkgdColor Then
    
        Dim tmpChunk As pdPNGChunk
        Set tmpChunk = New pdPNGChunk
        
        With tmpChunk
            
            .CreateChunkForExport "bKGD"
            
            'The format of the background chunk depends on the current color mode and bit-depth.
            Dim dstColor As Long
            
            'Grayscale is a single value, at the same bit-depth as the image
            If (m_Header.ColorType = png_Greyscale) Or (m_Header.ColorType = png_GreyscaleAlpha) Then
            
                dstColor = Colors.GetHQLuminance(Colors.ExtractRed(m_BkgdColor), Colors.ExtractGreen(m_BkgdColor), Colors.ExtractBlue(m_BkgdColor))
                
                If (m_Header.BitDepth = 16) Then
                    .BorrowData.WriteIntU_BE dstColor * 257
                ElseIf (m_Header.BitDepth = 8) Then
                    .BorrowData.WriteInt_BE dstColor
                ElseIf (m_Header.BitDepth = 4) Then
                    .BorrowData.WriteInt_BE dstColor \ 17
                ElseIf (m_Header.BitDepth = 2) Then
                    .BorrowData.WriteInt_BE dstColor \ 85
                ElseIf (m_Header.BitDepth = 1) Then
                    .BorrowData.WriteInt_BE dstColor \ 255
                End If
            
            'Color is an rgb triplet, two bytes per color, at the same bit-depth as the image
            ElseIf (m_Header.ColorType = png_Truecolor) Or (m_Header.ColorType = png_TruecolorAlpha) Then
                
                If (m_Header.BitDepth = 16) Then
                    .BorrowData.WriteIntU_BE Colors.ExtractRed(m_BkgdColor) * 257
                    .BorrowData.WriteIntU_BE Colors.ExtractGreen(m_BkgdColor) * 257
                    .BorrowData.WriteIntU_BE Colors.ExtractBlue(m_BkgdColor) * 257
                Else
                    .BorrowData.WriteInt_BE Colors.ExtractRed(m_BkgdColor)
                    .BorrowData.WriteInt_BE Colors.ExtractGreen(m_BkgdColor)
                    .BorrowData.WriteInt_BE Colors.ExtractBlue(m_BkgdColor)
                End If
            
            'Indexed mode has special requirements; the color is not actually a color, but an index into
            ' the palette array.
            Else
                dstColor = Palettes.GetNearestIndexRGB(m_ImgPalette, m_BkgdColor)
                .BorrowData.WriteByte dstColor
            End If
            
        End With
        
        tmpChunk.FinalizeChunk
        cStream.WriteStream tmpChunk.BorrowData
        
    End If

End Sub

'Various helper functions for prepping pixel data pre-filtering follow.  Note that source pixel data is never guaranteed
' to already match the requested color model - for example, the user can request that a color image be exported as grayscale.
' As such, analyses may have to be performed to determine how to best cram the original pixel data into the requested
' export format.  (Similarly, the source pixel data *may* already be in a compatible export format - so don't do silly
' things like forcibly quantize image data without first checking to see if it's already 256-color compatible!)
'
'The main thing each prep function needs to do is fill a finalBytes array (at the appropriate size), and populate a
' dstStride value that defines the stride of each line in the finalBytes array (which is actually declared as a 1D array).
' The stride value is critical because PNG filters operate on scanlines, so we need to know the distance between scanlines.
'
'Note that the finalBytes() array will already be initialized to the size specified by sizeOfPngBytes; this is performed
' externally as a convenience to the caller, so they can simply proceed with filling the array.
Private Function Step3Helper_TruecolorAlpha(ByRef srcImage As pdDIB, ByRef cStream As pdStream, ByRef finalBytes() As Byte, ByRef sizeOfPngBytes As Long, ByRef dstStride As Long, Optional ByVal writeAncillaryChunks As Boolean = True) As PD_PNGResult
    
    'TrueColor+Alpha mode supports 32/64-bpp output
    Dim tmpByte As Byte, srcStride As Long, srcHeightInPixels As Long
    Dim xLimit As Long, yDstOffset As Long, curOffset As Long
    Dim x As Long, y As Long
    
    'If the caller wants us to write a bKGD chunk, do so now.
    If m_EmbedBkgdColor And writeAncillaryChunks Then WriteBkgdChunk cStream
    
    'Some color-depths (e.g. HDR) require us to directly reference discrete pixels inside the
    ' source DIB; we cannot accelerate the process using pointer-based math, alas.
    Dim srcBytes() As Byte, srcSA1D As SafeArray1D
    
    'Some params are identical for all image types (such as number of scanlines)
    srcHeightInPixels = srcImage.GetDIBHeight
    
    '32-bpp path
    If (m_Header.BitDepth = 8) Then
    
        'Copy over *all* DIB data
        CopyMemoryStrict VarPtr(finalBytes(0)), srcImage.GetDIBPointer, sizeOfPngBytes
        
        srcStride = srcImage.GetDIBStride
        dstStride = srcStride
        xLimit = srcStride - 1
        
        'Manually swizzle r/b
        For y = 0 To srcHeightInPixels - 1
            yDstOffset = y * dstStride
            For x = 0 To xLimit Step 4
                tmpByte = finalBytes(yDstOffset + x)
                finalBytes(yDstOffset + x) = finalBytes(yDstOffset + x + 2)
                finalBytes(yDstOffset + x + 2) = tmpByte
            Next x
        Next y
        
    '64-bpp path
    ElseIf (m_Header.BitDepth = 16) Then
    
        srcStride = srcImage.GetDIBStride
        dstStride = srcStride * 2
        xLimit = srcStride - 1
        
        For y = 0 To srcHeightInPixels - 1
            
            yDstOffset = y * dstStride
            
            'Wrap an array around the source data
            srcImage.WrapArrayAroundScanline srcBytes, srcSA1D, y
            
            For x = 0 To xLimit Step 4
                
                'Swizzle R and B as part of the copy
                curOffset = yDstOffset + x * 2
                finalBytes(curOffset) = srcBytes(x + 2)
                finalBytes(curOffset + 2) = srcBytes(x + 1)
                finalBytes(curOffset + 4) = srcBytes(x)
                finalBytes(curOffset + 6) = srcBytes(x + 3)
                
            Next x
            
        Next y
        
        'Make sure our array is unwrapped from the source data
        srcImage.UnwrapArrayFromDIB srcBytes
        
    End If
    
    Step3Helper_TruecolorAlpha = png_Success

End Function

Private Function Step3Helper_Truecolor(ByRef srcImage As pdDIB, ByRef cStream As pdStream, ByRef finalBytes() As Byte, ByRef sizeOfPngBytes As Long, ByRef dstStride As Long, Optional ByVal writeAncillaryChunks As Boolean = True) As PD_PNGResult
    
    'TrueColor mode supports 24/48-bpp output
    Dim srcStride As Long, srcBPP As Long, srcHeightInPixels As Long, srcPtr As Long
    Dim xLimit As Long, yDstOffset As Long, curOffset As Long
    Dim x As Long, y As Long
    
    'Some color-depths (e.g. HDR) require us to directly reference discrete pixels inside the
    ' source DIB; we cannot accelerate the process using pointer-based math, alas.
    Dim srcBytes() As Byte, srcSA1D As SafeArray1D
    
    'Some params are identical for all image types (such as number of scanlines)
    srcHeightInPixels = srcImage.GetDIBHeight
    
    If writeAncillaryChunks Then
    
        'If the caller wants us to write a tRNS chunk, do so now.
        If m_WriteTrns Then WriteTRNSNow cStream
    
        'If the caller wants us to write a bKGD chunk, do so now.
        If m_EmbedBkgdColor Then WriteBkgdChunk cStream
        
    End If
    
    'Start by "removing" alpha by compositing the image against whatever backcolor the caller requested
    srcImage.CompositeBackgroundColor Colors.ExtractRed(m_CompositingColor), Colors.ExtractGreen(m_CompositingColor), Colors.ExtractBlue(m_CompositingColor)
    
    '24-bpp path
    If (m_Header.BitDepth = 8) Then
        
        'Calculate stride manually for source and destination lines
        srcStride = srcImage.GetDIBStride
        If (srcImage.GetDIBColorDepth = 32) Then srcBPP = 4 Else srcBPP = 3
        
        dstStride = srcImage.GetDIBWidth * 3
        xLimit = srcImage.GetDIBWidth - 1
        
        srcImage.WrapArrayAroundScanline srcBytes, srcSA1D, 0
        srcPtr = srcSA1D.pvData
        
        'Next, we want to copy all byte data into place - but note that the Windows DIB will have a different
        ' stride than the PNG DIB (which does *not* pad to increments of 4).
        For y = 0 To srcHeightInPixels - 1
            
            yDstOffset = y * dstStride
            
            'Wrap an array around the source data
            srcSA1D.pvData = srcPtr + srcStride * y
            
            'Copy bytes manually, remembering to swizzle r/b
            For x = 0 To xLimit
                curOffset = yDstOffset + x * 3
                finalBytes(curOffset) = srcBytes(x * srcBPP + 2)
                finalBytes(curOffset + 1) = srcBytes(x * srcBPP + 1)
                finalBytes(curOffset + 2) = srcBytes(x * srcBPP)
            Next x
            
        Next y
        
    '48-bpp path
    ElseIf (m_Header.BitDepth = 16) Then
    
        srcStride = srcImage.GetDIBStride
        dstStride = srcImage.GetDIBWidth * 6
        xLimit = srcImage.GetDIBWidth - 1
        
        For y = 0 To srcHeightInPixels - 1
            
            yDstOffset = y * dstStride
            
            'Wrap an array around the source data
            srcImage.WrapArrayAroundScanline srcBytes, srcSA1D, y
            
            For x = 0 To xLimit
                
                'Swizzle R and B as part of the copy
                curOffset = yDstOffset + x * 6
                finalBytes(curOffset) = srcBytes(x * 4 + 2)
                finalBytes(curOffset + 2) = srcBytes(x * 4 + 1)
                finalBytes(curOffset + 4) = srcBytes(x * 4)
                
            Next x
            
        Next y
        
    End If
    
    'Make sure our array is unwrapped from the source data
    srcImage.UnwrapArrayFromDIB srcBytes
    
    Step3Helper_Truecolor = png_Success
    
End Function

Private Function Step3Helper_Grayscale(ByRef srcImage As pdDIB, ByRef cStream As pdStream, ByRef finalBytes() As Byte, ByRef sizeOfPngBytes As Long, ByRef dstStride As Long, Optional ByVal writeAncillaryChunks As Boolean = True) As PD_PNGResult
    
    'Grayscale mode supports 1/2/4/8/16-bit output
    Dim tmpByte As Byte, srcHeightInPixels As Long
    Dim xLimit As Long, yDstOffset As Long, dstOffsetX As Long, curOffset As Long
    Dim x As Long, y As Long
    
    'Some params are identical for all image types (such as number of scanlines)
    srcHeightInPixels = srcImage.GetDIBHeight
    
    'If this is a static PNG (or the first frame of an animated PNG), write out any extra data before
    ' writing out pixel bits
    If writeAncillaryChunks Then
    
        'If the caller wants us to write a tRNS chunk, do so now.
        If m_WriteTrns Then WriteTRNSNow cStream
        
        'If the caller wants us to write a bKGD chunk, do so now.
        If m_EmbedBkgdColor Then WriteBkgdChunk cStream
        
    End If
    
    'If it isn't already, composite the source image against the specified backcolor
    srcImage.ConvertTo24bpp m_CompositingColor
    
    'Retrieve a grayscale-only copy of the source image's bytes.  (Note that this will forcibly convert the
    ' pixel data to grayscale if it isn't already, so it doesn't really matter whether the data is already gray!)
    Dim grayBytes() As Byte
    DIBs.GetDIBGrayscaleMapEx srcImage, grayBytes, m_ImgMaxColorCount
    
    'Cheap copy for 8-bits, manual transfer for 16, special encoding for sub-8-bit
    If (m_Header.BitDepth = 8) Then
        dstStride = srcImage.GetDIBWidth
        CopyMemoryStrict VarPtr(finalBytes(0)), VarPtr(grayBytes(0, 0)), sizeOfPngBytes
    ElseIf (m_Header.BitDepth = 16) Then
    
        dstStride = srcImage.GetDIBWidth * 2
        xLimit = (srcImage.GetDIBWidth - 1)
        
        For y = 0 To srcHeightInPixels - 1
            yDstOffset = y * dstStride
            For x = 0 To xLimit
                finalBytes(yDstOffset + x * 2) = grayBytes(x, y)
            Next x
        Next y
    
    ElseIf (m_Header.BitDepth = 4) Then
    
        'Because VB lacks shift operators (uuuugh) we have to write these bytes idiotically.
        ' (TODO: remove inner loop division by pre-calculating a lut with scanline indices)
        xLimit = (srcImage.GetDIBWidth - 1)
        For y = 0 To srcHeightInPixels - 1
            
            yDstOffset = y * dstStride
            
            For x = 0 To xLimit
                dstOffsetX = yDstOffset + x \ 2
                tmpByte = grayBytes(x, y) \ 17
                If (x And 1) Then
                    finalBytes(dstOffsetX) = finalBytes(dstOffsetX) Or tmpByte
                Else
                    finalBytes(dstOffsetX) = tmpByte * 16
                End If
            Next x
            
        Next y
        
    ElseIf (m_Header.BitDepth = 2) Then
    
        'Because VB lacks shift operators (uuuugh) we have to write these bytes idiotically.
        xLimit = (srcImage.GetDIBWidth - 1)
        For y = 0 To srcHeightInPixels - 1
            
            yDstOffset = y * dstStride
            curOffset = 0
            
            For x = 0 To xLimit
                dstOffsetX = yDstOffset + x \ 4
                tmpByte = grayBytes(x, y) \ 85
                If (curOffset = 0) Then
                    finalBytes(dstOffsetX) = tmpByte * 64
                ElseIf (curOffset = 1) Then
                    finalBytes(dstOffsetX) = finalBytes(dstOffsetX) Or (tmpByte * 16)
                ElseIf (curOffset = 2) Then
                    finalBytes(dstOffsetX) = finalBytes(dstOffsetX) Or (tmpByte * 4)
                Else
                    finalBytes(dstOffsetX) = finalBytes(dstOffsetX) Or tmpByte
                End If
                curOffset = curOffset + 1
                If (curOffset > 3) Then curOffset = 0
            Next x
            
        Next y
    
    ElseIf (m_Header.BitDepth = 1) Then
    
        'Because VB lacks shift operators (uuuugh) we have to write these bytes idiotically.
        ' TODO: build a lut of multipliers; use that instead of branches
        xLimit = (srcImage.GetDIBWidth - 1)
        For y = 0 To srcHeightInPixels - 1
            
            yDstOffset = y * dstStride
            curOffset = 0
            
            For x = 0 To xLimit
                dstOffsetX = yDstOffset + x \ 8
                If (grayBytes(x, y) < 128) Then tmpByte = 0 Else tmpByte = 1
                If (curOffset = 0) Then
                    finalBytes(dstOffsetX) = tmpByte * 128
                ElseIf (curOffset = 1) Then
                    finalBytes(dstOffsetX) = finalBytes(dstOffsetX) Or (tmpByte * 64)
                ElseIf (curOffset = 2) Then
                    finalBytes(dstOffsetX) = finalBytes(dstOffsetX) Or (tmpByte * 32)
                ElseIf (curOffset = 3) Then
                    finalBytes(dstOffsetX) = finalBytes(dstOffsetX) Or (tmpByte * 16)
                ElseIf (curOffset = 4) Then
                    finalBytes(dstOffsetX) = finalBytes(dstOffsetX) Or (tmpByte * 8)
                ElseIf (curOffset = 5) Then
                    finalBytes(dstOffsetX) = finalBytes(dstOffsetX) Or (tmpByte * 4)
                ElseIf (curOffset = 6) Then
                    finalBytes(dstOffsetX) = finalBytes(dstOffsetX) Or (tmpByte * 2)
                ElseIf (curOffset = 7) Then
                    finalBytes(dstOffsetX) = finalBytes(dstOffsetX) Or tmpByte
                End If
                curOffset = curOffset + 1
                If (curOffset > 7) Then curOffset = 0
            Next x
            
        Next y
    
    End If
    
    Step3Helper_Grayscale = png_Success
    
End Function

Private Function Step3Helper_GrayscaleAlpha(ByRef srcImage As pdDIB, ByRef cStream As pdStream, ByRef finalBytes() As Byte, ByRef sizeOfPngBytes As Long, ByRef dstStride As Long, Optional ByVal writeAncillaryChunks As Boolean = True) As PD_PNGResult
    
    'Grayscale+Alpha mode supports 16/32-bpp output
    Dim srcHeightInPixels As Long
    Dim xLimit As Long, yDstOffset As Long, curOffset As Long
    Dim x As Long, y As Long
    
    'If the caller wants us to write a bKGD chunk, do so now.
    If m_EmbedBkgdColor And writeAncillaryChunks Then WriteBkgdChunk cStream
    
    'Some params are identical for all image types (such as number of scanlines)
    srcHeightInPixels = srcImage.GetDIBHeight
    
    'Retrieve a grayscale+alpha copy of the source image's bytes
    Dim grayAlphaBytes() As Byte
    DIBs.GetDIBGrayscaleAndAlphaMap srcImage, grayAlphaBytes
    
    'Cheap copy for 8-bits, manual transfer for 16
    If (m_Header.BitDepth = 8) Then
        dstStride = srcImage.GetDIBWidth * 2
        CopyMemoryStrict VarPtr(finalBytes(0)), VarPtr(grayAlphaBytes(0, 0)), sizeOfPngBytes
    ElseIf (m_Header.BitDepth = 16) Then
    
        dstStride = srcImage.GetDIBWidth * 4
        xLimit = (srcImage.GetDIBWidth * 2 - 1)
        
        For y = 0 To srcHeightInPixels - 1
            
            yDstOffset = y * dstStride
            
            For x = 0 To xLimit Step 2
                curOffset = yDstOffset + x * 2
                finalBytes(curOffset) = grayAlphaBytes(x, y)
                finalBytes(curOffset + 2) = grayAlphaBytes(x + 1, y)
            Next x
            
        Next y
        
    End If
    
    Step3Helper_GrayscaleAlpha = png_Success
    
End Function

'Unlike other modes, indexed color images require a reference to the underlying pdStream object because they must
' write a separate palette chunk (and perhaps an optional transparency chunk)
Private Function Step3Helper_Indexed(ByRef srcImage As pdDIB, ByRef cStream As pdStream, ByRef finalBytes() As Byte, ByRef sizeOfPngBytes As Long, ByRef dstStride As Long, Optional ByVal writeAncillaryChunks As Boolean = True) As PD_PNGResult
    
    'Indexed mode supports 1/2/4/8-bit output
    Dim srcHeightInPixels As Long
    Dim xLimit As Long, yDstOffset As Long, dstOffsetX As Long, curOffset As Long
    Dim x As Long, y As Long
    
    'Unlike other color modes, note that dstStride has already been calculated for us; this was necessary
    ' to size the finalBytes() array.
    
    'Some params are identical for all image types (such as number of scanlines)
    srcHeightInPixels = srcImage.GetDIBHeight
    
    'Retrieve a copy of the image and its palette.  If the module-level palette was previously marked
    ' as "safe", it means the image analyzer has already produced a palette for us.  (If "auto" output
    ' settings are selected, the analyzer has to produce a palette anyway to determine how many
    ' unique colors are in the image - so we can just reuse it now!
    Dim imgPixels() As Byte, imgPalette() As RGBQuad, numColors As Long
    If m_ImgPaletteSafe Then
        numColors = DIBs.GetDIBAs8bpp_RGBA_SrcPalette(srcImage, m_ImgPalette, imgPixels)
    Else
        numColors = DIBs.GetDIBAs8bpp_RGBA(srcImage, imgPalette, imgPixels)
    End If
    
    'If the source image has more than the user's requested color count (256 by default), we need to
    ' forcibly reduce it before continuing.
    Dim maxColorCount As Long
    maxColorCount = m_ImgMaxColorCount
    If (maxColorCount > 256) Then maxColorCount = 256
    
    If (numColors > maxColorCount) Then
        m_ImgPaletteSafe = False
        numColors = DIBs.GetDIBAs8bpp_RGBA_Forcibly(srcImage, imgPalette, imgPixels, maxColorCount)
    End If
    
    'Ensure the module-level palette (m_ImgPalette) and our local palette (imgPalette) are in agreement.
    ' If we had to forcibly downsample the source image due to caller settings, our locally produced
    ' palette gets precedence - otherwise, the module-level palette will be used.
    
    '(It's critical that these two palettes are in agreement; if they vary, subsequent steps - like
    ' writing a bKGD chunk - will produce bad files because palette indices won't be in alignment.)
    If m_ImgPaletteSafe Then
        ReDim imgPalette(0 To numColors - 1) As RGBQuad
        CopyMemoryStrict VarPtr(imgPalette(0)), VarPtr(m_ImgPalette(0)), numColors * 4
    Else
        ReDim m_ImgPalette(0 To numColors - 1) As RGBQuad
        CopyMemoryStrict VarPtr(m_ImgPalette(0)), VarPtr(imgPalette(0)), numColors * 4
    End If
    
    'Before passing pixel data off to the encoder, we need to write a palette (and possibly transparency) chunk.
    ' (Note that the "writeAncillaryChunks" parameter here is not accurately named - PLTE is a *critical* chunk,
    '  but because we also support animated PNGs, I needed a way to flag chunks that don't need these extra,
    '  non-IDAT/fdAT chunks prepended!)
    If writeAncillaryChunks Then
        
        Dim palChunk As pdPNGChunk
        Set palChunk = New pdPNGChunk
        palChunk.CreateChunkForExport "PLTE", numColors * 3
        
        With palChunk.BorrowData
        
            Dim i As Long
            For i = 0 To numColors - 1
                .WriteByte imgPalette(i).Red
                .WriteByte imgPalette(i).Green
                .WriteByte imgPalette(i).Blue
            Next i
            
        End With
        
        palChunk.FinalizeChunk
        
        cStream.WriteStream palChunk.BorrowData
        Set palChunk = Nothing
        
    End If
        
    'If any of the palette indices are non-zero, we also need to write a tRNS chunk
    Dim palHasTransparency As Boolean
    palHasTransparency = False
    
    For i = 0 To numColors - 1
        If (imgPalette(i).Alpha <> 255) Then
            palHasTransparency = True
            Exit For
        End If
    Next i
    
    If palHasTransparency And writeAncillaryChunks Then
    
        Dim trnsChunk As pdPNGChunk
        Set trnsChunk = New pdPNGChunk
        trnsChunk.CreateChunkForExport "tRNS", numColors
        
        'The spec comments that the tRNS chunk doesn't need to be the same size as the palette chunk.
        ' Any palette entries without a matching tRNS entry are assumed to have opacity 255.  Because of
        ' this quirk, let's scan the palette from the *end* and look for the first non-255 value.
        ' (This may allow us to shave a few bytes off the file by not writing redundant 255s to the table.)
        Dim trnsTerminal As Long
        For i = numColors - 1 To 0 Step -1
            If (imgPalette(i).Alpha <> 255) Then
                trnsTerminal = i
                Exit For
            End If
        Next i
        
        With trnsChunk.BorrowData
            For i = 0 To trnsTerminal
                .WriteByte imgPalette(i).Alpha
            Next i
        End With
        
        trnsChunk.FinalizeChunk
        cStream.WriteStream trnsChunk.BorrowData
        Set trnsChunk = Nothing
        
    End If
    
    'If the caller wants us to write a bKGD chunk, do so now.
    If m_EmbedBkgdColor And writeAncillaryChunks Then WriteBkgdChunk cStream
    
    'Finally, we need to copy the pixel data into the final "ready to write" buffer.  For 8-bpp data,
    ' this is as simple as CopyMemory.  For lower bit-depths, we need to transform the data into
    ' its appropriate final bitness.
    If (m_Header.BitDepth = 8) Then
        CopyMemoryStrict VarPtr(finalBytes(0)), VarPtr(imgPixels(0, 0)), sizeOfPngBytes
    ElseIf (m_Header.BitDepth = 4) Then
        
        'Because VB lacks shift operators (uuuugh) we have to write these bytes idiotically.
        ' (TODO: remove inner loop division by pre-calculating a lut with scanline indices)
        xLimit = (srcImage.GetDIBWidth - 1)
        For y = 0 To srcHeightInPixels - 1
            
            yDstOffset = y * dstStride
            
            For x = 0 To xLimit
                dstOffsetX = yDstOffset + x \ 2
                If (x And 1) Then
                    finalBytes(dstOffsetX) = finalBytes(dstOffsetX) Or imgPixels(x, y)
                Else
                    finalBytes(dstOffsetX) = imgPixels(x, y) * 16
                End If
            Next x
            
        Next y
        
    ElseIf (m_Header.BitDepth = 2) Then
    
        'Because VB lacks shift operators (uuuugh) we have to write these bytes idiotically.
        xLimit = (srcImage.GetDIBWidth - 1)
        For y = 0 To srcHeightInPixels - 1
            
            yDstOffset = y * dstStride
            curOffset = 0
            
            For x = 0 To xLimit
                dstOffsetX = yDstOffset + x \ 4
                If (curOffset = 0) Then
                    finalBytes(dstOffsetX) = imgPixels(x, y) * 64
                ElseIf (curOffset = 1) Then
                    finalBytes(dstOffsetX) = finalBytes(dstOffsetX) Or (imgPixels(x, y) * 16)
                ElseIf (curOffset = 2) Then
                    finalBytes(dstOffsetX) = finalBytes(dstOffsetX) Or (imgPixels(x, y) * 4)
                Else
                    finalBytes(dstOffsetX) = finalBytes(dstOffsetX) Or imgPixels(x, y)
                End If
                curOffset = curOffset + 1
                If (curOffset > 3) Then curOffset = 0
            Next x
            
        Next y
    
    ElseIf (m_Header.BitDepth = 1) Then
    
        'Because VB lacks shift operators (uuuugh) we have to write these bytes idiotically.
        xLimit = (srcImage.GetDIBWidth - 1)
        For y = 0 To srcHeightInPixels - 1
            
            yDstOffset = y * dstStride
            curOffset = 0
            
            For x = 0 To xLimit
                dstOffsetX = yDstOffset + x \ 8
                If (curOffset = 0) Then
                    finalBytes(dstOffsetX) = imgPixels(x, y) * 128
                ElseIf (curOffset = 1) Then
                    finalBytes(dstOffsetX) = finalBytes(dstOffsetX) Or (imgPixels(x, y) * 64)
                ElseIf (curOffset = 2) Then
                    finalBytes(dstOffsetX) = finalBytes(dstOffsetX) Or (imgPixels(x, y) * 32)
                ElseIf (curOffset = 3) Then
                    finalBytes(dstOffsetX) = finalBytes(dstOffsetX) Or (imgPixels(x, y) * 16)
                ElseIf (curOffset = 4) Then
                    finalBytes(dstOffsetX) = finalBytes(dstOffsetX) Or (imgPixels(x, y) * 8)
                ElseIf (curOffset = 5) Then
                    finalBytes(dstOffsetX) = finalBytes(dstOffsetX) Or (imgPixels(x, y) * 4)
                ElseIf (curOffset = 6) Then
                    finalBytes(dstOffsetX) = finalBytes(dstOffsetX) Or (imgPixels(x, y) * 2)
                ElseIf (curOffset = 7) Then
                    finalBytes(dstOffsetX) = finalBytes(dstOffsetX) Or imgPixels(x, y)
                End If
                curOffset = curOffset + 1
                If (curOffset > 7) Then curOffset = 0
            Next x
            
        Next y
    
    End If
    
    Step3Helper_Indexed = png_Success
    
End Function

'After populating all class-level tRNS data, call this function to immediately write the actual tRNS chunk.
' (Note that tRNS order matters: it needs to be after PLTE but before IDAT.)
Private Sub WriteTRNSNow(ByRef cStream As pdStream)

    Dim tmpChunk As pdPNGChunk
    Set tmpChunk = New pdPNGChunk
    
    tmpChunk.CreateChunkForExport "tRNS"
    
    If (m_Header.ColorType = png_Truecolor) Then
        If (m_Header.BitDepth = 8) Then
            tmpChunk.BorrowData.WriteInt_BE m_TrnsValueR
            tmpChunk.BorrowData.WriteInt_BE m_TrnsValueG
            tmpChunk.BorrowData.WriteInt_BE m_TrnsValueB
        ElseIf (m_Header.BitDepth = 16) Then
            tmpChunk.BorrowData.WriteIntU_BE m_TrnsValueR * 256
            tmpChunk.BorrowData.WriteIntU_BE m_TrnsValueG * 256
            tmpChunk.BorrowData.WriteIntU_BE m_TrnsValueB * 256
        End If
    
    ElseIf (m_Header.ColorType = png_Greyscale) Then
        If (m_Header.BitDepth = 8) Then
            tmpChunk.BorrowData.WriteInt_BE m_TrnsValueG
        ElseIf (m_Header.BitDepth = 16) Then
            tmpChunk.BorrowData.WriteIntU_BE m_TrnsValueG * 256
        ElseIf (m_Header.BitDepth = 4) Then
            tmpChunk.BorrowData.WriteInt_BE m_TrnsValueG \ 17
        ElseIf (m_Header.BitDepth = 2) Then
            tmpChunk.BorrowData.WriteInt_BE m_TrnsValueG \ 85
        ElseIf (m_Header.BitDepth = 1) Then
            tmpChunk.BorrowData.WriteInt_BE m_TrnsValueG \ 255
        End If
    
    'Paletted images write tRNS data without our help; look there for data on how it constructs its specialized
    ' tRNS entry.
    End If
    
    tmpChunk.FinalizeChunk
    cStream.WriteStream tmpChunk.BorrowData
    
End Sub

'This function requires a temporary chunk object that has already filtered pixel bytes; this is handled by export step 3,
' which *MUST* be called before calling this function.
Friend Function ExportStep4_WriteIDAT(ByRef cStream As pdStream, ByRef tmpChunk As pdPNGChunk, Optional ByVal cmpLevel As Long = -1, Optional ByVal writefdATInstead As Boolean = False) As PD_PNGResult
    
    Dim startTime As Currency
    VBHacks.GetHighResTime startTime
    
    'Get a pointer and length of the filtered bytes
    Dim srcPtr As Long, srcLen As Long
    tmpChunk.GetFilteredBytesData srcPtr, srcLen
    
    'Prep a temporary buffer to receive the compressed bytes
    Dim tmpBytes() As Byte
    
    'Determine compression level.  Note that a libdeflate setting of 7 correlates pretty closely to a
    ' final deflate stream size of zlib level 9.  At matching levels, however, libdeflate will produce
    ' *faster* results.  (And of course, if you ratchet its level up to a max of 12 you'll get extremely
    ' good compression, at some penalty to processing speed.)
    Dim cmpSize As Long
    If (cmpLevel < 0) Then
        cmpLevel = Compression.GetDefaultCompressionLevel(cf_Zlib)
    
    'Note that old versions of libdeflate did not support "none" compression, so we needed to manually
    ' specify a minimum level of 1.  This was changed in late 2020; see https://github.com/ebiggers/libdeflate/issues/86
    'ElseIf (cmpLevel = 0) Then
    '    cmpLevel = 1
    ElseIf (cmpLevel > Compression.GetMaxCompressionLevel(cf_Zlib)) Then
        cmpLevel = Compression.GetMaxCompressionLevel(cf_Zlib)
    End If
    
    If PNG_DEBUG_VERBOSE And (Not writefdATInstead) Then PDDebug.LogAction "- Requested DEFLATE compression level: " & cmpLevel
    
    'This next bit of code is a failsafe so that this class can support two different write modes:
    ' memory-mapped files (which are faster but use slightly more memory) and traditional file writes
    ' (slower but minimal memory usage).  In June 2019, I switched over to the memory-mapped
    ' interface because of its obvious perf benefits, but I'm leaving the old interface in case I
    ' ever decide to revert back.
    If (cStream.GetStreamMode = PD_SM_FileBacked) Then
        
        'NOTE: this file-backed approach does *not* currently support animated PNGs!
        
        Compression.CompressPtrToDstArray tmpBytes, cmpSize, srcPtr, srcLen, cf_Zlib, cmpLevel
        
        If PNG_DEBUG_VERBOSE And (Not writefdATInstead) Then
            PDDebug.LogAction "- IDAT write: DEFLATE took " & VBHacks.GetTimeDiffNowAsString(startTime)
            VBHacks.GetHighResTime startTime
        End If
        
        'We're done with the temporary chunk, and because it consumes a *lot* of memory, free it immediately
        Set tmpChunk = Nothing
        
        'Normally, we would use a temporary chunk object to write the actual chunk data (including CRC calculation).
        ' Because IDAT chunks are potentially enormous, however, we want to manually embed the final, compressed data
        ' stream directly inside the target file.
        
        'Start with data length
        cStream.WriteLong_BE cmpSize
        
        'Next, the chunk name
        cStream.WriteString_ASCII "IDAT"
        
        'Next, the data itself
        cStream.WriteBytesFromPointer VarPtr(tmpBytes(0)), cmpSize
        
        If PNG_DEBUG_VERBOSE And (Not writefdATInstead) Then
            PDDebug.LogAction "- IDAT write: copy IDAT to file took " & VBHacks.GetTimeDiffNowAsString(startTime)
            VBHacks.GetHighResTime startTime
        End If
        
        'Finally, a CRC value.  We need to manually calculate this, and the calculation must include the
        ' chunk name (but not its length).  Start by calculating a CRC for the chunk name.
        Dim nameBytes() As Byte
        nameBytes = StrConv("IDAT", vbFromUnicode)
        
        Dim crcBase As Long, crcFinal As Long
        crcBase = Plugin_libdeflate.GetCrc32(0&, 0&, 0&, True)
        crcBase = Plugin_libdeflate.GetCrc32(VarPtr(nameBytes(0)), 4&, crcBase, False)
        
        'Next, compute a CRC over the compressed data - but make sure we pass the *existing* CRC value
        ' to ensure a correct final value!
        crcFinal = Plugin_libdeflate.GetCrc32(VarPtr(tmpBytes(0)), cmpSize, crcBase, False)
        
        If PNG_DEBUG_VERBOSE And (Not writefdATInstead) Then
            PDDebug.LogAction "- IDAT write: checksum calculation took " & VBHacks.GetTimeDiffNowAsString(startTime)
            VBHacks.GetHighResTime startTime
        End If
        
        'Write the final CRC out to file
        cStream.WriteLong_BE crcFinal
    
    'Memory-based approaches can use a faster approach, where we compress directly into the final buffer
    ' (instead of wasting time and energy on an intermediary one).
    Else
        
        'Ensure the stream has sufficient space for a worst-case compression, plus 16 bytes for
        ' a PNG chunk ID, size, CRC checksum, and (optionally) an animated PNG frame sequence number.
        Dim worstCaseCompressedSize As Long
        worstCaseCompressedSize = Compression.GetWorstCaseSize(srcLen, cf_Zlib, cmpLevel)
        If writefdATInstead Then worstCaseCompressedSize = worstCaseCompressedSize + 16 Else worstCaseCompressedSize = worstCaseCompressedSize + 12
        cStream.EnsureBufferSpaceAvailable worstCaseCompressedSize
        
        'Compress directly into the final buffer, and remember that it needs to be offset by 8
        ' in an IDAT chunk (for the PNG chunk ID and size, which preface the pixel data), and 12
        ' in an fdAT chunk (because the frame sequence will also be prepended)
        cmpSize = worstCaseCompressedSize
        
        Dim peekOffset As Long
        If writefdATInstead Then peekOffset = 12 Else peekOffset = 8
        Compression.CompressPtrToPtr cStream.Peek_PointerOnly(cStream.GetPosition + peekOffset, worstCaseCompressedSize - peekOffset), cmpSize, srcPtr, srcLen, cf_Zlib, cmpLevel
        
        If PNG_DEBUG_VERBOSE And (Not writefdATInstead) Then
            PDDebug.LogAction "- IDAT write: DEFLATE took " & VBHacks.GetTimeDiffNowAsString(startTime)
            VBHacks.GetHighResTime startTime
        End If
        
        'We're done with the temporary chunk, and because it consumes a *lot* of memory, free it immediately
        Set tmpChunk = Nothing
        
        'At the current (unmoved) pointer position, write the final compressed length.
        ' (On APNGs, we must manually append 4 for the frame sequence value)
        If writefdATInstead Then
            cStream.WriteLong_BE cmpSize + 4
        Else
            cStream.WriteLong_BE cmpSize
        End If
        
        'Next, the chunk name
        If writefdATInstead Then
            cStream.WriteString_ASCII "fdAT"
        Else
            cStream.WriteString_ASCII "IDAT"
        End If
        
        'For an fdAT stream, write the sequence number and increment it
        If writefdATInstead Then
            cStream.WriteLong_BE m_FrameSequence
            m_FrameSequence = m_FrameSequence + 1
        End If
        
        'Next, the compressed data stream already exists - so we don't need to write it!  We simply need to
        ' (manually) notify the stream that the data exists.
        Dim posData As Long
        posData = cStream.GetPosition()
        cStream.SetSizeExternally posData + cmpSize
        
        If PNG_DEBUG_VERBOSE And (Not writefdATInstead) Then
            PDDebug.LogAction "- IDAT write: copy IDAT to file took " & VBHacks.GetTimeDiffNowAsString(startTime)
            VBHacks.GetHighResTime startTime
        End If
        
        'Finally, a CRC value.  We need to manually calculate this, and the calculation must include the
        ' chunk name (but not its length) and obviously the compressed data itself.
        If writefdATInstead Then peekOffset = 8 Else peekOffset = 4
        crcFinal = Plugin_libdeflate.GetCrc32(cStream.Peek_PointerOnly(posData - peekOffset, cmpSize + peekOffset), cmpSize + peekOffset, 0&, True)
        
        If PNG_DEBUG_VERBOSE And (Not writefdATInstead) Then
            PDDebug.LogAction "- IDAT write: checksum calculation took " & VBHacks.GetTimeDiffNowAsString(startTime)
            VBHacks.GetHighResTime startTime
        End If
        
        'Manually move the file pointer to the end of the compressed data (remember - the pointer hasn't moved
        ' since we wrote the chunk name (and possible frame sequence), then write the final CRC out to file.
        cStream.SetPosition cmpSize, FILE_CURRENT
        cStream.WriteLong_BE crcFinal
        
    End If
    
    ExportStep4_WriteIDAT = png_Success
    
End Function

'Make sure you have populated the module-level m_Header struct before calling this function!
Friend Function ExportStep5_WriteIEND(ByRef cStream As pdStream) As PD_PNGResult

    'Initialize a chunk object
    Dim dstChunk As pdPNGChunk
    Set dstChunk = New pdPNGChunk
    
    'The end chunk is always 0-bytes long.  (Note that we never include length, chunk type, or CRC lengths -
    ' the chunk class handles all that for us.)
    dstChunk.CreateChunkForExport "IEND", 0
    
    'Immediately finalize the chunk and copy its contents into the final stream
    dstChunk.FinalizeChunk
    cStream.WriteStream dstChunk.BorrowData
    
    ExportStep5_WriteIEND = png_Success
    
End Function

'For non-paletted images (e.g. 16-bpp or more), call this function to convert an interlaced IDAT stream
' to a de-interlaced one.  (8-bpp are handled separately, using a function optimized specifically for them.)
Private Sub DeInterlaceNonPaletteData(ByVal idatIndex As Long, ByVal pxWidth As Long, ByVal pxHeight As Long)

    'This function only handles deinterlacing for bit-depths higher than 8-bpp
    If (m_Header.BitsPerPixel < 8) Or (Not m_Header.Interlaced) Then Exit Sub
    
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
    SeparateInterlacedIDATs idatIndex, tmpChunks, False
        
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
    Dim bytesPerPixel As Long, bppLoop As Long
    bytesPerPixel = (m_Header.BitsPerPixel \ 8)
    bppLoop = bytesPerPixel - 1
    
    Dim newIdatSize As Long
    newIdatSize = (pxWidth * pxHeight * bytesPerPixel) + pxHeight
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
    xFinal = pxWidth - 1
    yFinal = pxHeight - 1
    newStride = (pxWidth * bytesPerPixel) + 1
        
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
                    dstBytes(1 + (y * newStride) + (x * bytesPerPixel) + j) = srcBytes(srcPxOffset)
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
    m_XPixelCount(0) = pxWidth
    
    ReDim m_YPixelCount(0) As Long
    m_YPixelCount(0) = pxHeight
    
    ReDim m_XByteCount(0) As Long
    m_XByteCount(0) = (pxWidth * bytesPerPixel) + 1
    
End Sub

'LittleCMS is an integral part of the PNG load process.  If PNG images contain ICC profiles, we use LittleCMS
' to translate colors between spaces.  If PNG images do *not* contain ICC profiles, we can still use LittleCMS
' for faster swizzling than VB6 alone can provide.  (As such, we still prep generic ICC profiles as part of
' the preparation stage.)
Private Sub PrepICCData(ByRef srcProfile As pdLCMSProfile, ByRef dstProfile As pdLCMSProfile, ByRef dstTransform As pdLCMSTransform, ByRef srcImgProfileGood As Boolean)
    
    Set srcProfile = New pdLCMSProfile
    
    'In 2025, the PNG spec was updated with a newly defined cICP chunk which simplifies defining HDR data.
    ' This small chunk specifies a series of hard-coded IDs for known popular HDR color formats.  Decoders can
    ' use these chunks to handle pixel data accordingly.
    '
    'If a cICP chunk is present, it takes precedence over other color profile chunks.
    Dim idxCICP As Long
    If ColorManagement.UseEmbeddedICCProfiles() Then
        idxCICP = Me.GetIndexOfChunk("cICP")
    Else
        idxCICP = -1
    End If
    
    Dim cICPOK As Boolean
    If (idxCICP > 0) Then
        cICPOK = PrepcICPData(idxCICP, srcProfile, srcImgProfileGood)
    Else
        cICPOK = False
    End If
    
    If (Not cICPOK) Then
        
        'Users can set a preference for ignoring embedded profiles (this allows recovery of images with broken or faulty ICC)
        Dim iccIndexV2 As Long, iccIndexV4 As Long
        If ColorManagement.UseEmbeddedICCProfiles() Then
            iccIndexV4 = Me.GetIndexOfChunk("iCCN")
            iccIndexV2 = Me.GetIndexOfChunk("iCCP")
        Else
            iccIndexV4 = -1
            iccIndexV2 = -1
        End If
        
        'New in 2025 is the possibility of a v4 ICC profile; check that first (it has priority per the spec)
        If (iccIndexV4 > 0) Then
            srcImgProfileGood = srcProfile.CreateFromPointer(m_Chunks(iccIndexV4).BorrowData().Peek_PointerOnly(0), m_Chunks(iccIndexV4).BorrowData().GetStreamSize())
            
            'If the image does not contain color profile data, or if the embedded profile is invalid, assumed sRGB.
            ' (TODO: expose this to the user as a preference; untagged images may reside in a different space.)
            If (Not srcImgProfileGood) Then
                m_Warnings.AddString "WARNING!  ICC v4 profile found in file, but LittleCMS couldn't validate the profile.  Color management will *not* be applied using this profile!"
                iccIndexV4 = -1
            End If
            
        End If
            
        If (iccIndexV2 > 0) And (iccIndexV4 <= 0) Then
            srcImgProfileGood = srcProfile.CreateFromPointer(m_Chunks(iccIndexV2).BorrowData().Peek_PointerOnly(0), m_Chunks(iccIndexV2).BorrowData().GetStreamSize())
            
            'If the image does not contain color profile data, or if the embedded profile is invalid, assumed sRGB.
            ' (TODO: expose this to the user as a preference; untagged images may reside in a different space.)
            If (Not srcImgProfileGood) Then
                m_Warnings.AddString "WARNING!  ICC v2 profile found in file, but LittleCMS couldn't validate the profile.  Color management will *not* be applied using this profile!"
                iccIndexV2 = -1
            End If
            
        End If
        
    End If
    
    'If ICC processing failed, create a generic source profile.
    If (iccIndexV4 <= 0) And (iccIndexV2 <= 0) And (Not cICPOK) Then
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
    
    'In most (?) cases, we want to use LittleCMS to color-manage incoming PNG data.  There are exceptions, however;
    ' indexed images without embedded profiles don't require management, for example.  As such, we can skip specialized
    ' ICC handling under some scenarios.
    Dim iccRequired As Boolean: iccRequired = True
    
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
        
        'RGBA images can always use LittleCMS
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
            m_Warnings.AddString "WARNING!  pdPNG.ImportStage5_ConstructImage failed to create a valid ICC transform.  Color management disabled for this image."
            Set dstTransform = Nothing
        End If
    End If

End Sub

'Build a color profile for an embedded cICP chunk, if any.
Private Function PrepcICPData(ByVal idxCICP As Long, ByRef srcProfile As pdLCMSProfile, ByRef srcImgProfileGood As Boolean) As Boolean
    
    Const FUNC_NAME As String = "PrepcICPData"
    
    Dim cICPBytes() As Byte
    If m_Chunks(idxCICP).GetcICPData(cICPBytes, m_Warnings) Then
        
        If PNG_DEBUG_VERBOSE Then PDDebug.LogAction "cICP chunk found: " & cICPBytes(0) & ", " & cICPBytes(1) & ", " & cICPBytes(2) & ", " & cICPBytes(3)
        
        'Validate that the 3rd byte is 0 (required by the spec)
        If (cICPBytes(2) <> 0) Then
            PrepcICPData = False
            Exit Function
        End If
        
        'Now we need to build a color profile from the hard-coded list of potential colors and primaries defined by the
        ' list of cICP bytes.
        Dim wpValues() As Double, rgbValues() As Double
        ReDim wpValues(0 To 3) As Double
        ReDim rgbValues(0 To 9) As Double
        
        'The right-most column of all color definitions are 1.0
        wpValues(2) = 1#
        rgbValues(2) = 1#
        rgbValues(5) = 1#
        rgbValues(8) = 1#
        
        'Some indices are reserved, or deliberately marked as "do not use", or are otherwise invalid.
        Dim idxOK As Boolean
        idxOK = False
        
        'All values in this table are taken from https://www.itu.int/rec/T-REC-H.273-202407-I/en (update 07/2024)
        Select Case cICPBytes(0)
            
            Case 0
                'Reserved
            Case 1
                'sRGB
                idxOK = True
                rgbValues(0) = 0.64
                rgbValues(1) = 0.33
                
                rgbValues(3) = 0.3
                rgbValues(4) = 0.6
                
                rgbValues(6) = 0.15
                rgbValues(7) = 0.06
                
                wpValues(0) = 0.3127
                wpValues(1) = 0.329
            Case 2
                'Unspecified
            Case 3
                'Reserved
            Case 4
                '1953 NTSC
                idxOK = True
                rgbValues(0) = 0.67
                rgbValues(1) = 0.33
                
                rgbValues(3) = 0.21
                rgbValues(4) = 0.71
                
                rgbValues(6) = 0.14
                rgbValues(7) = 0.08
                
                wpValues(0) = 0.31
                wpValues(1) = 0.316
            Case 5
                'Historical PAL/SECAM
                idxOK = True
                rgbValues(0) = 0.64
                rgbValues(1) = 0.33
                
                rgbValues(3) = 0.29
                rgbValues(4) = 0.6
                
                rgbValues(6) = 0.15
                rgbValues(7) = 0.06
                
                wpValues(0) = 0.3127
                wpValues(1) = 0.329
            Case 6, 7
                'NTSC, SMPTE ST240
                idxOK = True
                rgbValues(0) = 0.63
                rgbValues(1) = 0.34
                
                rgbValues(3) = 0.31
                rgbValues(4) = 0.595
                
                rgbValues(6) = 0.155
                rgbValues(7) = 0.07
                
                wpValues(0) = 0.3127
                wpValues(1) = 0.329
            Case 8
                'Generic film
                idxOK = True
                rgbValues(0) = 0.681
                rgbValues(1) = 0.319
                
                rgbValues(3) = 0.243
                rgbValues(4) = 0.692
                
                rgbValues(6) = 0.145
                rgbValues(7) = 0.049
                
                wpValues(0) = 0.31
                wpValues(1) = 0.316
            Case 9
                'Modern HDTV (Rec BT 2020-2, 2100-2)
                idxOK = True
                rgbValues(0) = 0.708
                rgbValues(1) = 0.292
                
                rgbValues(3) = 0.17
                rgbValues(4) = 0.797
                
                rgbValues(6) = 0.131
                rgbValues(7) = 0.046
                
                wpValues(0) = 0.3127
                wpValues(1) = 0.329
            Case 10
                'CIE 1931 XYZ (should not be used in PNG due to strict RGB requirement)
                idxOK = True
                rgbValues(0) = 1#
                rgbValues(1) = 0#
                
                rgbValues(3) = 0#
                rgbValues(4) = 1#
                
                rgbValues(6) = 0#
                rgbValues(7) = 0#
                
                wpValues(0) = 0.333333333333333
                wpValues(1) = 0.333333333333333
            Case 11
                'SMPTE RP 431-2
                idxOK = True
                rgbValues(0) = 0.68
                rgbValues(1) = 0.32
                
                rgbValues(3) = 0.265
                rgbValues(4) = 0.69
                
                rgbValues(6) = 0.15
                rgbValues(7) = 0.06
                
                wpValues(0) = 0.314
                wpValues(1) = 0.351
            Case 12
                'SMPTE RP 432-1
                idxOK = True
                rgbValues(0) = 0.68
                rgbValues(1) = 0.32
                
                rgbValues(3) = 0.265
                rgbValues(4) = 0.69
                
                rgbValues(6) = 0.15
                rgbValues(7) = 0.06
                
                wpValues(0) = 0.3127
                wpValues(1) = 0.329
            '13-21 are all reserved
            Case 22
                'No corresponding spec
                idxOK = True
                rgbValues(0) = 0.63
                rgbValues(1) = 0.34
                
                rgbValues(3) = 0.295
                rgbValues(4) = 0.605
                
                rgbValues(6) = 0.155
                rgbValues(7) = 0.077
                
                wpValues(0) = 0.3127
                wpValues(1) = 0.329
            
        End Select
        
        'If the color primaries were bad, don't proceed with transfer functions
        If (Not idxOK) Then
            InternalError FUNC_NAME, "unknown color primaries index: " & cICPBytes(1)
            Exit Function
        End If
        
        'Transfer functions come next.
        idxOK = False
        
        '10 params max for an LCMS tone curve (not all params are needed by all functions, FYI,
        ' but there is no harm in null-entries).
        Dim listParams() As Double
        ReDim listParams(0 To 9) As Double
        
        'Technically transfer curves can vary per-channel, but in all cases we only intend to cover
        ' the "same transfer for all channels" case.
        Dim pTransferCurve As Long, pTransferCurve2 As Long, pTransferCurve3 As Long
        Select Case cICPBytes(1)
            
            'Transfer functions rely on built-in LCMS2 handling.  I currently only cover formats found "in the wild",
            ' as the potential list of transfer functions is huge and unwieldy, and will not be encountered in any of our lifetimes.
            Case 1, 6, 14, 15
                idxOK = True
                
                'Build a custom tone curve in LCMS format using hard-coded values from the spec
                ' (defined by https://www.itu.int/rec/T-REC-H.273-202407-I/en).
                listParams(0) = 1# / 0.45 'Power
                listParams(1) = 1# / 1.099 'alpha
                listParams(2) = 0.099 / 1.099
                listParams(3) = 1# / 4.5
                listParams(4) = 0.081   'beta
                
                pTransferCurve = LittleCMS.LCMS_GetAdvancedToneCurve(4, listParams)
                pTransferCurve2 = pTransferCurve
                pTransferCurve3 = pTransferCurve
                If PNG_DEBUG_VERBOSE Then PDDebug.LogAction "Using Rec.2020 via cICP for color management"
                
            'sRGB
            Case 13
                idxOK = True
                
                listParams(0) = 2.4
                listParams(1) = 1 / 1.055
                listParams(2) = 0.055 / 1.055
                listParams(3) = 1# / 12.92
                listParams(4) = 0.04045
                
                pTransferCurve = LittleCMS.LCMS_GetAdvancedToneCurve(4, listParams)
                pTransferCurve2 = pTransferCurve
                pTransferCurve3 = pTransferCurve
                If PNG_DEBUG_VERBOSE Then PDDebug.LogAction "Using sRGB via cICP for color management"
            
            Case Else
                'Any other indices are unexpected and *could* be covered in the futurs, but it's time-consuming
                ' to implement them because I have to track down third-party specs with magic numbers... so I'm
                ' leaving any further color spaces as "TODO if/when users complain or real-world examples appear!"
            
        End Select
        
        'If the color primaries were bad, don't proceed with transfer functions
        If (Not idxOK) Then
            InternalError FUNC_NAME, "unknown transfer function index: " & cICPBytes(1)
            Exit Function
        End If
        
        'Our next goal is to assemble an ICC profile that matches the chromaticity and/or gamma data.
        ' With this, we can use our normal color-management pipeline to color-correct this DIB.
        Set srcProfile = New pdLCMSProfile

        'Make sure the chromaticity data passes internal LCMS validation
        PrepcICPData = srcProfile.CreateCustomRGBProfile_Advanced(VarPtr(wpValues(0)), VarPtr(rgbValues(0)), pTransferCurve, pTransferCurve2, pTransferCurve3, True)
        srcImgProfileGood = PrepcICPData
        
    Else
        PrepcICPData = False
    End If
    
End Function

'1, 2, and 4-bit-per-pixel data must be upsampled to be easily handled in VB.  We do this prior to constructing a
' final image, as it allows us to use identical rendering code for 1/2/4/8-bpp indexed images.
' (NOTE: this function handles both interlaced and non-interlaced images just fine.)
Private Sub UpsampleLowBitDepthData(ByVal idatIndex As Long, ByVal pxWidth As Long, ByVal pxHeight As Long)
    
    'We only need to upsample bit-depths less than 8-bpp
    If (m_Header.BitsPerPixel > 8) Or ((m_Header.BitsPerPixel = 8) And (Not m_Header.Interlaced)) Then Exit Sub
    
    'For non-interlaced images, this sub is just a thin wrapper.  For interlaced images, however,
    ' we first need to separate the original data stream into its component images (one for each
    ' interlacing pass), up-sample each one separately, then reassemble the results into a normal,
    ' un-interlaced stream that subsequent assembly functions can more easily use.
    Dim xFinal As Long, yFinal As Long, xStride As Long
    
    'Anyway, the gist is that non-interlaced images are *way* easier to deal with.
    If (Not m_Header.Interlaced) Then
    
        xFinal = pxWidth - 1
        yFinal = pxHeight - 1
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
        SeparateInterlacedIDATs idatIndex, tmpChunks, (m_Header.BitDepth < 8)
        
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
        newIdatSize = (pxWidth + 1) * pxHeight
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
        xFinal = pxWidth - 1
        yFinal = pxHeight - 1
        newStride = pxWidth + 1
        
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
        m_XPixelCount(0) = pxWidth
        
        ReDim m_YPixelCount(0) As Long
        m_YPixelCount(0) = pxHeight
        
        ReDim m_XByteCount(0) As Long
        m_XByteCount(0) = pxWidth + 1
        
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
Private Sub SeparateInterlacedIDATs(ByVal idatIndex As Long, ByRef dstChunks() As pdPNGChunk, Optional ByVal alsoUpSampleBits As Boolean = True)
    
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
            dstChunks(i).CreateChunkForImport "IDAT", sizeOfPass
            
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
        
            Dim bitFlags(0 To 7) As Byte
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
            
            Dim shiftTable(0 To 3) As Byte
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

'Interlaced sub-images have varying sizes depending on the iteration pass.  This function returns the
' width and height of a given interlaced sub-image; m_Header *must* be populated for this to work, since it
' relies on color-mode and bit-depth data.
Private Function GetSizeInterlaced(ByVal srcWidth As Long, ByVal srcHeight As Long, ByRef dstX As Long, ByRef dstY As Long, ByVal intPass As Long) As Long

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
            dstX = Int((srcWidth + 7) \ 8)
            dstY = Int((srcHeight + 7) \ 8)
        
        'Second pass is a little weirder in the x-direction
        Case 2
            dstX = Int((srcWidth + 3) \ 8)
            dstY = Int((srcHeight + 7) \ 8)
        
        'Third pass introduces twice as many pixels in each scanline
        Case 3
            dstX = Int((srcWidth + 7) \ 8) + Int((srcWidth + 3) \ 8)
            dstY = Int((srcHeight + 3) \ 8)
        
        'Fourth pass introduces twice as many y-pixels in each 8x8 block
        Case 4
            dstX = Int((srcWidth + 5) \ 8) + Int((srcWidth + 1) \ 8)
            dstY = Int((srcHeight + 7) \ 8) + Int((srcHeight + 3) \ 8)
        
        'Fifth pass once again doubles the pixels in each scanline
        Case 5
            dstX = Int((srcWidth + 7) \ 8) + Int((srcWidth + 5) \ 8) + Int((srcWidth + 3) \ 8) + Int((srcWidth + 1) \ 8)
            dstY = Int((srcHeight + 5) \ 8) + Int((srcHeight + 1) \ 8)
        
        'Sixth pass now has 4 pixels in the x and y direction of each 8x8 block
        Case 6
            dstX = Int((srcWidth + 6) \ 8) + Int((srcWidth + 4) \ 8) + Int((srcWidth + 2) \ 8) + Int(srcWidth \ 8)
            dstY = Int((srcHeight + 7) \ 8) + Int((srcHeight + 5) \ 8) + Int((srcHeight + 3) \ 8) + Int((srcHeight + 1) \ 8)
        
        'Seventh (and final) pass has full pixels in the x-direction, and 4 pixels per 8x8 block in the y-direction
        Case 7
            dstX = srcWidth
            dstY = Int((srcHeight + 6) \ 8) + Int((srcHeight + 4) \ 8) + Int((srcHeight + 2) \ 8) + Int(srcHeight \ 8)
    
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

'Return the image's transparent color, if any.  Only relevant on grayscale and truecolor type PNGs.
' RETURNS: FALSE if no chunk exists; TRUE (with a filled dstColor value) otherwise.  For a grayscale image, the grayscale
' "shade" is automatically expanded to its RGB equivalent.
Friend Function GetTransparentColor(ByRef dstColor As Long, Optional ByVal scaleTo8Bits As Boolean = True) As Boolean
    
    GetTransparentColor = False
    
    'If a tRNS chunk exists, retrieve its information, too.  (Note that we only check for tRNS under certain
    ' color type combinations; in particular, it is forbidden if the image contains a full alpha channel.)
    Dim trnsRed As Long, trnsGreen As Long, trnsBlue As Long
    trnsRed = -1: trnsGreen = -1: trnsBlue = -1
    
    If (m_Header.ColorType = png_Greyscale) Or (m_Header.ColorType = png_Truecolor) Then
        
        Dim trnsIndex As Long
        trnsIndex = Me.GetIndexOfChunk("tRNS")
        If (trnsIndex >= 0) Then
            Dim tmpPal() As RGBQuad
            GetTransparentColor = m_Chunks(trnsIndex).GetTRNSData(tmpPal, trnsRed, trnsGreen, trnsBlue, m_Header, m_Warnings)
            If GetTransparentColor Then
                If (m_Header.BitDepth = 16) Then
                    If scaleTo8Bits Then
                        trnsRed = trnsRed \ 257
                        trnsGreen = trnsGreen \ 257
                        trnsBlue = trnsBlue \ 257
                    End If
                
                'We still need to check for valid 8-bit values; the 2025 PNG spec revision explicitly
                ' tests this case.
                Else
                    If (trnsRed > 255) Then trnsRed = trnsRed \ 257
                    If (trnsGreen > 255) Then trnsGreen = trnsGreen \ 257
                    If (trnsBlue > 255) Then trnsBlue = trnsBlue \ 257
                End If
                
                'Additional failsafes added 2025
                If (trnsRed <= 255) And (trnsGreen <= 255) And (trnsBlue <= 255) And (trnsRed >= 0) And (trnsGreen >= 0) And (trnsBlue >= 0) Then
                    dstColor = RGB(trnsRed, trnsGreen, trnsBlue)
                Else
                    dstColor = 0
                End If
            End If
        End If
        
    End If
    
End Function

'Request a copy of this PNG's header; returns a blank header if no PNG is loaded
Friend Sub GetHeader(ByRef dstHeader As PD_PNGHeader)
    dstHeader = m_Header
End Sub

Friend Function GetBitsPerPixel() As Long
    GetBitsPerPixel = m_Header.BitsPerPixel
End Function

Friend Function GetColorType() As PD_PNGColorType
    GetColorType = m_Header.ColorType
End Function

Friend Function HasAlpha() As Boolean
    HasAlpha = (m_Header.ColorType = png_GreyscaleAlpha) Or (m_Header.ColorType = png_TruecolorAlpha) Or (Me.GetIndexOfChunk("tRNS") >= 0)
End Function

Friend Sub SetIgnoreColorData(ByVal newSetting As Boolean)
    m_IgnoreEmbeddedColorData = newSetting
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
