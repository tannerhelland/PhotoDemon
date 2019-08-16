Attribute VB_Name = "Plugin_libdeflate"
'***************************************************************************
'File Compression Interface (via libdeflate)
'Copyright 2002-2016 by Tanner Helland
'Created: 3/02/02
'Last updated: 15/February/19
'Last update: wrap up initial build
'
'LibDeflate: https://github.com/ebiggers/libdeflate
' - "libdeflate is a library for fast, whole-buffer DEFLATE-based compression and decompression."
' - "libdeflate is heavily optimized. It is significantly faster than the zlib library, both for
'    compression and decompression, and especially on x86 processors."
'
'PhotoDemon uses libdeflate for reading/writing DEFLATE, zlib, and gzip data buffers.
'
'This wrapper class uses a shorthand implementation of DispCallFunc originally written by Olaf Schmidt.
' Many thanks to Olaf, whose original version can be found here (link good as of Feb 2019):
' http://www.vbforums.com/showthread.php?781595-VB6-Call-Functions-By-Pointer-(Universall-DLL-Calls)&p=4795471&viewfull=1#post4795471
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Enum LibDeflate_Result
    ld_Success = 0
    ld_BadData = 1
    ld_ShortOutput = 2
    ld_InsufficientSpace = 3
End Enum

#If False Then
    Private Const ld_Success = 0, ld_BadData = 1, ld_ShortOutput = 2, ld_InsufficientSpace = 3
#End If

'LibDeflate is zlib-compatible, but it exposes even higher compression levels (12 vs zlib's 9) for
' better-but-slower compression.  The default value remains 6; these are all declared in libdeflate.h
Private Const LIBDEFLATE_MIN_CLEVEL = 1
Private Const LIBDEFLATE_MAX_CLEVEL = 12
Private Const LIBDEFLATE_DEFAULT_CLEVEL = 6

'libdeflate has very specific compiler needs in order to produce maximum perf code, so rather than
' recompile myself, I've just grabbed the prebuilt Windows binaries and wrapped 'em using DispCallFunc
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef retVAR As Variant) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

'A single library handle is maintained for the life of a class instance; see Initialize and Release functions, below.
Private m_libDeflateHandle As Long

'To simplify interactions, we create temporary libdeflate compressor and decompressor instances
' "as we go".  Compressors in libdeflate are unique in-that they are specific to a given compression level
' (e.g. "compress level 1" compressor != "compress level 6" compressor) which makes interactions
' a little wonky as far as this class is concerned; for best results, you should only
Private m_hCompressor As Long, m_hDecompressor As Long

'At load-time, we cache a number of proc addresses (required for passing through DispCallFunc).
' This saves us a little time vs calling GetProcAddress on each call.
Private Enum LD_ProcAddress
    libdeflate_alloc_compressor
    libdeflate_deflate_compress
    libdeflate_deflate_compress_bound
    libdeflate_zlib_compress
    libdeflate_zlib_compress_bound
    libdeflate_gzip_compress
    libdeflate_gzip_compress_bound
    libdeflate_free_compressor
    libdeflate_alloc_decompressor
    libdeflate_deflate_decompress
    libdeflate_deflate_decompress_ex
    libdeflate_zlib_decompress
    libdeflate_gzip_decompress
    libdeflate_gzip_decompress_ex
    libdeflate_free_decompressor
    libdeflate_adler32
    libdeflate_crc32
    [last_address]
End Enum

Private m_ProcAddresses() As Long

'Rather than allocate new memory on each DispCallFunc invoke, just reuse a set of temp arrays declared
' to the maximum relevant size (see InitializeEngine, below).
Private Const MAX_PARAM_COUNT As Long = 10
Private m_vType() As Integer, m_vPtr() As Long

'Basic init/release functions
Public Function InitializeEngine(ByRef pathToDLLFolder As String) As Boolean

    Dim libDeflatePath As String
    libDeflatePath = pathToDLLFolder & "libdeflate.dll"
    m_libDeflateHandle = VBHacks.LoadLib(libDeflatePath)
    InitializeEngine = (m_libDeflateHandle <> 0)
    
    If InitializeEngine Then
    
        'Pre-load all relevant proc addresses
        ReDim m_ProcAddresses(0 To [last_address] - 1) As Long
        m_ProcAddresses(libdeflate_alloc_compressor) = GetProcAddress(m_libDeflateHandle, "libdeflate_alloc_compressor")
        m_ProcAddresses(libdeflate_deflate_compress) = GetProcAddress(m_libDeflateHandle, "libdeflate_deflate_compress")
        m_ProcAddresses(libdeflate_deflate_compress_bound) = GetProcAddress(m_libDeflateHandle, "libdeflate_deflate_compress_bound")
        m_ProcAddresses(libdeflate_zlib_compress) = GetProcAddress(m_libDeflateHandle, "libdeflate_zlib_compress")
        m_ProcAddresses(libdeflate_zlib_compress_bound) = GetProcAddress(m_libDeflateHandle, "libdeflate_zlib_compress_bound")
        m_ProcAddresses(libdeflate_gzip_compress) = GetProcAddress(m_libDeflateHandle, "libdeflate_gzip_compress")
        m_ProcAddresses(libdeflate_gzip_compress_bound) = GetProcAddress(m_libDeflateHandle, "libdeflate_gzip_compress_bound")
        m_ProcAddresses(libdeflate_free_compressor) = GetProcAddress(m_libDeflateHandle, "libdeflate_free_compressor")
        m_ProcAddresses(libdeflate_alloc_decompressor) = GetProcAddress(m_libDeflateHandle, "libdeflate_alloc_decompressor")
        m_ProcAddresses(libdeflate_deflate_decompress) = GetProcAddress(m_libDeflateHandle, "libdeflate_deflate_decompress")
        m_ProcAddresses(libdeflate_deflate_decompress_ex) = GetProcAddress(m_libDeflateHandle, "libdeflate_deflate_decompress_ex")
        m_ProcAddresses(libdeflate_zlib_decompress) = GetProcAddress(m_libDeflateHandle, "libdeflate_zlib_decompress")
        m_ProcAddresses(libdeflate_gzip_decompress) = GetProcAddress(m_libDeflateHandle, "libdeflate_gzip_decompress")
        m_ProcAddresses(libdeflate_gzip_decompress_ex) = GetProcAddress(m_libDeflateHandle, "libdeflate_gzip_decompress_ex")
        m_ProcAddresses(libdeflate_free_decompressor) = GetProcAddress(m_libDeflateHandle, "libdeflate_free_decompressor")
        m_ProcAddresses(libdeflate_adler32) = GetProcAddress(m_libDeflateHandle, "libdeflate_adler32")
        m_ProcAddresses(libdeflate_crc32) = GetProcAddress(m_libDeflateHandle, "libdeflate_crc32")
            
    Else
        PDDebug.LogAction "WARNING!  LoadLibraryW failed to load libdeflate.  Last DLL error: " & Err.LastDllError
    End If
    
    'Initialize all module-level arrays
    ReDim m_vType(0 To MAX_PARAM_COUNT - 1) As Integer
    ReDim m_vPtr(0 To MAX_PARAM_COUNT - 1) As Long
    
End Function

Public Sub ReleaseEngine()
    If (m_libDeflateHandle <> 0) Then
        VBHacks.FreeLib m_libDeflateHandle
        m_libDeflateHandle = 0
    End If
End Sub

'Actual compression/decompression functions.  Only arrays and pointers are standardized.  It's assumed
' that users can write simple wrappers for other data types, as necessary.
Public Function CompressPtrToDstArray(ByRef dstArray() As Byte, ByRef dstCompressedSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long, Optional ByVal compressionLevel As Long = -1, Optional ByVal dstArrayIsAlreadySized As Boolean = False, Optional ByVal trimCompressedArray As Boolean = False, Optional ByVal cmpFormat As PD_CompressionFormat = cf_Zlib) As Boolean

    ValidateCompressionLevel compressionLevel
    
    'Prep the destination array, as necessary
    If (Not dstArrayIsAlreadySized) Then
        dstCompressedSizeInBytes = GetWorstCaseSize(constSrcSizeInBytes)
        ReDim dstArray(0 To dstCompressedSizeInBytes - 1) As Byte
    End If
    
    'Compress the data
    CompressPtrToDstArray = LibDeflateCompress(VarPtr(dstArray(0)), dstCompressedSizeInBytes, constSrcPtr, constSrcSizeInBytes, compressionLevel, cmpFormat)
        
    'If compression was successful, trim the destination array, as requested
    If trimCompressedArray And CompressPtrToDstArray Then
        If (UBound(dstArray) <> dstCompressedSizeInBytes - 1) Then ReDim Preserve dstArray(0 To dstCompressedSizeInBytes - 1) As Byte
    End If
    
End Function

Public Function CompressPtrToPtr(ByVal constDstPtr As Long, ByRef dstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long, Optional ByVal compressionLevel As Long = -1, Optional ByVal cmpFormat As PD_CompressionFormat = cf_Zlib) As Boolean
    CompressPtrToPtr = LibDeflateCompress(constDstPtr, dstSizeInBytes, constSrcPtr, constSrcSizeInBytes, compressionLevel, cmpFormat)
End Function

Public Function DecompressPtrToDstArray(ByRef dstArray() As Byte, ByVal constDstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long, Optional ByVal dstArrayIsAlreadySized As Boolean = False, Optional ByVal cmpFormat As PD_CompressionFormat = cf_Zlib, Optional ByVal allowFallbacks As Boolean = True) As Boolean
    If (Not dstArrayIsAlreadySized) Then ReDim dstArray(0 To constDstSizeInBytes - 1) As Byte
    DecompressPtrToDstArray = LibDeflateDecompress(VarPtr(dstArray(0)), constDstSizeInBytes, constSrcPtr, constSrcSizeInBytes, cmpFormat, allowFallbacks)
End Function

Public Function DecompressPtrToPtr(ByVal constDstPtr As Long, ByVal constDstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long, Optional ByVal cmpFormat As PD_CompressionFormat = cf_Zlib, Optional ByVal allowFallbacks As Boolean = True) As Boolean
    DecompressPtrToPtr = LibDeflateDecompress(constDstPtr, constDstSizeInBytes, constSrcPtr, constSrcSizeInBytes, cmpFormat, allowFallbacks)
End Function

Public Function GetCrc32(ByVal srcPtr As Long, ByVal srcLen As Long, Optional ByVal startValue As Long = 0&, Optional ByVal calcStartForMe As Boolean = True) As Long
    
    'Get an initial default "seed"
    If calcStartForMe Then startValue = CallCDeclW(libdeflate_crc32, vbLong, 0&, 0&, 0&)
    
    'Use the seed to calculate an actual Crc32
    If (srcPtr <> 0) Then GetCrc32 = CallCDeclW(libdeflate_crc32, vbLong, startValue, srcPtr, srcLen) Else GetCrc32 = startValue
    
End Function

Public Function GetAdler32(ByVal srcPtr As Long, ByVal srcLen As Long, Optional ByVal startValue As Long = 0&) As Long
    
    'Get an initial default "seed"
    If (startValue = 0) Then startValue = CallCDeclW(libdeflate_adler32, vbLong, 0&, 0&, 0&)
    
    'Use the seed to calculate an actual Crc32
    GetAdler32 = CallCDeclW(libdeflate_adler32, vbLong, startValue, srcPtr, srcLen)
    
End Function

'Specialized wrappers follow.
'
'Return the precise return code from a zlib decompress operation; this is helpful for obnoxious
' edge-cases like PNG files where compressed chunks don't store the original data size, so we
' have to manually attempt ever-larger buffers.
Public Function Decompress_ZLib(ByVal constDstPtr As Long, ByVal constDstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long) As Long
    
    'Allocate a decompressor
    ' LIBDEFLATEAPI struct libdeflate_decompressor * libdeflate_alloc_decompressor(void)
    Dim hDecompress As Long
    hDecompress = CallCDeclW(libdeflate_alloc_decompressor, vbLong)
    If (hDecompress <> 0) Then
        
        'Perform decompression
        ' LIBDEFLATEAPI enum libdeflate_result libdeflate_zlib_decompress(struct libdeflate_decompressor *decompressor, const void *in, size_t in_nbytes, void *out, size_t out_nbytes_avail, size_t *actual_out_nbytes_ret)
        Decompress_ZLib = CallCDeclW(libdeflate_zlib_decompress, vbLong, hDecompress, constSrcPtr, constSrcSizeInBytes, constDstPtr, constDstSizeInBytes, 0&)
        
        ' Make sure we free the compressor before exiting
        'LIBDEFLATEAPI void libdeflate_free_decompressor(struct libdeflate_decompressor *decompressor);
        CallCDeclW libdeflate_free_decompressor, vbEmpty, hDecompress
        
    Else
        InternalError "Decompress_ZLib", "Failed to initialize a decompressor"
    End If
    
End Function

'Compression helper functions.  Worst-case size is generally required for sizing a destination array prior to compression,
' and the exact calculation method varies by compressor.
Private Function LibDeflateCompress(ByVal constDstPtr As Long, ByRef dstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long, Optional ByVal compressionLevel As Long = -1, Optional ByVal cmpFormat As PD_CompressionFormat = cf_Zlib) As Boolean
    
    ValidateCompressionLevel compressionLevel
    
    'Allocate a compressor
    ' LIBDEFLATEAPI struct libdeflate_compressor * libdeflate_alloc_compressor(int compression_level)
    Dim hCompress As Long
    hCompress = CallCDeclW(libdeflate_alloc_compressor, vbLong, compressionLevel)
    
    If (hCompress <> 0) Then
        
        'Perform compression
        ' LIBDEFLATEAPI size_t libdeflate_zlib_compress(struct libdeflate_compressor *compressor, const void *in, size_t in_nbytes, void *out, size_t out_nbytes_avail)
        ' LIBDEFLATEAPI size_t libdeflate_deflate_compress(struct libdeflate_compressor *compressor, const void *in, size_t in_nbytes, void *out, size_t out_nbytes_avail)
        ' LIBDEFLATEAPI size_t libdeflate_gzip_compress(struct libdeflate_compressor *compressor, const void *in, size_t in_nbytes, void *out, size_t out_nbytes_avail)
        Dim lReturn As Long
        If (cmpFormat = cf_Zlib) Then
            lReturn = CallCDeclW(libdeflate_zlib_compress, vbLong, hCompress, constSrcPtr, constSrcSizeInBytes, constDstPtr, dstSizeInBytes)
        ElseIf (cmpFormat = cf_Deflate) Then
            lReturn = CallCDeclW(libdeflate_deflate_compress, vbLong, hCompress, constSrcPtr, constSrcSizeInBytes, constDstPtr, dstSizeInBytes)
        ElseIf (cmpFormat = cf_Gzip) Then
            lReturn = CallCDeclW(libdeflate_gzip_compress, vbLong, hCompress, constSrcPtr, constSrcSizeInBytes, constDstPtr, dstSizeInBytes)
        End If
        
        LibDeflateCompress = (lReturn <> 0)
        If LibDeflateCompress Then
            dstSizeInBytes = lReturn
        Else
            InternalError "LibDeflateCompress", "operation failed"
        End If
        
        'Free the compressor before exiting
        ' LIBDEFLATEAPI void libdeflate_free_compressor(struct libdeflate_compressor *compressor)
        CallCDeclW libdeflate_free_compressor, vbEmpty, hCompress
        
    Else
        InternalError "LibDeflateCompress", "failed to initialize a compressor"
    End If
    
    
End Function

Private Function LibDeflateDecompress(ByVal constDstPtr As Long, ByVal constDstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long, Optional ByVal cmpFormat As PD_CompressionFormat = cf_Zlib, Optional ByVal allowFallbacks As Boolean = True) As Boolean
    
    'Allocate a decompressor
    ' LIBDEFLATEAPI struct libdeflate_decompressor * libdeflate_alloc_decompressor(void)
    Dim hDecompress As Long
    hDecompress = CallCDeclW(libdeflate_alloc_decompressor, vbLong)
    If (hDecompress <> 0) Then
        
        'Perform decompression
        ' LIBDEFLATEAPI enum libdeflate_result libdeflate_zlib_decompress(struct libdeflate_decompressor *decompressor, const void *in, size_t in_nbytes, void *out, size_t out_nbytes_avail, size_t *actual_out_nbytes_ret)
        Dim lReturn As Long
        If (cmpFormat = cf_Zlib) Then
            lReturn = CallCDeclW(libdeflate_zlib_decompress, vbLong, hDecompress, constSrcPtr, constSrcSizeInBytes, constDstPtr, constDstSizeInBytes, 0&)
        ElseIf (cmpFormat = cf_Deflate) Then
            lReturn = CallCDeclW(libdeflate_deflate_decompress, vbLong, hDecompress, constSrcPtr, constSrcSizeInBytes, constDstPtr, constDstSizeInBytes, 0&)
        ElseIf (cmpFormat = cf_Gzip) Then
            lReturn = CallCDeclW(libdeflate_gzip_decompress, vbLong, hDecompress, constSrcPtr, constSrcSizeInBytes, constDstPtr, constDstSizeInBytes, 0&)
        End If
        
        'If decompression failed, try it again, but with the explicit instruction to attempt to decompress enough
        ' data to fill the output buffer.  If there is still more compressed data past that point, it will be
        ' lost/ignored, but this may produce enough data for the caller to proceed normally.
        If (lReturn <> 0) And allowFallbacks Then
            InternalError "LibDeflateDecompress", "full decompress failed; attempting partial decompress instead..."
            Dim bytesWritten As Long
            If (cmpFormat = cf_Zlib) Then
                lReturn = CallCDeclW(libdeflate_zlib_decompress, vbLong, hDecompress, constSrcPtr, constSrcSizeInBytes, constDstPtr, constDstSizeInBytes, VarPtr(bytesWritten))
            ElseIf (cmpFormat = cf_Deflate) Then
                lReturn = CallCDeclW(libdeflate_deflate_decompress, vbLong, hDecompress, constSrcPtr, constSrcSizeInBytes, constDstPtr, constDstSizeInBytes, VarPtr(bytesWritten))
            ElseIf (cmpFormat = cf_Gzip) Then
                lReturn = CallCDeclW(libdeflate_gzip_decompress, vbLong, hDecompress, constSrcPtr, constSrcSizeInBytes, constDstPtr, constDstSizeInBytes, VarPtr(bytesWritten))
            End If
            If (constDstSizeInBytes = bytesWritten) Then
                lReturn = 0
                InternalError "LibDeflateDecompress", "Partial decompression successful."
            Else
                InternalError "LibDeflateDecompress", "Partial decompression failed; source data is likely corrupt. (" & lReturn & ", " & bytesWritten & ")"
            End If
        End If
        
        'If compression still failed, we're outta luck.
        LibDeflateDecompress = (lReturn = 0)
        If (Not LibDeflateDecompress) Then
            InternalError "LibDeflateDecompress", "operation failed; return was " & lReturn & " " & GetDecompressErrorText(lReturn)
            InternalError "LibDeflateDecompress", "FYI inputs were: " & constDstPtr & ", " & constDstSizeInBytes & ", " & constSrcPtr & ", " & constSrcSizeInBytes
        End If
        
        ' Make sure we free the compressor before exiting
        'LIBDEFLATEAPI void libdeflate_free_decompressor(struct libdeflate_decompressor *decompressor);
        CallCDeclW libdeflate_free_decompressor, vbEmpty, hDecompress
        
    Else
        InternalError "LibDeflateDecompress", "WARNING!  Failed to initialize a libdeflate decompressor."
    End If
    
End Function

'Magic numbers are taken from libdeflate.h
Private Function GetDecompressErrorText(ByVal ErrNum As Long) As String
    If (ErrNum = 1) Then
        GetDecompressErrorText = "bad input data"
    ElseIf (ErrNum = 2) Then
        GetDecompressErrorText = "short output"
    ElseIf (ErrNum = 3) Then
        GetDecompressErrorText = "insufficient destination buffer"
    End If
End Function

'Note that libdeflate exports its own "get worst-case dst size" function.  However, it requires you to
' pass a compressor handle that has been initialized to the target compression level... which creates
' problems for the way our ICompress interface works.  Because there's no good way to mimic this,
' we simply use the standard zlib "worst case" calculation, but add extra bytes for the gzip case
' (as gzip headers/trailers are larger than zlib ones).
Public Function GetWorstCaseSize(ByVal srcBufferSizeInBytes As Long, Optional ByVal compressionLevel As Long = -1, Optional ByVal cmpFormat As PD_CompressionFormat = cf_Zlib) As Long
    
    If (compressionLevel = -1) Then compressionLevel = Plugin_libdeflate.GetDefaultCompressionLevel()
    
    'libdeflate requires a compressor object in order to calculate a "worst-case" size
    Dim hCompress As Long
    hCompress = CallCDeclW(libdeflate_alloc_compressor, vbLong, compressionLevel)
    If (hCompress <> 0) Then
    
        If (cmpFormat = cf_Zlib) Then
            ' LIBDEFLATEAPI size_t libdeflate_deflate_compress_bound(struct libdeflate_compressor *compressor, size_t in_nbytes)
            GetWorstCaseSize = CallCDeclW(libdeflate_deflate_compress_bound, vbLong, hCompress, srcBufferSizeInBytes)
        ElseIf (cmpFormat = cf_Deflate) Then
            ' LIBDEFLATEAPI size_t libdeflate_zlib_compress_bound(struct libdeflate_compressor *compressor, size_t in_nbytes)
            GetWorstCaseSize = CallCDeclW(libdeflate_zlib_compress_bound, vbLong, hCompress, srcBufferSizeInBytes)
        ElseIf (cmpFormat = cf_Gzip) Then
            ' LIBDEFLATEAPI size_t libdeflate_gzip_compress_bound(struct libdeflate_compressor *compressor, size_t in_nbytes)
            GetWorstCaseSize = CallCDeclW(libdeflate_gzip_compress_bound, vbLong, hCompress, srcBufferSizeInBytes)
        End If
        
        'Free the compressor before exiting
        CallCDeclW libdeflate_free_compressor, vbEmpty, hCompress
        
    Else
        InternalError "GetWorstCaseSize", "failed to allocate a compressor"
    End If
    
End Function

Public Function GetDefaultCompressionLevel() As Long
    GetDefaultCompressionLevel = LIBDEFLATE_DEFAULT_CLEVEL
End Function

Public Function GetMinCompressionLevel() As Long
    GetMinCompressionLevel = LIBDEFLATE_MIN_CLEVEL
End Function

Public Function GetMaxCompressionLevel() As Long
    GetMaxCompressionLevel = LIBDEFLATE_MAX_CLEVEL
End Function

'Misc helper functions.  Name can be useful for user-facing reporting.
Public Function GetCompressorName() As String
    GetCompressorName = "libdeflate"
End Function

Public Function IsCompressorReady() As Boolean
    IsCompressorReady = (m_libDeflateHandle <> 0)
End Function

'libdeflate doesn't export a version function, but this class was designed against the v1.2 release.
Public Function GetCompressorVersion() As String
    GetCompressorVersion = "1.2"
End Function

'Private methods follow

'Clamp requested compression levels to valid inputs, and resolve negative numbers to the engine's default value.
Private Sub ValidateCompressionLevel(ByRef inputLevel As Long)
    If (inputLevel = -1) Then
        inputLevel = LIBDEFLATE_DEFAULT_CLEVEL
    ElseIf (inputLevel < LIBDEFLATE_MIN_CLEVEL) Then
        inputLevel = LIBDEFLATE_MIN_CLEVEL
    ElseIf (inputLevel > LIBDEFLATE_MAX_CLEVEL) Then
        inputLevel = LIBDEFLATE_MAX_CLEVEL
    End If
End Sub

'DispCallFunc wrapper originally by Olaf Schmidt, with a few minor modifications; see the top of this class
' for a link to his original, unmodified version
Private Function CallCDeclW(ByVal lProc As LD_ProcAddress, ByVal fRetType As VbVarType, ParamArray pa() As Variant) As Variant

    Dim i As Long, pFunc As Long, vTemp() As Variant, hResult As Long
    
    Dim numParams As Long
    If (UBound(pa) < LBound(pa)) Then numParams = 0 Else numParams = UBound(pa) + 1
    
    If IsMissing(pa) Then
        ReDim vTemp(0) As Variant
    Else
        vTemp = pa 'make a copy of the params, to prevent problems with VT_Byref-Members in the ParamArray
    End If
    
    For i = 0 To numParams - 1
        If VarType(pa(i)) = vbString Then vTemp(i) = StrPtr(pa(i))
        m_vType(i) = VarType(vTemp(i))
        m_vPtr(i) = VarPtr(vTemp(i))
    Next i
    
    Const CC_CDECL As Long = 1
    hResult = DispCallFunc(0, m_ProcAddresses(lProc), CC_CDECL, fRetType, i, m_vType(0), m_vPtr(0), CallCDeclW)
    If hResult Then Err.Raise hResult
    
End Function

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String, Optional ByVal writeDebugLog As Boolean = True)
    If UserPrefs.GenerateDebugLogs Then
        If writeDebugLog Then PDDebug.LogAction "Plugin_libdeflate." & funcName & "() reported an error: " & errDescription
    Else
        Debug.Print "Plugin_libdeflate." & funcName & "() reported an error: " & errDescription
    End If
End Sub
