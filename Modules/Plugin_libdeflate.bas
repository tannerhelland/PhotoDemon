Attribute VB_Name = "Plugin_libdeflate"
'***************************************************************************
'File Compression Interface (via libdeflate)
'Copyright 2002-2016 by Tanner Helland
'Created: 3/02/02
'Last updated: 22/May/20
'Last update: switch to new (official) stdcall builds; see https://github.com/ebiggers/libdeflate/blob/master/NEWS
'
'LibDeflate: https://github.com/ebiggers/libdeflate
' - "libdeflate is a library for fast, whole-buffer DEFLATE-based compression and decompression."
' - "libdeflate is heavily optimized. It is significantly faster than the zlib library, both for
'    compression and decompression, and especially on x86 processors."
'
'PhotoDemon uses libdeflate for reading/writing DEFLATE, zlib, and gzip data buffers.
'
'As of v1.4, libdeflate authors now provide "official" 32-bit stdcall builds of the library.
' This greatly simplifies our interactions with it (vs the old cdecl workarounds).
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
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
Private Const LIBDEFLATE_MIN_CLEVEL As Long = 0
Private Const LIBDEFLATE_MAX_CLEVEL As Long = 12
Private Const LIBDEFLATE_DEFAULT_CLEVEL As Long = 6

'A single library handle is maintained for the life of a class instance; see Initialize and Release functions, below.
Private m_libDeflateHandle As Long

Private Declare Function libdeflate_alloc_compressor Lib "libdeflate" (ByVal compression_level As Long) As Long
Private Declare Function libdeflate_deflate_compress Lib "libdeflate" (ByVal libdeflate_compressor As Long, ByVal ptr_in As Long, ByVal in_nbytes As Long, ByVal ptr_out As Long, ByVal out_nbytes_avail As Long) As Long
Private Declare Function libdeflate_deflate_compress_bound Lib "libdeflate" (ByVal libdeflate_compressor As Long, ByVal in_nbytes As Long) As Long
Private Declare Function libdeflate_zlib_compress Lib "libdeflate" (ByVal libdeflate_compressor As Long, ByVal ptr_in As Long, ByVal in_nbytes As Long, ByVal ptr_out As Long, ByVal out_nbytes_avail As Long) As Long
Private Declare Function libdeflate_zlib_compress_bound Lib "libdeflate" (ByVal libdeflate_compressor As Long, ByVal in_nbytes As Long) As Long
Private Declare Function libdeflate_gzip_compress Lib "libdeflate" (ByVal libdeflate_compressor As Long, ByVal ptr_in As Long, ByVal in_nbytes As Long, ByVal ptr_out As Long, ByVal out_nbytes_avail As Long) As Long
Private Declare Function libdeflate_gzip_compress_bound Lib "libdeflate" (ByVal libdeflate_compressor As Long, ByVal in_nbytes As Long) As Long
Private Declare Sub libdeflate_free_compressor Lib "libdeflate" (ByVal libdeflate_compressor As Long)
Private Declare Function libdeflate_alloc_decompressor Lib "libdeflate" () As Long
Private Declare Function libdeflate_deflate_decompress Lib "libdeflate" (ByVal libdeflate_decompressor As Long, ByVal ptr_in As Long, ByVal in_nbytes As Long, ByVal ptr_out As Long, ByVal out_nbytes_avail As Long, ByRef actual_out_nbytes_ret As Long) As LibDeflate_Result
Private Declare Function libdeflate_zlib_decompress Lib "libdeflate" (ByVal libdeflate_decompressor As Long, ByVal ptr_in As Long, ByVal in_nbytes As Long, ByVal ptr_out As Long, ByVal out_nbytes_avail As Long, ByRef actual_out_nbytes_ret As Long) As LibDeflate_Result
Private Declare Function libdeflate_gzip_decompress Lib "libdeflate" (ByVal libdeflate_decompressor As Long, ByVal ptr_in As Long, ByVal in_nbytes As Long, ByVal ptr_out As Long, ByVal out_nbytes_avail As Long, ByRef actual_out_nbytes_ret As Long) As LibDeflate_Result
Private Declare Sub libdeflate_free_decompressor Lib "libdeflate" (ByVal libdeflate_decompressor As Long)
Private Declare Function libdeflate_adler32 Lib "libdeflate" (ByVal adler32 As Long, ByVal ptr_buffer As Long, ByVal len_in_bytes As Long) As Long
Private Declare Function libdeflate_crc32 Lib "libdeflate" (ByVal crc As Long, ByVal ptr_buffer As Long, ByVal len_in_bytes As Long) As Long

'libdeflate provides _ex versions of all decompress functions, where it simply attempts to decompress
' until it can't anymore.  PD never uses these are there are strong security implications, and we only
' decompress in contexts where the decompressed size is known in advance.
'Private Declare Function libdeflate_deflate_decompress_ex Lib "libdeflate" (ByVal libdeflate_decompressor As Long, ByVal ptr_in As Long, ByVal in_nbytes As Long, ByVal ptr_out As Long, ByVal out_nbytes_avail As Long, ByRef actual_in_nbytes_ret As Long, ByRef actual_out_nbytes_ret As Long) As LibDeflate_Result
'Private Declare Function libdeflate_zlib_decompress_ex Lib "libdeflate" (ByVal libdeflate_decompressor As Long, ByVal ptr_in As Long, ByVal in_nbytes As Long, ByVal ptr_out As Long, ByVal out_nbytes_avail As Long, ByRef actual_in_nbytes_ret As Long, ByRef actual_out_nbytes_ret As Long) As LibDeflate_Result
'Private Declare Function libdeflate_gzip_decompress_ex Lib "libdeflate" (ByVal libdeflate_decompressor As Long, ByVal ptr_in As Long, ByVal in_nbytes As Long, ByVal ptr_out As Long, ByVal out_nbytes_avail As Long, ByRef actual_in_nbytes_ret As Long, ByRef actual_out_nbytes_ret As Long) As LibDeflate_Result

'Basic init/release functions
Public Function InitializeEngine(ByRef pathToDLLFolder As String) As Boolean

    Dim libDeflatePath As String
    libDeflatePath = pathToDLLFolder & "libdeflate.dll"
    m_libDeflateHandle = VBHacks.LoadLib(libDeflatePath)
    InitializeEngine = (m_libDeflateHandle <> 0)
    
    If (Not InitializeEngine) Then
        PDDebug.LogAction "WARNING!  LoadLibraryW failed to load libdeflate.  Last DLL error: " & Err.LastDllError
    End If
    
End Function

Public Sub ReleaseEngine()
    If (m_libDeflateHandle <> 0) Then
        VBHacks.FreeLib m_libDeflateHandle
        m_libDeflateHandle = 0
    End If
End Sub

Public Function CompressPtrToPtr(ByVal constDstPtr As Long, ByRef dstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long, Optional ByVal compressionLevel As Long = -1, Optional ByVal cmpFormat As PD_CompressionFormat = cf_Zlib) As Boolean
    CompressPtrToPtr = LibDeflateCompress(constDstPtr, dstSizeInBytes, constSrcPtr, constSrcSizeInBytes, compressionLevel, cmpFormat)
End Function

Public Function DecompressPtrToPtr(ByVal constDstPtr As Long, ByVal constDstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long, Optional ByVal cmpFormat As PD_CompressionFormat = cf_Zlib, Optional ByVal allowFallbacks As Boolean = True) As Boolean
    DecompressPtrToPtr = LibDeflateDecompress(constDstPtr, constDstSizeInBytes, constSrcPtr, constSrcSizeInBytes, cmpFormat, allowFallbacks)
End Function

Public Function GetCrc32(ByVal srcPtr As Long, ByVal srcLen As Long, Optional ByVal startValue As Long = 0&, Optional ByVal calcStartForMe As Boolean = True) As Long
    
    'Get an initial default "seed"
    If calcStartForMe Then startValue = libdeflate_crc32(0&, 0&, 0&)
    
    'Use the seed to calculate an actual Crc32
    If (srcPtr <> 0) Then GetCrc32 = libdeflate_crc32(startValue, srcPtr, srcLen) Else GetCrc32 = startValue
    
End Function

Public Function GetAdler32(ByVal srcPtr As Long, ByVal srcLen As Long, Optional ByVal startValue As Long = 0&) As Long
    
    'Get an initial default "seed"
    If (startValue = 0) Then startValue = libdeflate_adler32(0&, 0&, 0&)
    
    'Use the seed to calculate an actual Crc32
    GetAdler32 = libdeflate_adler32(startValue, srcPtr, srcLen)
    
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
    hDecompress = libdeflate_alloc_decompressor()
    If (hDecompress <> 0) Then
        
        'Perform decompression
        ' LIBDEFLATEAPI enum libdeflate_result libdeflate_zlib_decompress(struct libdeflate_decompressor *decompressor, const void *in, size_t in_nbytes, void *out, size_t out_nbytes_avail, size_t *actual_out_nbytes_ret)
        ' "If the actual uncompressed size is known, then pass the actual
        '  uncompressed size as 'out_nbytes_avail' and pass NULL for
        '  actual_out_nbytes_ret'.  This makes libdeflate_deflate_decompress() fail
        '  with LIBDEFLATE_SHORT_OUTPUT if the data decompressed to fewer than the
        '  specified number of bytes."
        Decompress_ZLib = libdeflate_zlib_decompress(hDecompress, constSrcPtr, constSrcSizeInBytes, constDstPtr, constDstSizeInBytes, 0&)
        
        ' Make sure we free the compressor before exiting
        'LIBDEFLATEAPI void libdeflate_free_decompressor(struct libdeflate_decompressor *decompressor);
        libdeflate_free_decompressor hDecompress
        
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
    hCompress = libdeflate_alloc_compressor(compressionLevel)
    
    If (hCompress <> 0) Then
        
        'Perform compression
        ' LIBDEFLATEAPI size_t libdeflate_zlib_compress(struct libdeflate_compressor *compressor, const void *in, size_t in_nbytes, void *out, size_t out_nbytes_avail)
        ' LIBDEFLATEAPI size_t libdeflate_deflate_compress(struct libdeflate_compressor *compressor, const void *in, size_t in_nbytes, void *out, size_t out_nbytes_avail)
        ' LIBDEFLATEAPI size_t libdeflate_gzip_compress(struct libdeflate_compressor *compressor, const void *in, size_t in_nbytes, void *out, size_t out_nbytes_avail)
        Dim lReturn As Long
        If (cmpFormat = cf_Zlib) Then
            lReturn = libdeflate_zlib_compress(hCompress, constSrcPtr, constSrcSizeInBytes, constDstPtr, dstSizeInBytes)
        ElseIf (cmpFormat = cf_Deflate) Then
            lReturn = libdeflate_deflate_compress(hCompress, constSrcPtr, constSrcSizeInBytes, constDstPtr, dstSizeInBytes)
        ElseIf (cmpFormat = cf_Gzip) Then
            lReturn = libdeflate_gzip_compress(hCompress, constSrcPtr, constSrcSizeInBytes, constDstPtr, dstSizeInBytes)
        End If
        
        LibDeflateCompress = (lReturn <> 0)
        If LibDeflateCompress Then
            dstSizeInBytes = lReturn
        Else
            InternalError "LibDeflateCompress", "operation failed"
        End If
        
        'Free the compressor before exiting
        ' LIBDEFLATEAPI void libdeflate_free_compressor(struct libdeflate_compressor *compressor)
        libdeflate_free_compressor hCompress
        
    Else
        InternalError "LibDeflateCompress", "failed to initialize a compressor"
    End If
    
    
End Function

Private Function LibDeflateDecompress(ByVal constDstPtr As Long, ByVal constDstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long, Optional ByVal cmpFormat As PD_CompressionFormat = cf_Zlib, Optional ByVal allowFallbacks As Boolean = True) As Boolean
    
    'Allocate a decompressor
    ' LIBDEFLATEAPI struct libdeflate_decompressor * libdeflate_alloc_decompressor(void)
    Dim hDecompress As Long
    hDecompress = libdeflate_alloc_decompressor()
    If (hDecompress <> 0) Then
        
        'Perform decompression
        ' LIBDEFLATEAPI enum libdeflate_result libdeflate_zlib_decompress(struct libdeflate_decompressor *decompressor, const void *in, size_t in_nbytes, void *out, size_t out_nbytes_avail, size_t *actual_out_nbytes_ret)
        Dim lReturn As Long
        If (cmpFormat = cf_Zlib) Then
            lReturn = libdeflate_zlib_decompress(hDecompress, constSrcPtr, constSrcSizeInBytes, constDstPtr, constDstSizeInBytes, 0&)
        ElseIf (cmpFormat = cf_Deflate) Then
            lReturn = libdeflate_deflate_decompress(hDecompress, constSrcPtr, constSrcSizeInBytes, constDstPtr, constDstSizeInBytes, 0&)
        ElseIf (cmpFormat = cf_Gzip) Then
            lReturn = libdeflate_gzip_decompress(hDecompress, constSrcPtr, constSrcSizeInBytes, constDstPtr, constDstSizeInBytes, 0&)
        End If
        
        'If decompression failed, try it again, but with the explicit instruction to attempt to decompress enough
        ' data to fill the output buffer.  If there is still more compressed data past that point, it will be
        ' lost/ignored, but this may produce enough data for the caller to proceed normally.
        If (lReturn <> 0) And allowFallbacks Then
            InternalError "LibDeflateDecompress", "full decompress failed; attempting partial decompress instead..."
            Dim bytesWritten As Long
            If (cmpFormat = cf_Zlib) Then
                lReturn = libdeflate_zlib_decompress(hDecompress, constSrcPtr, constSrcSizeInBytes, constDstPtr, constDstSizeInBytes, bytesWritten)
            ElseIf (cmpFormat = cf_Deflate) Then
                lReturn = libdeflate_deflate_decompress(hDecompress, constSrcPtr, constSrcSizeInBytes, constDstPtr, constDstSizeInBytes, bytesWritten)
            ElseIf (cmpFormat = cf_Gzip) Then
                lReturn = libdeflate_gzip_decompress(hDecompress, constSrcPtr, constSrcSizeInBytes, constDstPtr, constDstSizeInBytes, bytesWritten)
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
        libdeflate_free_decompressor hDecompress
        
    Else
        InternalError "LibDeflateDecompress", "WARNING!  Failed to initialize a libdeflate decompressor."
    End If
    
End Function

'Magic numbers are taken from libdeflate.h
Private Function GetDecompressErrorText(ByVal srcErrNum As Long) As String
    If (srcErrNum = 1) Then
        GetDecompressErrorText = "bad input data"
    ElseIf (srcErrNum = 2) Then
        GetDecompressErrorText = "short output"
    ElseIf (srcErrNum = 3) Then
        GetDecompressErrorText = "insufficient destination buffer"
    End If
End Function

'Note that libdeflate exports its own "get worst-case dst size" function.  However, it requires you to
' pass a compressor handle that has been initialized to the target compression level... which creates
' problems for the way our ICompress interface works.  Because there's no good way to mimic this,
' we simply use the standard zlib "worst case" calculation, but add extra bytes for the gzip case
' (as gzip headers/trailers are larger than zlib ones).
Public Function GetWorstCaseSize(ByVal srcBufferSizeInBytes As Long, Optional ByVal compressionLevel As Long = -1, Optional ByVal cmpFormat As PD_CompressionFormat = cf_Zlib) As Long
    
    ValidateCompressionLevel compressionLevel
    
    'libdeflate requires a compressor object in order to calculate a "worst-case" size
    Dim hCompress As Long
    hCompress = libdeflate_alloc_compressor(compressionLevel)
    If (hCompress <> 0) Then
    
        If (cmpFormat = cf_Zlib) Then
            ' LIBDEFLATEAPI size_t libdeflate_deflate_compress_bound(struct libdeflate_compressor *compressor, size_t in_nbytes)
            GetWorstCaseSize = libdeflate_deflate_compress_bound(hCompress, srcBufferSizeInBytes)
        ElseIf (cmpFormat = cf_Deflate) Then
            ' LIBDEFLATEAPI size_t libdeflate_zlib_compress_bound(struct libdeflate_compressor *compressor, size_t in_nbytes)
            GetWorstCaseSize = libdeflate_zlib_compress_bound(hCompress, srcBufferSizeInBytes)
        ElseIf (cmpFormat = cf_Gzip) Then
            ' LIBDEFLATEAPI size_t libdeflate_gzip_compress_bound(struct libdeflate_compressor *compressor, size_t in_nbytes)
            GetWorstCaseSize = libdeflate_gzip_compress_bound(hCompress, srcBufferSizeInBytes)
        End If
        
        'Free the compressor before exiting
        libdeflate_free_compressor hCompress
        
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

'libdeflate doesn't export a version function, but this class was last tested against the v1.9 release (released Jan 2022).
Public Function GetCompressorVersion() As String
    GetCompressorVersion = "1.9"
End Function

'Private methods follow

'Clamp requested compression levels to valid inputs, and resolve negative numbers to the engine's default value.
Private Sub ValidateCompressionLevel(ByRef inputLevel As Long)
    If (inputLevel = -1) Then inputLevel = LIBDEFLATE_DEFAULT_CLEVEL
    If (inputLevel < LIBDEFLATE_MIN_CLEVEL) Then inputLevel = LIBDEFLATE_MIN_CLEVEL
    If (inputLevel > LIBDEFLATE_MAX_CLEVEL) Then inputLevel = LIBDEFLATE_MAX_CLEVEL
End Sub

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String, Optional ByVal writeDebugLog As Boolean = True)
    If UserPrefs.GenerateDebugLogs Then
        If writeDebugLog Then PDDebug.LogAction "Plugin_libdeflate." & funcName & "() reported an error: " & errDescription
    Else
        Debug.Print "Plugin_libdeflate." & funcName & "() reported an error: " & errDescription
    End If
End Sub
