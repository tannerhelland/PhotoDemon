Attribute VB_Name = "Plugin_libdeflate"
'***************************************************************************
'File Compression Interface (via libdeflate)
'Copyright 2002-2016 by Tanner Helland
'Created: 3/02/02
'Last updated: 13/December/22
'Last update: libdeflate has switched calling conventions AGAIN so I'm now going with a callconv-agnostic solution
'             (so I don't have to keep rewriting this module against stdcall OR cdecl depending on the library
'              author's mood, ugh).
'
'LibDeflate: https://github.com/ebiggers/libdeflate
' - "libdeflate is a library for fast, whole-buffer DEFLATE-based compression and decompression."
' - "libdeflate is heavily optimized. It is significantly faster than the zlib library, both for
'    compression and decompression, and especially on x86 processors."
'
'PhotoDemon uses libdeflate for reading/writing DEFLATE, zlib, and gzip data buffers.
'
'In v1.4, libdeflate authors switched to using stdcall builds for their 32-bit Windows builds.  Unfortunately for
' me, they reversed this decision several years later in v1.13, so if you check the history of this module on GitHub
' you'll see me manually switching between callconv implementations over the years.  I am now using a
' callconv-agnostic solution because I never want to rewrite all this code again!
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

'libdeflate was originally built using cdecl.  They later changed this to stdcall, then years later changed *back* to cdecl.
' I now just use a DispCallFunc wrapper that allows for callconv-agnostic access to the library.
'Private Declare Function libdeflate_alloc_compressor Lib "libdeflate" (ByVal compression_level As Long) As Long
'Private Declare Function libdeflate_deflate_compress Lib "libdeflate" (ByVal libdeflate_compressor As Long, ByVal ptr_in As Long, ByVal in_nbytes As Long, ByVal ptr_out As Long, ByVal out_nbytes_avail As Long) As Long
'Private Declare Function libdeflate_deflate_compress_bound Lib "libdeflate" (ByVal libdeflate_compressor As Long, ByVal in_nbytes As Long) As Long
'Private Declare Function libdeflate_zlib_compress Lib "libdeflate" (ByVal libdeflate_compressor As Long, ByVal ptr_in As Long, ByVal in_nbytes As Long, ByVal ptr_out As Long, ByVal out_nbytes_avail As Long) As Long
'Private Declare Function libdeflate_zlib_compress_bound Lib "libdeflate" (ByVal libdeflate_compressor As Long, ByVal in_nbytes As Long) As Long
'Private Declare Function libdeflate_gzip_compress Lib "libdeflate" (ByVal libdeflate_compressor As Long, ByVal ptr_in As Long, ByVal in_nbytes As Long, ByVal ptr_out As Long, ByVal out_nbytes_avail As Long) As Long
'Private Declare Function libdeflate_gzip_compress_bound Lib "libdeflate" (ByVal libdeflate_compressor As Long, ByVal in_nbytes As Long) As Long
'Private Declare Sub libdeflate_free_compressor Lib "libdeflate" (ByVal libdeflate_compressor As Long)
'Private Declare Function libdeflate_alloc_decompressor Lib "libdeflate" () As Long
'Private Declare Function libdeflate_deflate_decompress Lib "libdeflate" (ByVal libdeflate_decompressor As Long, ByVal ptr_in As Long, ByVal in_nbytes As Long, ByVal ptr_out As Long, ByVal out_nbytes_avail As Long, ByRef actual_out_nbytes_ret As Long) As LibDeflate_Result
'Private Declare Function libdeflate_zlib_decompress Lib "libdeflate" (ByVal libdeflate_decompressor As Long, ByVal ptr_in As Long, ByVal in_nbytes As Long, ByVal ptr_out As Long, ByVal out_nbytes_avail As Long, ByRef actual_out_nbytes_ret As Long) As LibDeflate_Result
'Private Declare Function libdeflate_gzip_decompress Lib "libdeflate" (ByVal libdeflate_decompressor As Long, ByVal ptr_in As Long, ByVal in_nbytes As Long, ByVal ptr_out As Long, ByVal out_nbytes_avail As Long, ByRef actual_out_nbytes_ret As Long) As LibDeflate_Result
'Private Declare Sub libdeflate_free_decompressor Lib "libdeflate" (ByVal libdeflate_decompressor As Long)
'Private Declare Function libdeflate_adler32 Lib "libdeflate" (ByVal adler32 As Long, ByVal ptr_buffer As Long, ByVal len_in_bytes As Long) As Long
'Private Declare Function libdeflate_crc32 Lib "libdeflate" (ByVal crc As Long, ByVal ptr_buffer As Long, ByVal len_in_bytes As Long) As Long

'libdeflate provides _ex versions of all decompress functions, where it simply attempts to decompress
' until it can't anymore.  PD never uses these as there are strong security implications, and we only
' decompress in contexts where the decompressed size is known in advance.
'Private Declare Function libdeflate_deflate_decompress_ex Lib "libdeflate" (ByVal libdeflate_decompressor As Long, ByVal ptr_in As Long, ByVal in_nbytes As Long, ByVal ptr_out As Long, ByVal out_nbytes_avail As Long, ByRef actual_in_nbytes_ret As Long, ByRef actual_out_nbytes_ret As Long) As LibDeflate_Result
'Private Declare Function libdeflate_zlib_decompress_ex Lib "libdeflate" (ByVal libdeflate_decompressor As Long, ByVal ptr_in As Long, ByVal in_nbytes As Long, ByVal ptr_out As Long, ByVal out_nbytes_avail As Long, ByRef actual_in_nbytes_ret As Long, ByRef actual_out_nbytes_ret As Long) As LibDeflate_Result
'Private Declare Function libdeflate_gzip_decompress_ex Lib "libdeflate" (ByVal libdeflate_decompressor As Long, ByVal ptr_in As Long, ByVal in_nbytes As Long, ByVal ptr_out As Long, ByVal out_nbytes_avail As Long, ByRef actual_in_nbytes_ret As Long, ByRef actual_out_nbytes_ret As Long) As LibDeflate_Result

'This library has very specific compiler needs in order to produce maximum perf code, so rather than
' custom compile it, I use the official Windows binaries and wrap 'em using DispCallFunc.
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef retVAR As Variant) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

'At load-time, we cache a number of proc addresses (required for passing through DispCallFunc).
' This saves us a little time vs calling GetProcAddress on each call.
Private Enum libdeflate_ProcAddress
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
    libdeflate_zlib_decompress
    libdeflate_gzip_decompress
    libdeflate_free_decompressor
    libdeflate_adler32
    libdeflate_crc32
    [last_address]
End Enum

Private m_ProcAddresses() As Long

'Rather than allocate new memory on each DispCallFunc invoke, just reuse a set of temp arrays declared
' to a maximum relevant size (see InitializeEngine, below).
Private Const MAX_PARAM_COUNT As Long = 8
Private m_vType() As Integer, m_vPtr() As Long

'If for some reason the local copy of libdeflate is an (old) stdcall build, this flag will be set to TRUE
' by the library loader.
Private m_UseStdCallFallback As Boolean

'Basic init/release functions
Public Function InitializeEngine(ByRef pathToDLLFolder As String) As Boolean

    Dim libDeflatePath As String
    libDeflatePath = pathToDLLFolder & "libdeflate.dll"
    m_libDeflateHandle = VBHacks.LoadLib(libDeflatePath)
    InitializeEngine = (m_libDeflateHandle <> 0)
    
    If InitializeEngine Then
    
        'Pre-load all relevant proc addresses
        ReDim m_ProcAddresses(0 To [last_address] - 1) As Long
        m_ProcAddresses(libdeflate_alloc_compressor) = GetProcAddressHelper(m_libDeflateHandle, "libdeflate_alloc_compressor")
        m_ProcAddresses(libdeflate_deflate_compress) = GetProcAddressHelper(m_libDeflateHandle, "libdeflate_deflate_compress")
        m_ProcAddresses(libdeflate_deflate_compress_bound) = GetProcAddressHelper(m_libDeflateHandle, "libdeflate_deflate_compress_bound")
        m_ProcAddresses(libdeflate_zlib_compress) = GetProcAddressHelper(m_libDeflateHandle, "libdeflate_zlib_compress")
        m_ProcAddresses(libdeflate_zlib_compress_bound) = GetProcAddressHelper(m_libDeflateHandle, "libdeflate_zlib_compress_bound")
        m_ProcAddresses(libdeflate_gzip_compress) = GetProcAddressHelper(m_libDeflateHandle, "libdeflate_gzip_compress")
        m_ProcAddresses(libdeflate_gzip_compress_bound) = GetProcAddressHelper(m_libDeflateHandle, "libdeflate_gzip_compress_bound")
        m_ProcAddresses(libdeflate_free_compressor) = GetProcAddressHelper(m_libDeflateHandle, "libdeflate_free_compressor")
        m_ProcAddresses(libdeflate_alloc_decompressor) = GetProcAddressHelper(m_libDeflateHandle, "libdeflate_alloc_decompressor")
        m_ProcAddresses(libdeflate_deflate_decompress) = GetProcAddressHelper(m_libDeflateHandle, "libdeflate_deflate_decompress")
        m_ProcAddresses(libdeflate_zlib_decompress) = GetProcAddressHelper(m_libDeflateHandle, "libdeflate_zlib_decompress")
        m_ProcAddresses(libdeflate_gzip_decompress) = GetProcAddressHelper(m_libDeflateHandle, "libdeflate_gzip_decompress")
        m_ProcAddresses(libdeflate_free_decompressor) = GetProcAddressHelper(m_libDeflateHandle, "libdeflate_free_decompressor")
        m_ProcAddresses(libdeflate_adler32) = GetProcAddressHelper(m_libDeflateHandle, "libdeflate_adler32")
        m_ProcAddresses(libdeflate_crc32) = GetProcAddressHelper(m_libDeflateHandle, "libdeflate_crc32")
        
        'Initialize all module-level arrays
        ReDim m_vType(0 To MAX_PARAM_COUNT - 1) As Integer
        ReDim m_vPtr(0 To MAX_PARAM_COUNT - 1) As Long
        
    Else
        PDDebug.LogAction "WARNING!  LoadLibraryW failed to load libdeflate.  Last DLL error: " & Err.LastDllError
    End If
    
End Function

Private Function GetProcAddressHelper(ByVal hLibrary As Long, ByRef srcFuncName As String) As Long

    'Attempt to load the function as-is
    GetProcAddressHelper = GetProcAddress(hLibrary, srcFuncName)
    If (GetProcAddressHelper <> 0) Then
        m_UseStdCallFallback = False
        
    'If the load function failed, the most likely explanation is that an stdcall version of the library
    ' (with name-mangling) exists.  Try to find a new address w/name-mangling.
    Else
        
        Dim i As Long, tmpFuncName As String
        For i = 0 To 64 'Arbitrary upper limit!
            tmpFuncName = "_" & srcFuncName & "@" & Trim$(Str$(i))
            GetProcAddressHelper = GetProcAddress(hLibrary, tmpFuncName)
            If (GetProcAddressHelper <> 0) Then
                m_UseStdCallFallback = True
                Exit Function
            End If
        Next i
        
    End If

End Function

Public Sub ReleaseEngine()
    If (m_libDeflateHandle <> 0) Then
        VBHacks.FreeLib m_libDeflateHandle
        m_libDeflateHandle = 0
    End If
End Sub

'libdeflate doesn't export a version function, but this class was last tested against the v1.19 release.
Public Function GetCompressorVersion() As String
    GetCompressorVersion = "1.23"
End Function

Public Function CompressPtrToPtr(ByVal constDstPtr As Long, ByRef dstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long, Optional ByVal compressionLevel As Long = -1, Optional ByVal cmpFormat As PD_CompressionFormat = cf_Zlib) As Boolean
    CompressPtrToPtr = LibDeflateCompress(constDstPtr, dstSizeInBytes, constSrcPtr, constSrcSizeInBytes, compressionLevel, cmpFormat)
End Function

Public Function DecompressPtrToPtr(ByVal constDstPtr As Long, ByVal constDstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long, Optional ByVal cmpFormat As PD_CompressionFormat = cf_Zlib, Optional ByVal allowFallbacks As Boolean = True, Optional ByVal suppressErrors As Boolean = False) As Boolean
    DecompressPtrToPtr = LibDeflateDecompress(constDstPtr, constDstSizeInBytes, constSrcPtr, constSrcSizeInBytes, cmpFormat, allowFallbacks, suppressErrors)
End Function

Public Function GetCrc32(ByVal srcPtr As Long, ByVal srcLen As Long, Optional ByVal startValue As Long = 0&, Optional ByVal calcStartForMe As Boolean = True) As Long
    
    'Get an initial default "seed"
    If calcStartForMe Then startValue = CallCDeclW(libdeflate_crc32, vbLong, 0&, 0&, 0&)
    
    'Use the seed to calculate checksum
    If (srcPtr <> 0) Then GetCrc32 = CallCDeclW(libdeflate_crc32, vbLong, startValue, srcPtr, srcLen) Else GetCrc32 = startValue
    
End Function

Public Function GetAdler32(ByVal srcPtr As Long, ByVal srcLen As Long, Optional ByVal startValue As Long = 0&) As Long
    
    'Get an initial default "seed"
    If (startValue = 0) Then startValue = CallCDeclW(libdeflate_adler32, vbLong, 0&, 0&, 0&)
    
    'Use the seed to calculate checksum
    GetAdler32 = CallCDeclW(libdeflate_adler32, vbLong, startValue, srcPtr, srcLen)
    
End Function

'Specialized wrappers follow.
'
'Return the precise return code from specific decompress operations; this is helpful for obnoxious
' edge-cases like PNG files where compressed chunks don't store the original data size, so we
' have to manually attempt ever-larger buffers.
Public Function Decompress_GZip(ByVal constDstPtr As Long, ByVal constDstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long, ByRef actualBytesWritten As Long) As Long
    
    'Allocate a decompressor
    Dim hDecompress As Long
    hDecompress = CallCDeclW(libdeflate_alloc_decompressor, vbLong)
    If (hDecompress <> 0) Then
        actualBytesWritten = 0
        Decompress_GZip = CallCDeclW(libdeflate_gzip_decompress, vbLong, hDecompress, constSrcPtr, constSrcSizeInBytes, constDstPtr, constDstSizeInBytes, VarPtr(actualBytesWritten))
        CallCDeclW libdeflate_free_decompressor, vbEmpty, hDecompress
    Else
        InternalError "Decompress_GZip", "Failed to initialize a decompressor"
    End If
    
End Function

Public Function Decompress_ZLib(ByVal constDstPtr As Long, ByVal constDstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long) As Long
    
    'Allocate a decompressor
    ' LIBDEFLATEAPI struct libdeflate_decompressor * libdeflate_alloc_decompressor(void)
    Dim hDecompress As Long
    hDecompress = CallCDeclW(libdeflate_alloc_decompressor, vbLong)
    If (hDecompress <> 0) Then
        
        'Perform decompression
        ' LIBDEFLATEAPI enum libdeflate_result libdeflate_zlib_decompress(struct libdeflate_decompressor *decompressor, const void *in, size_t in_nbytes, void *out, size_t out_nbytes_avail, size_t *actual_out_nbytes_ret)
        ' "If the actual uncompressed size is known, then pass the actual
        '  uncompressed size as 'out_nbytes_avail' and pass NULL for
        '  actual_out_nbytes_ret'.  This makes libdeflate_deflate_decompress() fail
        '  with LIBDEFLATE_SHORT_OUTPUT if the data decompressed to fewer than the
        '  specified number of bytes."
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

Private Function LibDeflateDecompress(ByVal constDstPtr As Long, ByVal constDstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long, Optional ByVal cmpFormat As PD_CompressionFormat = cf_Zlib, Optional ByVal allowFallbacks As Boolean = True, Optional ByVal suppressErrors As Boolean = False) As Boolean
    
    'Allocate a decompressor
    ' LIBDEFLATEAPI struct libdeflate_decompressor * libdeflate_alloc_decompressor(void)
    Dim hDecompress As Long
    hDecompress = CallCDeclW(libdeflate_alloc_decompressor, vbLong)
    If (hDecompress <> 0) Then
        
        'Perform decompression
        ' LIBDEFLATEAPI enum libdeflate_result libdeflate_zlib_decompress(struct libdeflate_decompressor *decompressor, const void *in, size_t in_nbytes, void *out, size_t out_nbytes_avail, size_t *actual_out_nbytes_ret)
        Dim lReturn As LibDeflate_Result
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
        If (lReturn <> ld_Success) And allowFallbacks Then
            
            If (Not suppressErrors) Then InternalError "LibDeflateDecompress", "full decompress failed (" & lReturn & "); attempting partial decompress instead..."
            
            Dim bytesWritten As Long
            If (cmpFormat = cf_Zlib) Then
                lReturn = CallCDeclW(libdeflate_zlib_decompress, vbLong, hDecompress, constSrcPtr, constSrcSizeInBytes, constDstPtr, constDstSizeInBytes, VarPtr(bytesWritten))
            ElseIf (cmpFormat = cf_Deflate) Then
                lReturn = CallCDeclW(libdeflate_deflate_decompress, vbLong, hDecompress, constSrcPtr, constSrcSizeInBytes, constDstPtr, constDstSizeInBytes, VarPtr(bytesWritten))
            ElseIf (cmpFormat = cf_Gzip) Then
                lReturn = CallCDeclW(libdeflate_gzip_decompress, vbLong, hDecompress, constSrcPtr, constSrcSizeInBytes, constDstPtr, constDstSizeInBytes, VarPtr(bytesWritten))
            End If
            
            If (constDstSizeInBytes = bytesWritten) Then
                lReturn = ld_Success
                If (Not suppressErrors) Then InternalError "LibDeflateDecompress", "Partial decompression successful."
            Else
                If (lReturn = ld_Success) Then
                    If (Not suppressErrors) Then InternalError "LibDeflateDecompress", "Partial decompression worked but bytes written is unexpected (" & bytesWritten & " written vs " & constDstSizeInBytes & " expected)"
                Else
                    If (Not suppressErrors) Then InternalError "LibDeflateDecompress", "Partial decompression failed; source data is likely corrupt. (" & lReturn & ", " & bytesWritten & ")"
                End If
            End If
            
        End If
        
        'If compression still failed, we're outta luck.
        LibDeflateDecompress = (lReturn = 0)
        If (Not LibDeflateDecompress) Then
            If (Not suppressErrors) Then InternalError "LibDeflateDecompress", "operation failed; return was " & lReturn & " " & GetDecompressErrorText(lReturn)
            If (Not suppressErrors) Then InternalError "LibDeflateDecompress", "FYI inputs were: " & constDstPtr & ", " & constDstSizeInBytes & ", " & constSrcPtr & ", " & constSrcSizeInBytes
        End If
        
        ' Make sure we free the compressor before exiting
        'LIBDEFLATEAPI void libdeflate_free_decompressor(struct libdeflate_decompressor *decompressor);
        CallCDeclW libdeflate_free_decompressor, vbEmpty, hDecompress
        
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

'Private methods follow

'Clamp requested compression levels to valid inputs, and resolve negative numbers to the engine's default value.
Private Sub ValidateCompressionLevel(ByRef inputLevel As Long)
    If (inputLevel = -1) Then inputLevel = LIBDEFLATE_DEFAULT_CLEVEL
    If (inputLevel < LIBDEFLATE_MIN_CLEVEL) Then inputLevel = LIBDEFLATE_MIN_CLEVEL
    If (inputLevel > LIBDEFLATE_MAX_CLEVEL) Then inputLevel = LIBDEFLATE_MAX_CLEVEL
End Sub

'DispCallFunc wrapper originally by Olaf Schmidt, with a few minor modifications; see the top of this class
' for a link to his original, unmodified version
Private Function CallCDeclW(ByVal lProc As libdeflate_ProcAddress, ByVal fRetType As VbVarType, ParamArray pa() As Variant) As Variant
    
    'I normally don't like to use error-handling on perf-sensitive functions, but we may need to detect
    ' different calling conventions (because libdeflate has been built in various ways over the years).
    ' On a "bad calling convention" error, we'll try again with the other callconv.
    On Error GoTo TryOtherCallConv
    
    Dim i As Long, vTemp() As Variant, hResult As Long
    
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
    
    Const CC_CDECL As Long = 1, CC_STDCALL = 4
    If m_UseStdCallFallback Then
        hResult = DispCallFunc(0, m_ProcAddresses(lProc), CC_STDCALL, fRetType, i, m_vType(0), m_vPtr(0), CallCDeclW)
    Else
        hResult = DispCallFunc(0, m_ProcAddresses(lProc), CC_CDECL, fRetType, i, m_vType(0), m_vPtr(0), CallCDeclW)
    End If
    
    If (hResult <> 0) Then InternalError "CallCDeclW", "bad hresult: " & hResult
    Exit Function
    
TryOtherCallConv:
    On Error GoTo 0
    If m_UseStdCallFallback Then
        hResult = DispCallFunc(0, m_ProcAddresses(lProc), CC_CDECL, fRetType, i, m_vType(0), m_vPtr(0), CallCDeclW)
        If (hResult = 0) Then m_UseStdCallFallback = False
    Else
        hResult = DispCallFunc(0, m_ProcAddresses(lProc), CC_STDCALL, fRetType, i, m_vType(0), m_vPtr(0), CallCDeclW)
        If (hResult = 0) Then m_UseStdCallFallback = True
    End If
    
End Function

Private Sub InternalError(ByRef funcName As String, ByRef errDescription As String, Optional ByVal writeDebugLog As Boolean = True)
    If UserPrefs.GenerateDebugLogs Then
        If writeDebugLog Then PDDebug.LogAction "Plugin_libdeflate." & funcName & "() reported an error: " & errDescription
    Else
        Debug.Print "Plugin_libdeflate." & funcName & "() reported an error: " & errDescription
    End If
End Sub
