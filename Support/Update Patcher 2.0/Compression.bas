Attribute VB_Name = "Compression"
'***************************************************************************
'Unified Compression Interface for PhotoDemon
'Copyright 2016-2019 by Tanner Helland
'Created: 02/December/16
'Last updated: 13/February/19
'Last update: overhaul our internal compression API as part of embedding libdeflate for deflate/zlib/gzip tasks
'Dependencies: - standalone plugin modules for whatever compression engines you want to use.  This module
'              simply wraps those dedicated wrappers in a more convenient format.
'
'As of v7.0, PhotoDemon performs a *lot* of custom compression work.  There are a lot of different needs in
' image processing - for example, when the user saves a large, multi-layer image, it's okay to take plenty of time
' and squeeze every last bit of compression you can out of the finished file (which is potentially enormous).
' But when saving Undo/Redo data for rapid operations like paint strokes, you want to dump data out to file as
' quickly as humanly possible, with a compression strategy that's as close as possible to HDD performance limits.
'
'Different compression engines work better for different workloads, so PD currently ships a few different
' 3rd-party solutions.  Standard DEFLATE/zlib/gzip compression is provided by libdeflate (this is used by
' a number of image format parsers, e.g. PNG, PSD, OpenRaster).  For better compression results, zstd is
' available (and used by PD's native PDI format).  For the fastest possible compression, lz4 is also
' available (and used heavily by PD's Undo/Redo engine).
'
'The purpose of this module is to simplify compression tasks by exposing standardized function signatures.
' Simply specify the compressor you desire, and this module will silently plug in the right compression or
' decompression code.  (Note that - at present - you *must* request the correct decompressor at decompression
' time, meaning you can't just hand a compressed stream to this module and expect it to magically
' reverse-engineer which decompression engine to use.  That's *your* job.)
'
'All wrapper code in this function is written from scratch by me.  It is not based on any preexisting work.
' This module is, as usual, licensed under the same BSD license governing PD as a whole, so feel free to use
' it in any application, commercial or otherwise.  Bug reports are always welcome.
'
'Licenses for wrapped libraries include:
' libdeflate: MIT license (https://github.com/ebiggers/libdeflate/blob/master/COPYING)
' zstd: BSD 3-clause license (https://github.com/facebook/zstd/blob/dev/LICENSE)
' lz4/lz4-hc: BSD 2-clause license (https://github.com/lz4/lz4/blob/dev/LICENSE)
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'Currently supported compression formats.  Note that you must *always* use the same format for compression
' and decompression (e.g. there is no way to auto-detect the format of a previously compressed stream).
Public Enum PD_CompressionFormat
    
    '"None" just copies source data to destination data as-is.
    cf_None = 0
    
    'The following compression engines require a 3rd-party DLL.
    cf_Zlib = 1
    cf_Zstd = 2
    cf_Lz4 = 3
    cf_Lz4hc = 4
    cf_Deflate = 5
    cf_Gzip = 6
    
    [cf_Last] = 6
    
End Enum

#If False Then
    Private Const cf_None = 0, cf_Zlib = 1, cf_Zstd = 2, cf_Lz4 = 3, cf_Lz4hc = 4, cf_Deflate = 5, cf_Gzip = 6, cf_Last = 6
#End If

Private Enum PD_CompressionEngine
    ce_None = 0
    ce_LibDeflate = 1
    ce_Zstd = 2
    ce_Lz4 = 3
    [ce_Last] = 3
End Enum

#If False Then
    Private Const ce_None = 0, ce_LibDeflate = 1, ce_Zstd = 2, ce_Lz4 = 3, ce_Last = 3
#End If

'When a compression engine is initialized successfully, the matching value in this array will be set to TRUE.
' Note that there is *not* a 1:1 mapping between compression formats and compression engines, because some
' engines handle multiple formats.
Private m_CompressorAvailable() As Boolean

'Initialize all supported compression engines.  The path to the DLL folder *must* include a trailing slash.
'Returns: TRUE if initialization is successful for all engines; FALSE otherwise.  FALSE typically means the
'         path to the DLL folder is malformed, or it's correct but the program doesn't have access rights to it.
Public Function StartCompressionEngines(ByRef pathToDLLFolder As String) As Boolean
    
    'Keep track of which compression engines have been initialized
    If (Not VBHacks.IsArrayInitialized(m_CompressorAvailable)) Then
        ReDim m_CompressorAvailable(0 To ce_Last) As Boolean
        m_CompressorAvailable(ce_None) = True
    End If
    
    'Skip initialization if a compressor has already been initialized
   ' If (Not m_CompressorAvailable(ce_LibDeflate)) Then m_CompressorAvailable(ce_LibDeflate) = Plugin_libdeflate.InitializeEngine(pathToDLLFolder)
    If (Not m_CompressorAvailable(ce_Zstd)) Then m_CompressorAvailable(ce_Zstd) = Plugin_zstd.InitializeZStd(pathToDLLFolder)
    If (Not m_CompressorAvailable(ce_Lz4)) Then m_CompressorAvailable(ce_Lz4) = Plugin_lz4.InitializeLz4(pathToDLLFolder)
    
    StartCompressionEngines = True
    Dim i As Long
    For i = 0 To ce_Last
        If (Not m_CompressorAvailable(i)) Then StartCompressionEngines = False
    Next i
    
End Function

'Stop all compression engines.  You (obviously) cannot use compression abilities once the engines are released.
' You *must* call this function before your program terminates.
Public Sub StopCompressionEngines()

    'Keep track of which compression engines have been initialized
    If VBHacks.IsArrayInitialized(m_CompressorAvailable) Then
        
        'Skip termination if a compressor has already been shut down
        If m_CompressorAvailable(ce_LibDeflate) Then
            'Plugin_libdeflate.ReleaseEngine
            m_CompressorAvailable(ce_LibDeflate) = False
        End If
        
        If m_CompressorAvailable(ce_Zstd) Then
            Plugin_zstd.ReleaseZstd
            m_CompressorAvailable(ce_Zstd) = False
        End If
            
        If m_CompressorAvailable(ce_Lz4) Then
            Plugin_lz4.ReleaseLz4
            m_CompressorAvailable(ce_Lz4) = False
        End If
        
    End If
    
End Sub

'Want to know if a given compression engine is available?  Call this function.  It will (obviously) return FALSE for
' any engines that weren't initialized properly.
Public Function IsFormatSupported(ByVal cmpFormat As PD_CompressionFormat) As Boolean
    Select Case cmpFormat
        Case cf_None
            IsFormatSupported = m_CompressorAvailable(ce_None)
        Case cf_Zlib
            IsFormatSupported = m_CompressorAvailable(ce_LibDeflate)
        Case cf_Zstd
            IsFormatSupported = m_CompressorAvailable(ce_Zstd)
        Case cf_Lz4
            IsFormatSupported = m_CompressorAvailable(ce_Lz4)
        Case cf_Lz4hc
            IsFormatSupported = m_CompressorAvailable(ce_Lz4)
        Case cf_Deflate
            IsFormatSupported = m_CompressorAvailable(ce_LibDeflate)
        Case cf_Gzip
            IsFormatSupported = m_CompressorAvailable(ce_LibDeflate)
        Case Else
            IsFormatSupported = False
    End Select
End Function

'Compress some arbitrary pointer to a destination array.
'
'Required inputs:
' 1) ByRef destination array, declared As Byte.
' 2) ByRef final compressed size, as Long.  You need to cache this value with your compressed data,
'    so the decompression engine knows how large of a buffer to prepare later on.  (Some decompression engines
'    may be able to calculate this internally, but for that functionality, you'll need to call them directly.)
' 3) ByVal pointer to the source data.  This can be any valid pointer, aligned or not.
' 4) ByVal size of the source data.  This must be byte-accurate, NO EXCEPTIONS.
' 5) ByVal desired compression format.  Note that "cf_None" is a valid option; this module works just fine with
'    uncompressed data, and it will simply perform a fast copy instead (where destination size = source size).
'
'Optional inputs:
' 6) Desired compression level.  This parameter has different meanings for different compression engines.  -1 will use
'    each engine's default setting.  For zLib and zstd, higher values mean *slower but better* compression.  lz4 is the
'    opposite; higher values mean *faster but worse* compression.
' 7) If the caller has already prepared the destination array at an appropriate size, pass TRUE for dstArrayIsAlreadySized.
'    This spares us a memory allocation, which can improve performance.  (Note that no verifications are done on the
'    target array, so you *must* have resized the array to a size >= the maximum required size, as calculated by the
'    GetWorstCaseSize() function, ideally.)
' 8) If you want the destination array trimmed to the exact compressed size, pass TRUE for trimCompressedArray.  If you do
'    not specify this, dstArray() will be left at the worst-case size, and it is up to the caller to check the value of
'    dstCompressedSize to see how much size the compressed data actually consumed.
'
'Returns:
' - TRUE if compression was successful; FALSE otherwise.  Note that a FALSE return will still *always* copy the uncompressed
'   source bytes into the destination array, so you can potentially proceed with processing even if the function fails.
Public Function CompressPtrToDstArray(ByRef dstArray() As Byte, ByRef dstCompressedSizeInBytes As Long, ByVal ptrToSource As Long, ByVal srcBufferSizeInBytes As Long, ByVal cmpFormat As PD_CompressionFormat, Optional ByVal compressionLevel As Long = -1, Optional ByVal dstArrayIsAlreadySized As Boolean = False, Optional ByVal trimCompressedArray As Boolean = False) As Boolean

    'If the destination array isn't allocated, forcibly initialize it now
    If (Not dstArrayIsAlreadySized) Then
        dstCompressedSizeInBytes = GetWorstCaseSize(srcBufferSizeInBytes, cmpFormat)
        ReDim dstArray(0 To dstCompressedSizeInBytes - 1) As Byte
    End If
    
    'Now that our destination array is guaranteed sized correctly, use naked pointers for compression
    CompressPtrToDstArray = CompressPtrToPtr(VarPtr(dstArray(0)), dstCompressedSizeInBytes, ptrToSource, srcBufferSizeInBytes, cmpFormat, compressionLevel)
    
    'Trim the destination array, as requested
    If trimCompressedArray Then
        If (UBound(dstArray) <> dstCompressedSizeInBytes - 1) Then ReDim Preserve dstArray(0 To dstCompressedSizeInBytes - 1) As Byte
    End If
    
End Function

'All compression functions ultimately wrap this function.  You can also use it directly, but you *must* size your destination buffer
' correctly to avoid hard crashes.  Also, you *must* pass in the starting destination buffer size as dstSizeInBytes; the compressor
' needs to know this for security reasons.
Public Function CompressPtrToPtr(ByVal constDstPtr As Long, ByRef dstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long, ByVal cmpFormat As PD_CompressionFormat, Optional ByVal compressionLevel As Long = -1) As Boolean
    
    CompressPtrToPtr = False
    
    If (cmpFormat = cf_None) Then
        'Do nothing; the catch at the end of the function will handle this case for us
    ElseIf (cmpFormat = cf_Zlib) Then
        
        'libdeflate doesn't expose a "0" compression mode (it's treated as "default" compression),
        ' so for now, silently switch to compression mode 1.
        If (compressionLevel = 0) Then compressionLevel = 1
        'CompressPtrToPtr = Plugin_libdeflate.CompressPtrToPtr(constDstPtr, dstSizeInBytes, constSrcPtr, constSrcSizeInBytes, compressionLevel, cf_Zlib)
            
        'Want to compare against zLib?  FreeImage exposes a zlib default compress call
        'dstSizeInBytes = FreeImage_ZLibCompress(constDstPtr, dstSizeInBytes, constSrcPtr, constSrcSizeInBytes)
        CompressPtrToPtr = (dstSizeInBytes <> 0)
        
    ElseIf (cmpFormat = cf_Zstd) Then
        CompressPtrToPtr = Plugin_zstd.ZstdCompressNakedPointers(constDstPtr, dstSizeInBytes, constSrcPtr, constSrcSizeInBytes, compressionLevel)
    ElseIf (cmpFormat = cf_Lz4) Then
        CompressPtrToPtr = Plugin_lz4.Lz4CompressNakedPointers(constDstPtr, dstSizeInBytes, constSrcPtr, constSrcSizeInBytes, compressionLevel)
    ElseIf (cmpFormat = cf_Lz4hc) Then
        CompressPtrToPtr = Plugin_lz4.Lz4HCCompressNakedPointers(constDstPtr, dstSizeInBytes, constSrcPtr, constSrcSizeInBytes, compressionLevel)
    ElseIf (cmpFormat = cf_Deflate) Then
        'CompressPtrToPtr = Plugin_libdeflate.CompressPtrToPtr(constDstPtr, dstSizeInBytes, constSrcPtr, constSrcSizeInBytes, compressionLevel, cf_Deflate)
    ElseIf (cmpFormat = cf_Gzip) Then
        'CompressPtrToPtr = Plugin_libdeflate.CompressPtrToPtr(constDstPtr, dstSizeInBytes, constSrcPtr, constSrcSizeInBytes, compressionLevel, cf_Gzip)
    End If
    
    'If compression failed, perform a direct source-to-dst copy
    If (Not CompressPtrToPtr) Then
        If (cmpFormat <> ce_None) Then InternalErrorMsg "CompressPtrToPtr failed on compression format " & cmpFormat
        CopyMemoryStrict constDstPtr, constSrcPtr, constSrcSizeInBytes
        dstSizeInBytes = constSrcSizeInBytes
        CompressPtrToPtr = (cmpFormat = ce_None)
    End If

End Function

'Decompress some arbitrary pointer (containing compressed data, obviously) to a destination array.
'
'Required inputs:
' 1) ByRef destination array, declared As Byte.
' 2) ByVal final decompressed size, as Long.  You *must* pass this value to the function,
'    as the decompressed stream is unlikely to store this value independently.
' 3) ByVal pointer to the source data.  This can be any valid pointer, aligned or not.
' 4) ByVal size of the source data.  This must be byte-accurate, NO EXCEPTIONS.
' 5) ByVal desired compression engine.  Note that "no compression engine" is a valid option; this module works
'    just fine with uncompressed data, and it will simply perform a fast copy instead (where destination
'    size = source size).
'
'Optional inputs:
' 6) If the caller has already prepared the destination array at an appropriate size, pass TRUE for dstArrayIsAlreadySized.
'    This spares us a memory allocation, which can improve performance.  (Note that no verifications are done on the
'    target array, so you *must* have resized the array to a size >= the original decompressed size.)
'
'Returns:
' - TRUE if decompression was successful; FALSE otherwise.  Note that a FALSE return will still *always* copy the compressed
'   source bytes into the destination array, to mirror the behavior of the matching compression function, above.  (This allows
'   you to use the compression and decompression functions in "no compression" mode and have them behave as expected.)
'   If FALSE occurs, however, you may need to abandon further processing, as there's currently no way to decompress the
'   bytestream without help from the original decompression library.
Public Function DecompressPtrToDstArray(ByRef dstArray() As Byte, ByVal dstDecompressedSizeInBytes As Long, ByVal ptrToSource As Long, ByVal srcBufferSizeInBytes As Long, ByVal cmpFormat As PD_CompressionFormat, Optional ByVal dstArrayIsAlreadySized As Boolean = False) As Boolean
    
    'If the destination array isn't allocated, forcibly initialize it now
    If (Not dstArrayIsAlreadySized) Then ReDim dstArray(0 To dstDecompressedSizeInBytes - 1) As Byte
    
    'Now that our destination array is guaranteed sized correctly, use naked pointers for decompression
    DecompressPtrToDstArray = DecompressPtrToPtr(VarPtr(dstArray(0)), dstDecompressedSizeInBytes, ptrToSource, srcBufferSizeInBytes, cmpFormat)
    
End Function

'All decompression functions ultimately wrap this function.  You can also use it directly, but you *must* size your destination buffer
' correctly to avoid hard crashes.  Also, you *must* pass in the byte-accurate destination buffer size as dstSizeInBytes;
' most decompressors do not store this value independently.
Public Function DecompressPtrToPtr(ByVal constDstPtr As Long, ByVal dstSizeInBytes As Long, ByVal constSrcPtr As Long, ByVal constSrcSizeInBytes As Long, ByVal cmpFormat As PD_CompressionFormat) As Boolean
    
    DecompressPtrToPtr = False
    
    If (cmpFormat = cf_None) Then
        'Do nothing; the failsafe catch at the end of this function handles this case for us
    ElseIf (cmpFormat = cf_Zlib) Then
        
        'While I don't doubt libdeflate's capabilities, I've added an emergency zlib fallback for now... "just in case"
        'DecompressPtrToPtr = Plugin_libdeflate.DecompressPtrToPtr(constDstPtr, dstSizeInBytes, constSrcPtr, constSrcSizeInBytes, cf_Zlib)
        If (Not DecompressPtrToPtr) Then
        
            'libdeflate failed.  FreeImage embeds a copy of zlib; try it, and if it works, proceed as normal
            InternalErrorMsg "WARNING: libdeflate failed on a stream. cmp size: " & constSrcSizeInBytes & ", orig size: " & dstSizeInBytes
            
            Dim lRet As Long
            'lRet = FreeImage_ZLibUncompress(constDstPtr, dstSizeInBytes, constSrcPtr, constSrcSizeInBytes)
            DecompressPtrToPtr = (lRet = dstSizeInBytes)
            
            If DecompressPtrToPtr Then
                InternalErrorMsg "zlib fallback was successful; data decompressed OK"
            Else
                InternalErrorMsg "WARNING: zlib fallback failed!  Data is corrupt."
            End If
            
        End If
        
    ElseIf (cmpFormat = cf_Zstd) Then
        DecompressPtrToPtr = (Plugin_zstd.ZstdDecompress_UnsafePtr(constDstPtr, dstSizeInBytes, constSrcPtr, constSrcSizeInBytes) = dstSizeInBytes)
    ElseIf ((cmpFormat = cf_Lz4) Or (cmpFormat = cf_Lz4hc)) Then
        DecompressPtrToPtr = (Plugin_lz4.Lz4Decompress_UnsafePtr(constDstPtr, dstSizeInBytes, constSrcPtr, constSrcSizeInBytes) = dstSizeInBytes)
    ElseIf (cmpFormat = cf_Deflate) Then
        'DecompressPtrToPtr = Plugin_libdeflate.DecompressPtrToPtr(constDstPtr, dstSizeInBytes, constSrcPtr, constSrcSizeInBytes, cf_Deflate)
    ElseIf (cmpFormat = cf_Gzip) Then
        'DecompressPtrToPtr = Plugin_libdeflate.DecompressPtrToPtr(constDstPtr, dstSizeInBytes, constSrcPtr, constSrcSizeInBytes, cf_Gzip)
    End If
    
    'If compression failed, perform a direct source-to-dst copy
    If (Not DecompressPtrToPtr) Then
        If (cmpFormat <> cf_None) Then InternalErrorMsg "DecompressPtrToPtr failed on compression format " & cmpFormat
        CopyMemoryStrict constDstPtr, constSrcPtr, constSrcSizeInBytes
        DecompressPtrToPtr = (cmpFormat = cf_None)
    End If

End Function

'All compression functions require a destination buffer sized to the "worst-case" scenario size, which is the largest size
' the compressed data will consume if it is 100% incompressible.  You can almost always shrink the destination buffer after
' the fact (to the exact compressed size), but you must always start with a buffer at least this large.
'
'Obviously, you must pass the size of your source data, and you must also specify the desired compression engine (as they
' use different rules for formulating a "worst-case" size).
Public Function GetWorstCaseSize(ByVal srcBufferSizeInBytes As Long, ByVal cmpFormat As PD_CompressionFormat, Optional ByVal compressionLevel As Long = -1) As Long
    
    If (cmpFormat = cf_None) Then
        GetWorstCaseSize = srcBufferSizeInBytes
    ElseIf (cmpFormat = cf_Zlib) Then
        'GetWorstCaseSize = Plugin_libdeflate.GetWorstCaseSize(srcBufferSizeInBytes, compressionLevel, cf_Zlib)
    ElseIf (cmpFormat = cf_Zstd) Then
        GetWorstCaseSize = Plugin_zstd.ZstdGetMaxCompressedSize(srcBufferSizeInBytes)
    ElseIf (cmpFormat = cf_Lz4) Then
        GetWorstCaseSize = Plugin_lz4.Lz4GetMaxCompressedSize(srcBufferSizeInBytes)
    ElseIf (cmpFormat = cf_Lz4hc) Then
        GetWorstCaseSize = Plugin_lz4.Lz4GetMaxCompressedSize(srcBufferSizeInBytes)
    ElseIf (cmpFormat = cf_Deflate) Then
        'GetWorstCaseSize = Plugin_libdeflate.GetWorstCaseSize(srcBufferSizeInBytes, compressionLevel, cf_Deflate)
    ElseIf (cmpFormat = cf_Gzip) Then
        'GetWorstCaseSize = Plugin_libdeflate.GetWorstCaseSize(srcBufferSizeInBytes, compressionLevel, cf_Gzip)
    End If

End Function

'Retrieve default/min/max compression levels from a given library.  Note that these are *not* standardized,
' meaning each library has its own default/min/max levels, and the levels mean different things in different
' libraries.  For example, most libraries use the formula where "larger numbers = slower, better compression",
' but lz4 is the opposite - e.g. "larger lz4 numbers = faster, worse compression".
Public Function GetDefaultCompressionLevel(ByVal cmpFormat As PD_CompressionFormat) As Long

    If (cmpFormat = cf_Zlib) Then
        'GetDefaultCompressionLevel = Plugin_libdeflate.GetDefaultCompressionLevel()
    ElseIf (cmpFormat = cf_Zstd) Then
        GetDefaultCompressionLevel = Plugin_zstd.Zstd_GetDefaultCompressionLevel()
    ElseIf (cmpFormat = cf_Lz4) Then
        GetDefaultCompressionLevel = Plugin_lz4.Lz4_GetDefaultAccelerationLevel()
    ElseIf (cmpFormat = cf_Lz4hc) Then
        GetDefaultCompressionLevel = Plugin_lz4.Lz4HC_GetDefaultCompressionLevel()
    ElseIf (cmpFormat = cf_Deflate) Then
        'GetDefaultCompressionLevel = Plugin_libdeflate.GetDefaultCompressionLevel()
    ElseIf (cmpFormat = cf_Gzip) Then
        'GetDefaultCompressionLevel = Plugin_libdeflate.GetDefaultCompressionLevel()
    Else
        GetDefaultCompressionLevel = 0
    End If

End Function

Public Function GetMinCompressionLevel(ByVal cmpFormat As PD_CompressionFormat) As Long
    
    If (cmpFormat = cf_Zlib) Then
        'GetMinCompressionLevel = Plugin_libdeflate.GetMinCompressionLevel()
    ElseIf (cmpFormat = cf_Zstd) Then
        GetMinCompressionLevel = Plugin_zstd.Zstd_GetMinCompressionLevel()
    ElseIf (cmpFormat = cf_Lz4) Then
        GetMinCompressionLevel = Plugin_lz4.Lz4_GetMaxAccelerationLevel()
    ElseIf (cmpFormat = cf_Lz4hc) Then
        GetMinCompressionLevel = Plugin_lz4.Lz4HC_GetMinCompressionLevel()
    ElseIf (cmpFormat = cf_Deflate) Then
        'GetMinCompressionLevel = Plugin_libdeflate.GetMinCompressionLevel()
    ElseIf (cmpFormat = cf_Gzip) Then
        'GetMinCompressionLevel = Plugin_libdeflate.GetMinCompressionLevel()
    Else
        GetMinCompressionLevel = 0
    End If
    
End Function

Public Function GetMaxCompressionLevel(ByVal cmpFormat As PD_CompressionFormat) As Long

    If (cmpFormat = cf_Zlib) Then
        'GetMaxCompressionLevel = Plugin_libdeflate.GetMaxCompressionLevel()
    ElseIf (cmpFormat = cf_Zstd) Then
        GetMaxCompressionLevel = Plugin_zstd.Zstd_GetMaxCompressionLevel()
    
    'Remember that Lz4 does *not* expose a compression level.  Instead, it exposes an "acceleration" level,
    ' where higher values mean faster - but worse - compression.  Because of this, the way we report
    ' max/min values is opposite other libraries.
    ElseIf (cmpFormat = cf_Lz4) Then
        GetMaxCompressionLevel = Plugin_lz4.Lz4_GetMinAccelerationLevel()
    ElseIf (cmpFormat = cf_Lz4hc) Then
        GetMaxCompressionLevel = Plugin_lz4.Lz4HC_GetMaxCompressionLevel()
    ElseIf (cmpFormat = cf_Deflate) Then
        'GetMaxCompressionLevel = Plugin_libdeflate.GetMaxCompressionLevel()
    ElseIf (cmpFormat = cf_Gzip) Then
        'GetMaxCompressionLevel = Plugin_libdeflate.GetMaxCompressionLevel()
    Else
        GetMaxCompressionLevel = 0
    End If

End Function

'This function exists purely for debug purposes.  Feel free to remove it if you find it unnecessary.
Public Function GetFormatName(ByVal cmpFormat As PD_CompressionFormat) As String
    
    If (cmpFormat = cf_Zlib) Then
        GetFormatName = "zlib"
    ElseIf (cmpFormat = cf_Zstd) Then
        GetFormatName = "zstd"
    ElseIf (cmpFormat = cf_Lz4) Then
        GetFormatName = "lz4"
    ElseIf (cmpFormat = cf_Lz4hc) Then
        GetFormatName = "lz4_HC"
    ElseIf (cmpFormat = cf_Deflate) Then
        GetFormatName = "deflate"
    ElseIf (cmpFormat = cf_Gzip) Then
        GetFormatName = "gzip"
    Else
        GetFormatName = vbNullString
    End If

End Function

Private Sub InternalErrorMsg(ByVal errMsg As String)
    'PDDebug.LogAction "WARNING! Compression module error: " & errMsg
End Sub
