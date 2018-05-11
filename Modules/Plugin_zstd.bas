Attribute VB_Name = "Plugin_zstd"
'***************************************************************************
'Zstd Compression Library Interface
'Copyright 2016-2018 by Tanner Helland
'Created: 01/December/16
'Last updated: 04/December/16
'Last update: wrap up initial build
'
'Per its documentation (available at https://github.com/facebook/zstd), zstd is...
'
' "...a fast lossless compression algorithm, targeting real-time compression scenarios
'  at zlib-level and better compression ratios."
'
'zstd is BSD-licensed and sponsored by Facebook.  As of Dec 2016, development is very active and performance
' numbers are very favorable compared to zLib.  (3-4x faster at compressing, ~1.5x faster at decompressing
' depending on workload.)  As PD writes a ton of huge files, improved compression performance is a big win
' for us.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'These constants were originally declared in zstd.h
Private Const ZSTD_MIN_CLEVEL As Long = 1
Private Const ZSTD_DEFAULT_CLEVEL As Long = 3

'Zstd supports higher compression levels (e.g. >= 20), but these "ultra-mode" compression levels require
' additional memory during both compression *and* decompression.  This limits its usefulness in a project
' like ours, where we attempt to run even on extremely old, memory-limited PCs.  As such, I've artificially
' limited the maximum level to 19 for our usage.
Private Const ZSTD_MAX_CLEVEL As Long = 19

Private Declare Function ZSTD_versionNumber Lib "libzstd" Alias "_ZSTD_versionNumber@0" () As Long

'Basic compress/decompress functions.  Note that these create their own contexts on every call;
' for reduced memory churn, it's preferable to reuse one compression and decompression context per-session.
Private Declare Function ZSTD_compress Lib "libzstd" Alias "_ZSTD_compress@20" (ByVal ptrToDstBuffer As Long, ByVal dstBufferCapacityInBytes As Long, ByVal constPtrToSrcBuffer As Long, ByVal srcSizeInBytes As Long, ByVal cCompressionLevel As Long) As Long
Private Declare Function ZSTD_decompress Lib "libzstd" Alias "_ZSTD_decompress@16" (ByVal ptrToDstBuffer As Long, ByVal dstBufferCapacityInBytes As Long, ByVal constPtrToSrcBuffer As Long, ByVal srcSizeInBytes As Long) As Long

Private m_CompressionContext As Long, m_DecompressionContext As Long
Private Declare Function ZSTD_createCCtx Lib "libzstd" Alias "_ZSTD_createCCtx@0" () As Long
Private Declare Function ZSTD_freeCCtx Lib "libzstd" Alias "_ZSTD_freeCCtx@4" (ByVal srcCCtx As Long) As Long
Private Declare Function ZSTD_compressCCtx Lib "libzstd" Alias "_ZSTD_compressCCtx@24" (ByVal srcCCtx As Long, ByVal ptrToDstBuffer As Long, ByVal dstBufferCapacityInBytes As Long, ByVal constPtrToSrcBuffer As Long, ByVal srcSizeInBytes As Long, ByVal cCompressionLevel As Long) As Long

Private Declare Function ZSTD_createDCtx Lib "libzstd" Alias "_ZSTD_createDCtx@0" () As Long
Private Declare Function ZSTD_freeDCtx Lib "libzstd" Alias "_ZSTD_freeDCtx@4" (ByVal srcDCtx As Long) As Long
Private Declare Function ZSTD_decompressDCtx Lib "libzstd" Alias "_ZSTD_decompressDCtx@20" (ByVal srcDCtx As Long, ByVal ptrToDstBuffer As Long, ByVal dstBufferCapacityInBytes As Long, ByVal constPtrToSrcBuffer As Long, ByVal srcSizeInBytes As Long) As Long

'These functions are not as self-explanatory as the ones above:
Private Declare Function ZSTD_maxCLevel Lib "libzstd" Alias "_ZSTD_maxCLevel@0" () As Long  'Maximum compression level available
Private Declare Function ZSTD_compressBound Lib "libzstd" Alias "_ZSTD_compressBound@4" (ByVal inputSizeInBytes As Long) As Long 'Maximum compressed size in worst case scenario; use this to size your input array
Private Declare Function ZSTD_isError Lib "libzstd" Alias "_ZSTD_isError@4" (ByVal returnCode As Long) As Long 'Tells you if a function result is an error code or a valid size return
Private Declare Function ZSTD_getErrorName Lib "libzstd" Alias "_ZSTD_getErrorName@4" (ByVal returnCode As Long) As Long 'Returns a pointer to a const char string, with a human-readable string describing the given error code

'If you want, you can ask zstd to tell you how much size is require to decompress a given compression array.  PD doesn't need this
' (as we track compression sizes manually), but it's here if you need it.  Note that automatic calculations like this are generally
' discouraged, as a malicious user can send malformed streams with faulty compression sizes embedded, leading to buffer overflow
' exploits.  Be good, and always manually supply known buffer sizes to external libraries!
'unsigned long long ZSTD_getDecompressedSize(const void* src, size_t srcSize);

'A single zstd handle is maintained for the life of a PD instance; see InitializeZstd and ReleaseZstd, below.
Private m_ZstdHandle As Long

'Maximum compression level that the library currently supports.  This is cached at initialization time.
Private m_ZstdCompressLevelMax As Long

'Initialize zstd.  Do not call this until you have verified zstd's existence (typically via the PluginManager module)
Public Function InitializeZStd(ByRef pathToDLLFolder As String) As Boolean

    'Manually load the DLL from the plugin folder (should be App.Path\Data\Plugins)
    Dim zstdPath As String
    zstdPath = pathToDLLFolder & "libzstd.dll"
    m_ZstdHandle = VBHacks.LoadLib(zstdPath)
    InitializeZStd = (m_ZstdHandle <> 0)
    
    'If we initialized the library successfully, cache some zstd-specific data
    If InitializeZStd Then
        m_ZstdCompressLevelMax = ZSTD_maxCLevel()
        If (m_ZstdCompressLevelMax > ZSTD_MAX_CLEVEL) Then m_ZstdCompressLevelMax = ZSTD_MAX_CLEVEL
        m_CompressionContext = ZSTD_createCCtx()
        m_DecompressionContext = ZSTD_createDCtx()
        PDDebug.LogAction "zstd is ready.  Max compression level supported: " & CStr(m_ZstdCompressLevelMax)
    Else
        PDDebug.LogAction "WARNING!  LoadLibrary failed to load zstd.  Last DLL error: " & Err.LastDllError
        PDDebug.LogAction "(FYI, the attempted path was: " & zstdPath & ")"
    End If
    
End Function

'When PD closes, make sure to release our open zstd handle
Public Sub ReleaseZstd()

    If (m_ZstdHandle <> 0) Then
        
        If (m_CompressionContext <> 0) Then
            ZSTD_freeCCtx m_CompressionContext
            m_CompressionContext = 0
        End If
        
        If (m_DecompressionContext <> 0) Then
            ZSTD_freeDCtx m_DecompressionContext
            m_DecompressionContext = 0
        End If
        
        VBHacks.FreeLib m_ZstdHandle
        m_ZstdHandle = 0
        
    End If
    
End Sub

Public Function GetZstdVersion() As Long
    GetZstdVersion = ZSTD_versionNumber()
End Function

Public Function IsZstdAvailable() As Boolean
    IsZstdAvailable = (m_ZstdHandle <> 0)
End Function

'Determine the maximum possible size required by a compression operation.  The destination buffer should be at least
' this large (and if it's even bigger, that's okay too).
Public Function ZstdGetMaxCompressedSize(ByVal srcSize As Long) As Long
    ZstdGetMaxCompressedSize = ZSTD_compressBound(srcSize)
    If (ZSTD_isError(ZstdGetMaxCompressedSize) <> 0) Then
        InternalError "ZstdGetMaxCompressedSize failed", ZstdGetMaxCompressedSize
        ZstdGetMaxCompressedSize = 0
    End If
End Function

'Compress some arbitrary source pointer + length into a destination array.  Pass the optional "dstArrayIsReady" as TRUE
' (with a matching size descriptor) if you don't want us to auto-size the destination for you.
'
'RETURNS: final size of the compressed data, in bytes.  0 on failure.
'
'IMPORTANT!  The destination array is *not* resized to match the final compressed size.  The caller is responsible
' for this, if they want it.
Public Function ZstdCompressArray(ByRef dstArray() As Byte, ByVal ptrToSrcData As Long, ByVal srcDataSize As Long, Optional ByVal dstArrayIsReady As Boolean = False, Optional ByVal dstArraySizeInBytes As Long = 0, Optional ByVal compressionLevel As Long = -1) As Long
    
    'Validate the incoming compression level parameter
    If (compressionLevel < 1) Then
        compressionLevel = -1
    ElseIf (compressionLevel > m_ZstdCompressLevelMax) Then
        compressionLevel = m_ZstdCompressLevelMax
    End If
    
    'Prep the destination array, as necessary
    If (Not dstArrayIsReady) Or (dstArraySizeInBytes = 0) Then
        dstArraySizeInBytes = ZstdGetMaxCompressedSize(srcDataSize)
        ReDim dstArray(0 To dstArraySizeInBytes - 1) As Byte
    End If
    
    'Perform the compression, and attempt to reuse a compression context if one is available
    Dim finalSize As Long
    If (m_CompressionContext <> 0) Then
        finalSize = ZSTD_compressCCtx(m_CompressionContext, VarPtr(dstArray(0)), dstArraySizeInBytes, ptrToSrcData, srcDataSize, compressionLevel)
    Else
        finalSize = ZSTD_compress(VarPtr(dstArray(0)), dstArraySizeInBytes, ptrToSrcData, srcDataSize, compressionLevel)
    End If
    
    'Check for error returns
    If (ZSTD_isError(finalSize) <> 0) Then
        InternalError "ZSTD_compress failed", finalSize
        finalSize = 0
    End If
    
    ZstdCompressArray = finalSize

End Function

'Compress some arbitrary source buffer to an arbitrary destination buffer.  Caller is responsible for all allocations.
Public Function ZstdCompressNakedPointers(ByVal dstPointer As Long, ByRef dstSizeInBytes As Long, ByVal srcPointer As Long, ByVal srcSizeInBytes As Long, Optional ByVal compressionLevel As Long = -1) As Boolean
    
    'Validate the incoming compression level parameter
    If (compressionLevel < 1) Then
        compressionLevel = -1
    ElseIf (compressionLevel > m_ZstdCompressLevelMax) Then
        compressionLevel = m_ZstdCompressLevelMax
    End If
    
    'Perform the compression
    Dim finalSize As Long
    If (m_CompressionContext <> 0) Then
        finalSize = ZSTD_compressCCtx(m_CompressionContext, dstPointer, dstSizeInBytes, srcPointer, srcSizeInBytes, compressionLevel)
    Else
        finalSize = ZSTD_compress(dstPointer, dstSizeInBytes, srcPointer, srcSizeInBytes, compressionLevel)
    End If
    
    'Check for error returns
    ZstdCompressNakedPointers = (ZSTD_isError(finalSize) = 0)
    
    If ZstdCompressNakedPointers Then
        dstSizeInBytes = finalSize
    Else
        InternalError "ZSTD_compress failed", finalSize
        finalSize = 0
    End If
    
End Function

'Decompress some arbitrary source pointer + length into a destination array.  Pass the optional "dstArrayIsReady" as TRUE
' (with a matching size descriptor) if you don't want us to auto-size the destination for you.
'
'RETURNS: final size of the uncompressed data, in bytes.  0 on failure.
'
'IMPORTANT!  The destination array is *not* resized to match the returned size.  The caller is responsible for this.
Public Function ZstdDecompressArray(ByRef dstArray() As Byte, ByVal ptrToSrcData As Long, ByVal srcDataSize As Long, ByVal knownUncompressedSize As Long, Optional ByVal dstArrayIsReady As Boolean = False) As Long
    
    'Prep the destination array, as necessary
    If (Not dstArrayIsReady) Then
        ReDim dstArray(0 To knownUncompressedSize - 1) As Byte
    End If
    
    'Perform decompression, and attempt to reuse a decompression context if one is available
    Dim finalSize As Long
    If (m_DecompressionContext <> 0) Then
        finalSize = ZSTD_decompressDCtx(m_DecompressionContext, VarPtr(dstArray(0)), knownUncompressedSize, ptrToSrcData, srcDataSize)
    Else
        finalSize = ZSTD_decompress(VarPtr(dstArray(0)), knownUncompressedSize, ptrToSrcData, srcDataSize)
    End If
    
    'Check for error returns
    If (ZSTD_isError(finalSize) <> 0) Then
        PDDebug.LogAction "ZSTD_Decompress failure inputs: " & VarPtr(dstArray(0)) & ", " & knownUncompressedSize & ", " & ptrToSrcData & ", " & srcDataSize
        InternalError "ZSTD_decompress failed", finalSize
        finalSize = 0
    End If
    
    ZstdDecompressArray = finalSize

End Function

Public Function ZstdDecompress_UnsafePtr(ByVal ptrToDstBuffer As Long, ByVal knownUncompressedSize As Long, ByVal ptrToSrcData As Long, ByVal srcDataSize As Long) As Long
    
    'Perform decompression
    Dim finalSize As Long
    If (m_DecompressionContext <> 0) Then
        finalSize = ZSTD_decompressDCtx(m_DecompressionContext, ptrToDstBuffer, knownUncompressedSize, ptrToSrcData, srcDataSize)
    Else
        finalSize = ZSTD_decompress(ptrToDstBuffer, knownUncompressedSize, ptrToSrcData, srcDataSize)
    End If
    
    'Check for error returns
    If (ZSTD_isError(finalSize) <> 0) Then
        PDDebug.LogAction "ZSTD_Decompress failure inputs: " & ptrToDstBuffer & ", " & knownUncompressedSize & ", " & ptrToSrcData & ", " & srcDataSize
        InternalError "ZSTD_decompress failed", finalSize
        finalSize = 0
    End If
    
    ZstdDecompress_UnsafePtr = finalSize

End Function

Public Function Zstd_GetDefaultCompressionLevel() As Long
    Zstd_GetDefaultCompressionLevel = ZSTD_DEFAULT_CLEVEL
End Function

Public Function Zstd_GetMinCompressionLevel() As Long
    Zstd_GetMinCompressionLevel = ZSTD_MIN_CLEVEL
End Function

Public Function Zstd_GetMaxCompressionLevel() As Long
    Zstd_GetMaxCompressionLevel = ZSTD_MAX_CLEVEL
End Function

Private Sub InternalError(ByVal errString As String, Optional ByVal faultyReturnCode As Long = 256)
    
    If (faultyReturnCode <> 256) Then
        
        'Get a char pointer that describes this error
        Dim ptrChar As Long
        ptrChar = ZSTD_getErrorName(faultyReturnCode)
        
        'Convert the char * to a VB string
        Dim errDescription As String
        errDescription = Strings.StringFromCharPtr(ptrChar, False, 255)

        PDDebug.LogAction "zstd returned an error code (" & faultyReturnCode & "): " & errDescription, PDM_External_Lib
    Else
        PDDebug.LogAction "zstd experienced an error: " & errString, PDM_External_Lib
    End If
    
End Sub
