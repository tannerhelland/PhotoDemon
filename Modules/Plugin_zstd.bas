Attribute VB_Name = "Plugin_zstd"
'***************************************************************************
'Zstd Compression Library Interface
'Copyright 2016-2017 by Tanner Helland
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
Private Const ZSTD_DEFAULT_CLEVEL As Long = 1
Private Const ZSTD_MAX_CLEVEL As Long = 22

Private Declare Function ZSTD_VersionNumber Lib "libzstd" Alias "_ZSTD_versionNumber@0" () As Long
Private Declare Function ZSTD_compress Lib "libzstd" Alias "_ZSTD_compress@20" (ByVal ptrToDstBuffer As Long, ByVal dstBufferCapacityInBytes As Long, ByVal constPtrToSrcBuffer As Long, ByVal srcSizeInBytes As Long, ByVal cCompressionLevel As Long) As Long
Private Declare Function ZSTD_decompress Lib "libzstd" Alias "_ZSTD_decompress@16" (ByVal ptrToDstBuffer As Long, ByVal dstBufferCapacityInBytes As Long, ByVal constPtrToSrcBuffer As Long, ByVal srcSizeInBytes As Long) As Long

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
    m_ZstdHandle = LoadLibrary(StrPtr(zstdPath))
    InitializeZStd = CBool(m_ZstdHandle <> 0)
    
    'If we initialized the library successfully, cache some zstd-specific data
    If InitializeZStd Then
        m_ZstdCompressLevelMax = ZSTD_maxCLevel()
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "zstd is ready.  Max compression level supported: " & CStr(m_ZstdCompressLevelMax)
        #End If
    End If
    
    #If DEBUGMODE = 1 Then
        If (Not InitializeZStd) Then
            pdDebug.LogAction "WARNING!  LoadLibrary failed to load zstd.  Last DLL error: " & Err.LastDllError
            pdDebug.LogAction "(FYI, the attempted path was: " & zstdPath & ")"
        End If
    #End If
    
End Function

'When PD closes, make sure to release our open zstd handle
Public Sub ReleaseZstd()
    If (m_ZstdHandle <> 0) Then
        FreeLibrary m_ZstdHandle
        m_ZstdHandle = 0
    End If
End Sub

Public Function GetZstdVersion() As Long
    GetZstdVersion = ZSTD_VersionNumber()
End Function

Public Function IsZstdAvailable() As Boolean
    IsZstdAvailable = CBool(m_ZstdHandle <> 0)
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
    
    'Perform the compression
    Dim finalSize As Long
    finalSize = ZSTD_compress(VarPtr(dstArray(0)), dstArraySizeInBytes, ptrToSrcData, srcDataSize, compressionLevel)
    
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
    finalSize = ZSTD_compress(dstPointer, dstSizeInBytes, srcPointer, srcSizeInBytes, compressionLevel)
    
    'Check for error returns
    ZstdCompressNakedPointers = CBool(ZSTD_isError(finalSize) = 0)
    
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
    
    'Perform decompression
    Dim finalSize As Long
    finalSize = ZSTD_decompress(VarPtr(dstArray(0)), knownUncompressedSize, ptrToSrcData, srcDataSize)
    
    'Check for error returns
    If (ZSTD_isError(finalSize) <> 0) Then
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "ZSTD_Decompress failure inputs: " & VarPtr(dstArray(0)) & ", " & knownUncompressedSize & ", " & ptrToSrcData & ", " & srcDataSize
        #End If
        InternalError "ZSTD_decompress failed", finalSize
        finalSize = 0
    End If
    
    ZstdDecompressArray = finalSize

End Function

Public Function ZstdDecompress_UnsafePtr(ByVal ptrToDstBuffer As Long, ByVal knownUncompressedSize As Long, ByVal ptrToSrcData As Long, ByVal srcDataSize As Long) As Long
    
    'Perform decompression
    Dim finalSize As Long
    finalSize = ZSTD_decompress(ptrToDstBuffer, knownUncompressedSize, ptrToSrcData, srcDataSize)
    
    'Check for error returns
    If (ZSTD_isError(finalSize) <> 0) Then
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "ZSTD_Decompress failure inputs: " & ptrToDstBuffer & ", " & knownUncompressedSize & ", " & ptrToSrcData & ", " & srcDataSize
        #End If
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
    #If DEBUGMODE = 1 Then
        If (faultyReturnCode <> 256) Then
            
            'Get a char pointer that describes this error
            Dim ptrChar As Long
            ptrChar = ZSTD_getErrorName(faultyReturnCode)
            
            'Convert the char * to a VB string
            Dim errDescription As String
            
            Dim cUnicode As pdUnicode
            Set cUnicode = New pdUnicode
            errDescription = cUnicode.ConvertCharPointerToVBString(ptrChar, False, 255)
    
            pdDebug.LogAction "zstd returned an error code (" & faultyReturnCode & "): " & errDescription, PDM_EXTERNAL_LIB
        Else
            pdDebug.LogAction "zstd experienced an error: " & errString, PDM_EXTERNAL_LIB
        End If
    #End If
End Sub
