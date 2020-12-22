Attribute VB_Name = "Plugin_zstd"
'***************************************************************************
'Zstd Compression Library Interface
'Copyright 2016-2019 by Tanner Helland
'Created: 01/December/16
'Last updated: 08/March/19
'Last update: switch to callconv-agnostic implementation (so we can use "official" binaries)
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
'This wrapper class uses a shorthand implementation of DispCallFunc originally written by Olaf Schmidt.
' Many thanks to Olaf, whose original version can be found here (link good as of Feb 2019):
' http://www.vbforums.com/showthread.php?781595-VB6-Call-Functions-By-Pointer-(Universall-DLL-Calls)&p=4795471&viewfull=1#post4795471
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
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

'As recommended by the manual, PD reuses de/compression contexts for the lifetime of the project;
' this reduces the need for repeat allocations on every de/compression request.
Private m_CompressionContext As Long, m_DecompressionContext As Long

'The following functions are used in this module, but instead of being called directly, calls are routed
' through DispCallFunc (which allows us to use the prebuilt release DLLs provided by the library authors):
'Private Declare Function ZSTD_compress Lib "libzstd" Alias "_ZSTD_compress@20" (ByVal ptrToDstBuffer As Long, ByVal dstBufferCapacityInBytes As Long, ByVal constPtrToSrcBuffer As Long, ByVal srcSizeInBytes As Long, ByVal cCompressionLevel As Long) As Long
'Private Declare Function ZSTD_compressBound Lib "libzstd" Alias "_ZSTD_compressBound@4" (ByVal inputSizeInBytes As Long) As Long 'Maximum compressed size in worst case scenario; use this to size your input array
'Private Declare Function ZSTD_compressCCtx Lib "libzstd" Alias "_ZSTD_compressCCtx@24" (ByVal srcCCtx As Long, ByVal ptrToDstBuffer As Long, ByVal dstBufferCapacityInBytes As Long, ByVal constPtrToSrcBuffer As Long, ByVal srcSizeInBytes As Long, ByVal cCompressionLevel As Long) As Long
'Private Declare Function ZSTD_createCCtx Lib "libzstd" Alias "_ZSTD_createCCtx@0" () As Long
'Private Declare Function ZSTD_createDCtx Lib "libzstd" Alias "_ZSTD_createDCtx@0" () As Long
'Private Declare Function ZSTD_decompress Lib "libzstd" Alias "_ZSTD_decompress@16" (ByVal ptrToDstBuffer As Long, ByVal dstBufferCapacityInBytes As Long, ByVal constPtrToSrcBuffer As Long, ByVal srcSizeInBytes As Long) As Long
'Private Declare Function ZSTD_decompressDCtx Lib "libzstd" Alias "_ZSTD_decompressDCtx@20" (ByVal srcDCtx As Long, ByVal ptrToDstBuffer As Long, ByVal dstBufferCapacityInBytes As Long, ByVal constPtrToSrcBuffer As Long, ByVal srcSizeInBytes As Long) As Long
'Private Declare Function ZSTD_freeCCtx Lib "libzstd" Alias "_ZSTD_freeCCtx@4" (ByVal srcCCtx As Long) As Long
'Private Declare Function ZSTD_freeDCtx Lib "libzstd" Alias "_ZSTD_freeDCtx@4" (ByVal srcDCtx As Long) As Long
'Private Declare Function ZSTD_getErrorName Lib "libzstd" Alias "_ZSTD_getErrorName@4" (ByVal returnCode As Long) As Long 'Returns a pointer to a const char string, with a human-readable string describing the given error code
'Private Declare Function ZSTD_isError Lib "libzstd" Alias "_ZSTD_isError@4" (ByVal returnCode As Long) As Long 'Tells you if a function result is an error code or a valid size return
'Private Declare Function ZSTD_maxCLevel Lib "libzstd" Alias "_ZSTD_maxCLevel@0" () As Long  'Maximum compression level available
'Private Declare Function ZSTD_versionNumber Lib "libzstd" Alias "_ZSTD_versionNumber@0" () As Long

'If you want, you can ask zstd to tell you how much size is require to decompress a given compression array.
' PD doesn't need this (as we track compression sizes manually), but it's here if you need it.  Note that
' automatic calculations like this are generally discouraged, as a malicious user can send malformed streams
' with faulty compression sizes embedded, leading to buffer overflow exploits.  Be good, and always manually
' supply known buffer sizes to external libraries!
'unsigned long long ZSTD_getDecompressedSize(const void* src, size_t srcSize);

'A single zstd handle is maintained for the life of a PD instance; see InitializeZstd and ReleaseZstd, below.
Private m_ZstdHandle As Long

'Maximum compression level that the library currently supports.  This is cached at initialization time.
Private m_ZstdCompressLevelMax As Long

'zstd has very specific compiler needs in order to produce maximum perf code, so rather than
' recompile myself, I've just grabbed the prebuilt Windows binaries and wrapped 'em using DispCallFunc
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef retVAR As Variant) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

'At load-time, we cache a number of proc addresses (required for passing through DispCallFunc).
' This saves us a little time vs calling GetProcAddress on each call.
Private Enum Zstd_ProcAddress
    ZSTD_versionNumber
    ZSTD_compress
    ZSTD_decompress
    ZSTD_createCCtx
    ZSTD_freeCCtx
    ZSTD_compressCCtx
    ZSTD_createDCtx
    ZSTD_freeDCtx
    ZSTD_decompressDCtx
    ZSTD_maxCLevel
    ZSTD_compressBound
    ZSTD_isError
    ZSTD_getErrorName
    [last_address]
End Enum

Private m_ProcAddresses() As Long

'Rather than allocate new memory on each DispCallFunc invoke, just reuse a set of temp arrays declared
' to the maximum relevant size (see InitializeEngine, below).
Private Const MAX_PARAM_COUNT As Long = 8
Private m_vType() As Integer, m_vPtr() As Long

'If for some reason the local copy of zstd is an (old) stdcall build, this flag will be set to TRUE by
' the library loader.
Private m_UseStdCallFallback As Boolean

'Initialize zstd.  Do not call this until you have verified zstd's existence (typically via the PluginManager module)
Public Function InitializeZStd(ByRef pathToDLLFolder As String) As Boolean

    'Manually load the DLL from the plugin folder (should be App.Path\Data\Plugins)
    Dim zstdPath As String
    zstdPath = pathToDLLFolder & "libzstd.dll"
    m_ZstdHandle = VBHacks.LoadLib(zstdPath)
    InitializeZStd = (m_ZstdHandle <> 0)
    
    'If we initialized the library successfully, cache some zstd-specific data
    If InitializeZStd Then
    
        'Pre-load all relevant proc addresses
        ReDim m_ProcAddresses(0 To [last_address] - 1) As Long
        m_ProcAddresses(ZSTD_compress) = GetProcAddressHelper(m_ZstdHandle, "ZSTD_compress")
        m_ProcAddresses(ZSTD_compressBound) = GetProcAddressHelper(m_ZstdHandle, "ZSTD_compressBound")
        m_ProcAddresses(ZSTD_compressCCtx) = GetProcAddressHelper(m_ZstdHandle, "ZSTD_compressCCtx")
        m_ProcAddresses(ZSTD_createCCtx) = GetProcAddressHelper(m_ZstdHandle, "ZSTD_createCCtx")
        m_ProcAddresses(ZSTD_createDCtx) = GetProcAddressHelper(m_ZstdHandle, "ZSTD_createDCtx")
        m_ProcAddresses(ZSTD_decompress) = GetProcAddressHelper(m_ZstdHandle, "ZSTD_decompress")
        m_ProcAddresses(ZSTD_decompressDCtx) = GetProcAddressHelper(m_ZstdHandle, "ZSTD_decompressDCtx")
        m_ProcAddresses(ZSTD_freeCCtx) = GetProcAddressHelper(m_ZstdHandle, "ZSTD_freeCCtx")
        m_ProcAddresses(ZSTD_freeDCtx) = GetProcAddressHelper(m_ZstdHandle, "ZSTD_freeDCtx")
        m_ProcAddresses(ZSTD_getErrorName) = GetProcAddressHelper(m_ZstdHandle, "ZSTD_getErrorName")
        m_ProcAddresses(ZSTD_isError) = GetProcAddressHelper(m_ZstdHandle, "ZSTD_isError")
        m_ProcAddresses(ZSTD_maxCLevel) = GetProcAddressHelper(m_ZstdHandle, "ZSTD_maxCLevel")
        m_ProcAddresses(ZSTD_versionNumber) = GetProcAddressHelper(m_ZstdHandle, "ZSTD_versionNumber")
        
        'Initialize all module-level arrays
        ReDim m_vType(0 To MAX_PARAM_COUNT - 1) As Integer
        ReDim m_vPtr(0 To MAX_PARAM_COUNT - 1) As Long
        
        'Retrieve some zstd-specific data.  Note that we manually cap the compression level to avoid
        ' "ultra" settings (levels >= 20) because they require extremely large amounts of memory.
        m_ZstdCompressLevelMax = CallCDeclW(ZSTD_maxCLevel, vbLong)
        If (m_ZstdCompressLevelMax > ZSTD_MAX_CLEVEL) Then m_ZstdCompressLevelMax = ZSTD_MAX_CLEVEL
        m_CompressionContext = CallCDeclW(ZSTD_createCCtx, vbLong)
        m_DecompressionContext = CallCDeclW(ZSTD_createDCtx, vbLong)
        
        'PDDebug.LogAction "zstd is ready.  Max compression level supported: " & CStr(m_ZstdCompressLevelMax)
        
    Else
        'PDDebug.LogAction "WARNING!  LoadLibrary failed to load zstd.  Last DLL error: " & Err.LastDllError
        'PDDebug.LogAction "(FYI, the attempted path was: " & zstdPath & ")"
    End If
    
End Function

Private Function GetProcAddressHelper(ByVal zstdHandle As Long, ByRef srcFuncName As String) As Long

    'Attempt to load the function as-is
    GetProcAddressHelper = GetProcAddress(zstdHandle, srcFuncName)
    If (GetProcAddressHelper <> 0) Then
        m_UseStdCallFallback = False
    
    'If the load function failed, the most likely explanation is that an stdcall version of the library
    ' (with name-mangling) exists.  Try to find a new address w/name-mangling.
    Else
        
        Dim i As Long, tmpFuncName As String
        For i = 0 To 64 'Arbitrary upper limit!
            tmpFuncName = "_" & srcFuncName & "@" & Trim$(Str$(i))
            GetProcAddressHelper = GetProcAddress(zstdHandle, tmpFuncName)
            If (GetProcAddressHelper <> 0) Then
                m_UseStdCallFallback = True
                Exit Function
            End If
        Next i
        
    End If

End Function

'When PD closes, make sure to release our open zstd handle
Public Sub ReleaseZstd()

    If (m_ZstdHandle <> 0) Then
        
        If (m_CompressionContext <> 0) Then
            CallCDeclW ZSTD_freeCCtx, vbEmpty, m_CompressionContext
            m_CompressionContext = 0
        End If
        
        If (m_DecompressionContext <> 0) Then
            CallCDeclW ZSTD_freeDCtx, vbEmpty, m_DecompressionContext
            m_DecompressionContext = 0
        End If
        
        VBHacks.FreeLib m_ZstdHandle
        m_ZstdHandle = 0
        
    End If
    
End Sub

Public Function GetZstdVersion() As Long
    GetZstdVersion = CallCDeclW(ZSTD_versionNumber, vbLong)
End Function

Public Function IsZstdAvailable() As Boolean
    IsZstdAvailable = (m_ZstdHandle <> 0)
End Function

'Determine the maximum possible size required by a compression operation.  The destination buffer should be at least
' this large (and if it's even bigger, that's okay too).
Public Function ZstdGetMaxCompressedSize(ByVal srcSize As Long) As Long
    ZstdGetMaxCompressedSize = CallCDeclW(ZSTD_compressBound, vbLong, srcSize)
    If (CallCDeclW(ZSTD_isError, vbLong, ZstdGetMaxCompressedSize) <> 0) Then
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
        finalSize = CallCDeclW(ZSTD_compressCCtx, vbLong, m_CompressionContext, VarPtr(dstArray(0)), dstArraySizeInBytes, ptrToSrcData, srcDataSize, compressionLevel)
    Else
        finalSize = CallCDeclW(ZSTD_compress, vbLong, VarPtr(dstArray(0)), dstArraySizeInBytes, ptrToSrcData, srcDataSize, compressionLevel)
    End If
    
    'Check for error returns
    If (CallCDeclW(ZSTD_isError, vbLong, finalSize) <> 0) Then
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
        finalSize = CallCDeclW(ZSTD_compressCCtx, vbLong, m_CompressionContext, dstPointer, dstSizeInBytes, srcPointer, srcSizeInBytes, compressionLevel)
    Else
        finalSize = CallCDeclW(ZSTD_compress, vbLong, dstPointer, dstSizeInBytes, srcPointer, srcSizeInBytes, compressionLevel)
    End If
    
    'Check for error returns
    ZstdCompressNakedPointers = (CallCDeclW(ZSTD_isError, vbLong, finalSize) = 0)
    
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
        finalSize = CallCDeclW(ZSTD_decompressDCtx, vbLong, m_DecompressionContext, VarPtr(dstArray(0)), knownUncompressedSize, ptrToSrcData, srcDataSize)
    Else
        finalSize = CallCDeclW(ZSTD_decompress, vbLong, VarPtr(dstArray(0)), knownUncompressedSize, ptrToSrcData, srcDataSize)
    End If
    
    'Check for error returns
    If (CallCDeclW(ZSTD_isError, vbLong, finalSize) <> 0) Then
        'PDDebug.LogAction "ZSTD_Decompress failure inputs: " & VarPtr(dstArray(0)) & ", " & knownUncompressedSize & ", " & ptrToSrcData & ", " & srcDataSize
        InternalError "ZSTD_decompress failed", finalSize
        finalSize = 0
    End If
    
    ZstdDecompressArray = finalSize

End Function

Public Function ZstdDecompress_UnsafePtr(ByVal ptrToDstBuffer As Long, ByVal knownUncompressedSize As Long, ByVal ptrToSrcData As Long, ByVal srcDataSize As Long) As Long
    
    'Perform decompression
    Dim finalSize As Long
    If (m_DecompressionContext <> 0) Then
        finalSize = CallCDeclW(ZSTD_decompressDCtx, vbLong, m_DecompressionContext, ptrToDstBuffer, knownUncompressedSize, ptrToSrcData, srcDataSize)
    Else
        finalSize = CallCDeclW(ZSTD_decompress, vbLong, ptrToDstBuffer, knownUncompressedSize, ptrToSrcData, srcDataSize)
    End If
    
    'Check for error returns
    If (CallCDeclW(ZSTD_isError, vbLong, finalSize) <> 0) Then
        'PDDebug.LogAction "ZSTD_Decompress failure inputs: " & ptrToDstBuffer & ", " & knownUncompressedSize & ", " & ptrToSrcData & ", " & srcDataSize
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

'DispCallFunc wrapper originally by Olaf Schmidt, with a few minor modifications; see the top of this class
' for a link to his original, unmodified version
Private Function CallCDeclW(ByVal lProc As Zstd_ProcAddress, ByVal fRetType As VbVarType, ParamArray pa() As Variant) As Variant

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
    
    Const CC_CDECL As Long = 1, CC_STDCALL = 4
    If m_UseStdCallFallback Then
        hResult = DispCallFunc(0, m_ProcAddresses(lProc), CC_STDCALL, fRetType, i, m_vType(0), m_vPtr(0), CallCDeclW)
    Else
        hResult = DispCallFunc(0, m_ProcAddresses(lProc), CC_CDECL, fRetType, i, m_vType(0), m_vPtr(0), CallCDeclW)
    End If
    
    If hResult Then
        'FormPatch.TextOut "hResult failed"
        Err.Raise hResult
    End If
    
End Function

Private Sub InternalError(ByVal errString As String, Optional ByVal faultyReturnCode As Long = 256)
    
    If (faultyReturnCode <> 256) Then
        
        'Get a char pointer that describes this error
        Dim ptrChar As Long
        ptrChar = CallCDeclW(ZSTD_getErrorName, vbLong, faultyReturnCode)
        
        'Convert the char * to a VB string
        Dim errDescription As String
        errDescription = Strings.StringFromCharPtr(ptrChar, False, 255)

        'PDDebug.LogAction "zstd returned an error code (" & faultyReturnCode & "): " & errDescription, PDM_External_Lib
    Else
        'PDDebug.LogAction "zstd experienced an error: " & errString, PDM_External_Lib
    End If
    
End Sub
