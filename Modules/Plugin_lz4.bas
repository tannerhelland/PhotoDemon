Attribute VB_Name = "Plugin_lz4"
'***************************************************************************
'Lz4 Compression Library Interface
'Copyright 2016-2026 by Tanner Helland
'Created: 04/December/16
'Last updated: 08/March/19
'Last update: switch to callconv-agnostic implementation (so we can use "official" binaries)
'
'Per its documentation (available at https://github.com/lz4/lz4), lz4 is...
'
' "...a lossless compression algorithm, providing compression speed at 400 MB/s per core, scalable with
'  multi-cores CPU. It features an extremely fast decoder, with speed in multiple GB/s per core, typically
'  reaching RAM speed limits on multi-core systems."
'
'lz4 is BSD-licensed and written by Yann Collet, the same genius behind the zstd compression library.  As of
' Dec 2016, development is very active and performance numbers rank among the best available for open-source
' compression libraries.  As PD writes a ton of huge files, improved compression performance is a big win
' for us, particularly on old systems with 5400 RPM HDDs.
'
'lz4-hc support is also provided.  lz4-hc is a high-compression variant of lz4.  It is much slower
' (6-10x depending on workload), but provides compression levels close to zlib.  Decompression speed is
' identical to regular lz4, so it is a good fit for things like run-time resources, where you have ample
' time available during compression stages, but you still want decompression to be as fast as possible.
'
'As of v7.0, most internal PD temp files and caches are written using Lz4, so this library sees heavy usage
' during a typical session.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This constant was originally declared in lz4.c.
' Note that lz4 does *not* support variable compression levels.
' Instead, it supports variable *acceleration* levels.
' The difference is that bigger values = worse compression.
Private Const LZ4_MIN_ALEVEL As Long = 1
Private Const LZ4_DEFAULT_ALEVEL As Long = 1

'This value is not declared by the lz4 library, and technically, there is no maximum value.
' Compression just approaches 0% as you increase the level.  I provide a "magic number" cap
' simply so it supports the same default/min/max functions as other compression libraries in PD.
Private Const LZ4_MAX_ALEVEL As Long = 500

'These constants were originally declared in lz4_hc.h
Private Const LZ4HC_MIN_CLEVEL As Long = 3
Private Const LZ4HC_DEFAULT_CLEVEL As Long = 9
Private Const LZ4HC_MAX_CLEVEL As Long = 12

'The following functions are used in this module, but instead of being called directly, calls are routed
' through DispCallFunc (which allows us to use the prebuilt release DLLs provided by the library authors):
'Private Declare Function LZ4_versionNumber Lib "liblz4" Alias "_LZ4_versionNumber@0" () As Long
'Private Declare Function LZ4_compress_fast Lib "liblz4" Alias "_LZ4_compress_fast@20" (ByVal constPtrToSrcBuffer As Long, ByVal ptrToDstBuffer As Long, ByVal srcSizeInBytes As Long, ByVal dstBufferCapacityInBytes As Long, ByVal cAccelerationLevel As Long) As Long
'Private Declare Function LZ4_compress_HC Lib "liblz4" Alias "_LZ4_compress_HC@20" (ByVal constPtrToSrcBuffer As Long, ByVal ptrToDstBuffer As Long, ByVal srcSizeInBytes As Long, ByVal dstBufferCapacityInBytes As Long, ByVal cCompressionLevel As Long) As Long
'Private Declare Function LZ4_decompress_safe Lib "liblz4" Alias "_LZ4_decompress_safe@16" (ByVal constPtrToSrcBuffer As Long, ByVal ptrToDstBuffer As Long, ByVal srcSizeInBytes As Long, ByVal dstBufferCapacityInBytes As Long) As Long
'Private Declare Function LZ4_compressBound Lib "liblz4" Alias "_LZ4_compressBound@4" (ByVal inputSizeInBytes As Long) As Long 'Maximum compressed size in worst case scenario; use this to size your input array

'A single lz4 handle is maintained for the life of a PD instance; see InitializeLz4 and ReleaseLz4, below.
Private m_Lz4Handle As Long

'lz4 has very specific compiler needs in order to produce maximum perf code, so rather than
' recompile myself, I've just grabbed the prebuilt Windows binaries and wrapped 'em using DispCallFunc
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As Long, ByVal offsetinVft As Long, ByVal CallConv As Long, ByVal retTYP As Integer, ByVal paCNT As Long, ByRef paTypes As Integer, ByRef paValues As Long, ByRef retVAR As Variant) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

'At load-time, we cache a number of proc addresses (required for passing through DispCallFunc).
' This saves us a little time vs calling GetProcAddress on each call.
Private Enum LZ4_ProcAddress
    LZ4_versionNumber
    LZ4_compress_fast
    LZ4_compress_HC
    LZ4_decompress_safe
    LZ4_compressBound
    [last_address]
End Enum

Private m_ProcAddresses() As Long

'Rather than allocate new memory on each DispCallFunc invoke, just reuse a set of temp arrays declared
' to the maximum relevant size (see InitializeEngine, below).
Private Const MAX_PARAM_COUNT As Long = 8
Private m_vType() As Integer, m_vPtr() As Long

'Initialize lz4.  Do not call this until you have verified its existence (typically via the PluginManager module)
Public Function InitializeLz4(ByRef pathToDLLFolder As String) As Boolean

    'Manually load the DLL from the plugin folder (should be App.Path\Data\Plugins)
    Dim lz4Path As String
    lz4Path = pathToDLLFolder & "liblz4.dll"
    m_Lz4Handle = VBHacks.LoadLib(lz4Path)
    InitializeLz4 = (m_Lz4Handle <> 0)
    
    'If we initialized the library successfully, cache some lz4-specific data
    If InitializeLz4 Then
    
        'Pre-load all relevant proc addresses
        ReDim m_ProcAddresses(0 To [last_address] - 1) As Long
        m_ProcAddresses(LZ4_versionNumber) = GetProcAddress(m_Lz4Handle, "LZ4_versionNumber")
        m_ProcAddresses(LZ4_compress_fast) = GetProcAddress(m_Lz4Handle, "LZ4_compress_fast")
        m_ProcAddresses(LZ4_compress_HC) = GetProcAddress(m_Lz4Handle, "LZ4_compress_HC")
        m_ProcAddresses(LZ4_decompress_safe) = GetProcAddress(m_Lz4Handle, "LZ4_decompress_safe")
        m_ProcAddresses(LZ4_compressBound) = GetProcAddress(m_Lz4Handle, "LZ4_compressBound")
        
        'Initialize all module-level arrays
        ReDim m_vType(0 To MAX_PARAM_COUNT - 1) As Integer
        ReDim m_vPtr(0 To MAX_PARAM_COUNT - 1) As Long
        
        PDDebug.LogAction "lz4 and lz4hc compression engines are ready."
        
    Else
        PDDebug.LogAction "WARNING!  LoadLibrary failed to load lz4.  Last DLL error: " & Err.LastDllError
        PDDebug.LogAction "(FYI, the attempted path was: " & lz4Path & ")"
    End If
    
End Function

'When PD closes, make sure to release our open Lz4 handle
Public Sub ReleaseLz4()
    If (m_Lz4Handle <> 0) Then
        VBHacks.FreeLib m_Lz4Handle
        m_Lz4Handle = 0
    End If
End Sub

Public Function GetLz4Version() As String
    If (m_Lz4Handle <> 0) Then
        Dim ptrVersion As Long
        ptrVersion = CallCDeclW(LZ4_versionNumber, vbLong)
        GetLz4Version = ptrVersion
    End If
End Function

Public Function IsLz4Available() As Boolean
    IsLz4Available = (m_Lz4Handle <> 0)
End Function

'Determine the maximum possible size required by a compression operation.  The destination buffer should be at least
' this large (and if it's even bigger, that's okay too).
Public Function Lz4GetMaxCompressedSize(ByVal srcSize As Long) As Long
    Lz4GetMaxCompressedSize = CallCDeclW(LZ4_compressBound, vbLong, srcSize)
End Function

'Compress some arbitrary source pointer + length into a destination array.  Pass the optional "dstArrayIsReady" as TRUE
' (with a matching size descriptor) if you don't want us to auto-create the destination buffer for you.
'
'RETURNS: final size of the compressed data, in bytes.  0 on failure.
'
'IMPORTANT!  The destination array is *not* trimmed to match the final compressed size.  The caller is responsible
' for this, if they want it.
Public Function Lz4CompressArray(ByRef dstArray() As Byte, ByVal ptrToSrcData As Long, ByVal srcDataSize As Long, Optional ByVal dstArrayIsReady As Boolean = False, Optional ByVal dstArraySizeInBytes As Long = 0, Optional ByVal compressionAcceleration As Long = -1) As Long
    
    'Failsafe only
    If (m_Lz4Handle = 0) Then Exit Function
    
    'Normally, we would want to validate the incoming compression acceleration parameter, but lz4 automatically handles
    ' negative numbers and sets them to the acceleraton default (currently 1).  Similarly, there's no upper maximum for
    ' acceleration, as far as I can see.  It just asymptotically approaches the speed of a raw memory copy as request
    ' larger accelerations...
    
    'Prep the destination array, as necessary
    If (Not dstArrayIsReady) Or (dstArraySizeInBytes = 0) Then
        dstArraySizeInBytes = Lz4GetMaxCompressedSize(srcDataSize)
        ReDim dstArray(0 To dstArraySizeInBytes - 1) As Byte
    End If
    
    'Perform the compression
    Dim finalSize As Long
    finalSize = CallCDeclW(LZ4_compress_fast, vbLong, ptrToSrcData, VarPtr(dstArray(0)), srcDataSize, dstArraySizeInBytes, compressionAcceleration)
    
    Lz4CompressArray = finalSize

End Function

'Compress some arbitrary source buffer to an arbitrary destination buffer.  Caller is responsible for all allocations.
' Returns: success/failure, and size of the written data in dstSizeInBytes (passed ByRef).
Public Function Lz4CompressNakedPointers(ByVal dstPointer As Long, ByRef dstSizeInBytes As Long, ByVal srcPointer As Long, ByVal srcSizeInBytes As Long, Optional ByVal compressionAcceleration As Long = -1) As Boolean
    
    'Failsafe only
    If (m_Lz4Handle = 0) Then Exit Function
    
    Dim finalSize As Long
    finalSize = CallCDeclW(LZ4_compress_fast, vbLong, srcPointer, dstPointer, srcSizeInBytes, dstSizeInBytes, compressionAcceleration)
    
    'Check for error returns
    Lz4CompressNakedPointers = (finalSize <> 0)
    If Lz4CompressNakedPointers Then
        dstSizeInBytes = finalSize
    Else
        InternalError "lz4_compress failed", finalSize
        finalSize = 0
    End If
    
End Function

'High-compression variant.  Note that compression level has different meaning here - higher values result in SLOWER but BETTER compression
' (vs normal LZ4, where higher values result in FASTER but WORSE compression).
Public Function Lz4HCCompressArray(ByRef dstArray() As Byte, ByVal ptrToSrcData As Long, ByVal srcDataSize As Long, Optional ByVal dstArrayIsReady As Boolean = False, Optional ByVal dstArraySizeInBytes As Long = 0, Optional ByVal compressionLevel As Long = -1) As Long
    
    'Failsafe only
    If (m_Lz4Handle = 0) Then Exit Function
    
    'LZ4_HC provides its own validation of compression levels, but for the record...
    ' 4 is the recommended minimum (though levels as low as 1 are supported, but compression will be poor)
    ' 9 is the default, and any value < 1 resolves to this
    ' 16 is the current maximum level
    
    'Prep the destination array, as necessary
    If (Not dstArrayIsReady) Or (dstArraySizeInBytes = 0) Then
        dstArraySizeInBytes = Lz4GetMaxCompressedSize(srcDataSize)
        ReDim dstArray(0 To dstArraySizeInBytes - 1) As Byte
    End If
    
    'Perform the compression
    Dim finalSize As Long
    finalSize = CallCDeclW(LZ4_compress_HC, vbLong, ptrToSrcData, VarPtr(dstArray(0)), srcDataSize, dstArraySizeInBytes, compressionLevel)
    
    Lz4HCCompressArray = finalSize

End Function

'High-compression variant of the normal LZ4 compression function.
Public Function Lz4HCCompressNakedPointers(ByVal dstPointer As Long, ByRef dstSizeInBytes As Long, ByVal srcPointer As Long, ByVal srcSizeInBytes As Long, Optional ByVal compressionLevel As Long = -1) As Boolean
    
    'Failsafe only
    If (m_Lz4Handle = 0) Then Exit Function
    
    Dim finalSize As Long
    finalSize = CallCDeclW(LZ4_compress_HC, vbLong, srcPointer, dstPointer, srcSizeInBytes, dstSizeInBytes, compressionLevel)
    
    'Check for error returns
    Lz4HCCompressNakedPointers = (finalSize <> 0)
    
    If Lz4HCCompressNakedPointers Then
        dstSizeInBytes = finalSize
    Else
        InternalError "lz4_compress_HC failed", finalSize
        finalSize = 0
    End If
    
End Function

'Decompress some arbitrary source pointer + length into a destination array.  Pass the optional "dstArrayIsReady" as TRUE
' (with a matching size descriptor) if you don't want us to auto-size the destination for you.
'
'RETURNS: final size of the uncompressed data, in bytes.  0 on failure.
'
'IMPORTANT!  The destination array is *not* resized to match the returned size.  The caller is responsible for this.
Public Function Lz4DecompressArray(ByRef dstArray() As Byte, ByVal ptrToSrcData As Long, ByVal srcDataSize As Long, ByVal knownUncompressedSize As Long, Optional ByVal dstArrayIsReady As Boolean = False) As Long
    
    'Failsafe only
    If (m_Lz4Handle = 0) Then Exit Function
    
    'Prep the destination array, as necessary
    If (Not dstArrayIsReady) Then
        ReDim dstArray(0 To knownUncompressedSize - 1) As Byte
    End If
    
    'Perform decompression
    Dim finalSize As Long
    finalSize = CallCDeclW(LZ4_decompress_safe, vbLong, ptrToSrcData, VarPtr(dstArray(0)), srcDataSize, knownUncompressedSize)
    
    'Check for error returns
    If (finalSize <= 0) Then
        PDDebug.LogAction "lz4_decompress_safe failure inputs: " & VarPtr(dstArray(0)) & ", " & knownUncompressedSize & ", " & ptrToSrcData & ", " & srcDataSize
        InternalError "lz4_decompress failed", finalSize
        finalSize = 0
    End If
    
    Lz4DecompressArray = finalSize

End Function

Public Function Lz4Decompress_UnsafePtr(ByVal ptrToDstBuffer As Long, ByVal knownUncompressedSize As Long, ByVal ptrToSrcData As Long, ByVal srcDataSize As Long) As Long
    
    'Failsafe only
    If (m_Lz4Handle = 0) Then Exit Function
    
    'Perform decompression
    Dim finalSize As Long
    finalSize = CallCDeclW(LZ4_decompress_safe, vbLong, ptrToSrcData, ptrToDstBuffer, srcDataSize, knownUncompressedSize)
    
    'Check for error returns
    If (finalSize <= 0) Then
        PDDebug.LogAction "lz4_decompress_safe failure inputs: " & ptrToDstBuffer & ", " & knownUncompressedSize & ", " & ptrToSrcData & ", " & srcDataSize
        InternalError "lz4_decompress failed", finalSize
        finalSize = 0
    End If
    
    Lz4Decompress_UnsafePtr = finalSize
    
End Function

Public Function Lz4_GetDefaultAccelerationLevel() As Long
    Lz4_GetDefaultAccelerationLevel = LZ4_DEFAULT_ALEVEL
End Function

Public Function Lz4_GetMinAccelerationLevel() As Long
    Lz4_GetMinAccelerationLevel = LZ4_MIN_ALEVEL
End Function

Public Function Lz4_GetMaxAccelerationLevel() As Long
    Lz4_GetMaxAccelerationLevel = LZ4_MAX_ALEVEL
End Function

Public Function Lz4HC_GetDefaultCompressionLevel() As Long
    Lz4HC_GetDefaultCompressionLevel = LZ4HC_DEFAULT_CLEVEL
End Function

Public Function Lz4HC_GetMinCompressionLevel() As Long
    Lz4HC_GetMinCompressionLevel = LZ4HC_MIN_CLEVEL
End Function

Public Function Lz4HC_GetMaxCompressionLevel() As Long
    Lz4HC_GetMaxCompressionLevel = LZ4HC_MAX_CLEVEL
End Function

'DispCallFunc wrapper originally by Olaf Schmidt, with a few minor modifications; see the top of this class
' for a link to his original, unmodified version
Private Function CallCDeclW(ByVal lProc As LZ4_ProcAddress, ByVal fRetType As VbVarType, ParamArray pa() As Variant) As Variant

    Dim i As Long, vTemp() As Variant, hResult As Long
    
    Dim numParams As Long
    If (UBound(pa) < LBound(pa)) Then numParams = 0 Else numParams = UBound(pa) + 1
    
    If IsMissing(pa) Then
        ReDim vTemp(0) As Variant
    Else
        vTemp = pa 'make a copy of the params to prevent problems with VT_ByRef members in the ParamArray
    End If
    
    For i = 0 To numParams - 1
        m_vType(i) = VarType(vTemp(i))
        m_vPtr(i) = VarPtr(vTemp(i))
    Next i
    
    Const CC_CDECL As Long = 1
    hResult = DispCallFunc(0, m_ProcAddresses(lProc), CC_CDECL, fRetType, i, m_vType(0), m_vPtr(0), CallCDeclW)
    
End Function

Private Sub InternalError(ByVal errString As String, Optional ByVal faultyReturnCode As Long = 256)
    If (faultyReturnCode <> 256) Then
        PDDebug.LogAction "lz4 returned an error code: " & faultyReturnCode, PDM_External_Lib
    Else
        PDDebug.LogAction "lz4 experienced an error; additional explanation may be: " & errString, PDM_External_Lib
    End If
End Sub
