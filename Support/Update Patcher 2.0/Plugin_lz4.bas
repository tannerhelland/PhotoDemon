Attribute VB_Name = "Plugin_lz4"
'***************************************************************************
'Lz4 Compression Library Interface
'Copyright 2016-2018 by Tanner Helland
'Created: 04/December/16
'Last updated: 07/December/16
'Last update: add LZ4_HC compression algorithm.  (LZ4_HC has no special decompression algorithm;
'             it uses an identical LZ4 frame and block format, so standard LZ4 decompression can be used.)
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
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This constant was originally declared in lz4.c.  Note that lz4 does *not* support variable compression levels.
' Instead, it supports variable *acceleration* levels.  The difference is that bigger values = worse compression.
Private Const LZ4_MIN_ALEVEL As Long = 1
Private Const LZ4_DEFAULT_ALEVEL As Long = 1

'This value is not declared by the lz4 library, and technically, there is no maximum value.  Compression just
' approaches 0% as you increase the level.  I provide a "magic number" cap simply so it supports the same
' default/min/max functions as the other libraries
Private Const LZ4_MAX_ALEVEL As Long = 500

'These constants were originally declared in lz4_hc.h
Private Const LZ4HC_MIN_CLEVEL As Long = 3
Private Const LZ4HC_DEFAULT_CLEVEL As Long = 9
Private Const LZ4HC_MAX_CLEVEL As Long = 12

Private Declare Function LZ4_versionNumber Lib "liblz4" Alias "_LZ4_versionNumber@0" () As Long
Private Declare Function LZ4_compress_fast Lib "liblz4" Alias "_LZ4_compress_fast@20" (ByVal constPtrToSrcBuffer As Long, ByVal ptrToDstBuffer As Long, ByVal srcSizeInBytes As Long, ByVal dstBufferCapacityInBytes As Long, ByVal cAccelerationLevel As Long) As Long
Private Declare Function LZ4_compress_HC Lib "liblz4" Alias "_LZ4_compress_HC@20" (ByVal constPtrToSrcBuffer As Long, ByVal ptrToDstBuffer As Long, ByVal srcSizeInBytes As Long, ByVal dstBufferCapacityInBytes As Long, ByVal cCompressionLevel As Long) As Long
Private Declare Function LZ4_decompress_safe Lib "liblz4" Alias "_LZ4_decompress_safe@16" (ByVal constPtrToSrcBuffer As Long, ByVal ptrToDstBuffer As Long, ByVal srcSizeInBytes As Long, ByVal dstBufferCapacityInBytes As Long) As Long
Private Declare Function LZ4_compressBound Lib "liblz4" Alias "_LZ4_compressBound@4" (ByVal inputSizeInBytes As Long) As Long 'Maximum compressed size in worst case scenario; use this to size your input array

'A single lz4 handle is maintained for the life of a PD instance; see InitializeLz4 and ReleaseLz4, below.
Private m_Lz4Handle As Long

'Initialize lz4.  Do not call this until you have verified its existence (typically via the PluginManager module)
Public Function InitializeLz4(ByRef pathToDLLFolder As String) As Boolean

    'Manually load the DLL from the plugin folder (should be App.Path\Data\Plugins)
    Dim lz4Path As String
    lz4Path = pathToDLLFolder & "liblz4.dll"
    m_Lz4Handle = VBHacks.LoadLib(lz4Path)
    InitializeLz4 = (m_Lz4Handle <> 0)
    
    'If we initialized the library successfully, cache some lz4-specific data
    If InitializeLz4 Then
        Debug.Print "lz4 and lz4hc compression engines are ready."
    Else
        Debug.Print "WARNING!  LoadLibrary failed to load lz4.  Last DLL error: " & Err.LastDllError
        Debug.Print "(FYI, the attempted path was: " & lz4Path & ")"
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
    Dim ptrVersion As Long
    ptrVersion = LZ4_versionNumber()
    GetLz4Version = ptrVersion
End Function

Public Function IsLz4Available() As Boolean
    IsLz4Available = (m_Lz4Handle <> 0)
End Function

'Determine the maximum possible size required by a compression operation.  The destination buffer should be at least
' this large (and if it's even bigger, that's okay too).
Public Function Lz4GetMaxCompressedSize(ByVal srcSize As Long) As Long
    Lz4GetMaxCompressedSize = LZ4_compressBound(srcSize)
End Function

'Compress some arbitrary source pointer + length into a destination array.  Pass the optional "dstArrayIsReady" as TRUE
' (with a matching size descriptor) if you don't want us to auto-create the destination buffer for you.
'
'RETURNS: final size of the compressed data, in bytes.  0 on failure.
'
'IMPORTANT!  The destination array is *not* trimmed to match the final compressed size.  The caller is responsible
' for this, if they want it.
Public Function Lz4CompressArray(ByRef dstArray() As Byte, ByVal ptrToSrcData As Long, ByVal srcDataSize As Long, Optional ByVal dstArrayIsReady As Boolean = False, Optional ByVal dstArraySizeInBytes As Long = 0, Optional ByVal compressionAcceleration As Long = -1) As Long
    
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
    finalSize = LZ4_compress_fast(ptrToSrcData, VarPtr(dstArray(0)), srcDataSize, dstArraySizeInBytes, compressionAcceleration)
    
    Lz4CompressArray = finalSize

End Function

'Compress some arbitrary source buffer to an arbitrary destination buffer.  Caller is responsible for all allocations.
' Returns: success/failure, and size of the written data in dstSizeInBytes (passed ByRef).
Public Function Lz4CompressNakedPointers(ByVal dstPointer As Long, ByRef dstSizeInBytes As Long, ByVal srcPointer As Long, ByVal srcSizeInBytes As Long, Optional ByVal compressionAcceleration As Long = -1) As Boolean
    
    Dim finalSize As Long
    finalSize = LZ4_compress_fast(srcPointer, dstPointer, srcSizeInBytes, dstSizeInBytes, compressionAcceleration)
    
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
    finalSize = LZ4_compress_HC(ptrToSrcData, VarPtr(dstArray(0)), srcDataSize, dstArraySizeInBytes, compressionLevel)
    
    Lz4HCCompressArray = finalSize

End Function

'High-compression variant of the normal LZ4 compression function.
Public Function Lz4HCCompressNakedPointers(ByVal dstPointer As Long, ByRef dstSizeInBytes As Long, ByVal srcPointer As Long, ByVal srcSizeInBytes As Long, Optional ByVal compressionLevel As Long = -1) As Boolean
    
    Dim finalSize As Long
    finalSize = LZ4_compress_HC(srcPointer, dstPointer, srcSizeInBytes, dstSizeInBytes, compressionLevel)
    
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
    
    'Prep the destination array, as necessary
    If (Not dstArrayIsReady) Then
        ReDim dstArray(0 To knownUncompressedSize - 1) As Byte
    End If
    
    'Perform decompression
    Dim finalSize As Long
    finalSize = LZ4_decompress_safe(ptrToSrcData, VarPtr(dstArray(0)), srcDataSize, knownUncompressedSize)
    
    'Check for error returns
    If (finalSize <= 0) Then
        Debug.Print "lz4_decompress_safe failure inputs: " & VarPtr(dstArray(0)) & ", " & knownUncompressedSize & ", " & ptrToSrcData & ", " & srcDataSize
        InternalError "lz4_decompress failed", finalSize
        finalSize = 0
    End If
    
    Lz4DecompressArray = finalSize

End Function

Public Function Lz4Decompress_UnsafePtr(ByVal ptrToDstBuffer As Long, ByVal knownUncompressedSize As Long, ByVal ptrToSrcData As Long, ByVal srcDataSize As Long) As Long
    
    'Perform decompression
    Dim finalSize As Long
    finalSize = LZ4_decompress_safe(ptrToSrcData, ptrToDstBuffer, srcDataSize, knownUncompressedSize)
    
    'Check for error returns
    If (finalSize <= 0) Then
        Debug.Print "lz4_decompress_safe failure inputs: " & ptrToDstBuffer & ", " & knownUncompressedSize & ", " & ptrToSrcData & ", " & srcDataSize
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

Private Sub InternalError(ByVal errString As String, Optional ByVal faultyReturnCode As Long = 256)
    If (faultyReturnCode <> 256) Then
        Debug.Print "lz4 returned an error code: " & faultyReturnCode
    Else
        Debug.Print "lz4 experienced an error; additional explanation may be: " & errString
    End If
End Sub
