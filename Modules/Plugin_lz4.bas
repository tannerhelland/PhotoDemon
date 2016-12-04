Attribute VB_Name = "Plugin_lz4"
'***************************************************************************
'Lz4 Compression Library Interface
'Copyright 2016-2016 by Tanner Helland
'Created: 04/December/16
'Last updated: 04/December/16
'Last update: initial build
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
'As of v7.0, most internal PD temp files and caches are written using Lz4.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Declare Function LZ4_versionNumber Lib "liblz4" Alias "_LZ4_versionNumber@0" () As Long
Private Declare Function LZ4_compress_fast Lib "liblz4" Alias "_LZ4_compress_fast@20" (ByVal constPtrToSrcBuffer As Long, ByVal ptrToDstBuffer As Long, ByVal srcSizeInBytes As Long, ByVal dstBufferCapacityInBytes As Long, ByVal cAccelerationLevel As Long) As Long
Private Declare Function LZ4_decompress_safe Lib "liblz4" Alias "_LZ4_decompress_safe@16" (ByVal constPtrToSrcBuffer As Long, ByVal ptrToDstBuffer As Long, ByVal srcSizeInBytes As Long, ByVal dstBufferCapacityInBytes As Long) As Long
Private Declare Function LZ4_compressBound Lib "liblz4" Alias "_LZ4_compressBound@4" (ByVal inputSizeInBytes As Long) As Long 'Maximum compressed size in worst case scenario; use this to size your input array

'A single lz4 handle is maintained for the life of a PD instance; see InitializeLz4 and ReleaseLz4, below.
Private m_Lz4Handle As Long

'Initialize lz4.  Do not call this until you have verified its existence (typically via the PluginManager module)
Public Function InitializeLz4() As Boolean

    'Manually load the DLL from the "g_PluginPath" folder (should be App.Path\Data\Plugins)
    Dim lz4Path As String
    lz4Path = g_PluginPath & "liblz4.dll"
    m_Lz4Handle = LoadLibrary(StrPtr(lz4Path))
    InitializeLz4 = CBool(m_Lz4Handle <> 0)
    
    'If we initialized the library successfully, cache some lz4-specific data
    If InitializeLz4 Then
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "lz4 compression engine is ready."
        #End If
    End If
    
    #If DEBUGMODE = 1 Then
        If (Not InitializeLz4) Then
            pdDebug.LogAction "WARNING!  LoadLibrary failed to load lz4.  Last DLL error: " & Err.LastDllError
            pdDebug.LogAction "(FYI, the attempted path was: " & lz4Path & ")"
        End If
    #End If
    
End Function

'When PD closes, make sure to release our open Lz4 handle
Public Sub ReleaseLz4()
    If (m_Lz4Handle <> 0) Then FreeLibrary m_Lz4Handle
    g_Lz4Enabled = False
End Sub

Public Function GetLz4Version() As String
    Dim ptrVersion As Long
    ptrVersion = LZ4_versionNumber()
    GetLz4Version = ptrVersion
End Function

Public Function IsLz4Available() As Boolean
    IsLz4Available = CBool(m_Lz4Handle <> 0)
End Function

'Determine the maximum possible size required by a compression operation.  The destination buffer should be at least
' this large (and if it's even bigger, that's okay too).
Public Function Lz4GetMaxCompressedSize(ByVal srcSize As Long) As Long
    Lz4GetMaxCompressedSize = LZ4_compressBound(srcSize)
End Function

'Compress some arbitrary source pointer + length into a destination array.  Pass the optional "dstArrayIsReady" as TRUE
' (with a matching size descriptor) if you don't want us to auto-size the destination for you.
'
'RETURNS: final size of the compressed data, in bytes.  0 on failure.
'
'IMPORTANT!  The destination array is *not* resized to match the final compressed size.  The caller is responsible
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
    Lz4CompressNakedPointers = CBool(finalSize <> 0)
    
    If Lz4CompressNakedPointers Then
        dstSizeInBytes = finalSize
    Else
        InternalError "lz4_compress failed", finalSize
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
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "lz4_decompress_safe failure inputs: " & VarPtr(dstArray(0)) & ", " & knownUncompressedSize & ", " & ptrToSrcData & ", " & srcDataSize
        #End If
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
        #If DEBUGMODE = 1 Then
            pdDebug.LogAction "lz4_decompress_safe failure inputs: " & ptrToDstBuffer & ", " & knownUncompressedSize & ", " & ptrToSrcData & ", " & srcDataSize
        #End If
        InternalError "lz4_decompress failed", finalSize
        finalSize = 0
    End If
    
    Lz4Decompress_UnsafePtr = finalSize

End Function

Private Sub InternalError(ByVal errString As String, Optional ByVal faultyReturnCode As Long = 256)
    #If DEBUGMODE = 1 Then
        If (faultyReturnCode <> 256) Then
            pdDebug.LogAction "lz4 returned an error code: " & faultyReturnCode, PDM_EXTERNAL_LIB
        Else
            pdDebug.LogAction "lz4 experienced an error; additional explanation may be: " & errString, PDM_EXTERNAL_LIB
        End If
    #End If
End Sub

