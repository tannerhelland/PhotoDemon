Attribute VB_Name = "Plugin_zLib"
'***************************************************************************
'File Compression Interface (via zLib)
'Copyright 2002-2018 by Tanner Helland
'Created: 3/02/02
'Last updated: 08/December/16
'Last update: general code clean-up to better integrate with the new Compression wrapper module
'
'Module to handle file compression and decompression to a custom file format via the zLib compression library.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

Private Enum PD_ZLibReturn
    Z_OK = 0
    Z_STREAM_END = 1
    Z_NEED_DICT = 2
    Z_ERRNO = -1
    Z_STREAM_ERROR = -2
    Z_DATA_ERROR = -3
    Z_MEM_ERROR = -4
    Z_BUF_ERROR = -5
    Z_VERSION_ERROR = -6
End Enum

#If False Then
    Private Const Z_OK = 0, Z_STREAM_END = 1, Z_NEED_DICT = 2, Z_ERRNO = -1, Z_STREAM_ERROR = -2, Z_DATA_ERROR = -3, Z_MEM_ERROR = -4, Z_BUF_ERROR = -5, Z_VERSION_ERROR = -6
#End If

'These constants were originally declared in zlib.h.  Note that zLib weirdly supports level 0, which just performs
' a bare memory copy with no compression.  We deliberately omit that possibility here.
Private Const ZLIB_MIN_CLEVEL As Long = 1
Private Const ZLIB_MAX_CLEVEL As Long = 9

'This constant was originally declared (or rather, resolved) in deflate.c.
Private Const ZLIB_DEFAULT_CLEVEL As Long = 6

Private Declare Function compress2 Lib "zlibwapi" (ByVal ptrDstBuffer As Long, ByRef dstLen As Long, ByVal ptrSrcBuffer As Any, ByVal srcLen As Long, ByVal cmpLevel As Long) As PD_ZLibReturn
Private Declare Function uncompress Lib "zlibwapi" (ByVal ptrToDestBuffer As Long, ByRef dstLen As Long, ByVal ptrToSrcBuffer As Long, ByVal srcLen As Long) As PD_ZLibReturn
Private Declare Function crc32 Lib "zlibwapi" (ByVal initValue As Long, ByVal ptrDstBuffer As Long, ByVal dstBufferLen As Long) As Long
Private Declare Function zlibVersion Lib "zlibwapi" () As Long

'A single zLib handle is maintained for the life of a PD instance; see InitializeZLib and ReleaseZLib, below.
Private m_ZLibHandle As Long

'Initialize zLib.  Do not call this until you have verified zLib's existence (typically via the PluginManager module)
Public Function InitializeZLib() As Boolean
    
    'Manually load the DLL from the plugin folder (should be App.Path\Data\Plugins)
    Dim zLibPath As String
    zLibPath = App.Path & "zlibwapi.dll"
    m_ZLibHandle = VBHacks.LoadLib(zLibPath)
    InitializeZLib = (m_ZLibHandle <> 0)
    
    If (Not InitializeZLib) Then
        Debug.Print "WARNING!  LoadLibrary failed to load zLib.  Last DLL error: " & Err.LastDllError
        Debug.Print "(FYI, the attempted path was: " & zLibPath & ")"
    End If
        
End Function

'When PD closes, make sure to release our open zLib handle!
Public Sub ReleaseZLib()
    If (m_ZLibHandle <> 0) Then
        VBHacks.FreeLib m_ZLibHandle
        m_ZLibHandle = 0
    End If
End Sub

Public Function IsZLibAvailable() As Boolean
    IsZLibAvailable = (m_ZLibHandle <> 0)
End Function

'Return the current zLib version
Public Function GetZLibVersion() As String

    If (m_ZLibHandle <> 0) Then
        
        'Get a pointer to the version string
        Dim ptrZLibVer As Long
        ptrZLibVer = zlibVersion()
        
        'Convert the char * to a VB string
        GetZLibVersion = Strings.StringFromCharPtr(ptrZLibVer, False, 255)
        
    Else
        GetZLibVersion = vbNullString
    End If
    
End Function

'Fill a destination array with the compressed version of a source array.
' Returns: final size of the compressed data, in bytes.  0 on failure.
Public Function ZlibCompressArray(ByRef dstArray() As Byte, ByVal ptrToSrcData As Long, ByVal srcDataSize As Long, Optional ByVal dstArrayIsReady As Boolean = False, Optional ByVal dstArraySizeInBytes As Long = 0, Optional ByVal compressionLevel As Long = -1) As Long
    
    'Validate the requested compression level
    If (compressionLevel < ZLIB_MIN_CLEVEL) Then
        compressionLevel = ZLIB_DEFAULT_CLEVEL
    ElseIf (compressionLevel > ZLIB_MAX_CLEVEL) Then
        compressionLevel = ZLIB_MAX_CLEVEL
    End If
    
    'Prep the destination array, as necessary
    If (Not dstArrayIsReady) Or (dstArraySizeInBytes = 0) Then
        dstArraySizeInBytes = ZlibGetMaxCompressedSize(srcDataSize)
        ReDim dstArray(0 To dstArraySizeInBytes - 1) As Byte
    End If

    'Compress the data using zLib
    If (compress2(VarPtr(dstArray(0)), dstArraySizeInBytes, ptrToSrcData, srcDataSize, compressionLevel) = Z_OK) Then
        ZlibCompressArray = dstArraySizeInBytes
    Else
        ZlibCompressArray = 0
    End If

End Function

'Given arbitrary pointers to both source and destination buffers, compress a zLib stream.  Obviously, it's assumed the caller
' has knowledge of the size required by the destination buffer, because this function will not modify any buffer sizes.
'
'RETURNS: TRUE on success, FALSE on failure.  The dstLength parameter will be filled with the amount of data written to dstPoint
'         (in bytes, 1-based).
Public Function ZlibCompressNakedPointers(ByVal dstPointer As Long, ByRef dstLength As Long, ByVal srcPointer As Long, ByVal srcLength As Long, Optional ByVal compressionLevel As Long = -1) As Boolean
    If (compressionLevel < ZLIB_MIN_CLEVEL) Then
        compressionLevel = ZLIB_DEFAULT_CLEVEL
    ElseIf (compressionLevel > ZLIB_MAX_CLEVEL) Then
        compressionLevel = ZLIB_MAX_CLEVEL
    End If
    ZlibCompressNakedPointers = (compress2(dstPointer, dstLength, srcPointer, srcLength, compressionLevel) = Z_OK)
End Function

'Decompress some arbitrary source pointer + length into a destination array.  Pass the optional "dstArrayIsReady" as TRUE
' (with a matching size descriptor) if you don't want us to auto-size the destination for you.
'
'RETURNS: TRUE if successful, FALSE otherwise.  The knownUncompressedSize parameter is filled with the amount of data written
'         to the destination buffer, in bytes (1-based).
'
'IMPORTANT!  The destination array is *not* resized to match the returned size.  The caller is responsible for this.
Public Function ZlibDecompressArray(ByRef dstArray() As Byte, ByVal ptrToSrcData As Long, ByVal srcDataSize As Long, ByRef knownUncompressedSize As Long, Optional ByVal dstArrayIsReady As Boolean = False) As Long
    
    'Prep the destination array, as necessary
    If (Not dstArrayIsReady) Then
        ReDim dstArray(0 To knownUncompressedSize - 1) As Byte
    End If
    
    'Perform decompression
    ZlibDecompressArray = (uncompress(VarPtr(dstArray(0)), knownUncompressedSize, ptrToSrcData, srcDataSize) = Z_OK)
    
End Function

'Given arbitrary pointers to both source and destination buffers, decompress a zLib stream.  Obviously, it's assumed the caller
' has knowledge of the size required by the destination buffer (e.g. the decompressed data size was previously stored in a
' file or something), because this function will not modify any buffer sizes.
'
'RETURNS: TRUE on success, FALSE on failure.  The knownUncompressedSize parameter will be filled with the amount of data written
'         to the destination buffer, in bytes (1-based).
Public Function ZlibDecompress_UnsafePtr(ByVal ptrToDstBuffer As Long, ByRef knownUncompressedSize As Long, ByVal ptrToSrcData As Long, ByVal srcDataSize As Long) As Boolean
    Dim zlReturn As PD_ZLibReturn
    zlReturn = uncompress(ptrToDstBuffer, knownUncompressedSize, ptrToSrcData, srcDataSize)
    If (zlReturn < 0) Then InternalError "ZlibDecompress_UnsafePtr", "unknown", zlReturn
    ZlibDecompress_UnsafePtr = (zlReturn = Z_OK)
End Function

'If you don't know the size of the required decompression buffer in advance (shame on you), you can use this function
' to attempt a partial decompression.  ZLib returns code (-5) Z_BUF_ERROR if there is not enough output space for the
' full stream.  It's up to you to increase buffer size and try again.  (Similarly, this function does not log failures.)
'RETURNS: zLib return code, unmodified.  The knownUncompressedSize parameter will be filled with the amount of data written
'         to the destination buffer, in bytes (1-based).
Public Function ZlibDecompress_UnsafePtrEx(ByVal ptrToDstBuffer As Long, ByRef knownUncompressedSize As Long, ByVal ptrToSrcData As Long, ByVal srcDataSize As Long) As Long
    ZlibDecompress_UnsafePtrEx = uncompress(ptrToDstBuffer, knownUncompressedSize, ptrToSrcData, srcDataSize)
End Function

'Determine the maximum possible size required by a compression operation.  The destination buffer should be at least
' this large (and if it's even bigger, that's okay too).
Public Function ZlibGetMaxCompressedSize(ByVal srcSize As Long) As Long
    ZlibGetMaxCompressedSize = srcSize + (CSng(srcSize) * 0.01) + 12
End Function

Public Function ZLib_GetDefaultCompressionLevel() As Long
    ZLib_GetDefaultCompressionLevel = ZLIB_DEFAULT_CLEVEL
End Function

Public Function ZLib_GetMinCompressionLevel() As Long
    ZLib_GetMinCompressionLevel = ZLIB_MIN_CLEVEL
End Function

Public Function ZLib_GetMaxCompressionLevel() As Long
    ZLib_GetMaxCompressionLevel = ZLIB_MAX_CLEVEL
End Function

'ZLib also provides checksum functionality
Public Function ZLib_GetCRC32(ByVal ptrToData As Long, ByVal dataLength As Long, Optional ByVal startValue As Long = 0&) As Long
    If (startValue = 0&) Then startValue = crc32(0&, 0&, 0&)
    ZLib_GetCRC32 = crc32(startValue, ptrToData, dataLength)
End Function

'ZLib errors are automatically reported to PDDebug
Private Sub InternalError(ByRef funcName As String, errDescription As String, Optional ByVal errValue As Long = 0)
    Debug.Print "WARNING!  ZLib." & funcName & "() reported an error (" & CStr(errValue) & "): " & errDescription
End Sub
