Attribute VB_Name = "Plugin_zLib"
'***************************************************************************
'File Compression Interface (via zLib)
'Copyright 2002-2016 by Tanner Helland
'Created: 3/02/02
'Last updated: 08/December/16
'Last update: general code clean-up to better integrate with the new Compression wrapper module
'
'Module to handle file compression and decompression to a custom file format via the zLib compression library.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

Private Const ZLIB_OK = 0

'These constants were originally declared in zlib.h.  Note that zLib weirdly supports level 0, which just performs
' a bare memory copy with no compression.  We deliberately omit that possibility here.
Private Const ZLIB_MIN_CLEVEL = 1
Private Const ZLIB_MAX_CLEVEL = 9

'This constant was originally declared (or rather, resolved) in deflate.c.
Private Const ZLIB_DEFAULT_CLEVEL = 6

Private Declare Function compress Lib "zlibwapi" (ByVal ptrToDestBuffer As Long, ByRef dstLen As Long, ByVal ptrToSrcBuffer As Long, ByVal srcLen As Long) As Long
Private Declare Function compress2 Lib "zlibwapi" (ByVal ptrDstBuffer As Long, ByRef dstLen As Long, ByVal ptrSrcBuffer As Any, ByVal srcLen As Long, ByVal cmpLevel As Long) As Long
Private Declare Function uncompress Lib "zlibwapi" (ByVal ptrToDestBuffer As Long, ByRef dstLen As Long, ByVal ptrToSrcBuffer As Long, ByVal srcLen As Long) As Long
Private Declare Function zlibVersion Lib "zlibwapi" () As Long

'A single zLib handle is maintained for the life of a PD instance; see InitializeZLib and ReleaseZLib, below.
Private m_ZLibHandle As Long

'Initialize zLib.  Do not call this until you have verified zLib's existence (typically via the PluginManager module)
Public Function InitializeZLib(ByRef pathToDLLFolder As String) As Boolean
    
    'Manually load the DLL from the "g_PluginPath" folder (should be App.Path\Data\Plugins)
    Dim zLibPath As String
    zLibPath = pathToDLLFolder & "zlibwapi.dll"
    m_ZLibHandle = LoadLibrary(StrPtr(zLibPath))
    InitializeZLib = CBool(m_ZLibHandle <> 0)
    
    #If DEBUGMODE = 1 Then
        If (Not InitializeZLib) Then
            pdDebug.LogAction "WARNING!  LoadLibrary failed to load zLib.  Last DLL error: " & Err.LastDllError
            pdDebug.LogAction "(FYI, the attempted path was: " & zLibPath & ")"
        End If
    #End If
    
End Function

'When PD closes, make sure to release our open zLib handle!
Public Sub ReleaseZLib()
    If (m_ZLibHandle <> 0) Then
        FreeLibrary m_ZLibHandle
        m_ZLibHandle = 0
    End If
End Sub

Public Function IsZLibAvailable() As Boolean
    IsZLibAvailable = CBool(m_ZLibHandle <> 0)
End Function

'Return the current zLib version
Public Function GetZLibVersion() As String

    If (m_ZLibHandle <> 0) Then
        
        'Get a pointer to the version string
        Dim ptrZLibVer As Long
        ptrZLibVer = zlibVersion()
        
        'Convert the char * to a VB string
        Dim cUnicode As pdUnicode
        Set cUnicode = New pdUnicode
        GetZLibVersion = cUnicode.ConvertCharPointerToVBString(ptrZLibVer, False, 255)
        
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
    If CBool(compress2(VarPtr(dstArray(0)), dstArraySizeInBytes, ptrToSrcData, srcDataSize, compressionLevel) = ZLIB_OK) Then
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
    ZlibCompressNakedPointers = CBool(compress2(dstPointer, dstLength, srcPointer, srcLength, compressionLevel) = ZLIB_OK)
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
    ZlibDecompressArray = CBool(uncompress(VarPtr(dstArray(0)), knownUncompressedSize, ptrToSrcData, srcDataSize) = ZLIB_OK)
    
End Function

'Given arbitrary pointers to both source and destination buffers, decompress a zLib stream.  Obviously, it's assumed the caller
' has knowledge of the size required by the destination buffer (e.g. the decompressed data size was previously stored in a
' file or something), because this function will not modify any buffer sizes.
'
'RETURNS: TRUE on success, FALSE on failure.  The knownUncompressedSize parameter will be filled with the amount of data written
'         to the destination buffer, in bytes (1-based).
Public Function ZlibDecompress_UnsafePtr(ByVal ptrToDstBuffer As Long, ByRef knownUncompressedSize As Long, ByVal ptrToSrcData As Long, ByVal srcDataSize As Long) As Boolean
    ZlibDecompress_UnsafePtr = CBool(uncompress(ptrToDstBuffer, knownUncompressedSize, ptrToSrcData, srcDataSize) = ZLIB_OK)
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

