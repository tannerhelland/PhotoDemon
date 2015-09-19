Attribute VB_Name = "Plugin_zLib_Interface"
'***************************************************************************
'File Compression Interface (via zLib)
'Copyright 2002-2015 by Tanner Helland
'Created: 3/02/02
'Last updated: 05/August/13
'Last update: standalone functions for compressing and decompressing arrays.  I still need to tie the compress/decompress file
'              routines into these, to avoid duplicating code unnecessarily.
'
'Module to handle file compression and decompression to a custom file format via the zLib compression library.
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://photodemon.org/about/license/
'
'***************************************************************************

Option Explicit

'API Declarations
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function compress Lib "zlibwapi.dll" (Dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function uncompress Lib "zlibwapi.dll" (Dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function zlibVersion Lib "zlibwapi.dll" () As Long

'A single zLib handle is maintained for the life of a PD instance; see initializeZLib and releaseZLib, below.
Private m_ZLibHandle As Long

'Is zLib available as a plugin?  (NOTE: this is now determined separately from g_ZLibEnabled.)
Public Function isZLibAvailable() As Boolean
    
    Dim cFile As pdFSO
    Set cFile = New pdFSO
    
    If cFile.FileExist(g_PluginPath & "zlibwapi.dll") Then isZLibAvailable = True Else isZLibAvailable = False
    
End Function

'Initialize zLib.  Do not call this until you have verified zLib's existence (typically via isZLibAvailable(), above)
Public Function initializeZLib() As Boolean
    
    'Manually load the DLL from the "g_PluginPath" folder (should be App.Path\Data\Plugins)
    m_ZLibHandle = LoadLibrary(g_PluginPath & "zlibwapi.dll")
    initializeZLib = CBool(m_ZLibHandle <> 0)
    
End Function

'When PD closes, make sure to release our open zLib handle!
Public Sub releaseZLib()
    If m_ZLibHandle <> 0 Then FreeLibrary m_ZLibHandle
    g_ZLibEnabled = False
End Sub

'Return the current zLib version
Public Function getZLibVersion() As String

    If Not g_ZLibEnabled Then
        getZLibVersion = -1
        Exit Function
    End If
    
    'Get a pointer to the version string
    Dim ptrZLibVer As Long
    ptrZLibVer = zlibVersion()
    
    'Convert the char * to a VB string
    Dim cUnicode As pdUnicode
    Set cUnicode = New pdUnicode
    getZLibVersion = cUnicode.ConvertCharPointerToVBString(ptrZLibVer, False, 255)
    
End Function

'Fill a destination array with the compressed version of a source array.
Public Function compressArray(ByRef srcArray() As Byte, ByRef dstArray() As Byte, Optional ByRef origSize As Long = 0, Optional ByRef compressSize As Long = 0) As Boolean
    
    'Mark the original size
    origSize = UBound(srcArray) - LBound(srcArray) + 1

    'Allocate memory for a temporary compression array.  Per the zLib spec, the buffer should be slightly larger than
    ' the original array to allow space for generating the compression data.
    Dim bufferSize As Long
    bufferSize = origSize + (origSize * 0.01) + 12
    ReDim dstArray(0 To bufferSize) As Byte

    'Compress the data using zLib
    Dim zResult As Long
    zResult = compress(dstArray(0), bufferSize, srcArray(0), origSize)
    
    'Let VB repopulate its SafeArray structure by redimming the array.
    ReDim Preserve dstArray(0 To bufferSize - 1) As Byte
    
    'Return success or failure (zLib returns 0 upon a successful compression)
    If zResult = 0 Then
        compressSize = bufferSize
        compressArray = True
    Else
        compressSize = 0
        compressArray = False
    End If

End Function

'Given an arbitrary pointer and a length, compress all that data into a normal VB array.  This function will resize the destination
' array as necessary, but - obviously - the caller is responsible for verifying the source data.
'
'ALSO NOTE: this function WILL NOT PRECISELY SIZE THE DESTINATION ARRAY.  zLib requires the destination buffer to make extra
' space available for temporary use during compression.  Previously, we would Redim Preserve the compressed results so that
' the destination array is precisely sized, but this really isn't necessary for most use-cases.
'
'INSTEAD, the caller needs to pass a Long as the finalCompressedSize param.  This function will fill that with the final
' compression size, which the caller can use to call their own ReDim preserve as necessary.
Public Function compressNakedPointerToArray(ByVal srcPointer As Long, ByVal srcLength As Long, ByRef dstArray() As Byte, ByRef finalCompressedSize As Long) As Boolean
    
    'Allocate memory for a temporary compression array.  Per the zLib spec, the buffer should be slightly larger than
    ' the original array to allow space for generating the compression data.
    Dim bufferSize As Long
    bufferSize = srcLength + (CSng(srcLength) * 0.01) + 12
    ReDim dstArray(0 To bufferSize) As Byte

    'Compress the data.  (Note that zLib returns 0 upon a successful compression.)
    Dim zResult As Long
    If compress(dstArray(0), bufferSize, ByVal srcPointer, srcLength) = 0 Then
        finalCompressedSize = bufferSize
        compressNakedPointerToArray = True
    Else
        finalCompressedSize = 0
        compressNakedPointerToArray = False
    End If
    
End Function

'Given arbitrary pointers to both source and destination buffers, decompress a zLib stream.  Obviously, it's assumed the caller
' has knowledge of the size required by the destination buffer (e.g. the decompressed data size was previously stored in a
' file or something), because this function will not modify any buffer sizes.
Public Function decompressNakedPointers(ByVal srcPointer As Long, ByVal srcLength As Long, ByVal dstPointer As Long, ByVal dstLength As Long) As Boolean
    
    'Decompress the data using zLib
    decompressNakedPointers = CBool(uncompress(ByVal dstPointer, dstLength, ByVal srcPointer, srcLength) = 0)
    
End Function

'Fill a destination array with the compressed version of a source array.  Also, ask for the original size,
' which allows us to avoid wasting time creating poorly sized buffers.
Public Function decompressArray(ByRef srcArray() As Byte, ByRef dstArray() As Byte, ByRef origSize As Long) As Boolean
    
    'Calculate the size of the compressed array
    Dim compressedSize As Long
    compressedSize = UBound(srcArray) - LBound(srcArray) + 1

    'Allocate memory for a temporary decompression array.  Per the zLib spec, the buffer should be slightly larger than
    ' the original array to allow space for generating the compression data.
    Dim bufferSize As Long
    bufferSize = origSize + (origSize * 0.01) + 12
    ReDim dstArray(0 To bufferSize - 1) As Byte

    'Decompress the data using zLib
    Dim zResult As Long
    zResult = uncompress(dstArray(0), bufferSize, srcArray(0), compressedSize)
    
    'Let VB repopulate its SafeArray structure by redimming the array.
    ReDim Preserve dstArray(0 To bufferSize - 1) As Byte
    
    'Return success or failure (zLib returns 0 upon a successful compression)
    If zResult = 0 Then
        decompressArray = True
    Else
        decompressArray = False
    End If

End Function
