Attribute VB_Name = "Plugin_zLib_Interface"
'***************************************************************************
'File Compression Interface (via zLib)
'Copyright 2002-2014 by Tanner Helland
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

'Custom compressed file header
Type CompressionHeader
    Verification As String * 4
    OriginalExt As String * 3
    originalSize As Long
End Type

'Actual header variable
Private FileHeader As CompressionHeader

'Filename used for parsing
Private dstFilename As String

'Used to compare compression ratios
Private originalSize As Long, compressedSize As Long

'Is zLib available as a plugin?  (NOTE: this is now determined separately from g_ZLibEnabled.)
Public Function isZLibAvailable() As Boolean
    If FileExist(g_PluginPath & "zlibwapi.dll") Then isZLibAvailable = True Else isZLibAvailable = False
End Function

'Return the current zLib version
Public Function getZLibVersion() As Long

    If Not g_ZLibEnabled Then
        getZLibVersion = -1
        Exit Function
    End If

    'Manually load the DLL from the "g_PluginPath" folder (should be App.Path\Data\Plugins)
    Dim hLib As Long
    hLib = LoadLibrary(g_PluginPath & "zlibwapi.dll")
    
    'Check the version
    Dim zLibVer As Long
    zLibVer = zlibVersion()
    
    'Release the zLib library
    FreeLibrary hLib
    
    getZLibVersion = zLibVer

End Function

'Compress a file
Public Function CompressFile(ByVal srcFilename As String, Optional ByVal DispResults As Boolean = False) As Boolean
    
    'Manually load the DLL from the "g_PluginPath" folder (should be App.Path\Data\Plugins)
    Dim hLib As Long
    hLib = LoadLibrary(g_PluginPath & "zlibwapi.dll")

    'Used to strip the extension from the original filename
    Dim fExtension As String * 3
    fExtension = GetExtension(srcFilename)

    'Allocate an array to receive the data from a file
    Dim DataBytes() As Byte
    ReDim DataBytes(FileLen(srcFilename) - 1)

    'Copy the data from the source into a numerical array
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open srcFilename For Binary Access Read As #fileNum
        Get #fileNum, , DataBytes()
    Close #fileNum

    'Track the original size
    originalSize = UBound(DataBytes) + 1

    'Allocate memory for a temporary compression array
    Dim bufferSize As Long
    Dim TempBuffer() As Byte
    bufferSize = UBound(DataBytes) + 1
    bufferSize = bufferSize + (bufferSize * 0.01) + 12
    ReDim TempBuffer(bufferSize)

    'Compress the data using zLib
    Dim result As Long
    result = compress(TempBuffer(0), bufferSize, DataBytes(0), UBound(DataBytes) + 1)

    'Copy the compressed data back into our first array
    ReDim DataBytes(bufferSize - 1)
    CopyMemory DataBytes(0), TempBuffer(0), bufferSize

    'Kill the now useless buffer
    Erase TempBuffer

    'Some very simple error handling
    If result = 0 Then
        compressedSize = UBound(DataBytes) + 1
    Else
        pdMsgBox "An error (#%1) has occurred.  Compression halted.", vbCritical + vbOKOnly + vbApplicationModal, "zLib error", Err.Number
        Exit Function
    End If

    'Build the destination filename with a .pdi extension
    dstFilename = Left(srcFilename, Len(srcFilename) - 4)
    dstFilename = dstFilename & ".pdi"
    If FileExist(dstFilename) Then Kill dstFilename
    
    'Build our custom compressed file header
    FileHeader.Verification = "THZC"
    FileHeader.OriginalExt = fExtension
    FileHeader.originalSize = originalSize
    'Write the header and then the compressed data
    Open dstFilename For Binary Access Write As #fileNum
        Put #fileNum, 1, FileHeader
        Put #fileNum, , DataBytes()
    Close #fileNum

    'Kill the now unnecessary compressed data
    Erase DataBytes

    'Free the zLib library from memory
    FreeLibrary hLib

    'Kill the old file (may want to disable when debugging...?)
    'If SrcFilename <> DstFilename Then Kill SrcFilename

    'Report the compression ratio
    If DispResults Then pdMsgBox "File compressed from %1 bytes to %2 bytes.  Ratio: %3 %", vbInformation + vbOKOnly, "Compression results", originalSize, compressedSize, CStr(100 - (100 * (CDbl(compressedSize) / CDbl(originalSize))))

    'Return
    CompressFile = True

End Function

'Decompress a file
Public Function DecompressFile(ByVal srcFilename As String, Optional ByVal DispResults As Boolean = False) As Boolean
    
    On Error Resume Next
    
    'Manually load the DLL from the "g_PluginPath" folder (should be App.Path\Data\Plugins)
    Dim hLib As Long
    hLib = LoadLibrary(g_PluginPath & "zlibwapi.dll")

    'Allocate a temporary array for receiving the compressed data
    Dim DataBytes() As Byte
    ReDim DataBytes(FileLen(srcFilename) - Len(FileHeader) - 1)
    
    'Copy out the header and then the compressed data
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open srcFilename For Binary Access Read As #fileNum
        Get #fileNum, 1, FileHeader
        'Make sure that we've got a valid file
        If FileHeader.Verification <> "THZC" Then
            Close #fileNum
            DecompressFile = False
            Exit Function
        Else
            Get #fileNum, , DataBytes()
        End If
    Close #fileNum
    
    'Get the compressed size
    originalSize = UBound(DataBytes) + 1
    
    'Allocate memory for buffers
    Dim bufferSize As Long
    Dim TempBuffer() As Byte
    bufferSize = FileHeader.originalSize
    bufferSize = bufferSize + (bufferSize * 0.01) + 12
    ReDim TempBuffer(bufferSize)
    
    'Decompress the data using zLib
    Dim result As Long
    result = uncompress(TempBuffer(0), bufferSize, DataBytes(0), UBound(DataBytes) + 1)
    
    'Copy the uncompressed data back into our first array
    ReDim DataBytes(bufferSize - 1)
    CopyMemory DataBytes(0), TempBuffer(0), bufferSize
    
    'Some very simple error handling
    If result = 0 Then
        compressedSize = UBound(DataBytes) + 1
    Else
        pdMsgBox "An error (#%1) has occurred.  Compression halted.", vbCritical + vbOKOnly + vbApplicationModal, "zLib error", Err.Number
        Exit Function
    End If
    
    'Kill the now unnecessary buffer
    Erase TempBuffer
    
    'Free the zLib library from memory
    FreeLibrary hLib
    
    'Build the output path using the original filename
    dstFilename = Left(srcFilename, Len(srcFilename) - 3)
    dstFilename = dstFilename & FileHeader.OriginalExt
    
    'If that file exists, murder it
    If FileExist(dstFilename) Then Kill dstFilename
    
    'Write the uncompressed data back into its original format
    Open dstFilename For Binary Access Write As #fileNum
        Put #fileNum, , DataBytes()
    Close #fileNum
    
    'Kill the now unnecessary data array
    Erase DataBytes
        
    'Kill the original compressed file (note: may want to disable when debugging, so no important files are lost)
    If srcFilename <> dstFilename Then Kill srcFilename
    
    'Display decompression results
    If DispResults Then pdMsgBox "File decompressed from %1 bytes to %2 bytes.", vbInformation + vbOKOnly, "Compression results", originalSize, compressedSize
    
    'Return
    DecompressFile = True
    
End Function

'Fill a destination array with the compressed version of a source array.
Public Function compressArray(ByRef srcArray() As Byte, ByRef dstArray() As Byte, Optional ByRef origSize As Long = 0, Optional ByRef compressSize As Long = 0) As Boolean

    'Manually load the DLL from the "g_PluginPath" folder (should be App.Path\Data\Plugins)
    Dim hLib As Long
    hLib = LoadLibrary(g_PluginPath & "zlibwapi.dll")

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

    'Free the zLib library from memory
    FreeLibrary hLib

    'Return success or failure (zLib returns 0 upon a successful compression)
    If zResult = 0 Then
        compressSize = bufferSize
        compressArray = True
    Else
        compressSize = 0
        compressArray = False
    End If

End Function

'Fill a destination array with the compressed version of a source array.  Also, ask for the original size,
' which allows us to avoid wasting time creating poorly sized buffers.
Public Function decompressArray(ByRef srcArray() As Byte, ByRef dstArray() As Byte, ByRef origSize As Long) As Boolean

    'Manually load the DLL from the "g_PluginPath" folder (should be App.Path\Data\Plugins)
    Dim hLib As Long
    hLib = LoadLibrary(g_PluginPath & "zlibwapi.dll")

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

    'Free the zLib library from memory
    FreeLibrary hLib

    'Return success or failure (zLib returns 0 upon a successful compression)
    If zResult = 0 Then
        decompressArray = True
    Else
        decompressArray = False
    End If

End Function
