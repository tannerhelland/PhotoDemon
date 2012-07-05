Attribute VB_Name = "CompressionUtility"
'***************************************************************************
'File Compression Interface (via zLib)
'©2003-2012 Tanner 'DemonSpectre' Helland
'Created: 3/02/02
'Last updated: 24/October/07
'Last update: cleaned up error handling
'
'Module to handle file compression and decompression to a custom file format
'using the zLib compression library.
'
'***************************************************************************

Option Explicit

'API Declarations
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function compress Lib "zlibwapi.dll" (Dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function uncompress Lib "zlibwapi.dll" (Dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

'Custom compressed file header
Type CompressionHeader
    Verification As String * 4
    OriginalExt As String * 3
    OriginalSize As Long
End Type

'Actual header variable
Dim FileHeader As CompressionHeader

'Filename used for parsing
Dim DstFilename As String

'Used to compare compression ratios
Dim OriginalSize As Long, CompressedSize As Long

Public Function CompressFile(ByVal SrcFilename As String, Optional ByVal DispResults As Boolean = False) As Boolean
    
    'Manually load the DLL from the "PluginPath" folder (should be App.Path/Plugins)
    Dim hLib As Long
    hLib = LoadLibrary(PluginPath & "zlibwapi.dll")

    'Used to strip the extension from the original filename
    Dim fExtension As String * 3
    fExtension = GetExtension(SrcFilename)

    'Allocate an array to receive the data from a file
    Dim DataBytes() As Byte
    ReDim DataBytes(FileLen(SrcFilename) - 1)

    'Copy the data from the source into a numerical array
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open SrcFilename For Binary Access Read As #fileNum
        Get #fileNum, , DataBytes()
    Close #fileNum

    'Track the original size
    OriginalSize = UBound(DataBytes) + 1

    'Allocate memory for a temporary compression array
    Dim BufferSize As Long
    Dim TempBuffer() As Byte
    BufferSize = UBound(DataBytes) + 1
    BufferSize = BufferSize + (BufferSize * 0.01) + 12
    ReDim TempBuffer(BufferSize)

    'Compress the data using zLib
    Dim result As Long
    result = compress(TempBuffer(0), BufferSize, DataBytes(0), UBound(DataBytes) + 1)

    'Copy the compressed data back into our first array
    ReDim DataBytes(BufferSize - 1)
    CopyMemory DataBytes(0), TempBuffer(0), BufferSize

    'Kill the now useless buffer
    Erase TempBuffer

    'Some very simple error handling
    If result = 0 Then
        CompressedSize = UBound(DataBytes) + 1
    Else
        MsgBox "An error (#" & Err.Number & ") has occurred.  Compression halted."
        Exit Function
    End If

    'Build the destination filename with a .pdi extension
    DstFilename = Left(SrcFilename, Len(SrcFilename) - 4)
    DstFilename = DstFilename & ".pdi"
    If FileExist(DstFilename) Then Kill DstFilename
    
    'Build our custom compressed file header
    FileHeader.Verification = "THZC"
    FileHeader.OriginalExt = fExtension
    FileHeader.OriginalSize = OriginalSize
    'Write the header and then the compressed data
    Open DstFilename For Binary Access Write As #fileNum
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
    If DispResults = True Then MsgBox "File compressed from " & OriginalSize & " bytes to " & CompressedSize & " bytes.  Ratio: " & CStr(100 - (100 * (CDbl(CompressedSize) / CDbl(OriginalSize)))) & "%"

    'Return
    CompressFile = True

End Function

Public Function DecompressFile(ByVal SrcFilename As String, Optional ByVal DispResults As Boolean = False) As Boolean
    
    'Manually load the DLL from the "PluginPath" folder (should be App.Path/Plugins)
    Dim hLib As Long
    hLib = LoadLibrary(PluginPath & "zlibwapi.dll")

    'Allocate a temporary array for receiving the compressed data
    Dim DataBytes() As Byte
    ReDim DataBytes(FileLen(SrcFilename) - Len(FileHeader) - 1)
    
    'Copy out the header and then the compressed data
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open SrcFilename For Binary Access Read As #fileNum
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
    OriginalSize = UBound(DataBytes) + 1
    
    'Allocate memory for buffers
    Dim BufferSize As Long
    Dim TempBuffer() As Byte
    BufferSize = FileHeader.OriginalSize
    BufferSize = BufferSize + (BufferSize * 0.01) + 12
    ReDim TempBuffer(BufferSize)
    
    'Decompress the data using zLib
    Dim result As Long
    result = uncompress(TempBuffer(0), BufferSize, DataBytes(0), UBound(DataBytes) + 1)
    
    'Copy the uncompressed data back into our first array
    ReDim DataBytes(BufferSize - 1)
    CopyMemory DataBytes(0), TempBuffer(0), BufferSize
    
    'Some very simple error handling
    If result = 0 Then
        CompressedSize = UBound(DataBytes) + 1
    Else
        MsgBox "An error (#" & Err.Number & ") has occurred.  Decompression halted."
        Exit Function
    End If
    
    'Kill the now unnecessary buffer
    Erase TempBuffer
    
    'Build the output path using the original filename
    DstFilename = Left(SrcFilename, Len(SrcFilename) - 3)
    DstFilename = DstFilename & FileHeader.OriginalExt
    
    'If that file exists, murder it
    If FileExist(DstFilename) Then Kill DstFilename
    
    'Write the uncompressed data back into its original format
    Open DstFilename For Binary Access Write As #fileNum
        Put #fileNum, , DataBytes()
    Close #fileNum
    
    'Kill the now unnecessary data array
    Erase DataBytes
    
    'Free the zLib library from memory
    FreeLibrary hLib
    
    'Kill the original compressed file (note: may want to disable when debugging, so no important files are lost)
    If SrcFilename <> DstFilename Then Kill SrcFilename
    
    'Display decompression results
    If DispResults = True Then MsgBox "File decompressed from " & OriginalSize & " bytes to " & CompressedSize & " bytes."
    
    'Return
    DecompressFile = True
    
End Function
