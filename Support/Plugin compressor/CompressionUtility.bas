Attribute VB_Name = "CompressionUtility"
'***************************************************************************
'PhotoDemon Plugin Compression Tool (uses zLib)
'Copyright ©2002-2013 by Tanner Helland
'Created: 3/02/02
'Last updated: 18/June/2013
'Last update: moved project into main PD Git repository
'
'Module to handle file compression and decompression to a custom file format via the zLib compression library.
'
'NOTE: this project is intended only as a support tool for PhotoDemon.  It is not designed or tested for general-purpose use.
'       I do not have any intention of supporting this tool outside its intended use, so please do not submit bug reports
'       regarding this project unless they directly relate to its intended purpose (compressing PhotoDemon plugins).
'
'All source code in this file is licensed under a modified BSD license.  This means you may use the code in your own
' projects IF you provide attribution.  For more information, please visit http://www.tannerhelland.com/photodemon/#license
'
'***************************************************************************

Option Explicit

'API Declarations
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function compress Lib "zlibwapi.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function compress2 Lib "zlibwapi.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long, ByVal level As Integer) As Long
Private Declare Function uncompress Lib "zlibwapi.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long

'Custom compressed file header
Type CompressionHeader
    Verification As String * 4
    OriginalExt As String * 3
    OriginalSize As Long
End Type

'Actual header variable
Dim FileHeader As CompressionHeader

'Filename used for parsing
Dim dstFilename As String

'Used to compare compression ratios
Dim OriginalSize As Long, CompressedSize As Long

'Path to the zLib (WAPI-variant) dll
Dim PluginPath As String

'API calls for explicitly calling dlls.  This allows us to build DLL paths at runtime, and it also allows
' us to call any DLL we like without first passing them through regsvr32.
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

'Used to quickly check if a file (or folder) exists
Private Const ERROR_SHARING_VIOLATION As Long = 32
Private Declare Function GetFileAttributesW Lib "kernel32" (ByVal lpFileName As Long) As Long

'This initialize routine is particular to this support tool.  It must be called just once, so it can set up the proper zLib path.
Public Sub initializeZLib()

    'Determine where this .exe is located
    PluginPath = App.Path
    If Right(PluginPath, 1) <> "\" Then PluginPath = PluginPath & "\"

End Sub

'Compress a file
Public Function CompressFile(ByVal srcFilename As String, Optional ByVal deleteOriginal As Boolean = True, Optional ByVal DispResults As Boolean = False) As Boolean
    
    'Manually load the DLL from the "PluginPath" folder (should be App.Path\Data\Plugins)
    Dim hLib As Long
    hLib = LoadLibrary(PluginPath & "zlibwapi.dll")

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
    OriginalSize = UBound(DataBytes) + 1

    'Allocate memory for a temporary compression array
    Dim BufferSize As Long
    Dim TempBuffer() As Byte
    BufferSize = UBound(DataBytes) + 1
    BufferSize = BufferSize + (BufferSize * 0.01) + 12
    ReDim TempBuffer(BufferSize)

    'Compress the data using zLib
    Dim result As Long
    'result = compress(TempBuffer(0), BufferSize, DataBytes(0), UBound(DataBytes) + 1)
    'Note: by requesting a maximum compression level (9) we can shave a few bytes off the final compression stream
    result = compress2(TempBuffer(0), BufferSize, DataBytes(0), UBound(DataBytes) + 1, 9)

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
    dstFilename = Left(srcFilename, Len(srcFilename) - 4)
    dstFilename = dstFilename & ".pdc"
    If FileExist(dstFilename) Then Kill dstFilename
    
    'Build our custom compressed file header
    FileHeader.Verification = "THZC"
    FileHeader.OriginalExt = fExtension
    FileHeader.OriginalSize = OriginalSize
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
    If deleteOriginal Then Kill srcFilename

    'Report the compression ratio
    If DispResults Then MsgBox "File compressed from " & OriginalSize & " bytes to " & CompressedSize & " bytes.  Ratio: " & CStr(100 - (100 * (CDbl(CompressedSize) / CDbl(OriginalSize)))) & "%"

    'Return
    CompressFile = True

End Function

'Decompress a file
Public Function DecompressFile(ByVal srcFilename As String, Optional ByVal deleteOriginal As Boolean = True, Optional ByVal DispResults As Boolean = False) As Boolean
    
    'Manually load the DLL from the "PluginPath" folder (should be App.Path\Data\Plugins)
    Dim hLib As Long
    hLib = LoadLibrary(PluginPath & "zlibwapi.dll")

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
    
    'Free the zLib library from memory
    FreeLibrary hLib
    
    'Kill the original compressed file (note: may want to disable when debugging, so no important files are lost)
    If deleteOriginal Then Kill srcFilename
    
    'Display decompression results
    If DispResults Then MsgBox "File decompressed from " & OriginalSize & " bytes to its original size of " & CompressedSize & " bytes."
    
    'Return
    DecompressFile = True
    
End Function

'Function to strip the extension from a filename (taken long ago from the Internet; thank you to whoever wrote it!)
Private Function GetExtension(Filename As String) As String
    
    Dim pathLoc As Long, extLoc As Long
    Dim i As Long, j As Long

    For i = Len(Filename) To 1 Step -1
        If Mid(Filename, i, 1) = "." Then
            extLoc = i
            For j = Len(Filename) To 1 Step -1
                If Mid(Filename, j, 1) = "\" Then
                    pathLoc = j
                    Exit For
                End If
            Next j
            Exit For
        End If
    Next i
    
    If pathLoc > extLoc Then
        GetExtension = ""
    Else
        If extLoc = 0 Then GetExtension = ""
        GetExtension = Mid(Filename, extLoc + 1, Len(Filename) - extLoc)
    End If
            
End Function

'Returns a boolean as to whether or not a given file exists
Private Function FileExist(ByRef fName As String) As Boolean
    Select Case (GetFileAttributesW(StrPtr(fName)) And vbDirectory) = 0
        Case True: FileExist = True
        Case Else: FileExist = (Err.LastDllError = ERROR_SHARING_VIOLATION)
    End Select
End Function
