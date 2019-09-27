Attribute VB_Name = "PDPackaging"
Option Explicit

'When sending data to a pdPackageChunky instance, the caller must specify what type of pointer they're passing.
' pdPackageChunky will automatically optimize the storage format for the passed data (e.g. strings get
' auto-converted to UTF-8 at write-time, and auto-converted back to UTF-16 at read-time).
Public Enum PDPackage_DataType
    dt_Bytes = 0
    dt_StringToUTF8 = 1
End Enum

#If False Then
    Private Const dt_Bytes = 0, dt_StringToUTF8 = 1
#End If

Public Type PDPackage_ChunkData
    chunkID As String
    ptrChunkData As Long
    chunkDataLength As Long
    chunkDataFormat As PDPackage_DataType
    chunkCompressionFormat As PD_CompressionFormat
    chunkCompressionLevel As Long
End Type

'Return a PDPackage_ChunkData struct filled with the passed data; this simplifies the process of passing
' the data to a pdPackageChunky instance.
Public Function GetChunkDataStruct(ByVal srcChunkID As String, ByVal srcPtrChunkData As Long, ByVal srcChunkDataSizeBytes As Long, Optional ByVal srcChunkDataType As PDPackage_DataType = dt_Bytes, Optional ByVal dstCompressionFormat As PD_CompressionFormat = cf_None, Optional ByVal dstCompressionLevel As Long = -1) As PDPackage_ChunkData
    With GetChunkDataStruct
        .chunkID = ValidateChunkID(srcChunkID)
        .ptrChunkData = srcPtrChunkData
        .chunkDataLength = srcChunkDataSizeBytes
        .chunkDataFormat = srcChunkDataType
        .chunkCompressionFormat = dstCompressionFormat
        .chunkCompressionLevel = dstCompressionLevel
    End With
End Function

'Convert an arbitrary string into a usable 4-char chunk ID
Public Function ValidateChunkID(ByRef srcID As String) As String
    If (Len(srcID) = 4) Then
        ValidateChunkID = srcID
    ElseIf (Len(srcID) > 4) Then
        PDDebug.LogAction "ValidateChunkID reports: " & srcID & " is an invalid chunk ID - only 4 chars allowed.  ID will be truncated."
        ValidateChunkID = Left$(srcID, 4)
    
    'Lengths less than 4 are fine; they're just extended with spaces
    Else
        ValidateChunkID = srcID & String$(4 - Len(srcID), " ")
    End If
End Function
