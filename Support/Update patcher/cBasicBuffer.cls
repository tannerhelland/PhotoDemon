VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBasicBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Note: this file has been heavily modified for use within PhotoDemon.

'You may download the original version of this code from the following link (good as of April '14):
' http://www.vbforums.com/showthread.php?710155-VB6-BasicBuffer-Binary-Stream-Class

'This class was originally written by vbforums.com user dilettante.  To the best of my knowledge, the code has
' been released into the public domain.

'Many thanks to dilettante for this excellent minimal byte buffer interface, which PhotoDemon uses to simplify
' the process of reading/writing various internal file formats (including PDI files).


Option Explicit
Option Base 0


'A simple stream-style buffer class.
'
'This is a kind of binary stream, similar to an ADO Steam object
'in Type = adBinary mode or an OLE IStream object.  It accepts
'and returns Byte arrays, Integers, and Longs but not text String
'values.
'
'It can be useful whenever you want a data structure that supports
'something similar to concatenating Byte arrays when you need to
'accumulate data arriving in bits and pieces.
'
'The properties and methods are similar to those of an ADO Stream.
'
'Notes
'-----
'
'   o You can WriteBytes() passing a fixed length array but
'     ReadBytes() and ExtractBytes() return a value that must be
'     assigned to a dynamic array.
'
'   o ChunkSize is an important property to set for optimal
'     performance.  See the comments below, just above its
'     definition.
'
'This class has a lot of things in it to handle common needs.
'Remove anything you don't need in your programs.
'
'Properties
'----------
'
'ChunkSize As Long [R/W]
'EOS As Boolean [RO]
'HaveAtLeast(Length As Long) As Boolean [RO]
'IsOpen As Boolean [RO]
'Position As Long [R/W]
'Size As Long [RO]
'
'Methods
'-------
'
'CloseBuf()
'CurrentBytes() As Byte()
'DeleteBytes(Optional ByVal Length As Long = -1)
'ExtractBytes(Optional ByVal Length As Long = -1) As Byte()
'ExtractInteger() As Integer
'ExtractLong() As Long
'OpenBuf()
'ReadBytes(Optional ByVal Length As Long = -1) As Byte()
'ReadInteger() As Integer
'ReadLong() As Long
'ScanForBytes(ByRef Target() As Byte) As Long
'SetEOS()
'ShrinkBuf()
'WriteBytes(ByRef Bytes() As Byte)
'WriteInteger(ByVal Value As Integer)
'WriteLong(ByVal Value As Long)
'

'This class can be made faster by using a typelib to define this
'API entrypoint.  Edanmo's olelib.tlb is quite popular and used
'in many sophisticated VB6 programs.  It uses this same signature
'(below) so if you have referenced olelib.tlb you can just
'comment this out or delete it:
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef Destination As Any, _
    ByRef Source As Any, _
    ByVal Length As Long)

Private Const ALLOCATION_CHUNK_DEFAULT As Long = 8192

Private Buffer() As Byte

Private mChunkSize As Long
Private mOpen As Boolean
Private mPosition As Long 'Base 0 "cursor."
Private mSize As Long 'Base 1 number of bytes of live data in the buffer.

'ChunkSize
'
'Minimum number of bytes to expand Buffer by when space is needed.
'This avoids expensive concatenation as much as possible.  Can be
'changed while open but normally set once before opening or first
'WriteBytes() call.
'
'Choose a value that is larger than your expected "write sizes"
'will be, to reduce the number of times that a Redim Preserve of
'Buffer will need to be done.  Making it too large may result in
'excessive space allocation though.
'
'If you expect to buffer a total of about "X" bytes and you tend
'to write in chunks of (X / 4) or less you might consider a value
'of (X / 2) or even (3 * X / 4) for ChunkSize.
Public Property Get ChunkSize() As Long
    ChunkSize = mChunkSize
End Property

Public Property Let ChunkSize(ByVal RHS As Long)
    If 1 <= RHS Then
        mChunkSize = RHS
    Else
        Err.Raise 5, TypeName(Me), "ChunkSize must be at least 1 byte"
    End If
End Property

Public Property Get EOS() As Boolean
    'Also called by ReadBytes().
    
    EOS = mPosition >= mSize
End Property

Public Property Get HaveAtLeast(ByVal Length As Long) As Boolean
    HaveAtLeast = mSize - mPosition >= Length
End Property

Public Property Get IsOpen() As Boolean
    IsOpen = mOpen
End Property

'Position is 0-based, ranges from 0 to Size.  This is where the
'next read or write operation begins.
Public Property Get Position() As Long
    Position = mPosition
End Property

Public Property Let Position(ByVal RHS As Long)
    If 0 <= RHS And RHS <= mSize Then
        mPosition = RHS
    Else
        Err.Raise 5, TypeName(Me), "Can't set Position before beginning or after Size"
    End If
End Property

'Size is a 1-based "count" of live bytes in the buffer.  You
'can move to the end by setting .Position = .Size, which can be
'useful.
Public Property Get Size() As Long
    Size = mSize
End Property

Public Sub CloseBuf()
    'Get rid of Buffer contents, prepare for new open and reuse.
    
    Erase Buffer
    mOpen = False
    mPosition = 0
    mSize = 0
End Sub

Public Function CurrentBytes() As Byte()
    'Used when you want to get the entire current data as a Byte
    'array.
    '
    'Note that it calls ShrinkBuf() if needed.  This can have
    'performance implications if you will do further writes!
    '
    'You can set .Position = 0 and then call .SetEOS() to "empty"
    'the buffer after calling this.  That is more efficient than
    'calling CloseBuf() afterward if you will reuse the
    'BasicBuffer for more data as a sort of "ExtractCurrentBytes"
    'operation because it keeps the Buffer array intact for reuse.
    
    mPosition = mSize
    If UBound(Buffer) > mPosition Then ShrinkBuf
    
    CurrentBytes = Buffer
    
End Function

'Added by Tanner: return a pointer to the first position in the array.
Public Function CurrentBytesPointer() As Long
    CurrentBytesPointer = VarPtr(Buffer(0))
End Function

'Open a new buffer.
' Added by Tanner: allow the caller to specify a starting buffer size.  This can be helpful if we have some idea
'  of the buffer's size in advance, as we avoid the need for costly ReDim Preserve statements.
Public Sub OpenBuf(Optional ByVal startingBufferSize As Long = 0)

    If mOpen Then
        Err.Raise &H8004C302, TypeName(Me), "Already open"
    Else
        mOpen = True
        If startingBufferSize <= 0 Then
            ReDim Buffer(0 To mChunkSize - 1) As Byte
        Else
            ReDim Buffer(0 To startingBufferSize - 1) As Byte
        End If
    End If
    
End Sub

Public Function ReadBytes(Optional ByVal Length As Long = -1) As Byte()
    'Return requested number of bytes (or less if less in buffer) or
    'all bytes.
    '
    'Requesting Length = 0 bytes (or any bytes when the buffer is empty)
    'raises an exception since we can't return an "empty" Byte array.
    '
    'ReadBytes() leaves the bytes in the buffer, it just copies them
    'and adjusts Position.  Use ExtractBytes() if you want to "pull"
    'bytes from the buffer as you go - but it will be slower.
    Dim Bytes() As Byte
    
    If mOpen Then
        If Length < 0 Then
            'A Length = -1 (or < 0) means read all.
            If EOS Then
                Err.Raise 5, TypeName(Me), "Nothing to read"
            Else
                Length = mSize - mPosition
            End If
        ElseIf Length > 0 Then
            If mPosition + Length > mSize Then
                'Handle requests extending past EOS.
                Length = mSize - mPosition
            End If
        Else
            Err.Raise 5, TypeName(Me), "Can't request 0 (no) bytes"
        End If
    Else
        Err.Raise &H8004C300, TypeName(Me), "Not open"
    End If
    
    ReDim Bytes(Length - 1)
    MoveMemory Bytes(0), Buffer(mPosition), Length
    mPosition = mPosition + Length
    ReadBytes = Bytes
End Function

Public Function ScanForBytes(ByRef Target() As Byte) As Long
    'Scan forward from Position, looking for Target value which can be
    'one or more bytes.
    '
    'If found returns position index of first byte of Target, else
    'returns -1.
    
    If mOpen Then
        ScanForBytes = InStrB(mPosition + 1, Buffer, Target) - 1
        If ScanForBytes > mSize - (UBound(Target) - LBound(Target) + 1) Then ScanForBytes = -1
    Else
        Err.Raise &H8004C300, TypeName(Me), "Not open"
    End If
End Function

Public Sub SetEOS()
    'Trim the contents of Buffer() from Position to current end, making the Size marker smaller,
    ' *but not actually resizing the array*.
    
    If mOpen Then
        mSize = mPosition
    Else
        Err.Raise &H8004C300, TypeName(Me), "Not open"
    End If
    
End Sub

'Trim Buffer to its smallest possible size (given the current data)
Public Sub ShrinkBuf()
    ReDim Preserve Buffer(0 To mSize) As Byte
End Sub

Public Sub WriteBytes(ByRef Bytes() As Byte, Optional ByVal Length As Long = -1)
    'Copy provided bytes into Buffer at current Position, extending
    'Size as required, updating Position to next byte offset.
    Dim LBoundOfBytes As Long
    Dim WriteCountMinus1 As Long
    Dim WriteCount As Long
    Dim NeedAtLeastUBound As Long
    Dim UBoundOfBuffer As Long
    
    'If not open, first call here auto-opens.
    If Not mOpen Then OpenBuf
    
    LBoundOfBytes = LBound(Bytes)
    
    'Change by Tanner: If the user has specified a Length value, use that instead of the array's upper bound.
    If Length >= 0 Then
        WriteCountMinus1 = LBoundOfBytes + Length - 1
    Else
        WriteCountMinus1 = UBound(Bytes) - LBoundOfBytes
    End If
    
    'Original code follows...
    
    WriteCount = WriteCountMinus1 + 1
    NeedAtLeastUBound = mPosition + WriteCount
    UBoundOfBuffer = UBound(Buffer)
    
    If NeedAtLeastUBound > UBoundOfBuffer Then
        'Extend Buffer to make room, by at least ChunkSize.
        If WriteCount > mChunkSize Then
            ReDim Preserve Buffer(0 To NeedAtLeastUBound) As Byte
        Else
            ReDim Preserve Buffer(0 To UBoundOfBuffer + mChunkSize) As Byte
            'UBoundOfBuffer is now invalid, but we won't use it again.
        End If
    End If
    
    MoveMemory Buffer(mPosition), Bytes(LBoundOfBytes), WriteCount
    mPosition = mPosition + WriteCount
    If mPosition > mSize Then mSize = mPosition
    
End Sub

'Added by Tanner: add bytes to the buffer, using only a raw pointer (to the first entry of the array, not the SafeArray
' struct itself!) and a length, *which must be byte-accurate*!
Public Sub WriteBytesFromPointer(ByVal dataPointer As Long, ByVal dataLength As Long)
    
    Dim NeedAtLeastUBound As Long
    Dim UBoundOfBuffer As Long
    
    'If not open, first call here auto-opens.
    If Not mOpen Then OpenBuf
        
    'Length is supplied by the user; pray they calculated it correctly!
    NeedAtLeastUBound = mPosition + dataLength
    UBoundOfBuffer = UBound(Buffer)
    
    'If the buffer is not large enough to hold the data, resize it by the larger of mChunkSize or the size of the data.
    If NeedAtLeastUBound > UBoundOfBuffer Then
        
        If dataLength > mChunkSize Then
            ReDim Preserve Buffer(0 To NeedAtLeastUBound) As Byte
        Else
            ReDim Preserve Buffer(0 To UBoundOfBuffer + mChunkSize) As Byte
        End If
        
    End If
    
    'Copy the data into place
    MoveMemory Buffer(mPosition), ByVal dataPointer, dataLength
    
    'Increment the buffer pointer
    mPosition = mPosition + dataLength
    
    'Note the actual size of the current data set
    If mPosition > mSize Then mSize = mPosition
    
End Sub

Private Sub Class_Initialize()
    mChunkSize = ALLOCATION_CHUNK_DEFAULT
End Sub
