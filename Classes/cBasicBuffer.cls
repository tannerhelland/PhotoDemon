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
'Note: this file has been modified for use within PhotoDemon.

'This class was originally written by vbforums.com user dilettante.

'You may download the original version of this code from the following link (good as of April '14):
' http://www.vbforums.com/showthread.php?710155-VB6-BasicBuffer-Binary-Stream-Class

'To the best of my knowledge, this code has been released into the public domain.

'Many thanks to dilettante for this excellent minimal byte buffer interface, which PhotoDemon uses to simplify the process of
' reading/writing various internal file formats (including PDI files).


Option Explicit
'
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
    
    mPosition = mSize - 1
    If UBound(Buffer) > mPosition Then ShrinkBuf
    CurrentBytes = Buffer
End Function

Public Sub DeleteBytes(Optional ByVal Length As Long = -1)
    'Delete requested number of bytes (or less if less in buffer) or
    'all bytes starting at Position.  "All bytes", Length = -1 (or < 0)
    'is the same as calling SetEOS().
    
    If mOpen Then
        DeleteBytesFrom mPosition, Length
    Else
        Err.Raise &H8004C300, TypeName(Me), "Not open"
    End If
End Sub

Public Function ExtractBytes(Optional ByVal Length As Long = -1) As Byte()
    'Combine ReadBytes() and DeleteBytes() of same range of bytes.
    '
    'Not highly efficient, but some applications need to grab
    'things off the front as much as adding them to the end
    'without having the Buffer array grow too large.
    Dim CurrPos As Long
    
    If mOpen Then
        CurrPos = mPosition
        ExtractBytes = ReadBytes(Length)
        mPosition = CurrPos
        DeleteBytesFrom CurrPos, Length
    Else
        Err.Raise &H8004C300, TypeName(Me), "Not open"
    End If
End Function

Public Function ExtractInteger() As Integer
    'Combine ReadInteger() and DeleteBytes() of same range of bytes.
    Dim CurrPos As Long
    
    If mOpen Then
        CurrPos = mPosition
        ExtractInteger = ReadInteger()
        mPosition = CurrPos
        DeleteBytesFrom CurrPos, Len(ExtractInteger)
    Else
        Err.Raise &H8004C300, TypeName(Me), "Not open"
    End If
End Function

Public Function ExtractLong() As Long
    'Combine ReadLong() and DeleteBytes() of same range of bytes.
    Dim CurrPos As Long
    
    If mOpen Then
        CurrPos = mPosition
        ExtractLong = ReadLong()
        mPosition = CurrPos
        DeleteBytesFrom CurrPos, Len(ExtractLong)
    Else
        Err.Raise &H8004C300, TypeName(Me), "Not open"
    End If
End Function

Public Sub OpenBuf()
    If mOpen Then
        Err.Raise &H8004C302, TypeName(Me), "Already open"
    Else
        mOpen = True
        ReDim Buffer(mChunkSize - 1)
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

Public Function ReadInteger() As Integer
    'Return next 2 bytes as an Integer.
    
    If mOpen Then
        If HaveAtLeast(2) Then
            MoveMemory ReadInteger, Buffer(mPosition), 2
            mPosition = mPosition + 2
        Else
            Err.Raise &H8004C304, TypeName(Me), "Need 2 bytes in buffer"
        End If
    Else
        Err.Raise &H8004C300, TypeName(Me), "Not open"
    End If
End Function

Public Function ReadLong() As Long
    'Return next 4 bytes as a Long.
    
    If mOpen Then
        If HaveAtLeast(4) Then
            MoveMemory ReadLong, Buffer(mPosition), 4
            mPosition = mPosition + 4
        Else
            Err.Raise &H8004C304, TypeName(Me), "Need 4 bytes in buffer"
        End If
    Else
        Err.Raise &H8004C300, TypeName(Me), "Not open"
    End If
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
    'Trim the contents of the buffer from Position to current end,
    'making Size smaller.
    
    If mOpen Then
        mSize = mPosition
    Else
        Err.Raise &H8004C300, TypeName(Me), "Not open"
    End If
End Sub

Public Sub ShrinkBuf()
    'Trims Buffer back to current size actually in use.
    '
    'Rarely needed by itself.  Use only to reclaim "wasted" space
    'in the Buffer array if it will hang around for a while.
    '
    'Also called by CurrentBytes().
    
    ReDim Preserve Buffer(mSize - 1)
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
    NeedAtLeastUBound = mPosition + WriteCountMinus1
    UBoundOfBuffer = UBound(Buffer)
    
    If NeedAtLeastUBound > UBoundOfBuffer Then
        'Extend Buffer to make room, by at least ChunkSize.
        If WriteCount > mChunkSize Then
            ReDim Preserve Buffer(NeedAtLeastUBound)
        Else
            ReDim Preserve Buffer(UBoundOfBuffer + mChunkSize)
            'UBoundOfBuffer is now invalid, but we won't use it again.
        End If
    End If
    
    MoveMemory Buffer(mPosition), Bytes(LBoundOfBytes), WriteCount
    mPosition = mPosition + WriteCount
    If mPosition > mSize Then
        mSize = mPosition
    End If
End Sub

'Added by Tanner: add bytes to the buffer, using only a raw pointer (to the first entry of the array, not the SafeArray
' struct itself!) and a length, *which must be byte-accurate*!.
Public Sub WriteBytesFromPointer(ByVal dataPointer As Long, ByVal dataLength As Long)
    
    'Copy provided bytes into Buffer at current Position, extending
    'Size as required, updating Position to next byte offset.
    Dim WriteCountMinus1 As Long
    Dim WriteCount As Long
    Dim NeedAtLeastUBound As Long
    Dim UBoundOfBuffer As Long
    
    'If not open, first call here auto-opens.
    If Not mOpen Then OpenBuf
        
    'Length is supplied by the user; pray they calculated it correctly!
    WriteCountMinus1 = dataLength - 1
    
    'Original code follows...
    
    WriteCount = WriteCountMinus1 + 1
    NeedAtLeastUBound = mPosition + WriteCountMinus1
    UBoundOfBuffer = UBound(Buffer)
    
    If NeedAtLeastUBound > UBoundOfBuffer Then
        
        'Extend Buffer to make room, by at least ChunkSize.
        If WriteCount > mChunkSize Then
            ReDim Preserve Buffer(NeedAtLeastUBound)
        Else
            ReDim Preserve Buffer(UBoundOfBuffer + mChunkSize)
            'UBoundOfBuffer is now invalid, but we won't use it again.
        End If
        
    End If
    
    MoveMemory Buffer(mPosition), ByVal dataPointer, WriteCount
    mPosition = mPosition + WriteCount
    If mPosition > mSize Then
        mSize = mPosition
    End If
    
End Sub

Public Sub WriteInteger(ByVal Value As Integer)
    'Copy Value into Buffer at current Position, extending
    'Size as required, updating Position to next byte offset.
    Dim NeedAtLeastUBound As Long
    Dim UBoundOfBuffer As Long
    
    'If not open, first call here auto-opens.
    If Not mOpen Then OpenBuf
    
    NeedAtLeastUBound = mPosition + 1
    UBoundOfBuffer = UBound(Buffer)
    If NeedAtLeastUBound > UBoundOfBuffer Then
        'Extend Buffer to make room, by at least ChunkSize.
        If 2 > mChunkSize Then
            ReDim Preserve Buffer(NeedAtLeastUBound)
        Else
            ReDim Preserve Buffer(UBoundOfBuffer + mChunkSize)
            'UBoundOfBuffer is now invalid, but we won't use it again.
        End If
    End If
    
    MoveMemory Buffer(mPosition), Value, 2
    mPosition = mPosition + 2
    If mPosition > mSize Then
        mSize = mPosition
    End If
End Sub

Public Sub WriteLong(ByVal Value As Long)
    'Copy Value into Buffer at current Position, extending
    'Size as required, updating Position to next byte offset.
    Dim NeedAtLeastUBound As Long
    Dim UBoundOfBuffer As Long
    
    'If not open, first call here auto-opens.
    If Not mOpen Then OpenBuf
    
    NeedAtLeastUBound = mPosition + 3
    UBoundOfBuffer = UBound(Buffer)
    If NeedAtLeastUBound > UBoundOfBuffer Then
        'Extend Buffer to make room, by at least ChunkSize.
        If 4 > mChunkSize Then
            ReDim Preserve Buffer(NeedAtLeastUBound)
        Else
            ReDim Preserve Buffer(UBoundOfBuffer + mChunkSize)
            'UBoundOfBuffer is now invalid, but we won't use it again.
        End If
    End If
    
    MoveMemory Buffer(mPosition), Value, 4
    mPosition = mPosition + 4
    If mPosition > mSize Then
        mSize = mPosition
    End If
End Sub

Private Sub DeleteBytesFrom(ByVal Pos As Long, ByVal Length As Long)
    'Deletes requested number of bytes at requested Pos.
    '
    'Called by DeleteBytes() and by ExtractBytes().
    Dim BytesToMove As Long
        
    If Length < 0 Then
        mSize = Pos
    ElseIf Length > 0 Then
        If Pos + Length > mSize Then
            mSize = Pos
        Else
            BytesToMove = mSize - (Pos + Length)
            If BytesToMove > 0 Then
                MoveMemory Buffer(Pos), Buffer(Pos + Length), BytesToMove
            End If
            mSize = mSize - Length
        End If
    End If
End Sub

Private Sub Class_Initialize()
    mChunkSize = ALLOCATION_CHUNK_DEFAULT
End Sub
