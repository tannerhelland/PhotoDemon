VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SPBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Note: this file has been modified for use within PhotoDemon.

'This code was originally written by vbforums user "dilettante".

'You may download the original version of this code from the following link (good as of May '14):
' http://www.vbforums.com/showthread.php?660014-VB6-ShellPipe-quot-Shell-with-I-O-Redirection-quot-control

'This class is required by the ShellPipe user control, also written by dilettante.


Option Explicit
'
'SPBuffer
'========
'
'A buffer class for use with stream data that arrives in chunks,
'where the data is to be extracted in "lines" delimited by vbCr
'sequences or in entirety.
'

#Const SPB_DEBUG = False

Private Const CHUNK_SIZE As Long = 1000
Private Const CHUNK_SOFT_LIMIT As Long = CHUNK_SIZE * 10

Private Buffer As String
Private CharsInUse As Long
Private mHasLine As Boolean

Public Sub Append(ByVal Text As String)
    Dim TextLen As Long
    Dim BufferLen As Long
    Dim Temp As String
    
    TextLen = Len(Text)
    If TextLen > 0 Then
        BufferLen = Len(Buffer)
        Do While TextLen + CharsInUse > BufferLen
            BufferLen = BufferLen + CHUNK_SIZE
        Loop
        If BufferLen > Len(Buffer) Then
            Temp = Left$(Buffer, CharsInUse)
            Buffer = String$(BufferLen, 0)
            Mid$(Buffer, 1) = Temp
        End If
        Mid$(Buffer, CharsInUse + 1, TextLen) = Text
        CharsInUse = CharsInUse + TextLen
        mHasLine = mHasLine Or InStr(Text, vbCr) > 0
    End If
End Sub

Public Function Backspace(Optional ByVal Spaces As Long = 1) As Long
    If Spaces > CharsInUse Then Spaces = CharsInUse
    If Spaces > 0 Then
        CharsInUse = CharsInUse - Spaces
        If CharsInUse > 0 Then
            mHasLine = InStrRev(Buffer, vbCr, CharsInUse) > 0
        Else
            mHasLine = False
        End If
    End If
    Backspace = Spaces
End Function

Public Sub Clear()
    If Len(Buffer) > CHUNK_SOFT_LIMIT Then
        Buffer = String$(CHUNK_SOFT_LIMIT, 0)
    End If
    CharsInUse = 0
    mHasLine = False
End Sub

Private Sub DeleteChars(ByVal Length As Long)
    Dim strRemainder As String

    strRemainder = Mid$(Buffer, Length + 1, _
                        CharsInUse - Length)
    CharsInUse = Len(strRemainder)
    Mid$(Buffer, 1, CharsInUse) = strRemainder
    mHasLine = InStr(strRemainder, vbCr) > 0
End Sub

Public Sub DeleteData(ByVal MaxLen As Long)
    If MaxLen > 0 Then
        If MaxLen > CharsInUse Then MaxLen = CharsInUse
        DeleteChars MaxLen
    End If
End Sub

Public Function GetData(Optional ByVal MaxLen As Long = -1) As String
    If MaxLen < 0 Then
        GetData = Left$(Buffer, CharsInUse)
        Clear
    ElseIf MaxLen = 0 Then
        GetData = ""
    Else
        If MaxLen > CharsInUse Then MaxLen = CharsInUse
        GetData = Left$(Buffer, MaxLen)
        DeleteChars MaxLen
    End If
End Function

Public Function GetLine() As String
    'A "line" is anything delimited by a CR, but also filter
    'out any LF characters for cases where CRLF delimits lines.
    'We see both conventions used in StdIO programs.
    Dim lngLineEnd As Long
    
    If mHasLine Then
        If Left$(Buffer, 1) = vbLf Then DeleteChars 1 'LF arrived late.
        lngLineEnd = InStr(Buffer, vbCr)
        GetLine = Left$(Buffer, lngLineEnd - 1)
        DeleteChars lngLineEnd
        If CharsInUse > 0 Then
            If Left$(Buffer, 1) = vbLf Then DeleteChars 1 'LF on time.
        End If
    Else
        GetLine = ""
    End If
End Function

Public Property Get HasLine() As Boolean
    HasLine = mHasLine
End Property

Public Property Get Length() As Long
    Length = CharsInUse
End Property

Public Sub PeekBuffer(ByRef Data As String)
    Data = Left$(Buffer, CharsInUse)
End Sub

#If SPB_DEBUG Then
Public Property Get PeekDebug() As String
    'MEANT ONLY FOR DEBUGGING.
    'DUMPS BUFFER IN HEX.
    Dim CharX As Long

    PeekDebug = ""
    For CharX = 1 To CharsInUse
        PeekDebug = PeekDebug _
             & Right$("0" & Hex$(Asc(Mid$(Buffer, CharX, 1))), 2) & " "
    Next
End Property
#End If

Private Sub Class_Initialize()
    Clear
End Sub
