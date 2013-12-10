VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SmartBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Note: this file has been modified for use within PhotoDemon.

'This code was originally written by vbforums user "dilettante".

'You may download the original version of this code from the following link (good as of November '13):
' http://www.vbforums.com/showthread.php?660014-VB6-ShellPipe-quot-Shell-with-I-O-Redirection-quot-control

'This class is required by the ShellPipe user control, also written by dilettante.


Option Explicit
'
'SmartBuffer
'===========
'
'A buffer class for use with stream data that arrives in chunks,
'where the data is to be extracted in "lines" delimited by vbCr
'sequences or in entirety.
'

Private m_strBuffer As String
Private m_lngInUse As Long
Private m_blnHasLine As Boolean
Private Const CHUNK_SIZE As Long = 32768
Private Const CHUNK_SOFT_LIMIT As Long = CHUNK_SIZE * 10

Public Sub Append(ByVal Text As String)
    Dim lngTextLen As Long
    Dim lngBufferLen As Long
    Dim strTemp As String
    
    lngTextLen = Len(Text)
    If lngTextLen > 0 Then
        lngBufferLen = Len(m_strBuffer)
        Do While lngTextLen + m_lngInUse > lngBufferLen
            lngBufferLen = lngBufferLen + CHUNK_SIZE
        Loop
        If lngBufferLen > Len(m_strBuffer) Then
            strTemp = Left$(m_strBuffer, m_lngInUse)
            m_strBuffer = String$(lngBufferLen, 0)
            Mid$(m_strBuffer, 1) = strTemp
        End If
        Mid$(m_strBuffer, m_lngInUse + 1, lngTextLen) = Text
        m_lngInUse = m_lngInUse + lngTextLen
        m_blnHasLine = m_blnHasLine Or InStr(Text, vbCr) > 0
    End If
End Sub

Public Function Backspace(Optional ByVal Spaces As Long = 1) As Long
    If Spaces > m_lngInUse Then Spaces = m_lngInUse
    If Spaces > 0 Then
        m_lngInUse = m_lngInUse - Spaces
        If m_lngInUse > 0 Then
            m_blnHasLine = InStrRev(m_strBuffer, vbCr, m_lngInUse) > 0
        Else
            m_blnHasLine = False
        End If
    End If
    Backspace = Spaces
End Function

Private Sub DeleteChars(ByVal Length As Long)
    Dim strRemainder As String

    strRemainder = Mid$(m_strBuffer, Length + 1, _
                        m_lngInUse - Length)
    m_lngInUse = Len(strRemainder)
    Mid$(m_strBuffer, 1, m_lngInUse) = strRemainder
    m_blnHasLine = InStr(strRemainder, vbCr) > 0
End Sub

Public Sub DeleteData(ByVal MaxLen As Long)
    If MaxLen > 0 Then
        If MaxLen > m_lngInUse Then MaxLen = m_lngInUse
        DeleteChars MaxLen
    End If
End Sub

Public Sub Flush()
    If Len(m_strBuffer) > CHUNK_SOFT_LIMIT Then
        m_strBuffer = String$(CHUNK_SOFT_LIMIT, 0)
    End If
    m_lngInUse = 0
    m_blnHasLine = False
End Sub

Public Function GetData(Optional ByVal MaxLen As Long = -1) As String
    If MaxLen < 0 Then
        GetData = Left$(m_strBuffer, m_lngInUse)
        Flush
    ElseIf MaxLen = 0 Then
        GetData = ""
    Else
        If MaxLen > m_lngInUse Then MaxLen = m_lngInUse
        GetData = Left$(m_strBuffer, MaxLen)
        DeleteChars MaxLen
    End If
End Function

Public Function GetLine() As String
    'A "line" is anything delimited by a CR, but also filter
    'out any LF characters for cases where CRLF delimits lines.
    'We see both conventions used in StdIO programs.
    Dim lngLineEnd As Long
    
    If m_blnHasLine Then
        If Left$(m_strBuffer, 1) = vbLf Then DeleteChars 1 'LF arrived late.
        lngLineEnd = InStr(m_strBuffer, vbCr)
        GetLine = Left$(m_strBuffer, lngLineEnd - 1)
        DeleteChars lngLineEnd
        If m_lngInUse > 0 Then
            If Left$(m_strBuffer, 1) = vbLf Then DeleteChars 1 'LF on time.
        End If
    Else
        GetLine = ""
    End If
End Function

Public Property Get HasLine() As Boolean
    HasLine = m_blnHasLine
End Property

Public Property Get Length() As Long
    Length = m_lngInUse
End Property

Public Sub PeekBuffer(ByRef Data As String)
    Data = Left$(m_strBuffer, m_lngInUse)
End Sub
'
'Public Property Get PeekDebug() As String
'    'MEANT ONLY FOR DEBUGGING.
'    'DUMPS BUFFER IN HEX.
'    Dim lngChar As Long
'
'    PeekDebug = ""
'    For lngChar = 1 To m_lngInUse
'        PeekDebug = PeekDebug _
'             & Right$("0" & Hex$(Asc(Mid$(m_strBuffer, lngChar, 1))), 2) & " "
'    Next
'End Property

Private Sub Class_Initialize()
    Flush
End Sub
